import { Prisma } from '@prisma/client';
import { prisma } from '../../main';
import { AppError } from '../../middleware/error-handler';
import { createAuditLog } from '../../common/audit';
import { APP_CONFIG } from '@vybe/config';
import { v4 as uuidv4 } from 'uuid';
import crypto from 'crypto';

export const contentService = {
  // ─── Media Upload ───────────────────────────────────────
  async createUploadIntent(userId: string, data: { filename: string; mimeType: string; sizeBytes: number }) {
    const kind = data.mimeType.startsWith('image/') ? 'image'
      : data.mimeType.startsWith('video/') ? 'video'
      : data.mimeType.startsWith('audio/') ? 'audio'
      : null;

    if (!kind) {
      throw new AppError(400, 'UNSUPPORTED_MEDIA', 'Unsupported media type');
    }

    if (data.sizeBytes > APP_CONFIG.MAX_UPLOAD_SIZE_BYTES) {
      throw new AppError(400, 'FILE_TOO_LARGE', 'File exceeds maximum upload size');
    }

    const storageKey = `uploads/${userId}/${uuidv4()}/${data.filename}`;
    const checksum = crypto.randomBytes(16).toString('hex'); // Placeholder

    const asset = await prisma.mediaAsset.create({
      data: {
        ownerId: userId,
        kind,
        mimeType: data.mimeType,
        sizeBytes: BigInt(data.sizeBytes),
        storageKey,
        checksum,
        status: 'pending',
      },
    });

    // In production, generate pre-signed S3 URL
    const uploadUrl = `${process.env.S3_ENDPOINT || 'http://localhost:9000'}/${process.env.S3_BUCKET || 'vybe-media'}/${storageKey}`;

    return {
      assetId: asset.id,
      uploadUrl,
      expiresAt: new Date(Date.now() + 30 * 60 * 1000).toISOString(),
    };
  },

  async completeUpload(userId: string, assetId: string, data?: { width?: number; height?: number; durationMs?: number }) {
    const asset = await prisma.mediaAsset.findFirst({
      where: { id: assetId, ownerId: userId },
    });
    if (!asset) {
      throw new AppError(404, 'ASSET_NOT_FOUND', 'Media asset not found');
    }

    return prisma.mediaAsset.update({
      where: { id: assetId },
      data: {
        status: 'ready',
        width: data?.width,
        height: data?.height,
        durationMs: data?.durationMs,
      },
    });
  },

  // ─── Content CRUD ───────────────────────────────────────
  async createContent(userId: string, data: {
    type: string;
    caption?: string;
    audience: string;
    audienceCircleId?: string;
    mediaIds: string[];
    allowComments?: boolean;
    allowSharing?: boolean;
    allowDownload?: boolean;
    publishNow?: boolean;
  }) {
    // Verify media assets belong to user
    if (data.mediaIds.length > 0) {
      const assets = await prisma.mediaAsset.findMany({
        where: { id: { in: data.mediaIds }, ownerId: userId },
      });
      if (assets.length !== data.mediaIds.length) {
        throw new AppError(400, 'INVALID_MEDIA', 'Some media assets not found or not owned by you');
      }
    }

    // Extract hashtags from caption
    const hashtags = data.caption
      ? (data.caption.match(/#[\w]+/g) || []).map((t) => t.toLowerCase())
      : [];

    const content = await prisma.$transaction(async (tx) => {
      const item = await tx.contentItem.create({
        data: {
          authorId: userId,
          type: data.type,
          status: data.publishNow ? 'published' : 'draft',
          audience: data.audience,
          audienceCircleId: data.audienceCircleId,
          caption: data.caption || '',
          hashtags: JSON.stringify(hashtags),
          allowComments: data.allowComments ?? true,
          allowSharing: data.allowSharing ?? true,
          allowDownload: data.allowDownload ?? false,
          publishedAt: data.publishNow ? new Date() : null,
        },
      });

      // Link media
      for (let i = 0; i < data.mediaIds.length; i++) {
        await tx.contentMedia.create({
          data: {
            contentId: item.id,
            mediaId: data.mediaIds[i],
            position: i,
            role: i === 0 ? 'primary' : 'attachment',
          },
        });
      }

      if (data.publishNow) {
        await tx.userProfile.update({
          where: { userId },
          data: { postsCount: { increment: 1 } },
        });
      }

      return item;
    });

    return this.getContentById(content.id, userId);
  },

  async getContentById(contentId: string, viewerId?: string) {
    const content = await prisma.contentItem.findUnique({
      where: { id: contentId },
      include: {
        author: { include: { profile: true } },
        contentMedia: {
          include: { media: true },
          orderBy: { position: 'asc' },
        },
      },
    });

    if (!content || content.deletedAt) {
      throw new AppError(404, 'CONTENT_NOT_FOUND', 'Content not found');
    }

    // Check visibility
    if (viewerId && viewerId !== content.authorId) {
      const blocked = await prisma.blockEdge.findFirst({
        where: {
          OR: [
            { blockerId: content.authorId, blockedId: viewerId },
            { blockerId: viewerId, blockedId: content.authorId },
          ],
        },
      });
      if (blocked) {
        throw new AppError(403, 'BLOCKED', 'Cannot view this content');
      }
    }

    let isLiked = false;
    let isSaved = false;
    if (viewerId) {
      const reaction = await prisma.reaction.findUnique({
        where: { actorId_targetType_targetId: { actorId: viewerId, targetType: 'content', targetId: contentId } },
      });
      isLiked = !!reaction;

      const save = await prisma.saveItem.findUnique({
        where: { userId_contentId: { userId: viewerId, contentId } },
      });
      isSaved = !!save;
    }

    return {
      id: content.id,
      authorId: content.authorId,
      author: {
        id: content.author.id,
        username: content.author.username,
        displayName: content.author.displayName,
        avatarUrl: content.author.profile?.avatarMediaId,
      },
      type: content.type,
      status: content.status,
      audience: content.audience,
      caption: content.caption,
      hashtags: content.hashtags,
      location: content.location,
      allowComments: content.allowComments,
      allowSharing: content.allowSharing,
      allowDownload: content.allowDownload,
      media: content.contentMedia.map((cm) => ({
        id: cm.media.id,
        kind: cm.media.kind,
        mimeType: cm.media.mimeType,
        width: cm.media.width,
        height: cm.media.height,
        durationMs: cm.media.durationMs,
        url: `${process.env.S3_ENDPOINT || 'http://localhost:9000'}/${process.env.S3_BUCKET || 'vybe-media'}/${cm.media.storageKey}`,
        status: cm.media.status,
      })),
      likesCount: content.likesCount,
      commentsCount: content.commentsCount,
      sharesCount: content.sharesCount,
      savesCount: content.savesCount,
      isLiked,
      isSaved,
      publishedAt: content.publishedAt?.toISOString(),
      createdAt: content.createdAt.toISOString(),
    };
  },

  async updateContent(userId: string, contentId: string, data: {
    caption?: string;
    audience?: string;
    mediaIds?: string[];
  }) {
    const content = await prisma.contentItem.findFirst({
      where: { id: contentId, authorId: userId, deletedAt: null },
    });
    if (!content) {
      throw new AppError(404, 'CONTENT_NOT_FOUND', 'Content not found');
    }

    const hashtags = data.caption
      ? (data.caption.match(/#[\w]+/g) || []).map((t) => t.toLowerCase())
      : undefined;

    await prisma.contentItem.update({
      where: { id: contentId },
      data: {
        caption: data.caption,
        audience: data.audience,
        hashtags: hashtags ? JSON.stringify(hashtags) : undefined,
      },
    });

    return this.getContentById(contentId, userId);
  },

  async publishContent(userId: string, contentId: string) {
    const content = await prisma.contentItem.findFirst({
      where: { id: contentId, authorId: userId, status: 'draft', deletedAt: null },
    });
    if (!content) {
      throw new AppError(404, 'CONTENT_NOT_FOUND', 'Draft content not found');
    }

    await prisma.$transaction([
      prisma.contentItem.update({
        where: { id: contentId },
        data: { status: 'published', publishedAt: new Date() },
      }),
      prisma.userProfile.update({
        where: { userId },
        data: { postsCount: { increment: 1 } },
      }),
    ]);

    return this.getContentById(contentId, userId);
  },

  async deleteContent(userId: string, contentId: string) {
    const content = await prisma.contentItem.findFirst({
      where: { id: contentId, authorId: userId, deletedAt: null },
    });
    if (!content) {
      throw new AppError(404, 'CONTENT_NOT_FOUND', 'Content not found');
    }

    await prisma.$transaction([
      prisma.contentItem.update({
        where: { id: contentId },
        data: { deletedAt: new Date(), status: 'removed' },
      }),
      ...(content.status === 'published'
        ? [prisma.userProfile.update({
            where: { userId },
            data: { postsCount: { decrement: 1 } },
          })]
        : []),
    ]);

    await createAuditLog({
      actorId: userId,
      action: 'content.deleted',
      targetType: 'content',
      targetId: contentId,
    });
  },

  // ─── Comments ───────────────────────────────────────────
  async createComment(userId: string, contentId: string, body: string, parentId?: string) {
    const content = await prisma.contentItem.findFirst({
      where: { id: contentId, status: 'published', deletedAt: null },
    });
    if (!content) {
      throw new AppError(404, 'CONTENT_NOT_FOUND', 'Content not found');
    }
    if (!content.allowComments) {
      throw new AppError(403, 'COMMENTS_DISABLED', 'Comments are disabled for this content');
    }

    if (parentId) {
      const parent = await prisma.comment.findFirst({
        where: { id: parentId, contentId, status: 'active' },
      });
      if (!parent) {
        throw new AppError(404, 'PARENT_NOT_FOUND', 'Parent comment not found');
      }
    }

    const comment = await prisma.$transaction(async (tx) => {
      const c = await tx.comment.create({
        data: { contentId, authorId: userId, parentId, body },
        include: { author: { include: { profile: true } } },
      });

      await tx.contentItem.update({
        where: { id: contentId },
        data: { commentsCount: { increment: 1 } },
      });

      return c;
    });

    return {
      id: comment.id,
      contentId: comment.contentId,
      authorId: comment.authorId,
      author: {
        id: comment.author.id,
        username: comment.author.username,
        displayName: comment.author.displayName,
        avatarUrl: comment.author.profile?.avatarMediaId,
      },
      parentId: comment.parentId,
      body: comment.body,
      likesCount: comment.likesCount,
      createdAt: comment.createdAt.toISOString(),
    };
  },

  async getComments(contentId: string, cursor?: string, limit = 20) {
    const where: Record<string, unknown> = {
      contentId,
      status: 'active',
      parentId: null, // Top-level comments only
    };
    if (cursor) {
      where.createdAt = { lt: new Date(cursor) };
    }

    const comments = await prisma.comment.findMany({
      where,
      include: {
        author: { include: { profile: true } },
        replies: {
          where: { status: 'active' },
          include: { author: { include: { profile: true } } },
          take: 3,
          orderBy: { createdAt: 'asc' },
        },
      },
      take: limit + 1,
      orderBy: { createdAt: 'desc' },
    });

    const hasMore = comments.length > limit;
    if (hasMore) comments.pop();

    return {
      items: comments.map((c) => ({
        id: c.id,
        contentId: c.contentId,
        authorId: c.authorId,
        author: {
          id: c.author.id,
          username: c.author.username,
          displayName: c.author.displayName,
          avatarUrl: c.author.profile?.avatarMediaId,
        },
        body: c.body,
        likesCount: c.likesCount,
        createdAt: c.createdAt.toISOString(),
        replies: c.replies.map((r) => ({
          id: r.id,
          authorId: r.authorId,
          author: {
            id: r.author.id,
            username: r.author.username,
            displayName: r.author.displayName,
            avatarUrl: r.author.profile?.avatarMediaId,
          },
          body: r.body,
          likesCount: r.likesCount,
          createdAt: r.createdAt.toISOString(),
        })),
      })),
      cursor: comments.length > 0
        ? comments[comments.length - 1].createdAt.toISOString()
        : undefined,
      hasMore,
    };
  },

  async deleteComment(userId: string, commentId: string) {
    const comment = await prisma.comment.findFirst({
      where: { id: commentId, status: 'active' },
      include: { content: true },
    });
    if (!comment) {
      throw new AppError(404, 'COMMENT_NOT_FOUND', 'Comment not found');
    }
    if (comment.authorId !== userId && comment.content.authorId !== userId) {
      throw new AppError(403, 'FORBIDDEN', 'Cannot delete this comment');
    }

    await prisma.$transaction([
      prisma.comment.update({
        where: { id: commentId },
        data: { status: 'removed', deletedAt: new Date() },
      }),
      prisma.contentItem.update({
        where: { id: comment.contentId },
        data: { commentsCount: { decrement: 1 } },
      }),
    ]);
  },

  // ─── Reactions ──────────────────────────────────────────
  async toggleReaction(userId: string, targetType: string, targetId: string, reactionType = 'like') {
    const existing = await prisma.reaction.findUnique({
      where: { actorId_targetType_targetId: { actorId: userId, targetType, targetId } },
    });

    if (existing) {
      await prisma.reaction.delete({ where: { id: existing.id } });
      if (targetType === 'content') {
        await prisma.contentItem.update({
          where: { id: targetId },
          data: { likesCount: { decrement: 1 } },
        });
      } else if (targetType === 'comment') {
        await prisma.comment.update({
          where: { id: targetId },
          data: { likesCount: { decrement: 1 } },
        });
      }
      return { action: 'removed' };
    } else {
      await prisma.reaction.create({
        data: { actorId: userId, targetType, targetId, reactionType },
      });
      if (targetType === 'content') {
        await prisma.contentItem.update({
          where: { id: targetId },
          data: { likesCount: { increment: 1 } },
        });
      } else if (targetType === 'comment') {
        await prisma.comment.update({
          where: { id: targetId },
          data: { likesCount: { increment: 1 } },
        });
      }
      return { action: 'added' };
    }
  },

  // ─── Save Items ─────────────────────────────────────────
  async toggleSave(userId: string, contentId: string) {
    const existing = await prisma.saveItem.findUnique({
      where: { userId_contentId: { userId, contentId } },
    });

    if (existing) {
      await prisma.saveItem.delete({ where: { id: existing.id } });
      await prisma.contentItem.update({
        where: { id: contentId },
        data: { savesCount: { decrement: 1 } },
      });
      return { action: 'unsaved' };
    } else {
      await prisma.saveItem.create({ data: { userId, contentId } });
      await prisma.contentItem.update({
        where: { id: contentId },
        data: { savesCount: { increment: 1 } },
      });
      return { action: 'saved' };
    }
  },

  async getSavedContent(userId: string, cursor?: string, limit = 20) {
    const where: Record<string, unknown> = { userId };
    if (cursor) {
      where.createdAt = { lt: new Date(cursor) };
    }

    const saves = await prisma.saveItem.findMany({
      where,
      include: {
        content: {
          include: {
            author: { include: { profile: true } },
            contentMedia: { include: { media: true }, orderBy: { position: 'asc' } },
          },
        },
      },
      take: limit + 1,
      orderBy: { createdAt: 'desc' },
    });

    const hasMore = saves.length > limit;
    if (hasMore) saves.pop();

    return {
      items: saves.filter((s) => !s.content.deletedAt).map((s) => s.content),
      cursor: saves.length > 0 ? saves[saves.length - 1].createdAt.toISOString() : undefined,
      hasMore,
    };
  },

  // ─── Stories ────────────────────────────────────────────
  async createStory(userId: string, data: {
    mediaId: string;
    audience: string;
    interactiveType?: string;
    interactiveData?: Record<string, unknown>;
    saveToArchive?: boolean;
  }) {
    const asset = await prisma.mediaAsset.findFirst({
      where: { id: data.mediaId, ownerId: userId },
    });
    if (!asset) {
      throw new AppError(404, 'ASSET_NOT_FOUND', 'Media asset not found');
    }

    return prisma.storyItem.create({
      data: {
        authorId: userId,
        mediaId: data.mediaId,
        audience: data.audience,
        interactiveType: data.interactiveType,
        interactiveData: data.interactiveData as Prisma.InputJsonValue | undefined,
        saveToArchive: data.saveToArchive ?? false,
        expiresAt: new Date(Date.now() + APP_CONFIG.STORY_TTL_HOURS * 60 * 60 * 1000),
      },
    });
  },

  async getStories(viewerId: string) {
    // Get stories from followed users
    const follows = await prisma.followEdge.findMany({
      where: { followerId: viewerId, status: 'active' },
      select: { followeeId: true },
    });

    const followedIds = [...follows.map((f) => f.followeeId), viewerId];

    const stories = await prisma.storyItem.findMany({
      where: {
        authorId: { in: followedIds },
        expiresAt: { gt: new Date() },
        deletedAt: null,
      },
      include: { author: { include: { profile: true } } },
      orderBy: { createdAt: 'desc' },
    });

    // Group by author
    const grouped = new Map<string, typeof stories>();
    for (const story of stories) {
      const existing = grouped.get(story.authorId) || [];
      existing.push(story);
      grouped.set(story.authorId, existing);
    }

    return Array.from(grouped.entries()).map(([authorId, authorStories]) => ({
      author: {
        id: authorStories[0].author.id,
        username: authorStories[0].author.username,
        displayName: authorStories[0].author.displayName,
        avatarUrl: authorStories[0].author.profile?.avatarMediaId,
      },
      stories: authorStories.map((s) => ({
        id: s.id,
        mediaId: s.mediaId,
        audience: s.audience,
        viewsCount: s.viewsCount,
        expiresAt: s.expiresAt.toISOString(),
        createdAt: s.createdAt.toISOString(),
      })),
    }));
  },

  // ─── Highlights ─────────────────────────────────────────
  async createHighlight(userId: string, data: { title: string; storyIds: string[]; coverMediaId?: string }) {
    return prisma.highlightCollection.create({
      data: {
        ownerId: userId,
        title: data.title,
        storyIds: JSON.stringify(data.storyIds),
        coverMediaId: data.coverMediaId,
      },
    });
  },

  async getHighlights(userId: string) {
    return prisma.highlightCollection.findMany({
      where: { ownerId: userId },
      orderBy: { createdAt: 'desc' },
    });
  },
};

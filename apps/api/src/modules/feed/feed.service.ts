import { prisma } from '../../main';
import { APP_CONFIG } from '@vybe/config';

export const feedService = {
  async getMyVybeFeed(userId: string, cursor?: string, limit = 20) {
    // Get followed user IDs
    const follows = await prisma.followEdge.findMany({
      where: { followerId: userId, status: 'active' },
      select: { followeeId: true },
    });
    const followedIds = follows.map((f) => f.followeeId);

    if (followedIds.length === 0) {
      return { items: [], cursor: undefined, hasMore: false };
    }

    // Get blocked/muted users
    const blocks = await prisma.blockEdge.findMany({
      where: { blockerId: userId },
      select: { blockedId: true },
    });
    const mutes = await prisma.muteEdge.findMany({
      where: { ownerId: userId, targetType: 'user' },
      select: { targetId: true },
    });
    const excludeIds = new Set([
      ...blocks.map((b) => b.blockedId),
      ...mutes.map((m) => m.targetId),
    ]);

    const visibleAuthors = followedIds.filter((id) => !excludeIds.has(id));

    const where: Record<string, unknown> = {
      authorId: { in: visibleAuthors },
      status: 'published',
      deletedAt: null,
      audience: { in: ['public', 'followers'] },
    };

    if (cursor) {
      where.publishedAt = { lt: new Date(cursor) };
    }

    const items = await prisma.contentItem.findMany({
      where,
      include: {
        author: { include: { profile: true } },
        contentMedia: {
          include: { media: true },
          orderBy: { position: 'asc' },
        },
      },
      take: limit + 1,
      orderBy: { publishedAt: 'desc' },
    });

    const hasMore = items.length > limit;
    if (hasMore) items.pop();

    // Check viewer reactions/saves
    const contentIds = items.map((i) => i.id);
    const reactions = await prisma.reaction.findMany({
      where: { actorId: userId, targetType: 'content', targetId: { in: contentIds } },
    });
    const saves = await prisma.saveItem.findMany({
      where: { userId, contentId: { in: contentIds } },
    });
    const likedSet = new Set(reactions.map((r) => r.targetId));
    const savedSet = new Set(saves.map((s) => s.contentId));

    return {
      items: items.map((item) => ({
        content: {
          id: item.id,
          authorId: item.authorId,
          author: {
            id: item.author.id,
            username: item.author.username,
            displayName: item.author.displayName,
            avatarUrl: item.author.profile?.avatarMediaId,
          },
          type: item.type,
          status: item.status,
          audience: item.audience,
          caption: item.caption,
          hashtags: item.hashtags,
          media: item.contentMedia.map((cm) => ({
            id: cm.media.id,
            kind: cm.media.kind,
            mimeType: cm.media.mimeType,
            url: `${process.env.S3_ENDPOINT || 'http://localhost:9000'}/${process.env.S3_BUCKET || 'vybe-media'}/${cm.media.storageKey}`,
            status: cm.media.status,
          })),
          likesCount: item.likesCount,
          commentsCount: item.commentsCount,
          sharesCount: item.sharesCount,
          savesCount: item.savesCount,
          isLiked: likedSet.has(item.id),
          isSaved: savedSet.has(item.id),
          publishedAt: item.publishedAt?.toISOString(),
          createdAt: item.createdAt.toISOString(),
        },
      })),
      cursor: items.length > 0 ? items[items.length - 1].publishedAt?.toISOString() : undefined,
      hasMore,
    };
  },

  async getForYouFeed(userId: string, cursor?: string, limit = 20) {
    // Simple for-you: mix of public content from non-blocked users, sorted by recency + engagement
    const blocks = await prisma.blockEdge.findMany({
      where: { OR: [{ blockerId: userId }, { blockedId: userId }] },
    });
    const blockedIds = blocks.map((b) => b.blockerId === userId ? b.blockedId : b.blockerId);

    const where: Record<string, unknown> = {
      status: 'published',
      deletedAt: null,
      audience: 'public',
      authorId: { notIn: [...blockedIds, userId] },
    };

    if (cursor) {
      where.publishedAt = { lt: new Date(cursor) };
    }

    const items = await prisma.contentItem.findMany({
      where,
      include: {
        author: { include: { profile: true } },
        contentMedia: {
          include: { media: true },
          orderBy: { position: 'asc' },
        },
      },
      take: limit + 1,
      orderBy: [
        { likesCount: 'desc' },
        { publishedAt: 'desc' },
      ],
    });

    const hasMore = items.length > limit;
    if (hasMore) items.pop();

    return {
      items: items.map((item) => ({
        content: {
          id: item.id,
          authorId: item.authorId,
          author: {
            id: item.author.id,
            username: item.author.username,
            displayName: item.author.displayName,
            avatarUrl: item.author.profile?.avatarMediaId,
          },
          type: item.type,
          caption: item.caption,
          hashtags: item.hashtags,
          media: item.contentMedia.map((cm) => ({
            id: cm.media.id,
            kind: cm.media.kind,
            mimeType: cm.media.mimeType,
            url: `${process.env.S3_ENDPOINT || 'http://localhost:9000'}/${process.env.S3_BUCKET || 'vybe-media'}/${cm.media.storageKey}`,
            status: cm.media.status,
          })),
          likesCount: item.likesCount,
          commentsCount: item.commentsCount,
          sharesCount: item.sharesCount,
          publishedAt: item.publishedAt?.toISOString(),
          createdAt: item.createdAt.toISOString(),
        },
      })),
      cursor: items.length > 0 ? items[items.length - 1].publishedAt?.toISOString() : undefined,
      hasMore,
    };
  },

  async getProfileFeed(profileUserId: string, viewerId?: string, cursor?: string, limit = 20) {
    // Check block
    if (viewerId) {
      const blocked = await prisma.blockEdge.findFirst({
        where: {
          OR: [
            { blockerId: profileUserId, blockedId: viewerId },
            { blockerId: viewerId, blockedId: profileUserId },
          ],
        },
      });
      if (blocked) {
        return { items: [], cursor: undefined, hasMore: false };
      }
    }

    const audienceFilter: string[] = ['public'];
    if (viewerId) {
      const isFollowing = await prisma.followEdge.findFirst({
        where: { followerId: viewerId, followeeId: profileUserId, status: 'active' },
      });
      if (isFollowing) audienceFilter.push('followers');

      const [userAId, userBId] = [viewerId, profileUserId].sort();
      const isFriend = await prisma.friendEdge.findFirst({
        where: { userAId, userBId, status: 'active' },
      });
      if (isFriend) audienceFilter.push('friends');
    }

    if (viewerId === profileUserId) {
      // User can see all their own content
      audienceFilter.push('followers', 'friends', 'circle', 'private');
    }

    const where: Record<string, unknown> = {
      authorId: profileUserId,
      status: 'published',
      deletedAt: null,
      audience: { in: [...new Set(audienceFilter)] },
    };

    if (cursor) {
      where.publishedAt = { lt: new Date(cursor) };
    }

    const items = await prisma.contentItem.findMany({
      where,
      include: {
        author: { include: { profile: true } },
        contentMedia: {
          include: { media: true },
          orderBy: { position: 'asc' },
        },
      },
      take: limit + 1,
      orderBy: { publishedAt: 'desc' },
    });

    const hasMore = items.length > limit;
    if (hasMore) items.pop();

    return {
      items: items.map((item) => ({
        content: {
          id: item.id,
          authorId: item.authorId,
          author: {
            id: item.author.id,
            username: item.author.username,
            displayName: item.author.displayName,
            avatarUrl: item.author.profile?.avatarMediaId,
          },
          type: item.type,
          caption: item.caption,
          hashtags: item.hashtags,
          media: item.contentMedia.map((cm) => ({
            id: cm.media.id,
            kind: cm.media.kind,
            mimeType: cm.media.mimeType,
            url: `${process.env.S3_ENDPOINT || 'http://localhost:9000'}/${process.env.S3_BUCKET || 'vybe-media'}/${cm.media.storageKey}`,
            status: cm.media.status,
          })),
          likesCount: item.likesCount,
          commentsCount: item.commentsCount,
          publishedAt: item.publishedAt?.toISOString(),
          createdAt: item.createdAt.toISOString(),
        },
      })),
      cursor: items.length > 0 ? items[items.length - 1].publishedAt?.toISOString() : undefined,
      hasMore,
    };
  },

  async getTrending(timeframe = '24h', cursor?: string, limit = 20) {
    const timeMap: Record<string, number> = {
      '1h': 60 * 60 * 1000,
      '24h': 24 * 60 * 60 * 1000,
      '7d': 7 * 24 * 60 * 60 * 1000,
    };
    const since = new Date(Date.now() - (timeMap[timeframe] || timeMap['24h']));

    const where: Record<string, unknown> = {
      status: 'published',
      deletedAt: null,
      audience: 'public',
      publishedAt: { gte: since },
    };

    if (cursor) {
      where.id = { lt: cursor };
    }

    const items = await prisma.contentItem.findMany({
      where,
      include: {
        author: { include: { profile: true } },
        contentMedia: {
          include: { media: true },
          orderBy: { position: 'asc' },
        },
      },
      take: limit + 1,
      orderBy: [
        { likesCount: 'desc' },
        { commentsCount: 'desc' },
        { publishedAt: 'desc' },
      ],
    });

    const hasMore = items.length > limit;
    if (hasMore) items.pop();

    return {
      items: items.map((item) => ({
        content: {
          id: item.id,
          authorId: item.authorId,
          author: {
            id: item.author.id,
            username: item.author.username,
            displayName: item.author.displayName,
            avatarUrl: item.author.profile?.avatarMediaId,
          },
          type: item.type,
          caption: item.caption,
          likesCount: item.likesCount,
          commentsCount: item.commentsCount,
          publishedAt: item.publishedAt?.toISOString(),
        },
      })),
      cursor: items.length > 0 ? items[items.length - 1].id : undefined,
      hasMore,
    };
  },
};

import { prisma } from '../../main';
import { AppError } from '../../middleware/error-handler';
import { APP_CONFIG } from '@vybe/config';

export const messagingService = {
  async createConversation(userId: string, data: {
    type: string;
    participantIds?: string[];
    circleId?: string;
    title?: string;
  }) {
    if (data.type === 'direct') {
      if (!data.participantIds || data.participantIds.length !== 1) {
        throw new AppError(400, 'INVALID_PARTICIPANTS', 'Direct conversations require exactly one other participant');
      }

      const otherUserId = data.participantIds[0];
      if (otherUserId === userId) {
        throw new AppError(400, 'SELF_CONVERSATION', 'Cannot create conversation with yourself');
      }

      // Check if blocked
      const blocked = await prisma.blockEdge.findFirst({
        where: {
          OR: [
            { blockerId: userId, blockedId: otherUserId },
            { blockerId: otherUserId, blockedId: userId },
          ],
        },
      });
      if (blocked) {
        throw new AppError(403, 'BLOCKED', 'Cannot message this user');
      }

      // Check DM permissions
      const otherUser = await prisma.user.findUnique({
        where: { id: otherUserId },
        include: { settings: true },
      });
      if (!otherUser || otherUser.deletedAt) {
        throw new AppError(404, 'USER_NOT_FOUND', 'User not found');
      }

      const dmPerm = otherUser.settings?.dmPermission || 'followers';
      if (dmPerm === 'nobody') {
        throw new AppError(403, 'DM_DISABLED', 'This user does not accept direct messages');
      }
      if (dmPerm === 'followers') {
        const isFollowing = await prisma.followEdge.findFirst({
          where: { followerId: otherUserId, followeeId: userId, status: 'active' },
        });
        if (!isFollowing) {
          throw new AppError(403, 'DM_RESTRICTED', 'This user only accepts messages from people they follow');
        }
      }
      if (dmPerm === 'friends') {
        const [userAId, userBId] = [userId, otherUserId].sort();
        const isFriend = await prisma.friendEdge.findFirst({
          where: { userAId, userBId, status: 'active' },
        });
        if (!isFriend) {
          throw new AppError(403, 'DM_RESTRICTED', 'This user only accepts messages from friends');
        }
      }

      // Check if conversation already exists
      const existingMembers = await prisma.conversationMember.findMany({
        where: { userId: { in: [userId, otherUserId] } },
        select: { conversationId: true },
      });

      const conversationCounts = new Map<string, number>();
      for (const m of existingMembers) {
        conversationCounts.set(m.conversationId, (conversationCounts.get(m.conversationId) || 0) + 1);
      }

      for (const [convId, count] of conversationCounts) {
        if (count === 2) {
          const conv = await prisma.conversation.findFirst({
            where: { id: convId, type: 'direct' },
          });
          if (conv) {
            return this.getConversation(conv.id, userId);
          }
        }
      }
    }

    const conversation = await prisma.$transaction(async (tx) => {
      const conv = await tx.conversation.create({
        data: {
          type: data.type,
          title: data.title,
          createdBy: userId,
          circleId: data.circleId,
        },
      });

      // Add creator as member
      await tx.conversationMember.create({
        data: { conversationId: conv.id, userId, role: 'owner' },
      });

      // Add other participants
      if (data.participantIds) {
        for (const pid of data.participantIds) {
          await tx.conversationMember.create({
            data: { conversationId: conv.id, userId: pid },
          });
        }
      }

      return conv;
    });

    return this.getConversation(conversation.id, userId);
  },

  async getConversation(conversationId: string, userId: string) {
    const conv = await prisma.conversation.findUnique({
      where: { id: conversationId },
      include: {
        members: {
          include: { user: { include: { profile: true } } },
        },
      },
    });

    if (!conv) {
      throw new AppError(404, 'CONVERSATION_NOT_FOUND', 'Conversation not found');
    }

    const isMember = conv.members.some((m) => m.userId === userId && !m.leftAt);
    if (!isMember) {
      throw new AppError(403, 'NOT_A_MEMBER', 'You are not a member of this conversation');
    }

    // Get last message
    const lastMessage = await prisma.message.findFirst({
      where: { conversationId, deletedForAll: false },
      include: { sender: { include: { profile: true } } },
      orderBy: { createdAt: 'desc' },
    });

    // Get unread count
    const member = conv.members.find((m) => m.userId === userId);
    let unreadCount = 0;
    if (member?.lastReadMessageId) {
      const lastRead = await prisma.message.findUnique({
        where: { id: member.lastReadMessageId },
      });
      if (lastRead) {
        unreadCount = await prisma.message.count({
          where: {
            conversationId,
            createdAt: { gt: lastRead.createdAt },
            senderId: { not: userId },
            deletedForAll: false,
          },
        });
      }
    } else {
      unreadCount = await prisma.message.count({
        where: {
          conversationId,
          senderId: { not: userId },
          deletedForAll: false,
        },
      });
    }

    return {
      id: conv.id,
      type: conv.type,
      title: conv.title,
      participants: conv.members
        .filter((m) => !m.leftAt)
        .map((m) => ({
          id: m.user.id,
          username: m.user.username,
          displayName: m.user.displayName,
          avatarUrl: m.user.profile?.avatarMediaId,
        })),
      lastMessage: lastMessage
        ? {
            id: lastMessage.id,
            senderId: lastMessage.senderId,
            sender: {
              id: lastMessage.sender.id,
              username: lastMessage.sender.username,
              displayName: lastMessage.sender.displayName,
            },
            kind: lastMessage.kind,
            body: lastMessage.body,
            createdAt: lastMessage.createdAt.toISOString(),
          }
        : undefined,
      unreadCount,
      createdAt: conv.createdAt.toISOString(),
    };
  },

  async getUserConversations(userId: string, cursor?: string, limit = 20) {
    const memberships = await prisma.conversationMember.findMany({
      where: { userId, leftAt: null },
      select: { conversationId: true },
    });

    const conversationIds = memberships.map((m) => m.conversationId);
    if (conversationIds.length === 0) {
      return { items: [], cursor: undefined, hasMore: false };
    }

    const where: Record<string, unknown> = {
      id: { in: conversationIds },
    };
    if (cursor) {
      where.lastMessageAt = { lt: new Date(cursor) };
    }

    const conversations = await prisma.conversation.findMany({
      where,
      orderBy: { lastMessageAt: { sort: 'desc', nulls: 'last' } },
      take: limit + 1,
    });

    const hasMore = conversations.length > limit;
    if (hasMore) conversations.pop();

    const results = await Promise.all(
      conversations.map((c) => this.getConversation(c.id, userId))
    );

    return {
      items: results,
      cursor: conversations.length > 0
        ? conversations[conversations.length - 1].lastMessageAt?.toISOString()
        : undefined,
      hasMore,
    };
  },

  async sendMessage(userId: string, conversationId: string, data: {
    kind: string;
    body?: string;
    mediaIds?: string[];
    replyToId?: string;
  }) {
    // Verify membership
    const member = await prisma.conversationMember.findFirst({
      where: { conversationId, userId, leftAt: null },
    });
    if (!member) {
      throw new AppError(403, 'NOT_A_MEMBER', 'You are not a member of this conversation');
    }

    const message = await prisma.$transaction(async (tx) => {
      const msg = await tx.message.create({
        data: {
          conversationId,
          senderId: userId,
          kind: data.kind,
          body: data.body,
          replyToId: data.replyToId,
        },
        include: {
          sender: { include: { profile: true } },
          replyTo: { include: { sender: true } },
        },
      });

      // Link media
      if (data.mediaIds) {
        for (const mediaId of data.mediaIds) {
          await tx.messageMedia.create({
            data: { messageId: msg.id, mediaId },
          });
        }
      }

      // Update conversation lastMessageAt
      await tx.conversation.update({
        where: { id: conversationId },
        data: { lastMessageAt: new Date() },
      });

      // Update sender's last read
      await tx.conversationMember.update({
        where: { id: member.id },
        data: { lastReadMessageId: msg.id },
      });

      return msg;
    });

    return {
      id: message.id,
      conversationId: message.conversationId,
      senderId: message.senderId,
      sender: {
        id: message.sender.id,
        username: message.sender.username,
        displayName: message.sender.displayName,
        avatarUrl: message.sender.profile?.avatarMediaId,
      },
      kind: message.kind,
      body: message.body,
      replyTo: message.replyTo
        ? {
            id: message.replyTo.id,
            senderId: message.replyTo.senderId,
            body: message.replyTo.body,
          }
        : undefined,
      status: message.status,
      createdAt: message.createdAt.toISOString(),
    };
  },

  async getMessages(userId: string, conversationId: string, cursor?: string, limit = 30) {
    const member = await prisma.conversationMember.findFirst({
      where: { conversationId, userId, leftAt: null },
    });
    if (!member) {
      throw new AppError(403, 'NOT_A_MEMBER', 'You are not a member of this conversation');
    }

    const where: Record<string, unknown> = {
      conversationId,
      deletedForAll: false,
    };
    if (cursor) {
      where.createdAt = { lt: new Date(cursor) };
    }

    const messages = await prisma.message.findMany({
      where,
      include: {
        sender: { include: { profile: true } },
        replyTo: { include: { sender: true } },
        media: { include: { media: true } },
        reactions: true,
      },
      take: limit + 1,
      orderBy: { createdAt: 'desc' },
    });

    const hasMore = messages.length > limit;
    if (hasMore) messages.pop();

    return {
      items: messages.map((msg) => ({
        id: msg.id,
        conversationId: msg.conversationId,
        senderId: msg.senderId,
        sender: {
          id: msg.sender.id,
          username: msg.sender.username,
          displayName: msg.sender.displayName,
          avatarUrl: msg.sender.profile?.avatarMediaId,
        },
        kind: msg.kind,
        body: msg.body,
        media: msg.media.map((m) => ({
          id: m.media.id,
          kind: m.media.kind,
          url: `${process.env.S3_ENDPOINT || 'http://localhost:9000'}/${process.env.S3_BUCKET || 'vybe-media'}/${m.media.storageKey}`,
        })),
        replyTo: msg.replyTo
          ? { id: msg.replyTo.id, senderId: msg.replyTo.senderId, body: msg.replyTo.body }
          : undefined,
        reactions: msg.reactions.map((r) => ({
          actorId: r.actorId,
          emoji: r.emoji,
          createdAt: r.createdAt.toISOString(),
        })),
        status: msg.status,
        createdAt: msg.createdAt.toISOString(),
      })),
      cursor: messages.length > 0
        ? messages[messages.length - 1].createdAt.toISOString()
        : undefined,
      hasMore,
    };
  },

  async deleteMessage(userId: string, messageId: string, forAll = false) {
    const message = await prisma.message.findUnique({
      where: { id: messageId },
    });
    if (!message) {
      throw new AppError(404, 'MESSAGE_NOT_FOUND', 'Message not found');
    }
    if (message.senderId !== userId) {
      throw new AppError(403, 'FORBIDDEN', 'Cannot delete this message');
    }

    if (forAll) {
      const ageMinutes = (Date.now() - message.createdAt.getTime()) / (1000 * 60);
      if (ageMinutes > APP_CONFIG.DELETE_FOR_ALL_WINDOW_MINUTES) {
        throw new AppError(400, 'DELETE_WINDOW_EXPIRED', 'Delete for all window has expired');
      }
      await prisma.message.update({
        where: { id: messageId },
        data: { deletedForAll: true, status: 'deleted_for_all' },
      });
    } else {
      await prisma.message.update({
        where: { id: messageId },
        data: { status: 'deleted_for_sender' },
      });
    }
  },

  async reactToMessage(userId: string, messageId: string, emoji: string) {
    const message = await prisma.message.findUnique({ where: { id: messageId } });
    if (!message) {
      throw new AppError(404, 'MESSAGE_NOT_FOUND', 'Message not found');
    }

    // Verify membership
    const member = await prisma.conversationMember.findFirst({
      where: { conversationId: message.conversationId, userId, leftAt: null },
    });
    if (!member) {
      throw new AppError(403, 'NOT_A_MEMBER', 'You are not a member of this conversation');
    }

    const existing = await prisma.messageReaction.findUnique({
      where: { messageId_actorId: { messageId, actorId: userId } },
    });

    if (existing) {
      if (existing.emoji === emoji) {
        await prisma.messageReaction.delete({ where: { id: existing.id } });
        return { action: 'removed' };
      } else {
        await prisma.messageReaction.update({
          where: { id: existing.id },
          data: { emoji },
        });
        return { action: 'updated' };
      }
    }

    await prisma.messageReaction.create({
      data: { messageId, actorId: userId, emoji },
    });
    return { action: 'added' };
  },

  async markRead(userId: string, conversationId: string, messageId: string) {
    const member = await prisma.conversationMember.findFirst({
      where: { conversationId, userId, leftAt: null },
    });
    if (!member) {
      throw new AppError(403, 'NOT_A_MEMBER', 'You are not a member of this conversation');
    }

    await prisma.conversationMember.update({
      where: { id: member.id },
      data: { lastReadMessageId: messageId },
    });
  },
};

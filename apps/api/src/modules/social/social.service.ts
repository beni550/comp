import { prisma } from '../../main';
import { AppError } from '../../middleware/error-handler';
import { createAuditLog } from '../../common/audit';
import { APP_CONFIG } from '@vybe/config';

export const socialService = {
  // ─── Follow ─────────────────────────────────────────────
  async follow(followerId: string, followeeId: string) {
    if (followerId === followeeId) {
      throw new AppError(400, 'SELF_FOLLOW', 'Cannot follow yourself');
    }

    const followee = await prisma.user.findUnique({ where: { id: followeeId } });
    if (!followee || followee.deletedAt) {
      throw new AppError(404, 'USER_NOT_FOUND', 'User not found');
    }

    // Check if blocked
    const blocked = await prisma.blockEdge.findFirst({
      where: {
        OR: [
          { blockerId: followeeId, blockedId: followerId },
          { blockerId: followerId, blockedId: followeeId },
        ],
      },
    });
    if (blocked) {
      throw new AppError(403, 'BLOCKED', 'Cannot follow this user');
    }

    // Check existing follow
    const existing = await prisma.followEdge.findUnique({
      where: { followerId_followeeId: { followerId, followeeId } },
    });
    if (existing) {
      throw new AppError(409, 'ALREADY_FOLLOWING', 'Already following this user');
    }

    // If private account, create pending follow request
    const status = followee.privacyMode === 'public' ? 'active' : 'pending';

    const edge = await prisma.followEdge.create({
      data: { followerId, followeeId, status },
    });

    if (status === 'active') {
      // Update counters
      await prisma.$transaction([
        prisma.userProfile.update({
          where: { userId: followerId },
          data: { followingCount: { increment: 1 } },
        }),
        prisma.userProfile.update({
          where: { userId: followeeId },
          data: { followersCount: { increment: 1 } },
        }),
      ]);
    }

    return edge;
  },

  async unfollow(followerId: string, followeeId: string) {
    const edge = await prisma.followEdge.findUnique({
      where: { followerId_followeeId: { followerId, followeeId } },
    });
    if (!edge) {
      throw new AppError(404, 'NOT_FOLLOWING', 'Not following this user');
    }

    await prisma.followEdge.delete({
      where: { id: edge.id },
    });

    if (edge.status === 'active') {
      await prisma.$transaction([
        prisma.userProfile.update({
          where: { userId: followerId },
          data: { followingCount: { decrement: 1 } },
        }),
        prisma.userProfile.update({
          where: { userId: followeeId },
          data: { followersCount: { decrement: 1 } },
        }),
      ]);
    }
  },

  async acceptFollowRequest(userId: string, followEdgeId: string) {
    const edge = await prisma.followEdge.findFirst({
      where: { id: followEdgeId, followeeId: userId, status: 'pending' },
    });
    if (!edge) {
      throw new AppError(404, 'REQUEST_NOT_FOUND', 'Follow request not found');
    }

    await prisma.followEdge.update({
      where: { id: edge.id },
      data: { status: 'active' },
    });

    await prisma.$transaction([
      prisma.userProfile.update({
        where: { userId: edge.followerId },
        data: { followingCount: { increment: 1 } },
      }),
      prisma.userProfile.update({
        where: { userId: userId },
        data: { followersCount: { increment: 1 } },
      }),
    ]);
  },

  async declineFollowRequest(userId: string, followEdgeId: string) {
    const edge = await prisma.followEdge.findFirst({
      where: { id: followEdgeId, followeeId: userId, status: 'pending' },
    });
    if (!edge) {
      throw new AppError(404, 'REQUEST_NOT_FOUND', 'Follow request not found');
    }

    await prisma.followEdge.delete({ where: { id: edge.id } });
  },

  async getFollowers(userId: string, cursor?: string, limit = 20) {
    const where: Record<string, unknown> = { followeeId: userId, status: 'active' };
    if (cursor) {
      where.id = { gt: cursor };
    }

    const edges = await prisma.followEdge.findMany({
      where,
      include: { follower: { include: { profile: true } } },
      take: limit + 1,
      orderBy: { createdAt: 'desc' },
    });

    const hasMore = edges.length > limit;
    if (hasMore) edges.pop();

    return {
      items: edges.map((e) => ({
        id: e.follower.id,
        username: e.follower.username,
        displayName: e.follower.displayName,
        avatarUrl: e.follower.profile?.avatarMediaId,
        followedAt: e.createdAt.toISOString(),
      })),
      cursor: edges.length > 0 ? edges[edges.length - 1].id : undefined,
      hasMore,
    };
  },

  async getFollowing(userId: string, cursor?: string, limit = 20) {
    const where: Record<string, unknown> = { followerId: userId, status: 'active' };
    if (cursor) {
      where.id = { gt: cursor };
    }

    const edges = await prisma.followEdge.findMany({
      where,
      include: { followee: { include: { profile: true } } },
      take: limit + 1,
      orderBy: { createdAt: 'desc' },
    });

    const hasMore = edges.length > limit;
    if (hasMore) edges.pop();

    return {
      items: edges.map((e) => ({
        id: e.followee.id,
        username: e.followee.username,
        displayName: e.followee.displayName,
        avatarUrl: e.followee.profile?.avatarMediaId,
        followedAt: e.createdAt.toISOString(),
      })),
      cursor: edges.length > 0 ? edges[edges.length - 1].id : undefined,
      hasMore,
    };
  },

  async getFollowRequests(userId: string) {
    const edges = await prisma.followEdge.findMany({
      where: { followeeId: userId, status: 'pending' },
      include: { follower: { include: { profile: true } } },
      orderBy: { createdAt: 'desc' },
    });

    return edges.map((e) => ({
      id: e.id,
      follower: {
        id: e.follower.id,
        username: e.follower.username,
        displayName: e.follower.displayName,
        avatarUrl: e.follower.profile?.avatarMediaId,
      },
      createdAt: e.createdAt.toISOString(),
    }));
  },

  // ─── Friends ────────────────────────────────────────────
  async sendFriendRequest(fromId: string, toId: string) {
    if (fromId === toId) {
      throw new AppError(400, 'SELF_FRIEND', 'Cannot send friend request to yourself');
    }

    const target = await prisma.user.findUnique({ where: { id: toId } });
    if (!target || target.deletedAt) {
      throw new AppError(404, 'USER_NOT_FOUND', 'User not found');
    }

    // Sort IDs to maintain unique constraint
    const [userAId, userBId] = [fromId, toId].sort();

    const existing = await prisma.friendEdge.findUnique({
      where: { userAId_userBId: { userAId, userBId } },
    });
    if (existing) {
      if (existing.status === 'active') {
        throw new AppError(409, 'ALREADY_FRIENDS', 'Already friends');
      }
      if (existing.status === 'pending') {
        throw new AppError(409, 'REQUEST_PENDING', 'Friend request already pending');
      }
    }

    return prisma.friendEdge.create({
      data: { userAId, userBId, status: 'pending', initiatedBy: fromId },
    });
  },

  async acceptFriendRequest(userId: string, friendEdgeId: string) {
    const edge = await prisma.friendEdge.findFirst({
      where: {
        id: friendEdgeId,
        status: 'pending',
        OR: [{ userAId: userId }, { userBId: userId }],
      },
    });
    if (!edge || edge.initiatedBy === userId) {
      throw new AppError(404, 'REQUEST_NOT_FOUND', 'Friend request not found');
    }

    await prisma.friendEdge.update({
      where: { id: edge.id },
      data: { status: 'active' },
    });

    // Update friend counts
    await prisma.$transaction([
      prisma.userProfile.update({
        where: { userId: edge.userAId },
        data: { friendsCount: { increment: 1 } },
      }),
      prisma.userProfile.update({
        where: { userId: edge.userBId },
        data: { friendsCount: { increment: 1 } },
      }),
    ]);
  },

  async declineFriendRequest(userId: string, friendEdgeId: string) {
    const edge = await prisma.friendEdge.findFirst({
      where: {
        id: friendEdgeId,
        status: 'pending',
        OR: [{ userAId: userId }, { userBId: userId }],
      },
    });
    if (!edge) {
      throw new AppError(404, 'REQUEST_NOT_FOUND', 'Friend request not found');
    }

    await prisma.friendEdge.update({
      where: { id: edge.id },
      data: { status: 'declined' },
    });
  },

  async removeFriend(userId: string, friendId: string) {
    const [userAId, userBId] = [userId, friendId].sort();
    const edge = await prisma.friendEdge.findUnique({
      where: { userAId_userBId: { userAId, userBId } },
    });
    if (!edge || edge.status !== 'active') {
      throw new AppError(404, 'NOT_FRIENDS', 'Not friends with this user');
    }

    await prisma.friendEdge.delete({ where: { id: edge.id } });

    await prisma.$transaction([
      prisma.userProfile.update({
        where: { userId: userAId },
        data: { friendsCount: { decrement: 1 } },
      }),
      prisma.userProfile.update({
        where: { userId: userBId },
        data: { friendsCount: { decrement: 1 } },
      }),
    ]);
  },

  async getFriends(userId: string, cursor?: string, limit = 20) {
    const edges = await prisma.friendEdge.findMany({
      where: {
        status: 'active',
        OR: [{ userAId: userId }, { userBId: userId }],
      },
      include: {
        userA: { include: { profile: true } },
        userB: { include: { profile: true } },
      },
      take: limit + 1,
      orderBy: { createdAt: 'desc' },
    });

    const hasMore = edges.length > limit;
    if (hasMore) edges.pop();

    return {
      items: edges.map((e) => {
        const friend = e.userAId === userId ? e.userB : e.userA;
        return {
          id: friend.id,
          username: friend.username,
          displayName: friend.displayName,
          avatarUrl: friend.profile?.avatarMediaId,
          since: e.createdAt.toISOString(),
        };
      }),
      cursor: edges.length > 0 ? edges[edges.length - 1].id : undefined,
      hasMore,
    };
  },

  // ─── Block & Mute ──────────────────────────────────────
  async blockUser(blockerId: string, blockedId: string) {
    if (blockerId === blockedId) {
      throw new AppError(400, 'SELF_BLOCK', 'Cannot block yourself');
    }

    const existing = await prisma.blockEdge.findUnique({
      where: { blockerId_blockedId: { blockerId, blockedId } },
    });
    if (existing) {
      throw new AppError(409, 'ALREADY_BLOCKED', 'User already blocked');
    }

    // Remove follow edges and friend edges
    await prisma.$transaction([
      prisma.followEdge.deleteMany({
        where: {
          OR: [
            { followerId: blockerId, followeeId: blockedId },
            { followerId: blockedId, followeeId: blockerId },
          ],
        },
      }),
      prisma.friendEdge.deleteMany({
        where: {
          OR: [
            { userAId: blockerId, userBId: blockedId },
            { userAId: blockedId, userBId: blockerId },
          ],
        },
      }),
    ]);

    return prisma.blockEdge.create({
      data: { blockerId, blockedId },
    });
  },

  async unblockUser(blockerId: string, blockedId: string) {
    const edge = await prisma.blockEdge.findUnique({
      where: { blockerId_blockedId: { blockerId, blockedId } },
    });
    if (!edge) {
      throw new AppError(404, 'NOT_BLOCKED', 'User is not blocked');
    }

    await prisma.blockEdge.delete({ where: { id: edge.id } });
  },

  async muteUser(ownerId: string, targetId: string, targetType = 'user') {
    const existing = await prisma.muteEdge.findUnique({
      where: { ownerId_targetId_targetType: { ownerId, targetId, targetType } },
    });
    if (existing) {
      throw new AppError(409, 'ALREADY_MUTED', 'Already muted');
    }

    return prisma.muteEdge.create({
      data: { ownerId, targetId, targetType },
    });
  },

  async unmuteUser(ownerId: string, targetId: string, targetType = 'user') {
    const edge = await prisma.muteEdge.findUnique({
      where: { ownerId_targetId_targetType: { ownerId, targetId, targetType } },
    });
    if (!edge) {
      throw new AppError(404, 'NOT_MUTED', 'Not muted');
    }

    await prisma.muteEdge.delete({ where: { id: edge.id } });
  },

  // ─── Circles ────────────────────────────────────────────
  async createCircle(ownerId: string, data: { name: string; description?: string; type?: string; privacyMode?: string }) {
    const circle = await prisma.circle.create({
      data: {
        ownerId,
        name: data.name,
        description: data.description || '',
        type: data.type || 'general',
        privacyMode: data.privacyMode || 'private',
        memberLimit: APP_CONFIG.MAX_CIRCLE_MEMBERS,
      },
    });

    // Add owner as member
    await prisma.circleMember.create({
      data: { circleId: circle.id, userId: ownerId, role: 'owner' },
    });

    return circle;
  },

  async getCircle(circleId: string, userId: string) {
    const circle = await prisma.circle.findUnique({
      where: { id: circleId },
      include: { members: { include: { user: { include: { profile: true } } } } },
    });
    if (!circle || circle.deletedAt) {
      throw new AppError(404, 'CIRCLE_NOT_FOUND', 'Circle not found');
    }

    // Check membership
    const isMember = circle.members.some((m) => m.userId === userId);
    if (!isMember && circle.privacyMode !== 'invite_only') {
      throw new AppError(403, 'NOT_A_MEMBER', 'You are not a member of this circle');
    }

    return {
      ...circle,
      memberCount: circle.members.length,
      members: circle.members.map((m) => ({
        id: m.id,
        userId: m.userId,
        username: m.user.username,
        displayName: m.user.displayName,
        avatarUrl: m.user.profile?.avatarMediaId,
        role: m.role,
        joinedAt: m.joinedAt.toISOString(),
      })),
    };
  },

  async getUserCircles(userId: string) {
    const memberships = await prisma.circleMember.findMany({
      where: { userId },
      include: {
        circle: {
          include: { _count: { select: { members: true } } },
        },
      },
    });

    return memberships.map((m) => ({
      id: m.circle.id,
      name: m.circle.name,
      description: m.circle.description,
      type: m.circle.type,
      role: m.role,
      memberCount: m.circle._count.members,
      createdAt: m.circle.createdAt.toISOString(),
    }));
  },

  async updateCircle(circleId: string, userId: string, data: { name?: string; description?: string }) {
    const circle = await prisma.circle.findUnique({ where: { id: circleId } });
    if (!circle || circle.deletedAt) {
      throw new AppError(404, 'CIRCLE_NOT_FOUND', 'Circle not found');
    }

    const member = await prisma.circleMember.findFirst({
      where: { circleId, userId, role: { in: ['owner', 'admin'] } },
    });
    if (!member) {
      throw new AppError(403, 'INSUFFICIENT_ROLE', 'Must be owner or admin');
    }

    return prisma.circle.update({
      where: { id: circleId },
      data,
    });
  },

  async deleteCircle(circleId: string, userId: string) {
    const circle = await prisma.circle.findUnique({ where: { id: circleId } });
    if (!circle || circle.deletedAt) {
      throw new AppError(404, 'CIRCLE_NOT_FOUND', 'Circle not found');
    }
    if (circle.ownerId !== userId) {
      throw new AppError(403, 'NOT_OWNER', 'Only the owner can delete a circle');
    }

    await prisma.circle.update({
      where: { id: circleId },
      data: { deletedAt: new Date() },
    });
  },

  async addCircleMember(circleId: string, actorId: string, targetId: string) {
    const circle = await prisma.circle.findUnique({
      where: { id: circleId },
      include: { _count: { select: { members: true } } },
    });
    if (!circle || circle.deletedAt) {
      throw new AppError(404, 'CIRCLE_NOT_FOUND', 'Circle not found');
    }

    const actorMember = await prisma.circleMember.findFirst({
      where: { circleId, userId: actorId, role: { in: ['owner', 'admin'] } },
    });
    if (!actorMember) {
      throw new AppError(403, 'INSUFFICIENT_ROLE', 'Must be owner or admin to add members');
    }

    if (circle._count.members >= circle.memberLimit) {
      throw new AppError(400, 'CIRCLE_FULL', 'Circle has reached member limit');
    }

    const existing = await prisma.circleMember.findUnique({
      where: { circleId_userId: { circleId, userId: targetId } },
    });
    if (existing) {
      throw new AppError(409, 'ALREADY_MEMBER', 'User is already a member');
    }

    return prisma.circleMember.create({
      data: { circleId, userId: targetId },
    });
  },

  async removeCircleMember(circleId: string, actorId: string, targetId: string) {
    const actorMember = await prisma.circleMember.findFirst({
      where: { circleId, userId: actorId, role: { in: ['owner', 'admin'] } },
    });

    // Allow self-removal or admin removal
    if (!actorMember && actorId !== targetId) {
      throw new AppError(403, 'INSUFFICIENT_ROLE', 'Must be owner or admin');
    }

    const targetMember = await prisma.circleMember.findFirst({
      where: { circleId, userId: targetId },
    });
    if (!targetMember) {
      throw new AppError(404, 'NOT_A_MEMBER', 'User is not a member');
    }
    if (targetMember.role === 'owner') {
      throw new AppError(400, 'CANNOT_REMOVE_OWNER', 'Cannot remove the owner');
    }

    await prisma.circleMember.delete({ where: { id: targetMember.id } });
  },
};

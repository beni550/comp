import { prisma } from '../../main';
import { AppError } from '../../middleware/error-handler';
import { createAuditLog } from '../../common/audit';
import { moderationService } from '../moderation/moderation.service';

export const adminService = {
  async getDashboard() {
    const now = new Date();
    const twentyFourHoursAgo = new Date(now.getTime() - 24 * 60 * 60 * 1000);

    const [signups24h, totalUsers, totalContent, moderationQueueSize, totalReports] = await Promise.all([
      prisma.user.count({ where: { createdAt: { gte: twentyFourHoursAgo } } }),
      prisma.user.count({ where: { deletedAt: null } }),
      prisma.contentItem.count({ where: { status: 'published', deletedAt: null } }),
      prisma.moderationCase.count({ where: { status: { in: ['open', 'in_review'] } } }),
      prisma.report.count(),
    ]);

    return {
      signups24h,
      dau: 0, // Would require session/analytics tracking
      moderationQueueSize,
      storageUsageBytes: 0, // Would require S3 query
      activeUsers: totalUsers,
      totalContent,
      totalReports,
    };
  },

  async lookupUser(query: string) {
    const user = await prisma.user.findFirst({
      where: {
        OR: [
          { username: query },
          { email: query },
          { phone: query },
          { id: query },
        ],
      },
      include: { profile: true, settings: true },
    });

    if (!user) {
      throw new AppError(404, 'USER_NOT_FOUND', 'User not found');
    }

    return user;
  },

  async suspendUser(actorId: string, userId: string, reason: string) {
    const user = await prisma.user.findUnique({ where: { id: userId } });
    if (!user) throw new AppError(404, 'USER_NOT_FOUND', 'User not found');
    if (user.role === 'admin') throw new AppError(403, 'CANNOT_SUSPEND_ADMIN', 'Cannot suspend an admin');

    await prisma.user.update({
      where: { id: userId },
      data: { status: 'suspended' },
    });

    // Revoke all sessions
    await prisma.session.updateMany({
      where: { userId, revokedAt: null },
      data: { revokedAt: new Date() },
    });

    await createAuditLog({
      actorId,
      action: 'admin.user_suspended',
      targetType: 'user',
      targetId: userId,
      metadata: { reason },
    });
  },

  async banUser(actorId: string, userId: string, reason: string) {
    const user = await prisma.user.findUnique({ where: { id: userId } });
    if (!user) throw new AppError(404, 'USER_NOT_FOUND', 'User not found');
    if (user.role === 'admin') throw new AppError(403, 'CANNOT_BAN_ADMIN', 'Cannot ban an admin');

    await prisma.user.update({
      where: { id: userId },
      data: { status: 'banned' },
    });

    await prisma.session.updateMany({
      where: { userId, revokedAt: null },
      data: { revokedAt: new Date() },
    });

    await createAuditLog({
      actorId,
      action: 'admin.user_banned',
      targetType: 'user',
      targetId: userId,
      metadata: { reason },
    });
  },

  async reinstateUser(actorId: string, userId: string) {
    const user = await prisma.user.findUnique({ where: { id: userId } });
    if (!user) throw new AppError(404, 'USER_NOT_FOUND', 'User not found');

    await prisma.user.update({
      where: { id: userId },
      data: { status: 'active' },
    });

    await createAuditLog({
      actorId,
      action: 'admin.user_reinstated',
      targetType: 'user',
      targetId: userId,
    });
  },

  async getModerationQueue(status?: string, priority?: string, cursor?: string, limit = 20) {
    return moderationService.getModerationQueue(status, priority, cursor, limit);
  },

  async resolveCase(actorId: string, caseId: string, data: {
    resolution: string;
    notes?: string;
    actionType?: string;
  }) {
    return moderationService.resolveCase(actorId, caseId, data);
  },

  // ─── Feature Flags ──────────────────────────────────────
  async getFeatureFlags() {
    return prisma.featureFlag.findMany({ orderBy: { key: 'asc' } });
  },

  async createFeatureFlag(data: {
    key: string;
    description?: string;
    isEnabled?: boolean;
    rolloutPercentage?: number;
    environment?: string;
  }) {
    const existing = await prisma.featureFlag.findUnique({ where: { key: data.key } });
    if (existing) {
      throw new AppError(409, 'FLAG_EXISTS', 'Feature flag with this key already exists');
    }

    return prisma.featureFlag.create({
      data: {
        key: data.key,
        description: data.description || '',
        isEnabled: data.isEnabled ?? false,
        rolloutPercentage: data.rolloutPercentage ?? 0,
        environment: data.environment || 'development',
      },
    });
  },

  async updateFeatureFlag(flagId: string, data: {
    isEnabled?: boolean;
    rolloutPercentage?: number;
    description?: string;
  }) {
    const flag = await prisma.featureFlag.findUnique({ where: { id: flagId } });
    if (!flag) throw new AppError(404, 'FLAG_NOT_FOUND', 'Feature flag not found');

    return prisma.featureFlag.update({
      where: { id: flagId },
      data,
    });
  },

  async deleteFeatureFlag(flagId: string) {
    const flag = await prisma.featureFlag.findUnique({ where: { id: flagId } });
    if (!flag) throw new AppError(404, 'FLAG_NOT_FOUND', 'Feature flag not found');

    await prisma.featureFlag.delete({ where: { id: flagId } });
  },

  // ─── Audit Logs ─────────────────────────────────────────
  async getAuditLogs(filters: {
    actorId?: string;
    action?: string;
    cursor?: string;
    limit?: number;
  }) {
    const where: Record<string, unknown> = {};
    if (filters.actorId) where.actorId = filters.actorId;
    if (filters.action) where.action = { contains: filters.action };
    if (filters.cursor) where.createdAt = { lt: new Date(filters.cursor) };

    const limit = filters.limit || 50;
    const logs = await prisma.auditLog.findMany({
      where,
      include: { actor: { select: { id: true, username: true } } },
      take: limit + 1,
      orderBy: { createdAt: 'desc' },
    });

    const hasMore = logs.length > limit;
    if (hasMore) logs.pop();

    return {
      items: logs,
      cursor: logs.length > 0 ? logs[logs.length - 1].createdAt.toISOString() : undefined,
      hasMore,
    };
  },
};

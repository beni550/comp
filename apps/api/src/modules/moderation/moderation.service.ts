import { prisma } from '../../main';
import { AppError } from '../../middleware/error-handler';
import { createAuditLog } from '../../common/audit';
import { APP_CONFIG } from '@vybe/config';

export const moderationService = {
  async submitReport(userId: string, data: {
    targetType: string;
    targetId: string;
    reason: string;
    description?: string;
  }) {
    // Check if user already reported this target
    const existing = await prisma.report.findFirst({
      where: {
        reporterId: userId,
        targetType: data.targetType,
        targetId: data.targetId,
        status: { in: ['open', 'in_review'] },
      },
    });
    if (existing) {
      throw new AppError(409, 'ALREADY_REPORTED', 'You have already reported this');
    }

    // Capture evidence snapshot
    let evidenceSnapshot = {};
    if (data.targetType === 'content') {
      const content = await prisma.contentItem.findUnique({ where: { id: data.targetId } });
      if (content) {
        evidenceSnapshot = { caption: content.caption, type: content.type, authorId: content.authorId };
      }
    } else if (data.targetType === 'comment') {
      const comment = await prisma.comment.findUnique({ where: { id: data.targetId } });
      if (comment) {
        evidenceSnapshot = { body: comment.body, authorId: comment.authorId };
      }
    }

    const report = await prisma.$transaction(async (tx) => {
      const r = await tx.report.create({
        data: {
          reporterId: userId,
          targetType: data.targetType,
          targetId: data.targetId,
          reason: data.reason,
          description: data.description,
          evidenceSnapshot,
        },
      });

      // Auto-create moderation case
      await tx.moderationCase.create({
        data: {
          reportId: r.id,
          priority: data.reason === 'violence' || data.reason === 'hate_speech' ? 'high' : 'normal',
          slaDeadline: new Date(Date.now() + APP_CONFIG.REPORT_SLA_HOURS * 60 * 60 * 1000),
        },
      });

      return r;
    });

    await createAuditLog({
      actorId: userId,
      action: 'report.submitted',
      targetType: data.targetType,
      targetId: data.targetId,
      metadata: { reason: data.reason },
    });

    return report;
  },

  async getUserReports(userId: string) {
    return prisma.report.findMany({
      where: { reporterId: userId },
      orderBy: { createdAt: 'desc' },
    });
  },

  // ─── Admin/Mod endpoints ────────────────────────────────
  async getModerationQueue(status?: string, priority?: string, cursor?: string, limit = 20) {
    const where: Record<string, unknown> = {};
    if (status) where.status = status;
    if (priority) where.priority = priority;
    if (cursor) where.id = { gt: cursor };

    const cases = await prisma.moderationCase.findMany({
      where,
      include: {
        report: { include: { reporter: { select: { id: true, username: true } } } },
        assignee: { select: { id: true, username: true } },
      },
      take: limit + 1,
      orderBy: [{ priority: 'desc' }, { createdAt: 'asc' }],
    });

    const hasMore = cases.length > limit;
    if (hasMore) cases.pop();

    return {
      items: cases,
      cursor: cases.length > 0 ? cases[cases.length - 1].id : undefined,
      hasMore,
    };
  },

  async resolveCase(actorId: string, caseId: string, data: {
    resolution: string;
    notes?: string;
    actionType?: string;
  }) {
    const modCase = await prisma.moderationCase.findUnique({
      where: { id: caseId },
      include: { report: true },
    });
    if (!modCase) {
      throw new AppError(404, 'CASE_NOT_FOUND', 'Moderation case not found');
    }

    await prisma.$transaction(async (tx) => {
      await tx.moderationCase.update({
        where: { id: caseId },
        data: {
          status: 'resolved',
          resolution: data.resolution,
          notes: data.notes,
          assigneeId: actorId,
          resolvedAt: new Date(),
        },
      });

      await tx.report.update({
        where: { id: modCase.reportId },
        data: { status: 'resolved' },
      });

      if (data.actionType) {
        await tx.moderationAction.create({
          data: {
            caseId,
            actorId,
            actionType: data.actionType,
            targetType: modCase.report.targetType,
            targetId: modCase.report.targetId,
            reason: data.notes || '',
          },
        });

        // Apply action
        if (data.actionType === 'hide' && modCase.report.targetType === 'content') {
          await tx.contentItem.update({
            where: { id: modCase.report.targetId },
            data: { status: 'hidden_review' },
          });
        } else if (data.actionType === 'remove' && modCase.report.targetType === 'content') {
          await tx.contentItem.update({
            where: { id: modCase.report.targetId },
            data: { status: 'removed', deletedAt: new Date() },
          });
        } else if (data.actionType === 'suspend' && modCase.report.targetType === 'user') {
          await tx.user.update({
            where: { id: modCase.report.targetId },
            data: { status: 'suspended' },
          });
        } else if (data.actionType === 'ban' && modCase.report.targetType === 'user') {
          await tx.user.update({
            where: { id: modCase.report.targetId },
            data: { status: 'banned' },
          });
        }
      }
    });

    await createAuditLog({
      actorId,
      action: 'moderation.resolved',
      targetType: 'case',
      targetId: caseId,
      metadata: { resolution: data.resolution, actionType: data.actionType },
    });
  },
};

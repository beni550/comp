import { Prisma } from '@prisma/client';
import { prisma } from '../main';

export async function createAuditLog(params: {
  actorId?: string;
  action: string;
  targetType?: string;
  targetId?: string;
  metadata?: Record<string, unknown>;
  ip?: string;
}): Promise<void> {
  await prisma.auditLog.create({
    data: {
      actorId: params.actorId,
      action: params.action,
      targetType: params.targetType,
      targetId: params.targetId,
      metadata: (params.metadata || {}) as Prisma.InputJsonValue,
      ip: params.ip,
    },
  });
}

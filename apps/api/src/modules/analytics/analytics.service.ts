import { Prisma } from '@prisma/client';
import { prisma } from '../../main';

export const analyticsService = {
  async trackEvent(userId: string | undefined, data: {
    eventType: string;
    attributes: Record<string, unknown>;
    sessionId?: string;
  }) {
    return prisma.analyticsEvent.create({
      data: {
        eventType: data.eventType,
        userId,
        attributes: data.attributes as Prisma.InputJsonValue,
        sessionId: data.sessionId,
      },
    });
  },

  async getSummary(filters: {
    eventType?: string;
    startDate?: string;
    endDate?: string;
  }) {
    const where: Record<string, unknown> = {};
    if (filters.eventType) where.eventType = filters.eventType;
    if (filters.startDate || filters.endDate) {
      where.createdAt = {};
      if (filters.startDate) (where.createdAt as Record<string, unknown>).gte = new Date(filters.startDate);
      if (filters.endDate) (where.createdAt as Record<string, unknown>).lte = new Date(filters.endDate);
    }

    const total = await prisma.analyticsEvent.count({ where });

    // Group by event type
    const eventTypes = await prisma.analyticsEvent.groupBy({
      by: ['eventType'],
      where,
      _count: { id: true },
      orderBy: { _count: { id: 'desc' } },
      take: 20,
    });

    return {
      total,
      byEventType: eventTypes.map((e) => ({
        eventType: e.eventType,
        count: e._count.id,
      })),
    };
  },
};

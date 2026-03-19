import { Prisma } from '@prisma/client';
import { prisma } from '../../main';
import { AppError } from '../../middleware/error-handler';

export const notificationsService = {
  async createNotification(data: {
    recipientId: string;
    type: string;
    title: string;
    body: string;
    data?: Record<string, unknown>;
    groupKey?: string;
  }) {
    // Dedup check
    if (data.groupKey) {
      const recent = await prisma.notification.findFirst({
        where: {
          recipientId: data.recipientId,
          groupKey: data.groupKey,
          createdAt: { gte: new Date(Date.now() - 5 * 60 * 1000) },
        },
      });
      if (recent) return recent;
    }

    const notification = await prisma.notification.create({
      data: {
        recipientId: data.recipientId,
        type: data.type,
        title: data.title,
        body: data.body,
        data: (data.data || {}) as Prisma.InputJsonValue,
        groupKey: data.groupKey,
      },
    });

    // Create outbox entry for in-app delivery
    await prisma.notificationOutbox.create({
      data: {
        notificationId: notification.id,
        channel: 'in_app',
        status: 'sent',
        sentAt: new Date(),
      },
    });

    return notification;
  },

  async getNotifications(userId: string, cursor?: string, limit = 20, unreadOnly = false) {
    const where: Record<string, unknown> = { recipientId: userId };
    if (cursor) {
      where.createdAt = { lt: new Date(cursor) };
    }
    if (unreadOnly) {
      where.isRead = false;
    }

    const notifications = await prisma.notification.findMany({
      where,
      take: limit + 1,
      orderBy: { createdAt: 'desc' },
    });

    const hasMore = notifications.length > limit;
    if (hasMore) notifications.pop();

    return {
      items: notifications.map((n) => ({
        id: n.id,
        type: n.type,
        title: n.title,
        body: n.body,
        data: n.data,
        isRead: n.isRead,
        createdAt: n.createdAt.toISOString(),
      })),
      cursor: notifications.length > 0
        ? notifications[notifications.length - 1].createdAt.toISOString()
        : undefined,
      hasMore,
    };
  },

  async markRead(userId: string, notificationId: string) {
    const notification = await prisma.notification.findFirst({
      where: { id: notificationId, recipientId: userId },
    });
    if (!notification) {
      throw new AppError(404, 'NOTIFICATION_NOT_FOUND', 'Notification not found');
    }

    await prisma.notification.update({
      where: { id: notificationId },
      data: { isRead: true },
    });
  },

  async markAllRead(userId: string) {
    const result = await prisma.notification.updateMany({
      where: { recipientId: userId, isRead: false },
      data: { isRead: true },
    });
    return { updatedCount: result.count };
  },

  async getUnreadCount(userId: string) {
    return prisma.notification.count({
      where: { recipientId: userId, isRead: false },
    });
  },
};

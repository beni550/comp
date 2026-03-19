import { Router, Request, Response, NextFunction } from 'express';
import { notificationsService } from './notifications.service';
import { authenticate } from '../../middleware/auth';
import { success, paginated } from '../../common/response';

export const notificationsRouter = Router();

// GET /notifications
notificationsRouter.get('/', authenticate, async (req: Request, res: Response, next: NextFunction) => {
  try {
    const result = await notificationsService.getNotifications(
      req.user!.id,
      req.query.cursor as string,
      parseInt(req.query.limit as string) || 20,
      req.query.unreadOnly === 'true'
    );
    paginated(res, req, result.items, { cursor: result.cursor, hasMore: result.hasMore });
  } catch (err) {
    next(err);
  }
});

// GET /notifications/unread-count
notificationsRouter.get('/unread-count', authenticate, async (req: Request, res: Response, next: NextFunction) => {
  try {
    const count = await notificationsService.getUnreadCount(req.user!.id);
    success(res, req, { count });
  } catch (err) {
    next(err);
  }
});

// POST /notifications/:notificationId/read
notificationsRouter.post('/:notificationId/read', authenticate, async (req: Request, res: Response, next: NextFunction) => {
  try {
    await notificationsService.markRead(req.user!.id, req.params.notificationId);
    success(res, req, { success: true });
  } catch (err) {
    next(err);
  }
});

// POST /notifications/read-all
notificationsRouter.post('/read-all', authenticate, async (req: Request, res: Response, next: NextFunction) => {
  try {
    const result = await notificationsService.markAllRead(req.user!.id);
    success(res, req, result);
  } catch (err) {
    next(err);
  }
});

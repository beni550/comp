import { Router, Request, Response, NextFunction } from 'express';
import { analyticsService } from './analytics.service';
import { authenticate, optionalAuth, requireAdmin } from '../../middleware/auth';
import { success } from '../../common/response';
import { trackEventSchema } from '@vybe/validation';

export const analyticsRouter = Router();

// POST /analytics/events
analyticsRouter.post('/events', optionalAuth, async (req: Request, res: Response, next: NextFunction) => {
  try {
    const data = trackEventSchema.parse(req.body);
    const event = await analyticsService.trackEvent(req.user?.id, data);
    success(res, req, { id: event.id }, 201);
  } catch (err) {
    next(err);
  }
});

// GET /analytics/summary (admin only)
analyticsRouter.get('/summary', authenticate, requireAdmin, async (req: Request, res: Response, next: NextFunction) => {
  try {
    const summary = await analyticsService.getSummary({
      eventType: req.query.eventType as string,
      startDate: req.query.startDate as string,
      endDate: req.query.endDate as string,
    });
    success(res, req, summary);
  } catch (err) {
    next(err);
  }
});

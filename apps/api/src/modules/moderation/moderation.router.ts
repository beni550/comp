import { Router, Request, Response, NextFunction } from 'express';
import { moderationService } from './moderation.service';
import { authenticate, requireAdmin } from '../../middleware/auth';
import { success, paginated } from '../../common/response';
import { createReportSchema } from '@vybe/validation';

export const moderationRouter = Router();

// POST /reports
moderationRouter.post('/', authenticate, async (req: Request, res: Response, next: NextFunction) => {
  try {
    const data = createReportSchema.parse(req.body);
    const report = await moderationService.submitReport(req.user!.id, data);
    success(res, req, report, 201);
  } catch (err) {
    next(err);
  }
});

// GET /reports/mine
moderationRouter.get('/mine', authenticate, async (req: Request, res: Response, next: NextFunction) => {
  try {
    const reports = await moderationService.getUserReports(req.user!.id);
    success(res, req, reports);
  } catch (err) {
    next(err);
  }
});

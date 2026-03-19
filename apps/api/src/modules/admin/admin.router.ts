import { Router, Request, Response, NextFunction } from 'express';
import { adminService } from './admin.service';
import { authenticate, requireAdmin } from '../../middleware/auth';
import { success, paginated } from '../../common/response';
import { createFeatureFlagSchema, updateFeatureFlagSchema } from '@vybe/validation';

export const adminRouter = Router();

// All admin routes require authentication + admin/moderator role
adminRouter.use(authenticate, requireAdmin);

// GET /admin/dashboard
adminRouter.get('/dashboard', async (req: Request, res: Response, next: NextFunction) => {
  try {
    const stats = await adminService.getDashboard();
    success(res, req, stats);
  } catch (err) {
    next(err);
  }
});

// GET /admin/users/lookup
adminRouter.get('/users/lookup', async (req: Request, res: Response, next: NextFunction) => {
  try {
    const q = req.query.q as string;
    if (!q) {
      res.status(400).json({ error: { code: 'MISSING_QUERY', message: 'Query parameter required', requestId: req.requestId } });
      return;
    }
    const user = await adminService.lookupUser(q);
    success(res, req, user);
  } catch (err) {
    next(err);
  }
});

// POST /admin/users/:userId/suspend
adminRouter.post('/users/:userId/suspend', async (req: Request, res: Response, next: NextFunction) => {
  try {
    await adminService.suspendUser(req.user!.id, req.params.userId, req.body.reason || '');
    success(res, req, { success: true });
  } catch (err) {
    next(err);
  }
});

// POST /admin/users/:userId/ban
adminRouter.post('/users/:userId/ban', async (req: Request, res: Response, next: NextFunction) => {
  try {
    await adminService.banUser(req.user!.id, req.params.userId, req.body.reason || '');
    success(res, req, { success: true });
  } catch (err) {
    next(err);
  }
});

// POST /admin/users/:userId/reinstate
adminRouter.post('/users/:userId/reinstate', async (req: Request, res: Response, next: NextFunction) => {
  try {
    await adminService.reinstateUser(req.user!.id, req.params.userId);
    success(res, req, { success: true });
  } catch (err) {
    next(err);
  }
});

// GET /admin/moderation/queue
adminRouter.get('/moderation/queue', async (req: Request, res: Response, next: NextFunction) => {
  try {
    const result = await adminService.getModerationQueue(
      req.query.status as string,
      req.query.priority as string,
      req.query.cursor as string,
      parseInt(req.query.limit as string) || 20
    );
    paginated(res, req, result.items, { cursor: result.cursor, hasMore: result.hasMore });
  } catch (err) {
    next(err);
  }
});

// POST /admin/moderation/cases/:caseId/resolve
adminRouter.post('/moderation/cases/:caseId/resolve', async (req: Request, res: Response, next: NextFunction) => {
  try {
    await adminService.resolveCase(req.user!.id, req.params.caseId, req.body);
    success(res, req, { success: true });
  } catch (err) {
    next(err);
  }
});

// GET /admin/feature-flags
adminRouter.get('/feature-flags', async (req: Request, res: Response, next: NextFunction) => {
  try {
    const flags = await adminService.getFeatureFlags();
    success(res, req, flags);
  } catch (err) {
    next(err);
  }
});

// POST /admin/feature-flags
adminRouter.post('/feature-flags', async (req: Request, res: Response, next: NextFunction) => {
  try {
    const data = createFeatureFlagSchema.parse(req.body);
    const flag = await adminService.createFeatureFlag(data);
    success(res, req, flag, 201);
  } catch (err) {
    next(err);
  }
});

// PATCH /admin/feature-flags/:flagId
adminRouter.patch('/feature-flags/:flagId', async (req: Request, res: Response, next: NextFunction) => {
  try {
    const data = updateFeatureFlagSchema.parse(req.body);
    const flag = await adminService.updateFeatureFlag(req.params.flagId, data);
    success(res, req, flag);
  } catch (err) {
    next(err);
  }
});

// DELETE /admin/feature-flags/:flagId
adminRouter.delete('/feature-flags/:flagId', async (req: Request, res: Response, next: NextFunction) => {
  try {
    await adminService.deleteFeatureFlag(req.params.flagId);
    success(res, req, { success: true });
  } catch (err) {
    next(err);
  }
});

// GET /admin/audit-logs
adminRouter.get('/audit-logs', async (req: Request, res: Response, next: NextFunction) => {
  try {
    const result = await adminService.getAuditLogs({
      actorId: req.query.actorId as string,
      action: req.query.action as string,
      cursor: req.query.cursor as string,
      limit: parseInt(req.query.limit as string) || 50,
    });
    paginated(res, req, result.items, { cursor: result.cursor, hasMore: result.hasMore });
  } catch (err) {
    next(err);
  }
});

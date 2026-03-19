import { Router, Request, Response, NextFunction } from 'express';
import { usersService } from './users.service';
import { authenticate, optionalAuth } from '../../middleware/auth';
import { success } from '../../common/response';
import { updateUserSchema, updateProfileSchema, updateSettingsSchema } from '@vybe/validation';

export const usersRouter = Router();

// GET /users/me
usersRouter.get('/me', authenticate, async (req: Request, res: Response, next: NextFunction) => {
  try {
    const user = await usersService.getMe(req.user!.id);
    success(res, req, user);
  } catch (err) {
    next(err);
  }
});

// PATCH /users/me
usersRouter.patch('/me', authenticate, async (req: Request, res: Response, next: NextFunction) => {
  try {
    const data = updateUserSchema.parse(req.body);
    const user = await usersService.updateUser(req.user!.id, data);
    success(res, req, user);
  } catch (err) {
    next(err);
  }
});

// GET /users/:userId
usersRouter.get('/:userId', optionalAuth, async (req: Request, res: Response, next: NextFunction) => {
  try {
    const user = await usersService.getPublicUser(req.params.userId, req.user?.id);
    success(res, req, user);
  } catch (err) {
    next(err);
  }
});

// PATCH /users/me/profile
usersRouter.patch('/me/profile', authenticate, async (req: Request, res: Response, next: NextFunction) => {
  try {
    const data = updateProfileSchema.parse(req.body);
    const profile = await usersService.updateProfile(req.user!.id, data);
    success(res, req, profile);
  } catch (err) {
    next(err);
  }
});

// GET /users/me/settings
usersRouter.get('/me/settings', authenticate, async (req: Request, res: Response, next: NextFunction) => {
  try {
    const user = await usersService.getMe(req.user!.id);
    success(res, req, user.settings);
  } catch (err) {
    next(err);
  }
});

// PATCH /users/me/settings
usersRouter.patch('/me/settings', authenticate, async (req: Request, res: Response, next: NextFunction) => {
  try {
    const data = updateSettingsSchema.parse(req.body);
    const settings = await usersService.updateSettings(req.user!.id, data);
    success(res, req, settings);
  } catch (err) {
    next(err);
  }
});

// GET /users/search
usersRouter.get('/search', authenticate, async (req: Request, res: Response, next: NextFunction) => {
  try {
    const q = (req.query.q as string) || '';
    const limit = Math.min(parseInt(req.query.limit as string) || 20, 50);
    const users = await usersService.searchUsers(q, limit);
    success(res, req, users.map((u) => ({
      id: u.id,
      username: u.username,
      displayName: u.displayName,
      avatarUrl: u.profile?.avatarMediaId,
      bio: u.profile?.bio,
    })));
  } catch (err) {
    next(err);
  }
});

// DELETE /users/me
usersRouter.delete('/me', authenticate, async (req: Request, res: Response, next: NextFunction) => {
  try {
    await usersService.deleteUser(req.user!.id);
    success(res, req, { success: true });
  } catch (err) {
    next(err);
  }
});

// GET /users/me/onboarding
usersRouter.get('/me/onboarding', authenticate, async (req: Request, res: Response, next: NextFunction) => {
  try {
    const progress = await usersService.getOnboardingProgress(req.user!.id);
    success(res, req, progress);
  } catch (err) {
    next(err);
  }
});

// POST /users/me/onboarding/complete
usersRouter.post('/me/onboarding/complete', authenticate, async (req: Request, res: Response, next: NextFunction) => {
  try {
    const { step } = req.body;
    const progress = await usersService.completeOnboardingStep(req.user!.id, step);
    success(res, req, progress);
  } catch (err) {
    next(err);
  }
});

// POST /users/me/onboarding/skip
usersRouter.post('/me/onboarding/skip', authenticate, async (req: Request, res: Response, next: NextFunction) => {
  try {
    const { step } = req.body;
    const progress = await usersService.skipOnboardingStep(req.user!.id, step);
    success(res, req, progress);
  } catch (err) {
    next(err);
  }
});

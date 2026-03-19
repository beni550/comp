import { Router, Request, Response, NextFunction } from 'express';
import { authService } from './auth.service';
import { authenticate } from '../../middleware/auth';
import { success } from '../../common/response';
import { authRateLimiter, otpRateLimiter } from '../../middleware/rate-limit';
import {
  registerPhoneStartSchema,
  registerPhoneVerifySchema,
  registerEmailStartSchema,
  loginSchema,
} from '@vybe/validation';

export const authRouter = Router();

// POST /auth/register/phone/start
authRouter.post('/register/phone/start', otpRateLimiter, async (req: Request, res: Response, next: NextFunction) => {
  try {
    const data = registerPhoneStartSchema.parse(req.body);
    const result = await authService.registerPhoneStart(data.phone);
    success(res, req, result, 201);
  } catch (err) {
    next(err);
  }
});

// POST /auth/register/phone/verify
authRouter.post('/register/phone/verify', authRateLimiter, async (req: Request, res: Response, next: NextFunction) => {
  try {
    const data = registerPhoneVerifySchema.parse(req.body);
    const result = await authService.registerPhoneVerify(data);
    success(res, req, {
      user: result.user,
      accessToken: result.accessToken,
      refreshToken: result.refreshToken,
    }, 201);
  } catch (err) {
    next(err);
  }
});

// POST /auth/register/email/start
authRouter.post('/register/email/start', otpRateLimiter, async (req: Request, res: Response, next: NextFunction) => {
  try {
    const data = registerEmailStartSchema.parse(req.body);
    const result = await authService.registerEmailStart(data.email);
    success(res, req, result, 201);
  } catch (err) {
    next(err);
  }
});

// POST /auth/login
authRouter.post('/login', authRateLimiter, async (req: Request, res: Response, next: NextFunction) => {
  try {
    const data = loginSchema.parse(req.body);
    const result = await authService.login({
      ...data,
      ip: req.ip,
      userAgent: req.headers['user-agent'],
    });
    success(res, req, {
      user: result.user,
      accessToken: result.accessToken,
      refreshToken: result.refreshToken,
    });
  } catch (err) {
    next(err);
  }
});

// POST /auth/refresh
authRouter.post('/refresh', async (req: Request, res: Response, next: NextFunction) => {
  try {
    const { refreshToken } = req.body;
    if (!refreshToken) {
      res.status(400).json({ error: { code: 'MISSING_TOKEN', message: 'Refresh token required', requestId: req.requestId } });
      return;
    }
    const result = await authService.refreshToken(refreshToken);
    success(res, req, result);
  } catch (err) {
    next(err);
  }
});

// POST /auth/logout
authRouter.post('/logout', authenticate, async (req: Request, res: Response, next: NextFunction) => {
  try {
    await authService.logout(req.user!.id);
    success(res, req, { success: true });
  } catch (err) {
    next(err);
  }
});

// POST /auth/logout-all
authRouter.post('/logout-all', authenticate, async (req: Request, res: Response, next: NextFunction) => {
  try {
    const result = await authService.logoutAll(req.user!.id);
    success(res, req, result);
  } catch (err) {
    next(err);
  }
});

// GET /auth/sessions
authRouter.get('/sessions', authenticate, async (req: Request, res: Response, next: NextFunction) => {
  try {
    const sessions = await authService.getSessions(req.user!.id);
    success(res, req, {
      sessions: sessions.map((s) => ({
        id: s.id,
        deviceId: s.deviceId,
        platform: s.device.platform,
        ip: s.ip,
        userAgent: s.userAgent,
        createdAt: s.createdAt.toISOString(),
        isCurrent: false, // Would need session matching logic
      })),
    });
  } catch (err) {
    next(err);
  }
});

// DELETE /auth/sessions/:sessionId
authRouter.delete('/sessions/:sessionId', authenticate, async (req: Request, res: Response, next: NextFunction) => {
  try {
    await authService.revokeSession(req.user!.id, req.params.sessionId);
    success(res, req, { success: true });
  } catch (err) {
    next(err);
  }
});

// GET /auth/username/check/:username
authRouter.get('/username/check/:username', async (req: Request, res: Response, next: NextFunction) => {
  try {
    const result = await authService.checkUsernameAvailability(req.params.username);
    success(res, req, result);
  } catch (err) {
    next(err);
  }
});

// POST /auth/recovery/start
authRouter.post('/recovery/start', otpRateLimiter, async (req: Request, res: Response, next: NextFunction) => {
  try {
    const { identifier } = req.body;
    if (!identifier) {
      res.status(400).json({ error: { code: 'MISSING_IDENTIFIER', message: 'Identifier required', requestId: req.requestId } });
      return;
    }
    const result = await authService.startRecovery(identifier);
    success(res, req, result);
  } catch (err) {
    next(err);
  }
});

// POST /auth/2fa/challenge
authRouter.post('/2fa/challenge', authenticate, async (req: Request, res: Response, next: NextFunction) => {
  try {
    const { method } = req.body;
    const user = await (await import('../../main')).prisma.user.findUnique({
      where: { id: req.user!.id },
    });
    if (!user) {
      res.status(404).json({ error: { code: 'USER_NOT_FOUND', message: 'User not found', requestId: req.requestId } });
      return;
    }
    const target = method === 'email' ? user.email : user.phone;
    if (!target) {
      res.status(400).json({ error: { code: 'NO_TARGET', message: `No ${method} configured`, requestId: req.requestId } });
      return;
    }
    const result = method === 'email'
      ? await authService.registerEmailStart(target)
      : await authService.registerPhoneStart(target);
    success(res, req, { challengeId: result.challengeId });
  } catch (err) {
    next(err);
  }
});

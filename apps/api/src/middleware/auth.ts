import { Request, Response, NextFunction } from 'express';
import jwt from 'jsonwebtoken';
import { prisma } from '../main';
import { AppError } from './error-handler';

export interface AuthUser {
  id: string;
  username: string;
  role: string;
  isMinor: boolean;
  status: string;
}

declare global {
  namespace Express {
    interface Request {
      user?: AuthUser;
    }
  }
}

const JWT_SECRET = process.env.JWT_SECRET || 'dev-secret-change-me';

export function authenticate(req: Request, _res: Response, next: NextFunction): void {
  const authHeader = req.headers.authorization;
  if (!authHeader?.startsWith('Bearer ')) {
    next(new AppError(401, 'UNAUTHORIZED', 'Missing or invalid authorization header'));
    return;
  }

  const token = authHeader.slice(7);
  try {
    const payload = jwt.verify(token, JWT_SECRET) as AuthUser & { exp: number };
    req.user = {
      id: payload.id,
      username: payload.username,
      role: payload.role,
      isMinor: payload.isMinor,
      status: payload.status,
    };
    next();
  } catch {
    next(new AppError(401, 'TOKEN_EXPIRED', 'Access token is invalid or expired'));
  }
}

export function optionalAuth(req: Request, _res: Response, next: NextFunction): void {
  const authHeader = req.headers.authorization;
  if (!authHeader?.startsWith('Bearer ')) {
    next();
    return;
  }

  const token = authHeader.slice(7);
  try {
    const payload = jwt.verify(token, JWT_SECRET) as AuthUser & { exp: number };
    req.user = {
      id: payload.id,
      username: payload.username,
      role: payload.role,
      isMinor: payload.isMinor,
      status: payload.status,
    };
  } catch {
    // Token invalid, continue as guest
  }
  next();
}

export function requireRole(...roles: string[]) {
  return (req: Request, _res: Response, next: NextFunction): void => {
    if (!req.user) {
      next(new AppError(401, 'UNAUTHORIZED', 'Authentication required'));
      return;
    }
    if (!roles.includes(req.user.role)) {
      next(new AppError(403, 'FORBIDDEN', 'Insufficient permissions'));
      return;
    }
    next();
  };
}

export function requireAdmin(req: Request, _res: Response, next: NextFunction): void {
  if (!req.user) {
    next(new AppError(401, 'UNAUTHORIZED', 'Authentication required'));
    return;
  }
  if (!['admin', 'moderator'].includes(req.user.role)) {
    next(new AppError(403, 'FORBIDDEN', 'Admin access required'));
    return;
  }
  next();
}

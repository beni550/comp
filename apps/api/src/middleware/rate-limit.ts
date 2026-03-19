import rateLimit from 'express-rate-limit';
import { APP_CONFIG } from '@vybe/config';

/**
 * General API rate limiter: 100 requests per minute per IP.
 */
export const generalRateLimiter = rateLimit({
  windowMs: APP_CONFIG.RATE_LIMIT_WINDOW_MS,
  max: APP_CONFIG.RATE_LIMIT_MAX_REQUESTS,
  standardHeaders: true,
  legacyHeaders: false,
  message: {
    error: {
      code: 'RATE_LIMIT_EXCEEDED',
      message: 'Too many requests, please try again later',
    },
  },
});

/**
 * Auth endpoint rate limiter: 10 requests per minute per IP.
 * Protects login, register, and password recovery endpoints.
 */
export const authRateLimiter = rateLimit({
  windowMs: APP_CONFIG.RATE_LIMIT_WINDOW_MS,
  max: APP_CONFIG.AUTH_RATE_LIMIT_MAX,
  standardHeaders: true,
  legacyHeaders: false,
  message: {
    error: {
      code: 'AUTH_RATE_LIMIT_EXCEEDED',
      message: 'Too many authentication attempts, please try again later',
    },
  },
});

/**
 * OTP rate limiter: 3 requests per minute per IP.
 * Protects OTP send/verify endpoints from abuse.
 */
export const otpRateLimiter = rateLimit({
  windowMs: APP_CONFIG.RATE_LIMIT_WINDOW_MS,
  max: APP_CONFIG.OTP_RATE_LIMIT_MAX,
  standardHeaders: true,
  legacyHeaders: false,
  message: {
    error: {
      code: 'OTP_RATE_LIMIT_EXCEEDED',
      message: 'Too many OTP requests, please try again later',
    },
  },
});

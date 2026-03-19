import express from 'express';
import cors from 'cors';
import helmet from 'helmet';
import morgan from 'morgan';
import { createServer } from 'http';
import { PrismaClient } from '@prisma/client';
import swaggerUi from 'swagger-ui-express';

import { errorHandler } from './middleware/error-handler';
import { requestId } from './middleware/request-id';
import { generalRateLimiter } from './middleware/rate-limit';
import { swaggerSpec } from './swagger';
import { authRouter } from './modules/auth/auth.router';
import { usersRouter } from './modules/users/users.router';
import { socialRouter } from './modules/social/social.router';
import { contentRouter } from './modules/content/content.router';
import { feedRouter } from './modules/feed/feed.router';
import { messagingRouter } from './modules/messaging/messaging.router';
import { notificationsRouter } from './modules/notifications/notifications.router';
import { moderationRouter } from './modules/moderation/moderation.router';
import { adminRouter } from './modules/admin/admin.router';
import { analyticsRouter } from './modules/analytics/analytics.router';
import { setupWebSocket } from './modules/messaging/websocket';

export const prisma = new PrismaClient();

const app = express();
const httpServer = createServer(app);

// ─── Global Middleware ──────────────────────────────────────
app.use(helmet());
app.use(cors({
  origin: process.env.CORS_ORIGIN || 'http://localhost:5173',
  credentials: true,
}));
app.use(express.json({ limit: '10mb' }));
app.use(morgan('dev'));
app.use(requestId);
app.use(generalRateLimiter);

// ─── Swagger Documentation ──────────────────────────────────
app.use('/api/docs', swaggerUi.serve, swaggerUi.setup(swaggerSpec, {
  customCss: '.swagger-ui .topbar { display: none }',
  customSiteTitle: 'VYBE API Documentation',
}));
app.get('/api/docs.json', (_req, res) => {
  res.json(swaggerSpec);
});

// ─── Health Check ───────────────────────────────────────────
app.get('/api/v1/health', (_req, res) => {
  res.json({ status: 'ok', timestamp: new Date().toISOString() });
});

// ─── Routes ─────────────────────────────────────────────────
app.use('/api/v1/auth', authRouter);
app.use('/api/v1/users', usersRouter);
app.use('/api/v1/social', socialRouter);
app.use('/api/v1/circles', socialRouter);
app.use('/api/v1/content', contentRouter);
app.use('/api/v1/media', contentRouter);
app.use('/api/v1/stories', contentRouter);
app.use('/api/v1/highlights', contentRouter);
app.use('/api/v1/feed', feedRouter);
app.use('/api/v1/discover', feedRouter);
app.use('/api/v1/feedback', feedRouter);
app.use('/api/v1/conversations', messagingRouter);
app.use('/api/v1/messages', messagingRouter);
app.use('/api/v1/notifications', notificationsRouter);
app.use('/api/v1/reports', moderationRouter);
app.use('/api/v1/admin', adminRouter);
app.use('/api/v1/analytics', analyticsRouter);

// ─── Error Handler ──────────────────────────────────────────
app.use(errorHandler);

// ─── WebSocket ──────────────────────────────────────────────
setupWebSocket(httpServer);

// ─── Start ──────────────────────────────────────────────────
const PORT = process.env.PORT || 3000;

httpServer.listen(PORT, () => {
  console.log(`VYBE API server running on port ${PORT}`);
  console.log(`WebSocket server attached`);
});

export { app, httpServer };

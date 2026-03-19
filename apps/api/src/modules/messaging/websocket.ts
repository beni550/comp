import { Server as HttpServer } from 'http';
import { Server, Socket } from 'socket.io';
import jwt from 'jsonwebtoken';

const JWT_SECRET = process.env.JWT_SECRET || 'dev-secret-change-me';

interface AuthenticatedSocket extends Socket {
  userId?: string;
}

export function setupWebSocket(httpServer: HttpServer): Server {
  const io = new Server(httpServer, {
    cors: {
      origin: process.env.CORS_ORIGIN || 'http://localhost:5173',
      credentials: true,
    },
    path: '/ws',
  });

  // Authentication middleware
  io.use((socket: AuthenticatedSocket, next) => {
    const token = socket.handshake.auth.token || socket.handshake.headers.authorization?.replace('Bearer ', '');
    if (!token) {
      return next(new Error('Authentication required'));
    }

    try {
      const payload = jwt.verify(token, JWT_SECRET) as { id: string };
      socket.userId = payload.id;
      next();
    } catch {
      next(new Error('Invalid token'));
    }
  });

  io.on('connection', (socket: AuthenticatedSocket) => {
    const userId = socket.userId!;
    console.log(`User ${userId} connected via WebSocket`);

    // Join user's personal room for notifications
    socket.join(`user:${userId}`);

    // ─── Chat Events ──────────────────────────────────
    socket.on('chat:join', (conversationId: string) => {
      socket.join(`conversation:${conversationId}`);
    });

    socket.on('chat:leave', (conversationId: string) => {
      socket.leave(`conversation:${conversationId}`);
    });

    socket.on('chat:message', (data: { conversationId: string; kind: string; body?: string }) => {
      // Broadcast to conversation room
      socket.to(`conversation:${data.conversationId}`).emit('chat:new_message', {
        conversationId: data.conversationId,
        senderId: userId,
        kind: data.kind,
        body: data.body,
        timestamp: new Date().toISOString(),
      });
    });

    socket.on('chat:typing', (data: { conversationId: string }) => {
      socket.to(`conversation:${data.conversationId}`).emit('chat:typing', {
        conversationId: data.conversationId,
        userId,
      });
    });

    socket.on('chat:stop_typing', (data: { conversationId: string }) => {
      socket.to(`conversation:${data.conversationId}`).emit('chat:stop_typing', {
        conversationId: data.conversationId,
        userId,
      });
    });

    socket.on('chat:read', (data: { conversationId: string; messageId: string }) => {
      socket.to(`conversation:${data.conversationId}`).emit('chat:read_receipt', {
        conversationId: data.conversationId,
        userId,
        messageId: data.messageId,
      });
    });

    // ─── Presence Events ──────────────────────────────
    socket.on('presence:online', () => {
      io.emit('presence:status', { userId, status: 'online' });
    });

    socket.on('disconnect', () => {
      console.log(`User ${userId} disconnected`);
      io.emit('presence:status', { userId, status: 'offline' });
    });
  });

  return io;
}

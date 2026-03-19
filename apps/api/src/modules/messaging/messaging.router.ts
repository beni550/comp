import { Router, Request, Response, NextFunction } from 'express';
import { messagingService } from './messaging.service';
import { authenticate } from '../../middleware/auth';
import { success, paginated } from '../../common/response';
import { createConversationSchema, sendMessageSchema } from '@vybe/validation';

export const messagingRouter = Router();

// POST /conversations
messagingRouter.post('/', authenticate, async (req: Request, res: Response, next: NextFunction) => {
  try {
    const data = createConversationSchema.parse(req.body);
    const conv = await messagingService.createConversation(req.user!.id, data);
    success(res, req, conv, 201);
  } catch (err) {
    next(err);
  }
});

// GET /conversations
messagingRouter.get('/', authenticate, async (req: Request, res: Response, next: NextFunction) => {
  try {
    const result = await messagingService.getUserConversations(
      req.user!.id,
      req.query.cursor as string,
      parseInt(req.query.limit as string) || 20
    );
    paginated(res, req, result.items, { cursor: result.cursor, hasMore: result.hasMore });
  } catch (err) {
    next(err);
  }
});

// GET /conversations/:conversationId
messagingRouter.get('/:conversationId', authenticate, async (req: Request, res: Response, next: NextFunction) => {
  try {
    const conv = await messagingService.getConversation(req.params.conversationId, req.user!.id);
    success(res, req, conv);
  } catch (err) {
    next(err);
  }
});

// POST /conversations/:conversationId/messages
messagingRouter.post('/:conversationId/messages', authenticate, async (req: Request, res: Response, next: NextFunction) => {
  try {
    const data = sendMessageSchema.parse(req.body);
    const msg = await messagingService.sendMessage(req.user!.id, req.params.conversationId, data);
    success(res, req, msg, 201);
  } catch (err) {
    next(err);
  }
});

// GET /conversations/:conversationId/messages
messagingRouter.get('/:conversationId/messages', authenticate, async (req: Request, res: Response, next: NextFunction) => {
  try {
    const result = await messagingService.getMessages(
      req.user!.id,
      req.params.conversationId,
      req.query.cursor as string,
      parseInt(req.query.limit as string) || 30
    );
    paginated(res, req, result.items, { cursor: result.cursor, hasMore: result.hasMore });
  } catch (err) {
    next(err);
  }
});

// DELETE /messages/:messageId
messagingRouter.delete('/:messageId', authenticate, async (req: Request, res: Response, next: NextFunction) => {
  try {
    await messagingService.deleteMessage(req.user!.id, req.params.messageId, req.query.forAll === 'true');
    success(res, req, { success: true });
  } catch (err) {
    next(err);
  }
});

// POST /messages/:messageId/react
messagingRouter.post('/:messageId/react', authenticate, async (req: Request, res: Response, next: NextFunction) => {
  try {
    const result = await messagingService.reactToMessage(req.user!.id, req.params.messageId, req.body.emoji);
    success(res, req, result);
  } catch (err) {
    next(err);
  }
});

// POST /conversations/:conversationId/read
messagingRouter.post('/:conversationId/read', authenticate, async (req: Request, res: Response, next: NextFunction) => {
  try {
    await messagingService.markRead(req.user!.id, req.params.conversationId, req.body.messageId);
    success(res, req, { success: true });
  } catch (err) {
    next(err);
  }
});

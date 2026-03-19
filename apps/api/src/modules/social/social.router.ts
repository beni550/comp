import { Router, Request, Response, NextFunction } from 'express';
import { socialService } from './social.service';
import { authenticate } from '../../middleware/auth';
import { success, paginated } from '../../common/response';
import { createCircleSchema, updateCircleSchema } from '@vybe/validation';

export const socialRouter = Router();

// ─── Follow Endpoints ───────────────────────────────────────

// POST /social/follow/:userId
socialRouter.post('/follow/:userId', authenticate, async (req: Request, res: Response, next: NextFunction) => {
  try {
    const edge = await socialService.follow(req.user!.id, req.params.userId);
    success(res, req, edge, 201);
  } catch (err) {
    next(err);
  }
});

// DELETE /social/follow/:userId
socialRouter.delete('/follow/:userId', authenticate, async (req: Request, res: Response, next: NextFunction) => {
  try {
    await socialService.unfollow(req.user!.id, req.params.userId);
    success(res, req, { success: true });
  } catch (err) {
    next(err);
  }
});

// POST /social/follow-requests/:edgeId/accept
socialRouter.post('/follow-requests/:edgeId/accept', authenticate, async (req: Request, res: Response, next: NextFunction) => {
  try {
    await socialService.acceptFollowRequest(req.user!.id, req.params.edgeId);
    success(res, req, { success: true });
  } catch (err) {
    next(err);
  }
});

// POST /social/follow-requests/:edgeId/decline
socialRouter.post('/follow-requests/:edgeId/decline', authenticate, async (req: Request, res: Response, next: NextFunction) => {
  try {
    await socialService.declineFollowRequest(req.user!.id, req.params.edgeId);
    success(res, req, { success: true });
  } catch (err) {
    next(err);
  }
});

// GET /social/followers/:userId
socialRouter.get('/followers/:userId', authenticate, async (req: Request, res: Response, next: NextFunction) => {
  try {
    const result = await socialService.getFollowers(
      req.params.userId,
      req.query.cursor as string,
      parseInt(req.query.limit as string) || 20
    );
    paginated(res, req, result.items, { cursor: result.cursor, hasMore: result.hasMore });
  } catch (err) {
    next(err);
  }
});

// GET /social/following/:userId
socialRouter.get('/following/:userId', authenticate, async (req: Request, res: Response, next: NextFunction) => {
  try {
    const result = await socialService.getFollowing(
      req.params.userId,
      req.query.cursor as string,
      parseInt(req.query.limit as string) || 20
    );
    paginated(res, req, result.items, { cursor: result.cursor, hasMore: result.hasMore });
  } catch (err) {
    next(err);
  }
});

// GET /social/follow-requests
socialRouter.get('/follow-requests', authenticate, async (req: Request, res: Response, next: NextFunction) => {
  try {
    const requests = await socialService.getFollowRequests(req.user!.id);
    success(res, req, requests);
  } catch (err) {
    next(err);
  }
});

// ─── Friend Endpoints ───────────────────────────────────────

// POST /social/friends/request/:userId
socialRouter.post('/friends/request/:userId', authenticate, async (req: Request, res: Response, next: NextFunction) => {
  try {
    const edge = await socialService.sendFriendRequest(req.user!.id, req.params.userId);
    success(res, req, edge, 201);
  } catch (err) {
    next(err);
  }
});

// POST /social/friends/:edgeId/accept
socialRouter.post('/friends/:edgeId/accept', authenticate, async (req: Request, res: Response, next: NextFunction) => {
  try {
    await socialService.acceptFriendRequest(req.user!.id, req.params.edgeId);
    success(res, req, { success: true });
  } catch (err) {
    next(err);
  }
});

// POST /social/friends/:edgeId/decline
socialRouter.post('/friends/:edgeId/decline', authenticate, async (req: Request, res: Response, next: NextFunction) => {
  try {
    await socialService.declineFriendRequest(req.user!.id, req.params.edgeId);
    success(res, req, { success: true });
  } catch (err) {
    next(err);
  }
});

// DELETE /social/friends/:userId
socialRouter.delete('/friends/:userId', authenticate, async (req: Request, res: Response, next: NextFunction) => {
  try {
    await socialService.removeFriend(req.user!.id, req.params.userId);
    success(res, req, { success: true });
  } catch (err) {
    next(err);
  }
});

// GET /social/friends
socialRouter.get('/friends', authenticate, async (req: Request, res: Response, next: NextFunction) => {
  try {
    const result = await socialService.getFriends(
      req.user!.id,
      req.query.cursor as string,
      parseInt(req.query.limit as string) || 20
    );
    paginated(res, req, result.items, { cursor: result.cursor, hasMore: result.hasMore });
  } catch (err) {
    next(err);
  }
});

// ─── Block & Mute Endpoints ─────────────────────────────────

// POST /social/block/:userId
socialRouter.post('/block/:userId', authenticate, async (req: Request, res: Response, next: NextFunction) => {
  try {
    const edge = await socialService.blockUser(req.user!.id, req.params.userId);
    success(res, req, edge, 201);
  } catch (err) {
    next(err);
  }
});

// DELETE /social/block/:userId
socialRouter.delete('/block/:userId', authenticate, async (req: Request, res: Response, next: NextFunction) => {
  try {
    await socialService.unblockUser(req.user!.id, req.params.userId);
    success(res, req, { success: true });
  } catch (err) {
    next(err);
  }
});

// POST /social/mute/:userId
socialRouter.post('/mute/:userId', authenticate, async (req: Request, res: Response, next: NextFunction) => {
  try {
    const edge = await socialService.muteUser(req.user!.id, req.params.userId, req.body.targetType);
    success(res, req, edge, 201);
  } catch (err) {
    next(err);
  }
});

// DELETE /social/mute/:userId
socialRouter.delete('/mute/:userId', authenticate, async (req: Request, res: Response, next: NextFunction) => {
  try {
    await socialService.unmuteUser(req.user!.id, req.params.userId, req.body.targetType);
    success(res, req, { success: true });
  } catch (err) {
    next(err);
  }
});

// ─── Circle Endpoints ───────────────────────────────────────

// POST /circles
socialRouter.post('/', authenticate, async (req: Request, res: Response, next: NextFunction) => {
  try {
    const data = createCircleSchema.parse(req.body);
    const circle = await socialService.createCircle(req.user!.id, data);
    success(res, req, circle, 201);
  } catch (err) {
    next(err);
  }
});

// GET /circles
socialRouter.get('/my', authenticate, async (req: Request, res: Response, next: NextFunction) => {
  try {
    const circles = await socialService.getUserCircles(req.user!.id);
    success(res, req, circles);
  } catch (err) {
    next(err);
  }
});

// GET /circles/:circleId
socialRouter.get('/:circleId', authenticate, async (req: Request, res: Response, next: NextFunction) => {
  try {
    const circle = await socialService.getCircle(req.params.circleId, req.user!.id);
    success(res, req, circle);
  } catch (err) {
    next(err);
  }
});

// PATCH /circles/:circleId
socialRouter.patch('/:circleId', authenticate, async (req: Request, res: Response, next: NextFunction) => {
  try {
    const data = updateCircleSchema.parse(req.body);
    const circle = await socialService.updateCircle(req.params.circleId, req.user!.id, data);
    success(res, req, circle);
  } catch (err) {
    next(err);
  }
});

// DELETE /circles/:circleId
socialRouter.delete('/:circleId', authenticate, async (req: Request, res: Response, next: NextFunction) => {
  try {
    await socialService.deleteCircle(req.params.circleId, req.user!.id);
    success(res, req, { success: true });
  } catch (err) {
    next(err);
  }
});

// POST /circles/:circleId/members
socialRouter.post('/:circleId/members', authenticate, async (req: Request, res: Response, next: NextFunction) => {
  try {
    const { userId } = req.body;
    const member = await socialService.addCircleMember(req.params.circleId, req.user!.id, userId);
    success(res, req, member, 201);
  } catch (err) {
    next(err);
  }
});

// DELETE /circles/:circleId/members/:userId
socialRouter.delete('/:circleId/members/:userId', authenticate, async (req: Request, res: Response, next: NextFunction) => {
  try {
    await socialService.removeCircleMember(req.params.circleId, req.user!.id, req.params.userId);
    success(res, req, { success: true });
  } catch (err) {
    next(err);
  }
});

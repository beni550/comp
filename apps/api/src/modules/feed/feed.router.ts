import { Router, Request, Response, NextFunction } from 'express';
import { feedService } from './feed.service';
import { authenticate, optionalAuth } from '../../middleware/auth';
import { paginated } from '../../common/response';

export const feedRouter = Router();

// GET /feed/my-vybe
feedRouter.get('/my-vybe', authenticate, async (req: Request, res: Response, next: NextFunction) => {
  try {
    const result = await feedService.getMyVybeFeed(
      req.user!.id,
      req.query.cursor as string,
      parseInt(req.query.limit as string) || 20
    );
    paginated(res, req, result.items, { cursor: result.cursor, hasMore: result.hasMore });
  } catch (err) {
    next(err);
  }
});

// GET /feed/for-you
feedRouter.get('/for-you', authenticate, async (req: Request, res: Response, next: NextFunction) => {
  try {
    const result = await feedService.getForYouFeed(
      req.user!.id,
      req.query.cursor as string,
      parseInt(req.query.limit as string) || 20
    );
    paginated(res, req, result.items, { cursor: result.cursor, hasMore: result.hasMore });
  } catch (err) {
    next(err);
  }
});

// GET /feed/profile/:userId
feedRouter.get('/profile/:userId', optionalAuth, async (req: Request, res: Response, next: NextFunction) => {
  try {
    const result = await feedService.getProfileFeed(
      req.params.userId,
      req.user?.id,
      req.query.cursor as string,
      parseInt(req.query.limit as string) || 20
    );
    paginated(res, req, result.items, { cursor: result.cursor, hasMore: result.hasMore });
  } catch (err) {
    next(err);
  }
});

// GET /discover/trending
feedRouter.get('/trending', optionalAuth, async (req: Request, res: Response, next: NextFunction) => {
  try {
    const result = await feedService.getTrending(
      req.query.timeframe as string || '24h',
      req.query.cursor as string,
      parseInt(req.query.limit as string) || 20
    );
    paginated(res, req, result.items, { cursor: result.cursor, hasMore: result.hasMore });
  } catch (err) {
    next(err);
  }
});

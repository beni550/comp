import { Router, Request, Response, NextFunction } from 'express';
import { contentService } from './content.service';
import { authenticate, optionalAuth } from '../../middleware/auth';
import { success, paginated } from '../../common/response';
import {
  uploadIntentSchema, createContentSchema, updateContentSchema,
  commentBodySchema, createStorySchema, createHighlightSchema,
} from '@vybe/validation';

export const contentRouter = Router();

// ─── Media Upload ───────────────────────────────────────────

// POST /media/upload-intent
contentRouter.post('/upload-intent', authenticate, async (req: Request, res: Response, next: NextFunction) => {
  try {
    const data = uploadIntentSchema.parse(req.body);
    const result = await contentService.createUploadIntent(req.user!.id, data);
    success(res, req, result, 201);
  } catch (err) {
    next(err);
  }
});

// POST /media/:assetId/complete
contentRouter.post('/:assetId/complete', authenticate, async (req: Request, res: Response, next: NextFunction) => {
  try {
    const asset = await contentService.completeUpload(req.user!.id, req.params.assetId, req.body);
    success(res, req, asset);
  } catch (err) {
    next(err);
  }
});

// ─── Content CRUD ───────────────────────────────────────────

// POST /content
contentRouter.post('/', authenticate, async (req: Request, res: Response, next: NextFunction) => {
  try {
    const data = createContentSchema.parse(req.body);
    const content = await contentService.createContent(req.user!.id, data);
    success(res, req, content, 201);
  } catch (err) {
    next(err);
  }
});

// GET /content/:contentId
contentRouter.get('/:contentId', optionalAuth, async (req: Request, res: Response, next: NextFunction) => {
  try {
    const content = await contentService.getContentById(req.params.contentId, req.user?.id);
    success(res, req, content);
  } catch (err) {
    next(err);
  }
});

// PATCH /content/:contentId
contentRouter.patch('/:contentId', authenticate, async (req: Request, res: Response, next: NextFunction) => {
  try {
    const data = updateContentSchema.parse(req.body);
    const content = await contentService.updateContent(req.user!.id, req.params.contentId, data);
    success(res, req, content);
  } catch (err) {
    next(err);
  }
});

// POST /content/:contentId/publish
contentRouter.post('/:contentId/publish', authenticate, async (req: Request, res: Response, next: NextFunction) => {
  try {
    const content = await contentService.publishContent(req.user!.id, req.params.contentId);
    success(res, req, content);
  } catch (err) {
    next(err);
  }
});

// DELETE /content/:contentId
contentRouter.delete('/:contentId', authenticate, async (req: Request, res: Response, next: NextFunction) => {
  try {
    await contentService.deleteContent(req.user!.id, req.params.contentId);
    success(res, req, { success: true });
  } catch (err) {
    next(err);
  }
});

// ─── Comments ───────────────────────────────────────────────

// POST /content/:contentId/comments
contentRouter.post('/:contentId/comments', authenticate, async (req: Request, res: Response, next: NextFunction) => {
  try {
    const body = commentBodySchema.parse(req.body.body);
    const comment = await contentService.createComment(
      req.user!.id,
      req.params.contentId,
      body,
      req.body.parentId
    );
    success(res, req, comment, 201);
  } catch (err) {
    next(err);
  }
});

// GET /content/:contentId/comments
contentRouter.get('/:contentId/comments', optionalAuth, async (req: Request, res: Response, next: NextFunction) => {
  try {
    const result = await contentService.getComments(
      req.params.contentId,
      req.query.cursor as string,
      parseInt(req.query.limit as string) || 20
    );
    paginated(res, req, result.items, { cursor: result.cursor, hasMore: result.hasMore });
  } catch (err) {
    next(err);
  }
});

// DELETE /content/comments/:commentId
contentRouter.delete('/comments/:commentId', authenticate, async (req: Request, res: Response, next: NextFunction) => {
  try {
    await contentService.deleteComment(req.user!.id, req.params.commentId);
    success(res, req, { success: true });
  } catch (err) {
    next(err);
  }
});

// ─── Reactions ──────────────────────────────────────────────

// POST /content/:contentId/react
contentRouter.post('/:contentId/react', authenticate, async (req: Request, res: Response, next: NextFunction) => {
  try {
    const result = await contentService.toggleReaction(
      req.user!.id,
      'content',
      req.params.contentId,
      req.body.type || 'like'
    );
    success(res, req, result);
  } catch (err) {
    next(err);
  }
});

// POST /content/comments/:commentId/react
contentRouter.post('/comments/:commentId/react', authenticate, async (req: Request, res: Response, next: NextFunction) => {
  try {
    const result = await contentService.toggleReaction(
      req.user!.id,
      'comment',
      req.params.commentId,
      req.body.type || 'like'
    );
    success(res, req, result);
  } catch (err) {
    next(err);
  }
});

// ─── Save Items ─────────────────────────────────────────────

// POST /content/:contentId/save
contentRouter.post('/:contentId/save', authenticate, async (req: Request, res: Response, next: NextFunction) => {
  try {
    const result = await contentService.toggleSave(req.user!.id, req.params.contentId);
    success(res, req, result);
  } catch (err) {
    next(err);
  }
});

// GET /content/saved
contentRouter.get('/saved/list', authenticate, async (req: Request, res: Response, next: NextFunction) => {
  try {
    const result = await contentService.getSavedContent(
      req.user!.id,
      req.query.cursor as string,
      parseInt(req.query.limit as string) || 20
    );
    paginated(res, req, result.items, { cursor: result.cursor, hasMore: result.hasMore });
  } catch (err) {
    next(err);
  }
});

// ─── Stories ────────────────────────────────────────────────

// POST /stories
contentRouter.post('/stories', authenticate, async (req: Request, res: Response, next: NextFunction) => {
  try {
    const data = createStorySchema.parse(req.body);
    const story = await contentService.createStory(req.user!.id, data);
    success(res, req, story, 201);
  } catch (err) {
    next(err);
  }
});

// GET /stories
contentRouter.get('/stories', authenticate, async (req: Request, res: Response, next: NextFunction) => {
  try {
    const groups = await contentService.getStories(req.user!.id);
    success(res, req, groups);
  } catch (err) {
    next(err);
  }
});

// ─── Highlights ─────────────────────────────────────────────

// POST /highlights
contentRouter.post('/highlights', authenticate, async (req: Request, res: Response, next: NextFunction) => {
  try {
    const data = createHighlightSchema.parse(req.body);
    const highlight = await contentService.createHighlight(req.user!.id, data);
    success(res, req, highlight, 201);
  } catch (err) {
    next(err);
  }
});

// GET /highlights/:userId
contentRouter.get('/highlights/:userId', optionalAuth, async (req: Request, res: Response, next: NextFunction) => {
  try {
    const highlights = await contentService.getHighlights(req.params.userId);
    success(res, req, highlights);
  } catch (err) {
    next(err);
  }
});

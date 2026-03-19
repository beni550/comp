import { Response, Request } from 'express';

export function success<T>(res: Response, req: Request, data: T, statusCode = 200): void {
  res.status(statusCode).json({
    data,
    meta: {
      requestId: req.requestId,
      timestamp: new Date().toISOString(),
    },
  });
}

export function paginated<T>(
  res: Response,
  req: Request,
  data: T[],
  opts: { cursor?: string; hasMore: boolean; total?: number }
): void {
  res.status(200).json({
    data,
    meta: {
      requestId: req.requestId,
      cursor: opts.cursor,
      hasMore: opts.hasMore,
      total: opts.total,
    },
  });
}

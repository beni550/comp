export interface ApiResponse<T> {
  data: T;
  meta: {
    requestId: string;
    timestamp: string;
  };
}

export interface ApiError {
  error: {
    code: string;
    message: string;
    details?: Record<string, unknown>;
    requestId: string;
  };
}

export interface PaginatedResponse<T> {
  data: T[];
  meta: {
    requestId: string;
    cursor?: string;
    hasMore: boolean;
    total?: number;
  };
}

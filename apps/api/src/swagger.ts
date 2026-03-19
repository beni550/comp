import swaggerJsdoc from 'swagger-jsdoc';

const options: swaggerJsdoc.Options = {
  definition: {
    openapi: '3.0.3',
    info: {
      title: 'VYBE API',
      version: '1.0.0',
      description: 'VYBE Social Platform MVP API Documentation',
    },
    servers: [
      { url: '/api/v1', description: 'Local development' },
    ],
    components: {
      securitySchemes: {
        bearerAuth: {
          type: 'http',
          scheme: 'bearer',
          bearerFormat: 'JWT',
        },
      },
      schemas: {
        Error: {
          type: 'object',
          properties: {
            error: {
              type: 'object',
              properties: {
                code: { type: 'string' },
                message: { type: 'string' },
              },
            },
          },
        },
        User: {
          type: 'object',
          properties: {
            id: { type: 'string', format: 'uuid' },
            username: { type: 'string' },
            displayName: { type: 'string' },
            email: { type: 'string', format: 'email' },
            avatarUrl: { type: 'string', nullable: true },
            bio: { type: 'string', nullable: true },
            role: { type: 'string', enum: ['user', 'moderator', 'admin'] },
            status: { type: 'string' },
          },
        },
        ContentItem: {
          type: 'object',
          properties: {
            id: { type: 'string', format: 'uuid' },
            authorId: { type: 'string', format: 'uuid' },
            author: { $ref: '#/components/schemas/User' },
            type: { type: 'string', enum: ['video', 'photo', 'carousel', 'text_media_thread'] },
            status: { type: 'string', enum: ['draft', 'published', 'removed'] },
            caption: { type: 'string' },
            media: { type: 'array', items: { $ref: '#/components/schemas/Media' } },
            likesCount: { type: 'integer' },
            commentsCount: { type: 'integer' },
            sharesCount: { type: 'integer' },
            isLiked: { type: 'boolean' },
            isSaved: { type: 'boolean' },
            publishedAt: { type: 'string', format: 'date-time', nullable: true },
            createdAt: { type: 'string', format: 'date-time' },
          },
        },
        Media: {
          type: 'object',
          properties: {
            id: { type: 'string', format: 'uuid' },
            kind: { type: 'string', enum: ['image', 'video', 'audio'] },
            mimeType: { type: 'string' },
            url: { type: 'string' },
            width: { type: 'integer', nullable: true },
            height: { type: 'integer', nullable: true },
          },
        },
        Comment: {
          type: 'object',
          properties: {
            id: { type: 'string', format: 'uuid' },
            contentId: { type: 'string', format: 'uuid' },
            authorId: { type: 'string', format: 'uuid' },
            author: { $ref: '#/components/schemas/User' },
            body: { type: 'string' },
            likesCount: { type: 'integer' },
            createdAt: { type: 'string', format: 'date-time' },
          },
        },
        Conversation: {
          type: 'object',
          properties: {
            id: { type: 'string', format: 'uuid' },
            type: { type: 'string', enum: ['direct', 'circle_group'] },
            participants: { type: 'array', items: { $ref: '#/components/schemas/User' } },
            lastMessage: { type: 'object', nullable: true },
            unreadCount: { type: 'integer' },
          },
        },
        Message: {
          type: 'object',
          properties: {
            id: { type: 'string', format: 'uuid' },
            senderId: { type: 'string', format: 'uuid' },
            body: { type: 'string' },
            createdAt: { type: 'string', format: 'date-time' },
            status: { type: 'string' },
          },
        },
        Notification: {
          type: 'object',
          properties: {
            id: { type: 'string', format: 'uuid' },
            type: { type: 'string' },
            title: { type: 'string' },
            body: { type: 'string' },
            isRead: { type: 'boolean' },
            createdAt: { type: 'string', format: 'date-time' },
          },
        },
      },
    },
    paths: {
      // ─── Auth ────────────────────────────────────────────
      '/auth/register/email/start': {
        post: {
          tags: ['Auth'],
          summary: 'Start email registration (sends OTP)',
          requestBody: { content: { 'application/json': { schema: { type: 'object', properties: { email: { type: 'string', format: 'email' } }, required: ['email'] } } } },
          responses: { '200': { description: 'OTP sent' }, '429': { description: 'Rate limited' } },
        },
      },
      '/auth/register/phone/start': {
        post: {
          tags: ['Auth'],
          summary: 'Start phone registration (sends OTP)',
          requestBody: { content: { 'application/json': { schema: { type: 'object', properties: { phone: { type: 'string' } }, required: ['phone'] } } } },
          responses: { '200': { description: 'OTP sent' }, '429': { description: 'Rate limited' } },
        },
      },
      '/auth/register/phone/verify': {
        post: {
          tags: ['Auth'],
          summary: 'Verify phone OTP and create account',
          requestBody: { content: { 'application/json': { schema: { type: 'object', properties: { challengeId: { type: 'string' }, code: { type: 'string' }, username: { type: 'string' }, displayName: { type: 'string' }, birthDate: { type: 'string' } }, required: ['challengeId', 'code', 'username', 'displayName', 'birthDate'] } } } },
          responses: { '201': { description: 'Account created with tokens' } },
        },
      },
      '/auth/login': {
        post: {
          tags: ['Auth'],
          summary: 'Login with email/phone',
          requestBody: { content: { 'application/json': { schema: { type: 'object', properties: { method: { type: 'string', enum: ['email', 'phone'] }, identifier: { type: 'string' }, password: { type: 'string' }, code: { type: 'string' }, deviceInfo: { type: 'object', properties: { platform: { type: 'string', enum: ['ios', 'android', 'web'] } } } }, required: ['method', 'deviceInfo'] } } } },
          responses: { '200': { description: 'Login successful with tokens and user data' }, '401': { description: 'Invalid credentials' }, '429': { description: 'Rate limited' } },
        },
      },
      '/auth/refresh': {
        post: {
          tags: ['Auth'],
          summary: 'Refresh access token',
          requestBody: { content: { 'application/json': { schema: { type: 'object', properties: { refreshToken: { type: 'string' } }, required: ['refreshToken'] } } } },
          responses: { '200': { description: 'New token pair' }, '401': { description: 'Invalid refresh token' } },
        },
      },
      '/auth/logout': {
        post: {
          tags: ['Auth'],
          summary: 'Logout current session',
          security: [{ bearerAuth: [] }],
          responses: { '200': { description: 'Logged out' } },
        },
      },
      '/auth/logout-all': {
        post: {
          tags: ['Auth'],
          summary: 'Logout all sessions',
          security: [{ bearerAuth: [] }],
          responses: { '200': { description: 'All sessions invalidated' } },
        },
      },
      '/auth/sessions': {
        get: {
          tags: ['Auth'],
          summary: 'List active sessions',
          security: [{ bearerAuth: [] }],
          responses: { '200': { description: 'List of active sessions' } },
        },
      },
      '/auth/username/check/{username}': {
        get: {
          tags: ['Auth'],
          summary: 'Check if username is available',
          parameters: [{ in: 'path', name: 'username', required: true, schema: { type: 'string' } }],
          responses: { '200': { description: 'Availability result' } },
        },
      },

      // ─── Users ───────────────────────────────────────────
      '/users/me': {
        get: {
          tags: ['Users'],
          summary: 'Get current user profile',
          security: [{ bearerAuth: [] }],
          responses: { '200': { description: 'User profile', content: { 'application/json': { schema: { $ref: '#/components/schemas/User' } } } } },
        },
        patch: {
          tags: ['Users'],
          summary: 'Update current user',
          security: [{ bearerAuth: [] }],
          requestBody: { content: { 'application/json': { schema: { type: 'object', properties: { displayName: { type: 'string' }, bio: { type: 'string' } } } } } },
          responses: { '200': { description: 'Updated user' } },
        },
        delete: {
          tags: ['Users'],
          summary: 'Delete account',
          security: [{ bearerAuth: [] }],
          responses: { '200': { description: 'Account deleted' } },
        },
      },
      '/users/{userId}': {
        get: {
          tags: ['Users'],
          summary: 'Get user by ID',
          parameters: [{ in: 'path', name: 'userId', required: true, schema: { type: 'string' } }],
          responses: { '200': { description: 'User profile' } },
        },
      },
      '/users/me/profile': {
        patch: {
          tags: ['Users'],
          summary: 'Update user profile details',
          security: [{ bearerAuth: [] }],
          requestBody: { content: { 'application/json': { schema: { type: 'object', properties: { bio: { type: 'string' }, avatarMediaId: { type: 'string' }, theme: { type: 'string' } } } } } },
          responses: { '200': { description: 'Updated profile' } },
        },
      },
      '/users/me/settings': {
        get: {
          tags: ['Users'],
          summary: 'Get user settings',
          security: [{ bearerAuth: [] }],
          responses: { '200': { description: 'User settings' } },
        },
        patch: {
          tags: ['Users'],
          summary: 'Update user settings',
          security: [{ bearerAuth: [] }],
          requestBody: { content: { 'application/json': { schema: { type: 'object' } } } },
          responses: { '200': { description: 'Updated settings' } },
        },
      },
      '/users/search': {
        get: {
          tags: ['Users'],
          summary: 'Search users',
          security: [{ bearerAuth: [] }],
          parameters: [{ in: 'query', name: 'q', required: true, schema: { type: 'string' } }],
          responses: { '200': { description: 'Search results' } },
        },
      },
      '/users/me/onboarding': {
        get: {
          tags: ['Users'],
          summary: 'Get onboarding status',
          security: [{ bearerAuth: [] }],
          responses: { '200': { description: 'Onboarding progress' } },
        },
      },
      '/users/me/onboarding/complete': {
        post: {
          tags: ['Users'],
          summary: 'Mark onboarding as complete',
          security: [{ bearerAuth: [] }],
          responses: { '200': { description: 'Onboarding completed' } },
        },
      },

      // ─── Social ──────────────────────────────────────────
      '/social/follow/{userId}': {
        post: {
          tags: ['Social'],
          summary: 'Follow a user',
          security: [{ bearerAuth: [] }],
          parameters: [{ in: 'path', name: 'userId', required: true, schema: { type: 'string' } }],
          responses: { '200': { description: 'Follow created' } },
        },
        delete: {
          tags: ['Social'],
          summary: 'Unfollow a user',
          security: [{ bearerAuth: [] }],
          parameters: [{ in: 'path', name: 'userId', required: true, schema: { type: 'string' } }],
          responses: { '200': { description: 'Unfollowed' } },
        },
      },
      '/social/followers/{userId}': {
        get: {
          tags: ['Social'],
          summary: 'Get followers list',
          security: [{ bearerAuth: [] }],
          parameters: [{ in: 'path', name: 'userId', required: true, schema: { type: 'string' } }],
          responses: { '200': { description: 'Followers list' } },
        },
      },
      '/social/following/{userId}': {
        get: {
          tags: ['Social'],
          summary: 'Get following list',
          security: [{ bearerAuth: [] }],
          parameters: [{ in: 'path', name: 'userId', required: true, schema: { type: 'string' } }],
          responses: { '200': { description: 'Following list' } },
        },
      },
      '/social/block/{userId}': {
        post: {
          tags: ['Social'],
          summary: 'Block a user',
          security: [{ bearerAuth: [] }],
          parameters: [{ in: 'path', name: 'userId', required: true, schema: { type: 'string' } }],
          responses: { '200': { description: 'User blocked' } },
        },
        delete: {
          tags: ['Social'],
          summary: 'Unblock a user',
          security: [{ bearerAuth: [] }],
          parameters: [{ in: 'path', name: 'userId', required: true, schema: { type: 'string' } }],
          responses: { '200': { description: 'User unblocked' } },
        },
      },

      // ─── Content ─────────────────────────────────────────
      '/content': {
        post: {
          tags: ['Content'],
          summary: 'Create new content',
          security: [{ bearerAuth: [] }],
          requestBody: { content: { 'application/json': { schema: { type: 'object', properties: { type: { type: 'string', enum: ['video', 'photo', 'carousel', 'text_media_thread'] }, caption: { type: 'string' }, audience: { type: 'string', enum: ['public', 'followers', 'friends', 'circle', 'private'] }, mediaIds: { type: 'array', items: { type: 'string' } }, publishNow: { type: 'boolean' } }, required: ['type', 'audience', 'mediaIds'] } } } },
          responses: { '201': { description: 'Content created', content: { 'application/json': { schema: { $ref: '#/components/schemas/ContentItem' } } } } },
        },
      },
      '/content/{contentId}': {
        get: {
          tags: ['Content'],
          summary: 'Get content by ID',
          parameters: [{ in: 'path', name: 'contentId', required: true, schema: { type: 'string' } }],
          responses: { '200': { description: 'Content item', content: { 'application/json': { schema: { $ref: '#/components/schemas/ContentItem' } } } } },
        },
        patch: {
          tags: ['Content'],
          summary: 'Update content',
          security: [{ bearerAuth: [] }],
          parameters: [{ in: 'path', name: 'contentId', required: true, schema: { type: 'string' } }],
          requestBody: { content: { 'application/json': { schema: { type: 'object', properties: { caption: { type: 'string' }, audience: { type: 'string' } } } } } },
          responses: { '200': { description: 'Updated content' } },
        },
        delete: {
          tags: ['Content'],
          summary: 'Delete content',
          security: [{ bearerAuth: [] }],
          parameters: [{ in: 'path', name: 'contentId', required: true, schema: { type: 'string' } }],
          responses: { '200': { description: 'Content deleted' } },
        },
      },
      '/content/{contentId}/publish': {
        post: {
          tags: ['Content'],
          summary: 'Publish draft content',
          security: [{ bearerAuth: [] }],
          parameters: [{ in: 'path', name: 'contentId', required: true, schema: { type: 'string' } }],
          responses: { '200': { description: 'Content published' } },
        },
      },
      '/content/{contentId}/comments': {
        get: {
          tags: ['Content'],
          summary: 'Get comments for content',
          parameters: [{ in: 'path', name: 'contentId', required: true, schema: { type: 'string' } }, { in: 'query', name: 'cursor', schema: { type: 'string' } }, { in: 'query', name: 'limit', schema: { type: 'integer' } }],
          responses: { '200': { description: 'Comments list' } },
        },
        post: {
          tags: ['Content'],
          summary: 'Add a comment',
          security: [{ bearerAuth: [] }],
          parameters: [{ in: 'path', name: 'contentId', required: true, schema: { type: 'string' } }],
          requestBody: { content: { 'application/json': { schema: { type: 'object', properties: { body: { type: 'string' }, parentId: { type: 'string' } }, required: ['body'] } } } },
          responses: { '201': { description: 'Comment created', content: { 'application/json': { schema: { $ref: '#/components/schemas/Comment' } } } } },
        },
      },
      '/content/{contentId}/react': {
        post: {
          tags: ['Content'],
          summary: 'Toggle reaction on content',
          security: [{ bearerAuth: [] }],
          parameters: [{ in: 'path', name: 'contentId', required: true, schema: { type: 'string' } }],
          requestBody: { content: { 'application/json': { schema: { type: 'object', properties: { type: { type: 'string', default: 'like' } } } } } },
          responses: { '200': { description: 'Reaction toggled (action: added/removed)' } },
        },
      },
      '/content/{contentId}/save': {
        post: {
          tags: ['Content'],
          summary: 'Toggle save on content',
          security: [{ bearerAuth: [] }],
          parameters: [{ in: 'path', name: 'contentId', required: true, schema: { type: 'string' } }],
          responses: { '200': { description: 'Save toggled' } },
        },
      },

      // ─── Media Upload ────────────────────────────────────
      '/media/upload-intent': {
        post: {
          tags: ['Media'],
          summary: 'Create upload intent (get presigned URL)',
          security: [{ bearerAuth: [] }],
          requestBody: { content: { 'application/json': { schema: { type: 'object', properties: { filename: { type: 'string' }, mimeType: { type: 'string' }, sizeBytes: { type: 'integer' } }, required: ['filename', 'mimeType', 'sizeBytes'] } } } },
          responses: { '201': { description: 'Upload URL and asset ID' } },
        },
      },
      '/media/{assetId}/complete': {
        post: {
          tags: ['Media'],
          summary: 'Mark media upload as complete',
          security: [{ bearerAuth: [] }],
          parameters: [{ in: 'path', name: 'assetId', required: true, schema: { type: 'string' } }],
          responses: { '200': { description: 'Media marked ready' } },
        },
      },

      // ─── Feed ────────────────────────────────────────────
      '/feed/my-vybe': {
        get: {
          tags: ['Feed'],
          summary: 'Get personalized feed (followed users)',
          security: [{ bearerAuth: [] }],
          parameters: [{ in: 'query', name: 'cursor', schema: { type: 'string' } }, { in: 'query', name: 'limit', schema: { type: 'integer' } }],
          responses: { '200': { description: 'Feed items' } },
        },
      },
      '/feed/for-you': {
        get: {
          tags: ['Feed'],
          summary: 'Get discovery feed',
          security: [{ bearerAuth: [] }],
          parameters: [{ in: 'query', name: 'cursor', schema: { type: 'string' } }, { in: 'query', name: 'limit', schema: { type: 'integer' } }],
          responses: { '200': { description: 'Feed items' } },
        },
      },
      '/feed/trending': {
        get: {
          tags: ['Feed'],
          summary: 'Get trending content',
          responses: { '200': { description: 'Trending items' } },
        },
      },

      // ─── Messaging ───────────────────────────────────────
      '/conversations': {
        get: {
          tags: ['Messaging'],
          summary: 'List conversations',
          security: [{ bearerAuth: [] }],
          responses: { '200': { description: 'Conversations list' } },
        },
        post: {
          tags: ['Messaging'],
          summary: 'Create new conversation',
          security: [{ bearerAuth: [] }],
          requestBody: { content: { 'application/json': { schema: { type: 'object', properties: { type: { type: 'string', enum: ['direct', 'circle_group'] }, participantIds: { type: 'array', items: { type: 'string' } } }, required: ['type'] } } } },
          responses: { '201': { description: 'Conversation created' } },
        },
      },
      '/conversations/{conversationId}/messages': {
        get: {
          tags: ['Messaging'],
          summary: 'Get messages in conversation',
          security: [{ bearerAuth: [] }],
          parameters: [{ in: 'path', name: 'conversationId', required: true, schema: { type: 'string' } }],
          responses: { '200': { description: 'Messages list' } },
        },
        post: {
          tags: ['Messaging'],
          summary: 'Send a message',
          security: [{ bearerAuth: [] }],
          parameters: [{ in: 'path', name: 'conversationId', required: true, schema: { type: 'string' } }],
          requestBody: { content: { 'application/json': { schema: { type: 'object', properties: { kind: { type: 'string', enum: ['text', 'image', 'video', 'reply'] }, body: { type: 'string' } }, required: ['kind'] } } } },
          responses: { '201': { description: 'Message sent' } },
        },
      },
      '/conversations/{conversationId}/read': {
        post: {
          tags: ['Messaging'],
          summary: 'Mark conversation as read',
          security: [{ bearerAuth: [] }],
          parameters: [{ in: 'path', name: 'conversationId', required: true, schema: { type: 'string' } }],
          responses: { '200': { description: 'Marked as read' } },
        },
      },

      // ─── Notifications ───────────────────────────────────
      '/notifications': {
        get: {
          tags: ['Notifications'],
          summary: 'List notifications',
          security: [{ bearerAuth: [] }],
          parameters: [{ in: 'query', name: 'cursor', schema: { type: 'string' } }, { in: 'query', name: 'limit', schema: { type: 'integer' } }],
          responses: { '200': { description: 'Notifications list' } },
        },
      },
      '/notifications/unread-count': {
        get: {
          tags: ['Notifications'],
          summary: 'Get unread notification count',
          security: [{ bearerAuth: [] }],
          responses: { '200': { description: 'Unread count' } },
        },
      },
      '/notifications/{notificationId}/read': {
        post: {
          tags: ['Notifications'],
          summary: 'Mark notification as read',
          security: [{ bearerAuth: [] }],
          parameters: [{ in: 'path', name: 'notificationId', required: true, schema: { type: 'string' } }],
          responses: { '200': { description: 'Marked as read' } },
        },
      },
      '/notifications/read-all': {
        post: {
          tags: ['Notifications'],
          summary: 'Mark all notifications as read',
          security: [{ bearerAuth: [] }],
          responses: { '200': { description: 'All marked as read' } },
        },
      },

      // ─── Reports (Moderation) ────────────────────────────
      '/reports': {
        post: {
          tags: ['Moderation'],
          summary: 'Submit a report',
          security: [{ bearerAuth: [] }],
          requestBody: { content: { 'application/json': { schema: { type: 'object', properties: { targetType: { type: 'string', enum: ['user', 'content', 'comment', 'message', 'circle'] }, targetId: { type: 'string' }, reason: { type: 'string', enum: ['spam', 'harassment', 'hate_speech', 'violence', 'nudity', 'misinformation', 'other'] }, description: { type: 'string' } }, required: ['targetType', 'targetId', 'reason'] } } } },
          responses: { '201': { description: 'Report submitted' } },
        },
      },
      '/reports/mine': {
        get: {
          tags: ['Moderation'],
          summary: 'Get my submitted reports',
          security: [{ bearerAuth: [] }],
          responses: { '200': { description: 'My reports' } },
        },
      },

      // ─── Admin ───────────────────────────────────────────
      '/admin/dashboard': {
        get: {
          tags: ['Admin'],
          summary: 'Get admin dashboard stats',
          security: [{ bearerAuth: [] }],
          responses: { '200': { description: 'Dashboard statistics' } },
        },
      },
      '/admin/users/lookup': {
        get: {
          tags: ['Admin'],
          summary: 'Lookup users',
          security: [{ bearerAuth: [] }],
          parameters: [{ in: 'query', name: 'q', required: true, schema: { type: 'string' } }],
          responses: { '200': { description: 'User results' } },
        },
      },
      '/admin/users/{userId}/suspend': {
        post: {
          tags: ['Admin'],
          summary: 'Suspend a user',
          security: [{ bearerAuth: [] }],
          parameters: [{ in: 'path', name: 'userId', required: true, schema: { type: 'string' } }],
          responses: { '200': { description: 'User suspended' } },
        },
      },
      '/admin/users/{userId}/ban': {
        post: {
          tags: ['Admin'],
          summary: 'Ban a user',
          security: [{ bearerAuth: [] }],
          parameters: [{ in: 'path', name: 'userId', required: true, schema: { type: 'string' } }],
          responses: { '200': { description: 'User banned' } },
        },
      },
      '/admin/users/{userId}/reinstate': {
        post: {
          tags: ['Admin'],
          summary: 'Reinstate a user',
          security: [{ bearerAuth: [] }],
          parameters: [{ in: 'path', name: 'userId', required: true, schema: { type: 'string' } }],
          responses: { '200': { description: 'User reinstated' } },
        },
      },
      '/admin/moderation/queue': {
        get: {
          tags: ['Admin'],
          summary: 'Get moderation queue',
          security: [{ bearerAuth: [] }],
          responses: { '200': { description: 'Moderation queue items' } },
        },
      },
      '/admin/moderation/cases/{caseId}/resolve': {
        post: {
          tags: ['Admin'],
          summary: 'Resolve a moderation case',
          security: [{ bearerAuth: [] }],
          parameters: [{ in: 'path', name: 'caseId', required: true, schema: { type: 'string' } }],
          requestBody: { content: { 'application/json': { schema: { type: 'object', properties: { action: { type: 'string' } } } } } },
          responses: { '200': { description: 'Case resolved' } },
        },
      },
      '/admin/feature-flags': {
        get: {
          tags: ['Admin'],
          summary: 'List feature flags',
          security: [{ bearerAuth: [] }],
          responses: { '200': { description: 'Feature flags' } },
        },
        post: {
          tags: ['Admin'],
          summary: 'Create feature flag',
          security: [{ bearerAuth: [] }],
          requestBody: { content: { 'application/json': { schema: { type: 'object', properties: { key: { type: 'string' }, description: { type: 'string' }, isEnabled: { type: 'boolean' } }, required: ['key'] } } } },
          responses: { '201': { description: 'Flag created' } },
        },
      },

      // ─── Analytics ───────────────────────────────────────
      '/analytics/events': {
        post: {
          tags: ['Analytics'],
          summary: 'Track analytics event',
          requestBody: { content: { 'application/json': { schema: { type: 'object', properties: { eventType: { type: 'string' }, attributes: { type: 'object' } }, required: ['eventType', 'attributes'] } } } },
          responses: { '200': { description: 'Event tracked' } },
        },
      },
      '/analytics/summary': {
        get: {
          tags: ['Analytics'],
          summary: 'Get analytics summary (admin only)',
          security: [{ bearerAuth: [] }],
          responses: { '200': { description: 'Analytics summary' } },
        },
      },
    },
  },
  apis: [], // We define paths inline above instead of using JSDoc annotations
};

export const swaggerSpec = swaggerJsdoc(options);

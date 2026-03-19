import { describe, it, expect } from 'vitest';
import {
  usernameSchema,
  displayNameSchema,
  phoneSchema,
  emailSchema,
  birthDateSchema,
  bioSchema,
  commentBodySchema,
  registerPhoneStartSchema,
  loginSchema,
  createContentSchema,
  createReportSchema,
  sendMessageSchema,
  createStorySchema,
  trackEventSchema,
} from './index';

describe('usernameSchema', () => {
  it('accepts valid usernames', () => {
    expect(usernameSchema.safeParse('alice').success).toBe(true);
    expect(usernameSchema.safeParse('bob_smith').success).toBe(true);
    expect(usernameSchema.safeParse('user.name').success).toBe(true);
    expect(usernameSchema.safeParse('User123').success).toBe(true);
  });

  it('rejects too short usernames', () => {
    expect(usernameSchema.safeParse('ab').success).toBe(false);
  });

  it('rejects too long usernames', () => {
    expect(usernameSchema.safeParse('a'.repeat(25)).success).toBe(false);
  });

  it('rejects usernames starting with dot', () => {
    expect(usernameSchema.safeParse('.alice').success).toBe(false);
  });

  it('rejects usernames ending with dot', () => {
    expect(usernameSchema.safeParse('alice.').success).toBe(false);
  });

  it('rejects consecutive dots', () => {
    expect(usernameSchema.safeParse('ali..ce').success).toBe(false);
  });

  it('rejects special characters', () => {
    expect(usernameSchema.safeParse('alice@bob').success).toBe(false);
    expect(usernameSchema.safeParse('alice bob').success).toBe(false);
    expect(usernameSchema.safeParse('alice-bob').success).toBe(false);
  });
});

describe('phoneSchema', () => {
  it('accepts valid E.164 phone numbers', () => {
    expect(phoneSchema.safeParse('+1234567890').success).toBe(true);
    expect(phoneSchema.safeParse('+44123456789').success).toBe(true);
  });

  it('rejects invalid phone numbers', () => {
    expect(phoneSchema.safeParse('1234567890').success).toBe(false);
    expect(phoneSchema.safeParse('+0123456789').success).toBe(false);
    expect(phoneSchema.safeParse('phone').success).toBe(false);
    expect(phoneSchema.safeParse('').success).toBe(false);
  });
});

describe('emailSchema', () => {
  it('accepts valid emails and lowercases them', () => {
    const result = emailSchema.safeParse('Alice@Example.com');
    expect(result.success).toBe(true);
    if (result.success) {
      expect(result.data).toBe('alice@example.com');
    }
  });

  it('rejects invalid emails', () => {
    expect(emailSchema.safeParse('not-email').success).toBe(false);
    expect(emailSchema.safeParse('@example.com').success).toBe(false);
  });
});

describe('birthDateSchema', () => {
  it('accepts valid date format', () => {
    expect(birthDateSchema.safeParse('2000-01-15').success).toBe(true);
    expect(birthDateSchema.safeParse('1990-12-31').success).toBe(true);
  });

  it('rejects invalid formats', () => {
    expect(birthDateSchema.safeParse('01/15/2000').success).toBe(false);
    expect(birthDateSchema.safeParse('2000-1-5').success).toBe(false);
  });

  it('rejects future dates', () => {
    expect(birthDateSchema.safeParse('2099-01-01').success).toBe(false);
  });
});

describe('bioSchema', () => {
  it('accepts valid bios', () => {
    expect(bioSchema.safeParse('Hello world').success).toBe(true);
    expect(bioSchema.safeParse('').success).toBe(true);
  });

  it('rejects bios over 300 chars', () => {
    expect(bioSchema.safeParse('a'.repeat(301)).success).toBe(false);
  });
});

describe('commentBodySchema', () => {
  it('accepts valid comments', () => {
    expect(commentBodySchema.safeParse('Great post!').success).toBe(true);
  });

  it('rejects empty comments', () => {
    expect(commentBodySchema.safeParse('').success).toBe(false);
  });

  it('rejects comments over 2000 chars', () => {
    expect(commentBodySchema.safeParse('a'.repeat(2001)).success).toBe(false);
  });
});

describe('registerPhoneStartSchema', () => {
  it('accepts valid phone registration', () => {
    const result = registerPhoneStartSchema.safeParse({ phone: '+1234567890' });
    expect(result.success).toBe(true);
  });

  it('rejects missing phone', () => {
    expect(registerPhoneStartSchema.safeParse({}).success).toBe(false);
  });
});

describe('loginSchema', () => {
  it('accepts phone login', () => {
    const result = loginSchema.safeParse({
      method: 'phone',
      identifier: '+1234567890',
      code: '123456',
      deviceInfo: { platform: 'ios' },
    });
    expect(result.success).toBe(true);
  });

  it('accepts email login', () => {
    const result = loginSchema.safeParse({
      method: 'email',
      identifier: 'user@test.com',
      password: 'secret123',
      deviceInfo: { platform: 'web' },
    });
    expect(result.success).toBe(true);
  });

  it('rejects invalid method', () => {
    const result = loginSchema.safeParse({
      method: 'invalid',
      deviceInfo: { platform: 'web' },
    });
    expect(result.success).toBe(false);
  });

  it('rejects invalid platform', () => {
    const result = loginSchema.safeParse({
      method: 'email',
      deviceInfo: { platform: 'desktop' },
    });
    expect(result.success).toBe(false);
  });
});

describe('createContentSchema', () => {
  it('accepts valid content creation', () => {
    const result = createContentSchema.safeParse({
      type: 'photo',
      audience: 'public',
      mediaIds: [],
    });
    expect(result.success).toBe(true);
  });

  it('rejects invalid content type', () => {
    const result = createContentSchema.safeParse({
      type: 'blog',
      audience: 'public',
      mediaIds: [],
    });
    expect(result.success).toBe(false);
  });

  it('rejects too many media items', () => {
    const result = createContentSchema.safeParse({
      type: 'carousel',
      audience: 'public',
      mediaIds: Array.from({ length: 11 }, (_, i) => `00000000-0000-0000-0000-00000000000${i}`),
    });
    expect(result.success).toBe(false);
  });
});

describe('createReportSchema', () => {
  it('accepts valid report', () => {
    const result = createReportSchema.safeParse({
      targetType: 'user',
      targetId: '00000000-0000-0000-0000-000000000001',
      reason: 'spam',
    });
    expect(result.success).toBe(true);
  });

  it('rejects invalid target type', () => {
    const result = createReportSchema.safeParse({
      targetType: 'invalid',
      targetId: '00000000-0000-0000-0000-000000000001',
      reason: 'spam',
    });
    expect(result.success).toBe(false);
  });
});

describe('sendMessageSchema', () => {
  it('accepts text message', () => {
    const result = sendMessageSchema.safeParse({
      kind: 'text',
      body: 'Hello!',
    });
    expect(result.success).toBe(true);
  });

  it('accepts reply message', () => {
    const result = sendMessageSchema.safeParse({
      kind: 'reply',
      body: 'Replying to this',
      replyToId: '00000000-0000-0000-0000-000000000001',
    });
    expect(result.success).toBe(true);
  });
});

describe('createStorySchema', () => {
  it('accepts valid story', () => {
    const result = createStorySchema.safeParse({
      mediaId: '00000000-0000-0000-0000-000000000001',
      audience: 'followers',
    });
    expect(result.success).toBe(true);
  });

  it('accepts story with interactive data', () => {
    const result = createStorySchema.safeParse({
      mediaId: '00000000-0000-0000-0000-000000000001',
      audience: 'public',
      interactiveType: 'poll',
      interactiveData: { question: 'Yes or no?', options: ['Yes', 'No'] },
    });
    expect(result.success).toBe(true);
  });
});

describe('trackEventSchema', () => {
  it('accepts valid event', () => {
    const result = trackEventSchema.safeParse({
      eventType: 'page_view',
      attributes: { page: '/home' },
    });
    expect(result.success).toBe(true);
  });

  it('rejects empty event type', () => {
    const result = trackEventSchema.safeParse({
      eventType: '',
      attributes: {},
    });
    expect(result.success).toBe(false);
  });
});

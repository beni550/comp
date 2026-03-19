import bcrypt from 'bcryptjs';
import crypto from 'crypto';
import jwt from 'jsonwebtoken';
import { v4 as uuidv4 } from 'uuid';
import { prisma } from '../../main';
import { AppError } from '../../middleware/error-handler';
import { createAuditLog } from '../../common/audit';
import { APP_CONFIG, RESERVED_USERNAMES } from '@vybe/config';

const JWT_SECRET = process.env.JWT_SECRET || 'dev-secret-change-me';
const JWT_REFRESH_SECRET = process.env.JWT_REFRESH_SECRET || 'dev-refresh-secret-change-me';

function generateOtp(): string {
  return Math.floor(100000 + Math.random() * 900000).toString();
}

function calculateAge(birthDate: Date): number {
  const today = new Date();
  let age = today.getFullYear() - birthDate.getFullYear();
  const m = today.getMonth() - birthDate.getMonth();
  if (m < 0 || (m === 0 && today.getDate() < birthDate.getDate())) {
    age--;
  }
  return age;
}

function generateAccessToken(user: { id: string; username: string; role: string; isMinor: boolean; status: string }): string {
  return jwt.sign(
    { id: user.id, username: user.username, role: user.role, isMinor: user.isMinor, status: user.status },
    JWT_SECRET,
    { expiresIn: APP_CONFIG.ACCESS_TOKEN_TTL }
  );
}

function generateRefreshToken(sessionId: string): string {
  return jwt.sign(
    { sessionId },
    JWT_REFRESH_SECRET,
    { expiresIn: `${APP_CONFIG.REFRESH_TOKEN_TTL_DAYS}d` }
  );
}

/**
 * SHA-256 pre-hash for bcrypt input.
 * bcrypt silently truncates inputs longer than 72 bytes. JWTs exceed this limit,
 * so tokens that differ only after byte 72 would produce identical bcrypt hashes.
 * Pre-hashing with SHA-256 produces a fixed 64-char hex string, ensuring the
 * full token is considered.
 */
function sha256(input: string): string {
  return crypto.createHash('sha256').update(input).digest('hex');
}

export const authService = {
  async registerPhoneStart(phone: string) {
    // Check if phone already exists (generic response to prevent enumeration)
    const existing = await prisma.authIdentity.findFirst({
      where: { provider: 'phone', phoneE164: phone },
    });

    const code = generateOtp();
    const codeHash = await bcrypt.hash(code, 10);

    const challenge = await prisma.otpChallenge.create({
      data: {
        target: phone,
        channel: 'sms',
        codeHash,
        purpose: existing ? 'login' : 'registration',
        expiresAt: new Date(Date.now() + APP_CONFIG.OTP_TTL_MINUTES * 60 * 1000),
      },
    });

    // In dev, log OTP to console
    if (process.env.NODE_ENV !== 'production') {
      console.log(`[DEV] OTP for ${phone}: ${code}`);
    }

    return {
      challengeId: challenge.id,
      expiresAt: challenge.expiresAt.toISOString(),
    };
  },

  async registerPhoneVerify(data: {
    challengeId: string;
    code: string;
    username: string;
    displayName: string;
    birthDate: string;
  }) {
    const challenge = await prisma.otpChallenge.findUnique({
      where: { id: data.challengeId },
    });

    if (!challenge) {
      throw new AppError(400, 'INVALID_CHALLENGE', 'Challenge not found');
    }
    if (challenge.consumedAt) {
      throw new AppError(400, 'CHALLENGE_CONSUMED', 'Challenge already used');
    }
    if (challenge.expiresAt < new Date()) {
      throw new AppError(400, 'CHALLENGE_EXPIRED', 'Challenge has expired');
    }
    if (challenge.attempts >= APP_CONFIG.OTP_MAX_ATTEMPTS) {
      throw new AppError(429, 'TOO_MANY_ATTEMPTS', 'Too many verification attempts');
    }

    await prisma.otpChallenge.update({
      where: { id: challenge.id },
      data: { attempts: challenge.attempts + 1 },
    });

    const isValid = await bcrypt.compare(data.code, challenge.codeHash);
    if (!isValid) {
      throw new AppError(400, 'INVALID_CODE', 'Incorrect verification code');
    }

    // Validate username
    await this.validateUsername(data.username);

    // Age verification
    const birthDate = new Date(data.birthDate);
    const age = calculateAge(birthDate);
    if (age < APP_CONFIG.MIN_AGE) {
      throw new AppError(400, 'AGE_REQUIREMENT', 'You must be at least 13 years old to register');
    }
    const isMinor = age < 18;

    // Check if phone already linked
    const existingIdentity = await prisma.authIdentity.findFirst({
      where: { provider: 'phone', phoneE164: challenge.target },
    });
    if (existingIdentity) {
      throw new AppError(409, 'IDENTITY_EXISTS', 'This phone number is already registered');
    }

    // Create user + identity + profile + settings + onboarding in transaction
    const user = await prisma.$transaction(async (tx) => {
      const newUser = await tx.user.create({
        data: {
          username: data.username,
          displayName: data.displayName,
          phone: challenge.target,
          birthDate,
          isMinor,
          privacyMode: isMinor ? 'followers' : 'public',
        },
      });

      await tx.authIdentity.create({
        data: {
          userId: newUser.id,
          provider: 'phone',
          phoneE164: challenge.target,
          isPrimary: true,
          verifiedAt: new Date(),
        },
      });

      await tx.userProfile.create({
        data: { userId: newUser.id },
      });

      await tx.userSettings.create({
        data: {
          userId: newUser.id,
          discoverableByPhone: !isMinor,
          dmPermission: isMinor ? 'friends' : 'followers',
        },
      });

      await tx.onboardingProgress.create({
        data: {
          userId: newUser.id,
          currentStep: 'interests',
        },
      });

      await tx.otpChallenge.update({
        where: { id: challenge.id },
        data: { consumedAt: new Date(), userId: newUser.id },
      });

      return newUser;
    });

    // Create device + session
    const { accessToken, refreshToken } = await this.createSession(user, {
      platform: 'web',
      ip: undefined,
      userAgent: undefined,
    });

    await createAuditLog({
      actorId: user.id,
      action: 'user.registered',
      metadata: { method: 'phone' },
    });

    return { user, accessToken, refreshToken };
  },

  async registerEmailStart(email: string) {
    const code = generateOtp();
    const codeHash = await bcrypt.hash(code, 10);

    const challenge = await prisma.otpChallenge.create({
      data: {
        target: email.toLowerCase(),
        channel: 'email',
        codeHash,
        purpose: 'registration',
        expiresAt: new Date(Date.now() + APP_CONFIG.OTP_TTL_MINUTES * 60 * 1000),
      },
    });

    if (process.env.NODE_ENV !== 'production') {
      console.log(`[DEV] OTP for ${email}: ${code}`);
    }

    return {
      challengeId: challenge.id,
      expiresAt: challenge.expiresAt.toISOString(),
    };
  },

  async login(data: {
    method: string;
    identifier?: string;
    code?: string;
    password?: string;
    providerToken?: string;
    deviceInfo: { platform: string; appVersion?: string };
    ip?: string;
    userAgent?: string;
  }) {
    let user;

    if (data.method === 'phone' && data.identifier && data.code) {
      // Find the latest OTP challenge for this phone
      const challenge = await prisma.otpChallenge.findFirst({
        where: {
          target: data.identifier,
          channel: 'sms',
          purpose: 'login',
          consumedAt: null,
          expiresAt: { gt: new Date() },
        },
        orderBy: { createdAt: 'desc' },
      });

      if (!challenge) {
        throw new AppError(400, 'INVALID_CREDENTIALS', 'Invalid login credentials');
      }

      if (challenge.attempts >= APP_CONFIG.OTP_MAX_ATTEMPTS) {
        throw new AppError(429, 'TOO_MANY_ATTEMPTS', 'Too many attempts');
      }

      await prisma.otpChallenge.update({
        where: { id: challenge.id },
        data: { attempts: challenge.attempts + 1 },
      });

      const isValid = await bcrypt.compare(data.code, challenge.codeHash);
      if (!isValid) {
        throw new AppError(400, 'INVALID_CREDENTIALS', 'Invalid login credentials');
      }

      await prisma.otpChallenge.update({
        where: { id: challenge.id },
        data: { consumedAt: new Date() },
      });

      const identity = await prisma.authIdentity.findFirst({
        where: { provider: 'phone', phoneE164: data.identifier },
        include: { user: true },
      });

      if (!identity) {
        throw new AppError(400, 'INVALID_CREDENTIALS', 'Invalid login credentials');
      }
      user = identity.user;
    } else if (data.method === 'email' && data.identifier && data.password) {
      user = await prisma.user.findFirst({
        where: { email: data.identifier.toLowerCase(), deletedAt: null },
      });
      if (!user) {
        throw new AppError(400, 'INVALID_CREDENTIALS', 'Invalid login credentials');
      }

      const credential = await prisma.passwordCredential.findUnique({
        where: { userId: user.id },
      });
      if (!credential) {
        throw new AppError(400, 'INVALID_CREDENTIALS', 'Invalid login credentials');
      }

      const isValid = await bcrypt.compare(data.password, credential.hash);
      if (!isValid) {
        throw new AppError(400, 'INVALID_CREDENTIALS', 'Invalid login credentials');
      }
    } else {
      throw new AppError(400, 'INVALID_METHOD', 'Unsupported login method');
    }

    if (user.status === 'banned') {
      throw new AppError(403, 'ACCOUNT_BANNED', 'Your account has been banned');
    }
    if (user.status === 'suspended') {
      throw new AppError(403, 'ACCOUNT_SUSPENDED', 'Your account is suspended');
    }

    const { accessToken, refreshToken } = await this.createSession(user, {
      platform: data.deviceInfo.platform,
      appVersion: data.deviceInfo.appVersion,
      ip: data.ip,
      userAgent: data.userAgent,
    });

    await createAuditLog({
      actorId: user.id,
      action: 'auth.login',
      ip: data.ip,
      metadata: { method: data.method, platform: data.deviceInfo.platform },
    });

    return { user, accessToken, refreshToken };
  },

  async createSession(
    user: { id: string; username: string; role: string; isMinor: boolean; status: string },
    deviceInfo: { platform: string; appVersion?: string; ip?: string; userAgent?: string }
  ) {
    const device = await prisma.device.create({
      data: {
        userId: user.id,
        platform: deviceInfo.platform,
        appVersion: deviceInfo.appVersion,
        lastSeenAt: new Date(),
      },
    });

    const sessionId = uuidv4();
    const refreshToken = generateRefreshToken(sessionId);
    const refreshTokenHash = await bcrypt.hash(sha256(refreshToken), 10);

    await prisma.session.create({
      data: {
        id: sessionId,
        userId: user.id,
        deviceId: device.id,
        refreshTokenHash,
        ip: deviceInfo.ip,
        userAgent: deviceInfo.userAgent,
        expiresAt: new Date(Date.now() + APP_CONFIG.REFRESH_TOKEN_TTL_DAYS * 24 * 60 * 60 * 1000),
      },
    });

    const accessToken = generateAccessToken(user);

    return { accessToken, refreshToken };
  },

  async refreshToken(refreshToken: string) {
    let payload: { sessionId: string };
    try {
      payload = jwt.verify(refreshToken, JWT_REFRESH_SECRET) as { sessionId: string };
    } catch {
      throw new AppError(401, 'INVALID_REFRESH_TOKEN', 'Invalid or expired refresh token');
    }

    // Look up the specific session by ID from the JWT payload (O(1) instead of O(n))
    const session = await prisma.session.findUnique({
      where: { id: payload.sessionId },
      include: { user: true },
    });

    if (!session || session.revokedAt || session.expiresAt < new Date()) {
      throw new AppError(401, 'SESSION_EXPIRED', 'Session not found or expired');
    }

    // Verify the refresh token hash matches
    const isMatch = await bcrypt.compare(sha256(refreshToken), session.refreshTokenHash);
    if (!isMatch) {
      throw new AppError(401, 'INVALID_REFRESH_TOKEN', 'Invalid refresh token');
    }

    const user = session.user;
    const newAccessToken = generateAccessToken(user);
    const newSessionId = session.id; // reuse same session ID
    const newRefreshToken = generateRefreshToken(newSessionId);
    const newRefreshTokenHash = await bcrypt.hash(sha256(newRefreshToken), 10);

    await prisma.session.update({
      where: { id: session.id },
      data: { refreshTokenHash: newRefreshTokenHash },
    });

    return { accessToken: newAccessToken, refreshToken: newRefreshToken };
  },

  async logout(userId: string, sessionToken?: string) {
    // Revoke the most recent active session
    const session = await prisma.session.findFirst({
      where: { userId, revokedAt: null },
      orderBy: { createdAt: 'desc' },
    });

    if (session) {
      await prisma.session.update({
        where: { id: session.id },
        data: { revokedAt: new Date() },
      });
    }

    await createAuditLog({
      actorId: userId,
      action: 'auth.logout',
    });
  },

  async logoutAll(userId: string) {
    const result = await prisma.session.updateMany({
      where: { userId, revokedAt: null },
      data: { revokedAt: new Date() },
    });

    await createAuditLog({
      actorId: userId,
      action: 'auth.logout_all',
      metadata: { revokedCount: result.count },
    });

    return { revokedCount: result.count };
  },

  async getSessions(userId: string) {
    return prisma.session.findMany({
      where: { userId, revokedAt: null, expiresAt: { gt: new Date() } },
      include: { device: true },
      orderBy: { createdAt: 'desc' },
    });
  },

  async revokeSession(userId: string, sessionId: string) {
    const session = await prisma.session.findFirst({
      where: { id: sessionId, userId, revokedAt: null },
    });

    if (!session) {
      throw new AppError(404, 'SESSION_NOT_FOUND', 'Session not found');
    }

    await prisma.session.update({
      where: { id: sessionId },
      data: { revokedAt: new Date() },
    });

    await createAuditLog({
      actorId: userId,
      action: 'session.revoked',
      targetType: 'session',
      targetId: sessionId,
    });
  },

  async validateUsername(username: string) {
    if (RESERVED_USERNAMES.includes(username.toLowerCase() as typeof RESERVED_USERNAMES[number])) {
      throw new AppError(400, 'USERNAME_RESERVED', 'This username is reserved');
    }

    const existing = await prisma.user.findUnique({ where: { username } });
    if (existing) {
      throw new AppError(409, 'USERNAME_TAKEN', 'This username is already taken');
    }

    // Check blocked words
    const blockedWords = await prisma.blockedWord.findMany({
      where: { severity: 'high' },
    });
    const lowerUsername = username.toLowerCase();
    for (const bw of blockedWords) {
      if (lowerUsername.includes(bw.word.toLowerCase())) {
        throw new AppError(400, 'USERNAME_INAPPROPRIATE', 'This username contains inappropriate content');
      }
    }
  },

  async checkUsernameAvailability(username: string) {
    try {
      await this.validateUsername(username);
      return { available: true };
    } catch (err) {
      if (err instanceof AppError) {
        return { available: false, reason: err.message };
      }
      throw err;
    }
  },

  async startRecovery(identifier: string) {
    const user = await prisma.user.findFirst({
      where: {
        OR: [
          { email: identifier.toLowerCase() },
          { phone: identifier },
        ],
        deletedAt: null,
      },
    });

    // Always return generic response to prevent enumeration
    const code = generateOtp();
    const codeHash = await bcrypt.hash(code, 10);
    const channel = identifier.includes('@') ? 'email' : 'sms';

    const challenge = await prisma.otpChallenge.create({
      data: {
        userId: user?.id,
        target: identifier,
        channel,
        codeHash,
        purpose: 'recovery',
        expiresAt: new Date(Date.now() + APP_CONFIG.OTP_TTL_MINUTES * 60 * 1000),
      },
    });

    if (process.env.NODE_ENV !== 'production') {
      console.log(`[DEV] Recovery OTP for ${identifier}: ${code}`);
    }

    return {
      challengeId: challenge.id,
      method: channel,
    };
  },
};

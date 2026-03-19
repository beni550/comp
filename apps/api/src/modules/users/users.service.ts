import { prisma } from '../../main';
import { AppError } from '../../middleware/error-handler';
import { createAuditLog } from '../../common/audit';
import { APP_CONFIG, RESERVED_USERNAMES } from '@vybe/config';

export const usersService = {
  async getMe(userId: string) {
    const user = await prisma.user.findUnique({
      where: { id: userId },
      include: { profile: true, settings: true, onboardingProgress: true },
    });
    if (!user || user.deletedAt) {
      throw new AppError(404, 'USER_NOT_FOUND', 'User not found');
    }
    return user;
  },

  async getPublicUser(userId: string, viewerId?: string) {
    const user = await prisma.user.findUnique({
      where: { id: userId },
      include: { profile: true },
    });
    if (!user || user.deletedAt) {
      throw new AppError(404, 'USER_NOT_FOUND', 'User not found');
    }

    // Check if viewer is blocked
    if (viewerId) {
      const blocked = await prisma.blockEdge.findFirst({
        where: {
          OR: [
            { blockerId: userId, blockedId: viewerId },
            { blockerId: viewerId, blockedId: userId },
          ],
        },
      });
      if (blocked) {
        throw new AppError(403, 'BLOCKED', 'Cannot view this user');
      }
    }

    return {
      id: user.id,
      username: user.username,
      displayName: user.displayName,
      avatarUrl: user.profile?.avatarMediaId || undefined,
      bio: user.profile?.bio,
      profileMode: user.profile?.profileMode || 'standard',
      privacyMode: user.privacyMode,
      isVerified: false, // Would check verification_requests
      followersCount: user.profile?.followersCount,
      followingCount: user.profile?.followingCount,
      postsCount: user.profile?.postsCount,
    };
  },

  async updateUser(userId: string, data: {
    displayName?: string;
    username?: string;
    email?: string;
    phone?: string;
  }) {
    const user = await prisma.user.findUnique({ where: { id: userId } });
    if (!user) throw new AppError(404, 'USER_NOT_FOUND', 'User not found');

    if (data.username && data.username !== user.username) {
      // Check cooldown
      if (user.lastUsernameChange) {
        const daysSince = (Date.now() - user.lastUsernameChange.getTime()) / (1000 * 60 * 60 * 24);
        if (daysSince < APP_CONFIG.USERNAME_CHANGE_COOLDOWN_DAYS) {
          throw new AppError(429, 'USERNAME_COOLDOWN', `You can change your username again in ${Math.ceil(APP_CONFIG.USERNAME_CHANGE_COOLDOWN_DAYS - daysSince)} days`);
        }
      }

      // Validate new username
      if (RESERVED_USERNAMES.includes(data.username.toLowerCase() as typeof RESERVED_USERNAMES[number])) {
        throw new AppError(400, 'USERNAME_RESERVED', 'This username is reserved');
      }
      const existing = await prisma.user.findUnique({ where: { username: data.username } });
      if (existing) {
        throw new AppError(409, 'USERNAME_TAKEN', 'This username is already taken');
      }

      // Save history
      await prisma.usernameHistory.create({
        data: { userId, previousUsername: user.username },
      });
    }

    const updated = await prisma.user.update({
      where: { id: userId },
      data: {
        ...data,
        lastUsernameChange: data.username && data.username !== user.username ? new Date() : undefined,
      },
    });

    await createAuditLog({
      actorId: userId,
      action: 'user.updated',
      targetType: 'user',
      targetId: userId,
      metadata: { fields: Object.keys(data) },
    });

    return updated;
  },

  async updateProfile(userId: string, data: {
    bio?: string;
    avatarMediaId?: string | null;
    theme?: string;
    mood?: string | null;
    moodEmoji?: string | null;
    links?: string[];
    profileMode?: string;
    featuredContentIds?: string[];
  }) {
    const profile = await prisma.userProfile.findUnique({ where: { userId } });
    if (!profile) {
      throw new AppError(404, 'PROFILE_NOT_FOUND', 'Profile not found');
    }

    return prisma.userProfile.update({
      where: { userId },
      data: {
        bio: data.bio,
        avatarMediaId: data.avatarMediaId,
        theme: data.theme,
        mood: data.mood,
        moodEmoji: data.moodEmoji,
        links: data.links !== undefined ? JSON.stringify(data.links) : undefined,
        profileMode: data.profileMode,
        featuredContentIds: data.featuredContentIds !== undefined
          ? JSON.stringify(data.featuredContentIds)
          : undefined,
      },
    });
  },

  async updateSettings(userId: string, data: Record<string, unknown>) {
    const settings = await prisma.userSettings.findUnique({ where: { userId } });
    if (!settings) {
      throw new AppError(404, 'SETTINGS_NOT_FOUND', 'Settings not found');
    }

    const updateData: Record<string, unknown> = {};
    const allowedFields = [
      'defaultAudience', 'dmPermission', 'readReceipts', 'screenshotAlert',
      'downloadPermission', 'discoverableByPhone', 'discoverableByEmail',
      'sensitiveContentLevel', 'autoplay', 'textScale', 'reduceMotion',
      'dailyUsageLimit', 'breakReminders', 'quietHoursStart', 'quietHoursEnd',
      'locale',
    ];

    for (const field of allowedFields) {
      if (data[field] !== undefined) {
        updateData[field] = data[field];
      }
    }

    if (data.pushNotifications !== undefined) {
      updateData.pushNotifications = data.pushNotifications;
    }
    if (data.emailNotifications !== undefined) {
      updateData.emailNotifications = data.emailNotifications;
    }

    return prisma.userSettings.update({
      where: { userId },
      data: updateData,
    });
  },

  async searchUsers(query: string, limit = 20) {
    return prisma.user.findMany({
      where: {
        deletedAt: null,
        status: 'active',
        OR: [
          { username: { contains: query, mode: 'insensitive' } },
          { displayName: { contains: query, mode: 'insensitive' } },
        ],
      },
      include: { profile: true },
      take: limit,
    });
  },

  async deleteUser(userId: string) {
    await prisma.user.update({
      where: { id: userId },
      data: {
        status: 'deactivated',
        deletedAt: new Date(),
      },
    });

    // Revoke all sessions
    await prisma.session.updateMany({
      where: { userId, revokedAt: null },
      data: { revokedAt: new Date() },
    });

    await createAuditLog({
      actorId: userId,
      action: 'user.deleted',
      targetType: 'user',
      targetId: userId,
    });
  },

  async getOnboardingProgress(userId: string) {
    return prisma.onboardingProgress.findUnique({ where: { userId } });
  },

  async completeOnboardingStep(userId: string, step: string) {
    const progress = await prisma.onboardingProgress.findUnique({ where: { userId } });
    if (!progress) {
      throw new AppError(404, 'ONBOARDING_NOT_FOUND', 'Onboarding progress not found');
    }

    const completedSteps = progress.completedSteps as string[];
    if (!completedSteps.includes(step)) {
      completedSteps.push(step);
    }

    const allSteps = ['interests', 'avatar', 'contacts', 'follow_suggestions', 'first_post'];
    const nextStepIndex = allSteps.indexOf(step) + 1;
    const nextStep = nextStepIndex < allSteps.length ? allSteps[nextStepIndex] : step;
    const isComplete = completedSteps.length >= allSteps.length ||
      (completedSteps.length + (progress.skippedSteps as string[]).length >= allSteps.length);

    return prisma.onboardingProgress.update({
      where: { userId },
      data: {
        completedSteps,
        currentStep: isComplete ? 'done' : nextStep,
        isComplete,
      },
    });
  },

  async skipOnboardingStep(userId: string, step: string) {
    const progress = await prisma.onboardingProgress.findUnique({ where: { userId } });
    if (!progress) {
      throw new AppError(404, 'ONBOARDING_NOT_FOUND', 'Onboarding progress not found');
    }

    const skippedSteps = progress.skippedSteps as string[];
    if (!skippedSteps.includes(step)) {
      skippedSteps.push(step);
    }

    const allSteps = ['interests', 'avatar', 'contacts', 'follow_suggestions', 'first_post'];
    const currentIdx = allSteps.indexOf(step);
    const nextStep = currentIdx + 1 < allSteps.length ? allSteps[currentIdx + 1] : 'done';
    const completedSteps = progress.completedSteps as string[];
    const isComplete = completedSteps.length + skippedSteps.length >= allSteps.length;

    return prisma.onboardingProgress.update({
      where: { userId },
      data: {
        skippedSteps,
        currentStep: isComplete ? 'done' : nextStep,
        isComplete,
      },
    });
  },
};

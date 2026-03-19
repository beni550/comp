import { PrismaClient } from '@prisma/client';
import bcrypt from 'bcryptjs';
import { v4 as uuidv4 } from 'uuid';

const prisma = new PrismaClient();

async function main() {
  console.log('Seeding database...');

  // ─── Interest Categories ────────────────────────────────────
  const interests = [
    { name: 'Music', slug: 'music' },
    { name: 'Sports', slug: 'sports' },
    { name: 'Art', slug: 'art' },
    { name: 'Technology', slug: 'technology' },
    { name: 'Fashion', slug: 'fashion' },
    { name: 'Food', slug: 'food' },
    { name: 'Travel', slug: 'travel' },
    { name: 'Gaming', slug: 'gaming' },
    { name: 'Photography', slug: 'photography' },
    { name: 'Fitness', slug: 'fitness' },
  ];

  for (const interest of interests) {
    await prisma.interestCategory.upsert({
      where: { slug: interest.slug },
      update: {},
      create: {
        id: uuidv4(),
        name: interest.name,
        slug: interest.slug,
        sortOrder: interests.indexOf(interest),
      },
    });
  }
  console.log(`  Created ${interests.length} interest categories`);

  // ─── Feature Flags ──────────────────────────────────────────
  const flags = [
    { key: 'stories_enabled', description: 'Enable Stories feature', isEnabled: true },
    { key: 'circles_enabled', description: 'Enable Circles feature', isEnabled: true },
    { key: 'dm_enabled', description: 'Enable Direct Messages', isEnabled: true },
    { key: 'media_upload_enabled', description: 'Enable media uploads', isEnabled: true },
    { key: 'notifications_enabled', description: 'Enable push notifications', isEnabled: true },
    { key: 'analytics_enabled', description: 'Enable analytics event tracking', isEnabled: true },
    { key: 'minor_safety_mode', description: 'Enable safety restrictions for minors', isEnabled: true },
    { key: 'content_moderation_enabled', description: 'Enable content moderation', isEnabled: true },
  ];

  for (const flag of flags) {
    await prisma.featureFlag.upsert({
      where: { key: flag.key },
      update: {},
      create: {
        id: uuidv4(),
        key: flag.key,
        description: flag.description,
        isEnabled: flag.isEnabled,
        rolloutPercentage: 100,
        environment: 'development',
      },
    });
  }
  console.log(`  Created ${flags.length} feature flags`);

  // ─── Users ──────────────────────────────────────────────────
  const passwordHash = await bcrypt.hash('Password123!', 12);

  const adminUser = await prisma.user.upsert({
    where: { email: 'admin@vybe.app' },
    update: {},
    create: {
      id: uuidv4(),
      email: 'admin@vybe.app',
      phone: '+1555000001',
      username: 'admin',
      displayName: 'VYBE Admin',
      role: 'admin',
      dateOfBirth: new Date('1990-01-01'),
      isVerified: true,
      onboardingComplete: true,
      status: 'active',
    },
  });

  // Create auth identity for admin
  await prisma.authIdentity.upsert({
    where: { provider_providerKey: { provider: 'email', providerKey: 'admin@vybe.app' } },
    update: {},
    create: {
      id: uuidv4(),
      userId: adminUser.id,
      provider: 'email',
      providerKey: 'admin@vybe.app',
      passwordHash,
    },
  });

  // Create admin profile
  await prisma.userProfile.upsert({
    where: { userId: adminUser.id },
    update: {},
    create: {
      id: uuidv4(),
      userId: adminUser.id,
      bio: 'Platform administrator',
      location: 'San Francisco, CA',
    },
  });

  // Create admin settings
  await prisma.userSettings.upsert({
    where: { userId: adminUser.id },
    update: {},
    create: {
      id: uuidv4(),
      userId: adminUser.id,
    },
  });

  console.log(`  Created admin user: ${adminUser.username} (${adminUser.email})`);

  const moderatorUser = await prisma.user.upsert({
    where: { email: 'mod@vybe.app' },
    update: {},
    create: {
      id: uuidv4(),
      email: 'mod@vybe.app',
      phone: '+1555000002',
      username: 'moderator',
      displayName: 'VYBE Moderator',
      role: 'moderator',
      dateOfBirth: new Date('1992-06-15'),
      isVerified: true,
      onboardingComplete: true,
      status: 'active',
    },
  });

  await prisma.authIdentity.upsert({
    where: { provider_providerKey: { provider: 'email', providerKey: 'mod@vybe.app' } },
    update: {},
    create: {
      id: uuidv4(),
      userId: moderatorUser.id,
      provider: 'email',
      providerKey: 'mod@vybe.app',
      passwordHash,
    },
  });

  await prisma.userProfile.upsert({
    where: { userId: moderatorUser.id },
    update: {},
    create: {
      id: uuidv4(),
      userId: moderatorUser.id,
      bio: 'Community moderator',
    },
  });

  await prisma.userSettings.upsert({
    where: { userId: moderatorUser.id },
    update: {},
    create: {
      id: uuidv4(),
      userId: moderatorUser.id,
    },
  });

  console.log(`  Created moderator user: ${moderatorUser.username} (${moderatorUser.email})`);

  // Create test users
  const testUsers = [
    { username: 'alice', displayName: 'Alice Johnson', email: 'alice@test.com', phone: '+1555000010', bio: 'Photographer and traveler' },
    { username: 'bob', displayName: 'Bob Smith', email: 'bob@test.com', phone: '+1555000011', bio: 'Music producer' },
    { username: 'charlie', displayName: 'Charlie Brown', email: 'charlie@test.com', phone: '+1555000012', bio: 'Tech enthusiast' },
    { username: 'diana', displayName: 'Diana Prince', email: 'diana@test.com', phone: '+1555000013', bio: 'Fitness coach' },
    { username: 'eve', displayName: 'Eve Wilson', email: 'eve@test.com', phone: '+1555000014', bio: 'Food blogger' },
  ];

  const createdUsers = [];
  for (const userData of testUsers) {
    const user = await prisma.user.upsert({
      where: { email: userData.email },
      update: {},
      create: {
        id: uuidv4(),
        email: userData.email,
        phone: userData.phone,
        username: userData.username,
        displayName: userData.displayName,
        role: 'user',
        dateOfBirth: new Date('1995-03-20'),
        isVerified: true,
        onboardingComplete: true,
        status: 'active',
      },
    });

    await prisma.authIdentity.upsert({
      where: { provider_providerKey: { provider: 'email', providerKey: userData.email } },
      update: {},
      create: {
        id: uuidv4(),
        userId: user.id,
        provider: 'email',
        providerKey: userData.email,
        passwordHash,
      },
    });

    await prisma.userProfile.upsert({
      where: { userId: user.id },
      update: {},
      create: {
        id: uuidv4(),
        userId: user.id,
        bio: userData.bio,
      },
    });

    await prisma.userSettings.upsert({
      where: { userId: user.id },
      update: {},
      create: {
        id: uuidv4(),
        userId: user.id,
      },
    });

    createdUsers.push(user);
    console.log(`  Created test user: ${user.username} (${user.email})`);
  }

  // ─── Follow Relationships ──────────────────────────────────
  // Alice follows Bob, Charlie, Diana
  // Bob follows Alice, Charlie
  // Charlie follows everyone
  const followPairs = [
    [0, 1], [0, 2], [0, 3],
    [1, 0], [1, 2],
    [2, 0], [2, 1], [2, 3], [2, 4],
    [3, 0], [3, 4],
    [4, 0], [4, 1],
  ];

  for (const [followerIdx, followeeIdx] of followPairs) {
    const followerId = createdUsers[followerIdx].id;
    const followeeId = createdUsers[followeeIdx].id;
    await prisma.followEdge.upsert({
      where: { followerId_followeeId: { followerId, followeeId } },
      update: {},
      create: {
        id: uuidv4(),
        followerId,
        followeeId,
        status: 'active',
      },
    });
  }
  console.log(`  Created ${followPairs.length} follow relationships`);

  // ─── Sample Content ────────────────────────────────────────
  const samplePosts = [
    { authorIdx: 0, body: 'Just captured the most amazing sunset at the beach! The colors were unreal.', hashtags: ['photography', 'sunset', 'nature'] },
    { authorIdx: 1, body: 'New beat just dropped! Been working on this track for weeks. Let me know what you think.', hashtags: ['music', 'producer', 'newmusic'] },
    { authorIdx: 2, body: 'Anyone else excited about the new AI developments? The pace of innovation is incredible.', hashtags: ['technology', 'ai', 'innovation'] },
    { authorIdx: 3, body: 'Morning workout done! Remember: consistency beats intensity every time.', hashtags: ['fitness', 'motivation', 'health'] },
    { authorIdx: 4, body: 'Made the most delicious homemade pasta today. Recipe coming soon!', hashtags: ['food', 'cooking', 'homemade'] },
    { authorIdx: 0, body: 'Exploring the streets of Tokyo. Every corner has a story to tell.', hashtags: ['travel', 'tokyo', 'streetphotography'] },
    { authorIdx: 1, body: 'Collaboration is the key to great music. Grateful for my studio session today.', hashtags: ['music', 'collaboration', 'studio'] },
    { authorIdx: 2, body: 'Just set up my new home lab. Ready for some serious tinkering!', hashtags: ['technology', 'homelab', 'diy'] },
  ];

  for (const post of samplePosts) {
    const contentId = uuidv4();
    await prisma.contentItem.create({
      data: {
        id: contentId,
        authorId: createdUsers[post.authorIdx].id,
        type: 'post',
        body: post.body,
        audience: 'public',
        status: 'published',
        hashtags: JSON.stringify(post.hashtags),
        publishedAt: new Date(Date.now() - Math.random() * 7 * 24 * 60 * 60 * 1000),
      },
    });
  }
  console.log(`  Created ${samplePosts.length} sample posts`);

  // ─── Sample Comments ───────────────────────────────────────
  // Get published content for comments
  const publishedContent = await prisma.contentItem.findMany({
    where: { status: 'published' },
    take: 4,
  });

  if (publishedContent.length > 0) {
    const comments = [
      { contentIdx: 0, authorIdx: 1, body: 'This is stunning! Where was this taken?' },
      { contentIdx: 0, authorIdx: 2, body: 'Amazing colors! What camera do you use?' },
      { contentIdx: 1, authorIdx: 0, body: 'This beat is fire! Can we collab?' },
      { contentIdx: 2, authorIdx: 3, body: 'So true! AI is changing everything.' },
      { contentIdx: 3, authorIdx: 4, body: 'You inspire me to get moving!' },
    ];

    for (const comment of comments) {
      if (publishedContent[comment.contentIdx]) {
        await prisma.comment.create({
          data: {
            id: uuidv4(),
            contentId: publishedContent[comment.contentIdx].id,
            authorId: createdUsers[comment.authorIdx].id,
            body: comment.body,
          },
        });
      }
    }
    console.log(`  Created sample comments`);

    // Add some reactions
    for (let i = 0; i < Math.min(4, publishedContent.length); i++) {
      for (let j = 0; j < 3; j++) {
        const userIdx = (i + j + 1) % createdUsers.length;
        await prisma.reaction.upsert({
          where: {
            contentId_actorId: {
              contentId: publishedContent[i].id,
              actorId: createdUsers[userIdx].id,
            },
          },
          update: {},
          create: {
            id: uuidv4(),
            contentId: publishedContent[i].id,
            actorId: createdUsers[userIdx].id,
            emoji: ['like', 'love', 'fire'][j],
          },
        });
      }
    }
    console.log(`  Created sample reactions`);
  }

  // ─── Feed Config ───────────────────────────────────────────
  const feedConfigs = [
    { key: 'recency_weight', value: 0.4, description: 'Weight for content recency in feed ranking' },
    { key: 'engagement_weight', value: 0.3, description: 'Weight for engagement metrics in feed ranking' },
    { key: 'affinity_weight', value: 0.2, description: 'Weight for user affinity in feed ranking' },
    { key: 'diversity_weight', value: 0.1, description: 'Weight for content diversity in feed ranking' },
    { key: 'trending_threshold', value: 50, description: 'Minimum engagement score for trending content' },
  ];

  for (const config of feedConfigs) {
    await prisma.feedConfig.upsert({
      where: { key: config.key },
      update: {},
      create: {
        id: uuidv4(),
        key: config.key,
        value: config.value,
        description: config.description,
      },
    });
  }
  console.log(`  Created ${feedConfigs.length} feed config entries`);

  // ─── Blocked Words ─────────────────────────────────────────
  const blockedWords = ['spam', 'scam', 'phishing'];
  for (const word of blockedWords) {
    await prisma.blockedWord.upsert({
      where: { word },
      update: {},
      create: {
        id: uuidv4(),
        word,
        severity: 'high',
      },
    });
  }
  console.log(`  Created ${blockedWords.length} blocked words`);

  console.log('\nSeeding complete!');
  console.log('\nTest accounts (password: Password123!):');
  console.log('  Admin:     admin@vybe.app');
  console.log('  Moderator: mod@vybe.app');
  console.log('  Users:     alice@test.com, bob@test.com, charlie@test.com, diana@test.com, eve@test.com');
}

main()
  .catch((e) => {
    console.error('Seed error:', e);
    process.exit(1);
  })
  .finally(async () => {
    await prisma.$disconnect();
  });

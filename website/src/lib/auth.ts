import { betterAuth } from 'better-auth';
import { Kysely } from 'kysely';
import { D1Dialect } from 'kysely-d1';

export function createAuth(env: Env) {
  return betterAuth({
    database: {
      db: new Kysely({
        dialect: new D1Dialect({ database: env.DB }),
      }),
      type: 'sqlite',
    },
    secret: env.BETTER_AUTH_SECRET,
    baseURL: env.BETTER_AUTH_URL,
    socialProviders: {
      github: {
        clientId: env.OAUTH_GITHUB_CLIENT_ID,
        clientSecret: env.OAUTH_GITHUB_CLIENT_SECRET,
      },
    },
    emailAndPassword: { enabled: false },
  });
}

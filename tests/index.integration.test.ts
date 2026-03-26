/**
 * Integration tests for the clientCredential plugin.
 *
 * These tests make real HTTP requests to the Microsoft identity platform.
 * They require the following environment variables (populated from .env.test):
 *
 *   VITE_ENTRA_TENANT_ID      - Azure AD tenant ID
 *   VITE_ENTRA_CLIENT_ID      - App registration client ID
 *   VITE_ENTRA_CLIENT_SECRET  - App registration client secret
 *   VITE_ENTRA_OBO_SCOPE      - Application-level scope, e.g. api://<id>/.default
 *
 * Tests are skipped automatically when any of these variables is missing.
 */

import { describe, it, expect, beforeAll } from "vitest";
import { getTestInstance } from "better-auth/test";
import { clientCredential, type ClientCredential } from "../src";

// ---------------------------------------------------------------------------
// Env + skip guard
// ---------------------------------------------------------------------------

const tenantId = process.env.VITE_ENTRA_TENANT_ID;
const clientId = process.env.VITE_ENTRA_CLIENT_ID;
const clientSecret = process.env.VITE_ENTRA_CLIENT_SECRET;
const oboScope = process.env.VITE_ENTRA_OBO_SCOPE;

const hasCredentials = Boolean(tenantId && clientId && clientSecret && oboScope);

const itLive = hasCredentials
  ? it
  : it.skip.bind(it, "skipped: VITE_ENTRA_* env vars not set");

// ---------------------------------------------------------------------------
// Shared auth instance (created once; SQLite in-memory via getTestInstance)
// ---------------------------------------------------------------------------

const APP_NAME = "integration-test-app";

// ms social provider config — must match msSocialProviderConfigSchema shape
const MS_PROVIDER_OPTIONS = {
  tenantId: tenantId ?? "",
  clientId: clientId ?? "",
  clientSecret: clientSecret ?? "",
};

// The application-level scope used for client credentials
const APP_SCOPE = [oboScope ?? ""];

/**
 * Create a fully-typed auth instance with the clientCredential plugin.
 * Returning the result of getTestInstance directly (rather than assigning it
 * into a pre-declared variable with an explicit type annotation) lets
 * TypeScript infer the precise plugin-aware type, ensuring auth.api includes
 * getClientCredentialToken.
 */
async function createLiveAuthInstance(
  providerOptions = MS_PROVIDER_OPTIONS,
  pluginOptions: Parameters<typeof clientCredential>[0] = {
    [APP_NAME]: { scope: APP_SCOPE },
  },
) {
  const plugin = clientCredential(pluginOptions);
  return getTestInstance({
    socialProviders: { microsoft: providerOptions as never },
    plugins: [plugin],
  });
}

// Holds the shared instance for tests that need it across multiple assertions.
// Typed by inference from createLiveAuthInstance so the full plugin API is visible.
let testInstance: Awaited<ReturnType<typeof createLiveAuthInstance>>;

beforeAll(async () => {
  if (!hasCredentials) return;
  testInstance = await createLiveAuthInstance();
});

// ---------------------------------------------------------------------------
// Tests
// ---------------------------------------------------------------------------

describe("clientCredential plugin — integration (live Microsoft)", () => {
  describe("token acquisition", () => {
    itLive("returns a valid access token from Microsoft", async () => {
      const result = await testInstance.auth.api.getClientCredentialToken({
        body: { applicationName: APP_NAME },
      });

      // Token must be a non-empty JWT-like string
      expect(typeof result.accessToken).toBe("string");
      expect(result.accessToken.length).toBeGreaterThan(0);
      expect(result.accessToken.split(".").length).toBe(3); // header.payload.sig

      expect(result.applicationName).toBe(APP_NAME);
      expect(result.scope).toBe(APP_SCOPE.join(" "));
    });

    itLive("sets accessTokenExpiresAt to a future date", async () => {
      const result = await testInstance.auth.api.getClientCredentialToken({
        body: { applicationName: APP_NAME },
      });

      expect(result.accessTokenExpiresAt).toBeInstanceOf(Date);
      expect(result.accessTokenExpiresAt.getTime()).toBeGreaterThan(Date.now());
    });

    itLive("persists the token to the database", async () => {
      await testInstance.auth.api.getClientCredentialToken({
        body: { applicationName: APP_NAME },
      });

      const ctx = await testInstance.auth.$context;
      const stored = await ctx.adapter.findOne<ClientCredential>({
        model: "clientCredential",
        where: [
          { field: "applicationName", value: APP_NAME, operator: "eq" },
        ],
      });

      expect(stored).not.toBeNull();
      expect(stored?.accessToken).toBeTruthy();
      expect(stored?.accessTokenExpiresAt.getTime()).toBeGreaterThan(Date.now());
    });
  });

  describe("caching", () => {
    itLive(
      "returns the cached token on a second call without hitting Microsoft again",
      async () => {
        // First call — fetches from Microsoft and caches
        const first = await testInstance.auth.api.getClientCredentialToken({
          body: { applicationName: APP_NAME },
        });

        const start = Date.now();

        // Second call — must be served from cache (near-instant)
        const second = await testInstance.auth.api.getClientCredentialToken({
          body: { applicationName: APP_NAME },
        });

        const elapsed = Date.now() - start;

        expect(second.accessToken).toBe(first.accessToken);
        expect(second.id).toBe(first.id);

        // A real MS token round-trip takes 100ms+; a cache hit should be <50ms
        expect(elapsed).toBeLessThan(50);
      },
    );

    itLive(
      "refreshes the token when the cached entry is within the 60 s expiry buffer",
      async () => {
        // Ensure there is a cached entry first
        const first = await testInstance.auth.api.getClientCredentialToken({
          body: { applicationName: APP_NAME },
        });

        // Force the cached entry into the expiry buffer (30 s from now)
        const ctx = await testInstance.auth.$context;
        const stored = await ctx.adapter.findOne<ClientCredential>({
          model: "clientCredential",
          where: [
            { field: "applicationName", value: APP_NAME, operator: "eq" },
          ],
        });
        await ctx.adapter.update<ClientCredential>({
          model: "clientCredential",
          where: [{ field: "id", value: stored!.id, operator: "eq" }],
          update: { accessTokenExpiresAt: new Date(Date.now() + 30_000) },
        });

        // Next call should detect the near-expiry and fetch a fresh token
        const refreshed = await testInstance.auth.api.getClientCredentialToken({
          body: { applicationName: APP_NAME },
        });

        // We get a new access_token (Microsoft issues a new one each time)
        expect(typeof refreshed.accessToken).toBe("string");
        expect(refreshed.accessToken.length).toBeGreaterThan(0);

        // The new expiry must be further in the future than the 30 s we set
        expect(refreshed.accessTokenExpiresAt.getTime()).toBeGreaterThan(
          Date.now() + 30_000,
        );

        // The record was updated in place, not duplicated
        const ctx2 = await testInstance.auth.$context;
        const records = await ctx2.adapter.findMany<ClientCredential>({
          model: "clientCredential",
          where: [
            { field: "applicationName", value: APP_NAME, operator: "eq" },
          ],
        });
        expect(records.length).toBe(1);
        expect(records[0]!.id).toBe(first.id);
      },
    );
  });

  describe("error handling", () => {
    itLive(
      "throws MICROSOFT_APPLICATION_NOT_FOUND for an unregistered app name",
      async () => {
        await expect(
          testInstance.auth.api.getClientCredentialToken({
            body: { applicationName: "not-registered" },
          }),
        ).rejects.toMatchObject({
          body: expect.objectContaining({
            code: "MICROSOFT_APPLICATION_NOT_FOUND",
          }),
        });
      },
    );

    itLive(
      "throws MICROSOFT_TOKEN_REQUEST_FAILED with error detail when credentials are wrong",
      async () => {
        // Stand up a separate fully-typed instance with a bad client secret.
        // Using createLiveAuthInstance keeps the type inferred, so badAuth.api
        // also carries getClientCredentialToken.
        const { auth: badAuth } = await createLiveAuthInstance(
          { ...MS_PROVIDER_OPTIONS, clientSecret: "invalid-secret" },
          { "bad-app": { scope: APP_SCOPE } },
        );

        await expect(
          badAuth.api.getClientCredentialToken({
            body: { applicationName: "bad-app" },
          }),
        ).rejects.toMatchObject({
          body: expect.objectContaining({
            code: "MICROSOFT_TOKEN_REQUEST_FAILED",
          }),
        });
      },
    );
  });
});

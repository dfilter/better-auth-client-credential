import { describe, it, expect, vi, beforeEach } from "vitest";
import { getTestInstance } from "better-auth/test";
import { clientCredential, type ClientCredential } from "../src";

// ---------------------------------------------------------------------------
// Mock @better-fetch/fetch so we never make real HTTP calls to Microsoft.
// We expose `mockFetch` so individual tests can control the response.
// ---------------------------------------------------------------------------

const mockFetch = vi.fn();

vi.mock("@better-fetch/fetch", async (importOriginal) => {
  const actual =
    await importOriginal<typeof import("@better-fetch/fetch")>();
  return {
    ...actual,
    createFetch: () => mockFetch,
  };
});

// ---------------------------------------------------------------------------
// Helpers
// ---------------------------------------------------------------------------

const MS_TOKEN_RESPONSE = {
  access_token: "mock-access-token",
  expires_in: 3600,
  token_type: "Bearer",
};

/** Return a successful Microsoft token response. */
function mockSuccessfulTokenFetch() {
  mockFetch.mockResolvedValueOnce({
    data: MS_TOKEN_RESPONSE,
    error: null,
  });
}

/** Return a failed Microsoft token response. */
function mockFailedTokenFetch() {
  mockFetch.mockResolvedValueOnce({
    data: null,
    error: {
      error: "invalid_client",
      error_description: "AADSTS70011",
      error_codes: [70011],
      timestamp: new Date(),
      trace_id: "abc",
      correlation_id: "def",
    },
  });
}

/** The Microsoft social provider options used across tests. */
const MS_OPTIONS = {
  tenantId: "test-tenant-id",
  clientId: "test-client-id",
  clientSecret: "test-client-secret",
  authority: "https://login.microsoftonline.com/test-tenant-id",
};

const DEFAULT_SCOPE = ["https://graph.microsoft.com/.default"];

// ---------------------------------------------------------------------------
// Shared auth instance factory
// ---------------------------------------------------------------------------

async function createAuthInstance(
  overrides: {
    withMicrosoftProvider?: boolean;
    pluginOptions?: Parameters<typeof clientCredential>[0];
  } = {},
) {
  const {
    withMicrosoftProvider = true,
    pluginOptions = { "my-app": { scope: DEFAULT_SCOPE } },
  } = overrides;

  const plugin = clientCredential(pluginOptions);

  const { auth } = await getTestInstance({
    socialProviders: withMicrosoftProvider
      ? { microsoft: MS_OPTIONS as never }
      : undefined,
    plugins: [plugin],
  });

  return { auth, plugin };
}

// ---------------------------------------------------------------------------
// Tests
// ---------------------------------------------------------------------------

describe("clientCredential plugin", () => {
  beforeEach(() => {
    mockFetch.mockReset();
  });

  // ─── Plugin shape ──────────────────────────────────────────────────────────

  describe("plugin shape", () => {
    it("has the correct id", () => {
      const plugin = clientCredential({ "my-app": { scope: DEFAULT_SCOPE } });
      expect(plugin.id).toBe("client-credential");
    });

    it("exports $ERROR_CODES with all expected codes", () => {
      const plugin = clientCredential({ "my-app": { scope: DEFAULT_SCOPE } });
      expect(plugin.$ERROR_CODES).toMatchObject({
        MICROSOFT_PROVIDER_NOT_CONFIGURED: expect.objectContaining({
          code: "MICROSOFT_PROVIDER_NOT_CONFIGURED",
        }),
        MICROSOFT_PROVIDER_INVALID_CONFIG: expect.objectContaining({
          code: "MICROSOFT_PROVIDER_INVALID_CONFIG",
        }),
        MICROSOFT_TOKEN_REQUEST_FAILED: expect.objectContaining({
          code: "MICROSOFT_TOKEN_REQUEST_FAILED",
        }),
        MICROSOFT_APPLICATION_NOT_FOUND: expect.objectContaining({
          code: "MICROSOFT_APPLICATION_NOT_FOUND",
        }),
      });
    });

    it("registers a clientCredential schema model with required fields", () => {
      const plugin = clientCredential({ "my-app": { scope: DEFAULT_SCOPE } });
      const fields = plugin.schema.clientCredential.fields;

      expect(fields).toMatchObject({
        applicationName: { type: "string", required: true },
        accessToken: { type: "string", required: true },
        accessTokenExpiresAt: { type: "date", required: true },
        scope: { type: "string", required: true },
      });
    });

    it("does not include a providerId field in the schema", () => {
      const plugin = clientCredential({ "my-app": { scope: DEFAULT_SCOPE } });
      const fields = plugin.schema.clientCredential.fields as Record<
        string,
        unknown
      >;
      expect(fields["providerId"]).toBeUndefined();
    });

    it("does not define a defaultValue on the id schema field", () => {
      const plugin = clientCredential({ "my-app": { scope: DEFAULT_SCOPE } });
      const fields = plugin.schema.clientCredential.fields as Record<
        string,
        unknown
      >;
      // Better Auth's adapter handles ID generation; no override should exist.
      expect(fields["id"]).toBeUndefined();
    });

    it("registers a getClientCredentialToken endpoint", () => {
      const plugin = clientCredential({ "my-app": { scope: DEFAULT_SCOPE } });
      expect(plugin.endpoints).toHaveProperty("getClientCredentialToken");
    });
  });

  // ─── Token fetch — cache miss ──────────────────────────────────────────────

  describe("getClientCredentialToken — cache miss", () => {
    it("fetches a token from Microsoft and returns the credential", async () => {
      mockSuccessfulTokenFetch();

      const { auth } = await createAuthInstance();

      const result = await auth.api.getClientCredentialToken({
        body: { applicationName: "my-app" },
      });

      expect(result).toMatchObject({
        applicationName: "my-app",
        accessToken: "mock-access-token",
        scope: DEFAULT_SCOPE.join(" "),
      });
    });

    it("sets accessTokenExpiresAt to now + expires_in seconds", async () => {
      mockSuccessfulTokenFetch();

      const before = Date.now();
      const { auth } = await createAuthInstance();

      const result = await auth.api.getClientCredentialToken({
        body: { applicationName: "my-app" },
      });

      const after = Date.now();
      const expiresAt = result.accessTokenExpiresAt.getTime();

      // expires_in = 3600s; allow ±5 s of test overhead
      expect(expiresAt).toBeGreaterThanOrEqual(before + 3600 * 1000 - 5_000);
      expect(expiresAt).toBeLessThanOrEqual(after + 3600 * 1000 + 5_000);
    });

    it("persists the credential to the database", async () => {
      mockSuccessfulTokenFetch();

      const { auth } = await createAuthInstance();
      await auth.api.getClientCredentialToken({
        body: { applicationName: "my-app" },
      });

      const ctx = await auth.$context;
      const stored = await ctx.adapter.findOne<ClientCredential>({
        model: "clientCredential",
        where: [
          { field: "applicationName", value: "my-app", operator: "eq" },
        ],
      });

      expect(stored).not.toBeNull();
      expect(stored?.accessToken).toBe("mock-access-token");
      expect(stored?.applicationName).toBe("my-app");
    });

    it("uses the scope configured for the given applicationName", async () => {
      mockSuccessfulTokenFetch();

      const customScope = ["https://graph.microsoft.com/.default", "offline_access"];
      const { auth } = await createAuthInstance({
        pluginOptions: { "my-app": { scope: customScope } },
      });
      await auth.api.getClientCredentialToken({
        body: { applicationName: "my-app" },
      });

      const callBody = mockFetch.mock.calls[0]?.[1]?.body as
        | Record<string, string>
        | undefined;
      expect(callBody?.scope).toBe(customScope.join(" "));
    });

    it("stores separate cache entries for different applicationNames", async () => {
      mockFetch
        .mockResolvedValueOnce({
          data: { ...MS_TOKEN_RESPONSE, access_token: "token-a" },
          error: null,
        })
        .mockResolvedValueOnce({
          data: { ...MS_TOKEN_RESPONSE, access_token: "token-b" },
          error: null,
        });

      const { auth } = await createAuthInstance({
        pluginOptions: {
          "app-a": { scope: ["https://graph.microsoft.com/.default"] },
          "app-b": { scope: ["https://graph.microsoft.com/.default"] },
        },
      });

      const resultA = await auth.api.getClientCredentialToken({
        body: { applicationName: "app-a" },
      });
      const resultB = await auth.api.getClientCredentialToken({
        body: { applicationName: "app-b" },
      });

      expect(resultA.accessToken).toBe("token-a");
      expect(resultB.accessToken).toBe("token-b");
      expect(mockFetch).toHaveBeenCalledTimes(2);
    });

    it("uses the per-app scope when multiple apps are registered", async () => {
      mockFetch
        .mockResolvedValueOnce({
          data: { ...MS_TOKEN_RESPONSE, access_token: "token-a" },
          error: null,
        })
        .mockResolvedValueOnce({
          data: { ...MS_TOKEN_RESPONSE, access_token: "token-b" },
          error: null,
        });

      const { auth } = await createAuthInstance({
        pluginOptions: {
          "app-a": { scope: ["https://graph.microsoft.com/.default"] },
          "app-b": { scope: ["https://storage.azure.com/.default"] },
        },
      });

      await auth.api.getClientCredentialToken({
        body: { applicationName: "app-a" },
      });
      await auth.api.getClientCredentialToken({
        body: { applicationName: "app-b" },
      });

      const scopeA = (mockFetch.mock.calls[0]?.[1]?.body as Record<string, string>).scope;
      const scopeB = (mockFetch.mock.calls[1]?.[1]?.body as Record<string, string>).scope;
      expect(scopeA).toBe("https://graph.microsoft.com/.default");
      expect(scopeB).toBe("https://storage.azure.com/.default");
    });
  });

  // ─── Token fetch — cache hit ───────────────────────────────────────────────

  describe("getClientCredentialToken — cache hit", () => {
    it("returns the cached token and skips the Microsoft request", async () => {
      mockSuccessfulTokenFetch();

      const { auth } = await createAuthInstance();

      // First call — populates cache
      await auth.api.getClientCredentialToken({
        body: { applicationName: "my-app" },
      });

      // Second call — must use cache
      const result = await auth.api.getClientCredentialToken({
        body: { applicationName: "my-app" },
      });

      expect(mockFetch).toHaveBeenCalledTimes(1);
      expect(result.accessToken).toBe("mock-access-token");
    });

    it("refreshes the token when cached entry falls within the 60 s expiry buffer", async () => {
      mockSuccessfulTokenFetch();

      const { auth } = await createAuthInstance();

      // Populate cache
      await auth.api.getClientCredentialToken({
        body: { applicationName: "my-app" },
      });

      // Move the stored expiry to 30 s from now (inside the 60 s buffer)
      const ctx = await auth.$context;
      const stored = await ctx.adapter.findOne<ClientCredential>({
        model: "clientCredential",
        where: [
          { field: "applicationName", value: "my-app", operator: "eq" },
        ],
      });
      await ctx.adapter.update<ClientCredential>({
        model: "clientCredential",
        where: [{ field: "id", value: stored!.id, operator: "eq" }],
        update: { accessTokenExpiresAt: new Date(Date.now() + 30_000) },
      });

      // Next call should detect near-expiry and refresh
      mockFetch.mockResolvedValueOnce({
        data: { ...MS_TOKEN_RESPONSE, access_token: "refreshed-token" },
        error: null,
      });

      const result = await auth.api.getClientCredentialToken({
        body: { applicationName: "my-app" },
      });

      expect(mockFetch).toHaveBeenCalledTimes(2);
      expect(result.accessToken).toBe("refreshed-token");
    });
  });

  // ─── Error cases ───────────────────────────────────────────────────────────

  describe("getClientCredentialToken — errors", () => {
    it("throws MICROSOFT_APPLICATION_NOT_FOUND when applicationName is not in options", async () => {
      const { auth } = await createAuthInstance();

      await expect(
        auth.api.getClientCredentialToken({
          body: { applicationName: "unknown-app" },
        }),
      ).rejects.toMatchObject({
        body: expect.objectContaining({
          code: "MICROSOFT_APPLICATION_NOT_FOUND",
        }),
      });

      // No token fetch should have been attempted
      expect(mockFetch).not.toHaveBeenCalled();
    });

    it("throws MICROSOFT_PROVIDER_NOT_CONFIGURED when microsoft provider is absent", async () => {
      const { auth } = await createAuthInstance({
        withMicrosoftProvider: false,
      });

      await expect(
        auth.api.getClientCredentialToken({
          body: { applicationName: "my-app" },
        }),
      ).rejects.toMatchObject({
        body: expect.objectContaining({
          code: "MICROSOFT_PROVIDER_NOT_CONFIGURED",
        }),
      });
    });

    it("throws MICROSOFT_TOKEN_REQUEST_FAILED when Microsoft returns an error", async () => {
      mockFailedTokenFetch();

      const { auth } = await createAuthInstance();

      await expect(
        auth.api.getClientCredentialToken({
          body: { applicationName: "my-app" },
        }),
      ).rejects.toMatchObject({
        body: expect.objectContaining({
          code: "MICROSOFT_TOKEN_REQUEST_FAILED",
        }),
      });
    });
  });
});

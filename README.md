# better-auth-client-credential

A [Better Auth](https://better-auth.com) plugin that implements the
[OAuth 2.0 client credentials flow](https://learn.microsoft.com/en-us/entra/identity-platform/v2-oauth2-client-creds-grant-flow)
against **Microsoft Entra ID** (formerly Azure AD).

The plugin fetches application-level access tokens on behalf of your server,
caches them in the Better Auth database, and refreshes them automatically
within a 60-second window before expiry. No user interaction is required.

## How it works

```
your server  →  auth.api.getClientCredentialToken()
                   │
                   ├─ cache hit (token valid for > 60 s)?  →  return cached token
                   │
                   └─ cache miss / near-expiry
                          │
                          ↓
                   POST /oauth2/v2.0/token  (Microsoft Entra ID)
                          │
                          ↓
                   store / refresh token in DB  →  return token
```

## Prerequisites

### 1. Entra ID app registration

In the [Azure portal](https://portal.azure.com), create (or reuse) an app
registration and:

- Add **application permissions** (not delegated) for each downstream API
  you want to call.
- Grant **admin consent** for those permissions.
- Create a **client secret** under *Certificates & secrets*.

You will need the **Tenant ID**, **Client ID**, and **Client Secret**.

### 2. Peer dependencies

```bash
npm install better-auth @better-fetch/fetch zod
```

### 3. Database migration

After adding the plugin, run the Better Auth CLI to create the
`clientCredential` table:

```bash
npx @better-auth/cli@latest migrate
# or, for Prisma / Drizzle:
npx @better-auth/cli@latest generate
```

Re-run this command whenever you update the plugin.

## Installation

```bash
npm install better-auth-client-credential
```

## Setup

### Configure `auth.ts`

Add the `microsoft` social provider and the `clientCredential` plugin to your
Better Auth instance. Each key in the plugin options is an **application name**
— a label you choose to identify a set of scopes. You can register as many
applications as you need.

```ts
// auth.ts
import { betterAuth } from "better-auth";
import { clientCredential } from "better-auth-client-credential";

export const auth = betterAuth({
  database: yourDatabaseAdapter,

  socialProviders: {
    microsoft: {
      tenantId: process.env.ENTRA_TENANT_ID!,
      clientId: process.env.ENTRA_CLIENT_ID!,
      clientSecret: process.env.ENTRA_CLIENT_SECRET!,
    },
  },

  plugins: [
    clientCredential({
      // Register one entry per downstream API.
      // The key becomes the `applicationName` you pass at call-time.
      "graph-api": {
        scope: ["https://graph.microsoft.com/.default"],
      },
      "my-internal-api": {
        scope: ["api://<downstream-app-client-id>/.default"],
      },
    }),
  ],
});
```

### Environment variables

```bash
ENTRA_TENANT_ID=<your-tenant-id>
ENTRA_CLIENT_ID=<your-client-id>
ENTRA_CLIENT_SECRET=<your-client-secret>
```

> The plugin reads these through Better Auth's `microsoft` social provider
> configuration — they are **not** read directly from `process.env`.

## Usage

`getClientCredentialToken` is a **server-only** endpoint. No HTTP route is
registered for it; call it directly via `auth.api` in your server code.

```ts
import { auth } from "./auth";

const credential = await auth.api.getClientCredentialToken({
  body: { applicationName: "graph-api" },
});

// Use the token in a downstream request
const response = await fetch("https://graph.microsoft.com/v1.0/users", {
  headers: {
    Authorization: `Bearer ${credential.accessToken}`,
  },
});
```

### Return value

The call returns a `ClientCredential` record:

| Field | Type | Description |
|---|---|---|
| `id` | `string` | Row primary key |
| `applicationName` | `string` | The name passed in the request body |
| `accessToken` | `string` | The Bearer token to send to the downstream API |
| `accessTokenExpiresAt` | `Date` | When `accessToken` expires |
| `scope` | `string` | Space-separated scopes that were granted |
| `createdAt` | `Date` | When this record was first written |
| `updatedAt` | `Date` | When the token was last refreshed |

### Caching behaviour

- A fresh token is fetched from Microsoft on the **first call** for each
  `applicationName`.
- Subsequent calls return the **cached token** with no network request, as long
  as `accessTokenExpiresAt` is more than **60 seconds** in the future.
- When the token falls inside that 60-second buffer, the plugin transparently
  fetches a new token and **updates the existing database row** — the record
  count stays at one per `applicationName`.

### Multiple applications

Each entry in the plugin options map is independent. You can fetch tokens for
different APIs in the same request:

```ts
const [graphToken, storageToken] = await Promise.all([
  auth.api.getClientCredentialToken({ body: { applicationName: "graph-api" } }),
  auth.api.getClientCredentialToken({ body: { applicationName: "storage-api" } }),
]);
```

## Error reference

| Code | HTTP status | Cause |
|---|---|---|
| `MICROSOFT_APPLICATION_NOT_FOUND` | 404 | `applicationName` is not registered in the plugin options |
| `MICROSOFT_PROVIDER_NOT_CONFIGURED` | 500 | The `microsoft` social provider is missing from the Better Auth config |
| `MICROSOFT_PROVIDER_INVALID_CONFIG` | 500 | The `microsoft` provider options are missing `tenantId`, `clientId`, or `clientSecret` |
| `MICROSOFT_TOKEN_REQUEST_FAILED` | 502 | The Microsoft token endpoint returned an error; the raw error body is included in the message |

## API reference

### `clientCredential(options)`

Creates the Better Auth plugin.

```ts
import { clientCredential } from "better-auth-client-credential";
```

**Parameters**

- `options` — `ClientCredentialOptions`  
  A map of application names to their scope configuration (see [Setup](#setup)).

**Returns** a `BetterAuthPlugin` to pass to the `plugins` array of `betterAuth`.

---

### `ClientCredentialOptions`

```ts
type ClientCredentialOptions = {
  [applicationName: string]: {
    scope: string[];
  };
};
```

---

### `ClientCredential`

The shape of the cached token record returned by `getClientCredentialToken` and
stored in the database.

```ts
type ClientCredential = {
  id: string;
  applicationName: string;
  accessToken: string;
  accessTokenExpiresAt: Date;
  scope: string;
  createdAt: Date;
  updatedAt: Date;
};
```

## Development

```bash
# Install dependencies
pnpm install

# Run unit + integration tests
pnpm test

# Type-check src and tests
pnpm typecheck

# Build the library
pnpm build
```

### Integration tests

The integration tests make real HTTP requests to Microsoft Entra ID. Copy
`.env.test.example` to `.env.test` and fill in your credentials:

```bash
VITE_ENTRA_TENANT_ID=<tenant-id>
VITE_ENTRA_CLIENT_ID=<client-id>
VITE_ENTRA_CLIENT_SECRET=<client-secret>
VITE_ENTRA_OBO_SCOPE=api://<downstream-app-client-id>/.default
```

If these variables are absent the integration tests are skipped automatically,
so `pnpm test` is always safe to run without credentials.

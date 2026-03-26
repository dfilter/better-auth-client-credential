import { createFetch, createSchema } from "@better-fetch/fetch";
import { defineErrorCodes, type BetterAuthPlugin } from "better-auth";
import { APIError, createAuthEndpoint } from "better-auth/api";
import { z } from "zod";

const inputSchema = z.object({
  tenant: z.string(),
  client_id: z.string(),
  scope: z.string(),
  client_secret: z.string(),
  grant_type: z.literal("client_credentials"),
});

const outputSchema = z.object({
  access_token: z.string(),
  expires_in: z.number(),
  token_type: z.string(),
});

const clientCredentialBodySchema = z.object({
  applicationName: z.string(),
});

const defaultErrorSchema = z.object({
  error: z.string(),
  error_description: z.string(),
  error_codes: z.number().array(),
  timestamp: z.coerce.date(),
  trace_id: z.string(),
  correlation_id: z.string(),
});

const msSocialProviderConfigSchema = z.object({
  tenantId: z.string(),
  clientId: z.string(),
  clientSecret: z.string(),
});

const CLIENT_CREDENTIAL_ERRORS = defineErrorCodes({
  MICROSOFT_PROVIDER_NOT_CONFIGURED:
    "Microsoft social provider is not configured",
  MICROSOFT_PROVIDER_INVALID_CONFIG:
    "Microsoft provider configuration is invalid",
  MICROSOFT_TOKEN_REQUEST_FAILED: "Failed to fetch Microsoft OAuth token",
  MICROSOFT_APPLICATION_NOT_FOUND: "Provided Application not found",
});

export type ClientCredentialOptions = {
  [applicationName: string]: {
    scope: string[];
  };
};

export type ClientCredential = {
  createdAt: Date;
  updatedAt: Date;
  id: string;
  applicationName: string;
  accessToken: string;
  accessTokenExpiresAt: Date;
  scope: string;
};

export const clientCredential = (options: ClientCredentialOptions) => {
  const $fetch = createFetch({
    defaultError: defaultErrorSchema,
    schema: createSchema({
      "@post/token": {
        input: inputSchema,
        output: outputSchema,
      },
    }),
  });

  return {
    id: "client-credential",
    $ERROR_CODES: CLIENT_CREDENTIAL_ERRORS,
    schema: {
      clientCredential: {
        modelName: "clientCredential",
        fields: {
          createdAt: {
            type: "date",
            required: true,
            defaultValue: () => new Date(),
          },
          updatedAt: {
            type: "date",
            required: true,
            defaultValue: () => new Date(),
            onUpdate: () => new Date(),
          },
          applicationName: {
            type: "string",
            required: true,
            index: true,
          },
          accessToken: {
            type: "string",
            required: true,
          },
          accessTokenExpiresAt: {
            type: "date",
            required: true,
          },
          scope: {
            type: "string",
            required: true,
          },
        },
      },
    },
    endpoints: {
      getClientCredentialToken: createAuthEndpoint(
        {
          method: "POST",
          body: clientCredentialBodySchema,
        },
        async ({
          body: { applicationName },
          context: { adapter, socialProviders },
        }) => {
          const scope = options[applicationName]?.scope.join(" ");
          if (!scope) {
            throw APIError.fromStatus(
              "NOT_FOUND",
              CLIENT_CREDENTIAL_ERRORS.MICROSOFT_APPLICATION_NOT_FOUND,
            );
          }

          const cachedClientCredential =
            await adapter.findOne<ClientCredential>({
              model: "clientCredential",
              where: [
                {
                  field: "applicationName",
                  value: applicationName,
                  operator: "eq",
                },
              ],
            });

          const refreshThreshold = Date.now() + 60_000;
          if (
            cachedClientCredential &&
            cachedClientCredential.accessTokenExpiresAt &&
            cachedClientCredential.accessTokenExpiresAt.getTime() >
              refreshThreshold
          ) {
            return cachedClientCredential;
          }

          const oAuthProvider = socialProviders.find(
            ({ id }) => id === "microsoft",
          );

          if (!oAuthProvider?.options) {
            throw APIError.from(
              "INTERNAL_SERVER_ERROR",
              CLIENT_CREDENTIAL_ERRORS.MICROSOFT_PROVIDER_NOT_CONFIGURED,
            );
          }

          const { error: msConfigError, data: msOptions } =
            msSocialProviderConfigSchema.safeParse(oAuthProvider.options);
          if (msConfigError) {
            throw APIError.from("INTERNAL_SERVER_ERROR", {
              code: CLIENT_CREDENTIAL_ERRORS.MICROSOFT_PROVIDER_INVALID_CONFIG
                .code,
              message: `${
                CLIENT_CREDENTIAL_ERRORS.MICROSOFT_PROVIDER_INVALID_CONFIG
                  .message
              } ${msConfigError.message}`,
            });
          }

          const tenantId = msOptions.tenantId;

          // The Microsoft token endpoint requires application/x-www-form-urlencoded.
          const { error, data } = await $fetch("@post/token", {
            baseURL: `https://login.microsoftonline.com/${tenantId}/oauth2/v2.0`,
            headers: { "Content-Type": "application/x-www-form-urlencoded" },
            body: {
              tenant: tenantId,
              client_id: msOptions.clientId,
              client_secret: msOptions.clientSecret,
              scope: scope,
              grant_type: "client_credentials",
            },
          });

          if (error) {
            throw APIError.from("BAD_GATEWAY", {
              code: CLIENT_CREDENTIAL_ERRORS.MICROSOFT_TOKEN_REQUEST_FAILED
                .code,
              message: `${CLIENT_CREDENTIAL_ERRORS.MICROSOFT_TOKEN_REQUEST_FAILED.message}\n${JSON.stringify(error)}`,
            });
          }

          const now = new Date();
          const accessTokenExpiresAt = new Date(
            Date.now() + data.expires_in * 1000,
          );

          if (cachedClientCredential) {
            const updated = await adapter.update<ClientCredential>({
              model: "clientCredential",
              where: [
                {
                  field: "id",
                  value: cachedClientCredential.id,
                  operator: "eq",
                },
              ],
              update: {
                accessToken: data.access_token,
                accessTokenExpiresAt,
                updatedAt: now,
              },
            });
            return (
              updated ?? {
                ...cachedClientCredential,
                accessToken: data.access_token,
                accessTokenExpiresAt,
                updatedAt: now,
              }
            );
          }

          const created = await adapter.create<ClientCredential>({
            model: "clientCredential",
            data: {
              createdAt: now,
              updatedAt: now,
              applicationName,
              accessToken: data.access_token,
              accessTokenExpiresAt,
              scope,
            },
          });

          return created;
        },
      ),
    },
  } satisfies BetterAuthPlugin;
};

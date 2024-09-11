import { Client } from "@microsoft/microsoft-graph-client";
import { ClientSecretCredential, DeviceCodeCredential } from "@azure/identity";
import { TokenCredentialAuthenticationProvider } from "@microsoft/microsoft-graph-client/authProviders/azureTokenCredentials";
import * as dotenv from "dotenv";
import { z } from "zod";

const envSchema = z.object({
  AZURE_TENANT_ID: z.string().min(1),
  AZURE_APP_ID: z.string().min(1),
  AZURE_AUTH_SECRET: z.string().min(1),
});

const main = async () => {
  // läd alles aus der .env in die process.env
  dotenv.config();

  const config = envSchema.parse(process.env);
  console.log(config);

  // @azure/identity
  // const credential = new ClientSecretCredential(
  //   config.AZURE_TENANT_ID,
  //   config.AZURE_APP_ID,
  //   config.AZURE_AUTH_SECRET
  // );

  const credential = new DeviceCodeCredential({
    clientId: config.AZURE_APP_ID,
    tenantId: config.AZURE_TENANT_ID,
    userPromptCallback: console.log,
  });

  // get auth token
  const token = await credential.getToken(["offline_access", "User.Read"]);
  // console.log(token);

  // @microsoft/microsoft-graph-client/authProviders/azureTokenCredentials
  const authProvider = new TokenCredentialAuthenticationProvider(credential, {
    scopes: ["offline_access", "User.Read", "Tasks.ReadWrite"],
  });

  // replace authProvider with another one that caches the token / uses the refreshtoken to automatically get new tokens

  const graphClient = Client.initWithMiddleware({ authProvider: authProvider });

  const taskListId =
    "AQMkADAwATM3ZmYAZS1lZQA2MC04NjM0LTAwAi0wMAoALgAAAyVIhJQRxJVPi0-75UZXF38BANpnlmN9-SNJrDJThYJa9skAAAGcOY8AAAA=";

  const response = await graphClient
    .api(`/me/todo/lists/${taskListId}/tasks`)
    .get();
  console.log(response);
};
main();

import { Client } from "@microsoft/microsoft-graph-client";
import { ClientSecretCredential } from "@azure/identity";
import { TokenCredentialAuthenticationProvider } from "@microsoft/microsoft-graph-client/authProviders/azureTokenCredentials";
import * as dotenv from "dotenv";
import { log } from "console";
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

  // @azure/identity
  const credential = new ClientSecretCredential(
    config.AZURE_TENANT_ID,
    config.AZURE_APP_ID,
    config.AZURE_AUTH_SECRET
  );

  // @microsoft/microsoft-graph-client/authProviders/azureTokenCredentials
  const authProvider = new TokenCredentialAuthenticationProvider(credential, {
    // The client credentials flow requires that you request the
    // /.default scope, and pre-configure your permissions on the
    // app registration in Azure. An administrator must grant consent
    // to those permissions beforehand.
    scopes: ["https://graph.microsoft.com/.default"],
  });

  const graphClient = Client.initWithMiddleware({ authProvider: authProvider });

  const baseURL = "https://graph.microsoft.com/v1.0";
  //   const userId = "krikelkr4kel_gmail.com#EXT#@krikelkr4kelgmail.onmicrosoft.com";
  const userId = "be8f005e-f2d2-4d0c-88a5-7cf452153413";
  const response = await graphClient
    .api(`${baseURL}/users/${userId}`)
    .get();
  log(response);

  // await graphClient.api(`${baseURL}/users/${userId}/todo/lists`).post({
  //     displayName: 'Travel items'
  //   })
};

main();

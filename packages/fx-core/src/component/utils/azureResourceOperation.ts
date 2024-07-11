// Copyright (c) Microsoft Corporation.
// Licensed under the MIT license.

import { AzureResourceInfo } from "../driver/interface/commonArgs";
import { ExternalApiCallError, PrerequisiteError } from "../error/componentError";
import { AzureAccountProvider } from "@microsoft/teamsfx-api";
import { BlobServiceClient } from "@azure/storage-blob";
import { StorageAccounts, StorageManagementClient, AccountSasParameters } from "@azure/arm-storage";
import { DeployConstant } from "../constant/deployConstant";
import { TokenCredential } from "@azure/identity";
import { wrapAzureOperation } from "./azureSdkErrorHandler";

/**
 * parse Azure resource id into subscriptionId, resourceGroupName and resourceName
 * @param resourceId Azure resource id
 * @param pattern the pattern that used to parse resource id and extract info from it
 */
export function parseAzureResourceId(resourceId: string, pattern: RegExp): AzureResourceInfo {
  const result = resourceId.trim().match(pattern);
  if (!result || result.length != 4) {
    throw PrerequisiteError.somethingIllegal("Deploy", "resourceId", "plugins.bot.InvalidValue", [
      "resourceId",
      resourceId,
    ]);
  }
  return {
    subscriptionId: result[1].trim(),
    resourceGroupName: result[2].trim(),
    instanceId: result[3].trim(),
  };
}

/**
 * get Azure credential from Azure account provider
 * @param tokenProvider Azure account provider
 */
export async function getAzureAccountCredential(
  tokenProvider: AzureAccountProvider
): Promise<TokenCredential> {
  const credential = await wrapAzureOperation(
    () => tokenProvider.getIdentityCredentialAsync(),
    (e) => ExternalApiCallError.getAzureCredentialRemoteError(DeployConstant.DEPLOY_ERROR_TYPE, e),
    (e) => ExternalApiCallError.getAzureCredentialError(DeployConstant.DEPLOY_ERROR_TYPE, e)
  );
  if (!credential) {
    throw PrerequisiteError.somethingIllegal(
      "Deploy",
      "azureCredential",
      "plugin.hosting.FailRetrieveAzureCredentials",
      undefined,
      "plugin.hosting.LoginToAzure"
    );
  }
  return credential;
}

/**
 * create Azure Storage Blob Service Client
 * @param azureResource azure resource info
 * @param azureCredential azure user credential
 */
export function createBlobServiceClient(
  azureResource: AzureResourceInfo,
  azureCredential: TokenCredential
): BlobServiceClient {
  const blobUri = getBlobUri(azureResource.instanceId);
  return getBlobServiceClient(blobUri, azureCredential);
}

function getBlobUri(storageName: string): string {
  return `https://${storageName}.blob.core.windows.net`;
}

function getBlobServiceClient(blobUri: string, token: TokenCredential): BlobServiceClient {
  return new BlobServiceClient(blobUri, token);
}

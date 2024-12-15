import 'isomorphic-fetch';
import { ClientSecretCredential } from '@azure/identity';
import { Client } from '@microsoft/microsoft-graph-client';
import { TokenCredentialAuthenticationProvider } from '@microsoft/microsoft-graph-client/authProviders/azureTokenCredentials';
import settings from './appSettings';

let _clientSecretCredential: ClientSecretCredential | undefined = undefined;
let _appClient: Client | undefined = undefined;

export function initializeGraphForAppOnlyAuth() {
  if (!settings) {
    throw new Error('Settings cannot be undefined');
  }

  _clientSecretCredential = new ClientSecretCredential(
    settings.tenantId,
    settings.clientId,
    settings.clientSecret
  );

  const authProvider = new TokenCredentialAuthenticationProvider(_clientSecretCredential, {
    scopes: ['https://graph.microsoft.com/.default'],
  });

  _appClient = Client.initWithMiddleware({ authProvider });
}

export async function getUsers() {
  if (!_appClient) {
    throw new Error('Graph has not been initialized for app-only auth');
  }

  return _appClient
    .api('/users')
    .select(['displayName', 'id', 'mail'])
    .top(25)
    .orderby('displayName')
    .get();
}

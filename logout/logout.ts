/*
 * Copyright (c) Microsoft Corporation. All rights reserved. Licensed under the MIT license.
 * See LICENSE in the project root for license information.
 */

import { PublicClientApplication } from '@azure/msal-browser';
import settings from '../appSettings';

let pca;

Office.onReady(async () => {
  Office.context.ui.addHandlerAsync(Office.EventType.DialogParentMessageReceived,
    onMessageFromParent);
  pca = new PublicClientApplication({
    auth: {
      clientId: settings.clientId,
      authority: 'https://login.microsoftonline.com/common',
      redirectUri: `${window.location.origin}/login/login.html` // Must be registered as "spa" type.
    },
    cache: {
      cacheLocation: 'localStorage' // Needed to avoid a "login required" error.
    }
  });
  await pca.initialize();
});

async function onMessageFromParent(arg) {
  const messageFromParent = JSON.parse(arg.message);
  const currentAccount = pca.getAccountByHomeId(messageFromParent.userName);
  // You can select which account application should sign out.
  const logoutRequest = {
    account: currentAccount,
    postLogoutRedirectUri: `${window.location.origin}/logoutcomplete/logoutcomplete.html`,
  };

  try {
    // Attempt silent token retrieval
    const silentResponse = await pca.acquireTokenSilent({
      scopes: ['user.read', 'files.read.all'],
      account: currentAccount
    });
    console.log('Silent token retrieval successful:', silentResponse);
  } catch (silentError) {
    // Silent token retrieval failed, invoke login
    console.error('Silent token retrieval failed:', silentError);
    await pca.loginRedirect({
      scopes: ['user.read', 'files.read.all']
    });
  }

  await pca.logoutRedirect(logoutRequest);
  const messageObject = { messageType: "dialogClosed" };
  const jsonMessage = JSON.stringify(messageObject);
  Office.context.ui.messageParent(jsonMessage);
}

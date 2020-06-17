// Copyright (c) Microsoft Corporation.
// Licensed under the MIT License.

var graph = require('@microsoft/microsoft-graph-client');
require('isomorphic-fetch');
const fs = require("fs");
const open =  require("open");

// addEventListener("load", this.uploadFile);

module.exports = {
  getUserDetails: async function(accessToken) {
    const client = getAuthenticatedClient(accessToken);

    const user = await client.api('/me').get();
    return user;
  },

  uploadFile: async function(accessToken) {
    try {
      const client = getAuthenticatedClient(accessToken);
      let readStream = fs.createReadStream("./ExcelWorkbookWithTaskPane.xlsx");
      client
      .api('/me/drive/root/children/ExcelWorkbookWithTaskPane.xlsx/content')
      .putStream(readStream)
      .then((res) => {
          console.log(res.webUrl);
          open(res.webUrl);
      })
    } catch (err) {
      console.log(err);
    }
  },

  // <GetEventsSnippet>
  getEvents: async function(accessToken) {
    const client = getAuthenticatedClient(accessToken);

    const events = await client
      .api('/me/events')
      .select('subject,organizer,start,end')
      .orderby('createdDateTime DESC')
      .get();

    return events;
  }
  // </GetEventsSnippet>
};

function getAuthenticatedClient(accessToken) {
  // Initialize Graph client
  const client = graph.Client.init({
    // Use the provided access token to authenticate
    // requests
    authProvider: (done) => {
      done(null, accessToken);
    }
  });

  return client;
}
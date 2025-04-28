// RUN LOCAL TESTS
// npx tsx debug.ts

import * as spauth from 'node-sp-auth';
import request from 'request-promise';  // Change this import
import * as dotenv from 'dotenv';
dotenv.config();

const clientId = process.env.SHAREPOINT_CLIENT_ID || '';
const secret = process.env.SHAREPOINT_CLIENT_SECRET || '';
const tenantID = process.env.SHAREPOINT_TENANT_ID || '';
const url = process.env.SHAREPOINT_SITE_URL || '';

console.log("Starting Calculator MCP Server...");

// Use async/await instead of promise chaining
const authData = await spauth.getAuth(url, {
  clientId: clientId,
  clientSecret: secret,
  realm: tenantID
});

var headers = authData.headers;
headers['Accept'] = 'application/json;odata=verbose';

// Use request-promise correctly
const response = await request({
  url: `${url}/_api/web`,
  headers: headers,
  json: true,
  method: 'GET'
});

console.log("SharePoint site title: ", response.d.Title);
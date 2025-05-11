
# SharePoint Online MCP Server

This is a MCP server for Claude Desktop that allows you to interact with SharePoint Online using the SharePoint REST API. It is designed to be used with the [Claude Desktop](https://claude.ai/download) app, but could be used by other MCP clients as well.

## Implementation

| Component          | Operation                  | Available |
|--------------------|----------------------------|-----------|
| Users              |                            | ✅        |
|                    | List Site Users by Role    | ✅        |
|                    | List Site Groups           | ✅        |
|                    | Get Group Members          | ✅        |
|                    | Add Group Member           | ✅        |
|                    | Remove Group Member        | ✅        |
| Sites              |                            | ✅        |
|                    | Get Site Details           | ✅        |
|                    | Update Site Properties     | ✅        |
|                    | Get Global Navigation Links| ✅        |
|                    | Get Quick Navigation Links | ✅        |
|                    | Add Navigation Link        | ✅        |
|                    | Update Navigation Link     | ✅        |
|                    | Delete Navigation Link     | ✅        |
|                    | Get Subsites               | ✅        |
|                    | Delete Subsite             | ✅        |
| Lists              |                            | ✅        |
|                    | Get All Lists              | ✅        |
|                    | Create List                | ✅        |
|                    | Get List Fields            | ✅        |
|                    | Create List Field          | ✅        |
|                    | Update List Field          | ✅        |
|                    | Delete List Field          | ✅        |
|                    | Set List Field formatting  | ❌        |
|                    | Set List View formatting   | ❌        |
|                    | Get List Views             | ✅        |
|                    | Create List View           | ✅        |
|                    | Update List View           | ✅        |
|                    | Delete List View           | ✅        |
|                    | Delete List                | ✅        |
| List Views         |                            | ✅        |
|                    | Get View Fields            | ✅        |
|                    | Add View Field             | ✅        |
|                    | Remove View Field          | ✅        |
|                    | Remove All View Fields     | ✅        |
|                    | Move View Field            | ✅        |
| List Items         |                            | ✅        |
|                    | Get All List Items         | ✅        |
|                    | Create List Item           | ✅        |
|                    | Update List Item           | ✅        |
|                    | Delete List Item           | ✅        |

## Configuration steps

⚠️ You need to be a SharePoint tenant admin or a global admin to create a SharePoint app.

### Allow SharePoint-only apps

- Download and install the SharePoint Online Management Shell from [here](https://www.microsoft.com/en-ca/download/details.aspx?id=35588).
- Open PowerShell and run the following command to make sure the SharePoint-only apps are enabled:

```powershell
Connect-SPOService -Url https://<your-tenant>-admin.sharepoint.com/
Set-SPOTenant -DisableCustomAppAuthentication $false
```

### Create SharePoint-only App

- Open `https://<your-tenant>.sharepoint.com/_layouts/15/appregnew.aspx`
- Click `Generate` to create a new client ID and secret.
- Fill in the following fields:
  - App Domain: `www.example.com` (or any domain you own)
  - Redirect URL: `http://example.com`
![App registration](IMG/image-3.png)
- Click `Create` to create the app.
- Copy the `Client ID` and `Client Secret` values to a safe place. You will need them later.

### Grant App Permissions

- Open `https://<your-tenant>.sharepoint.com/_layouts/15/appinv.aspx`
- Paste the `Client ID` value in the `App ID` field and click `Lookup`.
![Lookup](IMG/image-4.png)
- Paste the following XML in the `App Permission Request XML` field:

```xml
<AppPermissionRequests AllowAppOnlyPolicy="true">
   <AppPermissionRequest Scope="http://sharepoint/content/tenant" Right="FullControl" />
</AppPermissionRequests>
```

- Click `Create` and then `Trust It` to grant the app permissions.

![Trust it](IMG/image-5.png)

### Install Node.js

 Install Node.js version: [22.14.0 or later](https://nodejs.org/en/download)

### Install and Configure Claude Desktop

- Download [Claude Desktop](https://claude.ai/download) and install it.
- In Claude Desktop, go to `File` > `Settings` > `Developer`.
![File > Settings](IMG/image.png)
- Click Edit Config
![Config](IMG/image-1.png)
- Open the claude_desktop_config.json config file in the editor.
- Paste the following to the config and update the `mcpServers` section with your own values:

```json
{
  "mcpServers": {
    "server-sharepoint": {
      "command": "npx",
      "args": [
        "-y",
        "server-sharepoint"
      ],
      "env": {
        "SHAREPOINT_CLIENT_ID": "<your-client-id>",
        "SHAREPOINT_CLIENT_SECRET": "<your-client-secret>",
        "SHAREPOINT_TENANT_ID": "<your-tenant-id>"
      }
    }
  }
} 

```

> 💡[Learn how to get the tenant ID from Entra ID](https://learn.microsoft.com/en-us/entra/fundamentals/how-to-find-tenant)

- Save the config file and restart Claude Desktop. Make sure to quit the app completely via the tasks tray icon:
![Quick Claude App](IMG/image-2.png)

## How to "talk" to the MCP Server

- Open Claude Desktop and ask a question like `Get me the list of lists from https://<your-tenant>.sharepoint.com/sites/Dashboard-Communication. Present the results in a table form. don't use code`.

![Sample result](IMG/image-6.png)

## Available Tools

### Site Management
- `getSite` - Get detailed information about a SharePoint site
- `updateSite` - Update site properties including Title, Description, and Logo URL
- `getSubsites` - Get all subsites from a site
- `deleteSubsite` - Delete a subsite (requires confirmation)
- `getGlobalNavigationLinks` - Get global (top) navigation links
- `getQuickNavigationLinks` - Get quick (left) navigation links
- `addNavigationLink` - Add a navigation link to global or quick navigation
- `updateNavigationLink` - Update an existing navigation link
- `deleteNavigationLink` - Delete a navigation link
- `getSiteUsers` - Get users from a site with optional filtering by role (Owners, Members, Visitors)
- `getSiteGroups` - Get all SharePoint groups for a site

### List Management
- `getLists` - Get all available lists on a SharePoint site
- `createList` - Create a new list or document library with custom settings and optional custom URL
- `deleteList` - Delete a list or document library (requires confirmation)
- `getListFields` - Get detailed information about fields/columns in a list
- `createListField` - Create a new field/column with a two-step process: first with a clean name for proper internal name generation, then updating to the desired display name with spaces
- `updateListField` - Update a field including display name, choices, and other properties
- `deleteListField` - Delete a field/column from a list (requires confirmation)
- `getListViews` - Get all views for a list with optional field details
- `createListView` - Create a custom view with selected fields and filters
- `updateListView` - Update an existing view's settings and fields
- `deleteListView` - Delete a view from a list (cannot delete default views)

### List View Field Management
- `getViewFields` - Get all fields in a specific list view
- `addViewField` - Add a field to a list view
- `removeViewField` - Remove a field from a list view
- `removeAllViewFields` - Remove all fields from a list view
- `moveViewFieldTo` - Move a field to a specific position in a list view

### List Item Management
- `getListItems` - Get all items from a specific list
- `createListItem` - Create a new item with custom field values
- `updateListItem` - Update an existing item's field values
- `deleteListItem` - Remove an item from a list

### User Management
- `getGroupMembers` - Get all members of a specific SharePoint group
- `addGroupMember` - Add a user to a SharePoint group
- `removeGroupMember` - Remove a user from a SharePoint group

## For contributors

- Build a new npm package

```bash
npm run build
npm publish
npx clear-npx-cache
npm cache clean --force
```

## Run MCP Inspector locally

```bash
npx @modelcontextprotocol/inspector node build/src/index.js
```

## MCP Server Context 7 Documentation for LLMs

[https://context7.com/modelcontextprotocol/servers](MCP Server Context 7 Documentation for LLMs)

## Local Development

- Clone the repo
- `nvm use 22.14.0`
- `npm install`
- Update MCP config in `claude_desktop_config.json` file. Sample:

```json
{
"mcpServers": {
  "sharepoint": {
      "command": "node",
      "args": [
        "C:\\trash\\server-sharepoint\\build\\index.js"
      ],
      "env": {
        "SHAREPOINT_CLIENT_ID": "aaa2ce0f-7c97-4e51-acd9-4ac99e6831d0",
        "SHAREPOINT_CLIENT_SECRET": "bbb0OFF+aWkua2NJRERUMGduMmhsWDhGUTVPUHJZLXd5dldSemNPUw==",
        "SHAREPOINT_TENANT_ID": "22bcd981-eaa6-475f-aac3-3afca0c35261"
      }
    }
  }
}
```

- Make changes to the code
- Build the code `npm run build`

- `npx @modelcontextprotocol/inspector node build/src/index.js -y`
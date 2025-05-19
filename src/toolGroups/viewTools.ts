// src/toolGroups/viewTools.ts
import { McpServer } from "@modelcontextprotocol/sdk/server/mcp.js";
import { z } from "zod";
import { SharePointConfig } from '../config';

// Import view-related tools
import {
    getListViews,
    createListView,
    updateListView,
    deleteListView,
    getViewFields,
    addViewField,
    removeViewField,
    removeAllViewFields,
    moveViewFieldTo,
    // Types
    GetListViewsParams,
    CreateListViewParams,
    UpdateListViewParams,
    DeleteListViewParams,
    GetViewFieldsParams,
    AddViewFieldParams,
    RemoveViewFieldParams,
    RemoveAllViewFieldsParams,
    MoveViewFieldToParams
} from '../tools';

/**
 * Register view management tools with the MCP server
 * @param server The MCP server instance
 * @param config SharePoint configuration
 * @returns void
 */
export function registerViewTools(server: McpServer, config: SharePointConfig): void {
    console.error("Registering view management tools...");

    // Add getListViews tool
    server.tool(
        "getListViews",
        "Get all views from a SharePoint list with optional field details",
        {
            url: z.string().url().describe("URL of the SharePoint website"),
            listTitle: z.string().describe("Title of the SharePoint list to retrieve views from"),
            includeFields: z.boolean().optional().describe("Whether to include the fields for each view"),
            includeHidden: z.boolean().optional().describe("Whether to include hidden views")
        },
        async (params: GetListViewsParams) => {
            return await getListViews(params, config);
        }
    );

    // Add createListView tool
    server.tool(
        "createListView",
        "Create a new view for a SharePoint list with specified fields and settings",
        {
            url: z.string().url().describe("URL of the SharePoint website"),
            listTitle: z.string().describe("Title of the SharePoint list"),
            viewData: z.object({
                Title: z.string().describe("Title of the new view"),
                ViewQuery: z.string().optional().describe("CAML query for filtering items"),
                RowLimit: z.number().int().optional().describe("Maximum number of items to display per page"),
                ViewFields: z.array(z.string()).optional().describe("Array of field internal names to include in the view"),
                PersonalView: z.boolean().optional().describe("Whether this is a personal view"),
                SetAsDefaultView: z.boolean().optional().describe("Whether to set this as the default view")
            }).describe("Properties for the new view")
        },
        async (params: CreateListViewParams) => {
            return await createListView(params, config);
        }
    );

    // Add updateListView tool
    server.tool(
        "updateListView",
        "Update an existing view for a SharePoint list",
        {
            url: z.string().url().describe("URL of the SharePoint website"),
            listTitle: z.string().describe("Title of the SharePoint list"),
            viewTitle: z.string().describe("Title of the view to update"),
            updateData: z.object({
                Title: z.string().optional().describe("New title for the view"),
                ViewQuery: z.string().optional().describe("New CAML query for filtering items"),
                RowLimit: z.number().int().optional().describe("New maximum number of items to display per page"),
                ViewFields: z.array(z.string()).optional().describe("New array of field internal names to include in the view"),
                PersonalView: z.boolean().optional().describe("Whether this is a personal view"),
                SetAsDefaultView: z.boolean().optional().describe("Whether to set this as the default view")
            }).describe("Properties to update in the view"),
            appendFields: z.boolean().optional().describe("Whether to append fields to existing ones instead of replacing them")
        },
        async (params: UpdateListViewParams) => {
            return await updateListView(params, config);
        }
    );
    
    // Add deleteListView tool
    server.tool(
        "deleteListView",
        "Delete a view from a SharePoint list",
        {
            url: z.string().url().describe("URL of the SharePoint website"),
            listTitle: z.string().describe("Title of the SharePoint list"),
            viewTitle: z.string().describe("Title of the view to delete")
        },
        async (params: DeleteListViewParams) => {
            return await deleteListView(params, config);
        }
    );

    // Add getViewFields tool
    server.tool(
        "getViewFields",
        "Get all fields from a specific SharePoint list view",
        {
            url: z.string().url().describe("URL of the SharePoint website"),
            listTitle: z.string().describe("Title of the SharePoint list"),
            viewTitle: z.string().describe("Title of the view to get fields from")
        },
        async (params: GetViewFieldsParams) => {
            return await getViewFields(params, config);
        }
    );

    // Add addViewField tool
    server.tool(
        "addViewField",
        "Add a field to a SharePoint list view",
        {
            url: z.string().url().describe("URL of the SharePoint website"),
            listTitle: z.string().describe("Title of the SharePoint list"),
            viewTitle: z.string().describe("Title of the view to modify"),
            fieldName: z.string().describe("Internal name of the field to add to the view")
        },
        async (params: AddViewFieldParams) => {
            return await addViewField(params, config);
        }
    );

    // Add removeViewField tool
    server.tool(
        "removeViewField",
        "Remove a field from a SharePoint list view",
        {
            url: z.string().url().describe("URL of the SharePoint website"),
            listTitle: z.string().describe("Title of the SharePoint list"),
            viewTitle: z.string().describe("Title of the view to modify"),
            fieldName: z.string().describe("Internal name of the field to remove from the view")
        },
        async (params: RemoveViewFieldParams) => {
            return await removeViewField(params, config);
        }
    );

    // Add removeAllViewFields tool
    server.tool(
        "removeAllViewFields",
        "Remove all fields from a SharePoint list view",
        {
            url: z.string().url().describe("URL of the SharePoint website"),
            listTitle: z.string().describe("Title of the SharePoint list"),
            viewTitle: z.string().describe("Title of the view to modify")
        },
        async (params: RemoveAllViewFieldsParams) => {
            return await removeAllViewFields(params, config);
        }
    );

    // Add moveViewFieldTo tool
    server.tool(
        "moveViewFieldTo",
        "Move a field to a specific position in a SharePoint list view",
        {
            url: z.string().url().describe("URL of the SharePoint website"),
            listTitle: z.string().describe("Title of the SharePoint list"),
            viewTitle: z.string().describe("Title of the view to modify"),
            fieldName: z.string().describe("Internal name of the field to move"),
            index: z.number().int().min(0).describe("New position index (0-based)")
        },
        async (params: MoveViewFieldToParams) => {
            return await moveViewFieldTo(params, config);
        }
    );
}

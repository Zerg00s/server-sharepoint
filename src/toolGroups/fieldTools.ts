// src/toolGroups/fieldTools.ts
import { McpServer } from "@modelcontextprotocol/sdk/server/mcp.js";
import { z } from "zod";
import { SharePointConfig } from '../config';

// Import field-related tools
import {
    createListField,
    updateListField,
    deleteListField,
    // Types
    CreateListFieldParams,
    UpdateListFieldParams,
    DeleteListFieldParams
} from '../tools';

/**
 * Register field management tools with the MCP server
 * @param server The MCP server instance
 * @param config SharePoint configuration
 * @returns void
 */
export function registerFieldTools(server: McpServer, config: SharePointConfig): void {
    console.error("Registering field management tools...");

    // Add createListField tool
    server.tool(
        "createListField",
        "Create a new field (column) in a SharePoint list",
        {
            url: z.string().url().describe("URL of the SharePoint website"),
            listTitle: z.string().describe("Title of the SharePoint list"),
            fieldData: z.object({
                Title: z.string().describe("Display name for the field (can contain spaces)"),
                CleanName: z.string().optional().describe("Clean name without spaces (used for internal name generation)"),
                FieldTypeKind: z.number().int().describe(
                  "Field type value: 0=Invalid, 1=Integer, 2=Text, 3=Note, 4=DateTime, 5=Choice, 6=Lookup, " +
                  "7=Boolean (according to docs, but may not work), 8=Boolean, 9=Number, 10=Currency, 11=URL, " +
                  "15=MultiChoice, 17=Calculated, 19=User"),
                Required: z.boolean().optional().describe("Whether the field is required"),
                EnforceUniqueValues: z.boolean().optional().describe("Whether the field must have unique values"),
                StaticName: z.string().optional().describe("Static name, if not provided will be generated from CleanName or Title"),
                Description: z.string().optional().describe("Description for the field"),
                Choices: z.array(z.string()).optional().describe("For choice fields (FieldTypeKind=5) or MultiChoice fields (FieldTypeKind=15)"),
                DefaultValue: z.union([z.string(), z.number(), z.boolean()]).optional().describe("Default value for the field")
            }).passthrough().describe("Properties for the new field")
        },
        async (params: CreateListFieldParams) => {
            return await createListField(params, config);
        }
    );

    // Add updateListField tool
    server.tool(
        "updateListField",
        "Update a field/column in a SharePoint list including display name, choices, etc.",
        {
            url: z.string().url().describe("URL of the SharePoint website"),
            listTitle: z.string().describe("Title of the SharePoint list"),
            fieldInternalName: z.string().describe("Internal name of the field to update"),
            updateData: z.object({
                Title: z.string().optional().describe("New display name for the field"),
                Description: z.string().optional().describe("New description for the field"),
                Required: z.boolean().optional().describe("Whether the field is required"),
                EnforceUniqueValues: z.boolean().optional().describe("Whether the field must have unique values"),
                Choices: z.array(z.string()).optional().describe("New choices for choice fields"),
                DefaultValue: z.string().optional().describe("New default value for the field")
            }).describe("Field properties to update")
        },
        async (params: UpdateListFieldParams) => {
            return await updateListField(params, config);
        }
    );

    // Add deleteListField tool
    server.tool(
        "deleteListField",
        "Delete a field (column) from a SharePoint list",
        {
            url: z.string().url().describe("URL of the SharePoint website"),
            listTitle: z.string().describe("Title of the SharePoint list"),
            fieldInternalName: z.string().describe("Internal name of the field to delete"),
            confirmation: z.string().describe("Confirmation string that must match the field internal name exactly")
        },
        async (params: DeleteListFieldParams) => {
            return await deleteListField(params, config);
        }
    );
}

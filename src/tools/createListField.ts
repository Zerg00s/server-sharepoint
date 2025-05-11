// src/tools/createListField.ts
import request from 'request-promise';
import { IToolResult, IFieldUpdateData } from '../interfaces';
import { getSharePointHeaders, getRequestDigest } from '../auth_factory';
import { SharePointConfig } from '../config';

export interface CreateListFieldParams {
    url: string;
    listTitle: string;
    fieldData: {
        Title: string; // Display name for the field (can contain spaces)
        CleanName?: string; // Clean name without spaces (used for internal name generation)
        FieldTypeKind: number; // Field type values: 0=Invalid, 1=Integer, 2=Text, 3=Note, 4=DateTime, 5=Choice, 6=Lookup, 7=Boolean (docs), 8=Boolean, 9=Number, 10=Currency, 11=URL, 15=MultiChoice, 17=Calculated, 19=User
        Required?: boolean;
        EnforceUniqueValues?: boolean;
        StaticName?: string; // Static name, if not provided will be generated from CleanName or Title
        Description?: string;
        Choices?: string[]; // For choice fields
        DefaultValue?: string | number | boolean;
        [key: string]: any; // Any other field properties
    };
}

/**
 * Create a new field (column) in a SharePoint list
 * @param params Parameters including site URL, list title, and field data
 * @param config SharePoint configuration
 * @returns Tool result with creation status and new field info
 */
export async function createListField(
    params: CreateListFieldParams,
    config: SharePointConfig
): Promise<IToolResult> {
    const { url, listTitle, fieldData } = params;
    console.error(`createListField tool called with URL: ${url}, List Title: ${listTitle}, Field Title: ${fieldData.Title}`);

    try {
        // Validate required parameters
        if (!fieldData.Title) {
            throw new Error("Field Title is required");
        }
        
        if (fieldData.FieldTypeKind === undefined || fieldData.FieldTypeKind === null) {
            throw new Error("FieldTypeKind is required");
        }
        
        // Authenticate with SharePoint
        const headers = await getSharePointHeaders(url, config);
        console.error("Headers prepared with authentication");

        // Get request digest for POST operations
        headers['X-RequestDigest'] = await getRequestDigest(url, headers);
        headers['Content-Type'] = 'application/json;odata=verbose';
        console.error("Headers prepared with request digest");

        // Encode the list title to handle special characters
        const encodedListTitle = encodeURIComponent(listTitle);
        
        // First, verify the list exists
        console.error(`Verifying list "${listTitle}" exists...`);
        try {
            await request({
                url: `${url}/_api/web/lists/getByTitle('${encodedListTitle}')`,
                headers: { ...headers, 'Content-Type': undefined },
                json: true,
                method: 'GET',
                timeout: 15000
            });
        } catch (error) {
            throw new Error(`List "${listTitle}" not found`);
        }
        
        // Check if a field with the same title already exists
        console.error(`Checking if field "${fieldData.Title}" already exists...`);
        try {
            const fieldsResponse = await request({
                url: `${url}/_api/web/lists/getByTitle('${encodedListTitle}')/fields?$filter=Title eq '${fieldData.Title}'`,
                headers: { ...headers, 'Content-Type': undefined },
                json: true,
                method: 'GET',
                timeout: 30000
            });
            
            if (fieldsResponse.d.results && fieldsResponse.d.results.length > 0) {
                throw new Error(`A field with title "${fieldData.Title}" already exists in list "${listTitle}"`);
            }
        } catch (err: unknown) {
            // If the error is due to the field not being found, we can continue
            // Otherwise, re-throw the error
            const errorMessage = err instanceof Error ? err.message : String(err);
            if (errorMessage && !errorMessage.includes("already exists")) {
                throw err;
            }
        }
        
        // Determine the initial title to use (for internal name generation)
        // If CleanName is provided, use it for the initial creation
        const originalTitle = fieldData.Title;
        const temporaryTitle = fieldData.CleanName || fieldData.Title.replace(/\s/g, '');
        
        console.error(`Using temporary title "${temporaryTitle}" for initial field creation to control internal name generation`);
        
        // Variable to hold the final field data
        let finalField: any;
        
        // Handle choice fields using CreateFieldAsXml for better reliability
        if (fieldData.FieldTypeKind === 5 || fieldData.FieldTypeKind === 15) { // Choice or MultiChoice
            if (!fieldData.Choices || !Array.isArray(fieldData.Choices) || fieldData.Choices.length === 0) {
                throw new Error("Choices array is required for Choice and MultiChoice fields");
            }
            
            // Use the CreateFieldAsXml endpoint instead of the standard field creation
            console.error(`Creating Choice field using CreateFieldAsXml approach...`);
            
            // Build CHOICES XML element
            let choicesXml = "<CHOICES>";
            for (const choice of fieldData.Choices) {
                choicesXml += `<CHOICE>${choice}</CHOICE>`;
            }
            choicesXml += "</CHOICES>";
            
            // Build the full SchemaXml for the field
            const fieldDisplayName = originalTitle;
            const fieldName = temporaryTitle.replace(/\s/g, '');
            
            // Set the field type based on FieldTypeKind
            const fieldType = fieldData.FieldTypeKind === 5 ? 'Choice' : 'MultiChoice';
            
            const schemaXml = `<Field DisplayName='${fieldDisplayName}' FillInChoice='FALSE' IsModern='TRUE' Name='${fieldName}' Title='${fieldDisplayName}' Type='${fieldType}'>${choicesXml}</Field>`;
            
            const createXmlPayload = {
                parameters: {
                    __metadata: { type: "SP.XmlSchemaFieldCreationInformation" },
                    SchemaXml: schemaXml,
                    Options: 12
                }
            };
            
            console.error(`Creating field with XML payload: ${JSON.stringify(createXmlPayload)}`);
            
            try {
                // Use the CreateFieldAsXml endpoint
                const createXmlResponse = await request({
                    url: `${url}/_api/web/lists/getByTitle('${encodedListTitle}')/fields/CreateFieldAsXml`,
                    headers: headers,
                    json: true,
                    method: 'POST',
                    body: createXmlPayload,
                    timeout: 30000
                });
                
                console.error(`Field created successfully with CreateFieldAsXml`);
                finalField = createXmlResponse.d;
                
                // Skip the standard field creation for Choice fields
                const newField = {
                    Title: finalField.Title,
                    InternalName: finalField.InternalName,
                    StaticName: finalField.StaticName,
                    Type: finalField.TypeAsString,
                    Id: finalField.Id,
                    Required: finalField.Required,
                    Description: finalField.Description || ''
                };
                
                return {
                    content: [{
                        type: "text",
                        text: JSON.stringify({
                            success: true,
                            message: `${fieldType} field "${originalTitle}" successfully created in list "${listTitle}" with internal name "${newField.InternalName}"`,
                            newField: newField
                        }, null, 2)
                    }]
                } as IToolResult;
            } catch (xmlError) {
                console.error(`Error creating ${fieldType} field with CreateFieldAsXml: ${xmlError instanceof Error ? xmlError.message : String(xmlError)}`);
                console.error("Falling back to standard field creation method");
                // Continue with standard field creation as fallback
            }
        }
        
        // Standard field creation approach for non-Choice fields or as fallback
        // Prepare the initial field creation payload - use CleanName or space-less Title for internal name generation
        const createPayload: any = {
            __metadata: { type: 'SP.Field' },
            Title: temporaryTitle,
            FieldTypeKind: fieldData.FieldTypeKind
        };
        
        // Add optional parameters if provided
        if (fieldData.Required !== undefined) {
            createPayload.Required = fieldData.Required;
        }
        
        if (fieldData.EnforceUniqueValues !== undefined) {
            createPayload.EnforceUniqueValues = fieldData.EnforceUniqueValues;
        }
        
        // If StaticName is provided, include it
        if (fieldData.StaticName) {
            createPayload.StaticName = fieldData.StaticName;
        }
        
        if (fieldData.Description) {
            createPayload.Description = fieldData.Description;
        }
        
        if (fieldData.DefaultValue !== undefined) {
            createPayload.DefaultValue = fieldData.DefaultValue.toString();
        }
        
        // For choice fields (if CreateFieldAsXml failed)
        if (fieldData.FieldTypeKind === 5 || fieldData.FieldTypeKind === 15) { 
            createPayload.Choices = {
                __metadata: { type: 'Collection(Edm.String)' },
                results: fieldData.Choices
            };
        }
        
        // Add any other properties from fieldData except Title/CleanName which we're handling specially
        Object.entries(fieldData).forEach(([key, value]) => {
            if (!createPayload[key] && key !== 'Choices' && key !== 'Title' && key !== 'CleanName') {
                createPayload[key] = value;
            }
        });
        
        console.error(`Creating field with payload: ${JSON.stringify(createPayload)}`);
        
        // Step 1: Create the field with the temporary title
        const createResponse = await request({
            url: `${url}/_api/web/lists/getByTitle('${encodedListTitle}')/Fields`,
            headers: headers,
            json: true,
            method: 'POST',
            body: createPayload,
            timeout: 30000
        });
        
        console.error(`Field created with internal name "${createResponse.d.InternalName}"`);
        
        // Start with the response data
        finalField = createResponse.d;
        
        // Check if we need to update the title (if we used a temporary title)
        if (temporaryTitle !== originalTitle) {
            console.error(`Updating field title from "${temporaryTitle}" to original title "${originalTitle}"`);
            
            // Get a fresh request digest for the update operation
            const updateHeaders = { ...headers };
            updateHeaders['X-RequestDigest'] = await getRequestDigest(url, headers);
            updateHeaders['IF-MATCH'] = '*';
            updateHeaders['X-HTTP-Method'] = 'MERGE';
            
            // Step 2: Update the field title to the desired value with spaces
            const updatePayload = {
                __metadata: { type: 'SP.Field' },
                Title: originalTitle
            };
            
            // Encode the internal name to handle special characters
            const encodedInternalName = encodeURIComponent(createResponse.d.InternalName);
            
            await request({
                url: `${url}/_api/web/lists/getByTitle('${encodedListTitle}')/fields/getByInternalNameOrTitle('${encodedInternalName}')`,
                headers: updateHeaders,
                json: true,
                method: 'POST',
                body: updatePayload,
                timeout: 20000
            });
            
            // Get the updated field information
            const getHeaders = { ...headers };
            // Remove any headers that could cause issues with GET
            delete getHeaders['X-HTTP-Method'];
            delete getHeaders['IF-MATCH'];
            
            const updatedFieldResponse = await request({
                url: `${url}/_api/web/lists/getByTitle('${encodedListTitle}')/fields/getByInternalNameOrTitle('${encodedInternalName}')`,
                headers: { ...getHeaders, 'Content-Type': undefined },
                json: true,
                method: 'GET',
                timeout: 15000
            });
            
            finalField = updatedFieldResponse.d;
        }
        
        // Get the new field details
        const newField = {
            Title: finalField.Title,
            InternalName: finalField.InternalName,
            StaticName: finalField.StaticName,
            Type: finalField.TypeAsString,
            Id: finalField.Id,
            Required: finalField.Required,
            Description: finalField.Description || ''
        };
        
        return {
            content: [{
                type: "text",
                text: JSON.stringify({
                    success: true,
                    message: `Field "${originalTitle}" successfully created in list "${listTitle}" with internal name "${newField.InternalName}"`,
                    newField: newField
                }, null, 2)
            }]
        } as IToolResult;
    } catch (error: unknown) {
        // Type-safe error handling
        let errorMessage: string;

        if (error instanceof Error) {
            errorMessage = error.message;
        } else if (typeof error === 'string') {
            errorMessage = error;
        } else {
            errorMessage = "Unknown error occurred";
        }

        console.error("Error in createListField tool:", errorMessage);

        return {
            content: [{
                type: "text",
                text: `Error creating list field: ${errorMessage}`
            }],
            isError: true
        } as IToolResult;
    }
}

export default createListField;

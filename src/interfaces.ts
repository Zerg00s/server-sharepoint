// src/interfaces.ts

/**
 * Interface representing the raw SharePoint list API response
 */
export interface ISharePointListResponse {
  Title: string;
  Id: string;
  ItemCount: number;
  LastItemModifiedDate: string;
  Description: string;
  BaseTemplate: number;
  Hidden: boolean;
  IsSystemList: boolean;
  RootFolder: {
    ServerRelativeUrl: string;
  };
  [key: string]: any; // For any other fields that might be present
}

/**
 * Interface representing a processed SharePoint list
 */
export interface IList {
  Title: string;
  URL: string;
  ItemCount: number;
  LastModified: string;
  Description: string;
  BaseTemplateID: number;
}

/**
 * Interface representing a SharePoint content type
 */
export interface ISharePointContentType {
  Id: {
    StringValue: string;
  };
  Name: string;
  Group: string;
  Description: string;
  Hidden: boolean;
  ReadOnly: boolean;
  Sealed: boolean;
  [key: string]: any; // For any other fields that might be present
}

/**
 * Interface representing a processed SharePoint content type
 */
export interface IContentType {
  Id: string;
  Name: string;
  Group: string;
  Description: string;
  Hidden: boolean;
  ReadOnly: boolean;
  Sealed: boolean;
}

/**
 * Interface representing a SharePoint list item
 */
export interface ISharePointListItem {
  ID: number;
  Title?: string;
  Created: string;
  Modified: string;
  AuthorId?: number;
  EditorId?: number;
  GUID: string;
  [key: string]: any; // Allow for dynamic custom fields
}

/**
 * Interface representing a formatted list item for return
 */
export interface IFormattedListItem {
  [key: string]: any; // Dynamic fields based on the list schema
}

/**
 * Interface for the response from the SharePoint web API
 */
export interface ISharePointWebResponse {
  d: {
    Title: string;
    [key: string]: any;
  };
}

/**
 * Interface for the response from creating a list item
 */
export interface ISharePointCreateItemResponse {
  d: ISharePointListItem;
}

/**
 * Interface for SharePoint field definition
 */
export interface ISharePointField {
  InternalName: string;
  Title: string;
  TypeAsString: string;
  TypeDisplayName?: string;
  ReadOnlyField: boolean;
  Hidden: boolean;
  LookupList?: string;
  LookupField?: string;
  Choices?: {
    results: string[];
  };
  [key: string]: any;
}

/**
 * Interface for lookup field data
 */
export interface ILookupFieldValue {
  ID: number;
  Value: string;
}

/**
 * Interface for SharePoint lookup data collection
 */
export interface ILookupData {
  [fieldName: string]: ILookupFieldValue[];
}

/**
 * Interface for mock data generation result
 */
export interface IMockDataResult {
  listTitle: string;
  writeableFields: {
    name: string;
    title: string;
    type: string;
  }[];
  lookupFields: {
    name: string;
    valuesFound: number;
  }[];
  requested: number;
  successful: number;
  failed: number;
  successfulItems: number[];
  failedItems: { index: number, error: string }[];
}

/**
 * Interface for tool function result content item (text)
 */
export interface IToolResultTextContent {
  type: "text";
  text: string;
  [key: string]: unknown;
}

/**
 * Interface for tool function result content item (image)
 */
export interface IToolResultImageContent {
  type: "image";
  data: string;
  mimeType: string;
  [key: string]: unknown;
}

/**
 * Interface for tool function result content item (audio)
 */
export interface IToolResultAudioContent {
  type: "audio";
  data: string;
  mimeType: string;
  [key: string]: unknown;
}

/**
 * Interface for tool function result content item (resource with text)
 */
export interface IToolResultResourceTextContent {
  type: "resource";
  resource: {
    uri: string;
    text: string;
    mimeType?: string;
    [key: string]: unknown;
  };
  [key: string]: unknown;
}

/**
 * Interface for tool function result content item (resource with blob)
 */
export interface IToolResultResourceBlobContent {
  type: "resource";
  resource: {
    uri: string;
    blob: string;
    mimeType?: string;
    [key: string]: unknown;
  };
  [key: string]: unknown;
}

/**
 * Union type for all possible content item types
 */
export type ToolResultContentItem = 
  | IToolResultTextContent 
  | IToolResultImageContent 
  | IToolResultAudioContent 
  | IToolResultResourceTextContent 
  | IToolResultResourceBlobContent;

/**
 * Interface for tool function result
 */
export interface IToolResult {
  content: ToolResultContentItem[];
  isError?: boolean;
  meta?: Record<string, unknown>;
  [key: string]: unknown;
}

/**
 * Interface for functions that implement tool functionality
 */
export interface IToolFunction<T> {
  (params: T): Promise<IToolResult>;
}

/**
 * Interface for SharePoint list fields response
 */
export interface ISharePointListFieldsResponse {
  d: {
    results: ISharePointField[];
  };
}

/**
 * Interface for list field update parameters
 */
export interface IFieldUpdateData {
  Title?: string;
  Description?: string;
  Required?: boolean;
  EnforceUniqueValues?: boolean;
  Choices?: string[];
  DefaultValue?: string;
  [key: string]: any;
}

/**
 * Interface for list creation parameters
 */
export interface IListCreationData {
  Title: string;
  Description?: string;
  TemplateType?: number; // e.g., 100 for generic list, 101 for document library
  Url?: string; // Relative URL for the list (used in browser URLs)
  ContentTypesEnabled?: boolean;
  AllowContentTypes?: boolean;
  EnableVersioning?: boolean;
  EnableMinorVersions?: boolean;
  EnableModeration?: boolean;
  [key: string]: any;
}

/**
 * Interface for list view parameters
 */
export interface IListViewData {
  Title?: string;
  ViewQuery?: string;
  RowLimit?: number;
  ViewFields?: string[];
  PersonalView?: boolean;
  SetAsDefaultView?: boolean;
  [key: string]: any;
}

/**
 * Interface for create list view parameters (Title is required)
 */
export interface ICreateListViewData extends Omit<IListViewData, 'Title'> {
  Title: string;
}

/**
 * Interface for SharePoint user
 */
export interface ISharePointUser {
  Id: number;
  Title: string;
  Email: string;
  LoginName: string;
  IsSiteAdmin?: boolean;
  [key: string]: any;
}

/**
 * Interface for SharePoint group
 */
export interface ISharePointGroup {
  Id: number;
  Title: string;
  Description?: string;
  OwnerTitle?: string;
  AllowMembersEditMembership?: boolean;
  OnlyAllowMembersViewMembership?: boolean;
  [key: string]: any;
}

/**
 * Interface for SharePoint group member
 */
export interface ISharePointGroupMember {
  Id: number;
  Title: string;
  Email?: string;
  LoginName: string;
  [key: string]: any;
}
// Existing interfaces would be here...

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

export interface IFormattedListItem {
  [key: string]: any; // Dynamic fields based on the list schema
}
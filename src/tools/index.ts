// src/tools/index.ts
import getSite from './getSite';
import getLists from './getLists';
import getListItems from './getListItems';
import addMockData from './addMockData';

// Export all tools
export {
    getSite as getSite,
    getLists,
    getListItems,
    addMockData
};

// Also export the parameter interfaces for better type safety
export type { GetSiteParams } from './getSite';
export type { GetListsParams } from './getLists';
export type { GetListItemsParams } from './getListItems';
export type { AddMockDataParams } from './addMockData';
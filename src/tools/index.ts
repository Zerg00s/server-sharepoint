// src/tools/index.ts
import getTitle from './getTitle';
import getLists from './getLists';
import getListItems from './getListItems';
import addMockData from './addMockData';

// Export all tools
export {
    getTitle,
    getLists,
    getListItems,
    addMockData
};

// Also export the parameter interfaces for better type safety
export type { GetTitleParams } from './getTitle';
export type { GetListsParams } from './getLists';
export type { GetListItemsParams } from './getListItems';
export type { AddMockDataParams } from './addMockData';
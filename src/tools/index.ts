// src/tools/index.ts
import getSite from './getSite';
import getLists from './getLists';
import getListItems from './getListItems';
import getListFields from './getListFields';
import updateListField from './updateListField';
import updateListItem from './updateListItem';
import createListItem from './createListItem';
import batchCreateListItems from './batchCreateListItems';
import batchUpdateListItems from './batchUpdateListItems';
import batchDeleteListItems from './batchDeleteListItems';
import createList from './createList';
import createListView from './createListView';
import updateListView from './updateListView';
import deleteListItem from './deleteListItem';
import getSiteUsers from './getSiteUsers';
import getSiteGroups from './getSiteGroups';
import addGroupMember from './addGroupMember';
import removeGroupMember from './removeGroupMember';
import getListViews from './getListViews';
import deleteListView from './deleteListView';
import deleteList from './deleteList';
import createListField from './createListField';
import deleteListField from './deleteListField';
import getGroupMembers from './getGroupMembers';
import getGlobalNavigationLinks from './getGlobalNavigationLinks';
import getQuickNavigationLinks from './getQuickNavigationLinks';
import getSubsites from './getSubsites';
import deleteSubsite from './deleteSubsite';
import updateSite from './updateSite';
import updateList from './updateList';
import addNavigationLink from './addNavigationLink';
import updateNavigationLink from './updateNavigationLink';
import deleteNavigationLink from './deleteNavigationLink';
import getViewFields from './getViewFields';
import addViewField from './addViewField';
import removeViewField from './removeViewField';
import removeAllViewFields from './removeAllViewFields';
import moveViewFieldTo from './moveViewFieldTo';
import createModernPage from './createModernPage';
import getModernPages from './getModernPages';
import getModernPage from './getModernPage'; // Tool for retrieving a specific page
import deleteModernPage from './deleteModernPage';
// Import content type tools
import getListContentTypes from './getListContentTypes';
import createListContentType from './createListContentType';
import updateListContentType from './updateListContentType';
import deleteListContentType from './deleteListContentType';

// Export all tools
export {
    getSite,
    getLists,
    getListItems,
    getListFields,
    updateListField,
    updateListItem,
    createListItem,
    batchCreateListItems,
    batchUpdateListItems,
    batchDeleteListItems,
    createList,
    createListView,
    updateListView,
    deleteListItem,
    getSiteUsers,
    getSiteGroups,
    addGroupMember,
    removeGroupMember,
    getListViews,
    deleteListView,
    deleteList,
    createListField,
    deleteListField,
    getGroupMembers,
    getGlobalNavigationLinks,
    getQuickNavigationLinks,
    getSubsites,
    deleteSubsite,
    updateSite,
    updateList,
    addNavigationLink,
    updateNavigationLink,
    deleteNavigationLink,
    // New tools for view field management
    getViewFields,
    addViewField,
    removeViewField,
    removeAllViewFields,
    moveViewFieldTo,
    // Page management tools
    createModernPage,
    getModernPages,
    getModernPage,
    deleteModernPage,
    // Content type management tools
    getListContentTypes,
    createListContentType,
    updateListContentType,
    deleteListContentType
};

// Also export the parameter interfaces for better type safety
export type { GetSiteParams } from './getSite';
export type { GetListsParams } from './getLists';
export type { GetListItemsParams } from './getListItems';
export type { GetListFieldsParams } from './getListFields';
export type { UpdateListFieldParams } from './updateListField';
export type { UpdateListItemParams } from './updateListItem';
export type { CreateListItemParams } from './createListItem';
export type { BatchCreateListItemsParams } from './batchCreateListItems';
export type { BatchUpdateListItemsParams } from './batchUpdateListItems';
export type { BatchDeleteListItemsParams } from './batchDeleteListItems';
export type { CreateListParams } from './createList';
export type { CreateListViewParams } from './createListView';
export type { UpdateListViewParams } from './updateListView';
export type { DeleteListItemParams } from './deleteListItem';
export type { GetSiteUsersParams } from './getSiteUsers';
export type { GetSiteGroupsParams } from './getSiteGroups';
export type { AddGroupMemberParams } from './addGroupMember';
export type { RemoveGroupMemberParams } from './removeGroupMember';
export type { GetListViewsParams } from './getListViews';
export type { DeleteListViewParams } from './deleteListView';
export type { DeleteListParams } from './deleteList';
export type { CreateListFieldParams } from './createListField';
export type { DeleteListFieldParams } from './deleteListField';
export type { GetGroupMembersParams } from './getGroupMembers';
export type { GetGlobalNavigationLinksParams } from './getGlobalNavigationLinks';
export type { GetQuickNavigationLinksParams } from './getQuickNavigationLinks';
export type { GetSubsitesParams } from './getSubsites';
export type { DeleteSubsiteParams } from './deleteSubsite';
export type { UpdateSiteParams } from './updateSite';
export type { UpdateListParams } from './updateList';
export type { AddNavigationLinkParams } from './addNavigationLink';
export type { UpdateNavigationLinkParams } from './updateNavigationLink';
export type { DeleteNavigationLinkParams } from './deleteNavigationLink';
// New tool parameter interfaces
export type { GetViewFieldsParams } from './getViewFields';
export type { AddViewFieldParams } from './addViewField';
export type { RemoveViewFieldParams } from './removeViewField';
export type { RemoveAllViewFieldsParams } from './removeAllViewFields';
export type { MoveViewFieldToParams } from './moveViewFieldTo';
// Page management tool params
export type { CreateModernPageParams } from './createModernPage';
export type { GetModernPagesParams } from './getModernPages';
export type { GetModernPageParams } from './getModernPage';


export type { DeleteModernPageParams } from './deleteModernPage';
// Content type management tool params
export type { GetListContentTypesParams } from './getListContentTypes';
export type { CreateListContentTypeParams } from './createListContentType';
export type { UpdateListContentTypeParams } from './updateListContentType';
export type { DeleteListContentTypeParams } from './deleteListContentType';

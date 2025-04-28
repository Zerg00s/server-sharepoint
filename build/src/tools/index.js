"use strict";
var __importDefault = (this && this.__importDefault) || function (mod) {
    return (mod && mod.__esModule) ? mod : { "default": mod };
};
Object.defineProperty(exports, "__esModule", { value: true });
exports.addMockData = exports.getListItems = exports.getLists = exports.getTitle = void 0;
// src/tools/index.ts
const getTitle_1 = __importDefault(require("./getTitle"));
exports.getTitle = getTitle_1.default;
const getLists_1 = __importDefault(require("./getLists"));
exports.getLists = getLists_1.default;
const getListItems_1 = __importDefault(require("./getListItems"));
exports.getListItems = getListItems_1.default;
const addMockData_1 = __importDefault(require("./addMockData"));
exports.addMockData = addMockData_1.default;

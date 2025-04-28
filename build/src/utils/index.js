"use strict";
var __importDefault = (this && this.__importDefault) || function (mod) {
    return (mod && mod.__esModule) ? mod : { "default": mod };
};
Object.defineProperty(exports, "__esModule", { value: true });
exports.generateMockValueForField = exports.mockDataGenerator = void 0;
// src/utils/index.ts
var mockDataGenerator_1 = require("./mockDataGenerator");
Object.defineProperty(exports, "mockDataGenerator", { enumerable: true, get: function () { return __importDefault(mockDataGenerator_1).default; } });
Object.defineProperty(exports, "generateMockValueForField", { enumerable: true, get: function () { return mockDataGenerator_1.generateMockValueForField; } });

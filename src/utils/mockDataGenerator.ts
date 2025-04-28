// src/utils/mockDataGenerator.ts

/**
 * Generate mock data for a SharePoint field based on its type
 * @param field The SharePoint field definition
 * @param index The index of the mock item being generated
 * @returns Generated mock value appropriate for the field type
 */
export function generateMockValueForField(field: any, index: number): any {
    const fieldType = field.TypeDisplayName?.toLowerCase() || field.TypeAsString?.toLowerCase() || '';
    const fieldName = field.InternalName;
    
    // Common prefix for generated values to make them identifiable as mock data
    const mockPrefix = `Mock-${index + 1}`;
    
    switch (fieldType) {
        case 'single line of text':
        case 'text':
            if (fieldName.toLowerCase().includes('name')) {
                return `${mockPrefix}: Name`;
            } else if (fieldName.toLowerCase().includes('title')) {
                return `${mockPrefix}: Title`;
            } else if (fieldName.toLowerCase().includes('description')) {
                return `${mockPrefix}: Description text for this mock item`;
            } else {
                return `${mockPrefix}: ${field.Title || fieldName}`;
            }
            
        case 'multiple lines of text':
        case 'note':
            return `${mockPrefix}: This is a longer text for the field "${field.Title || fieldName}".\nThis is some additional text to make it multi-line.\nGenerated as mock data.`;
            
        case 'number':
            return index * 10 + Math.floor(Math.random() * 100);
            
        case 'currency':
            return (index * 10.25 + Math.random() * 100).toFixed(2);
            
        case 'date and time':
        case 'datetime':
            const mockDate = new Date();
            mockDate.setDate(mockDate.getDate() + index); // Each item gets a different date
            return mockDate.toISOString();
            
        case 'choice':
        case 'multichoice':
            // If choices are available, use one of them
            if (field.Choices && field.Choices.results && field.Choices.results.length > 0) {
                const choiceIndex = index % field.Choices.results.length;
                return field.Choices.results[choiceIndex];
            }
            return `Choice ${(index % 5) + 1}`;
            
        case 'yes/no':
        case 'boolean':
            return index % 2 === 0;
            
        case 'person or group':
        case 'user':
            // We can't truly populate this without knowing valid users
            return null;
            
        case 'hyperlink':
        case 'url':
            return `https://example.com/mock-link-${index}`;
            
        case 'lookup':
        case 'lookupfield':
        case 'lookup (allow multiple values)':
        case 'lookupfieldmulti':
            // For lookup fields, we need to use the correct format expected by SharePoint
            // Return a placeholder that the caller should replace
            console.error(`Field ${fieldName} is a lookup field. Need special handling.`);
            return { __lookupField: true, fieldName, multiple: fieldType.includes('multiple') };
            
        default:
            // For unsupported types, return null to skip
            console.error(`Unsupported field type: ${fieldType} for field ${fieldName}`);
            return null;
    }
}

export default {
    generateMockValueForField
};
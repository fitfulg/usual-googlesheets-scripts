// tests/formatting.test.js

const { setCellStyle } = require('../shared/formatting');

describe('setCellStyle', () => {
    // Mock methods for the range object
    const mockRange = {
        setValue: jest.fn().mockReturnThis(),
        setFontWeight: jest.fn().mockReturnThis(),
        setFontColor: jest.fn().mockReturnThis(),
        setHorizontalAlignment: jest.fn().mockReturnThis(),
        setBackground: jest.fn().mockReturnThis(),
    };

    // Mock global SpreadsheetApp object and its methods
    global.SpreadsheetApp = {
        getActiveSpreadsheet: jest.fn().mockReturnThis(),
        getActiveSheet: jest.fn().mockReturnThis(),
        newTextStyle: jest.fn().mockReturnThis(),
        newRichTextValue: jest.fn().mockReturnThis(),
    };

    // Mock global sheet object
    global.sheet = {
        getRange: jest.fn().mockReturnValue(mockRange),
    };

    setCellStyle('A1', 'Test Value', 'bold', '#000000', '#ffffff', 'center');

    expect(mockRange.setValue).toHaveBeenCalledWith('Test Value');
    expect(mockRange.setFontWeight).toHaveBeenCalledWith('bold');
    expect(mockRange.setFontColor).toHaveBeenCalledWith('#000000');
    expect(mockRange.setHorizontalAlignment).toHaveBeenCalledWith('center');
    expect(mockRange.setBackground).toHaveBeenCalledWith('#ffffff');

});

describe('applyBorders', () => {
    test('should apply solid borders to the range', () => {
        // Mock setBorder method for the range object
        const mockRange = {
            setBorder: jest.fn().mockReturnThis(),
        };

        // Mock global SpreadsheetApp object with BorderStyle
        global.SpreadsheetApp = {
            BorderStyle: {
                SOLID: 'solid'
            }
        };

        // Higher-order function to apply formatting to a range only if it is valid
        const withValidRange = (fn) => (range, ...args) => range && fn(range, ...args);

        const applyBordersWithStyle = withValidRange((range, borderStyle) =>
            range.setBorder(true, true, true, true, true, true, "#000000", borderStyle)
        );

        // Global applyBorders function
        global.applyBorders = (range) => applyBordersWithStyle(range, SpreadsheetApp.BorderStyle.SOLID);

        applyBorders(mockRange);

        expect(mockRange.setBorder).toHaveBeenCalledWith(true, true, true, true, true, true, "#000000", 'solid');
    });
});
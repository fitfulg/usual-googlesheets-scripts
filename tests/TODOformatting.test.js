const { exampleTextTODO } = require('../TODOsheet/TODOformatting');

describe('Google Sheets Functions from TODOsheet/TODOformatting.js', () => {
    let mockRange, mockSheet, mockDataRange;

    beforeEach(() => {
        // Mocking the methods for the range object
        mockRange = {
            getValue: jest.fn(),
            getValues: jest.fn(),
            setValue: jest.fn(),
            setBorder: jest.fn(),
            setBackground: jest.fn()
        };

        // Mocking the global sheet object
        mockSheet = {
            getRange: jest.fn().mockReturnValue(mockRange),
            getMaxRows: jest.fn().mockReturnValue(10)
        };

        // Mocking the global SpreadsheetApp object and its methods
        global.SpreadsheetApp = {
            getActiveSpreadsheet: jest.fn().mockReturnThis(),
            getActiveSheet: jest.fn().mockReturnThis(),
            newTextStyle: jest.fn().mockReturnThis(),
            newRichTextValue: jest.fn().mockReturnThis(),
            BorderStyle: {
                SOLID: 'solid',
                SOLID_MEDIUM: 'solid_medium'
            },
            getUi: jest.fn().mockReturnValue({
                alert: jest.fn()
            })
        };

        global.sheet = mockSheet;

        // Mocking the getDataRange method
        mockDataRange = {
            getLastRow: jest.fn().mockReturnValue(10),
        };

        global.getDataRange = jest.fn().mockReturnValue(mockDataRange);
    });

    describe('exampleTextTODO', () => {
        test('should set example text for column B if cells are empty', () => {
            // Setup mock values
            mockRange.getValue.mockReturnValue('');
            mockRange.getValues.mockReturnValue([[''], [''], [''], [''], [''], [''], [''], [''], ['']]);

            exampleTextTODO('B', 'Example Text');

            expect(mockSheet.getRange).toHaveBeenCalledWith('B2');
            expect(mockSheet.getRange).toHaveBeenCalledWith('B4:B7');
            expect(mockSheet.getRange).toHaveBeenCalledWith('B9:B10');
            expect(mockRange.setValue).toHaveBeenCalledWith('Example Text');
        });

        test('should not set example text for column B if cells are not empty', () => {
            // Setup mock values
            mockRange.getValue.mockReturnValue('Existing Value');
            mockRange.getValues.mockReturnValue([['Existing Value'], ['Existing Value'], ['Existing Value'], ['Existing Value'], ['Existing Value'], ['Existing Value'], ['Existing Value'], ['Existing Value'], ['Existing Value']]);

            exampleTextTODO('B', 'Example Text');

            expect(mockRange.setValue).not.toHaveBeenCalled();
        });

        test('should set example text for other columns if cells are empty', () => {
            // Setup mock values
            mockRange.getValues.mockReturnValue([[''], [''], [''], [''], [''], [''], [''], [''], ['']]);

            exampleTextTODO('C', 'Example Text');

            expect(mockSheet.getRange).toHaveBeenCalledWith('C2:C10');
            expect(mockRange.setValue).toHaveBeenCalledWith('Example Text');
        });

        test('should not set example text for other columns if cells are not empty', () => {
            // Setup mock values
            mockRange.getValues.mockReturnValue([['Existing Value'], ['Existing Value'], ['Existing Value'], ['Existing Value'], ['Existing Value'], ['Existing Value'], ['Existing Value'], ['Existing Value'], ['Existing Value']]);

            exampleTextTODO('C', 'Example Text');

            expect(mockRange.setValue).not.toHaveBeenCalled();
        });
    });
});

const { updateCellCommentTODO,
    exampleTextTODO,
    applyFormatToAllTODO,
    checkAndSetColumnTODO,
    setColumnBackground,
    customCellBGColorTODO,
    setCellContentAndStyleTODO,
    setupDropdownTODO,
    pushUpEmptyCellsTODO,
    updateRichTextTODO,
    shiftCellsUpTODO } = require('../TODOsheet/TODOformatting');

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

        // Mocking the Logger object
        global.Logger = {
            log: jest.fn()  // Mock the log method
        };
    });

    describe('Function Existence', () => {
        test('updateCellCommentTODO function exists', () => {
            expect(typeof updateCellCommentTODO).toBe('function');
        });
        test('exampleTextTODO function exists', () => {
            expect(typeof exampleTextTODO).toBe('function');
        });
        test('applyFormatToAllTODO function exists', () => {
            expect(typeof applyFormatToAllTODO).toBe('function');
        });
        test('checkAndSetColumnTODO function exists', () => {
            expect(typeof checkAndSetColumnTODO).toBe('function');
        });
        test('setColumnBackground function exists', () => {
            expect(typeof setColumnBackground).toBe('function');
        });
        test('customCellBGColorTODO function exists', () => {
            expect(typeof customCellBGColorTODO).toBe('function');
        });
        test('setCellContentAndStyleTODO function exists', () => {
            expect(typeof setCellContentAndStyleTODO).toBe('function');
        });
        test('setupDropdownTODO function exists', () => {
            expect(typeof setupDropdownTODO).toBe('function');
        });
        test('pushUpEmptyCellsTODO function exists', () => {
            expect(typeof pushUpEmptyCellsTODO).toBe('function');
        });
        test('updateRichTextTODO function exists', () => {
            expect(typeof updateRichTextTODO).toBe('function');
        });
        test('shiftCellsUpTODO function exists', () => {
            expect(typeof shiftCellsUpTODO).toBe('function');
        });
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

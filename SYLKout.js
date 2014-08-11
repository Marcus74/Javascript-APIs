// Author: Mark Huff II 8.8.2014
// Description: This javascript library is used to create an SYLK file.

/* 
The MIT License (MIT)

Copyright (c) 2014 Mark C. Huff II

Permission is hereby granted, free of charge, to any person obtaining a copy
of this software and associated documentation files (the "Software"), to deal
in the Software without restriction, including without limitation the rights
to use, copy, modify, merge, publish, distribute, sublicense, and/or sell
copies of the Software, and to permit persons to whom the Software is
furnished to do so, subject to the following conditions:

The above copyright notice and this permission notice shall be included in all
copies or substantial portions of the Software.

THE SOFTWARE IS PROVIDED "AS IS", WITHOUT WARRANTY OF ANY KIND, EXPRESS OR
IMPLIED, INCLUDING BUT NOT LIMITED TO THE WARRANTIES OF MERCHANTABILITY,
FITNESS FOR A PARTICULAR PURPOSE AND NONINFRINGEMENT. IN NO EVENT SHALL THE
AUTHORS OR COPYRIGHT HOLDERS BE LIABLE FOR ANY CLAIM, DAMAGES OR OTHER
LIABILITY, WHETHER IN AN ACTION OF CONTRACT, TORT OR OTHERWISE, ARISING FROM,
OUT OF OR IN CONNECTION WITH THE SOFTWARE OR THE USE OR OTHER DEALINGS IN THE
SOFTWARE.
*/
/* EXAMPLE:
        SYLKout.init();
        SYLKout.addHeaderColumnRow(["Buyer", "Name", "Address"]);
        SYLKout.addDataRow(["1223344498765", "Mark", "Somewhere St."]);
        SYLKout.setColumnFormat(SYLKout.cellFormatingCodes.GENERAL, 0, SYLKout.cellAlignmentCodes.RIGHT_JUSTIFY,1);
        SYLKout.closeSpreadsheet();
        console.log(SYLKout.getSpreadsheetfile());
*/


var SYLKout = {
    mySpreadsheetoutput: 'ID;P\n' + 
                         'P;PGeneral\n' +
                         'P;P_(* #,##0_);;_(* \-#,##0_);;_(* "-"_);;_(@_)\n',
    headerX: 1,
    headerY: 1,
    rowX: 1,
    rowY: 2,
    columnFormattingCodes: {
        GENERAL: "P0",
        NUMBER_NO_DECIMAL: "P1",
    },
    cellFormatingCodes: {
        DEFAULT: "D",
        CONTINUOUS: "C",
        SCIENTIFIC: "E",
        FIXED_DECIMAL_POINT: "F",
        GENERAL: "G"
    },
    cellAlignmentCodes: {
        DEFAULT: "D",
        CENTER: "C",
        GENERAL: "G",
        LEFT_JUSTIFY: "L",
        RIGHT_JUSTIFY: "R"
    },
    init: function() {
        this.mySpreadsheetoutput = 'ID;P\n' +
                                   'P;PGeneral\n' +
                                   'P;P_(* #,##0_);;_(* \-#,##0_);;_(* "-"_);;_(@_)\n';
    },
    getSpreadsheetfile: function() {
        return this.mySpreadsheetoutput;
    },
    addHeaderColumn: function(newHeaderColumnName) {
        this.mySpreadsheetoutput = this.mySpreadsheetoutput.concat('C;Y' + this.headerY + ';X' + this.headerX + ';K"' + newHeaderColumnName + '"\n');
        this.headerX++;
    },
    addHeaderColumnRow: function (newHeaderColumnsNames) {
        for (var i = 0; i < newHeaderColumnsNames.length; i++) {
            this.mySpreadsheetoutput = this.mySpreadsheetoutput.concat('C;Y' + this.headerY + ';X' + this.headerX++ + ';K"' + newHeaderColumnsNames[i] + '"\n');
        }
    },
    addDataCell: function(cellData, cellFormattingCode, cellNumberOfDigits, cellAlignmentCode, x, y) {
        this.mySpreadsheetoutput = this.mySpreadsheetoutput.concat('C;Y' + y + ';X' + x + ';K"' + cellData + '"\n');
        this.mySpreadsheetoutput = this.mySpreadsheetoutput.concat('F;Y' + y + ';X' + x + ';F' + cellFormattingCode + cellNumberOfDigits + cellAlignmentCode + '\n');
    },
    addDataRow: function (rowData) {
        console.log(this.rowX);
        for (var i = 0; i < rowData.length; i++) {
            this.mySpreadsheetoutput = this.mySpreadsheetoutput.concat('C;Y' + this.rowY + ';X' + (this.rowX+i) + ';K"' + rowData[i] + '"\n');
        }
        this.rowY++;
    },
    // This seems to be buggy in Excel 2013, the cells don't seem to reflect formatting till after you click on them.
    //setColumnFormat: function(columnFormat, columnNumber) {
    //    this.mySpreadsheetoutput = this.mySpreadsheetoutput.concat('F;' + columnFormat + ';C' + columnNumber + '\n');
    //},
    
    // Lets try it by setting each cell that we want formatted, this will increase the file size but the other function doesn't work as expected for me.
    setColumnFormat: function(cellFormattingCode, cellNumberOfDigits, cellAlignmentCode, columnNumber) {
        for (var y = 2; y < this.rowY; y++) {
            this.mySpreadsheetoutput = this.mySpreadsheetoutput.concat('F;Y' + y + ';X' + columnNumber + ';F' + cellFormattingCode + cellNumberOfDigits + cellAlignmentCode + '\n');
        }
    },
    setCellFormat: function(cellFormattingCode, cellNumberOfDigits, cellAlignmentCode, x, y) {
        this.mySpreadsheetoutput = this.mySpreadsheetoutput.concat('F;Y' + y + ';X' + x + ';F' + cellFormattingCode + cellNumberOfDigits + cellAlignmentCode + '\n');
    },
    closeSpreadsheet: function() { this.mySpreadsheetoutput = this.mySpreadsheetoutput.concat('E'); }
}
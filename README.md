# Biff12Writer
A VB6 library for consuming/producing BIFF12 (.xlsb) spreadsheets

2018-08-13  Make some changes to cBiff12Writer.cls,
            Added function AddSheet to cBiff12Writer.cls (works only after the previous sheet is complete,
            no return to edit more i think. Optional parameter sheetname in function Init and AddSheet.
            Some changes to cBiff12Part.cls and cBiff12Container.cls to work with multiple sheets.

Attribute VB_Name = "GatherSheetsSummary"
Sub Main()

    Dim sheet_index As Integer
    Dim index_offset As Integer ' sheet number to start with
    
    ' user param ' can become a user form
    index_offset = 5
    
    Dim offset As Integer
    
    offset_row = 0
    
    Dim rng As Range
    
    For sheet_index = index_offset To Application.Sheets.Count
        
        Sheets(sheet_index).Select
        
        Set rng = Range("A116:G128")
              
        ' Grab Some Data and Store it in a "Range" variable
          
        Sheets("synthese_auto").Select
        
        offset_row = 14 * (sheet_index - index_offset)
        ' 14 is the number of rows contained in the range + 1 (to leave one row blank between tables)
        
        ' from : https://www.thespreadsheetguru.com/the-code-vault/best-way-to-copy-pastespecial-values-only-with-vba
        
        ' Transfer Values to same spot in another worksheet (Mimics PasteSpecial Values Only)
        Range("A1").offset(offset_row, 0).Resize(rng.Rows.Count, rng.Columns.Count).Cells.Value = rng.Cells.Value

    Next
    
End Sub




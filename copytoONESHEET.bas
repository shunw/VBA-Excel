Attribute VB_Name = "Module1"
Sub copytoONESHEET()

Sheets("Combine sheet").Cells.Clear
Cells(1, 1) = "Sheet Name"
Cells(1, 2) = "M"
Dim shtName As String

s_col = 9 'I col
col = 2 'combine sheet col
col_1 = 1

Set current_wb = ActiveWorkbook
total_sht = current_wb.Sheets.Count

begin_r = 2
For i = 1 To total_sht
        '==============================
        'get the sheet name
        '==============================
        shtName = current_wb.Sheets(i).Name
        
        '==============================
        'find other sheet except the "combine sheet"
        '==============================
        If Not shtName = "Combine sheet" Then
                
                s_begin_r = 5
                s_end_r = copy_row(shtName)
                'Debug.Print s_begin_r, s_end_r
                end_r = begin_r + (s_end_r - s_begin_r)
                
                
                With Sheets(shtName)
                        Range(.Cells(s_begin_r, s_col), .Cells(s_end_r, s_col)).Copy
                End With
    
                With Sheets("Combine sheet")
                        Range(.Cells(begin_r, col_1), .Cells(end_r, col_1)) = shtName
                        Cells(begin_r, col).PasteSpecial xlPasteValues
                End With

                begin_r = end_r + 1
        End If
                
        
        
Next i



End Sub


Function copy_row(sht_name As String)
'find the end row of the I column
row = 5
Do While Not IsEmpty(Sheets(sht_name).Cells(row, 9))
        row = row + 1
Loop
'return the end row in that sheet
copy_row = row - 1
End Function


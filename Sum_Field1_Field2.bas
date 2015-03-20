Attribute VB_Name = "Module1"
Sub cal_Field1_Field2()
'step 1: sort by D Name and Start Job
'--- F
'step 2: find the begin row and end row for a same D
'--- F
'step 3: find how many D exist
'step 4: sum the Field1 page if Input="Field1"; sum the Field2 page in input = "Field2"

j = "J_Sheet"
'######################################
'check if the current sheet name is the correct one
'######################################
right_sheet = ActiveSheet.Name
If right_sheet <> j Then
    'Debug.Print "I'm in"
    Exit Sub
End If

'######################################
'sort
'######################################
Selection.Sort Key1:=Range("C1"), Order1:=xlAscending, Key2:=Range("s1") _
, Order2:=xlAscending, Header:= _
xlGuess, OrderCustom:=1, MatchCase:=False, Orientation:=xlTopToBottom

'######################################
'find col
'######################################
d_col = Col_Find("D")
input_col = Col_Find("I")
mode_col = Col_Find("M")
p_qty_col = Col_Find("#ofP")

'######################################
'set up two col for two new items
'######################################

Field1_col = new_col(1)
Cells(1, Field1_col).Value = "Field1 count"
Field2_col = new_col(1)
Cells(1, Field2_col).Value = "Field2 count"

'######################################
'loop for different items
'######################################
outside_row = 2
Do While Cells(outside_row, d_col) <> ""

    '######################################
    'get begin and end row
    '######################################
    begin_row = find_row(Cells(outside_row, d_col), d_col)(0)
    end_row = find_row(Cells(outside_row, d_col), d_col)(1)
    Field1_page = 0
    Field2_page = 0
    
    '######################################
    'judge the input is Field1 or Field2
    '######################################
    For inside_row = begin_row To end_row
        p_qty = Cells(inside_row, p_qty_col).Value
        '######################################
        'judge the input is Field1 or Field2
        '######################################
        If Cells(inside_row, input_col).Value = "Field1" Then
            
            '######################################
            'judge the Field1 2 mode
            '######################################
            If Right(Cells(inside_row, mode_col).Value, 5) = "2" Then
                Field1_page = Field1_page + p_qty * 2
            Else
                Field1_page = Field1_page + p_qty
            End If
        Else
            Field2_page = Field2_page + p_qty
        End If
        Cells(inside_row, Field1_col).Value = Field1_page
        Cells(inside_row, Field2_col).Value = Field2_page
    Next inside_row
    
outside_row = end_row + 1
Loop
End Sub

Function find_row(D, col)
'find the begin row and the end row for one device
b_row = 2
Do While Cells(b_row, col) <> D Or Cells(b_row, col).Value = ""
    b_row = b_row + 1
Loop

e_row = b_row
Do While Cells(e_row + 1, col) = D And Cells(e_row + 1, col).Value <> ""
    e_row = e_row + 1
Loop

Dim arr(2)
arr(0) = b_row
arr(1) = e_row

find_row = arr
End Function


Function new_col(index)
'this is to add a new col in the right side
'index is to define the row you want to add the col
col = 1
Do While Cells(index, col).Value <> ""
    col = col + 1
Loop
new_col = col
End Function

Function Col_Find(TN As String)
'find the field name
row = 1
i = 1
Do While Cells(row, i).Value <> TN And Cells(row, i).Value <> ""
    i = i + 1
Loop
'Debug.Print i
Col_Find = i
End Function

Sub test()
Field1_col = new_col(1)
Cells(1, Field1_col).Value = "Field1 count"
Field2_col = new_col(1)
Cells(1, Field2_col).Value = "Field2 count"
End Sub



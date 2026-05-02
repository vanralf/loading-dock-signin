Private Sub Worksheet_Change(ByVal Target As Range) 

  ' SAFETY NET 1: If anything goes wrong, jump to the bottom and turn events back on 
  On Error GoTo CleanUp 
  
  ' Only react to a single cell change in columns B, F, or H 
  If Target.Cells.CountLarge > 1 Then Exit Sub 
  If Intersect(Target, Me.Range("B:B,F:F,H:H")) Is Nothing Then Exit Sub 

  Application.EnableEvents = False 

  ' Set row height whenever something is entered 
  If Target.Value <> "" Then Target.EntireRow.RowHeight = 30 
  Select Case Target.Column 

  ' ----------------------------- 
  ' NAME ENTRY (Column B) 
  ' ----------------------------- 
    Case 2 
      If Target.Value <> "" Then 
        ' Highlight row 
        Me.Rows(Target.Row).Interior.Color = RGB(255, 255, 0) 
        ' Print Date in Column K (Offset 9 columns from B) 
        Target.Offset(0, 9).Value = Date 
        Target.Offset(0, 9).NumberFormat = "mm/dd/yyyy" 
      End If 

  ' ----------------------------- 
  ' NEW VISITOR (Column F) 
  ' ----------------------------- 
  Case 6 
    If Target.Value = True Then 
      Dim wsContact As Worksheet 
      Dim tbl As ListObject 
      Dim newRow As ListRow 
      Dim phoneValue As String 
      Dim dupCount As Long 
      Dim phoneCell As Range 
      
      Set wsContact = ThisWorkbook.Sheets("ContactList") 
      Set tbl = wsContact.ListObjects("tblContactList") 
      
      ' Capture phone value from entry sheet (Column D) 
      phoneValue = Trim(Me.Cells(Target.Row, "D").Value) 
      
      ' Add new row to contact list table 
      Set newRow = tbl.ListRows.Add 
      newRow.Range(1, tbl.ListColumns("Name").Index).Value = _ Me.Cells(Target.Row, "B").Value 
      newRow.Range(1, tbl.ListColumns("Company").Index).Value = _ Me.Cells(Target.Row, "C").Value 
      newRow.Range(1, tbl.ListColumns("Phone Number").Index).Value = _ phoneValue 
      
  ' ---------------------------------------- 
  ' DUPLICATE PHONE NUMBER CHECK 
  ' ---------------------------------------- 
    If phoneValue <> "" Then 
      dupCount = Application.WorksheetFunction.CountIf( _ tbl.ListColumns("Phone Number").DataBodyRange, _ phoneValue _ ) 
      
    Set phoneCell = Me.Cells(Target.Row, "D") 
    
    If dupCount > 1 Then 
      ' Highlight duplicate phone number in entry sheet 
      phoneCell.Interior.Color = RGB(255, 0, 0) 
      phoneCell.Font.Color = RGB(255, 255, 255) 
      End If 
    End If 
  End If 
  ' ----------------------------- 
  ' CHECK-OUT (Column H) 
  ' ----------------------------- 
  Case 8 
    If Target.Value = True Then 
      ' Check-out time 
      Target.Offset(0, 1).Value = Now 
      Target.Offset(0, 1).NumberFormat = "h:mm AM/PM" 
      
      ' Remove highlight 
      Target.EntireRow.RowHeight = 15 
      Me.Rows(Target.Row).Interior.ColorIndex = xlNone 
      End If 
    End Select 
    
CleanUp: 
  ' Excel ALWAYS wakes back up, even if there is an error 
  Application.EnableEvents = True
  
End Sub

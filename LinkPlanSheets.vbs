Sub Link_plan_sheets()

Dim objExcel As Object
Dim objSheet As Object
Set objExcel = GetObject(, "Excel.Application")
Set objSheet = ThisWorkbook.Worksheets("plan")

Dim ArrSheets As Object
Set ArrSheets = CreateObject("Scripting.Dictionary")

For Each Sheet In ThisWorkbook.Worksheets
    ArrSheets.Add Sheet.Name, ""
Next

planSheet = "plan"
For i = 1 To 1000    
    nameSheet = objSheet.Cells(i, 1)
    If ArrSheets.Exists(nameSheet) Then        
      destinationAddress = CStr("'" + nameSheet + "'!A1")
      Worksheets(planSheet).Hyperlinks.Add Anchor:=Worksheets(planSheet).Cells(i, 1), Address:="", SubAddress:= _        
      destinationAddress, TextToDisplay:=nameSheet
      
      returnAddress = CStr("'" + planSheet + "'" + "!A" + CStr(i))
      Worksheets(nameSheet).Hyperlinks.Add Anchor:=Worksheets(nameSheet).Cells(1, 1), Address:="", SubAddress:= _        
      returnAddress, TextToDisplay:=planSheet
    End If
Next

End Sub

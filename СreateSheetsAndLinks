Sub CreateSheetsLinkPlanSheets()

Dim objExcel As Object
Dim objSheet As Object
Dim newSheet As Object
Set objExcel = GetObject(, "Excel.Application")
Set objSheet = ThisWorkbook.Worksheets("plan")
Dim DSheets As Object
Set DSheets = CreateObject("Scripting.Dictionary")
Dim DRows As Object
Set DRows = CreateObject("Scripting.Dictionary")

For Each Sheet In ThisWorkbook.Worksheets    
  DSheets.Add Sheet.Name, ""
Next

objSheet.Activate

Set fcell = Columns("A:H").Find("Выходные данные", LookAt:=xlPart)
If Not fcell Is Nothing Then
    columnOutput = fcell.Column    
    rowHead = fcell.Row
End If
'only rows with a filled 'output results' column are considered as individual steps

rowStart = rowHead + 1
Do While objSheet.Cells(rowStart, 1) <> ""
        If objSheet.Cells(rowStart, columnOutput) <> "" Then
            nameSheet = objSheet.Cells(rowStart, 1)        
            If Not DSheets.Exists(nameSheet) Then            
                Sheets.Add(After:=Sheets(Sheets.Count)).Name = nameSheet            
                Set newSheet = Worksheets(nameSheet)

                With newSheet
                  .Columns(1).ColumnWidth = objSheet.Columns(1).ColumnWidth                
                  .Columns(2).ColumnWidth = objSheet.Columns(2).ColumnWidth
                  .Columns(3).ColumnWidth = objSheet.Columns(3).ColumnWidth                
                  .Columns(4).ColumnWidth = objSheet.Columns(4).ColumnWidth
                  .Columns(5).ColumnWidth = objSheet.Columns(5).ColumnWidth                
                  .Columns(6).ColumnWidth = objSheet.Columns(6).ColumnWidth
                  .Columns(7).ColumnWidth = objSheet.Columns(7).ColumnWidth                
                  .Columns(8).ColumnWidth = objSheet.Columns(8).ColumnWidth
                  
                  For counter = 0 To DRows.Count - 1
                      rowFromArr = DRows.Keys                    
                      rowTo = counter + 2
                      
                      objSheet.Rows(rowFromArr(counter)).Copy _
                      Destination:=.Rows(rowTo)
                      'copy rows from plan to the sheet of step (all headers. starting from the second line of the page) 
                  Next
                  objSheet.Rows(rowStart).Copy _
                  Destination:=.Rows(rowTo + 1)        
                  'copy the row of the step
              End With        
          End If    
      Else
        If objSheet.Cells(rowStart - 1, columnOutput) <> "" Then            
            DRows.RemoveAll
            DRows.Add rowHead, ""        
            'if current row in result column is null then check current - 1 row, if it is not null, then there is the start of new block
            'all old headers are deleted 
            'collect new headers for new block 
            'each block has the same main header 
        End If
        DRows.Add rowStart, ""    
        'add subtitle to the dict  
      End If
      rowStart = rowStart + 1
Loop


planSheet = "plan"
For i = 1 To Cells(Rows.Count, 1).End(xlUp).Row    
    nameSheet = objSheet.Cells(i, 1)
    If DSheets.Exists(nameSheet) Then        
        destinationAddress = CStr("'" + nameSheet + "'!A1")
        Worksheets(planSheet).Hyperlinks.Add Anchor:=Worksheets(planSheet).Cells(i, 1), Address:="", SubAddress:= _
        destinationAddress, TextToDisplay:=nameSheet
    
        returnAddress = CStr("'" + planSheet + "'" + "!A" + CStr(i))
        Worksheets(nameSheet).Hyperlinks.Add Anchor:=Worksheets(nameSheet).Cells(1, 1), Address:="", SubAddress:= _
        returnAddress, TextToDisplay:=planSheet
    End If
Next

Beep
Shell "msg /TIME:" & 8 & " " & Environ("Username") & " " & "Листы созданы"

End Sub

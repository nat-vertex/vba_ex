Sub Address()

Dim oE As Object
Dim oSh As Object
Dim arrPatterns As Object
Dim patReg As Object
Dim newPatterns As Object

Set oE = GetObject(, "Excel.Application")
Set oSh = ThisWorkbook.Worksheets(Лист1.Name)
Set patReg = New RegExp
Set arrPatterns = CreateObject("Scripting.Dictionary")

patReg.IgnoreCase = False
patReg.Global = True

first_begin = "(^)[йцукенгшщзхъфывапролджэячсмитьбюё.-]{1,100}\s[\S]{1,100}"
second_begin = "(^)[йцукенгшщзхъфывапролджэячсмитьбюё.-]{1,100}\s"
first_end = "[\S]{1,100}\s[йцукенгшщзхъфывапролджэячсмитьбюё.-]{1,100}$"
second_end = "\s[йцукенгшщзхъфывапролджэячсмитьбюё.-]{1,100}$"

counter = 2             'первая строка с данными 
column_d = 1            'столбец с данными 
Do While oSh.Cells(counter, column_d) <> ""        
    ourAddress = oSh.Cells(counter, column_d)
    allSubs = Split(ourAddress, ", ")
    For Each subs In allSubs
        patReg.Pattern = first_begin
        Set newPatterns = patReg.Execute(subs)
        If newPatterns.Count > 0 Then
        
            For Each newPattern In newPatterns
            
                patReg.Pattern = second_begin
                Set resPatterns = patReg.Execute(newPattern)
                
                For Each resPattern In resPatterns
                    If Not arrPatterns.Exists(resPattern.Value) Then
                        arrPatterns.Add resPattern.Value, ""
                    End If
                Next
            Next
        Else
            patReg.Pattern = first_end
            Set newPatterns = patReg.Execute(subs)
            For Each newPattern In newPatterns
                patReg.Pattern = second_end
                Set resPatterns = patReg.Execute(newPattern)
                For Each resPattern In resPatterns
                    If Not arrPatterns.Exists(resPattern.Value) Then
                        arrPatterns.Add resPattern.Value, ""
                    End If
                Next
            Next
        End If
    Next
    counter = counter + 1
Loop

freeColumn = 10
counter = 2
Do While oSh.Cells(counter, column_d) <> ""
    ourAddress = oSh.Cells(counter, column_d)
    oSh.Cells(counter, freeColumn) = CleanAddress(ourAddress, arrPatterns)
    counter = counter + 1
Loop
End Sub


Function CleanAddress(ourAddress, arrPatterns) As String
resultAddress = ""
If arrPatterns.Exists("лесхоза ") Then
    arrPatterns.Remove "лесхоза "
End If
If arrPatterns.Exists("муниципальный ") Then
    arrPatterns.Remove "муниципальный "
If Not arrPatterns.Exists(" округ") Then
        arrPatterns.Add " округ", ""
    End If
End If

fullOurAddress = Split(ourAddress, ", ")

For Each partAddress In arrPatterns
    positionSpace = InStr(partAddress, " ")
    If positionSpace > 1 Then
        addSymbol = " *"
    Else
        addSymbol = "* "
    End If
    
    partAddressWithStar = Replace(partAddress, " ", addSymbol)
    st = partAddressWithStar
    partAddressWithStars = "*" + partAddressWithStar + "*"
    If ourAddress Like partAddressWithStars Then
    
                                        
        For Each subOurAddress In fullOurAddress
            If (subOurAddress Like st) Or (Left(st, 2) = "* " And subOurAddress Like (st + " *")) Or (Right(st, 2) = " *" And subOurAddress Like ("* " + st)) Then
                If resultAddress <> "" Then
                    resultAddress = resultAddress + ", " + subOurAddress
                Else
                    resultAddress = resultAddress + subOurAddress
                End If
                Exit For
            End If
        Next
    Else
        If resultAddress <> "" Then
            resultAddress = resultAddress + ", " + st
        Else
            resultAddress = resultAddress + st
        End If
    End If
Next
CleanAddress = resultAddress

End Function

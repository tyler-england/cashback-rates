Attribute VB_Name = "Archive"
Option Explicit
Dim bInProg As Boolean

Sub ArchiveData()
    Dim wsSource As Worksheet, wsPub As Worksheet
    Dim oDictVals As Object
    Dim iRowPub As Integer, iRowDate As Integer, icol As Integer, i As Integer
    Dim sCol As String, sPub As String, sNamePart As String, sOK As String, sSucc As String, sErr As String, sMsg As String
    Dim vRange As Variant, vSheet As Variant
    
    bInProg = True
    Set wsSource = ActiveSheet
    If ActiveSheet.Name <> ThisWorkbook.Worksheets(1).Name Then Set wsSource = ThisWorkbook.Worksheets(1)
    icol = wsSource.Range("AA1").End(xlToLeft).Column
    sCol = Split(Cells(1, icol).Address, "$")(1) 'last column with competitor
    vRange = wsSource.Range("A1:" & sCol & "100").Value2
    iRowPub = 2
    sPub = vRange(iRowPub, 1)
    Set oDictVals = CreateObject("Scripting.Dictionary")
    Do While sPub <> ""
        Application.StatusBar = "Updating " & sPub & "..."
        Set wsPub = Nothing
        For Each vSheet In ThisWorkbook.Worksheets
            If LCase(vSheet.Name) = LCase(sPub) Then
                Set wsPub = vSheet
                Exit For
            End If
        Next
        If wsPub Is Nothing Then 'not exact name -> try matching first word
            sNamePart = LCase(Split(sPub, " ")(1))
            For Each vSheet In ThisWorkbook
                If LCase(vSheet.Name) Like "*" & sNamePart & "*" Then
                    Set wsPub = vSheet
                    If sErr <> "" Then sErr = sErr & vbCrLf
                    sErr = sErr & Chr(149) & " " & sPub & " - May not be correct worksheet (" & vSheet.Name & ")"
                    Exit For
                End If
            Next
        End If
        If wsPub Is Nothing And Len(sPub) > 5 Then 'try matching beginning
            sNamePart = LCase(Left(sPub, 5))
            For Each vSheet In ThisWorkbook
                If LCase(vSheet.Name) Like "*" & sNamePart & "*" Then
                    Set wsPub = vSheet
                    If sErr <> "" Then sErr = sErr & vbCrLf
                    sErr = sErr & Chr(149) & " " & sPub & " - May not be correct worksheet (" & vSheet.Name & ")"
                    Exit For
                End If
            Next
        End If
        Debug.Print sPub
        If wsPub Is Nothing Then 'no sheet for this publisher
            If sErr <> "" Then sErr = sErr & vbCrLf
            sErr = sErr & Chr(149) & " Unable to find tab for " & sPub & " in row " & Str(iRowPub)
        Else 'wspub was found
            iRowDate = wsPub.Range("A10000").End(xlUp).Row
            If wsPub.Range("A" & iRowDate).Value <> Date Then
                iRowDate = iRowDate + 1
                wsPub.Range("A" & iRowDate).Value = Date
            End If
            For i = 2 To icol
                oDictVals.Add vRange(1, i), vRange(iRowPub, i)
            Next
            sOK = Replace(sPub, " ", "")
            With wsPub 'archive the values
                For i = 2 To icol
                    If oDictVals.exists(.Cells(1, i).Value) Then 'brand found
                        .Cells(iRowDate, i).Value = oDictVals(.Cells(1, i).Value)
                        If sOK = Replace(sPub, " ", "") Then
                            sOK = sOK & ": "
                        Else
                            sOK = sOK & ", "
                        End If
                        sOK = sOK & .Cells(1, i).Value
                    Else 'no value for this brand
                        If sErr <> "" Then sErr = sErr & vbCrLf
                        sErr = sErr & Chr(149) & " Error archiving value for " & Replace(sPub, " ", "") & "/" & .Cells(1, i).Value
                    End If
                Next
                If sOK <> sPub Then 'add to success list
                    If sSucc <> "" Then sSucc = sSucc & vbCrLf
                    sSucc = sSucc & Chr(149) & " " & sOK
                End If
            End With
        End If
        
        oDictVals.RemoveAll
        iRowPub = iRowPub + 1
        sPub = vRange(iRowPub, 1)
    Loop
    
    If sErr <> "" Then sMsg = "The following errors occurred:" & vbCrLf & vbCrLf & sErr & vbCrLf
    If sSucc <> "" Then
        If sMsg <> "" Then sMsg = sMsg & vbCrLf
        sMsg = sMsg & "The following items were successfully archived:" & vbCrLf & vbCrLf & sSucc
    End If
    bInProg = False
    MsgBox sMsg
    wsSource.Range("A1").Value = wsSource.Range("A1").Value
End Sub

Sub CheckArchivedData()
    Dim wsSource As Worksheet
    Dim rngAll As Range
    Dim vVals As Variant, var As Variant
    Dim iRow As Integer, icol As Integer, iRec As Integer, i As Integer
    Dim wSheet As Worksheet
    Dim sStatus As String, sSheet As String
    
    If bInProg Then Exit Sub
    Set wsSource = ThisWorkbook.Worksheets(1)
    iRow = wsSource.Range("A10000").End(xlUp).Row
    icol = wsSource.Range("AA1").End(xlToLeft).Column
    Set rngAll = wsSource.Range("B2:" & Split(Cells(1, icol).Address, "$")(1) & iRow)
    vVals = rngAll.Value2
    On Error Resume Next
    For Each var In vVals 'check if data is all there
        If var = "" Then
            sStatus = "Missing data"
            Exit For
        End If
    Next
    If sStatus = "" Then 'check against archived
        For Each var In ThisWorkbook.Worksheets 'check dates
            If WorksheetFunction.Max(var.Range("A:A")) < Date And var.Name <> wsSource.Name Then
                sStatus = "Archive not up to date"
                Exit For
            End If
        Next
        If sStatus = "" Then 'check actual data
            sStatus = "Archive not up to date" 'default value
            On Error GoTo errhandler
            iRow = 2
            Do While wsSource.Range("A" & iRow).Value > 0
                sSheet = wsSource.Range("A" & iRow).Value
                iRec = Worksheets(sSheet).Range("A10000").End(xlUp).Row
                For i = 2 To icol
                    If Worksheets(sSheet).Cells(iRec, i).Value = wsSource.Cells(iRow, i).Value Then 'equal (necessary for N/A values)
                        Err.Clear
                    ElseIf Round(Worksheets(sSheet).Cells(iRec, i).Value, 4) <> Round(wsSource.Cells(iRow, i).Value, 4) Then
                        Err.Raise (555)
                    End If
                Next
                iRow = iRow + 1
            Loop
            sStatus = "Up to date"
        End If
    End If
errhandler:
    If wsSource.Range("K5").Value <> sStatus Then
        bInProg = True
        wsSource.Range("K5").Value = sStatus
        bInProg = False
    End If
    Application.StatusBar = False
End Sub

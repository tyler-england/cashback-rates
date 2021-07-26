Attribute VB_Name = "Pull_Data"
Option Explicit

Sub GetData()
    Dim sBrands() As String, sPubs() As String, sURL As String
    Dim sPublisher As String, sPubWrds() As String, sDonePubs() As String
    Dim sSucc As String, sCol As String, sLine As String
    Dim i As Integer, j As Integer, k As Integer, x As Integer
    Dim icol As Integer, iRow As Integer
    Dim vVals As Variant, vHTML As Variant, vVal As Variant, vOut As Variant
    Dim bFound As Boolean, bDone As Boolean
    
    Application.ScreenUpdating = False
    With ThisWorkbook
        ReDim sBrands(.Worksheets(1).Range("B1").End(xlToRight).Column - 2)
        For i = 2 To .Worksheets(1).Range("B1").End(xlToRight).Column
            sBrands(i - 2) = .Worksheets(1).Cells(1, i).Value
        Next
        ReDim sPubs(Worksheets.Count - 2)
        For i = 2 To Worksheets.Count
            sPubs(i - 2) = Worksheets(i).Name
        Next
        .Worksheets(1).Range("B2:Z100").ClearContents
    End With
    For i = 0 To UBound(sBrands)
        Debug.Print sBrands(i)
        ReDim sDonePubs(0) 'reset "done publishers"
        ReDim vOut(UBound(sPubs)) 'output numbers
        x = 0
        icol = WorksheetFunction.Match("*" & Replace(sBrands(i), "kitchen", "") & "*", ThisWorkbook.Worksheets(1).Range("1:1"), 0)
        sCol = Split(Cells(1, icol).Address, "$")(1)
        sURL = "https://www.cashbackmonitor.com/cashback-store/" & sBrands(i)
        Set vHTML = CreateObject("htmlfile")
        vHTML.body.innerhtml = GetHTML(sURL)
        Set vVals = vHTML.body.getelementsbytagname("tr")
        On Error Resume Next
        For j = 2 To vVals.Length 'for each line on the site
            sLine = Replace(LCase(vVals(j).innertext), " ", "")
            If UBound(sPubs) <> UBound(sDonePubs) Then
                For k = 0 To UBound(sPubs) 'for each publisher we want
                    bFound = False
                    If x > 0 Then 'get rid of publishers already found
                        For Each vVal In sDonePubs
                            If LCase(vVal) = LCase(sPubs(k)) Then
                                bFound = True
                                Exit For
                            End If
                        Next
                    End If
                    If Not bFound Then
                        sPubWrds = Split(LCase(sPubs(k)), " ")
                        If UBound(sPubWrds) > 0 Then
                            If sLine Like "*" & sPubWrds(0) & "*" & sPubWrds(1) & "*" Then 'match
                                sPublisher = sPubs(k)
                            End If
                        End If
                        If sPublisher = "" Then
                            If sLine Like "*" & sPubWrds(0) & "*" Then 'match
                                sPublisher = sPubs(k)
                            End If
                        End If
                        If sPublisher <> "" Then 'match
                            ReDim Preserve sDonePubs(x)
                            sDonePubs(x) = sPubs(k)
                            x = x + 1
                            Exit For
                        End If
                    End If
                Next
                If sPublisher <> "" Then 'publisher found
                    If sBrands(i) = "Dyson" Then Debug.Print sPublisher
                    iRow = WorksheetFunction.Match("*" & sPublisher & "*", ThisWorkbook.Worksheets(1).Range("A:A"), 0) - 2
                    vOut(iRow) = GetValue(vVals(j).innertext) 'get the % or $/mi value
                End If
            End If
            sPublisher = "" 'reset for next iteration
        Next
        vVal = GetHoneyValue(sBrands(i)) 'honey
        If vVal > 0 Then vOut(UBound(vOut)) = vVal 'put into vOut
        For j = 0 To UBound(vOut) 'go through vout --> fill in N/A for any missing
            If vOut(j) = "" Then vOut(j) = "N/A"
        Next
        If InStr(sBrands(i), "kitchen") > 0 Then 'check ninja values that are already there
            For j = 0 To UBound(vOut)
                If IsNumeric(vOut(j)) Then
                    If IsNumeric(ThisWorkbook.Worksheets(1).Range(sCol & j + 2).Value) Then 'compare
                        If vOut(j) < ThisWorkbook.Worksheets(1).Range(sCol & j + 2).Value Then vOut(j) = ThisWorkbook.Worksheets(1).Range(sCol & j + 2).Value
                    End If
                Else
                    vOut(j) = ThisWorkbook.Worksheets(1).Range(sCol & j + 2).Value
                End If
            Next
        End If
        vOut = Application.WorksheetFunction.Transpose(vOut)
        ThisWorkbook.Worksheets(1).Range(sCol & "2:" & sCol & UBound(vOut) + 1).Value = vOut
        If Len(sSucc) > 0 Then sSucc = sSucc & vbCrLf
        sSucc = sSucc & Chr(149) & " " & sBrands(i)
        Set vHTML = Nothing
        If InStr(sBrands(i), "Ninja") > 0 And InStr(sBrands(i), "kitchen") = 0 Then 'check ninjakitchen
            sBrands(i) = "Ninjakitchen"
            i = i - 1
        End If
    Next
    sSucc = "The following brand numbers were found successfully:" & vbCrLf & vbCrLf & sSucc
    Application.ScreenUpdating = True
    MsgBox sSucc
End Sub

Function GetHTML(sURL As String) As String 'returns site HTML
    Dim oIE As Object
    Set oIE = CreateObject("InternetExplorer.Application")
    oIE.Visible = False
    oIE.Navigate sURL
    Do Until oIE.ReadyState = 4
    DoEvents
    Loop
    GetHTML = oIE.Document.DocumentElement.outerhtml
    oIE.Quit
    Set oIE = Nothing
End Function

Function GetValue(sPubAndVal As String) As Single 'returns numeric value of cashback
    Dim sWrds() As String
    Dim vVal As Variant
    Dim siMax As Single
    Dim i As Integer, j As Integer
    Dim bPercent As Boolean
    sWrds = Split(sPubAndVal, " ")
    For Each vVal In sWrds
        For i = 1 To Len(vVal)
            If IsNumeric(Mid(vVal, i, 1)) Then
                If Mid(vVal, i, 1) > siMax Then siMax = Mid(vVal, i, 1)
                j = i + 1
                Do While IsNumeric(Mid(vVal, i, j - i)) And j < 10
                    If IsNumeric(Mid(vVal, i, j - i)) And Mid(vVal, i, j - i) > siMax Then siMax = Mid(vVal, i, j - i)
                    j = j + 1
                Loop
                i = j
            End If
        Next
    Next
    If Mid(sPubAndVal, InStr(sPubAndVal, Trim(Str(siMax))), Len(Trim(Str(siMax))) + 1) = Trim(Str(siMax)) & "%" Then siMax = siMax / 100
    GetValue = siMax
    If InStr(sPubAndVal, "yson") > 0 Then Debug.Print sPubAndVal, siMax
End Function

Function GetHoneyValue(sBrand As String) As Single
    Dim sURL As String, sHTML As String, i As Integer, siVal As Single
    GetHoneyValue = 0
    sURL = "http://www.joinhoney.com/shop/" & sBrand & "/new/savings"
    sHTML = GetHTML(sURL)
    i = InStr(sHTML, "discountAmount")
    If i > 0 Then
        sHTML = Mid(sHTML, i, 30)
        siVal = GetValue(sHTML)
        If siVal > 0 Then GetHoneyValue = siVal
    End If
End Function

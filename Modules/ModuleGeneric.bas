Attribute VB_Name = "ModuleGeneric"
Option Explicit
Option Base 1

'Standard variables
Global strApplicationName As String
Global strApplicationEXEName As String

Global arrCompanyData(10) As String
Global arrData(13) As String
Global arrMenu() As Integer
Global blnErrors As Boolean

'Databases
Global wrkCurrent As DAO.Workspace
Global CommonDB As Database
Global PrintersDB As Database
Global UsersDB As Database
Global dBaseTables As TableDefs
Global TempQuery As QueryDef
Global TempSummaryPerPortQuery As QueryDef

'���������
Global strPrinterName As String
Global strPrinterType As String
Global intPrinterReportDetailLines As Integer
Global intPrinterReportTopMargin As Integer
Global intPrinterReportLeftMargin As Integer
Global strPrinterFontName As String
Global strPrinterFontSize As String

'Variables
Global strStandardMessages(30) As String
Global strAppMessages(20) As String
Global strCurrentUser As String
Global strFullPathName As String
Global strPathName As String
Global strReportsPathName As String
Global strCompanyName As String
Global strUnicodeFile As String
Global strAsciiFile As String
Global blnAppIsRunning As Boolean

'Indexes
Public Type typTableData
    strCode As String
    strFirstField As String
    strSecondField As String
    strThirdField As String
    strFourthField As String
    strFifthField As String
    strSixthField As String
    strSeventhField As String
    strEighthField As String
End Type

'Processes
Public glPid As Long
Public glHandle As Long
Public colHandle As New Collection
Public Const WM_CLOSE = &H10
Public Const WM_DESTROY = &H2

Declare Function SendMessage Lib "user32" Alias "SendMessageA" (ByVal hwnd As Long, ByVal wMsg As Long, ByVal wParam As Long, lParam As Any) As Long
Declare Function GetParent Lib "user32" (ByVal hwnd As Long) As Long
Declare Function GetWindowThreadProcessId Lib "user32" (ByVal hwnd As Long, lpdwProcessId As Long) As Long
Declare Function EnumWindows Lib "user32" (ByVal lpEnumFunc As Long, ByVal lParam As Long) As Long
Public Function CheckForSpecialCharacter(strCharacter)

    Select Case strCharacter
        Case "�": CheckForSpecialCharacter = "�"
        Case "�": CheckForSpecialCharacter = "�"
        Case "�": CheckForSpecialCharacter = "�"
        Case "�": CheckForSpecialCharacter = "�"
        Case "�": CheckForSpecialCharacter = "�"
        Case "�": CheckForSpecialCharacter = "�"
        Case "�": CheckForSpecialCharacter = "�"
        Case "�": CheckForSpecialCharacter = "�"
        Case "�": CheckForSpecialCharacter = "�"
        Case "�":: CheckForSpecialCharacter = "�"
        Case Else
            CheckForSpecialCharacter = strCharacter
    End Select

End Function

Public Function ConvertToSpecialUpperCase(someString)

    Dim intLoop As Integer
    Dim strConvertedString As String
    
    For intLoop = 1 To Len(someString)
        strConvertedString = strConvertedString & CheckForSpecialCharacter(Mid(someString, intLoop, 1))
    Next intLoop
    
    ConvertToSpecialUpperCase = UCase(strConvertedString)

End Function

Function HideObjects(ParamArray tmpObjects())

    Dim intLoop As Integer
    
    For intLoop = 0 To UBound(tmpObjects)
        tmpObjects(intLoop).Visible = False
    Next intLoop

End Function

Function InvertColorForNegativeNumbers(grdGrid As iGrid, lngCurrentRow As Long)

    Dim lngCol As Long
    
    For lngCol = 1 To grdGrid.colCount
        grdGrid.CellForeColor(lngCurrentRow, lngCol) = IIf(grdGrid.CellValue(lngCurrentRow, lngCol) < 0, &H8080FF, vbWhite)
    Next lngCol

End Function
Function LinesHaveBeenSelected(grdGrid As iGrid)

    Dim lngRow As Long
    
    LinesHaveBeenSelected = False
    
    For lngRow = 1 To grdGrid.RowCount
        If grdGrid.CellIcon(lngRow, "Selected") >= 1 Then
            LinesHaveBeenSelected = True
            Exit Function
        End If
    Next lngRow

End Function


Function CheckForMatch(DBToUse, TableToUse, FieldNames, FieldTypes, ParamArray FieldValues() As Variant)

    On Error GoTo ErrTrap
    
    Dim intLoop As Integer
    Dim strCriteria As String
    Dim arrFieldNames() As String
    Dim arrFieldTypes() As String
    Dim strSingleQuotes As String
    Dim strFieldValue As String
    
    Dim rstTempRecordset As Recordset
    
    If DBToUse = "CommonDB" Then Set TempQuery = CommonDB.CreateQueryDef("") Else Set TempQuery = PrintersDB.CreateQueryDef("")
    
    arrFieldNames() = Split(Replace(FieldNames, " ", ""), ",")
    arrFieldTypes() = Split(Replace(FieldTypes, " ", ""), ",")
    
    For intLoop = 0 To UBound(arrFieldNames)
        FieldValues(intLoop) = Replace(FieldValues(intLoop), "'", """") 'Replace single quotes, which causes things to break, with double
    Next intLoop
    
    For intLoop = 0 To UBound(arrFieldNames)
    
        If Len(FieldValues(intLoop)) >= 1 Then
        
            If arrFieldTypes(intLoop) = "String" Then strSingleQuotes = "'" Else strSingleQuotes = "" 'Add quotes if type is string, add nothing if type is numeric
             
            If Left(FieldValues(intLoop), 1) <> "*" Then 'If the leftmost character is not "star"
                If arrFieldTypes(intLoop) = "String" Then 'If the field type is a string
                    strCriteria = strCriteria & "Left(" & arrFieldNames(intLoop) & ", " & Len(FieldValues(intLoop)) & ")" & " = " & strSingleQuotes & FieldValues(intLoop) & strSingleQuotes 'Assemble the criteria with left characters as input with quotes
                End If
                If arrFieldTypes(intLoop) = "Numeric" Then 'If the field type is numeric
                    strCriteria = strCriteria & arrFieldNames(intLoop) & " = " & FieldValues(intLoop) 'Assemble the criteria with whole field as input with no quotes
                End If
            End If
            
            If Left(FieldValues(intLoop), 1) = "*" Then 'If the leftmost character is a "star"
                If arrFieldTypes(intLoop) = "String" Then 'If the field type is a string
                    strFieldValue = Right(FieldValues(intLoop), Len(FieldValues(intLoop)) - 1) 'Keep the field value without the leading star
                    strCriteria = strCriteria & "InStr(" & arrFieldNames(intLoop) & ", '" & strFieldValue & "')" 'Look if the given field value is contained inside the field name
                End If
            End If
            
            strCriteria = strCriteria & IIf(intLoop + 1 <= UBound(arrFieldNames), " AND ", "") 'If there are more fields, add logical condition
            
        End If
            
    Next intLoop
    
    TempQuery.SQL = "SELECT * FROM " & TableToUse & IIf(strCriteria <> "", " WHERE " & strCriteria, "")
    
    Set rstTempRecordset = TempQuery.OpenRecordset()
    
    Set CheckForMatch = rstTempRecordset
    
    Exit Function
    
ErrTrap:
    DisplayErrorMessage True, Err.Description

End Function

Function ClearNumberFormat(strInput)

    ClearNumberFormat = Replace(strInput, ".", "")

End Function

Function FullNumber(myNumber)
    
    On Error GoTo ErrTrap
    
    'Local ����������
    Dim intLoop As Byte
    Dim aArray(9, 10) As String
    Dim strTotalGross As String
    Dim strSubNumber As String
    Dim tmpDecNumber As String
    Dim strFullNumber As String
    Dim strDecNumber As String
    Dim bytArrayIndex As Byte
    Dim tmpIntNumber As Long
    Dim tmpNumber As String
    Dim aFullNumber(9) As String
    
    '������� �����
    bytArrayIndex = 1
   
    aArray(1, 1) = " "
    aArray(1, 2) = "������ "
    aArray(1, 3) = "�������� "
    aArray(1, 4) = "��������� "
    aArray(1, 5) = "���������� "
    aArray(1, 6) = "���������� "
    aArray(1, 7) = "�������� "
    aArray(1, 8) = "��������� "
    aArray(1, 9) = "��������� "
    aArray(1, 10) = "���������� "
    
    aArray(2, 1) = " "
    aArray(2, 2) = "���� "
    aArray(2, 3) = "������ "
    aArray(2, 4) = "������� "
    aArray(2, 5) = "������� "
    aArray(2, 6) = "������� "
    aArray(2, 7) = "������ "
    aArray(2, 8) = "��������� "
    aArray(2, 9) = "������� "
    aArray(2, 10) = "�������� "
    
    aArray(3, 1) = " "
    aArray(3, 2) = "��� "
    aArray(3, 3) = "��� "
    aArray(3, 4) = "���� "
    aArray(3, 5) = "������� "
    aArray(3, 6) = "����� "
    aArray(3, 7) = "��� "
    aArray(3, 8) = "���� "
    aArray(3, 9) = "���� "
    aArray(3, 10) = "����� "
    
    aArray(4, 1) = " "
    aArray(4, 2) = "������ "
    aArray(4, 3) = "��������� "
    aArray(4, 4) = "���������� "
    aArray(4, 5) = "����������� "
    aArray(4, 6) = "����������� "
    aArray(4, 7) = "��������� "
    aArray(4, 8) = "���������� "
    aArray(4, 9) = "���������� "
    aArray(4, 10) = "����������� "
    
    aArray(5, 1) = " "
    aArray(5, 2) = "���� "
    aArray(5, 3) = "������ "
    aArray(5, 4) = "������� "
    aArray(5, 5) = "������� "
    aArray(5, 6) = "������� "
    aArray(5, 7) = "������ "
    aArray(5, 8) = "��������� "
    aArray(5, 9) = "������� "
    aArray(5, 10) = "�������� "
    
    aArray(6, 1) = " "
    aArray(6, 2) = "��� "
    aArray(6, 3) = "��� "
    aArray(6, 4) = "����� "
    aArray(6, 5) = "�������� "
    aArray(6, 6) = "����� "
    aArray(6, 7) = "��� "
    aArray(6, 8) = "���� "
    aArray(6, 9) = "���� "
    aArray(6, 10) = "����� "
    
    aArray(7, 1) = " "
    aArray(7, 2) = "������ "
    aArray(7, 3) = "�������� "
    aArray(7, 4) = "��������� "
    aArray(7, 5) = "���������� "
    aArray(7, 6) = "���������� "
    aArray(7, 7) = "�������� "
    aArray(7, 8) = "��������� "
    aArray(7, 9) = "��������� "
    aArray(7, 10) = "���������� "
    
    aArray(8, 1) = " "
    aArray(8, 2) = "���� "
    aArray(8, 3) = "������ "
    aArray(8, 4) = "������� "
    aArray(8, 5) = "������� "
    aArray(8, 6) = "������� "
    aArray(8, 7) = "������ "
    aArray(8, 8) = "��������� "
    aArray(8, 9) = "������� "
    aArray(8, 10) = "�������� "
    
    aArray(9, 1) = " "
    aArray(9, 2) = "��� "
    aArray(9, 3) = "��� "
    aArray(9, 4) = "���� "
    aArray(9, 5) = "������� "
    aArray(9, 6) = "����� "
    aArray(9, 7) = "��� "
    aArray(9, 8) = "���� "
    aArray(9, 9) = "���� "
    aArray(9, 10) = "����� "
    
    For intLoop = 1 To 14
        If Mid(myNumber, intLoop, 1) <> "." Then
            tmpNumber = tmpNumber + Mid(myNumber, intLoop, 1)
        End If
    Next intLoop
    
    tmpIntNumber = Int(Val(tmpNumber))
    
    For intLoop = 1 To 9 - Len(Trim(tmpIntNumber))
        strTotalGross = strTotalGross + "0"
    Next intLoop
    strTotalGross = strTotalGross + Trim(tmpNumber)

    For intLoop = 1 To 9
        strSubNumber = Mid(strTotalGross, intLoop, 1)
        aFullNumber(intLoop) = aArray(bytArrayIndex, Val(strSubNumber) + 1)
        bytArrayIndex = bytArrayIndex + 1
    Next intLoop
    
    '�����������
    If aFullNumber(1) <> " " Or aFullNumber(2) <> " " Or aFullNumber(3) <> " " Then
        If aFullNumber(2) = "���� " Then
            If aFullNumber(3) = "��� " Then
                aFullNumber(2) = ""
                aFullNumber(3) = "������ "
            End If
            If aFullNumber(3) = "��� " Then
                aFullNumber(2) = ""
                aFullNumber(3) = "������ "
            End If
        End If
    End If
    
    '��������
    If aFullNumber(4) <> " " Or aFullNumber(5) <> " " Or aFullNumber(6) <> " " Then
        If aFullNumber(5) = "���� " Then
            If aFullNumber(6) = "��� " Then
                aFullNumber(5) = ""
                aFullNumber(6) = "������ "
            End If
            If aFullNumber(6) = "��� " Then
                aFullNumber(5) = ""
                aFullNumber(6) = "������ "
            End If
        End If
    End If
    
    '�����������
    If aFullNumber(7) <> " " Or aFullNumber(8) <> " " Or aFullNumber(9) <> " " Then
        If aFullNumber(8) = "���� " Then
            If aFullNumber(9) = "��� " Then
                aFullNumber(8) = ""
                aFullNumber(9) = "������ "
            End If
            If aFullNumber(9) = "��� " Then
                aFullNumber(8) = ""
                aFullNumber(9) = "������ "
            End If
        End If
    End If
    
    '�����������
    If aFullNumber(1) <> " " Or aFullNumber(2) <> " " Or aFullNumber(3) <> " " Then
        If aFullNumber(1) = "������ " And aFullNumber(2) = " " And aFullNumber(3) = " " Then
            aFullNumber(1) = "����� "
        End If
        If aFullNumber(1) = " " And aFullNumber(2) = " " And aFullNumber(3) = "��� " Then
            aFullNumber(3) = aFullNumber(3) + "����������� "
        Else
            aFullNumber(3) = aFullNumber(3) + "����������� "
        End If
    End If
    
    '��������
    If aFullNumber(4) <> " " Or aFullNumber(5) <> " " Or aFullNumber(6) <> " " Then
        If aFullNumber(4) = "������ " And aFullNumber(5) = " " And aFullNumber(6) = " " Then
            aFullNumber(4) = "����� "
        End If
        If aFullNumber(4) = " " And aFullNumber(5) = " " And aFullNumber(6) = "��� " Then
            aFullNumber(6) = "����� "
        End If
        If aFullNumber(6) <> "����� " Then
            aFullNumber(6) = aFullNumber(6) + "�������� "
        End If
    End If
    
    '�����������
    If aFullNumber(7) = "������ " And aFullNumber(8) = " " And aFullNumber(9) = " " Then
        aFullNumber(7) = "����� "
    End If
    
    For intLoop = 1 To 9
        If Trim(aFullNumber(intLoop)) <> "" Then
            strFullNumber = strFullNumber + aFullNumber(intLoop)
        End If
    Next intLoop
    
    If strFullNumber = "" Then strFullNumber = "����� "
    strFullNumber = strFullNumber + "���� "
    
    bytArrayIndex = 8
    tmpDecNumber = Mid(strTotalGross, 11, 2)
     
    If tmpDecNumber = "00" Or tmpDecNumber = "" Then
        FullNumber = strFullNumber
        Exit Function
    End If
        
    strFullNumber = IIf(strFullNumber <> "����� ���� ", strFullNumber + "��� ", "")
    
    For intLoop = 1 To 2
        strSubNumber = Mid(tmpDecNumber, intLoop, 1)
        aFullNumber(intLoop) = aArray(bytArrayIndex, Val(strSubNumber) + 1)
        bytArrayIndex = bytArrayIndex + 1
    Next intLoop
    
    If aFullNumber(1) <> " " Or aFullNumber(2) <> " " Then
        If aFullNumber(1) = "���� " Then
            If aFullNumber(2) = "��� " Then
                aFullNumber(1) = " "
                aFullNumber(2) = "������ "
            End If
            If aFullNumber(2) = "��� " Then
                aFullNumber(1) = " "
                aFullNumber(2) = "������ "
            End If
        End If
    End If
    
    For intLoop = 1 To 2
        If Len(Trim(aFullNumber(intLoop))) <> 0 Then
            strFullNumber = strFullNumber + aFullNumber(intLoop)
        End If
    Next intLoop
    
    If tmpDecNumber = "01" Then
        strFullNumber = strFullNumber + "����� "
    Else
        strFullNumber = strFullNumber + "����� "
    End If
            
    FullNumber = strFullNumber
    
    Exit Function
    
ErrTrap:
    FullNumber = "�� ���� ��� ������ �� ����������� ���������!"

End Function


Public Function CreateUnisexPDF(fileName As String, reportName As ActiveReport, fontSize As Integer)

    On Error GoTo ErrTrap
    
    Dim pdf As New ARExportPDF

    With reportName
        .oneLongField.Font.Size = fontSize
        .Restart
        .Run False
        pdf.AcrobatVersion = 2
        pdf.SemiDelimitedNeverEmbedFonts = ""
        pdf.fileName = Replace(fileName, "/", "-")
        pdf.fileName = Replace(pdf.fileName, "[", "")
        pdf.fileName = Replace(pdf.fileName, "]", "")
        pdf.fileName = Replace(pdf.fileName, "  ", " ")
        pdf.fileName = strReportsPathName & Replace(pdf.fileName, ":", "") & ".pdf"
        .Export pdf
    End With
    
    CreateUnisexPDF = True
    
    Exit Function
    
ErrTrap:
    CreateUnisexPDF = False
    DisplayErrorMessage True, Err.Description

End Function


Function ChangeEditButtonStatus(grdGrid, strTag, lngRow, lngCol)

    ChangeEditButtonStatus = False
    
    If grdGrid.RowCount = 0 Or lngRow = 0 Or strTag = "Blank" Then Exit Function
    
    If grdGrid.CellValue(lngRow, lngCol) <> "" Then ChangeEditButtonStatus = True

End Function

Function DisplayMessageRecordsNotFound()

    If MyMsgBox(1, strApplicationName, strStandardMessages(7), 1) Then
    End If

End Function

Function EnableGrid(grid As iGrid, canEdit As Boolean)

    With grid
        .Enabled = True
        .Redraw = True
        .Editable = canEdit
        .RowMode = Not canEdit
    End With

End Function

Function AddColumnsToGrid(grdGrid As iGrid, blnShowGridLines As Boolean, headerHeight, strLayoutCol, tmpElements, tmpTitles)

    On Error GoTo ErrTrap
    
    Dim intLoop As Integer
    Dim intNoOfElements As Integer
    Dim strKey As String
    Dim strHeader As String
    Dim intOuter As Integer
    Dim lngCol As Long
    
    intNoOfElements = 0
    
    With grdGrid
        .Clear True
        .Redraw = False
        .GridLines = IIf(blnShowGridLines, igGridLinesBoth, igGridLinesNone)
        .Visible = False
    End With
    
    ReDim arrWidth(1)
    ReDim arrJustification(1)
    ReDim arrFormat(1)
    ReDim arrKey(1)
    ReDim arrAllowSizing(1)
    ReDim arrHeaderTitle(1)
    
    For intOuter = 1 To Len(tmpElements)
        intNoOfElements = intNoOfElements + 1
        '������
        ReDim Preserve arrWidth(intNoOfElements)
        arrWidth(intNoOfElements) = Mid(tmpElements, intOuter, 2)
        intOuter = intOuter + 2
        '����������� � ������ �������
        ReDim Preserve arrAllowSizing(intNoOfElements)
        arrAllowSizing(intNoOfElements) = Mid(tmpElements, intOuter, 1)
        intOuter = intOuter + 1
        '��������
        ReDim Preserve arrJustification(intNoOfElements)
        arrJustification(intNoOfElements) = Mid(tmpElements, intOuter, 1)
        intOuter = intOuter + 1
        '�����
        ReDim Preserve arrFormat(intNoOfElements)
        arrFormat(intNoOfElements) = Mid(tmpElements, intOuter, 1)
        intOuter = intOuter + 1
        'ColKey
        ReDim Preserve arrKey(intNoOfElements)
        Do Until Mid(tmpElements, intOuter, 1) = ","
            If intOuter <= Len(tmpElements) Then
                strKey = strKey + Mid(tmpElements, intOuter, 1)
                intOuter = intOuter + 1
            Else
                Exit Do
            End If
        Loop
        arrKey(intNoOfElements) = strKey
        strKey = ""
    Next intOuter
    
    intNoOfElements = 0
    
    For intOuter = 1 To Len(tmpTitles)
        intNoOfElements = intNoOfElements + 1
        ReDim Preserve arrHeaderTitle(intNoOfElements)
        Do Until Mid(tmpTitles, intOuter, 1) = ","
            If intOuter <= Len(tmpTitles) Then
                strHeader = strHeader + Mid(tmpTitles, intOuter, 1)
                intOuter = intOuter + 1
            Else
                Exit Do
            End If
        Loop
        arrHeaderTitle(intNoOfElements) = strHeader
        strHeader = ""
    Next intOuter

    For intLoop = 1 To intNoOfElements
        strHeader = arrHeaderTitle(intLoop)
        With grdGrid.AddCol(sKey:=IIf(Left(arrKey(intLoop), 1) <> "X", arrKey(intLoop), Right(arrKey(intLoop), Len(arrKey(intLoop)) - 1)), sHeader:=strHeader, lWidth:=arrWidth(intLoop), eHdrTextFlags:=igTextCenter)
            Select Case arrJustification(intLoop)
                Case "L": .eTextFlags = 0
                Case "C": .eTextFlags = 1:
                Case "R": .eTextFlags = 2
            End Select
            Select Case arrFormat(intLoop)
                Case "I"
                    .sFmtString = "#,##0"
                Case "F"
                    .sFmtString = "#,##0.00"
                Case "D"
                    .sFmtString = "dd/mm/yyyy"
                Case "T"
                    .sFmtString = "hh:mm"
            End Select
        End With
        grdGrid.ColHeaderTextFlags(intLoop) = 32821
        grdGrid.ColTag(intLoop) = arrAllowSizing(intLoop)
        If Left(arrKey(intLoop), 1) = "X" Then
            grdGrid.ColHeaderTextFlags(intLoop) = 32789
        End If
    Next intLoop
    
    With grdGrid
        .LayoutCol = strLayoutCol
        .Header.Height = headerHeight
        .Redraw = True
        .Visible = True
    End With
    
    Exit Function

ErrTrap:
    AddColumnsToGrid = False
    DisplayErrorMessage True, Err.Description
    
    Exit Function

End Function


Public Function FormatDateAsFileName(myDate)

    If IsDate(myDate) Then
        FormatDateAsFileName = Right(myDate, 4) & "-" & Mid(myDate, 4, 2) & "-" & Left(myDate, 2)
    Else
        FormatDateAsFileName = myDate
    End If

End Function

Function HighlightRow(grdGrid As iGrid, lngSelectedRow, lngColumn, strID, blnRowMode)

    Dim lngIndex As Long
    
    If strID <> "" Then
        With grdGrid
            For lngIndex = 1 To .RowCount
                If (.CellText(lngIndex, lngColumn) = strID) Then
                    .EnsureVisibleRow lngIndex
                    .SetCurCell lngIndex, lngColumn
                    .RowMode = blnRowMode
                    .SetFocus
                    Exit Function
                End If
            Next lngIndex
        End With
    End If
    
    If strID = "" Then
        If grdGrid.RowCount > 0 Then
            If lngSelectedRow - 1 = 0 Then
                grdGrid.SetCurCell 1, lngColumn
                grdGrid.EnsureVisibleRow 1
            Else
                grdGrid.SetCurCell lngSelectedRow - 1, lngColumn
                grdGrid.EnsureVisibleRow lngSelectedRow - 1
            End If
            grdGrid.RowMode = blnRowMode
            grdGrid.SetFocus
        End If
    End If

End Function

'Public Function ShowMonthlyCalendar(myFormName As Form, myMonthyCalendar As MonthView)

'    With myMonthyCalendar
'        .Visible = True
'        .Left = myFormName.Width / 2 - .Width / 2
'        .Top = myFormName.Height / 2 - .Height / 2
'        .ZOrder 0
'        .Value = Date
'        .SetFocus
'    End With

'End Function

Function ToggleInfoPanel(thisForm As Form)

    With thisForm.frmInfo
        If .Visible = True Then
            .Visible = False
        Else
            .Visible = True
            .Left = 10100
            .Top = 1100
            .ZOrder 0
        End If
    End With

End Function

Function UpdateColors(thisForm As Form, formFullScreen As Boolean, Optional grdGrid As iGrid, Optional customColours As Boolean)

    Dim ctl As Control

    '�������� ����������
    For Each ctl In thisForm.Controls
        If ctl.Name = "lblToday" Then thisForm.lblToday.Caption = format(Date, "dddd dd/mm/yyyy")
    Next ctl
    
    '�����������
    For Each ctl In thisForm.Controls
        If ctl.Name = "frmInfo" Then thisForm.frmInfo.Visible = False
    Next ctl
    
    '�������
    For Each ctl In thisForm.Controls
        If ctl.Name = "frmProgress" Then
            With thisForm.frmProgress
                .Visible = False
                .ZOrder 1
                .Top = ((thisForm.Height + thisForm.Top) / 2) - (.Height / 2)
                .Left = (thisForm.Width / 2) - (.Width / 2)
            End With
        End If
    Next ctl
    
    '������ �����
    If formFullScreen Then
        '�����
        With thisForm
            .BackColor = GetSetting(strApplicationName, "Colors", "Background Full Screen Forms")
            .Top = 350
            .Height = CommonMain.Height - (.Top * 1.2)
            .Width = CommonMain.Width
            .Left = -10
        End With
        'Container
        With thisForm.frmContainer
            .BackColor = GetSetting(strApplicationName, "Colors", "Background Full Screen Forms")
            .Height = thisForm.Height - 510
            .Top = (thisForm.Height / 2) - (.Height / 2)
            .Left = (thisForm.Width / 2) - (.Width / 2)
        End With
        '�������
        With thisForm.frmButtonFrame
            .BackColor = GetSetting(strApplicationName, "Colors", "Background Full Screen Forms")
            .Top = thisForm.frmContainer.Height - 750
            .Left = (thisForm.frmContainer.Width / 2) - (.Width / 2)
        End With
        '��������� �������
        With thisForm.shpBackground
            .BackColor = GetSetting(strApplicationName, "Colors", "Background Containers")
            .Top = 975
            .Left = 0
            .Width = thisForm.Width
            .Height = thisForm.frmButtonFrame.Top - 1400 - .Top
        End With
        '������
        grdGrid.Height = thisForm.shpBackground.Height - grdGrid.Top + (thisForm.Top * 2)
        '������� ��� ������� �� ������
        For Each ctl In thisForm.Controls
            If ctl.Name = "frmFrameForGridButtons" Then
                With thisForm.frmFrameForGridButtons
                    .Top = thisForm.shpBackground.Height + 300
                    .Left = (thisForm.frmContainer.Width / 2) - (.Width / 2)
                    .BackColor = GetSetting(strApplicationName, "Colors", "Background Containers")
                End With
                grdGrid.Height = thisForm.Height - 3150 - thisForm.frmFrameForGridButtons.Height
            End If
        Next ctl
    End If
    
    '��� ������ ����� - ������������
    If Not formFullScreen Then
        thisForm.Width = thisForm.shpRightEdge.Left + thisForm.shpRightEdge.Width
        thisForm.Height = thisForm.shpBottomEdge.Top + thisForm.shpBottomEdge.Height - 90
        thisForm.Left = CommonMain.Width / 2 - thisForm.Width / 2
        thisForm.Top = CommonMain.Height / 2 - thisForm.Height / 2
        '�������
        With thisForm.frmButtonFrame
            .Left = (thisForm.Width / 2) - (thisForm.frmButtonFrame.Width / 2)
        End With
        '��������� �������
        With thisForm.shpBackground
            .Top = 900
            .Left = 225
            .Width = thisForm.Width - 470
            .Height = thisForm.frmButtonFrame.Top - 270 - .Top
        End With
    End If
    
    '��� ������ ����� - �������
    If Not formFullScreen And Not customColours Then
        thisForm.BackColor = GetSetting(strApplicationName, "Colors", "Forms Centered Background")
        thisForm.shpBackground.BackColor = GetSetting(strApplicationName, "Colors", "Background Containers")
        thisForm.frmButtonFrame.BackColor = GetSetting(strApplicationName, "Colors", "Forms Centered Background")
    End If
        
    '��������
    For Each ctl In thisForm.Controls
        If ctl.Name = "frmCriteria" Then
            With thisForm.frmCriteria
                .BackColor = GetSetting(strApplicationName, "Colors", "Background Criteria")
                .Visible = True
                .ZOrder 0
                .Top = ((grdGrid.Height) / 2) - (.Height / 2) + grdGrid.Top
                .Left = (grdGrid.Width / 2) - (.Width / 2) + grdGrid.Left
            End With
        End If
    Next ctl
    
    '�������
    For Each ctl In thisForm.Controls
        '��������
        If TypeOf ctl Is Label Then
            Select Case ctl.Name
                '������� �� �����
                Case "lblLabel"
                    ctl.ForeColor = GetSetting(strApplicationName, "Colors", "Labels Normal Foreground")
                    ctl.BackStyle = 0
                '������� �� ������� ���������
                Case "lblCriteriaLabel"
                    ctl.ForeColor = GetSetting(strApplicationName, "Colors", "Labels Criteria Foreground")
                    ctl.BackStyle = 0
            End Select
        End If
        '�������� ������
        If TypeOf ctl Is Label And Not customColours Then
            Select Case ctl.Name
                Case "lblTitle"
                    '�������� ������
                    Dim objFont As StdFont
                    ctl.ForeColor = GetSetting(strApplicationName, "Colors", "Labels Title Foreground")
                    Set objFont = New StdFont
                    objFont.Name = GetSetting(strApplicationName, "Colors", "Labels Title Font")
                    objFont.Size = 30
                    objFont.Bold = True
                    objFont.Charset = 161
                    Set ctl.Font = objFont
                    Set objFont = Nothing
            End Select
        End If
        'Checkboxes
        If TypeOf ctl Is CheckBox And Not customColours Then
            'Checkbox �� �����
            If Left(ctl.Name, 11) <> "chkCriteria" Then
                ctl.ForeColor = GetSetting(strApplicationName, "Colors", "Checkbox Normal Foreground")
                ctl.BackColor = GetSetting(strApplicationName, "Colors", "Checkbox Normal Background")
            End If
            'Checkbox �� ������� ���������
            If Left(ctl.Name, 11) = "chkCriteria" Then
                ctl.ForeColor = GetSetting(strApplicationName, "Colors", "Labels Criteria Foreground")
                ctl.BackColor = GetSetting(strApplicationName, "Colors", "Labels Criteria Background")
            End If
        End If
        'Radios
        If TypeOf ctl Is OptionButton And Not customColours Then
            'Radios �� �����
            If Left(ctl.Name, 11) <> "optCriteria" Then
                ctl.ForeColor = GetSetting(strApplicationName, "Colors", "OptionButton Normal Foreground")
                ctl.BackColor = GetSetting(strApplicationName, "Colors", "OptionButton Normal Background")
            End If
            'Radios �� ������� ���������
            If Left(ctl.Name, 11) = "optCriteria" Then
                ctl.ForeColor = GetSetting(strApplicationName, "Colors", "Labels Criteria Foreground")
                ctl.BackColor = GetSetting(strApplicationName, "Colors", "Labels Criteria Background")
            End If
        End If
        'Frames
        If TypeOf ctl Is Frame And Not customColours Then
            If ctl.Tag = "SameColorAsBackground" Then
                ctl.ForeColor = GetSetting(strApplicationName, "Colors", "Frames Foreground")
                ctl.BackColor = GetSetting(strApplicationName, "Colors", "Frames Background")
            End If
        End If
        '������� ��������
        If TypeOf ctl Is dcButton Then
            ctl.ForeColor = vbBlack
        End If
    Next

End Function


Function CheckForLoadedForm(thisForm As String)

    Dim loadedForm As Form
    
    On Error Resume Next
    
    CheckForLoadedForm = False
    
    For Each loadedForm In Forms
        If loadedForm.Name = thisForm Then
            CheckForLoadedForm = True
            Exit For
        End If
    Next loadedForm
    
End Function




Function PrinterExists(strPrinterName)

    Dim blnPrinterExists As Boolean
    Dim strPrinter As Printer
    
    blnPrinterExists = False
    
    For Each strPrinter In Printers
        If strPrinter.DeviceName = strPrinterName Then
            Set Printer = strPrinter
            blnPrinterExists = True
            Exit For
        End If
    Next
    
    If Not blnPrinterExists Then
        MyMsgBox 4, strApplicationName, strStandardMessages(18), 1
        Exit Function
    Else
        PrinterExists = True
    End If

End Function


Function KillProcess(appName)

    Dim process As Object

    For Each process In GetObject("winmgmts:").ExecQuery("Select * from Win32_Process")
        If process.Caption = appName Then
            process.Terminate (0)
        End If
    Next

End Function

Function UpdateRegistryWithUserData(company, user, password)

    SaveSetting strApplicationName, "Settings", "LastCompany", company
    SaveSetting strApplicationName, "Settings", "LastUserNameIndex", user
    SaveSetting strApplicationName, "Settings", "LastPassword", password

End Function



Function SelectPrinter(whatPrinterPrints)

    With CommonSelectPrinter
        .Tag = "True"
        .txtShowInList.text = whatPrinterPrints & "ID"
        .Show 1
    End With
    
    SelectPrinter = IIf(strPrinterName <> "", True, False)
    
End Function


Sub PrintColumnHeadings(ParamArray columns() As Variant)

    'Local variables
    Dim bytLoop As Byte
    
    For bytLoop = 0 To UBound(columns) - 1 Step 2
        Print #1, Tab(columns(bytLoop)); columns(bytLoop + 1);
    Next bytLoop
    
    Print #1, ""

End Sub

Function PrintHeadings(tmpColumns, tmpPageNo, tmpReportTitle, tmpReportSubTitle1)

    Dim bytLeft As Byte
    Dim bytPageLen As Byte
    
    bytPageLen = 6 + Len(tmpPageNo)
    
    Print #1, arrCompanyData(7); Tab(tmpColumns - bytPageLen); "������ " & tmpPageNo
    Print #1, arrCompanyData(8)
    Print #1, arrCompanyData(9)
    Print #1, arrCompanyData(10)
    
    Print #1, ""
    
    bytLeft = (tmpColumns / 2) - (Len(tmpReportTitle) / 2)
    Print #1, Space(bytLeft) & ConvertToSpecialUpperCase(tmpReportTitle)
    bytLeft = (tmpColumns / 2) - (Len(tmpReportSubTitle1) / 2)
    Print #1, Space(bytLeft) & ConvertToSpecialUpperCase(tmpReportSubTitle1)
    
    Print #1, "^"
    
End Function


Function CaptureNumbers(strString, tmpRow, tmpCol, tmpKeyAscii, blnDecimals)

    If (tmpKeyAscii = 46 Or tmpKeyAscii = 44) And blnDecimals Then
        If InStr(strString, ".") Or InStr(strString, ",") Then
            tmpKeyAscii = 0
        Else
            tmpKeyAscii = 44
            Exit Function
        End If
    End If
    
    If (tmpKeyAscii < 48 Or tmpKeyAscii > 58) And tmpKeyAscii <> 8 And tmpKeyAscii <> 13 Then
        tmpKeyAscii = 0
    End If

End Function

Function SimpleSeek(Table, index, ParamArray Indexes() As Variant)

    On Error GoTo ErrTrap
    
    Dim intLoop As Integer
    Dim intInnerLoop As Integer
    Dim strField()
    Dim intUpper As Integer
    Dim intArrayindex As Integer
    Dim strNewField As String
    Dim rsTable As Recordset
    
    SimpleSeek = False
    
    Set rsTable = CommonDB.OpenRecordset(Table)

    With rsTable
        .index = index
        If UBound(Indexes) = 0 Then .Seek "=", Indexes(0)
        If UBound(Indexes) = 1 Then .Seek "=", Indexes(0), Indexes(1)
        If .NoMatch Then SimpleSeek = True '�� � ������� ��� ������, ����� �� ��� ��������
        .Close
    End With
    
    Exit Function
    
ErrTrap:
    SimpleSeek = False
    DisplayErrorMessage True, Err.Description

End Function


Function SetUpGrid(myIconList As vbalImageList, ParamArray myGrid() As Variant)
    
    Dim intLoop As Integer
    
    For intLoop = 0 To UBound(myGrid)
        With myGrid(intLoop)
            .Editable = False
            .DefaultRowHeight = 22
            .RowMode = True
            .GridLines = igGridLinesBoth
            .GridLinesExtend = igGridLinesExtendBoth
            .ScrollBarStyle = 2
            .Top = .Top - 6
            With .Font
                .Name = "Ubuntu Condensed"
                .Size = 11
                .Bold = False
            End With
            With .Header
                .Flat = True
                .Buttons = False
                .BackColor = GetSetting(appName:=strApplicationName, Section:="Colors", Key:="Grid Header BackColor")
                .ForeColor = GetSetting(appName:=strApplicationName, Section:="Colors", Key:="Grid Header ForeColor")
                .SortInfoStyle = igSortInfoNone
                With .Font
                    .Name = "Ubuntu Condensed"
                    .Size = 10
                End With
            End With
            .ImageList = myIconList
        End With
    Next intLoop

End Function


Sub ClearFields(ParamArray tmpFields())

    Dim bytLoop As Byte
    
    For bytLoop = 0 To UBound(tmpFields)
        
        If TypeOf tmpFields(bytLoop) Is TextBox Or TypeOf tmpFields(bytLoop) Is newText Or TypeOf tmpFields(bytLoop) Is newInteger Or TypeOf tmpFields(bytLoop) Is newDate Or TypeOf tmpFields(bytLoop) Is newFloat Then
            tmpFields(bytLoop).text = ""
        End If
        If TypeOf tmpFields(bytLoop) Is Label Then
            tmpFields(bytLoop).Caption = ""
        End If
        If TypeOf tmpFields(bytLoop) Is CheckBox Then
            tmpFields(bytLoop).Value = 0
        End If
        If TypeOf tmpFields(bytLoop) Is OptionButton Then
            tmpFields(bytLoop).Value = False
        End If
        If TypeOf tmpFields(bytLoop) Is iGrid Then
            tmpFields(bytLoop).Clear
            tmpFields(bytLoop).TabStop = False
        End If
        If TypeOf tmpFields(bytLoop) Is Frame Then
            tmpFields(bytLoop).Visible = False
        End If
    Next bytLoop

End Sub


Sub InitializeFields(ParamArray tmpFields())

    Dim bytLoop As Byte
    
    For bytLoop = 0 To UBound(tmpFields)
        If TypeOf tmpFields(bytLoop) Is newDate Then
            tmpFields(bytLoop).text = format(Date, "dd/mm/yyyy")
        End If
        If TypeOf tmpFields(bytLoop) Is newFloat Then
            tmpFields(bytLoop).text = "0,00"
        End If
        If TypeOf tmpFields(bytLoop) Is newInteger Then
            tmpFields(bytLoop).text = "0"
        End If
        If TypeOf tmpFields(bytLoop) Is Label Then
            tmpFields(bytLoop).Caption = "0"
        End If
    Next bytLoop

End Sub

Sub InitializeProgressBar(frmForm, lblTitle, tmpRecordset)
    
    On Error GoTo ErrTrap
    
    With frmForm
        If Not tmpRecordset.EOF Then
            frmForm.lblMaster.Caption = lblTitle
            frmForm.frmProgress.Top = (frmForm.Height / 2) - (frmForm.frmProgress.Height / 2)
            frmForm.frmProgress.Left = (frmForm.Width / 2) - (frmForm.frmProgress.Width / 2)
            frmForm.prgProgressBar.Value = 0
            frmForm.prgProgressBar.Min = 0
            If Not IsNumeric(tmpRecordset) Then
                tmpRecordset.MoveLast
                frmForm.prgProgressBar.Max = tmpRecordset.RecordCount
                tmpRecordset.MoveFirst
            Else
                frmForm.prgProgressBar.Max = tmpRecordset
            End If
            frmForm.frmProgress.Visible = True
            frmForm.frmProgress.ZOrder 0
            frmForm.Refresh
        End If
    End With
    
    Exit Sub
    
ErrTrap:
    If Err.Number = 424 Then
        Resume Next
    End If

End Sub

Sub DisableFields(ParamArray tmpFields())

    Dim bytLoop As Byte
    
    For bytLoop = 0 To UBound(tmpFields)
        tmpFields(bytLoop).Enabled = False
    Next bytLoop

End Sub

Sub EnableFields(ParamArray tmpFields())
    
    Dim bytLoop As Byte
    
    For bytLoop = 0 To UBound(tmpFields)
        tmpFields(bytLoop).Enabled = True
    Next bytLoop

End Sub

Function MainDeleteRecord(SelectedDB, Table, FormTitle, IndexField, CodeToSeek, AskConfirmation)

    On Error GoTo ErrTrap
    
    Dim rsTable As Recordset
    
    Select Case SelectedDB
        Case "CommonDB"
            Set rsTable = CommonDB.OpenRecordset(Table)
        Case "PrintersDB"
            Set PrintersDB = DBEngine.OpenDataBase(App.Path + "\Data\" + "Printers.mdb", False, False)
            Set rsTable = PrintersDB.OpenRecordset(Table)
        Case "UsersDB"
            Set UsersDB = DBEngine.OpenDataBase(strPathName + "Users.mdb", False, False)
            Set rsTable = UsersDB.OpenRecordset(Table)
    End Select

    With rsTable
        .index = IndexField
        .Seek "=", CodeToSeek
        If Not .NoMatch Then
            If AskConfirmation = False Then
                .Delete
                .Close
                MainDeleteRecord = True
                Exit Function
            End If
            If MyMsgBox(3, FormTitle, strStandardMessages(4), 2) Then
                .Delete
                .Close
                MainDeleteRecord = True
            Else
                .Close
                MainDeleteRecord = False
            End If
        Else
            If MyMsgBox(4, FormTitle, strStandardMessages(9), 1) Then
            End If
        End If
    End With
    
    Exit Function
    
ErrTrap:
    MainDeleteRecord = False
    DisplayErrorMessage True, Err.Description
    
End Function

Function MainSeekRecord(SelectedDB, Table, IndexField, CodeToSeek, DisplayNotFoundMessage, ParamArray Fields())

    On Error GoTo ErrTrap
    
    Dim bytLoop As Byte
    Dim rsTable As Recordset
    
    Select Case SelectedDB
        Case "CommonDB"
            Set rsTable = CommonDB.OpenRecordset(Table)
        Case "PrintersDB"
            Set PrintersDB = DBEngine.OpenDataBase(App.Path + "\Data\" + "Printers.mdb", False, False)
            Set rsTable = PrintersDB.OpenRecordset(Table)
        Case "UsersDB"
            Set UsersDB = DBEngine.OpenDataBase(strPathName + "Users.mdb", False, False)
            Set rsTable = UsersDB.OpenRecordset(Table)
    End Select
    
    MainSeekRecord = True
    
    With rsTable
        .index = IndexField
        .Seek "=", CodeToSeek
        If Not .NoMatch Then
            For bytLoop = 0 To UBound(Fields)
                If TypeOf Fields(bytLoop) Is TextBox Or TypeOf Fields(bytLoop) Is newText Then
                    Fields(bytLoop).text = IIf(Not IsNull(rsTable.Fields(bytLoop)), rsTable.Fields(bytLoop), "")
                End If
                If TypeOf Fields(bytLoop) Is newFloat Then
                    Fields(bytLoop).text = format(rsTable.Fields(bytLoop), "#,##0.00")
                End If
                If TypeOf Fields(bytLoop) Is newInteger Then
                    Fields(bytLoop).text = format(rsTable.Fields(bytLoop), "#,##0")
                End If
                If TypeOf Fields(bytLoop) Is Label Then
                    Fields(bytLoop).Caption = rsTable.Fields(bytLoop)
                End If
                If TypeOf Fields(bytLoop) Is CheckBox Then
                    Fields(bytLoop).Value = IIf(rsTable.Fields(bytLoop), 1, 0)
                End If
                If TypeOf Fields(bytLoop) Is OptionButton Then
                    Fields(bytLoop).Value = IIf(rsTable.Fields(bytLoop), 1, 0)
                End If
                If TypeOf Fields(bytLoop) Is newDate Then
                    Fields(bytLoop).text = format(rsTable.Fields(bytLoop), "dd/mm/yyyy")
                End If
            Next bytLoop
        Else
            If DisplayNotFoundMessage Then
                If MyMsgBox(4, strApplicationName, strStandardMessages(9), 1) Then
                End If
                MainSeekRecord = False
            End If
        End If
        .Close
    End With
    
    Exit Function
    
ErrTrap:
    MainSeekRecord = False
    DisplayErrorMessage True, Err.Description
    
    Exit Function

End Function
Function ColorizeGrid(ParamArray tmpFields())
    
    Dim bytLoop As Byte
    
    For bytLoop = 0 To UBound(tmpFields)
        tmpFields(bytLoop).ForeColor = vbBlack
    Next bytLoop

End Function

Function DisplayErrorMessage(displayMessage, errorDescription, Optional progress As Frame, Optional grid As iGrid, Optional CloseThisConnection As Boolean = True)

    If displayMessage Then
        If Not progress Is Nothing Then progress.Visible = False
        If Not grid Is Nothing Then grid.Redraw = True
        If MyMsgBox(4, strApplicationName, strStandardMessages(13), 1, errorDescription) Then
        End If
    End If
    
    UpdateLogFile errorDescription

End Function


Function UpdateLogFile(errorDescription)

    On Error GoTo ErrTrap
    
    Open strReportsPathName & "Errors.txt" For Append As #2
        Print #2, format(Date, "dd/mm/yyyy") & " " & format(Time, "hh:mm") & " " & errorDescription; ""
    Close #2
    
    Exit Function
    
ErrTrap:
    
    If Err.Number = 55 Then
        Close #2
        Open strPathName & "Errors.txt" For Append As #2
        Resume Next
    Else
        Exit Function
    End If
    
End Function


Function FillGridFromDB(SelectedDB, grdGrid, strTable, Fields, joins, criteriaString, sortColumn, ParamArray arguments())
    
    On Error GoTo ErrTrap
    
    Dim intLoop As Integer
    Dim lngRow As Long
    Dim lngCol As Long
    Dim strSQL As String

    Dim rstTempRecordset As Recordset
    
    strPrinterName = ""
    FillGridFromDB = False
    
    strSQL = "SELECT " & IIf(Fields = "", "*", Fields) & " FROM " & strTable & " " & joins & IIf(criteriaString <> "", "WHERE " & criteriaString, "")
    
    Select Case SelectedDB
        Case "CommonDB"
            Set rstTempRecordset = CommonDB.OpenRecordset(strSQL)
        Case "PrintersDB"
            Set PrintersDB = DBEngine.OpenDataBase(App.Path + "\Data\" + "Printers.mdb", False, False)
            Set rstTempRecordset = PrintersDB.OpenRecordset(strSQL)
        Case "UsersDB"
            Set UsersDB = DBEngine.OpenDataBase(strPathName + "Users.mdb", False, False)
            Set rstTempRecordset = UsersDB.OpenRecordset(strSQL)
    End Select
    
    With grdGrid
        .Clear
        .Redraw = False
    End With
    
    Do Until rstTempRecordset.EOF
        grdGrid.AddRow
        intLoop = 0
        lngRow = grdGrid.RowCount
        For lngCol = 1 To UBound(arguments) + 1
            grdGrid.CellValue(lngRow, lngCol) = rstTempRecordset.Fields(arguments(intLoop))
            intLoop = intLoop + 1
        Next lngCol
        rstTempRecordset.MoveNext
    Loop
    
    grdGrid.Redraw = True
    
    If grdGrid.RowCount > 0 Then
        FillGridFromDB = True
        With grdGrid
            .Sort sortColumn
            .Enabled = True
        End With
    End If
    
    Exit Function
    
ErrTrap:
    FillGridFromDB = False
    DisplayErrorMessage True, Err.Description
    
End Function


Function MainSaveRecord(SelectedDB, Table, Status, FormTitle, IndexField, CodeToSeek, ParamArray Fields() As Variant)

    On Error GoTo ErrTrap
    
    Dim lngFieldNo As Long
    Dim rsTable As Recordset
    
    Select Case SelectedDB
        Case "CommonDB"
            Set rsTable = CommonDB.OpenRecordset(Table)
        Case "PrintersDB"
            Set PrintersDB = DBEngine.OpenDataBase(App.Path + "\Data\" + "Printers.mdb", False, False)
            Set rsTable = PrintersDB.OpenRecordset(Table)
        Case "UsersDB"
            Set UsersDB = DBEngine.OpenDataBase(strPathName + "Users.mdb", False, False)
            Set rsTable = UsersDB.OpenRecordset(Table)
    End Select
    
    With rsTable
        .index = IndexField
        If Status Then
            .AddNew
        Else
            .Seek "=", CodeToSeek
            If Not .NoMatch Then
                .Edit
            Else
                If MyMsgBox(4, FormTitle, strStandardMessages(9), 1) Then
                End If
                MainSaveRecord = 0
                Exit Function
            End If
        End If
        For lngFieldNo = 0 To UBound(Fields)
            'Debug.Print .Fields(lngFieldNo + 1).Name & " " & Fields(lngFieldNo)
            .Fields(lngFieldNo + 1).Value = Trim(Fields(lngFieldNo))
        Next
        .Update
        If Status Then
            .MoveLast
        End If
        MainSaveRecord = .Fields(0).Value
        .Close
    End With
    
    Exit Function
    
ErrTrap:
    MainSaveRecord = 0
    DisplayErrorMessage True, Err.Description
    
End Function

Function MoveToNextColumn(grdGrid As iGrid, lngRow, lngCol)

    On Error GoTo ErrTrap
    
    Do While True
        If lngCol + 1 <= grdGrid.colCount Then
            If grdGrid.ColTag(lngCol + 1) = "Y" Then
                grdGrid.SetCurCell lngRow, lngCol + 1
                Exit Function
            End If
        Else
            lngCol = 1
            Do While True
                grdGrid.SetCurCell lngRow + 1, lngCol
                If grdGrid.ColTag(lngCol) = "Y" Then
                    Exit Function
                End If
                lngCol = lngCol + 1
            Loop
        End If
        lngCol = lngCol + 1
    Loop
    
ErrTrap:
    If Err.Number = -2147220991 Then Exit Function

End Function


Sub UpdateButtons(formName, Max, ParamArray Buttons() As Variant)
    
    Dim intLoop As Integer
    
    For intLoop = 0 To Max
        formName.cmdButton(intLoop).Enabled = Buttons(intLoop)
    Next intLoop
    
    formName.Refresh
    
End Sub

Sub CheckForArrows(KeyCode)
    
    'Up
    If KeyCode = 38 Then
        Sendkeys "+{TAB}"
        KeyCode = 0
    End If
    
    'Down
    If KeyCode = 40 Then
        Sendkeys "{TAB}"
        KeyCode = 0
    End If
    
End Sub

Sub UpdateProgressBar(frmForm)
    
    frmForm.prgProgressBar.Value = frmForm.prgProgressBar.Value + 1
       
End Sub

Function SelectRow(grdGrid As iGrid, colorIndex, strKeyCode, lngRow, lngCol)

    '������
    If grdGrid.RowCount = 0 Then Exit Function
    If lngRow = 0 Then Exit Function
    If grdGrid.CellValue(lngRow, lngCol) = "" Then SelectRow = 1: Exit Function
    
    '������� - ��������� �� ������ �� space
    With grdGrid
        If strKeyCode = 32 Then
            If .CellIcon(lngRow, "Selected") = "-1" Or .CellIcon(lngRow, "Selected") = "0" Then
                SelectRow = colorIndex
            Else
                SelectRow = 1
            End If
        End If
    End With

End Function

Sub ValidateInput(KeyAscii)
    
    If KeyAscii = 13 Then
        KeyAscii = 0
        Sendkeys "{tab}"
    End If

End Sub

Function AddTitle(sheet As Object, Title As String, colCount As Long)

    'Excel
    With sheet
        .Range("A6:" & Chr(colCount + 64) & "6").MergeCells = True
        .Range("A6").Value = Title
        .Range("A6").HorizontalAlignment = 3
        .Range("A6").VerticalAlignment = 2
        .rows("6").RowHeight = 24
    End With

End Function



Function AdjustColumnWidths(sheet As Object, ParamArray columns() As Variant)

    Dim X As Integer
    
    'Excel
    With sheet
        For X = 0 To UBound(columns) - 1 / 2 Step 2
            .columns(columns(X)).columnWidth = columns(X + 1)
        Next X
    End With

End Function
Function AddCriteria(sheet As Object, criteria As String, colCount As Long)

    'Excel
    With sheet
        .Range("A7:" & Chr(colCount + 64) & "7").MergeCells = True
        .Range("A7:" & Chr(colCount + 64) & "7").WrapText = True
        .Range("A7").Value = criteria
        .Range("A7").HorizontalAlignment = 3
        .Range("A7").VerticalAlignment = 2
        .rows("7").RowHeight = 40
    End With

End Function

Function AddHeaders(sheet As Object, grid As iGrid, colCount As Long, ParamArray columns() As Variant)

    On Error Resume Next
    
    Dim X As Integer
    Dim z As Integer
    Dim lngColCount As Long
    lngColCount = UBound(columns) + 1
    
    X = 0
    z = 1
    
    With sheet
        .Range("A9:" & Chr(colCount + 64) & "9").WrapText = True
        .Range("A9:" & Chr(colCount + 64) & "9").HorizontalAlignment = 3
        .Range("A9:" & Chr(colCount + 64) & "9").VerticalAlignment = 2
        Do While X < lngColCount
            .Range("" & columns(X) & "9").Value = grid.ColHeaderText(z)
            X = X + 2
            z = z + 1
        Loop
        .rows("9").RowHeight = 30
    End With

End Function


Sub LoadMessages()

    strStandardMessages(1) = Chr(13) & "�� ����� ����� �����������." & Chr(13)
    strStandardMessages(2) = Chr(13) & "�� ����� ��� ����� �����." & Chr(13)
    strStandardMessages(3) = "�� ������������� ��� �����������" & Chr(13) & "�� ������ ��� �� ����������." & Chr(13) & "������ ������� �� �������������;"
    strStandardMessages(4) = "� ������� �� ��������� ��������." & Chr(13) & "����� �������� ��� ������" & Chr(13) & "�� ���������� ��� �������;"
    strStandardMessages(5) = "� ������� ��� ������������."
    strStandardMessages(6) = Chr(13) & "��� ����� �������� ��������."
    strStandardMessages(7) = Chr(13) & "��� �������� ��������."
    strStandardMessages(8) = Chr(13) & "� ���������� ������������."
    strStandardMessages(9) = Chr(13) & "� ������� ��� �������."
    strStandardMessages(10) = Chr(13) & "� ����� ��� - ��� ��� ����� �����." & Chr(13)
    strStandardMessages(11) = "�� ����� ��� ������" & Chr(13) & "�/��� � �������" & Chr(13) & "��� ������ ����� �����."
    strStandardMessages(13) = "� ������� ������������ �������� ��� ���" & Chr(13) & " ������������. ������� �� ������ ����� ��� ���� ������������."
    strStandardMessages(14) = "�� ����� '���� �������' ������" & Chr(13) & "�� ����� ���� ��" & Chr(13) & "�� ����� '����������� ���� �������'."
    strStandardMessages(15) = Chr(13) & "� �������� ���������� ���." & Chr(13)
    strStandardMessages(16) = Chr(13) & "������ �� ����������� ��� ��������;" & Chr(13)
    strStandardMessages(17) = Chr(13) & "��� ������� ��������� ��������." & Chr(13)
    strStandardMessages(18) = "� ��������� ��� ��������� ���" & Chr(13) & "������� ��� �������." & Chr(13) & "������� �� ����� ��� ���������������."
    strStandardMessages(19) = Chr(13) & "��� ������� ��������� ������������." & Chr(13)
    strStandardMessages(20) = "� �������� ��������. ����� ���� �������!"
    strStandardMessages(21) = "��� �� �������� ����� �������" & Chr(13) & "��� ������, ������ ��" & Chr(13) & "����� ������������ ��� ���������."
    strStandardMessages(22) = Chr(13) & "� �������� ������������ ����� ����."
    strStandardMessages(23) = Chr(13) & "� ������� ������������ ��������."
    strStandardMessages(24) = "� �������� �����������"
    strStandardMessages(25) = "�������� ��� ����������� ������"
    strStandardMessages(26) = "����� ���������"
    strStandardMessages(27) = Chr(13) & "� ���������� ���������"
    strStandardMessages(28) = Chr(13) & "�� ����������� ����� ��� ������������"
    
    strAppMessages(1) = Chr(13) & "��� ������� ����������� �� �� ���� ���������."
    strAppMessages(2) = "�������� ������ �������� �� �����" & Chr(13) & "������ �� ���������� �����." & Chr(13) & "������� �� ������ ����� ��� ���� ������������."
    strAppMessages(3) = "� ������� ��� ���� " & Chr(13) & "������������." & Chr(13) & "������ �� ������������� ���� ���;"
    strAppMessages(4) = "��� �������� �� ������������" & Chr(13) & "�� ����������" & Chr(13) & "��������� ��� "
    strAppMessages(5) = "��� �������� �� ������������" & Chr(13) & "�� ����������" & Chr(13) & "���������� ��� ���������."
    strAppMessages(6) = "� ���������� ��� ������������" & Chr(13) & "������ �������� ����." & Chr(13)
    strAppMessages(7) = "� ������� ������������." & Chr(13) & "������ �� ���������" & Chr(13) & "�� �����������;"
    strAppMessages(8) = Chr(13) & "������ �� ��������� ���� ���� �����"
    strAppMessages(9) = Chr(13) & "������ �� ������������ ��� �� ��������"
    strAppMessages(10) = "�� ����������� �������� �� ����������" & Chr(13) & " ��������. ����� �������� ��� ������" & Chr(13) & "�� ��� ����������;"
    strAppMessages(11) = Chr(13) & "�� ����, ��������� ����� ���"
    strAppMessages(12) = "� ���������� �� ������������" & Chr(13) & "�������� �� �� ������� ��� ������." & Chr(13) & "������ �� ����������;"
    strAppMessages(13) = Chr(13) & "� ������� ��� ����� ��������."
    
End Sub

Function UpdateRecordCount(myLabel As Label, myRecordCount)

    myLabel.Caption = "�������� " & myRecordCount & " ��������"

End Function

Function CountSelected(myGrid As iGrid)

    Dim lngRow As Long
    Dim intSelected As Integer
    
    For lngRow = 1 To myGrid.RowCount
        If myGrid.CellIcon(lngRow, "Selected") > 0 Then
            intSelected = intSelected + 1
        End If
    Next lngRow
    
    CountSelected = IIf(intSelected > 0, "����������� " & intSelected & " ��������", "")

End Function

Function SumSelectedGridRows(myGrid As iGrid, myLastColumnIsSpecial, ParamArray myColumns() As Variant)

    Dim lngRow As Long
    Dim intLoop As Integer
    Dim blnSelected As Boolean
    Dim strDummy As String
    ReDim curGridColumnTotals(UBound(myColumns) + 1)
    Dim strFormat As String
    Dim strHeaderText As String
    
    For lngRow = 1 To myGrid.RowCount
        If myGrid.CellIcon(lngRow, "Selected") > 0 Then
            blnSelected = True
            For intLoop = 0 To UBound(myColumns) + IIf(myLastColumnIsSpecial, -3, 0) Step 3
                curGridColumnTotals(intLoop + 1) = curGridColumnTotals(intLoop + 1) + myGrid.CellValue(lngRow, myColumns(intLoop))
            Next intLoop
            If myLastColumnIsSpecial Then
                curGridColumnTotals(intLoop + 1) = curGridColumnTotals(intLoop + 1) + myGrid.CellValue(lngRow, myColumns(intLoop - 6)) - myGrid.CellValue(lngRow, myColumns(intLoop - 3))
            End If
        End If
    Next lngRow
    
    If blnSelected Then
        For intLoop = 2 To UBound(myColumns) Step 3
            strFormat = IIf(myColumns(intLoop) = "integer", "#,##0", "#,##0.00")
            strDummy = strDummy & myColumns(intLoop - 1) & " " & format(curGridColumnTotals(intLoop - 1), strFormat) & " "
        Next intLoop
        SumSelectedGridRows = Replace(Replace(Left(strDummy, Len(strDummy) - 1), Chr(13), " "), "   ", " ")
    End If

End Function

Function MyMsgBox(intPictureIndex, txtTitle, txtLine, intNoOfButtons, Optional errorDescription = "")

    With CommonMessages
        .frmButtonFrame(1).Visible = False
        .frmButtonFrame(2).Visible = False
        .imgImage.Picture = .lslIcons.ItemPicture(intPictureIndex)
        .imgImage.ToolTipText = errorDescription
        .lblTitle = txtTitle
        .lblLine = txtLine
        .frmButtonFrame(intNoOfButtons).Visible = True
        .Show 1
        If .cmdButton(0).Tag = "Pressed" Then
            MyMsgBox = True
            Exit Function
        Else
            MyMsgBox = False
            Exit Function
        End If
        If .cmdButton(2).Tag = "Pressed" Then
            MyMsgBox = True
        End If
    End With
    
End Function

Function OpenDataBase(tmpCompany)

    On Error GoTo TrapError
    
    OpenDataBase = False
    
    Set wrkCurrent = DBEngine.Workspaces(0)
    
    strFullPathName = strPathName & tmpCompany
    Set CommonDB = DBEngine.OpenDataBase(strFullPathName, False, False)
    OpenDataBase = True
    Set dBaseTables = CommonDB.TableDefs
    
    Exit Function
    
TrapError:
    If Err.Number = 3031 Or Err.Number = 3029 Then
        Exit Function
    Else
        Exit Function
    End If
    
End Function

Public Sub Sendkeys(text As Variant, Optional wait As Boolean = False)
   
    Dim WshShell As Object
   
    Set WshShell = CreateObject("wscript.shell")
   
    WshShell.Sendkeys CStr(text), wait
   
    Set WshShell = Nothing
   
End Sub

Private Function pvCryptXor(ByVal lI As Long, ByVal lJ As Long) As Long
    
    If lI = lJ Then
        pvCryptXor = lJ
    Else
        pvCryptXor = lI Xor lJ
    End If
    
End Function

Public Function CryptRC4(username, password) As String
    
    Dim baS(0 To 255) As Byte
    Dim baK(0 To 255) As Byte
    Dim bytSwap     As Byte
    Dim lI As Long
    Dim lJ As Long
    Dim lIdx As Long

    For lIdx = 0 To 255
        baS(lIdx) = lIdx
        baK(lIdx) = Asc(Mid$(password, 1 + (lIdx Mod Len(password)), 1))
    Next
    
    For lI = 0 To 255
        lJ = (lJ + baS(lI) + baK(lI)) Mod 256
        bytSwap = baS(lI)
        baS(lI) = baS(lJ)
        baS(lJ) = bytSwap
    Next
    
    lI = 0
    lJ = 0
    
    For lIdx = 1 To Len(username)
        lI = (lI + 1) Mod 256
        lJ = (lJ + baS(lI)) Mod 256
        bytSwap = baS(lI)
        baS(lI) = baS(lJ)
        baS(lJ) = bytSwap
        CryptRC4 = CryptRC4 & Chr$((pvCryptXor(baS((CLng(baS(lI)) + baS(lJ)) Mod 256), Asc(Mid$(username, lIdx, 1)))))
    Next
    
End Function

Public Function ToHexDump(sText As String) As String
    
    Dim lIdx As Long

    For lIdx = 1 To Len(sText)
        ToHexDump = ToHexDump & Right$("0" & Hex(Asc(Mid(sText, lIdx, 1))), 2)
    Next
    
End Function

Function IsCorrectPassword(strUsername, strPassword As String)

    Dim rstUsers As Recordset
    Dim strUserInput As String
    
    strPathName = GetSetting(appName:=strApplicationName, Section:="Path Names", Key:="Database Path Name")
    Set UsersDB = DBEngine.OpenDataBase(strPathName + "Users.mdb", False, False)
    
    Set TempQuery = UsersDB.CreateQueryDef("")
    
    TempQuery.SQL = "SELECT * FROM Users WHERE Username = '" & strUsername & "' AND PasswordHash = '" & HashPassword(strUsername, strPassword) & "'"
    
    Set rstUsers = TempQuery.OpenRecordset()
    
    If Not rstUsers.EOF Then
        IsCorrectPassword = True
    Else
        IsCorrectPassword = False
    End If
    
    UsersDB.Close
    
End Function

Public Function HashPassword(username, password)
    
    HashPassword = ToHexDump(CryptRC4(GetNewPID(username), password))

End Function


Private Function GetNewPID(username)

    Dim strPID As String
    
    strPID = username
    
    If (Len(strPID) > 20) Then
        strPID = Left$(strPID, 20)
    Else
        While (Len(strPID) < 4)
            strPID = strPID & "_"
        Wend
    End If
    
    GetNewPID = strPID
    
End Function


Function DisplayIndex(tmpRecordset, lngOrder, blnShowList, tmpGroupElements, ParamArray tmpArguments()) As typTableData

    On Error GoTo TrapError
    
    Dim bytLoop As Byte
    
    Dim lngRow As Long
    Dim lngCol As Long
    
    Dim TempFields As typTableData
    
    If Not tmpRecordset.EOF Then
        tmpRecordset.MoveFirst
        GoSub InitializeGrid
        While tmpRecordset.EOF = False
            With CommonIndex.grdGrid
                .AddRow
                bytLoop = 0
                lngRow = .RowCount
                For lngCol = 1 To tmpGroupElements
                    .CellValue(lngRow, lngCol) = tmpRecordset.Fields(tmpArguments(bytLoop))
                    bytLoop = bytLoop + 1
                Next lngCol
            End With
            tmpRecordset.MoveNext
        Wend
        
        If CommonIndex.grdGrid.RowCount > 1 Then
            If blnShowList Then
                CommonIndex.grdGrid.Redraw = True
                If CommonIndex.grdGrid.HScrollBar.Visible Then
                    Do Until Not CommonIndex.grdGrid.HScrollBar.Visible
                        CommonIndex.grdGrid.Width = CommonIndex.grdGrid.Width + 90
                    Loop
                    GoSub ResizeForm
                End If
                With CommonIndex
                    .grdGrid.Sort lngOrder
                    .grdGrid.Enabled = True
                    .grdGrid.Redraw = True
                    .grdGrid.SetCurCell 1, 1
                    .Show 1
                End With
            End If
        Else
            CommonIndex.grdGrid.CurRow = 1
        End If
    End If
    
    TempFields.strCode = CommonIndex.grdGrid.CellValue(CommonIndex.grdGrid.CurRow, 1)
    TempFields.strFirstField = CommonIndex.grdGrid.CellValue(CommonIndex.grdGrid.CurRow, 2)
    TempFields.strSecondField = CommonIndex.grdGrid.CellValue(CommonIndex.grdGrid.CurRow, 3)
    TempFields.strThirdField = CommonIndex.grdGrid.CellValue(CommonIndex.grdGrid.CurRow, 4)
    TempFields.strFourthField = CommonIndex.grdGrid.CellValue(CommonIndex.grdGrid.CurRow, 5)
    TempFields.strFifthField = CommonIndex.grdGrid.CellValue(CommonIndex.grdGrid.CurRow, 6)
    TempFields.strSixthField = CommonIndex.grdGrid.CellValue(CommonIndex.grdGrid.CurRow, 7)
    TempFields.strSeventhField = CommonIndex.grdGrid.CellValue(CommonIndex.grdGrid.CurRow, 8)
    TempFields.strEighthField = CommonIndex.grdGrid.CellValue(CommonIndex.grdGrid.CurRow, 9)
    
    DisplayIndex = TempFields
    
    Unload CommonIndex
    
    Exit Function
    
TrapError:
    If Err.Number = 3021 Or Err.Number = 91 Or Err.Number = -2147220991 Or Err.Number = 3265 Or Err.Number = 3075 Then
        DisplayIndex = TempFields
        Unload CommonIndex
        Exit Function
    Else
        If Err.Number = 94 Then
            Resume Next
        End If
    End If

InitializeGrid:
    
    ReDim arrFirstElements(1)
    ReDim arrSecondElements(1)
    ReDim arrThirdElements(1)
    ReDim arrFourthElements(1)
    
    Dim bytGroupStart As Byte
    Dim bytArrayIndex As Byte
    
    For bytLoop = 0 To UBound(tmpArguments) + 1
        '�����������
        bytGroupStart = tmpGroupElements
        bytArrayIndex = 1
        While bytLoop < tmpGroupElements
            ReDim Preserve arrFirstElements(UBound(arrFirstElements))
            arrFirstElements(bytArrayIndex) = tmpRecordset(tmpArguments(bytLoop))
            bytLoop = bytLoop + 1
        Wend
        '������ ������
        bytGroupStart = tmpGroupElements + bytGroupStart
        bytArrayIndex = 1
        While bytLoop < bytGroupStart
            ReDim Preserve arrSecondElements(UBound(arrSecondElements) + 1)
            arrSecondElements(bytArrayIndex) = tmpArguments(bytLoop)
            bytArrayIndex = bytArrayIndex + 1
            bytLoop = bytLoop + 1
        Wend
        '������ ������
        bytGroupStart = tmpGroupElements + bytGroupStart
        bytArrayIndex = 1
        While bytLoop < bytGroupStart
            ReDim Preserve arrThirdElements(UBound(arrThirdElements) + 1)
            arrThirdElements(bytArrayIndex) = tmpArguments(bytLoop)
            bytArrayIndex = bytArrayIndex + 1
            bytLoop = bytLoop + 1
        Wend
        '�������� ������
        bytGroupStart = tmpGroupElements + bytGroupStart
        bytArrayIndex = 1
        While bytLoop < bytGroupStart
            ReDim Preserve arrFourthElements(UBound(arrFourthElements) + 1)
            arrFourthElements(bytArrayIndex) = tmpArguments(bytLoop)
            bytArrayIndex = bytArrayIndex + 1
            bytLoop = bytLoop + 1
        Wend
    Next bytLoop
    
    '�������� ������ - ������� - �����
    CommonIndex.grdGrid.Width = 0
    For bytLoop = 1 To tmpGroupElements
        CommonIndex.grdGrid.AddCol.eTextFlags = arrFourthElements(bytLoop)
        CommonIndex.grdGrid.ColHeaderText(bytLoop) = arrSecondElements(bytLoop)
        CommonIndex.grdGrid.ColWidth(bytLoop) = 7 * (arrThirdElements(bytLoop) + 1)
        If arrThirdElements(bytLoop) = 0 Then CommonIndex.grdGrid.ColVisible(bytLoop) = False
        CommonIndex.grdGrid.ColHeaderTextFlags(bytLoop) = 1
    Next bytLoop
    
    With CommonIndex.grdGrid
        .Header.Flat = True
        .Header.Height = 25
    End With
        
    Return
    
ResizeForm:
    
    With CommonIndex
        .shpShape.Width = .grdGrid.Width + 160
        .Width = .shpShape.Width + 470
        .frmButtonFrame.Left = (.Width / 2) - (.frmButtonFrame.Width / 2)
    End With

    Return

End Function

Sub AddDummyLines(grdGrid, ParamArray columns() As Variant)

    Dim lngRow As Long
    Dim lngCol As Long
    Dim lngLoop As Long
    
    For lngRow = 1 To 50
        With grdGrid
            .AddRow
            For lngCol = 1 To (UBound(columns) + 1)
                .CellValue(lngRow, lngCol) = columns(lngCol - 1)
            Next lngCol
        End With
    Next lngRow

End Sub

Function ResetKeyCode(KeyCode As Integer, Shift As Integer)

    Dim CtrlDown
    
    CtrlDown = Shift + vbCtrlMask
    
    If _
        (KeyCode = vbKeyEscape) Or _
        (KeyCode = vbKeyM And CtrlDown > 2) Or _
        (KeyCode = vbKeyN And CtrlDown > 2) Or _
        (KeyCode = vbKeyS And CtrlDown > 2) Or _
        (KeyCode = vbKeyD And CtrlDown > 2) Or _
        (KeyCode = vbKeyP And CtrlDown > 2) Or _
        (KeyCode = vbKeyC And CtrlDown > 2) Or _
        (KeyCode = vbKeyV And CtrlDown > 2) Or _
        (KeyCode = vbKeyF And CtrlDown) > 2 Then KeyCode = 0
    
    ResetKeyCode = KeyCode
    
End Function
Function EditableFields(ParamArray tmpFields())
    
    Dim bytLoop As Byte
    
    For bytLoop = 0 To UBound(tmpFields)
        tmpFields(bytLoop).Editable = True
    Next bytLoop

End Function

Function EnableTabStop(ParamArray tmpFields())
    
    Dim bytLoop As Byte
    
    For bytLoop = 0 To UBound(tmpFields)
        tmpFields(bytLoop).TabStop = True
    Next bytLoop

End Function

Function CheckForAcceptableKey(myKeyCode)

    CheckForAcceptableKey = IIf((myKeyCode >= 48 And myKeyCode <= 57) Or myKeyCode = 46 Or myKeyCode = 44 Or myKeyCode = 45 Or myKeyCode = 8 Or myKeyCode = 13, True, False)

End Function

Function PositionControls(thisForm As Form, formFullScreen As Boolean, Optional grdGrid As iGrid)

    On Error GoTo ErrTrap
    
    Dim ctl As Control
    Dim intLoop As Integer
    
    intLoop = 0
    
    '��� - ���
    For Each ctl In thisForm.Controls
        '�� ����� ������
        If ctl.Name = "frmInfo" Then
            thisForm.frmInfo.Visible = False
        End If
        '�������
        If ctl.Name = "cmdButton" Then
            thisForm.cmdButton(intLoop).ButtonStyle = ebsXPOliveGreen
            intLoop = intLoop + 1
        End If
    Next ctl
    
    '������ �����
    If formFullScreen Then PositionFullScreenControls thisForm, True, grdGrid
    
    '��� ������ �����
    If Not formFullScreen Then PositionCenteredScreenControls thisForm, True, grdGrid
    
    '�������
    For Each ctl In thisForm.Controls
        If ctl.Name = "frmProgress" Then
            With thisForm.frmProgress
                .Visible = False
                .ZOrder 1
                .Top = (thisForm.Height / 2) - (.Height / 2)
                .Left = (thisForm.Width / 2) - (.Width / 2)
                Exit For
            End With
        End If
        If ctl.Name = "frmTotals" Then
            With thisForm.frmTotals
                .Left = (thisForm.frmContainer.Width / 2) - (.Width / 2)
            End With
        End If
    Next ctl
    
    '�������� ����������
    For Each ctl In thisForm.Controls
        If ctl.Name = "lblToday" Then thisForm.lblToday.Caption = format(Date, "dddd dd/mm/yyyy")
    Next ctl
    
    '��������
    Dim intIndex As Integer
    intIndex = 0
    For Each ctl In thisForm.Controls
        If Left(ctl.Name, 11) = "frmCriteria" Then
            With thisForm.frmCriteria(intIndex)
                .Visible = True
                .ZOrder 0
                .Top = ((grdGrid.Height) / 2) - (.Height / 2) + grdGrid.Top
                .Left = (grdGrid.Width / 2) - (.Width / 2) + grdGrid.Left
                intIndex = intIndex + 1
            End With
        End If
    Next ctl

    Exit Function
    
ErrTrap:
    If Err.Number = 438 Then Resume Next '�� ����������� ��� �������

End Function


Function PositionFullScreenControls(thisForm As Form, formFullScreen As Boolean, Optional grdGrid As iGrid, Optional customColours As Boolean)

    Dim ctl As Control
    
    '�����
    With thisForm
        .Top = 350
        .Height = CommonMain.Height - (.Top * 1.2)
        .Width = CommonMain.Width
        .Left = -100
    End With
    
    'Container
    With thisForm.frmContainer
        .Height = thisForm.Height - 520
        .Top = (thisForm.Height / 2) - (.Height / 2)
        .Left = (thisForm.Width / 2) - (.Width / 2)
    End With
    
    '�������
    With thisForm.frmButtonFrame
        .Top = thisForm.frmContainer.Height - 690
        .Left = (thisForm.frmContainer.Width / 2) - (.Width / 2)
    End With
    
    '��������� �������
    With thisForm.shpBackground
        .Top = 975
        .Left = 0
        .Width = thisForm.Width
        .Height = thisForm.frmButtonFrame.Top - 300 - .Top
    End With
    
    '������
    With grdGrid
        .Height = thisForm.shpBackground.Height + 180 - .Top + (thisForm.Top * 2)
        '.ForeColor = vbWhite
        '.HighlightForeColor = vbBlack
        '.HighlightBackColor = &HC0FFC0
    End With
    
    For Each ctl In thisForm.Controls
        '������� ��� ������� �� ������
        If ctl.Name = "frmFrameForGridButtons" Then
            With thisForm.frmFrameForGridButtons
                .Top = thisForm.shpBackground.Height + 550
                .Left = (thisForm.frmContainer.Width / 2) - (.Width / 2)
            End With
            grdGrid.Height = thisForm.Height - 3150 - thisForm.frmFrameForGridButtons.Height
        End If
        '������ ������ - ��������
        If ctl.Name = "frmTotals" Then
            With thisForm.frmTotals
                .Top = thisForm.shpBackground.Height - 190
                .Left = (thisForm.frmContainer.Width / 2) - (.Width / 2)
            End With
            With thisForm.frmDetails
                .Top = thisForm.frmTotals.Top - .Height - 90
                .Left = (thisForm.frmContainer.Width / 2) - (.Width / 2)
            End With
            grdGrid.Height = thisForm.Height - 6190 - thisForm.frmDetails.Height
        End If
    Next ctl
    
End Function


Function PositionCenteredScreenControls(thisForm As Form, formFullScreen As Boolean, Optional grdGrid As iGrid, Optional customColours As Boolean)

    '�����
    thisForm.Width = thisForm.shpRightEdge.Left + thisForm.shpRightEdge.Width
    thisForm.Height = thisForm.shpBottomEdge.Top + thisForm.shpBottomEdge.Height - 90
    thisForm.Left = CommonMain.Width / 2 - thisForm.Width / 2 - 100
    thisForm.Top = CommonMain.Height / 2 - thisForm.Height / 2
    
    '�������
    With thisForm.frmButtonFrame
        .Left = (thisForm.Width / 2) - (thisForm.frmButtonFrame.Width / 2)
    End With
    
    '��������� �������
    With thisForm.shpBackground
        .Top = 900
        .Left = 225
        .Width = thisForm.Width - 470
        .Height = thisForm.frmButtonFrame.Top - 270 - .Top
    End With
    
End Function
Function ColorizeControls(thisForm As Form, Optional fullScreen As Boolean, Optional customColours As Boolean)

    Dim ctl As Control
    Dim objFont As StdFont
    
    If Not customColours Then
        thisForm.BackColor = IIf(fullScreen, GetSetting(strApplicationName, "Colors", "Background Full Screen Forms"), GetSetting(strApplicationName, "Colors", "Forms Centered Background"))
    End If
    
    For Each ctl In thisForm.Controls
        '�������
        If ctl.Name = "cmdButton" Then
            ctl.ForeColor = vbBlack
        End If
        '��������
        If ctl.Name = "frmCriteria" Then
            ctl.BackColor = GetSetting(strApplicationName, "Colors", "Background Criteria")
        End If
        'Container
        If ctl.Name = "frmContainer" Then
            ctl.BackColor = IIf(fullScreen, GetSetting(strApplicationName, "Colors", "Forms FullScreen Background"), GetSetting(strApplicationName, "Colors", "Background Containers"))
        End If
        '�����
        If ctl.Name = "shpBackground" Then
            ctl.BackColor = IIf(fullScreen, GetSetting(strApplicationName, "Colors", "Forms FullScreen Background"), GetSetting(strApplicationName, "Colors", "Frames Background"))
        End If
        '������� ��������
        If ctl.Name = "frmButtonFrame" Or ctl.Name = "frmFrameForGridButtons" Or ctl.Name = "frmTotals" Or ctl.Name = "frmDetails" Then
            ctl.BackColor = thisForm.BackColor
        End If
        '������
        If TypeOf ctl Is iGrid And Not customColours Then
            ctl.BackColor = IIf(fullScreen, GetSetting(appName:=strApplicationName, Section:="Colors", Key:="Grid FullScreen BackColor"), GetSetting(appName:=strApplicationName, Section:="Colors", Key:="Grid BackColor"))
            ctl.GridLines = IIf(fullScreen, GetSetting(appName:=strApplicationName, Section:="Colors", Key:="Grid FullScreen GridLines"), GetSetting(appName:=strApplicationName, Section:="Colors", Key:="Grid GridLines"))
            ctl.ForeColor = IIf(fullScreen, GetSetting(appName:=strApplicationName, Section:="Colors", Key:="Grid FullScreen ForeColor"), GetSetting(appName:=strApplicationName, Section:="Colors", Key:="Grid ForeColor"))
            ctl.HighlightForeColor = IIf(fullScreen, GetSetting(appName:=strApplicationName, Section:="Colors", Key:="Grid FullScreen Highlight ForeColor"), GetSetting(appName:=strApplicationName, Section:="Colors", Key:="Grid Highlight ForeColor"))
            ctl.HighlightBackColor = IIf(fullScreen, GetSetting(appName:=strApplicationName, Section:="Colors", Key:="Grid FullScreen Highlight BackColor"), GetSetting(appName:=strApplicationName, Section:="Colors", Key:="Grid Highlight BackColor"))
        End If
        '��������
        If TypeOf ctl Is Label Then
            Select Case ctl.Name
                '������� �� ����� ��� ������� ������
                Case "lblLabel"
                    ctl.ForeColor = GetSetting(strApplicationName, "Colors", "Labels Normal Foreground")
                    ctl.BackStyle = 0
                '������� �� ������� ���������
                Case "lblCriteriaLabel"
                    ctl.ForeColor = GetSetting(strApplicationName, "Colors", "Labels Criteria Foreground")
                    ctl.BackStyle = 0
                Case "lblSimple"
                    ctl.ForeColor = vbWhite
                    ctl.BackStyle = 0
                    Set objFont = New StdFont
                    objFont.Name = GetSetting(strApplicationName, "Colors", "Labels Title Font")
                    objFont.Size = 10
                    objFont.Bold = False
                    Set ctl.Font = objFont
            End Select
        End If
        '�������� ������������ ������
        If TypeOf ctl Is Label Then
            Select Case ctl.Name
                '�������� ������
                Case "lblTitle"
                    If Not customColours Then
                        ctl.ForeColor = GetSetting(strApplicationName, "Colors", "Labels Title Foreground")
                    End If
                    Set objFont = New StdFont
                    objFont.Name = GetSetting(strApplicationName, "Colors", "Labels Title Font")
                    objFont.Size = 30
                    objFont.Bold = True
                    objFont.Charset = 161
                    Set ctl.Font = objFont
                    Set objFont = Nothing
                Case "lblCriteria"
                    ctl.ForeColor = GetSetting(strApplicationName, "Colors", "Labels Totals Criteria")
            End Select
        End If
        
        'Checkboxes
        If TypeOf ctl Is CheckBox And Not customColours Then
            'Checkbox �� �����
            If Left(ctl.Name, 11) <> "chkCriteria" Then
                ctl.ForeColor = GetSetting(strApplicationName, "Colors", "Checkbox Normal Foreground")
                ctl.BackColor = GetSetting(strApplicationName, "Colors", "Checkbox Normal Background")
            End If
            'Checkbox �� ������� ���������
            If Left(ctl.Name, 11) = "chkCriteria" Then
                ctl.ForeColor = GetSetting(strApplicationName, "Colors", "Labels Criteria Foreground")
                ctl.BackColor = GetSetting(strApplicationName, "Colors", "Background Criteria")
            End If
        End If
        
        'Radios
        If TypeOf ctl Is OptionButton And Not customColours Then
            'Radios �� �����
            If Left(ctl.Name, 11) <> "optCriteria" Then
                ctl.ForeColor = GetSetting(strApplicationName, "Colors", "OptionButton Normal Foreground")
                ctl.BackColor = GetSetting(strApplicationName, "Colors", "OptionButton Normal Background")
            End If
            'Radios �� ������� ���������
            If Left(ctl.Name, 11) = "optCriteria" Then
                ctl.ForeColor = GetSetting(strApplicationName, "Colors", "Labels Criteria Foreground")
                ctl.BackColor = GetSetting(strApplicationName, "Colors", "Labels Criteria Background")
            End If
        End If
        
        'Frames
        If TypeOf ctl Is Frame And Not customColours Then
            If ctl.Tag = "SameColorAsBackground" Then
                ctl.ForeColor = GetSetting(strApplicationName, "Colors", "Frames Foreground")
                ctl.BackColor = GetSetting(strApplicationName, "Colors", "Frames Background")
            End If
        End If
        
    Next
    
End Function




Attribute VB_Name = "ModuleParticular"
Option Explicit

'���. ���������
Global intAccountsCodeLength As Byte
Global blnCustomerCheckTaxNo As Boolean
Global strSalesAccountsCode As String
Global strVATAccountsCode As String
Global intVAT As Integer
Global strAccountsFileName As String
Global strCashAccountsCode As String

'��������
Global blnPreviewInvoices As Boolean
Global intUsualPaymentTermID As Integer
Global strUsualRemarks As String
Global intInvoiceCopies As Byte

'��������
Global intPreviewReports As Integer

Function AddCompanyData(sheet As Object, colCount As Long)

    'Excel
    With sheet
        .Range("A1:" & Chr(colCount + 64) & "1").MergeCells = True
        .Range("A2:" & Chr(colCount + 64) & "2").MergeCells = True
        .Range("A3:" & Chr(colCount + 64) & "3").MergeCells = True
        .Range("A4:" & Chr(colCount + 64) & "4").MergeCells = True
        .Range("A1").Value = arrCompanyData(7)
        .Range("A2").Value = arrCompanyData(8)
        .Range("A3").Value = arrCompanyData(9)
        .Range("A4").Value = arrCompanyData(10)
    End With

End Function

Function AddNumberFormats(sheet As Object, grid As iGrid, format As String, rowOffsetFromTop As Long, ParamArray columns() As Variant)

    Dim column As Long
    Dim row As Long
    
    'Excel
    With sheet
        For column = 0 To UBound(columns)
            Select Case format
                Case "Floats"
                    For row = 1 To grid.RowCount
                        .Range(columns(column) & row + rowOffsetFromTop).NumberFormat = "#,##0.00_);[Red]-#,##0.00 "
                    Next row
                Case "Integers"
                    For row = 1 To grid.RowCount
                        .Range(columns(column) & row + rowOffsetFromTop).NumberFormat = "#,##0_);[Red]-#,##0 "
                    Next row
                Case "Dates"
                    For row = 1 To grid.RowCount
                        .Range(columns(column) & row + rowOffsetFromTop).NumberFormat = "dd-mm-yyyy"
                    Next row
            End Select
        Next column
    End With

End Function

Function CreateSELECTStatement(InvoiceMasterRefersTo As String)

    Dim strSQL As String
    
    '�����
    If InvoiceMasterRefersTo = "1" Then
        strSQL = "SELECT " _
            & "InvoiceID, InvoiceMasterRefersTo, InvoiceSecondaryRefersTo, InvoiceDateIssue, InvoiceTrnID, InvoiceNo, " _
            & "Description, " _
            & "CodeShortDescriptionB, CodeBatch, CodeSuppliers, CodeDescription, CodeCustomers, CodeDescription, " _
            & "InvoiceInAmount, " _
            & "Amount, " _
            & "PaymentTermCreditID, " _
            & "ExpenseCategoryDescription " _
            & "FROM ((((((Invoices " _
            & "INNER JOIN Suppliers ON Invoices.InvoicePersonID = Suppliers.ID) " _
            & "INNER JOIN Codes ON Invoices.InvoiceCodeID = Codes.CodeID) " _
            & "LEFT JOIN InvoicesIn ON Invoices.InvoiceTrnID = InvoicesIn.InvoiceInTrnID) " _
            & "LEFT JOIN PaymentOut ON Invoices.InvoiceTrnID = PaymentOut.TrnID) " _
            & "LEFT JOIN PaymentTerms ON InvoicesIn.InvoiceInPaymentTermID = PaymentTerms.PaymentTermID) " _
            & "LEFT JOIN ExpensesCategories ON InvoicesIn.InvoiceInExpenseCategoryID = ExpensesCategories.ExpenseCategoryID) "
        End If
    
    '�����
    If InvoiceMasterRefersTo = "2" Then
        strSQL = "SELECT " _
            & "InvoiceID, InvoiceMasterRefersTo, InvoiceSecondaryRefersTo, InvoiceDateIssue, InvoiceTrnID, InvoiceNo, " _
            & "Description, " _
            & "CodeShortDescriptionB, CodeBatch, CodeCustomers, CodeDescription, " _
            & "InvoiceOutAdultsWithTransfer, InvoiceOutKidsWithTransfer, InvoiceOutFreeWithTransfer, InvoiceOutAdultsWithoutTransfer, InvoiceOutKidsWithoutTransfer, InvoiceOutFreeWithoutTransfer, InvoiceOutAdultsAmountWithTransfer, InvoiceOutAdultsAmountWithoutTransfer, InvoiceOutKidsAmountWithTransfer, InvoiceOutKidsAmountWithoutTransfer, InvoiceOutDirectAmount, " _
            & "Amount, " _
            & "PaymentTermCreditID, " _
            & "DestinationDescription " _
            & "FROM ((((((Invoices " _
            & "INNER JOIN Customers ON Invoices.InvoicePersonID = Customers.ID) " _
            & "INNER JOIN Codes ON Invoices.InvoiceCodeID = Codes.CodeID) " _
            & "LEFT JOIN InvoicesOut ON Invoices.InvoiceTrnID = InvoicesOut.InvoiceOutTrnID) " _
            & "LEFT JOIN PaymentIn ON Invoices.InvoiceTrnID = PaymentIn.TrnID) " _
            & "LEFT JOIN PaymentTerms ON InvoicesOut.InvoiceOutPaymentTermID = PaymentTerms.PaymentTermID) " _
            & "LEFT JOIN Destinations ON InvoicesOut.InvoiceOutDestinationID = Destinations.DestinationID) "
        End If
        
    CreateSELECTStatement = strSQL

End Function


Function AddOneToTheLastRecord(myTable)

    On Error GoTo ErrTrap
    
    Dim strSQL As String
    Dim rsInvoices As Recordset
    
    strSQL = "SELECT InvoiceTrnID FROM " & myTable & " ORDER BY InvoiceTrnID"
    Set rsInvoices = CommonDB.OpenRecordset(strSQL)
    
    With rsInvoices
        If .EOF Then
            AddOneToTheLastRecord = 1
        Else
            .MoveLast
            AddOneToTheLastRecord = !InvoiceTrnID + 1
        End If
        .Close
    End With
    
    Set rsInvoices = Nothing
    
    Exit Function
    
ErrTrap:
    Exit Function

End Function


Function CalculateFields(rstTransactions As Recordset, strSign As String, ParamArray Fields() As Variant)

    Dim intLoop As Integer
    Dim curAmount
    
    curAmount = 0
    
    With rstTransactions
        For intLoop = 0 To UBound(Fields)
            curAmount = curAmount + rstTransactions.Fields(Fields(intLoop))
        Next intLoop
    End With
    
    curAmount = IIf(strSign = "+", curAmount, -curAmount)
    
    CalculateFields = curAmount

End Function


Function CheckForDuplicateInvoice(myDate, myCodeID, myInvoiceNo)

    On Error GoTo ErrTrap
    
    'Local variables
    Dim intIndex As Byte
    Dim strThisQuery As String
    Dim strParameters As String
    Dim strParFields As String
    Dim strThisParameter As String
    Dim strOrder As String
    Dim strLogic As String
    Dim arrQuery() As Variant
    Dim strSQL As String
    Dim lngRow As Long
    Dim rstTrips As Recordset
    Dim intYear As Integer
    Dim intInvoiceNo As Integer
    Dim lngCodeID As Long
    
    '������� �����
    intIndex = 0
    lngRow = 0
    Set TempQuery = CommonDB.CreateQueryDef("")
    
    '����� SQL - 1� ������� - ������ �� � ������� ��� ���� ����� ��� �� Codes.CodeLastNo + 1 ����� ��� �������������
    strSQL = "SELECT InvoiceDateIssue, InvoiceCodeID, InvoiceNo " _
        & "FROM Invoices "
        
    '�����
    strThisParameter = "intYear Integer"
    strThisQuery = "Year(InvoiceDateIssue) = intYear"
    strLogic = " AND "
    GoSub UpdateSQLString
    arrQuery(intIndex) = Year(myDate)
    
    'ID ���������
    strThisParameter = "lngInvoiceID Long"
    strThisQuery = "InvoiceCodeID = lngInvoiceID"
    strLogic = " AND "
    GoSub UpdateSQLString
    arrQuery(intIndex) = Val(myCodeID)
    
    '�� ���������
    strThisParameter = "strInvoiceNo String"
    strThisQuery = "InvoiceNo = strInvoiceNo"
    strLogic = " AND "
    GoSub UpdateSQLString
    arrQuery(intIndex) = myInvoiceNo
    
    '�������� �� ��������
    If strThisParameter <> "" Then
        strParameters = "PARAMETERS " & strParameters & "; "
        strParFields = "WHERE " & strParFields
        strSQL = strParameters & strSQL & strParFields
    End If
    
    TempQuery.SQL = strSQL & strOrder
    
    If strThisParameter <> "" Then
        For intIndex = 1 To UBound(arrQuery)
            TempQuery.Parameters(intIndex - 1) = arrQuery(intIndex)
        Next intIndex
    End If
    
    '������ �� recordset
    Set rstTrips = TempQuery.OpenRecordset()
    
    '������ ��� �������������
    With rstTrips
        If Not .EOF Then
            rstTrips.MoveLast
            CheckForDuplicateInvoice = True
            Exit Function
        End If
    End With
    
    '����� SQL - 2� ������� - ������ �� � ������� ��� ���� ����� ��� �� Codes.CodeLastNo + 1 ����� � �������� ��� ��� ��������� ������������ ��� Sales.TripInvoiceNo
    Set TempQuery = CommonDB.CreateQueryDef("")
    
    strSQL = "SELECT InvoiceDateIssue, InvoiceCodeID, InvoiceNo " _
        & "FROM Invoices "
    
    strOrder = " ORDER BY InvoiceDateIssue, val(InvoiceNo)"
    
    intIndex = 0
    strParameters = ""
    strParFields = ""
    
    '�����
    strThisParameter = "intYear Integer"
    strThisQuery = "Year(InvoiceDateIssue) = intYear"
    strLogic = " AND "
    GoSub UpdateSQLString
    arrQuery(intIndex) = Year(myDate)
    
    'ID ���������
    strThisParameter = "lngInvoiceID Long"
    strThisQuery = "InvoiceCodeID = lngInvoiceID"
    strLogic = " AND "
    GoSub UpdateSQLString
    arrQuery(intIndex) = Val(myCodeID)
    
    '�������� �� ��������
    If strThisParameter <> "" Then
        strParameters = "PARAMETERS " & strParameters & "; "
        strParFields = "WHERE " & strParFields
        strSQL = strParameters & strSQL & strParFields
    End If
    
    TempQuery.SQL = strSQL & strOrder
    
    For intIndex = 1 To UBound(arrQuery)
        TempQuery.Parameters(intIndex - 1) = arrQuery(intIndex)
    Next intIndex
    
    '������ �� recordset
    Set rstTrips = TempQuery.OpenRecordset()
    
    '������ ��� �������������
    With rstTrips
        If .RecordCount > 0 Then
            .MoveLast
            If rstTrips!InvoiceNo + 1 <> Int(myInvoiceNo) Then CheckForDuplicateInvoice = True
        Else
            CheckForDuplicateInvoice = False
        End If
        .Close
    End With
    
    Exit Function

UpdateSQLString:
    intIndex = intIndex + 1
    strParameters = IIf(intIndex > 1, strParameters & ", ", strParameters)
    strParFields = IIf(intIndex > 1, strParFields & strLogic, strParFields)
    strParameters = strParameters & strThisParameter
    strParFields = strParFields & strThisQuery
    ReDim Preserve arrQuery(intIndex)
    
    Return
    
ErrTrap:
    blnErrors = True
    DisplayErrorMessage True, Err.Description
    
End Function


Function CheckDate(myDate, myTitle)

    CheckDate = False
    
    '����������
    If myDate = "" Then
        If MyMsgBox(4, myTitle, strStandardMessages(1), 1) Then
        End If
        Exit Function
    End If
    
    '����������
    If Len(myDate) <> 10 Then
        If MyMsgBox(4, myTitle, strStandardMessages(2), 1) Then
        End If
        Exit Function
    End If
    
    CheckDate = True

End Function


Function CheckTaxNo(tmpVAT)

    On Error GoTo ErrTrap
    
    'Local variables
    Dim lngSum As Long
    Dim lngRemainder As Long
    
    CheckTaxNo = True
    
    If Len(tmpVAT.text) <> 9 Then
        CheckTaxNo = False
        Exit Function
    End If
    
    lngSum = 256 * Mid(tmpVAT.text, 1, 1) + 128 * Mid(tmpVAT.text, 2, 1) + 64 * Mid(tmpVAT.text, 3, 1) + 32 * Mid(tmpVAT.text, 4, 1) + 16 * Mid(tmpVAT.text, 5, 1) + 8 * Mid(tmpVAT.text, 6, 1) + 4 * Mid(tmpVAT.text, 7, 1) + 2 * Mid(tmpVAT.text, 8, 1)
    lngRemainder = lngSum Mod 11
    If lngRemainder = 10 Then
        lngRemainder = 0
    End If
    If Val(Right(tmpVAT.text, 1)) <> lngRemainder Then
        CheckTaxNo = False
    End If
    Exit Function
    
ErrTrap:
    If Err.Number = 13 Then
        CheckTaxNo = False
        Exit Function
    End If
    
End Function


Function FindWeekDay(myDate)
    
    FindWeekDay = ""
    
    If IsDate(myDate) Then FindWeekDay = WeekdayName(Weekday(myDate, vbUseSystemDayOfWeek))

End Function


Function FullInvoice(CodeShortDescriptionB, CodeBatch, InvoiceNo)

    Dim strFullInvoice As String
    
    strFullInvoice = CodeShortDescriptionB & Space(3 - Len(CodeShortDescriptionB)) & " "
    strFullInvoice = strFullInvoice & IIf(CodeBatch <> "", CodeBatch, "0") & " "
    strFullInvoice = strFullInvoice & Right("00000" & InvoiceNo, 5)
    
    FullInvoice = strFullInvoice

End Function

Function LoadParameters()
    
    On Error GoTo ErrTrap
    
    Dim intLoop As Integer
    Dim intUpper As Integer
    
    Dim TableSettings As TableDef
    
    Dim rsParameters As Recordset
    
    Set TableSettings = dBaseTables("Settings")
    Set rsParameters = TableSettings.OpenRecordset()
    
    With rsParameters
        .MoveFirst
        '���������� ��������
        arrCompanyData(1) = ![Line01]
        arrCompanyData(2) = ![Line02]
        arrCompanyData(3) = ![Line03]
        arrCompanyData(4) = ![Line04]
        arrCompanyData(5) = ![Line05]
        arrCompanyData(6) = ![Line06]
        '������������ ��������
        arrCompanyData(7) = ![Line07]
        arrCompanyData(8) = ![Line08]
        arrCompanyData(9) = ![Line09]
        arrCompanyData(10) = ![Line10]
        '���. ���������
        intAccountsCodeLength = !AccountsCodeLength
        strSalesAccountsCode = !SalesAccountsCode
        strVATAccountsCode = !VATAccountsCode
        intVAT = !VAT
        strCashAccountsCode = !CashAccountsCode
        '��������
        intUsualPaymentTermID = !UsualPaymentTermID
        strUsualRemarks = !UsualRemarks
        '�������� ����������
        strAccountsFileName = !fileName
        intInvoiceCopies = !InvoiceCopies
        blnPreviewInvoices = !PreviewInvoicesID
        '��������
        intPreviewReports = !PreviewReportsID
        .Close
    End With
    
    LoadParameters = True
    
    Exit Function
    
ErrTrap:
    LoadParameters = False
    DisplayErrorMessage True, Err.Description
    
End Function



Function CreateReferenceNo(DestinationShortDescription, transferDate, TransferID)

    Dim DayOfWeek As String
    
    Dim Days(7) As String
    
    Days(1) = "SUNDAY"
    Days(2) = "MONDAY"
    Days(3) = "TUESDAY"
    Days(4) = "WEDNESDAY"
    Days(5) = "THURSDAY"
    Days(6) = "FRIDAY"
    Days(7) = "SATURDAY"
    
    DayOfWeek = Left(Days(Weekday(transferDate)), 1)
    
    CreateReferenceNo = DestinationShortDescription & "" & DayOfWeek & "-" & Right("00000" & TransferID, 5)

End Function


Function AddNumbers(ParamArray Numbers() As Variant)

    'Local variables
    Dim intLoop As Integer
    Dim curTotal As Currency
    
    '������ ������
    For intLoop = 0 To UBound(Numbers)
        curTotal = curTotal + CCur(Numbers(intLoop))
    Next intLoop
    
    '���������
    AddNumbers = curTotal

End Function



Function CalculateExcursionCharges(lngCustomerID, lngDestinationID, mskDate, mskPersons, lngFieldNo)

    'Local recordsets
    Dim rstPrices As Recordset
    
    '������� �����
    Set TempQuery = CommonDB.CreateQueryDef("")
    CalculateExcursionCharges = 0
    
    '�� ��� ����� ������� ��� ���������
    If lngCustomerID <> "" And lngDestinationID <> "" And mskDate <> "" And Not IsNull(mskPersons) And mskPersons <> "" Then
        '������ ��� �������� �� ��� �������� ��� ��������
        TempQuery.SQL = "PARAMETERS lngCustomerID Long, lngDestinationID Long; " _
        & "SELECT PriceFrom, PriceTo, PriceAdultWithTransfer, PriceKidWithTransfer, PriceAdultWithoutTransfer, PriceKidWithoutTransfer " _
        & "FROM Prices " _
        & "WHERE PriceCustomerID = lngCustomerID AND PriceDestinationID = lngDestinationID"
        TempQuery![lngCustomerID] = lngCustomerID
        TempQuery![lngDestinationID] = lngDestinationID
    
        Set rstPrices = TempQuery.OpenRecordset
        
        With rstPrices
            '�� ������� �������������
            If Not .EOF Then
                .MoveFirst
                While Not .EOF
                    '������ ��� ������� ���� � ���������� �������� ����� ���� ��� ���� ��� �������������
                    If CDate(mskDate) >= ![PriceFrom] And CDate(mskDate) <= ![PriceTo] Then
                        CalculateExcursionCharges = .Fields(lngFieldNo) * CCur(mskPersons)
                        .Close
                        Exit Function
                    Else
                        .MoveNext
                    End If
                Wend
            End If
        End With
    End If
    
End Function



Function SetFontNameAndSize(sheet As Object, fontName As String, fontSize As Integer)

    'Excel
    With sheet
        .Range("A1:Z9999").Font.Name = fontName
        .Range("A1:Z9999").Font.Size = fontSize
    End With

End Function



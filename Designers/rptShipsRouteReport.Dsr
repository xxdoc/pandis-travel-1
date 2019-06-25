VERSION 5.00
Begin {9EB8768B-CDFA-44DF-8F3E-857A8405E1DB} rptShipsRouteReport 
   Caption         =   "����� ������������"
   ClientHeight    =   15630
   ClientLeft      =   120
   ClientTop       =   450
   ClientWidth     =   28560
   StartUpPosition =   2  'CenterScreen
   _ExtentX        =   50377
   _ExtentY        =   27570
   SectionData     =   "rptShipsRouteReport.dsx":0000
End
Attribute VB_Name = "rptShipsRouteReport"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Dim lngRow As Long
Dim blnError As Boolean
Private blnLastPage As Boolean

Private Sub ActiveReport_DataInitialize()

    lngRow = 0
    blnError = False
    
    Fields.RemoveAll
    
    Fields.Add "AA"
    Fields.Add "LastName"
    Fields.Add "FirstName"
    Fields.Add "GenderDescription"
    Fields.Add "AgeDescription"
    Fields.Add "OccupantDescription"
    Fields.Add "Care"
    Fields.Add "Remarks"
    
End Sub

Private Sub ActiveReport_FetchData(EOF As Boolean)

    If blnError Then Exit Sub
    
    lngRow = lngRow + 1
    
    With ShipsRouteReport
        
        If lngRow > .grdShipsRouteReport.RowCount Then
            blnLastPage = True
            EOF = True
            Exit Sub
        End If
        
        Fields("AA") = .grdShipsRouteReport.CellValue(lngRow, "AA")
        Fields("LastName") = .grdShipsRouteReport.CellValue(lngRow, "LastName")
        Fields("FirstName") = .grdShipsRouteReport.CellValue(lngRow, "FirstName")
        Fields("GenderDescription") = .grdShipsRouteReport.CellText(lngRow, "Gender")
        Fields("AgeDescription") = .grdShipsRouteReport.CellText(lngRow, "Age")
        Fields("OccupantDescription") = .grdShipsRouteReport.CellText(lngRow, "OccupantDescription")
        Fields("Care") = .grdShipsRouteReport.CellValue(lngRow, "Care")
        Fields("Remarks") = .grdShipsRouteReport.CellValue(lngRow, "Remarks")
        
        EOF = False
        blnLastPage = False
        
    End With
        
End Sub

Private Sub ReportHeader_Format()

    On Error GoTo ErrTrap
    
    With ShipsRouteReport
        
        '���������� �������� ��������
        lblCompanyData.Caption = arrCompanyData(1) & Chr(13) & arrCompanyData(2) & Chr(13) & arrCompanyData(3) & Chr(13) & arrCompanyData(4) & Chr(13) & arrCompanyData(5) & Chr(13) & arrCompanyData(6)
        
        '����������
        lblDate.Caption = "���������� ��� " + .mskDate.text
        
        '������ �� �������� ��� ������ ��� ��� ��������� ���������� ��������
        Dim strSQL As String
        Dim rstRecordset As Recordset
        
        strSQL = "SELECT ShipDescription, ShipFlag, ShipRegistryNo, ShipIMO, ShipManagerName, ShipManagerNameInGreece, ShipManagerAgent, " _
            & "ShipManagerAPersonName, ShipManagerAPersonPhones, ShipManagerAPersonEmail, ShipManagerAPersonFax, ShipManagerAPersonAddress, " _
            & "ShipManagerBPersonName, ShipManagerBPersonPhones, ShipManagerBPersonEmail, ShipManagerBPersonFax, ShipManagerBPersonAddress " _
        & "FROM Ships " _
        & "INNER JOIN ShipsManagers ON Ships.ShipID = ShipsManagers.ShipManagerShipID " _
        & "WHERE Ships.ShipID = " & ShipsRouteReport.txtShipID.text
        
        Set TempQuery = CommonDB.CreateQueryDef(""): TempQuery.SQL = strSQL
        Set rstRecordset = TempQuery.OpenRecordset()
        
        '�����
        lblShipDescription.Caption = rstRecordset!ShipDescription
        lblShipFlag.Caption = rstRecordset!ShipFlag
        lblShipRegistryNo.Caption = rstRecordset!ShipRegistryNo
        lblShipIMO.Caption = rstRecordset!ShipIMO
        
        '����������
        lblRouteFrom.Caption = ShipsRouteReport.txtFrom.text
        lblRouteVia.Caption = ShipsRouteReport.txtVia.text
        lblRouteTo.Caption = ShipsRouteReport.txtTo.text
        
        '�������
        lblManager.Caption = rstRecordset!ShipManagerName
        lblManagerInGreece.Caption = rstRecordset!ShipManagerNameInGreece
        lblAgent.Caption = rstRecordset!ShipManagerAgent
        
        '��������� ����������
        lblAPersonName.Caption = rstRecordset!ShipManagerAPersonName
        lblAPersonPhones.Caption = rstRecordset!ShipManagerAPersonPhones
        lblAPersonEmail.Caption = rstRecordset!ShipManagerAPersonEmail
        lblAPersonFax.Caption = rstRecordset!ShipManagerAPersonFax
        lblAPersonAddress.Caption = rstRecordset!ShipManagerAPersonAddress
        
        '�������������� ��������� ����������
        lblBPersonName.Caption = rstRecordset!ShipManagerBPersonName
        lblBPersonPhones.Caption = rstRecordset!ShipManagerBPersonPhones
        lblBPersonEmail.Caption = rstRecordset!ShipManagerBPersonEmail
        lblBPersonFax.Caption = rstRecordset!ShipManagerBPersonFax
        lblBPersonAddress.Caption = rstRecordset!ShipManagerBPersonAddress
        
        '�������� ����� ��� �� ���������
        lblPerson.Caption = rstRecordset!ShipManagerName
        
        '�����������
        lblPlace.Caption = "�������"
        lblFooterDate.Caption = format(Date, "dd/mm/yyyy")
        lblTime.Caption = format(Time, "hh:mm")
        
    End With
    
    Exit Sub
    
ErrTrap:
    blnError = True
    DisplayErrorMessage True, Err.Description

End Sub



VERSION 5.00
Object = "{396F7AC0-A0DD-11D3-93EC-00C0DFE7442A}#1.0#0"; "ImageList.ocx"
Object = "{158C2A77-1CCD-44C8-AF42-AA199C5DCC6C}#1.0#0"; "dcButton.ocx"
Object = "{55473EAC-7715-4257-B5EF-6E14EBD6A5DD}#1.0#0"; "ProgressBar.ocx"
Object = "{839D0F5D-B7D7-41B7-A3B4-85D69300B8C1}#98.0#0"; "iGrid300_10Tec.ocx"
Object = "{FFE4AEB4-0D55-4004-ADF2-3C1C84D17A72}#1.0#0"; "userControls.ocx"
Begin VB.Form UtilsBatchPriceListEdit 
   Appearance      =   0  'Flat
   BackColor       =   &H00C0C0C0&
   BorderStyle     =   0  'None
   ClientHeight    =   9930
   ClientLeft      =   0
   ClientTop       =   0
   ClientWidth     =   16665
   ControlBox      =   0   'False
   ForeColor       =   &H00800000&
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   9930
   ScaleWidth      =   16665
   ShowInTaskbar   =   0   'False
   Begin VB.Frame frmProgress 
      BackColor       =   &H00800080&
      BorderStyle     =   0  'None
      Caption         =   "Frame1"
      Height          =   990
      Left            =   5175
      TabIndex        =   15
      Top             =   5325
      Width           =   4065
      Begin vbalProgBarLib6.vbalProgressBar prgProgressBar 
         Height          =   465
         Left            =   150
         TabIndex        =   16
         TabStop         =   0   'False
         Top             =   375
         Width           =   3765
         _ExtentX        =   6641
         _ExtentY        =   820
         Picture         =   "UtilsBatchPriceListEdit.frx":0000
         ForeColor       =   0
         Appearance      =   0
         BarPicture      =   "UtilsBatchPriceListEdit.frx":001C
         BarPictureMode  =   0
         BackPictureMode =   0
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   161
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
      End
      Begin VB.Label lblMaster 
         Alignment       =   2  'Center
         BackStyle       =   0  'Transparent
         Caption         =   "������"
         BeginProperty Font 
            Name            =   "Ubuntu Condensed"
            Size            =   9.75
            Charset         =   161
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00FFFFFF&
         Height          =   240
         Left            =   150
         TabIndex        =   17
         Top             =   75
         Width           =   3765
      End
   End
   Begin VB.Frame frmContainer 
      Appearance      =   0  'Flat
      BackColor       =   &H00E0E0E0&
      BorderStyle     =   0  'None
      ForeColor       =   &H80000008&
      Height          =   9765
      Left            =   75
      TabIndex        =   0
      Top             =   75
      Width           =   16515
      Begin VB.Frame frmButtonFrame 
         BackColor       =   &H00FF8080&
         BorderStyle     =   0  'None
         Height          =   690
         Left            =   75
         TabIndex        =   19
         Top             =   9000
         Width           =   6090
         Begin Dacara_dcButton.dcButton cmdButton 
            Height          =   690
            Index           =   0
            Left            =   225
            TabIndex        =   20
            TabStop         =   0   'False
            Top             =   0
            Width           =   1365
            _ExtentX        =   2408
            _ExtentY        =   1217
            BackColor       =   12640511
            ButtonShape     =   3
            ButtonStyle     =   4
            Caption         =   "��������"
            BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
               Name            =   "Ubuntu Condensed"
               Size            =   9.75
               Charset         =   161
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            PicOpacity      =   0
         End
         Begin Dacara_dcButton.dcButton cmdButton 
            Height          =   690
            Index           =   3
            Left            =   4500
            TabIndex        =   21
            TabStop         =   0   'False
            Top             =   0
            Width           =   1365
            _ExtentX        =   2408
            _ExtentY        =   1217
            BackColor       =   8421631
            ButtonShape     =   3
            ButtonStyle     =   4
            Caption         =   "��������"
            BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
               Name            =   "Ubuntu Condensed"
               Size            =   9.75
               Charset         =   161
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            PicOpacity      =   0
         End
         Begin Dacara_dcButton.dcButton cmdButton 
            Height          =   690
            Index           =   1
            Left            =   1650
            TabIndex        =   22
            TabStop         =   0   'False
            Top             =   0
            Width           =   1365
            _ExtentX        =   2408
            _ExtentY        =   1217
            BackColor       =   12640511
            ButtonShape     =   3
            ButtonStyle     =   4
            Caption         =   "����������"
            BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
               Name            =   "Ubuntu Condensed"
               Size            =   9.75
               Charset         =   161
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            PicOpacity      =   0
         End
         Begin Dacara_dcButton.dcButton cmdButton 
            Height          =   690
            Index           =   2
            Left            =   3075
            TabIndex        =   23
            TabStop         =   0   'False
            Top             =   0
            Width           =   1365
            _ExtentX        =   2408
            _ExtentY        =   1217
            BackColor       =   12640511
            ButtonShape     =   3
            ButtonStyle     =   4
            Caption         =   "��������"
            BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
               Name            =   "Ubuntu Condensed"
               Size            =   9.75
               Charset         =   161
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            PicOpacity      =   0
         End
      End
      Begin VB.Frame frmCriteria 
         BackColor       =   &H00FFC0FF&
         BorderStyle     =   0  'None
         Height          =   2565
         Left            =   300
         TabIndex        =   6
         Top             =   6300
         Width           =   7440
         Begin UserControls.newText txtCustomerDescription 
            Height          =   465
            Left            =   1575
            TabIndex        =   2
            Top             =   1350
            Width           =   4965
            _ExtentX        =   8758
            _ExtentY        =   820
            ForeColor       =   0
            MaxLength       =   40
            Text            =   "AAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAA"
            BackColor       =   4210688
            BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
               Name            =   "Ubuntu Condensed"
               Size            =   11.25
               Charset         =   161
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
         End
         Begin UserControls.newText txtYear 
            Height          =   465
            Left            =   1575
            TabIndex        =   1
            Top             =   825
            Width           =   690
            _ExtentX        =   1217
            _ExtentY        =   820
            Alignment       =   2
            ForeColor       =   4194304
            MaxLength       =   4
            Text            =   "9999"
            BackColor       =   4210688
            BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
               Name            =   "Ubuntu Condensed"
               Size            =   11.25
               Charset         =   161
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
         End
         Begin Dacara_dcButton.dcButton cmdIndex 
            Height          =   465
            Index           =   0
            Left            =   6600
            TabIndex        =   18
            TabStop         =   0   'False
            Top             =   1350
            Width           =   390
            _ExtentX        =   688
            _ExtentY        =   820
            BackColor       =   14742518
            ButtonShape     =   3
            ButtonStyle     =   8
            Caption         =   ""
            BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
               Name            =   "Ubuntu Condensed"
               Size            =   9.75
               Charset         =   161
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   0
            PicNormal       =   "UtilsBatchPriceListEdit.frx":0038
            PicSizeH        =   16
            PicSizeW        =   16
         End
         Begin VB.Label Label1 
            BackColor       =   &H00808000&
            BeginProperty Font 
               Name            =   "Arial"
               Size            =   14.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H00FFFFFF&
            Height          =   465
            Index           =   4
            Left            =   0
            TabIndex        =   11
            Top             =   2100
            Width           =   7440
         End
         Begin VB.Label lblToday 
            Alignment       =   1  'Right Justify
            BackColor       =   &H00808000&
            Caption         =   "01/05/2017"
            BeginProperty Font 
               Name            =   "Aka-Acid-Steelfish"
               Size            =   14.25
               Charset         =   161
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H0000FFFF&
            Height          =   345
            Left            =   2925
            TabIndex        =   10
            Top             =   75
            Width           =   4365
         End
         Begin VB.Label Label1 
            BackColor       =   &H00808000&
            Caption         =   "�������� ����������"
            BeginProperty Font 
               Name            =   "Aka-Acid-Steelfish"
               Size            =   14.25
               Charset         =   161
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H00FFFFFF&
            Height          =   390
            Index           =   3
            Left            =   150
            TabIndex        =   9
            Top             =   75
            Width           =   1665
         End
         Begin VB.Label lblLabel 
            BackColor       =   &H000080FF&
            Caption         =   "�������"
            BeginProperty Font 
               Name            =   "Ubuntu Condensed"
               Size            =   9.75
               Charset         =   161
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H00FFFFFF&
            Height          =   315
            Index           =   0
            Left            =   450
            TabIndex        =   8
            Top             =   1425
            Width           =   690
         End
         Begin VB.Label lblLabel 
            BackColor       =   &H000080FF&
            Caption         =   "����"
            BeginProperty Font 
               Name            =   "Ubuntu Condensed"
               Size            =   9.75
               Charset         =   161
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H00FFFFFF&
            Height          =   315
            Index           =   1
            Left            =   450
            TabIndex        =   7
            Top             =   900
            Width           =   690
         End
         Begin VB.Shape shpWedge 
            BackColor       =   &H0000FFFF&
            BackStyle       =   1  'Opaque
            BorderStyle     =   0  'Transparent
            FillColor       =   &H00008000&
            Height          =   840
            Index           =   0
            Left            =   0
            Top             =   900
            Visible         =   0   'False
            Width           =   465
         End
         Begin VB.Shape shpWedge 
            BackColor       =   &H0000FFFF&
            BackStyle       =   1  'Opaque
            BorderStyle     =   0  'Transparent
            FillColor       =   &H00008000&
            Height          =   840
            Index           =   1
            Left            =   1125
            Top             =   825
            Visible         =   0   'False
            Width           =   465
         End
         Begin VB.Shape shpWedge 
            BackColor       =   &H0000FFFF&
            BackStyle       =   1  'Opaque
            BorderStyle     =   0  'Transparent
            FillColor       =   &H00008000&
            Height          =   840
            Index           =   2
            Left            =   6975
            Top             =   1200
            Visible         =   0   'False
            Width           =   465
         End
         Begin VB.Label Label1 
            BackColor       =   &H00808000&
            BeginProperty Font 
               Name            =   "Arial"
               Size            =   14.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H00FFFFFF&
            Height          =   540
            Index           =   0
            Left            =   0
            TabIndex        =   12
            Top             =   0
            Width           =   7440
         End
      End
      Begin VB.Frame frmInfo 
         Appearance      =   0  'Flat
         BackColor       =   &H8000000D&
         BorderStyle     =   0  'None
         Enabled         =   0   'False
         ForeColor       =   &H80000008&
         Height          =   1065
         Left            =   300
         TabIndex        =   3
         Top             =   5175
         Width           =   4740
         Begin VB.TextBox txtCustomerID 
            Appearance      =   0  'Flat
            BackColor       =   &H00C0C0FF&
            BorderStyle     =   0  'None
            BeginProperty Font 
               Name            =   "Ubuntu Condensed"
               Size            =   9.75
               Charset         =   161
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H00000000&
            Height          =   315
            Left            =   3675
            TabIndex        =   5
            TabStop         =   0   'False
            Top             =   75
            Width           =   1000
         End
         Begin VB.TextBox Text2 
            Appearance      =   0  'Flat
            BackColor       =   &H00C0C0FF&
            BorderStyle     =   0  'None
            BeginProperty Font 
               Name            =   "Ubuntu Condensed"
               Size            =   9.75
               Charset         =   161
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H00000000&
            Height          =   315
            Left            =   75
            TabIndex        =   4
            TabStop         =   0   'False
            Text            =   "Customers.ID"
            Top             =   75
            Width           =   3540
         End
         Begin vbalIml6.vbalImageList lstIconList 
            Left            =   75
            Top             =   450
            _ExtentX        =   953
            _ExtentY        =   953
            Size            =   2296
            Images          =   "UtilsBatchPriceListEdit.frx":05D2
            Version         =   131072
            KeyCount        =   2
            Keys            =   "�"
         End
      End
      Begin iGrid300_10Tec.iGrid grdBatchPriceEdit 
         Height          =   7440
         Left            =   225
         TabIndex        =   13
         TabStop         =   0   'False
         Top             =   1500
         Width           =   16065
         _ExtentX        =   28337
         _ExtentY        =   13123
         Appearance      =   0
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Ubuntu Condensed"
            Size            =   9.75
            Charset         =   161
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   -2147483631
      End
      Begin VB.Label lblTitle 
         AutoSize        =   -1  'True
         BackColor       =   &H00800000&
         BackStyle       =   0  'Transparent
         Caption         =   "������ ����������� �������������"
         BeginProperty Font 
            Name            =   "Ubuntu Condensed"
            Size            =   30
            Charset         =   161
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H000080FF&
         Height          =   720
         Left            =   225
         TabIndex        =   14
         Top             =   75
         Width           =   8445
      End
      Begin VB.Shape shpBackground 
         BackColor       =   &H00C0FFFF&
         BackStyle       =   1  'Opaque
         BorderStyle     =   0  'Transparent
         Height          =   840
         Left            =   0
         Top             =   0
         Width           =   840
      End
   End
   Begin VB.Menu mnuHdrPopUp 
      Caption         =   "mnuHdrPopUp"
      Visible         =   0   'False
      Begin VB.Menu mnu����������������������� 
         Caption         =   "���������� ������� ������"
      End
   End
End
Attribute VB_Name = "UtilsBatchPriceListEdit"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private Function FindRecordsAndPopulateGrid()

    If ValidateFields Then
        If RefreshList Then
            EnableGrid grdBatchPriceEdit, True
            HighlightRow grdBatchPriceEdit, 0, 4, "", False
            UpdateButtons Me, 3, 0, 1, 1, 0
            Exit Function
        Else
            If Not blnErrors Then DisplayMessageRecordsNotFound
            frmCriteria.Visible = True
            txtYear.SetFocus
        End If
    End If

End Function

Private Function DeleteRecord()
    
    'Local variables
    Dim lngID As Long
    Dim lngRow As Long
    
    '������� �����
    lngID = grdBatchPriceEdit.CellValue(grdBatchPriceEdit.CurRow, "PriceID")
    lngRow = grdBatchPriceEdit.CurRow
    
    '�������� ��������
    If MainDeleteRecord("CommonDB", "Prices", strApplicationName, "ID", lngID, "True") Then
        With grdBatchPriceEdit
            '�������� �������
            .RemoveRow (lngRow)
            If .RowCount > 0 Then
                If lngRow - 1 >= 1 Then
                    .CurRow = lngRow - 1
                Else
                    .CurRow = 1
                End If
                .Redraw = True
                .EnsureVisibleRow .CurRow
                .SetFocus
            Else
                'Redraw
                grdBatchPriceEdit.Redraw = True
                '������
                If MyMsgBox(4, strApplicationName, strStandardMessages(7), 1) Then
                End If
                '�����
                EnableFields txtYear
                '���� focus
                txtYear.SetFocus
                '�������� �� �������
                UpdateButtons Me, 3, 1, 0, 0, 1
            End If
        End With
    Else
        grdBatchPriceEdit.SetFocus
    End If

End Function

Private Function ValidateFields()

    ValidateFields = False
    
    '����
    If Len(txtYear.text) = 0 Then
        If MyMsgBox(4, strApplicationName, strStandardMessages(1), 1) Then
        End If
        txtYear.SetFocus
        Exit Function
    End If
    If Len(txtYear.text) <> 4 Then
        If MyMsgBox(4, strApplicationName, strStandardMessages(2), 1) Then
        End If
        txtYear.SetFocus
        Exit Function
    End If
    If Not IsNumeric(txtYear.text) Then
        If MyMsgBox(4, strApplicationName, strStandardMessages(2), 1) Then
        End If
        txtYear.SetFocus
        Exit Function
    End If
    
    ValidateFields = True
    
End Function

Private Function ValidateGrid()

    ValidateGrid = True
    
End Function

Private Sub cmdButton_Click(index As Integer)

    Select Case index
        Case 0
            FindRecordsAndPopulateGrid
        Case 1
            DeleteRecord
        Case 2
            AbortProcedure False
        Case 3
            AbortProcedure True
    End Select
    
End Sub

Private Function AbortProcedure(blnStatus)

    If grdBatchPriceEdit.TextEditText <> "" Then
        grdBatchPriceEdit.CancelEdit
        Exit Function
    End If
    
    If Not blnStatus Then
        ClearFields grdBatchPriceEdit
        frmCriteria.Visible = True
        txtYear.SetFocus
        UpdateButtons Me, 3, 1, 0, 0, 1
    End If
    
    If blnStatus Then
        Unload Me
    End If

End Function

Private Function RefreshList()

    On Error GoTo ErrTrap

    'SQL
    Dim intIndex As Byte
    Dim strThisQuery As String
    Dim strParameters As String
    Dim strParFields As String
    Dim strThisParameter As String
    Dim strOrder As String
    Dim strLogic As String
    Dim arrQuery() As Variant
    Dim strSQL As String
    
    'Local variables
    Dim lngRow As Long
    
    'Recordsets
    Dim rstRecordset As Recordset

    '������� �����
    intIndex = 0
    lngRow = 0
    frmCriteria.Visible = False
    
    '������
    With grdBatchPriceEdit
        .Clear
        .Redraw = False
    End With
    
    '������ ����������
    strSQL = "SELECT PriceID, PriceCustomerID, PriceDestinationID, PriceFrom, PriceTo, PriceAdultWithTransfer, PriceKidWithTransfer, PriceAdultWithoutTransfer, PriceKidWithoutTransfer, Prices.ShowInList, Description, DestinationDescription " _
    & "FROM ((Prices " _
    & "INNER JOIN Destinations ON Prices.PriceDestinationID = Destinations.DestinationID) " _
    & "INNER JOIN Customers ON Prices.PriceCustomerID = Customers.ID) "

    '����
    strThisParameter = "intCurrentYear Integer"
    strThisQuery = "(Year(PriceFrom) = intCurrentYear"
    strLogic = " AND "
    GoSub UpdateSQLString
    arrQuery(intIndex) = Val(txtYear.text)
    
    '����������� ����
    strThisParameter = "intPreviousYear Integer"
    strThisQuery = "Year(PriceFrom) = intPreviousYear)"
    strLogic = " OR "
    GoSub UpdateSQLString
    arrQuery(intIndex) = Val(txtYear.text) - 1
    
    '�������
    If txtCustomerID.text <> "" Then
        strThisParameter = "intCustomer Integer"
        strThisQuery = "PriceCustomerID = intCustomer"
        strLogic = " AND "
        GoSub UpdateSQLString
        arrQuery(intIndex) = Val(txtCustomerID.text)
    End If
    
    '����������
    strOrder = " ORDER BY Description, DestinationDescription, PriceFrom"
    
    Set TempQuery = CommonDB.CreateQueryDef("")
    
    '�������� �� ��������
    If strThisParameter <> "" Then
        strParameters = "PARAMETERS " & strParameters & "; "
        strParFields = "WHERE " & strParFields
        TempQuery.SQL = strSQL & strOrder
    End If
    
    '��������
    If strThisParameter <> "" Then
        For intIndex = 1 To UBound(arrQuery)
            TempQuery.Parameters(intIndex - 1) = arrQuery(intIndex)
        Next intIndex
    End If
    
    '������ �� recordset
    Set rstRecordset = TempQuery.OpenRecordset()
    
    '�� ��� ��� ��������, ������
    If rstRecordset.RecordCount = 0 Then blnErrors = False: RefreshList = False: Exit Function
    
    '����������� �� ����� �������
    InitializeProgressBar Me, strApplicationName, rstRecordset
    
    '������ �� ������
    With rstRecordset
        grdBatchPriceEdit.AddRow , , , , , , , rstRecordset.RecordCount
        Do Until .EOF
            lngRow = lngRow + 1
            UpdateProgressBar Me
            grdBatchPriceEdit.CellValue(lngRow, "PriceID") = !PriceID
            grdBatchPriceEdit.CellValue(lngRow, "CustomerID") = !PriceCustomerID
            grdBatchPriceEdit.CellValue(lngRow, "DestinationID") = !PriceDestinationID
            grdBatchPriceEdit.CellValue(lngRow, "CustomerDescription") = !CustomerDescription
            grdBatchPriceEdit.CellValue(lngRow, "DestinationDescription") = !DestinationDescription
            grdBatchPriceEdit.CellValue(lngRow, "PriceFrom") = !PriceFrom
            grdBatchPriceEdit.CellValue(lngRow, "PriceTo") = !PriceTo
            grdBatchPriceEdit.CellValue(lngRow, "PriceAdultWithTransfer") = !PriceAdultWithTransfer
            grdBatchPriceEdit.CellValue(lngRow, "PriceKidWithTransfer") = !PriceKidWithTransfer
            grdBatchPriceEdit.CellValue(lngRow, "PriceAdultWithoutTransfer") = !PriceAdultWithoutTransfer
            grdBatchPriceEdit.CellValue(lngRow, "PriceKidWithoutTransfer") = !PriceKidWithoutTransfer
            grdBatchPriceEdit.CellValue(lngRow, "ShowInList") = !ShowInList
            .MoveNext
        Loop
    End With
    
    '������ ����������
    RefreshList = True
    
    '������� ���������
    frmProgress.Visible = False
    
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
    ClearFields grdBatchPriceEdit, frmProgress
    DisplayErrorMessage True, Err.Description
    
End Function

Private Sub cmdIndex_Click(index As Integer)

    'Local variables
    Dim tmpTableData As typTableData
    Dim tmpRecordset As Recordset
    
    Select Case index
        Case 0
            '�������
            Set tmpRecordset = CheckForMatch("CommonDB", "Customers", "CustomerDescription", "String", txtCustomerDescription.text)
            If tmpRecordset.RecordCount > 0 Then
                tmpTableData = DisplayIndex(tmpRecordset, 2, True, 2, 0, 1, "ID", "��������", 0, 40, 1, 0)
                txtCustomerID.text = tmpTableData.strCode
                txtCustomerDescription.text = tmpTableData.strFirstField
            End If
    End Select

End Sub

Private Sub Form_Activate()
                
    If Me.Tag = "True" Then
        Me.Tag = "False"
        AddColumnsToGrid grdBatchPriceEdit, False, 44, GetSetting(strApplicationName, "Layout Strings", "grdBatchPriceEdit"), _
        "05NCIPriceID,05NCICustomerID,05NCIDestinationID,40NLNCustomerDescription,40NLNDestinationDescription,10NRFPriceFrom,10NRFPriceTo,10NRFXPriceAdultWithTransfer,10NRFXPriceKidWithTransfer,10NRFXPriceAdultWithoutTransfer,10NRFXPriceKidWithoutTransfer,05ShowInList", _
        "ID,CustomerID,DestinationID,�������,����������,���,���,�������� �� ��������,������ �� ��������,�������� ����� ��������,������ ����� ��������,�����"
        Me.Refresh
        frmCriteria.Visible = True
        txtYear.SetFocus
    End If
    
    'AddDummyLines grdBatchPriceEdit, "99999", "99999", "99999", "AAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAA", "AAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAA", "A99/99/9999A", "A99/99/9999A", "999999"
    
End Sub

Private Sub Form_KeyDown(KeyCode As Integer, Shift As Integer)

    CheckFunctionKeys KeyCode, Shift
    
    KeyCode = ResetKeyCode(KeyCode, Shift)

End Sub

Private Function CheckFunctionKeys(KeyCode, Shift)

    Dim ShiftDown, AltDown, CtrlDown
    
    ShiftDown = (Shift And vbShiftMask) > 0
    AltDown = (Shift And vbAltMask) > 0
    CtrlDown = (Shift And vbCtrlMask) > 0
    
    Select Case KeyCode
        Case vbKeyF10 And cmdButton(0).Enabled, vbKeyC And CtrlDown And cmdButton(0).Enabled
            cmdButton_Click 0
        Case vbKeyF3 And cmdButton(1).Enabled, vbKeyD And CtrlDown And cmdButton(1).Enabled
            cmdButton_Click 1
        Case vbKeyEscape
            If cmdButton(2).Enabled Then cmdButton_Click 2: Exit Function
            If cmdButton(3).Enabled Then cmdButton_Click 3
        Case vbKeyF12 And CtrlDown
            ToggleInfoPanel Me
    End Select

End Function

Private Sub Form_Load()

    UpdateColors Me, True, grdBatchPriceEdit
    SetUpGrid lstIconList, , grdBatchPriceEdit
    ClearFields txtYear, txtCustomerID, txtCustomerDescription
    EnableFields txtYear, txtCustomerDescription
    UpdateButtons Me, 3, 1, 0, 0, 1

End Sub

Private Sub grdBatchPriceEdit_AfterCommitEdit(ByVal lRow As Long, ByVal lCol As Long)

    Dim lngRow As Long
    
    If Not ValidateGrid Then Exit Sub
    
    lngRow = grdBatchPriceEdit.CurRow
    
    With grdBatchPriceEdit
        .CellValue(lngRow, "PriceID") = MainSaveRecord("CommonDB", "Prices", False, strApplicationName, "ID", _
        .CellValue(lngRow, "PriceID"), _
        .CellValue(lngRow, "CustomerID"), _
        .CellValue(lngRow, "DestinationID"), _
        .CellValue(lngRow, "PriceFrom"), _
        .CellValue(lngRow, "PriceTo"), _
        .CellValue(lngRow, "PriceAdultWithTransfer"), _
        .CellValue(lngRow, "PriceKidWithTransfer"), _
        .CellValue(lngRow, "PriceAdultWithoutTransfer"), _
        .CellValue(lngRow, "PriceKidWithoutTransfer"), _
        .CellValue(lngRow, "ShowInList"), _
        strCurrentUser)
    End With
    
    grdBatchPriceEdit.SetFocus

End Sub

Private Sub grdBatchPriceEdit_BeforeCommitEdit(ByVal lRow As Long, ByVal lCol As Long, eResult As iGrid300_10Tec.EEditResults, ByVal sNewText As String, vNewValue As Variant, ByVal lConvErr As Long)

    '��� - ���
    If lCol = 6 Or lCol = 7 Then
        If Not IsDate(vNewValue) Then
            vNewValue = ""
        End If
    End If
    
    '����
    If lCol = 8 Or lCol = 9 Or lCol = 10 Or lCol = 11 Then
        If (Val(vNewValue) < 0 Or Val(vNewValue) > 999.99) Or vNewValue = "" Or Not IsNumeric(vNewValue) Then
            vNewValue = "0,00"
        End If
    End If

End Sub

Private Sub grdBatchPriceEdit_ColHeaderMouseEnter(ByVal lCol As Long)

    grdBatchPriceEdit.Header.Buttons = True

End Sub

Private Sub grdBatchPriceEdit_ColHeaderMouseLeave(ByVal lCol As Long)

    grdBatchPriceEdit.Header.Buttons = False
    
End Sub

Private Sub grdBatchPriceEdit_CurCellChange(ByVal lRow As Long, ByVal lCol As Long)

    Dim lngCol As Long
    Dim lngRow As Long
    Dim lngColCount As Long
    Dim lngRowCount As Long
    
    lngColCount = grdBatchPriceEdit.colCount
    lngRowCount = grdBatchPriceEdit.RowCount
    
    If grdBatchPriceEdit.RowCount = 0 Then Exit Sub
    
    If grdBatchPriceEdit.CurRow = 0 Then Exit Sub
    
    grdBatchPriceEdit.Redraw = False
    
    For lngCol = 1 To lngColCount
        For lngRow = 1 To lngRowCount
            grdBatchPriceEdit.CellBackColor(lngRow, lngCol) = grdBatchPriceEdit.BackColor
        Next lngRow
    Next lngCol
    
    For lngCol = 1 To lngColCount
        grdBatchPriceEdit.CellBackColor(grdBatchPriceEdit.CurRow, lngCol) = &HC0C0FF
    Next lngCol
    
    grdBatchPriceEdit.Redraw = True
        
End Sub

Private Sub grdBatchPriceEdit_HeaderRightClick(ByVal lCol As Long, ByVal Shift As Integer, ByVal X As Long, ByVal Y As Long)

    PopupMenu mnuHdrPopUp

End Sub

Private Sub grdBatchPriceEdit_RequestEdit(ByVal lRow As Long, ByVal lCol As Long, ByVal iKeyAscii As Integer, bCancel As Boolean, sText As String, lMaxLength As Long, eTextEditOpt As iGrid300_10Tec.ETextEditFlags)

    '������ ���� ��������
    If lCol < 6 Or lCol > 11 Then bCancel = True

End Sub

Private Sub mnu�����������������������_Click()

    SaveSetting strApplicationName, "Layout Strings", "grdBatchPriceEdit", grdBatchPriceEdit.LayoutCol

End Sub

Private Function SaveRecords()

    On Error GoTo ErrTrap
    
    'Local variables
    Dim lngRow As Long
    Dim lngID As Long
    
    '����������� �� ����� �������
    InitializeProgressBar Me, strApplicationName, grdBatchPriceEdit.RowCount
    
    '������
    With grdBatchPriceEdit
        For lngRow = 1 To .RowCount
            '�������
            UpdateProgressBar Me
            '���������
            lngID = MainSaveRecord("CommonDB", "Prices", False, strApplicationName, "ID", .CellValue(lngRow, "ID"), .CellValue(lngRow, "Description"), .CellValue(lngRow, "Profession"), .CellValue(lngRow, "Address"), .CellValue(lngRow, "TaxNo"), .CellValue(lngRow, "TaxOfficeDescription"), .CellValue(lngRow, "VATStateID"), .CellValue(lngRow, "AccountCode"), .CellValue(lngRow, "PersonInCharge"), .CellValue(lngRow, "Phones"), 1, strCurrentUser)
        Next lngRow
    End With
    
    '�������
    frmProgress.Visible = False
    
    '������ �����������
    If MyMsgBox(1, strApplicationName, strStandardMessages(8), 1) Then
    End If
    
    '����������
    cmdButton_Click 2
    
    '������
    Exit Function
    
ErrTrap:
    If MyMsgBox(4, strApplicationName, strStandardMessages(20), 1) Then
    End If
    '�������
    frmProgress.Visible = False
    '����������
    cmdButton_Click 2
    '������
    Exit Function
    
End Function

Private Sub txtCustomerDescription_Change()

    If txtCustomerDescription.text = "" Then
        ClearFields txtCustomerID
    End If

End Sub

Private Sub txtCustomerDescription_KeyDown(KeyCode As Integer, Shift As Integer)

    If KeyCode = vbKeyF2 Then cmdIndex_Click 0

End Sub

Private Sub txtCustomerDescription_Validate(Cancel As Boolean)

    If txtCustomerID.text = "" And txtCustomerDescription.text <> "" Then cmdIndex_Click 0: If txtCustomerID.text = "" Then Cancel = True

End Sub


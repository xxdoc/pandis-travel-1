VERSION 5.00
Object = "{396F7AC0-A0DD-11D3-93EC-00C0DFE7442A}#1.0#0"; "ImageList.ocx"
Object = "{158C2A77-1CCD-44C8-AF42-AA199C5DCC6C}#1.0#0"; "dcButton.ocx"
Object = "{55473EAC-7715-4257-B5EF-6E14EBD6A5DD}#1.0#0"; "ProgressBar.ocx"
Object = "{839D0F5D-B7D7-41B7-A3B4-85D69300B8C1}#98.0#0"; "iGrid300_10Tec.ocx"
Object = "{FFE4AEB4-0D55-4004-ADF2-3C1C84D17A72}#1.0#0"; "userControls.ocx"
Begin VB.Form ChecksIndex 
   Appearance      =   0  'Flat
   BackColor       =   &H00C0E0FF&
   BorderStyle     =   0  'None
   ClientHeight    =   10875
   ClientLeft      =   15
   ClientTop       =   15
   ClientWidth     =   19170
   ControlBox      =   0   'False
   ForeColor       =   &H00800000&
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   10875
   ScaleWidth      =   19170
   ShowInTaskbar   =   0   'False
   Begin VB.Frame frmProgress 
      BackColor       =   &H000080FF&
      BorderStyle     =   0  'None
      Caption         =   "Frame1"
      Height          =   1140
      Left            =   12750
      TabIndex        =   25
      Top             =   7650
      Width           =   4065
      Begin vbalProgBarLib6.vbalProgressBar prgProgressBar 
         Height          =   615
         Left            =   150
         TabIndex        =   26
         TabStop         =   0   'False
         Top             =   375
         Width           =   3765
         _ExtentX        =   6641
         _ExtentY        =   1085
         Picture         =   "ChecksIndex.frx":0000
         ForeColor       =   0
         Appearance      =   0
         BarPicture      =   "ChecksIndex.frx":001C
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
         Caption         =   "Τίτλος"
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
         Height          =   240
         Left            =   150
         TabIndex        =   27
         Top             =   75
         Width           =   3765
      End
   End
   Begin VB.Frame frmContainer 
      BackColor       =   &H00808000&
      BorderStyle     =   0  'None
      Height          =   9615
      Left            =   75
      TabIndex        =   0
      Top             =   75
      Width           =   18990
      Begin VB.Frame frmButtonFrame 
         BackColor       =   &H00FF8080&
         BorderStyle     =   0  'None
         Height          =   690
         Left            =   75
         TabIndex        =   17
         Top             =   8850
         Width           =   6015
         Begin Dacara_dcButton.dcButton cmdButton 
            Height          =   690
            Index           =   0
            Left            =   225
            TabIndex        =   18
            TabStop         =   0   'False
            Top             =   0
            Width           =   1365
            _ExtentX        =   2408
            _ExtentY        =   1217
            ButtonShape     =   3
            ButtonStyle     =   4
            Caption         =   "Συνέχεια"
            BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
               Name            =   "Ubuntu Condensed"
               Size            =   9.75
               Charset         =   161
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   8388736
            PicOpacity      =   0
         End
         Begin Dacara_dcButton.dcButton cmdButton 
            Height          =   690
            Index           =   3
            Left            =   4500
            TabIndex        =   19
            TabStop         =   0   'False
            Top             =   0
            Width           =   1365
            _ExtentX        =   2408
            _ExtentY        =   1217
            BackColor       =   8421631
            ButtonShape     =   3
            ButtonStyle     =   4
            Caption         =   "Κλείσιμο"
            BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
               Name            =   "Ubuntu Condensed"
               Size            =   9.75
               Charset         =   161
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   8388736
            PicOpacity      =   0
         End
         Begin Dacara_dcButton.dcButton cmdButton 
            Height          =   690
            Index           =   2
            Left            =   3075
            TabIndex        =   20
            TabStop         =   0   'False
            Top             =   0
            Width           =   1365
            _ExtentX        =   2408
            _ExtentY        =   1217
            ButtonShape     =   3
            ButtonStyle     =   4
            Caption         =   "Νέα αναζήτηση"
            BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
               Name            =   "Ubuntu Condensed"
               Size            =   9.75
               Charset         =   161
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   8388736
            PicOpacity      =   0
         End
         Begin Dacara_dcButton.dcButton cmdButton 
            Height          =   690
            Index           =   1
            Left            =   1650
            TabIndex        =   30
            TabStop         =   0   'False
            Top             =   0
            Width           =   1365
            _ExtentX        =   2408
            _ExtentY        =   1217
            ButtonShape     =   3
            ButtonStyle     =   4
            Caption         =   "Δημιουργία αρχείου XLS"
            BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
               Name            =   "Ubuntu Condensed"
               Size            =   9.75
               Charset         =   161
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   8388736
            PicOpacity      =   0
         End
      End
      Begin VB.Frame frmInfo 
         Appearance      =   0  'Flat
         BackColor       =   &H8000000D&
         BorderStyle     =   0  'None
         Enabled         =   0   'False
         ForeColor       =   &H80000008&
         Height          =   2190
         Left            =   7575
         TabIndex        =   10
         Top             =   6525
         Width           =   5040
         Begin VB.TextBox Text3 
            Appearance      =   0  'Flat
            BackColor       =   &H00C0E0FF&
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
            TabIndex        =   32
            TabStop         =   0   'False
            Text            =   "CustomersOrSuppliers"
            Top             =   450
            Width           =   3540
         End
         Begin VB.TextBox txtCustomersOrSuppliers 
            Appearance      =   0  'Flat
            BackColor       =   &H00C0E0FF&
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
            TabIndex        =   31
            TabStop         =   0   'False
            Text            =   "2"
            Top             =   450
            Width           =   1305
         End
         Begin VB.TextBox txtPaymentInOrPaymentOut 
            Appearance      =   0  'Flat
            BackColor       =   &H00C0E0FF&
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
            TabIndex        =   29
            TabStop         =   0   'False
            Text            =   "2"
            Top             =   75
            Width           =   1305
         End
         Begin VB.TextBox Text4 
            Appearance      =   0  'Flat
            BackColor       =   &H00C0E0FF&
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
            TabIndex        =   28
            TabStop         =   0   'False
            Text            =   "PaymentInOrPaymentOut"
            Top             =   75
            Width           =   3540
         End
         Begin VB.TextBox txtBankID 
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
            TabIndex        =   16
            TabStop         =   0   'False
            Top             =   1200
            Width           =   1305
         End
         Begin VB.TextBox Text1 
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
            TabIndex        =   15
            TabStop         =   0   'False
            Text            =   "Banks.BankID"
            Top             =   1200
            Width           =   3540
         End
         Begin vbalIml6.vbalImageList lstIconList 
            Left            =   75
            Top             =   1575
            _ExtentX        =   953
            _ExtentY        =   953
            Size            =   2296
            Images          =   "ChecksIndex.frx":0038
            Version         =   131072
            KeyCount        =   2
            Keys            =   "ÿ"
         End
      End
      Begin VB.Frame frmCriteria 
         BackColor       =   &H00FFC0FF&
         BorderStyle     =   0  'None
         Height          =   2640
         Index           =   0
         Left            =   150
         TabIndex        =   6
         Top             =   6075
         Width           =   7365
         Begin UserControls.newDate mskCheckExpireDateFrom 
            Height          =   465
            Left            =   1500
            TabIndex        =   1
            Top             =   825
            Width           =   1455
            _ExtentX        =   2672
            _ExtentY        =   820
            ForeColor       =   0
            Text            =   "01/01/2017"
            BackColor       =   0
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
         Begin UserControls.newDate mskCheckExpireDateTo 
            Height          =   465
            Left            =   3000
            TabIndex        =   2
            Top             =   825
            Width           =   1455
            _ExtentX        =   2672
            _ExtentY        =   820
            ForeColor       =   0
            Text            =   "01/01/2017"
            BackColor       =   0
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
         Begin UserControls.newText txtBankDescription 
            Height          =   465
            Left            =   1500
            TabIndex        =   3
            Top             =   1350
            Width           =   4965
            _ExtentX        =   8758
            _ExtentY        =   820
            ForeColor       =   0
            MaxLength       =   40
            Text            =   "ΑΑΑΑΑΑΑΑΑΑΑΑΑΑΑΑΑΑΑΑΑΑΑΑΑΑΑΑΑΑΑΑΑΑΑΑΑΑΑΑ"
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
            Left            =   6525
            TabIndex        =   14
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
            PicNormal       =   "ChecksIndex.frx":0950
            PicSizeH        =   16
            PicSizeW        =   16
         End
         Begin VB.Shape shpWedge 
            BackColor       =   &H0000FFFF&
            BackStyle       =   1  'Opaque
            BorderStyle     =   0  'Transparent
            FillColor       =   &H00008000&
            Height          =   315
            Index           =   3
            Left            =   2175
            Top             =   1800
            Visible         =   0   'False
            Width           =   465
         End
         Begin VB.Shape shpWedge 
            BackColor       =   &H0000FFFF&
            BackStyle       =   1  'Opaque
            BorderStyle     =   0  'Transparent
            FillColor       =   &H00008000&
            Height          =   315
            Index           =   4
            Left            =   2400
            Top             =   525
            Visible         =   0   'False
            Width           =   465
         End
         Begin VB.Label lblLabel 
            AutoSize        =   -1  'True
            BackColor       =   &H000080FF&
            Caption         =   "Τράπεζα"
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
            TabIndex        =   13
            Top             =   1425
            Width           =   615
         End
         Begin VB.Shape shpWedge 
            BackColor       =   &H0000FFFF&
            BackStyle       =   1  'Opaque
            BorderStyle     =   0  'Transparent
            FillColor       =   &H00008000&
            Height          =   840
            Index           =   2
            Left            =   6900
            Top             =   1125
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
            Left            =   1050
            Top             =   675
            Visible         =   0   'False
            Width           =   465
         End
         Begin VB.Shape shpWedge 
            BackColor       =   &H0000FFFF&
            BackStyle       =   1  'Opaque
            BorderStyle     =   0  'Transparent
            FillColor       =   &H00008000&
            Height          =   840
            Index           =   0
            Left            =   0
            Top             =   675
            Visible         =   0   'False
            Width           =   465
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
            Height          =   390
            Left            =   4275
            TabIndex        =   12
            Top             =   75
            Width           =   2940
         End
         Begin VB.Label Label1 
            BackColor       =   &H00808000&
            Caption         =   "Κριτήρια αναζήτησης"
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
            TabIndex        =   11
            Top             =   0
            Width           =   7440
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
            Index           =   4
            Left            =   0
            TabIndex        =   8
            Top             =   2100
            Width           =   7440
         End
         Begin VB.Label lblLabel 
            BackColor       =   &H000080FF&
            Caption         =   "Λήξη"
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
            TabIndex        =   7
            Top             =   900
            Width           =   615
         End
      End
      Begin iGrid300_10Tec.iGrid grdChecksIndex 
         Height          =   7290
         Left            =   75
         TabIndex        =   4
         TabStop         =   0   'False
         Top             =   1500
         Width           =   18840
         _ExtentX        =   33232
         _ExtentY        =   12859
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
      Begin VB.Label lblCriteria 
         Alignment       =   1  'Right Justify
         BackColor       =   &H00C0FFFF&
         BackStyle       =   0  'Transparent
         Caption         =   "Κριτήρια αναζήτησης"
         BeginProperty Font 
            Name            =   "Ubuntu Condensed"
            Size            =   11.25
            Charset         =   161
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H000080FF&
         Height          =   315
         Left            =   3975
         TabIndex        =   24
         Top             =   1125
         Width           =   14940
      End
      Begin VB.Label lblSelectedGridLines 
         Alignment       =   1  'Right Justify
         BackColor       =   &H00C0C0FF&
         BackStyle       =   0  'Transparent
         Caption         =   "Επιλεγμένες 0 εγγραφές"
         BeginProperty Font 
            Name            =   "Ubuntu Condensed"
            Size            =   11.25
            Charset         =   161
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00C0FFFF&
         Height          =   315
         Left            =   3975
         TabIndex        =   23
         Top             =   525
         Width           =   14940
      End
      Begin VB.Label lblSelectedGridTotals 
         Alignment       =   1  'Right Justify
         BackColor       =   &H000080FF&
         BackStyle       =   0  'Transparent
         Caption         =   "Σύνολα πάνε εδώ"
         BeginProperty Font 
            Name            =   "Ubuntu Condensed"
            Size            =   11.25
            Charset         =   161
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H0080C0FF&
         Height          =   315
         Left            =   3975
         TabIndex        =   22
         Top             =   825
         Width           =   14940
      End
      Begin VB.Label lblRecordCount 
         BackColor       =   &H000080FF&
         BackStyle       =   0  'Transparent
         Caption         =   "Βρέθηκαν 99.999 εγγραφές"
         BeginProperty Font 
            Name            =   "Ubuntu Condensed"
            Size            =   12
            Charset         =   161
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00FFFF00&
         Height          =   315
         Left            =   75
         TabIndex        =   21
         Top             =   1125
         Width           =   2565
      End
      Begin VB.Label lblTitle 
         AutoSize        =   -1  'True
         BackColor       =   &H00800000&
         BackStyle       =   0  'Transparent
         Caption         =   "Ημερολόγιο αξιογράφων"
         BeginProperty Font 
            Name            =   "Ubuntu Condensed"
            Size            =   30
            Charset         =   161
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H000080FF&
         Height          =   720
         Left            =   75
         TabIndex        =   5
         Top             =   75
         Width           =   5505
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
      Begin VB.Menu mnuΑποθήκευσηΠλάτουςΣτηλών 
         Caption         =   "Αποθήκευση πλάτους στηλών"
      End
   End
End
Attribute VB_Name = "ChecksIndex"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Dim lngRowCount As Long
Dim blnError As Boolean
Dim blnProcessing As Boolean

'Ζητούμενη περίοδος
Dim curTotalAmount As Currency

Private Function AddGrandTotalsToGrid(lngLastRow As Long)

    With grdChecksIndex
        If .RowCount > 0 Then
            'Προσθέτω τα γενικά σύνολα (τελευταία γραμμή)
            .AddRow
            .AddRow
            .CellValue(lngLastRow, "Description") = "ΓΕΝΙΚΑ ΣΥΝΟΛΑ"
            .CellValue(lngLastRow, "Amount") = curTotalAmount
        End If
    End With

End Function

Private Function FindRecordsAndPopulateGrid()

    If ValidateFields Then
        If RefreshList > 0 Then
            UpdateRecordCount lblRecordCount, lngRowCount
            UpdateCriteriaLabels mskCheckExpireDateFrom.text, mskCheckExpireDateTo.text, txtBankDescription.text
            EnableGrid grdChecksIndex, False
            HighlightRow grdChecksIndex, 1, 1, "", True
            UpdateButtons Me, 3, 0, 1, 1, 0
            Exit Function
        Else
            UpdateButtons Me, 3, 1, 0, 0, 1
            If Not blnError Then
                If blnProcessing Then
                    If MyMsgBox(4, strApplicationName, strStandardMessages(27), 1) Then
                    End If
                Else
                    If MyMsgBox(1, strApplicationName, strStandardMessages(7), 1) Then
                    End If
                End If
            End If
            blnError = False
            blnProcessing = False
            frmCriteria(0).Visible = True
            mskCheckExpireDateFrom.SetFocus
        End If
    End If

End Function

Private Function UpdateCriteriaLabels(InvoiceDateIssueFrom, InvoiceDateIssueTo, FilterDescription)

    Dim strCriteriaA As String

    strCriteriaA = IIf(InvoiceDateIssueFrom = "", "Από [ ΟΛΑ ] ", "Από [ " & InvoiceDateIssueFrom & " ] ")
    strCriteriaA = strCriteriaA & IIf(InvoiceDateIssueTo = "", "Εως [ ΟΛΑ ] ", "Εως [ " & InvoiceDateIssueTo & " ] ")
    strCriteriaA = strCriteriaA & "Τράπεζα [ " & IIf(FilterDescription <> "", FilterDescription, "ΟΛΕΣ") & " ]"
    
    lblCriteria.Caption = strCriteriaA
    
End Function

Private Sub cmdButton_Click(index As Integer)

    Select Case index
        Case 0
            FindRecordsAndPopulateGrid
        Case 1
            ExportToExcel
        Case 2
            AbortProcedure False
        Case 3
            AbortProcedure True
    End Select
   
End Sub

Private Function ExportToExcel()

    On Error GoTo ErrTrap
    
    Dim lngRow As Long
    Dim lngCol As Long
    Dim xlsRowOffsetFromTop As Long
    Dim xlsColCount As Long
    
    Dim oExcel As Object
    Dim oBook As Object
    Dim oSheet As Object

    Set oExcel = CreateObject("Excel.Application")
    Set oBook = oExcel.Workbooks.Add
    Set oSheet = oBook.Worksheets(1)
    
    xlsRowOffsetFromTop = 10
    xlsColCount = 5
    
    With oSheet
    
        SetFontNameAndSize oSheet, "Ubuntu Condensed", 11
        AddCompanyData oSheet, xlsColCount
        AddTitle oSheet, lblTitle.Caption, xlsColCount
        AddCriteria oSheet, lblCriteria.Caption, xlsColCount
        AddHeaders oSheet, grdChecksIndex, xlsColCount, "A", "BankDescription", "B", "PersonDescription", "C", "CheckNo", "D", "ExpireDate", "E", "Amount"
        AdjustColumnWidths oSheet, "A", 60, "B", 60, "C", 25, "D", 25, "E", 25
                
        For lngRow = 1 To grdChecksIndex.RowCount
            .Range("A" & lngRow + xlsRowOffsetFromTop) = grdChecksIndex.CellValue(lngRow, "BankDescription")
            .Range("B" & lngRow + xlsRowOffsetFromTop) = grdChecksIndex.CellValue(lngRow, "PersonDescription")
            .Range("C" & lngRow + xlsRowOffsetFromTop) = grdChecksIndex.CellValue(lngRow, "CheckNo")
            .Range("D" & lngRow + xlsRowOffsetFromTop) = grdChecksIndex.CellValue(lngRow, "ExpireDate")
            .Range("E" & lngRow + xlsRowOffsetFromTop) = grdChecksIndex.CellValue(lngRow, "Amount")
        Next lngRow
        
        AddNumberFormats oSheet, grdChecksIndex, "Floats", 10, "E"
    
    End With
    
    oBook.SaveAs strReportsPathName & " " & UCase(CommonMain.lblCompany.Caption) & " " & lblTitle.Caption & " " & format(Date, "yyyy.mm.dd") & "-" & format(Time, "hh.mm.ss") & ".xlsx"
    oExcel.Quit
    
    grdChecksIndex.SetFocus
    
    MyMsgBox 1, strApplicationName, strStandardMessages(8), 1
    
    Exit Function
    
ErrTrap:
    oBook.Close False
    oExcel.Quit

    grdChecksIndex.SetFocus
    
    If Err.Number = 1004 Then
        MyMsgBox 4, strApplicationName, strStandardMessages(27), 1
    Else
        DisplayErrorMessage True, Err.Description
    End If
    
    Exit Function
    
End Function

Private Function CreateUnicodeFile(strReportTitle, strReportSubTitle1, intReportDetailLines)

    'Εκτυπωτής
    Dim lngRow As Long
    Dim intProcessedDetailLines As Integer
    Dim intPageNo As Integer
    
    'Μετρητές
    Dim curDebitSoFar As Currency
    Dim curCreditSoFar As Currency
    Dim curBalanceSoFar As Currency
    Dim curDebitPeriod As Currency
    Dim curCreditPeriod As Currency
    Dim curBalance As Currency

    'Αρχικές τιμές
    curDebitSoFar = 0
    curCreditSoFar = 0
    curBalanceSoFar = 0
    curDebitPeriod = 0
    curCreditPeriod = 0
    curBalance = 0
    intPageNo = 1
    
    intPageNo = 0
    intProcessedDetailLines = 0
    
    Open strUnicodeFile For Output As #1
    
    GoSub Headers
    
    'Εγγραφές
    With grdChecksIndex
        For lngRow = 1 To grdChecksIndex.RowCount
            
            'Εκτυπώνω τη γραμμή
            Print #1, .CellText(lngRow, "Description"); _
                Tab(55 - Len((format(.CellText(lngRow, "DebitSoFar"), "#,##0.00")))); format(.CellText(lngRow, "DebitSoFar"), "#,##0.00"); _
                Tab(69 - Len((format(.CellText(lngRow, "CreditSoFar"), "#,##0.00")))); format(.CellText(lngRow, "CreditSoFar"), "#,##0.00"); _
                Tab(83 - Len((format(.CellText(lngRow, "BalanceSoFar"), "#,##0.00")))); format(.CellText(lngRow, "BalanceSoFar"), "#,##0.00"); _
                Tab(97 - Len((format(.CellText(lngRow, "DebitPeriod"), "#,##0.00")))); format(.CellText(lngRow, "DebitPeriod"), "#,##0.00"); _
                Tab(111 - Len((format(.CellText(lngRow, "CreditPeriod"), "#,##0.00")))); format(.CellText(lngRow, "CreditPeriod"), "#,##0.00"); _
                Tab(125 - Len((format(.CellText(lngRow, "Balance"), "#,##0.00")))); format(.CellText(lngRow, "Balance"), "#,##0.00")
            
            intProcessedDetailLines = intProcessedDetailLines + 1
            
            'Eject
            If intProcessedDetailLines > Val(intReportDetailLines) Then
                Print #1, ""
                Print #1, "ΣΕ ΜΕΤΑΦΟΡΑ"; _
                    Tab(55 - Len(format(curDebitSoFar, "#,##0.00"))); format(curDebitSoFar, "#,##0.00"); _
                    Tab(69 - Len(format(curCreditSoFar, "#,##0.00"))); format(curCreditSoFar, "#,##0.00"); _
                    Tab(83 - Len(format(curBalanceSoFar, "#,##0.00"))); format(curBalanceSoFar, "#,##0.00"); _
                    Tab(97 - Len(format(curDebitPeriod, "#,##0.00"))); format(curDebitPeriod, "#,##0.00"); _
                    Tab(111 - Len(format(curCreditPeriod, "#,##0.00"))); format(curCreditPeriod, "#,##0.00"); _
                    Tab(125 - Len(format(curBalance, "#,##0.00"))); format(curBalance, "#,##0.00")
                    
                GoSub Headers

                Print #1, "ΑΠΟ ΜΕΤΑΦΟΡΑ"; _
                    Tab(55 - Len(format(curDebitSoFar, "#,##0.00"))); format(curDebitSoFar, "#,##0.00"); _
                    Tab(69 - Len(format(curCreditSoFar, "#,##0.00"))); format(curCreditSoFar, "#,##0.00"); _
                    Tab(83 - Len(format(curBalanceSoFar, "#,##0.00"))); format(curBalanceSoFar, "#,##0.00"); _
                    Tab(97 - Len(format(curDebitPeriod, "#,##0.00"))); format(curDebitPeriod, "#,##0.00"); _
                    Tab(111 - Len(format(curCreditPeriod, "#,##0.00"))); format(curCreditPeriod, "#,##0.00"); _
                    Tab(125 - Len(format(curBalance, "#,##0.00"))); format(curBalance, "#,##0.00")
                Print #1, ""
                intProcessedDetailLines = intProcessedDetailLines + 2
            End If
            
            'Σύνολα
            If .CellText(lngRow, "ID") <> "" Then
                curDebitSoFar = curDebitSoFar + .CellValue(lngRow, "DebitSoFar")
                curCreditSoFar = curCreditSoFar + .CellValue(lngRow, "CreditSoFar")
                curBalanceSoFar = curBalanceSoFar + .CellValue(lngRow, "BalanceSoFar")
                curDebitPeriod = curDebitPeriod + .CellValue(lngRow, "DebitPeriod")
                curCreditPeriod = curCreditPeriod + .CellValue(lngRow, "CreditPeriod")
                curBalance = curBalance + .CellValue(lngRow, "Balance")
            End If
            
        Next lngRow
    End With
    
    Close #1
    
    Exit Function
    
Headers:
    intPageNo = intPageNo + 1
    PrintHeadings 124, intPageNo, strReportTitle, strReportSubTitle1
    PrintColumnHeadings 42, "--------- ΠΡΟΗΓΟΥΜΕΝΗ ΠΕΡΙΟΔΟΣ ---------- ---------- ΖΗΤΟΥΜΕΝΗ ΠΕΡΙΟΔΟΣ -----------"
    PrintColumnHeadings 1, "ΕΠΩΝΥΜΙΑ", 49, "ΧΡΕΩΣΗ       ΠΙΣΤΩΣΗ      ΥΠΟΛΟΙΠΟ        ΧΡΕΩΣΗ       ΠΙΣΤΩΣΗ      ΥΠΟΛΟΙΠΟ"
    Print #1, ""
    intProcessedDetailLines = 7
    
    Return
    
End Function

Private Function ValidateFields()

    'Αρχικές τιμές
    ValidateFields = False
    
    'Σωστό διάστημα
    If IsDate(mskCheckExpireDateFrom.text) And IsDate(mskCheckExpireDateTo.text) Then
        If CDate(mskCheckExpireDateFrom.text) > CDate(mskCheckExpireDateTo.text) Then
            If MyMsgBox(4, strApplicationName, strStandardMessages(10), 1) Then
            End If
            mskCheckExpireDateFrom.SetFocus
            Exit Function
        End If
    End If

    ValidateFields = True
    
End Function

Private Function AbortProcedure(blnStatus)

    If blnProcessing Then blnProcessing = False: Exit Function

    If Not blnStatus Then
        ClearFields lblSelectedGridTotals, lblSelectedGridLines, lblCriteria, lblRecordCount
        ClearFields grdChecksIndex
        frmCriteria(0).Visible = True
        mskCheckExpireDateFrom.SetFocus
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
    Dim curTotalAmount As Currency
    
    'Recordsets
    Dim rstRecordset As Recordset
    
    'Αρχικές τιμές
    intIndex = 0
    lngRow = 0
    lngRowCount = 0
    frmCriteria(0).Visible = False
    
    'Πλέγμα
    With grdChecksIndex
        .Clear
        .Redraw = False
    End With
    
    'Κυρίως διαδικασία
    strSQL = "SELECT " _
        & txtPaymentInOrPaymentOut.text & ".TrnID, CheckNo, CheckExpireDate, Amount, BankDescription, Description " _
        & "FROM ((" & txtPaymentInOrPaymentOut.text & " " _
        & "INNER JOIN Invoices ON " & txtPaymentInOrPaymentOut.text & ".TrnID = Invoices.InvoiceTrnID) " _
        & "INNER JOIN Banks ON " & txtPaymentInOrPaymentOut.text & ".BankID = Banks.BankID) " _
        & "INNER JOIN " & txtCustomersOrSuppliers.text & " ON Invoices.InvoicePersonID = " & txtCustomersOrSuppliers.text & ".ID "
    
    'Λήξη Από
    If mskCheckExpireDateFrom.text <> "" Then
        strThisParameter = "datDateIssueFrom Date"
        strThisQuery = "CheckExpireDate >= datDateIssueFrom"
        strLogic = " AND "
        GoSub UpdateSQLString
        arrQuery(intIndex) = mskCheckExpireDateFrom.text
    End If
        
    'Λήξη Εως
    If mskCheckExpireDateTo.text <> "" Then
        strThisParameter = "datDateIssueTo Date"
        strThisQuery = "CheckExpireDate <= datDateIssueTo"
        strLogic = " AND "
        GoSub UpdateSQLString
        arrQuery(intIndex) = mskCheckExpireDateTo.text
    End If
    
    'Τράπεζα
    If txtBankID.text <> "" Then
        strThisParameter = "lngBankID Long"
        strThisQuery = "PaymentOut.BankID = lngBankID"
        strLogic = " AND "
        GoSub UpdateSQLString
        arrQuery(intIndex) = Val(txtBankID.text)
    End If
    
    'Ταξινόμηση
    strOrder = " ORDER BY BankDescription, CheckExpireDate"
    
    Set TempQuery = CommonDB.CreateQueryDef("")
    
    'Προσθέτω τα κριτήρια
    If strThisParameter <> "" Then
        strParameters = "PARAMETERS " & strParameters & "; "
        strParFields = "WHERE " & strParFields
        strSQL = strParameters & strSQL & strParFields
    End If
    
    TempQuery.SQL = strSQL & strOrder
    
    'Κριτήρια
    If strThisParameter <> "" Then
        For intIndex = 1 To UBound(arrQuery)
            TempQuery.Parameters(intIndex - 1) = arrQuery(intIndex)
        Next intIndex
    End If
    
    'Ανοίγω το recordset
    Set rstRecordset = TempQuery.OpenRecordset()
    
    'Αν δεν έχω εγγραφές, βγαίνω
    If rstRecordset.RecordCount = 0 Then blnError = False: RefreshList = False: Exit Function
    
    'Προετοιμάζω τη μπάρα προόδου
    InitializeProgressBar Me, strApplicationName, rstRecordset
    
    'Προσωρινά
    UpdateButtons Me, 3, 0, 0, 1, 0
    cmdButton(3).Caption = "Διακοπή επεξεργασίας"
    blnProcessing = True
    
    'Γεμίζω το πλέγμα
    With rstRecordset
        grdChecksIndex.AddRow , , , , , , , rstRecordset.RecordCount
        lngRowCount = rstRecordset.RecordCount
        Do Until .EOF
            lngRow = lngRow + 1
            UpdateProgressBar Me
            grdChecksIndex.CellValue(lngRow, "TrnID") = !TrnID
            grdChecksIndex.CellValue(lngRow, "BankDescription") = !BankDescription
            grdChecksIndex.CellValue(lngRow, "PersonDescription") = !Description
            grdChecksIndex.CellValue(lngRow, "CheckNo") = !CheckNo
            grdChecksIndex.CellValue(lngRow, "ExpireDate") = !CheckExpireDate
            grdChecksIndex.CellValue(lngRow, "Amount") = !Amount
            curTotalAmount = curTotalAmount + !Amount
            rstRecordset.MoveNext
            DoEvents
            If Not blnProcessing Then Exit Do
        Loop
        rstRecordset.Close
    End With
    
    'Ακύρωση επεξεργασίας
    If Not blnProcessing Then
        blnProcessing = True
        ClearFields grdChecksIndex
        RefreshList = 0
    Else
        RefreshList = lngRowCount
        blnProcessing = False
    End If
    
    'Σύνολα
    If Not blnProcessing Then
        With grdChecksIndex
            .AddRow , , , , , , , 2
            .CellValue(grdChecksIndex.RowCount, "Amount") = curTotalAmount
        End With
    End If
    
    'Τελικές ενέργειες
    cmdButton(2).Caption = "Νέα αναζήτηση"
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
    blnError = True
    ClearFields grdChecksIndex, frmProgress
    DisplayErrorMessage True, Err.Description

    
End Function

Private Sub cmdIndex_Click(index As Integer)

    Dim tmpTableData As typTableData
    Dim tmpRecordset As Recordset
    
    Select Case index
        Case 0
            'Εγγραφές - F2
            Set tmpRecordset = CheckForMatch("CommonDB", "Banks", "BankDescription", "String", txtBankDescription.text)
            If tmpRecordset.RecordCount > 0 Then
                tmpTableData = DisplayIndex(tmpRecordset, 2, True, 2, 0, 1, "ID", "Περιγραφή", 0, 40, 0, 0)
                txtBankID.text = tmpTableData.strCode
                txtBankDescription.text = tmpTableData.strFirstField
            End If
    End Select

End Sub

Private Sub Form_Activate()
                
    If Me.Tag = "True" Then
        Me.Tag = "False"
        AddColumnsToGrid grdChecksIndex, False, 44, GetSetting(strApplicationName, "Layout Strings", "grdChecksIndex"), _
            "40NLNBankDescription,40NLNPersonDescription,10NLNCheckNo,12NCDXExpireDate,10NRFAmount,05NCNTrnID,04NCNSelected", _
            "Τράπεζα,Συναλλασόμενος,Νο επιταγής,Ημερομηνία λήξης,Ποσό,ID,E"
        Me.Refresh
        frmCriteria(0).Visible = True
        mskCheckExpireDateFrom.SetFocus
    End If
            
    'AddDummyLines grdChecksIndex, "AAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAA", "AAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAA", "AAAAAAAAAAAAAAAAAAAA", "99/99/9999", "99.999.999,99", "99999", "E"
    
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
        Case vbKeyEscape
            If cmdButton(2).Enabled Then cmdButton_Click 2: Exit Function
            If cmdButton(3).Enabled Then cmdButton_Click 3
        Case vbKeyF12 And CtrlDown
            ToggleInfoPanel Me
    End Select

End Function

Private Sub Form_Load()

    SetUpGrid lstIconList, grdChecksIndex
    PositionControls Me, True, grdChecksIndex
    ColorizeControls Me, True
    ClearFields lblSelectedGridTotals, lblSelectedGridLines, lblCriteria, lblRecordCount
    ClearFields txtBankID
    ClearFields mskCheckExpireDateFrom, mskCheckExpireDateTo, txtBankDescription
    ClearFields grdChecksIndex
    EnableFields mskCheckExpireDateFrom, mskCheckExpireDateTo, txtBankDescription
    UpdateButtons Me, 3, 1, 0, 0, 1
    
End Sub

Private Sub grdChecksIndex_ColHeaderClick(ByVal lCol As Long, bDoDefault As Boolean, ByVal Shift As Integer, ByVal X As Long, ByVal Y As Long)

    bDoDefault = False

End Sub

Private Sub grdChecksIndex_HeaderRightClick(ByVal lCol As Long, ByVal Shift As Integer, ByVal X As Long, ByVal Y As Long)

    PopupMenu mnuHdrPopUp

End Sub

Private Sub grdChecksIndex_KeyDown(KeyCode As Integer, Shift As Integer, bDoDefault As Boolean)

    If KeyCode = vbKeySpace And grdChecksIndex.RowCount > 0 Then
        grdChecksIndex.CellIcon(grdChecksIndex.CurRow, "Selected") = lstIconList.ItemIndex(SelectRow(grdChecksIndex, 2, KeyCode, grdChecksIndex.CurRow, "TrnID"))
        lblSelectedGridLines.Caption = CountSelected(grdChecksIndex)
        lblSelectedGridTotals.Caption = SumSelectedGridRows(grdChecksIndex, False, "Ποσό", "Amount", "Ποσό", "decimal")
    End If

End Sub

Private Sub mnuΑποθήκευσηΠλάτουςΣτηλών_Click()

    SaveSetting strApplicationName, "Layout Strings", "grdChecksIndex", grdChecksIndex.LayoutCol

End Sub

Private Sub txtBankDescription_Change()

    If txtBankDescription.text = "" Then
        ClearFields txtBankID
    End If

End Sub

Private Sub txtBankDescription_KeyDown(KeyCode As Integer, Shift As Integer)

    If KeyCode = vbKeyF2 Then cmdIndex_Click 0

End Sub

Private Sub txtBankDescription_Validate(Cancel As Boolean)

    If txtBankID = "" And txtBankDescription.text <> "" Then cmdIndex_Click 0

End Sub


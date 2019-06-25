VERSION 5.00
Object = "{396F7AC0-A0DD-11D3-93EC-00C0DFE7442A}#1.0#0"; "ImageList.ocx"
Object = "{158C2A77-1CCD-44C8-AF42-AA199C5DCC6C}#1.0#0"; "dcButton.ocx"
Object = "{55473EAC-7715-4257-B5EF-6E14EBD6A5DD}#1.0#0"; "ProgressBar.ocx"
Object = "{839D0F5D-B7D7-41B7-A3B4-85D69300B8C1}#98.0#0"; "iGrid300_10Tec.ocx"
Object = "{FFE4AEB4-0D55-4004-ADF2-3C1C84D17A72}#1.0#0"; "userControls.ocx"
Begin VB.Form InvoicesInIndex 
   BackColor       =   &H00C0E0FF&
   BorderStyle     =   0  'None
   ClientHeight    =   10875
   ClientLeft      =   -30
   ClientTop       =   -420
   ClientWidth     =   19170
   ControlBox      =   0   'False
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   ScaleHeight     =   10875
   ScaleWidth      =   19170
   ShowInTaskbar   =   0   'False
   Begin VB.Frame frmProgress 
      BackColor       =   &H000080FF&
      BorderStyle     =   0  'None
      Caption         =   "Frame1"
      Height          =   1140
      Left            =   12900
      TabIndex        =   8
      Top             =   7650
      Width           =   4065
      Begin vbalProgBarLib6.vbalProgressBar prgProgressBar 
         Height          =   615
         Left            =   150
         TabIndex        =   9
         TabStop         =   0   'False
         Top             =   375
         Width           =   3765
         _ExtentX        =   6641
         _ExtentY        =   1085
         Picture         =   "InvoicesInIndex.frx":0000
         ForeColor       =   0
         Appearance      =   0
         BarPicture      =   "InvoicesInIndex.frx":001C
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
         ForeColor       =   &H00000000&
         Height          =   240
         Left            =   150
         TabIndex        =   10
         Top             =   75
         Width           =   3765
      End
   End
   Begin VB.Frame frmContainer 
      BackColor       =   &H00808000&
      BorderStyle     =   0  'None
      Height          =   9615
      Left            =   75
      TabIndex        =   11
      Top             =   75
      Width           =   18990
      Begin VB.Frame frmCriteria 
         BackColor       =   &H00FFC0FF&
         BorderStyle     =   0  'None
         Height          =   4740
         Index           =   0
         Left            =   150
         TabIndex        =   19
         Top             =   3975
         Width           =   8040
         Begin UserControls.newDate mskInvoiceDateIssueFrom 
            Height          =   465
            Left            =   2175
            TabIndex        =   0
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
         Begin UserControls.newDate mskInvoiceDateIssueTo 
            Height          =   465
            Left            =   3675
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
         Begin UserControls.newText txtExpenseDescription 
            Height          =   465
            Left            =   2175
            TabIndex        =   4
            Top             =   1875
            Width           =   4965
            _ExtentX        =   8758
            _ExtentY        =   820
            ForeColor       =   0
            MaxLength       =   40
            Text            =   "����������������������������������������"
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
            Left            =   7200
            TabIndex        =   20
            TabStop         =   0   'False
            Top             =   1875
            Width           =   390
            _ExtentX        =   688
            _ExtentY        =   820
            BackColor       =   16777215
            ButtonShape     =   3
            ButtonStyle     =   2
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
            PicNormal       =   "InvoicesInIndex.frx":0038
            PicSizeH        =   16
            PicSizeW        =   16
         End
         Begin UserControls.newText txtInvoiceNo 
            Height          =   465
            Left            =   2175
            TabIndex        =   7
            Top             =   3450
            Width           =   615
            _ExtentX        =   1085
            _ExtentY        =   820
            Alignment       =   2
            ForeColor       =   0
            MaxLength       =   6
            Text            =   "99999"
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
         Begin UserControls.newDate mskInvoiceDateInFrom 
            Height          =   465
            Left            =   2175
            TabIndex        =   2
            Top             =   1350
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
         Begin UserControls.newDate mskInvoiceDateInTo 
            Height          =   465
            Left            =   3675
            TabIndex        =   3
            Top             =   1350
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
         Begin UserControls.newText txtSupplierDescription 
            Height          =   465
            Left            =   2175
            TabIndex        =   5
            Top             =   2400
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
         Begin UserControls.newText txtCodeShortDescriptionA 
            Height          =   465
            Left            =   2175
            TabIndex        =   6
            Top             =   2925
            Width           =   615
            _ExtentX        =   1085
            _ExtentY        =   820
            Alignment       =   2
            ForeColor       =   0
            MaxLength       =   3
            Text            =   "���"
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
            Index           =   2
            Left            =   2850
            TabIndex        =   21
            TabStop         =   0   'False
            Top             =   2925
            Width           =   390
            _ExtentX        =   688
            _ExtentY        =   820
            BackColor       =   16777215
            ButtonShape     =   3
            ButtonStyle     =   2
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
            PicNormal       =   "InvoicesInIndex.frx":05D2
            PicSizeH        =   16
            PicSizeW        =   16
         End
         Begin Dacara_dcButton.dcButton cmdIndex 
            Height          =   465
            Index           =   1
            Left            =   7200
            TabIndex        =   22
            TabStop         =   0   'False
            Top             =   2400
            Width           =   390
            _ExtentX        =   688
            _ExtentY        =   820
            BackColor       =   16777215
            ButtonShape     =   3
            ButtonStyle     =   2
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
            PicNormal       =   "InvoicesInIndex.frx":0B6C
            PicSizeH        =   16
            PicSizeW        =   16
         End
         Begin VB.Shape shpWedge 
            BackColor       =   &H0000FFFF&
            BackStyle       =   1  'Opaque
            BorderStyle     =   0  'Transparent
            FillColor       =   &H00008000&
            Height          =   315
            Index           =   4
            Left            =   2250
            Top             =   3900
            Visible         =   0   'False
            Width           =   465
         End
         Begin VB.Shape shpWedge 
            BackColor       =   &H0000FFFF&
            BackStyle       =   1  'Opaque
            BorderStyle     =   0  'Transparent
            FillColor       =   &H00008000&
            Height          =   315
            Index           =   3
            Left            =   4650
            Top             =   525
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
            Index           =   4
            Left            =   0
            TabIndex        =   32
            Top             =   4200
            Width           =   8040
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
            Left            =   5475
            TabIndex        =   31
            Top             =   75
            Width           =   2415
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
            TabIndex        =   30
            Top             =   75
            Width           =   1665
         End
         Begin VB.Label lblLabel 
            AutoSize        =   -1  'True
            BackColor       =   &H000080FF&
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
            Height          =   315
            Index           =   0
            Left            =   450
            TabIndex        =   29
            Top             =   900
            Width           =   1290
         End
         Begin VB.Shape shpWedge 
            BackColor       =   &H0000FFFF&
            BackStyle       =   1  'Opaque
            BorderStyle     =   0  'Transparent
            FillColor       =   &H00008000&
            Height          =   840
            Index           =   0
            Left            =   0
            Top             =   600
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
            Left            =   1725
            Top             =   1425
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
            Left            =   7575
            Top             =   1950
            Visible         =   0   'False
            Width           =   465
         End
         Begin VB.Label lblLabel 
            BackColor       =   &H000080FF&
            Caption         =   "��������� ������"
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
            Index           =   3
            Left            =   450
            TabIndex        =   28
            Top             =   1950
            Width           =   1290
         End
         Begin VB.Label lblLabel 
            BackColor       =   &H000080FF&
            Caption         =   "�� ������������"
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
            Index           =   6
            Left            =   450
            TabIndex        =   27
            Top             =   3525
            Width           =   1290
         End
         Begin VB.Label lblLabel 
            AutoSize        =   -1  'True
            BackColor       =   &H000080FF&
            Caption         =   "����������"
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
            TabIndex        =   26
            Top             =   1425
            Width           =   1290
         End
         Begin VB.Label lblLabel 
            BackColor       =   &H000080FF&
            Caption         =   "��������������"
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
            Index           =   4
            Left            =   450
            TabIndex        =   25
            Top             =   2475
            Width           =   1290
         End
         Begin VB.Label lblLabel 
            BackColor       =   &H000080FF&
            Caption         =   "�����������"
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
            Index           =   5
            Left            =   450
            TabIndex        =   24
            Top             =   3000
            Width           =   1290
         End
         Begin VB.Label lblCodeDescription 
            BackColor       =   &H00E0E0E0&
            BackStyle       =   0  'Transparent
            Caption         =   "AAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAA"
            BeginProperty Font 
               Name            =   "Ubuntu Condensed"
               Size            =   9.75
               Charset         =   161
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H00800080&
            Height          =   315
            Left            =   3300
            TabIndex        =   23
            Top             =   3000
            Width           =   4215
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
            TabIndex        =   33
            Top             =   0
            Width           =   8040
         End
      End
      Begin VB.Frame frmInfo 
         Appearance      =   0  'Flat
         BackColor       =   &H00000000&
         BorderStyle     =   0  'None
         Enabled         =   0   'False
         ForeColor       =   &H80000008&
         Height          =   2565
         Left            =   8250
         TabIndex        =   34
         Top             =   6150
         Width           =   4515
         Begin VB.TextBox txtExpenseID 
            Appearance      =   0  'Flat
            BackColor       =   &H00FFFFC0&
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
            TabIndex        =   44
            TabStop         =   0   'False
            Text            =   "999"
            Top             =   1575
            Width           =   780
         End
         Begin VB.TextBox Text9 
            Appearance      =   0  'Flat
            BackColor       =   &H00FFFFC0&
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
            TabIndex        =   43
            TabStop         =   0   'False
            Text            =   "InvoiceInExpenseCategoryID"
            Top             =   1575
            Width           =   3540
         End
         Begin VB.TextBox txtCodeID 
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
            TabIndex        =   42
            TabStop         =   0   'False
            Text            =   "999"
            Top             =   1200
            Width           =   780
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
            TabIndex        =   41
            TabStop         =   0   'False
            Text            =   "InvoiceCodeID"
            Top             =   1200
            Width           =   3540
         End
         Begin VB.TextBox Text8 
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
            TabIndex        =   40
            TabStop         =   0   'False
            Text            =   "InvoiceSecondaryRefersTo"
            Top             =   450
            Width           =   3540
         End
         Begin VB.TextBox txtInvoiceSecondaryRefersTo 
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
            TabIndex        =   39
            TabStop         =   0   'False
            Text            =   "999"
            Top             =   450
            Width           =   780
         End
         Begin VB.TextBox Text20 
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
            TabIndex        =   38
            TabStop         =   0   'False
            Text            =   "InvoiceMasterRefersTo"
            Top             =   75
            Width           =   3540
         End
         Begin VB.TextBox txtInvoiceMasterRefersTo 
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
            TabIndex        =   37
            TabStop         =   0   'False
            Text            =   "999"
            Top             =   75
            Width           =   780
         End
         Begin VB.TextBox Text1 
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
            TabIndex        =   36
            TabStop         =   0   'False
            Text            =   "InvoicePersonID"
            Top             =   825
            Width           =   3540
         End
         Begin VB.TextBox txtPersonID 
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
            TabIndex        =   35
            TabStop         =   0   'False
            Text            =   "999"
            Top             =   825
            Width           =   780
         End
         Begin vbalIml6.vbalImageList lstIconList 
            Left            =   75
            Top             =   1950
            _ExtentX        =   953
            _ExtentY        =   953
            Size            =   4592
            Images          =   "InvoicesInIndex.frx":1106
            Version         =   131072
            KeyCount        =   4
            Keys            =   "���"
         End
      End
      Begin VB.Frame frmButtonFrame 
         BackColor       =   &H00FF8080&
         BorderStyle     =   0  'None
         Height          =   690
         Left            =   75
         TabIndex        =   14
         Top             =   8850
         Width           =   7515
         Begin Dacara_dcButton.dcButton cmdButton 
            Height          =   690
            Index           =   0
            Left            =   225
            TabIndex        =   15
            TabStop         =   0   'False
            Top             =   0
            Width           =   1365
            _ExtentX        =   2408
            _ExtentY        =   1217
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
            ForeColor       =   0
            PicOpacity      =   0
         End
         Begin Dacara_dcButton.dcButton cmdButton 
            Height          =   690
            Index           =   4
            Left            =   5925
            TabIndex        =   16
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
            ForeColor       =   0
            PicOpacity      =   0
         End
         Begin Dacara_dcButton.dcButton cmdButton 
            Height          =   690
            Index           =   1
            Left            =   1650
            TabIndex        =   17
            TabStop         =   0   'False
            Top             =   0
            Width           =   1365
            _ExtentX        =   2408
            _ExtentY        =   1217
            ButtonShape     =   3
            ButtonStyle     =   4
            Caption         =   "����������� ��������"
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
            PicOpacity      =   0
         End
         Begin Dacara_dcButton.dcButton cmdButton 
            Height          =   690
            Index           =   3
            Left            =   4500
            TabIndex        =   18
            TabStop         =   0   'False
            Top             =   0
            Width           =   1365
            _ExtentX        =   2408
            _ExtentY        =   1217
            ButtonShape     =   3
            ButtonStyle     =   4
            Caption         =   "��� ���������"
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
            PicOpacity      =   0
         End
         Begin Dacara_dcButton.dcButton cmdButton 
            Height          =   690
            Index           =   2
            Left            =   3075
            TabIndex        =   49
            TabStop         =   0   'False
            Top             =   0
            Width           =   1365
            _ExtentX        =   2408
            _ExtentY        =   1217
            ButtonShape     =   3
            ButtonStyle     =   4
            Caption         =   "���������� ������� xls"
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
            PicOpacity      =   0
         End
      End
      Begin iGrid300_10Tec.iGrid grdInvoicesInIndex 
         Height          =   7290
         Left            =   75
         TabIndex        =   13
         TabStop         =   0   'False
         Top             =   1500
         Width           =   18840
         _ExtentX        =   33232
         _ExtentY        =   12859
         Appearance      =   0
         BackColor       =   14737632
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
      Begin VB.Label lblRecordCount 
         BackColor       =   &H000080FF&
         BackStyle       =   0  'Transparent
         Caption         =   "�������� 99.999 ��������"
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
         TabIndex        =   48
         Top             =   1125
         Width           =   2565
      End
      Begin VB.Label lblSelectedGridTotals 
         Alignment       =   1  'Right Justify
         BackColor       =   &H000080FF&
         BackStyle       =   0  'Transparent
         Caption         =   "������ ���� ���"
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
         TabIndex        =   47
         Top             =   525
         Width           =   14940
      End
      Begin VB.Label lblSelectedGridLines 
         Alignment       =   1  'Right Justify
         BackColor       =   &H00C0C0FF&
         BackStyle       =   0  'Transparent
         Caption         =   "����������� 0 ��������"
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
         TabIndex        =   46
         Top             =   825
         Width           =   14940
      End
      Begin VB.Label lblCriteria 
         Alignment       =   1  'Right Justify
         BackColor       =   &H00C0FFFF&
         BackStyle       =   0  'Transparent
         Caption         =   "�������� ����������"
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
         Left            =   2550
         TabIndex        =   45
         Top             =   1125
         Width           =   16365
      End
      Begin VB.Label lblTitle 
         AutoSize        =   -1  'True
         BackColor       =   &H00800000&
         BackStyle       =   0  'Transparent
         Caption         =   "���������� ������"
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
         TabIndex        =   12
         Top             =   75
         Width           =   4410
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
   Begin iGrid300_10Tec.iGrid grdTotals 
      Height          =   2565
      Left            =   5400
      TabIndex        =   50
      TabStop         =   0   'False
      Top             =   675
      Visible         =   0   'False
      Width           =   5715
      _ExtentX        =   10081
      _ExtentY        =   4524
      Appearance      =   0
      BackColor       =   14737632
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
   Begin VB.Menu mnuHdrPopUp 
      Caption         =   "mnuHdrPopUp"
      Visible         =   0   'False
      Begin VB.Menu mnu����������������������� 
         Caption         =   "���������� ������� ������"
      End
   End
End
Attribute VB_Name = "InvoicesInIndex"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Dim lngRowCount As Long
Dim blnError As Boolean
Dim blnProcessing As Boolean

Private Function AbortProcedure(blnStatus)

    If blnProcessing Then blnProcessing = False: Exit Function

    If Not blnStatus Then
        ClearFields grdInvoicesInIndex, grdTotals, lblRecordCount, lblCriteria, lblSelectedGridLines, lblSelectedGridTotals
        frmCriteria(0).Visible = True
        mskInvoiceDateIssueFrom.SetFocus
        UpdateButtons Me, 4, 1, 0, 0, 0, 1
    End If
    
    If blnStatus Then
        Unload Me
    End If

End Function

Private Function CreateTotals()

    Dim lngRow As Long
    Dim lngCol As Long
    Dim strSQL As String
    Dim lngID As Long
    Dim rsTable As Recordset
    
    Set rsTable = CommonDB.OpenRecordset("TestMe")
    strSQL = "DELETE * FROM TestMe"
    CommonDB.Execute (strSQL)
    
    With grdInvoicesInIndex
        For lngRow = 1 To .RowCount - 2
            lngID = MainSaveRecord("CommonDB", "TestMe", True, strApplicationName, "ID", 0, _
                .CellValue(lngRow, "Supplier"), _
                .CellValue(lngRow, "InvoiceNet"), _
                .CellValue(lngRow, "InvoiceVAT"), _
                .CellValue(lngRow, "InvoiceGross"))
        Next lngRow
    End With

End Function

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
    xlsColCount = 4
    
    With oSheet
        
        SetFontNameAndSize oSheet, "Ubuntu Condensed", 11
        AddCompanyData oSheet, xlsColCount
        AddTitle oSheet, lblTitle.Caption, xlsColCount
        AddCriteria oSheet, lblCriteria.Caption, xlsColCount
        AddHeaders oSheet, grdTotals, xlsColCount, _
            "A", "Description", _
            "B", "Net", _
            "C", "VAT", _
            "D", "Gross"
        AdjustColumnWidths oSheet, "A", 40, "B", 15, "C", 15, "D", 15
                
        InitializeProgressBar Me, strApplicationName, grdTotals.RowCount
        
        '���������
        UpdateButtons Me, 4, 0, 0, 0, 1, 0
        cmdButton(3).Caption = "������� ������������"
        blnProcessing = True
                
        For lngRow = 1 To grdTotals.RowCount
            UpdateProgressBar Me
            .Range("A" & lngRow + xlsRowOffsetFromTop) = grdTotals.CellValue(lngRow, "Supplier")
            .Range("B" & lngRow + xlsRowOffsetFromTop) = grdTotals.CellValue(lngRow, "TotalNet")
            .Range("C" & lngRow + xlsRowOffsetFromTop) = grdTotals.CellValue(lngRow, "TotalVAT")
            .Range("D" & lngRow + xlsRowOffsetFromTop) = grdTotals.CellValue(lngRow, "TotalGross")
            DoEvents
            If Not blnProcessing Then Exit For
        Next lngRow
        
        AddNumberFormats oSheet, grdTotals, "Floats", 10, "B", "C", "D"
        
    End With
    
    If blnProcessing Then
        blnError = False
        GoSub DoFinals
        oBook.SaveAs strReportsPathName & format(Date, "yyyy.mm.dd") & "-" & format(Time, "hh.mm.ss") & ".xlsx"
        oExcel.Quit
        MyMsgBox 1, strApplicationName, strStandardMessages(8), 1
    Else
        GoSub DoFinals
        MyMsgBox 1, strApplicationName, strStandardMessages(27), 1
    End If
    
    Exit Function
    
DoFinals:
    blnProcessing = False
    ClearFields frmProgress
    UpdateButtons Me, 4, 0, 1, 1, 1, 0
    cmdButton(3).Caption = "��� ���������"
    grdInvoicesInIndex.SetFocus
    
    Return

ErrTrap:
    oBook.Close False
    oExcel.Quit
    GoSub DoFinals
    If blnError Then
        DisplayErrorMessage True, Err.Description
    End If
    
End Function



Private Function EditRecord()

    If Not grdInvoicesInIndex.Enabled Then Exit Function
        
    Dim rstRecordset As Recordset
    Dim rstExpensesPerVAT As Recordset
    
    Set rstRecordset = InvoicesIn.SeekRecord(grdInvoicesInIndex.CellValue(grdInvoicesInIndex.CurRow, "TrnID"))
                
    If rstRecordset.RecordCount = 0 Then
        If MyMsgBox(4, strApplicationName, strStandardMessages(9), 1) Then
        End If
        Exit Function
    End If
    
    Set rstExpensesPerVAT = InvoicesIn.FindExpensesPerVAT(grdInvoicesInIndex.CellValue(grdInvoicesInIndex.CurRow, "TrnID"))
    
    If rstExpensesPerVAT.RecordCount = 0 Then
        If MyMsgBox(4, strApplicationName, strStandardMessages(9), 1) Then
        End If
        Exit Function
    End If
    
    InvoicesIn.DoPostFoundJobs rstRecordset, rstExpensesPerVAT
    
    If Not InvoicesIn.Visible Then
        InvoicesIn.Show 1, Me
    Else
        Unload Me
    End If
    
End Function

Private Function FindRecordsAndPopulateGrid()

    If ValidateFields Then
        If RefreshList > 0 Then
            UpdateRecordCount lblRecordCount, lngRowCount
            UpdateCriteriaLabels mskInvoiceDateIssueFrom.text, mskInvoiceDateIssueTo.text, mskInvoiceDateInFrom.text, mskInvoiceDateInTo.text, txtExpenseDescription.text, txtSupplierDescription.text, txtCodeShortDescriptionA.text, txtInvoiceNo.text
            EnableGrid grdInvoicesInIndex, False
            HighlightRow grdInvoicesInIndex, 1, 1, "", True
            UpdateButtons Me, 4, 0, 1, 1, 1, 0
            Exit Function
        Else
            UpdateButtons Me, 4, 1, 0, 0, 0, 1
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
            mskInvoiceDateIssueFrom.SetFocus
        End If
    End If

End Function

Private Function UpdateCriteriaLabels(InvoiceDateIssueFrom, InvoiceDateIssueTo, InvoiceDateInFrom, InvoiceDateInTo, ExpenseDescription, SupplierDescription, CodeShortDescriptionA, InvoiceNo)

    Dim strCriteriaA As String

    strCriteriaA = "������ ���" & IIf(InvoiceDateIssueFrom = "", " [ ��� ] ", " [ " & InvoiceDateIssueFrom & " ] ")
    strCriteriaA = strCriteriaA & "���" & IIf(InvoiceDateIssueTo = "", " [ ��� ] ", " [ " & InvoiceDateIssueTo & " ] ")
    
    strCriteriaA = strCriteriaA & "���������� ���" & IIf(InvoiceDateInFrom = "", " [ ��� ] ", " [ " & InvoiceDateInFrom & " ] ")
    strCriteriaA = strCriteriaA & "���" & IIf(InvoiceDateInTo = "", " [ ��� ] ", " [ " & InvoiceDateInTo & " ] ")
    
    strCriteriaA = strCriteriaA & "��������� ������" & IIf(ExpenseDescription = "", " [ ��� ] ", " [ " & ExpenseDescription & " ] ")
    strCriteriaA = strCriteriaA & "��������" & IIf(SupplierDescription = "", " [ ���� ] ", " [ " & SupplierDescription & " ] ")
    
    strCriteriaA = strCriteriaA & "�����������" & IIf(CodeShortDescriptionA = "", " [ ��� ] ", " [ " & CodeShortDescriptionA & " ] ")
    strCriteriaA = strCriteriaA & "�� ������������" & IIf(InvoiceNo = "", " [ ��� ]", " [ " & InvoiceNo & " ]")
    
    lblCriteria.Caption = strCriteriaA
    
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
    Dim strFullInvoice As String
    Dim curTotalNet As Currency
    Dim curTotalVAT As Currency
    Dim curTotalGross As Currency
    Dim lngTotalPersons As Long
    
    'Recordsets
    Dim rstRecordset As Recordset
    Dim rstExpensesPerVAT As Recordset
    
    '������� �����
    intIndex = 0
    lngRow = 0
    lngRowCount = 0
    frmCriteria(0).Visible = False
    
    '������
    With grdInvoicesInIndex
        .Clear
        .Redraw = False
    End With
    
    '������ ����������
    strSQL = "SELECT " _
        & "InvoiceTrnID, InvoiceDateIssue, InvoiceNo, " _
        & "CodeShortDescriptionB, CodeBatch, CodeDescription, CodeSuppliers, " _
        & "Description " _
        & "FROM ((Invoices " _
        & "INNER JOIN Codes ON Invoices.InvoiceCodeID = Codes.CodeID) " _
        & "INNER JOIN Suppliers ON Invoices.InvoicePersonID = Suppliers.ID) "
    
    '�������� ������
    strThisParameter = "strMasterRefersTo String"
    strThisQuery = "InvoiceMasterRefersTo = strMasterRefersTo"
    strLogic = " AND "
    GoSub UpdateSQLString
    arrQuery(intIndex) = txtInvoiceMasterRefersTo.text
    
    '������ ���
    If mskInvoiceDateIssueFrom.text <> "" Then
        strThisParameter = "datDateIssueFrom Date"
        strThisQuery = "InvoiceDateIssue >= datDateIssueFrom"
        strLogic = " AND "
        GoSub UpdateSQLString
        arrQuery(intIndex) = mskInvoiceDateIssueFrom.text
    End If
        
    '������ ���
    If mskInvoiceDateIssueTo.text <> "" Then
        strThisParameter = "datDateIssueTo Date"
        strThisQuery = "InvoiceDateIssue <= datDateIssueTo"
        strLogic = " AND "
        GoSub UpdateSQLString
        arrQuery(intIndex) = mskInvoiceDateIssueTo.text
    End If
    
    '���������� ���
    If mskInvoiceDateInFrom.text <> "" Then
        strThisParameter = "datDateInFrom Date"
        strThisQuery = "InvoiceDateIn >= datDateInFrom"
        strLogic = " AND "
        GoSub UpdateSQLString
        arrQuery(intIndex) = mskInvoiceDateInFrom.text
    End If
        
    '���������� ���
    If mskInvoiceDateInTo.text <> "" Then
        strThisParameter = "datDateInTo Date"
        strThisQuery = "InvoiceDateIn <= datDateInTo"
        strLogic = " AND "
        GoSub UpdateSQLString
        arrQuery(intIndex) = mskInvoiceDateInTo.text
    End If
    
    '��������� ������
    If txtExpenseID.text <> "" Then
        strThisParameter = "lngExpenseID Long"
        strThisQuery = "InvoiceInExpenseCategoryID = lngExpenseID"
        strLogic = " AND "
        GoSub UpdateSQLString
        arrQuery(intIndex) = Val(txtExpenseID.text)
    End If
    
    '��������
    If txtPersonID.text <> "" Then
        strThisParameter = "lngPersonID Long"
        strThisQuery = "InvoicePersonID = lngPersonID"
        strLogic = " AND "
        GoSub UpdateSQLString
        arrQuery(intIndex) = Val(txtPersonID.text)
    End If
    
    '����� ������������
    If txtCodeID.text <> "" Then
        strThisParameter = "lngCodeID Long"
        strThisQuery = "InvoiceCodeID = lngCodeID"
        strLogic = " AND "
        GoSub UpdateSQLString
        arrQuery(intIndex) = Val(txtCodeID.text)
    End If
    
    '�� ������������
    If txtInvoiceNo.text <> "" Then
        strThisParameter = "intInvoiceNo Integer"
        strThisQuery = "InvoiceNo = intInvoiceNo"
        strLogic = " AND "
        GoSub UpdateSQLString
        arrQuery(intIndex) = Val(txtInvoiceNo.text)
    End If
    
    '����������
    strOrder = " ORDER BY Description, InvoiceDateIssue, CodeDescription, InvoiceNo"
    
    Set TempQuery = CommonDB.CreateQueryDef("")
    
    '�������� �� ��������
    If strThisParameter <> "" Then
        strParameters = "PARAMETERS " & strParameters & "; "
        strParFields = "WHERE " & strParFields
        strSQL = strParameters & strSQL & strParFields
    End If
    
    TempQuery.SQL = strSQL & strOrder
    
    '��������
    If strThisParameter <> "" Then
        For intIndex = 1 To UBound(arrQuery)
            TempQuery.Parameters(intIndex - 1) = arrQuery(intIndex)
        Next intIndex
    End If
    
    '������ �� recordset
    Set rstRecordset = TempQuery.OpenRecordset()
    
    '�� ��� ��� ��������, ������
    If rstRecordset.RecordCount = 0 Then blnError = False: RefreshList = False: Exit Function
    
    '����������� �� ����� �������
    InitializeProgressBar Me, strApplicationName, rstRecordset
    
    '���������
    UpdateButtons Me, 4, 0, 0, 0, 1, 0
    cmdButton(3).Caption = "������� ������������"
    blnProcessing = True
    
    '������ �� ������
    With rstRecordset
        grdInvoicesInIndex.AddRow , , , , , , , rstRecordset.RecordCount
        lngRowCount = rstRecordset.RecordCount
        Do Until .EOF
            lngRow = lngRow + 1
            UpdateProgressBar Me
            grdInvoicesInIndex.CellValue(lngRow, "TrnID") = !InvoiceTrnID
            grdInvoicesInIndex.CellValue(lngRow, "DateIssue") = !InvoiceDateIssue
            grdInvoicesInIndex.CellValue(lngRow, "Supplier") = !Description
            grdInvoicesInIndex.CellValue(lngRow, "InvoiceNo") = !InvoiceNo
            grdInvoicesInIndex.CellValue(lngRow, "CodeDescription") = !CodeDescription
            grdInvoicesInIndex.CellValue(lngRow, "CodeSuppliers") = !CodeSuppliers
            
            strSQL = "SELECT " _
                & "SUM(ExpensePerVATNetAmount) AS TotalNet, " _
                & "SUM(ExpensePerVATVATAmount) AS TotalVAT, " _
                & "SUM(ExpensePerVATGrossAmount) AS TotalGross " _
                & "FROM ExpensesPerVAT WHERE ExpensePerVATTrnID = " & !InvoiceTrnID
            TempQuery.SQL = strSQL
            Set rstExpensesPerVAT = TempQuery.OpenRecordset()
            
            If rstExpensesPerVAT.RecordCount = 1 Then
                grdInvoicesInIndex.CellValue(lngRow, "InvoiceNet") = IIf(!CodeSuppliers = "+", rstExpensesPerVAT!TotalNet, rstExpensesPerVAT!TotalNet * -1)
                grdInvoicesInIndex.CellValue(lngRow, "InvoiceVAT") = IIf(!CodeSuppliers = "+", rstExpensesPerVAT!TotalVAT, rstExpensesPerVAT!TotalVAT * -1)
                grdInvoicesInIndex.CellValue(lngRow, "InvoiceGross") = IIf(!CodeSuppliers = "+", rstExpensesPerVAT!TotalGross, rstExpensesPerVAT!TotalGross * -1)
                curTotalNet = curTotalNet + grdInvoicesInIndex.CellValue(lngRow, "InvoiceNet")
                curTotalVAT = curTotalVAT + grdInvoicesInIndex.CellValue(lngRow, "InvoiceVAT")
                curTotalGross = curTotalGross + grdInvoicesInIndex.CellValue(lngRow, "InvoiceGross")
            End If
            
            InvertColorForNegativeNumbers grdInvoicesInIndex, lngRow
            
            rstRecordset.MoveNext
            DoEvents
            If Not blnProcessing Then Exit Do
        Loop
        rstRecordset.Close
    End With
    
    '������� ������������
    If Not blnProcessing Then
        blnProcessing = True
        ClearFields grdInvoicesInIndex, grdTotals
        RefreshList = 0
    Else
        RefreshList = lngRowCount
        blnProcessing = False
    End If
    
    '������
    If Not blnProcessing Then
        With grdInvoicesInIndex
            .AddRow , , , , , , , 2
            .CellValue(grdInvoicesInIndex.RowCount, "InvoiceNet") = curTotalNet
            .CellValue(grdInvoicesInIndex.RowCount, "InvoiceVAT") = curTotalVAT
            .CellValue(grdInvoicesInIndex.RowCount, "InvoiceGross") = curTotalGross
        End With
    End If
    
    '������� ���������
    cmdButton(3).Caption = "��� ���������"
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
    ClearFields grdInvoicesInIndex, grdTotals, frmProgress
    DisplayErrorMessage True, Err.Description

End Function

Private Function UpdateGridWithTotals()

    Dim strSQL As String
    Dim lngRow As Long
    Dim rstRecordset As Recordset
    
    strSQL = "SELECT Description, SUM(Net) AS TotalNet, SUM(Vat) AS TotalVat, SUM(Gross) AS TotalGross FROM TestMe GROUP BY Description"
    
    Set TempQuery = CommonDB.CreateQueryDef("")
    
    TempQuery.SQL = strSQL
    
    Set rstRecordset = TempQuery.OpenRecordset()
    
    grdTotals.Clear
    
    With rstRecordset
        While Not .EOF
            grdTotals.AddRow , , , , , , , .RecordCount
            Do Until .EOF
                lngRow = lngRow + 1
                grdTotals.CellValue(lngRow, "Supplier") = !Description
                grdTotals.CellValue(lngRow, "TotalNet") = !TotalNet
                grdTotals.CellValue(lngRow, "TotalVAT") = !TotalVAT
                grdTotals.CellValue(lngRow, "TotalGross") = !TotalGross
                .MoveNext
            Loop
        Wend
    End With
    
    strSQL = "SELECT SUM(Net) AS TotalNet, SUM(Vat) AS TotalVat, SUM(Gross) AS TotalGross FROM TestMe"

    Set TempQuery = CommonDB.CreateQueryDef("")
    
    TempQuery.SQL = strSQL
    
    Set rstRecordset = TempQuery.OpenRecordset()
    
   With rstRecordset
        While Not .EOF
            grdTotals.AddRow , , , , , , , 2
            Do Until .EOF
                grdTotals.CellValue(grdTotals.RowCount, "TotalNet") = !TotalNet
                grdTotals.CellValue(grdTotals.RowCount, "TotalVAT") = !TotalVAT
                grdTotals.CellValue(grdTotals.RowCount, "TotalGross") = !TotalGross
                .MoveNext
            Loop
        Wend
    End With
                
End Function

Private Sub cmdButton_Click(index As Integer)

    Select Case index
        Case 0
            FindRecordsAndPopulateGrid
        Case 1
            EditRecord
        Case 2
            CreateTotals
            UpdateGridWithTotals
            ExportToExcel
        Case 3
            AbortProcedure False
        Case 4
            AbortProcedure True
    End Select
    
End Sub

Private Function ValidateFields()

    'OK
    ValidateFields = False
    
    '����� �������� �������
    If IsDate(mskInvoiceDateIssueFrom.text) And IsDate(mskInvoiceDateIssueTo.text) Then
        If CDate(mskInvoiceDateIssueFrom.text) > CDate(mskInvoiceDateIssueTo.text) Then
            If MyMsgBox(4, strApplicationName, strStandardMessages(10), 1) Then
            End If
            mskInvoiceDateIssueFrom.SetFocus
            Exit Function
        End If
    End If
    
    '����� �������� �����������
    If IsDate(mskInvoiceDateInFrom.text) And IsDate(mskInvoiceDateInTo.text) Then
        If CDate(mskInvoiceDateInFrom.text) > CDate(mskInvoiceDateInTo.text) Then
            If MyMsgBox(4, strApplicationName, strStandardMessages(10), 1) Then
            End If
            mskInvoiceDateInFrom.SetFocus
            Exit Function
        End If
    End If
    
    ValidateFields = True

End Function

Private Sub cmdButton_KeyDown(index As Integer, KeyCode As Integer, Shift As Integer)

    CheckForArrows (KeyCode)

End Sub

Private Sub cmdIndex_Click(index As Integer)

    Dim tmpTableData As typTableData
    Dim tmpRecordset As Recordset
    
    Select Case index
        Case 0
            '��������� ������ - F2
            Set tmpRecordset = CheckForMatch("CommonDB", "ExpensesCategories", "ExpenseCategoryDescription", "String", txtExpenseDescription.text)
            If tmpRecordset.RecordCount > 0 Then
                tmpTableData = DisplayIndex(tmpRecordset, 2, True, 2, 0, 1, "ID", "���������", 0, 40, 1, 0)
                txtExpenseID.text = tmpTableData.strCode
                txtExpenseDescription.text = tmpTableData.strFirstField
            End If
        Case 1
            '����������� - F2
            Set tmpRecordset = CheckForMatch("CommonDB", "Suppliers", "Description", "String", txtSupplierDescription.text)
            If tmpRecordset.RecordCount > 0 Then
                tmpTableData = DisplayIndex(tmpRecordset, 2, True, 2, 0, 1, "ID", "���������", 0, 40, 1, 0)
                txtPersonID.text = tmpTableData.strCode
                txtSupplierDescription.text = tmpTableData.strFirstField
            End If
        Case 2
            '����������� - F2
            Set tmpRecordset = CheckForMatch("CommonDB", "Codes", "CodeShortDescriptionA, CodeMasterRefersTo", "String, String", txtCodeShortDescriptionA.text, "1")
            If tmpRecordset.RecordCount > 0 Then
                tmpTableData = DisplayIndex(tmpRecordset, 3, True, 8, 0, 3, 5, 6, 7, 9, 10, 11, "ID", "����. �'", "���������", "�����", "����������", "�����������", "��������� ��", "����������", 0, 6, 40, 6, 0, 0, 0, 0, 1, 1, 0, 1, 1, 1, 1, 1)
                txtCodeID.text = tmpTableData.strCode
                txtCodeShortDescriptionA.text = tmpTableData.strFirstField
                lblCodeDescription.Caption = tmpTableData.strSecondField
            End If
    End Select

End Sub

Private Sub Form_Activate()

    If Me.Tag = "True" Then
        Me.Tag = "False"
        AddColumnsToGrid grdInvoicesInIndex, 44, GetSetting(strApplicationName, "Layout Strings", "grdInvoicesInIndex"), _
            "05�CNTrnID,12NCDXDateIssue,40NLNSupplier,10NCNInvoiceNo,05NCNCodeSuppliers,40NLNCodeDescription,10NRFInvoiceNet,10NRFInvoiceVAT,10NRFInvoiceGross,05NCNSelected", _
            "TrnID,���������� �������,��������,�� ������������,+/-,�����������,������ ����, ���� �.�.�.,�������� ����,E"
        AddColumnsToGrid grdTotals, 44, GetSetting(strApplicationName, "Layout Strings", "grdInvoicesInIndex"), _
            "40NLNSupplier,10NRFTotalNet,10NRFTotalVAT,10NRFTotalGross", _
            "��������,������ ����, ���� �.�.�.,�������� ����"
        Me.Refresh
        frmCriteria(0).Visible = True
        mskInvoiceDateIssueFrom.SetFocus
    End If
            
    'AddDummyLines grdInvoicesInIndex, "99999", "A99/99/9999A", "����������������������������������������", "AAAAAA", "+", "AAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAA", "-9.999.999,99", "-9.999.999,99", "-9.999.999,99"

End Sub

    
Private Sub Form_KeyDown(KeyCode As Integer, Shift As Integer)

    CheckFunctionKeys KeyCode, Shift
    
    KeyCode = ResetKeyCode(KeyCode, Shift)
    
End Sub

Private Function CheckFunctionKeys(KeyCode, Shift)
    
    Dim CtrlDown
    
    CtrlDown = Shift + vbCtrlMask
    
    Select Case KeyCode
        Case vbKeyF10 And cmdButton(0).Enabled, vbKeyC And CtrlDown = 4 And cmdButton(0).Enabled
            cmdButton_Click 0
        Case vbKeyE And CtrlDown = 4 And cmdButton(1).Enabled
            cmdButton_Click 1
        Case vbKeyEscape
            If cmdButton(3).Enabled Then cmdButton_Click 3: Exit Function
            If cmdButton(4).Enabled Then cmdButton_Click 4
        Case vbKeyF12 And CtrlDown = 4
            ToggleInfoPanel Me
    End Select
    
End Function

Private Sub Form_Load()

    SetUpGrid lstIconList, grdInvoicesInIndex
    PositionControls Me, True, grdInvoicesInIndex
    ColorizeControls Me, True
    ClearFields lblRecordCount, lblCriteria, lblSelectedGridLines, lblSelectedGridTotals
    ClearFields txtExpenseID, txtPersonID, txtCodeID
    ClearFields mskInvoiceDateIssueFrom, mskInvoiceDateIssueTo, mskInvoiceDateInFrom, mskInvoiceDateInTo, txtExpenseDescription, txtSupplierDescription, txtCodeShortDescriptionA, txtInvoiceNo
    EnableFields mskInvoiceDateIssueFrom, mskInvoiceDateIssueTo, mskInvoiceDateInFrom, mskInvoiceDateInTo, txtExpenseDescription, txtSupplierDescription, txtCodeShortDescriptionA, txtInvoiceNo
    EnableFields cmdIndex(0), cmdIndex(1), cmdIndex(2)
    UpdateButtons Me, 4, 1, 0, 0, 0, 1

End Sub

Private Sub grdInvoicesInIndex_ColHeaderMouseEnter(ByVal lCol As Long)

    grdInvoicesInIndex.Header.Buttons = True

End Sub

Private Sub grdInvoicesInIndex_ColHeaderMouseLeave(ByVal lCol As Long)

    grdInvoicesInIndex.Header.Buttons = False
    
End Sub

Private Sub grdInvoicesInIndex_CurCellChange(ByVal lRow As Long, ByVal lCol As Long)

    cmdButton(1).Enabled = ChangeEditButtonStatus(grdInvoicesInIndex, Me.Tag, lRow, 1)

End Sub

Private Sub grdInvoicesInIndex_DblClick(ByVal lRow As Long, ByVal lCol As Long, bRequestEdit As Boolean)

    If cmdButton(1).Enabled Then cmdButton_Click 1
    
End Sub

Private Sub grdInvoicesInIndex_HeaderRightClick(ByVal lCol As Long, ByVal Shift As Integer, ByVal X As Long, ByVal Y As Long)

    PopupMenu mnuHdrPopUp

End Sub

Private Sub grdInvoicesInIndex_KeyDown(KeyCode As Integer, Shift As Integer, bDoDefault As Boolean)

    If KeyCode = vbKeySpace And grdInvoicesInIndex.RowCount > 0 Then
        grdInvoicesInIndex.CellIcon(grdInvoicesInIndex.CurRow, "Selected") = lstIconList.ItemIndex(SelectRow(grdInvoicesInIndex, 4, KeyCode, grdInvoicesInIndex.CurRow, "TrnID"))
        lblSelectedGridLines.Caption = CountSelected(grdInvoicesInIndex)
        lblSelectedGridTotals.Caption = SumSelectedGridRows(grdInvoicesInIndex, False, "", "InvoiceNet", "decimal", "InvoiceVAT", "decimal", "InvoiceGross", "decimal")
    End If

End Sub

Private Sub grdInvoicesInIndex_KeyPress(KeyAscii As Integer)

    If KeyAscii = vbKeyReturn And cmdButton(1).Enabled Then cmdButton_Click 1
    
End Sub

Private Sub mnu�����������������������_Click()

    SaveSetting strApplicationName, "Layout Strings", "grdInvoicesInIndex", grdInvoicesInIndex.LayoutCol

End Sub

Private Sub txtCodeShortDescriptionA_Change()

    If txtCodeShortDescriptionA.text = "" Then ClearFields txtCodeID, txtCodeShortDescriptionA, lblCodeDescription

End Sub

Private Sub txtCodeShortDescriptionA_KeyDown(KeyCode As Integer, Shift As Integer)

    If KeyCode = vbKeyF2 Then cmdIndex_Click 2
    
End Sub

Private Sub txtCodeShortDescriptionA_Validate(Cancel As Boolean)

    If txtCodeID.text = "" And txtCodeShortDescriptionA.text <> "" Then cmdIndex_Click 2

End Sub

Private Sub txtExpenseDescription_Change()
    
    If txtExpenseDescription.text = "" Then ClearFields txtExpenseID
    
End Sub

Private Sub txtExpenseDescription_KeyDown(KeyCode As Integer, Shift As Integer)

    If KeyCode = vbKeyF2 Then cmdIndex_Click 0

End Sub

Private Sub txtExpenseDescription_Validate(Cancel As Boolean)

    If txtExpenseID.text = "" And txtExpenseDescription.text <> "" Then cmdIndex_Click 0

End Sub

Private Sub txtSupplierDescription_Change()

    If txtSupplierDescription.text = "" Then ClearFields txtPersonID
    
End Sub

Private Sub txtSupplierDescription_KeyDown(KeyCode As Integer, Shift As Integer)

    If KeyCode = vbKeyF2 Then cmdIndex_Click 1
    
End Sub

Private Sub txtSupplierDescription_Validate(Cancel As Boolean)

    If txtPersonID.text = "" And txtSupplierDescription.text <> "" Then cmdIndex_Click 1

End Sub


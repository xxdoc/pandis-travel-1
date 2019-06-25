VERSION 5.00
Object = "{396F7AC0-A0DD-11D3-93EC-00C0DFE7442A}#1.0#0"; "ImageList.ocx"
Object = "{158C2A77-1CCD-44C8-AF42-AA199C5DCC6C}#1.0#0"; "dcButton.ocx"
Object = "{55473EAC-7715-4257-B5EF-6E14EBD6A5DD}#1.0#0"; "ProgressBar.ocx"
Object = "{839D0F5D-B7D7-41B7-A3B4-85D69300B8C1}#98.0#0"; "iGrid300_10Tec.ocx"
Object = "{FFE4AEB4-0D55-4004-ADF2-3C1C84D17A72}#1.0#0"; "userControls.ocx"
Begin VB.Form Transfers 
   Appearance      =   0  'Flat
   BackColor       =   &H00E0E0E0&
   BorderStyle     =   0  'None
   ClientHeight    =   10875
   ClientLeft      =   -30
   ClientTop       =   15
   ClientWidth     =   19170
   ControlBox      =   0   'False
   FillColor       =   &H00800000&
   FillStyle       =   0  'Solid
   BeginProperty Font 
      Name            =   "Ubuntu Condensed"
      Size            =   9.75
      Charset         =   161
      Weight          =   400
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
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
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   161
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   1140
      Left            =   7800
      TabIndex        =   49
      Top             =   3675
      Width           =   4065
      Begin vbalProgBarLib6.vbalProgressBar prgProgressBar 
         Height          =   615
         Left            =   150
         TabIndex        =   50
         TabStop         =   0   'False
         Top             =   375
         Width           =   3765
         _ExtentX        =   6641
         _ExtentY        =   1085
         Picture         =   "Transfers.frx":0000
         ForeColor       =   0
         Appearance      =   0
         BarPicture      =   "Transfers.frx":001C
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
         ForeColor       =   &H00000000&
         Height          =   240
         Left            =   150
         TabIndex        =   51
         Top             =   75
         Width           =   3765
      End
   End
   Begin VB.Frame frmContainer 
      Appearance      =   0  'Flat
      BackColor       =   &H00808080&
      BorderStyle     =   0  'None
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   161
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H80000008&
      Height          =   10440
      Left            =   75
      TabIndex        =   10
      Top             =   75
      Width           =   18990
      Begin VB.Frame frmCriteria 
         BackColor       =   &H00FFC0FF&
         BorderStyle     =   0  'None
         Height          =   2190
         Index           =   0
         Left            =   7800
         TabIndex        =   66
         Top             =   4800
         Width           =   7290
         Begin UserControls.newText txtDriverDescriptionForRoutes 
            Height          =   465
            Left            =   1425
            TabIndex        =   67
            TabStop         =   0   'False
            Top             =   825
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
            Left            =   6450
            TabIndex        =   68
            TabStop         =   0   'False
            Top             =   825
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
            PicNormal       =   "Transfers.frx":0038
            PicSizeH        =   16
            PicSizeW        =   16
         End
         Begin Dacara_dcButton.dcButton cmdButton 
            Height          =   465
            Index           =   10
            Left            =   1575
            TabIndex        =   74
            TabStop         =   0   'False
            Top             =   1650
            Width           =   2115
            _ExtentX        =   3731
            _ExtentY        =   820
            BackColor       =   12640511
            ButtonShape     =   3
            ButtonStyle     =   4
            Caption         =   "Ακυρο"
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
            Height          =   465
            Index           =   11
            Left            =   3750
            TabIndex        =   75
            TabStop         =   0   'False
            Top             =   1650
            Width           =   2115
            _ExtentX        =   3731
            _ExtentY        =   820
            BackColor       =   12640511
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
            PicOpacity      =   0
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
            Height          =   615
            Index           =   4
            Left            =   0
            TabIndex        =   73
            Top             =   1575
            Width           =   7440
         End
         Begin VB.Label Label1 
            BackColor       =   &H00808000&
            Caption         =   "Απόδοση διαδρομών σε οδηγό"
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
            TabIndex        =   71
            Top             =   75
            Width           =   3615
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
            Left            =   4200
            TabIndex        =   70
            Top             =   75
            Width           =   2940
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
         Begin VB.Shape shpWedge 
            BackColor       =   &H0000FFFF&
            BackStyle       =   1  'Opaque
            BorderStyle     =   0  'Transparent
            FillColor       =   &H00008000&
            Height          =   840
            Index           =   1
            Left            =   975
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
            Index           =   2
            Left            =   6825
            Top             =   750
            Visible         =   0   'False
            Width           =   465
         End
         Begin VB.Label lblCriteriaLabel 
            AutoSize        =   -1  'True
            BackColor       =   &H000080FF&
            Caption         =   "Οδηγός"
            ForeColor       =   &H00000000&
            Height          =   315
            Index           =   11
            Left            =   450
            TabIndex        =   69
            Top             =   900
            Width           =   540
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
         Begin VB.Shape shpWedge 
            BackColor       =   &H0000FFFF&
            BackStyle       =   1  'Opaque
            BorderStyle     =   0  'Transparent
            FillColor       =   &H00008000&
            Height          =   315
            Index           =   3
            Left            =   2475
            Top             =   1275
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
            TabIndex        =   72
            Top             =   0
            Width           =   7440
         End
      End
      Begin VB.CheckBox chkAllVisibleTransfers 
         Appearance      =   0  'Flat
         BackColor       =   &H00000000&
         Caption         =   "Επιλογή όλων"
         ForeColor       =   &H00FFFFFF&
         Height          =   315
         Left            =   7725
         TabIndex        =   65
         TabStop         =   0   'False
         Top             =   600
         Width           =   2340
      End
      Begin VB.PictureBox Seperator 
         BackColor       =   &H80000003&
         BorderStyle     =   0  'None
         Height          =   50
         Left            =   3150
         MousePointer    =   7  'Size N S
         ScaleHeight     =   45
         ScaleWidth      =   5865
         TabIndex        =   63
         TabStop         =   0   'False
         Top             =   7275
         Width           =   5865
      End
      Begin VB.Frame frmSummaries 
         BackColor       =   &H00000000&
         BorderStyle     =   0  'None
         Height          =   2115
         Left            =   75
         TabIndex        =   52
         Top             =   7500
         Width           =   18840
         Begin VB.CheckBox chkAllDrivers 
            Appearance      =   0  'Flat
            BackColor       =   &H80000005&
            Caption         =   "Επιλογή όλων"
            ForeColor       =   &H80000008&
            Height          =   315
            Left            =   14400
            TabIndex        =   56
            TabStop         =   0   'False
            Top             =   0
            Width           =   2340
         End
         Begin VB.CheckBox chkAllRoutes 
            Appearance      =   0  'Flat
            BackColor       =   &H80000005&
            Caption         =   "Επιλογή όλων"
            ForeColor       =   &H80000008&
            Height          =   315
            Left            =   9600
            TabIndex        =   55
            TabStop         =   0   'False
            Top             =   0
            Width           =   2340
         End
         Begin VB.CheckBox chkAllCustomers 
            Appearance      =   0  'Flat
            BackColor       =   &H80000005&
            Caption         =   "Επιλογή όλων"
            ForeColor       =   &H80000008&
            Height          =   315
            Left            =   4800
            TabIndex        =   54
            TabStop         =   0   'False
            Top             =   0
            Width           =   2340
         End
         Begin VB.CheckBox chkAllDestinations 
            Appearance      =   0  'Flat
            BackColor       =   &H80000005&
            Caption         =   "Επιλογή όλων"
            ForeColor       =   &H80000008&
            Height          =   315
            Left            =   0
            TabIndex        =   53
            TabStop         =   0   'False
            Top             =   0
            Width           =   2340
         End
         Begin iGrid300_10Tec.iGrid grdSummaryPerRoute 
            Height          =   1720
            Left            =   9600
            TabIndex        =   57
            TabStop         =   0   'False
            Top             =   375
            Width           =   4740
            _ExtentX        =   8361
            _ExtentY        =   3043
            Appearance      =   0
            ForeColor       =   -2147483631
         End
         Begin iGrid300_10Tec.iGrid grdSummaryPerDriver 
            Height          =   1720
            Left            =   14400
            TabIndex        =   58
            TabStop         =   0   'False
            Top             =   375
            Width           =   4440
            _ExtentX        =   7832
            _ExtentY        =   3043
            Appearance      =   0
            ForeColor       =   -2147483631
         End
         Begin iGrid300_10Tec.iGrid grdSummaryPerCustomer 
            Height          =   1720
            Left            =   4800
            TabIndex        =   59
            TabStop         =   0   'False
            Top             =   375
            Width           =   4740
            _ExtentX        =   8361
            _ExtentY        =   3043
            Appearance      =   0
            ForeColor       =   -2147483631
         End
         Begin iGrid300_10Tec.iGrid grdSummaryPerDestination 
            Height          =   1720
            Left            =   0
            TabIndex        =   60
            TabStop         =   0   'False
            Top             =   375
            Width           =   4740
            _ExtentX        =   8361
            _ExtentY        =   3043
            Appearance      =   0
            ForeColor       =   -2147483631
         End
      End
      Begin VB.Frame frmInfo 
         Appearance      =   0  'Flat
         BackColor       =   &H8000000D&
         BorderStyle     =   0  'None
         Enabled         =   0   'False
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   161
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H80000008&
         Height          =   3690
         Left            =   14325
         TabIndex        =   33
         Top             =   1050
         Width           =   4515
         Begin VB.TextBox Text5 
            Appearance      =   0  'Flat
            BackColor       =   &H0000FF00&
            BorderStyle     =   0  'None
            ForeColor       =   &H00000000&
            Height          =   315
            Left            =   75
            TabIndex        =   77
            TabStop         =   0   'False
            Text            =   "DriverIDForRoutes"
            Top             =   2700
            Width           =   3540
         End
         Begin VB.TextBox txtDriverIDForRoutes 
            Appearance      =   0  'Flat
            BackColor       =   &H0000FF00&
            BorderStyle     =   0  'None
            ForeColor       =   &H00000000&
            Height          =   315
            Left            =   3675
            TabIndex        =   76
            TabStop         =   0   'False
            Top             =   2700
            Width           =   780
         End
         Begin VB.TextBox txtDriverID 
            Appearance      =   0  'Flat
            BackColor       =   &H00C0C0FF&
            BorderStyle     =   0  'None
            ForeColor       =   &H00000000&
            Height          =   315
            Left            =   3675
            TabIndex        =   48
            TabStop         =   0   'False
            Top             =   1950
            Width           =   780
         End
         Begin VB.TextBox Text9 
            Appearance      =   0  'Flat
            BackColor       =   &H00C0C0FF&
            BorderStyle     =   0  'None
            ForeColor       =   &H00000000&
            Height          =   315
            Left            =   75
            TabIndex        =   47
            TabStop         =   0   'False
            Text            =   "Transfers.DriverID"
            Top             =   1950
            Width           =   3540
         End
         Begin VB.TextBox Text4 
            Appearance      =   0  'Flat
            BackColor       =   &H00C0C0FF&
            BorderStyle     =   0  'None
            ForeColor       =   &H00000000&
            Height          =   315
            Left            =   75
            TabIndex        =   45
            TabStop         =   0   'False
            Text            =   "Transfers.TransferID"
            Top             =   75
            Width           =   3540
         End
         Begin VB.TextBox Text3 
            Appearance      =   0  'Flat
            BackColor       =   &H00C0C0FF&
            BorderStyle     =   0  'None
            ForeColor       =   &H00000000&
            Height          =   315
            Left            =   75
            TabIndex        =   44
            TabStop         =   0   'False
            Text            =   "Transfers.TransferPickupPointID"
            Top             =   1575
            Width           =   3540
         End
         Begin VB.TextBox Text1 
            Appearance      =   0  'Flat
            BackColor       =   &H00C0C0FF&
            BorderStyle     =   0  'None
            ForeColor       =   &H00000000&
            Height          =   315
            Left            =   75
            TabIndex        =   43
            TabStop         =   0   'False
            Text            =   "Transfers.TransferCustomerID"
            Top             =   825
            Width           =   3540
         End
         Begin VB.TextBox txtTransferID 
            Appearance      =   0  'Flat
            BackColor       =   &H00C0C0FF&
            BorderStyle     =   0  'None
            ForeColor       =   &H00000000&
            Height          =   315
            Left            =   3675
            TabIndex        =   42
            TabStop         =   0   'False
            Top             =   75
            Width           =   780
         End
         Begin VB.TextBox txtCustomerID 
            Appearance      =   0  'Flat
            BackColor       =   &H00C0C0FF&
            BorderStyle     =   0  'None
            ForeColor       =   &H00000000&
            Height          =   315
            Left            =   3675
            TabIndex        =   41
            TabStop         =   0   'False
            Top             =   825
            Width           =   780
         End
         Begin VB.TextBox txtPickupPointID 
            Appearance      =   0  'Flat
            BackColor       =   &H00C0C0FF&
            BorderStyle     =   0  'None
            ForeColor       =   &H00000000&
            Height          =   315
            Left            =   3675
            TabIndex        =   40
            TabStop         =   0   'False
            Top             =   1575
            Width           =   780
         End
         Begin VB.TextBox Text6 
            Appearance      =   0  'Flat
            BackColor       =   &H00C0C0FF&
            BorderStyle     =   0  'None
            ForeColor       =   &H00000000&
            Height          =   315
            Left            =   75
            TabIndex        =   39
            TabStop         =   0   'False
            Text            =   "Transfers.TransferRouteID"
            Top             =   1200
            Width           =   3540
         End
         Begin VB.TextBox txtRouteID 
            Appearance      =   0  'Flat
            BackColor       =   &H00C0C0FF&
            BorderStyle     =   0  'None
            ForeColor       =   &H00000000&
            Height          =   315
            Left            =   3675
            TabIndex        =   38
            TabStop         =   0   'False
            Top             =   1200
            Width           =   780
         End
         Begin VB.TextBox Text7 
            Appearance      =   0  'Flat
            BackColor       =   &H00C0C0FF&
            BorderStyle     =   0  'None
            ForeColor       =   &H00000000&
            Height          =   315
            Left            =   75
            TabIndex        =   37
            TabStop         =   0   'False
            Text            =   "Transfers.TransferDestinationID"
            Top             =   450
            Width           =   3540
         End
         Begin VB.TextBox txtDestinationID 
            Appearance      =   0  'Flat
            BackColor       =   &H00C0C0FF&
            BorderStyle     =   0  'None
            ForeColor       =   &H00000000&
            Height          =   315
            Left            =   3675
            TabIndex        =   36
            TabStop         =   0   'False
            Top             =   450
            Width           =   780
         End
         Begin VB.TextBox Text8 
            Appearance      =   0  'Flat
            BackColor       =   &H00FFFF00&
            BorderStyle     =   0  'None
            ForeColor       =   &H00000000&
            Height          =   315
            Left            =   75
            TabIndex        =   35
            TabStop         =   0   'False
            Text            =   "SaveAndNew"
            Top             =   2325
            Width           =   3540
         End
         Begin VB.TextBox txtCoachSaveAndNewID 
            Appearance      =   0  'Flat
            BackColor       =   &H00FFFF00&
            BorderStyle     =   0  'None
            ForeColor       =   &H00000000&
            Height          =   315
            Left            =   3675
            TabIndex        =   34
            TabStop         =   0   'False
            Top             =   2325
            Width           =   780
         End
         Begin vbalIml6.vbalImageList lstIconList 
            Left            =   75
            Top             =   3075
            _ExtentX        =   953
            _ExtentY        =   953
            IconSizeX       =   26
            IconSizeY       =   32
            Size            =   24612
            Images          =   "Transfers.frx":05D2
            Version         =   131072
            KeyCount        =   7
            Keys            =   "ÿÿÿÿÿÿ"
         End
      End
      Begin VB.Frame frmButtonFrame 
         BackColor       =   &H00FF8080&
         BorderStyle     =   0  'None
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   161
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   615
         Left            =   150
         TabIndex        =   13
         Top             =   9750
         Width           =   16440
         Begin Dacara_dcButton.dcButton cmdButton 
            Height          =   465
            Index           =   9
            Left            =   13125
            TabIndex        =   14
            TabStop         =   0   'False
            Top             =   75
            Width           =   3165
            _ExtentX        =   5583
            _ExtentY        =   820
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
            ForeColor       =   0
            PicOpacity      =   0
         End
         Begin Dacara_dcButton.dcButton cmdButton 
            Height          =   465
            Index           =   6
            Left            =   3450
            TabIndex        =   15
            TabStop         =   0   'False
            Top             =   75
            Width           =   3165
            _ExtentX        =   5583
            _ExtentY        =   820
            BackColor       =   12640511
            ButtonShape     =   3
            ButtonStyle     =   4
            Caption         =   "Εκτύπωση"
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
            Height          =   465
            Index           =   7
            Left            =   6675
            TabIndex        =   16
            TabStop         =   0   'False
            Top             =   75
            Width           =   3165
            _ExtentX        =   5583
            _ExtentY        =   820
            BackColor       =   12640511
            ButtonShape     =   3
            ButtonStyle     =   4
            Caption         =   "Εκτύπωση για οδηγό"
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
            Height          =   465
            Index           =   8
            Left            =   9900
            TabIndex        =   46
            TabStop         =   0   'False
            Top             =   75
            Width           =   3165
            _ExtentX        =   5583
            _ExtentY        =   820
            BackColor       =   12640511
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
            PicOpacity      =   0
         End
         Begin Dacara_dcButton.dcButton cmdButton 
            Height          =   465
            Index           =   5
            Left            =   225
            TabIndex        =   78
            TabStop         =   0   'False
            Top             =   75
            Width           =   3165
            _ExtentX        =   5583
            _ExtentY        =   820
            BackColor       =   12640511
            ButtonShape     =   3
            ButtonStyle     =   4
            Caption         =   "Απόδοση δρομολογίου σε οδηγό"
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
      Begin iGrid300_10Tec.iGrid grdCoachesReport 
         Height          =   6090
         Left            =   7725
         TabIndex        =   11
         TabStop         =   0   'False
         Top             =   975
         Width           =   11190
         _ExtentX        =   19738
         _ExtentY        =   10742
         Appearance      =   0
         ForeColor       =   -2147483631
      End
      Begin UserControls.newDate mskDate 
         Height          =   465
         Left            =   1875
         TabIndex        =   0
         Top             =   975
         Width           =   1515
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
      Begin UserControls.newText txtCustomerDescription 
         Height          =   465
         Left            =   1875
         TabIndex        =   2
         Top             =   2025
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
      Begin UserControls.newText txtDestinationDescription 
         Height          =   465
         Left            =   1875
         TabIndex        =   1
         Top             =   1500
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
      Begin UserControls.newText txtPickupPointDescription 
         Height          =   465
         Left            =   1875
         TabIndex        =   4
         Top             =   3075
         Width           =   4965
         _ExtentX        =   8758
         _ExtentY        =   820
         ForeColor       =   0
         MaxLength       =   50
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
      Begin UserControls.newText txtRemarks 
         Height          =   465
         Left            =   1875
         TabIndex        =   8
         Top             =   5175
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
      Begin UserControls.newText txtRouteShortDescription 
         Height          =   465
         Left            =   1875
         TabIndex        =   3
         Top             =   2550
         Width           =   1515
         _ExtentX        =   2672
         _ExtentY        =   820
         Alignment       =   2
         ForeColor       =   0
         MaxLength       =   10
         Text            =   "AAAAAAAAAA"
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
      Begin UserControls.newInteger mskAdults 
         Height          =   465
         Left            =   1875
         TabIndex        =   5
         Top             =   3600
         Width           =   690
         _ExtentX        =   1217
         _ExtentY        =   820
         ForeColor       =   0
         MaxLength       =   3
         Text            =   "999"
         BackColor       =   4210688
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Ubuntu Condensed"
            Size            =   12
            Charset         =   161
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
      End
      Begin UserControls.newInteger mskKids 
         Height          =   465
         Left            =   1875
         TabIndex        =   6
         Top             =   4125
         Width           =   690
         _ExtentX        =   1217
         _ExtentY        =   820
         ForeColor       =   0
         MaxLength       =   3
         Text            =   "999"
         BackColor       =   4210688
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Ubuntu Condensed"
            Size            =   12
            Charset         =   161
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
      End
      Begin UserControls.newInteger mskFree 
         Height          =   465
         Left            =   1875
         TabIndex        =   7
         Top             =   4650
         Width           =   690
         _ExtentX        =   1217
         _ExtentY        =   820
         ForeColor       =   0
         MaxLength       =   3
         Text            =   "999"
         BackColor       =   4210688
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Ubuntu Condensed"
            Size            =   12
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
         Left            =   6900
         TabIndex        =   18
         TabStop         =   0   'False
         Top             =   2025
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
         PicNormal       =   "Transfers.frx":6616
         PicSizeH        =   16
         PicSizeW        =   16
      End
      Begin Dacara_dcButton.dcButton cmdIndex 
         Height          =   465
         Index           =   4
         Left            =   6900
         TabIndex        =   19
         TabStop         =   0   'False
         Top             =   3075
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
         PicNormal       =   "Transfers.frx":6BB0
         PicSizeH        =   16
         PicSizeW        =   16
      End
      Begin Dacara_dcButton.dcButton cmdIndex 
         Height          =   465
         Index           =   3
         Left            =   3450
         TabIndex        =   20
         TabStop         =   0   'False
         Top             =   2550
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
         PicNormal       =   "Transfers.frx":714A
         PicSizeH        =   16
         PicSizeW        =   16
      End
      Begin UserControls.newText txtDriverDescription 
         Height          =   465
         Left            =   1875
         TabIndex        =   9
         Top             =   5700
         Width           =   4965
         _ExtentX        =   8758
         _ExtentY        =   820
         ForeColor       =   0
         MaxLength       =   50
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
         Index           =   5
         Left            =   6900
         TabIndex        =   31
         TabStop         =   0   'False
         Top             =   5700
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
         PicNormal       =   "Transfers.frx":76E4
         PicSizeH        =   16
         PicSizeW        =   16
      End
      Begin Dacara_dcButton.dcButton cmdIndex 
         Height          =   465
         Index           =   1
         Left            =   6900
         TabIndex        =   61
         TabStop         =   0   'False
         Top             =   1500
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
         PicNormal       =   "Transfers.frx":7C7E
         PicSizeH        =   16
         PicSizeW        =   16
      End
      Begin Dacara_dcButton.dcButton cmdButton 
         Height          =   465
         Index           =   0
         Left            =   3450
         TabIndex        =   79
         Top             =   975
         Width           =   465
         _ExtentX        =   820
         _ExtentY        =   820
         BackColor       =   15133676
         ButtonShape     =   3
         ButtonStyle     =   7
         Caption         =   ""
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Ubuntu Condensed"
            Size            =   12
            Charset         =   161
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   12583104
         PicNormal       =   "Transfers.frx":8218
         PicSizeH        =   16
         PicSizeW        =   16
      End
      Begin Dacara_dcButton.dcButton cmdButton 
         Height          =   465
         Index           =   1
         Left            =   825
         TabIndex        =   82
         TabStop         =   0   'False
         Top             =   6450
         Width           =   1365
         _ExtentX        =   2408
         _ExtentY        =   820
         BackColor       =   12640511
         ButtonShape     =   3
         ButtonStyle     =   4
         Caption         =   "Δημιουργία"
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
         Height          =   465
         Index           =   3
         Left            =   3675
         TabIndex        =   83
         TabStop         =   0   'False
         Top             =   6450
         Width           =   1365
         _ExtentX        =   2408
         _ExtentY        =   820
         BackColor       =   12640511
         ButtonShape     =   3
         ButtonStyle     =   4
         Caption         =   "Διαγραφή"
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
         Height          =   465
         Index           =   4
         Left            =   5100
         TabIndex        =   84
         TabStop         =   0   'False
         Top             =   6450
         Width           =   1365
         _ExtentX        =   2408
         _ExtentY        =   820
         BackColor       =   12640511
         ButtonShape     =   3
         ButtonStyle     =   4
         Caption         =   "Ακύρωση"
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
         Height          =   465
         Index           =   2
         Left            =   2250
         TabIndex        =   85
         TabStop         =   0   'False
         Top             =   6450
         Width           =   1365
         _ExtentX        =   2408
         _ExtentY        =   820
         BackColor       =   12640511
         ButtonShape     =   3
         ButtonStyle     =   4
         Caption         =   "Αποθήκευση"
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
      Begin VB.Label mskTotal 
         BackStyle       =   0  'Transparent
         Caption         =   "100"
         BeginProperty Font 
            Name            =   "Ubuntu Condensed"
            Size            =   18
            Charset         =   161
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H0000FFFF&
         Height          =   465
         Left            =   3750
         TabIndex        =   86
         Top             =   4190
         Width           =   1215
      End
      Begin VB.Label lblBraces 
         AutoSize        =   -1  'True
         BackColor       =   &H000080FF&
         BackStyle       =   0  'Transparent
         Caption         =   "}"
         BeginProperty Font 
            Name            =   "Ubuntu Condensed"
            Size            =   48
            Charset         =   161
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00FFFF80&
         Height          =   1335
         Left            =   2700
         TabIndex        =   81
         Top             =   3750
         Width           =   255
      End
      Begin VB.Label lblRouteDescription 
         BackColor       =   &H000080FF&
         BackStyle       =   0  'Transparent
         Caption         =   "Διαδρομή"
         BeginProperty Font 
            Name            =   "Ubuntu Condensed"
            Size            =   11.25
            Charset         =   161
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00FFFFFF&
         Height          =   315
         Left            =   3900
         TabIndex        =   80
         Top             =   2625
         Width           =   3390
      End
      Begin VB.Label lblTotalPersonsForVisibleRows 
         Alignment       =   1  'Right Justify
         BackColor       =   &H00C0C0FF&
         BackStyle       =   0  'Transparent
         Caption         =   "Σύνολο φιλτραρισμένων"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   9.75
            Charset         =   161
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00C0C0FF&
         Height          =   315
         Left            =   13875
         TabIndex        =   64
         Top             =   300
         Width           =   5040
      End
      Begin VB.Label lblSelectedGridLines 
         Alignment       =   1  'Right Justify
         BackColor       =   &H00C0C0FF&
         BackStyle       =   0  'Transparent
         Caption         =   "Σύνολο επιλεγμένων"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   9.75
            Charset         =   161
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00FFFF00&
         Height          =   315
         Left            =   13875
         TabIndex        =   62
         Top             =   600
         Width           =   5040
      End
      Begin VB.Label lblLabel 
         BackColor       =   &H000080FF&
         Caption         =   "Οδηγός"
         ForeColor       =   &H00FFFFFF&
         Height          =   315
         Index           =   4
         Left            =   75
         TabIndex        =   32
         Top             =   5775
         Width           =   1365
      End
      Begin VB.Label lblLabel 
         AutoSize        =   -1  'True
         BackColor       =   &H000080FF&
         Caption         =   "Ημερομηνία"
         ForeColor       =   &H00FFFFFF&
         Height          =   315
         Index           =   3
         Left            =   75
         TabIndex        =   30
         Top             =   1050
         Width           =   1365
      End
      Begin VB.Label lblLabel 
         AutoSize        =   -1  'True
         BackColor       =   &H000080FF&
         Caption         =   "Πελάτης"
         ForeColor       =   &H00FFFFFF&
         Height          =   315
         Index           =   5
         Left            =   75
         TabIndex        =   29
         Top             =   2100
         Width           =   1365
      End
      Begin VB.Label lblLabel 
         AutoSize        =   -1  'True
         BackColor       =   &H000080FF&
         Caption         =   "Προορισμός"
         ForeColor       =   &H00FFFFFF&
         Height          =   315
         Index           =   7
         Left            =   75
         TabIndex        =   28
         Top             =   1575
         Width           =   1365
      End
      Begin VB.Label lblLabel 
         AutoSize        =   -1  'True
         BackColor       =   &H000080FF&
         Caption         =   "Σημείο παραλαβής"
         ForeColor       =   &H00FFFFFF&
         Height          =   315
         Index           =   8
         Left            =   75
         TabIndex        =   27
         Top             =   3150
         Width           =   1365
      End
      Begin VB.Label lblLabel 
         AutoSize        =   -1  'True
         BackColor       =   &H000080FF&
         Caption         =   "Δρομολόγιο"
         ForeColor       =   &H00FFFFFF&
         Height          =   255
         Index           =   9
         Left            =   75
         TabIndex        =   26
         Top             =   2625
         Width           =   840
      End
      Begin VB.Label lblLabel 
         AutoSize        =   -1  'True
         BackColor       =   &H000080FF&
         Caption         =   "Ενήλικες"
         ForeColor       =   &H00FFFFFF&
         Height          =   315
         Index           =   10
         Left            =   75
         TabIndex        =   25
         Top             =   3675
         Width           =   1365
      End
      Begin VB.Label lblLabel 
         AutoSize        =   -1  'True
         BackColor       =   &H000080FF&
         Caption         =   "Παιδιά"
         ForeColor       =   &H00FFFFFF&
         Height          =   315
         Index           =   6
         Left            =   75
         TabIndex        =   24
         Top             =   4200
         Width           =   1365
      End
      Begin VB.Label lblLabel 
         AutoSize        =   -1  'True
         BackColor       =   &H000080FF&
         Caption         =   "Δωρεάν"
         ForeColor       =   &H00FFFFFF&
         Height          =   315
         Index           =   2
         Left            =   75
         TabIndex        =   23
         Top             =   4725
         Width           =   1365
      End
      Begin VB.Label lblLabel 
         AutoSize        =   -1  'True
         BackColor       =   &H000080FF&
         Caption         =   "Σύνολο"
         ForeColor       =   &H00FFFFFF&
         Height          =   315
         Index           =   1
         Left            =   3075
         TabIndex        =   22
         Top             =   4275
         Width           =   615
      End
      Begin VB.Label lblLabel 
         AutoSize        =   -1  'True
         BackColor       =   &H000080FF&
         Caption         =   "Παρατηρήσεις"
         ForeColor       =   &H00FFFFFF&
         Height          =   315
         Index           =   0
         Left            =   75
         TabIndex        =   21
         Top             =   5250
         Width           =   1365
      End
      Begin VB.Shape shpWedge 
         BackColor       =   &H0000FFFF&
         BackStyle       =   1  'Opaque
         BorderStyle     =   0  'Transparent
         FillColor       =   &H00008000&
         Height          =   840
         Index           =   6
         Left            =   7275
         Top             =   3450
         Visible         =   0   'False
         Width           =   465
      End
      Begin VB.Shape shpWedge 
         BackColor       =   &H0000FFFF&
         BackStyle       =   1  'Opaque
         BorderStyle     =   0  'Transparent
         FillColor       =   &H00008000&
         Height          =   840
         Index           =   5
         Left            =   1425
         Top             =   1950
         Visible         =   0   'False
         Width           =   465
      End
      Begin VB.Label lblTotalPersons 
         Alignment       =   1  'Right Justify
         BackColor       =   &H00C0C0FF&
         BackStyle       =   0  'Transparent
         Caption         =   "Σύνολο ημέρας"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   9.75
            Charset         =   161
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00C0FFFF&
         Height          =   315
         Left            =   13875
         TabIndex        =   17
         Top             =   0
         Width           =   5040
      End
      Begin VB.Label lblTitle 
         AutoSize        =   -1  'True
         BackColor       =   &H00800000&
         BackStyle       =   0  'Transparent
         Caption         =   "Επιβαίνοντες λεωφορείων"
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
         Top             =   0
         Width           =   5760
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
Attribute VB_Name = "Transfers"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Dim blnStatus As Boolean
Dim blnCancel As Boolean
Dim lngRowCount As Long
Dim blnError As Boolean
Dim blnProcessing As Boolean

Dim lngMinimumSeperatorTop As Long
Dim lngMaximumSeperatorTop As Long
Dim lngOldSeperatorTop As Long
Dim blnIsMoving As Boolean

Dim lngTotalPersonsForVisibleRows As Long
Dim lngTotalPersonsForSelectedRows As Long

Dim lngCurrentRow As Long



Private Function AssignRoutesToDriver()

    Dim lngRow As Long
    Dim IsRowSelected As Boolean
    Dim IsError As Boolean
    Dim strDummy As String
    
    If txtDriverIDForRoutes.text = "" Then
        If MyMsgBox(4, strApplicationName, strStandardMessages(2), 1) Then
        End If
        txtDriverDescriptionForRoutes.SetFocus
        Exit Function
    End If
    
    frmCriteria(0).Visible = False
    UpdateButtons Me, 11, 0, 1, 0, 0, 1, 1, 1, 1, 0, 0, 0, 0
    
    BeginTrans
    
    For lngRow = 1 To grdCoachesReport.RowCount
        DoEvents
        If grdCoachesReport.CellIcon(lngRow, "Selected") = 3 Then
            AssignDriverToThisRoute grdCoachesReport.CellValue(lngRow, "TransferID")
        End If
    Next lngRow
    
    If IsError Then
        Rollback
        strDummy = MyMsgBox(4, strApplicationName, strStandardMessages(13), 1)
    Else
        CommitTrans
        FindRecordsAndPopulateGrid
        strDummy = MyMsgBox(1, strApplicationName, strStandardMessages(8), 1)
    End If

End Function

Private Function AssignDriverToThisRoute(TransferID)

    Dim lngTransferID As Long
    Dim rsTable As Recordset
    
    Set rsTable = CommonDB.OpenRecordset("Transfers")
    
    With rsTable
        .index = "ID"
        .Seek "=", TransferID
        If Not .NoMatch Then
            .Edit
            !TransferDriverID = txtDriverIDForRoutes.text
            .Update
        End If
    End With

End Function

Private Function CalculateSummaryPerCustomer()

    'SQL
    Dim intIndex As Byte
    Dim strThisQuery As String
    Dim strParameters As String
    Dim strParFields As String
    Dim strThisParameter As String
    Dim strOrder As String
    Dim strGroupBy As String
    Dim strLogic As String
    Dim arrQuery() As Variant
    Dim strSQL As String
    
    Dim lngRow As Long
    
    'Recordsets
    Dim rstRecordset As Recordset
    
    'SQL
    strSQL = "SELECT " _
        & "Customers.ID, Customers.Description, Sum(Transfers.TransferAdults+Transfers.TransferKids+Transfers.TransferFree) AS SumOfTransferPersons " _
        & "FROM Transfers INNER JOIN Customers ON Transfers.TransferCustomerID = Customers.ID " _
            
    'Ημερομηνία
    strThisParameter = "datDate Date"
    strThisQuery = "Transfers.TransferDate = datDate"
    strLogic = " AND "
    GoSub UpdateSQLString
    arrQuery(intIndex) = CDate(mskDate.text)
               
    strGroupBy = " GROUP BY Customers.ID, Customers.Description "
    strOrder = " ORDER BY Sum(Transfers.TransferAdults+Transfers.TransferKids+Transfers.TransferFree) DESC, Description"
    
    Set TempQuery = CommonDB.CreateQueryDef("")
    
    'Προσθέτω τα κριτήρια
    If strThisParameter <> "" Then
        strParameters = "PARAMETERS " & strParameters & "; "
        strParFields = "WHERE " & strParFields
        strSQL = strParameters & strSQL & strParFields
        TempQuery.SQL = strSQL & strGroupBy & strOrder
    End If
    
    'Κριτήρια
    If strThisParameter <> "" Then
        For intIndex = 1 To UBound(arrQuery)
            TempQuery.Parameters(intIndex - 1) = arrQuery(intIndex)
        Next intIndex
    End If
    
    'Ανοίγω το recordset
    Set rstRecordset = TempQuery.OpenRecordset()
    
    ClearFields grdSummaryPerCustomer
    
    'Γεμίζω το πλέγμα
    With rstRecordset
        DoEvents
        Do While Not .EOF
            grdSummaryPerCustomer.AddRow
            lngRow = lngRow + 1
            grdSummaryPerCustomer.CellIcon(lngRow, "Selected") = lstIconList.ItemIndex(3)
            grdSummaryPerCustomer.CellValue(lngRow, "CustomerID") = !ID
            grdSummaryPerCustomer.CellValue(lngRow, "CustomerDescription") = !Description
            grdSummaryPerCustomer.CellValue(lngRow, "TotalPersons") = !SumOfTransferPersons
            rstRecordset.MoveNext
        Loop
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
    blnError = True
    ClearFields grdSummaryPerCustomer
    DisplayErrorMessage True, Err.Description
    
End Function

Private Function CalculateSummaryPerDestination()

    'SQL
    Dim intIndex As Byte
    Dim strThisQuery As String
    Dim strParameters As String
    Dim strParFields As String
    Dim strThisParameter As String
    Dim strOrder As String
    Dim strGroupBy As String
    Dim strLogic As String
    Dim arrQuery() As Variant
    Dim strSQL As String
    
    Dim lngRow As Long
    
    'Recordsets
    Dim rstRecordset As Recordset
    
    'SQL
    strSQL = "SELECT " _
        & "DestinationID, DestinationDescription, Sum(Transfers.TransferAdults+Transfers.TransferKids+Transfers.TransferFree) AS SumOfTransferPersons " _
        & "FROM Transfers INNER JOIN Destinations ON Transfers.TransferDestinationID = Destinations.DestinationID " _
            
    'Ημερομηνία
    strThisParameter = "datDate Date"
    strThisQuery = "Transfers.TransferDate = datDate"
    strLogic = " AND "
    GoSub UpdateSQLString
    arrQuery(intIndex) = CDate(mskDate.text)
               
    strGroupBy = " GROUP BY DestinationID, DestinationDescription "
    strOrder = " ORDER BY Sum(Transfers.TransferAdults+Transfers.TransferKids+Transfers.TransferFree) DESC, DestinationDescription"
    
    Set TempQuery = CommonDB.CreateQueryDef("")
    
    'Προσθέτω τα κριτήρια
    If strThisParameter <> "" Then
        strParameters = "PARAMETERS " & strParameters & "; "
        strParFields = "WHERE " & strParFields
        strSQL = strParameters & strSQL & strParFields
        TempQuery.SQL = strSQL & strGroupBy & strOrder
    End If
    
    'Κριτήρια
    If strThisParameter <> "" Then
        For intIndex = 1 To UBound(arrQuery)
            TempQuery.Parameters(intIndex - 1) = arrQuery(intIndex)
        Next intIndex
    End If
    
    'Ανοίγω το recordset
    Set rstRecordset = TempQuery.OpenRecordset()
    
    ClearFields grdSummaryPerDestination
    
    'Γεμίζω το πλέγμα
    With rstRecordset
        Do While Not .EOF
            DoEvents
            grdSummaryPerDestination.AddRow
            lngRow = lngRow + 1
            grdSummaryPerDestination.CellIcon(lngRow, "Selected") = lstIconList.ItemIndex(2)
            grdSummaryPerDestination.CellValue(lngRow, "DestinationID") = !DestinationID
            grdSummaryPerDestination.CellValue(lngRow, "DestinationDescription") = !DestinationDescription
            grdSummaryPerDestination.CellValue(lngRow, "TotalPersons") = !SumOfTransferPersons
            rstRecordset.MoveNext
        Loop
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
    blnError = True
    ClearFields grdSummaryPerDestination
    DisplayErrorMessage True, Err.Description

End Function

Private Function CalculateSummaryPerDriver()

    'SQL
    Dim intIndex As Byte
    Dim strThisQuery As String
    Dim strParameters As String
    Dim strParFields As String
    Dim strThisParameter As String
    Dim strOrder As String
    Dim strGroupBy As String
    Dim strLogic As String
    Dim arrQuery() As Variant
    Dim strSQL As String
    
    Dim lngRow As Long
    
    'Recordsets
    Dim rstRecordset As Recordset
    
    'SQL
    strSQL = "SELECT " _
        & "TransferDriverID, DriverDescription, Sum(Transfers.TransferAdults+Transfers.TransferKids+Transfers.TransferFree) AS SumOfTransferPersons " _
        & "FROM Transfers LEFT JOIN Drivers ON Transfers.TransferDriverID = Drivers.DriverID " _
            
    'Ημερομηνία
    strThisParameter = "datDate Date"
    strThisQuery = "Transfers.TransferDate = datDate"
    strLogic = " AND "
    GoSub UpdateSQLString
    arrQuery(intIndex) = CDate(mskDate.text)
               
    strGroupBy = " GROUP BY TransferDriverID, DriverDescription "
    strOrder = " ORDER BY Sum(Transfers.TransferAdults+Transfers.TransferKids+Transfers.TransferFree) DESC, DriverDescription"
    
    Set TempQuery = CommonDB.CreateQueryDef("")
    
    'Προσθέτω τα κριτήρια
    If strThisParameter <> "" Then
        strParameters = "PARAMETERS " & strParameters & "; "
        strParFields = "WHERE " & strParFields
        strSQL = strParameters & strSQL & strParFields
        TempQuery.SQL = strSQL & strGroupBy & strOrder
    End If
    
    'Κριτήρια
    If strThisParameter <> "" Then
        For intIndex = 1 To UBound(arrQuery)
            TempQuery.Parameters(intIndex - 1) = arrQuery(intIndex)
        Next intIndex
    End If
    
    'Ανοίγω το recordset
    Set rstRecordset = TempQuery.OpenRecordset()
    
    ClearFields grdSummaryPerDriver
    
    'Γεμίζω το πλέγμα
    With rstRecordset
        Do While Not .EOF
            DoEvents
            grdSummaryPerDriver.AddRow
            lngRow = lngRow + 1
            grdSummaryPerDriver.CellIcon(lngRow, "Selected") = lstIconList.ItemIndex(5)
            grdSummaryPerDriver.CellValue(lngRow, "DriverID") = !TransferDriverID
            grdSummaryPerDriver.CellValue(lngRow, "DriverDescription") = IIf(IsNull(!DriverDescription), "-", !DriverDescription)
            grdSummaryPerDriver.CellValue(lngRow, "TotalPersons") = !SumOfTransferPersons
            rstRecordset.MoveNext
        Loop
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
    blnError = True
    ClearFields grdSummaryPerDriver
    DisplayErrorMessage True, Err.Description

End Function

Private Function CalculateSummaryPerRoute()

    'SQL
    Dim intIndex As Byte
    Dim strThisQuery As String
    Dim strParameters As String
    Dim strParFields As String
    Dim strThisParameter As String
    Dim strOrder As String
    Dim strGroupBy As String
    Dim strLogic As String
    Dim arrQuery() As Variant
    Dim strSQL As String
    
    Dim lngRow As Long
    
    'Recordsets
    Dim rstRecordset As Recordset
    
    'SQL
    strSQL = "SELECT Transfers.TransferRouteID, PickupRoutes.PickupRouteShortDescription, Sum(Transfers.TransferAdults+Transfers.TransferKids+Transfers.TransferFree) AS SumOfTransferPersons " _
        & "FROM Transfers INNER JOIN PickupRoutes ON Transfers.TransferRouteID = PickupRoutes.PickupRouteID" _

    'Ημερομηνία
    strThisParameter = "datDate Date"
    strThisQuery = "Transfers.TransferDate = datDate"
    strLogic = " AND "
    GoSub UpdateSQLString
    arrQuery(intIndex) = CDate(mskDate.text)
               
    strGroupBy = "GROUP BY Transfers.TransferRouteID, PickupRoutes.PickupRouteShortDescription"
    strOrder = "ORDER BY Sum(Transfers.TransferAdults+Transfers.TransferKids+Transfers.TransferFree) DESC, PickupRouteShortDescription"
    
    Set TempQuery = CommonDB.CreateQueryDef("")
    
    'Προσθέτω τα κριτήρια
    If strThisParameter <> "" Then
        strParameters = "PARAMETERS " & strParameters & "; "
        strParFields = " WHERE " & strParFields
        strSQL = strParameters & strSQL & strParFields
        TempQuery.SQL = strSQL & " " & strGroupBy & " " & strOrder
    End If
    
    'Κριτήρια
    If strThisParameter <> "" Then
        For intIndex = 1 To UBound(arrQuery)
            TempQuery.Parameters(intIndex - 1) = arrQuery(intIndex)
        Next intIndex
    End If
    
    'Ανοίγω το recordset
    Set rstRecordset = TempQuery.OpenRecordset()
    
    ClearFields grdSummaryPerRoute
    
    'Γεμίζω το πλέγμα
    With rstRecordset
        Do While Not .EOF
            DoEvents
            grdSummaryPerRoute.AddRow
            lngRow = lngRow + 1
            grdSummaryPerRoute.CellIcon(lngRow, "Selected") = lstIconList.ItemIndex(4)
            grdSummaryPerRoute.CellValue(lngRow, "RouteID") = !TransferRouteID
            grdSummaryPerRoute.CellValue(lngRow, "RouteDescription") = !PickupRouteShortDescription
            grdSummaryPerRoute.CellValue(lngRow, "TotalPersons") = !SumOfTransferPersons
            rstRecordset.MoveNext
        Loop
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
    blnError = True
    ClearFields grdSummaryPerRoute
    DisplayErrorMessage True, Err.Description
    
End Function

Private Function CalculateTotalPersons()

    'SQL
    Dim intIndex As Byte
    Dim strThisQuery As String
    Dim strParameters As String
    Dim strParFields As String
    Dim strThisParameter As String
    Dim strOrder As String
    Dim strGroupBy As String
    Dim strLogic As String
    Dim arrQuery() As Variant
    Dim strSQL As String
    
    Dim lngTotalPersons As Long
    
    'Recordsets
    Dim rstRecordset As Recordset
    
    'SQL
    strSQL = "SELECT " _
        & "Sum(Transfers.TransferAdults+Transfers.TransferKids+Transfers.TransferFree) AS SumOfTransferPersons " _
        & "FROM Transfers " _
            
    'Ημερομηνία
    strThisParameter = "datDate Date"
    strThisQuery = "Transfers.TransferDate = datDate"
    strLogic = " AND "
    GoSub UpdateSQLString
    arrQuery(intIndex) = CDate(mskDate.text)
               
    strGroupBy = ""
    strOrder = ""
    
    Set TempQuery = CommonDB.CreateQueryDef("")
    
    'Προσθέτω τα κριτήρια
    If strThisParameter <> "" Then
        strParameters = "PARAMETERS " & strParameters & "; "
        strParFields = "WHERE " & strParFields
        strSQL = strParameters & strSQL & strParFields
        TempQuery.SQL = strSQL & strGroupBy & strOrder
    End If
    
    'Κριτήρια
    If strThisParameter <> "" Then
        For intIndex = 1 To UBound(arrQuery)
            TempQuery.Parameters(intIndex - 1) = arrQuery(intIndex)
        Next intIndex
    End If
    
    'Ανοίγω το recordset
    Set rstRecordset = TempQuery.OpenRecordset()
    
    CalculateTotalPersons = rstRecordset.Fields(0)
    lngTotalPersonsForVisibleRows = rstRecordset.Fields(0)
    
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
    ClearFields grdSummaryPerDestination
    DisplayErrorMessage True, Err.Description

End Function

Private Function DeleteRecord()

    If MainDeleteRecord("CommonDB", "Transfers", strApplicationName, "ID", txtTransferID.text, "True") Then
        blnCancel = True
        ClearFields txtDestinationID, txtCustomerID, txtRouteID, txtPickupPointID, txtDriverID
        ClearFields txtDestinationDescription, txtCustomerDescription, txtRouteShortDescription, lblRouteDescription, txtPickupPointDescription, mskAdults, mskKids, mskFree, txtRemarks, txtDriverDescription
        ClearFields mskTotal
        DisableFields mskDate, txtCustomerDescription, txtDestinationDescription, txtPickupPointDescription, txtRouteShortDescription, mskAdults, mskKids, mskFree, txtRemarks, txtDriverDescription
        DisableFields cmdIndex(0), cmdIndex(1), cmdIndex(2), cmdIndex(3), cmdIndex(4), cmdIndex(5)
        UpdateButtons Me, 11, 0, 1, 0, 0, 0, 1, 1, 1, 1, 1, 0, 0
        FindRecordsAndPopulateGrid
        If Not blnStatus Then ClearFields txtTransferID
        blnStatus = True
    End If

End Function

Private Function DisplayAssignRoutesToDriverDialog()

    Dim lngRow As Long
    Dim IsRowSelected As Boolean
    
    For lngRow = 1 To grdCoachesReport.RowCount
        If grdCoachesReport.CellIcon(lngRow, "Selected") = 3 Then
            IsRowSelected = True
            Exit For
        End If
    Next lngRow
    
    If Not IsRowSelected Then
        If MyMsgBox(4, strApplicationName, strStandardMessages(6), 1) Then
        End If
        Exit Function
    End If

    ClearFields txtDriverIDForRoutes, txtDriverDescriptionForRoutes
    frmCriteria(0).Visible = True
    
    UpdateButtons Me, 11, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 1, 1
    txtDriverDescriptionForRoutes.SetFocus

End Function

Private Function FindRecordsAndPopulateGrid()

    If ValidateFields Then
        If RefreshList > 0 Then
            lblTotalPersons.Caption = "Σύνολο ημέρας: " & format(CalculateTotalPersons, "#,##0")
            lblTotalPersonsForVisibleRows.Caption = "Συνολο φιλτραρισμένων: " & format(lngTotalPersonsForVisibleRows, "#,##0")
            lblSelectedGridLines.Caption = "Σύνολο επιλεγμένων: 0"
            'Σύνολα
            CalculateSummaryPerCustomer
            CalculateSummaryPerDestination
            CalculateSummaryPerRoute
            CalculateSummaryPerDriver
            'Επιλογή όλων
            chkAllDestinations.Value = 1
            chkAllCustomers.Value = 1
            chkAllRoutes.Value = 1
            chkAllDrivers.Value = 1
            'Εμφανίζω ή όχι εγγραφές
            ScanGridsForSelectedRows
            'Υπόλοιπα
            EnableGrid grdCoachesReport, False
            EnableFields chkAllVisibleTransfers, chkAllDestinations, chkAllCustomers, chkAllRoutes, chkAllDrivers
            DisableFields mskDate
            HighlightRow txtTransferID.text
            UpdateButtons Me, 11, 0, 1, 0, 0, 0, 1, 1, 1, 1, 0, 0, 0
            Exit Function
        Else
            UpdateButtons Me, 11, 1, 1, 0, 0, 0, 0, 0, 0, 0, 1, 0, 0
            If Not blnError Then
                If blnProcessing Then
                    If MyMsgBox(4, strApplicationName, strStandardMessages(27), 1) Then
                    End If
                Else
                    If MyMsgBox(1, strApplicationName, strStandardMessages(7), 1) Then
                    End If
                    mskDate.SetFocus
                End If
            End If
            blnError = False
            blnProcessing = False
        End If
    End If

End Function
Private Function HighlightRow(TransferID)

    Dim lngRow As Long
    
    lngRow = grdCoachesReport.FindSearchMatchRow("TransferID", TransferID)
    
    'Αν δεν έχω διαγράψει την εγγραφή
    If lngRow <> 0 Then
        grdCoachesReport.EnsureVisibleRow (lngRow)
        grdCoachesReport.CurRow = lngRow
    Else
        grdCoachesReport.EnsureVisibleRow (IIf(lngCurrentRow - 1 > 0, lngCurrentRow - 1, 1))
        grdCoachesReport.CurRow = (IIf(lngCurrentRow - 1 > 0, lngCurrentRow - 1, 1))
    End If
    
    grdCoachesReport.SetFocus
    
End Function

Private Function PositionSeperator()
    
    Seperator.Left = 75
    Seperator.Width = frmContainer.Width - Seperator.Left - 75
    Seperator.Top = GetSetting(appName:=strApplicationName, Section:="Settings", Key:="SeperatorTop")

End Function

Private Function RecolorizeControls()

    Dim intIndex As Integer
    
    For intIndex = 0 To lblLabel.UBound
        lblLabel(intIndex).BackColor = vbBlack
        lblLabel(intIndex).ForeColor = vbWhite
    Next intIndex
    
    chkAllVisibleTransfers.BackColor = vbBlack
    chkAllDestinations.BackColor = vbBlack
    chkAllCustomers.BackColor = vbBlack
    chkAllRoutes.BackColor = vbBlack
    chkAllDrivers.BackColor = vbBlack
    
    chkAllVisibleTransfers.ForeColor = vbWhite
    chkAllDestinations.ForeColor = vbWhite
    chkAllCustomers.ForeColor = vbWhite
    chkAllRoutes.ForeColor = vbWhite
    chkAllDrivers.ForeColor = vbWhite

End Function

Private Function RunActiveReport()

    On Error GoTo ErrTrap
    
    'With rptDriverReport
    '    .Caption = lblTitle.Caption
    '    .Restart
    '    If intPreviewReports = 1 Then
    '        .Zoom = -2
    '        .Printer.ColorMode = vbPRCMMonochrome
    '        .WindowState = vbMaximized
    '        .Run False
    '        .Show 1
    '    Else
    '        If GetSetting(appName:=strApplicationName, Section:="Settings", Key:="IsDevelopment") = "1" Then
    '            MsgBox "Development Mode: Will not print!", vbInformation
    '            Exit Function
    '        Else
    '            .Printer.DeviceName = strPrinterName
    '            .PrintReport False
    '            .Run True
    '        End If
    '    End If
    'End With
    
    RunActiveReport = True
    
    Exit Function
    
ErrTrap:
    RunActiveReport = False
    DisplayErrorMessage True, Err.Description

End Function

Private Function DoControlBreak(gridName As iGrid, totalName, ParamArray levelName() As Variant)

    Dim lngRow As Long
    Dim intLoop As Integer
    Dim level As Integer
    Dim curRouteTotal As Currency
    Dim curDailyTotal As Currency
    Dim curGrandTotal As Currency
    ReDim oldArea(UBound(levelName)) As String
    
    gridName.Redraw = False
    blnProcessing = True
    
    gridName.AddRow 1
    
    For intLoop = 0 To UBound(levelName)
        oldArea(intLoop) = gridName.CellValue(1, levelName(intLoop))
    Next intLoop
    
    lngRow = 1
    level = UBound(levelName)
    
    Do While True
        Do While oldArea(level) = gridName.CellValue(lngRow, levelName(level))
            If oldArea(level) = gridName.CellValue(lngRow, levelName(level)) Then
                curRouteTotal = curRouteTotal + gridName.CellValue(lngRow, totalName)
            Else
                GoSub AddTotalLineAndUpdateLevels
            End If
            lngRow = lngRow + 1
            If lngRow >= gridName.RowCount Then Exit Do
            DoEvents
            If Not blnProcessing Then
                gridName.Redraw = True
                DoControlBreak = False
                Exit Function
            End If
        Loop
        GoSub AddTotalLineAndUpdateLevels
        If level - 1 >= 0 Then
            If gridName.CellValue(lngRow, levelName(level - 1)) <> "" Then
                If oldArea(level - 1) <> gridName.CellValue(lngRow, levelName(level - 1)) Then
                    GoSub AddTotalLineAndUpdateLevels
                End If
            End If
        End If
        If lngRow >= gridName.RowCount Then
            Exit Do
        End If
    Loop
    
    GoSub AddDailyTotal
    GoSub AddGrandTotalLine
    
    gridName.Redraw = True
    
    Exit Function
    
AddTotalLineAndUpdateLevels:
    gridName.AddRow "", lngRow, , , , , , 1
    gridName.CellValue(lngRow, 3) = "     ΣΥΝΟΛΟ: " & curRouteTotal
    curDailyTotal = curDailyTotal + curRouteTotal
    gridName.CellForeColor(lngRow, 3) = vbCyan
    oldArea(level) = gridName.CellValue(lngRow + 1, levelName(level))
    If level - 1 >= 0 Then
        If gridName.CellValue(lngRow + 1, levelName(level - 1)) <> "" Then
            If oldArea(level - 1) <> gridName.CellValue(lngRow + 1, levelName(level - 1)) Then
                GoSub AddDailyTotal
            End If
            oldArea(level - 1) = gridName.CellValue(lngRow + 1, levelName(level - 1))
        End If
    End If
    lngRow = lngRow + 1
    
    curRouteTotal = 0
    
    Return
    
AddDailyTotal:
    If lngRow >= gridName.RowCount Then
        lngRow = gridName.RowCount
    Else
        lngRow = lngRow + 1
        gridName.AddRow "", lngRow
    End If
    gridName.CellForeColor(lngRow, 3) = vbCyan
    gridName.CellValue(lngRow, 3) = "     ΣΥΝΟΛΟ ΗΜΕΡΑΣ: " & curDailyTotal
    curGrandTotal = curGrandTotal + curDailyTotal
    curDailyTotal = 0
    
    Return

AddGrandTotalLine:
    gridName.AddRow ""
    gridName.CellValue(gridName.RowCount, 3) = "     ΓΕΝΙΚΟ ΣΥΝΟΛΟ: " & curGrandTotal
    gridName.CellForeColor(gridName.RowCount, 3) = vbCyan
    
    Return
    
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
    
    'Αρχικές τιμές
    intIndex = 0
    lngRow = 0
    
    'Πλέγμα
    With grdCoachesReport
        .Clear
        .Redraw = False
    End With
    
    'Κυρίως διαδικασία
    strSQL = "SELECT " _
        & "TransferID, TransferDate, TransferAdults, TransferKids, TransferFree, TransferRemarks, " _
        & "PickUpPointHotelDescription, PickUpPointExactPoint, PickUpPointTime, " _
        & "PickupRouteShortDescription, PickupRouteDescription, " _
        & "Description, " _
        & "DestinationShortDescription, DestinationDescription, " _
        & "DriverDescription " _
        & "FROM (((((Transfers " _
        & "LEFT JOIN PickupPoints ON Transfers.TransferPickupPointID = PickupPoints.PickUpPointID) " _
        & "LEFT JOIN PickupRoutes ON Transfers.TransferRouteID = PickupRoutes.PickupRouteID) " _
        & "LEFT JOIN Customers ON Transfers.TransferCustomerID = Customers.ID) " _
        & "LEFT JOIN Drivers ON Transfers.TransferDriverID = Drivers.DriverID) " _
        & "INNER JOIN Destinations ON Transfers.TransferDestinationID = Destinations.DestinationID) "
    
    'Ημερομηνία
    strThisParameter = "datDate Date"
    strThisQuery = "Transfers.TransferDate = datDate"
    strLogic = " AND "
    GoSub UpdateSQLString
    arrQuery(intIndex) = CDate(mskDate.text)
    
    'Ταξινόμηση
    strOrder = "ORDER BY PIckupRouteDescription, PickUpPointHotelDescription, PickupPointTime"

    Set TempQuery = CommonDB.CreateQueryDef("")
    
    'Προσθέτω τα κριτήρια
    If strThisParameter <> "" Then
        strParameters = "PARAMETERS " & strParameters & "; "
        strParFields = "WHERE " & strParFields
        strSQL = strParameters & strSQL & strParFields
        TempQuery.SQL = strSQL & " " & strOrder
    Else
        TempQuery.SQL = strSQL & " " & strOrder
    End If
    
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
    UpdateButtons Me, 11, 0, 0, 0, 0, 0, 0, 0, 0, 1, 0, 0, 0
    cmdButton(8).Caption = "Διακοπή επεξεργασίας"
    blnProcessing = True
    
    'Γεμίζω το πλέγμα
    With rstRecordset
        Do While Not .EOF
            grdCoachesReport.AddRow
            lngRowCount = rstRecordset.RecordCount
            UpdateProgressBar Me
            lngRow = lngRow + 1
            grdCoachesReport.CellValue(lngRow, "TransferID") = !TransferID
            grdCoachesReport.CellValue(lngRow, "TransferDate") = !transferDate
            grdCoachesReport.CellValue(lngRow, "CustomerDescription") = !Description
            grdCoachesReport.CellValue(lngRow, "DestinationShortDescription") = !DestinationShortDescription
            grdCoachesReport.CellValue(lngRow, "DestinationDescription") = !DestinationDescription
            grdCoachesReport.CellValue(lngRow, "RouteShortDescription") = !PickupRouteShortDescription
            grdCoachesReport.CellValue(lngRow, "RouteDescription") = !PickupRouteDescription
            grdCoachesReport.CellValue(lngRow, "PickupPointHotelDescription") = !PickupPointHotelDescription
            grdCoachesReport.CellValue(lngRow, "PickUpPointExactPoint") = !PickupPointExactPoint
            grdCoachesReport.CellValue(lngRow, "PickUpPointTime") = !PickupPointTime
            grdCoachesReport.CellValue(lngRow, "TransferAdults") = IIf(!TransferAdults > 0, !TransferAdults, "")
            grdCoachesReport.CellValue(lngRow, "TransferKids") = IIf(!TransferKids > 0, !TransferKids, "")
            grdCoachesReport.CellValue(lngRow, "TransferFree") = IIf(!TransferFree > 0, !TransferFree, "")
            grdCoachesReport.CellValue(lngRow, "TransferTotal") = !TransferAdults + !TransferKids + !TransferFree
            grdCoachesReport.CellValue(lngRow, "TransferRemarks") = !TransferRemarks
            grdCoachesReport.CellValue(lngRow, "DriverDescription") = IIf(IsNull(!DriverDescription), "-", !DriverDescription)
            grdCoachesReport.CellValue(lngRow, "TotalCriteria") = 0
            rstRecordset.MoveNext
            DoEvents
            If Not blnProcessing Then Exit Do
        Loop
    End With
    
    'Ακύρωση επεξεργασίας
    If Not blnProcessing Then
        blnProcessing = True
        ClearFields grdCoachesReport
        RefreshList = 0
    Else
        blnProcessing = False
        RefreshList = lngRowCount
    End If
    
    'Τελικές ενέργειες
    frmProgress.Visible = False
    Me.Refresh
    cmdButton(8).Caption = "Νέα αναζήτηση"
   
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
    ClearFields grdCoachesReport
    DisplayErrorMessage True, Err.Description

End Function

Private Function RemoveTotals(gridName As iGrid)

    Dim lngRow As Long
    
    gridName.Redraw = False
    
    For lngRow = 1 To gridName.RowCount
        If gridName.CellValue(lngRow, "TransferDate") = "" Then
            gridName.RemoveRow (lngRow)
            lngRow = lngRow - 1
            If lngRow = gridName.RowCount Then
                Exit For
            End If
        End If
    Next lngRow
    
    gridName.Redraw = False

End Function

Private Function ToggleRowVisibility(fieldDescription, lookupString, showOrHide)

    Dim lngRow As Long
    
    For lngRow = 1 To grdCoachesReport.RowCount
        If grdCoachesReport.CellValue(lngRow, grdCoachesReport.ColIndex(fieldDescription)) = lookupString Then
            grdCoachesReport.CellValue(lngRow, "TotalCriteria") = grdCoachesReport.CellValue(lngRow, "TotalCriteria") + 1
            'grdCoachesReport.RowVisible(lngRow) = showOrHide
        End If
    Next lngRow
    
    'lblTotalPersonsForVisibleRows.Caption = "Συνολο φιλτραρισμένων: " & format(lngTotalPersonsForVisibleRows, "#,##0")
    
End Function

Private Sub chkAllCustomers_Click()

    Dim lngRow As Long
    
    grdCoachesReport.Redraw = False
    lngTotalPersonsForVisibleRows = 0
    
    For lngRow = 1 To grdSummaryPerCustomer.RowCount
        grdSummaryPerCustomer.CellIcon(lngRow, "Selected") = lstIconList.ItemIndex(IIf(chkAllCustomers.Value = 0, 1, 3))
        ScanGridsForSelectedRows
    Next lngRow
    
    grdCoachesReport.Redraw = True

End Sub

Private Function ScanGridsForSelectedRows()

    Dim lngRow As Long
    Dim lngCoachesReport As Long
    Dim lngTotalPersonsForVisibleRows As Long
    
    'Μηδενίζω το πλέγμα των εγγραφών
    chkAllVisibleTransfers.Value = 0
    For lngRow = 1 To grdCoachesReport.RowCount
        grdCoachesReport.CellValue(lngRow, "TotalCriteria") = 0
    Next lngRow
    
    'Για κάθε μία εγγραφή των προορισμών
    For lngRow = 1 To grdSummaryPerDestination.RowCount
        'Αν την έχω επιλέξει
        If grdSummaryPerDestination.CellIcon(lngRow, "Selected") <> 0 Then
            'Σκανάρω το πλέγμα των εγγραφών
            For lngCoachesReport = 1 To grdCoachesReport.RowCount
                'Αν ο προορισμός είναι ο ίδιος
                If grdCoachesReport.CellValue(lngCoachesReport, "DestinationDescription") = grdSummaryPerDestination.CellValue(lngRow, "DestinationDescription") Then
                    'Αυξάνω το μετρητή κατά 1
                    grdCoachesReport.CellValue(lngCoachesReport, "TotalCriteria") = grdCoachesReport.CellValue(lngCoachesReport, "TotalCriteria") + 1
                End If
            Next lngCoachesReport
        End If
    Next lngRow

    'Για κάθε μία εγγραφή των πελατών
    For lngRow = 1 To grdSummaryPerCustomer.RowCount
        'Αν την έχω επιλέξει
        If grdSummaryPerCustomer.CellIcon(lngRow, "Selected") <> 0 Then
            'Σκανάρω το πλέγμα των εγγραφών
            For lngCoachesReport = 1 To grdCoachesReport.RowCount
                'Αν ο πελάτης είναι ο ίδιος
                If grdCoachesReport.CellValue(lngCoachesReport, "CustomerDescription") = grdSummaryPerCustomer.CellValue(lngRow, "CustomerDescription") Then
                    'Αυξάνω το μετρητή κατά 1
                    grdCoachesReport.CellValue(lngCoachesReport, "TotalCriteria") = grdCoachesReport.CellValue(lngCoachesReport, "TotalCriteria") + 1
                End If
            Next lngCoachesReport
        End If
    Next lngRow

    'Για κάθε μία εγγραφή των δρομολογίων
    For lngRow = 1 To grdSummaryPerRoute.RowCount
        'Αν την έχω επιλέξει
        If grdSummaryPerRoute.CellIcon(lngRow, "Selected") <> 0 Then
            'Σκανάρω το πλέγμα των εγγραφών
            For lngCoachesReport = 1 To grdCoachesReport.RowCount
                'Αν η δρομολόγιο είναι η ίδια
                If grdCoachesReport.CellValue(lngCoachesReport, "RouteShortDescription") = grdSummaryPerRoute.CellValue(lngRow, "RouteDescription") Then
                    'Αυξάνω το μετρητή κατά 1
                    grdCoachesReport.CellValue(lngCoachesReport, "TotalCriteria") = grdCoachesReport.CellValue(lngCoachesReport, "TotalCriteria") + 1
                End If
            Next lngCoachesReport
        End If
    Next lngRow

    'Για κάθε μία εγγραφή των οδηγών
    For lngRow = 1 To grdSummaryPerDriver.RowCount
        'Αν την έχω επιλέξει
        If grdSummaryPerDriver.CellIcon(lngRow, "Selected") <> 0 Then
            'Σκανάρω το πλέγμα των εγγραφών
            For lngCoachesReport = 1 To grdCoachesReport.RowCount
                'Αν ο οδηγός είναι ο ίδιος
                If grdCoachesReport.CellValue(lngCoachesReport, "DriverDescription") = grdSummaryPerDriver.CellValue(lngRow, "DriverDescription") Then
                    'Αυξάνω το μετρητή κατά 1
                    grdCoachesReport.CellValue(lngCoachesReport, "TotalCriteria") = grdCoachesReport.CellValue(lngCoachesReport, "TotalCriteria") + 1
                End If
            Next lngCoachesReport
        End If
    Next lngRow
    
    'Σκανάρω το πλέγμα των εγγραφών για >=3  στο άθροισμα
    For lngCoachesReport = 1 To grdCoachesReport.RowCount
        grdCoachesReport.RowVisible(lngCoachesReport) = IIf(grdCoachesReport.CellValue(lngCoachesReport, "TotalCriteria") = 4, True, False)
        If grdCoachesReport.RowVisible(lngCoachesReport) Then lngTotalPersonsForVisibleRows = lngTotalPersonsForVisibleRows + grdCoachesReport.CellValue(lngCoachesReport, "TransferTotal")
    Next lngCoachesReport
    
    lblTotalPersonsForVisibleRows.Caption = "Συνολο με βάση τα επιλεγμένα φίλτρα: " & format(lngTotalPersonsForVisibleRows, "#,##0")

End Function


Private Sub chkAllDestinations_Click()

    Dim lngRow As Long
    
    grdCoachesReport.Redraw = False
    lngTotalPersonsForVisibleRows = 0
    
    For lngRow = 1 To grdSummaryPerDestination.RowCount
        grdSummaryPerDestination.CellIcon(lngRow, "Selected") = lstIconList.ItemIndex(IIf(chkAllDestinations.Value = 0, 1, 2))
        ScanGridsForSelectedRows
    Next lngRow
    
    grdCoachesReport.Redraw = True

End Sub

Private Sub chkAllDrivers_Click()

    Dim lngRow As Long
    
    grdCoachesReport.Redraw = False
    lngTotalPersonsForVisibleRows = 0
    
    For lngRow = 1 To grdSummaryPerDriver.RowCount
        grdSummaryPerDriver.CellIcon(lngRow, "Selected") = lstIconList.ItemIndex(IIf(chkAllDrivers.Value = 0, 1, 5))
        ScanGridsForSelectedRows
    Next lngRow
    
    grdCoachesReport.Redraw = True

End Sub


Private Sub chkAllRoutes_Click()
    
    Dim lngRow As Long
    
    grdCoachesReport.Redraw = False
    lngTotalPersonsForVisibleRows = 0
    
    For lngRow = 1 To grdSummaryPerRoute.RowCount
        grdSummaryPerRoute.CellIcon(lngRow, "Selected") = lstIconList.ItemIndex(IIf(chkAllRoutes.Value = 0, 1, 4))
        ScanGridsForSelectedRows
    Next lngRow
    
    grdCoachesReport.Redraw = True

End Sub

Private Sub chkAllVisibleTransfers_Click()

    Dim lngRow As Long
    
    grdCoachesReport.Redraw = True
    lngTotalPersonsForSelectedRows = 0
    
    For lngRow = 1 To grdCoachesReport.RowCount
        grdCoachesReport.CellIcon(lngRow, "Selected") = IIf(grdCoachesReport.RowVisible(lngRow), lstIconList.ItemIndex(IIf(chkAllVisibleTransfers.Value = 0, 1, 4)), lstIconList.ItemIndex(1))
        lngTotalPersonsForSelectedRows = lngTotalPersonsForSelectedRows + IIf(grdCoachesReport.CellIcon(lngRow, "Selected") = 3, grdCoachesReport.CellValue(lngRow, "TransferTotal"), 0)
    Next lngRow
    
    grdCoachesReport.Redraw = True
    
    lblSelectedGridLines.Caption = "Σύνολο επιλεγμένων: " & lngTotalPersonsForSelectedRows

End Sub


Private Sub cmdButton_Click(index As Integer)
                                
    Select Case index
        Case 0
            FindRecordsAndPopulateGrid
        Case 1
            NewRecord
        Case 2
            If SaveRecord And blnStatus Then CheckToCreateNewRecord
        Case 3
            DeleteRecord
        Case 4
            AbortProcedure True
        Case 5
            DisplayAssignRoutesToDriverDialog
        Case 6
            'DoReport "Print"
        Case 7
            DoReport "Print"
        Case 8
            AbortProcedure False
        Case 9
            AbortProcedure False
        Case 10
            AbortProcedure False
        Case 11
            AssignRoutesToDriver
    End Select

End Sub

Private Function NewRecord()

    If True Then
        If txtTransferID.text <> "" Then
            DisplayLastRecord txtTransferID.text
            ClearFields txtTransferID, txtCustomerID, txtRouteID, txtPickupPointID, txtDriverID
            ClearFields txtCustomerDescription, lblRouteDescription, txtPickupPointDescription, txtRouteShortDescription, mskAdults, mskKids, mskFree, txtRemarks, txtDriverDescription
            ClearFields mskTotal
            txtCustomerDescription.SetFocus
        End If
    Else
        ClearFields txtTransferID, txtDestinationID, txtCustomerID, txtRouteID, txtPickupPointID, txtDriverID
        ClearFields txtDestinationDescription, txtCustomerDescription, lblRouteDescription, txtPickupPointDescription, txtRouteShortDescription, mskAdults, mskKids, mskFree, txtRemarks, txtDriverDescription
        ClearFields mskTotal
        txtCustomerDescription.SetFocus
    End If

    blnStatus = True
    blnCancel = False
    DisableFields mskDate
    EnableFields txtDestinationDescription, txtCustomerDescription, txtRouteShortDescription, txtPickupPointDescription, mskAdults, mskKids, mskFree, txtRemarks, txtDriverDescription
    EnableFields cmdIndex(0), cmdIndex(1), cmdIndex(2), cmdIndex(3), cmdIndex(4), cmdIndex(5)
    UpdateButtons Me, 11, 0, 0, 1, 0, 1, 0, 0, 0, 0, 0, 0, 0, 0
        
    InitializeFields mskAdults, mskKids, mskFree, mskTotal
    
    txtDestinationDescription.SetFocus
        
End Function


Private Function DisplayLastRecord(lngLastRecord)

    If Not SeekRecord(lngLastRecord) Then Exit Function

End Function



Private Function SaveRecord()

    If Not ValidateFields Then Exit Function
    
    txtTransferID.text = MainSaveRecord("CommonDB", "Transfers", blnStatus, strApplicationName, "ID", txtTransferID.text, mskDate.text, txtDestinationID.text, txtCustomerID.text, txtRouteID.text, txtPickupPointID.text, mskAdults.text, mskKids.text, mskFree.text, txtRemarks.text, IIf(txtDriverID.text = "", Null, txtDriverID.text), 1, strCurrentUser)
    
    If txtTransferID.text <> "" Then
        SaveRecord = True
        blnCancel = True
        ClearFields txtDestinationID, txtCustomerID, txtRouteID, txtPickupPointID, txtDriverID
        ClearFields txtDestinationDescription, txtCustomerDescription, txtRouteShortDescription, lblRouteDescription, txtPickupPointDescription, mskAdults, mskKids, mskFree, txtRemarks, txtDriverDescription
        ClearFields mskTotal
        DisableFields mskDate, txtCustomerDescription, txtDestinationDescription, txtPickupPointDescription, txtRouteShortDescription, mskAdults, mskKids, mskFree, txtRemarks, txtDriverDescription
        DisableFields cmdIndex(0), cmdIndex(1), cmdIndex(2), cmdIndex(3), cmdIndex(4), cmdIndex(5)
        UpdateButtons Me, 11, 0, 1, 0, 0, 0, 1, 1, 1, 1, 1, 0, 0
        FindRecordsAndPopulateGrid
        HighlightRow txtTransferID.text
        If Not blnStatus Then ClearFields txtTransferID
        blnStatus = True
    Else
        DisplayErrorMessage True, strStandardMessages(5)
    End If

End Function


Private Function CheckToCreateNewRecord()

    If txtCoachSaveAndNewID.text = "1" Then
        cmdButton_Click 0
    End If

End Function



Private Function DoReport(action As String)
    
    On Error GoTo ErrTrap
    
    If action = "Print" Then
        If SelectPrinter("PrinterPrintsReports") Then
            CreateUnicodeFile "Αναφορά παραλαβών για " & mskDate.text, "Οδηγός: ", "", intPrinterReportDetailLines
            With rptOneLiner
                .oneLongField.Font.Size = 8
                If intPreviewReports = 1 Then
                    .Restart
                    .Zoom = -2
                    .WindowState = vbMaximized
                    .Show 1
                Else
                    .Printer.DeviceName = strPrinterName
                    .PrintReport False
                    .Run True
                End If
            End With
        End If
    End If
    
    If action = "CreatePDF" Then
        'CreateUnicodeFileForCustomers lblTitle.Caption & " " & txtPersonDescription.text, " από " & mskInvoiceDateIssueFrom.text & " έως " & mskInvoiceDateIssueTo.text, "", GetSetting(strApplicationName, "Settings", "Export Report Height")
        'CreateUnisexPDF lblTitle.Caption & " " & txtPersonDescription.text
        If MyMsgBox(1, strApplicationName, strStandardMessages(8), 1) Then
        End If
    End If
    
    Exit Function
    
ErrTrap:
    Close #1
    DisplayErrorMessage True, Err.Description
    
End Function

Private Function CreatePDF(fileName)

    On Error GoTo ErrTrap
    
    Dim pdf As New ARExportPDF
    
    'With rptCoachesReport
    '    .Restart
    '    .Run False
    '    pdf.SemiDelimitedNeverEmbedFonts = ""
    '    pdf.fileName = Replace(fileName, "/", "-")
    '    pdf.fileName = Replace(pdf.fileName, "[", "")
    '    pdf.fileName = Replace(pdf.fileName, "]", "")
    '    pdf.fileName = Replace(pdf.fileName, "  ", " ")
    '    pdf.fileName = strReportsPathName & Replace(pdf.fileName, ":", "") & ".pdf"
    '    pdf.Export .Pages
    'End With
    
    CreatePDF = True
    
    Exit Function
    
ErrTrap:
    CreatePDF = False
    DisplayErrorMessage True, Err.Description
    
End Function

Private Function ValidateFields()

    ValidateFields = False
    
    'Ημερομηνία
    If mskDate.text = "" Then
        If MyMsgBox(4, strApplicationName, strStandardMessages(1), 1) Then
        End If
        mskDate.SetFocus
        Exit Function
    End If
    If Not IsDate(mskDate.text) Then
        If MyMsgBox(4, strApplicationName, strStandardMessages(2), 1) Then
        End If
        mskDate.SetFocus
        Exit Function
    End If
    
    ValidateFields = True

End Function

Private Function AbortProcedure(blnStatus)

    If blnProcessing Then blnProcessing = False: Exit Function

    'Πλαίσιο απόδοσης οδηγού σε δρομολόγιο
    If frmCriteria(0).Visible Then
        frmCriteria(0).Visible = False
        UpdateButtons Me, 11, 0, 1, 0, 0, 0, 1, 1, 1, 1, 0, 0, 0
        grdCoachesReport.SetFocus
        Exit Function
    End If
    
    'Επεξεργάζομαι εγγραφή (νέα ή μεταβολή)
    If blnStatus Then
        If MyMsgBox(3, strApplicationName, strStandardMessages(3), 2) Then
            blnStatus = False
            blnCancel = True
            ClearFields txtTransferID, txtCustomerID, txtPickupPointID, txtRouteID, txtDestinationID, txtDriverID
            ClearFields txtCustomerDescription, txtDestinationDescription, txtPickupPointDescription, txtRouteShortDescription, lblRouteDescription, mskAdults, mskKids, mskFree, txtRemarks, txtDriverDescription
            ClearFields mskTotal
            DisableFields txtCustomerDescription, txtDestinationDescription, txtPickupPointDescription, txtRouteShortDescription, mskAdults, mskKids, mskFree, txtRemarks, txtDriverDescription
            DisableFields cmdIndex(1), cmdIndex(2), cmdIndex(3), cmdIndex(4), cmdIndex(5)
            UpdateButtons Me, 11, 0, 1, 0, 0, 0, 1, 1, 1, 1, 0, 0, 0
            grdCoachesReport.SetFocus
            blnStatus = True
            Exit Function
        End If
        Exit Function
    End If
    
    'Νέα αναζήτηση
    If Not blnStatus And Not cmdButton(9).Enabled Then
        ClearFields grdCoachesReport
        ClearFields grdSummaryPerDestination, grdSummaryPerCustomer, grdSummaryPerRoute, grdSummaryPerDriver
        ClearFields lblTotalPersons, lblTotalPersonsForVisibleRows, lblSelectedGridLines
        ClearFields chkAllVisibleTransfers, chkAllVisibleTransfers, chkAllDestinations, chkAllCustomers, chkAllRoutes, chkAllDrivers
        DisableFields chkAllVisibleTransfers, chkAllVisibleTransfers, chkAllDestinations, chkAllCustomers, chkAllRoutes, chkAllDrivers
        UpdateButtons Me, 11, 1, 0, 0, 0, 0, 0, 0, 0, 0, 1, 0, 0
        EnableFields mskDate
        mskDate.SetFocus
        Exit Function
    End If
    
    Unload Me

End Function

Private Sub cmdIndex_Click(index As Integer)

    Dim strShowInList As String
    Dim tmpTableData As typTableData
    Dim tmpRecordset As Recordset
    Dim strSQL As String
    Dim intSize As Integer

    Select Case index
        'Οδηγός σε σύνδεση δρομολογίου
        Case 0
            Set tmpRecordset = CheckForMatch("CommonDB", "Drivers", "DriverDescription", "String", txtDriverDescriptionForRoutes.text)
            If tmpRecordset.RecordCount > 0 Then
                tmpTableData = DisplayIndex(tmpRecordset, 2, True, 2, 0, 1, "ID", "Περιγραφή", 0, 40, 1, 0)
                txtDriverIDForRoutes.text = tmpTableData.strCode
                txtDriverDescriptionForRoutes.text = tmpTableData.strFirstField
                cmdButton(10).SetFocus
            End If
        Case 1
            'Προορισμός
            Set tmpRecordset = CheckForMatch("CommonDB", "Destinations", "DestinationDescription", "String", txtDestinationDescription.text)
            If tmpRecordset.RecordCount > 0 Then
                tmpTableData = DisplayIndex(tmpRecordset, 2, True, 2, 0, 2, "ID", "Περιγραφή", 0, 40, 1, 0)
                txtDestinationID.text = tmpTableData.strCode
                txtDestinationDescription.text = tmpTableData.strFirstField
            End If
        Case 2
            'Πελάτης
            Set tmpRecordset = CheckForMatch("CommonDB", "Customers", "Description", "String", txtCustomerDescription.text)
            If tmpRecordset.RecordCount > 0 Then
                tmpTableData = DisplayIndex(tmpRecordset, 2, True, 2, 0, 1, "ID", "Περιγραφή", 0, 40, 1, 0)
                txtCustomerID.text = tmpTableData.strCode
                txtCustomerDescription.text = tmpTableData.strFirstField
            End If
        Case 3
            'Δρομολόγιο - Αν έχω δώσει προορισμό βρίσκω τα δρομολόγια για το δοσμένο προορισμό
            If txtDestinationID.text <> "" And txtRouteID.text <> "" Then
                'intSize = Len(txtPickupPointDescription.text)
                'strSQL = "SELECT DestinationID, RouteID, DestinationsRoutesPickupPoints.PickupPointID, PickupPointHotelDescription, PickupPointTime " _
                '    & "FROM DestinationsRoutesPickupPoints " _
                '    & "INNER JOIN PickupPoints ON DestinationsRoutesPickupPoints.PickupPointID = PickupPoints.PickupPointID " _
                '    & "WHERE RouteID = " & txtRouteID.text & " AND DestinationID = " & txtDestinationID.text & " " _
                '    & IIf(txtPickupPointDescription.text <> "", "AND Left(PickupPointHotelDescription, " & intSize & ") = '" & txtPickupPointDescription.text & "'", "")
                'Set tmpRecordset = FindAndReturnRecords(strSQL)
                'If tmpRecordset.RecordCount > 0 Then
                '    tmpTableData = DisplayIndex(tmpRecordset, 2, True, 3, 2, 3, 4, "ID", "Περιγραφή", "Ωρα", 0, 40, 7, 1, 0, 1)
                '    txtPickupPointID.text = tmpTableData.strCode
                '    txtPickupPointDescription.text = tmpTableData.strFirstField
                'End If
            End If
            'Δρομολόγιο - έχω δώσει προορισμό - δεν έχω δώσει σημείο παραλαβής
            If txtDestinationID.text <> "" And txtRouteID.text = "" Then
                intSize = Len(txtRouteShortDescription.text)
                strSQL = "SELECT DISTINCT DestinationsRoutesPickupPoints.RouteID, DestinationID, PickupRouteShortDescription, PickupRouteDescription " _
                    & "FROM DestinationsRoutesPickupPoints " _
                    & "INNER JOIN PickupRoutes ON DestinationsRoutesPickupPoints.RouteID = PickupRoutes.PickupRouteID " _
                    & "WHERE DestinationID = " & txtDestinationID.text & " " _
                    & "AND Left(PickupRouteShortDescription, " & intSize & ") = '" & txtRouteShortDescription.text & "'"
                Set tmpRecordset = FindAndReturnRecords(strSQL)
                If tmpRecordset.RecordCount > 0 Then
                    tmpTableData = DisplayIndex(tmpRecordset, 2, True, 3, 0, 2, 3, "ID", "Συντ.", "Δρομολόγιο", 0, 5, 40, 1, 1, 0)
                    txtRouteID.text = tmpTableData.strCode
                    txtRouteShortDescription.text = tmpTableData.strFirstField
                    lblRouteDescription.Caption = tmpTableData.strSecondField
                End If
                
                'intSize = Len(txtPickupPointDescription.text)
                'strSQL = "SELECT DestinationID, RouteID, DestinationsRoutesPickupPoints.PickupPointID, PickupPointHotelDescription, PickupPointTime " _
                    & "FROM DestinationsRoutesPickupPoints " _
                    & "INNER JOIN PickupPoints ON DestinationsRoutesPickupPoints.PickupPointID = PickupPoints.PickupPointID " _
                    & "WHERE DestinationID = " & txtDestinationID.text & " " _
                    & IIf(txtPickupPointDescription.text <> "", "AND Left(PickupPointHotelDescription, " & intSize & ") = '" & txtPickupPointDescription.text & "'", "")
                'Set tmpRecordset = FindAndReturnRecords(strSQL)
                'If tmpRecordset.RecordCount > 0 Then
                '    tmpTableData = DisplayIndex(tmpRecordset, 2, True, 4, 1, 2, 3, 4, "ID", "RouteID", "Περιγραφή", "Ωρα", 0, 0, 40, 7, 1, 0, 0, 1)
                '    txtPickupPointID.text = tmpTableData.strFirstField
                '    txtPickupPointDescription.text = tmpTableData.strSecondField
                '    txtRouteID.text = tmpTableData.strCode
                '    FindRoute
                'End If
            End If
        Case 4
            'Σημείο παραλαβής
            'Εχω δώσει προορισμό
            If txtDestinationID.text <> "" Then
                'Εχω δώσει δρομολόγιο
                If txtRouteID.text <> "" Then
                    'Βρίσκω τα σημεία παραλαβής που είναι συνδεδεμένα με τον δοσμένο προορισμό και τη δρομολόγιο
                    intSize = Len(txtRouteShortDescription.text)
                    strSQL = "SELECT DestinationID, RouteID, DestinationsRoutesPickupPoints.PickupPointID, PickupPointHotelDescription, PickupPointTime " _
                        & "FROM DestinationsRoutesPickupPoints " _
                        & "INNER JOIN PickupPoints ON DestinationsRoutesPickupPoints.PickupPointID = PickupPoints.PickupPointID " _
                        & "WHERE DestinationID = " & txtDestinationID.text & " " _
                        & "AND RouteID = " & txtRouteID.text & " " _
                        & "ORDER BY PickUpPointTime"
                    Set tmpRecordset = FindAndReturnRecords(strSQL)
                    If tmpRecordset.RecordCount > 0 Then
                        tmpTableData = DisplayIndex(tmpRecordset, 4, True, 4, 1, 2, 3, 4, "ID", "RouteID", "Περιγραφή", "Ωρα", 0, 0, 40, 7, 1, 0, 0, 1)
                        txtPickupPointID.text = tmpTableData.strFirstField
                        txtPickupPointDescription.text = tmpTableData.strSecondField
                        txtRouteID.text = tmpTableData.strCode
                    End If
                Else
                    'Δεν έχω δώσει δρομολόγιο, βρίσκω τα σημεία παραλαβής που είναι συνδεδεμένα με τον δοσμένο προορισμό
                    If txtRouteID.text = "" Then
                        intSize = Len(txtPickupPointDescription.text)
                        strSQL = "SELECT DestinationID, RouteID, DestinationsRoutesPickupPoints.PickupPointID, PickupPointHotelDescription, PickupPointTime " _
                            & "FROM DestinationsRoutesPickupPoints " _
                            & "INNER JOIN PickupPoints ON DestinationsRoutesPickupPoints.PickupPointID = PickupPoints.PickupPointID " _
                            & "WHERE DestinationID = " & txtDestinationID.text & " " _
                            & "AND Left(PickupPointHotelDescription, " & intSize & ") = '" & txtPickupPointDescription.text & "' " _
                            & "ORDER BY PickUpPointTime"
                        Set tmpRecordset = FindAndReturnRecords(strSQL)
                        If tmpRecordset.RecordCount > 0 Then
                            tmpTableData = DisplayIndex(tmpRecordset, 4, True, 4, 1, 2, 3, 4, "ID", "RouteID", "Περιγραφή", "Ωρα", 0, 0, 40, 7, 1, 0, 0, 1)
                            txtPickupPointID.text = tmpTableData.strFirstField
                            txtPickupPointDescription.text = tmpTableData.strSecondField
                            txtRouteID.text = tmpTableData.strCode
                            FindRoute
                        End If
                    End If
                End If
            End If
        Case 5
            'Οδηγός
            Set tmpRecordset = CheckForMatch("CommonDB", "Drivers", "DriverDescription", "String", txtDriverDescription.text)
            If tmpRecordset.RecordCount > 0 Then
                tmpTableData = DisplayIndex(tmpRecordset, 2, True, 2, 0, 1, "ID", "Περιγραφή", 0, 40, 1, 0)
                txtDriverID.text = tmpTableData.strCode
                txtDriverDescription.text = tmpTableData.strFirstField
            End If
        Case 5
            'Ημερομηνία (Αναζήτηση)
            FindRecordsAndPopulateGrid
    End Select

End Sub

Private Function FindRoute()

    Dim rsTable As Recordset
    
    Set rsTable = CommonDB.OpenRecordset("PickupRoutes")
    With rsTable
        .index = "PickupRouteID"
        .Seek "=", Val(txtRouteID.text)
        If Not .NoMatch Then
            txtRouteID.text = !PickupRouteID
            txtRouteShortDescription.text = !PickupRouteShortDescription
            lblRouteDescription.Caption = !PickupRouteDescription
            txtRouteShortDescription.Locked = True
        Else
            txtRouteID.text = ""
            txtRouteShortDescription.text = ""
            lblRouteDescription.Caption = ""
            txtRouteShortDescription.Locked = False
        End If
        .Close
    End With

End Function


Private Function FindAndReturnRecords(strSQL) As Recordset

   Dim tmpRecordset As Recordset
   
   Set tmpRecordset = CommonDB.OpenRecordset(strSQL, dbOpenSnapshot)
   Set FindAndReturnRecords = tmpRecordset

End Function


Private Sub Form_Activate()
    
    If Me.Tag = "True" Then
    
        Me.Tag = "False"
        
        AddColumnsToGrid grdCoachesReport, 44, GetSetting(strApplicationName, "Layout Strings", grdCoachesReport.Tag), _
            "05NCNTransferID,12NCDTransferDate,40NLNCustomerDescription,40NCNDestinationShortDescription,40NLNDestinationDescription,50NCNRouteShortDescription,50NLNRouteDescription,40NLNPickupPointHotelDescription,10NLNPickUpPointExactPoint,10NCTPickupPointTime,10NRITransferAdults,10NRITransferKids,10NRITransferFree,10NLNTransferRemarks,10NLNDriverDescription,10NRIXTransferTotal,04NCNTotalCriteria,04NCNSelected", _
            "TransferID,Ημερομηνία,Πελάτης,Π,Προορισμός,Δρομολόγιο,Δρομολόγιο,Σημείο παραλαβής,Ακριβές σημείο,Ωρα,Ε,Π,Δ,Παρατηρήσεις,Οδηγός,Σύνολο ατόμων,Κριτήρια,Ε"
        AddColumnsToGrid grdSummaryPerDestination, 24, GetSetting(strApplicationName, "Layout Strings", "grdCoachesReportSummaryPerDestination"), _
            "04NCNSelected,05NCNDestinationID,40NLNDestinationDescription,10NRITotalPersons", _
            "E,DestinationID,Προορισμός,Ατομα"
        AddColumnsToGrid grdSummaryPerCustomer, 24, GetSetting(strApplicationName, "Layout Strings", "grdCoachesReportSummaryPerCustomer"), _
            "04NCNSelected,05NCNCustomerID,40NLNCustomerDescription,10NRITotalPersons", _
            "E,CustomerID,Πελάτης,Ατομα"
        AddColumnsToGrid grdSummaryPerRoute, 24, GetSetting(strApplicationName, "Layout Strings", "grdCoachesReportSummaryPerRoute"), _
            "04NCNSelected,05NCNRouteID,40NLNRouteDescription,10NRITotalPersons", _
            "E,RouteID,Δρομολόγιο,Ατομα"
        AddColumnsToGrid grdSummaryPerDriver, 24, GetSetting(strApplicationName, "Layout Strings", "grdCoachesReportSummaryPerDriver"), _
            "04NCNSelected,05NCNDriverID,40NLNDriverDescription,10NRITotalPersons", _
            "E,DriverID,Οδηγός,Ατομα"
        Me.Refresh
        
        'mskDate.text = "21/05/2018"
        mskDate.SetFocus
        
    End If
    
    'AddDummyLines grdCoachesReport, "99999", "Α99/99/9999Α", "ΠΡΟΟΡΙΣΜΟΣ", "ΠΕΛΑΤΗΣ", "ΔΡΟΜΟΛΟΓΙΟ", "ΣΗΜΕΙΟ ΠΑΡΑΛΑΒΗΣ", "ΑΑΑΑΑΑΑΑΑΑΑΑΑΑΑΑΑΑΑΑΑΑΑΑΑΑΑΑΑΑΑΑΑΑΑΑΑΑΑΑΑΑΑΑΑΑΑΑΑΑ", "Α00:00Α", "999999", "999999", "999999", "999999", "Αριθμός αναφοράς", "Παρατηρήσεις", "Σημείο παραλαβής"
        
End Sub

Private Sub Form_KeyDown(KeyCode As Integer, Shift As Integer)
    
    CheckFunctionKeys KeyCode, Shift
    
    KeyCode = ResetKeyCode(KeyCode, Shift)

End Sub

Private Function CheckFunctionKeys(KeyCode, Shift)

    Dim CtrlDown
    
    CtrlDown = Shift + vbCtrlMask
    
    Select Case KeyCode
        Case vbKeyC And CtrlDown = 4 And cmdButton(0).Enabled
            cmdButton_Click 0
        Case vbKeyN And CtrlDown = 4 And cmdButton(1).Enabled
            cmdButton_Click 1
        Case vbKeyS And CtrlDown = 4 And cmdButton(2).Enabled
            cmdButton_Click 2
        Case vbKeyD And CtrlDown = 4 And cmdButton(3).Enabled
            cmdButton_Click 3
        Case vbKeyP And CtrlDown = 4 And cmdButton(4).Enabled
            cmdButton_Click 4
        Case vbKeyP And CtrlDown = 5 And cmdButton(5).Enabled
            cmdButton_Click 5
        Case vbKeyEscape
            If cmdButton(4).Enabled Then cmdButton_Click 4: Exit Function 'Ακύρωση επεξεργασίας
            If cmdButton(8).Enabled Then cmdButton_Click 8: Exit Function 'Νέα αναζήτηση (επιστροφή στην ημερομηνία)
            If cmdButton(9).Enabled Then cmdButton_Click 9: Exit Function 'Εξοδος
            If cmdButton(10).Enabled Then cmdButton_Click 10 'Κλείσιμο φόρμας Σύνδεσης δρομολογίου με οδηγό
        Case vbKeyM And CtrlDown = 4 And grdCoachesReport.Enabled
            grdCoachesReport.SetFocus
            grdCoachesReport.EnsureVisibleRow (grdCoachesReport.CurRow)
        Case vbKeyC And CtrlDown = 4 And grdSummaryPerCustomer.RowCount > 0
            grdSummaryPerCustomer.SetCurCell 1, 1
            grdSummaryPerCustomer.SetFocus
        Case vbKeyD And CtrlDown = 4 And grdSummaryPerDestination.RowCount > 0
            grdSummaryPerDestination.SetCurCell 1, 1
            grdSummaryPerDestination.SetFocus
        Case vbKeyR And CtrlDown = 4 And grdSummaryPerRoute.RowCount > 0
            grdSummaryPerRoute.SetCurCell 1, 1
            grdSummaryPerRoute.SetFocus
        Case vbKeyV And CtrlDown = 4 And grdSummaryPerDriver.RowCount > 0
            grdSummaryPerDriver.SetCurCell 1, 1
            grdSummaryPerDriver.SetFocus
        Case vbKeyF12 And CtrlDown = 4
            ToggleInfoPanel Me
    End Select

End Function

Private Sub Form_Load()

    blnCancel = True
    lngMinimumSeperatorTop = 7220
    lngMaximumSeperatorTop = 11585
    
    SetUpGrid lstIconList, grdCoachesReport, grdSummaryPerDestination, grdSummaryPerCustomer, grdSummaryPerRoute, grdSummaryPerDriver
    PositionControls Me, True, grdCoachesReport
    PositionSeperator
    RepositionMainGrid
    ColorizeControls Me, True
    RecolorizeControls
    PositionGrids
    
    frmCriteria(0).Visible = False
    
    ClearFields txtTransferID, txtCustomerID, txtPickupPointID, txtRouteID, txtDestinationID, txtDriverID
    ClearFields mskDate, txtCustomerDescription, txtDestinationDescription, txtPickupPointDescription, txtRouteShortDescription, lblRouteDescription, mskAdults, mskKids, mskFree, txtRemarks, txtDriverDescription
    ClearFields mskTotal
    ClearFields chkAllVisibleTransfers, chkAllVisibleTransfers, chkAllDestinations, chkAllCustomers, chkAllRoutes, chkAllDrivers
    ClearFields lblTotalPersons, lblTotalPersonsForVisibleRows, lblSelectedGridLines
    
    DisableFields txtCustomerDescription, txtDestinationDescription, txtPickupPointDescription, txtRouteShortDescription, mskAdults, mskKids, mskFree, txtRemarks, txtDriverDescription
    DisableFields cmdIndex(1), cmdIndex(2), cmdIndex(3), cmdIndex(4), cmdIndex(5)
    DisableFields chkAllVisibleTransfers, chkAllVisibleTransfers, chkAllDestinations, chkAllCustomers, chkAllRoutes, chkAllDrivers
    
    UpdateButtons Me, 11, 1, 0, 0, 0, 0, 0, 0, 0, 0, 1, 0, 0
    
End Sub

Private Function RepositionMainGrid()

    grdCoachesReport.Height = grdCoachesReport.Height - frmSummaries.Height - 150

End Function


Private Function SeekRecord(TransferID)
    
    Dim tmpRecordset As Recordset
    Dim tmpTableData As typTableData
    
    ClearFields txtDestinationID, txtCustomerID, txtRouteID, txtPickupPointID, txtDriverID
    ClearFields mskDate, txtDestinationDescription, txtCustomerDescription, txtRouteShortDescription, lblRouteDescription, txtPickupPointDescription, mskAdults, mskKids, mskFree, txtRemarks, txtDriverDescription
    ClearFields mskTotal
    DisableFields mskDate, txtCustomerDescription, txtDestinationDescription, txtPickupPointDescription, txtRouteShortDescription, mskAdults, mskKids, mskFree, txtRemarks
    DisableFields cmdIndex(0), cmdIndex(1), cmdIndex(2), cmdIndex(3), cmdIndex(4), cmdIndex(5)
    
    SeekRecord = False
    
    If MainSeekRecord("CommonDB", "Transfers", "ID", TransferID, True, txtTransferID, mskDate, txtDestinationID, txtCustomerID, txtRouteID, txtPickupPointID, mskAdults, mskKids, mskFree, txtRemarks, txtDriverID) Then
        'Προορισμός
        Set tmpRecordset = CheckForMatch("CommonDB", "Destinations", "DestinationID", "Numeric", txtDestinationID.text)
        txtDestinationID.text = tmpRecordset.Fields(0)
        txtDestinationDescription.text = tmpRecordset.Fields(2)
        'Πελάτης
        Set tmpRecordset = CheckForMatch("CommonDB", "Customers", "ID", "Numeric", txtCustomerID.text)
        txtCustomerID.text = tmpRecordset.Fields(0)
        txtCustomerDescription.text = tmpRecordset.Fields(1)
        'Δρομολόγιο
        Set tmpRecordset = CheckForMatch("CommonDB", "PickupRoutes", "PickupRouteID", "Numeric", txtRouteID.text)
        txtRouteID.text = tmpRecordset.Fields(0)
        txtRouteShortDescription.text = tmpRecordset.Fields(1)
        lblRouteDescription.Caption = tmpRecordset.Fields(2)
        'Σημείο παραλαβής
        Set tmpRecordset = CheckForMatch("CommonDB", "PickupPoints", "PickupPointID", "Numeric", txtPickupPointID.text)
        txtPickupPointID.text = tmpRecordset.Fields(0)
        txtPickupPointDescription.text = tmpRecordset.Fields(2)
        'Οδηγός (Αν έχω)
        If txtDriverID.text <> "" Then
            Set tmpRecordset = CheckForMatch("CommonDB", "Drivers", "DriverID", "Numeric", txtDriverID.text)
            txtDriverID.text = tmpRecordset.Fields(0)
            txtDriverDescription.text = tmpRecordset.Fields(1)
        End If
        'Τα υπόλοιπα
        EnableFields mskDate, txtDestinationDescription, txtCustomerDescription, txtRouteShortDescription, txtPickupPointDescription, mskAdults, mskKids, mskFree, txtRemarks, txtDriverDescription
        EnableFields cmdIndex(0), cmdIndex(1), cmdIndex(2), cmdIndex(3), cmdIndex(4), cmdIndex(5)
        mskTotal.Caption = AddNumbers(mskAdults.text, mskKids.text, mskFree.text)
        blnCancel = False
        blnStatus = False
        SeekRecord = True
        lngCurrentRow = grdCoachesReport.CurRow
        UpdateButtons Me, 11, 0, 0, 1, 1, 1, 0, 0, 0, 1, 0, 0, 0
        mskDate.SetFocus
    End If
    
End Function

Private Sub grdCoachesReport_ColHeaderMouseEnter(ByVal lCol As Long)

    grdCoachesReport.Header.Buttons = True

End Sub

Private Sub grdCoachesReport_ColHeaderMouseLeave(ByVal lCol As Long)

    grdCoachesReport.Header.Buttons = False
    
End Sub

Private Sub grdCoachesReport_DblClick(ByVal lRow As Long, ByVal lCol As Long, bRequestEdit As Boolean)

    On Error GoTo ErrTrap
    
    Dim TransferID As Long
    
    TransferID = grdCoachesReport.CellValue(lRow, "TransferID")
    
    SeekRecord TransferID
    
    Exit Sub
    
ErrTrap:
    Exit Sub
    
End Sub

Private Sub grdCoachesReport_HeaderRightClick(ByVal lCol As Long, ByVal Shift As Integer, ByVal X As Long, ByVal Y As Long)

    PopupMenu mnuHdrPopUp
    
End Sub

Private Sub grdCoachesReport_KeyDown(KeyCode As Integer, Shift As Integer, bDoDefault As Boolean)

    If KeyCode = vbKeySpace And grdCoachesReport.RowCount > 0 Then
        grdCoachesReport.CellIcon(grdCoachesReport.CurRow, "Selected") = lstIconList.ItemIndex(SelectRow(grdCoachesReport, 4, KeyCode, grdCoachesReport.CurRow, "TransferID"))
        lblSelectedGridLines.Caption = SumSelectedGridRows(grdCoachesReport, False, "Σύνολο επιλεγμένων: ", "TransferTotal", "integer")
    End If

End Sub

Private Sub grdCoachesReport_KeyPress(KeyAscii As Integer)

    On Error GoTo ErrTrap
    
    Dim TransferID As Long
    
    If KeyAscii = vbKeyReturn And grdCoachesReport.RowCount > 0 Then
        TransferID = grdCoachesReport.CellValue(grdCoachesReport.CurRow, "TransferID")
        SeekRecord TransferID
        Exit Sub
    End If
    
ErrTrap:
    Exit Sub
    
End Sub

Private Sub grdSummaryPerCustomer_DblClick(ByVal lRow As Long, ByVal lCol As Long, bRequestEdit As Boolean)

    grdCoachesReport.Redraw = False

    If grdSummaryPerCustomer.RowCount > 0 Then
        grdSummaryPerCustomer.CellIcon(grdSummaryPerCustomer.CurRow, "Selected") = lstIconList.ItemIndex(IIf(grdSummaryPerCustomer.CellIcon(grdSummaryPerCustomer.CurRow, "Selected") = 2, 1, 3))
        ScanGridsForSelectedRows
    End If
    
    grdCoachesReport.Redraw = True

End Sub

Private Sub grdSummaryPerCustomer_HeaderRightClick(ByVal lCol As Long, ByVal Shift As Integer, ByVal X As Long, ByVal Y As Long)

    PopupMenu mnuHdrPopUp
    
End Sub

Private Sub grdSummaryPerCustomer_KeyDown(KeyCode As Integer, Shift As Integer, bDoDefault As Boolean)

    grdCoachesReport.Redraw = False
    
    If KeyCode = vbKeySpace And grdSummaryPerCustomer.RowCount > 0 Then
        grdSummaryPerCustomer.CellIcon(grdSummaryPerCustomer.CurRow, "Selected") = lstIconList.ItemIndex(SelectRow(grdSummaryPerCustomer, 3, KeyCode, grdSummaryPerCustomer.CurRow, "CustomerID"))
        ScanGridsForSelectedRows
    End If
    
    grdCoachesReport.Redraw = True

End Sub


Private Sub grdSummaryPerDestination_DblClick(ByVal lRow As Long, ByVal lCol As Long, bRequestEdit As Boolean)

    grdCoachesReport.Redraw = False

    If grdSummaryPerDestination.RowCount > 0 Then
        grdSummaryPerDestination.CellIcon(grdSummaryPerDestination.CurRow, "Selected") = lstIconList.ItemIndex(IIf(grdSummaryPerDestination.CellIcon(grdSummaryPerDestination.CurRow, "Selected") = 1, 1, 2))
        ScanGridsForSelectedRows
    End If
    
    grdCoachesReport.Redraw = True

End Sub

Private Sub grdSummaryPerDestination_HeaderRightClick(ByVal lCol As Long, ByVal Shift As Integer, ByVal X As Long, ByVal Y As Long)

    PopupMenu mnuHdrPopUp
    
End Sub


Private Sub grdSummaryPerDestination_KeyDown(KeyCode As Integer, Shift As Integer, bDoDefault As Boolean)

    If KeyCode = vbKeySpace And grdSummaryPerDestination.RowCount > 0 Then
        grdSummaryPerDestination.CellIcon(grdSummaryPerDestination.CurRow, "Selected") = lstIconList.ItemIndex(SelectRow(grdSummaryPerDestination, 2, KeyCode, grdSummaryPerDestination.CurRow, "DestinationID"))
        ScanGridsForSelectedRows
    End If

End Sub


Private Sub grdSummaryPerDriver_DblClick(ByVal lRow As Long, ByVal lCol As Long, bRequestEdit As Boolean)

    grdCoachesReport.Redraw = False

    If grdSummaryPerDriver.RowCount > 0 Then
        grdSummaryPerDriver.CellIcon(grdSummaryPerDriver.CurRow, "Selected") = lstIconList.ItemIndex(IIf(grdSummaryPerDriver.CellIcon(grdSummaryPerDriver.CurRow, "Selected") = 4, 1, 5))
        ScanGridsForSelectedRows
    End If
    
    grdCoachesReport.Redraw = True

End Sub

Private Sub grdSummaryPerDriver_HeaderRightClick(ByVal lCol As Long, ByVal Shift As Integer, ByVal X As Long, ByVal Y As Long)

    PopupMenu mnuHdrPopUp
    
End Sub


Private Sub grdSummaryPerDriver_KeyDown(KeyCode As Integer, Shift As Integer, bDoDefault As Boolean)

    grdCoachesReport.Redraw = False
    
    If KeyCode = vbKeySpace And grdSummaryPerDriver.RowCount > 0 Then
        grdSummaryPerDriver.CellIcon(grdSummaryPerDriver.CurRow, "Selected") = lstIconList.ItemIndex(SelectRow(grdSummaryPerDriver, 5, KeyCode, grdSummaryPerDriver.CurRow, "DriverID"))
        ScanGridsForSelectedRows
    End If

    grdCoachesReport.Redraw = True

End Sub


Private Sub grdSummaryPerRoute_DblClick(ByVal lRow As Long, ByVal lCol As Long, bRequestEdit As Boolean)

    grdCoachesReport.Redraw = False

    If grdSummaryPerRoute.RowCount > 0 Then
        grdSummaryPerRoute.CellIcon(grdSummaryPerRoute.CurRow, "Selected") = lstIconList.ItemIndex(IIf(grdSummaryPerRoute.CellIcon(grdSummaryPerRoute.CurRow, "Selected") = 3, 1, 4))
        ScanGridsForSelectedRows
    End If
    
    grdCoachesReport.Redraw = True

End Sub

Private Sub grdSummaryPerRoute_HeaderRightClick(ByVal lCol As Long, ByVal Shift As Integer, ByVal X As Long, ByVal Y As Long)

    PopupMenu mnuHdrPopUp
    
End Sub


Private Sub grdSummaryPerRoute_KeyDown(KeyCode As Integer, Shift As Integer, bDoDefault As Boolean)

    grdCoachesReport.Redraw = False
    
    If KeyCode = vbKeySpace And grdSummaryPerRoute.RowCount > 0 Then
        grdSummaryPerRoute.CellIcon(grdSummaryPerRoute.CurRow, "Selected") = lstIconList.ItemIndex(SelectRow(grdSummaryPerRoute, 4, KeyCode, grdSummaryPerRoute.CurRow, "RouteID"))
        ScanGridsForSelectedRows
    End If

    grdCoachesReport.Redraw = True

End Sub


Private Sub mnuΑποθήκευσηΠλάτουςΣτηλών_Click()

    SaveSetting strApplicationName, "Layout Strings", grdCoachesReport.Tag, grdCoachesReport.LayoutCol
    
    SaveSetting strApplicationName, "Layout Strings", "grdCoachesReportSummaryPerDestination", grdSummaryPerDestination.LayoutCol
    SaveSetting strApplicationName, "Layout Strings", "grdCoachesReportSummaryPerCustomer", grdSummaryPerCustomer.LayoutCol
    SaveSetting strApplicationName, "Layout Strings", "grdCoachesReportSummaryPerRoute", grdSummaryPerRoute.LayoutCol
    SaveSetting strApplicationName, "Layout Strings", "grdCoachesReportSummaryPerDriver", grdSummaryPerDriver.LayoutCol

End Sub

Private Function CreateUnicodeFile(strReportTitle, strReportSubTitle1, strReportSubTitle2, intReportDetailLines)

    On Error GoTo ErrTrap
    
    'Εκτυπωτής
    Dim lngRow As Long
    Dim intProcessedDetailLines As Integer
    Dim intPageNo As Integer
    
    'PickupPoint
    Dim lngPickupPointAdults As Integer
    Dim lngPickupPointKids As Integer
    Dim lngPickupPointFree As Integer
    Dim lngPickupPointPersons As Integer
    'Total
    Dim lngTotalAdults As Integer
    Dim lngTotalKids As Integer
    Dim lngTotalFree As Integer
    Dim lngTotalPersons As Integer
    'PickupPoint
    Dim intPickupPointCount As Integer
    Dim strPickupPoint As String

    'Αρχικές τιμές
    intPageNo = 1
    lngTotalAdults = 0
    lngTotalKids = 0
    lngTotalFree = 0
    lngTotalPersons = 0
    
    Open strUnicodeFile For Output As #1

    'Επικεφαλίδες
    PrintHeadings 124, intPageNo, strReportTitle, strReportSubTitle1, strReportSubTitle2
    PrintColumnHeadings 1, "ΩΡΑ", 7, "ΣΗΜΕΙΟ ΠΑΡΑΛΑΒΗΣ", 39, "Ε", 42, "Π", 45, "Δ", 49, "Σ", 51, "ΠΕΛΑΤΗΣ", 72, "ΠΑΡΑΤΗΡΗΣΕΙΣ", 123, "Π"
    Print #1, ""
    
    'Εγγραφές
    intProcessedDetailLines = 10
    intPickupPointCount = 0
    
    With grdCoachesReport
        For lngRow = 1 To .RowCount
            'Αν η γραμμή είναι ορατή
            If .RowVisible(lngRow) Then
                'Αν το σημείο παραλαβής που βρίσκομαι είναι ίδιο με αυτό που έχω ήδη κρατήσει ή βρίσκομαι στην πρώτη εγγραφή
                If .CellValue(lngRow, "PickupPointHotelDescription") = strPickupPoint Or strPickupPoint = "" Then
                    'Ενημερώνω τη μεταβλητή που κρατάει το σημείο παραλαβής
                    strPickupPoint = .CellValue(lngRow, "PickupPointHotelDescription")
                    'Αυξάνω τα σημεία παραλαβής κατά ένα
                    intPickupPointCount = intPickupPointCount + 1
                Else
                    'Αν το σημείο παραλαβής που βρίσκομαι δεν είναι ίδιο με αυτό που έχω κρατήσει
                    'Αυτό συμβαίνει όταν έχω αλλαγή σημείου παραλαβής
                    'Αν έχω μετρήσει παραπάνω από ένα σημεία παραλαβής
                    If intPickupPointCount > 1 Then
                        'Τυπώνω τα σύνολα του σημείου παραλαβής
                        Print #1, _
                            Tab(12); "ΣΥΝΟΛΑ " & Left(strPickupPoint, 18); _
                            Tab(40 - Len(format(lngPickupPointAdults, "#,##0"))); IIf(lngPickupPointAdults > 0, format(lngPickupPointAdults, "#,##0"), ""); _
                            Tab(43 - Len(format(lngPickupPointKids, "#,##0"))); IIf(lngPickupPointKids > 0, format(lngPickupPointKids, "#,##0"), ""); _
                            Tab(46 - Len(format(lngPickupPointFree, "#,##0"))); IIf(lngPickupPointFree > 0, format(lngPickupPointFree, "#,##0"), ""); _
                            Tab(50 - Len(format(lngPickupPointPersons, "#,##0"))); IIf(lngPickupPointPersons > 0, format(lngPickupPointPersons, "#,##0"), "")
                        'Εκτυπωμένες γραμμές
                        intProcessedDetailLines = intProcessedDetailLines + 1
                        'Ελέγχω για αλλαγή σελίδας
                        GoSub CheckToEject
                        'Δίνω αρχική τιμή στα σημεία παραλαβής
                        intPickupPointCount = 1
                    End If
                    'Ενημερώνω τη μεταβλητή που κρατάει το σημείο παραλαβής
                    strPickupPoint = .CellValue(lngRow, "PickupPointHotelDescription")
                    'Μηδενίζω τα σύνολα του σημείου παραλαβής
                    lngPickupPointAdults = 0
                    lngPickupPointKids = 0
                    lngPickupPointFree = 0
                    lngPickupPointPersons = 0
                End If
            
                'Τυπώνω το σημείο παραλαβής που βρίσκομαι
                Print #1, _
                    Tab(1); .CellText(lngRow, "PickupPointTime"); _
                    Tab(7); Left(.CellText(lngRow, "PickupPointHotelDescription"), 20); _
                    Tab(40 - Len((format(.CellText(lngRow, "TransferAdults"), "#,##0")))); format(.CellText(lngRow, "TransferAdults"), "#,##0"); _
                    Tab(43 - Len((format(.CellText(lngRow, "TransferKids"), "#,##0")))); format(.CellText(lngRow, "TransferKids"), "#,##0"); _
                    Tab(46 - Len((format(.CellText(lngRow, "TransferFree"), "#,##0")))); format(.CellText(lngRow, "TransferFree"), "#,##0"); _
                    Tab(50 - Len((format(.CellText(lngRow, "TransferTotal"), "#,##0")))); format(.CellText(lngRow, "TransferTotal"), "#,##0"); _
                    Tab(51); Left(.CellText(lngRow, "CustomerDescription"), 20); _
                    Tab(72); Left(.CellText(lngRow, "TransferRemarks"), 50); _
                    Tab(123); Left(.CellText(lngRow, "DestinationShortDescription"), 2)
                
                'Σύνολα σημείου παραλαβής
                lngPickupPointAdults = lngPickupPointAdults + IIf(.CellValue(lngRow, "TransferAdults") <> "", .CellValue(lngRow, "TransferAdults"), 0)
                lngPickupPointKids = lngPickupPointKids + IIf(.CellValue(lngRow, "TransferKids") <> "", .CellValue(lngRow, "TransferKids"), 0)
                lngPickupPointFree = lngPickupPointFree + IIf(.CellValue(lngRow, "TransferFree") <> "", .CellValue(lngRow, "TransferFree"), 0)
                lngPickupPointPersons = lngPickupPointAdults + lngPickupPointKids + lngPickupPointFree

                'Σύνολα οδηγού
                lngTotalAdults = lngTotalAdults + IIf(.CellValue(lngRow, "TransferAdults") <> "", .CellValue(lngRow, "TransferAdults"), 0)
                lngTotalKids = lngTotalKids + IIf(.CellValue(lngRow, "TransferKids") <> "", .CellValue(lngRow, "TransferKids"), 0)
                lngTotalFree = lngTotalFree + IIf(.CellValue(lngRow, "TransferFree") <> "", .CellValue(lngRow, "TransferFree"), 0)
                lngTotalPersons = lngTotalAdults + lngTotalKids + lngTotalFree

                intProcessedDetailLines = intProcessedDetailLines + 1
                
                GoSub CheckToEject

            End If
            
        Next lngRow
        
    End With
        
    'Αν έχω μετρήσει παραπάνω από ένα σημεία παραλαβής
    If intPickupPointCount > 1 Then
        'Τυπώνω τα σύνολα του σημείου παραλαβής
        Print #1, _
            Tab(12); "ΣΥΝΟΛΑ " & Left(strPickupPoint, 18); _
            Tab(40 - Len(format(lngPickupPointAdults, "#,##0"))); IIf(lngPickupPointAdults > 0, format(lngPickupPointAdults, "#,##0"), ""); _
            Tab(43 - Len(format(lngPickupPointKids, "#,##0"))); IIf(lngPickupPointKids > 0, format(lngPickupPointKids, "#,##0"), ""); _
            Tab(46 - Len(format(lngPickupPointFree, "#,##0"))); IIf(lngPickupPointFree > 0, format(lngPickupPointFree, "#,##0"), ""); _
            Tab(50 - Len(format(lngPickupPointPersons, "#,##0"))); IIf(lngPickupPointPersons > 0, format(lngPickupPointPersons, "#,##0"), "")
    End If
    
    'Τυπώνω τα σύνολα του οδηγού
    Print #1, ""
    Print #1, _
        Tab(40 - Len(format(lngTotalAdults, "#,##0"))); format(lngTotalAdults, "#,##0"); _
        Tab(43 - Len(format(lngTotalKids, "#,##0"))); format(lngTotalKids, "#,##0"); _
        Tab(46 - Len(format(lngTotalFree, "#,##0"))); format(lngTotalFree, "#,##0"); _
        Tab(50 - Len(format(lngTotalPersons, "#,##0"))); format(lngTotalPersons, "#,##0")
        
    Print #1, Tab(7); "ΤΕΛΟΣ ΕΚΤΥΠΩΣΗΣ"
    
    Close #1

    CreateUnicodeFile = True

    Exit Function
    
ErrTrap:
    CreateUnicodeFile = False
    DisplayErrorMessage True, Err.Description
    
    Return
    
CheckToEject:
    If intProcessedDetailLines > intReportDetailLines Then
        Print #1, ""
        Print #1, Tab(7); "Η ΕΚΤΥΠΩΣΗ ΣΥΝΕΧΙΖΕΤΑΙ..."
        intPageNo = intPageNo + 1
        PrintHeadings 124, intPageNo, strReportTitle, strReportSubTitle1, strReportSubTitle2
        PrintColumnHeadings 1, "ΩΡΑ", 7, "ΣΗΜΕΙΟ ΠΑΡΑΛΑΒΗΣ", 39, "Ε", 42, "Π", 45, "Δ", 49, "Σ", 51, "ΠΕΛΑΤΗΣ", 72, "ΠΑΡΑΤΗΡΗΣΕΙΣ", 123, "Π"
        Print #1, ""
        Print #1, Tab(7); "ΣΥΝΕΧΕΙΑ ΕΚΤΥΠΩΣΗΣ ΑΠΟ ΠΡΟΗΓΟΥΜΕΝΗ ΣΕΛΙΔΑ..."
        Print #1, ""
        intProcessedDetailLines = 12
    End If
    
    Return

End Function

Private Sub mskAdults_Validate(Cancel As Boolean)
    
    If Not blnCancel Then
        mskTotal.Caption = AddNumbers(mskAdults.text, mskKids.text, mskFree.text)
    End If

End Sub


Private Sub mskDate_KeyDown(KeyCode As Integer, Shift As Integer)

    If KeyCode = 13 And cmdButton(0).Enabled Then
        cmdButton(0).SetFocus
    End If

End Sub


Private Sub mskFree_Validate(Cancel As Boolean)

    If Not blnCancel Then
        mskTotal.Caption = AddNumbers(mskAdults.text, mskKids.text, mskFree.text)
    End If

End Sub


Private Sub mskKids_Validate(Cancel As Boolean)

    If Not blnCancel Then
        mskTotal.Caption = AddNumbers(mskAdults.text, mskKids.text, mskFree.text)
    End If

End Sub


Private Sub Seperator_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)

    If Button = vbLeftButton Then
        lngOldSeperatorTop = Y
        blnIsMoving = True
    End If

End Sub


Private Sub Seperator_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)

    Dim lngNewTop As Long
    Dim IsMaximumReached As Boolean
    
    lngNewTop = Seperator.Top - (lngOldSeperatorTop - Y)
    
    If blnIsMoving Then
        
        'Check for maximums
        If lngNewTop <= lngMaximumSeperatorTop Then
            Seperator.Top = lngNewTop
        Else
            Seperator.Top = lngMaximumSeperatorTop
            Exit Sub
        End If
        
        'Check for minimums
        If lngNewTop >= lngMinimumSeperatorTop Then
            Seperator.Top = lngNewTop
        Else
            Seperator.Top = lngMinimumSeperatorTop
        End If
        
        PositionGrids
        
    End If

End Sub

Private Function PositionGrids()

    grdCoachesReport.Height = Seperator.Top - grdCoachesReport.Top - 255
    
    frmSummaries.Top = Seperator.Top + 150
    frmSummaries.Height = shpBackground.Height - frmSummaries.Top + shpBackground.Top
    
    grdSummaryPerDestination.Height = frmSummaries.Height - grdSummaryPerDestination.Top - 5
    grdSummaryPerCustomer.Height = frmSummaries.Height - grdSummaryPerCustomer.Top - 5
    grdSummaryPerRoute.Height = frmSummaries.Height - grdSummaryPerRoute.Top - 5
    grdSummaryPerDriver.Height = frmSummaries.Height - grdSummaryPerDriver.Top - 5
    
End Function

Private Sub Seperator_MouseUp(Button As Integer, Shift As Integer, X As Single, Y As Single)

    blnIsMoving = False
    
    SaveSetting strApplicationName, "Settings", "SeperatorTop", Seperator.Top

End Sub


Private Sub txtCustomerDescription_Change()

    If txtCustomerDescription.text = "" Then ClearFields txtCustomerID
    
End Sub

Private Sub txtCustomerDescription_KeyDown(KeyCode As Integer, Shift As Integer)

    If KeyCode = vbKeyF2 Then cmdIndex_Click 2
    
End Sub


Private Sub txtCustomerDescription_Validate(Cancel As Boolean)

    If txtCustomerID.text = "" And txtCustomerDescription.text <> "" Then cmdIndex_Click 2
    
End Sub

Private Sub txtDestinationDescription_Change()

    If txtDestinationDescription.text = "" Then ClearFields txtDestinationID

End Sub


Private Sub txtDestinationDescription_KeyDown(KeyCode As Integer, Shift As Integer)

    If KeyCode = vbKeyF2 Then cmdIndex_Click 1

End Sub


Private Sub txtDestinationDescription_Validate(Cancel As Boolean)

    If txtDestinationID.text = "" And txtDestinationDescription.text <> "" Then cmdIndex_Click 1
    
End Sub


Private Sub txtDriverDescription_Change()

    If txtDriverDescription.text = "" Then ClearFields txtDriverID

End Sub

Private Sub txtDriverDescription_KeyDown(KeyCode As Integer, Shift As Integer)

    If KeyCode = vbKeyF2 Then cmdIndex_Click 5

End Sub


Private Sub txtDriverDescription_Validate(Cancel As Boolean)

    If txtDriverID.text = "" And txtDriverDescription.text <> "" Then cmdIndex_Click 5

End Sub

Private Sub txtDriverDescriptionForRoutes_Change()

    If txtDriverDescriptionForRoutes.text = "" Then
        ClearFields txtDriverIDForRoutes
    End If

End Sub


Private Sub txtDriverDescriptionForRoutes_KeyDown(KeyCode As Integer, Shift As Integer)

    If KeyCode = vbKeyF2 Then cmdIndex_Click 0
    
End Sub


Private Sub txtDriverDescriptionForRoutes_Validate(Cancel As Boolean)

    If txtDriverIDForRoutes.text = "" And txtDriverDescriptionForRoutes.text <> "" Then cmdIndex_Click 0

End Sub


Private Sub txtPickupPointDescription_Change()

    If txtPickupPointDescription.text = "" Then
        ClearFields txtRouteID, txtPickupPointID, txtRouteShortDescription, lblRouteDescription
        txtRouteShortDescription.Locked = False
    End If

End Sub

Private Sub txtPickupPointDescription_KeyDown(KeyCode As Integer, Shift As Integer)

    If KeyCode = vbKeyF2 Then cmdIndex_Click 4

End Sub


Private Sub txtPickupPointDescription_Validate(Cancel As Boolean)

    If txtPickupPointID.text = "" And txtPickupPointDescription.text <> "" Then cmdIndex_Click 4

End Sub


Private Sub txtRouteShortDescription_Change()

    If txtRouteShortDescription.text = "" Then
        ClearFields txtRouteID, lblRouteDescription, txtPickupPointID, txtPickupPointDescription
    End If
    
End Sub


Private Sub txtRouteShortDescription_KeyDown(KeyCode As Integer, Shift As Integer)

    If KeyCode = vbKeyF2 Then cmdIndex_Click 3

End Sub


Private Sub txtRouteShortDescription_Validate(Cancel As Boolean)

    If txtRouteID.text = "" And txtRouteShortDescription.text <> "" Then cmdIndex_Click 3

End Sub



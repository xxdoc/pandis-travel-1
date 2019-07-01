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
   ClientHeight    =   10590
   ClientLeft      =   -30
   ClientTop       =   15
   ClientWidth     =   18690
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
   ScaleHeight     =   10590
   ScaleWidth      =   18690
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
      Left            =   8025
      TabIndex        =   45
      Top             =   3825
      Width           =   4065
      Begin vbalProgBarLib6.vbalProgressBar prgProgressBar 
         Height          =   615
         Left            =   150
         TabIndex        =   46
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
         TabIndex        =   47
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
      TabIndex        =   9
      Top             =   75
      Width           =   18540
      Begin VB.Frame frmCriteria 
         BackColor       =   &H0080C0FF&
         BorderStyle     =   0  'None
         Height          =   2190
         Index           =   0
         Left            =   7950
         TabIndex        =   61
         Top             =   4950
         Width           =   7290
         Begin UserControls.newText txtDriverDescriptionForRoutes 
            Height          =   465
            Left            =   1425
            TabIndex        =   62
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
            TabIndex        =   63
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
            TabIndex        =   69
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
            ForeColor       =   0
            PicOpacity      =   0
         End
         Begin Dacara_dcButton.dcButton cmdButton 
            Height          =   465
            Index           =   11
            Left            =   3750
            TabIndex        =   70
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
            ForeColor       =   0
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
            TabIndex        =   68
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
            TabIndex        =   66
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
            TabIndex        =   65
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
            BackStyle       =   0  'Transparent
            Caption         =   "Οδηγός"
            ForeColor       =   &H00000000&
            Height          =   315
            Index           =   11
            Left            =   450
            TabIndex        =   64
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
            TabIndex        =   67
            Top             =   0
            Width           =   7440
         End
      End
      Begin VB.CheckBox chkAllTransfers 
         Appearance      =   0  'Flat
         BackColor       =   &H00000000&
         Caption         =   "Επιλογή όλων"
         ForeColor       =   &H00FFFFFF&
         Height          =   315
         Left            =   7875
         TabIndex        =   60
         TabStop         =   0   'False
         Top             =   1050
         Width           =   2340
      End
      Begin VB.PictureBox Seperator 
         BackColor       =   &H80000003&
         BorderStyle     =   0  'None
         Height          =   50
         Left            =   300
         MousePointer    =   7  'Size N S
         ScaleHeight     =   45
         ScaleWidth      =   5865
         TabIndex        =   59
         TabStop         =   0   'False
         Top             =   6825
         Width           =   5865
      End
      Begin VB.Frame frmSummaries 
         BackColor       =   &H00000000&
         BorderStyle     =   0  'None
         Height          =   2115
         Left            =   225
         TabIndex        =   48
         Top             =   7500
         Width           =   18090
         Begin VB.CheckBox chkAllDrivers 
            Appearance      =   0  'Flat
            BackColor       =   &H80000005&
            Caption         =   "Επιλογή όλων"
            ForeColor       =   &H80000008&
            Height          =   315
            Left            =   13725
            TabIndex        =   52
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
            Left            =   9150
            TabIndex        =   51
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
            Left            =   4575
            TabIndex        =   50
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
            TabIndex        =   49
            TabStop         =   0   'False
            Top             =   0
            Width           =   2340
         End
         Begin iGrid300_10Tec.iGrid grdSummaryPerRoute 
            Height          =   1725
            Left            =   9150
            TabIndex        =   53
            TabStop         =   0   'False
            Top             =   375
            Width           =   4500
            _ExtentX        =   7938
            _ExtentY        =   3043
            Appearance      =   0
            ForeColor       =   -2147483631
         End
         Begin iGrid300_10Tec.iGrid grdSummaryPerDriver 
            Height          =   1725
            Left            =   13725
            TabIndex        =   54
            TabStop         =   0   'False
            Top             =   375
            Width           =   4350
            _ExtentX        =   7673
            _ExtentY        =   3043
            Appearance      =   0
            ForeColor       =   -2147483631
         End
         Begin iGrid300_10Tec.iGrid grdSummaryPerCustomer 
            Height          =   1725
            Left            =   4575
            TabIndex        =   55
            TabStop         =   0   'False
            Top             =   375
            Width           =   4500
            _ExtentX        =   7938
            _ExtentY        =   3043
            Appearance      =   0
            ForeColor       =   -2147483631
         End
         Begin iGrid300_10Tec.iGrid grdSummaryPerDestination 
            Height          =   1725
            Left            =   0
            TabIndex        =   56
            TabStop         =   0   'False
            Top             =   375
            Width           =   4500
            _ExtentX        =   7938
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
         Left            =   13725
         TabIndex        =   29
         Top             =   1575
         Width           =   4515
         Begin VB.TextBox Text5 
            Appearance      =   0  'Flat
            BackColor       =   &H0000FF00&
            BorderStyle     =   0  'None
            ForeColor       =   &H00000000&
            Height          =   315
            Left            =   75
            TabIndex        =   72
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
            TabIndex        =   71
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
            TabIndex        =   44
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
            TabIndex        =   43
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
            TabIndex        =   41
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
            TabIndex        =   40
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
            TabIndex        =   39
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
            TabIndex        =   38
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
            TabIndex        =   37
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
            TabIndex        =   36
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
            TabIndex        =   35
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
            TabIndex        =   34
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
            TabIndex        =   33
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
            TabIndex        =   32
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
            TabIndex        =   31
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
            TabIndex        =   30
            TabStop         =   0   'False
            Top             =   2325
            Width           =   780
         End
         Begin vbalIml6.vbalImageList lstIconList 
            Left            =   75
            Top             =   3075
            _ExtentX        =   953
            _ExtentY        =   953
            ColourDepth     =   8
            Size            =   6888
            Images          =   "Transfers.frx":05D2
            Version         =   131072
            KeyCount        =   6
            Keys            =   "ÿÿÿÿÿ"
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
         TabIndex        =   12
         Top             =   9750
         Width           =   16440
         Begin Dacara_dcButton.dcButton cmdButton 
            Height          =   465
            Index           =   9
            Left            =   13125
            TabIndex        =   13
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
            TabIndex        =   14
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
            Index           =   8
            Left            =   9900
            TabIndex        =   42
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
            ForeColor       =   0
            PicOpacity      =   0
         End
         Begin Dacara_dcButton.dcButton cmdButton 
            Height          =   465
            Index           =   5
            Left            =   225
            TabIndex        =   73
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
         Begin Dacara_dcButton.dcButton cmdButton 
            Height          =   465
            Index           =   7
            Left            =   6675
            TabIndex        =   81
            TabStop         =   0   'False
            Top             =   75
            Width           =   3165
            _ExtentX        =   5583
            _ExtentY        =   820
            BackColor       =   12640511
            ButtonShape     =   3
            ButtonStyle     =   4
            Caption         =   "Δημιουργία αρχείου PDF"
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
         Height          =   5790
         Left            =   7875
         TabIndex        =   10
         TabStop         =   0   'False
         Top             =   1425
         Width           =   10440
         _ExtentX        =   18415
         _ExtentY        =   10213
         Appearance      =   0
         ForeColor       =   -2147483631
      End
      Begin UserControls.newDate mskDate 
         Height          =   465
         Left            =   2025
         TabIndex        =   0
         Top             =   1125
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
         Left            =   2025
         TabIndex        =   2
         Top             =   2175
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
         Left            =   2025
         TabIndex        =   1
         Top             =   1650
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
         Left            =   2025
         TabIndex        =   3
         Top             =   2700
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
         Left            =   2025
         TabIndex        =   7
         Top             =   4800
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
      Begin UserControls.newInteger mskAdults 
         Height          =   465
         Left            =   2025
         TabIndex        =   4
         Top             =   3225
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
         Left            =   2025
         TabIndex        =   5
         Top             =   3750
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
         Left            =   2025
         TabIndex        =   6
         Top             =   4275
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
         Left            =   7050
         TabIndex        =   16
         TabStop         =   0   'False
         Top             =   2175
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
         PicNormal       =   "Transfers.frx":20DA
         PicSizeH        =   16
         PicSizeW        =   16
      End
      Begin Dacara_dcButton.dcButton cmdIndex 
         Height          =   465
         Index           =   3
         Left            =   7050
         TabIndex        =   17
         TabStop         =   0   'False
         Top             =   2700
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
         PicNormal       =   "Transfers.frx":2674
         PicSizeH        =   16
         PicSizeW        =   16
      End
      Begin UserControls.newText txtDriverDescription 
         Height          =   465
         Left            =   2025
         TabIndex        =   8
         Top             =   5325
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
         Index           =   4
         Left            =   7050
         TabIndex        =   27
         TabStop         =   0   'False
         Top             =   5325
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
         PicNormal       =   "Transfers.frx":2C0E
         PicSizeH        =   16
         PicSizeW        =   16
      End
      Begin Dacara_dcButton.dcButton cmdIndex 
         Height          =   465
         Index           =   1
         Left            =   7050
         TabIndex        =   57
         TabStop         =   0   'False
         Top             =   1650
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
         PicNormal       =   "Transfers.frx":31A8
         PicSizeH        =   16
         PicSizeW        =   16
      End
      Begin Dacara_dcButton.dcButton cmdButton 
         Height          =   465
         Index           =   0
         Left            =   3600
         TabIndex        =   74
         Top             =   1125
         Width           =   390
         _ExtentX        =   688
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
         PicNormal       =   "Transfers.frx":3742
         PicSizeH        =   16
         PicSizeW        =   16
      End
      Begin Dacara_dcButton.dcButton cmdButton 
         Height          =   465
         Index           =   1
         Left            =   1050
         TabIndex        =   76
         TabStop         =   0   'False
         Top             =   6075
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
         ForeColor       =   0
         PicOpacity      =   0
      End
      Begin Dacara_dcButton.dcButton cmdButton 
         Height          =   465
         Index           =   3
         Left            =   3900
         TabIndex        =   77
         TabStop         =   0   'False
         Top             =   6075
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
         ForeColor       =   0
         PicOpacity      =   0
      End
      Begin Dacara_dcButton.dcButton cmdButton 
         Height          =   465
         Index           =   4
         Left            =   5325
         TabIndex        =   78
         TabStop         =   0   'False
         Top             =   6075
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
         ForeColor       =   0
         PicOpacity      =   0
      End
      Begin Dacara_dcButton.dcButton cmdButton 
         Height          =   465
         Index           =   2
         Left            =   2475
         TabIndex        =   79
         TabStop         =   0   'False
         Top             =   6075
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
         ForeColor       =   0
         PicOpacity      =   0
      End
      Begin VB.Shape shpWedge 
         BackColor       =   &H0000FFFF&
         BackStyle       =   1  'Opaque
         BorderStyle     =   0  'Transparent
         FillColor       =   &H00008000&
         Height          =   840
         Index           =   10
         Left            =   18300
         Top             =   7875
         Visible         =   0   'False
         Width           =   240
      End
      Begin VB.Shape shpWedge 
         BackColor       =   &H0000FFFF&
         BackStyle       =   1  'Opaque
         BorderStyle     =   0  'Transparent
         FillColor       =   &H00008000&
         Height          =   840
         Index           =   9
         Left            =   0
         Top             =   7800
         Visible         =   0   'False
         Width           =   240
      End
      Begin VB.Shape shpWedge 
         BackColor       =   &H0000FFFF&
         BackStyle       =   1  'Opaque
         BorderStyle     =   0  'Transparent
         FillColor       =   &H00008000&
         Height          =   840
         Index           =   8
         Left            =   18300
         Top             =   1800
         Visible         =   0   'False
         Width           =   240
      End
      Begin VB.Shape shpWedge 
         BackColor       =   &H0000FFFF&
         BackStyle       =   1  'Opaque
         BorderStyle     =   0  'Transparent
         FillColor       =   &H00008000&
         Height          =   840
         Index           =   7
         Left            =   0
         Top             =   1125
         Visible         =   0   'False
         Width           =   240
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
         Left            =   3900
         TabIndex        =   80
         Top             =   3810
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
         Left            =   2850
         TabIndex        =   75
         Top             =   3375
         Width           =   255
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
         Left            =   13500
         TabIndex        =   58
         Top             =   600
         Width           =   5040
      End
      Begin VB.Label lblLabel 
         BackColor       =   &H000080FF&
         Caption         =   "Οδηγός"
         ForeColor       =   &H00FFFFFF&
         Height          =   315
         Index           =   4
         Left            =   225
         TabIndex        =   28
         Top             =   5400
         Width           =   1365
      End
      Begin VB.Label lblLabel 
         AutoSize        =   -1  'True
         BackColor       =   &H000080FF&
         Caption         =   "Ημερομηνία"
         ForeColor       =   &H00FFFFFF&
         Height          =   315
         Index           =   3
         Left            =   225
         TabIndex        =   26
         Top             =   1200
         Width           =   1365
      End
      Begin VB.Label lblLabel 
         AutoSize        =   -1  'True
         BackColor       =   &H000080FF&
         Caption         =   "Πελάτης"
         ForeColor       =   &H00FFFFFF&
         Height          =   315
         Index           =   5
         Left            =   225
         TabIndex        =   25
         Top             =   2250
         Width           =   1365
      End
      Begin VB.Label lblLabel 
         AutoSize        =   -1  'True
         BackColor       =   &H000080FF&
         Caption         =   "Προορισμός"
         ForeColor       =   &H00FFFFFF&
         Height          =   315
         Index           =   7
         Left            =   225
         TabIndex        =   24
         Top             =   1725
         Width           =   1365
      End
      Begin VB.Label lblLabel 
         AutoSize        =   -1  'True
         BackColor       =   &H000080FF&
         Caption         =   "Σημείο παραλαβής"
         ForeColor       =   &H00FFFFFF&
         Height          =   315
         Index           =   8
         Left            =   225
         TabIndex        =   23
         Top             =   2775
         Width           =   1365
      End
      Begin VB.Label lblLabel 
         AutoSize        =   -1  'True
         BackColor       =   &H000080FF&
         Caption         =   "Ενήλικες"
         ForeColor       =   &H00FFFFFF&
         Height          =   315
         Index           =   9
         Left            =   225
         TabIndex        =   22
         Top             =   3300
         Width           =   1365
      End
      Begin VB.Label lblLabel 
         AutoSize        =   -1  'True
         BackColor       =   &H000080FF&
         Caption         =   "Παιδιά"
         ForeColor       =   &H00FFFFFF&
         Height          =   315
         Index           =   6
         Left            =   225
         TabIndex        =   21
         Top             =   3825
         Width           =   1365
      End
      Begin VB.Label lblLabel 
         AutoSize        =   -1  'True
         BackColor       =   &H000080FF&
         Caption         =   "Δωρεάν"
         ForeColor       =   &H00FFFFFF&
         Height          =   315
         Index           =   2
         Left            =   225
         TabIndex        =   20
         Top             =   4350
         Width           =   1365
      End
      Begin VB.Label lblLabel 
         AutoSize        =   -1  'True
         BackColor       =   &H000080FF&
         Caption         =   "Σύνολο"
         ForeColor       =   &H00FFFFFF&
         Height          =   315
         Index           =   1
         Left            =   3225
         TabIndex        =   19
         Top             =   3900
         Width           =   615
      End
      Begin VB.Label lblLabel 
         AutoSize        =   -1  'True
         BackColor       =   &H000080FF&
         Caption         =   "Παρατηρήσεις"
         ForeColor       =   &H00FFFFFF&
         Height          =   315
         Index           =   0
         Left            =   225
         TabIndex        =   18
         Top             =   4875
         Width           =   1365
      End
      Begin VB.Shape shpWedge 
         BackColor       =   &H0000FFFF&
         BackStyle       =   1  'Opaque
         BorderStyle     =   0  'Transparent
         FillColor       =   &H00008000&
         Height          =   840
         Index           =   6
         Left            =   7425
         Top             =   3300
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
         Left            =   1575
         Top             =   2025
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
         Left            =   13500
         TabIndex        =   15
         Top             =   300
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
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H0000FF00&
         Height          =   765
         Left            =   0
         TabIndex        =   11
         Top             =   0
         Width           =   6165
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

Dim lngTotalPersonsForSelectedRows As Long

Dim lngCurrentRow As Long
Dim IsFirstReadFromDatabase As Boolean

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
    
    BeginTrans
    
    For lngRow = 1 To grdCoachesReport.RowCount
        DoEvents
        If grdCoachesReport.CellIcon(lngRow, "Selected") > 0 Then
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

Private Function CalculateSummaryPerCustomerForSelectedDestinations()

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
    Dim IsFirstItemProcessed As Boolean

    Dim lngDestinationItem  As Long
    Dim lngCustomerItem As Long
    Dim IsDestinationSelected As Boolean
    Dim IsCustomerSelected As Boolean
    
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
    
    'Προορισμός
    IsDestinationSelected = False
    IsFirstItemProcessed = True
    For lngDestinationItem = 1 To grdSummaryPerDestination.RowCount
        If grdSummaryPerDestination.CellIcon(lngDestinationItem, "Selected") <> 0 Then
            IsDestinationSelected = True
            strThisParameter = "lngDestinationID" & lngDestinationItem & " Long"
            strThisQuery = "Transfers.TransferDestinationID = " & "lngDestinationID" & lngDestinationItem
            strLogic = IIf(IsFirstItemProcessed, " AND (", " OR ")
            IsFirstItemProcessed = False
            GoSub UpdateSQLString
            arrQuery(intIndex) = grdSummaryPerDestination.CellValue(lngDestinationItem, "DestinationID")
        End If
    Next lngDestinationItem
    If IsDestinationSelected Then GoSub AddClosingParenthesis
               
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
    
    'Καθαρίζω το πλέγμα
    ClearFields grdSummaryPerCustomer
    
    'Αν δεν έχω επιλέξει κανέναν προορισμό, βγαίνω
    If Not IsDestinationSelected Then
        blnError = False
        grdCoachesReport.Redraw = True
        Exit Function
    End If
    
    'Ανοίγω το recordset
    Set rstRecordset = TempQuery.OpenRecordset()
    
    'Γεμίζω το πλέγμα των πελατών
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
    
AddClosingParenthesis:
    strParFields = strParFields & ")"
    
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

Private Function CalculateSummaryPerDriverForSelectedDestinationsAndCustomersAndRoutes()

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
    Dim IsFirstItemProcessed As Boolean

    Dim lngDestinationItem  As Long
    Dim lngCustomerItem As Long
    Dim IsDestinationSelected As Boolean
    Dim IsCustomerSelected As Boolean
    Dim lngRouteItem As Long
    Dim IsRouteSelected As Boolean
    
    'Recordsets
    Dim rstRecordset As Recordset
    
    'SQL
    strSQL = "SELECT " _
        & "DriverID, DriverDescription, Sum(Transfers.TransferAdults+Transfers.TransferKids+Transfers.TransferFree) AS SumOfTransferPersons " _
        & "FROM Transfers " _
        & "LEFT JOIN Drivers ON Transfers.TransferDriverID = Drivers.DriverID "
   
    'Ημερομηνία
    strThisParameter = "datDate Date"
    strThisQuery = "Transfers.TransferDate = datDate"
    strLogic = " AND "
    GoSub UpdateSQLString
    arrQuery(intIndex) = CDate(mskDate.text)
    
    'Προορισμός
    IsDestinationSelected = False
    IsFirstItemProcessed = True
    For lngDestinationItem = 1 To grdSummaryPerDestination.RowCount
        If grdSummaryPerDestination.CellIcon(lngDestinationItem, "Selected") <> 0 Then
            IsDestinationSelected = True
            strThisParameter = "lngDestinationID" & lngDestinationItem & " Long"
            strThisQuery = "Transfers.TransferDestinationID = " & "lngDestinationID" & lngDestinationItem
            strLogic = IIf(IsFirstItemProcessed, " AND (", " OR ")
            IsFirstItemProcessed = False
            GoSub UpdateSQLString
            arrQuery(intIndex) = grdSummaryPerDestination.CellValue(lngDestinationItem, "DestinationID")
        End If
    Next lngDestinationItem
    If IsDestinationSelected Then GoSub AddClosingParenthesis
    
    'Πελάτης
    IsCustomerSelected = False
    IsFirstItemProcessed = True
    For lngCustomerItem = 1 To grdSummaryPerCustomer.RowCount
        If grdSummaryPerCustomer.CellIcon(lngCustomerItem, "Selected") <> 0 Then
            IsCustomerSelected = True
            strThisParameter = "lngCustomerID" & lngCustomerItem & " Long"
            strThisQuery = "Transfers.TransferCustomerID = " & "lngCustomerID" & lngCustomerItem
            strLogic = IIf(IsFirstItemProcessed, " AND (", " OR ")
            IsFirstItemProcessed = False
            GoSub UpdateSQLString
            arrQuery(intIndex) = grdSummaryPerCustomer.CellValue(lngCustomerItem, "CustomerID")
        End If
    Next lngCustomerItem
    If IsCustomerSelected Then GoSub AddClosingParenthesis
    
    'Δρομολόγιο
    IsRouteSelected = False
    IsFirstItemProcessed = True
    For lngRouteItem = 1 To grdSummaryPerRoute.RowCount
        If grdSummaryPerRoute.CellIcon(lngRouteItem, "Selected") <> 0 Then
            IsRouteSelected = True
            strThisParameter = "lngRouteID" & lngRouteItem & " Long"
            strThisQuery = "Transfers.TransferRouteID = " & "lngRouteID" & lngRouteItem
            strLogic = IIf(IsFirstItemProcessed, " AND (", " OR ")
            IsFirstItemProcessed = False
            GoSub UpdateSQLString
            arrQuery(intIndex) = grdSummaryPerRoute.CellValue(lngRouteItem, "RouteID")
        End If
    Next lngRouteItem
    If IsRouteSelected Then GoSub AddClosingParenthesis
    
    strGroupBy = " GROUP BY DriverID, DriverDescription "
    strOrder = " ORDER BY Sum(Transfers.TransferAdults+Transfers.TransferKids+Transfers.TransferFree) DESC, DriverDescription "
    
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
    
    'Καθαρίζω το πλέγμα
    ClearFields grdSummaryPerDriver
    
    'Αν δεν έχω επιλέξει τουλάχιστον ένα προορισμό, έναν πελάτη και μία διαδρομή, βγαίνω
    If Not IsDestinationSelected Or Not IsCustomerSelected Or Not IsRouteSelected Then
        blnError = False
        grdCoachesReport.Redraw = True
        Exit Function
    End If
    
    'Ανοίγω το recordset
    Set rstRecordset = TempQuery.OpenRecordset()
    
    'Γεμίζω το πλέγμα
    With rstRecordset
        DoEvents
        Do While Not .EOF
            grdSummaryPerDriver.AddRow
            lngRow = lngRow + 1
            grdSummaryPerDriver.CellIcon(lngRow, "Selected") = lstIconList.ItemIndex(5)
            grdSummaryPerDriver.CellValue(lngRow, "DriverID") = !DriverID
            grdSummaryPerDriver.CellValue(lngRow, "DriverDescription") = !DriverDescription
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
    
AddClosingParenthesis:
    strParFields = strParFields & ")"
    
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
            grdSummaryPerRoute.CellValue(lngRow, "RouteShortDescription") = !PickupRouteShortDescription
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

Private Function CalculateSummaryPerRouteForSelectedDestinationsAndCustomers()

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
    Dim IsFirstItemProcessed As Boolean

    Dim lngDestinationItem  As Long
    Dim lngCustomerItem As Long
    Dim IsDestinationSelected As Boolean
    Dim IsCustomerSelected As Boolean
    
    'Recordsets
    Dim rstRecordset As Recordset
    
    'SQL
    strSQL = "SELECT " _
        & "PickupRouteID, PickupRouteShortDescription, Sum(Transfers.TransferAdults+Transfers.TransferKids+Transfers.TransferFree) AS SumOfTransferPersons " _
        & "FROM Transfers " _
        & "INNER JOIN PickupRoutes ON Transfers.TransferRouteID = PickupRoutes.PickupRouteID "
   
    'Ημερομηνία
    strThisParameter = "datDate Date"
    strThisQuery = "Transfers.TransferDate = datDate"
    strLogic = " AND "
    GoSub UpdateSQLString
    arrQuery(intIndex) = CDate(mskDate.text)
    
    'Προορισμός
    IsDestinationSelected = False
    IsFirstItemProcessed = True
    For lngDestinationItem = 1 To grdSummaryPerDestination.RowCount
        If grdSummaryPerDestination.CellIcon(lngDestinationItem, "Selected") <> 0 Then
            IsDestinationSelected = True
            strThisParameter = "lngDestinationID" & lngDestinationItem & " Long"
            strThisQuery = "Transfers.TransferDestinationID = " & "lngDestinationID" & lngDestinationItem
            strLogic = IIf(IsFirstItemProcessed, " AND (", " OR ")
            IsFirstItemProcessed = False
            GoSub UpdateSQLString
            arrQuery(intIndex) = grdSummaryPerDestination.CellValue(lngDestinationItem, "DestinationID")
        End If
    Next lngDestinationItem
    If IsDestinationSelected Then GoSub AddClosingParenthesis
    
    'Πελάτης
    IsCustomerSelected = False
    IsFirstItemProcessed = True
    For lngCustomerItem = 1 To grdSummaryPerCustomer.RowCount
        If grdSummaryPerCustomer.CellIcon(lngCustomerItem, "Selected") <> 0 Then
            IsCustomerSelected = True
            strThisParameter = "lngCustomerID" & lngCustomerItem & " Long"
            strThisQuery = "Transfers.TransferCustomerID = " & "lngCustomerID" & lngCustomerItem
            strLogic = IIf(IsFirstItemProcessed, " AND (", " OR ")
            IsFirstItemProcessed = False
            GoSub UpdateSQLString
            arrQuery(intIndex) = grdSummaryPerCustomer.CellValue(lngCustomerItem, "CustomerID")
        End If
    Next lngCustomerItem
    If IsCustomerSelected Then GoSub AddClosingParenthesis
    
    strGroupBy = " GROUP BY PickupRouteID, PickupRouteShortDescription "
    strOrder = " ORDER BY Sum(Transfers.TransferAdults+Transfers.TransferKids+Transfers.TransferFree) DESC, PickupRouteShortDescription "
    
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
    
    'Καθαρίζω το πλέγμα
    ClearFields grdSummaryPerRoute
    
    'Αν δεν έχω επιλέξει τουλάχιστον ένα προορισμό και έναν πελάτη, βγαίνω
    If Not IsCustomerSelected Or Not IsDestinationSelected Then
        blnError = False
        grdCoachesReport.Redraw = True
        Exit Function
    End If
    
    'Ανοίγω το recordset
    Set rstRecordset = TempQuery.OpenRecordset()
    
    'Γεμίζω το πλέγμα
    With rstRecordset
        DoEvents
        Do While Not .EOF
            grdSummaryPerRoute.AddRow
            lngRow = lngRow + 1
            grdSummaryPerRoute.CellIcon(lngRow, "Selected") = lstIconList.ItemIndex(4)
            grdSummaryPerRoute.CellValue(lngRow, "RouteID") = !PickupRouteID
            grdSummaryPerRoute.CellValue(lngRow, "RouteShortDescription") = !PickupRouteShortDescription
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
    
AddClosingParenthesis:
    strParFields = strParFields & ")"
    
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
        ClearFields txtDestinationDescription, txtCustomerDescription, txtPickupPointDescription, mskAdults, mskKids, mskFree, txtRemarks, txtDriverDescription
        ClearFields mskTotal
        DisableFields mskDate, txtCustomerDescription, txtDestinationDescription, txtPickupPointDescription, mskAdults, mskKids, mskFree, txtRemarks, txtDriverDescription
        DisableFields cmdIndex(0), cmdIndex(1), cmdIndex(2), cmdIndex(3), cmdIndex(4)
        FindRecordsAndPopulateGrid
        If Not blnStatus Then ClearFields txtTransferID
        blnStatus = True
    End If

End Function

Private Function DisplayAssignRoutesToDriverDialog()

    Dim lngRow As Long
    Dim IsRowSelected As Boolean
    
    For lngRow = 1 To grdCoachesReport.RowCount
        If grdCoachesReport.CellIcon(lngRow, "Selected") > 0 Then
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

    If ValidateFields(False) Then
        'Init
        IsFirstReadFromDatabase = True
        'Σύνολα
        CalculateSummaryPerDestination
        CalculateSummaryPerCustomerForSelectedDestinations
        CalculateSummaryPerRoute
        CalculateSummaryPerDriver
        'Επιλογή όλων
        chkAllDestinations.Value = 1
        chkAllCustomers.Value = 1
        chkAllRoutes.Value = 1
        chkAllDrivers.Value = 1
        'Upate Init
        IsFirstReadFromDatabase = False
        'Κεντρικό πλέγμα
        RefreshList
        If grdCoachesReport.RowCount > 0 Then
            lblTotalPersons.Caption = "Σύνολο ημέρας: " & format(CalculateTotalPersons, "#,##0")
            lblSelectedGridLines.Caption = "Σύνολο επιλεγμένων: 0"
            'Υπόλοιπα
            EnableGrid grdCoachesReport, False
            EnableFields chkAllTransfers, chkAllDestinations, chkAllCustomers, chkAllRoutes, chkAllDrivers
            DisableFields mskDate
            HighlightRow txtTransferID.text
            UpdateButtons Me, 11, 0, 1, 0, 0, 0, 1, 1, 1, 1, 0, 0, 0
            Exit Function
        Else
            UpdateButtons Me, 11, 1, 1, 0, 0, 0, 0, 0, 0, 1, 0, 0, 0
            If Not blnError Then
                If blnProcessing Then
                    If MyMsgBox(4, strApplicationName, strStandardMessages(27), 1) Then
                    End If
                Else
                    If MyMsgBox(1, strApplicationName, strStandardMessages(7), 1) Then
                    End If
                    If mskDate.Enabled Then mskDate.SetFocus
                End If
            End If
            blnError = False
            blnProcessing = False
        End If
    End If

End Function

Private Function GetDriverName()

    Dim lngRow As Long
    Dim intDriversSelected As Integer
    
    For lngRow = 1 To grdSummaryPerDriver.RowCount
        If grdSummaryPerDriver.CellIcon(lngRow, "Selected") > 0 Then
            intDriversSelected = intDriversSelected + 1
            GetDriverName = grdSummaryPerDriver.CellValue(lngRow, "DriverDescription")
        End If
    Next lngRow
    
    'If intDriversSelected > 1 Then
    '    MyMsgBox 4, strApplicationName, strAppMessages(8), 1
    '    GetDriverName = ""
    'End If

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

Private Function PositionFormButtons()
    
    Dim intLoop As Integer
    
    For intLoop = 1 To 4
        cmdButton(intLoop).Top = grdCoachesReport.Top + grdCoachesReport.Height - cmdButton(intLoop).Height - 150
    Next intLoop

End Function

Private Function PositionSeperator()
    
    Seperator.Left = 225
    Seperator.Width = frmContainer.Width - 450
    Seperator.Top = GetSetting(appName:=strApplicationName, Section:="Settings", Key:="SeperatorTop")

End Function

Private Function RecolorizeControls()

    Dim intIndex As Integer
    
    Me.BackColor = GetSetting(strApplicationName, "Colors", "Forms Centered Background")
    Me.shpBackground.BackColor = GetSetting(strApplicationName, "Colors", "Background Containers")
    Me.frmButtonFrame.BackColor = GetSetting(strApplicationName, "Colors", "Forms Centered Background")
    
    frmContainer.BackColor = Me.BackColor
   
    For intIndex = 0 To lblLabel.UBound
        lblLabel(intIndex).BackColor = Me.shpBackground.BackColor
        lblLabel(intIndex).ForeColor = vbBlack
    Next intIndex
            
    shpBackground.BackColor = GetSetting(strApplicationName, "Colors", "Background Containers")
    
    Seperator.BackColor = Me.BackColor
            
    chkAllTransfers.BackColor = shpBackground.BackColor
    chkAllDestinations.BackColor = shpBackground.BackColor
    chkAllCustomers.BackColor = shpBackground.BackColor
    chkAllRoutes.BackColor = shpBackground.BackColor
    chkAllDrivers.BackColor = shpBackground.BackColor
    
    chkAllTransfers.ForeColor = vbWhite
    chkAllDestinations.ForeColor = vbWhite
    chkAllCustomers.ForeColor = vbWhite
    chkAllRoutes.ForeColor = vbWhite
    chkAllDrivers.ForeColor = vbWhite
    
    frmSummaries.BackColor = shpBackground.BackColor

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
    Dim IsFirstItemProcessed As Boolean
    Dim lngDestinationItem As Long
    Dim IsDestinationSelected  As Boolean
    Dim lngCustomerItem As Long
    Dim IsCustomerSelected  As Boolean
    Dim lngRouteItem As Long
    Dim IsRouteSelected  As Boolean
    Dim lngDriverItem As Long
    Dim IsDriverSelected  As Boolean
    
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
    
    'Προορισμός
    IsDestinationSelected = False
    IsFirstItemProcessed = True
    For lngDestinationItem = 1 To grdSummaryPerDestination.RowCount
        If grdSummaryPerDestination.CellIcon(lngDestinationItem, "Selected") <> 0 Then
            IsDestinationSelected = True
            strThisParameter = "lngDestinationID" & lngDestinationItem & " Long"
            strThisQuery = "Transfers.TransferDestinationID = " & "lngDestinationID" & lngDestinationItem
            strLogic = IIf(IsFirstItemProcessed, " AND (", " OR ")
            IsFirstItemProcessed = False
            GoSub UpdateSQLString
            arrQuery(intIndex) = grdSummaryPerDestination.CellValue(lngDestinationItem, "DestinationID")
        End If
    Next lngDestinationItem
    If IsDestinationSelected Then GoSub AddClosingParenthesis
    
    'Πελάτης
    IsCustomerSelected = False
    IsFirstItemProcessed = True
    For lngCustomerItem = 1 To grdSummaryPerCustomer.RowCount
        If grdSummaryPerCustomer.CellIcon(lngCustomerItem, "Selected") <> 0 Then
            IsCustomerSelected = True
            strThisParameter = "lngCustomerID" & lngCustomerItem & " Long"
            strThisQuery = "Transfers.TransferCustomerID = " & "lngCustomerID" & lngCustomerItem
            strLogic = IIf(IsFirstItemProcessed, " AND (", " OR ")
            IsFirstItemProcessed = False
            GoSub UpdateSQLString
            arrQuery(intIndex) = grdSummaryPerCustomer.CellValue(lngCustomerItem, "CustomerID")
        End If
    Next lngCustomerItem
    If IsCustomerSelected Then GoSub AddClosingParenthesis
    
    'Δρομολόγιο
    IsRouteSelected = False
    IsFirstItemProcessed = True
    For lngRouteItem = 1 To grdSummaryPerRoute.RowCount
        If grdSummaryPerRoute.CellIcon(lngRouteItem, "Selected") <> 0 Then
            IsRouteSelected = True
            strThisParameter = "lngRouteID" & lngRouteItem & " Long"
            strThisQuery = "Transfers.TransferRouteID = " & "lngRouteID" & lngRouteItem
            strLogic = IIf(IsFirstItemProcessed, " AND (", " OR ")
            IsFirstItemProcessed = False
            GoSub UpdateSQLString
            arrQuery(intIndex) = grdSummaryPerRoute.CellValue(lngRouteItem, "RouteID")
        End If
    Next lngRouteItem
    If IsRouteSelected Then GoSub AddClosingParenthesis
    
    'Οδηγός
    IsDriverSelected = False
    IsFirstItemProcessed = True
    For lngDriverItem = 1 To grdSummaryPerDriver.RowCount
        If grdSummaryPerDriver.CellIcon(lngDriverItem, "Selected") <> 0 Then
            IsDriverSelected = True
            strThisParameter = "lngDriverID" & lngDriverItem & " Long"
            strThisQuery = "Transfers.TransferDriverID = " & "lngDriverID" & lngDriverItem
            strLogic = IIf(IsFirstItemProcessed, " AND (", " OR ")
            IsFirstItemProcessed = False
            GoSub UpdateSQLString
            arrQuery(intIndex) = grdSummaryPerDriver.CellValue(lngDriverItem, "DriverID")
        End If
    Next lngDriverItem
    If IsDriverSelected Then GoSub AddClosingParenthesis
    
    'Ταξινόμηση
    strOrder = "ORDER BY PickupRouteDescription, PickupPointTime, PickUpPointHotelDescription"

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
    
    'Αν έχω μόνο την ημερομηνία και κανένα άλλο κριτήριο, βγαίνω
    If Not IsDestinationSelected Or Not IsCustomerSelected Or Not IsRouteSelected Or Not IsDriverSelected Then
        blnError = False
        RefreshList = False
        grdCoachesReport.Redraw = True
        Exit Function
    End If
    
    'Ανοίγω το recordset
    Set rstRecordset = TempQuery.OpenRecordset()
    
    'Αν δεν έχω εγγραφές, βγαίνω
    If rstRecordset.RecordCount = 0 Then blnError = False: RefreshList = False: grdCoachesReport.Redraw = True: Exit Function
    
    'Προετοιμάζω τη μπάρα προόδου
    'InitializeProgressBar Me, strApplicationName, rstRecordset
    
    'Προσωρινά
    'UpdateButtons Me, 11, 0, 0, 0, 0, 0, 0, 0, 0, 1, 0, 0, 0
    'cmdButton(8).Caption = "Διακοπή επεξεργασίας"
    blnProcessing = True
    
    'Γεμίζω το πλέγμα
    With rstRecordset
        Do While Not .EOF
            grdCoachesReport.AddRow
            lngRowCount = rstRecordset.RecordCount
            'UpdateProgressBar Me
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
        grdCoachesReport.Redraw = True
        grdCoachesReport.SetCurCell 1, 1
    End If
    
    'Τελικές ενέργειες
    chkAllTransfers_Click
    Me.Refresh
   
    Exit Function
    
UpdateSQLString:
    intIndex = intIndex + 1
    strParameters = IIf(intIndex > 1, strParameters & ", ", strParameters)
    strParFields = IIf(intIndex > 1, strParFields & strLogic, strParFields)
    strParameters = strParameters & strThisParameter
    strParFields = strParFields & strThisQuery
    ReDim Preserve arrQuery(intIndex)
    
    Return

AddClosingParenthesis:
    strParFields = strParFields & ")"
    
    Return

ErrTrap:
    blnError = True
    ClearFields grdCoachesReport
    DisplayErrorMessage True, Err.Description

End Function

Private Function ToggleCheckBox(grid As iGrid, initialValue)

    Dim lngRow As Long
    Dim lngSelectedRows As Long
    Dim lngNotSelectedRows As Long
    Dim lngRowCount As Long
    
    lngRowCount = grid.RowCount
    
    For lngRow = 1 To grid.RowCount
        If grid.CellIcon(lngRow, "Selected") > 0 Then
            lngSelectedRows = lngSelectedRows + 1
        End If
        If grid.CellIcon(lngRow, "Selected") = 0 Then
            lngNotSelectedRows = lngNotSelectedRows + 1
        End If
    Next lngRow
    
    ToggleCheckBox = initialValue
    
    If lngSelectedRows = lngRowCount Then ToggleCheckBox = 1
    If lngNotSelectedRows = lngRowCount Then ToggleCheckBox = 0

End Function

Private Sub chkAllCustomers_Click()

    Dim lngRow As Long
    
    For lngRow = 1 To grdSummaryPerCustomer.RowCount
        grdSummaryPerCustomer.CellIcon(lngRow, "Selected") = lstIconList.ItemIndex(IIf(chkAllCustomers.Value <= 0, 1, 3))
    Next lngRow
    
    CalculateSummaryPerRouteForSelectedDestinationsAndCustomers
    
    chkAllRoutes.Value = chkAllCustomers.Value
    
    If Not IsFirstReadFromDatabase Then RefreshList

End Sub

Private Sub chkAllDestinations_Click()

    Dim lngRow As Long
    
    For lngRow = 1 To grdSummaryPerDestination.RowCount
        grdSummaryPerDestination.CellIcon(lngRow, "Selected") = lstIconList.ItemIndex(IIf(chkAllDestinations.Value <= 0, 1, 2))
    Next lngRow
    
    CalculateSummaryPerCustomerForSelectedDestinations
    
    chkAllCustomers.Value = chkAllDestinations.Value
    
    If Not IsFirstReadFromDatabase Then RefreshList

End Sub

Private Sub chkAllDrivers_Click()

    Dim lngRow As Long
    
    For lngRow = 1 To grdSummaryPerDriver.RowCount
        grdSummaryPerDriver.CellIcon(lngRow, "Selected") = lstIconList.ItemIndex(IIf(chkAllDrivers.Value <= 0, 1, 5))
    Next lngRow
    
    If Not IsFirstReadFromDatabase Then RefreshList
    
End Sub

Private Sub chkAllRoutes_Click()
    
    Dim lngRow As Long
    
    For lngRow = 1 To grdSummaryPerRoute.RowCount
        grdSummaryPerRoute.CellIcon(lngRow, "Selected") = lstIconList.ItemIndex(IIf(chkAllRoutes.Value <= 0, 1, 4))
    Next lngRow
    
    CalculateSummaryPerDriverForSelectedDestinationsAndCustomersAndRoutes
    
    chkAllDrivers.Value = chkAllRoutes.Value
    
    If Not IsFirstReadFromDatabase Then RefreshList
    
End Sub

Private Sub chkAllTransfers_Click()

    Dim lngRow As Long
    
    grdCoachesReport.Redraw = True
    lngTotalPersonsForSelectedRows = 0
    
    For lngRow = 1 To grdCoachesReport.RowCount
        grdCoachesReport.CellIcon(lngRow, "Selected") = lstIconList.ItemIndex(IIf(chkAllTransfers.Value = 0, 1, 6))
        lngTotalPersonsForSelectedRows = lngTotalPersonsForSelectedRows + IIf(grdCoachesReport.CellIcon(lngRow, "Selected") > 0, grdCoachesReport.CellValue(lngRow, "TransferTotal"), 0)
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
            DoReport "Print"
        Case 7
            DoReport "CreatePDF"
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

    blnCancel = False
    DisableFields mskDate
    EnableFields txtDestinationDescription, txtCustomerDescription, txtPickupPointDescription, mskAdults, mskKids, mskFree, txtRemarks, txtDriverDescription
    EnableFields cmdIndex(0), cmdIndex(1), cmdIndex(2), cmdIndex(3), cmdIndex(4)
    UpdateButtons Me, 11, 0, 0, 1, 0, 1, 0, 0, 0, 0, 0, 0, 0, 0
        
    If True Then
        If txtTransferID.text <> "" Then
            DisplayLastRecord txtTransferID.text
            blnStatus = True
            ClearFields txtTransferID, txtCustomerID, txtRouteID, txtPickupPointID, txtDriverID
            ClearFields txtCustomerDescription, txtPickupPointDescription, mskAdults, mskKids, mskFree, txtRemarks, txtDriverDescription
            ClearFields mskTotal
            txtCustomerDescription.SetFocus
        Else
            blnStatus = True
            ClearFields txtTransferID, txtDestinationID, txtCustomerID, txtRouteID, txtPickupPointID, txtDriverID
            ClearFields txtDestinationDescription, txtCustomerDescription, txtPickupPointDescription, mskAdults, mskKids, mskFree, txtRemarks, txtDriverDescription
            ClearFields mskTotal
            txtDestinationDescription.SetFocus
        End If
    End If
    
    InitializeFields mskAdults, mskKids, mskFree, mskTotal
        
End Function

Private Function DisplayLastRecord(lngLastRecord)

    If Not SeekRecord(lngLastRecord) Then Exit Function

End Function

Private Function SaveRecord()

    If Not ValidateFields(True) Then Exit Function
    
    txtTransferID.text = MainSaveRecord("CommonDB", "Transfers", blnStatus, strApplicationName, "ID", txtTransferID.text, mskDate.text, txtDestinationID.text, txtCustomerID.text, txtRouteID.text, txtPickupPointID.text, mskAdults.text, mskKids.text, mskFree.text, txtRemarks.text, IIf(txtDriverID.text = "", "28", txtDriverID.text), 1, strCurrentUser)
    
    If txtTransferID.text <> "" Then
        SaveRecord = True
        blnCancel = True
        ClearFields txtDestinationID, txtCustomerID, txtRouteID, txtPickupPointID, txtDriverID
        ClearFields txtDestinationDescription, txtCustomerDescription, txtPickupPointDescription, mskAdults, mskKids, mskFree, txtRemarks, txtDriverDescription
        ClearFields mskTotal
        DisableFields mskDate, txtCustomerDescription, txtDestinationDescription, txtPickupPointDescription, mskAdults, mskKids, mskFree, txtRemarks, txtDriverDescription
        DisableFields cmdIndex(0), cmdIndex(1), cmdIndex(2), cmdIndex(3), cmdIndex(4)
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
    
    'Οδηγοί
    Dim lngDriverRow As Long
    
    Dim strDriverName As String
    
    strDriverName = GetDriverName
    If strDriverName = "" Then Exit Function
    
    grdCoachesReport.SortObject.colCount = 3
    grdCoachesReport.SortObject.SortCol(1) = grdCoachesReport.ColIndex("DriverDescription")
    grdCoachesReport.SortObject.SortCol(2) = grdCoachesReport.ColIndex("PickupPointTime")
    grdCoachesReport.SortObject.SortCol(3) = grdCoachesReport.ColIndex("PickupPointHotelDescription")
    grdCoachesReport.Sort
    
    If action = "Print" Then
        If SelectPrinter("PrinterPrintsReports") Then
            For lngDriverRow = 1 To grdSummaryPerDriver.RowCount
                If grdSummaryPerDriver.CellIcon(lngDriverRow, "Selected") > 0 Then
                    CreateUnicodeFile "ΑΝΑΦΟΡΑ ΠΑΡΑΛΑΒΩΝ ΓΙΑ : " & mskDate.text, "ΟΔΗΓΟΣ: " & grdSummaryPerDriver.CellValue(lngDriverRow, "DriverDescription"), GetSetting(strApplicationName, "Settings", "Export Report Transfers Height") - 11, lngDriverRow
                    With rptTransfers
                        .oneLongField.Font.Size = 11
                        If intPreviewReports = 1 Then
                            .Restart
                            .Zoom = -2
                            .WindowState = vbMaximized
                            .Show 1
                        Else
                            .Restart
                            .Printer.DeviceName = strPrinterName
                            .PrintReport False
                            .Run True
                        End If
                    End With
                End If
            Next lngDriverRow
        End If
    End If
    
    If action = "CreatePDF" Then
        For lngDriverRow = 1 To grdSummaryPerDriver.RowCount
            If grdSummaryPerDriver.CellIcon(lngDriverRow, "Selected") > 0 Then
                CreateUnicodeFile "ΑΝΑΦΟΡΑ ΠΑΡΑΛΑΒΩΝ ΓΙΑ : " & mskDate.text, "ΟΔΗΓΟΣ: " & grdSummaryPerDriver.CellValue(lngDriverRow, "DriverDescription"), GetSetting(strApplicationName, "Settings", "Export Report Transfers Height") - 11, lngDriverRow
                CreateUnisexPDF "ΑΝΑΦΟΡΑ ΠΑΡΑΛΑΒΩΝ ΓΙΑ : " & mskDate.text & " ΟΔΗΓΟΣ: " & grdSummaryPerDriver.CellValue(lngDriverRow, "DriverDescription"), rptTransfers, 11
            End If
        Next lngDriverRow
        If MyMsgBox(1, strApplicationName, strStandardMessages(8), 1) Then
        End If
    End If
    
    Exit Function
    
ErrTrap:
    Close #1
    DisplayErrorMessage True, Err.Description
    
End Function

Private Function ValidateFields(IsSavingRecord)

    ValidateFields = False
    
    'Ημερομηνία
    If mskDate.text = "" Then
        mskDate.SetFocus
        Exit Function
    End If
    If Not IsDate(mskDate.text) Then
        If MyMsgBox(4, strApplicationName, strStandardMessages(2), 1) Then
        End If
        mskDate.SetFocus
        Exit Function
    End If
    
    If Not IsSavingRecord Then ValidateFields = True: Exit Function
    
    'Προορισμός
    If txtDestinationID.text = "" Then
        If MyMsgBox(4, strApplicationName, strStandardMessages(1), 1) Then
        End If
        txtDestinationDescription.SetFocus
        Exit Function
    End If

    'Πελάτης
    If txtCustomerID.text = "" Then
        If MyMsgBox(4, strApplicationName, strStandardMessages(1), 1) Then
        End If
        txtCustomerDescription.SetFocus
        Exit Function
    End If

    'Σημείο παραλαβής
    If txtPickupPointID.text = "" Then
        If MyMsgBox(4, strApplicationName, strStandardMessages(1), 1) Then
        End If
        txtPickupPointDescription.SetFocus
        Exit Function
    End If

    'Ενήλικες
    If mskAdults.text = "" Then
        If MyMsgBox(4, strApplicationName, strStandardMessages(1), 1) Then
        End If
        mskAdults.SetFocus
        Exit Function
    End If
    
    'Παιδιά
    If mskKids.text = "" Then
        If MyMsgBox(4, strApplicationName, strStandardMessages(1), 1) Then
        End If
        mskKids.SetFocus
        Exit Function
    End If
    
    'Δωρεάν
    If mskFree.text = "" Then
        If MyMsgBox(4, strApplicationName, strStandardMessages(1), 1) Then
        End If
        mskFree.SetFocus
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
            ClearFields txtCustomerDescription, txtDestinationDescription, txtPickupPointDescription, mskAdults, mskKids, mskFree, txtRemarks, txtDriverDescription
            ClearFields mskTotal
            DisableFields txtCustomerDescription, txtDestinationDescription, txtPickupPointDescription, mskAdults, mskKids, mskFree, txtRemarks, txtDriverDescription
            DisableFields cmdIndex(1), cmdIndex(2), cmdIndex(3), cmdIndex(4)
            UpdateButtons Me, 11, 0, 1, 0, 0, 0, IIf(grdCoachesReport.RowCount > 0, 1, 0), IIf(grdCoachesReport.RowCount > 0, 1, 0), IIf(grdCoachesReport.RowCount > 0, 1, 0), 1, 0, 0, 0
            grdCoachesReport.SetFocus
            blnStatus = True
            Exit Function
        End If
        Exit Function
    End If
    
    'Νέα αναζήτηση
    If Not blnStatus And Not cmdButton(9).Enabled Then
        IsFirstReadFromDatabase = True
        ClearFields grdCoachesReport
        ClearFields grdSummaryPerDestination, grdSummaryPerCustomer, grdSummaryPerRoute, grdSummaryPerDriver
        ClearFields lblTotalPersons, lblSelectedGridLines
        ClearFields chkAllTransfers, chkAllTransfers, chkAllDestinations, chkAllCustomers, chkAllRoutes, chkAllDrivers
        DisableFields chkAllTransfers, chkAllTransfers, chkAllDestinations, chkAllCustomers, chkAllRoutes, chkAllDrivers
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
            'Σημείο παραλαβής
            If txtDestinationID.text <> "" Then
                txtPickupPointDescription.text = Replace(txtPickupPointDescription.text, "'", "")
                'Δεν έχω δώσει δρομολόγιο, βρίσκω τα σημεία παραλαβής που είναι συνδεδεμένα με τον δοσμένο προορισμό
                If txtRouteID.text = "" Then
                    intSize = Len(txtPickupPointDescription.text)
                    If intSize = 0 Then Exit Sub
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
        Case 4
            'Οδηγός
            Set tmpRecordset = CheckForMatch("CommonDB", "Drivers", "DriverDescription", "String", txtDriverDescription.text)
            If tmpRecordset.RecordCount > 0 Then
                tmpTableData = DisplayIndex(tmpRecordset, 2, True, 2, 0, 1, "ID", "Περιγραφή", 0, 40, 1, 0)
                txtDriverID.text = tmpTableData.strCode
                txtDriverDescription.text = tmpTableData.strFirstField
            End If
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
        Else
            txtRouteID.text = ""
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
        
        AddColumnsToGrid grdCoachesReport, True, 44, GetSetting(strApplicationName, "Layout Strings", grdCoachesReport.Tag), _
            "05NCNTransferID,12NCDTransferDate,40NLNCustomerDescription,40NCNDestinationShortDescription,40NLNDestinationDescription,50NCNRouteShortDescription,50NLNRouteDescription,40NLNPickupPointHotelDescription,10NLNPickUpPointExactPoint,10NCTPickupPointTime,10NRITransferAdults,10NRITransferKids,10NRITransferFree,10NLNTransferRemarks,10NLNDriverDescription,10NRITransferTotal,04NCNTotalCriteria,04NCNSelected", _
            "TransferID,Ημερομηνία,Πελάτης,Π,Προορισμός,Δρομολόγιο,Δρομολόγιο,Σημείο παραλαβής,Ακριβές σημείο,Ωρα,Ε,Π,Δ,Παρατηρήσεις,Οδηγός,Σύνολο,Κριτήρια,Ε"
        AddColumnsToGrid grdSummaryPerDestination, True, 24, GetSetting(strApplicationName, "Layout Strings", "grdCoachesReportSummaryPerDestination"), _
            "04NCNSelected,05NCNDestinationID,40NLNDestinationDescription,10NRITotalPersons", _
            "E,DestinationID,Προορισμός,Ατομα"
        AddColumnsToGrid grdSummaryPerCustomer, True, 24, GetSetting(strApplicationName, "Layout Strings", "grdCoachesReportSummaryPerCustomer"), _
            "04NCNSelected,05NCNCustomerID,40NLNCustomerDescription,10NRITotalPersons", _
            "E,CustomerID,Πελάτης,Ατομα"
        AddColumnsToGrid grdSummaryPerRoute, True, 24, GetSetting(strApplicationName, "Layout Strings", "grdCoachesReportSummaryPerRoute"), _
            "04NCNSelected,05NCNRouteID,40NLNRouteShortDescription,10NRITotalPersons", _
            "E,RouteID,Δρομολόγιο,Ατομα"
        AddColumnsToGrid grdSummaryPerDriver, True, 24, GetSetting(strApplicationName, "Layout Strings", "grdCoachesReportSummaryPerDriver"), _
            "04NCNSelected,05NCNDriverID,40NLNDriverDescription,10NRITotalPersons", _
            "E,DriverID,Οδηγός,Ατομα"
        
        Me.Refresh
        mskDate.SetFocus
        
    End If
    
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
        Case vbKeyC And CtrlDown And cmdButton(0).Enabled
            cmdButton_Click 0
        Case vbKeyN And CtrlDown And cmdButton(1).Enabled
            cmdButton_Click 1
        Case vbKeyS And CtrlDown And cmdButton(2).Enabled
            cmdButton_Click 2
        Case vbKeyD And CtrlDown And cmdButton(3).Enabled
            cmdButton_Click 3
        Case vbKeyP And CtrlDown And Not AltDown And cmdButton(6).Enabled
            cmdButton_Click 6
        Case vbKeyP And CtrlDown And AltDown And cmdButton(7).Enabled
            cmdButton_Click 7
        Case vbKeyEscape
            If cmdButton(4).Enabled Then cmdButton_Click 4: Exit Function 'Ακύρωση επεξεργασίας
            If cmdButton(8).Enabled Then cmdButton_Click 8: Exit Function 'Νέα αναζήτηση (επιστροφή στην ημερομηνία)
            If cmdButton(9).Enabled Then cmdButton_Click 9: Exit Function 'Εξοδος
            If cmdButton(10).Enabled Then cmdButton_Click 10 'Κλείσιμο φόρμας Σύνδεσης δρομολογίου με οδηγό
        Case vbKey0 And CtrlDown And grdCoachesReport.RowCount > 0
            grdSummaryPerCustomer.SetCurCell 1, 1
            grdCoachesReport.SetFocus
        Case vbKey1 And CtrlDown And grdSummaryPerDestination.RowCount > 0
            grdSummaryPerDestination.SetCurCell 1, 1
            grdSummaryPerDestination.SetFocus
        Case vbKey2 And CtrlDown And grdSummaryPerCustomer.RowCount > 0
            grdSummaryPerCustomer.SetCurCell 1, 1
            grdSummaryPerCustomer.SetFocus
        Case vbKey3 And CtrlDown And grdSummaryPerRoute.RowCount > 0
            grdSummaryPerRoute.SetCurCell 1, 1
            grdSummaryPerRoute.SetFocus
        Case vbKey4 And CtrlDown And grdSummaryPerDriver.RowCount > 0
            grdSummaryPerDriver.SetCurCell 1, 1
            grdSummaryPerDriver.SetFocus
        Case vbKeyF12 And CtrlDown
            ToggleInfoPanel Me
    End Select

End Function

Private Sub Form_Load()

    blnCancel = True
    
    lngMinimumSeperatorTop = 6825
    lngMaximumSeperatorTop = 11585
    
    SetUpGrid lstIconList, grdCoachesReport, grdSummaryPerDestination, grdSummaryPerCustomer, grdSummaryPerRoute, grdSummaryPerDriver
    PositionControls Me, True, grdCoachesReport
    PositionSeperator
    RepositionMainGrid
    RecolorizeControls
    PositionGrids
    PositionFormButtons
    
    frmCriteria(0).Visible = False
    
    ClearFields txtTransferID, txtCustomerID, txtPickupPointID, txtRouteID, txtDestinationID, txtDriverID
    ClearFields mskDate, txtCustomerDescription, txtDestinationDescription, txtPickupPointDescription, mskAdults, mskKids, mskFree, txtRemarks, txtDriverDescription
    ClearFields mskTotal
    ClearFields chkAllTransfers, chkAllTransfers, chkAllDestinations, chkAllCustomers, chkAllRoutes, chkAllDrivers
    ClearFields lblTotalPersons, lblSelectedGridLines
    
    DisableFields txtCustomerDescription, txtDestinationDescription, txtPickupPointDescription, mskAdults, mskKids, mskFree, txtRemarks, txtDriverDescription
    DisableFields cmdIndex(1), cmdIndex(2), cmdIndex(3), cmdIndex(4)
    DisableFields chkAllTransfers, chkAllTransfers, chkAllDestinations, chkAllCustomers, chkAllRoutes, chkAllDrivers
    
    UpdateButtons Me, 11, 1, 0, 0, 0, 0, 0, 0, 0, 0, 1, 0, 0
    
End Sub

Private Function RepositionMainGrid()

    grdCoachesReport.Height = grdCoachesReport.Height - frmSummaries.Height - 150

End Function

Private Function SeekRecord(TransferID)
    
    Dim tmpRecordset As Recordset
    Dim tmpTableData As typTableData
    
    ClearFields txtDestinationID, txtCustomerID, txtRouteID, txtPickupPointID, txtDriverID
    ClearFields mskDate, txtDestinationDescription, txtCustomerDescription, txtPickupPointDescription, mskAdults, mskKids, mskFree, txtRemarks, txtDriverDescription
    ClearFields mskTotal
    DisableFields mskDate, txtCustomerDescription, txtDestinationDescription, txtPickupPointDescription, mskAdults, mskKids, mskFree, txtRemarks
    DisableFields cmdIndex(0), cmdIndex(1), cmdIndex(2), cmdIndex(3), cmdIndex(4)
    
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
        EnableFields mskDate, txtDestinationDescription, txtCustomerDescription, txtPickupPointDescription, mskAdults, mskKids, mskFree, txtRemarks, txtDriverDescription
        EnableFields cmdIndex(0), cmdIndex(1), cmdIndex(2), cmdIndex(3), cmdIndex(4)
        mskTotal.Caption = AddNumbers(mskAdults.text, mskKids.text, mskFree.text)
        blnCancel = False
        blnStatus = False
        SeekRecord = True
        lngCurrentRow = grdCoachesReport.CurRow
        UpdateButtons Me, 11, 0, 0, 1, 1, 1, 0, 0, 0, 0, 0, 0, 0
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
        grdCoachesReport.CellIcon(grdCoachesReport.CurRow, "Selected") = lstIconList.ItemIndex(SelectRow(grdCoachesReport, 6, KeyCode, grdCoachesReport.CurRow, "TransferID"))
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

Private Sub grdSummaryPerCustomer_ColHeaderMouseEnter(ByVal lCol As Long)

    grdSummaryPerCustomer.Header.Buttons = True
    
End Sub

Private Sub grdSummaryPerCustomer_ColHeaderMouseLeave(ByVal lCol As Long)

    grdSummaryPerCustomer.Header.Buttons = False
    
End Sub


Private Sub grdSummaryPerCustomer_DblClick(ByVal lRow As Long, ByVal lCol As Long, bRequestEdit As Boolean)

    'Customers Grid
    
    If grdSummaryPerDestination.RowCount > 0 Then
        
        'Toggle selected line
        grdSummaryPerCustomer.CellIcon(grdSummaryPerCustomer.CurRow, "Selected") = lstIconList.ItemIndex(SelectRow(grdSummaryPerCustomer, 3, 32, grdSummaryPerCustomer.CurRow, "CustomerDescription"))
        
        'Routes
        CalculateSummaryPerRouteForSelectedDestinationsAndCustomers
        
        'Drivers
        CalculateSummaryPerDriverForSelectedDestinationsAndCustomersAndRoutes
        
        'Toggle checkbox
        chkAllCustomers.Value = ToggleCheckBox(grdSummaryPerCustomer, chkAllCustomers.Value)
        
        'Main grid
        RefreshList
        
    End If
End Sub

Private Sub grdSummaryPerCustomer_HeaderRightClick(ByVal lCol As Long, ByVal Shift As Integer, ByVal X As Long, ByVal Y As Long)

    PopupMenu mnuHdrPopUp
    
End Sub

Private Sub grdSummaryPerCustomer_KeyDown(KeyCode As Integer, Shift As Integer, bDoDefault As Boolean)

    'Customers Grid
    
    If KeyCode = vbKeySpace And grdSummaryPerDestination.RowCount > 0 Then
        
        'Toggle selected line
        grdSummaryPerCustomer.CellIcon(grdSummaryPerCustomer.CurRow, "Selected") = lstIconList.ItemIndex(SelectRow(grdSummaryPerCustomer, 3, KeyCode, grdSummaryPerCustomer.CurRow, "CustomerDescription"))
        
        'Routes
        CalculateSummaryPerRouteForSelectedDestinationsAndCustomers
        
        'Drivers
        CalculateSummaryPerDriverForSelectedDestinationsAndCustomersAndRoutes
        
        'Toggle checkbox
        chkAllCustomers.Value = ToggleCheckBox(grdSummaryPerCustomer, chkAllCustomers.Value)
        
        'Main grid
        RefreshList
        
    End If

End Sub

Private Sub grdSummaryPerDestination_ColHeaderMouseEnter(ByVal lCol As Long)

    grdSummaryPerDestination.Header.Buttons = True
    
End Sub

Private Sub grdSummaryPerDestination_ColHeaderMouseLeave(ByVal lCol As Long)

    grdSummaryPerDestination.Header.Buttons = False
    
End Sub


Private Sub grdSummaryPerDestination_DblClick(ByVal lRow As Long, ByVal lCol As Long, bRequestEdit As Boolean)

    'Destinations Grid
        
    If grdSummaryPerDestination.RowCount > 0 Then
        
        'Toggle selected line
        grdSummaryPerDestination.CellIcon(grdSummaryPerDestination.CurRow, "Selected") = lstIconList.ItemIndex(SelectRow(grdSummaryPerDestination, 2, 32, grdSummaryPerDestination.CurRow, "DestinationDescription"))
        
        'Customers
        CalculateSummaryPerCustomerForSelectedDestinations
        
        'Routes
        CalculateSummaryPerRouteForSelectedDestinationsAndCustomers
        
        'Drivers
        CalculateSummaryPerDriverForSelectedDestinationsAndCustomersAndRoutes
        
        'Toggle checkbox
        chkAllDestinations.Value = ToggleCheckBox(grdSummaryPerDestination, chkAllDestinations.Value)
        
        'Main grid
        RefreshList
    
    End If
End Sub

Private Sub grdSummaryPerDestination_HeaderRightClick(ByVal lCol As Long, ByVal Shift As Integer, ByVal X As Long, ByVal Y As Long)

    PopupMenu mnuHdrPopUp
    
End Sub

Private Sub grdSummaryPerDestination_KeyDown(KeyCode As Integer, Shift As Integer, bDoDefault As Boolean)

    'Destinations Grid
        
    If KeyCode = vbKeySpace And grdSummaryPerDestination.RowCount > 0 Then
        
        'Toggle selected line
        grdSummaryPerDestination.CellIcon(grdSummaryPerDestination.CurRow, "Selected") = lstIconList.ItemIndex(SelectRow(grdSummaryPerDestination, 2, KeyCode, grdSummaryPerDestination.CurRow, "DestinationDescription"))
        
        'Customers
        CalculateSummaryPerCustomerForSelectedDestinations
        
        'Routes
        CalculateSummaryPerRouteForSelectedDestinationsAndCustomers
        
        'Drivers
        CalculateSummaryPerDriverForSelectedDestinationsAndCustomersAndRoutes
        
        'Toggle checkbox
        chkAllDestinations.Value = ToggleCheckBox(grdSummaryPerDestination, chkAllDestinations.Value)
        
        'Main grid
        RefreshList
    
    End If
 
End Sub

Private Sub grdSummaryPerDriver_ColHeaderMouseEnter(ByVal lCol As Long)

    grdSummaryPerDriver.Header.Buttons = True
    
End Sub

Private Sub grdSummaryPerDriver_ColHeaderMouseLeave(ByVal lCol As Long)

    grdSummaryPerDriver.Header.Buttons = False
    
End Sub


Private Sub grdSummaryPerDriver_DblClick(ByVal lRow As Long, ByVal lCol As Long, bRequestEdit As Boolean)

    If grdSummaryPerRoute.RowCount > 0 Then
        
        'Toggle selected line
        grdSummaryPerDriver.CellIcon(grdSummaryPerDriver.CurRow, "Selected") = lstIconList.ItemIndex(SelectRow(grdSummaryPerDriver, 5, 32, grdSummaryPerDriver.CurRow, "DriverDescription"))
        
        'Toggle checkbox
        chkAllDrivers.Value = ToggleCheckBox(grdSummaryPerDriver, chkAllDrivers.Value)
        
        'Main grid
        RefreshList
    
    End If

End Sub

Private Sub grdSummaryPerDriver_HeaderRightClick(ByVal lCol As Long, ByVal Shift As Integer, ByVal X As Long, ByVal Y As Long)

    PopupMenu mnuHdrPopUp
    
End Sub

Private Sub grdSummaryPerDriver_KeyDown(KeyCode As Integer, Shift As Integer, bDoDefault As Boolean)

    'Drivers Grid
    
    If KeyCode = vbKeySpace And grdSummaryPerRoute.RowCount > 0 Then
        
        'Toggle selected line
        grdSummaryPerDriver.CellIcon(grdSummaryPerDriver.CurRow, "Selected") = lstIconList.ItemIndex(SelectRow(grdSummaryPerDriver, 5, KeyCode, grdSummaryPerDriver.CurRow, "DriverDescription"))
        
        'Toggle checkbox
        chkAllDrivers.Value = ToggleCheckBox(grdSummaryPerDriver, chkAllDrivers.Value)
        
        'Main grid
        RefreshList
    
    End If

End Sub


Private Sub grdSummaryPerRoute_ColHeaderMouseEnter(ByVal lCol As Long)

    grdSummaryPerRoute.Header.Buttons = True
    
End Sub

Private Sub grdSummaryPerRoute_ColHeaderMouseLeave(ByVal lCol As Long)

    grdSummaryPerRoute.Header.Buttons = False
    
End Sub


Private Sub grdSummaryPerRoute_DblClick(ByVal lRow As Long, ByVal lCol As Long, bRequestEdit As Boolean)

    'Routes Grid
    
    If grdSummaryPerRoute.RowCount > 0 Then
        
        'Toggle selected line
        grdSummaryPerRoute.CellIcon(grdSummaryPerRoute.CurRow, "Selected") = lstIconList.ItemIndex(SelectRow(grdSummaryPerRoute, 4, 32, grdSummaryPerRoute.CurRow, "RouteShortDescription"))
        
        'Drivers
        CalculateSummaryPerDriverForSelectedDestinationsAndCustomersAndRoutes
        
        'Toggle checkbox
        chkAllRoutes.Value = ToggleCheckBox(grdSummaryPerRoute, chkAllRoutes.Value)
        
        'Main grid
        RefreshList
    
    End If

End Sub

Private Sub grdSummaryPerRoute_HeaderRightClick(ByVal lCol As Long, ByVal Shift As Integer, ByVal X As Long, ByVal Y As Long)

    PopupMenu mnuHdrPopUp
    
End Sub

Private Sub grdSummaryPerRoute_KeyDown(KeyCode As Integer, Shift As Integer, bDoDefault As Boolean)

    'Routes Grid
    
    If KeyCode = vbKeySpace And grdSummaryPerRoute.RowCount > 0 Then
        
        'Toggle selected line
        grdSummaryPerRoute.CellIcon(grdSummaryPerRoute.CurRow, "Selected") = lstIconList.ItemIndex(SelectRow(grdSummaryPerRoute, 4, KeyCode, grdSummaryPerRoute.CurRow, "RouteShortDescription"))
        
        'Drivers
        CalculateSummaryPerDriverForSelectedDestinationsAndCustomersAndRoutes
        
        'Toggle checkbox
        chkAllRoutes.Value = ToggleCheckBox(grdSummaryPerRoute, chkAllRoutes.Value)
        
        'Main grid
        RefreshList
    
    End If

End Sub

Private Sub mnuΑποθήκευσηΠλάτουςΣτηλών_Click()

    SaveSetting strApplicationName, "Layout Strings", grdCoachesReport.Tag, grdCoachesReport.LayoutCol
    
    SaveSetting strApplicationName, "Layout Strings", "grdCoachesReportSummaryPerDestination", grdSummaryPerDestination.LayoutCol
    SaveSetting strApplicationName, "Layout Strings", "grdCoachesReportSummaryPerCustomer", grdSummaryPerCustomer.LayoutCol
    SaveSetting strApplicationName, "Layout Strings", "grdCoachesReportSummaryPerRoute", grdSummaryPerRoute.LayoutCol
    SaveSetting strApplicationName, "Layout Strings", "grdCoachesReportSummaryPerDriver", grdSummaryPerDriver.LayoutCol

End Sub

Private Function CreateUnicodeFile(strReportTitle, strReportSubTitle1, intReportDetailLines, lngDriverRow)

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
    
    Dim blnMustPrintSeperator As Boolean
    Dim strSeperator As String
    
    'Αρχικές τιμές
    intPageNo = 1
    lngTotalAdults = 0
    lngTotalKids = 0
    lngTotalFree = 0
    lngTotalPersons = 0
    strSeperator = "^"
    
    Open strUnicodeFile For Output As #1

    'Επικεφαλίδες
    PrintHeadings 91, intPageNo, strReportTitle, strReportSubTitle1
    PrintColumnHeadings 1, "ΩΡΑ", 7, "ΣΗΜΕΙΟ ΠΑΡΑΛΑΒΗΣ", 35, "Ε", 38, "Π", 41, "Δ", 45, "Σ", 47, "ΠΕΛΑΤΗΣ", 68, "ΠΑΡΑΤΗΡΗΣΕΙΣ", 89, "Π ^"
    Print #1, "^"
    
    'Εγγραφές
    intProcessedDetailLines = 10
    intPickupPointCount = 0
    
    'Πλέγμα εγγραφών
    With grdCoachesReport
        For lngRow = 1 To .RowCount
            'Αν ο οδηγός είναι ο ίδιος με αυτόν του πλέγματος των συνόλων
            If .CellValue(lngRow, "DriverDescription") = grdSummaryPerDriver.CellValue(lngDriverRow, "DriverDescription") Then
                'Αν το σημείο παραλαβής που βρίσκομαι είναι ίδιο με αυτό που έχω ήδη κρατήσει ή βρίσκομαι στην πρώτη εγγραφή
                If .CellValue(lngRow, "PickupPointHotelDescription") = strPickupPoint Or strPickupPoint = "" Then
                    'Ενημερώνω τη μεταβλητή που κρατάει το σημείο παραλαβής
                    strPickupPoint = .CellValue(lngRow, "PickupPointHotelDescription")
                    'Αυξάνω τα σημεία παραλαβής κατά ένα
                    intPickupPointCount = intPickupPointCount + 1
                Else
                    'Αν το σημείο παραλαβής που βρίσκομαι δεν είναι ίδιο με αυτό που έχω κρατήσει
                    'Αν έχω μετρήσει παραπάνω από ένα σημεία παραλαβής
                    If intPickupPointCount > 1 Then
                        'Τυπώνω τα σύνολα του σημείου παραλαβής
                        Print #1, _
                            Tab(7); "ΣΥΝΟΛΑ " & Left(strPickupPoint, 18); _
                            Tab(36 - Len(format(lngPickupPointAdults, "#,##0"))); IIf(lngPickupPointAdults > 0, format(lngPickupPointAdults, "#,##0"), ""); _
                            Tab(39 - Len(format(lngPickupPointKids, "#,##0"))); IIf(lngPickupPointKids > 0, format(lngPickupPointKids, "#,##0"), ""); _
                            Tab(42 - Len(format(lngPickupPointFree, "#,##0"))); IIf(lngPickupPointFree > 0, format(lngPickupPointFree, "#,##0"), ""); _
                            Tab(46 - Len(format(lngPickupPointPersons, "#,##0"))); IIf(lngPickupPointPersons > 0, format(lngPickupPointPersons, "#,##0"), ""); _
                            Tab(91); strSeperator
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
                'Αν το σημείο παραλαβής της επόμενης γραμμής είναι διαφορετικό από το τρέχον, η γραμμη θα τυπωθεί με διαχωριστική γραμμή
                If lngRow + 1 <= .RowCount Then
                    blnMustPrintSeperator = IIf(.CellValue(lngRow + 1, "PickupPointHotelDescription") <> strPickupPoint, True, False)
                End If
                'Αν είμαι στην τελευταία γραμμή του πλέγματος, η γραμμη θα τυπωθεί με διαχωριστική γραμμή
                If lngRow = .RowCount Then
                    blnMustPrintSeperator = True
                End If
                'Τυπώνω το σημείο παραλαβής που βρίσκομαι
                Print #1, _
                    Tab(1); .CellText(lngRow, "PickupPointTime"); _
                    Tab(7); Left(.CellText(lngRow, "PickupPointHotelDescription"), 20); _
                    Tab(36 - Len((format(.CellText(lngRow, "TransferAdults"), "#,##0")))); format(.CellText(lngRow, "TransferAdults"), "#,##0"); _
                    Tab(39 - Len((format(.CellText(lngRow, "TransferKids"), "#,##0")))); format(.CellText(lngRow, "TransferKids"), "#,##0"); _
                    Tab(42 - Len((format(.CellText(lngRow, "TransferFree"), "#,##0")))); format(.CellText(lngRow, "TransferFree"), "#,##0"); _
                    Tab(46 - Len((format(.CellText(lngRow, "TransferTotal"), "#,##0")))); format(.CellText(lngRow, "TransferTotal"), "#,##0"); _
                    Tab(47); Left(.CellText(lngRow, "CustomerDescription"), 20); _
                    Tab(68); Left(.CellText(lngRow, "TransferRemarks"), 20); _
                    Tab(89); Left(.CellText(lngRow, "DestinationShortDescription"), 2); _
                    Tab(91); IIf(blnMustPrintSeperator, strSeperator, "")
                'Εκτυπωμένες γραμμές
                intProcessedDetailLines = intProcessedDetailLines + 1
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
                'Eject (Y/N)
                GoSub CheckToEject
            End If
        Next lngRow
            
        'Αν έχω μετρήσει παραπάνω από ένα σημεία παραλαβής
        If intPickupPointCount > 1 Then
            'Τυπώνω τα σύνολα του σημείου παραλαβής
            Print #1, _
                Tab(7); "ΣΥΝΟΛΑ " & Left(strPickupPoint, 18); _
                Tab(36 - Len(format(lngPickupPointAdults, "#,##0"))); IIf(lngPickupPointAdults > 0, format(lngPickupPointAdults, "#,##0"), ""); _
                Tab(39 - Len(format(lngPickupPointKids, "#,##0"))); IIf(lngPickupPointKids > 0, format(lngPickupPointKids, "#,##0"), ""); _
                Tab(42 - Len(format(lngPickupPointFree, "#,##0"))); IIf(lngPickupPointFree > 0, format(lngPickupPointFree, "#,##0"), ""); _
                Tab(46 - Len(format(lngPickupPointPersons, "#,##0"))); IIf(lngPickupPointPersons > 0, format(lngPickupPointPersons, "#,##0"), ""); _
                Tab(91); strSeperator
            'Εκτυπωμένες γραμμές
            intProcessedDetailLines = intProcessedDetailLines + 1
        End If
        
        'Τυπώνω τα σύνολα του οδηγού
        Print #1, "", _
            Tab(7); "ΣΥΝΟΛΑ ΟΔΗΓΟΥ "; _
            Tab(36 - Len(format(lngTotalAdults, "#,##0"))); format(lngTotalAdults, "#,##0"); _
            Tab(39 - Len(format(lngTotalKids, "#,##0"))); format(lngTotalKids, "#,##0"); _
            Tab(42 - Len(format(lngTotalFree, "#,##0"))); format(lngTotalFree, "#,##0"); _
            Tab(46 - Len(format(lngTotalPersons, "#,##0"))); format(lngTotalPersons, "#,##0")
        
    End With
    
    Close #1

    CreateUnicodeFile = True

    Exit Function
    
ErrTrap:
    CreateUnicodeFile = False
    DisplayErrorMessage True, Err.Description
    
    Return
    
CheckToEject:
    If intProcessedDetailLines > CInt(intReportDetailLines) Then
        Print #1, ""
        Print #1, Tab(7); "Η ΕΚΤΥΠΩΣΗ ΣΥΝΕΧΙΖΕΤΑΙ..."
        intPageNo = intPageNo + 1
        PrintHeadings 91, intPageNo, strReportTitle, strReportSubTitle1
        PrintColumnHeadings 1, "ΩΡΑ", 7, "ΣΗΜΕΙΟ ΠΑΡΑΛΑΒΗΣ", 35, "Ε", 38, "Π", 41, "Δ", 45, "Σ", 47, "ΠΕΛΑΤΗΣ", 68, "ΠΑΡΑΤΗΡΗΣΕΙΣ", 93, "Π ^"
        Print #1, ""
        Print #1, Tab(7); "ΣΥΝΕΧΕΙΑ ΕΚΤΥΠΩΣΗΣ ΑΠΟ ΠΡΟΗΓΟΥΜΕΝΗ ΣΕΛΙΔΑ..."
        Print #1, "^"
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
        PositionFormButtons
        
    End If

End Sub

Private Function PositionGrids()

    On Error GoTo ErrTrap
    
    grdCoachesReport.Height = Seperator.Top - grdCoachesReport.Top - 255
    
    frmSummaries.Top = Seperator.Top + 150
    frmSummaries.Height = shpBackground.Height - frmSummaries.Top + shpBackground.Top - 150
    
    grdSummaryPerDestination.Height = frmSummaries.Height - grdSummaryPerDestination.Top - 5
    grdSummaryPerCustomer.Height = frmSummaries.Height - grdSummaryPerCustomer.Top - 5
    grdSummaryPerRoute.Height = frmSummaries.Height - grdSummaryPerRoute.Top - 5
    grdSummaryPerDriver.Height = frmSummaries.Height - grdSummaryPerDriver.Top - 5
    
    Exit Function
    
ErrTrap:
    Resume Next
    
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

    If txtCustomerID.text = "" And txtCustomerDescription.text <> "" Then cmdIndex_Click 2: If txtCustomerID.text = "" Then Cancel = True
    
End Sub

Private Sub txtDestinationDescription_Change()

    If txtDestinationDescription.text = "" Then ClearFields txtDestinationID

End Sub

Private Sub txtDestinationDescription_KeyDown(KeyCode As Integer, Shift As Integer)

    If KeyCode = vbKeyF2 Then cmdIndex_Click 1

End Sub

Private Sub txtDestinationDescription_Validate(Cancel As Boolean)

    If txtDestinationID.text = "" And txtDestinationDescription.text <> "" Then cmdIndex_Click 1: If txtDestinationID.text = "" Then Cancel = True
    
End Sub

Private Sub txtDriverDescription_Change()

    If txtDriverDescription.text = "" Then ClearFields txtDriverID

End Sub

Private Sub txtDriverDescription_KeyDown(KeyCode As Integer, Shift As Integer)

    If KeyCode = vbKeyF2 Then cmdIndex_Click 4

End Sub

Private Sub txtDriverDescription_Validate(Cancel As Boolean)

    If txtDriverID.text = "" And txtDriverDescription.text <> "" Then cmdIndex_Click 4: If txtDriverID.text = "" Then Cancel = True
    
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

    If txtDriverIDForRoutes.text = "" And txtDriverDescriptionForRoutes.text <> "" Then cmdIndex_Click 0: If txtDriverIDForRoutes.text = "" Then Cancel = True

End Sub

Private Sub txtPickupPointDescription_Change()

    If txtPickupPointDescription.text = "" Then
        ClearFields txtRouteID, txtPickupPointID
    End If

End Sub

Private Sub txtPickupPointDescription_KeyDown(KeyCode As Integer, Shift As Integer)

    If KeyCode = vbKeyF2 Then cmdIndex_Click 3

End Sub

Private Sub txtPickupPointDescription_Validate(Cancel As Boolean)

    If txtPickupPointID.text = "" And txtPickupPointDescription.text <> "" Then cmdIndex_Click 3: If txtPickupPointID.text = "" Then Cancel = True

End Sub



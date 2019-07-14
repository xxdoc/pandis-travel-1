VERSION 5.00
Object = "{396F7AC0-A0DD-11D3-93EC-00C0DFE7442A}#1.0#0"; "ImageList.ocx"
Object = "{158C2A77-1CCD-44C8-AF42-AA199C5DCC6C}#1.0#0"; "dcButton.ocx"
Object = "{55473EAC-7715-4257-B5EF-6E14EBD6A5DD}#1.0#0"; "ProgressBar.ocx"
Object = "{839D0F5D-B7D7-41B7-A3B4-85D69300B8C1}#98.0#0"; "iGrid300_10Tec.ocx"
Object = "{FFE4AEB4-0D55-4004-ADF2-3C1C84D17A72}#1.0#0"; "userControls.ocx"
Begin VB.Form PersonsLedger 
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
      Left            =   13050
      TabIndex        =   20
      Top             =   7650
      Width           =   4065
      Begin vbalProgBarLib6.vbalProgressBar prgProgressBar 
         Height          =   615
         Left            =   150
         TabIndex        =   21
         TabStop         =   0   'False
         Top             =   375
         Width           =   3765
         _ExtentX        =   6641
         _ExtentY        =   1085
         Picture         =   "PersonsLedger.frx":0000
         ForeColor       =   0
         Appearance      =   0
         BarPicture      =   "PersonsLedger.frx":001C
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
         TabIndex        =   22
         Top             =   75
         Width           =   3765
      End
   End
   Begin VB.Frame frmContainer 
      Appearance      =   0  'Flat
      BackColor       =   &H00808000&
      BorderStyle     =   0  'None
      ForeColor       =   &H80000008&
      Height          =   9615
      Left            =   75
      TabIndex        =   0
      Top             =   75
      Width           =   18990
      Begin VB.Frame frmCriteria 
         BackColor       =   &H00FFC0FF&
         BorderStyle     =   0  'None
         Height          =   2640
         Index           =   1
         Left            =   5550
         TabIndex        =   46
         Top             =   525
         Width           =   7665
         Begin UserControls.newDate mskInvoiceDateIssueFrom 
            Height          =   465
            Index           =   1
            Left            =   1800
            TabIndex        =   47
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
            Index           =   1
            Left            =   3300
            TabIndex        =   48
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
         Begin UserControls.newText txtDestinationDescription 
            Height          =   465
            Index           =   1
            Left            =   1800
            TabIndex        =   49
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
            Index           =   2
            Left            =   6825
            TabIndex        =   50
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
            PicNormal       =   "PersonsLedger.frx":0038
            PicSizeH        =   16
            PicSizeW        =   16
         End
         Begin VB.Label lblLabel 
            BackColor       =   &H000080FF&
            Caption         =   "Εκδοση"
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
            TabIndex        =   54
            Top             =   900
            Width           =   915
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
            Index           =   2
            Left            =   150
            TabIndex        =   53
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
            Index           =   1
            Left            =   0
            TabIndex        =   52
            Top             =   2100
            Width           =   7665
         End
         Begin VB.Shape shpWedge 
            BackColor       =   &H0000FFFF&
            BackStyle       =   1  'Opaque
            BorderStyle     =   0  'Transparent
            FillColor       =   &H00008000&
            Height          =   840
            Index           =   9
            Left            =   0
            Top             =   975
            Visible         =   0   'False
            Width           =   465
         End
         Begin VB.Shape shpWedge 
            BackColor       =   &H0000FFFF&
            BackStyle       =   1  'Opaque
            BorderStyle     =   0  'Transparent
            FillColor       =   &H00008000&
            Height          =   840
            Index           =   8
            Left            =   1350
            Top             =   1050
            Visible         =   0   'False
            Width           =   465
         End
         Begin VB.Shape shpWedge 
            BackColor       =   &H0000FFFF&
            BackStyle       =   1  'Opaque
            BorderStyle     =   0  'Transparent
            FillColor       =   &H00008000&
            Height          =   840
            Index           =   7
            Left            =   7200
            Top             =   975
            Visible         =   0   'False
            Width           =   465
         End
         Begin VB.Shape shpWedge 
            BackColor       =   &H0000FFFF&
            BackStyle       =   1  'Opaque
            BorderStyle     =   0  'Transparent
            FillColor       =   &H00008000&
            Height          =   315
            Index           =   6
            Left            =   2250
            Top             =   525
            Visible         =   0   'False
            Width           =   465
         End
         Begin VB.Label lblLabel 
            BackColor       =   &H000080FF&
            Caption         =   "Προορισμός"
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
            TabIndex        =   51
            Top             =   1425
            Width           =   915
         End
         Begin VB.Shape shpWedge 
            BackColor       =   &H0000FFFF&
            BackStyle       =   1  'Opaque
            BorderStyle     =   0  'Transparent
            FillColor       =   &H00008000&
            Height          =   315
            Index           =   5
            Left            =   5850
            Top             =   1800
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
            Index           =   5
            Left            =   0
            TabIndex        =   55
            Top             =   0
            Width           =   7665
         End
      End
      Begin VB.Frame frmCriteria 
         BackColor       =   &H00FFC0FF&
         BorderStyle     =   0  'None
         Height          =   3165
         Index           =   0
         Left            =   150
         TabIndex        =   31
         Top             =   5550
         Width           =   7665
         Begin UserControls.newText txtPersonDescription 
            Height          =   465
            Left            =   1800
            TabIndex        =   1
            Top             =   825
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
         Begin UserControls.newDate mskInvoiceDateIssueFrom 
            Height          =   465
            Index           =   0
            Left            =   1800
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
         Begin UserControls.newDate mskInvoiceDateIssueTo 
            Height          =   465
            Index           =   0
            Left            =   3300
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
         Begin Dacara_dcButton.dcButton cmdIndex 
            Height          =   465
            Index           =   0
            Left            =   6825
            TabIndex        =   32
            TabStop         =   0   'False
            Top             =   825
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
            PicNormal       =   "PersonsLedger.frx":05D2
            PicSizeH        =   16
            PicSizeW        =   16
         End
         Begin UserControls.newText txtDestinationDescription 
            Height          =   465
            Index           =   0
            Left            =   1800
            TabIndex        =   4
            Top             =   1875
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
            Index           =   1
            Left            =   6825
            TabIndex        =   39
            TabStop         =   0   'False
            Top             =   1875
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
            PicNormal       =   "PersonsLedger.frx":0B6C
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
            Left            =   5850
            Top             =   2325
            Visible         =   0   'False
            Width           =   465
         End
         Begin VB.Label lblLabel 
            BackColor       =   &H000080FF&
            Caption         =   "Προορισμός"
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
            TabIndex        =   40
            Top             =   1950
            Width           =   915
         End
         Begin VB.Shape shpWedge 
            BackColor       =   &H0000FFFF&
            BackStyle       =   1  'Opaque
            BorderStyle     =   0  'Transparent
            FillColor       =   &H00008000&
            Height          =   315
            Index           =   4
            Left            =   2250
            Top             =   525
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
            Left            =   7200
            Top             =   975
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
            Left            =   1350
            Top             =   1050
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
            Top             =   975
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
            TabIndex        =   37
            Top             =   2625
            Width           =   7665
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
            Left            =   2700
            TabIndex        =   36
            Top             =   75
            Width           =   4815
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
            TabIndex        =   35
            Top             =   75
            Width           =   1665
         End
         Begin VB.Label lblLabel 
            BackColor       =   &H000080FF&
            Caption         =   "Επωνυμία"
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
            Index           =   2
            Left            =   450
            TabIndex        =   34
            Top             =   900
            Width           =   915
         End
         Begin VB.Label lblLabel 
            BackColor       =   &H000080FF&
            Caption         =   "Εκδοση"
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
            TabIndex        =   33
            Top             =   1425
            Width           =   915
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
            TabIndex        =   38
            Top             =   0
            Width           =   7665
         End
      End
      Begin VB.Frame frmButtonFrame 
         BackColor       =   &H00FF8080&
         BorderStyle     =   0  'None
         Height          =   690
         Left            =   75
         TabIndex        =   23
         Top             =   8850
         Width           =   10290
         Begin Dacara_dcButton.dcButton cmdButton 
            Height          =   690
            Index           =   0
            Left            =   225
            TabIndex        =   24
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
            Index           =   6
            Left            =   8775
            TabIndex        =   25
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
            Index           =   1
            Left            =   1650
            TabIndex        =   26
            TabStop         =   0   'False
            Top             =   0
            Width           =   1365
            _ExtentX        =   2408
            _ExtentY        =   1217
            ButtonShape     =   3
            ButtonStyle     =   4
            Caption         =   "Επεξεργασία εγγραφής"
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
            Index           =   5
            Left            =   7350
            TabIndex        =   27
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
            Index           =   2
            Left            =   3080
            TabIndex        =   28
            TabStop         =   0   'False
            Top             =   0
            Width           =   1365
            _ExtentX        =   2408
            _ExtentY        =   1217
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
            ForeColor       =   8388736
            PicOpacity      =   0
         End
         Begin Dacara_dcButton.dcButton cmdButton 
            Height          =   690
            Index           =   3
            Left            =   4500
            TabIndex        =   29
            TabStop         =   0   'False
            Top             =   0
            Width           =   1365
            _ExtentX        =   2408
            _ExtentY        =   1217
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
            ForeColor       =   8388736
            PicOpacity      =   0
         End
         Begin Dacara_dcButton.dcButton cmdButton 
            Height          =   690
            Index           =   4
            Left            =   5925
            TabIndex        =   43
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
         Height          =   2940
         Left            =   7875
         TabIndex        =   7
         Top             =   5775
         Width           =   5040
         Begin VB.TextBox txtBatchReport 
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
            TabIndex        =   45
            TabStop         =   0   'False
            Text            =   "999"
            Top             =   1950
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
            TabIndex        =   44
            TabStop         =   0   'False
            Text            =   "BatchReport"
            Top             =   1950
            Width           =   3540
         End
         Begin VB.TextBox Text3 
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
            TabIndex        =   42
            TabStop         =   0   'False
            Text            =   "InvoiceOutDestinationID"
            Top             =   1575
            Width           =   3540
         End
         Begin VB.TextBox txtDestinationID 
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
            TabIndex        =   41
            TabStop         =   0   'False
            Text            =   "999"
            Top             =   1575
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
            TabIndex        =   15
            TabStop         =   0   'False
            Text            =   "PaymentInOrPaymentOut"
            Top             =   450
            Width           =   3540
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
            TabIndex        =   14
            TabStop         =   0   'False
            Text            =   "2"
            Top             =   450
            Width           =   1305
         End
         Begin VB.TextBox Text13 
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
            TabIndex        =   13
            TabStop         =   0   'False
            Text            =   "CustomersOrSuppliers"
            Top             =   825
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
            TabIndex        =   12
            TabStop         =   0   'False
            Text            =   "3"
            Top             =   825
            Width           =   1305
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
            TabIndex        =   11
            TabStop         =   0   'False
            Text            =   "1"
            Top             =   75
            Width           =   1305
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
            TabIndex        =   10
            TabStop         =   0   'False
            Text            =   "Invoices.InvoiceMasterRefersTo"
            Top             =   75
            Width           =   3540
         End
         Begin VB.TextBox txtInvoicePersonID 
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
            TabIndex        =   9
            TabStop         =   0   'False
            Text            =   "4"
            Top             =   1200
            Width           =   1305
         End
         Begin VB.TextBox Text2 
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
            TabIndex        =   8
            TabStop         =   0   'False
            Text            =   "Invoices.InvoicePersonID"
            Top             =   1200
            Width           =   3540
         End
         Begin vbalIml6.vbalImageList lstIconList 
            Left            =   75
            Top             =   2325
            _ExtentX        =   953
            _ExtentY        =   953
            Size            =   2296
            Images          =   "PersonsLedger.frx":1106
            Version         =   131072
            KeyCount        =   2
            Keys            =   "ÿ"
         End
      End
      Begin iGrid300_10Tec.iGrid grdCustomersLedger 
         Height          =   7290
         Left            =   75
         TabIndex        =   5
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
      Begin iGrid300_10Tec.iGrid grdSuppliersLedger 
         Height          =   7290
         Left            =   75
         TabIndex        =   30
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
      Begin iGrid300_10Tec.iGrid grdPersonsIndex 
         Height          =   7290
         Left            =   75
         TabIndex        =   56
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
         TabIndex        =   19
         Top             =   1125
         Width           =   2565
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
         TabIndex        =   18
         Top             =   525
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
         TabIndex        =   17
         Top             =   825
         Width           =   14940
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
         TabIndex        =   16
         Top             =   1125
         Width           =   14940
      End
      Begin VB.Label lblTitle 
         AutoSize        =   -1  'True
         BackColor       =   &H00800000&
         BackStyle       =   0  'Transparent
         Caption         =   "Καρτέλα συναλλασόμενου"
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
         TabIndex        =   6
         Top             =   75
         Width           =   5730
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
Attribute VB_Name = "PersonsLedger"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Dim lngRowCount As Long
Dim blnError As Boolean
Dim blnProcessing As Boolean
Dim printerHasAlreadyBeenSelected As Boolean

'Προοδευτικό υπόλοιπο
    Dim curAccBalance As Currency

'Προηγούμενη περίοδος
    Dim blnSoFarHasData As Boolean
    'Ατομα
    Dim lngAdultsSoFar As Long
    Dim lngKidsSoFar As Long
    Dim lngFreeSoFar As Long
    'Ποσά
    Dim curAdultsAmountSoFar As Currency
    Dim curKidsAmountSoFar As Currency
    Dim curDirectAmountSoFar As Currency
    Dim curDebitSoFar As Currency
    Dim curCreditSoFar As Currency
    Dim curBalanceSoFar As Currency

'Τρέχουσα εγγραφή
    Dim lngAdultsLine As Long
    Dim lngKidsLine As Long
    Dim lngFreeLine As Long
    Dim curAdultsAmountLine As Currency
    Dim curKidsAmountLine As Currency
    Dim curDirectAmountLine As Currency
    Dim curTotalDebitLine As Currency
    Dim curTotalCreditLine As Currency
    Dim curBalanceLine As Currency

'Ζητούμενη περίοδος
    Dim blnPeriodHasData As Boolean
    'Ατομα
    Dim lngAdultsPeriod As Long
    Dim lngKidsPeriod As Long
    Dim lngFreePeriod As Long
    'Ποσά
    Dim curAdultsAmountPeriod As Currency
    Dim curKidsAmountPeriod As Currency
    Dim curDirectAmountPeriod As Currency
    Dim curDebitPeriod As Currency
    Dim curCreditPeriod As Currency
    Dim curBalancePeriod As Currency

'Γενικά σύνολα
    'Ατομα
    Dim lngAdultsGrandTotal As Long
    Dim lngKidsGrandTotal As Long
    Dim lngFreeGrandTotal As Long
    'Ποσά
    Dim curAdultsAmountGrandTotal As Currency
    Dim curKidsAmountGrandTotal As Currency
    Dim curDirectAmountGrandTotal As Currency
    Dim curDebitGrandTotal As Currency
    Dim curCreditGrandTotal As Currency

Private Function AddCurrentLineForExpensesToGrid(rstTransactions As Recordset)

    With grdSuppliersLedger
        .AddRow
        .CellValue(.RowCount, "TrnID") = rstTransactions!InvoiceTrnID
        .CellValue(.RowCount, "Date") = rstTransactions!InvoiceDateIssue
        .CellValue(.RowCount, "InvoiceDetails") = FullInvoice(rstTransactions!CodeShortDescriptionB, rstTransactions!CodeBatch, rstTransactions!InvoiceNo)
        .CellValue(.RowCount, "ExpenseDescription") = IIf(IsNull(rstTransactions!ExpenseCategoryDescription), rstTransactions!CodeDescription, rstTransactions!ExpenseCategoryDescription)
        .CellValue(.RowCount, "Debit") = curTotalDebitLine
        .CellValue(.RowCount, "Credit") = curTotalCreditLine
        .CellValue(.RowCount, "Balance") = curAccBalance
        .CellValue(.RowCount, "MasterRefersTo") = rstTransactions!InvoiceMasterRefersTo
        .CellValue(.RowCount, "SecondaryRefersTo") = rstTransactions!InvoiceSecondaryRefersTo
        lngRowCount = lngRowCount + 1
    End With
    
    InvertColorForNegativeNumbers grdSuppliersLedger, grdSuppliersLedger.RowCount
    
End Function

Private Function AddCurrentLineForSalesToGrid(rstTransactions As Recordset)

    With grdCustomersLedger
        .AddRow
        .CellValue(.RowCount, "TrnID") = rstTransactions!InvoiceTrnID
        .CellValue(.RowCount, "Date") = rstTransactions!InvoiceDateIssue
        .CellValue(.RowCount, "InvoiceDetails") = FullInvoice(rstTransactions!CodeShortDescriptionB, rstTransactions!CodeBatch, rstTransactions!InvoiceNo)
        .CellValue(.RowCount, "Destination") = IIf(IsNull(rstTransactions!DestinationDescription), rstTransactions!CodeDescription, rstTransactions!DestinationDescription)
        .CellValue(.RowCount, "Adults") = lngAdultsLine
        .CellValue(.RowCount, "Kids") = lngKidsLine
        .CellValue(.RowCount, "Free") = lngFreeLine
        .CellValue(.RowCount, "AdultsAmount") = curAdultsAmountLine
        .CellValue(.RowCount, "KidsAmount") = curKidsAmountLine
        .CellValue(.RowCount, "DirectAmount") = curDirectAmountLine
        .CellValue(.RowCount, "Debit") = curTotalDebitLine
        .CellValue(.RowCount, "Credit") = curTotalCreditLine
        .CellValue(.RowCount, "Balance") = curAccBalance
        .CellValue(.RowCount, "MasterRefersTo") = rstTransactions!InvoiceMasterRefersTo
        .CellValue(.RowCount, "SecondaryRefersTo") = rstTransactions!InvoiceSecondaryRefersTo
        lngRowCount = lngRowCount + 1
    End With
    
    InvertColorForNegativeNumbers grdCustomersLedger, grdCustomersLedger.RowCount

End Function

Private Function AddCurrentLineToGrid(rstTransactions As Recordset)

    If txtInvoiceMasterRefersTo.text = "1" Then AddCurrentLineForExpensesToGrid rstTransactions 'Αγορές
    If txtInvoiceMasterRefersTo.text = "2" Then AddCurrentLineForSalesToGrid rstTransactions 'Πωλήσεις

End Function

Private Function AddTotalsSoFarForExpensesToGrid()

    With grdSuppliersLedger
        grdSuppliersLedger.AddRow
        .CellValue(.RowCount, "ExpenseDescription") = "ΠΡΟΗΓΟΥΜΕΝΗ ΠΕΡΙΟΔΟΣ"
        .CellValue(.RowCount, "Debit") = curDebitSoFar
        .CellValue(.RowCount, "Credit") = curCreditSoFar
        .CellValue(.RowCount, "Balance") = curAccBalance
        grdSuppliersLedger.AddRow
    End With
    
    InvertColorForNegativeNumbers grdSuppliersLedger, grdSuppliersLedger.RowCount - 1

End Function

Private Function AddTotalsSoFarForSalesToGrid()

    With grdCustomersLedger
        grdCustomersLedger.AddRow
        .CellValue(.RowCount, "Destination") = "ΠΡΟΗΓΟΥΜΕΝΗ ΠΕΡΙΟΔΟΣ"
        .CellValue(.RowCount, "Adults") = lngAdultsSoFar
        .CellValue(.RowCount, "Kids") = lngKidsSoFar
        .CellValue(.RowCount, "Free") = lngFreeSoFar
        .CellValue(.RowCount, "AdultsAmount") = curAdultsAmountSoFar
        .CellValue(.RowCount, "KidsAmount") = curKidsAmountSoFar
        .CellValue(.RowCount, "DirectAmount") = curDirectAmountSoFar
        .CellValue(.RowCount, "Debit") = curDebitSoFar
        .CellValue(.RowCount, "Credit") = curCreditSoFar
        .CellValue(.RowCount, "Balance") = curAccBalance
        grdCustomersLedger.AddRow
    End With
    
    InvertColorForNegativeNumbers grdCustomersLedger, grdCustomersLedger.RowCount - 1

End Function

Private Function CalculateCurrentLineForExpenses(rstTransactions As Recordset)

    curTotalDebitLine = 0
    curTotalCreditLine = 0
    
    With rstTransactions
        If !CodeSuppliers = "+" Then curTotalCreditLine = !InvoiceInAmount 'Αν το υπόλοιπο του προμηθευτή αυξάνεται, βάζω το ποσό στην πίστωση
        If !CodeSuppliers = "-" Then curTotalCreditLine = -!InvoiceInAmount 'Αν το υπόλοιπο του προμηθευτή μειώνεται, βάζω το ποσό στην πίστωση με μείον!
        If !CodeSuppliers = "+" And !PaymentTermCreditID = 0 Then curTotalDebitLine = !InvoiceInAmount 'Αν το υπόλοιπο του προμηθευτή αυξάνεται και πληρώθηκε, βάζω το ποσό και στη χρέωση
        If !CodeSuppliers = "-" And !PaymentTermCreditID = 0 Then curTotalDebitLine = -!InvoiceInAmount 'Αν το υπόλοιπο του προμηθευτή μειώνεται και πληρώθηκε, βάζω το ποσό και στη χρέωση με μείον
    End With

End Function

Private Function CalculateCurrentLineForPaymentsOut(rstTransactions As Recordset)

    With rstTransactions
        If !CodeSuppliers = "-" Then curTotalDebitLine = !Amount
        If !CodeSuppliers = "+" Then curTotalCreditLine = !Amount
    End With

End Function

Private Function CalculateCurrentLineForPaymentsIn(rstTransactions As Recordset)

    With rstTransactions
        lngAdultsLine = 0
        lngKidsLine = 0
        lngFreeLine = 0
        curAdultsAmountLine = 0
        curKidsAmountLine = 0
        curDirectAmountLine = 0
        curTotalDebitLine = 0
        curTotalCreditLine = IIf(!CodeCustomers = "-", !Amount, -!Amount)
    End With
    
End Function

Private Function CalculateCurrentLineForSales(rstTransactions As Recordset)

    'Helper
    Dim curTotals As Currency
    
    curTotalDebitLine = 0
    curTotalCreditLine = 0
    
    With rstTransactions
        'Ατομα
        lngAdultsLine = CalculateFields(rstTransactions, !CodeCustomers, "InvoiceOutAdultsWithTransfer", "InvoiceOutAdultsWithoutTransfer")
        lngKidsLine = CalculateFields(rstTransactions, !CodeCustomers, "InvoiceOutKidsWithTransfer", "InvoiceOutKidsWithoutTransfer")
        lngFreeLine = CalculateFields(rstTransactions, !CodeCustomers, "InvoiceOutFreeWithTransfer", "InvoiceOutFreeWithoutTransfer")
        'Ποσά
        curAdultsAmountLine = CalculateFields(rstTransactions, !CodeCustomers, "InvoiceOutAdultsAmountWithTransfer", "InvoiceOutAdultsAmountWithoutTransfer")
        curKidsAmountLine = CalculateFields(rstTransactions, !CodeCustomers, "InvoiceOutKidsAmountWithTransfer", "InvoiceOutKidsAmountWithoutTransfer")
        curDirectAmountLine = CalculateFields(rstTransactions, !CodeCustomers, "InvoiceOutDirectAmount")
        curTotalDebitLine = curAdultsAmountLine + curKidsAmountLine + curDirectAmountLine
        'Αν η κίνηση είναι μετρητοίς βάζω το ποσό και στην πίστωση
        curTotalCreditLine = IIf(!PaymentTermCreditID = 0, curTotalDebitLine, 0)
    End With
    
End Function

Private Function CalculatePeriodTotals(rstTransactions As Recordset)

    If rstTransactions!InvoiceMasterRefersTo = "1" Or rstTransactions!InvoiceMasterRefersTo = "3" Then CalculatePeriodTotalsForExpenses rstTransactions 'Expenses
    If rstTransactions!InvoiceMasterRefersTo = "2" Or rstTransactions!InvoiceMasterRefersTo = "4" Then CalculatePeriodTotalsForSales rstTransactions 'Sales

End Function

Private Function CalculatePeriodTotalsForExpenses(rstTransactions As Recordset)
    
    'Helper
    Dim curTotals As Currency
    
    With rstTransactions
        'Εξοδα
        If !InvoiceMasterRefersTo = "1" Then
            'Helper
            curTotals = CalculateFields(rstTransactions, !CodeSuppliers, "InvoiceInAmount")
            'Πίστωση
            curCreditPeriod = curCreditPeriod + curTotals
            'Αν η κίνηση είναι μετρητοίς βάζω το ποσό και στη χρέωση
            curDebitPeriod = IIf(!PaymentTermCreditID = 0, curDebitPeriod + curTotals, curDebitPeriod)
        End If
        'Πληρωμή ή πιστωτική απογραφή
        If rstTransactions!InvoiceMasterRefersTo = "3" Then
            'Κίνηση αυξάνει το υπόλοιπο
            If !CodeSuppliers = "+" Then
                'Αυξάνω την πίστωση
                curCreditPeriod = curCreditPeriod + curTotalCreditLine
            End If
            'Κίνηση μειώνει το υπόλοιπο
            If !CodeSuppliers = "-" Then
                'Αυξάνω τη χρέωση
                curDebitPeriod = curDebitPeriod + curTotalDebitLine
            End If
        End If
    End With
    
    curBalancePeriod = curBalancePeriod + curTotalDebitLine - curTotalCreditLine
    
    curTotalDebitLine = 0
    curTotalCreditLine = 0

End Function

Private Function CalculateSoFarTotalsForExpenses(rstTransactions As Recordset)

    'Helper
    Dim curTotals As Currency
    
    With rstTransactions
        'Εξοδο - Στήλη πίστωσης
        If !InvoiceMasterRefersTo = "1" Then
            'Helper
            curTotals = CalculateFields(rstTransactions, !CodeSuppliers, "InvoiceInAmount")
            'Ποσό
            curCreditSoFar = curCreditSoFar + IIf(!CodeSuppliers = "+", !InvoiceInAmount, -!InvoiceInAmount)
            'curCreditSoFar = curCreditSoFar + !InvoiceInAmount
            'Αν η κίνηση είναι μετρητοίς βάζω το ποσό και στη χρέωση
            curDebitSoFar = IIf(!PaymentTermCreditID = 0, curDebitSoFar + Abs(curTotals), curDebitSoFar)
        End If
        'Πληρωμή - Στήλη χρέωσης
        If !InvoiceMasterRefersTo = "3" Then
            'Helper
            curTotals = CalculateFields(rstTransactions, !CodeSuppliers, "Amount")
            If !CodeSuppliers = "+" Then
                'Αυξάνω την πίστωση
                curCreditSoFar = curCreditSoFar + curTotals
            End If
            If !CodeSuppliers = "-" Then
                'Αυξάνω τη χρέωση
                curDebitSoFar = curDebitSoFar + Abs(curTotals)
            End If
        End If
    End With

End Function

Private Function CalculateSoFarTotalsForSales(rstTransactions As Recordset)

    'Helper
    Dim curTotals As Currency
    
    With rstTransactions
        'Πώληση - Στήλη χρέωσης
        If !InvoiceMasterRefersTo = "2" Then
            'Helper
            curTotals = CalculateFields(rstTransactions, !CodeCustomers, "InvoiceOutAdultsAmountWithTransfer", "InvoiceOutAdultsAmountWithoutTransfer", "InvoiceOutKidsAmountWithTransfer", "InvoiceOutKidsAmountWithoutTransfer", "InvoiceOutDirectAmount")
            'Ατομα
            lngAdultsSoFar = lngAdultsSoFar + CalculateFields(rstTransactions, !CodeCustomers, "InvoiceOutAdultsWithTransfer", "InvoiceOutAdultsWithoutTransfer")
            lngKidsSoFar = lngKidsSoFar + CalculateFields(rstTransactions, !CodeCustomers, "InvoiceOutKidsWithTransfer", "InvoiceOutKidsWithoutTransfer")
            lngFreeSoFar = lngFreeSoFar + CalculateFields(rstTransactions, !CodeCustomers, "InvoiceOutFreeWithTransfer", "InvoiceOutFreeWithoutTransfer")
            'Ποσά
            curAdultsAmountSoFar = curAdultsAmountSoFar + CalculateFields(rstTransactions, !CodeCustomers, "InvoiceOutAdultsAmountWithTransfer", "InvoiceOutAdultsAmountWithoutTransfer")
            curKidsAmountSoFar = curKidsAmountSoFar + CalculateFields(rstTransactions, !CodeCustomers, "InvoiceOutKidsAmountWithTransfer", "InvoiceOutKidsAmountWithoutTransfer")
            'Απευθείας χρέωση
            curDirectAmountSoFar = curDirectAmountSoFar + CalculateFields(rstTransactions, !CodeCustomers, "InvoiceOutDirectAmount")
            'Χρέωση
            curDebitSoFar = curDebitSoFar + curTotals
            'Αν η κίνηση είναι μετρητοίς βάζω το ποσό και στην πίστωση
            curCreditSoFar = IIf(!PaymentTermCreditID = 0, curCreditSoFar + curTotals, curCreditSoFar)
        End If
        'Είσπραξη - Στήλη πίστωσης
        If !InvoiceMasterRefersTo = "4" Then
            If !CodeCustomers = "+" Then
                'Αυξάνω τη χρέωση
                curDebitSoFar = curDebitSoFar + CalculateFields(rstTransactions, !CodeCustomers, "Amount")
            End If
            If !CodeCustomers = "-" Then
                'Αυξάνω την πίστωση
                curCreditSoFar = curCreditSoFar + Abs(CalculateFields(rstTransactions, !CodeCustomers, "Amount"))
            End If
        End If
    End With

End Function

Private Function CalculateCurrentLine(rstTransactions As Recordset)

    
    If rstTransactions!InvoiceMasterRefersTo = "1" Then CalculateCurrentLineForExpenses rstTransactions 'Εξοδα
    If rstTransactions!InvoiceMasterRefersTo = "2" Then CalculateCurrentLineForSales rstTransactions 'Πωλήσεις
    
    If rstTransactions!InvoiceMasterRefersTo = "3" Then CalculateCurrentLineForPaymentsOut rstTransactions 'Πληρωμές
    If rstTransactions!InvoiceMasterRefersTo = "4" Then CalculateCurrentLineForPaymentsIn rstTransactions 'Εισπράξεις
    
    'Υπόλοιπο γραμμής
    curBalanceLine = curTotalDebitLine - curTotalCreditLine
    
    'Μηδενίζω
    'curTotalDebitLine = 0
    'curTotalCreditLine = 0
    
    'Προοδευτικό υπόλοιπο
    curAccBalance = curAccBalance + curBalanceLine

End Function

Private Function CalculateGrandTotals()

    lngAdultsGrandTotal = lngAdultsSoFar + lngAdultsPeriod
    lngKidsGrandTotal = lngKidsSoFar + lngKidsPeriod
    lngFreeGrandTotal = lngFreeSoFar + lngFreePeriod
    
    curAdultsAmountGrandTotal = curAdultsAmountSoFar + curAdultsAmountPeriod
    curKidsAmountGrandTotal = curKidsAmountSoFar + curKidsAmountPeriod
    curDirectAmountGrandTotal = curDirectAmountSoFar + curDirectAmountPeriod
    curDebitGrandTotal = curDebitSoFar + curDebitPeriod
    curCreditGrandTotal = curCreditSoFar + curCreditPeriod
    
End Function

Private Function CalculatePeriodTotalsForSales(rstTransactions As Recordset)

    Dim curTotals As Currency
    
    With rstTransactions
        If !InvoiceMasterRefersTo = "2" Then
            'Helper
            curTotals = CalculateFields(rstTransactions, !CodeCustomers, "InvoiceOutAdultsAmountWithTransfer", "InvoiceOutAdultsAmountWithoutTransfer", "InvoiceOutKidsAmountWithTransfer", "InvoiceOutKidsAmountWithoutTransfer", "InvoiceOutDirectAmount")
            'Ατομα
            lngAdultsPeriod = lngAdultsPeriod + CalculateFields(rstTransactions, !CodeCustomers, "InvoiceOutAdultsWithTransfer", "InvoiceOutAdultsWithoutTransfer")
            lngKidsPeriod = lngKidsPeriod + CalculateFields(rstTransactions, !CodeCustomers, "InvoiceOutKidsWithTransfer", "InvoiceOutKidsWithoutTransfer")
            lngFreePeriod = lngFreePeriod + CalculateFields(rstTransactions, !CodeCustomers, "InvoiceOutFreeWithTransfer", "InvoiceOutFreeWithoutTransfer")
            'Ποσά ατόμων
            curAdultsAmountPeriod = curAdultsAmountPeriod + CalculateFields(rstTransactions, !CodeCustomers, "InvoiceOutAdultsAmountWithTransfer", "InvoiceOutAdultsAmountWithoutTransfer")
            curKidsAmountPeriod = curKidsAmountPeriod + CalculateFields(rstTransactions, !CodeCustomers, "InvoiceOutKidsAmountWithTransfer", "InvoiceOutKidsAmountWithoutTransfer")
            'Απευθείας χρέωση
            curDirectAmountPeriod = curDirectAmountPeriod + CalculateFields(rstTransactions, !CodeCustomers, "InvoiceOutDirectAmount")
            'Χρέωση
            curDebitPeriod = curDebitPeriod + curTotals
            'Αν η κίνηση είναι μετρητοίς βάζω το ποσό και στην πίστωση
            curCreditPeriod = IIf(!PaymentTermCreditID = 0, curCreditPeriod + curTotals, curCreditPeriod)
        End If
    
        If !InvoiceMasterRefersTo = "4" Then
            curCreditPeriod = curCreditPeriod + curTotalCreditLine
        End If
    
    End With
    
    curBalancePeriod = curBalancePeriod + curAdultsAmountLine + curKidsAmountLine + curDirectAmountLine - curTotalCreditLine

End Function

Private Function CalculateSoFarTotals(fromDate As String, rstTransactions As Recordset)
    
    'Ατομα
    lngAdultsSoFar = 0
    lngKidsSoFar = 0
    lngFreeSoFar = 0
    
    'Ποσά
    curAdultsAmountSoFar = 0
    curKidsAmountSoFar = 0
    curDirectAmountSoFar = 0
    curDebitSoFar = 0
    curCreditSoFar = 0
    curBalanceSoFar = 0

    'Υπόλοιπο
    curAccBalance = 0
    
    CalculateSoFarTotals = False
    
    With rstTransactions
        While Not .EOF
            If Not blnProcessing Then Exit Function
            If !InvoiceDateIssue < CDate(fromDate) Then
                'Εξοδο (Χρεωστικό ή πιστωτικό) - Στήλη πίστωσης
                If txtInvoiceMasterRefersTo.text = "1" Then CalculateSoFarTotalsForExpenses rstTransactions
                'Πώληση (Χρεωστική ή πιστωτική) - Στήλη χρέωσης
                If txtInvoiceMasterRefersTo.text = "2" Then CalculateSoFarTotalsForSales rstTransactions
                'Εχω εγγραφές!
                CalculateSoFarTotals = True
                'Επόμενη εγγραφή
                rstTransactions.MoveNext
                'Async!
                DoEvents
                'Πρόοδος
                UpdateProgressBar Me
            Else
                curAccBalance = curDebitSoFar - curCreditSoFar
                Exit Function
            End If
        Wend
        'Υπόλοιπο
        curAccBalance = curDebitSoFar - curCreditSoFar
    End With

End Function

Private Function CreateUnicodeFileForSuppliers(strReportTitle, strReportSubTitle1, intReportDetailLines)

    'Εκτυπωτής
    Dim lngRow As Long
    Dim intProcessedDetailLines As Integer
    Dim intPageNo As Integer
    
    'Μετρητές
    Dim curDebit As Currency
    Dim curCredit As Currency
    Dim curBalance As Currency

    intPageNo = 0
    intProcessedDetailLines = 0
    
    Open strUnicodeFile For Output As #1
    GoSub Headers
    
    'Πλέγμα
    With grdSuppliersLedger
        For lngRow = 1 To .RowCount
            
            'Εκτυπώνω τη γραμμή
            Print #1, _
                format(.CellText(lngRow, "Date"), "dd/mm/yy"); _
                Tab(10); .CellText(lngRow, "InvoiceDetails"); _
                Tab(24); .CellText(lngRow, "ExpenseDescription"); _
                Tab(105 - Len((format(.CellText(lngRow, "Debit"), "#,##0.00")))); format(.CellText(lngRow, "Debit"), "#,##0.00"); _
                Tab(116 - Len((format(.CellText(lngRow, "Credit"), "#,##0.00")))); format(.CellText(lngRow, "Credit"), "#,##0.00"); _
                Tab(128 - Len((format(.CellText(lngRow, "Balance"), "#,##0.00")))); format(.CellText(lngRow, "Balance"), "#,##0.00")
            
            'Σύνολα
            If .CellText(lngRow, "TrnID") <> "" Then
                curDebit = curDebit + .CellValue(lngRow, "Debit")
                curCredit = curCredit + .CellValue(lngRow, "Credit")
                curBalance = curDebit - curCredit
            End If
            
            intProcessedDetailLines = intProcessedDetailLines + 1
            
            'Eject
            If intProcessedDetailLines > intReportDetailLines Then
                Print #1, ""
                Print #1, Space(23) & "ΣΕ ΜΕΤΑΦΟΡΑ"; _
                Tab(105 - Len(format(curDebit, "#,##0.00"))); format(curDebit, "#,##0.00"); _
                Tab(116 - Len(format(curCredit, "#,##0.00"))); format(curCredit, "#,##0.00"); _
                Tab(128 - Len(format(curBalance, "#,##0.00"))); format(curBalance, "#,##0.00")
                
                GoSub Headers
                
                Print #1, Space(23) & "ΑΠΟ ΜΕΤΑΦΟΡΑ"; _
                    Tab(105 - Len(format(curDebit, "#,##0.00"))); format(curDebit, "#,##0.00"); _
                    Tab(116 - Len(format(curCredit, "#,##0.00"))); format(curCredit, "#,##0.00"); _
                    Tab(128 - Len(format(curBalance, "#,##0.00"))); format(curBalance, "#,##0.00")
                Print #1, ""
                intProcessedDetailLines = intProcessedDetailLines + 2
            End If
            
        Next lngRow
    End With
    
    Close #1
    
    CreateUnicodeFileForSuppliers = True
    
    Exit Function
    
Headers:

    intPageNo = intPageNo + 1
    PrintHeadings 127, intPageNo, strReportTitle, strReportSubTitle1
    PrintColumnHeadings 10, "ΣΤΟΙΧΕΙΟ"
    PrintColumnHeadings 1, "ΗΜΕΡ/ΝΙΑ", 10, "ΣΕΙΡΑ - Νο", 24, "ΠΕΡΙΓΡΑΦΗ ΕΞΟΔΟΥ", 99, "ΧΡΕΩΣΗ", 109, "ΠΙΣΤΩΣΗ", 120, "ΥΠΟΛΟΙΠΟ"
    Print #1, ""
    intProcessedDetailLines = 7
      
    Return

End Function

Private Function DisplayCustomersOrSuppliersGrid()

    If txtBatchReport = "No" Then
        If txtInvoiceMasterRefersTo.text = "1" Then
            grdCustomersLedger.Visible = False
            grdSuppliersLedger.Visible = True
        End If
        If txtInvoiceMasterRefersTo.text = "2" Then
            grdCustomersLedger.Visible = True
            grdSuppliersLedger.Visible = False
        End If
    Else
        grdCustomersLedger.Visible = False
        grdSuppliersLedger.Visible = False
        grdPersonsIndex.Visible = True
    End If
    
End Function

Private Function DisplayDialog()
    
    frmCriteria(1).Visible = True
    mskInvoiceDateIssueFrom(1).SetFocus
    UpdateButtons Me, 6, 1, 0, 0, 0, 0, 1, 0

End Function

Private Function EditInvoiceInRecord()

    Dim rstRecordset As Recordset
    Dim rstExpensesPerVAT As Recordset
    
    If grdSuppliersLedger.CellValue(grdSuppliersLedger.CurRow, "MasterRefersTo") = "1" Then Set rstRecordset = InvoicesIn.SeekRecord(grdSuppliersLedger.CellValue(grdSuppliersLedger.CurRow, "TrnID"))
    If grdSuppliersLedger.CellValue(grdSuppliersLedger.CurRow, "MasterRefersTo") = "3" Then Set rstRecordset = PersonsTransactions.SeekRecord(grdSuppliersLedger.CellValue(grdSuppliersLedger.CurRow, "TrnID"), txtPaymentInOrPaymentOut.text, txtCustomersOrSuppliers.text)
                
    If rstRecordset.RecordCount = 0 Then
        If MyMsgBox(4, strApplicationName, strStandardMessages(9), 1) Then
        End If
        Exit Function
    End If
    
    If grdSuppliersLedger.CellValue(grdSuppliersLedger.CurRow, "MasterRefersTo") = "1" Then
        Set rstExpensesPerVAT = InvoicesIn.FindExpensesPerVAT(grdSuppliersLedger.CellValue(grdSuppliersLedger.CurRow, "TrnID"))
        If rstExpensesPerVAT.RecordCount = 0 Then
            If MyMsgBox(4, strApplicationName, strStandardMessages(9), 1) Then
            End If
            Exit Function
        End If
        InvoicesIn.DoPostFoundJobs rstRecordset, rstExpensesPerVAT
        InvoicesIn.Show 1, Me
    End If
    
    If grdSuppliersLedger.CellValue(grdSuppliersLedger.CurRow, "MasterRefersTo") = "3" Then PersonsTransactions.DoPostFoundJobs rstRecordset, txtPaymentInOrPaymentOut.text, txtCustomersOrSuppliers.text: PersonsTransactions.Show 1, Me

End Function

Private Function EditInvoiceOutRecord()

    Dim rstRecordset As Recordset
    
    If grdCustomersLedger.CellValue(grdCustomersLedger.CurRow, "MasterRefersTo") = "2" Then Set rstRecordset = InvoicesOut.SeekRecord(grdCustomersLedger.CellValue(grdCustomersLedger.CurRow, "TrnID"))
    If grdCustomersLedger.CellValue(grdCustomersLedger.CurRow, "MasterRefersTo") = "4" Then Set rstRecordset = PersonsTransactions.SeekRecord(grdCustomersLedger.CellValue(grdCustomersLedger.CurRow, "TrnID"), txtPaymentInOrPaymentOut.text, txtCustomersOrSuppliers.text)
                
    If rstRecordset.RecordCount = 0 Then
        If MyMsgBox(4, strApplicationName, strStandardMessages(9), 1) Then
        End If
        Exit Function
    End If
    
    If grdCustomersLedger.CellValue(grdCustomersLedger.CurRow, "MasterRefersTo") = "2" Then InvoicesOut.DoPostFoundJobs rstRecordset: InvoicesOut.Show 1, Me
    If grdCustomersLedger.CellValue(grdCustomersLedger.CurRow, "MasterRefersTo") = "4" Then PersonsTransactions.DoPostFoundJobs rstRecordset, txtPaymentInOrPaymentOut.text, txtCustomersOrSuppliers.text: PersonsTransactions.Show 1, Me

End Function

Private Function FindRecordsAndPopulateGrid()

    If ValidateFields Then
        If RefreshList(txtInvoicePersonID.text, mskInvoiceDateIssueFrom(0).text, mskInvoiceDateIssueTo(0).text, txtDestinationDescription(0).text) > 0 Then
            UpdateRecordCount lblRecordCount, lngRowCount
            UpdateCriteriaLabels mskInvoiceDateIssueFrom(0).text, mskInvoiceDateIssueTo(0).text, txtPersonDescription.text, txtDestinationDescription(0).text
            If txtInvoiceMasterRefersTo.text = "1" Then
                EnableGrid grdSuppliersLedger, False
                HighlightRow grdSuppliersLedger, 1, 1, "", True
                UpdateButtons Me, 6, 0, ChangeEditButtonStatus(grdSuppliersLedger, Me.Tag, 1, 1), 1, 1, 1, 1, 0
            End If
            If txtInvoiceMasterRefersTo.text = "2" Then
                EnableGrid grdCustomersLedger, False
                HighlightRow grdCustomersLedger, 1, 1, "", True
                UpdateButtons Me, 6, 0, ChangeEditButtonStatus(grdCustomersLedger, Me.Tag, 1, 1), 1, 1, 1, 1, 0
            End If
            Exit Function
        Else
            UpdateButtons Me, 6, 1, 0, 0, 0, 0, 0, 1
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
            txtPersonDescription.SetFocus
        End If
    End If

End Function

Private Function AddGrandTotalsToGrid()
    
    If txtInvoiceMasterRefersTo.text = "1" Then
        With grdSuppliersLedger
            .AddRow: .RowHeight(.RowCount) = 5: .AddRow
            .CellValue(.RowCount, "ExpenseDescription") = "ΓΕΝΙΚΑ ΣΥΝΟΛΑ"
            .CellValue(.RowCount, "Debit") = curDebitGrandTotal
            .CellValue(.RowCount, "Credit") = curCreditGrandTotal
            .CellValue(.RowCount, "Balance") = curAccBalance
        End With
    End If
    
    InvertColorForNegativeNumbers grdSuppliersLedger, grdSuppliersLedger.RowCount

    If txtInvoiceMasterRefersTo.text = "2" Then
        With grdCustomersLedger
            .AddRow: .RowHeight(.RowCount) = 5: .AddRow
            .CellValue(.RowCount, "Destination") = "ΓΕΝΙΚΑ ΣΥΝΟΛΑ"
            .CellValue(.RowCount, "Adults") = lngAdultsGrandTotal
            .CellValue(.RowCount, "Kids") = lngKidsGrandTotal
            .CellValue(.RowCount, "Free") = lngFreeGrandTotal
            .CellValue(.RowCount, "AdultsAmount") = curAdultsAmountGrandTotal
            .CellValue(.RowCount, "KidsAmount") = curKidsAmountGrandTotal
            .CellValue(.RowCount, "DirectAmount") = curDirectAmountGrandTotal
            .CellValue(.RowCount, "Debit") = curDebitGrandTotal
            .CellValue(.RowCount, "Credit") = curCreditGrandTotal
            .CellValue(.RowCount, "Balance") = curAccBalance
        End With
    End If
    
    InvertColorForNegativeNumbers grdCustomersLedger, grdCustomersLedger.RowCount

End Function

Private Function AddPeriodTotalsToGrid()

    With grdCustomersLedger
        grdCustomersLedger.AddRow
        .CellValue(.RowCount, "Destination") = "ΖΗΤΟΥΜΕΝΗ ΠΕΡΙΟΔΟΣ"
        .CellValue(.RowCount, "Adults") = lngAdultsPeriod
        .CellValue(.RowCount, "Kids") = lngKidsPeriod
        .CellValue(.RowCount, "Free") = lngFreePeriod
        .CellValue(.RowCount, "AdultsAmount") = curAdultsAmountPeriod
        .CellValue(.RowCount, "KidsAmount") = curKidsAmountPeriod
        .CellValue(.RowCount, "DirectAmount") = curDirectAmountPeriod
        .CellValue(.RowCount, "Debit") = curDebitPeriod
        .CellValue(.RowCount, "Credit") = curCreditPeriod
        .CellValue(.RowCount, "Balance") = curBalancePeriod
    End With
    
    InvertColorForNegativeNumbers grdCustomersLedger, grdCustomersLedger.RowCount
    
    With grdSuppliersLedger
        grdSuppliersLedger.AddRow
        grdSuppliersLedger.AddRow
        .CellValue(.RowCount, "ExpenseDescription") = "ΖΗΤΟΥΜΕΝΗ ΠΕΡΙΟΔΟΣ"
        .CellValue(.RowCount, "Debit") = curDebitPeriod
        .CellValue(.RowCount, "Credit") = curCreditPeriod
        .CellValue(.RowCount, "Balance") = curBalancePeriod
    End With
    
    InvertColorForNegativeNumbers grdSuppliersLedger, grdSuppliersLedger.RowCount
    
End Function

Private Function AddTotalsSoFarToGrid()

    If txtInvoiceMasterRefersTo.text = "1" Then AddTotalsSoFarForExpensesToGrid
    If txtInvoiceMasterRefersTo.text = "2" Then AddTotalsSoFarForSalesToGrid
        
End Function

Private Function HideOrDisplayDestinationCriteria()

    If txtInvoiceMasterRefersTo.text = "1" Then
        lblLabel(3).Visible = False
        txtDestinationDescription(0).Visible = False
        cmdIndex(1).Visible = False
        Label1(4).Top = 2100
        frmCriteria(0).Height = 2640
    Else
        lblLabel(3).Visible = True
        txtDestinationDescription(0).Visible = True
        cmdIndex(1).Visible = True
        Label1(4).Top = 2625
        frmCriteria(0).Height = 3165
    End If

End Function

Private Function CheckForSelectedRows()

    Dim lngRow As Long
    
    For lngRow = 1 To grdPersonsIndex.RowCount
        If grdPersonsIndex.CellIcon(lngRow, "Selected") > 0 Then
            CheckForSelectedRows = True
            Exit Function
        End If
    Next lngRow
    
End Function

Private Function PopulatePersonsIndex()

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
    With grdPersonsIndex
        .Clear
        .Editable = False
        .Redraw = False
        .RowMode = True
    End With
    
    'Κυρίως διαδικασία
    strSQL = "SELECT " _
        & "ID, Description, Profession, Address, Phones, PersonInCharge, Email, TaxNo, TaxOfficeID, VATStateID, AccountCode " _
        & "FROM " & txtCustomersOrSuppliers.text & " "

    'Ταξινόμηση
    strOrder = " ORDER BY Description"
    
    Set TempQuery = CommonDB.CreateQueryDef("")
    
    TempQuery.SQL = strSQL & strOrder
    
    'Ανοίγω το recordset
    Set rstRecordset = TempQuery.OpenRecordset()
    
    'Αν δεν έχω εγγραφές, βγαίνω
    If rstRecordset.RecordCount = 0 Then blnError = False: PopulatePersonsIndex = False: Exit Function
    
    'Προετοιμάζω τη μπάρα προόδου
    InitializeProgressBar Me, strApplicationName, rstRecordset
    
    'Προσωρινά
    UpdateButtons Me, 6, 0, 0, 0, 0, 0, 1, 0
    cmdButton(5).Caption = "Διακοπή επεξεργασίας"
    blnProcessing = True
    
    'Γεμίζω το πλέγμα
    With rstRecordset
        grdPersonsIndex.AddRow , , , , , , , rstRecordset.RecordCount
        lngRowCount = rstRecordset.RecordCount
        Do While Not .EOF
            lngRow = lngRow + 1
            UpdateProgressBar Me
            grdPersonsIndex.CellValue(lngRow, "ID") = !ID
            grdPersonsIndex.CellValue(lngRow, "Description") = !Description
            grdPersonsIndex.CellValue(lngRow, "Profession") = !Profession
            grdPersonsIndex.CellValue(lngRow, "Address") = !Address
            grdPersonsIndex.CellValue(lngRow, "Phones") = !Phones
            grdPersonsIndex.CellValue(lngRow, "PersonInCharge") = !PersonInCharge
            grdPersonsIndex.CellValue(lngRow, "Email") = !Email
            grdPersonsIndex.CellValue(lngRow, "TaxNo") = !TaxNo
            grdPersonsIndex.CellValue(lngRow, "TaxOfficeID") = !TaxOfficeID
            grdPersonsIndex.CellValue(lngRow, "VATStateID") = !VATStateID
            grdPersonsIndex.CellValue(lngRow, "AccountCode") = !AccountCode
            .MoveNext
        Loop
    End With
    
    'Ακύρωση επεξεργασίας
    If Not blnProcessing Then
        blnProcessing = True
        ClearFields grdPersonsIndex
        PopulatePersonsIndex = 0
    Else
        PopulatePersonsIndex = lngRowCount
        blnProcessing = False
        grdPersonsIndex.Redraw = True
    End If
    
    'Τελικές ενέργειες
    cmdButton(5).Caption = "Νέα αναζήτηση"
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
    ClearFields grdPersonsIndex, frmProgress
    DisplayErrorMessage True, Err.Description

End Function

Private Function ProcessIndex(printOrCreatePDF As String)

    Dim lngRow As Long
    
    If ValidateFields Then
        If printOrCreatePDF = "Print" Then
            If Not SelectPrinter("PrinterPrintsReports") Then Exit Function
            If Not PrinterExists(strPrinterName) Then Exit Function
            GoSub Continue
        Else
            GoSub Continue
        End If
    End If
    
    grdPersonsIndex.SetFocus
    
    If printOrCreatePDF = "CreatePDF" Then
        MyMsgBox 1, strApplicationName, strStandardMessages(8), 1
    End If
    
    Exit Function
    
Continue:
    With grdPersonsIndex
        For lngRow = 1 To .RowCount
            If .CellIcon(lngRow, "Selected") >= 1 Then
                If RefreshList(.CellValue(lngRow, "ID"), mskInvoiceDateIssueFrom(1).text, mskInvoiceDateIssueTo(1).text, txtDestinationDescription(1).text) > 0 Then
                    If printOrCreatePDF = "Print" Then DoReport "Print", txtCustomersOrSuppliers.text, .CellValue(lngRow, "Description"), mskInvoiceDateIssueFrom(1).text, mskInvoiceDateIssueTo(1).text
                    If printOrCreatePDF = "CreatePDF" Then DoReport "CreatePDF", txtCustomersOrSuppliers.text, .CellValue(lngRow, "Description"), mskInvoiceDateIssueFrom(1).text, mskInvoiceDateIssueTo(1).text
                End If
            End If
        Next lngRow
    End With
    
    Return

End Function

Private Function UpdateCriteriaLabels(DateIssueFrom, DateIssueTo, Person, Destination)

    Dim strCriteriaA As String

    strCriteriaA = IIf(DateIssueFrom = "", "Από [ ΟΛΑ ] ", "Από [ " & DateIssueFrom & " ] ")
    strCriteriaA = strCriteriaA & IIf(DateIssueTo = "", "Εως [ ΟΛΑ ] ", "Εως [ " & DateIssueTo & " ] ")
    strCriteriaA = strCriteriaA & IIf(Person = "", "Συναλλασόμενος [ ΟΛΟΙ ] ", "Συναλλασόμενος [ " & Person & " ] ")
    strCriteriaA = strCriteriaA & IIf(Destination = "", "Προορισμός [ ΟΛΟΙ ]", "Προορισμός [ " & Destination & " ]")
    
    lblCriteria.Caption = strCriteriaA
    
End Function

Private Function EditRecord()

    If txtInvoiceMasterRefersTo.text = "1" Then EditInvoiceInRecord:: grdSuppliersLedger.SetFocus
    If txtInvoiceMasterRefersTo.text = "2" Then EditInvoiceOutRecord: grdCustomersLedger.SetFocus
    
End Function

Private Function CreateUnicodeFileForCustomers(strReportTitle, strReportSubTitle1, intReportDetailLines)

    'Εκτυπωτής
    Dim lngRow As Long
    Dim intProcessedDetailLines As Integer
    Dim intPageNo As Integer
    
    'Μετρητές
    Dim intAdults As Integer
    Dim intKids As Integer
    Dim intFree As Integer
    Dim curAdultsAmount As Currency
    Dim curKidsAmount As Currency
    Dim curDebit As Currency
    Dim curCredit As Currency
    Dim curBalance As Currency

    intPageNo = 0
    intProcessedDetailLines = 0
    
    Open strUnicodeFile For Output As #1
    GoSub Headers
    
    'Πλέγμα
    With grdCustomersLedger
        For lngRow = 1 To grdCustomersLedger.RowCount
            
            'Εκτυπώνω τη γραμμή
            Print #1, _
                format(.CellText(lngRow, "Date"), "dd/mm/yy"); _
                Tab(10); .CellText(lngRow, "InvoiceDetails"); _
                Tab(24); Left(.CellText(lngRow, "Destination"), 21); _
                Tab(53 - Len((format(.CellText(lngRow, "Adults"), "#,##0")))); format(.CellText(lngRow, "Adults"), "#,##0"); _
                Tab(61 - Len((format(.CellText(lngRow, "Kids"), "#,##0")))); format(.CellText(lngRow, "Kids"), "#,##0"); _
                Tab(67 - Len((format(.CellText(lngRow, "Free"), "#,##0")))); format(.CellText(lngRow, "Free"), "#,##0"); _
                Tab(81 - Len((format(.CellText(lngRow, "AdultsAmount"), "#,##0.00")))); format(.CellText(lngRow, "AdultsAmount"), "#,##0.00"); _
                Tab(95 - Len((format(.CellText(lngRow, "KidsAmount"), "#,##0.00")))); format(.CellText(lngRow, "KidsAmount"), "#,##0.00"); _
                Tab(109 - Len((format(.CellText(lngRow, "Debit"), "#,##0.00")))); format(.CellText(lngRow, "Debit"), "#,##0.00"); _
                Tab(123 - Len((format(.CellText(lngRow, "Credit"), "#,##0.00")))); format(.CellText(lngRow, "Credit"), "#,##0.00"); _
                Tab(137 - Len((format(.CellText(lngRow, "Balance"), "#,##0.00")))); format(.CellText(lngRow, "Balance"), "#,##0.00")
            
            'Σύνολα
            If .CellText(lngRow, "TrnID") <> "" Then
                intAdults = intAdults + .CellValue(lngRow, "Adults")
                intKids = intKids + .CellValue(lngRow, "Kids")
                intFree = intFree + .CellValue(lngRow, "Free")
                curAdultsAmount = curAdultsAmount + .CellValue(lngRow, "AdultsAmount")
                curKidsAmount = curKidsAmount + .CellValue(lngRow, "KidsAmount")
                curDebit = curDebit + .CellValue(lngRow, "Debit")
                curCredit = curCredit + .CellValue(lngRow, "Credit")
                curBalance = curDebit - curCredit
            End If
            
            intProcessedDetailLines = intProcessedDetailLines + 1
            
            'Eject
            If intProcessedDetailLines > intReportDetailLines Then
                Print #1, ""
                Print #1, Space(23) & "ΣΕ ΜΕΤΑΦΟΡΑ"; _
                Tab(53 - Len(format(intAdults, "#,##0"))); format(intAdults, "#,##0"); _
                Tab(61 - Len(format(intKids, "#,##0"))); format(intKids, "#,##0"); _
                Tab(67 - Len(format(intFree, "#,##0"))); format(intFree, "#,##0"); _
                Tab(81 - Len(format(curAdultsAmount, "#,##0.00"))); format(curAdultsAmount, "#,##0.00"); _
                Tab(95 - Len(format(curKidsAmount, "#,##0.00"))); format(curKidsAmount, "#,##0.00"); _
                Tab(109 - Len(format(curDebit, "#,##0.00"))); format(curDebit, "#,##0.00"); _
                Tab(123 - Len(format(curCredit, "#,##0.00"))); format(curCredit, "#,##0.00"); _
                Tab(137 - Len(format(curBalance, "#,##0.00"))); format(curBalance, "#,##0.00")
                
                GoSub Headers
                
                Print #1, Space(23) & "ΑΠΟ ΜΕΤΑΦΟΡΑ"; _
                    Tab(53 - Len(format(intAdults, "#,##0"))); format(intAdults, "#,##0"); _
                    Tab(61 - Len(format(intKids, "#,##0"))); format(intKids, "#,##0"); _
                    Tab(67 - Len(format(intFree, "#,##0"))); format(intFree, "#,##0"); _
                    Tab(81 - Len(format(curAdultsAmount, "#,##0.00"))); format(curAdultsAmount, "#,##0.00"); _
                    Tab(95 - Len(format(curKidsAmount, "#,##0.00"))); format(curKidsAmount, "#,##0.00"); _
                    Tab(109 - Len(format(curDebit, "#,##0.00"))); format(curDebit, "#,##0.00"); _
                    Tab(123 - Len(format(curCredit, "#,##0.00"))); format(curCredit, "#,##0.00"); _
                    Tab(137 - Len(format(curBalance, "#,##0.00"))); format(curBalance, "#,##0.00")
                Print #1, ""
                intProcessedDetailLines = intProcessedDetailLines + 2
            End If
            
        Next lngRow
    End With
    
    Close #1
    
    CreateUnicodeFileForCustomers = True
    
    Exit Function
    
Headers:
    intPageNo = intPageNo + 1
    PrintHeadings 136, intPageNo, strReportTitle, strReportSubTitle1
    PrintColumnHeadings 10, "ΣΤΟΙΧΕΙΟ", 47, "ΕΝΗΛΙ-", 57, "ΠΑΙ-", 64, "ΔΩ-", 73, "ΧΡΕΩΣΕΙΣ", 87, "ΧΡΕΩΣΕΙΣ", 103, "ΣΥΝΟΛΟ"
    PrintColumnHeadings 1, "ΗΜΕΡ/ΝΙΑ", 10, "ΣΕΙΡΑ - Νο", 24, "ΠΡΟΟΡΙΣΜΟΣ", 50, "ΚΕΣ", 58, "ΔΙΑ", 63, "ΡΕΑΝ", 73, "ΕΝΗΛΙΚΩΝ", 88, "ΠΑΙΔΙΩΝ", 102, "ΧΡΕΩΣΗΣ", 116, "ΠΙΣΤΩΣΗ", 129, "ΥΠΟΛΟΙΠΟ"
    Print #1, ""
    intProcessedDetailLines = 11
      
    Return
    
End Function

Private Sub cmdButton_Click(index As Integer)

    Select Case index
        Case 0
            'Μία καρτέλα
            If txtBatchReport.text = "No" Then FindRecordsAndPopulateGrid
            'Πολλές καρτέλες
            If txtBatchReport.text = "Yes" Then
                If ValidateFields Then
                    If PopulatePersonsIndex > 0 Then
                        UpdateRecordCount lblRecordCount, lngRowCount
                        UpdateCriteriaLabels mskInvoiceDateIssueFrom(1).text, mskInvoiceDateIssueTo(1).text, txtPersonDescription.text, txtDestinationDescription(1).text
                        frmCriteria(1).Visible = False
                        grdPersonsIndex.SetCurCell 1, 1
                        grdPersonsIndex.SetFocus
                        UpdateButtons Me, 6, 0, 0, 1, 1, 1, 1, 0
                    Else
                        MyMsgBox 1, strApplicationName, strStandardMessages(7), 1
                        UpdateButtons Me, 6, 1, 0, 0, 0, 0, 0, 1
                    End If
                End If
            End If
        Case 1
            EditRecord
        Case 2
            'Μία καρτέλα
            If txtBatchReport.text = "No" Then
                If Not SelectPrinter("PrinterPrintsReports") Then Exit Sub
                If Not PrinterExists(strPrinterName) Then Exit Sub
                DoReport "Print", txtCustomersOrSuppliers.text, txtPersonDescription.text, mskInvoiceDateIssueFrom(0).text, mskInvoiceDateIssueTo(0).text
            End If
            'Πολλές καρτέλες
            If txtBatchReport.text = "Yes" Then
                If CheckForSelectedRows Then
                    ProcessIndex "Print"
                    UpdateButtons Me, 6, 0, 0, 1, 1, 1, 1, 0
                Else
                    MyMsgBox 4, strApplicationName, strStandardMessages(6), 1
                    grdPersonsIndex.SetFocus
                End If
            End If
        Case 3
            'Μία καρτέλα
            If txtBatchReport.text = "No" Then
                If DoReport("CreatePDF", txtCustomersOrSuppliers.text, txtPersonDescription.text, mskInvoiceDateIssueFrom(0).text, mskInvoiceDateIssueTo(0).text) Then
                    MyMsgBox 1, strApplicationName, strStandardMessages(8), 1
                End If
            End If
            'Πολλές καρτέλες
            If txtBatchReport.text = "Yes" Then
                If CheckForSelectedRows Then
                    ProcessIndex "CreatePDF"
                    grdPersonsIndex.SetFocus
                    UpdateButtons Me, 6, 0, 0, 1, 1, 1, 1, 0
                Else
                    MyMsgBox 4, strApplicationName, strStandardMessages(6), 1
                    grdPersonsIndex.SetFocus
                End If
            End If
        Case 4
            ExportToExcel
        Case 5
            AbortProcedure False
        Case 6
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
    xlsColCount = 6
    
    With oSheet
    
        SetFontNameAndSize oSheet, "Ubuntu Condensed", 11
        AddCompanyData oSheet, xlsColCount
        AddTitle oSheet, lblTitle.Caption, xlsColCount
        AddCriteria oSheet, lblCriteria.Caption, xlsColCount
        
        If txtCustomersOrSuppliers.text = "Customers" Then
            AddHeaders oSheet, grdCustomersLedger, xlsColCount, "A", "Date", "B", "InvoiceDetails", "C", "Destination", "D", "Adults", "E", "Kids", "F", "Free", "G", "AdultsAmount", "H", "KidsAmount", "I", "DirectAmount", "J", "Debit", "K", "Credit", "L", "Balance"
            AdjustColumnWidths oSheet, "A", 10, "B", 15, "C", 40, "D", 10, "E", 10, "F", 10, "G", 10, "H", 10, "I", 10, "J", 10, "K", 10, "L", 10
            For lngRow = 1 To grdCustomersLedger.RowCount
                .Range("A" & lngRow + xlsRowOffsetFromTop) = grdCustomersLedger.CellValue(lngRow, "Date")
                .Range("B" & lngRow + xlsRowOffsetFromTop) = grdCustomersLedger.CellValue(lngRow, "InvoiceDetails")
                .Range("C" & lngRow + xlsRowOffsetFromTop) = grdCustomersLedger.CellValue(lngRow, "Destination")
                .Range("D" & lngRow + xlsRowOffsetFromTop) = grdCustomersLedger.CellValue(lngRow, "Adults")
                .Range("E" & lngRow + xlsRowOffsetFromTop) = grdCustomersLedger.CellValue(lngRow, "Kids")
                .Range("F" & lngRow + xlsRowOffsetFromTop) = grdCustomersLedger.CellValue(lngRow, "Free")
                .Range("G" & lngRow + xlsRowOffsetFromTop) = grdCustomersLedger.CellValue(lngRow, "AdultsAmount")
                .Range("H" & lngRow + xlsRowOffsetFromTop) = grdCustomersLedger.CellValue(lngRow, "KidsAmount")
                .Range("I" & lngRow + xlsRowOffsetFromTop) = grdCustomersLedger.CellValue(lngRow, "DirectAmount")
                .Range("J" & lngRow + xlsRowOffsetFromTop) = grdCustomersLedger.CellValue(lngRow, "Debit")
                .Range("K" & lngRow + xlsRowOffsetFromTop) = grdCustomersLedger.CellValue(lngRow, "Credit")
                .Range("L" & lngRow + xlsRowOffsetFromTop) = grdCustomersLedger.CellValue(lngRow, "Balance")
            Next lngRow
            AddNumberFormats oSheet, grdCustomersLedger, "Dates", 10, "A"
            AddNumberFormats oSheet, grdCustomersLedger, "Integers", 10, "D", "E", "F"
            AddNumberFormats oSheet, grdCustomersLedger, "Floats", 10, "G", "H", "I", "J", "K", "L"
        End If
    
        If txtCustomersOrSuppliers.text = "Suppliers" Then
            AddHeaders oSheet, grdSuppliersLedger, xlsColCount, "A", "Date", "B", "InvoiceDetails", "C", "ExpenseDescription", "D", "Debit", "E", "Credit", "F", "Balance"
            AdjustColumnWidths oSheet, "A", 10, "B", 15, "C", 40, "D", 10, "E", 10, "F", 10
            For lngRow = 1 To grdSuppliersLedger.RowCount
                .Range("A" & lngRow + xlsRowOffsetFromTop) = grdSuppliersLedger.CellValue(lngRow, "Date")
                .Range("B" & lngRow + xlsRowOffsetFromTop) = grdSuppliersLedger.CellValue(lngRow, "InvoiceDetails")
                .Range("C" & lngRow + xlsRowOffsetFromTop) = grdSuppliersLedger.CellValue(lngRow, "ExpenseDescription")
                .Range("D" & lngRow + xlsRowOffsetFromTop) = grdSuppliersLedger.CellValue(lngRow, "Debit")
                .Range("E" & lngRow + xlsRowOffsetFromTop) = grdSuppliersLedger.CellValue(lngRow, "Credit")
                .Range("F" & lngRow + xlsRowOffsetFromTop) = grdSuppliersLedger.CellValue(lngRow, "Balance")
            Next lngRow
            AddNumberFormats oSheet, grdSuppliersLedger, "Dates", 10, "A"
            AddNumberFormats oSheet, grdSuppliersLedger, "Floats", 10, "D", "E", "F"
        End If
    
    End With
    
    oBook.SaveAs strReportsPathName & lblTitle.Caption & ".xlsx"
    
    oExcel.Quit
    
    MyMsgBox 1, strApplicationName, strStandardMessages(8), 1
    
    Exit Function
    
ErrTrap:
    oBook.Close False
    oExcel.Quit

    If Err.Number = 1004 Then
        MyMsgBox 4, strApplicationName, strStandardMessages(27), 1
    Else
        DisplayErrorMessage True, Err.Description
    End If
    
    Exit Function
    
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
                        .Range(columns(column) & row + rowOffsetFromTop).NumberFormat = "#,##0.00_);[Red]#,##0.00 "
                    Next row
                Case "Integers"
                    For row = 1 To grid.RowCount
                        .Range(columns(column) & row + rowOffsetFromTop).NumberFormat = "#,##0_);[Red]#,##0 "
                    Next row
                Case "Dates"
                    For row = 1 To grid.RowCount
                        .Range(columns(column) & row + rowOffsetFromTop).NumberFormat = "dd-mm-yyyy"
                    Next row
            End Select
        Next column
    End With

End Function

Private Function DoReport(action As String, Persons As String, PersonDescription As String, fromDate As String, toDate As String)
    
    On Error GoTo ErrTrap
    
    If action = "Print" Then
        If Persons = "Customers" Then CreateUnicodeFileForCustomers "ΚΑΡΤΕΛΑ ΠΕΛΑΤΗ " & PersonDescription, " από " & fromDate & " έως " & toDate, intPrinterReportDetailLines - 11
        If Persons = "Suppliers" Then CreateUnicodeFileForSuppliers "ΚΑΡΤΕΛΑ ΠΡΟΜΗΘΕΥΤΗ " & PersonDescription, " από " & fromDate & " έως " & toDate, intPrinterReportDetailLines - 15
        With rptOneLiner
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
    
    If action = "CreatePDF" Then
        If Persons = "Customers" Then
            If CreateUnicodeFileForCustomers("ΚΑΡΤΕΛΑ ΠΕΛΑΤΗ " & PersonDescription, " από " & fromDate & " έως " & toDate, GetSetting(strApplicationName, "Settings", "Export Report Height")) Then
                If CreateUnisexPDF("ΚΑΡΤΕΛΑ ΠΕΛΑΤΗ " & " " & PersonDescription, rptOneLiner, 7) Then
                    DoReport = True
                End If
            End If
        End If
        If Persons = "Suppliers" Then
            If CreateUnicodeFileForSuppliers("ΚΑΡΤΕΛΑ ΠΡΟΜΗΘΕΥΤΗ " & txtPersonDescription.text, " από " & fromDate & " έως " & toDate, GetSetting(strApplicationName, "Settings", "Export Report Height") - 4) Then
                If CreateUnisexPDF("ΚΑΡΤΕΛΑ ΠΡΟΜΗΘΕΥΤΗ " & txtPersonDescription.text, rptOneLiner, 7) Then
                    DoReport = True
                End If
            End If
        End If
    End If
    
    Exit Function
    
ErrTrap:
    Close #1
    DisplayErrorMessage True, Err.Description

End Function

Private Function ValidateFields()

    'Αρχικές τιμές
    ValidateFields = False
    
    'Συναλλασόμενος
    If txtInvoicePersonID.text = "" And txtBatchReport = "No" Then
        If MyMsgBox(4, strApplicationName, strStandardMessages(1), 1) Then
        End If
        txtPersonDescription.SetFocus
        Exit Function
    End If
    
    'Από
    If mskInvoiceDateIssueFrom(0).text = "" And txtBatchReport = "No" Then
        If MyMsgBox(4, strApplicationName, strStandardMessages(1), 1) Then
        End If
        mskInvoiceDateIssueFrom(0).SetFocus
        Exit Function
    End If
    
    'Εως
    If mskInvoiceDateIssueTo(0).text = "" And txtBatchReport = "No" Then
        If MyMsgBox(4, strApplicationName, strStandardMessages(1), 1) Then
        End If
        mskInvoiceDateIssueTo(0).SetFocus
        Exit Function
    End If
    
    'Σωστό διάστημα
    If IsDate(mskInvoiceDateIssueFrom(0).text) And IsDate(mskInvoiceDateIssueTo(0).text) And txtBatchReport = "No" Then
        If CDate(mskInvoiceDateIssueFrom(0).text) > CDate(mskInvoiceDateIssueTo(0).text) Then
            If MyMsgBox(4, strApplicationName, strStandardMessages(10), 1) Then
            End If
            mskInvoiceDateIssueFrom(0).SetFocus
            Exit Function
        End If
    End If
    
    'Από
    If mskInvoiceDateIssueFrom(1).text = "" And txtBatchReport = "Yes" Then
        If MyMsgBox(4, strApplicationName, strStandardMessages(1), 1) Then
        End If
        mskInvoiceDateIssueFrom(1).SetFocus
        Exit Function
    End If
    
    'Εως
    If mskInvoiceDateIssueTo(1).text = "" And txtBatchReport = "Yes" Then
        If MyMsgBox(4, strApplicationName, strStandardMessages(1), 1) Then
        End If
        mskInvoiceDateIssueTo(1).SetFocus
        Exit Function
    End If
    
    'Σωστό διάστημα
    If IsDate(mskInvoiceDateIssueFrom(1).text) And IsDate(mskInvoiceDateIssueTo(1).text) And txtBatchReport = "Yes" Then
        If CDate(mskInvoiceDateIssueFrom(1).text) > CDate(mskInvoiceDateIssueTo(1).text) Then
            If MyMsgBox(4, strApplicationName, strStandardMessages(10), 1) Then
            End If
            mskInvoiceDateIssueFrom(1).SetFocus
            Exit Function
        End If
    End If
    
    ValidateFields = True
    
End Function

Private Function AbortProcedure(blnStatus)

    If blnProcessing Then blnProcessing = False: Exit Function

    If Not blnStatus And txtBatchReport = "No" Then
        ClearFields lblSelectedGridTotals, lblSelectedGridLines, lblCriteria, lblRecordCount
        ClearFields grdCustomersLedger, grdSuppliersLedger
        frmCriteria(0).Visible = True
        txtPersonDescription.SetFocus
        UpdateButtons Me, 6, 1, 0, 0, 0, 0, 0, 1
    End If
    
    If Not blnStatus And txtBatchReport = "Yes" Then
        frmCriteria(1).Visible = True
        ClearFields lblSelectedGridTotals, lblSelectedGridLines, lblCriteria, lblRecordCount
        ClearFields grdPersonsIndex
        mskInvoiceDateIssueFrom(1).SetFocus
        UpdateButtons Me, 6, 1, 0, 0, 0, 0, 0, 1
    End If
    
    If blnStatus Then
        Unload Me
    End If

End Function

Private Function RefreshList(personID As String, fromDate As String, toDate As String, DestinationID As String)

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
    Dim rstTransactions As Recordset

    'Helpers
    Dim strFullInvoice As String

    Dim blnSoFarHasData As Boolean
    Dim blnPeriodHasData As Boolean
    
    'Αρχικές τιμές
    intIndex = 0
    lngRow = 0
    lngRowCount = 0
    
    lngAdultsPeriod = 0
    lngKidsPeriod = 0
    lngFreePeriod = 0
    
    curAdultsAmountPeriod = 0
    curKidsAmountPeriod = 0
    curDirectAmountPeriod = 0
    curDebitPeriod = 0
    curCreditPeriod = 0
    curBalancePeriod = 0
    
    curTotalDebitLine = 0
    curTotalCreditLine = 0
    
    curAccBalance = 0
    
    frmCriteria(0).Visible = False
    blnPeriodHasData = False
    
    'Πλέγμα
    With grdCustomersLedger
        .Clear
        .Redraw = False
    End With
    With grdSuppliersLedger
        .Clear
        .Redraw = False
    End With
    
    'Κυρίως διαδικασία
    strSQL = CreateSELECTStatement(txtInvoiceMasterRefersTo.text)
 
    'Αγορές = 1, Πωλήσεις = 2
    strThisParameter = "strMasterA String"
    strThisQuery = "(Invoices.InvoiceMasterRefersTo = strMasterA"
    strLogic = " AND "
    GoSub UpdateSQLString
    arrQuery(intIndex) = txtInvoiceMasterRefersTo.text
    
    'Πληρωμές = 3, Εισπράξεις = 4
    strThisParameter = "strMasterB String"
    strThisQuery = "Invoices.InvoiceMasterRefersTo = strMasterB)"
    strLogic = " OR "
    GoSub UpdateSQLString
    arrQuery(intIndex) = Trim(Str(Val(txtInvoiceMasterRefersTo.text) + 2))
    
    'Εκδοση Εως
    If toDate <> "" Then
        strThisParameter = "datToDate Date"
        strThisQuery = "InvoiceDateIssue <= datToDate"
        strLogic = " AND "
        GoSub UpdateSQLString
        arrQuery(intIndex) = toDate
    End If
    
    'Συναλλασόμενος
    strThisParameter = "intPersonID Integer"
    strThisQuery = "Invoices.InvoicePersonID = intPersonID"
    strLogic = " AND "
    GoSub UpdateSQLString
    arrQuery(intIndex) = Val(personID)
    
    'Προορισμός (Μόνο για πελάτες)
    If txtInvoiceMasterRefersTo.text = "2" Then
        If txtDestinationID.text <> "" Then
            strThisParameter = "intDestinationID Integer"
            strThisQuery = "InvoiceOutDestinationID = intDestinationID "
            strLogic = " AND "
            GoSub UpdateSQLString
            arrQuery(intIndex) = Val(txtDestinationID.text)
        End If
    End If
    
    'Ταξινόμηση
    strOrder = " ORDER BY InvoiceDateIssue, InvoiceNo, InvoiceID"
    
    Set TempQuery = CommonDB.CreateQueryDef("")
    
    'Προσθέτω τα κριτήρια
    If strThisParameter <> "" Then
        strParameters = "PARAMETERS " & strParameters & "; "
        strParFields = "WHERE " & strParFields
        strSQL = strParameters & strSQL & strParFields
        TempQuery.SQL = strSQL & strOrder
    End If
    
    'Κριτήρια
    If strThisParameter <> "" Then
        For intIndex = 1 To UBound(arrQuery)
            TempQuery.Parameters(intIndex - 1) = arrQuery(intIndex)
        Next intIndex
    End If
    
    'Ανοίγω το recordset
    Set rstTransactions = TempQuery.OpenRecordset()
    
    'Αν δεν έχω εγγραφές, βγαίνω
    If rstTransactions.RecordCount = 0 Then blnError = False: RefreshList = False: Exit Function
    
    'Προετοιμάζω τη μπάρα προόδου
    InitializeProgressBar Me, strApplicationName, rstTransactions
    
    'Προσωρινά
    UpdateButtons Me, 6, 0, 0, 0, 0, 0, 1, 0
    cmdButton(5).Caption = "Διακοπή επεξεργασίας"
    blnProcessing = True
    
    'Γεμίζω το πλέγμα
    With rstTransactions
        If .EOF = False Then
            If CalculateSoFarTotals(fromDate, rstTransactions) Then
                blnSoFarHasData = True
                AddTotalsSoFarToGrid
            End If
            Do While Not .EOF
                If Not blnProcessing Then Exit Do 'Async!
                blnPeriodHasData = True 'We have a live one!
                CalculateCurrentLine rstTransactions 'Υπολογίζω την τρέχουσα γραμμή
                AddCurrentLineToGrid rstTransactions 'Εμφανίζω την τρέχουσα γραμμή
                CalculatePeriodTotals rstTransactions 'Υπολογίζω τα σύνολα περιόδου
                UpdateProgressBar Me 'Πρόοδος
                rstTransactions.MoveNext 'Επόμενη εγγραφή
                DoEvents 'Async!
            Loop
        End If
    End With
    
    'Ακύρωση επεξεργασίας
    If Not blnProcessing Then
        blnProcessing = True
        ClearFields grdCustomersLedger, grdSuppliersLedger
        RefreshList = False
    Else
        RefreshList = lngRowCount
        blnProcessing = False
    End If
    
    'Σύνολα
    If Not blnProcessing Then
        If blnPeriodHasData Then grdCustomersLedger.AddRow
        If blnPeriodHasData Or curAccBalance <> 0 Then
            RefreshList = 1
            AddPeriodTotalsToGrid
            If blnSoFarHasData Then
                CalculateGrandTotals
                AddGrandTotalsToGrid
            End If
        Else
            ClearFields grdCustomersLedger, grdSuppliersLedger
        End If
    End If
        
    'Τελικές ενέργειες
    cmdButton(5).Caption = "Νέα αναζήτηση"
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
    ClearFields grdCustomersLedger, grdSuppliersLedger, frmProgress
    DisplayErrorMessage True, Err.Description
    
End Function
Private Sub cmdIndex_Click(index As Integer)

    'Local variables
    Dim tmpTableData As typTableData
    Dim tmpRecordset As Recordset
    
    Select Case index
        Case 0
            'Customers - F2
            Set tmpRecordset = CheckForMatch("CommonDB", txtCustomersOrSuppliers.text, "Description", "String", txtPersonDescription.text)
            If tmpRecordset.RecordCount > 0 Then
                tmpTableData = DisplayIndex(tmpRecordset, 2, True, 2, 0, 1, "ID", "Επωνυμία", 0, 40, 1, 0)
                txtInvoicePersonID.text = tmpTableData.strCode
                txtPersonDescription.text = tmpTableData.strFirstField
            End If
        Case 1
            'Destination - F2
            Set tmpRecordset = CheckForMatch("CommonDB", "Destinations", "DestinationDescription", "String", txtDestinationDescription(0).text)
            If tmpRecordset.RecordCount > 0 Then
                tmpTableData = DisplayIndex(tmpRecordset, 2, True, 2, 0, 2, "ID", "Περιγραφή", 0, 40, 1, 0)
                txtDestinationID.text = tmpTableData.strCode
                txtDestinationDescription(0).text = tmpTableData.strFirstField
            End If
        Case 2
            'Destination - F2
            Set tmpRecordset = CheckForMatch("CommonDB", "Destinations", "DestinationDescription", "String", txtDestinationDescription(1).text)
            If tmpRecordset.RecordCount > 0 Then
                tmpTableData = DisplayIndex(tmpRecordset, 2, True, 2, 0, 2, "ID", "Περιγραφή", 0, 40, 1, 0)
                txtDestinationID.text = tmpTableData.strCode
                txtDestinationDescription(1).text = tmpTableData.strFirstField
            End If
    End Select

End Sub

Private Sub Form_Activate()
                
    If Me.Tag = "True" Then
        Me.Tag = "False"
        AddColumnsToGrid grdCustomersLedger, False, 44, GetSetting(strApplicationName, "Layout Strings", "grdCustomersLedger"), _
            "12NCDDate,50NLNInvoiceDetails,40NLNDestination,10NRIAdults,10NRIKids,10NRIFree,10NRFXAdultsAmount,10NRFXKidsAmount,10NRFXDirectAmount,10NRFXDebit,10NRFCredit,10NRFBalance,04NCNMasterRefersTo,04NCNSecondaryRefersTo,04NCNSelected,05NCITrnID", _
            "Εκδοση,Παραστατικό,Προορισμός,Ενήλικες,Παιδιά,Δωρεάν,Χρέωση ενηλίκων,Χρέωση παιδιών,Απευθείας ποσό,Σύνολο χρέωσης,Πίστωση,Υπόλοιπο,A,B,E,TrnID"
        AddColumnsToGrid grdSuppliersLedger, False, 44, GetSetting(strApplicationName, "Layout Strings", "grdSuppliersLedger"), _
            "12NCDDate,50NLNInvoiceDetails,40NLNExpenseDescription,10NRFDebit,10NRFCredit,10NRFBalance,04NCNMasterRefersTo,04NCNSecondaryRefersTo,04NCNSelected,05NCITrnID", _
            "Εκδοση,Παραστατικό,Κατηγορία εξόδου,Χρέωση,Πίστωση,Υπόλοιπο,A,B,E,TrnID"
        AddColumnsToGrid grdPersonsIndex, False, 44, GetSetting(strApplicationName, "Layout Strings", "grdPersonsIndex"), _
            "04NCIID,40NLNDescription,50NLNProfession,50NLNAddress,50NLNPhones,50NLNPersonInCharge,15NLNEmail,15NCNTaxNo,05NCNTaxOfficeID,05NCNXVATStateID,15NCNXAccountCode,05NCNSelected", _
            "ID,Επωνυμία,Δραστηριότητα,Διεύθυνση,Τηλέφωνα,Υπεύθυνος,E-mail,Α.Φ.Μ.,Δ.Ο.Υ.,Καθεστώς Φ.Π.Α.,Κωδ. Γεν. Λογιστικής,Ε"
        Me.Refresh
        DisplayCustomersOrSuppliersGrid
        HideOrDisplayDestinationCriteria
        If txtBatchReport.text = "No" Then
            txtPersonDescription.SetFocus
            grdPersonsIndex.Visible = False
        Else
            'frmCriteria(1).Visible = True
        End If
    End If
            
    'AddDummyLines grdCustomersLedger, "99/99/9999", "ΑΑΑΑΑΑΑΑΑΑΑΑ", "ΑΑΑΑΑΑΑΑΑΑΑΑΑΑΑΑΑΑΑΑΑΑΑΑΑΑΑΑΑΑΑΑΑΑΑΑΑΑΑΑ", "999999", "999999", "999999", "-9999999", "-9999999", "-9999999", "-9999999", "-9999999", "-9999999", "", "", "", ""
    'AddDummyLines grdSuppliersLedger, "A99/99/9999A", "ΑΑΑΑΑΑΑΑΑΑΑΑ", "ΑΑΑΑΑΑΑΑΑΑΑΑΑΑΑΑΑΑΑΑΑΑΑΑΑΑΑΑΑΑΑΑΑΑΑΑΑΑΑΑ", "-9999999", "-9999999", "-9999999"
    
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
        Case vbKeyE And CtrlDown And cmdButton(1).Enabled
            cmdButton_Click 1
        Case vbKeyP And CtrlDown And Not AltDown And cmdButton(2).Enabled
            cmdButton_Click 2
        Case vbKeyP And CtrlDown And AltDown And cmdButton(3).Enabled
            cmdButton_Click 3
        Case vbKeyEscape
            If cmdButton(5).Enabled Then cmdButton_Click 5: Exit Function
            If cmdButton(6).Enabled Then cmdButton_Click 6
        Case vbKeyF12 And CtrlDown
            ToggleInfoPanel Me
    End Select

End Function

Private Sub Form_Load()

    SetUpGrid lstIconList, grdCustomersLedger, grdSuppliersLedger, grdPersonsIndex
    PositionControls Me, True, grdCustomersLedger
    PositionControls Me, True, grdSuppliersLedger
    PositionControls Me, True, grdPersonsIndex
    ColorizeControls Me, True
    ClearFields lblRecordCount, lblCriteria, lblSelectedGridLines, lblSelectedGridTotals
    ClearFields txtInvoicePersonID, txtDestinationID
    ClearFields mskInvoiceDateIssueFrom(0), mskInvoiceDateIssueTo(0), mskInvoiceDateIssueFrom(1), mskInvoiceDateIssueTo(1), txtPersonDescription, txtDestinationDescription(0), txtDestinationDescription(1)
    ClearFields grdCustomersLedger, grdSuppliersLedger, grdPersonsIndex
    EnableFields mskInvoiceDateIssueFrom(0), mskInvoiceDateIssueTo(0), txtPersonDescription, txtDestinationDescription(0), mskInvoiceDateIssueFrom(1), mskInvoiceDateIssueTo(1), txtDestinationDescription(1)
    If txtBatchReport = "Yes" Then
        UpdateButtons Me, 6, 1, 0, 0, 0, 0, 0, 1
    Else
        UpdateButtons Me, 6, 1, 0, 0, 0, 0, 0, 1
    End If
    
End Sub

Private Sub grdCustomersLedger_ColHeaderClick(ByVal lCol As Long, bDoDefault As Boolean, ByVal Shift As Integer, ByVal X As Long, ByVal Y As Long)

    bDoDefault = False

End Sub

Private Sub grdCustomersLedger_CurCellChange(ByVal lRow As Long, ByVal lCol As Long)
    
    cmdButton(1).Enabled = ChangeEditButtonStatus(grdCustomersLedger, Me.Tag, lRow, 1)

End Sub

Private Sub grdCustomersLedger_DblClick(ByVal lRow As Long, ByVal lCol As Long, bRequestEdit As Boolean)

    If cmdButton(1).Enabled Then cmdButton_Click 1

End Sub

Private Sub grdCustomersLedger_HeaderRightClick(ByVal lCol As Long, ByVal Shift As Integer, ByVal X As Long, ByVal Y As Long)

    PopupMenu mnuHdrPopUp

End Sub

Private Sub grdCustomersLedger_KeyDown(KeyCode As Integer, Shift As Integer, bDoDefault As Boolean)

    If KeyCode = vbKeySpace And grdCustomersLedger.RowCount > 0 Then
        grdCustomersLedger.CellIcon(grdCustomersLedger.CurRow, "Selected") = lstIconList.ItemIndex(SelectRow(grdCustomersLedger, 2, KeyCode, grdCustomersLedger.CurRow, "TrnID"))
        lblSelectedGridLines.Caption = CountSelected(grdCustomersLedger)
        lblSelectedGridTotals.Caption = SumSelectedGridRows(grdCustomersLedger, True, "AdultsAmount", "Χρέωση ενηλίκων", "decimal", "KidsAmount", "Χρέωση παιδιών", "decimal", "DirectAmount", "Απευθείας χρέωση", "decimal", "Debit", "Σύνολο χρέωσης", "decimal", "Credit", "Πίστωση", "decimal", "Balance", "Υπόλοιπο", "decimal")
    End If

End Sub

Private Sub grdCustomersLedger_KeyPress(KeyAscii As Integer)

    If KeyAscii = vbKeyReturn And cmdButton(1).Enabled Then cmdButton_Click 1

End Sub

Private Sub grdPersonsIndex_ColHeaderMouseEnter(ByVal lCol As Long)

    grdPersonsIndex.Header.Buttons = True

End Sub


Private Sub grdPersonsIndex_ColHeaderMouseLeave(ByVal lCol As Long)

    grdPersonsIndex.Header.Buttons = False

End Sub


Private Sub grdPersonsIndex_HeaderRightClick(ByVal lCol As Long, ByVal Shift As Integer, ByVal X As Long, ByVal Y As Long)

    PopupMenu mnuHdrPopUp

End Sub

Private Sub grdPersonsIndex_KeyDown(KeyCode As Integer, Shift As Integer, bDoDefault As Boolean)

    If KeyCode = vbKeySpace And grdPersonsIndex.RowCount > 0 Then
        grdPersonsIndex.CellIcon(grdPersonsIndex.CurRow, "Selected") = lstIconList.ItemIndex(SelectRow(grdPersonsIndex, 2, KeyCode, grdPersonsIndex.CurRow, "ID"))
        lblSelectedGridLines.Caption = CountSelected(grdPersonsIndex)
    End If

End Sub


Private Sub grdSuppliersLedger_ColHeaderClick(ByVal lCol As Long, bDoDefault As Boolean, ByVal Shift As Integer, ByVal X As Long, ByVal Y As Long)

    bDoDefault = False

End Sub

Private Sub grdSuppliersLedger_CurCellChange(ByVal lRow As Long, ByVal lCol As Long)

    cmdButton(1).Enabled = ChangeEditButtonStatus(grdSuppliersLedger, Me.Tag, lRow, 1)

End Sub

Private Sub grdSuppliersLedger_DblClick(ByVal lRow As Long, ByVal lCol As Long, bRequestEdit As Boolean)

    If cmdButton(1).Enabled Then cmdButton_Click 1

End Sub

Private Sub grdSuppliersLedger_HeaderRightClick(ByVal lCol As Long, ByVal Shift As Integer, ByVal X As Long, ByVal Y As Long)

    PopupMenu mnuHdrPopUp

End Sub

Private Sub grdSuppliersLedger_KeyDown(KeyCode As Integer, Shift As Integer, bDoDefault As Boolean)

    If KeyCode = vbKeySpace And grdSuppliersLedger.RowCount > 0 Then
        grdSuppliersLedger.CellIcon(grdSuppliersLedger.CurRow, "Selected") = lstIconList.ItemIndex(SelectRow(grdSuppliersLedger, 2, KeyCode, grdSuppliersLedger.CurRow, "TrnID"))
        lblSelectedGridLines.Caption = CountSelected(grdSuppliersLedger)
        lblSelectedGridTotals.Caption = SumSelectedGridRows(grdSuppliersLedger, True, "Debit", "Χρέωση", "decimal", "Credit", "Πίστωση", "decimal", "Balance", "Υπόλοιπο", "decimal")
    End If

End Sub

Private Sub grdSuppliersLedger_KeyPress(KeyAscii As Integer)

    If KeyAscii = vbKeyReturn And cmdButton(1).Enabled Then cmdButton_Click 1

End Sub

Private Sub mnuΑποθήκευσηΠλάτουςΣτηλών_Click()

    SaveSetting strApplicationName, "Layout Strings", "grdCustomersLedger", grdCustomersLedger.LayoutCol
    SaveSetting strApplicationName, "Layout Strings", "grdSuppliersLedger", grdSuppliersLedger.LayoutCol
    SaveSetting strApplicationName, "Layout Strings", "grdPersonsIndex", grdPersonsIndex.LayoutCol

End Sub

Private Sub txtDestinationDescription_Change(index As Integer)

    If txtDestinationDescription(index).text = "" Then txtDestinationID.text = ""

End Sub

Private Sub txtDestinationDescription_KeyDown(KeyCode As Integer, Shift As Integer, index As Integer)

    If KeyCode = vbKeyF2 And index = 0 Then cmdIndex_Click 1
    If KeyCode = vbKeyF2 And index = 1 Then cmdIndex_Click 2

End Sub


Private Sub txtDestinationDescription_Validate(index As Integer, Cancel As Boolean)

    If txtDestinationID.text = "" And index = 0 And txtDestinationDescription(index).text <> "" Then cmdIndex_Click 1: If txtDestinationID.text = "" Then Cancel = True
    If txtDestinationID.text = "" And index = 1 And txtDestinationDescription(index).text <> "" Then cmdIndex_Click 2: If txtDestinationID.text = "" Then Cancel = True

End Sub

Private Sub txtPersonDescription_Change()

    If txtPersonDescription.text = "" Then txtInvoicePersonID.text = ""

End Sub

Private Sub txtPersonDescription_KeyDown(KeyCode As Integer, Shift As Integer)

    If KeyCode = vbKeyF2 Then cmdIndex_Click 0

End Sub

Private Sub txtPersonDescription_Validate(Cancel As Boolean)

    If txtInvoicePersonID.text = "" And txtPersonDescription.text <> "" Then cmdIndex_Click 0: If txtInvoicePersonID.text = "" Then Cancel = True

End Sub


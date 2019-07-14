VERSION 5.00
Object = "{158C2A77-1CCD-44C8-AF42-AA199C5DCC6C}#1.0#0"; "dcButton.ocx"
Object = "{55473EAC-7715-4257-B5EF-6E14EBD6A5DD}#1.0#0"; "ProgressBar.ocx"
Begin VB.Form UtilsCheckFiles 
   Appearance      =   0  'Flat
   BackColor       =   &H00E0E0E0&
   BorderStyle     =   1  'Fixed Single
   ClientHeight    =   8925
   ClientLeft      =   15
   ClientTop       =   15
   ClientWidth     =   12300
   ControlBox      =   0   'False
   ForeColor       =   &H00800000&
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   8925
   ScaleWidth      =   12300
   Begin VB.Frame frmProgress 
      BackColor       =   &H000080FF&
      BorderStyle     =   0  'None
      Caption         =   "Frame1"
      Height          =   1140
      Left            =   3450
      TabIndex        =   4
      Top             =   4350
      Width           =   4065
      Begin vbalProgBarLib6.vbalProgressBar prgProgressBar 
         Height          =   615
         Left            =   150
         TabIndex        =   5
         TabStop         =   0   'False
         Top             =   375
         Width           =   3765
         _ExtentX        =   6641
         _ExtentY        =   1085
         Picture         =   "UtilsCheckFiles.frx":0000
         ForeColor       =   0
         Appearance      =   0
         BarPicture      =   "UtilsCheckFiles.frx":001C
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
         TabIndex        =   6
         Top             =   75
         Width           =   3765
      End
   End
   Begin VB.Frame frmButtonFrame 
      BackColor       =   &H00FF8080&
      BorderStyle     =   0  'None
      Height          =   690
      Left            =   75
      TabIndex        =   1
      Top             =   7650
      Width           =   3240
      Begin Dacara_dcButton.dcButton cmdButton 
         Height          =   690
         Index           =   0
         Left            =   225
         TabIndex        =   2
         TabStop         =   0   'False
         Top             =   0
         Width           =   1365
         _ExtentX        =   2408
         _ExtentY        =   1217
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
      Begin Dacara_dcButton.dcButton cmdButton 
         Height          =   690
         Index           =   1
         Left            =   1650
         TabIndex        =   3
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
         PicOpacity      =   0
      End
   End
   Begin VB.Shape shpRightEdge 
      BackColor       =   &H00800080&
      BackStyle       =   1  'Opaque
      BorderStyle     =   0  'Transparent
      Height          =   840
      Left            =   10950
      Top             =   3000
      Visible         =   0   'False
      Width           =   465
   End
   Begin VB.Shape shpBottomEdge 
      BackColor       =   &H00800080&
      BackStyle       =   1  'Opaque
      BorderStyle     =   0  'Transparent
      Height          =   465
      Left            =   1575
      Top             =   8325
      Visible         =   0   'False
      Width           =   840
   End
   Begin VB.Label lblTitle 
      AutoSize        =   -1  'True
      BackColor       =   &H00800000&
      BackStyle       =   0  'Transparent
      Caption         =   "Ελεγχος αρχείων"
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
      TabIndex        =   0
      Top             =   75
      Width           =   3870
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
Attribute VB_Name = "UtilsCheckFiles"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private Function CheckTables()

    On Error GoTo ErrTrap
        
    Dim strSQL As String
    Dim intLoop As Integer
    Dim intFieldCount As Integer
    Dim intNoOfTables As Integer
    Dim strTables() As String
    Dim rstTemp As Recordset
    Dim rstInvoice As Recordset
    Dim blnErrorFound As Boolean
    Dim InvoiceQuery As QueryDef
    
    ReDim strTables(0)
    Set TempQuery = CommonDB.CreateQueryDef("")
    Set InvoiceQuery = CommonDB.CreateQueryDef("")
    
    'Βρίσκω τον αριθμό των πινάκων
    With CommonDB
        For intLoop = 0 To .TableDefs.Count - 1
            If Left(.TableDefs(intLoop).Name, 4) <> "MSys" Then
                ReDim Preserve strTables(UBound(strTables) + 1)
                strTables(UBound(strTables)) = .TableDefs(intLoop).Name
                intNoOfTables = intNoOfTables + 1
            End If
        Next intLoop
    End With
    
    InitializeProgressBar Me, "Ελεγχος αρχείων", intNoOfTables
    
    Open strReportsPathName & "Errors.txt" For Append As #1
    
    'Ελέγχω τον κάθε πίνακα
    For intLoop = 1 To UBound(strTables)
        'Επιλογή όλων των εγγραφών
        TempQuery.SQL = "SELECT * FROM " & strTables(intLoop)
        'Ανοίγω το recordset
        Set rstTemp = TempQuery.OpenRecordset()
        Do While Not rstTemp.EOF
            For intFieldCount = 0 To rstTemp.Fields.Count - 1
                'Ελεγχος εγγραφής για null πεδία
                If IsNull(rstTemp.Fields(intFieldCount)) Then
                    Print #1, "Table: " & strTables(intLoop) & " Field: " & rstTemp.Fields(intFieldCount).Name & " | Rec ID: " & rstTemp.Fields(0).Value & " Field is NULL"
                    blnErrorFound = True
                End If
            Next intFieldCount
            'Αν είμαι στον πίνακα InvoicesOut, γίνεται έλεγχος αν έχω δώσει ποσά ΚΑΙ στα άτομα ΚΑΙ στην απευθείας χρέωση
            If strTables(intLoop) = "InvoicesOut" Then
                If (rstTemp.Fields("InvoiceOutAdultsAmountWithTransfer") <> 0 Or rstTemp.Fields("InvoiceOutKidsAmountWithTransfer") <> 0 Or rstTemp.Fields("InvoiceOutAdultsAmountWithoutTransfer") <> 0 Or rstTemp.Fields("InvoiceOutKidsAmountWithoutTransfer") <> 0) And rstTemp.Fields("InvoiceOutDirectAmount") <> 0 Then
                    Print #1, "Table: " & strTables(intLoop) & " Rec ID: " & rstTemp.Fields(0).Value & " >> Incorrect amounts"
                    blnErrorFound = True
                End If
            End If
            'Αν είμαι στον πίνακα InvoicesOut, γίνεται έλεγχος αν έχω δώσει ποσά χωρίς άτομα (div/0)
            If strTables(intLoop) = "InvoicesOut" Then
                If _
                    (rstTemp.Fields("InvoiceOutAdultsAmountWithTransfer") <> 0 And rstTemp.Fields("InvoiceOutAdultsWithTransfer") = 0) Or _
                    (rstTemp.Fields("InvoiceOutKidsAmountWithTransfer") <> 0 And rstTemp.Fields("InvoiceOutKidsWithTransfer") = 0) Or _
                    (rstTemp.Fields("InvoiceOutAdultsAmountWithoutTransfer") <> 0 And rstTemp.Fields("InvoiceOutAdultsWithoutTransfer") = 0) Or _
                    (rstTemp.Fields("InvoiceOutKidsAmountWithoutTransfer") <> 0 And rstTemp.Fields("InvoiceOutKidsWithoutTransfer") = 0) _
                Then
                    Print #1, "Table: " & strTables(intLoop) & " Rec ID: " & rstTemp.Fields(0).Value & " >> Division by zero"
                    blnErrorFound = True
                End If
            End If
            rstTemp.MoveNext
        Loop
        'Αν είμαι σε πίνακα Invoices(...) ή Payment(...), ελέγχω για εγγραφές με ίδιο TrnID
        If Left(strTables(intLoop), 8) = "Invoices" Or strTables(intLoop) = "PaymentIn" Or strTables(intLoop) = "PaymentOut" Then
            Select Case strTables(intLoop)
                Case Is = "Invoices"
                    strSQL = "SELECT InvoiceTrnID From Invoices GROUP BY InvoiceTrnID HAVING COUNT(*) > 1"
                Case Is = "InvoicesIn"
                    strSQL = "SELECT InvoiceInTrnID From InvoicesIn GROUP BY InvoiceInTrnID HAVING COUNT(*) > 1"
                Case Is = "InvoicesOut"
                    strSQL = "SELECT InvoiceOutTrnID From InvoicesOut GROUP BY InvoiceOutTrnID HAVING COUNT(*) > 1"
                Case Is = "PaymentIn"
                    strSQL = "SELECT TrnID From PaymentIn GROUP BY TrnId HAVING COUNT(*) > 1"
                Case Is = "PaymentOut"
                    strSQL = "SELECT TrnID From PaymentOut GROUP BY TrnId HAVING COUNT(*) > 1"
            End Select
            TempQuery.SQL = strSQL
            Set rstInvoice = TempQuery.OpenRecordset()
            If Not rstInvoice.EOF Then
                Do While Not rstInvoice.EOF
                    Print #1, "Table: " & strTables(intLoop) & " TrnID: " & rstInvoice(0).Value & " >> Duplicate InvoiceTrnID"
                    rstInvoice.MoveNext
                Loop
            End If
        End If
        UpdateProgressBar Me
    Next intLoop
    
    'Τέλος
    Close #1
        
    If blnErrorFound Then CheckTables = False Else CheckTables = True
    
    'Εξοδος
    Exit Function
    
ErrTrap:
    DisplayErrorMessage True, Err.Description

End Function

Private Function StartProcess()

    frmProgress.Visible = True
    
    If CheckTables Then
        frmProgress.Visible = False
        If MyMsgBox(1, strApplicationName, strAppMessages(13), 1) Then
        End If
    Else
        frmProgress.Visible = False
        If MyMsgBox(4, strApplicationName, strAppMessages(2), 1) Then
        End If
    End If
    
End Function

Private Sub cmdButton_Click(index As Integer)

    Select Case index
        Case 0
            StartProcess
        Case 1
            AbortProcedure True
    End Select
    
End Sub

Private Function AbortProcedure(blnStatus)

    If blnStatus Then
        Unload Me
    End If

End Function

Private Sub Form_Activate()
                
    If Me.Tag = "True" Then
        Me.Tag = "False"
        UpdateButtons Me, 1, 1, 1
        frmProgress.Visible = False
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
        Case vbKeyF10 And cmdButton(0).Enabled, vbKeyC And CtrlDown And cmdButton(0).Enabled
            cmdButton_Click 0
        Case vbKeyEscape
            cmdButton_Click 1
    End Select

End Function

Private Sub Form_Load()

    UpdateColors Me, False
    UpdateButtons Me, 1, 1, 1

End Sub



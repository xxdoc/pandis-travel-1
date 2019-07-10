VERSION 5.00
Begin {9EB8768B-CDFA-44DF-8F3E-857A8405E1DB} rptTransfers 
   ClientHeight    =   10950
   ClientLeft      =   120
   ClientTop       =   450
   ClientWidth     =   17430
   StartUpPosition =   2  'CenterScreen
   _ExtentX        =   30745
   _ExtentY        =   19315
   SectionData     =   "rptTransfers.dsx":0000
End
Attribute VB_Name = "rptTransfers"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private Sub ActiveReport_DataInitialize()

    On Error GoTo ErrTrap
    
    Open strUnicodeFile For Input As #1
    
    Fields.RemoveAll
    
    Fields.Add "OneLongField"
    
    Exit Sub
    
ErrTrap:
    Exit Sub
    
End Sub

Private Sub ActiveReport_FetchData(EOF As Boolean)

    Dim strLine As String
    Dim arr As String
    
    If VBA.EOF(1) Then
        EOF = True
        Exit Sub
    Else
        EOF = False
    End If
    
    Line Input #1, strLine
    
    If strLine <> "" Then
        If Right(strLine, 1) = "^" Then
            Fields("OneLongField").Value = RTrim(Left(strLine, Len(strLine) - 1))
        Else
            Fields("OneLongField").Value = strLine
        End If
    Else
        Fields("OneLongField").Value = strLine
    End If
    
    DateSeperator.Visible = IIf(Mid(strLine, 3, 1) = ":", True, False)
    DestinationSeperator.Visible = IIf(Mid(strLine, 3, 1) = ":", True, False)
    
    Seperator.Visible = IIf(Right(strLine, 1) = "^", True, False)
    
End Sub

Private Sub ActiveReport_ReportEnd()

    Close #1

End Sub


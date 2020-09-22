VERSION 5.00
Object = "{5E9E78A0-531B-11CF-91F6-C2863C385E30}#1.0#0"; "MSFLXGRD.OCX"
Begin VB.Form frmHeadings 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Change Field Headers"
   ClientHeight    =   6795
   ClientLeft      =   45
   ClientTop       =   435
   ClientWidth     =   10080
   HelpContextID   =   1018
   Icon            =   "frmHeadings.frx":0000
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   MaxButton       =   0   'False
   MinButton       =   0   'False
   Picture         =   "frmHeadings.frx":038A
   ScaleHeight     =   6795
   ScaleWidth      =   10080
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  'CenterScreen
   Begin Reports_Factory.ucButtons_H cmdOK 
      Height          =   375
      Left            =   3120
      TabIndex        =   1
      Top             =   6240
      Width           =   1815
      _ExtentX        =   3201
      _ExtentY        =   661
      Caption         =   "&OK  "
      CapAlign        =   2
      BackStyle       =   2
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Arial"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      cFore           =   0
      cFHover         =   0
      cGradient       =   14737632
      Mode            =   0
      Value           =   0   'False
      ImgAlign        =   3
      Image           =   "frmHeadings.frx":3389B
      cBack           =   -2147483633
      mPointer        =   99
      mIcon           =   "frmHeadings.frx":33C35
   End
   Begin MSFlexGridLib.MSFlexGrid MSFlexGrid1 
      DragIcon        =   "frmHeadings.frx":33F4F
      Height          =   5775
      Left            =   240
      TabIndex        =   0
      Top             =   240
      Width           =   9615
      _ExtentX        =   16960
      _ExtentY        =   10186
      _Version        =   393216
      Cols            =   3
      FixedCols       =   0
      BackColor       =   16777215
      BackColorFixed  =   12615680
      ForeColorFixed  =   -2147483628
      AllowUserResizing=   3
      FormatString    =   $"frmHeadings.frx":34259
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
   End
   Begin Reports_Factory.ucButtons_H cmdHelp 
      Height          =   375
      Left            =   5160
      TabIndex        =   2
      Top             =   6240
      Width           =   1815
      _ExtentX        =   3201
      _ExtentY        =   661
      Caption         =   "&Help"
      CapAlign        =   2
      BackStyle       =   2
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Arial"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      cFore           =   0
      cFHover         =   0
      cGradient       =   14737632
      Mode            =   0
      Value           =   0   'False
      ImgAlign        =   3
      Image           =   "frmHeadings.frx":342E4
      cBack           =   -2147483633
      mPointer        =   99
      mIcon           =   "frmHeadings.frx":37366
   End
End
Attribute VB_Name = "frmHeadings"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private Sub cmdHelp_Click()
Call ShowAppHelp(1018)
End Sub

Private Sub cmdOK_Click()
Unload Me
End Sub

Private Sub Form_Activate()
boolUnload = False
With MSFlexGrid1
.Redraw = False
Dim lngX As Long

.Row = 0
.Col = 0
.CellAlignment = 4
.Col = 1
.CellAlignment = 4
.Col = 2
.CellAlignment = 4

lngX = 1
While lngX + 1 < .Rows
    .Row = lngX
    
    .Col = 0
    .CellAlignment = 1
    .CellFontBold = False
    
    .Col = 1
    .CellAlignment = 1
    .CellFontBold = False
    
    .Col = 2
    .CellAlignment = 1
    .CellFontBold = False
    lngX = lngX + 1
Wend
.Redraw = True

End With
End Sub

Private Sub Form_Unload(Cancel As Integer)
If boolUnload = False Then
    Call setHeadings
    Me.Hide
    Cancel = 1
    frmReport.Enabled = True
    frmReport.SetFocus
End If
End Sub

Private Sub MSFlexGrid1_DblClick()
Call SortFlexiArrows(MSFlexGrid1, True, True)
End Sub

Private Sub MSFlexGrid1_EnterCell()
With MSFlexGrid1
If .CellFontBold = True Then
    Exit Sub
End If
If .Col = 2 And .Row <> 0 And .Row <> .Rows - 1 Then
    .CellBackColor = &HC0FFFF
End If
End With
End Sub

Private Sub MSFlexGrid1_KeyPress(KeyAscii As Integer)
With MSFlexGrid1
If .CellFontBold = True Then
    Exit Sub
End If
If .CellBackColor = &HC0FFFF Then
    Select Case KeyAscii
        Case vbKeyReturn
            .Row = .Row + 1
            
        Case vbKeyBack
            If Trim(.Text) <> "" Then
                .Text = Mid(.Text, 1, Len(.Text) - 1)
            End If
        Case Is < 32
        
        Case Else
            If .Col = 0 Or .Row = 0 Then
                Exit Sub
            Else
                .Text = .Text & Chr(KeyAscii)
            End If
    End Select
End If
End With
End Sub

Private Sub MSFlexGrid1_LeaveCell()
If (MSFlexGrid1.Col = 1 Or MSFlexGrid1.Col = 2) And MSFlexGrid1.Row <> 0 Then
    MSFlexGrid1.CellBackColor = vbWhite
End If
If Left(UCase(MSFlexGrid1.Text), 6) = "[NEW]." And MSFlexGrid1.Col = 2 Then
    MSFlexGrid1.Text = ""
    MsgBox "You cannot use '[NEW].' as a column alias since it is a reserved word", vbExclamation
ElseIf MSFlexGrid1.Col = 2 Then
    Dim lngX As Long
    lngX = 1
    While lngX + 1 < MSFlexGrid1.Rows
        If UCase(MSFlexGrid1.TextMatrix(lngX, 1)) = UCase(MSFlexGrid1.Text) And Len(Trim(MSFlexGrid1.Text)) > 0 Then
            MSFlexGrid1.Text = ""
            MsgBox "The column heading specified by you already exist. Please enter a new heading", vbExclamation
        ElseIf UCase(MSFlexGrid1.TextMatrix(lngX, 2)) = UCase(MSFlexGrid1.Text) And Len(Trim(MSFlexGrid1.Text)) > 0 And lngX <> MSFlexGrid1.Row Then
            MSFlexGrid1.Text = ""
            MsgBox "The column heading specified by you already exist. Please enter a new heading", vbExclamation

        End If
        lngX = lngX + 1
    Wend
End If
End Sub

Sub setHeadings()
With frmReport
MSFlexGrid1.Redraw = False
.MSFlexGrid1.Redraw = False
Dim lngX As Long
Dim lngY As Long
lngX = 1
MSFlexGrid1.Col = 1
.MSFlexGrid1.Row = 1
While lngX < MSFlexGrid1.Rows
    MSFlexGrid1.Row = lngX
    lngY = 1
    While lngY < .MSFlexGrid1.Cols
        .MSFlexGrid1.Col = lngY
        If UCase(.MSFlexGrid1.Text) = UCase(MSFlexGrid1.Text) Then
            MSFlexGrid1.Col = 2
            If Len(Trim(MSFlexGrid1.Text)) > 0 Then
                .MSFlexGrid1.Text = MSFlexGrid1.Text
                MSFlexGrid1.Text = ""
                MSFlexGrid1.Col = 1
                MSFlexGrid1.Text = .MSFlexGrid1.Text
            End If
            MSFlexGrid1.Col = 1
        End If
        lngY = lngY + 1
    Wend
    lngX = lngX + 1
Wend
MSFlexGrid1.Redraw = True
.MSFlexGrid1.Redraw = True

Call FG_AutosizeCols(frmReport.MSFlexGrid1, frmReport, , , True)
End With
End Sub

Private Sub MSFlexGrid1_LostFocus()
Call MSFlexGrid1_LeaveCell
End Sub

VERSION 5.00
Begin VB.Form frmFind 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Find"
   ClientHeight    =   2190
   ClientLeft      =   45
   ClientTop       =   435
   ClientWidth     =   6240
   Icon            =   "frmFind.frx":0000
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   MaxButton       =   0   'False
   MinButton       =   0   'False
   Picture         =   "frmFind.frx":038A
   ScaleHeight     =   2190
   ScaleWidth      =   6240
   StartUpPosition =   2  'CenterScreen
   Begin VB.CheckBox chkMatchCase 
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   225
      Left            =   1320
      TabIndex        =   7
      Top             =   1200
      Width           =   200
   End
   Begin VB.ComboBox cboCol 
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   315
      ItemData        =   "frmFind.frx":9511
      Left            =   1320
      List            =   "frmFind.frx":951E
      Style           =   2  'Dropdown List
      TabIndex        =   3
      Top             =   720
      Width           =   1695
   End
   Begin VB.TextBox txtFind 
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   315
      Left            =   1320
      MaxLength       =   150
      TabIndex        =   1
      Top             =   240
      Width           =   4695
   End
   Begin Reports_Factory.ucButtons_H cmdExit 
      Height          =   375
      Left            =   4680
      TabIndex        =   6
      Top             =   1680
      Width           =   1335
      _ExtentX        =   2355
      _ExtentY        =   661
      Caption         =   "E&xit  "
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
      Image           =   "frmFind.frx":9544
      cBack           =   -2147483633
      mPointer        =   99
      mIcon           =   "frmFind.frx":98DE
   End
   Begin Reports_Factory.ucButtons_H cmdFindNext 
      Height          =   375
      Left            =   3000
      TabIndex        =   5
      Top             =   1680
      Width           =   1455
      _ExtentX        =   2566
      _ExtentY        =   661
      Caption         =   "Find &Next"
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
      Image           =   "frmFind.frx":9BF8
      Enabled         =   0   'False
      cBack           =   -2147483633
      mPointer        =   99
      mIcon           =   "frmFind.frx":9F92
   End
   Begin Reports_Factory.ucButtons_H cmdFind 
      Height          =   375
      Left            =   1320
      TabIndex        =   4
      Top             =   1680
      Width           =   1455
      _ExtentX        =   2566
      _ExtentY        =   661
      Caption         =   "&Find"
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
      Image           =   "frmFind.frx":A2AC
      cBack           =   -2147483633
      mPointer        =   99
      mIcon           =   "frmFind.frx":A646
   End
   Begin VB.Label Label2 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Match Case"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00000000&
      Height          =   195
      Left            =   1560
      TabIndex        =   8
      Top             =   1230
      UseMnemonic     =   0   'False
      Width           =   1020
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Find in Col:"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   195
      Left            =   240
      TabIndex        =   2
      Top             =   840
      Width           =   975
   End
   Begin VB.Label lblServer 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Find:"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   195
      Left            =   240
      TabIndex        =   0
      Top             =   360
      Width           =   435
   End
End
Attribute VB_Name = "frmFind"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Dim lngRow As Long, lngCol As Long

Private Sub cmdExit_Click()
Unload Me
End Sub

Private Sub cmdFind_Click()
If frmReport.MSFlexGrid1.Visible = True Then
    cmdFindNext.Enabled = True
    Call findText(frmReport.MSFlexGrid1, False)
ElseIf frmReport.MSFlexGrid2.Visible = True Then
    cmdFindNext.Enabled = True
    Call findText(frmReport.MSFlexGrid2, False)
End If
End Sub

Private Sub cmdFindNext_Click()
If frmReport.MSFlexGrid1.Visible = True Then
    Call findText(frmReport.MSFlexGrid1, True)
ElseIf frmReport.MSFlexGrid2.Visible = True Then
    Call findText(frmReport.MSFlexGrid2, True)
End If
End Sub

Private Sub Form_KeyPress(KeyAscii As Integer)
If KeyAscii = 27 Then
    Unload Me
ElseIf KeyAscii = 13 Then
    If cmdFindNext.Enabled = True Then
        Call cmdFindNext_Click
    ElseIf cmdFind.Enabled = True Then
        Call cmdFind_Click
    End If
End If
End Sub

Private Sub Form_Load()
Dim lngX As Long
lngX = 1
cboCol.Clear
cboCol.AddItem "All"
If frmReport.MSFlexGrid1.Visible = True Then
    While lngX < frmReport.MSFlexGrid1.Cols
        cboCol.AddItem frmReport.MSFlexGrid1.TextMatrix(0, lngX)
        lngX = lngX + 1
    Wend
ElseIf frmReport.MSFlexGrid2.Visible = True Then
    While lngX < frmReport.MSFlexGrid2.Cols
        cboCol.AddItem frmReport.MSFlexGrid2.TextMatrix(0, lngX)
        lngX = lngX + 1
    Wend
End If
cboCol.Text = cboCol.List(0)
End Sub

Private Sub Form_Unload(Cancel As Integer)
On Error Resume Next
frmReport.SetFocus
End Sub

Sub findText(msfFlexi As MSFlexGrid, boolFindNext As Boolean)
On Error GoTo err:
With msfFlexi
.Redraw = False
.SetFocus
If boolFindNext = False Then
    lngRow = 2
ElseIf boolFindNext = True And lngCol = .Cols Then
    lngCol = 1
    lngRow = lngRow + 1
End If
If cboCol.ListIndex = 0 Then
    While lngRow + 1 < .Rows
        If boolFindNext = False Or lngCol = .Cols Then
            lngCol = 1
        End If
        .Row = lngRow
        While lngCol < .Cols
            .Col = lngCol
            If chkMatchCase.Value = 1 Then
                If InStr(.Text, txtFind.Text) <> 0 Then
                    .Redraw = True
                    lngCol = lngCol + 1
                    Exit Sub
                End If
            ElseIf chkMatchCase.Value = 0 Then
                If InStr(UCase(.Text), UCase(txtFind.Text)) <> 0 Then
                    .Redraw = True
                    lngCol = lngCol + 1
                    Exit Sub
                End If
            End If
            lngCol = lngCol + 1
            SendKeys "{RIGHT}", True
        Wend
        lngRow = lngRow + 1
        SendKeys "{DOWN}", True
    Wend
    .Redraw = True
    cmdFindNext.Enabled = False
    MsgBox "Searched item not found", vbExclamation
Else
    .Col = cboCol.ListIndex
    While lngRow + 1 < .Rows
        .Row = lngRow
        If chkMatchCase.Value = 1 Then
            If InStr(.Text, txtFind.Text) <> 0 Then
                lngRow = lngRow + 1
                .Redraw = True
                Exit Sub
            End If
        ElseIf chkMatchCase.Value = 0 Then
            If InStr(UCase(.Text), UCase(txtFind.Text)) <> 0 Then
                lngRow = lngRow + 1
                .Redraw = True
                Exit Sub
            End If
        End If
        lngRow = lngRow + 1
        SendKeys "{DOWN}", True
    Wend
    .Redraw = True
    cmdFindNext.Enabled = False
    MsgBox "Searched item not found", vbExclamation
End If
Exit Sub

err:
MsgBox err.Description, vbExclamation
.Redraw = True

End With
End Sub

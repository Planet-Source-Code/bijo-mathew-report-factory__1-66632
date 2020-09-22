VERSION 5.00
Begin VB.Form frmCalculate 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Calculated Field"
   ClientHeight    =   3660
   ClientLeft      =   45
   ClientTop       =   435
   ClientWidth     =   5700
   HelpContextID   =   1019
   Icon            =   "frmCalculate.frx":0000
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   MaxButton       =   0   'False
   MinButton       =   0   'False
   Picture         =   "frmCalculate.frx":038A
   ScaleHeight     =   3660
   ScaleWidth      =   5700
   StartUpPosition =   2  'CenterScreen
   Begin VB.TextBox txtCol 
      Height          =   315
      Index           =   4
      Left            =   1920
      Locked          =   -1  'True
      TabIndex        =   16
      Top             =   2400
      Visible         =   0   'False
      Width           =   1815
   End
   Begin VB.TextBox txtCol 
      Height          =   315
      Index           =   3
      Left            =   1920
      Locked          =   -1  'True
      TabIndex        =   13
      Top             =   1920
      Visible         =   0   'False
      Width           =   1815
   End
   Begin VB.TextBox txtCol 
      Height          =   315
      Index           =   2
      Left            =   1920
      Locked          =   -1  'True
      TabIndex        =   10
      Top             =   1440
      Visible         =   0   'False
      Width           =   1815
   End
   Begin VB.TextBox txtCol 
      Height          =   315
      Index           =   1
      Left            =   1920
      Locked          =   -1  'True
      TabIndex        =   7
      Top             =   960
      Visible         =   0   'False
      Width           =   1815
   End
   Begin VB.TextBox txtCol 
      Height          =   315
      Index           =   0
      Left            =   1920
      Locked          =   -1  'True
      TabIndex        =   4
      Top             =   480
      Visible         =   0   'False
      Width           =   1815
   End
   Begin VB.ComboBox cboFrom 
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
      Index           =   4
      ItemData        =   "frmCalculate.frx":9511
      Left            =   360
      List            =   "frmCalculate.frx":951B
      Sorted          =   -1  'True
      Style           =   2  'Dropdown List
      TabIndex        =   15
      Top             =   2400
      Width           =   1455
   End
   Begin VB.ComboBox cboFrom 
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
      Index           =   3
      ItemData        =   "frmCalculate.frx":9533
      Left            =   360
      List            =   "frmCalculate.frx":953D
      Sorted          =   -1  'True
      Style           =   2  'Dropdown List
      TabIndex        =   12
      Top             =   1920
      Width           =   1455
   End
   Begin VB.ComboBox cboFrom 
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
      Index           =   2
      ItemData        =   "frmCalculate.frx":9555
      Left            =   360
      List            =   "frmCalculate.frx":955F
      Sorted          =   -1  'True
      Style           =   2  'Dropdown List
      TabIndex        =   9
      Top             =   1440
      Width           =   1455
   End
   Begin VB.ComboBox cboFrom 
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
      Index           =   1
      ItemData        =   "frmCalculate.frx":9577
      Left            =   360
      List            =   "frmCalculate.frx":9581
      Sorted          =   -1  'True
      Style           =   2  'Dropdown List
      TabIndex        =   6
      Top             =   960
      Width           =   1455
   End
   Begin VB.ComboBox cboFrom 
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
      Index           =   0
      ItemData        =   "frmCalculate.frx":9599
      Left            =   360
      List            =   "frmCalculate.frx":95A3
      Sorted          =   -1  'True
      Style           =   2  'Dropdown List
      TabIndex        =   3
      Top             =   480
      Width           =   1455
   End
   Begin VB.TextBox txtVal 
      Height          =   315
      Index           =   4
      Left            =   1920
      Locked          =   -1  'True
      TabIndex        =   24
      Top             =   2400
      Width           =   1815
   End
   Begin VB.TextBox txtVal 
      Height          =   315
      Index           =   3
      Left            =   1920
      Locked          =   -1  'True
      TabIndex        =   23
      Top             =   1920
      Width           =   1815
   End
   Begin VB.ComboBox cboOP 
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
      Index           =   3
      ItemData        =   "frmCalculate.frx":95BB
      Left            =   3840
      List            =   "frmCalculate.frx":95D4
      Sorted          =   -1  'True
      Style           =   2  'Dropdown List
      TabIndex        =   14
      Top             =   1920
      Width           =   1695
   End
   Begin VB.TextBox txtVal 
      Height          =   315
      Index           =   2
      Left            =   1920
      Locked          =   -1  'True
      TabIndex        =   22
      Top             =   1440
      Width           =   1815
   End
   Begin VB.ComboBox cboOP 
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
      Index           =   2
      ItemData        =   "frmCalculate.frx":9608
      Left            =   3840
      List            =   "frmCalculate.frx":9621
      Sorted          =   -1  'True
      Style           =   2  'Dropdown List
      TabIndex        =   11
      Top             =   1440
      Width           =   1695
   End
   Begin VB.TextBox txtVal 
      Height          =   315
      Index           =   1
      Left            =   1920
      Locked          =   -1  'True
      TabIndex        =   21
      Top             =   960
      Width           =   1815
   End
   Begin VB.ComboBox cboOP 
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
      Index           =   1
      ItemData        =   "frmCalculate.frx":9655
      Left            =   3840
      List            =   "frmCalculate.frx":966E
      Sorted          =   -1  'True
      Style           =   2  'Dropdown List
      TabIndex        =   8
      Top             =   960
      Width           =   1695
   End
   Begin VB.TextBox txtVal 
      Height          =   315
      Index           =   0
      Left            =   1920
      Locked          =   -1  'True
      TabIndex        =   20
      Top             =   480
      Width           =   1815
   End
   Begin VB.ComboBox cboOP 
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
      Index           =   0
      ItemData        =   "frmCalculate.frx":96A2
      Left            =   3840
      List            =   "frmCalculate.frx":96BB
      Sorted          =   -1  'True
      Style           =   2  'Dropdown List
      TabIndex        =   5
      Top             =   480
      Width           =   1695
   End
   Begin Reports_Factory.ucButtons_H cmdCancel 
      Height          =   375
      Left            =   3000
      TabIndex        =   18
      Top             =   3120
      Width           =   1455
      _ExtentX        =   2566
      _ExtentY        =   661
      Caption         =   "&Cancel  "
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
      Image           =   "frmCalculate.frx":96EF
      cBack           =   -2147483633
      mPointer        =   99
      mIcon           =   "frmCalculate.frx":9A89
   End
   Begin Reports_Factory.ucButtons_H cmdOK 
      Height          =   375
      Left            =   1200
      TabIndex        =   17
      Top             =   3120
      Width           =   1575
      _ExtentX        =   2778
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
      Image           =   "frmCalculate.frx":9DA3
      cBack           =   -2147483633
      mPointer        =   99
      mIcon           =   "frmCalculate.frx":A13D
   End
   Begin VB.Label lblColInv 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Col"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   240
      Left            =   960
      TabIndex        =   1
      Top             =   120
      UseMnemonic     =   0   'False
      Visible         =   0   'False
      Width           =   360
   End
   Begin VB.Label Label2 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "="
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   240
      Left            =   120
      TabIndex        =   2
      Top             =   520
      Width           =   135
   End
   Begin VB.Label lblCol 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Col"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   240
      Left            =   960
      TabIndex        =   19
      Top             =   120
      Width           =   360
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Col:"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   240
      Left            =   480
      TabIndex        =   0
      Top             =   120
      Width           =   420
   End
End
Attribute VB_Name = "frmCalculate"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Public intIndex As Integer

Private Sub cboFrom_Change(Index As Integer)
If cboFrom(Index).Text = "From Report" Then
    txtVal(Index).Text = ""
    txtVal(Index).Locked = True
Else
    txtVal(Index).Text = ""
    txtVal(Index).Locked = False
End If
End Sub

Private Sub cboFrom_Click(Index As Integer)
Call cboFrom_Change(Index)
End Sub

Private Sub cboFrom_GotFocus(Index As Integer)
intIndex = Index
End Sub

Private Sub cboOP_Change(Index As Integer)
intIndex = -1
End Sub

Private Sub cmdCancel_Click()
Unload Me
frmReport.SetFocus
End Sub

Public Sub cmdOK_Click()
On Error GoTo Err
With frmReport
.MSFlexGrid1.Redraw = False
Dim lngX As Long
Dim lngY As Long
Dim strHeading, strOperations As String
strHeading = ""
strOperations = ""

lngX = 2
.MSFlexGrid1.Col = CLng(lblColInv.Caption)
strHeading = .MSFlexGrid1.TextMatrix(1, CLng(lblColInv.Caption))
While lngX + 1 < .MSFlexGrid1.Rows
    .MSFlexGrid1.Row = lngX
    If Len(txtVal(0).Text) > 0 Then
        If cboFrom(0).Text = "From Report" Then
            .MSFlexGrid1.Text = .MSFlexGrid1.TextMatrix(lngX, CInt(txtCol(0).Text))
            strOperations = "Report." & getBaseColName(.MSFlexGrid1.TextMatrix(1, CInt(txtCol(0).Text)))
        Else
            .MSFlexGrid1.Text = txtVal(0).Text
            strOperations = "Value." & txtVal(0).Text
        End If
    End If
    
    lngY = 1
    While lngY <= 4
        If Len(txtVal(lngY).Text) > 0 Then
            If cboFrom(lngY).Text = "From Report" Then
                Select Case cboOP(lngY - 1).Text
                    Case "+"
                        .MSFlexGrid1.Text = CDbl(.MSFlexGrid1.Text) + CDbl(.MSFlexGrid1.TextMatrix(lngX, CInt(txtCol(lngY).Text)))
                        strOperations = strOperations & vbCrLf & "+" & vbCrLf & "Report." & getBaseColName(.MSFlexGrid1.TextMatrix(1, CInt(txtCol(lngY).Text)))
                    Case "-"
                        .MSFlexGrid1.Text = CDbl(.MSFlexGrid1.Text) - CDbl(.MSFlexGrid1.TextMatrix(lngX, CInt(txtCol(lngY).Text)))
                        strOperations = strOperations & vbCrLf & "-" & vbCrLf & "Report." & getBaseColName(.MSFlexGrid1.TextMatrix(1, CInt(txtCol(lngY).Text)))
                    Case "*"
                        .MSFlexGrid1.Text = CDbl(.MSFlexGrid1.Text) * CDbl(.MSFlexGrid1.TextMatrix(lngX, CInt(txtCol(lngY).Text)))
                        strOperations = strOperations & vbCrLf & "*" & vbCrLf & "Report." & getBaseColName(.MSFlexGrid1.TextMatrix(1, CInt(txtCol(lngY).Text)))
                    Case "/"
                        .MSFlexGrid1.Text = CDbl(.MSFlexGrid1.Text) / CDbl(.MSFlexGrid1.TextMatrix(lngX, CInt(txtCol(lngY).Text)))
                        strOperations = strOperations & vbCrLf & "/" & vbCrLf & "Report." & getBaseColName(.MSFlexGrid1.TextMatrix(1, CInt(txtCol(lngY).Text)))
                    Case "Concatenate"
                        .MSFlexGrid1.Text = .MSFlexGrid1.Text & .MSFlexGrid1.TextMatrix(lngX, CInt(txtCol(lngY).Text))
                        strOperations = strOperations & vbCrLf & "Concatenate" & vbCrLf & "Report." & getBaseColName(.MSFlexGrid1.TextMatrix(1, CInt(txtCol(lngY).Text)))
                    Case "Greater Of"
                        strOperations = strOperations & vbCrLf & "Greater Of" & vbCrLf & "Report." & getBaseColName(.MSFlexGrid1.TextMatrix(1, CInt(txtCol(lngY).Text)))
                        If .MSFlexGrid1.TextMatrix(lngX, CInt(txtCol(lngY).Text)) > .MSFlexGrid1.Text Then
                            .MSFlexGrid1.Text = .MSFlexGrid1.TextMatrix(lngX, CInt(txtCol(lngY).Text))
                        End If
                    Case "Lesser Of"
                        strOperations = strOperations & vbCrLf & "Lesser Of" & vbCrLf & "Report." & getBaseColName(.MSFlexGrid1.TextMatrix(1, CInt(txtCol(lngY).Text)))
                        If .MSFlexGrid1.TextMatrix(lngX, CInt(txtCol(lngY).Text)) < .MSFlexGrid1.Text Then
                            .MSFlexGrid1.Text = .MSFlexGrid1.TextMatrix(lngX, CInt(txtCol(lngY).Text))
                        End If
                End Select
            Else
                Select Case cboOP(lngY - 1).Text
                    Case "+"
                        .MSFlexGrid1.Text = CDbl(.MSFlexGrid1.Text) + CDbl(txtVal(lngY).Text)
                        strOperations = strOperations & vbCrLf & "+" & vbCrLf & "Value." & CDbl(txtVal(lngY).Text)
                    Case "-"
                        .MSFlexGrid1.Text = CDbl(.MSFlexGrid1.Text) - CDbl(txtVal(lngY).Text)
                        strOperations = strOperations & vbCrLf & "-" & vbCrLf & "Value." & CDbl(txtVal(lngY).Text)
                    Case "*"
                        .MSFlexGrid1.Text = CDbl(.MSFlexGrid1.Text) * CDbl(txtVal(lngY).Text)
                        strOperations = strOperations & vbCrLf & "*" & vbCrLf & "Value." & CDbl(txtVal(lngY).Text)
                    Case "/"
                        .MSFlexGrid1.Text = CDbl(.MSFlexGrid1.Text) / CDbl(txtVal(lngY).Text)
                        strOperations = strOperations & vbCrLf & "/" & vbCrLf & "Value." & CDbl(txtVal(lngY).Text)
                    Case "Concatenate"
                        .MSFlexGrid1.Text = .MSFlexGrid1.Text & txtVal(lngY).Text
                        strOperations = strOperations & vbCrLf & "Concatenate" & vbCrLf & "Value." & txtVal(lngY).Text
                    Case "Greater Of"
                        strOperations = strOperations & vbCrLf & "Greater Of" & vbCrLf & "Value." & txtVal(lngY).Text
                        If txtVal(lngY).Text > .MSFlexGrid1.Text Then
                            .MSFlexGrid1.Text = txtVal(lngY).Text
                        End If
                    Case "Lesser Of"
                        strOperations = strOperations & vbCrLf & "Lesser Of" & vbCrLf & "Value." & txtVal(lngY).Text
                        If txtVal(lngY).Text < .MSFlexGrid1.Text Then
                            .MSFlexGrid1.Text = txtVal(lngY).Text
                        End If
                End Select
            End If
        End If
        lngY = lngY + 1
    Wend
    lngX = lngX + 1
Wend
.MSFlexGrid1.Redraw = True

Call FG_AutosizeCols(frmReport.MSFlexGrid1, frmReport, , , True)
Call frmReport.setColNums(frmReport.MSFlexGrid1)

'Add operations to additional invisible grid
lngX = 1
.MSFlexGrid3.Col = 0
While lngX + 1 < .MSFlexGrid3.Rows
    .MSFlexGrid3.Row = lngX
    If UCase(.MSFlexGrid3.Text) = UCase(strHeading) Then
        .MSFlexGrid3.Col = 1
        .MSFlexGrid3.Text = strOperations
    End If
    lngX = lngX + 1
Wend

Unload Me
Exit Sub

Err:
Me.Visible = True
txtVal(lngY).SetFocus
txtVal(lngY).SelStart = 0
txtVal(lngY).SelLength = Len(txtVal(lngY).Text)
MsgBox Err.Description & vbCrLf & "Check value " & txtVal(lngY).Text, vbExclamation
.MSFlexGrid1.Redraw = True
End With
End Sub

Private Sub Form_KeyPress(KeyAscii As Integer)
If KeyAscii = 27 Then
    Unload Me
End If
End Sub

Private Sub Form_Load()
lblCol.Caption = frmReport.MSFlexGrid1.TextMatrix(0, frmReport.MSFlexGrid1.Col)
lblColInv.Caption = frmReport.lngCurCol

cboOP(0).Text = cboOP(0).List(0)
cboOP(1).Text = cboOP(1).List(0)
cboOP(2).Text = cboOP(2).List(0)
cboOP(3).Text = cboOP(3).List(0)

cboFrom(0).Text = cboFrom(0).List(0)
cboFrom(1).Text = cboFrom(1).List(0)
cboFrom(2).Text = cboFrom(2).List(0)
cboFrom(3).Text = cboFrom(3).List(0)
cboFrom(4).Text = cboFrom(4).List(0)
End Sub

Private Sub Form_Unload(Cancel As Integer)
If frmReport.Visible = True Then
    frmReport.SetFocus
End If
End Sub

Private Sub txtVal_GotFocus(Index As Integer)
intIndex = Index
End Sub

Function getBaseColName(strName As String) As String
getBaseColName = ""
Dim lngX As Long
lngX = 1
While lngX + 1 < frmHeadings.MSFlexGrid1.Rows
    If UCase(frmHeadings.MSFlexGrid1.TextMatrix(lngX, 1)) = UCase(strName) Then
        getBaseColName = UCase(frmHeadings.MSFlexGrid1.TextMatrix(lngX, 0))
        Exit Function
    End If
    lngX = lngX + 1
Wend
getBaseColName = strName
End Function

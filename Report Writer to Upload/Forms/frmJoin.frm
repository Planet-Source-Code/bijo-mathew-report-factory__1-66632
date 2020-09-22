VERSION 5.00
Begin VB.Form frmJoin 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Join"
   ClientHeight    =   2385
   ClientLeft      =   45
   ClientTop       =   435
   ClientWidth     =   6465
   HelpContextID   =   1011
   Icon            =   "frmJoin.frx":0000
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   MaxButton       =   0   'False
   MinButton       =   0   'False
   Picture         =   "frmJoin.frx":038A
   ScaleHeight     =   2385
   ScaleWidth      =   6465
   StartUpPosition =   2  'CenterScreen
   Begin VB.ComboBox cboTab2 
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
      Left            =   3720
      Sorted          =   -1  'True
      Style           =   2  'Dropdown List
      TabIndex        =   4
      Top             =   480
      Width           =   2655
   End
   Begin VB.ComboBox cboTab1 
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
      Left            =   120
      Sorted          =   -1  'True
      Style           =   2  'Dropdown List
      TabIndex        =   1
      Top             =   480
      Width           =   2655
   End
   Begin Reports_Factory.ucButtons_H cmdCancel 
      Height          =   375
      Left            =   3360
      TabIndex        =   7
      Top             =   1800
      Width           =   1695
      _ExtentX        =   2990
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
      Image           =   "frmJoin.frx":9511
      cBack           =   -2147483633
      mPointer        =   99
      mIcon           =   "frmJoin.frx":98AB
   End
   Begin Reports_Factory.ucButtons_H cmdConfirm 
      Height          =   375
      Left            =   1560
      TabIndex        =   5
      Top             =   1200
      Width           =   3495
      _ExtentX        =   6165
      _ExtentY        =   661
      Caption         =   "&Confirm  "
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
      Image           =   "frmJoin.frx":9BC5
      cBack           =   -2147483633
      mPointer        =   99
      mIcon           =   "frmJoin.frx":9F5F
   End
   Begin Reports_Factory.ucButtons_H cmdHelp 
      Height          =   375
      Left            =   1560
      TabIndex        =   6
      Top             =   1800
      Width           =   1695
      _ExtentX        =   2990
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
      Image           =   "frmJoin.frx":A279
      cBack           =   -2147483633
      mPointer        =   99
      mIcon           =   "frmJoin.frx":D2FB
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "<< === >>"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00404040&
      Height          =   195
      Left            =   2800
      TabIndex        =   2
      Top             =   550
      UseMnemonic     =   0   'False
      Width           =   870
   End
   Begin VB.Label lblTable2 
      BackStyle       =   0  'Transparent
      Caption         =   "Table 2"
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
      Left            =   3720
      TabIndex        =   3
      Top             =   240
      UseMnemonic     =   0   'False
      Width           =   2580
      WordWrap        =   -1  'True
   End
   Begin VB.Label lblTable1 
      BackStyle       =   0  'Transparent
      Caption         =   "Table 1"
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
      Left            =   120
      TabIndex        =   0
      Top             =   240
      UseMnemonic     =   0   'False
      Width           =   2580
      WordWrap        =   -1  'True
   End
End
Attribute VB_Name = "frmJoin"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private Sub cmdCancel_Click()
boolCancelled = True
Unload Me
End Sub

Private Sub cmdConfirm_Click()
Dim intX As Long
Dim strTemp As String
With frmConnect.MSFlexGrid1
    intX = 1: strTemp = ""
    While intX < .Rows
        .Row = intX
        .Col = 1
        If UCase(.Text) = UCase(lblTable1.Caption) Or UCase(.Text) = UCase(lblTable2.Caption) Then
            .Col = 3
            If UCase(.Text) = UCase(lblTable1.Caption) Or UCase(.Text) = UCase(lblTable2.Caption) Then
                .RemoveItem intX
                GoTo lblG
            End If
        End If
        intX = intX + 1
    Wend
End With

lblG: frmConnect.MSFlexGrid1.AddItem "" & vbTab & lblTable1.Caption & vbTab & cboTab1.Text _
                             & vbTab & lblTable2.Caption & vbTab & cboTab2.Text & vbTab _
                             & frmConnect.fromList & vbTab & frmConnect.toList, 1
Unload Me
End Sub

Private Sub cmdHelp_Click()
Call ShowAppHelp(1011)
End Sub

Private Sub Form_Activate()
boolCancelled = False

Dim intX As Long

intX = 0
cboTab1.Clear
lblTable1.Caption = frmConnect.lblTable(frmConnect.fromList).ToolTipText
While intX < frmConnect.lstFields(frmConnect.fromList).ListCount
    cboTab1.AddItem frmConnect.lstFields(frmConnect.fromList).List(intX)
    intX = intX + 1
Wend
cboTab1.Text = cboTab1.List(0)

intX = 0
cboTab2.Clear
lblTable2.Caption = frmConnect.lblTable(frmConnect.toList).ToolTipText
While intX < frmConnect.lstFields(frmConnect.toList).ListCount
    cboTab2.AddItem frmConnect.lstFields(frmConnect.toList).List(intX)
    intX = intX + 1
Wend
cboTab2.Text = cboTab2.List(0)

intX = 0
Dim intY As Long
intY = 0

While intX < cboTab1.ListCount
    intY = 0
    While intY <= cboTab2.ListCount
        If UCase(cboTab1.List(intX)) = UCase(cboTab2.List(intY)) Then
            cboTab1.Text = cboTab1.List(intX)
            cboTab2.Text = cboTab2.List(intY)
            Exit Sub
        End If
        intY = intY + 1
    Wend
    intX = intX + 1
Wend
End Sub

Private Sub Form_Unload(Cancel As Integer)
boolJoin = True
frmConnect.Enabled = True
frmConnect.SetFocus
End Sub

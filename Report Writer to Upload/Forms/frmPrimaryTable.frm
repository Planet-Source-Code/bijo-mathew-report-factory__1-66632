VERSION 5.00
Begin VB.Form frmPrimaryTable 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Select Primary Table"
   ClientHeight    =   2310
   ClientLeft      =   45
   ClientTop       =   435
   ClientWidth     =   3240
   HelpContextID   =   1012
   Icon            =   "frmPrimaryTable.frx":0000
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   MaxButton       =   0   'False
   MinButton       =   0   'False
   Picture         =   "frmPrimaryTable.frx":038A
   ScaleHeight     =   2310
   ScaleWidth      =   3240
   StartUpPosition =   1  'CenterOwner
   Begin VB.CheckBox chkShowDistinct 
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
      Left            =   240
      TabIndex        =   6
      Top             =   810
      Width           =   200
   End
   Begin VB.ComboBox cboTables 
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
      ItemData        =   "frmPrimaryTable.frx":3389B
      Left            =   240
      List            =   "frmPrimaryTable.frx":338A5
      Sorted          =   -1  'True
      Style           =   2  'Dropdown List
      TabIndex        =   1
      Top             =   360
      Width           =   2895
   End
   Begin Reports_Factory.ucButtons_H cmdCancel 
      Height          =   375
      Left            =   1800
      TabIndex        =   5
      Top             =   1800
      Width           =   1335
      _ExtentX        =   2355
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
      Image           =   "frmPrimaryTable.frx":338C2
      cBack           =   -2147483633
      mPointer        =   99
      mIcon           =   "frmPrimaryTable.frx":33C5C
   End
   Begin Reports_Factory.ucButtons_H cmdConfirm 
      Height          =   375
      Left            =   240
      TabIndex        =   3
      Top             =   1200
      Width           =   2895
      _ExtentX        =   5106
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
      Image           =   "frmPrimaryTable.frx":33F76
      cBack           =   -2147483633
      mPointer        =   99
      mIcon           =   "frmPrimaryTable.frx":34310
   End
   Begin Reports_Factory.ucButtons_H cmdHelp 
      Height          =   375
      Left            =   240
      TabIndex        =   4
      Top             =   1800
      Width           =   1455
      _ExtentX        =   2566
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
      Image           =   "frmPrimaryTable.frx":3462A
      cBack           =   -2147483633
      mPointer        =   99
      mIcon           =   "frmPrimaryTable.frx":376AC
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Show distinct data only"
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
      Left            =   480
      TabIndex        =   2
      Top             =   840
      UseMnemonic     =   0   'False
      Width           =   1995
   End
   Begin VB.Label Label2 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Primary Table:"
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
      Left            =   1080
      TabIndex        =   0
      Top             =   120
      UseMnemonic     =   0   'False
      Width           =   1230
   End
End
Attribute VB_Name = "frmPrimaryTable"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private Sub cmdCancel_Click()
Unload Me
End Sub

Private Sub cmdConfirm_Click()
strPrimaryTable = cboTables.Text
frmStatus.Show , Me
Me.Hide
End Sub

Private Sub cmdHelp_Click()
Call ShowAppHelp(1012)
End Sub

Private Sub Form_Activate()
On Error Resume Next
Me.Visible = False
cboTables.Clear
Dim lngX As Long

While lngX <= frmConnect.lblTable.UBound
    cboTables.AddItem frmConnect.lblTable(lngX).ToolTipText
    lngX = lngX + 1
Wend
cboTables.Text = cboTables.List(0)

Dim intX As Long
intX = 1
Dim strLabels(0 To 500) As Integer
While intX + 1 < frmConnect.MSFlexGrid1.Rows
    frmConnect.MSFlexGrid1.Row = intX
    frmConnect.MSFlexGrid1.Col = 5
    strLabels(CInt(frmConnect.MSFlexGrid1.Text)) = strLabels(CInt(frmConnect.MSFlexGrid1.Text)) + 1
    frmConnect.MSFlexGrid1.Col = 6
    strLabels(CInt(frmConnect.MSFlexGrid1.Text)) = strLabels(CInt(frmConnect.MSFlexGrid1.Text)) + 1
    intX = intX + 1
Wend

Dim intMax As Integer, intPos As Integer
intMax = 0
intX = 0
intPos = 0
While intX <= UBound(strLabels)
    If intMax < strLabels(intX) Then
        intMax = strLabels(intX)
        intPos = intX
    End If
    intX = intX + 1
Wend

cboTables.Text = frmConnect.lblTable(intPos).ToolTipText

If boolFromRun = True And Len(strPrimaryTable) > 0 Then
    cboTables.Text = strPrimaryTable
    If intDistictData = 0 Or intDistictData = 1 Then
        chkShowDistinct.Value = intDistictData
    End If
    Call cmdConfirm_Click
Else
    Me.Visible = True
End If
End Sub

Private Sub Form_Unload(Cancel As Integer)
frmConnect.Enabled = True
End Sub

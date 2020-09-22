VERSION 5.00
Begin VB.Form frmFont 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Set Font"
   ClientHeight    =   1875
   ClientLeft      =   45
   ClientTop       =   435
   ClientWidth     =   3660
   HelpContextID   =   1022
   Icon            =   "frmFont.frx":0000
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   MaxButton       =   0   'False
   MinButton       =   0   'False
   Picture         =   "frmFont.frx":038A
   ScaleHeight     =   1875
   ScaleWidth      =   3660
   StartUpPosition =   1  'CenterOwner
   Begin VB.ComboBox cboFontSize 
      BeginProperty DataFormat 
         Type            =   1
         Format          =   "0"
         HaveTrueFalseNull=   0
         FirstDayOfWeek  =   0
         FirstWeekOfYear =   0
         LCID            =   1033
         SubFormatType   =   1
      EndProperty
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
      ItemData        =   "frmFont.frx":9511
      Left            =   1320
      List            =   "frmFont.frx":951E
      Style           =   2  'Dropdown List
      TabIndex        =   3
      Top             =   720
      Width           =   1215
   End
   Begin VB.ComboBox cboFont 
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
      ItemData        =   "frmFont.frx":9544
      Left            =   1320
      List            =   "frmFont.frx":9551
      Sorted          =   -1  'True
      Style           =   2  'Dropdown List
      TabIndex        =   1
      Top             =   240
      Width           =   2055
   End
   Begin Reports_Factory.ucButtons_H cmdCancel 
      Height          =   375
      Left            =   2040
      TabIndex        =   5
      Top             =   1320
      Width           =   1455
      _extentx        =   2566
      _extenty        =   661
      caption         =   "&Cancel  "
      capalign        =   2
      backstyle       =   2
      cgradient       =   14737632
      cfore           =   0
      font            =   "frmFont.frx":9577
      mode            =   0
      value           =   0
      image           =   "frmFont.frx":959B
      imgalign        =   3
      cfhover         =   0
      cback           =   -2147483633
      micon           =   "frmFont.frx":9935
      mpointer        =   99
   End
   Begin Reports_Factory.ucButtons_H cmdOK 
      Height          =   375
      Left            =   240
      TabIndex        =   4
      Top             =   1320
      Width           =   1575
      _extentx        =   2778
      _extenty        =   661
      caption         =   "&OK  "
      capalign        =   2
      backstyle       =   2
      cgradient       =   14737632
      cfore           =   0
      font            =   "frmFont.frx":9C4F
      mode            =   0
      value           =   0
      image           =   "frmFont.frx":9C73
      imgalign        =   3
      cfhover         =   0
      cback           =   -2147483633
      micon           =   "frmFont.frx":A00D
      mpointer        =   99
   End
   Begin VB.Label Label2 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Font Size:"
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
      Width           =   870
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Font Name:"
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
      Width           =   990
   End
End
Attribute VB_Name = "frmFont"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private Sub cmdCancel_Click()
Unload Me
End Sub

Private Sub cmdOK_Click()
If frmReport.MSFlexGrid1.Visible = True Then
    frmReport.MSFlexGrid1.Font.Name = cboFont.Text
    frmReport.MSFlexGrid1.Font.Size = cboFontSize.Text
    
    Call FG_AutosizeCols(frmReport.MSFlexGrid1, frmReport, , , True)
    Call FG_AutosizeRows(frmReport.MSFlexGrid1, frmReport, , , True)
ElseIf frmReport.MSFlexGrid2.Visible = True Then
    frmReport.MSFlexGrid2.Font.Name = cboFont.Text
    frmReport.MSFlexGrid2.Font.Size = cboFontSize.Text
    
    Call FG_AutosizeCols(frmReport.MSFlexGrid2, frmReport, , , True)
    Call FG_AutosizeRows(frmReport.MSFlexGrid2, frmReport, , , True)
End If
Unload Me
End Sub

Private Sub Form_KeyPress(KeyAscii As Integer)
If KeyAscii = 27 Then
    Unload Me
End If
End Sub

Private Sub Form_Load()
Screen.MousePointer = vbHourglass
Dim intX As Long

intX = 0
cboFont.Clear
While intX <= Screen.FontCount
    If Len(Trim(Screen.Fonts(intX))) > 0 Then
        cboFont.AddItem Screen.Fonts(intX)
    End If
    intX = intX + 1
Wend
cboFont.Text = frmReport.MSFlexGrid1.Font.Name

intX = 1
cboFontSize.Clear
While intX <= 50
    cboFontSize.AddItem intX
    intX = intX + 1
Wend

If chkComboMatch(cboFontSize, Round(frmReport.MSFlexGrid1.Font.Size, 0)) = False Then
    cboFontSize.AddItem Round(frmReport.MSFlexGrid1.Font.Size, 0)
End If

cboFontSize.Text = Round(frmReport.MSFlexGrid1.Font.Size, 0)

Screen.MousePointer = vbNormal
End Sub

Private Sub Form_Unload(Cancel As Integer)
frmReport.SetFocus
End Sub

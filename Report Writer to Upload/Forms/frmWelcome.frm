VERSION 5.00
Begin VB.Form frmWelcome 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Reports Factory"
   ClientHeight    =   4245
   ClientLeft      =   45
   ClientTop       =   435
   ClientWidth     =   7365
   HelpContextID   =   1008
   Icon            =   "frmWelcome.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   Picture         =   "frmWelcome.frx":038A
   ScaleHeight     =   4245
   ScaleWidth      =   7365
   StartUpPosition =   2  'CenterScreen
   Begin VB.Frame Frame1 
      Appearance      =   0  'Flat
      BackColor       =   &H8000000D&
      Caption         =   " I want to . . .  "
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFFFF&
      Height          =   2175
      Left            =   480
      TabIndex        =   4
      Top             =   1320
      Width           =   6375
      Begin VB.OptionButton optHelp 
         BackColor       =   &H8000000D&
         Caption         =   "Get help on Reports Factory"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00FFFFFF&
         Height          =   255
         Left            =   480
         MaskColor       =   &H8000000D&
         TabIndex        =   7
         Top             =   1680
         Width           =   4815
      End
      Begin VB.OptionButton optRun 
         BackColor       =   &H8000000D&
         Caption         =   "Delete / Run a published Reports Factory Template"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00FFFFFF&
         Height          =   495
         Left            =   480
         MaskColor       =   &H8000000D&
         TabIndex        =   6
         Top             =   960
         Width           =   5775
      End
      Begin VB.OptionButton optNew 
         BackColor       =   &H8000000D&
         Caption         =   "Create a new report using Reports Factory Wizard"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00FFFFFF&
         Height          =   240
         Left            =   480
         MaskColor       =   &H8000000D&
         TabIndex        =   5
         Top             =   480
         Width           =   5610
      End
   End
   Begin Reports_Factory.ucButtons_H cmdExit 
      Height          =   375
      Left            =   5160
      TabIndex        =   9
      Top             =   3720
      Width           =   1695
      _ExtentX        =   2990
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
      Image           =   "frmWelcome.frx":9511
      cBack           =   -2147483633
      mPointer        =   99
      mIcon           =   "frmWelcome.frx":98AB
   End
   Begin Reports_Factory.ucButtons_H cmdNext 
      Height          =   375
      Left            =   2760
      TabIndex        =   8
      Top             =   3720
      Width           =   1695
      _ExtentX        =   2990
      _ExtentY        =   661
      Caption         =   "&Next   "
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
      LockHover       =   2
      cGradient       =   -2147483635
      Mode            =   0
      Value           =   0   'False
      ImgAlign        =   3
      Image           =   "frmWelcome.frx":9BC5
      cBack           =   -2147483633
      mPointer        =   99
      mIcon           =   "frmWelcome.frx":9F5F
   End
   Begin VB.Label Label4 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Version 1.0"
      BeginProperty Font 
         Name            =   "Times New Roman"
         Size            =   15.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   -1  'True
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00800000&
      Height          =   360
      Left            =   5520
      TabIndex        =   3
      Top             =   600
      Width           =   1485
   End
   Begin VB.Label Label3 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Welcome to"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   21.75
         Charset         =   222
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00C00000&
      Height          =   525
      Left            =   960
      TabIndex        =   0
      Top             =   120
      Width           =   2595
   End
   Begin VB.Label Label2 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Reports are just a click away!!!"
      BeginProperty Font 
         Name            =   "Times New Roman"
         Size            =   15.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   -1  'True
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00000000&
      Height          =   360
      Left            =   1200
      TabIndex        =   2
      Top             =   600
      Width           =   4080
   End
   Begin VB.Image Image1 
      Height          =   990
      Left            =   120
      Picture         =   "frmWelcome.frx":A279
      Top             =   120
      Width           =   765
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Reports Factory"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   21.75
         Charset         =   222
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00C00000&
      Height          =   525
      Left            =   3675
      TabIndex        =   1
      Top             =   120
      Width           =   3480
   End
End
Attribute VB_Name = "frmWelcome"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private Sub cmdExit_Click()
Unload Me
End Sub

Private Sub cmdNext_Click()
If optHelp.Value = True Then
   Call ShowAppHelp(1007)
ElseIf optNew.Value = True Then
    frmDB.Show
    Me.Hide
ElseIf optRun.Value = True Then
    frmSelectReport.Show
    Me.Hide
End If
End Sub

Private Sub Form_Load()
On Error GoTo err:

Set FSO = CreateObject("Scripting.FileSystemObject")
App.HelpFile = App.Path & "\Help\REPORTSFACTORYHELP.HLP"
Call createLocalDB
Exit Sub

err:
MsgBox err.Description, vbExclamation
Me.Hide
Unload Me
Call unloadAllForms
End Sub

Private Sub Form_Unload(Cancel As Integer)
If Me.Visible = True Then
    boolConfirm = MsgBox("Are you sure you want to exit ?", vbYesNoCancel + vbDefaultButton3 + vbQuestion)
    If boolConfirm <> vbYes Then
        Cancel = 1
        Exit Sub
    End If
    Call unloadAllForms
End If
End Sub

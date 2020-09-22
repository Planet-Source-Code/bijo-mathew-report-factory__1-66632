VERSION 5.00
Object = "{EAB22AC0-30C1-11CF-A7EB-0000C05BAE0B}#1.1#0"; "shdocvw.dll"
Object = "{F9043C88-F6F2-101A-A3C9-08002B2F49FB}#1.2#0"; "Comdlg32.ocx"
Begin VB.Form frmPrint 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Print Preview"
   ClientHeight    =   9240
   ClientLeft      =   45
   ClientTop       =   435
   ClientWidth     =   15210
   HelpContextID   =   1025
   Icon            =   "frmPrint.frx":0000
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   MaxButton       =   0   'False
   Picture         =   "frmPrint.frx":038A
   ScaleHeight     =   9240
   ScaleWidth      =   15210
   StartUpPosition =   2  'CenterScreen
   Begin Reports_Factory.ucButtons_H cmdUp 
      Height          =   1095
      Left            =   14880
      TabIndex        =   2
      Top             =   135
      Width           =   375
      _ExtentX        =   661
      _ExtentY        =   1931
      CapAlign        =   2
      BackStyle       =   4
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   178
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      cBhover         =   -2147483635
      Focus           =   0   'False
      cGradient       =   -2147483635
      Gradient        =   4
      Mode            =   0
      Value           =   0   'False
      Image           =   "frmPrint.frx":3389B
      cBack           =   16777215
      mPointer        =   99
      mIcon           =   "frmPrint.frx":33C35
   End
   Begin MSComDlg.CommonDialog CD 
      Left            =   120
      Top             =   8760
      _ExtentX        =   847
      _ExtentY        =   847
      _Version        =   393216
   End
   Begin VB.Frame Frame1 
      BorderStyle     =   0  'None
      Height          =   8295
      Left            =   120
      TabIndex        =   0
      Top             =   120
      Width           =   14760
      Begin SHDocVwCtl.WebBrowser WebBrowser1 
         CausesValidation=   0   'False
         Height          =   8295
         Left            =   0
         TabIndex        =   1
         Top             =   0
         Width           =   15015
         ExtentX         =   26485
         ExtentY         =   14631
         ViewMode        =   0
         Offline         =   0
         Silent          =   0
         RegisterAsBrowser=   0
         RegisterAsDropTarget=   0
         AutoArrange     =   0   'False
         NoClientEdge    =   0   'False
         AlignLeft       =   0   'False
         NoWebView       =   0   'False
         HideFileNames   =   0   'False
         SingleClick     =   0   'False
         SingleSelection =   0   'False
         NoFolders       =   0   'False
         Transparent     =   0   'False
         ViewID          =   "{0057D0E0-3573-11CF-AE69-08002B2E1262}"
         Location        =   "http:///"
      End
   End
   Begin Reports_Factory.ucButtons_H cmdPrint 
      Height          =   375
      Left            =   4440
      TabIndex        =   5
      Top             =   8760
      Width           =   1815
      _ExtentX        =   3201
      _ExtentY        =   661
      Caption         =   "&Print    "
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
      Image           =   "frmPrint.frx":33F4F
      cBack           =   -2147483633
      mPointer        =   99
      mIcon           =   "frmPrint.frx":342E9
   End
   Begin Reports_Factory.ucButtons_H cmdSave 
      Height          =   375
      Left            =   6480
      TabIndex        =   6
      Top             =   8760
      Width           =   1815
      _ExtentX        =   3201
      _ExtentY        =   661
      Caption         =   "&Save Report  "
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
      Image           =   "frmPrint.frx":34603
      cBack           =   -2147483633
      mPointer        =   99
      mIcon           =   "frmPrint.frx":3499D
   End
   Begin Reports_Factory.ucButtons_H cmdExit 
      Height          =   375
      Left            =   10560
      TabIndex        =   8
      Top             =   8760
      Width           =   1815
      _ExtentX        =   3201
      _ExtentY        =   661
      Caption         =   "&Finish  "
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
      Image           =   "frmPrint.frx":34CB7
      cBack           =   -2147483633
      mPointer        =   99
      mIcon           =   "frmPrint.frx":35051
   End
   Begin Reports_Factory.ucButtons_H cmdHelp 
      Height          =   375
      Left            =   8520
      TabIndex        =   7
      Top             =   8760
      Width           =   1815
      _ExtentX        =   3201
      _ExtentY        =   661
      Caption         =   "&Help  "
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
      Image           =   "frmPrint.frx":3536B
      cBack           =   -2147483633
      mPointer        =   99
      mIcon           =   "frmPrint.frx":383ED
   End
   Begin Reports_Factory.ucButtons_H cmdPrevious 
      Height          =   375
      Left            =   2400
      TabIndex        =   4
      Top             =   8760
      Width           =   1815
      _ExtentX        =   3201
      _ExtentY        =   661
      Caption         =   "   &Previous"
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
      ImgAlign        =   1
      Image           =   "frmPrint.frx":38707
      cBack           =   -2147483633
      mPointer        =   99
      mIcon           =   "frmPrint.frx":38AA1
   End
   Begin Reports_Factory.ucButtons_H cmdDown 
      Height          =   1095
      Left            =   14880
      TabIndex        =   3
      Top             =   1200
      Width           =   375
      _ExtentX        =   661
      _ExtentY        =   1931
      CapAlign        =   2
      BackStyle       =   4
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   178
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      cBhover         =   -2147483635
      Focus           =   0   'False
      cGradient       =   -2147483635
      Gradient        =   3
      Mode            =   0
      Value           =   0   'False
      Image           =   "frmPrint.frx":38DBB
      cBack           =   16777215
      mPointer        =   99
      mIcon           =   "frmPrint.frx":39155
   End
End
Attribute VB_Name = "frmPrint"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private Sub cmdDown_Click()
Frame1.Enabled = True
WebBrowser1.SetFocus
SendKeys "{PGDN}", True
Frame1.Enabled = False
cmdDown.SetFocus
End Sub

Private Sub cmdHelp_Click()
Call ShowAppHelp(1025)
End Sub

Private Sub cmdPrevious_Click()
frmReport.Visible = True
Me.Hide
End Sub

Private Sub cmdPrint_Click()
On Error Resume Next
Screen.MousePointer = vbHourglass
Dim strFooter As String
Dim strHeader As String

'store HEADER & FOOTER
strFooter = QueryValue("Software\Microsoft\Internet Explorer\PageSetup", "footer")
strHeader = QueryValue("Software\Microsoft\Internet Explorer\PageSetup", "header")

'our HEADER & FOOTER
SetKeyValue "Software\Microsoft\Internet Explorer\PageSetup", "header", "", REG_SZ
SetKeyValue "Software\Microsoft\Internet Explorer\PageSetup", "footer", "Report generated using Reports Factory on " & Format(Now, "dd-MMM-yyyy"), REG_SZ

WebBrowser1.ExecWB OLECMDID_PRINT, OLECMDEXECOPT_PROMPTUSER

'replace HEADER & FOOTER with old value
SetKeyValue "Software\Microsoft\Internet Explorer\PageSetup", "footer", strFooter, REG_SZ
SetKeyValue "Software\Microsoft\Internet Explorer\PageSetup", "header", strHeader, REG_SZ

Screen.MousePointer = vbDefault
End Sub

Private Sub cmdExit_Click()
Unload Me
End Sub

Private Sub cmdSave_Click()
Dim strFileName As String
Dim boolOnce As Boolean
boolOnce = False
strFileName = Replace(LCase(WebBrowser1.LocationURL), "file:///", "")
strFileName = Replace(strFileName, "%20", " ")
strFileName = Replace(strFileName, "/", "\")

CD.CancelError = True
CD.DialogTitle = "Save report..."
CD.Filter = "MS Excel (*.xls)|*.xls|HTM (*.htm)|*.htm|MS Word (*.doc)|*.doc"

CD.FileName = Replace(GetVirtualFileName(strFileName), "." & chkFileExtension(strFileName), "")
If Len(Trim(frmReport.txtReportTitle.Text)) > 0 Then
    CD.FileName = frmReport.txtReportTitle.Text
End If
On Error GoTo err1
CD.ShowSave

If Len(CD.FileName) > 0 Then
    If chkFilePath(CD.FileName) = False Then
        FSO.CopyFile strFileName, CD.FileName
    ElseIf chkFilePath(CD.FileName) = True Then
        boolConfirm = MsgBox("This file already exist. Do you want to overwrite this file ?", vbYesNoCancel + vbDefaultButton3 + vbQuestion)
        If boolConfirm = vbYes Then
            On Error GoTo err1
            FSO.CopyFile strFileName, CD.FileName, True
        Else
            Exit Sub
        End If
    End If
    MsgBox "File saved to " & CD.FileName, vbInformation
End If
Exit Sub
    
err1:
If boolOnce = False And UCase(err.Description) <> "CANCEL WAS SELECTED." Then
    CD.FileName = Replace(GetVirtualFileName(strFileName), "." & chkFileExtension(strFileName), "")
    boolOnce = True
    Resume
End If

MsgBox err.Description, vbExclamation
Exit Sub

End Sub

Private Sub cmdUp_Click()
Frame1.Enabled = True
WebBrowser1.SetFocus
SendKeys "{PGUP}", True
Frame1.Enabled = False
cmdUp.SetFocus
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

Private Sub WebBrowser1_ProgressChange(ByVal Progress As Long, ByVal ProgressMax As Long)
If Progress = 0 Then
'    WebBrowser1.ExecWB OLECMDID_PRINTPREVIEW, OLECMDEXECOPT_PROMPTUSER, Null, Null
'    Unload Me
End If
End Sub

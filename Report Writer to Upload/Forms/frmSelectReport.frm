VERSION 5.00
Object = "{5E9E78A0-531B-11CF-91F6-C2863C385E30}#1.0#0"; "Msflxgrd.ocx"
Begin VB.Form frmSelectReport 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Published Reports"
   ClientHeight    =   5805
   ClientLeft      =   45
   ClientTop       =   435
   ClientWidth     =   11445
   HelpContextID   =   1014
   Icon            =   "frmSelectReport.frx":0000
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   MaxButton       =   0   'False
   MinButton       =   0   'False
   Picture         =   "frmSelectReport.frx":038A
   ScaleHeight     =   5805
   ScaleWidth      =   11445
   StartUpPosition =   2  'CenterScreen
   Begin MSFlexGridLib.MSFlexGrid MSFlexGrid1 
      Height          =   4815
      Left            =   120
      TabIndex        =   0
      Top             =   240
      Width           =   11175
      _ExtentX        =   19711
      _ExtentY        =   8493
      _Version        =   393216
      Cols            =   4
      BackColorFixed  =   12615680
      ForeColorFixed  =   16777215
      FocusRect       =   0
      AllowUserResizing=   1
      FormatString    =   $"frmSelectReport.frx":3389B
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   178
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
   End
   Begin Reports_Factory.ucButtons_H cmdExit 
      Height          =   375
      Left            =   9600
      TabIndex        =   5
      Top             =   5280
      Width           =   1695
      _ExtentX        =   2990
      _ExtentY        =   661
      Caption         =   "&E&xit  "
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
      Image           =   "frmSelectReport.frx":3393A
      cBack           =   -2147483633
      mPointer        =   99
      mIcon           =   "frmSelectReport.frx":33CD4
   End
   Begin Reports_Factory.ucButtons_H cmdHelp 
      Height          =   375
      Left            =   7680
      TabIndex        =   4
      Top             =   5280
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
      Image           =   "frmSelectReport.frx":33FEE
      cBack           =   -2147483633
      mPointer        =   99
      mIcon           =   "frmSelectReport.frx":37070
   End
   Begin Reports_Factory.ucButtons_H cmdDelete 
      Height          =   375
      Left            =   2520
      TabIndex        =   2
      Top             =   5280
      Width           =   1695
      _ExtentX        =   2990
      _ExtentY        =   661
      Caption         =   "&Delete Report  "
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
      Image           =   "frmSelectReport.frx":3738A
      cBack           =   -2147483633
      mPointer        =   99
      mIcon           =   "frmSelectReport.frx":37724
   End
   Begin Reports_Factory.ucButtons_H cmdRun 
      Height          =   375
      Left            =   5040
      TabIndex        =   3
      Top             =   5280
      Width           =   1695
      _ExtentX        =   2990
      _ExtentY        =   661
      Caption         =   "&Run Report   "
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
      Image           =   "frmSelectReport.frx":37A3E
      cBack           =   -2147483633
      mPointer        =   99
      mIcon           =   "frmSelectReport.frx":37DD8
   End
   Begin Reports_Factory.ucButtons_H cmdPrevious 
      Height          =   375
      Left            =   480
      TabIndex        =   1
      Top             =   5280
      Width           =   1695
      _ExtentX        =   2990
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
      Image           =   "frmSelectReport.frx":380F2
      cBack           =   -2147483633
      mPointer        =   99
      mIcon           =   "frmSelectReport.frx":3848C
   End
End
Attribute VB_Name = "frmSelectReport"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private Sub cmdDelete_Click()
Dim lngX As Long
lngX = 1
MSFlexGrid1.Redraw = False
MSFlexGrid1.Col = 0
While lngX + 1 <= MSFlexGrid1.Rows
    If MSFlexGrid1.TextMatrix(lngX, 0) <> "" Then
        boolConfirm = MsgBox("Are you sure you want to delete this Reports Factory Template : " & MSFlexGrid1.TextMatrix(lngX, 1) & " ?", vbYesNoCancel + vbDefaultButton3 + vbQuestion)
        If boolConfirm = vbYes Then
            FSO.deletefile App.Path & "\Publish\" & MSFlexGrid1.TextMatrix(lngX, 1), True
            If lngX > 1 Then
                MSFlexGrid1.RemoveItem lngX
            Else
                MSFlexGrid1.TextMatrix(1, 0) = ""
                MSFlexGrid1.TextMatrix(1, 1) = ""
                MSFlexGrid1.TextMatrix(1, 2) = ""
                MSFlexGrid1.TextMatrix(1, 3) = ""
            End If
            MsgBox "Report template deleted", vbInformation
        Else
            MSFlexGrid1.Redraw = True
            Exit Sub
        End If
        MSFlexGrid1.Redraw = True
        Exit Sub
    End If
    lngX = lngX + 1
Wend

Call AltFlexiColors(MSFlexGrid1, 1, 1)
MSFlexGrid1.Redraw = True

MsgBox "Please select a report to delete", vbExclamation

End Sub

Private Sub cmdExit_Click()
Unload Me
End Sub

Private Sub cmdHelp_Click()
Call ShowAppHelp(1014)
End Sub

Private Sub cmdPrevious_Click()
frmWelcome.Show
Me.Hide
End Sub

Private Sub cmdRun_Click()
Dim lngX As Long
lngX = 1
MSFlexGrid1.Redraw = False
MSFlexGrid1.Col = 0
While lngX + 1 <= MSFlexGrid1.Rows
    If MSFlexGrid1.TextMatrix(lngX, 0) <> "" Then
        strTemplateFileName = MSFlexGrid1.TextMatrix(lngX, 1)
        boolConfirm = MsgBox("Are you sure you want to run this Reports Factory Template : " & MSFlexGrid1.TextMatrix(lngX, 1) & " ?", vbYesNoCancel + vbDefaultButton3 + vbQuestion)
        If boolConfirm = vbYes Then
            If checkPassword = True Then
                MSFlexGrid1.Redraw = True
                boolFromRun = True
                boolFromPrevious = False
                frmConnect.Show
                Me.Hide
                Exit Sub
            Else
                MSFlexGrid1.Redraw = True
                Exit Sub
            End If
        Else
            MSFlexGrid1.Redraw = True
            Exit Sub
        End If
    End If
    lngX = lngX + 1
Wend
Call AltFlexiColors(MSFlexGrid1, 1, 1)
MSFlexGrid1.Redraw = True

MsgBox "Please select a report to run", vbExclamation

End Sub

Private Sub Form_Activate()
Dim oFolder, oFiles, oFile
Dim myFile
Dim strRptName As String, strRptDesc As String
Dim strFileRead As String
MSFlexGrid1.Rows = 1
MSFlexGrid1.Rows = 2
If fchkFolderPath(App.Path & "\Publish", True) = True Then
    Set oFolder = FSO.GetFolder(App.Path & "\Publish")
    Set oFiles = oFolder.Files
    For Each oFile In oFiles
        If Left(UCase(oFile.Type), 3) = "RFT" Then
            Set myFile = FSO.OpenTextFile(App.Path & "\Publish\" & oFile.Name, 1, -2)
            strFileRead = myFile.readline
            strRptName = "": strRptDesc = ""
            Do While myFile.AtEndOfStream <> True
                If Left(UCase(strFileRead), Len("REPORT NAME:=")) = "REPORT NAME:=" Then
                    strRptName = Mid(strFileRead, Len("REPORT NAME:=") + 1)
                ElseIf Left(UCase(strFileRead), Len("REPORT DESCRIPTION:=")) = "REPORT DESCRIPTION:=" Then
                    strRptDesc = Mid(strFileRead, Len("REPORT DESCRIPTION:=") + 1)
                End If
                strFileRead = myFile.readline
            Loop
            MSFlexGrid1.AddItem "" & vbTab & oFile.Name & vbTab & strRptName & vbTab & strRptDesc, 1
        End If
    Next
End If

On Error Resume Next
MSFlexGrid1.Rows = MSFlexGrid1.Rows - 1

Dim lngX As Long
lngX = 1
MSFlexGrid1.Col = 0
While lngX + 1 <= MSFlexGrid1.Rows
    MSFlexGrid1.Row = lngX
    MSFlexGrid1.CellFontName = "Wingdings"
    lngX = lngX + 1
Wend

lngX = 1
While lngX + 1 <= MSFlexGrid1.Cols
    MSFlexGrid1.ColAlignment(lngX) = 1
    lngX = lngX + 1
Wend

Call AltFlexiColors(MSFlexGrid1, 1, 1)
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

Private Sub MSFlexGrid1_Click()
With MSFlexGrid1
.Redraw = False

If .Row = 0 Or .MouseCol <> 0 Then
    'nothing
Else
    Dim lngTemp As Long
    Dim boolTicked As Boolean
    lngTemp = MSFlexGrid1.Row

    If .TextMatrix(lngTemp, 0) = "ü" Then
        boolTicked = True
        .TextMatrix(lngTemp, 0) = ""
    Else
        boolTicked = False
        .TextMatrix(lngTemp, 0) = "ü"
    End If
    
    Dim lngX As Long
    lngX = 1
    MSFlexGrid1.Col = 0
    While lngX + 1 <= MSFlexGrid1.Rows
        MSFlexGrid1.TextMatrix(lngX, 0) = ""
        lngX = lngX + 1
    Wend
    
    If boolTicked = True Then
        .TextMatrix(lngTemp, 0) = ""
    Else
        .TextMatrix(lngTemp, 0) = "ü"
    End If
End If
.Redraw = True

End With
End Sub

Private Sub MSFlexGrid1_DblClick()
If MSFlexGrid1.Col = 0 Then
    Exit Sub
End If

Call SortFlexiArrows(MSFlexGrid1, False, False)
Call AltFlexiColors(MSFlexGrid1, 1, 1)
End Sub

Function checkPassword() As Boolean
checkPassword = True
Dim strPassword As String, strEnterPwd As String
Dim myFile, strFileRead
strPassword = "": strEnterPwd = ""

Set myFile = FSO.OpenTextFile(App.Path & "\Publish\" & strTemplateFileName, 1, -2)
strFileRead = Trim(myFile.readline)
While myFile.AtEndOfStream <> True
    If Left(UCase(strFileRead), Len("Password:=")) = UCase("Password:=") Then
        strPassword = Mid(strFileRead, Len("Password:=") + 1)
        If Len(strPassword) > 0 Then
            strEnterPwd = InputBox("This report is password protected. You have to enter the password to run this report", "Enter Password")
            If strEnterPwd <> strPassword Then
                MsgBox "The password entered by you is invalid", vbExclamation
                checkPassword = False
                Exit Function
            End If
        End If
    End If
    strFileRead = Trim(myFile.readline)
Wend
Set myFile = Nothing
End Function

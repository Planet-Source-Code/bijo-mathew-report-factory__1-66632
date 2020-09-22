VERSION 5.00
Begin VB.Form frmSaveReport 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Publish Report"
   ClientHeight    =   2760
   ClientLeft      =   45
   ClientTop       =   435
   ClientWidth     =   6150
   HelpContextID   =   1024
   Icon            =   "frmSaveReport.frx":0000
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   MaxButton       =   0   'False
   MinButton       =   0   'False
   Picture         =   "frmSaveReport.frx":038A
   ScaleHeight     =   2760
   ScaleWidth      =   6150
   StartUpPosition =   2  'CenterScreen
   Begin VB.ListBox lstTables 
      Height          =   255
      ItemData        =   "frmSaveReport.frx":9511
      Left            =   4440
      List            =   "frmSaveReport.frx":9513
      TabIndex        =   0
      Top             =   0
      Visible         =   0   'False
      Width           =   1815
   End
   Begin VB.TextBox txtPassword 
      Height          =   315
      Left            =   2160
      MaxLength       =   150
      TabIndex        =   6
      Top             =   1440
      Width           =   3735
   End
   Begin VB.TextBox txtReportDesc 
      Height          =   315
      Left            =   2160
      MaxLength       =   150
      TabIndex        =   4
      Top             =   840
      Width           =   3735
   End
   Begin VB.TextBox txtReportName 
      Height          =   315
      Left            =   2160
      MaxLength       =   50
      TabIndex        =   2
      Top             =   240
      Width           =   3735
   End
   Begin Reports_Factory.ucButtons_H cmdSave 
      Height          =   375
      Left            =   2160
      TabIndex        =   8
      Top             =   2160
      Width           =   1815
      _ExtentX        =   3201
      _ExtentY        =   661
      Caption         =   "&Publish Report "
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
      Image           =   "frmSaveReport.frx":9515
      cBack           =   -2147483633
      mPointer        =   99
      mIcon           =   "frmSaveReport.frx":98AF
   End
   Begin Reports_Factory.ucButtons_H cmdExit 
      Height          =   375
      Left            =   4200
      TabIndex        =   9
      Top             =   2160
      Width           =   1815
      _ExtentX        =   3201
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
      Image           =   "frmSaveReport.frx":9BC9
      cBack           =   -2147483633
      mPointer        =   99
      mIcon           =   "frmSaveReport.frx":9F63
   End
   Begin Reports_Factory.ucButtons_H cmdHelp 
      Height          =   375
      Left            =   120
      TabIndex        =   7
      Top             =   2160
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
      Image           =   "frmSaveReport.frx":A27D
      cBack           =   -2147483633
      mPointer        =   99
      mIcon           =   "frmSaveReport.frx":D2FF
   End
   Begin VB.Label Label3 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Password:"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   178
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   195
      Left            =   240
      TabIndex        =   5
      Top             =   1560
      Width           =   885
   End
   Begin VB.Label Label2 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Report Description:"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   178
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   195
      Left            =   240
      TabIndex        =   3
      Top             =   960
      Width           =   1665
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Report Name:"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   178
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   195
      Left            =   240
      TabIndex        =   1
      Top             =   360
      Width           =   1185
   End
End
Attribute VB_Name = "frmSaveReport"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private Sub cmdExit_Click()
Unload Me
End Sub

Private Sub cmdHelp_Click()
Call ShowAppHelp(1024)
End Sub

Private Sub cmdSave_Click()
If Len(Trim(txtReportName.Text)) <= 0 Then
    MsgBox "You have to enter a report name to save", vbExclamation
ElseIf Len(Trim(txtReportDesc.Text)) <= 0 Then
    MsgBox "You have to enter a report description to save", vbExclamation
Else
    Dim lngX As Long
    Call fchkFolderPath(App.Path & "\Publish", True)
    lngX = 1
    While chkFilePath(App.Path & "\Publish\Published " & lngX & ".rft") = True
        lngX = lngX + 1
    Wend
    Dim myFile
    Set myFile = FSO.CreateTextFile(App.Path & "\Publish\Published " & lngX & ".rft", False)
    With myFile
        .writeline "*** Warning : Do not edit this file ***"
        .writeline "Report Name:=" & Trim(txtReportName.Text)
        .writeline "Report Description:=" & Trim(txtReportDesc.Text)
        If Len(Trim(txtPassword.Text)) > 0 Then
            .writeline "Password:=" & Trim(txtPassword.Text)
        End If
        .writeline "Report Title:=" & Trim(frmReport.txtReportTitle.Text)
        .writeline "------------------------------------------------------------------------------------------"
        
        .writeline "[DB START]"
            .writeline "1:=" & strConString(0)
            .writeline "2:=" & strConString(1)
            .writeline "3:=" & strConString(2)
            .writeline "4:=" & strConString(3)
            .writeline "5:=" & strConString(4)
            .writeline "6:=" & strConString(5)
            .writeline "7:=" & strConString(6)
            .writeline "8:=" & strConString(7)
            .writeline "9:=" & strConString(8)
        .writeline "[DB END]"
        .writeline "------------------------------------------------------------------------------------------"
        
        Dim intY As Integer
        Dim strTables() As String
        Dim lngA As Long, lngB As Long
        intY = 0
        While intY <= 9
            If Len(Trim(strConString(intY).ConnectionString)) > 0 Then
                .writeline "[DB START]:=" & intY + 1
                    lngA = 1
                    frmHeadings.MSFlexGrid1.Col = 0
                    lstTables.Clear
                    'add tables concerned with each DB
                    While lngA + 1 < frmHeadings.MSFlexGrid1.Rows
                        frmHeadings.MSFlexGrid1.Row = lngA
                        strTables = Split(frmHeadings.MSFlexGrid1.Text, ".")
                        If CInt(strTables(0)) = intY + 1 Then
                            If chkListMatch(Me.lstTables, strTables(1)) = False Then
                                lstTables.AddItem strTables(1), 0
                            End If
                        End If
                        lngA = lngA + 1
                    Wend
                    
                    'add fields to each table and db
                    lngB = 0
                    While lngB <= lstTables.ListCount
                        .writeline "[TABLE START]:=" & lstTables.List(lngB)
                            lngA = 1
                            frmHeadings.MSFlexGrid1.Col = 0
                            While lngA + 1 < frmHeadings.MSFlexGrid1.Rows
                                frmHeadings.MSFlexGrid1.Row = lngA
                                strTables = Split(frmHeadings.MSFlexGrid1.Text, ".")
                                If CInt(strTables(0)) = intY + 1 And UCase(strTables(1)) = UCase(lstTables.List(lngB)) Then
                                    .writeline UCase(strTables(2))
                                End If
                                lngA = lngA + 1
                            Wend
                        .writeline "[TABLE END]"
                        lngB = lngB + 1
                    Wend
                .writeline "[DB END]:=" & intY + 1
            End If
            intY = intY + 1
        Wend
        .writeline "------------------------------------------------------------------------------------------"
        
        'write all relationships
        lngA = 1
        frmConnect.MSFlexGrid1.Row = 1
        .writeline "[RELATIONSHIP START]"
        While lngA + 1 < frmConnect.MSFlexGrid1.Rows
            frmConnect.MSFlexGrid1.Row = lngA
            .writeline frmConnect.MSFlexGrid1.TextMatrix(lngA, 1) & vbTab _
                    & frmConnect.MSFlexGrid1.TextMatrix(lngA, 2) & vbTab _
                    & frmConnect.MSFlexGrid1.TextMatrix(lngA, 3) & vbTab _
                    & frmConnect.MSFlexGrid1.TextMatrix(lngA, 4) & vbTab _
                    & frmConnect.MSFlexGrid1.TextMatrix(lngA, 5) & vbTab _
                    & frmConnect.MSFlexGrid1.TextMatrix(lngA, 6)
            
            lngA = lngA + 1
        Wend
        .writeline "[RELATIONSHIP END]"
        .writeline "------------------------------------------------------------------------------------------"
        .writeline "[PRIMARY TABLE]:=" & strPrimaryTable
        .writeline "------------------------------------------------------------------------------------------"
        
        'Distinct Data
        .writeline "[DISTINCT DATA]:=" & intDistictData
        .writeline "------------------------------------------------------------------------------------------"
        
        'write table aliases
        .writeline "[TABLE ALIAS START]"
        lngA = 1
        While lngA + 1 < frmHeadings.MSFlexGrid1.Rows
            frmHeadings.MSFlexGrid1.Row = lngA
            .writeline UCase(frmHeadings.MSFlexGrid1.TextMatrix(lngA, 0)) & vbTab _
                    & UCase(frmHeadings.MSFlexGrid1.TextMatrix(lngA, 0)) & vbTab _
                    & UCase(frmHeadings.MSFlexGrid1.TextMatrix(lngA, 1))
            lngA = lngA + 1
        Wend
        .writeline "[TABLE ALIAS END]"
        
        .writeline "------------------------------------------------------------------------------------------"
        'write calculated cols
        .writeline "[CALCULATED COLUMNS START]"
        lngA = 1
        While lngA + 1 < frmReport.MSFlexGrid3.Rows
            frmReport.MSFlexGrid3.Row = lngA
            frmReport.MSFlexGrid3.Col = 0
            .writeline "[CALCULATED COLUMN]:=" & frmReport.MSFlexGrid3.Text
            frmReport.MSFlexGrid3.Col = 1
            .writeline "[CALCULATION START]"
            .writeline frmReport.MSFlexGrid3.Text
            .writeline "[CALCULATION END]"
            lngA = lngA + 1
        Wend
        .writeline "[CALCULATED COLUMNS END]"
        
        .writeline "------------------------------------------------------------------------------------------"
        .writeline "[DELETED COLS START]"
        lngA = 1
        frmHeadings.MSFlexGrid1.Col = 1
        While lngA + 1 < frmHeadings.MSFlexGrid1.Rows
            frmHeadings.MSFlexGrid1.Row = lngA
            lngB = 1
            frmReport.MSFlexGrid1.Row = 1
            While lngB + 1 <= frmReport.MSFlexGrid1.Cols
                frmReport.MSFlexGrid1.Col = lngB
                If UCase(frmReport.MSFlexGrid1.Text) = UCase(frmHeadings.MSFlexGrid1.Text) Then
                    GoTo lblCont
                End If
                lngB = lngB + 1
            Wend
            frmHeadings.MSFlexGrid1.Col = 0
            .writeline UCase(frmHeadings.MSFlexGrid1.Text)
            frmHeadings.MSFlexGrid1.Col = 1
lblCont:    lngA = lngA + 1
        Wend
        
        .writeline "[DELETED COLS END]"
        .writeline "------------------------------------------------------------------------------------------"
        
        .Close
        MsgBox "Report template saved successfully as : " & "Published " & lngX & ".rft", vbInformation
        Unload Me
    End With
End If
End Sub

Private Sub Form_Unload(Cancel As Integer)
frmReport.Enabled = True
frmReport.SetFocus
End Sub

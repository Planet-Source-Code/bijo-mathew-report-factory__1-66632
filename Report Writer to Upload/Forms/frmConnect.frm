VERSION 5.00
Object = "{5E9E78A0-531B-11CF-91F6-C2863C385E30}#1.0#0"; "Msflxgrd.ocx"
Begin VB.Form frmConnect 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Step 3 >>>  Define Relationship / Connections"
   ClientHeight    =   10200
   ClientLeft      =   45
   ClientTop       =   435
   ClientWidth     =   15270
   HelpContextID   =   1011
   Icon            =   "frmConnect.frx":0000
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   MaxButton       =   0   'False
   MinButton       =   0   'False
   Picture         =   "frmConnect.frx":038A
   ScaleHeight     =   10200
   ScaleWidth      =   15270
   StartUpPosition =   2  'CenterScreen
   Begin VB.ListBox lstExecuted 
      Height          =   840
      ItemData        =   "frmConnect.frx":3389B
      Left            =   12960
      List            =   "frmConnect.frx":3389D
      TabIndex        =   10
      Top             =   960
      Visible         =   0   'False
      Width           =   2295
   End
   Begin VB.ListBox lstCheck 
      Height          =   840
      ItemData        =   "frmConnect.frx":3389F
      Left            =   12960
      List            =   "frmConnect.frx":338A1
      TabIndex        =   3
      Top             =   0
      Visible         =   0   'False
      Width           =   2295
   End
   Begin VB.Frame Frame1 
      Caption         =   "Relationships"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   178
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00000080&
      Height          =   2295
      Left            =   120
      TabIndex        =   5
      Top             =   7200
      Width           =   15015
      Begin MSFlexGridLib.MSFlexGrid MSFlexGrid1 
         Height          =   1935
         Left            =   240
         TabIndex        =   4
         Top             =   240
         Width           =   14535
         _ExtentX        =   25638
         _ExtentY        =   3413
         _Version        =   393216
         Cols            =   8
         BackColorFixed  =   12615680
         ForeColorFixed  =   16777215
         FocusRect       =   0
         ScrollBars      =   2
         SelectionMode   =   1
         Appearance      =   0
         FormatString    =   $"frmConnect.frx":338A3
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
      End
   End
   Begin VB.TextBox lblTable 
      Alignment       =   2  'Center
      Appearance      =   0  'Flat
      BackColor       =   &H00800000&
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   178
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFFFF&
      Height          =   525
      Index           =   0
      Left            =   360
      Locked          =   -1  'True
      MultiLine       =   -1  'True
      ScrollBars      =   2  'Vertical
      TabIndex        =   1
      Top             =   840
      Width           =   1810
   End
   Begin VB.ListBox lstFields 
      Appearance      =   0  'Flat
      DragIcon        =   "frmConnect.frx":339B0
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   178
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   1200
      Index           =   0
      ItemData        =   "frmConnect.frx":33CBA
      Left            =   360
      List            =   "frmConnect.frx":33CBC
      Sorted          =   -1  'True
      TabIndex        =   2
      Top             =   1365
      Width           =   1815
   End
   Begin Reports_Factory.ucButtons_H cmdExit 
      Height          =   375
      Left            =   13440
      TabIndex        =   9
      Top             =   9720
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
      Image           =   "frmConnect.frx":33CBE
      cBack           =   -2147483633
      mPointer        =   99
      mIcon           =   "frmConnect.frx":34058
   End
   Begin Reports_Factory.ucButtons_H cmdNext 
      Height          =   375
      Left            =   7800
      TabIndex        =   7
      Top             =   9720
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
      cGradient       =   14737632
      Mode            =   0
      Value           =   0   'False
      ImgAlign        =   3
      Image           =   "frmConnect.frx":34372
      cBack           =   -2147483633
      mPointer        =   99
      mIcon           =   "frmConnect.frx":3470C
   End
   Begin Reports_Factory.ucButtons_H cmdPrevious 
      Height          =   375
      Left            =   5880
      TabIndex        =   6
      Top             =   9720
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
      Image           =   "frmConnect.frx":34A26
      cBack           =   -2147483633
      mPointer        =   99
      mIcon           =   "frmConnect.frx":34DC0
   End
   Begin Reports_Factory.ucButtons_H cmdHelp 
      Height          =   375
      Left            =   11520
      TabIndex        =   8
      Top             =   9720
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
      Image           =   "frmConnect.frx":350DA
      cBack           =   -2147483633
      mPointer        =   99
      mIcon           =   "frmConnect.frx":3815C
   End
   Begin VB.Line l 
      BorderWidth     =   2
      Index           =   0
      Visible         =   0   'False
      X1              =   480
      X2              =   120
      Y1              =   480
      Y2              =   480
   End
   Begin VB.Label Label2 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "You can drag tables around or drag and drop fields to create relationships"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   178
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00000000&
      Height          =   195
      Left            =   120
      TabIndex        =   0
      Top             =   120
      UseMnemonic     =   0   'False
      Width           =   6300
   End
   Begin VB.Menu mnuOption 
      Caption         =   "Options"
      Visible         =   0   'False
      Begin VB.Menu mnuDel 
         Caption         =   "Delete Relationship"
      End
   End
End
Attribute VB_Name = "frmConnect"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Private Declare Function SendMessage Lib "user32" Alias "SendMessageA" (ByVal hwnd As Long, ByVal wMsg As Long, ByVal wParam As Long, lParam As Any) As Long
Private Declare Function ReleaseCapture Lib "user32" () As Long
Private Const WM_NCLBUTTONDOWN = &HA1
Private Const HTCAPTION = 2
Public fromList As Integer, toList As Integer

Private Sub cmdExit_Click()
Unload Me
End Sub

Private Sub cmdHelp_Click()
Call ShowAppHelp(1011)
End Sub

Private Sub cmdNext_Click()
frmPrimaryTable.Show , Me
Me.Enabled = False
End Sub

Private Sub cmdPrevious_Click()
If boolFromRun = True Then
    frmSelectReport.Visible = True
    Me.Hide
Else
    frmTables.Visible = True
    Me.Hide
End If
End Sub

Private Sub Form_Activate()
On Error GoTo err
strReportTitle = ""
If boolFromRun = True Then
    Dim myFile
    Dim strFileRead As String
    Dim strTableName As String
    Dim strFields() As String
    Dim intDB As Integer
    Call clearAllControls
    intDB = 0
    Set myFile = FSO.OpenTextFile(App.Path & "\Publish\" & strTemplateFileName, 1, -2)
    While intDB <= 9
        strTableName = ""
        strFileRead = Trim(myFile.readline)
        While myFile.AtEndOfStream <> True
        
            If Left(UCase(strFileRead), Len("Report Title:=")) = UCase("Report Title:=") Then
                strReportTitle = Mid(UCase(strFileRead), Len("Report Title:=") + 1)
            ElseIf UCase(strFileRead) = "[DB START]" Then
                Dim intDBCount As Long
                intDBCount = 0
                While intDBCount <= 8
                    strFileRead = Mid(Trim(myFile.readline), 4)
                    If Len(strFileRead) > 0 Then
                        strConString(intDBCount).ConnectionString = strFileRead
                    End If
                    intDBCount = intDBCount + 1
                Wend
            ElseIf UCase(strFileRead) = UCase("[DB START]:=" & intDB) Then
                
            ElseIf Left(UCase(strFileRead), Len("[TABLE START]:=")) = "[TABLE START]:=" Then
                
                strTableName = Mid(strFileRead, Len("[TABLE START]:=") + 1)
                ReDim strFields(0) As String
                strFileRead = Trim(myFile.readline)
                While strFileRead <> "[TABLE END]"
                    strFields(UBound(strFields)) = strFileRead
                    strFileRead = Trim(myFile.readline)
                    ReDim Preserve strFields(UBound(strFields) + 1) As String
                Wend
                
                If Len(Trim(strTableName)) > 0 Then
                    Call createControls
                    lblTable(lblTable.UBound).ToolTipText = intDB + 1 & "." & strTableName
                    lblTable(lblTable.UBound).Text = "DB " & intDB + 1 & vbCrLf & strTableName
                    Dim lngY As Long
                    lngY = 0
                    lstFields(lblTable.UBound).Clear
                    While lngY < UBound(strFields)
                        lstFields(lblTable.UBound).AddItem UCase(strFields(lngY))
                        lngY = lngY + 1
                    Wend
                End If
               
            ElseIf Left(UCase(strFileRead), Len("[DB END]:=" & intDB)) = UCase("[DB END]:=" & intDB) Then
                GoTo lblNextDB
            ElseIf UCase(strFileRead) = UCase("[RELATIONSHIP START]") Then
                strFileRead = Trim(myFile.readline)
                While strFileRead <> "[RELATIONSHIP END]"
                    MSFlexGrid1.AddItem vbTab & strFileRead, 1
                    strFileRead = Trim(myFile.readline)
                Wend
            ElseIf Left(UCase(strFileRead), Len("[DISTINCT DATA]:=")) = UCase("[DISTINCT DATA]:=") Then
                intDistictData = Mid(UCase(strFileRead), Len("[DISTINCT DATA]:=") + 1, 1)
            ElseIf Left(UCase(strFileRead), Len("[PRIMARY TABLE]:=")) = UCase("[PRIMARY TABLE]:=") Then
                strPrimaryTable = Mid(UCase(strFileRead), Len("[PRIMARY TABLE]:=") + 1)
            ElseIf UCase(strFileRead) = UCase("[TABLE ALIAS START]") Then
                strFileRead = Trim(myFile.readline)
                ReDim strAlias(0) As String
                While UCase(strFileRead) <> "[TABLE ALIAS END]"
                    strAlias(UBound(strAlias)) = UCase(strFileRead)
                    strFileRead = Trim(myFile.readline)
                    ReDim Preserve strAlias(UBound(strAlias) + 1) As String
                Wend
                If UBound(strAlias) > 0 Then
                    ReDim Preserve strAlias(UBound(strAlias) - 1) As String
                End If
            ElseIf UCase(strFileRead) = UCase("[CALCULATED COLUMNS START]") Then
                strFileRead = Trim(myFile.readline)
                ReDim strCalcField(0) As String
                While UCase(strFileRead) <> "[CALCULATED COLUMNS END]"
                    If Left(UCase(strFileRead), Len("[CALCULATED COLUMN]:=")) = UCase("[CALCULATED COLUMN]:=") Then
                        ReDim Preserve strCalcField(UBound(strCalcField) + 1)
                        strCalcField(UBound(strCalcField)) = Mid(UCase(strFileRead), Len("[CALCULATED COLUMN]:=") + 1)
                        strFileRead = Trim(myFile.readline)
                        If UCase(strFileRead) = "[CALCULATION START]" Then
                            strFileRead = myFile.readline
                            ReDim Preserve strCalcField(UBound(strCalcField) + 1)
                            While UCase(strFileRead) <> "[CALCULATION END]"
                                If Len(strCalcField(UBound(strCalcField))) > 0 Then
                                    strCalcField(UBound(strCalcField)) = strCalcField(UBound(strCalcField)) & vbCrLf & strFileRead
                                Else
                                    strCalcField(UBound(strCalcField)) = strFileRead
                                End If
                                strFileRead = myFile.readline
                            Wend
                        End If
                    End If
                    strFileRead = Trim(myFile.readline)
                Wend
            ElseIf UCase(strFileRead) = UCase("[DELETED COLS START]") Then
                strFileRead = Trim(myFile.readline)
                ReDim strDeleteCols(0) As String
                While UCase(strFileRead) <> "[DELETED COLS END]"
                    strDeleteCols(UBound(strDeleteCols)) = UCase(strFileRead)
                    strFileRead = Trim(myFile.readline)
                    ReDim Preserve strDeleteCols(UBound(strDeleteCols) + 1) As String
                Wend
                If UBound(strDeleteCols) > 0 Then
                    ReDim Preserve strDeleteCols(UBound(strDeleteCols) - 1) As String
                End If
            End If
            
            strFileRead = Trim(myFile.readline)
            If myFile.AtEndOfStream = True Then
                GoTo lblCloseFile
            End If
        Wend
lblNextDB:        intDB = intDB + 1
    Wend

lblCloseFile:     myFile.Close

Call drawLines
If boolFromPrevious = False Then
    Call cmdNext_Click
End If
Exit Sub
End If

If boolJoin = False Then
    Call clearAllControls
    Dim intX As Long, intY As Long
    intX = 0
    intY = 0
    While intX <= frmTables.lblTable.UBound
        Call createControls
        lblTable(lblTable.UBound).ToolTipText = frmTables.lblTable(intX).ToolTipText
        lblTable(lblTable.UBound).Text = frmTables.lblTable(intX).Text
        intY = 0
        While intY < frmTables.lstFields(intX).ListCount
            If frmTables.lstFields(intX).Selected(intY) = True Then
                lstFields(lstFields.UBound).AddItem frmTables.lstFields(intX).List(intY)
            End If
            intY = intY + 1
        Wend
        intX = intX + 1
    Wend
    
    'clear flexi
    MSFlexGrid1.Rows = 1
    MSFlexGrid1.Rows = 2
Else
    Call drawLines
End If

Exit Sub


err:
MsgBox "There is an unexpected error in this report. Please try again", vbExclamation
Call cmdPrevious_Click
Exit Sub
End Sub

Sub createControls()
If lblTable(0).Visible = True And lblTable(0).Text = "" Then
    Exit Sub
End If

Load lblTable(lblTable.UBound + 1)
lblTable(lblTable.UBound).Left = lblTable(lblTable.UBound - 1).Left + lblTable(lblTable.UBound - 1).Width + 500
lblTable(lblTable.UBound).Top = lblTable(lblTable.UBound - 1).Top

Load lstFields(lstFields.UBound + 1)
lstFields(lstFields.UBound).Left = lblTable(lblTable.UBound).Left
lstFields(lstFields.UBound).Top = lblTable(lblTable.UBound).Top + lblTable(lblTable.UBound).Height

If lblTable.UBound Mod 5 = 0 Then
    lblTable(lblTable.UBound).Left = lblTable(lblTable.lBound).Left
    lblTable(lblTable.UBound).Top = lstFields(lstFields.UBound - 1).Top + lstFields(lstFields.UBound - 1).Height + 500
    
    lstFields(lstFields.UBound).Left = lstFields(lstFields.lBound).Left
    lstFields(lstFields.UBound).Top = lblTable(lblTable.UBound).Top + lblTable(lstFields.UBound).Height
End If

lblTable(lblTable.UBound).Visible = True
lstFields(lstFields.UBound).Visible = True
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

Sub clearAllControls()
Dim intX As Integer
intX = 0

While intX <= lblTable.UBound
    If intX <> 0 Then
        Unload lblTable(intX)
        Unload lstFields(intX)
    Else
        lblTable(intX).Text = ""
        lblTable(intX).ToolTipText = ""
        lstFields(intX).Clear
    End If
    intX = intX + 1
Wend

intX = 1
While intX <= l.UBound
    Unload l(intX)
    intX = intX + 1
Wend

lblTable(0).Top = 840
lblTable(0).Left = 360
lstFields(0).Top = 1365
lstFields(0).Left = 360

MSFlexGrid1.Rows = 1
MSFlexGrid1.Rows = 2
End Sub

Private Sub lblTable_MouseDown(Index As Integer, Button As Integer, Shift As Integer, X As Single, Y As Single)
If Button = vbLeftButton Then
    ReleaseCapture
    lstFields(Index).Visible = False
    Call SendMessage(lblTable(Index).hwnd, WM_NCLBUTTONDOWN, HTCAPTION, 0&)
    
    lstFields(Index).Top = lblTable(Index).Top + lblTable(Index).Height
    lstFields(Index).Left = lblTable(Index).Left
    lstFields(Index).Visible = True
    
    Call drawLines
End If
End Sub

Sub removeLine(strLstFrom As String, strLstTo As String)
On Error Resume Next
Dim intX As Long
intX = 1
While intX <= l.UBound
   If l(intX).Tag = strLstFrom Or l(intX).Tag = strLstTo Or l(intX).Tag = strLstFrom & "." & strLstTo Then
        Unload l(intX)
    End If
    intX = intX + 1
Wend
End Sub

Private Sub lstFields_DragDrop(Index As Integer, Source As Control, X As Single, Y As Single)
toList = Index
lstFields(Index).Drag vbEndDrag
If toList <> fromList Then
    Me.Enabled = False
    frmJoin.Show , Me
End If
End Sub

Private Sub lstFields_MouseUp(Index As Integer, Button As Integer, Shift As Integer, X As Single, Y As Single)
If Button = 2 Then
    fromList = Index
    lstFields(Index).Drag vbBeginDrag
End If
End Sub

Sub drawLines()
Call removeAllLines
Dim intX As Integer, intB As Integer, intA As Integer
intX = 1
With MSFlexGrid1
    While intX + 1 < .Rows
        intA = 0: intB = 0
        .Row = intX
        .Col = 5
        intA = CInt(.Text)
        .Col = 6
        intB = CInt(.Text)
        
        If intA + intB > 0 Then
            Call DrawNow(intA, intB)
        End If
        intX = intX + 1
    Wend
End With
End Sub

Sub DrawNow(lstFrom As Integer, lstTo As Integer)
Load l(l.UBound + 1)
l(l.UBound).Tag = lstFrom
l(l.UBound).Visible = True
l(l.UBound).Y1 = lstFields(lstFrom).Top + 50
l(l.UBound).Y2 = lstFields(lstFrom).Top + 50
l(l.UBound).X1 = lstFields(lstFrom).Left - 35
l(l.UBound).X2 = lstFields(lstFrom).Left - 250

Load l(l.UBound + 1)
l(l.UBound).Tag = lstTo
l(l.UBound).Visible = True
l(l.UBound).Y1 = lstFields(lstTo).Top + 50
l(l.UBound).Y2 = lstFields(lstTo).Top + 50
l(l.UBound).X1 = lstFields(lstTo).Left - 35
l(l.UBound).X2 = lstFields(lstTo).Left - 250


Load l(l.UBound + 1)
l(l.UBound).Tag = lstFrom & "." & lstTo
l(l.UBound).Visible = True
l(l.UBound).X1 = l(l.UBound - 2).X2
l(l.UBound).Y1 = l(l.UBound - 2).Y1
l(l.UBound).X2 = l(l.UBound - 1).X2
l(l.UBound).Y2 = l(l.UBound - 1).Y2
End Sub

Sub removeAllLines()
On Error Resume Next
Dim intX As Long
intX = 1
While intX <= l.UBound
    Unload l(intX)
    intX = intX + 1
Wend
End Sub

Private Sub mnuDel_Click()
MSFlexGrid1.Col = 1
If Len(MSFlexGrid1.Text) > 0 Then
    boolConfirm = MsgBox("Are you sure you want to delete this relationship ?", vbYesNoCancel + vbDefaultButton3 + vbQuestion)
    If boolConfirm = vbYes Then
        MSFlexGrid1.RemoveItem MSFlexGrid1.Row
        Call drawLines
    End If
End If
End Sub

Private Sub MSFlexGrid1_MouseUp(Button As Integer, Shift As Integer, X As Single, Y As Single)
If Button = 2 Then
    PopupMenu mnuOption
End If
End Sub

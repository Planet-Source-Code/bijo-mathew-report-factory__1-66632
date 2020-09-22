VERSION 5.00
Begin VB.Form frmStatus 
   BorderStyle     =   1  'Fixed Single
   Caption         =   " "
   ClientHeight    =   2880
   ClientLeft      =   45
   ClientTop       =   435
   ClientWidth     =   3585
   ControlBox      =   0   'False
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   MaxButton       =   0   'False
   MinButton       =   0   'False
   Picture         =   "frmStatus.frx":0000
   ScaleHeight     =   2880
   ScaleWidth      =   3585
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  'CenterScreen
   Begin VB.PictureBox Picture1 
      DrawMode        =   6  'Mask Pen Not
      Height          =   375
      Left            =   120
      Picture         =   "frmStatus.frx":9187
      ScaleHeight     =   315
      ScaleWidth      =   3315
      TabIndex        =   2
      Top             =   1800
      Width           =   3375
   End
   Begin Reports_Factory.ucButtons_H cmdOK 
      Height          =   375
      Left            =   1080
      TabIndex        =   3
      Top             =   2400
      Width           =   1455
      _ExtentX        =   2566
      _ExtentY        =   661
      Caption         =   "  &OK  "
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
      Image           =   "frmStatus.frx":1230E
      cBack           =   -2147483633
      mPointer        =   99
      mIcon           =   "frmStatus.frx":126A8
   End
   Begin VB.Image imgSuccess 
      Height          =   480
      Left            =   1440
      Picture         =   "frmStatus.frx":129C2
      Top             =   120
      Width           =   480
   End
   Begin VB.Label lblWarn 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "WARNING"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   178
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   240
      Left            =   1440
      TabIndex        =   0
      Top             =   360
      Visible         =   0   'False
      Width           =   1095
   End
   Begin VB.Image imgWarn 
      Height          =   480
      Left            =   840
      Picture         =   "frmStatus.frx":138C6
      Top             =   120
      Visible         =   0   'False
      Width           =   480
   End
   Begin VB.Label lblStatus 
      BackStyle       =   0  'Transparent
      Caption         =   "Transforming Data..."
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   178
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   915
      Left            =   120
      TabIndex        =   1
      Top             =   720
      Width           =   3315
      WordWrap        =   -1  'True
   End
End
Attribute VB_Name = "frmStatus"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Dim intCon As Integer
Dim strQuery As String
Dim strTableName As String

Sub PerCnt(iNewValue As Integer)
If iNewValue > 100 Or iNewValue < 0 Then
    Exit Sub
Else
    Picture1.Cls
    Picture1.FontSize = 10
    Picture1.ScaleMode = 0
    Picture1.ScaleWidth = 100
    Picture1.ScaleHeight = 10
    Picture1.CurrentY = 2
    Picture1.CurrentX = Picture1.ScaleWidth / 2 - (Picture1.ScaleWidth / 15)
    Picture1.Print Str(iNewValue) & "%"
    Picture1.Line (0, 0)-(iNewValue, Picture1.ScaleHeight), Picture1.FillColor, BF
End If
End Sub

Function checkRelationship() As Boolean
Picture1.Visible = True
Me.Caption = "Checking relationships..."
checkRelationship = True
With frmConnect
frmConnect.MSFlexGrid1.Redraw = False
Dim intX As Integer, intY As Long, intZ As Long

intX = 1
While intX <= frmConnect.lblTable.UBound
    frmConnect.lstCheck.Clear
    frmConnect.lstCheck.AddItem intX
    
    intY = 0
    While intY <= frmConnect.lstCheck.ListCount
        intZ = 1
        While intZ + 1 < frmConnect.MSFlexGrid1.Rows
            frmConnect.MSFlexGrid1.Row = intZ
            
            frmConnect.MSFlexGrid1.Col = 5
            If CStr(frmConnect.MSFlexGrid1.Text) = frmConnect.lstCheck.List(intY) Then
                frmConnect.MSFlexGrid1.Col = 6
                If chkListMatch(frmConnect.lstCheck, CStr(frmConnect.MSFlexGrid1.Text)) = False Then
                    frmConnect.lstCheck.AddItem frmConnect.MSFlexGrid1.Text
                End If
            End If
            
            frmConnect.MSFlexGrid1.Col = 6
            If CStr(frmConnect.MSFlexGrid1.Text) = frmConnect.lstCheck.List(intY) Then
                frmConnect.MSFlexGrid1.Col = 5
                If chkListMatch(frmConnect.lstCheck, CStr(frmConnect.MSFlexGrid1.Text)) = False Then
                    frmConnect.lstCheck.AddItem frmConnect.MSFlexGrid1.Text
                End If
            End If
            intZ = intZ + 1
        Wend
        intY = intY + 1
    Wend
    intX = intX + 1
    If chkListMatch(frmConnect.lstCheck, CStr(0)) = False Then
        Picture1.Visible = False
        lblStatus.Caption = "There are tables without relationship. " _
        & vbCrLf & "For eg.: " & Replace(frmConnect.lblTable(intX - 1).Text, vbCrLf, ".") _
        & vbCrLf & "Please correct this to continue."
        frmConnect.MSFlexGrid1.Redraw = True
        checkRelationship = False
        Me.Caption = " "
        Exit Function
    End If
Wend
Call PerCnt(CInt(25))

frmConnect.MSFlexGrid1.Redraw = True
End With
End Function

Private Sub cmdOK_Click()
If lblStatus.Caption = "Report created successfully. Click OK to continue" Then
    frmReport.Show
    frmConnect.Enabled = True
    frmConnect.Hide
    Unload frmPrimaryTable
    Unload frmStatus
Else
    boolFromRun = False
    frmConnect.Enabled = True
    frmConnect.SetFocus
    Unload frmPrimaryTable
    Unload frmStatus
End If
Unload Me
End Sub

Private Sub Form_Activate()
intDistictData = 0
cmdOK.Enabled = False
Call fWait(0.1)
Call PerCnt(CInt(0))
If checkRelationship = True Then
    Call fWait(0.1)
    If createDB = True Then
        Call fWait(0.1)
        If addData = True Then
            If frmPrimaryTable.chkShowDistinct.Value = 1 Then
                Call fWait(0.1)
                intDistictData = 1
                Call setDistinctData
            End If
            lblStatus.Caption = "Report created successfully. Click OK to continue"
            Me.Caption = " "
            Call PerCnt(100)
            Picture1.Visible = False
        End If
    End If
End If
cmdOK.Enabled = True
End Sub

Private Sub Form_Unload(Cancel As Integer)
frmConnect.Enabled = True
End Sub

Function createDB() As Boolean
On Error GoTo err:
Me.Caption = "Creating data mapping..."
createDB = True
Dim rsTables As New ADODB.Recordset
Picture1.Visible = True

Set rsTables = dbLocal.OpenSchema(adSchemaTables)
While rsTables.EOF = False
    If UCase(Trim(rsTables!TABLE_TYPE)) = "TABLE" Then
        dbLocal.Execute "drop table " & rsTables!TABLE_NAME
    End If
    rsTables.MoveNext
Wend
dbLocal.Execute "Create Table [data]"

Dim intX As Integer, intZ As Long, intY As Long, intA As Long, intB As Long
Dim strDB() As String
Dim strFields As String

intX = 0
With frmConnect
While intX <= .lblTable.UBound
    'checking tables & fields
    strFields = ""
    strDB = Split(.lblTable(intX).ToolTipText, ".", 2)
    intZ = 0
    While intZ <= .lstFields(intX).ListCount
        strFields = strFields & .lstFields(intX).List(intZ) & ","
        intZ = intZ + 1
    Wend
    strFields = "Select " & Mid(strFields, 1, Len(strFields) - 2) & " from " & strDB(1)
    'creating fields
    Set rsTables = Nothing
    strConString(CInt(strDB(0)) - 1).Open
    rsTables.Open strFields, strConString(CInt(strDB(0)) - 1), adOpenDynamic, adLockOptimistic
    intY = 0
    While intY < rsTables.Fields.Count
        dbLocal.Execute "alter table [data] add column [" & strDB(0) & "__" & strDB(1) & "__" & rsTables.Fields.Item(intY).Name & "]" _
        & dataType(rsTables.Fields.Item(intY).Type)
        intY = intY + 1
    Wend
     strConString(CInt(strDB(0)) - 1).Close
    intX = intX + 1
Wend

End With

Call PerCnt(CInt(50))
Exit Function

err:
createDB = False
lblStatus.Caption = "Unable to access local database." & vbCrLf & "Try again later"
Picture1.Visible = False
Me.Caption = " "
End Function

Function addData() As Boolean
On Error GoTo err
Dim strLinkFieldTemp As String
strLinkFieldTemp = ""
addData = True
Dim lngX As Long
Dim strDB() As String
With frmConnect
.MSFlexGrid1.Redraw = False
Me.Caption = "Transforming data..."
'add data
.lstCheck.Clear
.lstExecuted.Clear
lngX = 0
While lngX <= .lblTable.UBound
    If UCase(.lblTable(lngX).ToolTipText) = UCase(strPrimaryTable) Then
        .lstCheck.AddItem lngX, 0
        Call setBuildQuery(.lblTable(lngX).ToolTipText, lngX)
        strConString(intCon).Open
        Dim rsNewRec As New ADODB.Recordset
        rsNewRec.Open "Select * from [Data]", dbLocal, adOpenDynamic, adLockOptimistic
        Dim rsOldRec As New ADODB.Recordset
        rsOldRec.Open strQuery, strConString(intCon)
        'Add primary table data
        Call addCorrData(rsOldRec, rsNewRec, strTableName, intCon + 1)
        .lstExecuted.AddItem lngX
        GoTo lblE
    End If
    lngX = lngX + 1
Wend

'Add immeadiate subs for primary table
lblE: Dim lngA As Long
.lstCheck.Clear
lngA = 1
While lngA + 1 < .MSFlexGrid1.Rows
    If CLng(.MSFlexGrid1.TextMatrix(lngA, 5)) = lngX Then
        If chkListMatch(.lstCheck, .MSFlexGrid1.TextMatrix(lngA, 6)) = False Then
            .lstCheck.AddItem .MSFlexGrid1.TextMatrix(lngA, 6), .lstCheck.ListCount
        End If
    End If
    
    If CLng(.MSFlexGrid1.TextMatrix(lngA, 6)) = lngX Then
        If chkListMatch(.lstCheck, .MSFlexGrid1.TextMatrix(lngA, 5)) = False Then
            .lstCheck.AddItem .MSFlexGrid1.TextMatrix(lngA, 5), .lstCheck.ListCount
        End If
    End If
    lngA = lngA + 1
Wend

'add other subs
lngA = 0
Dim lngB As Long
lngB = 1
While lngA < .lstCheck.ListCount
    lngB = 1
    While lngB + 1 < .MSFlexGrid1.Rows
        If CLng(.MSFlexGrid1.TextMatrix(lngB, 5)) = CLng(.lstCheck.List(lngA)) Then
            If chkListMatch(.lstCheck, .MSFlexGrid1.TextMatrix(lngB, 6)) = False And CLng(.MSFlexGrid1.TextMatrix(lngB, 6)) <> CLng(lngX) Then
                .lstCheck.AddItem .MSFlexGrid1.TextMatrix(lngB, 6), lngA + 1
            End If
        End If

        If CLng(.MSFlexGrid1.TextMatrix(lngB, 6)) = CLng(.lstCheck.List(lngA)) Then
            If chkListMatch(.lstCheck, .MSFlexGrid1.TextMatrix(lngB, 5)) = False And CLng(.MSFlexGrid1.TextMatrix(lngB, 5)) <> CLng(lngX) Then
                .lstCheck.AddItem .MSFlexGrid1.TextMatrix(lngB, 5), lngA + 1
            End If
        End If
        lngB = lngB + 1
    Wend
    lngA = lngA + 1
Wend
Call PerCnt(CInt(75))
'-----------------add sub data now----------------------------------

lngA = 0: lngB = 0
Dim strLinkFieldNew As String
Dim strLinkFieldOld As String
Dim rsNewQry As New ADODB.Recordset
Dim rsOldQry As New ADODB.Recordset
Set rsNewQry = Nothing
Dim intPercent As Integer
rsNewQry.Open "Select * from data", dbLocal, adOpenDynamic, adLockOptimistic

While lngA < .lstCheck.ListCount
    intPercent = CInt((((lngA + 1) / .lstCheck.ListCount) * 100))
    intPercent = Round(CInt((intPercent / 100) * 25), 0)
    Call PerCnt(75 + intPercent)
    
    Call setBuildQuery(.lblTable(CInt(.lstCheck.List(lngA))).ToolTipText, CInt(.lstCheck.List(lngA)))
    Set rsOldQry = Nothing
    strConString(intCon).Open
    rsOldQry.Open strQuery, strConString(intCon)

    'set link fields
    Dim lngRow As Long
    lngRow = 1
    strLinkFieldNew = "": strLinkFieldOld = ""
    While lngRow + 1 < .MSFlexGrid1.Rows
        .MSFlexGrid1.Row = lngRow
        If UCase(.MSFlexGrid1.TextMatrix(lngRow, 7)) <> "Y" Then
            If CInt(.MSFlexGrid1.TextMatrix(lngRow, 5)) = CInt(.lstCheck.List(lngA)) And chkListMatch(.lstExecuted, .MSFlexGrid1.TextMatrix(lngRow, 6)) = True Then
                strLinkFieldOld = .MSFlexGrid1.TextMatrix(lngRow, 2)
                strLinkFieldNew = .MSFlexGrid1.TextMatrix(lngRow, 3)
                strLinkFieldNew = Replace(strLinkFieldNew, ".", "__") & "__" & .MSFlexGrid1.TextMatrix(lngRow, 4)
                .MSFlexGrid1.TextMatrix(lngRow, 7) = "Y"
                .lstExecuted.AddItem .MSFlexGrid1.TextMatrix(lngRow, 5)
            End If
            
            If CInt(.MSFlexGrid1.TextMatrix(lngRow, 6)) = CInt(.lstCheck.List(lngA)) And chkListMatch(.lstExecuted, .MSFlexGrid1.TextMatrix(lngRow, 5)) = True Then
                strLinkFieldOld = .MSFlexGrid1.TextMatrix(lngRow, 4)
                strLinkFieldNew = .MSFlexGrid1.TextMatrix(lngRow, 1)
                strLinkFieldNew = Replace(strLinkFieldNew, ".", "__") & "__" & .MSFlexGrid1.TextMatrix(lngRow, 2)
                .MSFlexGrid1.TextMatrix(lngRow, 7) = "Y"
                .lstExecuted.AddItem .MSFlexGrid1.TextMatrix(lngRow, 6)
            End If
        End If
        
        If Len(strLinkFieldNew) > 0 And Len(strLinkFieldOld) > 0 Then
            On Error Resume Next
            rsNewQry.MoveFirst
            While rsNewQry.EOF = False
'                If rsOldQry.EOF = False Then
                    On Error Resume Next
                    rsOldQry.MoveFirst
'                End If
                While rsOldQry.EOF = False
                    If UCase(Trim(rsOldQry(strLinkFieldOld))) = UCase(Trim(rsNewQry(strLinkFieldNew))) Then
                        lngB = 0
                        While lngB < rsOldQry.Fields.Count
                            rsNewQry(intCon + 1 & "__" & strTableName & "__" & rsOldQry.Fields(lngB).Name) = rsOldQry(lngB)
                            rsNewQry.Update
                            lngB = lngB + 1
                        Wend
                    End If
                    rsOldQry.MoveNext
                Wend
                rsNewQry.MoveNext
            Wend
            
            GoTo lblIncrList
        End If
        lngRow = lngRow + 1
    Wend
lblIncrList: strConString(intCon).Close
    lngA = lngA + 1
Wend

Call setMSFNils
.MSFlexGrid1.Redraw = True
Me.Caption = " "
Call PerCnt(CInt(100))
Picture1.Visible = False

End With
Exit Function

err:
Call setMSFNils
addData = False
lblStatus.Caption = "Error in processing data"
Me.Caption = " "
Picture1.Visible = False
frmConnect.MSFlexGrid1.Redraw = True
End Function

Sub setBuildQuery(strToolTipText As String, lngCtrl As Long)
Dim strDB() As String, strFields As String
Dim intZ As Long
intCon = -1
strQuery = ""
strTableName = "'"

strDB = Split(strToolTipText, ".", 2)
intZ = 0
While intZ <= frmConnect.lstFields(lngCtrl).ListCount
    strFields = strFields & frmConnect.lstFields(lngCtrl).List(intZ) & ","
    intZ = intZ + 1
Wend
    intCon = (CInt(strDB(0)) - 1)
    strQuery = "Select " & Mid(strFields, 1, Len(strFields) - 2) & " from " & strDB(1)
    strTableName = strDB(1)
End Sub

Sub addCorrData(rsOldTables As ADODB.Recordset, rsNewTables As ADODB.Recordset, strTableNameOld As String, intDB As Integer)
Dim lngA As Long, lngB As Long
lngA = 0: lngB = 0
If rsOldTables.EOF = False Then
    rsOldTables.MoveFirst
    While rsOldTables.EOF = False
        lngA = 0
        rsNewTables.AddNew
        While lngA < rsOldTables.Fields.Count
            lngB = 0
            While lngB < rsNewTables.Fields.Count
                If UCase(rsNewTables.Fields(lngB).Name) = UCase(intDB & "__" & strTableNameOld & "__" & rsOldTables.Fields.Item(lngA).Name) Then
                    rsNewTables(lngB) = rsOldTables(lngA)
                End If
                lngB = lngB + 1
            Wend
            lngA = lngA + 1
        Wend
        rsNewTables.Update
        rsOldTables.MoveNext
    Wend
End If

strConString(intCon).Close
End Sub

Private Sub lblStatus_Change()
If lblStatus.Caption = "Report created successfully. Click OK to continue" Or lblStatus.Caption = "" Then
    lblWarn.Visible = False
    imgWarn.Visible = False
    imgSuccess.Visible = True
Else
    imgSuccess.Visible = False
    lblWarn.Visible = True
    imgWarn.Visible = True
End If
End Sub

Sub setMSFNils()
With frmConnect.MSFlexGrid1
.Redraw = False
Dim lngX As Long
lngX = 1
While lngX + 1 < .Rows
    .TextMatrix(lngX, 7) = ""
    lngX = lngX + 1
Wend
.Redraw = True
End With
End Sub

Sub setDistinctData()
Picture1.Visible = True
Call PerCnt(0)
dbLocal.Execute "Select Distinct * into Data1 from Data"
Call PerCnt(25)
Call fWait(0.2)
dbLocal.Execute "Drop Table Data"
Call fWait(0.2)
Call PerCnt(50)
dbLocal.Execute "Select Distinct * into Data from Data1"
Call fWait(0.2)
Call PerCnt(75)
dbLocal.Execute "Drop Table Data1"
Call PerCnt(100)
End Sub


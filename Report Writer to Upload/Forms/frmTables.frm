VERSION 5.00
Begin VB.Form frmTables 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Step 2  >>  Select Tables"
   ClientHeight    =   7185
   ClientLeft      =   45
   ClientTop       =   435
   ClientWidth     =   14040
   HelpContextID   =   1010
   Icon            =   "frmTables.frx":0000
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   MaxButton       =   0   'False
   MinButton       =   0   'False
   Picture         =   "frmTables.frx":038A
   ScaleHeight     =   7185
   ScaleWidth      =   14040
   StartUpPosition =   2  'CenterScreen
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
      Left            =   4680
      Locked          =   -1  'True
      MultiLine       =   -1  'True
      ScrollBars      =   2  'Vertical
      TabIndex        =   4
      Top             =   795
      Width           =   1810
   End
   Begin VB.ListBox lstFields 
      Appearance      =   0  'Flat
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   178
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   1380
      Index           =   0
      ItemData        =   "frmTables.frx":23044
      Left            =   4680
      List            =   "frmTables.frx":23046
      Sorted          =   -1  'True
      Style           =   1  'Checkbox
      TabIndex        =   5
      Top             =   1320
      Width           =   1815
   End
   Begin VB.ComboBox cboDB 
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   178
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   315
      ItemData        =   "frmTables.frx":23048
      Left            =   1200
      List            =   "frmTables.frx":23055
      Style           =   2  'Dropdown List
      TabIndex        =   1
      Top             =   0
      Width           =   12855
   End
   Begin VB.ListBox lstTables 
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   178
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   5910
      ItemData        =   "frmTables.frx":2307B
      Left            =   0
      List            =   "frmTables.frx":2307D
      Sorted          =   -1  'True
      Style           =   1  'Checkbox
      TabIndex        =   3
      Top             =   720
      Width           =   4335
   End
   Begin Reports_Factory.ucButtons_H cmdExit 
      Height          =   375
      Left            =   12240
      TabIndex        =   9
      Top             =   6720
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
      Image           =   "frmTables.frx":2307F
      cBack           =   -2147483633
      mPointer        =   99
      mIcon           =   "frmTables.frx":23419
   End
   Begin Reports_Factory.ucButtons_H cmdNext 
      Height          =   375
      Left            =   7320
      TabIndex        =   7
      Top             =   6720
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
      Image           =   "frmTables.frx":23733
      cBack           =   -2147483633
      mPointer        =   99
      mIcon           =   "frmTables.frx":23ACD
   End
   Begin Reports_Factory.ucButtons_H cmdPrevious 
      Height          =   375
      Left            =   5400
      TabIndex        =   6
      Top             =   6720
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
      Image           =   "frmTables.frx":23DE7
      cBack           =   -2147483633
      mPointer        =   99
      mIcon           =   "frmTables.frx":24181
   End
   Begin Reports_Factory.ucButtons_H cmdHelp 
      Height          =   375
      Left            =   10320
      TabIndex        =   8
      Top             =   6720
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
      Image           =   "frmTables.frx":2449B
      cBack           =   -2147483633
      mPointer        =   99
      mIcon           =   "frmTables.frx":2751D
   End
   Begin VB.Label Label2 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Database:"
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
      Left            =   105
      TabIndex        =   0
      Top             =   75
      UseMnemonic     =   0   'False
      Width           =   885
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Tables in Database"
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
      Left            =   105
      TabIndex        =   2
      Top             =   480
      UseMnemonic     =   0   'False
      Width           =   1665
   End
   Begin VB.Menu mnuOptions 
      Caption         =   "&Options"
      Visible         =   0   'False
      Begin VB.Menu mnuDel 
         Caption         =   "&Delete Table"
      End
   End
End
Attribute VB_Name = "frmTables"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Dim boolSupress As Boolean
Dim dbCon As New ADODB.Connection
Dim intTempX As Integer

Private Sub cboDB_Change()
On Error GoTo err:

Dim strDBText() As String
Dim strCon() As String
Dim strUID As String
Dim strPWD As String
Set dbCon = Nothing
cboDB.ToolTipText = cboDB.Text
lstTables.Clear

strUID = "": strPWD = ""
strUID = frmDB.lstDBUser.List(((cboDB.ListIndex) * 3) + 1)
strPWD = frmDB.lstDBUser.List(((cboDB.ListIndex) * 3) + 2)
strDBText = Split(cboDB.Text, " Database -  > ")

If UCase(strDBText(0)) = "MS SQL SERVER" Then
    strCon = Split(strDBText(1), " on ")
    dbCon.Open ("Provider=SQLOLEDB;data Source=" & strCon(1) & ";Initial Catalog=" & strCon(0) & ";User Id=" & strUID & ";Password=" & strPWD & ";")
ElseIf UCase(strDBText(0)) = "MS ACCESS" Then
    dbCon.Open ("Provider=Microsoft.Jet.OLEDB.4.0;Data Source=" & strDBText(1) & ";Jet OLEDB:Database Password=" & strPWD)
ElseIf UCase(strDBText(0)) = "ORACLE" Then
    dbCon.Open "DRIVER={Microsoft ODBC For Oracle};UID=" & strUID & ";PWD=" & strPWD & ";SERVER=" & strDBText(1)
ElseIf UCase(strDBText(0)) = "MYSQL" Then
    strCon = Split(strDBText(1), " on ")
    dbCon.Open "DRIVER={MySQL ODBC 3.51 Driver};SERVER=" & strCon(1) & ";DATABASE=" & strCon(0) & ";UID=" & strUID & ";PWD=" & strPWD
ElseIf UCase(strDBText(0)) = "POSTGRESQL" Then
    strCon = Split(strDBText(1), " on ")
    dbCon.Open "DRIVER={PostgreSQL Unicode};SERVER=" & strCon(1) & ";DATABASE=" & strCon(0) & ";UID=" & strUID & ";PWD=" & strPWD
End If

Dim rsTables As New ADODB.Recordset
If UCase(strDBText(0)) = "MS SQL SERVER" Or UCase(strDBText(0)) = "MS ACCESS" Or UCase(strDBText(0)) = "MYSQL" Or UCase(strDBText(0)) = "POSTGRESQL" Then
    Set rsTables = dbCon.OpenSchema(adSchemaTables)
    If rsTables.EOF = False Then
    rsTables.MoveFirst
    lstTables.Visible = False
    While rsTables.EOF = False
        If UCase(Trim(rsTables!TABLE_TYPE)) = "TABLE" Then
            lstTables.AddItem rsTables!TABLE_NAME
        End If
        rsTables.MoveNext
    Wend
    lstTables.Visible = True
    End If
ElseIf UCase(strDBText(0)) = "ORACLE" Then
    'oracle takes other table also if schema is selected
    rsTables.Open "Select TNAME from tab", dbCon
    If rsTables.EOF = False Then
    rsTables.MoveFirst
    lstTables.Visible = False
    While rsTables.EOF = False
        lstTables.AddItem UCase(rsTables!TNAME)
        rsTables.MoveNext
    Wend
    lstTables.Visible = True
End If
End If

Call checkExitingItems
Exit Sub

err:
MsgBox err.Description, vbExclamation
Exit Sub
End Sub

Private Sub cboDB_Click()
Call cboDB_Change
End Sub

Private Sub cmdExit_Click()
Unload Me
End Sub

Private Sub cmdHelp_Click()
Call ShowAppHelp(1010)
End Sub

Private Sub cmdNext_Click()
Dim lngX As Long, lngY As Long
lngX = 0: lngY = 0
If lblTable.UBound = 0 And lblTable(0).Text = "" Then
    MsgBox "Please select atleast 1 table to continue", vbExclamation
    Exit Sub
Else
    While lngX <= lblTable.UBound
        lngY = 0
        While lngY < lstFields(lngX).ListCount
            If lstFields(lngX).Selected(lngY) = True Then
                GoTo lbl
            End If
            lngY = lngY + 1
        Wend
        'check last item
        If lstFields(lngX).Selected(lngY - 1) = True Then
            GoTo lbl
        End If
        MsgBox "There are tables without fields selected. Please remove or select fields to continue" _
        & vbCrLf & "For eg: " & "DB" & lblTable(lngX).ToolTipText, vbExclamation
        Exit Sub
        
lbl:     lngX = lngX + 1
    Wend
    boolJoin = False
    boolFromRun = False
    frmConnect.Show
    Me.Hide
End If
End Sub

Private Sub cmdPrevious_Click()
Me.Hide
frmDB.Show
End Sub

Private Sub Form_Activate()
Call Form_Load
End Sub

Private Sub Form_Load()
Dim intX As Integer
cboDB.Clear
intX = 0

With frmDB
    While intX < .lstDB.ListCount
        cboDB.AddItem .lstDB.List(intX) & " - " & .lstDB.List(intX + 1)
        intX = intX + 3
    Wend
End With
cboDB.Text = cboDB.List(0)

'set connection strings
Dim strUID As String, strPWD As String
Dim strDBText() As String
Dim strCon() As String
intX = 0
While intX <= 9
    strConString(intX) = ""
    intX = intX + 1
Wend

intX = 0
While intX < cboDB.ListCount
    strUID = "": strPWD = ""
    strUID = frmDB.lstDBUser.List(((intX) * 3) + 1)
    strPWD = frmDB.lstDBUser.List(((intX) * 3) + 2)
    strDBText = Split(cboDB.List(intX), " Database -  > ")
    
    If UCase(strDBText(0)) = "MS SQL SERVER" Then
        strCon = Split(strDBText(1), " on ")
        strConString(intX) = "Provider=SQLOLEDB;data Source=" & strCon(1) & ";Initial Catalog=" & strCon(0) & ";User Id=" & strUID & ";Password=" & strPWD & ";"
    ElseIf UCase(strDBText(0)) = "MS ACCESS" Then
        strConString(intX).ConnectionString = "Provider=Microsoft.Jet.OLEDB.4.0;Data Source=" & strDBText(1) & ";Jet OLEDB:Database Password=" & strPWD
    ElseIf UCase(strDBText(0)) = "ORACLE" Then
        strConString(intX) = "DRIVER={Microsoft ODBC For Oracle};UID=" & strUID & ";PWD=" & strPWD & ";SERVER=" & strDBText(1)
    ElseIf UCase(strDBText(0)) = "MYSQL" Then
        strCon = Split(strDBText(1), " on ")
        strConString(intX) = "DRIVER={MySQL ODBC 3.51 Driver};SERVER=" & strCon(1) & ";DATABASE=" & strCon(0) & ";UID=" & strUID & ";PWD=" & strPWD
    ElseIf UCase(strDBText(0)) = "POSTGRESQL" Then
        strCon = Split(strDBText(1), " on ")
        strConString(intX) = "DRIVER={PostgreSQL Unicode};SERVER=" & strCon(1) & ";DATABASE=" & strCon(0) & ";UID=" & strUID & ";PWD=" & strPWD
    End If
    
    intX = intX + 1
Wend
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

Private Sub lblTable_MouseUp(Index As Integer, Button As Integer, Shift As Integer, X As Single, Y As Single)
If Button = 1 Then
    intTempX = Index
    PopupMenu mnuOptions()
End If
End Sub

Private Sub lstTables_Click()
Dim strTemp As String
Dim rsRec As New ADODB.Recordset
Dim intX As Integer, tempX As Integer, intY As Integer, lngZ As Long

If lstTables.Selected(lstTables.ListIndex) = True Then
    'supress messages if auto selecting / check for existing
    If boolSupress = False Then
        intX = 0
        While intX <= lblTable.UBound
            If UCase(lblTable(intX).ToolTipText) = cboDB.ListIndex + 1 & "." & lstTables.Text Then
                MsgBox "This table already exist", vbExclamation
                Exit Sub
            End If
            intX = intX + 1
        Wend
    Else
        boolSupress = False
        Exit Sub
    End If
    
    'create and add text items
    Call createControls
    lblTable(lblTable.UBound).ToolTipText = cboDB.ListIndex + 1 & "." & lstTables.Text
    lblTable(lblTable.UBound).Text = "DB " & cboDB.ListIndex + 1 & vbCrLf
    lblTable(lblTable.UBound).Text = lblTable(lblTable.UBound).Text & lstTables.Text

    
    'move selected items to top
    boolSupress = True
    strTemp = ""
    strTemp = lstTables.Text
    
    'set fields
    rsRec.Open "select * from " & strTemp, dbCon, adOpenDynamic, adLockOptimistic
    lstFields(lstFields.UBound).Clear
    Dim intRec As Integer
    intRec = 0
    While intRec < rsRec.Fields.Count
        lstFields(lstFields.UBound).AddItem rsRec.Fields(intRec).Name
        intRec = intRec + 1
    Wend
    
    lstTables.RemoveItem lstTables.ListIndex
    lstTables.AddItem strTemp, 0
    lstTables.Selected(0) = True
ElseIf lstTables.Selected(lstTables.ListIndex) = False Then
    'remove control / clear also
    intX = 0
    While intX <= lblTable.UBound
        If UCase(lblTable(intX).ToolTipText) = cboDB.ListIndex + 1 & "." & lstTables.Text Then
            intX = intX + 1
            'trnasfer text to preceeding ones
            While intX <= lblTable.UBound
                lblTable(intX - 1).Text = lblTable(intX).Text
                lngZ = 0
                lstFields(intX - 1).Clear
                While lngZ < lstFields(intX).ListCount
                    lstFields(intX - 1).AddItem lstFields(intX).List(lngZ)
                    lngZ = lngZ + 1
                Wend
                lblTable(intX - 1).ToolTipText = lblTable(intX).ToolTipText
                intX = intX + 1
            Wend
            If lblTable.Count <> 1 Then
                Unload lblTable(lblTable.UBound)
                Unload lstFields(lstFields.UBound)
            Else
                lblTable(lblTable.UBound).Text = ""
                lstFields(lstFields.UBound).Clear
                lblTable(lblTable.UBound).ToolTipText = ""
            End If
            Exit Sub
        End If
        intX = intX + 1
    Wend
End If

End Sub

Sub createControls()
If lblTable.UBound >= 14 Then
    Exit Sub
ElseIf lblTable(0).Visible = True And lblTable(0).Text = "" Then
    Exit Sub
End If

Load lblTable(lblTable.UBound + 1)
lblTable(lblTable.UBound).Left = lblTable(lblTable.UBound - 1).Left + lblTable(lblTable.UBound - 1).Width + 55
lblTable(lblTable.UBound).Top = lblTable(lblTable.UBound - 1).Top

Load lstFields(lstFields.UBound + 1)
lstFields(lstFields.UBound).Left = lstFields(lstFields.UBound - 1).Left + lstFields(lstFields.UBound - 1).Width + 50
lstFields(lstFields.UBound).Top = lstFields(lstFields.UBound - 1).Top

If lblTable.UBound Mod 5 = 0 Then
    lblTable(lblTable.UBound).Left = lblTable(lblTable.lBound).Left
    lblTable(lblTable.UBound).Top = lstFields(lstFields.UBound - 1).Top + lstFields(lstFields.UBound - 1).Height + 50
    
    lstFields(lstFields.UBound).Left = lstFields(lstFields.lBound).Left
    lstFields(lstFields.UBound).Top = lblTable(lblTable.UBound).Top + lblTable(lstFields.UBound).Height
End If

lblTable(lblTable.UBound).Visible = True
lstFields(lstFields.UBound).Visible = True
End Sub

Sub checkExitingItems()
Dim lngX As Long, lngY As Long
lngX = 0: lngY = 0
While lngX <= lblTable.UBound
    lngY = 0
    While lngY <= lstTables.ListCount
        If UCase(lblTable(lngX).ToolTipText) = cboDB.ListIndex + 1 & "." & UCase(lstTables.List(lngY)) Then
            boolSupress = True
            lstTables.Selected(lngY) = True
        End If
        lngY = lngY + 1
    Wend
    lngX = lngX + 1
Wend
End Sub

Private Sub mnuDel_Click()
Dim lngX As Long
lngX = 0
While lngX <= lstTables.ListCount
    If UCase(cboDB.ListIndex + 1 & "." & lstTables.List(lngX)) = UCase(lblTable(intTempX).ToolTipText) Then
        boolSupress = True
        lstTables.ListIndex = lngX
        lstTables.Selected(lngX) = False
        Exit Sub
    End If
    lngX = lngX + 1
Wend
End Sub

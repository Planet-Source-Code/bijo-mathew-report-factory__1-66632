VERSION 5.00
Begin VB.Form frmDB 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Step 1 >  Select Databases"
   ClientHeight    =   3990
   ClientLeft      =   45
   ClientTop       =   435
   ClientWidth     =   9960
   BeginProperty Font 
      Name            =   "MS Sans Serif"
      Size            =   8.25
      Charset         =   178
      Weight          =   700
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   HelpContextID   =   1009
   Icon            =   "frmDB.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   Picture         =   "frmDB.frx":038A
   ScaleHeight     =   3990
   ScaleWidth      =   9960
   StartUpPosition =   2  'CenterScreen
   Begin VB.ListBox lstDB 
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   9
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   2760
      ItemData        =   "frmDB.frx":23044
      Left            =   5760
      List            =   "frmDB.frx":23046
      TabIndex        =   13
      Top             =   240
      Width           =   3975
   End
   Begin VB.ComboBox cboDBType 
      Height          =   315
      ItemData        =   "frmDB.frx":23048
      Left            =   1920
      List            =   "frmDB.frx":2305B
      Style           =   2  'Dropdown List
      TabIndex        =   1
      Top             =   360
      Width           =   1815
   End
   Begin VB.TextBox txtServer 
      Height          =   315
      Left            =   1920
      TabIndex        =   3
      Top             =   960
      Width           =   1815
   End
   Begin VB.TextBox txtUID 
      Height          =   315
      Left            =   1920
      TabIndex        =   7
      Top             =   2160
      Width           =   1815
   End
   Begin VB.TextBox txtPWD 
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   9
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   330
      IMEMode         =   3  'DISABLE
      Left            =   1920
      PasswordChar    =   "×"
      TabIndex        =   9
      Top             =   2760
      Width           =   1815
   End
   Begin VB.TextBox txtDB 
      Height          =   315
      Left            =   1920
      TabIndex        =   5
      Top             =   1560
      Width           =   1815
   End
   Begin VB.CheckBox chkSave 
      BackColor       =   &H00FFC0C0&
      Caption         =   "Save all my database connection details "
      Height          =   255
      Left            =   5760
      TabIndex        =   21
      Top             =   3000
      Width           =   210
   End
   Begin VB.ListBox lstDBUser 
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   9
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   2085
      ItemData        =   "frmDB.frx":23094
      Left            =   120
      List            =   "frmDB.frx":23096
      TabIndex        =   20
      Top             =   4200
      Visible         =   0   'False
      Width           =   2655
   End
   Begin Reports_Factory.ucButtons_H cmdAdd 
      Height          =   375
      Left            =   4080
      TabIndex        =   10
      Top             =   1200
      Width           =   1335
      _ExtentX        =   2355
      _ExtentY        =   661
      Caption         =   "&Add         "
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
      Image           =   "frmDB.frx":23098
      cBack           =   -2147483633
      mPointer        =   99
      mIcon           =   "frmDB.frx":23352
   End
   Begin Reports_Factory.ucButtons_H cmdRemove 
      Height          =   375
      Left            =   4080
      TabIndex        =   11
      Top             =   1920
      Width           =   1335
      _ExtentX        =   2355
      _ExtentY        =   661
      Caption         =   "&Remove  "
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
      Image           =   "frmDB.frx":2366C
      Enabled         =   0   'False
      cBack           =   -2147483633
      mPointer        =   99
      mIcon           =   "frmDB.frx":23926
   End
   Begin Reports_Factory.ucButtons_H cmdTestConnection 
      Height          =   375
      Left            =   240
      TabIndex        =   15
      Top             =   3480
      Width           =   1935
      _ExtentX        =   3413
      _ExtentY        =   661
      Caption         =   "&Test Connection"
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
      Image           =   "frmDB.frx":23C40
      cBack           =   -2147483633
      mPointer        =   99
      mIcon           =   "frmDB.frx":23FDA
   End
   Begin Reports_Factory.ucButtons_H cmdExit 
      Height          =   375
      Left            =   8040
      TabIndex        =   19
      Top             =   3480
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
      Image           =   "frmDB.frx":242F4
      cBack           =   -2147483633
      mPointer        =   99
      mIcon           =   "frmDB.frx":2468E
   End
   Begin Reports_Factory.ucButtons_H cmdNext 
      Height          =   375
      Left            =   4320
      TabIndex        =   17
      Top             =   3480
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
      Image           =   "frmDB.frx":249A8
      cBack           =   -2147483633
      mPointer        =   99
      mIcon           =   "frmDB.frx":24D42
   End
   Begin Reports_Factory.ucButtons_H cmdHelp 
      Height          =   375
      Left            =   6240
      TabIndex        =   18
      Top             =   3480
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
      Image           =   "frmDB.frx":2505C
      cBack           =   -2147483633
      mPointer        =   99
      mIcon           =   "frmDB.frx":280DE
   End
   Begin Reports_Factory.ucButtons_H cmdPrevious 
      Height          =   375
      Left            =   2520
      TabIndex        =   16
      Top             =   3480
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
      Image           =   "frmDB.frx":283F8
      cBack           =   -2147483633
      mPointer        =   99
      mIcon           =   "frmDB.frx":28792
   End
   Begin Reports_Factory.ucButtons_H cmdHint 
      Height          =   375
      Left            =   4080
      TabIndex        =   12
      ToolTipText     =   "Click here to find out how to create database connections"
      Top             =   2760
      Width           =   1335
      _ExtentX        =   2355
      _ExtentY        =   661
      Caption         =   "&Hints"
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
      Image           =   "frmDB.frx":28AAC
      cBack           =   -2147483633
      mPointer        =   99
      mIcon           =   "frmDB.frx":2BB2E
   End
   Begin VB.Label Label2 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Save all my database connection details "
      Height          =   195
      Left            =   6000
      TabIndex        =   14
      Top             =   3020
      Width           =   3510
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Type of Database:"
      Height          =   195
      Left            =   240
      TabIndex        =   0
      Top             =   480
      Width           =   1590
   End
   Begin VB.Label lblServer 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Server Name:"
      Height          =   195
      Left            =   240
      TabIndex        =   2
      Top             =   1080
      Width           =   1170
   End
   Begin VB.Label Label3 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "User Name:"
      Height          =   195
      Left            =   240
      TabIndex        =   6
      Top             =   2280
      Width           =   1005
   End
   Begin VB.Label Label4 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Password:"
      Height          =   195
      Left            =   240
      TabIndex        =   8
      Top             =   2880
      Width           =   885
   End
   Begin VB.Label lblDB 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Database:"
      Height          =   195
      Left            =   240
      TabIndex        =   4
      Top             =   1680
      Width           =   885
   End
End
Attribute VB_Name = "frmDB"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private Sub cboDBType_Change()
If UCase(cboDBType.Text) = "MS ACCESS" Then
    lblServer.Enabled = False
    txtServer.Text = ""
    txtServer.Enabled = False
    txtServer.BackColor = &H8000000F
    
    lblDB.Enabled = True
    txtDB.Text = ""
    txtDB.Enabled = True
    txtDB.BackColor = &H80000005
ElseIf UCase(cboDBType.Text) = "ORACLE" Then
    lblDB.Enabled = False
    txtDB.Text = ""
    txtDB.Enabled = False
    txtDB.BackColor = &H8000000F
    
    lblServer.Enabled = True
    txtServer.Text = ""
    txtServer.Enabled = True
    txtServer.BackColor = &H80000005
Else
    lblServer.Enabled = True
    txtServer.Text = ""
    txtServer.Enabled = True
    txtServer.BackColor = &H80000005
    
    lblDB.Enabled = True
    txtDB.Text = ""
    txtDB.Enabled = True
    txtDB.BackColor = &H80000005
End If
End Sub

Private Sub cboDBType_Click()
Call cboDBType_Change
End Sub

Private Sub cmdAdd_Click()
If lstDB.ListCount >= 27 Then
    MsgBox "You cannot add more than 9 databases", vbExclamation
    Exit Sub
End If
If checkExistingDB = False Then
    If testDBConnection(cboDBType.Text, txtServer.Text, txtDB.Text, txtUID.Text, txtPWD.Text) = True Then
        lstDB.AddItem cboDBType.Text & " Database"
        If UCase(cboDBType.Text) = "MS SQL SERVER" Then
            lstDB.AddItem " > " & txtDB.Text & " on " & txtServer.Text
        ElseIf UCase(cboDBType.Text) = "MS ACCESS" Then
            lstDB.AddItem " > " & txtDB.Text
        ElseIf UCase(cboDBType.Text) = "ORACLE" Then
            lstDB.AddItem " > " & txtServer.Text
        ElseIf UCase(cboDBType.Text) = "MYSQL" Then
            lstDB.AddItem " > " & txtDB.Text & " on " & txtServer.Text
        ElseIf UCase(cboDBType.Text) = "POSTGRESQL" Then
            lstDB.AddItem " > " & txtDB.Text & " on " & txtServer.Text
        End If
            lstDB.AddItem " "
            lstDBUser.AddItem " "
            lstDBUser.AddItem txtUID.Text
            lstDBUser.AddItem txtPWD.Text
            
        txtServer.Text = ""
        txtDB.Text = ""
        txtUID.Text = ""
        txtPWD.Text = ""
    Else
        MsgBox "Connection failed..." & vbCrLf & strError, vbExclamation
    End If
Else
    MsgBox "The database specified by you already exist", vbExclamation
End If
End Sub

Private Sub cmdExit_Click()
Unload Me
End Sub

Private Sub cmdHelp_Click()
Call ShowAppHelp(1009)
End Sub

Private Sub cmdHint_Click()
MsgBox "The databases that are currently supported (but increasing soon) are. Please note that you will require the specified drivers to connect to the databases listed:" & vbCrLf _
& "1. MS Access" & vbCrLf & "2. MS SQL Server" & vbCrLf & "3. Oracle" & vbCrLf _
& "4. MySQL" & vbCrLf & "5. PostgreSQL" & vbCrLf & vbCrLf & vbCrLf _
& "For MS Access: a. Ignore the Server Name b. The database name should be the full path of the database eg: 'C:\Windows\Test.mdb' c. The user name can be left blank or enter 'Admin' d. If your database has password enter it or leave it blank" & vbCrLf & vbCrLf _
& "For MS SQL Server: a. Enter the Server Name as machine name or IP eg: '\\TestMachine' or '127.0.0.1' b. The database name should be the name of the database eg: 'Northwind' c. The user name should be the user name to the database usually 'sa' d. If your database has password enter it or leave it blank", vbInformation

MsgBox "For Oracle: a. Ignore the Server Name because this will be taken care by the instance manager b. The database name should be the instance name as specified by the Net Configuration Wizard c. The user name should be the user name to the database default is 'system' d. If your database has password enter it or leave it blank for user 'system' the default is 'manager'" & vbCrLf & vbCrLf _
& "For MySQL: a. Enter the Server Name as machine name or IP eg: '\\TestMachine' or '127.0.0.1' b. The database name should be the name database eg: 'mysql' which will be a default DB in MySQL c. The user name should be the user name to the database usually 'root' d. If your database has password enter it or leave it blank" & vbCrLf & vbCrLf _
& "For PostgreSQL: a. Enter the Server Name as machine name or IP eg: '\\TestMachine' or '127.0.0.1' b. The database name should be the name database eg: 'postgres' which will be a default DB in PostgreSQL c. The user name should be the user name to the database d. If your database has password enter it or leave it blank", vbInformation

MsgBox "Note that the general assumption is that the databases run on their default ports For eg: " & vbCrLf _
& "1. MS Access does not have port based architecture" & vbCrLf _
& "2. MS SQL Server on 1433 but the port will be taken care by the SQL Server" & vbCrLf _
& "3. Oracle on 1521 but this will be taken care by the net assistant that created the instance" & vbCrLf _
& "4. MySQL on 3306" & vbCrLf _
& "5. PostgreSQL on 5432", vbInformation

End Sub

Private Sub cmdNext_Click()
Dim intX As Integer
Dim intY As Integer

If lstDB.ListCount >= 3 Then
        Call deleteExistingDB
    If chkSave.Value = 1 Then
        intX = 0
        intY = 1
        SaveSetting App.EXEName, "Auto", "Auto", "True"
        While intX < lstDB.ListCount
            SaveSetting App.EXEName, "Database" & intY, "Type", lstDB.List(intX)
            SaveSetting App.EXEName, "Database" & intY, "Server", lstDB.List(intX + 1)
            SaveSetting App.EXEName, "Database" & intY, "UID", lstDBUser.List(intX + 1)
            SaveSetting App.EXEName, "Database" & intY, "PWD", lstDBUser.List(intX + 2)
            intX = intX + 3
            intY = intY + 1
        Wend
    Else
        SaveSetting App.EXEName, "Auto", "Auto", "False"
    End If
    frmTables.Show
    Me.Hide
Else
    MsgBox "Please select at least 1 database to proceed", vbExclamation
End If
End Sub

Private Sub cmdPrevious_Click()
frmWelcome.Show
Me.Hide
End Sub

Private Sub cmdRemove_Click()
Dim lngX As Long
lngX = 0
lngX = lstDB.ListIndex

'remove dbs
lstDB.RemoveItem lngX + 2
lstDB.RemoveItem lngX + 1
lstDB.RemoveItem lngX

'remove users & password
lstDBUser.RemoveItem lngX + 2
lstDBUser.RemoveItem lngX + 1
lstDBUser.RemoveItem lngX

cmdRemove.Enabled = False
End Sub

Private Sub cmdTestConnection_Click()
If testDBConnection(cboDBType.Text, txtServer.Text, txtDB.Text, txtUID.Text, txtPWD.Text) = True Then
    MsgBox "Connection successful", vbInformation
Else
    MsgBox "Connection failed..." & vbCrLf & strError, vbExclamation
End If
End Sub

Private Sub Form_Load()

cboDBType.Text = cboDBType.List(0)
lstDB.Clear
lstDBUser.Clear

Call loadListBoxes
End Sub

Private Sub Form_Unload(Cancel As Integer)
If Me.Visible = True Then
    boolConfirm = MsgBox("Are you sure you want to exit ?", vbYesNoCancel + vbDefaultButton3 + vbQuestion)
    If boolConfirm <> vbYes Then
        Cancel = 1
    End If
    Call unloadAllForms
End If
End Sub

Private Sub lstDB_Click()
lstDB.ToolTipText = lstDB.Text
If UCase(lstDB.Text) = "MS SQL SERVER DATABASE" Or _
   UCase(lstDB.Text) = "MS ACCESS DATABASE" Or _
   UCase(lstDB.Text) = "ORACLE DATABASE" Or _
   UCase(lstDB.Text) = "MYSQL DATABASE" Then
    
    cmdRemove.Enabled = True
Else
    cmdRemove.Enabled = False
End If
End Sub

Function checkExistingDB() As Boolean
Dim intX As Integer
intX = 0
While intX <= lstDB.ListCount
    checkExistingDB = False
    If UCase(cboDBType.Text) = "MS SQL SERVER" Then
        If UCase(lstDB.List(intX)) = UCase(" > " & txtDB.Text & " on " & txtServer.Text) Then
            checkExistingDB = True
            Exit Function
        End If
    ElseIf UCase(cboDBType.Text) = "MS ACCESS" Then
        If UCase(lstDB.List(intX)) = UCase(" > " & txtDB.Text) Then
            checkExistingDB = True
            Exit Function
        End If
    ElseIf UCase(cboDBType.Text) = "ORACLE" Then
        If UCase(lstDB.List(intX)) = UCase(" > " & txtServer.Text) Then
            If UCase(lstDBUser.List(intX + 0)) = UCase(txtUID.Text) And UCase(lstDBUser.List(intX + 1)) = UCase(txtPWD.Text) Then
                checkExistingDB = True
                Exit Function
            End If
        End If
    End If
    intX = intX + 1
Wend
End Function

Sub deleteExistingDB()
On Error GoTo err
DeleteSetting App.EXEName

err:
Exit Sub
End Sub

Sub loadListBoxes()
On Error GoTo err
Dim strAuto As String
Dim strType As String
Dim strServer As String
Dim strUID As String
Dim strPWD As String
Dim intX As Integer
strAuto = "": strType = "": strServer = "": strUID = "": strPWD = ""

chkSave.Value = 0
strAuto = GetSetting(App.EXEName, "Auto", "Auto")
If strAuto = "True" Then
    chkSave.Value = 1
    intX = 1
    While intX <= 1000
        strType = GetSetting(App.EXEName, "Database" & intX, "Type")
        If Len(Trim(strType)) = 0 Then
            Exit Sub
        End If
        strServer = GetSetting(App.EXEName, "Database" & intX, "Server")
        strUID = GetSetting(App.EXEName, "Database" & intX, "UID")
        strPWD = GetSetting(App.EXEName, "Database" & intX, "PWD")
        
        lstDB.AddItem strType
        lstDB.AddItem strServer
        
        lstDB.AddItem " "
        lstDBUser.AddItem " "
        lstDBUser.AddItem strUID
        lstDBUser.AddItem strPWD
        
        intX = intX + 1
    Wend
End If
Exit Sub

err:
Exit Sub

End Sub

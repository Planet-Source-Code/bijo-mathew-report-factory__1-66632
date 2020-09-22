Attribute VB_Name = "modGeneral"
Option Explicit

'registry
Public Const REG_SZ As Long = 1
   Public Const REG_DWORD As Long = 4

   Public Const HKEY_CLASSES_ROOT = &H80000000
   Public Const HKEY_CURRENT_USER = &H80000001
   Public Const HKEY_LOCAL_MACHINE = &H80000002
   Public Const HKEY_USERS = &H80000003

   Public Const ERROR_NONE = 0
   Public Const ERROR_BADDB = 1
   Public Const ERROR_BADKEY = 2
   Public Const ERROR_CANTOPEN = 3
   Public Const ERROR_CANTREAD = 4
   Public Const ERROR_CANTWRITE = 5
   Public Const ERROR_OUTOFMEMORY = 6
   Public Const ERROR_ARENA_TRASHED = 7
   Public Const ERROR_ACCESS_DENIED = 8
   Public Const ERROR_INVALID_PARAMETERS = 87
   Public Const ERROR_NO_MORE_ITEMS = 259

   Public Const KEY_QUERY_VALUE = &H1
   Public Const KEY_SET_VALUE = &H2
   Public Const KEY_ALL_ACCESS = &H3F

   Public Const REG_OPTION_NON_VOLATILE = 0

   Declare Function RegCloseKey Lib "advapi32.dll" (ByVal hKey As Long) As Long
   Declare Function RegCreateKeyEx Lib "advapi32.dll" Alias "RegCreateKeyExA" (ByVal hKey As Long, ByVal lpSubKey As String, ByVal Reserved As Long, ByVal lpClass As String, ByVal dwOptions As Long, ByVal samDesired As Long, ByVal lpSecurityAttributes As Long, phkResult As Long, lpdwDisposition As Long) As Long
   Declare Function RegOpenKeyEx Lib "advapi32.dll" Alias "RegOpenKeyExA" (ByVal hKey As Long, ByVal lpSubKey As String, ByVal ulOptions As Long, ByVal samDesired As Long, phkResult As Long) As Long
   Declare Function RegQueryValueExString Lib "advapi32.dll" Alias "RegQueryValueExA" (ByVal hKey As Long, ByVal lpValueName As String, ByVal lpReserved As Long, lpType As Long, ByVal lpData As String, lpcbData As Long) As Long
   Declare Function RegQueryValueExLong Lib "advapi32.dll" Alias "RegQueryValueExA" (ByVal hKey As Long, ByVal lpValueName As String, ByVal lpReserved As Long, lpType As Long, lpData As Long, lpcbData As Long) As Long
   Declare Function RegQueryValueExNULL Lib "advapi32.dll" Alias "RegQueryValueExA" (ByVal hKey As Long, ByVal lpValueName As String, ByVal lpReserved As Long, lpType As Long, ByVal lpData As Long, lpcbData As Long) As Long
   Declare Function RegSetValueExString Lib "advapi32.dll" Alias "RegSetValueExA" (ByVal hKey As Long, ByVal lpValueName As String, ByVal Reserved As Long, ByVal dwType As Long, ByVal lpValue As String, ByVal cbData As Long) As Long
   Declare Function RegSetValueExLong Lib "advapi32.dll" Alias "RegSetValueExA" (ByVal hKey As Long, ByVal lpValueName As String, ByVal Reserved As Long, ByVal dwType As Long, lpValue As Long, ByVal cbData As Long) As Long

'ends


Dim mboolAsc As Boolean
Dim mctlTxt As VB.TextBox
Public FSO
Public strConString(9) As New ADODB.Connection
Public boolJoin As Boolean
Public boolCancelled As Boolean
Public strError As String
Public boolConfirm
Public dbLocal As New ADODB.Connection
Public boolUnload As Boolean

Public strPrimaryTable As String
Public boolFromRun As Boolean, boolFromPrevious As Boolean
Public strReportTitle As String
Public strTemplateFileName As String
Public intDistictData As Integer
Public strAlias() As String
Public strDeleteCols() As String
Public strCalcField() As String
'wait
    Public Declare Function GetTickCount Lib "kernel32.dll" () As Long
'ends

'wait for given seconds
Public Function fWait(ByVal SecondsToWait As Double) 'Time In seconds
Dim EndTime As Double
EndTime = GetTickCount + SecondsToWait * 1000 '* 1000 Cause u give seconds and GetTickCount uses Milliseconds
Do Until GetTickCount > EndTime
    DoEvents
Loop
End Function

'write to registry
Public Sub SetKeyValue(sKeyName As String, sValueName As String, vValueSetting As Variant, lValueType As Long)
        
    Dim lRetVal As Long         'result of the SetValueEx function
    Dim hKey As Long         'handle of open key

    'open the specified key
    lRetVal = RegOpenKeyEx(HKEY_CURRENT_USER, sKeyName, 0, KEY_SET_VALUE, hKey)
    lRetVal = SetValueEx(hKey, sValueName, lValueType, vValueSetting)
    RegCloseKey (hKey)
End Sub

'to write to registry
Public Function SetValueEx(ByVal hKey As Long, sValueName As String, lType As Long, vValue As Variant) As Long
Dim lValue As Long
Dim sValue As String
Select Case lType
    Case REG_SZ
        sValue = vValue & Chr$(0)
        SetValueEx = RegSetValueExString(hKey, sValueName, 0&, lType, sValue, Len(sValue))
    
    Case REG_DWORD
        lValue = vValue
        SetValueEx = RegSetValueExLong(hKey, sValueName, 0&, lType, lValue, 4)
End Select
End Function

'To read from registry
Function QueryValueEx(ByVal lhKey As Long, ByVal szValueName As String, vValue As Variant) As Long
     Dim cch As Long
     Dim lrc As Long
     Dim lType As Long
     Dim lValue As Long
     Dim sValue As String

     On Error GoTo QueryValueExError

     ' Determine the size and type of data to be read
     lrc = RegQueryValueExNULL(lhKey, szValueName, 0&, lType, 0&, cch)
     If lrc <> ERROR_NONE Then Error 5

     Select Case lType
         ' For strings
         Case REG_SZ:
             sValue = String(cch, 0)

 lrc = RegQueryValueExString(lhKey, szValueName, 0&, lType, sValue, cch)
             If lrc = ERROR_NONE Then
                 vValue = Left$(sValue, cch - 1)
             Else
                 vValue = Empty
             End If
         ' For DWORDS
         Case REG_DWORD:
 lrc = RegQueryValueExLong(lhKey, szValueName, 0&, lType, lValue, cch)
             If lrc = ERROR_NONE Then vValue = lValue
         Case Else
             'all other data types not supported
             lrc = -1
     End Select

QueryValueExExit:
     QueryValueEx = lrc
     Exit Function

QueryValueExError:
     Resume QueryValueExExit
 End Function
   
'To read from registry
Function QueryValue(sKeyName As String, sValueName As String) As String
    Dim lRetVal As Long         'result of the API functions
    Dim hKey As Long         'handle of opened key
    Dim vValue As Variant      'setting of queried value
    
    lRetVal = RegOpenKeyEx(HKEY_CURRENT_USER, sKeyName, 0, KEY_QUERY_VALUE, hKey)
    lRetVal = QueryValueEx(hKey, sValueName, vValue)
    QueryValue = vValue
    RegCloseKey (hKey)
End Function

Function testDBConnection(strDBType As String, strServer As String, strDB As String, strUID As String, strPWD As String) As Boolean
On Error GoTo err:
strError = ""
Dim dbTest As New ADODB.Connection
testDBConnection = True
If UCase(strDBType) <> "MS ACCESS" And UCase(strDBType) <> "MS SQL SERVER" And UCase(strDBType) <> "ORACLE" And UCase(strDBType) <> "MYSQL" And UCase(strDBType) <> "POSTGRESQL" Then
    MsgBox "The database type " & strDBType & " is not supported", vbExclamation
    testDBConnection = False
    Exit Function
Else
    If UCase(strDBType) = "MS ACCESS" Then
        dbTest.Open ("Provider=Microsoft.Jet.OLEDB.4.0;Data Source=" & strDB & ";Jet OLEDB:Database Password=" & strPWD)
    ElseIf UCase(strDBType) = "MS SQL SERVER" Then
        dbTest.Open ("Provider=SQLOLEDB;data Source=" & strServer & ";Initial Catalog=" & strDB & ";User Id=" & strUID & ";Password=" & strPWD & ";")
    ElseIf UCase(strDBType) = "ORACLE" Then
        dbTest.Open "DRIVER={Microsoft ODBC For Oracle};UID=" & strUID & ";PWD=" & strPWD & ";SERVER=" & strServer
    ElseIf UCase(strDBType) = "MYSQL" Then
        dbTest.Open "DRIVER={MySQL ODBC 3.51 Driver};SERVER=" & strServer & ";DATABASE=" & strDB & ";UID=" & strUID & ";PWD=" & strPWD
    ElseIf UCase(strDBType) = "POSTGRESQL" Then
        dbTest.Open "DRIVER={PostgreSQL Unicode};SERVER=" & strServer & ";DATABASE=" & strDB & ";UID=" & strUID & ";PWD=" & strPWD
    Else
        testDBConnection = False
    End If
End If
Exit Function

err:
strError = err.Description
testDBConnection = False
End Function

Sub unloadAllForms()
boolUnload = True
Dim frmForm As Form
For Each frmForm In Forms
    Unload frmForm
    Set frmForm = Nothing
Next
For Each frmForm In Forms
    Unload frmForm
    Set frmForm = Nothing
Next
Set dbLocal = Nothing
End Sub

Sub delDups(lstBox As ListBox)
Dim lngX As Long, lngY As Long
Dim strTemp As String
strTemp = "": lngX = 0: lngY = 0

While lngX <= lstBox.ListCount
    strTemp = lstBox.List(lngX)
    lngY = 0
    While lngY <= lstBox.ListCount
        If UCase(lstBox.List(lngY)) = UCase(lstBox.List(lngX)) And lngX <> lngY Then
            lstBox.RemoveItem lngY
            lngY = lngY - 1
        End If
        lngY = lngY + 1
    Wend
    lngX = lngX + 1
Wend
End Sub

Function chkListMatch(lstBox As ListBox, strVal As String) As Boolean
chkListMatch = False
Dim intX As Long
While intX <= lstBox.ListCount
    If UCase(lstBox.List(intX)) = UCase(strVal) Then
        chkListMatch = True
        Exit Function
    End If
    intX = intX + 1
Wend
End Function

Function chkComboMatch(cboBox As ComboBox, strVal As String) As Boolean
chkComboMatch = False
Dim intX As Long
While intX <= cboBox.ListCount
    If UCase(cboBox.List(intX)) = UCase(strVal) Then
        chkComboMatch = True
        Exit Function
    End If
    intX = intX + 1
Wend
End Function

Function dataType(intType As Long) As String
   If CInt(intType) = 3 Or CInt(intType) = 139 Then
        dataType = "Long"
    ElseIf CInt(intType) = 6 Then
        dataType = "Currency"
    ElseIf CInt(intType) = 7 Or CInt(intType) = 135 Then
        dataType = "Date"
    ElseIf CInt(intType) = 11 Then
        dataType = "YesNo"
    ElseIf CInt(intType) = 203 Then
        dataType = "Memo"
    Else
        dataType = "VarChar"
    End If
End Function

Sub SortFlexiNoArrows(MSFGrid As MSFlexGrid, boolLastRowBlank As Boolean, Optional intColNo As Integer)
With MSFGrid
.FormatString = Replace(.FormatString, " (+)", "")
.FormatString = Replace(.FormatString, " (-)", "")

'set the col no if passed as parameter
If intColNo > 0 And intColNo <= MSFGrid.Cols Then
    MSFGrid.Col = intColNo
End If

'remove blank row
If boolLastRowBlank = True Then
    MSFGrid.Rows = MSFGrid.Rows - 1
End If

'sort
If mboolAsc = False Then
    .Sort = 6
    mboolAsc = True
    .Row = 0
    .Text = .Text & " (-)"
Else
    .Sort = 7
    mboolAsc = False
    .Row = 0
    .Text = .Text & " (+)"
End If

'add blank row
If boolLastRowBlank = True Then
    MSFGrid.Rows = MSFGrid.Rows + 1
End If

Call AltFlexiColors(MSFGrid, 2, 1)
End With
End Sub

Sub SortFlexiArrows(MSFGrid As MSFlexGrid, boolLastRowBlank As Boolean, boolAltFlexiColors As Boolean, Optional sortColNo As Integer)
With MSFGrid
'generate text box randomly
On Error GoTo err
Set mctlTxt = MSFGrid.Parent.Controls.Add("VB.TextBox", "txt_txt_txt")
Set mctlTxt.Container = MSFGrid.Container
mctlTxt.Appearance = 0
mctlTxt.BorderStyle = 0
mctlTxt.Font = "Wingdings"
mctlTxt.Text = "ê"
mctlTxt.BackColor = MSFGrid.BackColorFixed
mctlTxt.ZOrder (0)
mctlTxt.TabStop = False

mctlTxt.Height = 175
mctlTxt.Width = 175
mctlTxt.Locked = True
mctlTxt.Enabled = False
mctlTxt.Visible = False

'set text box top
mctlTxt.Top = MSFGrid.Top + 60
'set the text box left
mctlTxt.Left = MSFGrid.Left + MSFGrid.CellLeft + MSFGrid.CellWidth - 225

'set the col no if passed as parameter
If sortColNo > 0 And sortColNo <= MSFGrid.Cols Then
    MSFGrid.Col = sortColNo
End If

'remove blank row
If boolLastRowBlank = True Then
    MSFGrid.Rows = MSFGrid.Rows - 1
End If

'sort
If mboolAsc = False Then
    .Sort = 6
    .Row = 0
    mctlTxt.Text = "ê"
    mboolAsc = True
Else
    .Sort = 7
    .Row = 0
    mctlTxt.Text = "é"
    mboolAsc = False
End If

'add blank row
If boolLastRowBlank = True Then
    MSFGrid.Rows = MSFGrid.Rows + 1
End If

mctlTxt.Visible = True

If boolAltFlexiColors = True Then
    Call AltFlexiColors(MSFGrid, 2, 1)
End If
Exit Sub

err:
Call fWait(0.2)
MSFGrid.Parent.Controls.Remove "txt_txt_txt"
Resume
End With
End Sub

Sub AltFlexiColors(MSFlexi As MSFlexGrid, startRow As Integer, startCol As Integer)
Dim intX As Integer
Dim intY As Integer
Dim boolY As Boolean

boolY = True
intX = startRow
intY = startCol
MSFlexi.Redraw = False
While intX < MSFlexi.Rows
    MSFlexi.Row = intX
    While intY < MSFlexi.Cols
        MSFlexi.Col = intY
        If boolY = True Then
            MSFlexi.CellBackColor = &HC0FFFF
        Else
            MSFlexi.CellBackColor = vbWhite
        End If
        intY = intY + 1
    Wend
    If boolY = True Then
        boolY = False
    Else
        boolY = True
    End If
    intY = startCol
    intX = intX + 1
Wend
MSFlexi.Redraw = True
End Sub

Public Function FG_AutosizeCols(myGrid As MSFlexGrid, frmForm As Form, _
                                Optional ByVal lFirstCol As Long = -1, _
                                Optional ByVal lLastCol As Long = -1, _
                                Optional bCheckFont As Boolean = False)
  
  Dim lCol As Long, lRow As Long, lCurCol As Long, lCurRow As Long
  Dim lCellWidth As Long, lColWidth As Long
  Dim bFontBold As Boolean
  Dim dFontSize As Double
  Dim sFontName As String
  myGrid.Redraw = False
  If bCheckFont Then
    ' save the forms font settings
    bFontBold = frmForm.FontBold
    sFontName = frmForm.FontName
    dFontSize = frmForm.FontSize
  End If
  
  With myGrid
    If bCheckFont Then
      lCurRow = .Row
      lCurCol = .Col
    End If
    
    If lFirstCol = -1 Then lFirstCol = 0
    If lLastCol = -1 Then lLastCol = .Cols - 1
    
    For lCol = lFirstCol To lLastCol
      lColWidth = 0
      If bCheckFont Then .Col = lCol
      For lRow = 0 To .Rows - 1
        If bCheckFont Then
          .Row = lRow
          frmForm.FontBold = .CellFontBold
          frmForm.FontName = .CellFontName
          frmForm.FontSize = .CellFontSize
        End If
        lCellWidth = frmForm.TextWidth(.TextMatrix(lRow, lCol))
        If lCellWidth > lColWidth Then lColWidth = lCellWidth
      Next lRow
      .ColWidth(lCol) = lColWidth + frmForm.TextWidth("W")
    Next lCol
    
    If bCheckFont Then
      .Row = lCurRow
      .Col = lCurCol
    End If
  End With
  
  If bCheckFont Then
    ' restore the forms font settings
    frmForm.FontBold = bFontBold
    frmForm.FontName = sFontName
    frmForm.FontSize = dFontSize
  End If

'Call SortFlexiArrows(myGrid, True, 1)
myGrid.Redraw = True
End Function

Public Function FG_AutosizeRows(myGrid As MSFlexGrid, frmForm As Form, _
                                Optional ByVal lFirstRow As Long = -1, _
                                Optional ByVal lLastRow As Long = -1, _
                                Optional bCheckFont As Boolean = False)
                                
  ' This will only work for Cells with a Chr(13)
  ' To have it working with WordWrap enabled
  ' you need some other routine
  ' Which has been added too
  myGrid.Redraw = False
  Dim lCol As Long, lRow As Long, lCurCol As Long, lCurRow As Long
  Dim lCellHeight As Long, lRowHeight As Long
  Dim bFontBold As Boolean
  Dim dFontSize As Double
  Dim sFontName As String
  
  If bCheckFont Then
    ' save the forms font settings
    bFontBold = frmForm.FontBold
    sFontName = frmForm.FontName
    dFontSize = frmForm.FontSize
  End If
  
  With myGrid
    If bCheckFont Then
      lCurCol = .Col
      lCurRow = .Row
    End If
    
    If lFirstRow = -1 Then lFirstRow = 0
    If lLastRow = -1 Then lLastRow = .Rows - 1
    
    For lRow = lFirstRow To lLastRow
      lRowHeight = 0
      If bCheckFont Then .Row = lRow
      For lCol = 0 To .Cols - 1
        If bCheckFont Then
          .Col = lCol
          frmForm.FontBold = .CellFontBold
          frmForm.FontName = .CellFontName
          frmForm.FontSize = .CellFontSize
        End If
        lCellHeight = frmForm.TextHeight(.TextMatrix(lRow, lCol))
        If lCellHeight > lRowHeight Then lRowHeight = lCellHeight
      Next lCol
      .RowHeight(lRow) = lRowHeight + frmForm.TextHeight("Wg") / 5
    Next lRow
    
    If bCheckFont Then
      .Row = lCurRow
      .Col = lCurCol
    End If
  End With
  
  If bCheckFont Then
    ' restore the forms font settings
    frmForm.FontBold = bFontBold
    frmForm.FontName = sFontName
    frmForm.FontSize = dFontSize
  End If
myGrid.Redraw = True
End Function

Public Function FG_RemoveColumn(myGrid As MSFlexGrid, ByVal lColumn As Long)
  With myGrid
    .Redraw = False
    If lColumn < .Cols Then
      .ColPosition(lColumn) = .Cols - 1
      .Cols = .Cols - 1
    End If
    .Redraw = True
  End With
End Function

Function chkArrayDups(objArr() As Variant, strString As String, boolCaseSensitive As Boolean) As Boolean
chkArrayDups = False
Dim lngX As Long
lngX = 0
While lngX <= UBound(objArr)
    If boolCaseSensitive = True Then
        If objArr(lngX) = strString Then
            chkArrayDups = True
        End If
    ElseIf boolCaseSensitive = False Then
        If UCase(objArr(lngX)) = UCase(strString) Then
            chkArrayDups = True
        End If
    End If
    lngX = lngX + 1
Wend
End Function

Public Function fchkFolderPath(strFilePath As String, boolCreateFolder As Boolean) As Boolean
If (FSO.folderexists(strFilePath)) Then
    fchkFolderPath = True
Else
    fchkFolderPath = False
    If boolCreateFolder = True Then
        On Error GoTo err:
        FSO.CreateFolder (strFilePath)
        fchkFolderPath = True
    End If
End If
Exit Function

err:
fchkFolderPath = False
End Function

Sub PrintFlexi(strReportTitle As String, msfFlexi As MSFlexGrid)
msfFlexi.Redraw = False
Dim intFontSize As Integer
Dim strFontFace As String
Dim dblCellWidthTot As Double
Dim intInv As Integer
Dim msearchResult

intFontSize = msfFlexi.Font.Size - 7
strFontFace = msfFlexi.Font.Name
intInv = 1
Call fchkFolderPath(App.Path & "\Print", True)
On Error GoTo err:
Set msearchResult = FSO.CreateTextFile(App.Path & "\Print\Report " & intInv & ".htm", True)
msearchResult.writeline "<html>"
msearchResult.writeline "<title>" & strReportTitle & "</title>"
msearchResult.writeline "<body>"

Dim intCol As Integer
Dim intRow As Integer
msearchResult.writeline "<font size='2' face='Arial'><center><b>" & strReportTitle & "</center></b></font>"
msearchResult.writeline "<BR>"
msearchResult.writeline "<table border='1' width='100%' cellspacing='0' cellpadding='0'>"
intCol = 0
dblCellWidthTot = 0
msfFlexi.Row = 0

'calculate space
While intCol < msfFlexi.Cols
    msfFlexi.Col = intCol
    dblCellWidthTot = dblCellWidthTot + msfFlexi.CellWidth
    intCol = intCol + 1
Wend

'set col headings - ignore first contains A,B,C etc
intCol = 0
msfFlexi.Row = 1
msearchResult.writeline "<tr>"
While intCol < msfFlexi.Cols
    msfFlexi.Col = intCol
    msearchResult.writeline "<td align='center' width=" & Format(msfFlexi.CellWidth * 100 / dblCellWidthTot, "##0.00") & "%><font size='" & intFontSize & "' face='" & strFontFace & "'><b> " & msfFlexi.Text & " </font></b></td>"
    intCol = intCol + 1
Wend
msearchResult.writeline "</tr>"

'add data - 2 since 1st col is added in above
intRow = 2
While intRow < msfFlexi.Rows
    msfFlexi.Row = intRow
    intCol = 0
    msearchResult.writeline "<tr>"
    While intCol < msfFlexi.Cols
        msfFlexi.Col = intCol
        msearchResult.writeline "<td align='left' width=" & Format(msfFlexi.CellWidth * 100 / dblCellWidthTot, "##0.00") & "%><font size='" & intFontSize & "' face='" & strFontFace & "'>&nbsp;" & msfFlexi.Text & " </font></td>"
        intCol = intCol + 1
    Wend
    msearchResult.writeline "</tr>"
    intRow = intRow + 1
Wend


msearchResult.writeline "</body>"
msearchResult.writeline "</html>"
frmPrint.WebBrowser1.Navigate (App.Path & "\Print\Report " & intInv & ".htm")
Set msearchResult = Nothing
frmPrint.Show
msfFlexi.Redraw = True
Exit Sub

err:
intInv = intInv + 1
Resume
End Sub

Public Function GetVirtualFileName(strFilePath As String) As String
Dim arrFile() As String
If Len(strFilePath) > 0 Then
    arrFile = Split(strFilePath, "\")
    GetVirtualFileName = arrFile(UBound(arrFile))
Else
    GetVirtualFileName = ""
End If
Exit Function

err:
GetVirtualFileName = ""
End Function

Function chkFileExtension(strFileInfo As String) As String
If strFileInfo <> "" Then
    chkFileExtension = Mid(strFileInfo, Len(strFileInfo) - 2, 3)
End If
End Function

Public Function chkFilePath(strFilePath As String) As Boolean
If (FSO.fileexists(strFilePath)) Then
    chkFilePath = True
Else
    chkFilePath = False
End If
End Function

Public Sub ShowAppHelp(Optional IndexID As Integer)

     On Error GoTo err

    With frmPrint.CD
        .HelpFile = App.HelpFile
        
        If IndexID = 0 Then
            .HelpCommand = cdlHelpContents
        Else
            .HelpContext = IndexID
            .HelpCommand = cdlHelpContext
        End If
        
        .ShowHelp
    End With
    
    Exit Sub
 
err:
   MsgBox err.Description, vbExclamation

End Sub

Sub createLocalDB()
Dim intX As Integer
intX = 1
While intX <= 1000
    On Error Resume Next
    FSO.deletefile App.Path & "\ReportsFactory " & intX & ".mdb"
    intX = intX + 1
Wend


intX = 1
On Error GoTo err:
DBEngine.CreateDatabase App.Path & "\ReportsFactory " & intX & ".mdb", dbLangGeneral & ";pwd=sslmmm"
Set dbLocal = Nothing
dbLocal.Open ("Provider=Microsoft.Jet.OLEDB.4.0;Data Source=" & App.Path & "\ReportsFactory " & intX & ".mdb" & ";Jet OLEDB:Database Password=sslmmm")

Exit Sub

err:
intX = intX + 1
Resume
End Sub

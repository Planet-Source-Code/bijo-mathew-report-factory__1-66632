Attribute VB_Name = "modHTML"
'Option Explicit
'
'Private Declare Function GetSysColor Lib "user32" (ByVal nIndex As Long) As Long
'
'Private Const EXTRAWIDTH As Double = 1.2
'
'Private Function HTMLColor(ByVal lColor As Long) As String
'  Dim sTemp As String
'
'  ' convert to hex
'  sTemp = Hex$(lColor)
'
'  ' handle system colors
'  If Len(sTemp) > 6 Then
'    If Left$(sTemp, 1) = "8" Then
'      lColor = Val("&H" & Mid$(sTemp, 2))
'      lColor = GetSysColor(lColor)
'      sTemp = Hex$(lColor)
'    End If
'  End If
'
'  ' build format
'  sTemp = String(6 - Len(sTemp), "0") & sTemp
'  HTMLColor = """#" & Right$(sTemp, 2) & Mid$(sTemp, 3, 2) & Left$(sTemp, 2) & """"
'End Function
'
'Private Function HTMLText(ByVal sLine As String) As String
'
'  If Len(sLine) = 0 Then
'    HTMLText = " "
'  Else
'    HTMLText = Replace$(sLine, "&", "&")
'    HTMLText = Replace$(HTMLText, "<", "<")
'    HTMLText = Replace$(HTMLText, ">", ">")
'  End If
'
'End Function
'
'Public Function FlexGridToHTML(FG As MSFlexGrid) As String
'  Dim sData As String, sLine As String
'  Dim dTblWidth As Double
'  Dim i As Long, lRow As Long, lCol As Long
'  Dim sSpan As String
'  Dim lRow1 As Long, lRow2 As Long, lCol1 As Long, lCol2 As Long
'  Dim dWidth As Double
'  Dim sText As String, sTemp As String
'  Dim sBackGround As String, lColor As Long
'  Dim sFont As String, sBorder As String, sFontFX As String
'  Dim sAlign As String, sCell As String
'  Dim bProcessCell As Boolean
'
'  With FG
'    .Redraw = False
'
'    ' get total table width in pixels
'    dTblWidth = 0
'    For lCol = 0 To .Cols - 1
'      dTblWidth = dTblWidth + .ColWidth(lCol)
'    Next lCol
'    dTblWidth = EXTRAWIDTH * dTblWidth / Screen.TwipsPerPixelX
'
'    ' save table header
'    sData = "<table border cellspacing=0 cellpadding=2 vAlign=center" & _
'            " bgcolor=" & HTMLColor(.BackColor) & _
'            " width=" & Format(Int(dTblWidth)) & _
'            ">" & vbCrLf
'
'    ' loop through the rows
'    For lRow = 0 To .Rows - 1
'
'      sLine = ""
'
'      ' skip hidden rows
'      If .RowHeight(lRow) > 0 Then
'
'        ' start row
'        sLine = "<tr>"
'
'        ' loop through the columns
'        For lCol = 0 To .Cols - 1
'
'          ' skip hidden cols
'          If .ColWidth(lCol) > 0 Then
'
'            .Col = lCol
'            .Row = lRow
'            bProcessCell = True
'
'            ' handle merges
'            sSpan = ""
'            GetMergedCols FG, lRow, lCol, lCol1, lCol2
'            GetMergedRows FG, lRow, lCol, lRow1, lRow2
'            If lCol1 < lCol Then bProcessCell = False
'            If lRow1 < lRow Then bProcessCell = False
'
'            If bProcessCell Then
'              If lCol2 > lCol Then sSpan = " colspan=" & (lCol2 - lCol + 1)
'              If lRow2 > lRow Then sSpan = sSpan & " rowspan=" & (lRow2 - lRow + 1)
'
'              ' get col width
'              dWidth = 0
'              For i = lCol1 To lCol2
'                dWidth = dWidth + .ColWidth(i)
'              Next
'              dWidth = EXTRAWIDTH * dWidth / Screen.TwipsPerPixelX
'
'              ' get cell text
'              sText = HTMLText(.TextMatrix(lRow, lCol))
'
'              ' get back color
'              sBackGround = ""
'              lColor = .CellBackColor
'              If lColor <> 0 Then
'                sBackGround = " bgcolor=" & HTMLColor(lColor)
'              ElseIf lRow < .FixedRows Or lCol < .FixedCols Then
'                sBackGround = " bgcolor=" & HTMLColor(.BackColorFixed)
'              End If
'
'              ' get border color
'              sBorder = ""
'              If lRow < .FixedRows Or lCol < .FixedCols Then
'                sBorder = " bordercolor=" & HTMLColor(.GridColorFixed)
'              Else
'                sBorder = " bordercolor=" & HTMLColor(.GridColor)
'              End If
'
'              ' get fore color and font name
'              sFont = " size=2"
'              sTemp = .CellFontName
'              If sTemp <> .FontName Then
'                sFont = " face=" & """" & sTemp & """"
'              End If
'              lColor = .CellForeColor
'              If lColor <> 0 Then sFont = " color=" & HTMLColor(lColor)
'
'              ' get font effects
'              sFontFX = ""
'              If .CellFontBold Then sFontFX = sFontFX & "<B>"
'              If .CellFontItalic Then sFontFX = sFontFX & "<I>"
'              If .CellFontUnderline Then sFontFX = sFontFX & "<U>"
'
'              ' get alignment
'              sAlign = ""
'              Select Case .CellAlignment
'                Case flexAlignCenterBottom
'                  sAlign = " align=center valign=bottom"
'                Case flexAlignCenterCenter
'                  sAlign = " align=center"
'                Case flexAlignCenterTop
'                  sAlign = " align=center valign=top"
'                Case flexAlignLeftBottom
'                  sAlign = " valign=bottom"
'                Case flexAlignLeftCenter
'                  sAlign = ""
'                Case flexAlignLeftTop
'                  sAlign = " valign=top"
'                Case flexAlignRightBottom
'                  sAlign = " align=right valign=bottom"
'                Case flexAlignRightCenter
'                  sAlign = " align=right"
'                Case flexAlignRightTop
'                  sAlign = " align=right valign=top"
'                Case Else
'                  If IsNumeric(.TextMatrix(lRow, lCol)) Then
'                    sAlign = " align=right valign=bottom"
'                  End If
'              End Select
'
'              ' build HTML cell string
'              sTemp = """" & Format(dWidth / dTblWidth, "#%") & """"
'              sCell = "<td width=" & sTemp & sBackGround & sAlign & sBorder & sSpan & ">"
'              If sFont <> "" Then sCell = sCell & "<FONT" & sFont & ">"
'              sCell = sCell & sFontFX & sText
'              If InStr(sFontFX, "B") > 0 Then sCell = sCell & "</B>"
'              If InStr(sFontFX, "I") > 0 Then sCell = sCell & "</I>"
'              If InStr(sFontFX, "U") > 0 Then sCell = sCell & "</U>"
'              If sFont <> "" Then sCell = sCell & "</font>"
'
'              ' end cell
'              sCell = sCell & "</td>"
'              sLine = sLine & sCell
'
'            End If ' ProcessCell
'          End If ' .ColWidth(lCol) > 0 Then
'
'        Next lCol
'
'        ' end row
'        If Len(sLine) > 0 Then sData = sData & sLine & "</tr>" & vbCrLf
'
'      End If ' .RowHeight(lRow) > 0 Then
'    Next lRow
'
'    .Redraw = True
'  End With
'  ' table end
'  sData = sData & "</table></font>"
'
'  ' return success
'  FlexGridToHTML = sData
'End Function
'
'Private Sub GetMergedCols(FG As MSFlexGrid, ByVal Row As Long, _
'  ByVal Col As Long, ByRef lStart As Long, ByRef lEnd As Long)
'  Dim lCol As Long
'  Dim lCnt As Long
'
'  lStart = Col
'  lEnd = Col
'
'  With FG
'    If Row < .FixedRows Then
'      For lCol = Col - 1 To 0 Step -1
'        If .ColWidth(lCol) <> 0 Then
'          If .TextMatrix(Row, lCol) = .TextMatrix(Row, Col) Then
'            lCnt = lCnt + 1
'          Else
'            Exit For
'          End If
'        End If
'      Next lCol
'      If lCnt > 0 Then lStart = Col - lCnt
'
'      lCnt = 0
'      For lCol = Col + 1 To .Cols - 1
'        If .ColWidth(lCol) <> 0 Then
'          If .TextMatrix(Row, lCol) = .TextMatrix(Row, Col) Then
'            lCnt = lCnt + 1
'          Else
'            Exit For
'          End If
'        End If
'      Next lCol
'      If lCnt > 0 Then lEnd = Col + lCnt
'
'    End If
'  End With
'End Sub
'
'Private Sub GetMergedRows(FG As MSFlexGrid, ByVal Row As Long, _
'  ByVal Col As Long, ByRef lStart As Long, ByRef lEnd As Long)
'  Dim lRow As Long
'  Dim lCnt As Long
'
'  lStart = Row
'  lEnd = Row
'
'  With FG
'    If Col < .FixedCols Then
'      For lRow = Row - 1 To 0 Step -1
'        If .RowHeight(lRow) <> 0 Then
'          If .TextMatrix(lRow, Col) = .TextMatrix(Row, Col) Then
'            lCnt = lCnt + 1
'          Else
'            Exit For
'          End If
'        End If
'      Next lRow
'      If lCnt > 0 Then lStart = Row - lCnt
'
'      For lRow = Row + 1 To .Rows - 1
'        If .RowHeight(lRow) <> 0 Then
'          If .TextMatrix(lRow, Col) = .TextMatrix(Row, Col) Then
'            lCnt = lCnt + 1
'          Else
'            Exit For
'          End If
'        End If
'      Next lRow
'      If lCnt > 0 Then lEnd = Row + lCnt
'    End If
'
'  End With
'End Sub
'

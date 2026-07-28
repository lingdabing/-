Attribute VB_Name = "模块1"


' ==================================================
' 唯一国家判定（核心修复）
' ==================================================
Function GetMatchedCountry(addr As String) As String
    Dim c As Variant
    addr = LCase(addr)
    GetMatchedCountry = ""

    ' ---- 高优先级（必须唯一）----
    If InStr(addr, "united kingdom") > 0 Then
        GetMatchedCountry = "United Kingdom"
        Exit Function
    End If

    If InStr(addr, "united states") > 0 Then
        GetMatchedCountry = "United States"
        Exit Function
    End If

    If InStr(addr, "south korea") > 0 Then
        GetMatchedCountry = "Korea"
        Exit Function
    End If

    ' ---- 欧洲 DHL ----
    Dim euDHLList As Variant
    euDHLList = Array( _
        "austria", "france", "italy", "spain", "netherlands", "belgium", _
        "czech republic", "poland", "sweden", "finland", "denmark", _
        "ireland", "portugal", "luxembourg", "greece", "hungary", _
        "romania", "slovakia", "slovenia", "croatia", "estonia", _
        "lithuania", "bulgaria", "latvia", "germany" _
    )

    For Each c In euDHLList
        If InStr(addr, c) > 0 Then
            GetMatchedCountry = StrConv(c, vbProperCase)
            Exit Function
        End If
    Next c

    ' ---- EUB ----
    Dim eubList As Variant
    eubList = Array( _
        "andorra", "armenia", "azerbaijan", "cyprus", _
        "iceland", "liechtenstein", "malta", "norway", _
        "san marino", "switzerland", "serbia", _
        "greenland", "faroe islands" _
    )

    For Each c In eubList
        If InStr(addr, c) > 0 Then
            GetMatchedCountry = StrConv(c, vbProperCase)
            Exit Function
        End If
    Next c

    ' ---- 顺丰 ----
    Dim sfList As Variant
    sfList = Array("Hongkong", "korea", "singapore", "thailand", "taiwan", "malaysia")

    For Each c In sfList
        If InStr(addr, c) > 0 Then
            GetMatchedCountry = StrConv(c, vbProperCase)
            Exit Function
        End If
    Next c
End Function


Sub ProcessExcelData()

    Dim ws As Worksheet
    Dim lastRow As Long
    Dim r As Long
    Dim cell As Range
    Dim rngG As Range, rngHIJ As Range, rngIJ As Range, rngD As Range
    Dim items As Object, itemArray As Variant, Item As Variant
    Dim currentK As String
    Dim matchedCountry As String

    Dim euDHLList As Variant, eubCountryList As Variant, c As Variant
    Dim hasPrice As Boolean
    Dim priceValue As Double

    Dim pVal As String, qVal As String, rVal As String, sVal As String
    Dim mergedVal As String
    Dim rowHasData As Boolean

    Set ws = ActiveSheet
    lastRow = ws.Cells(ws.Rows.Count, "G").End(xlUp).Row

    ' ==================================================
    ' ① 合并 P-Q-R-S → 写入 D → 清空
    ' ==================================================
    For r = 2 To lastRow
        pVal = Trim(ws.Cells(r, "P").Value)
        qVal = Trim(ws.Cells(r, "Q").Value)
        rVal = Trim(ws.Cells(r, "R").Value)
        sVal = Trim(ws.Cells(r, "S").Value)

        mergedVal = pVal
        If qVal <> "" Then mergedVal = mergedVal & "," & qVal
        If rVal <> "" Then mergedVal = mergedVal & "," & rVal
        If sVal <> "" Then mergedVal = mergedVal & "," & sVal

        If mergedVal <> "" Then ws.Cells(r, "D").Value = mergedVal

        ws.Cells(r, "P").ClearContents
        ws.Cells(r, "Q").ClearContents
        ws.Cells(r, "R").ClearContents
        ws.Cells(r, "S").ClearContents
    Next r

    ' 国家列表
    euDHLList = Array("Austria", "France", "Italy", "Spain", "Netherlands", "Belgium", _
                      "Czech Republic", "Poland", "Sweden", "Finland", "Denmark", _
                      "Ireland", "Portugal", "Luxembourg", "Greece", "Hungary", _
                      "Romania", "Slovakia", "Slovenia", "Croatia", "Estonia", _
                      "Lithuania", "Bulgaria", "Latvia", "Germany")

    eubCountryList = Array("Andorra", "Armenia", "Azerbaijan", "Cyprus", "United Kingdom", _
                           "Iceland", "Liechtenstein", "Malta", "Norway", "San Marino", _
                           "Switzerland", "Serbia", "Greenland", "Faroe Islands")

    ' ==================================================
    ' ② G 列（带电）
    ' ==================================================
    Set rngG = ws.Range("G2:G" & lastRow)
    For Each cell In rngG
        If cell.Value <> "" Then
            cell.Value = Replace(cell.Value, ",", vbLf)
            cell.Value = Replace(cell.Value, "*1", "")
            cell.WrapText = True

            If InStr(cell.Value, "FW-SH-17M45X") > 0 Then
                currentK = ws.Cells(cell.Row, "K").Value
                If InStr(currentK, "带电") = 0 Then
                    ws.Cells(cell.Row, "K").Value = IIf(currentK = "", "带电", currentK & vbLf & "带电")
                End If
            End If
        End If
    Next cell

    ' ==================================================
    ' ③ H-I-J
    ' ==================================================
    Set rngHIJ = ws.Range("H2:J" & lastRow)
    For Each cell In rngHIJ
        If cell.Value <> "" Then
            cell.Value = Replace(cell.Value, ";", vbLf)
            cell.WrapText = True
        End If
    Next cell

    ' ==================================================
    ' ④ I-J 去重
    ' ==================================================
    Set rngIJ = ws.Range("I2:J" & lastRow)
    For Each cell In rngIJ
        If cell.Value <> "" Then
            Set items = CreateObject("Scripting.Dictionary")
            itemArray = Split(cell.Value, vbLf)
            For Each Item In itemArray
                If Trim(Item) <> "" Then
                    If Not items.Exists(Trim(Item)) Then items.Add Trim(Item), Nothing
                End If
            Next Item
            cell.Value = Join(items.Keys, vbLf)
        End If
    Next cell

    ' ==================================================
    ' ⑤ 国家 & 物流（唯一国家版）
    ' ==================================================
    Set rngD = ws.Range("D2:D" & lastRow)

    For Each cell In rngD
        If cell.Value <> "" Then

            matchedCountry = GetMatchedCountry(cell.Value)

            ' 统一价格
            hasPrice = False
            If IsNumeric(ws.Cells(cell.Row, "N").Value) Then
                priceValue = ws.Cells(cell.Row, "N").Value
                hasPrice = True
            ElseIf IsNumeric(ws.Cells(cell.Row, "L").Value) Then
                priceValue = ws.Cells(cell.Row, "L").Value
                hasPrice = True
            End If

            ' 欧洲 DHL
            For Each c In euDHLList
                If matchedCountry = c And hasPrice And priceValue < 149 And ws.Cells(cell.Row, "O").Value < 35 Then
                    currentK = ws.Cells(cell.Row, "K").Value
                    If InStr(currentK, "出欧洲DHL专线小包") = 0 Then
                        ws.Cells(cell.Row, "K").Value = IIf(currentK = "", "出欧洲DHL专线小包", currentK & vbLf & "出欧洲DHL专线小包")
                    End If
                End If
            Next c

            ' EUB
            For Each c In eubCountryList
                If matchedCountry = c And hasPrice And priceValue < 149 And ws.Cells(cell.Row, "O").Value < 35 Then
                    currentK = ws.Cells(cell.Row, "K").Value
                    If InStr(currentK, "出EUB") = 0 Then
                        ws.Cells(cell.Row, "K").Value = IIf(currentK = "", "出EUB", currentK & vbLf & "出EUB")
                    End If
                End If
            Next c

            ' 美国
            If matchedCountry = "United States" And hasPrice Then

                If priceValue < 50 And ws.Cells(cell.Row, "O").Value <= 20 Then
                    currentK = ws.Cells(cell.Row, "K").Value
                    If InStr(currentK, "出EUB") = 0 Then
                        ws.Cells(cell.Row, "K").Value = IIf(currentK = "", "出EUB", currentK & vbLf & "出EUB")
                    End If
                End If

                If IsNumeric(ws.Cells(cell.Row, "N").Value) Then
                    ws.Cells(cell.Row, "L").Value = ws.Cells(cell.Row, "N").Value
                    ws.Cells(cell.Row, "N").ClearContents
                End If

                currentK = ws.Cells(cell.Row, "K").Value
                If InStr(currentK, "带磁") = 0 Then
                    ws.Cells(cell.Row, "K").Value = IIf(currentK = "", "带磁", currentK & vbLf & "带磁")
                End If
            End If

            ' 顺丰
            If matchedCountry = "HongKong" Or matchedCountry = "Korea" Or _
               matchedCountry = "Singapore" Or matchedCountry = "Thailand" Or _
               matchedCountry = "Taiwan" Or matchedCountry = "Malaysia" Then

                currentK = ws.Cells(cell.Row, "K").Value
                If InStr(currentK, "发顺丰") = 0 Then
                    ws.Cells(cell.Row, "K").Value = IIf(currentK = "", "发顺丰", currentK & vbLf & "发顺丰")
                End If
            End If

            ws.Cells(cell.Row, "K").WrapText = True
        End If
    Next cell

    ' ==================================================
    ' ⑥ 样式
    ' ==================================================
    For r = 2 To lastRow
        rowHasData = Application.WorksheetFunction.CountA(ws.Range("A" & r & ":L" & r)) > 0
        If rowHasData Then
            With ws.Range("A" & r & ":L" & r)
                .Borders.LineStyle = xlContinuous
                .Borders.Weight = xlThin
                .HorizontalAlignment = xlCenter
                .VerticalAlignment = xlCenter
            End With
        End If
    Next r

   ' ==================================================
' ⑦ 将 T 列剪切到 D 列前面，并给 D 列加边框
' ==================================================
ws.Columns("T").Cut
ws.Columns("D").Insert Shift:=xlToRight
Application.CutCopyMode = False

' 给新的 D 列加边框并居中
With ws.Range("D1:D" & lastRow)
    .Borders.LineStyle = xlContinuous
    .Borders.Weight = xlThin
    .HorizontalAlignment = xlCenter
    .VerticalAlignment = xlCenter
End With

    MsgBox "处理完成：唯一国家判定 + 物流规则已稳定运行", vbInformation

End Sub


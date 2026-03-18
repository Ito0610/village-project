Attribute VB_Name = "Module1"
Option Explicit

' ======================================================================================
' サービス記録自動入力マクロ 【完成版①】 (VBA版 - Mac完全対応修正v17)
'
' 本ファイル: 完成版①として保存。X〜AC列の数式を保持する書き戻し対応済み。
'
' 変更点 (Mac対応v17):
'   - 増コードを使う条件を限定。181分未満または時間跨ぎなしの場合は1コードで請求。
'     増コードは「181分以上かつ時間跨ぎ（日中→夜間など）」のときのみ使用。
'
' 変更点 (Mac対応v16):
'   - コードマッチングで「基本+増」の分割を優先するよう変更。
'     例: 181分以上かつ時間跨ぎのとき、日中2.5＋夜間1.0 → 160185＋160271を採用。
'
' 変更点 (Mac対応v15):
'   - 結果の書き戻しを「P〜W」と「AD〜AI」の2ブロックのみに変更。
'     X〜AC には一切書き込まないため、入っている数式が保持される。
'
' 変更点 (Mac対応v14):
'   - X列〜AC列(24〜29列)をマクロで上書きしないよう修正。
'     Flags の書き込み対象を P,Q および AD〜AI に限定。
'
' 変更点 (Mac対応v13):
'   - 2人介助コードの出力先を条件分岐:
'     A) 時間重複ルールで「2人介助」になった場合 -> R列(Main)に出力、U列は空。
'     B) 元から「2人介助」だった場合 -> U列(Add)に出力、R列は1人介助コード(従来通り)。
'
' 変更点 (Mac対応v12):
'   - Helper=2の出力先変更 (v13で条件付きに修正)
' ======================================================================================

' --- カスタム型定義 (Masterデータの構造体) ---
Private Type MasterRecord
    code As String          ' A列
    Name As String          ' B列
    Price As String         ' J列 (単価/備考)
    helpers As String       ' D列 (人数)
    Early As Double         ' E列 (早朝)
    Day As Double           ' F列 (日中)
    Night As Double         ' G列 (夜間)
    Deep As Double          ' H列 (深夜)
    IsIncrease As Boolean   ' J列に"増"が含まれるか
    Flags() As Variant      ' P-AI列 (フラグ群)
End Type

Private Type TimeBucket
    Early As Double
    Day As Double
    Night As Double
    Deep As Double
End Type

' ======================================================================================
' メイン処理
' ======================================================================================
Sub ProcessAllSheets()
    Dim wb As Workbook
    Dim wsMaster As Worksheet
    Dim ws As Worksheet
    Dim processedCount As Long
    Dim writtenCount As Long
    Dim masterData() As MasterRecord

    Set wb = ThisWorkbook

    On Error Resume Next
    Set wsMaster = wb.Worksheets("Master")
    On Error GoTo 0

    If wsMaster Is Nothing Then
        MsgBox "エラー: 'Master' シートが見つかりません。", vbCritical
        Exit Sub
    End If

    If Not LoadMasterData(wsMaster, masterData) Then
        MsgBox "エラー: Masterデータの読み込みに失敗しました。", vbCritical
        Exit Sub
    End If

    processedCount = 0
    writtenCount = 0

    For Each ws In wb.Worksheets
        If ws.Name <> "Master" Then
            Application.StatusBar = "処理中: " & ws.Name
            If ProcessSheet(ws, masterData) Then writtenCount = writtenCount + 1
            processedCount = processedCount + 1
        End If
    Next ws

    Application.StatusBar = False
    MsgBox processedCount & " 枚のシート処理が完了しました。" & vbCrLf & vbCrLf & "記入を行ったシート数: " & writtenCount, vbInformation
End Sub

' ======================================================================================
' Masterデータの読み込み処理
' ======================================================================================
Function LoadMasterData(ws As Worksheet, ByRef outData() As MasterRecord) As Boolean
    Dim lastRow As Long
    Dim dataRange As Variant
    Dim i As Long, j As Long
    Dim count As Long
    Dim bodyCheck As String
    Dim flagArr() As Variant

    lastRow = ws.Cells(ws.Rows.Count, "A").End(xlUp).Row
    If lastRow < 2 Then
        LoadMasterData = False
        Exit Function
    End If

    dataRange = ws.Range("A2").Resize(lastRow - 1, 35).Value

    ReDim outData(1 To UBound(dataRange, 1))
    count = 0

    For i = 1 To UBound(dataRange, 1)
        bodyCheck = VBA.Trim(CStr(dataRange(i, 3)))

        If bodyCheck = "あり" Then
            count = count + 1

            With outData(count)
                .code = CStr(dataRange(i, 1))
                .Name = CStr(dataRange(i, 2))
                .helpers = ParseHelperCount(CStr(dataRange(i, 4)))
                .Early = ParseDuration(CStr(dataRange(i, 5)))
                .Day = ParseDuration(CStr(dataRange(i, 6)))
                .Night = ParseDuration(CStr(dataRange(i, 7)))
                .Deep = ParseDuration(CStr(dataRange(i, 8)))
                .Price = CStr(dataRange(i, 10))
                .IsIncrease = (VBA.InStr(.Price, "増") > 0)

                ReDim flagArr(0 To 19)
                For j = 0 To 19
                    flagArr(j) = dataRange(i, 16 + j)
                Next j
                .Flags = flagArr
            End With
        End If
    Next i

    If count > 0 Then
        ReDim Preserve outData(1 To count)
        LoadMasterData = True
    End If
End Function

Function ParseDuration(strIn As String) As Double
    If strIn = "" Then
        ParseDuration = 0
        Exit Function
    End If

    Dim s As String
    Dim result As String
    Dim i As Long
    Dim c As String
    Dim hasDot As Boolean

    s = ToHalfWidth(strIn)
    result = ""
    hasDot = False

    For i = 1 To VBA.Len(s)
        c = VBA.Mid(s, i, 1)
        If IsNumeric(c) Then
            result = result & c
        ElseIf c = "." Then
            If Not hasDot And VBA.Len(result) > 0 Then
                result = result & c
                hasDot = True
            End If
        Else
            If VBA.Len(result) > 0 Then Exit For
        End If
    Next i

    If result <> "" Then
        ParseDuration = VBA.Val(result)
    Else
        ParseDuration = 0
    End If
End Function

Function ParseHelperCount(strIn As String) As String
    If strIn = "" Then
        ParseHelperCount = ""
        Exit Function
    End If
    ParseHelperCount = VBA.Trim(ToHalfWidth(strIn))
End Function

Function ToHalfWidth(strIn As String) As String
    Dim s As String
    Dim i As Long
    Dim c As String
    Dim code As Long

    s = ""
    For i = 1 To VBA.Len(strIn)
        c = VBA.Mid(strIn, i, 1)
        code = AscW(c) And &HFFFF&
        If code >= 65296 And code <= 65305 Then
            s = s & VBA.ChrW(code - 65248)
        ElseIf code = 65294 Then
            s = s & "."
        ElseIf code = 12288 Then
            s = s & " "
        Else
            s = s & c
        End If
    Next i
    ToHalfWidth = s
End Function

' ======================================================================================
' シート単位の処理
' ======================================================================================
Function ProcessSheet(ws As Worksheet, ByRef masterData() As MasterRecord) As Boolean
    Dim startRow As Long, numRows As Long
    Dim inputData As Variant
    Dim outputData As Variant
    Dim i As Long
    Dim rDate As Variant, rStart As Variant, rEnd As Variant
    Dim dateObj As Date, startObj As Date, endObj As Date
    Dim helperCount As String

    Dim currentGroup As Collection
    Dim groups As Collection
    Dim rowArr() As Variant
    Dim diffMinutes As Long
    Dim hadEntry As Boolean
    hadEntry = False

    startRow = 16
    numRows = 20

    If ws.UsedRange.Rows.Count < startRow Then ProcessSheet = False: Exit Function

    inputData = ws.Range(ws.Cells(startRow, 1), ws.Cells(startRow + numRows - 1, 21)).Value
    outputData = ws.Range(ws.Cells(startRow, 16), ws.Cells(startRow + numRows - 1, 35)).Value

    Dim calcR() As String, calcU() As String
    ReDim calcR(1 To numRows)
    ReDim calcU(1 To numRows)

    Set groups = New Collection
    Set currentGroup = New Collection

    Dim prevDate As Date, prevStart As Date, prevEnd As Date
    prevDate = 0: prevStart = 0: prevEnd = 0

    ' --- 1. グルーピング ---
    For i = 1 To numRows
        rDate = inputData(i, 1)
        rStart = inputData(i, 10)   ' J列
        rEnd = inputData(i, 11)     ' K列
        helperCount = ParseHelperCount(CStr(inputData(i, 13))) ' M列

        If Not TryParseDate(rDate, dateObj) Then GoTo CheckGroup
        If Not TryNormalizeTime(dateObj, rStart, startObj) Then GoTo CheckGroup
        If Not TryNormalizeTime(dateObj, rEnd, endObj) Then GoTo CheckGroup

        If endObj < startObj Then endObj = DateAdd("d", 1, endObj)

        Dim isOverlap As Boolean: isOverlap = False

        ' 【追加ルール】同一日時（時間が重複している）の場合
        ' 上の行はそのまま、下の行を「2人介助」として扱う (Helper="2"に変更)
        ' M列が "1" の場合のみ適用 (元々 "2" ならそのまま)
        ' 重複判定: (開始A < 終了B) かつ (開始B < 終了A) ※重複があればTrue
        If i > 1 And helperCount = "1" Then
             If IsSameDay(dateObj, prevDate) Then
                If startObj < prevEnd And prevStart < endObj Then
                    helperCount = "2"
                    isOverlap = True
                End If
             End If
        End If

        ' 次回判定用に現在値を保存
        prevDate = dateObj
        prevStart = startObj
        prevEnd = endObj

        ReDim rowArr(0 To 5)
        rowArr(0) = i: rowArr(1) = dateObj: rowArr(2) = startObj: rowArr(3) = endObj: rowArr(4) = helperCount
        rowArr(5) = isOverlap

        If currentGroup.Count = 0 Then
            currentGroup.Add rowArr
        Else
            Dim lastArr As Variant
            lastArr = currentGroup.Item(currentGroup.Count)

            If Not IsSameDay(CDate(lastArr(1)), dateObj) Then
                groups.Add currentGroup
                Set currentGroup = New Collection
                currentGroup.Add rowArr
            Else
                diffMinutes = DateDiff("n", CDate(lastArr(3)), startObj)

                If diffMinutes <= 119 And diffMinutes > -1 Then
                    currentGroup.Add rowArr
                Else
                    groups.Add currentGroup
                    Set currentGroup = New Collection
                    currentGroup.Add rowArr
                End If
            End If
        End If
        GoTo NextRow

CheckGroup:
        If currentGroup.Count > 0 Then
            groups.Add currentGroup
            Set currentGroup = New Collection
        End If
NextRow:
    Next i

    If currentGroup.Count > 0 Then groups.Add currentGroup

    ' --- 2. 計算 & マッチング ---
    Dim g As Variant, item As Variant
    Dim lastItemIdx As Long

    For Each g In groups
        Dim rawBuckets As TimeBucket
        rawBuckets.Early = 0: rawBuckets.Day = 0: rawBuckets.Night = 0: rawBuckets.Deep = 0

        Dim firstItem As Variant: firstItem = g.Item(1)
        Dim grpHelper As String: grpHelper = CStr(firstItem(4))
        Dim isForced2P As Boolean: isForced2P = CBool(firstItem(5))

        For Each item In g
            Dim b As TimeBucket
            b = CalculateBuckets(CDate(item(2)), CDate(item(3)))
            rawBuckets.Early = rawBuckets.Early + b.Early
            rawBuckets.Day = rawBuckets.Day + b.Day
            rawBuckets.Night = rawBuckets.Night + b.Night
            rawBuckets.Deep = rawBuckets.Deep + b.Deep
            lastItemIdx = item(0)
        Next item

        Dim existingR As String: existingR = CStr(inputData(lastItemIdx, 18))
        Dim existingU As String: existingU = CStr(inputData(lastItemIdx, 21))
        Dim proposedR As String: proposedR = ""
        Dim proposedU As String: proposedU = ""

        Dim totalRaw As Double
        totalRaw = rawBuckets.Early + rawBuckets.Day + rawBuckets.Night + rawBuckets.Deep

        If totalRaw > 0 Then
            Dim roundedUnits As Long
            roundedUnits = Int((totalRaw + 15) / 30)
            If roundedUnits = 0 Then roundedUnits = 1

            Dim finalBuckets As TimeBucket, remBuckets As TimeBucket
            ' 【重要修正】 finalBuckets を明示的に初期化 (ループ変数の再利用対策)
            finalBuckets.Early = 0: finalBuckets.Day = 0: finalBuckets.Night = 0: finalBuckets.Deep = 0

            remBuckets = rawBuckets

            Dim u As Long
            For u = 1 To roundedUnits
                Dim maxKey As String, maxVal As Double
                maxVal = -1
                If remBuckets.Early > maxVal Then maxVal = remBuckets.Early: maxKey = "Early"
                If remBuckets.Day > maxVal Then maxVal = remBuckets.Day: maxKey = "Day"
                If remBuckets.Night > maxVal Then maxVal = remBuckets.Night: maxKey = "Night"
                If remBuckets.Deep > maxVal Then maxVal = remBuckets.Deep: maxKey = "Deep"
                Select Case maxKey
                    Case "Early": finalBuckets.Early = finalBuckets.Early + 0.5: remBuckets.Early = remBuckets.Early - 30
                    Case "Day":   finalBuckets.Day = finalBuckets.Day + 0.5:     remBuckets.Day = remBuckets.Day - 30
                    Case "Night": finalBuckets.Night = finalBuckets.Night + 0.5: remBuckets.Night = remBuckets.Night - 30
                    Case "Deep":  finalBuckets.Deep = finalBuckets.Deep + 0.5:   remBuckets.Deep = remBuckets.Deep - 30
                End Select
            Next u

            Dim matchFound As Boolean: matchFound = False

            ' Case A: 2人介助
            If grpHelper = "2" Then
                Dim code2P As String
                code2P = FindMatchCode(masterData, "2", finalBuckets, False)

                If isForced2P Then
                    ' 重複ルールで強制2人介助になった場合: R列に2人介助コード、U列なし
                    proposedR = IIf(code2P <> "", code2P, "(NoMatch 2P)")
                    proposedU = ""
                Else
                    ' 通常の2人介助 (元々M列が"2"だった場合): R列に1人介助コード、U列に2人介助コード
                    Dim code1P As String
                    code1P = FindMatchCode(masterData, "1", finalBuckets, False)
                    proposedR = IIf(code1P <> "", code1P, "(NoMatch 1P)")
                    proposedU = IIf(code2P <> "", code2P, "(NoMatch 2P)")
                End If
                matchFound = True
            End If

            ' Case B: 通常（1人 or 1人分のマッチング）
            ' 【考え方】マスターに1コードで一致するものは1コードで請求。
            ' 増コードは「181分以上かつ時間跨ぎ（日中→夜間など）」のときのみ使用。
            If Not matchFound Then
                Dim strictCode As String
                strictCode = FindMatchCode(masterData, grpHelper, finalBuckets, False)
                If strictCode <> "" Then
                    proposedR = strictCode
                Else
                    ' 1コードで一致しない場合のみ、増コードの可否を判定
                    Dim totalRawMin As Long: totalRawMin = CLng(totalRaw)
                    Dim hasTimeCrossing As Boolean
                    hasTimeCrossing = (IIf(rawBuckets.Early > 0, 1, 0) + IIf(rawBuckets.Day > 0, 1, 0) + IIf(rawBuckets.Night > 0, 1, 0) + IIf(rawBuckets.Deep > 0, 1, 0)) >= 2
                    If totalRawMin >= 181 And hasTimeCrossing Then
                        Dim splitRes As Variant
                        splitRes = FindSplitMatch(masterData, grpHelper, finalBuckets)
                        If IsArray(splitRes) Then
                            proposedR = splitRes(0)
                            proposedU = splitRes(1)
                        Else
                            proposedR = "(NoMatch H" & grpHelper & ")"
                        End If
                    Else
                        proposedR = "(NoMatch H" & grpHelper & ")"
                    End If
                End If
            End If
        End If

        If existingR <> "" Then calcR(lastItemIdx) = existingR Else calcR(lastItemIdx) = proposedR
        If existingU <> "" Then calcU(lastItemIdx) = existingU Else calcU(lastItemIdx) = proposedU
    Next g

    ' --- 3. 結果出力 ---
    Dim cR As String, cU As String, recR As MasterRecord, recU As MasterRecord
    Dim fIdx As Long

    For i = 1 To numRows
        cR = calcR(i): cU = calcU(i)

        If cR <> "" And Left(cR, 1) <> "(" Then
            hadEntry = True
            outputData(i, 3) = cR
            If FindMasterRecord(masterData, cR, recR) Then
                outputData(i, 4) = recR.Name
                outputData(i, 5) = recR.Price
                ' P,Q および AD〜AI のみフラグを書き、R〜W と X〜AC は触らない
                For fIdx = 0 To 19
                    If (fIdx < 2 Or fIdx > 7) And (fIdx < 8 Or fIdx > 13) Then
                        If recR.Flags(fIdx) <> "" Then outputData(i, fIdx + 1) = recR.Flags(fIdx)
                    End If
                Next fIdx
            End If
        Else
            If cR <> "" Then outputData(i, 3) = cR
        End If

        If cU <> "" And Left(cU, 1) <> "(" Then
            hadEntry = True
            outputData(i, 6) = cU
            If FindMasterRecord(masterData, cU, recU) Then
                outputData(i, 7) = recU.Name
                outputData(i, 8) = recU.Price
                ' P,Q および AD〜AI のみフラグを書き、R〜W と X〜AC は触らない
                For fIdx = 0 To 19
                    If (fIdx < 2 Or fIdx > 7) And (fIdx < 8 Or fIdx > 13) Then
                        If recU.Flags(fIdx) <> "" Then outputData(i, fIdx + 1) = recU.Flags(fIdx)
                    End If
                Next fIdx
            End If
        Else
            If cU <> "" Then outputData(i, 6) = cU
        End If
    Next i

    ' X列〜AC列(24〜29列)は数式が入っているため書き込まない。P〜W と AD〜AI のみ書き戻す
    Dim outP2W As Variant, outAD2AI As Variant
    Dim r As Long, c As Long
    ReDim outP2W(1 To numRows, 1 To 8)   ' P〜W (8列)
    ReDim outAD2AI(1 To numRows, 1 To 6) ' AD〜AI (6列)
    For r = 1 To numRows
        For c = 1 To 8
            outP2W(r, c) = outputData(r, c)
        Next c
        For c = 1 To 6
            outAD2AI(r, c) = outputData(r, 14 + c)
        Next c
    Next r
    ws.Range(ws.Cells(startRow, 16), ws.Cells(startRow + numRows - 1, 23)).Value = outP2W
    ws.Range(ws.Cells(startRow, 30), ws.Cells(startRow + numRows - 1, 35)).Value = outAD2AI
    ProcessSheet = hadEntry
End Function

Function CalculateBuckets(startTime As Date, endTime As Date) As TimeBucket
    Dim result As TimeBucket
    Dim totalMins As Long
    totalMins = DateDiff("n", startTime, endTime)
    Dim cur As Date, h As Integer, i As Long
    cur = startTime
    For i = 0 To totalMins - 1
        h = Hour(cur)
        If h >= 6 And h < 8 Then
            result.Early = result.Early + 1
        ElseIf h >= 8 And h < 18 Then
            result.Day = result.Day + 1
        ElseIf h >= 18 And h < 22 Then
            result.Night = result.Night + 1
        Else
            result.Deep = result.Deep + 1
        End If
        cur = DateAdd("n", 1, cur)
    Next i
    CalculateBuckets = result
End Function

Function FindMatchCode(data() As MasterRecord, helpers As String, b As TimeBucket, isInc As Boolean) As String
    Dim i As Long
    Dim eps As Double: eps = 0.01
    For i = LBound(data) To UBound(data)
        If data(i).IsIncrease = isInc And data(i).helpers = helpers Then
            If Abs(data(i).Early - b.Early) < eps And _
               Abs(data(i).Day - b.Day) < eps And _
               Abs(data(i).Night - b.Night) < eps And _
               Abs(data(i).Deep - b.Deep) < eps Then
               FindMatchCode = data(i).code
               Exit Function
            End If
        End If
    Next i
    FindMatchCode = ""
End Function

Function FindSplitMatch(data() As MasterRecord, helpers As String, b As TimeBucket) As Variant
    Dim i As Long, j As Long
    Dim eps As Double: eps = 0.01
    Dim pList As Collection: Set pList = New Collection

    For i = LBound(data) To UBound(data)
        If Not data(i).IsIncrease And data(i).helpers = helpers Then
            If data(i).Early <= b.Early + eps And _
               data(i).Day <= b.Day + eps And _
               data(i).Night <= b.Night + eps And _
               data(i).Deep <= b.Deep + eps Then
                pList.Add i
            End If
        End If
    Next i
    If pList.Count = 0 Then Exit Function

    Dim sIdx() As Long: ReDim sIdx(1 To pList.Count)
    For i = 1 To pList.Count: sIdx(i) = pList.Item(i): Next i

    ' 合計時間の大きい順。同点の場合はコード番号の大きい順（例: 160185 を 160183 より優先）
    Dim t1 As Double, t2 As Double, tmp As Long
    Dim c1 As String, c2 As String
    For i = 1 To UBound(sIdx) - 1
        For j = i + 1 To UBound(sIdx)
            t1 = data(sIdx(i)).Early + data(sIdx(i)).Day + data(sIdx(i)).Night + data(sIdx(i)).Deep
            t2 = data(sIdx(j)).Early + data(sIdx(j)).Day + data(sIdx(j)).Night + data(sIdx(j)).Deep
            If t2 > t1 Then
                tmp = sIdx(i): sIdx(i) = sIdx(j): sIdx(j) = tmp
            ElseIf Abs(t2 - t1) < 0.01 Then
                c1 = data(sIdx(i)).code: c2 = data(sIdx(j)).code
                If c2 > c1 Then
                    tmp = sIdx(i): sIdx(i) = sIdx(j): sIdx(j) = tmp
                End If
            End If
        Next j
    Next i

    For i = 1 To UBound(sIdx)
        Dim idx As Long: idx = sIdx(i)
        Dim rE As Double, rD As Double, rN As Double, rDp As Double
        rE = Application.Max(0, b.Early - data(idx).Early)
        rD = Application.Max(0, b.Day - data(idx).Day)
        rN = Application.Max(0, b.Night - data(idx).Night)
        rDp = Application.Max(0, b.Deep - data(idx).Deep)
        Dim rTotal As Double: rTotal = rE + rD + rN + rDp
        If rTotal < eps Then GoTo NextCand

        For j = LBound(data) To UBound(data)
            If data(j).IsIncrease And data(j).helpers = helpers Then
                If Abs(data(j).Early - rE) < eps And _
                   Abs(data(j).Day - rD) < eps And _
                   Abs(data(j).Night - rN) < eps And _
                   Abs(data(j).Deep - rDp) < eps Then
                   FindSplitMatch = Array(data(idx).code, data(j).code)
                   Exit Function
                End If
            End If
        Next j
NextCand:
    Next i
    FindSplitMatch = Empty
End Function

Function FindMasterRecord(data() As MasterRecord, code As String, ByRef outRec As MasterRecord) As Boolean
    Dim i As Long
    For i = LBound(data) To UBound(data)
        If data(i).code = code Then
            outRec = data(i)
            FindMasterRecord = True
            Exit Function
        End If
    Next i
End Function

Function TryParseDate(strIn As Variant, ByRef outDate As Date) As Boolean
    If IsDate(strIn) Then
        outDate = strIn: TryParseDate = True
    ElseIf IsNumeric(strIn) Then
        Dim d As Long: d = CLng(strIn)
        If d >= 1 And d <= 31 Then
            outDate = DateSerial(2000, 1, d): TryParseDate = True
        End If
    End If
End Function

Function TryNormalizeTime(dBase As Date, tVal As Variant, ByRef outDate As Date) As Boolean
    Dim h As Integer, m As Integer
    If IsDate(tVal) Then
        h = Hour(tVal): m = Minute(tVal)
    ElseIf VarType(tVal) = vbString Then
        Dim p() As String: p = Split(tVal, ":")
        If UBound(p) >= 1 Then h = CInt(p(0)): m = CInt(p(1)) Else Exit Function
    ElseIf IsNumeric(tVal) Then
        Dim t As Date: t = CDate(tVal): h = Hour(t): m = Minute(t)
    Else
        Exit Function
    End If
    outDate = DateSerial(Year(dBase), Month(dBase), Day(dBase)) + TimeSerial(h, m, 0)
    TryNormalizeTime = True
End Function

Function IsSameDay(d1 As Date, d2 As Date) As Boolean
    IsSameDay = (Int(CDbl(d1)) = Int(CDbl(d2)))
End Function

Option Explicit
' ============================================================
' サービスコード集計マクロ（Mac対応版・Dictionary不使用）
' 要件: R16:W35のデータをE5で同一人物とみなし合算し、
'       45行目以降に R=コード, S=内容, T=単位数, U=回数 を出力
' 対象: シート名に「様」を含むシートのみ（例: 山田太郎様）
'
' 追加: 各「〇〇〇〇様」シートのV44（単位数合計）を、
'       受給者番号（E5）をキーにシート「集計」のA列と照合し、
'       対応する行のK列5行目以降に合計値を出力する。
' ============================================================

Private Const OUTPUT_START_ROW As Long = 45
Private Const SUMMARY_SHEET_NAME As String = "集計"
Private Const SUMMARY_START_ROW As Long = 5
Private Const SUMMARY_END_ROW As Long = 1000  ' 集計シートの検索・書き込み上限

Public Sub サービスコード集計()
    Dim ws As Worksheet
    Dim e5Val As String
    Dim groups As Collection      ' 各要素: Array(e5Val, Collection of シートIndex)
    Dim sheetIndices As Collection
    Dim firstSheetIndex As Long
    Dim i As Long, j As Long
    Dim codeKeys() As Variant
    Dim codeData() As Variant
    Dim nCodes As Long
    Dim keys() As Variant
    Dim arr As Variant
    Dim outWs As Worksheet
    Dim found As Boolean
    Dim idx As Long
    ' --- 集計シートへのV44単位数合計出力用 ---
    Dim wsSummary As Worksheet
    Dim totalUnits As Double
    Dim si As Variant
    Dim v44 As Variant
    Dim summaryRow As Long
    Dim summaryJukyu As String
    ' --- 処理①を実行したシート名一覧・スキップ判定用 ---
    Dim processedGroupIndices As Collection  ' 処理①を実行したグループの j
    Dim executedSheetNames As Collection    ' 処理①を実行したシート名
    Dim msg As String
    Dim nm As Variant

    On Error GoTo ErrHandler

    Set groups = New Collection
    Set processedGroupIndices = New Collection
    Set executedSheetNames = New Collection

    ' --- 1) シート名に「様」を含むシートのみ走査し、E5の値ごとにシートIndexをグループ化 ---
    For i = 1 To ThisWorkbook.Worksheets.Count
        Set ws = ThisWorkbook.Worksheets(i)
        If InStr(1, ws.Name, "様", vbBinaryCompare) = 0 Then GoTo NextSheet
        e5Val = ""
        If Not IsEmpty(ws.Range("E5").Value) Then e5Val = CStr(ws.Range("E5").Value)
        found = False
        For j = 1 To groups.Count
            If CStr(groups(j)(0)) = e5Val Then
                groups(j)(1).Add i
                found = True
                Exit For
            End If
        Next j
        If Not found Then
            Set sheetIndices = New Collection
            sheetIndices.Add i
            groups.Add Array(e5Val, sheetIndices)
        End If
NextSheet:
    Next i

    ' --- 2) 各E5グループについて、先頭シートのR45を確認。値があればスキップ。なければ集計し先頭シートに出力 ---
    For j = 1 To groups.Count
        Set sheetIndices = groups(j)(1)
        firstSheetIndex = GetMinFromCollection(sheetIndices)
        Set outWs = ThisWorkbook.Worksheets(firstSheetIndex)
        ' R45に値が入力されている場合はこの利用者様には何もしない
        If Not IsEmpty(outWs.Range("R45").Value) And Trim(CStr(outWs.Range("R45").Value)) <> "" Then
            GoTo NextGroup
        End If

        nCodes = 0
        ReDim codeKeys(0 To 0)
        ReDim codeData(0 To 0)

        For i = 1 To ThisWorkbook.Worksheets.Count
            If CollectionContains(sheetIndices, i) Then
                Set ws = ThisWorkbook.Worksheets(i)
                CollectRange ws, 16, 35, "R", "S", "T", codeKeys, codeData, nCodes
                CollectRange ws, 16, 35, "U", "V", "W", codeKeys, codeData, nCodes
            End If
        Next i

        keys = GetSortedCodeKeysArray(codeKeys, nCodes)
        ' nCodes > 0 のときだけR45以降にマクロで入力する（1行以上書き込む場合のみ「入力を行った」とみなす）
        If nCodes > 0 Then
            For i = LBound(keys) To UBound(keys)
                idx = FindCodeIndex(codeKeys, nCodes, keys(i))
                If idx >= 0 Then
                    arr = codeData(idx)
                    outWs.Cells(OUTPUT_START_ROW + i - LBound(keys), "R").Value = keys(i)
                    outWs.Cells(OUTPUT_START_ROW + i - LBound(keys), "S").Value = arr(0)
                    outWs.Cells(OUTPUT_START_ROW + i - LBound(keys), "T").Value = arr(1)
                    outWs.Cells(OUTPUT_START_ROW + i - LBound(keys), "U").Value = arr(2)
                End If
            Next i
            processedGroupIndices.Add j
            executedSheetNames.Add outWs.Name
        End If
NextGroup:
    Next j

    ' --- 3) シート「集計」のA列（受給者番号）をキーに、各E5グループのV44単位数合計をK列に出力 ---
    On Error Resume Next
    Set wsSummary = ThisWorkbook.Worksheets(SUMMARY_SHEET_NAME)
    On Error GoTo ErrHandler
    If Not wsSummary Is Nothing Then
        For j = 1 To groups.Count
            ' 処理①を実行したグループのみ集計シートに反映する
            If Not CollectionContains(processedGroupIndices, j) Then GoTo NextSummary
            e5Val = CStr(groups(j)(0))
            Set sheetIndices = groups(j)(1)
            totalUnits = 0
            For Each si In sheetIndices
                Set ws = ThisWorkbook.Worksheets(CLng(si))
                v44 = ws.Range("V44").Value
                If Not IsEmpty(v44) And IsNumeric(v44) Then totalUnits = totalUnits + CDbl(v44)
            Next si
            ' 集計シートのA列（5行目以降）で受給者番号が一致する行を探し、K列に書き込む
            For summaryRow = SUMMARY_START_ROW To SUMMARY_END_ROW
                summaryJukyu = ""
                If Not IsEmpty(wsSummary.Cells(summaryRow, "A").Value) Then summaryJukyu = Trim(CStr(wsSummary.Cells(summaryRow, "A").Value))
                If summaryJukyu = "" Then Exit For
                If summaryJukyu = Trim(e5Val) Then
                    wsSummary.Cells(summaryRow, "K").Value = totalUnits
                    Exit For
                End If
            Next summaryRow
NextSummary:
        Next j
    End If

    ' 完了メッセージ（入力を行ったシート名のみ表示）
    msg = "入力を行ったシート名：" & vbCrLf
    If executedSheetNames.Count > 0 Then
        For Each nm In executedSheetNames
            msg = msg & CStr(nm) & vbCrLf
        Next nm
    Else
        msg = msg & "（なし）"
    End If
    MsgBox msg, vbInformation
    Exit Sub

ErrHandler:
    MsgBox "エラーが発生しました: " & Err.Description, vbCritical
End Sub

' Collection に指定した値が含まれるか
Private Function CollectionContains(col As Collection, val As Long) As Boolean
    Dim v As Variant
    On Error Resume Next
    For Each v In col
        If CLng(v) = val Then CollectionContains = True: Exit Function
    Next v
    CollectionContains = False
End Function

' Collection 内の最小値（Long）を返す
Private Function GetMinFromCollection(col As Collection) As Long
    Dim v As Variant
    Dim first As Boolean
    first = True
    GetMinFromCollection = 0
    For Each v In col
        If first Then GetMinFromCollection = CLng(v): first = False
        If CLng(v) < GetMinFromCollection Then GetMinFromCollection = CLng(v)
    Next v
End Function

Private Sub CollectRange(ws As Worksheet, startRow As Long, endRow As Long, _
    colCode As String, colUnits As String, colContent As String, _
    ByRef codeKeys() As Variant, ByRef codeData() As Variant, ByRef nCodes As Long)
    Dim r As Long
    Dim code As Variant, units As Variant, content As Variant
    Dim key As String
    Dim idx As Long
    Dim arr As Variant

    For r = startRow To endRow
        code = ws.Range(colCode & r).Value
        If Not IsEmpty(code) And Trim(CStr(code)) <> "" Then
            units = ws.Range(colUnits & r).Value
            content = ws.Range(colContent & r).Value
            If IsEmpty(content) Then content = ""
            If IsEmpty(units) Then units = ""
            key = CStr(code)
            idx = FindCodeIndex(codeKeys, nCodes, key)
            If idx >= 0 Then
                arr = codeData(idx)
                arr(2) = arr(2) + 1
                codeData(idx) = arr
            Else
                nCodes = nCodes + 1
                ReDim Preserve codeKeys(0 To nCodes - 1)
                ReDim Preserve codeData(0 To nCodes - 1)
                codeKeys(nCodes - 1) = key
                codeData(nCodes - 1) = Array(content, units, 1)
            End If
        End If
    Next r
End Sub

' codeKeys(0..nCodes-1) から key のインデックスを返す。なければ -1
Private Function FindCodeIndex(codeKeys() As Variant, nCodes As Long, key As Variant) As Long
    Dim i As Long
    FindCodeIndex = -1
    For i = 0 To nCodes - 1
        If CStr(codeKeys(i)) = CStr(key) Then FindCodeIndex = i: Exit Function
    Next i
End Function

' コード配列を数値昇順でソートしたキー配列を返す
Private Function GetSortedCodeKeysArray(codeKeys() As Variant, nCodes As Long) As Variant
    Dim keys() As Variant
    Dim keysNum() As Double
    Dim i As Long, j As Long
    Dim tmp As Variant, tmpNum As Double

    If nCodes <= 0 Then
        GetSortedCodeKeysArray = Array()
        Exit Function
    End If

    ReDim keys(0 To nCodes - 1)
    ReDim keysNum(0 To nCodes - 1)
    For i = 0 To nCodes - 1
        keys(i) = codeKeys(i)
        If IsNumeric(codeKeys(i)) Then keysNum(i) = CDbl(codeKeys(i)) Else keysNum(i) = 0
    Next i

    For i = 0 To nCodes - 2
        For j = i + 1 To nCodes - 1
            If keysNum(j) < keysNum(i) Then
                tmp = keys(i): keys(i) = keys(j): keys(j) = tmp
                tmpNum = keysNum(i): keysNum(i) = keysNum(j): keysNum(j) = tmpNum
            End If
        Next j
    Next i
    GetSortedCodeKeysArray = keys
End Function

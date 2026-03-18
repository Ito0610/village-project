Option Explicit
' ============================================================
' 集計シート単位数出力マクロ
' 根拠: サービスコード集計_要件整理.md  §6 処理②
' 制約: Mac対応・Dictionary 不使用（Collection と配列のみ使用）
'
' 【§2 対象シート】
'   シート名に「様」を含むシートのみ対象（例: 山田太郎様、佐藤花子様）
'
' 【§3 同一人物の判定】
'   キー: 各シートの E5 の値（受給者番号）。E5 が同じシートは 1 グループで集計。
'
' 【§6 処理②：集計シートへの単位数合計の出力】
'   前提: シート名「集計」が存在する場合のみ実行。存在しない場合は行わない（エラーにしない）。
'   集計ロジック:
'     1. 各 E5 グループに属する「〇〇様」シートについて、V44 の値を合計する
'     2. V44 が空または数値でないセルは 0 として扱う
'     3. シート「集計」の A 列（5 行目以降）を上から検索し、受給者番号（E5 の値）と一致する行を探す
'     4. 一致した行の K 列に、上記で求めた V44 の合計値を書き込む
'   検索・書き込みの範囲: 検索開始行 5、検索終了行 1000、照合列 A、出力列 K
'   注意: A 列が空の行が現れた時点で、それ以降の検索は打ち切る（連続データを想定）
'   追加: K 列（5～1000 行）はいったんクリアしてから上書きする。
'
' 【§7 定数一覧】SUMMARY_SHEET_NAME, SUMMARY_START_ROW, SUMMARY_END_ROW
' 【§9 エラー時】ErrHandler で捕捉。集計シート取得時のみ On Error Resume Next で存在しない場合を無視。
' ============================================================

' §7 定数一覧（処理②で使用するもの）
Private Const SUMMARY_SHEET_NAME As String = "集計"   ' 単位数合計を書き込むシート名
Private Const SUMMARY_START_ROW As Long = 5           ' 集計シートの検索開始行
Private Const SUMMARY_END_ROW As Long = 1000          ' 集計シートの検索・書き込み上限行

Public Sub 集計シート単位数出力()
    Dim ws As Worksheet
    Dim wsSummary As Worksheet
    Dim e5Val As String
    Dim groups As Collection      ' 各要素: Array(e5Val, Collection of シートIndex)
    Dim sheetIndices As Collection
    Dim i As Long, j As Long
    Dim totalUnits As Double
    Dim si As Variant
    Dim v44 As Variant
    Dim summaryRow As Long
    Dim summaryJukyu As String
    Dim found As Boolean

    On Error GoTo ErrHandler

    ' §6 前提: シート「集計」が存在する場合のみ実行。§9 取得時のみ On Error Resume Next
    On Error Resume Next
    Set wsSummary = ThisWorkbook.Worksheets(SUMMARY_SHEET_NAME)
    On Error GoTo ErrHandler
    If wsSummary Is Nothing Then Exit Sub

    Set groups = New Collection

    ' §2・§3 シート名に「様」を含むシートのみ走査し、E5 の値ごとにシートをグループ化
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

    ' 集計シートの K 列（検索・書き込み範囲）をクリアしてから上書き
    wsSummary.Range(wsSummary.Cells(SUMMARY_START_ROW, "K"), wsSummary.Cells(SUMMARY_END_ROW, "K")).ClearContents

    ' §6 集計ロジック 1〜4: 各 E5 グループの V44 合計を求め、A 列照合で一致行の K 列に書き込む
    For j = 1 To groups.Count
        e5Val = CStr(groups(j)(0))
        Set sheetIndices = groups(j)(1)
        totalUnits = 0
        For Each si In sheetIndices
            Set ws = ThisWorkbook.Worksheets(CLng(si))
            v44 = ws.Range("V44").Value
            ' ロジック2: V44 が空または数値でないセルは 0 として扱う
            If Not IsEmpty(v44) And IsNumeric(v44) Then
                totalUnits = totalUnits + CDbl(v44)
            End If
        Next si

        ' ロジック3・4: A 列（5 行目以降）を上から検索し、受給者番号一致行の K 列に合計値を書き込む
        For summaryRow = SUMMARY_START_ROW To SUMMARY_END_ROW
            summaryJukyu = ""
            If Not IsEmpty(wsSummary.Cells(summaryRow, "A").Value) Then
                summaryJukyu = Trim(CStr(wsSummary.Cells(summaryRow, "A").Value))
            End If
            ' §6 注意: A 列が空の行が現れた時点で、それ以降の検索は打ち切る
            If summaryJukyu = "" Then Exit For
            If summaryJukyu = Trim(e5Val) Then
                wsSummary.Cells(summaryRow, "K").Value = totalUnits
                Exit For
            End If
        Next summaryRow
    Next j

    MsgBox "集計シートへの単位数合計の出力が完了しました。", vbInformation
    Exit Sub

ErrHandler:  ' §9 エラー時: ErrHandler で捕捉し、MsgBox で「エラーが発生しました: [説明]」を表示
    MsgBox "エラーが発生しました: " & Err.Description, vbCritical
End Sub

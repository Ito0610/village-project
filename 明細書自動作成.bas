Option Explicit
' ============================================================
' 明細書自動作成マクロ
'
' シート「集計」を参照（5行目以降）。K列が1以上の行について、
' シート「明細_原本」を複製し、1利用者につき1枚の明細書を作成する。
'
' ・複製シート名: 「明細_〇〇〇〇」（〇〇〇〇＝利用者名）
' ・挿入位置: 対象の「〇〇〇〇様」シートの頭（最初の様シートの直前）
' ・集計→明細: A→D7(受給者番号), B→D9(保護者氏名), C→D11(児童氏名)
'                B1→L6(年号), B2→N6(月), D→S3(利用者負担上限月額), J→S4(上限管理後の利用者負担額)
' ・各「〇〇〇〇様」シートの R45:U 以降 → 明細 Q16:T 以降へ転記
' ・14行を超えるレコードは30行目に行を挿入して記載（既存の明細シートに十分な行がある場合は挿入しない）。
'   挿入行にはA列の結合拡張、B〜C,D〜G,H〜J,K〜L,M〜Oの結合セルと式を引き継ぐ。
' ・印刷範囲: A1:O37 が漏れなく印刷できるように設定。
' ・既に「明細_〇〇〇〇」シートがある利用者は複製せずそのシートを上書き。上書き時は16〜29行目・Q〜T列
'   （および列追加分）の値をすべて削除してから転記する。
' ============================================================

Private Const SUMMARY_SHEET As String = "集計"
Private Const TEMPLATE_SHEET As String = "明細_原本"
Private Const SUMMARY_START_ROW As Long = 5
Private Const SUMMARY_MAX_ROW As Long = 1000
Private Const MEISAI_DATA_START_ROW As Long = 16      ' 明細のデータ開始行
Private Const MEISAI_FIRST_BLOCK_ROWS As Long = 14    ' 1ブロックの最大行数（16〜29）
Private Const MEISAI_INSERT_ROW As Long = 30         ' 2ブロック目開始行
Private Const SAMA_SOURCE_START_ROW As Long = 45     ' 様シートのデータ開始行 R45,S45,T45,U45
Private Const MEISAI_PRINT_END_ROW As Long = 37

Public Sub 明細書自動作成()
    Dim wb As Workbook
    Dim wsSummary As Worksheet
    Dim wsTemplate As Worksheet
    Dim wsMeisai As Worksheet
    Dim wsSama As Worksheet
    Dim r As Long
    Dim kVal As Variant
    Dim jukyu As String
    Dim userName As String
    Dim meisaiName As String
    Dim errPhase As String

    On Error GoTo ErrHandler
    Set wb = ThisWorkbook
    errPhase = "シート取得"

    On Error Resume Next
    Set wsSummary = wb.Worksheets(SUMMARY_SHEET)
    Set wsTemplate = wb.Worksheets(TEMPLATE_SHEET)
    On Error GoTo ErrHandler

    If wsSummary Is Nothing Then
        MsgBox "エラー: シート「" & SUMMARY_SHEET & "」が見つかりません。", vbCritical
        Exit Sub
    End If
    If wsTemplate Is Nothing Then
        MsgBox "エラー: シート「" & TEMPLATE_SHEET & "」が見つかりません。", vbCritical
        Exit Sub
    End If

    Application.ScreenUpdating = False

    For r = SUMMARY_START_ROW To SUMMARY_MAX_ROW
        jukyu = Trim(CStr(wsSummary.Cells(r, "A").Value))
        If jukyu = "" Then Exit For

        kVal = wsSummary.Cells(r, "K").Value
        If IsEmpty(kVal) Then GoTo NextSummaryRow
        If Not IsNumeric(kVal) Then GoTo NextSummaryRow
        If CDbl(kVal) < 1 Then GoTo NextSummaryRow

        errPhase = "利用者「" & jukyu & "」の明細書作成"
        Set wsSama = FindFirstSamaSheetByJukyu(wb, jukyu)
        If wsSama Is Nothing Then GoTo NextSummaryRow

        userName = GetUserNameFromSamaSheetName(wsSama.Name)
        meisaiName = "明細_" & userName

        ' 既に明細シートがある場合は複製せず上書き、なければ複製して挿入
        Set wsMeisai = Nothing
        On Error Resume Next
        Set wsMeisai = wb.Worksheets(meisaiName)
        On Error GoTo ErrHandler
        If wsMeisai Is Nothing Then
            errPhase = "明細シート複製（" & meisaiName & "）"
            Set wsMeisai = DuplicateMeisaiBeforeSama(wb, wsTemplate, wsSama, meisaiName)
        Else
            ' 既存シートを上書きするため、16〜29行・Q列以降のデータ領域を削除（列追加分も含む）
            errPhase = "明細シートのデータ領域クリア（" & meisaiName & "）"
            Call ClearMeisaiDataArea(wsMeisai)
        End If

        ' 集計の値を明細に書き込む
        errPhase = "集計データの転記"
        Call WriteSummaryToMeisai(wsSummary, wsMeisai, r)

        ' 年号・月を集計B1,B2から明細L6,N6へ
        wsMeisai.Range("L6").Value = wsSummary.Range("B1").Value
        wsMeisai.Range("N6").Value = wsSummary.Range("B2").Value

        ' 利用者負担上限月額・上限管理後の利用者負担額
        wsMeisai.Range("S3").Value = wsSummary.Cells(r, "D").Value
        wsMeisai.Range("S4").Value = wsSummary.Cells(r, "J").Value

        ' 該当受給者番号の全ての「様」シートから R45:U を集め、明細 Q16:T へ転記（14行超は30行目以降、15行以上で列追加）
        errPhase = "様シートからのサービスデータ転記"
        Call CopyServiceDataFromSamaToMeisai(wb, jukyu, wsMeisai)

        ' 印刷範囲の設定（A〜Oで値がある最終行まで）と1ページに収める設定
        Call SetMeisaiPrintArea(wsMeisai)

NextSummaryRow:
    Next r

    Application.ScreenUpdating = True
    MsgBox "明細書の自動作成が完了しました。", vbInformation
    Exit Sub

ErrHandler:
    Application.ScreenUpdating = True
    MsgBox "エラーが発生しました（" & errPhase & "）。" & vbCrLf & Err.Description, vbCritical
End Sub

' ============================================================
' 受給者番号に一致する最初の「〇〇〇〇様」シートを返す
' ============================================================
Private Function FindFirstSamaSheetByJukyu(wb As Workbook, ByVal jukyu As String) As Worksheet
    Dim ws As Worksheet
    Dim e5 As String
    Dim i As Long
    jukyu = Trim(jukyu)
    For i = 1 To wb.Worksheets.Count
        Set ws = wb.Worksheets(i)
        If InStr(1, ws.Name, "様", vbBinaryCompare) > 0 Then
            e5 = Trim(CStr(GetMergedCellValue(ws, 5, 5)))
            If e5 = jukyu Then
                Set FindFirstSamaSheetByJukyu = ws
                Exit Function
            End If
        End If
    Next i
    Set FindFirstSamaSheetByJukyu = Nothing
End Function

' シート名「山田太郎様」→「山田太郎」
Private Function GetUserNameFromSamaSheetName(ByVal sheetName As String) As String
    If Right(sheetName, 1) = "様" Then
        GetUserNameFromSamaSheetName = Left(sheetName, Len(sheetName) - 1)
    Else
        GetUserNameFromSamaSheetName = sheetName
    End If
End Function

' ============================================================
' 明細_原本を複製し、指定した「様」シートの直前に挿入。複製シートを返す。
' ============================================================
Private Function DuplicateMeisaiBeforeSama(wb As Workbook, wsTemplate As Worksheet, wsSama As Worksheet, newName As String) As Worksheet
    wsTemplate.Copy Before:=wsSama
    Set DuplicateMeisaiBeforeSama = wb.Worksheets(wsSama.Index - 1)
    DuplicateMeisaiBeforeSama.Name = newName
End Function

' ============================================================
' 既存明細シートのデータ領域を削除。16〜29行・Q列〜T列および列追加分（T列より右）、
' さらに30行目以降の同範囲をすべてクリアする。
' ============================================================
Private Sub ClearMeisaiDataArea(wsMeisai As Worksheet)
    Dim lastCol As Long
    Dim endCol As String
    Dim r As Long
    Dim c As Long
    ' 16〜29行・30行以降で使用されている最大列を取得（Q=17 以上）
    lastCol = 20
    For c = 17 To 50
        For r = MEISAI_DATA_START_ROW To 100
            If Not IsEmpty(wsMeisai.Cells(r, c).Value) Or Trim(CStr(wsMeisai.Cells(r, c).Value)) <> "" Then
                lastCol = c
                Exit For
            End If
        Next r
    Next c
    endCol = ColLetter(lastCol)
    wsMeisai.Range("Q" & MEISAI_DATA_START_ROW & ":" & endCol & "29").ClearContents
    wsMeisai.Range("Q" & MEISAI_INSERT_ROW & ":" & endCol & "500").ClearContents
End Sub

' ============================================================
' 集計の指定行から明細へ D7(受給者番号), D9(保護者氏名), D11(児童氏名) を書き込む
' ============================================================
Private Sub WriteSummaryToMeisai(wsSummary As Worksheet, wsMeisai As Worksheet, summaryRow As Long)
    SetMergedCellValue wsMeisai, 7, 4, wsSummary.Cells(summaryRow, "A").Value   ' D7 受給者番号
    SetMergedCellValue wsMeisai, 9, 4, wsSummary.Cells(summaryRow, "B").Value   ' D9 保護者氏名
    SetMergedCellValue wsMeisai, 11, 4, wsSummary.Cells(summaryRow, "C").Value  ' D11 児童氏名
End Sub

' ============================================================
' 受給者番号に一致する全「様」シートの R45:U を集め、明細 Q16:T に転記。
' 14行超は30行目以降に行挿入し、B:O の式を挿入行に反映する。印刷範囲は後で SetMeisaiPrintArea で設定。
' ============================================================
Private Sub CopyServiceDataFromSamaToMeisai(wb As Workbook, ByVal jukyu As String, wsMeisai As Worksheet)
    ' VBAのReDim Preserveは「最後の次元」のみ変更可能なため、list(列, 行) で定義する
    Dim list() As Variant   ' list(1, n)=コード, list(2,n)=内容, list(3,n)=単位数, list(4,n)=回数
    Dim nRec As Long
    Dim ws As Worksheet
    Dim e5 As String
    Dim i As Long, r As Long
    Dim srcRow As Long
    Dim insertCount As Long
    Dim currentCapacity As Long
    Dim existingExtraRows As Long
    Dim rowsToInsert As Long
    Dim lastDataRow As Long
    Dim sumRow As Long
    Dim extraRows As Long

    jukyu = Trim(jukyu)
    nRec = 0
    ReDim list(1 To 4, 1 To 1)

    ' 該当受給者番号の様シートを走査し、R45以降で R,S,T,U のいずれかがある行を収集
    For i = 1 To wb.Worksheets.Count
        Set ws = wb.Worksheets(i)
        If InStr(1, ws.Name, "様", vbBinaryCompare) = 0 Then GoTo NextWs
        e5 = Trim(CStr(GetMergedCellValue(ws, 5, 5)))
        If e5 <> jukyu Then GoTo NextWs

        For srcRow = SAMA_SOURCE_START_ROW To SAMA_SOURCE_START_ROW + 500
            If IsRowEmptyRSTU(ws, srcRow) Then Exit For
            nRec = nRec + 1
            ReDim Preserve list(1 To 4, 1 To nRec)
            list(1, nRec) = ws.Cells(srcRow, "R").Value
            list(2, nRec) = ws.Cells(srcRow, "S").Value
            list(3, nRec) = ws.Cells(srcRow, "T").Value
            list(4, nRec) = ws.Cells(srcRow, "U").Value
        Next srcRow
NextWs:
    Next i

    If nRec = 0 Then Exit Sub

    ' 明細の既存データ領域をクリア（Q16:T と 30行目以降）
    wsMeisai.Range("Q16:T" & (MEISAI_PRINT_END_ROW + 100)).ClearContents
    On Error Resume Next
    wsMeisai.Range("Q" & MEISAI_INSERT_ROW & ":T" & wsMeisai.Rows.Count).ClearContents
    On Error GoTo 0

    ' 1〜14件目: Q16:T29
    For i = 1 To nRec
        If i <= MEISAI_FIRST_BLOCK_ROWS Then
            r = MEISAI_DATA_START_ROW + i - 1
            wsMeisai.Cells(r, "Q").Value = list(1, i)
            wsMeisai.Cells(r, "R").Value = list(2, i)
            wsMeisai.Cells(r, "S").Value = list(3, i)
            wsMeisai.Cells(r, "T").Value = list(4, i)
        End If
    Next i

    ' 15件目以降: 明細シートの現在のデータ行数を検知し、足りない分だけ行を挿入してからQ〜Tに値を記載
    If nRec > MEISAI_FIRST_BLOCK_ROWS Then
        insertCount = nRec - MEISAI_FIRST_BLOCK_ROWS
        currentCapacity = GetMeisaiDataRowCapacity(wsMeisai)
        ' 30行目以降に既にある行数（枠のうち、14行を超えた分）
        existingExtraRows = currentCapacity - MEISAI_FIRST_BLOCK_ROWS
        If existingExtraRows < 0 Then existingExtraRows = 0
        ' 足りない分だけ挿入
        If insertCount > existingExtraRows Then
            rowsToInsert = insertCount - existingExtraRows
            wsMeisai.Range(wsMeisai.Rows(MEISAI_INSERT_ROW), wsMeisai.Rows(MEISAI_INSERT_ROW + rowsToInsert - 1)).Insert Shift:=xlDown
            Call ApplyMeisaiMergedCellsAndFormulas(wsMeisai, MEISAI_INSERT_ROW, rowsToInsert, insertCount)
        End If
        For i = 1 To insertCount
            r = MEISAI_INSERT_ROW + i - 1
            wsMeisai.Cells(r, "Q").Value = list(1, MEISAI_FIRST_BLOCK_ROWS + i)
            wsMeisai.Cells(r, "R").Value = list(2, MEISAI_FIRST_BLOCK_ROWS + i)
            wsMeisai.Cells(r, "S").Value = list(3, MEISAI_FIRST_BLOCK_ROWS + i)
            wsMeisai.Cells(r, "T").Value = list(4, MEISAI_FIRST_BLOCK_ROWS + i)
        Next i
    End If

    ' 合計行（M列）: 行を追加していない場合はM30固定。行を追加した場合は1行ずつ下にずらす（1行追加→M31、2行→M32…）
    extraRows = 0
    If nRec > MEISAI_FIRST_BLOCK_ROWS Then extraRows = nRec - MEISAI_FIRST_BLOCK_ROWS
    lastDataRow = 29 + extraRows
    sumRow = 30 + extraRows
    wsMeisai.Cells(sumRow, "M").Formula = "=SUM(M16:O" & lastDataRow & ")"
End Sub

' ============================================================
' 明細シートの「サービス費用の計算欄」のデータ行数を返す。
' A15の結合範囲（A15:A??）から判定。結合範囲が15行なら14、29行なら28を返す。
' （15行目はヘッダのためデータ行は16行目以降。データ行数 = 結合行数 - 1）
' 結合されていない場合は14（16〜29の1ブロック分）を返す。
' ============================================================
Private Function GetMeisaiDataRowCapacity(wsMeisai As Worksheet) As Long
    Dim rng As Range
    On Error Resume Next
    Set rng = wsMeisai.Range("A15")
    If rng Is Nothing Then
        GetMeisaiDataRowCapacity = MEISAI_FIRST_BLOCK_ROWS
        Exit Function
    End If
    If rng.MergeCells Then
        GetMeisaiDataRowCapacity = rng.MergeArea.Rows.Count - 1
        If GetMeisaiDataRowCapacity < MEISAI_FIRST_BLOCK_ROWS Then GetMeisaiDataRowCapacity = MEISAI_FIRST_BLOCK_ROWS
    Else
        GetMeisaiDataRowCapacity = MEISAI_FIRST_BLOCK_ROWS
    End If
    On Error GoTo 0
End Function

' ============================================================
' 挿入行（30行目以降）に結合セル・式・書式（罫線・配置等）を適用する。
' rowsToSetup: 今回セットアップする行数（新規挿入した行数）
' totalExtraRows: 30行目以降の総行数（A列結合の範囲用＝29+totalExtraRows まで拡張）
' ============================================================
Private Sub ApplyMeisaiMergedCellsAndFormulas(wsMeisai As Worksheet, startRow As Long, rowsToSetup As Long, totalExtraRows As Long)
    Dim lastRow As Long
    Dim r As Long
    Const FORMAT_REF_ROW As Long = 29   ' 書式の参照行（16〜29行目と同じ書式）

    lastRow = startRow + rowsToSetup - 1

    ' A列: 既存の結合を解除し、A15:A(29+totalExtraRows) に拡張（30行目以降の全データ行を含める）
    On Error Resume Next
    If wsMeisai.Range("A15").MergeCells Then wsMeisai.Range("A15").MergeArea.UnMerge
    On Error GoTo 0
    wsMeisai.Range("A15:A" & (29 + totalExtraRows)).Merge
    wsMeisai.Range("A15").Value = "サービス費用の計算欄"

    ' 今回挿入した行に B〜C, D〜G, H〜J, K〜L, M〜O の結合・式・書式を設定
    For r = startRow To lastRow
        ' B〜C 結合
        On Error Resume Next
        wsMeisai.Range(wsMeisai.Cells(r, 2), wsMeisai.Cells(r, 3)).UnMerge
        On Error GoTo 0
        wsMeisai.Range(wsMeisai.Cells(r, 2), wsMeisai.Cells(r, 3)).Merge
        wsMeisai.Cells(r, 2).Formula = "=IF(Q" & r & "="""", """", Q" & r & ")"
        wsMeisai.Range(wsMeisai.Cells(FORMAT_REF_ROW, 2), wsMeisai.Cells(FORMAT_REF_ROW, 3)).Copy
        wsMeisai.Range(wsMeisai.Cells(r, 2), wsMeisai.Cells(r, 3)).PasteSpecial Paste:=xlPasteFormats
        Application.CutCopyMode = False

        ' D〜G 結合
        On Error Resume Next
        wsMeisai.Range(wsMeisai.Cells(r, 4), wsMeisai.Cells(r, 7)).UnMerge
        On Error GoTo 0
        wsMeisai.Range(wsMeisai.Cells(r, 4), wsMeisai.Cells(r, 7)).Merge
        wsMeisai.Cells(r, 4).Formula = "=IF(R" & r & "="""", """", R" & r & ")"
        wsMeisai.Range(wsMeisai.Cells(FORMAT_REF_ROW, 4), wsMeisai.Cells(FORMAT_REF_ROW, 7)).Copy
        wsMeisai.Range(wsMeisai.Cells(r, 4), wsMeisai.Cells(r, 7)).PasteSpecial Paste:=xlPasteFormats
        Application.CutCopyMode = False

        ' H〜J 結合
        On Error Resume Next
        wsMeisai.Range(wsMeisai.Cells(r, 8), wsMeisai.Cells(r, 10)).UnMerge
        On Error GoTo 0
        wsMeisai.Range(wsMeisai.Cells(r, 8), wsMeisai.Cells(r, 10)).Merge
        wsMeisai.Cells(r, 8).Formula = "=IF(S" & r & "="""", """", S" & r & ")"
        wsMeisai.Range(wsMeisai.Cells(FORMAT_REF_ROW, 8), wsMeisai.Cells(FORMAT_REF_ROW, 10)).Copy
        wsMeisai.Range(wsMeisai.Cells(r, 8), wsMeisai.Cells(r, 10)).PasteSpecial Paste:=xlPasteFormats
        Application.CutCopyMode = False

        ' K〜L 結合
        On Error Resume Next
        wsMeisai.Range(wsMeisai.Cells(r, 11), wsMeisai.Cells(r, 12)).UnMerge
        On Error GoTo 0
        wsMeisai.Range(wsMeisai.Cells(r, 11), wsMeisai.Cells(r, 12)).Merge
        wsMeisai.Cells(r, 11).Formula = "=IF(T" & r & "="""", """", T" & r & ")"
        wsMeisai.Range(wsMeisai.Cells(FORMAT_REF_ROW, 11), wsMeisai.Cells(FORMAT_REF_ROW, 12)).Copy
        wsMeisai.Range(wsMeisai.Cells(r, 11), wsMeisai.Cells(r, 12)).PasteSpecial Paste:=xlPasteFormats
        Application.CutCopyMode = False

        ' M〜O 結合
        On Error Resume Next
        wsMeisai.Range(wsMeisai.Cells(r, 13), wsMeisai.Cells(r, 15)).UnMerge
        On Error GoTo 0
        wsMeisai.Range(wsMeisai.Cells(r, 13), wsMeisai.Cells(r, 15)).Merge
        wsMeisai.Cells(r, 13).Formula = "=IFERROR(H" & r & "*K" & r & ", 0)"
        wsMeisai.Range(wsMeisai.Cells(FORMAT_REF_ROW, 13), wsMeisai.Cells(FORMAT_REF_ROW, 15)).Copy
        wsMeisai.Range(wsMeisai.Cells(r, 13), wsMeisai.Cells(r, 15)).PasteSpecial Paste:=xlPasteFormats
        Application.CutCopyMode = False
    Next r
End Sub

' 様シートの指定行で R,S,T,U がすべて空かどうか
Private Function IsRowEmptyRSTU(ws As Worksheet, rowNum As Long) As Boolean
    Dim v As Variant
    v = ws.Cells(rowNum, "R").Value
    If Not IsEmpty(v) And Trim(CStr(v)) <> "" Then Exit Function
    v = ws.Cells(rowNum, "S").Value
    If Not IsEmpty(v) And Trim(CStr(v)) <> "" Then Exit Function
    v = ws.Cells(rowNum, "T").Value
    If Not IsEmpty(v) And Trim(CStr(v)) <> "" Then Exit Function
    v = ws.Cells(rowNum, "U").Value
    If Not IsEmpty(v) And Trim(CStr(v)) <> "" Then Exit Function
    IsRowEmptyRSTU = True
End Function

' ============================================================
' 明細シートの印刷範囲を設定。A列〜O列に値が入っている最終行までを含め、
' 常に1ページに収まるように拡大/縮小する。
' （37行のときは37行まで、1行増えたら38行まで、2行増えたら39行まで…を1ページで印刷）
' ============================================================
Private Sub SetMeisaiPrintArea(wsMeisai As Worksheet)
    Dim lastRow As Long
    Dim r As Long
    Dim c As Long
    Dim hasValue As Boolean

    lastRow = MEISAI_PRINT_END_ROW
    For r = 1 To 500
        hasValue = False
        For c = 1 To 15
            If Not IsEmpty(wsMeisai.Cells(r, c).Value) Or Trim(CStr(wsMeisai.Cells(r, c).Value)) <> "" Then
                hasValue = True
                Exit For
            End If
        Next c
        If hasValue Then lastRow = r
    Next r
    wsMeisai.PageSetup.PrintArea = "A1:O" & lastRow

    ' 印刷範囲を常に1ページに収める（行数が37でも38でも39でも1ページに拡大/縮小）
    With wsMeisai.PageSetup
        .Zoom = False
        .FitToPagesWide = 1
        .FitToPagesTall = 1
    End With
End Sub

' 列番号をアルファベットに（1=A, 27=AA など）
Private Function ColLetter(colNum As Long) As String
    Dim s As String
    Dim n As Long
    n = colNum
    ColLetter = ""
    Do While n > 0
        s = Chr(65 + ((n - 1) Mod 26)) & s
        n = Int((n - 1) / 26)
    Loop
    ColLetter = s
End Function

' ============================================================
' 結合セルヘルパー
' ============================================================
Private Function GetMergedCellValue(ws As Worksheet, rowNum As Long, colNum As Long) As Variant
    Dim rng As Range
    Set rng = ws.Cells(rowNum, colNum)
    If rng.MergeCells Then
        GetMergedCellValue = rng.MergeArea.Cells(1, 1).Value
    Else
        GetMergedCellValue = rng.Value
    End If
End Function

Private Sub SetMergedCellValue(ws As Worksheet, rowNum As Long, colNum As Long, ByVal val As Variant)
    Dim rng As Range
    Set rng = ws.Cells(rowNum, colNum)
    If rng.MergeCells Then
        rng.MergeArea.Cells(1, 1).Value = val
    Else
        rng.Value = val
    End If
End Sub

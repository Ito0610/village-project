Option Explicit

' ======================================================================================
' 各利用者様シートから通知書情報を取得
' 各利用者シートから集計シートへデータを収集するマクロ
'
' 対象シート: シート名に「様」を含むシート（〇〇〇〇様・〇〇〇〇様(2) など）
' 集計先: シート「集計」の5行目からデータを追加
'
' 集計シートの列:
'   A: 受給者番号（対象シート E5・結合セル）
'   B: 支給決定障害者（保護者）氏名（対象シート E9・結合セル）形式「名字　名前」（全角スペース）
'   C: 支給決定に係る児童氏名（対象シート J9・結合セル・空白可）
'   D: 利用者負担額（対象シート J5・結合セル・空白可）
'   E: 移動支援の時間数（対象シート J7・結合セル・半角で取得）
'   F: 通所支援の時間数（対象シート K8・半角で取得）
'   G: 社会参加の時間数 = E - F（計算値）
'
' 同一受給者番号は複数シートに存在する場合があるため、受給者番号ごとに1行のみ集計する。
' 集計結果は受給者番号順に5行目から並べる。
' ======================================================================================

' 収集用の1行分のデータ（可変長配列の要素として使う型は使わず、配列で列を表現）
Private Const COL_JUKYU As Long = 1   ' 受給者番号
Private Const COL_GUARDIAN As Long = 2 ' 保護者氏名
Private Const COL_JIDO As Long = 3     ' 児童氏名
Private Const COL_FUTAN As Long = 4   ' 利用者負担額
Private Const COL_IDO As Long = 5     ' 移動支援時間数
Private Const COL_TOSHO As Long = 6   ' 通所支援時間数
Private Const NUM_COLS As Long = 6    ' 収集列数（Gは後で計算）

' ======================================================================================
' メイン処理
' ======================================================================================
Public Sub 各利用者様シートから通知書情報を取得()
    Dim wb As Workbook
    Dim wsSummary As Worksheet
    Dim ws As Worksheet
    Dim list() As Variant   ' 1-based, 行×列(NUM_COLS)
    Dim listRows As Long
    Dim i As Long, j As Long
    Dim r As Long
    Dim jukyu As String
    Dim guardian As String, jido As String, futan As String, ido As String, tosho As String
    Dim alreadyAdded As Boolean
    Dim lastRow As Long
    Dim vE5 As Variant, vE9 As Variant, vJ9 As Variant, vJ5 As Variant, vJ7 As Variant, vK8 As Variant
    Dim idoNum As Double, toshoNum As Double, sankaNum As Double
    Dim errPhase As String

    errPhase = "初期化"
    Set wb = ThisWorkbook

    On Error Resume Next
    Set wsSummary = wb.Worksheets("集計")
    On Error GoTo 0
    If wsSummary Is Nothing Then
        MsgBox "エラー: シート「集計」が見つかりません。", vbCritical
        Exit Sub
    End If

    ' 収集用配列（最大でシート数と同じ行数、重複除外するので実際はそれ以下）
    listRows = 0
    ReDim list(1 To wb.Worksheets.Count, 1 To NUM_COLS)

    errPhase = "対象シートの走査"
    For Each ws In wb.Worksheets
        If ws.Name = "集計" Then GoTo NextSheet
        ' シート名が「〇〇〇〇様」形式（「様」を含む＝〇〇様・〇〇様(2) なども対象）
        If InStr(ws.Name, "様") = 0 Then GoTo NextSheet

        vE5 = GetMergedCellValue(ws, 5, 5)   ' E5 = 受給者番号
        jukyu = Trim(CStr(vE5))
        If jukyu = "" Then GoTo NextSheet

        ' 既に同じ受給者番号で追加済みならスキップ（重複排除）
        alreadyAdded = False
        For i = 1 To listRows
            If Trim(CStr(list(i, COL_JUKYU))) = jukyu Then
                alreadyAdded = True
                Exit For
            End If
        Next i
        If alreadyAdded Then GoTo NextSheet

        vE9 = GetMergedCellValue(ws, 9, 5)   ' E9 = 支給決定障害者（保護者）氏名
        vJ9 = GetMergedCellValue(ws, 9, 10) ' J9 = 支給決定に係る児童氏名
        vJ5 = GetMergedCellValue(ws, 5, 10) ' J5 = 利用者負担額
        vJ7 = GetMergedCellValue(ws, 7, 10) ' J7 = 移動支援の時間数
        vK8 = GetMergedCellValue(ws, 8, 11) ' K8 = 通所支援の時間数

        guardian = FormatGuardianName(Trim(CStr(vE9)))
        jido = FormatGuardianName(Trim(CStr(vJ9)))
        futan = Trim(CStr(vJ5))
        ido = ToHalfWidth(Trim(CStr(vJ7)))
        tosho = ToHalfWidth(Trim(CStr(vK8)))

        listRows = listRows + 1
        list(listRows, COL_JUKYU) = jukyu
        list(listRows, COL_GUARDIAN) = guardian
        list(listRows, COL_JIDO) = jido
        list(listRows, COL_FUTAN) = futan
        list(listRows, COL_IDO) = ido
        list(listRows, COL_TOSHO) = tosho
NextSheet:
    Next ws

    If listRows = 0 Then
        MsgBox "「〇〇〇〇様」形式のシートから取得できるデータがありませんでした。", vbInformation
        Exit Sub
    End If

    ' 受給者番号順にソート（バブルソート・文字列として比較）
    errPhase = "受給者番号でソート"
    Call SortListByJukyu(list, listRows)

    ' 集計シートの5行目以降を一旦クリア（A〜G列）
    errPhase = "集計シートのクリア"
    lastRow = wsSummary.Cells(wsSummary.Rows.Count, 1).End(xlUp).Row
    If lastRow >= 5 Then
        wsSummary.Range(wsSummary.Cells(5, 1), wsSummary.Cells(lastRow, 7)).ClearContents
    End If

    ' 集計シートの5行目から書き出し
    errPhase = "集計シートへの書き出し"
    Application.ScreenUpdating = False
    On Error GoTo ErrHandler

    For r = 1 To listRows
        ido = CStr(list(r, COL_IDO))
        tosho = CStr(list(r, COL_TOSHO))
        idoNum = ParseTimeNumber(ido)
        toshoNum = ParseTimeNumber(tosho)
        sankaNum = idoNum - toshoNum

        With wsSummary
            .Cells(4 + r, 1).Value = list(r, COL_JUKYU)                    ' A: 受給者番号
            .Cells(4 + r, 2).Value = list(r, COL_GUARDIAN)                ' B: 保護者氏名
            .Cells(4 + r, 3).Value = list(r, COL_JIDO)                     ' C: 児童氏名
            .Cells(4 + r, 4).Value = list(r, COL_FUTAN)                   ' D: 利用者負担額
            .Cells(4 + r, 5).Value = list(r, COL_IDO)                     ' E: 移動支援（半角）
            .Cells(4 + r, 6).Value = list(r, COL_TOSHO)                   ' F: 通所支援（半角）
            If sankaNum <> 0 Then
                .Cells(4 + r, 7).Value = FormatTimeNumber(sankaNum)       ' G: 社会参加
            Else
                .Cells(4 + r, 7).Value = ""
            End If
        End With
    Next r

    Application.ScreenUpdating = True
    MsgBox "集計シートに " & listRows & " 件のデータを反映しました（5行目から受給者番号順）。", vbInformation
    Exit Sub

ErrHandler:
    Application.ScreenUpdating = True
    MsgBox "エラーが発生しました（処理箇所: " & errPhase & "）。" & vbCrLf & Err.Description, vbCritical
End Sub

' ======================================================================================
' 結合セルの値を取得（結合範囲の左上セルの値を返す）
' ======================================================================================
Private Function GetMergedCellValue(ws As Worksheet, rowNum As Long, colNum As Long) As Variant
    Dim rng As Range
    Dim topLeft As Range
    Set rng = ws.Cells(rowNum, colNum)
    If rng.MergeCells Then
        Set topLeft = rng.MergeArea.Cells(1, 1)
        GetMergedCellValue = topLeft.Value
    Else
        GetMergedCellValue = rng.Value
    End If
End Function

' ======================================================================================
' 保護者・児童氏名の形式を「名字　名前」（全角スペース区切り）に統一
' ======================================================================================
Private Function FormatGuardianName(ByVal s As String) As String
    If Len(s) = 0 Then Exit Function
    ' 半角スペースを全角スペース（ChrW(12288)）に統一
    FormatGuardianName = Replace(s, " ", ChrW(12288))
End Function

' ======================================================================================
' 文字列を半角に変換（数字・英字・カナ・スペース）
' ======================================================================================
Private Function ToHalfWidth(ByVal s As String) As String
    Dim i As Long
    Dim c As String
    Dim code As Long
    Dim result As String
    result = ""
    For i = 1 To Len(s)
        c = Mid(s, i, 1)
        code = AscW(c)
        ' 全角数字 ０(FF10)〜９(FF19) → 0〜9
        If code >= 65296 And code <= 65305 Then
            result = result & Chr(code - 65296 + 48)
        ' 全角英大文字 Ａ(FF21)〜Ｚ(FF3A)
        ElseIf code >= 65313 And code <= 65338 Then
            result = result & Chr(code - 65313 + 65)
        ' 全角英小文字 ａ(FF41)〜ｚ(FF5A)
        ElseIf code >= 65345 And code <= 65370 Then
            result = result & Chr(code - 65345 + 97)
        ' 全角スペース
        ElseIf code = 12288 Then
            result = result & " "
        Else
            result = result & c
        End If
    Next i
    ToHalfWidth = result
End Function

' ======================================================================================
' 時間数文字列を数値に変換（空は0、小数点対応）
' ======================================================================================
Private Function ParseTimeNumber(ByVal s As String) As Double
    Dim d As Double
    If Trim(s) = "" Then
        ParseTimeNumber = 0
        Exit Function
    End If
    On Error Resume Next
    d = CDbl(Replace(Trim(s), ",", "."))
    If Err.Number <> 0 Then ParseTimeNumber = 0 Else ParseTimeNumber = d
    On Error GoTo 0
End Function

' ======================================================================================
' 時間数値を半角文字列で整形（小数点以下があればそのまま、整数なら整数表示）
' ======================================================================================
Private Function FormatTimeNumber(ByVal d As Double) As String
    If d = Int(d) Then
        FormatTimeNumber = CStr(CLng(d))
    Else
        FormatTimeNumber = CStr(d)
    End If
End Function

' ======================================================================================
' list を受給者番号（1列目）で昇順ソート
' ======================================================================================
Private Sub SortListByJukyu(ByRef list() As Variant, ByVal n As Long)
    Dim i As Long, j As Long
    Dim tmp As Variant
    Dim a As String, b As String
    For i = 1 To n - 1
        For j = i + 1 To n
            a = Trim(CStr(list(i, COL_JUKYU)))
            b = Trim(CStr(list(j, COL_JUKYU)))
            If CompareJukyu(b, a) < 0 Then
                ' list(j) と list(i) を入れ替え
                Dim col As Long
                For col = 1 To NUM_COLS
                    tmp = list(i, col)
                    list(i, col) = list(j, col)
                    list(j, col) = tmp
                Next col
            End If
        Next j
    Next i
End Sub

' 受給者番号の比較（数値として比較可能なら数値、そうでなければ文字列）
' a < b なら負、a = b なら 0、a > b なら正
Private Function CompareJukyu(ByVal a As String, ByVal b As String) As Long
    Dim na As Long, nb As Long
    a = Trim(a)
    b = Trim(b)
    On Error Resume Next
    na = CLng(a)
    nb = CLng(b)
    If Err.Number = 0 Then
        CompareJukyu = na - nb
    Else
        If a < b Then CompareJukyu = -1 Else If a > b Then CompareJukyu = 1 Else CompareJukyu = 0
    End If
    On Error GoTo 0
End Function

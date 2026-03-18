Option Explicit

' ======================================================================================
' コピペシート → 各利用者シート 転送マクロ
'
' シート「コピペシート」2行目以降のデータを、K列が「大田区」の行のみ
' 各利用者様シートの16〜35行目へ転記する。B列の値（例: 藤井舞純様）がそのままシート名。
'
' コピペシート: A=サービス提供者名, B=利用者様名, C=開始, D=終了, E=（空き）,
'               F=目的地, G=（空き）, H=日付(YYYY/M/D), I=目的コード,
'               J=派遣人数, K=請求
'
' 利用者シート: 16〜35行目がレコード欄（20件分）。21件目からは新シートを作成。
' 転記先列: A=日(のみ), C〜E=目的地(結合), F=目的コード,
'           G,H=開始・終了, J,K=開始・終了, M=派遣人数, P=サービス提供者名
' 新シート作成時: 元の「〇〇〇〇様」を複製し、上記列の16〜35行以外は全く同じ内容。
'                16〜35行の転記用エリアのみクリアして新規用とする。
'
' ・K列≠「大田区」→ A〜K を薄いグレーで塗る
' ・K列=「大田区」でシート不在等エラー → A〜K を赤で塗る
' ・K列=「大田区」で同一利用者内の時間帯重複（終了=開始は除く）→ C〜D を黄で塗る
' ・20行超はシート複製「〇〇〇〇様(2)」「〇〇〇〇様(3)」…
'
' ・コピペシートは並べ替えず、転記先の「〇〇〇〇様」シートには H列日付 → C列開始時間 の時間軸順で書き込む。
'   利用者ごとにコピペシートの該当行をH・Cでソートしてから転記するため、入力順に関係なくシート上は時間順になる。
' ======================================================================================

' ======================================================================================
' メイン処理
' ======================================================================================
Public Sub CopyPasteToUserSheets()
    ' 利用者シート: 16〜35行がデータエリア（20行）。20行以下でも20行超でも同じ処理で対応
    Dim dataStartRow As Long
    Dim dataEndRow As Long
    Dim maxRowsPerSheet As Long
    dataStartRow = 16
    dataEndRow = 35
    maxRowsPerSheet = 20

    Dim wb As Workbook
    Dim wsCopy As Worksheet
    Dim lastRow As Long
    Dim r As Long
    Dim billing As String
    Dim userName As String
    Dim wsUser As Worksheet
    Dim destRow As Long
    Dim destSheetBaseName As String
    Dim sheetSuffix As Long
    Dim writeRow As Long
    Dim providerNames As String
    Dim numStaff As Long
    Dim namesToWrite As String
    Dim dayOnly As Variant
    Dim timeOverlapRows() As Long
    Dim otaRows() As Long   ' 大田区の行番号のみ（重複検出の対象を絞って負荷軽減）
    Dim numOta As Long
    Dim nOverlap As Long
    Dim i As Long, j As Long
    Dim r2 As Long
    Dim c1 As Double, d1 As Double, c2 As Double, d2 As Double
    Dim hasError As Boolean
    Dim errPhase As String  ' エラー発生箇所の特定用
    Dim wsAfter As Worksheet  ' 複製シートを挿入する「直後の元になるシート」

    errPhase = "初期化"
    Set wb = ThisWorkbook

    On Error Resume Next
    Set wsCopy = wb.Worksheets("コピペシート")
    On Error GoTo 0

    If wsCopy Is Nothing Then
        MsgBox "エラー: シート「コピペシート」が見つかりません。", vbCritical
        Exit Sub
    End If

    lastRow = wsCopy.Cells(wsCopy.Rows.Count, 1).End(xlUp).Row
    If lastRow < 2 Then
        MsgBox "コピペシートにデータがありません（2行目以降を入力してください）。", vbInformation
        Exit Sub
    End If

    If MsgBox("コピペシートから各利用者シートへ転送しますか？", vbYesNo + vbQuestion, "転送確認") <> vbYes Then
        Exit Sub
    End If

    Application.ScreenUpdating = False
    On Error GoTo ErrHandler

    ' まず全行の色をリセット（塗りつぶし解除）してから判定
    errPhase = "色リセット"
    For r = 2 To lastRow
        wsCopy.Range(wsCopy.Cells(r, 1), wsCopy.Cells(r, 11)).Interior.Pattern = xlNone
    Next r

    ' 時間重複の黄色塗り: 次の3条件を満たすときだけ、その行のC列・D列を黄色にする
    '  (1) B列が同じ利用者  (2) H列の日付が同じ  (3) C列開始〜D列終了の時間帯がかぶる（終了=開始は重複としない）
    ' 大田区の行だけ対象。別の利用者・別の日は塗らない。
    errPhase = "時間重複検出"
    If lastRow >= 2 Then
        ReDim otaRows(1 To lastRow)
        numOta = 0
        For r = 2 To lastRow
            If Trim(CStr(wsCopy.Cells(r, 11).Value)) = "大田区" Then
                numOta = numOta + 1
                otaRows(numOta) = r
            End If
        Next r
        ReDim timeOverlapRows(1 To numOta)
        nOverlap = 0
        For i = 1 To numOta
            r = otaRows(i)
            userName = Trim(CStr(wsCopy.Cells(r, 2).Value))   ' B列＝利用者様名
            If userName = "" Then GoTo NextRowMain2
            c1 = TimeToSerial(wsCopy.Cells(r, 3).Value)       ' C列＝開始
            d1 = TimeToSerial(wsCopy.Cells(r, 4).Value)       ' D列＝終了
            For j = 1 To numOta
                If i = j Then GoTo NextJ2
                r2 = otaRows(j)
                ' (1) B列が同じ利用者どうしか
                If Trim(CStr(wsCopy.Cells(r2, 2).Value)) <> userName Then GoTo NextJ2
                ' (2) H列の日付が同じ日か（別の日なら時間がかぶっても重複にしない）
                If Not IsSameDate(wsCopy.Cells(r, 8).Value, wsCopy.Cells(r2, 8).Value) Then GoTo NextJ2
                c2 = TimeToSerial(wsCopy.Cells(r2, 3).Value) ' C列＝開始
                d2 = TimeToSerial(wsCopy.Cells(r2, 4).Value) ' D列＝終了
                ' (3) 時間帯がかぶっているか
                If IsOverlap(c1, d1, c2, d2) Then
                    nOverlap = nOverlap + 1
                    timeOverlapRows(nOverlap) = r
                    Exit For
                End If
NextJ2:
            Next j
NextRowMain2:
        Next i
    End If

    ' 同一利用者内で時間がかぶっている行の C列・D列 のみ黄色に塗る
    errPhase = "重複行の黄色塗り"
    If nOverlap > 0 Then
        For i = 1 To nOverlap
            With wsCopy.Range(wsCopy.Cells(timeOverlapRows(i), 3), wsCopy.Cells(timeOverlapRows(i), 4))
                .Interior.Color = RGB(255, 255, 0)
            End With
        Next i
    End If

    ' 色付けのみ（コピペシートは並べ替えない）
    errPhase = "色付け"
    For r = 2 To lastRow
        billing = Trim(CStr(wsCopy.Cells(r, 11).Value))
        If billing <> "大田区" Then
            With wsCopy.Range(wsCopy.Cells(r, 1), wsCopy.Cells(r, 11))
                .Interior.Color = RGB(220, 220, 220)
            End With
        Else
            userName = Trim(CStr(wsCopy.Cells(r, 2).Value))
            If userName = "" Then
                wsCopy.Range(wsCopy.Cells(r, 1), wsCopy.Cells(r, 11)).Interior.Color = RGB(255, 0, 0)
            Else
                On Error Resume Next
                Set wsUser = wb.Worksheets(userName)
                On Error GoTo 0
                If wsUser Is Nothing Then
                    wsCopy.Range(wsCopy.Cells(r, 1), wsCopy.Cells(r, 11)).Interior.Color = RGB(255, 0, 0)
                End If
            End If
        End If
    Next r

    ' 利用者ごとに転記対象行を集め、H列日付→C列開始時間の昇順でソートしてから転記する
    Dim userNames() As String
    Dim userCount() As Long
    Dim userRows() As Variant   ' userRows(u) = 行番号の配列（Long）
    Dim tmpRowList() As Long
    Dim n As Long
    Dim ii As Long
    Dim numUsers As Long
    Dim u As Long
    Dim foundUser As Boolean
    Dim sheetNameWithSuffix As String

    ReDim userNames(1 To lastRow)
    ReDim userCount(1 To lastRow)
    numUsers = 0

    errPhase = "利用者別に行を集計"
    For r = 2 To lastRow
        billing = Trim(CStr(wsCopy.Cells(r, 11).Value))
        If billing = "大田区" Then
            userName = Trim(CStr(wsCopy.Cells(r, 2).Value))
            If userName <> "" Then
                On Error Resume Next
                Set wsUser = wb.Worksheets(userName)
                On Error GoTo 0
                If Not wsUser Is Nothing Then
                    destSheetBaseName = wsUser.Name
                    foundUser = False
                    For u = 1 To numUsers
                        If userNames(u) = destSheetBaseName Then
                            userCount(u) = userCount(u) + 1
                            foundUser = True
                            Exit For
                        End If
                    Next u
                    If Not foundUser Then
                        numUsers = numUsers + 1
                        u = numUsers
                        userNames(u) = destSheetBaseName
                        userCount(u) = 1
                    End If
                End If
            End If
        End If
    Next r

    If numUsers = 0 Then GoTo NoTransfer

    ReDim userRows(1 To numUsers)
    For u = 1 To numUsers
        ReDim tmpRowList(1 To userCount(u))
        userRows(u) = tmpRowList
    Next u
    For u = 1 To numUsers
        userCount(u) = 0
    Next u

    errPhase = "利用者別に行番号を格納"
    For r = 2 To lastRow
        billing = Trim(CStr(wsCopy.Cells(r, 11).Value))
        If billing = "大田区" Then
            userName = Trim(CStr(wsCopy.Cells(r, 2).Value))
            If userName <> "" Then
                On Error Resume Next
                Set wsUser = wb.Worksheets(userName)
                On Error GoTo 0
                If Not wsUser Is Nothing Then
                    destSheetBaseName = wsUser.Name
                    For u = 1 To numUsers
                        If userNames(u) = destSheetBaseName Then
                            userCount(u) = userCount(u) + 1
                            userRows(u)(userCount(u)) = r
                            Exit For
                        End If
                    Next u
                End If
            End If
        End If
    Next r

    ' 利用者ごとの行番号配列を、H列日付→C列開始時間の昇順でソート（メイン内で直接入れ替え）
    errPhase = "利用者別にH列・C列で時間軸順にソート"
    Dim keyI As Double, keyJ As Double
    Dim tRow As Long
    For u = 1 To numUsers
        n = UBound(userRows(u))
        If n >= 2 Then
            For i = 1 To n - 1
                For j = i + 1 To n
                    keyI = GetRowSortKey(wsCopy, CLng(userRows(u)(i)))
                    keyJ = GetRowSortKey(wsCopy, CLng(userRows(u)(j)))
                    If keyI > keyJ Then
                        tRow = userRows(u)(i)
                        userRows(u)(i) = userRows(u)(j)
                        userRows(u)(j) = tRow
                    End If
                Next j
            Next i
        End If
    Next u

    errPhase = "転記ループ（時間軸順で各利用者シートへ書き込み）"
    For u = 1 To numUsers
        destSheetBaseName = userNames(u)
        On Error Resume Next
        Set wsUser = wb.Worksheets(destSheetBaseName)
        On Error GoTo 0
        If wsUser Is Nothing Then GoTo NextUser

        errPhase = "GetLastDataRow 利用者:" & destSheetBaseName
        destRow = GetLastDataRow(wsUser, 16, 35) + 1
        If destRow < 16 Then destRow = 16

        For ii = 1 To UBound(userRows(u))
            r = userRows(u)(ii)
            userName = Trim(CStr(wsCopy.Cells(r, 2).Value))

            If destRow > 35 Then
                sheetSuffix = Int((destRow - 16) / 20)
            Else
                sheetSuffix = 0
            End If
            If sheetSuffix = 0 Then
                On Error Resume Next
                Set wsUser = wb.Worksheets(destSheetBaseName)
                On Error GoTo 0
                writeRow = destRow
            Else
                sheetNameWithSuffix = destSheetBaseName & "(" & (sheetSuffix + 1) & ")"
                Set wsUser = Nothing
                On Error Resume Next
                Set wsUser = wb.Worksheets(sheetNameWithSuffix)
                On Error GoTo 0
                If wsUser Is Nothing Then
                    If sheetSuffix = 1 Then
                        Set wsAfter = wb.Worksheets(destSheetBaseName)
                    Else
                        Set wsAfter = wb.Worksheets(destSheetBaseName & "(" & sheetSuffix & ")")
                    End If
                    CreateUserSheetCopy wb, destSheetBaseName, sheetNameWithSuffix, 16, 35, wsUser, wsAfter
                End If
                writeRow = 16 + ((destRow - 16) Mod 20)
                If writeRow < 16 Then writeRow = 16
            End If

            If wsUser Is Nothing Then GoTo NextUser

            dayOnly = DayOnlyFromDate(wsCopy.Cells(r, 8).Value)
            providerNames = Trim(CStr(wsCopy.Cells(r, 1).Value))
            numStaff = CLng(Val(Trim(CStr(wsCopy.Cells(r, 10).Value))))
            If numStaff < 1 Then numStaff = 1
            namesToWrite = GetFirstNNames(providerNames, numStaff)

            errPhase = "転記実行 行r=" & r & " 利用者:" & userName & " writeRow=" & writeRow
            wsUser.Cells(writeRow, 1).Value = dayOnly
            On Error Resume Next
            wsUser.Range(wsUser.Cells(writeRow, 3), wsUser.Cells(writeRow, 5)).UnMerge
            On Error GoTo 0
            wsUser.Range(wsUser.Cells(writeRow, 3), wsUser.Cells(writeRow, 5)).Merge
            wsUser.Cells(writeRow, 3).Value = Trim(CStr(wsCopy.Cells(r, 6).Value))
            wsUser.Cells(writeRow, 6).Value = wsCopy.Cells(r, 9).Value
            wsUser.Cells(writeRow, 7).Value = wsCopy.Cells(r, 3).Value
            wsUser.Cells(writeRow, 8).Value = wsCopy.Cells(r, 4).Value
            wsUser.Cells(writeRow, 10).Value = wsCopy.Cells(r, 3).Value
            wsUser.Cells(writeRow, 11).Value = wsCopy.Cells(r, 4).Value
            wsUser.Cells(writeRow, 13).Value = wsCopy.Cells(r, 10).Value
            wsUser.Cells(writeRow, 16).Value = namesToWrite

            destRow = destRow + 1
        Next ii
NextUser:
    Next u

NoTransfer:

    Application.ScreenUpdating = True
    MsgBox "転送が完了しました。", vbInformation
    Exit Sub

ErrHandler:
    Application.ScreenUpdating = True
    MsgBox "エラーが発生しました: " & Err.Description & vbCrLf & vbCrLf & _
           "【特定用】" & vbCrLf & _
           "処理段階: " & errPhase & vbCrLf & _
           "コピペシートの行: " & r & vbCrLf & _
           "エラー番号: " & Err.Number, vbCritical, "エラー詳細"
End Sub

' ======================================================================================
' 日付セルから「日」のみ返す（2026年1月1日 → 1）
' ======================================================================================
Private Function DayOnlyFromDate(v As Variant) As Variant
    Dim d As Date
    On Error Resume Next
    If IsDate(v) Then
        d = CDate(v)
        DayOnlyFromDate = Day(d)
    Else
        DayOnlyFromDate = v
    End If
    On Error GoTo 0
End Function

' ======================================================================================
' 時刻をシリアル値に変換（比較用）
' ======================================================================================
Private Function TimeToSerial(v As Variant) As Double
    Dim s As String
    On Error Resume Next
    If IsEmpty(v) Or IsNull(v) Then
        TimeToSerial = 0
        Exit Function
    End If
    ' Excelの時刻シリアル（0〜1の小数）
    If VarType(v) = vbDouble Or VarType(v) = vbSingle Then
        If v >= 0 And v < 1 Then
            TimeToSerial = CDbl(v)
            Exit Function
        End If
        ' 日付+時刻のシリアルなら時刻部分のみ
        If v >= 1 Then
            TimeToSerial = v - Fix(v)
            Exit Function
        End If
    End If
    ' 日付型なら時刻部分のみ（日付+時刻のセル用）
    If IsDate(v) Then
        TimeToSerial = CDbl(CDate(v)) - Fix(CDbl(CDate(v)))
        Exit Function
    End If
    ' 文字列 "8:15" "9:30" "14:30" "17:00" など（24時間表記、全角コロンも可）
    s = Trim(CStr(v))
    s = Replace(s, ChrW(65306), ":")   ' 全角コロン（：）を半角に
    If s = "" Then
        TimeToSerial = 0
        Exit Function
    End If
    TimeToSerial = CDbl(TimeValue(s))
    On Error GoTo 0
End Function

' 行のソート用キー（H列日付のシリアル整数＋C列開始時間の端数 0〜1）。昇順＝早い日付・早い時刻が先（例: 28日9:30→28日14:30）
Private Function GetRowSortKey(wsCopy As Worksheet, r As Long) As Double
    Dim vDate As Variant, vTime As Variant
    Dim datePart As Double
    Dim timePart As Double
    On Error Resume Next
    vDate = wsCopy.Cells(r, 8).Value   ' H列＝日付
    vTime = wsCopy.Cells(r, 3).Value  ' C列＝開始時間
    If IsDate(vDate) Then datePart = Fix(CDbl(CDate(vDate))) Else datePart = 0
    timePart = TimeToSerial(vTime)
    If timePart < 0 Or timePart >= 1 Then timePart = 0
    GetRowSortKey = datePart + timePart
    On Error GoTo 0
End Function

' ======================================================================================
' 2つのセル値が同じ日付（日のみ比較、時刻は無視）か
' ======================================================================================
Private Function IsSameDate(v1 As Variant, v2 As Variant) As Boolean
    Dim d1 As Date, d2 As Date
    On Error Resume Next
    If Not IsDate(v1) Or Not IsDate(v2) Then
        IsSameDate = False
        Exit Function
    End If
    d1 = CDate(v1)
    d2 = CDate(v2)
    IsSameDate = (Year(d1) = Year(d2) And Month(d1) = Month(d2) And Day(d1) = Day(d2))
    On Error GoTo 0
End Function

' ======================================================================================
' 2つの時間帯が重なっているか（終了=開始の接続は重複としない）
' ======================================================================================
Private Function IsOverlap(c1 As Double, d1 As Double, c2 As Double, d2 As Double) As Boolean
    If c1 >= d2 Or c2 >= d1 Then
        IsOverlap = False
    Else
        IsOverlap = True
    End If
End Function

' ======================================================================================
' スペース区切り名前から先頭 n 人分を返す（半角・全角スペース対応）
' ======================================================================================
Private Function GetFirstNNames(providerNames As String, n As Long) As String
    Dim arr() As String
    Dim i As Long
    Dim k As Long
    Dim result As String
    Dim s As String
    If n <= 0 Then n = 1
    ' 全角スペースを半角に統一してから分割
    s = Replace(providerNames, ChrW(12288), " ")
    arr = Split(s, " ")
    result = ""
    k = 0
    For i = LBound(arr) To UBound(arr)
        If Trim(arr(i)) <> "" Then
            If k > 0 Then result = result & " "
            result = result & Trim(arr(i))
            k = k + 1
            If k >= n Then Exit For
        End If
    Next i
    GetFirstNNames = result
End Function

' ======================================================================================
' 利用者シートのデータ最終行（A列で判定、指定行範囲内）
Private Function GetLastDataRow(ws As Worksheet, startRow As Long, endRow As Long) As Long
' ======================================================================================
    Dim r As Long
    Dim last As Long
    last = 0
    For r = startRow To endRow
        If Trim(CStr(ws.Cells(r, 1).Value)) <> "" Then
            last = r
        End If
    Next r
    GetLastDataRow = last
End Function

' ======================================================================================
' 利用者シートの複製を作成（〇〇〇〇様 → 〇〇〇〇様(2) など）
' 元シートをそのまま複製し、転記用列（A,C〜E,F,G,H,J,K,M,P）の16〜35行のみクリアする。
' 並びは 〇〇様→(2)→(3)→(4)… になるよう、afterSheet の直後に挿入する。
' ======================================================================================
Private Sub CreateUserSheetCopy(wb As Workbook, baseName As String, newName As String, startRow As Long, endRow As Long, ByRef outWs As Worksheet, ByVal afterSheet As Worksheet)
    Dim wsTemplate As Worksheet
    Dim r As Long
    Dim colIdx As Long
    Dim colNum As Long
    Dim cell As Range
    Dim mergeTopLeft As Range
    ' 転記用列のみクリア（A,C,D,E,F,G,H,J,K,M,P = 1,3,4,5,6,7,8,10,11,13,16）
    Dim colList(0 To 10) As Long

    On Error Resume Next
    Set wsTemplate = wb.Worksheets(baseName)
    On Error GoTo 0
    If wsTemplate Is Nothing Then Exit Sub
    If wb.Worksheets.Count < 1 Then Exit Sub
    If afterSheet Is Nothing Then Exit Sub

    ' 元シートを複製。afterSheet の直後に挿入して 〇〇様→(2)→(3)→(4)… の並びにする
    wsTemplate.Copy After:=afterSheet
    Set outWs = wb.Worksheets(afterSheet.Index + 1)
    If outWs Is Nothing Then Exit Sub
    outWs.Name = newName

    colList(0) = 1: colList(1) = 3: colList(2) = 4: colList(3) = 5: colList(4) = 6
    colList(5) = 7: colList(6) = 8: colList(7) = 10: colList(8) = 11: colList(9) = 13: colList(10) = 16

    On Error Resume Next
    For r = startRow To endRow
        For colIdx = 0 To 10
            colNum = colList(colIdx)
            Set cell = outWs.Cells(r, colNum)
            If Not cell Is Nothing Then
                If cell.MergeCells Then
                    Set mergeTopLeft = cell.MergeArea.Cells(1, 1)
                    If cell.Row = mergeTopLeft.Row And cell.Column = mergeTopLeft.Column Then
                        mergeTopLeft.Value = ""
                    End If
                Else
                    cell.ClearContents
                End If
            End If
        Next colIdx
    Next r
    On Error GoTo 0
End Sub

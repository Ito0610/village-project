Option Explicit

' ======================================================================================
' マクロ名: 新規利用者様追加・時間数変更
'
' シート「新規追加&変更」の内容を元に、
' ・受給者番号(C2)が集計シートにない → 新規追加
'     - 同じ名前「〇〇〇〇様」のシートがあり、そのE5(受給者番号)が今回と一致: 実績_原本は複製せず既存シートに書き込む
'     - 同じ名前だがE5が異なる（別の利用者）: 「〇〇〇〇様(2)」「〇〇〇〇様(3)」…で新規シートを作成
'     - 同じ名前のシートがない場合: 「〇〇〇〇様」で実績_原本を複製し、E5(受給者番号)順の位置に挿入
' ・受給者番号(C2)が集計シートにある → 変更（集計と対象「〇〇〇〇様」シートの記載あり項目のみ上書き）
'
' 新規追加&変更シートの構成:
'   C2: 受給者番号
'   C4: 支給決定障害者(保護者)氏名・苗字  D4: 名前
'   C6: 支給決定に係る児童氏名・名字     D6: 名前（空白可）
'   C8: 利用者負担額（空白可）
'   C10: 移動支援の時間数  C11: 通所支援の時間数
'
' 対象シート（実績_原本 / 〇〇〇〇様）の書き込み先:
'   E5: 受給者番号  E9: 保護者氏名  J9: 児童氏名
'   J5: 利用者負担額  J7: 移動支援  K8: 通所支援
'
' 集計シートG列: 社会参加の時間数（必ず追加）
'   = 移動支援の時間数(E列) - 通所支援の時間数(F列) で計算
'
' その他仕様:
'   - 受給者番号(C2)は10桁の半角数字であること。それ以外はエラーで終了
'   - 利用者負担額(C8)が未記入の場合は自動的に0を設定
'   - マクロ完了後、新規追加&変更の入力セル(C2,C4,D4,C6,D6,C8,C10,C11)をクリア
' ======================================================================================

' マクロの実行はこのプロシージャ名で行います（日本語名は環境により赤字エラーになるため英名にしています）
Public Sub NewUserAddOrChange()
    Dim wb As Workbook
    Dim wsInput As Worksheet    ' 新規追加&変更
    Dim wsSummary As Worksheet ' 集計
    Dim wsTemplate As Worksheet ' 実績_原本
    Dim wsTarget As Worksheet   ' 対象 〇〇〇〇様
    Dim jukyu As String
    Dim lastSummaryRow As Long
    Dim summaryRow As Long      ' 集計で該当する行（変更時）
    Dim isChange As Boolean    ' True=変更, False=新規追加
    Dim errPhase As String
    Dim guardianName As String
    Dim sheetNameBase As String
    Dim newSheetName As String
    Dim suffix As Long
    Dim r As Long

    errPhase = "初期化"
    Set wb = ThisWorkbook

    On Error Resume Next
    Set wsInput = wb.Worksheets("新規追加&変更")
    On Error GoTo 0
    If wsInput Is Nothing Then
        MsgBox "エラー: シート「新規追加&変更」が見つかりません。", vbCritical
        Exit Sub
    End If

    jukyu = ToHalfWidth(Trim(CStr(wsInput.Cells(2, 3).Value)))  ' C2 = 受給者番号（半角に統一）
    If jukyu = "" Then
        MsgBox "エラー: シート「新規追加&変更」のC2に受給者番号を入力してください。", vbCritical
        Exit Sub
    End If
    If Not IsValidJukyu(jukyu) Then
        MsgBox "エラー: 受給者番号は10桁の半角数字で入力してください。" & vbCrLf & "入力値: " & jukyu, vbCritical
        Exit Sub
    End If

    On Error Resume Next
    Set wsSummary = wb.Worksheets("集計")
    On Error GoTo 0
    If wsSummary Is Nothing Then
        MsgBox "エラー: シート「集計」が見つかりません。", vbCritical
        Exit Sub
    End If

    lastSummaryRow = wsSummary.Cells(wsSummary.Rows.Count, 1).End(xlUp).Row
    If lastSummaryRow < 5 Then lastSummaryRow = 4

    ' 集計のA列(5行目〜)で受給者番号を検索
    summaryRow = 0
    For r = 5 To lastSummaryRow
        If Trim(CStr(wsSummary.Cells(r, 1).Value)) = jukyu Then
            summaryRow = r
            Exit For
        End If
    Next r

    isChange = (summaryRow >= 5)

    If isChange Then
        ' ========== 変更処理 ==========
        errPhase = "変更時・対象シートの検索"
        Set wsTarget = FindSheetByJukyu(wb, jukyu)
        If wsTarget Is Nothing Then
            MsgBox "エラー: 受給者番号「" & jukyu & "」に該当する「〇〇〇〇様」形式のシートが見つかりません。", vbCritical
            Exit Sub
        End If

        Application.ScreenUpdating = False
        On Error GoTo ErrHandler
        errPhase = "変更時・上書き"
        Call ApplyChangeFromInput(wsInput, wsTarget, wsSummary, summaryRow)
        Application.ScreenUpdating = True
        Call ClearInputSheet(wsInput)

        MsgBox "【変更が完了しました】" & vbCrLf & vbCrLf & _
               "受給者番号: " & jukyu & vbCrLf & _
               "シート「集計」と「" & wsTarget.Name & "」の、入力のあった項目を上書きしました。", vbInformation, "処理完了"
    Else
        ' ========== 新規追加処理 ==========
        On Error Resume Next
        Set wsTemplate = wb.Worksheets("実績_原本")
        On Error GoTo 0
        If wsTemplate Is Nothing Then
            MsgBox "エラー: シート「実績_原本」が見つかりません。新規追加にはこのシートが必要です。", vbCritical
            Exit Sub
        End If

        guardianName = BuildGuardianName( _
            Trim(CStr(wsInput.Cells(4, 3).Value)), _
            Trim(CStr(wsInput.Cells(4, 4).Value)))
        If guardianName = "" Then
            MsgBox "エラー: 新規追加の場合はC4・D4に支給決定障害者(保護者)氏名を入力してください。", vbCritical
            Exit Sub
        End If

        sheetNameBase = guardianName & "様"
        On Error Resume Next
        Set wsTarget = wb.Worksheets(sheetNameBase)
        On Error GoTo 0

        Application.ScreenUpdating = False
        On Error GoTo ErrHandler

        ' 同じ名前のシートがあり、かつそのシートのE5(受給者番号)が今回の受給者番号と一致 → 同一利用者、既存シートに書き込む
        If Not wsTarget Is Nothing Then
            If Trim(CStr(GetMergedCellValue(wsTarget, 5, 5))) = jukyu Then
                errPhase = "新規追加・既存シートへの入力"
                Call WriteInputToTargetSheet(wsInput, wsTarget)
                errPhase = "新規追加・集計シートへの行挿入"
                Call InsertSummaryRowInOrder(wsSummary, wsInput, wsTarget)
                newSheetName = sheetNameBase
            Else
                ' 同じ名前だがE5が違う → 別の利用者。〇〇〇〇様(2)、(3)... で新規シート作成
                newSheetName = GetUniqueSheetName(wb, sheetNameBase)
                errPhase = "新規追加・シート複製（受給者番号順に挿入）"
                Call DuplicateTemplateInJukyuOrder(wb, wsTemplate, jukyu, wsTarget)
                wsTarget.Name = newSheetName
                errPhase = "新規追加・対象シートへの入力"
                Call WriteInputToTargetSheet(wsInput, wsTarget)
                errPhase = "新規追加・集計シートへの行挿入"
                Call InsertSummaryRowInOrder(wsSummary, wsInput, wsTarget)
            End If
        Else
            ' 同じ名前のシートがない → 実績_原本を複製し、受給者番号順の位置に挿入
            newSheetName = sheetNameBase
            errPhase = "新規追加・シート複製（受給者番号順に挿入）"
            Call DuplicateTemplateInJukyuOrder(wb, wsTemplate, jukyu, wsTarget)
            wsTarget.Name = sheetNameBase

            errPhase = "新規追加・対象シートへの入力"
            Call WriteInputToTargetSheet(wsInput, wsTarget)

            errPhase = "新規追加・集計シートへの行挿入"
            Call InsertSummaryRowInOrder(wsSummary, wsInput, wsTarget)
        End If

        Application.ScreenUpdating = True
        Call ClearInputSheet(wsInput)

        MsgBox "【新規追加が完了しました】" & vbCrLf & vbCrLf & _
               "受給者番号: " & jukyu & vbCrLf & _
               "シート「" & newSheetName & "」にデータを反映し、集計シートに受給者番号順で1行追加しました。", vbInformation, "処理完了"
    End If
    Exit Sub

ErrHandler:
    Application.ScreenUpdating = True
    MsgBox "エラーが発生しました（処理箇所: " & errPhase & "）。" & vbCrLf & Err.Description, vbCritical
End Sub

' ======================================================================================
' 実績_原本を複製し、全「〇〇〇〇様」シートのE5(受給者番号)順になる位置に挿入する
' 複製したシートを outWs に返す
' ======================================================================================
Private Sub DuplicateTemplateInJukyuOrder(wb As Workbook, wsTemplate As Worksheet, ByVal jukyu As String, ByRef outWs As Worksheet)
    Dim wsBefore As Worksheet  ' このシートの直前に挿入（受給者番号が自分より大きい最初の「様」シート）
    Dim wsAfter As Worksheet   ' このシートの直後に挿入（全「様」のうち受給者番号が自分より小さい最後のシート）
    Dim ws As Worksheet
    Dim i As Long
    Dim e5 As String
    Dim e5After As String
    jukyu = Trim(jukyu)

    ' タブ順で、E5 > jukyu となる最初の「様」シートを探す → その直前に挿入
    Set wsBefore = Nothing
    For i = 1 To wb.Worksheets.Count
        Set ws = wb.Worksheets(i)
        If InStr(ws.Name, "様") > 0 Then
            e5 = Trim(CStr(GetMergedCellValue(ws, 5, 5)))
            If e5 <> "" And CompareJukyu(e5, jukyu) > 0 Then
                Set wsBefore = ws
                Exit For
            End If
        End If
    Next i

    If Not wsBefore Is Nothing Then
        wsTemplate.Copy Before:=wsBefore
        Set outWs = wb.Worksheets(wsBefore.Index - 1)
    Else
        ' 自分より大きい受給者番号の「様」シートがない → E5がjukyuより小さいもののうち最大のシートの直後に挿入
        Set wsAfter = Nothing
        e5After = ""
        For i = 1 To wb.Worksheets.Count
            Set ws = wb.Worksheets(i)
            If InStr(ws.Name, "様") > 0 Then
                e5 = Trim(CStr(GetMergedCellValue(ws, 5, 5)))
                If e5 <> "" And CompareJukyu(e5, jukyu) < 0 Then
                    If wsAfter Is Nothing Or CompareJukyu(e5, e5After) > 0 Then
                        Set wsAfter = ws
                        e5After = e5
                    End If
                End If
            End If
        Next i
        If Not wsAfter Is Nothing Then
            wsTemplate.Copy After:=wsAfter
            Set outWs = wb.Worksheets(wsAfter.Index + 1)
        Else
            ' 「様」シートが一つもない、または全て空 → 実績_原本の直後に挿入
            wsTemplate.Copy After:=wsTemplate
            Set outWs = wb.Worksheets(wsTemplate.Index + 1)
        End If
    End If
End Sub

' ======================================================================================
' 受給者番号が10桁の半角数字かどうか判定
' ======================================================================================
Private Function IsValidJukyu(ByVal s As String) As Boolean
    Dim i As Long
    Dim c As String
    If Len(s) <> 10 Then Exit Function
    For i = 1 To 10
        c = Mid(s, i, 1)
        If c < "0" Or c > "9" Then Exit Function
    Next i
    IsValidJukyu = True
End Function

' ======================================================================================
' 新規追加&変更シートの入力セル（C2,C4,D4,C6,D6,C8,C10,C11）をクリア
' ======================================================================================
Private Sub ClearInputSheet(wsInput As Worksheet)
    wsInput.Cells(2, 3).ClearContents   ' C2 受給者番号
    wsInput.Cells(4, 3).ClearContents   ' C4 保護者氏名・苗字
    wsInput.Cells(4, 4).ClearContents   ' D4 保護者氏名・名前
    wsInput.Cells(6, 3).ClearContents   ' C6 児童氏名・名字
    wsInput.Cells(6, 4).ClearContents   ' D6 児童氏名・名前
    wsInput.Cells(8, 3).ClearContents   ' C8 利用者負担額
    wsInput.Cells(10, 3).ClearContents  ' C10 移動支援の時間数
    wsInput.Cells(11, 3).ClearContents  ' C11 通所支援の時間数
End Sub

' ======================================================================================
' 受給者番号(E5)が一致する「〇〇〇〇様」シートを返す
' ======================================================================================
Private Function FindSheetByJukyu(wb As Workbook, ByVal jukyu As String) As Worksheet
    Dim ws As Worksheet
    Dim v As Variant
    jukyu = Trim(jukyu)
    For Each ws In wb.Worksheets
        If InStr(ws.Name, "様") > 0 Then
            v = GetMergedCellValue(ws, 5, 5)
            If Trim(CStr(v)) = jukyu Then
                Set FindSheetByJukyu = ws
                Exit Function
            End If
        End If
    Next ws
    Set FindSheetByJukyu = Nothing
End Function

' ======================================================================================
' 変更時: 新規追加&変更の「記載があった項目だけ」対象シートと集計を上書き
' ======================================================================================
Private Sub ApplyChangeFromInput(wsInput As Worksheet, wsTarget As Worksheet, wsSummary As Worksheet, summaryRow As Long)
    Dim v As String
    Dim guardian As String, jido As String, futan As String, ido As String, tosho As String
    Dim idoNum As Double, toshoNum As Double, sankaNum As Double

    ' 受給者番号は変更では更新しない（C2は必ずあるが、集計・シートのキーなので上書きしない想定）
    ' 保護者氏名 C4, D4
    v = Trim(CStr(wsInput.Cells(4, 3).Value)) & Trim(CStr(wsInput.Cells(4, 4).Value))
    If Len(Replace(v, ChrW(12288), "")) > 0 Then
        guardian = BuildGuardianName(Trim(CStr(wsInput.Cells(4, 3).Value)), Trim(CStr(wsInput.Cells(4, 4).Value)))
        SetMergedCellValue wsTarget, 9, 5, guardian
        wsSummary.Cells(summaryRow, 2).Value = guardian
    End If

    ' 児童氏名 C6, D6（どちらかでもあれば更新）
    v = Trim(CStr(wsInput.Cells(6, 3).Value)) & Trim(CStr(wsInput.Cells(6, 4).Value))
    If Len(Replace(v, ChrW(12288), "")) > 0 Then
        jido = BuildGuardianName(Trim(CStr(wsInput.Cells(6, 3).Value)), Trim(CStr(wsInput.Cells(6, 4).Value)))
        SetMergedCellValue wsTarget, 9, 10, jido
        wsSummary.Cells(summaryRow, 3).Value = jido
    End If

    ' 利用者負担額 C8（記入がない場合は0を設定）
    v = Trim(CStr(wsInput.Cells(8, 3).Value))
    If v = "" Then v = "0"
    SetMergedCellValue wsTarget, 5, 10, v
    wsSummary.Cells(summaryRow, 4).Value = v

    ' 移動支援 C10
    v = ToHalfWidth(Trim(CStr(wsInput.Cells(10, 3).Value)))
    If v <> "" Then
        SetMergedCellValue wsTarget, 7, 10, v
        wsSummary.Cells(summaryRow, 5).Value = v
    End If

    ' 通所支援 C11
    v = ToHalfWidth(Trim(CStr(wsInput.Cells(11, 3).Value)))
    If v <> "" Then
        wsTarget.Cells(8, 11).Value = v
        wsSummary.Cells(summaryRow, 6).Value = v
    End If

    ' 集計G列: 社会参加の時間数 = 移動支援(E列) - 通所支援(F列) で常に再計算
    ido = Trim(CStr(wsSummary.Cells(summaryRow, 5).Value))
    tosho = Trim(CStr(wsSummary.Cells(summaryRow, 6).Value))
    idoNum = ParseTimeNumber(ToHalfWidth(ido))
    toshoNum = ParseTimeNumber(ToHalfWidth(tosho))
    sankaNum = idoNum - toshoNum
    If sankaNum <> 0 Then
        wsSummary.Cells(summaryRow, 7).Value = FormatTimeNumber(sankaNum)
    Else
        wsSummary.Cells(summaryRow, 7).Value = ""
    End If
End Sub

' ======================================================================================
' 新規追加時: 新規追加&変更の内容を対象シート（複製済み）に書き込む
' ======================================================================================
Private Sub WriteInputToTargetSheet(wsInput As Worksheet, wsTarget As Worksheet)
    Dim jukyu As String
    Dim guardian As String, jido As String, futan As String, ido As String, tosho As String
    jukyu = Trim(CStr(wsInput.Cells(2, 3).Value))
    guardian = BuildGuardianName(Trim(CStr(wsInput.Cells(4, 3).Value)), Trim(CStr(wsInput.Cells(4, 4).Value)))
    jido = BuildGuardianName(Trim(CStr(wsInput.Cells(6, 3).Value)), Trim(CStr(wsInput.Cells(6, 4).Value)))
    futan = Trim(CStr(wsInput.Cells(8, 3).Value))
    If futan = "" Then futan = "0"
    ido = ToHalfWidth(Trim(CStr(wsInput.Cells(10, 3).Value)))
    tosho = ToHalfWidth(Trim(CStr(wsInput.Cells(11, 3).Value)))

    SetMergedCellValue wsTarget, 5, 5, jukyu
    SetMergedCellValue wsTarget, 9, 5, guardian
    SetMergedCellValue wsTarget, 9, 10, jido
    SetMergedCellValue wsTarget, 5, 10, futan
    SetMergedCellValue wsTarget, 7, 10, ido
    wsTarget.Cells(8, 11).Value = tosho
End Sub

' ======================================================================================
' 集計シートに1行を「受給者番号順」で挿入し、新規分のデータを書き込む
' ======================================================================================
Private Sub InsertSummaryRowInOrder(wsSummary As Worksheet, wsInput As Worksheet, wsTarget As Worksheet)
    Dim lastRow As Long
    Dim list() As Variant
    Dim listRows As Long
    Dim newRow(1 To 7) As Variant
    Dim insertAt As Long
    Dim r As Long
    Dim jukyu As String
    Dim guardian As String, jido As String, futan As String, ido As String, tosho As String
    Dim idoNum As Double, toshoNum As Double, sankaNum As Double

    jukyu = Trim(CStr(wsInput.Cells(2, 3).Value))
    guardian = BuildGuardianName(Trim(CStr(wsInput.Cells(4, 3).Value)), Trim(CStr(wsInput.Cells(4, 4).Value)))
    jido = BuildGuardianName(Trim(CStr(wsInput.Cells(6, 3).Value)), Trim(CStr(wsInput.Cells(6, 4).Value)))
    futan = Trim(CStr(wsInput.Cells(8, 3).Value))
    If futan = "" Then futan = "0"
    ido = ToHalfWidth(Trim(CStr(wsInput.Cells(10, 3).Value)))
    tosho = ToHalfWidth(Trim(CStr(wsInput.Cells(11, 3).Value)))
    idoNum = ParseTimeNumber(ido)
    toshoNum = ParseTimeNumber(tosho)
    ' G列 社会参加の時間数 = 移動支援(E) - 通所支援(F)
    sankaNum = idoNum - toshoNum

    newRow(1) = jukyu
    newRow(2) = guardian
    newRow(3) = jido
    newRow(4) = futan
    newRow(5) = ido
    newRow(6) = tosho
    If sankaNum <> 0 Then newRow(7) = FormatTimeNumber(sankaNum) Else newRow(7) = ""

    lastRow = wsSummary.Cells(wsSummary.Rows.Count, 1).End(xlUp).Row
    If lastRow < 5 Then lastRow = 4

    listRows = lastRow - 4
    If listRows <= 0 Then
        wsSummary.Cells(5, 1).Value = newRow(1)
        wsSummary.Cells(5, 2).Value = newRow(2)
        wsSummary.Cells(5, 3).Value = newRow(3)
        wsSummary.Cells(5, 4).Value = newRow(4)
        wsSummary.Cells(5, 5).Value = newRow(5)
        wsSummary.Cells(5, 6).Value = newRow(6)
        wsSummary.Cells(5, 7).Value = newRow(7)
        Exit Sub
    End If

    ReDim list(1 To listRows + 1, 1 To 7)
    For r = 1 To listRows
        list(r, 1) = wsSummary.Cells(4 + r, 1).Value
        list(r, 2) = wsSummary.Cells(4 + r, 2).Value
        list(r, 3) = wsSummary.Cells(4 + r, 3).Value
        list(r, 4) = wsSummary.Cells(4 + r, 4).Value
        list(r, 5) = wsSummary.Cells(4 + r, 5).Value
        list(r, 6) = wsSummary.Cells(4 + r, 6).Value
        list(r, 7) = wsSummary.Cells(4 + r, 7).Value
    Next r

    ' 挿入位置を受給者番号順で決定
    insertAt = listRows + 1
    For r = 1 To listRows
        If CompareJukyu(jukyu, Trim(CStr(list(r, 1)))) < 0 Then
            insertAt = r
            Exit For
        End If
    Next r

    ' insertAt の位置に新行を挿入（後ろをずらす）
    For r = listRows To insertAt Step -1
        list(r + 1, 1) = list(r, 1)
        list(r + 1, 2) = list(r, 2)
        list(r + 1, 3) = list(r, 3)
        list(r + 1, 4) = list(r, 4)
        list(r + 1, 5) = list(r, 5)
        list(r + 1, 6) = list(r, 6)
        list(r + 1, 7) = list(r, 7)
    Next r
    list(insertAt, 1) = newRow(1)
    list(insertAt, 2) = newRow(2)
    list(insertAt, 3) = newRow(3)
    list(insertAt, 4) = newRow(4)
    list(insertAt, 5) = newRow(5)
    list(insertAt, 6) = newRow(6)
    list(insertAt, 7) = newRow(7)

    ' 集計シート 5行目〜を上書き
    For r = 1 To listRows + 1
        wsSummary.Cells(4 + r, 1).Value = list(r, 1)
        wsSummary.Cells(4 + r, 2).Value = list(r, 2)
        wsSummary.Cells(4 + r, 3).Value = list(r, 3)
        wsSummary.Cells(4 + r, 4).Value = list(r, 4)
        wsSummary.Cells(4 + r, 5).Value = list(r, 5)
        wsSummary.Cells(4 + r, 6).Value = list(r, 6)
        wsSummary.Cells(4 + r, 7).Value = list(r, 7)
    Next r
End Sub

' ======================================================================================
' 保護者氏名を「名字　名前」（全角スペース）で結合
' ======================================================================================
Private Function BuildGuardianName(ByVal myoji As String, ByVal namae As String) As String
    myoji = Trim(myoji)
    namae = Trim(namae)
    If myoji = "" And namae = "" Then
        BuildGuardianName = ""
    ElseIf namae = "" Then
        BuildGuardianName = Replace(myoji, " ", ChrW(12288))
    ElseIf myoji = "" Then
        BuildGuardianName = Replace(namae, " ", ChrW(12288))
    Else
        BuildGuardianName = Replace(myoji, " ", ChrW(12288)) & ChrW(12288) & Replace(namae, " ", ChrW(12288))
    End If
End Function

' ======================================================================================
' シート名が既に存在する場合は "〇〇様(2)", "〇〇様(3)" などを返す
' （同じ利用者名で受給者番号が異なる場合に使用）
' ======================================================================================
Private Function GetUniqueSheetName(wb As Workbook, ByVal baseName As String) As String
    Dim s As String
    Dim n As Long
    On Error Resume Next
    If wb.Worksheets(baseName) Is Nothing Then
        GetUniqueSheetName = baseName
        Exit Function
    End If
    n = 2
    Do
        s = baseName & "(" & n & ")"
        If wb.Worksheets(s) Is Nothing Then
            GetUniqueSheetName = s
            Exit Function
        End If
        n = n + 1
    Loop
    On Error GoTo 0
    GetUniqueSheetName = baseName & "(" & n & ")"
End Function

' ======================================================================================
' 結合セルの値を取得（結合範囲の左上セル）
' ======================================================================================
Private Function GetMergedCellValue(ws As Worksheet, rowNum As Long, colNum As Long) As Variant
    Dim rng As Range
    Set rng = ws.Cells(rowNum, colNum)
    If rng.MergeCells Then
        GetMergedCellValue = rng.MergeArea.Cells(1, 1).Value
    Else
        GetMergedCellValue = rng.Value
    End If
End Function

' ======================================================================================
' 結合セルに書き込み（結合範囲の左上に書き込む）
' ======================================================================================
Private Sub SetMergedCellValue(ws As Worksheet, rowNum As Long, colNum As Long, ByVal val As Variant)
    Dim rng As Range
    Set rng = ws.Cells(rowNum, colNum)
    If rng.MergeCells Then
        rng.MergeArea.Cells(1, 1).Value = val
    Else
        rng.Value = val
    End If
End Sub

' ======================================================================================
' 文字列を半角に変換
' ======================================================================================
Private Function ToHalfWidth(ByVal s As String) As String
    Dim i As Long, c As String, code As Long, result As String
    result = ""
    For i = 1 To Len(s)
        c = Mid(s, i, 1)
        code = AscW(c)
        If code >= 65296 And code <= 65305 Then
            result = result & Chr(code - 65296 + 48)
        ElseIf code >= 65313 And code <= 65338 Then
            result = result & Chr(code - 65313 + 65)
        ElseIf code >= 65345 And code <= 65370 Then
            result = result & Chr(code - 65345 + 97)
        ElseIf code = 12288 Then
            result = result & " "
        Else
            result = result & c
        End If
    Next i
    ToHalfWidth = result
End Function

Private Function ParseTimeNumber(ByVal s As String) As Double
    Dim d As Double
    If Trim(s) = "" Then ParseTimeNumber = 0: Exit Function
    On Error Resume Next
    d = CDbl(Replace(Trim(s), ",", "."))
    If Err.Number <> 0 Then ParseTimeNumber = 0 Else ParseTimeNumber = d
    On Error GoTo 0
End Function

Private Function FormatTimeNumber(ByVal d As Double) As String
    If d = Int(d) Then FormatTimeNumber = CStr(CLng(d)) Else FormatTimeNumber = CStr(d)
End Function

' 受給者番号比較: a < b → 負, a = b → 0, a > b → 正
Private Function CompareJukyu(ByVal a As String, ByVal b As String) As Long
    Dim na As Long, nb As Long
    a = Trim(a): b = Trim(b)
    On Error Resume Next
    na = CLng(a): nb = CLng(b)
    If Err.Number = 0 Then
        CompareJukyu = na - nb
    Else
        If a < b Then CompareJukyu = -1 Else If a > b Then CompareJukyu = 1 Else CompareJukyu = 0
    End If
    On Error GoTo 0
End Function

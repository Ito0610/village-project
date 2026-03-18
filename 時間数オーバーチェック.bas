Option Explicit

' ======================================================================================
' マクロ名：時間数オーバーチェック
'
' 各「〇〇〇〇様」シートを参照し、シート「集計」のH列・I列に目的コード別の
' 合計時間数を入力し、F列・G列の許容時間と比較してオーバー時にH・I列を赤塗りする。
'
' 集計シート:
'   A列: 受給者番号（5行目〜）
'   F列: 通所支援の時間数（許容）→ H列の実績と比較
'   G列: 社会参加の時間数（許容）→ I列の実績と比較
'   H列: 目的コードFの合計時間数（本マクロで算出・記入）
'   I列: 目的コードAの合計時間数（本マクロで算出・記入）
'
' 各利用者様シート:
'   E5: 受給者番号（結合セル対応）
'   F列: 目的コード（半角 F または A）
'   L列: 算定時間数（0.5, 1.0, 2.5 など）
'   M列: 派遣人数（1 または 2）→ 実時間 = L×M
'
' 同一受給者番号は複数シートに存在する場合があるため、全「様」シートを
' 集計したうえで受給者番号ごとにF合計・A合計を集計シートのH・Iに反映する。
'
' オーバーチェック:
'   H > F のとき H列のセルを赤塗り（Fが空の場合は0とみなす）。F・H両方空はスキップ。
'   I > G のとき I列のセルを赤塗り（Gが空の場合は0とみなす）。G・I両方空はスキップ。
' ======================================================================================

' データ開始行（利用者様シートのF,L,Mのデータが始まる行）
Private Const DATA_START_ROW As Long = 16
' 目的コード（半角）
Private Const PURPOSE_F As String = "F"
Private Const PURPOSE_A As String = "A"

' ======================================================================================
' メイン処理
' ======================================================================================
Public Sub 時間数オーバーチェック()
    Dim wb As Workbook
    Dim wsSummary As Worksheet
    Dim ws As Worksheet
    Dim jukyu As String
    Dim lastDataRow As Long
    Dim r As Long, dr As Long
    Dim vF As Variant, vL As Variant, vM As Variant
    Dim codeF As String
    Dim timeL As Double, numM As Double, addTime As Double
    Dim lastSummaryRow As Long
    Dim sumF As Double, sumA As Double
    Dim vColF As Variant, vColG As Variant, vColH As Variant, vColI As Variant
    Dim valF As Double, valG As Double, valH As Double, valI As Double
    Dim errPhase As String
    Dim sheetJukyu As String

    errPhase = "初期化"
    Set wb = ThisWorkbook

    On Error Resume Next
    Set wsSummary = wb.Worksheets("集計")
    On Error GoTo 0
    If wsSummary Is Nothing Then
        MsgBox "エラー: シート「集計」が見つかりません。", vbCritical
        Exit Sub
    End If

    ' --- 集計シートの5行目〜最終行を取得 ---
    lastSummaryRow = wsSummary.Cells(wsSummary.Rows.Count, 1).End(xlUp).Row
    If lastSummaryRow < 5 Then
        MsgBox "集計シートにデータがありません（5行目以降に受給者番号を入力してください）。", vbInformation
        Exit Sub
    End If

    Application.ScreenUpdating = False
    On Error GoTo ErrHandler

    ' --- 集計シートの行ごとに、同じ受給者番号の「様」シートを回してF合計・A合計を算出 ---
    errPhase = "集計シートへの反映とオーバーチェック"
    For r = 5 To lastSummaryRow
        jukyu = Trim(CStr(wsSummary.Cells(r, 1).Value))

        sumF = 0
        sumA = 0
        ' 同じ受給者番号の「様」シートをすべて走査して合算
        For Each ws In wb.Worksheets
            If ws.Name = "集計" Then GoTo NextSheetLoop
            If InStr(ws.Name, "様") = 0 Then GoTo NextSheetLoop

            sheetJukyu = Trim(CStr(GetMergedCellValue(ws, 5, 5)))
            If sheetJukyu <> jukyu Then GoTo NextSheetLoop

            lastDataRow = GetLastDataRow(ws)
            If lastDataRow < DATA_START_ROW Then GoTo NextSheetLoop

            For dr = DATA_START_ROW To lastDataRow
                vF = ws.Cells(dr, 6).Value   ' F列 = 目的コード
                vL = ws.Cells(dr, 12).Value  ' L列 = 算定時間数
                vM = ws.Cells(dr, 13).Value  ' M列 = 派遣人数

                codeF = ToHalfWidth(Trim(CStr(vF)))
                If codeF <> PURPOSE_F And codeF <> PURPOSE_A Then GoTo NextDataRow

                timeL = ParseTimeNumber(CStr(vL))
                numM = ParseTimeNumber(CStr(vM))
                If numM <= 0 Then numM = 1
                addTime = timeL * numM

                If codeF = PURPOSE_F Then
                    sumF = sumF + addTime
                Else
                    sumA = sumA + addTime
                End If
NextDataRow:
            Next dr
NextSheetLoop:
        Next ws

        With wsSummary
            .Cells(r, 8).Value = IIf(sumF = 0, "", FormatTimeNumber(sumF))  ' H列
            .Cells(r, 9).Value = IIf(sumA = 0, "", FormatTimeNumber(sumA))  ' I列

            ' 一度塗りを解除
            .Cells(r, 8).Interior.ColorIndex = xlNone
            .Cells(r, 9).Interior.ColorIndex = xlNone

            vColF = .Cells(r, 6).Value
            vColG = .Cells(r, 7).Value
            vColH = .Cells(r, 8).Value
            vColI = .Cells(r, 9).Value

            ' F列とH列の比較（両方空ならスキップ）
            If Not (IsEmptyOrBlank(vColF) And IsEmptyOrBlank(vColH)) Then
                valF = ParseTimeNumber(CStr(vColF))
                valH = ParseTimeNumber(CStr(vColH))
                If valH > valF Then
                    .Cells(r, 8).Interior.Color = RGB(255, 0, 0)
                End If
            End If

            ' G列とI列の比較（両方空ならスキップ）
            If Not (IsEmptyOrBlank(vColG) And IsEmptyOrBlank(vColI)) Then
                valG = ParseTimeNumber(CStr(vColG))
                valI = ParseTimeNumber(CStr(vColI))
                If valI > valG Then
                    .Cells(r, 9).Interior.Color = RGB(255, 0, 0)
                End If
            End If
        End With
    Next r

    Application.ScreenUpdating = True
    MsgBox "時間数オーバーチェックを完了しました。" & vbCrLf & _
           "集計シートのH列（目的F合計）・I列（目的A合計）を記入し、F・G列を超えた行は赤で塗りました。", vbInformation
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
    Set rng = ws.Cells(rowNum, colNum)
    If rng.MergeCells Then
        GetMergedCellValue = rng.MergeArea.Cells(1, 1).Value
    Else
        GetMergedCellValue = rng.Value
    End If
End Function

' ======================================================================================
' 文字列を半角に変換（目的コードの比較用）
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
' 時間数値を半角文字列で整形
' ======================================================================================
Private Function FormatTimeNumber(ByVal d As Double) As String
    If d = Int(d) Then
        FormatTimeNumber = CStr(CLng(d))
    Else
        FormatTimeNumber = CStr(d)
    End If
End Function

' ======================================================================================
' セルが空または空白かどうか
' ======================================================================================
Private Function IsEmptyOrBlank(v As Variant) As Boolean
    If IsEmpty(v) Then
        IsEmptyOrBlank = True
        Exit Function
    End If
    If IsNull(v) Then
        IsEmptyOrBlank = True
        Exit Function
    End If
    IsEmptyOrBlank = (Trim(CStr(v)) = "")
End Function

' ======================================================================================
' シートのF列・L列・M列のうち、データが入っている最終行を返す
' ======================================================================================
Private Function GetLastDataRow(ws As Worksheet) As Long
    Dim lastF As Long, lastL As Long, lastM As Long
    lastF = ws.Cells(ws.Rows.Count, 6).End(xlUp).Row
    lastL = ws.Cells(ws.Rows.Count, 12).End(xlUp).Row
    lastM = ws.Cells(ws.Rows.Count, 13).End(xlUp).Row
    GetLastDataRow = Application.WorksheetFunction.Max(lastF, lastL, lastM)
    If GetLastDataRow < DATA_START_ROW Then GetLastDataRow = DATA_START_ROW - 1
End Function

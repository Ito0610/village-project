Attribute VB_Name = "ChangeAndResetMacro"
Option Explicit

' ======================================================================================
' 変更＆リセットマクロ
'
' 処理1: シート「月初変更＆リセット」のC2→対象シートのD3、C3→対象シートのF3 を入力
' 処理2: 対象シートの A,C,F,G,H,J,K,M,P 列および R,S,T,U,V,W 列の16〜35行目の値を削除
'        対象シート: シート名に「様」が含まれるシート
'
' 処理3: 「明細_」で始まるシート全て → S3,S4 の値削除、Q〜T列の16〜500行目の値削除
' 処理4: 「様」が含まれるシート全て → R〜U列の45〜100行目の値削除
'        （結合セル対応）
' ======================================================================================

' ======================================================================================
' メイン処理
' ======================================================================================
Public Sub ChangeAndReset()
    Dim wb As Workbook
    Dim wsSource As Worksheet
    Dim ws As Worksheet
    Dim valC2 As Variant
    Dim valC3 As Variant
    Dim targetCount As Long
    Dim writtenCount As Long   ' 記入（D3,F3への入力）を行ったシート数

    ' 実行前の確認
    If MsgBox("変更＆リセットを行っていいですか？", vbYesNo + vbQuestion, "変更＆リセット") <> vbYes Then
        Exit Sub
    End If

    Set wb = ThisWorkbook

    ' シート「月初変更＆リセット」を取得
    On Error Resume Next
    Set wsSource = wb.Worksheets("月初変更＆リセット")
    On Error GoTo 0

    If wsSource Is Nothing Then
        MsgBox "エラー: シート「月初変更＆リセット」が見つかりません。", vbCritical
        Exit Sub
    End If

    ' 処理1用の値を取得
    valC2 = wsSource.Range("C2").Value
    valC3 = wsSource.Range("C3").Value

    targetCount = 0
    writtenCount = 0

    Application.ScreenUpdating = False
    On Error GoTo ErrHandler

    For Each ws In wb.Worksheets
        ' 対象: シート名に「様」が含まれるシート
        If InStr(ws.Name, "様") > 0 Then
            Application.StatusBar = "処理中: " & ws.Name

            ' --- 処理1: C2→D3、C3→F3（記入） ---
            ws.Range("D3").Value = valC2
            ws.Range("F3").Value = valC3
            writtenCount = writtenCount + 1

            ' --- 処理2: A,C,F,G,H,J,K,M,P および R,S,T,U,V,W 列の16〜35行目の値を削除 ---
            Call ClearTargetRange(ws)

            ' --- 処理4: R〜U列の45〜100行目の値を削除 ---
            Call ClearRangeContentsSafe(ws, 45, 100, 18, 21)

            targetCount = targetCount + 1
        End If

        ' --- 処理3: 「明細_」で始まるシート → S3,S4 と Q16:T500 の値を削除 ---
        If Len(ws.Name) >= 3 And Left(ws.Name, 3) = "明細_" Then
            Application.StatusBar = "処理中: " & ws.Name
            Call ClearCellSafe(ws, 3, 19)   ' S3 (S=19列)
            Call ClearCellSafe(ws, 4, 19)   ' S4
            Call ClearRangeContentsSafe(ws, 16, 500, 17, 20)   ' Q=17〜T=20
            targetCount = targetCount + 1
        End If
    Next ws

Done:
    Application.StatusBar = False
    Application.ScreenUpdating = True
    MsgBox "変更＆リセットが完了しました。" & vbCrLf & vbCrLf & "対象シート数: " & targetCount & vbCrLf & "記入を行ったシート数: " & writtenCount, vbInformation
    Exit Sub

ErrHandler:
    Application.StatusBar = False
    Application.ScreenUpdating = True
    MsgBox "エラーが発生しました: " & Err.Description, vbCritical
End Sub

' ======================================================================================
' 対象シートの指定列 16〜35行目の値を削除
' 対象列: A,C,F,G,H,J,K,M,P および R,S,T,U,V,W
' 結合セル対応: 結合範囲は左上セルの値のみクリア（書式・結合は維持）
' ======================================================================================
Private Sub ClearTargetRange(ws As Worksheet)
    Dim r As Long
    Dim colIdx As Variant
    Dim cols As Variant
    Dim cell As Range
    Dim mergeTopLeft As Range

    ' 対象列: A=1, C=3, F=6, G=7, H=8, J=10, K=11, M=13, P=16, R=18, S=19, T=20, U=21, V=22, W=23
    cols = Array(1, 3, 6, 7, 8, 10, 11, 13, 16, 18, 19, 20, 21, 22, 23)

    For r = 16 To 35
        For Each colIdx In cols
            Set cell = ws.Cells(r, colIdx)
            If cell.MergeCells Then
                ' 結合セル: 結合範囲の左上セルが今回のセルなら、その値だけクリア（.Value = "" は結合セルでも可）
                Set mergeTopLeft = cell.MergeArea.Cells(1, 1)
                If cell.Row = mergeTopLeft.Row And cell.Column = mergeTopLeft.Column Then
                    mergeTopLeft.Value = ""
                End If
            Else
                cell.ClearContents
            End If
        Next colIdx
    Next r
End Sub

' ======================================================================================
' 単一セルの値を削除（結合セル対応）
' ======================================================================================
Private Sub ClearCellSafe(ws As Worksheet, rowNum As Long, colNum As Long)
    Dim cell As Range
    Set cell = ws.Cells(rowNum, colNum)
    If cell.MergeCells Then
        cell.MergeArea.Cells(1, 1).Value = ""
    Else
        cell.ClearContents
    End If
End Sub

' ======================================================================================
' 指定範囲の値を削除（結合セル対応: 結合範囲は左上セルの値のみクリア）
' ======================================================================================
Private Sub ClearRangeContentsSafe(ws As Worksheet, rowStart As Long, rowEnd As Long, colStart As Long, colEnd As Long)
    Dim r As Long
    Dim c As Long
    Dim cell As Range
    Dim mergeTopLeft As Range

    For r = rowStart To rowEnd
        For c = colStart To colEnd
            Set cell = ws.Cells(r, c)
            If cell.MergeCells Then
                Set mergeTopLeft = cell.MergeArea.Cells(1, 1)
                If cell.Row = mergeTopLeft.Row And cell.Column = mergeTopLeft.Column Then
                    mergeTopLeft.Value = ""
                End If
            Else
                cell.ClearContents
            End If
        Next c
    Next r
End Sub

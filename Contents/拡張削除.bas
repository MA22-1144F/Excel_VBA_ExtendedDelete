Attribute VB_Name = "拡張削除"

Sub 拡張削除()
    Dim wb As Workbook
    Dim ws As Worksheet
    Set wb = ActiveWorkbook
    Set ws = ActiveSheet
    
    If ws.Application.selection Is Nothing Then
        MsgBox "セルを選択してください。", vbExclamation, "拡張削除"
        Exit Sub
    End If
    
    ' 選択範囲を変数に保存
    Dim selectedRange As Range
    Set selectedRange = ws.Application.selection
    
    ' ユーザフォームを表示して削除方法を選択
    Dim deleteForm As UserForm削除選択
    Set deleteForm = New UserForm削除選択
    
    deleteForm.Show
    
    ' ユーザがキャンセルした場合
    If deleteForm.selectedOption = 0 Then
        Unload deleteForm
        Exit Sub
    End If
    
    ' 選択されたオプションを取得
    Dim selectedOption As Integer
    selectedOption = deleteForm.selectedOption
    Unload deleteForm
    
    ' 確認メッセージ
    Dim confirmMsg As String
    Select Case selectedOption
        Case 1: confirmMsg = "選択範囲の空白セルを削除して左方向にシフトします。"
        Case 2: confirmMsg = "選択範囲の空白セルを削除して上方向にシフトします。"
        Case 3: confirmMsg = "選択範囲の空白セルがある行全体を削除します。"
        Case 4: confirmMsg = "選択範囲の空白セルがある列全体を削除します。"
    End Select
    
    If MsgBox(confirmMsg & vbCrLf & "実行しますか？", vbYesNo + vbQuestion, "拡張削除") = vbNo Then
        Exit Sub
    End If
    
    ' 画面更新を停止
    Application.ScreenUpdating = False
    Application.EnableEvents = False
    
    ' 選択した機能を実行
    Select Case selectedOption
        Case 1: Call 空白セル削除_左シフト(selectedRange, ws)
        Case 2: Call 空白セル削除_上シフト(selectedRange, ws)
        Case 3: Call 空白行削除(selectedRange, ws)
        Case 4: Call 空白列削除(selectedRange, ws)
    End Select
    
    Application.ScreenUpdating = True
    Application.EnableEvents = True
    
    MsgBox "処理が完了しました。", vbInformation, "拡張削除"
End Sub

' 1. 空白セルを削除して左方向にシフト
Sub 空白セル削除_左シフト(targetRange As Range, ws As Worksheet)
    Dim cell As Range
    Dim blankCells As Range
    
    ' 空白セルを特定
    For Each cell In targetRange
        If IsEmpty(cell.value) And cell.Formula = "" Then
            If blankCells Is Nothing Then
                Set blankCells = cell
            Else
                Set blankCells = Union(blankCells, cell)
            End If
        End If
    Next cell
    
    ' 空白セルが見つかった場合、削除実行
    If Not blankCells Is Nothing Then
        blankCells.Delete Shift:=xlToLeft
    End If
End Sub

' 2. 空白セルを削除して上方向にシフト
Sub 空白セル削除_上シフト(targetRange As Range, ws As Worksheet)
    Dim cell As Range
    Dim blankCells As Range
    
    ' 空白セルを特定
    For Each cell In targetRange
        If IsEmpty(cell.value) And cell.Formula = "" Then
            If blankCells Is Nothing Then
                Set blankCells = cell
            Else
                Set blankCells = Union(blankCells, cell)
            End If
        End If
    Next cell
    
    ' 空白セルが見つかった場合、削除実行
    If Not blankCells Is Nothing Then
        blankCells.Delete Shift:=xlUp
    End If
End Sub

' 3. 空白セルがある行全体を削除
Sub 空白行削除(targetRange As Range, ws As Worksheet)
    Dim cell As Range
    Dim i As Long
    
    ' 下から上に向かって処理（行番号の変更を避けるため）
    For i = targetRange.row + targetRange.Rows.count - 1 To targetRange.row Step -1
        ' その行内の選択範囲部分で空白セルがあるかチェック
        Dim rowRange As Range
        Set rowRange = ws.Application.Intersect(targetRange, ws.Rows(i))
        
        If Not rowRange Is Nothing Then
            For Each cell In rowRange
                If IsEmpty(cell.value) And cell.Formula = "" Then
                    ' 空白セルが見つかったら、その行を削除
                    ws.Rows(i).Delete
                    Exit For
                End If
            Next cell
        End If
    Next i
End Sub

' 4. 空白セルがある列全体を削除
Sub 空白列削除(targetRange As Range, ws As Worksheet)
    Dim cell As Range
    Dim i As Long
    
    ' 右から左に向かって処理（列番号の変更を避けるため）
    For i = targetRange.Column + targetRange.Columns.count - 1 To targetRange.Column Step -1
        ' その列内の選択範囲部分で空白セルがあるかチェック
        Dim colRange As Range
        Set colRange = ws.Application.Intersect(targetRange, ws.Columns(i))
        
        If Not colRange Is Nothing Then
            For Each cell In colRange
                If IsEmpty(cell.value) And cell.Formula = "" Then
                    ' 空白セルが見つかったら、その列を削除
                    ws.Columns(i).Delete
                    Exit For
                End If
            Next cell
        End If
    Next i
End Sub

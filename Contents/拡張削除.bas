Attribute VB_Name = "�g���폜"

Sub �g���폜()
    Dim wb As Workbook
    Dim ws As Worksheet
    Set wb = ActiveWorkbook
    Set ws = ActiveSheet
    
    If ws.Application.selection Is Nothing Then
        MsgBox "�Z����I�����Ă��������B", vbExclamation, "�g���폜"
        Exit Sub
    End If
    
    ' �I��͈͂�ϐ��ɕۑ�
    Dim selectedRange As Range
    Set selectedRange = ws.Application.selection
    
    ' ���[�U�t�H�[����\�����č폜���@��I��
    Dim deleteForm As UserForm�폜�I��
    Set deleteForm = New UserForm�폜�I��
    
    deleteForm.Show
    
    ' ���[�U���L�����Z�������ꍇ
    If deleteForm.selectedOption = 0 Then
        Unload deleteForm
        Exit Sub
    End If
    
    ' �I�����ꂽ�I�v�V�������擾
    Dim selectedOption As Integer
    selectedOption = deleteForm.selectedOption
    Unload deleteForm
    
    ' �m�F���b�Z�[�W
    Dim confirmMsg As String
    Select Case selectedOption
        Case 1: confirmMsg = "�I��͈͂̋󔒃Z�����폜���č������ɃV�t�g���܂��B"
        Case 2: confirmMsg = "�I��͈͂̋󔒃Z�����폜���ď�����ɃV�t�g���܂��B"
        Case 3: confirmMsg = "�I��͈͂̋󔒃Z��������s�S�̂��폜���܂��B"
        Case 4: confirmMsg = "�I��͈͂̋󔒃Z���������S�̂��폜���܂��B"
    End Select
    
    If MsgBox(confirmMsg & vbCrLf & "���s���܂����H", vbYesNo + vbQuestion, "�g���폜") = vbNo Then
        Exit Sub
    End If
    
    ' ��ʍX�V���~
    Application.ScreenUpdating = False
    Application.EnableEvents = False
    
    ' �I�������@�\�����s
    Select Case selectedOption
        Case 1: Call �󔒃Z���폜_���V�t�g(selectedRange, ws)
        Case 2: Call �󔒃Z���폜_��V�t�g(selectedRange, ws)
        Case 3: Call �󔒍s�폜(selectedRange, ws)
        Case 4: Call �󔒗�폜(selectedRange, ws)
    End Select
    
    Application.ScreenUpdating = True
    Application.EnableEvents = True
    
    MsgBox "�������������܂����B", vbInformation, "�g���폜"
End Sub

' 1. �󔒃Z�����폜���č������ɃV�t�g
Sub �󔒃Z���폜_���V�t�g(targetRange As Range, ws As Worksheet)
    Dim cell As Range
    Dim blankCells As Range
    
    ' �󔒃Z�������
    For Each cell In targetRange
        If IsEmpty(cell.value) And cell.Formula = "" Then
            If blankCells Is Nothing Then
                Set blankCells = cell
            Else
                Set blankCells = Union(blankCells, cell)
            End If
        End If
    Next cell
    
    ' �󔒃Z�������������ꍇ�A�폜���s
    If Not blankCells Is Nothing Then
        blankCells.Delete Shift:=xlToLeft
    End If
End Sub

' 2. �󔒃Z�����폜���ď�����ɃV�t�g
Sub �󔒃Z���폜_��V�t�g(targetRange As Range, ws As Worksheet)
    Dim cell As Range
    Dim blankCells As Range
    
    ' �󔒃Z�������
    For Each cell In targetRange
        If IsEmpty(cell.value) And cell.Formula = "" Then
            If blankCells Is Nothing Then
                Set blankCells = cell
            Else
                Set blankCells = Union(blankCells, cell)
            End If
        End If
    Next cell
    
    ' �󔒃Z�������������ꍇ�A�폜���s
    If Not blankCells Is Nothing Then
        blankCells.Delete Shift:=xlUp
    End If
End Sub

' 3. �󔒃Z��������s�S�̂��폜
Sub �󔒍s�폜(targetRange As Range, ws As Worksheet)
    Dim cell As Range
    Dim i As Long
    
    ' �������Ɍ������ď����i�s�ԍ��̕ύX������邽�߁j
    For i = targetRange.row + targetRange.Rows.count - 1 To targetRange.row Step -1
        ' ���̍s���̑I��͈͕����ŋ󔒃Z�������邩�`�F�b�N
        Dim rowRange As Range
        Set rowRange = ws.Application.Intersect(targetRange, ws.Rows(i))
        
        If Not rowRange Is Nothing Then
            For Each cell In rowRange
                If IsEmpty(cell.value) And cell.Formula = "" Then
                    ' �󔒃Z��������������A���̍s���폜
                    ws.Rows(i).Delete
                    Exit For
                End If
            Next cell
        End If
    Next i
End Sub

' 4. �󔒃Z���������S�̂��폜
Sub �󔒗�폜(targetRange As Range, ws As Worksheet)
    Dim cell As Range
    Dim i As Long
    
    ' �E���獶�Ɍ������ď����i��ԍ��̕ύX������邽�߁j
    For i = targetRange.Column + targetRange.Columns.count - 1 To targetRange.Column Step -1
        ' ���̗���̑I��͈͕����ŋ󔒃Z�������邩�`�F�b�N
        Dim colRange As Range
        Set colRange = ws.Application.Intersect(targetRange, ws.Columns(i))
        
        If Not colRange Is Nothing Then
            For Each cell In colRange
                If IsEmpty(cell.value) And cell.Formula = "" Then
                    ' �󔒃Z��������������A���̗���폜
                    ws.Columns(i).Delete
                    Exit For
                End If
            Next cell
        End If
    Next i
End Sub

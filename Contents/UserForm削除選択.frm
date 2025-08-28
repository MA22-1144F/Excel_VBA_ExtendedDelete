VERSION 5.00
Begin {C62A69F0-16DC-11CE-9E98-00AA00574A4F} UserForm削除選択 
   Caption         =   "拡張削除"
   ClientHeight    =   3015
   ClientLeft      =   120
   ClientTop       =   465
   ClientWidth     =   3960
   OleObjectBlob   =   "UserForm削除選択.frx":0000
   StartUpPosition =   1  'オーナー フォームの中央
End
Attribute VB_Name = "UserForm削除選択"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False

' フォームレベル変数
Private m_SelectedOption As Integer

' プロパティ: 選択されたオプション（0=キャンセル, 1-4=削除方法）
Public Property Get selectedOption() As Integer
    selectedOption = m_SelectedOption
End Property

' フォーム初期化イベント
Private Sub UserForm_Initialize()
    ' 初期値設定
    m_SelectedOption = 0
    OptionButton1.value = True  ' デフォルトで最初のオプションを選択
    
    ' フォームを画面中央に配置
    Me.StartUpPosition = 0  ' Manual
    Me.Left = (Application.width - Me.width) / 2
    Me.Top = (Application.height - Me.height) / 2
End Sub

' OKボタンクリックイベント
Private Sub CommandButton1_Click()
    ' 選択されたオプションを特定
    If OptionButton1.value Then
        m_SelectedOption = 1
    ElseIf OptionButton2.value Then
        m_SelectedOption = 2
    ElseIf OptionButton3.value Then
        m_SelectedOption = 3
    ElseIf OptionButton4.value Then
        m_SelectedOption = 4
    Else
        m_SelectedOption = 1  ' デフォルト
    End If
    
    ' フォームを非表示
    Me.Hide
End Sub

' キャンセルボタンクリックイベント
Private Sub CommandButton2_Click()
    m_SelectedOption = 0  ' キャンセルを示す
    Me.Hide
End Sub

' Escキー対応
Private Sub UserForm_KeyDown(ByVal KeyCode As MSForms.ReturnInteger, ByVal Shift As Integer)
    If KeyCode = 27 Then  ' Escキー
        CommandButton2_Click  ' キャンセルボタンと同じ動作
    End If
End Sub

' ×ボタン（閉じるボタン）対応
Private Sub UserForm_QueryClose(Cancel As Integer, CloseMode As Integer)
    ' 右上の×ボタンが押された場合
    If CloseMode = 0 Then  ' vbFormControlMenu (×ボタン)
        Cancel = True  ' 通常の閉じる処理をキャンセル
        m_SelectedOption = 0  ' キャンセルを示す
        Me.Hide  ' フォームを非表示にする
    End If
End Sub

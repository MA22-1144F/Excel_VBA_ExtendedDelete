VERSION 5.00
Begin {C62A69F0-16DC-11CE-9E98-00AA00574A4F} UserForm�폜�I�� 
   Caption         =   "�g���폜"
   ClientHeight    =   3015
   ClientLeft      =   120
   ClientTop       =   465
   ClientWidth     =   3960
   OleObjectBlob   =   "UserForm�폜�I��.frx":0000
   StartUpPosition =   1  '�I�[�i�[ �t�H�[���̒���
End
Attribute VB_Name = "UserForm�폜�I��"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False

' �t�H�[�����x���ϐ�
Private m_SelectedOption As Integer

' �v���p�e�B: �I�����ꂽ�I�v�V�����i0=�L�����Z��, 1-4=�폜���@�j
Public Property Get selectedOption() As Integer
    selectedOption = m_SelectedOption
End Property

' �t�H�[���������C�x���g
Private Sub UserForm_Initialize()
    ' �����l�ݒ�
    m_SelectedOption = 0
    OptionButton1.value = True  ' �f�t�H���g�ōŏ��̃I�v�V������I��
    
    ' �t�H�[������ʒ����ɔz�u
    Me.StartUpPosition = 0  ' Manual
    Me.Left = (Application.width - Me.width) / 2
    Me.Top = (Application.height - Me.height) / 2
End Sub

' OK�{�^���N���b�N�C�x���g
Private Sub CommandButton1_Click()
    ' �I�����ꂽ�I�v�V���������
    If OptionButton1.value Then
        m_SelectedOption = 1
    ElseIf OptionButton2.value Then
        m_SelectedOption = 2
    ElseIf OptionButton3.value Then
        m_SelectedOption = 3
    ElseIf OptionButton4.value Then
        m_SelectedOption = 4
    Else
        m_SelectedOption = 1  ' �f�t�H���g
    End If
    
    ' �t�H�[�����\��
    Me.Hide
End Sub

' �L�����Z���{�^���N���b�N�C�x���g
Private Sub CommandButton2_Click()
    m_SelectedOption = 0  ' �L�����Z��������
    Me.Hide
End Sub

' Esc�L�[�Ή�
Private Sub UserForm_KeyDown(ByVal KeyCode As MSForms.ReturnInteger, ByVal Shift As Integer)
    If KeyCode = 27 Then  ' Esc�L�[
        CommandButton2_Click  ' �L�����Z���{�^���Ɠ�������
    End If
End Sub

' �~�{�^���i����{�^���j�Ή�
Private Sub UserForm_QueryClose(Cancel As Integer, CloseMode As Integer)
    ' �E��́~�{�^���������ꂽ�ꍇ
    If CloseMode = 0 Then  ' vbFormControlMenu (�~�{�^��)
        Cancel = True  ' �ʏ�̕��鏈�����L�����Z��
        m_SelectedOption = 0  ' �L�����Z��������
        Me.Hide  ' �t�H�[�����\���ɂ���
    End If
End Sub

Attribute VB_Name = "�V�[�g���ꗗtoClipBoard"
'����
' ���̃��W���[���́ADataObject���g�p���ăN���b�v�{�[�h�ɕ�����𑗂�܂��B
' DataObject���g�p����ɂ́uMicrosoft Forms 2.0 Object Library�v�ւ̎Q�Ƃ��K�v�ł��B
' Visual Basic Editor�̃��j���[����m�c�[���n���m�Q�Ɛݒ�n�R�}���h��I����
' �m�Q�Ɛݒ�n�_�C�A���O�{�b�N�X�ŁuMicrosoft Forms 2.0 Object Library�v�Ƀ`�F�b�N�����āA
' �mOK�n�{�^�����N���b�N���A�Q�Ɛݒ���s���܂��B
'
' �u�Q�Ɖ\�ȃ��C�u���� �t�@�C���v�̃��X�g�ɂȂ��ꍇ�́A
' �m�Q�Ɛݒ�n�_�C�A���O�{�b�N�X�Łm�Q�Ɓn�{�^�����N���b�N����
' �uC:\WINNT(�܂��� Windows)\system32\FM20.DLL�v��I�����܂��Br

'
'
'�J���Ă���u�b�N�̃V�[�g�ꗗ���N���b�v�{�[�h�ɓ\��t���܂�
'�N���b�v�{�[�h�ւ̓\��t����setClipBoad�̃R�����g���Q��
Sub �V�[�g���ꗗtoClipBoard()
    '�V�[�g���̕������ێ����܂�
    Dim workSheetNames As String
      
    For Each targetWorkSheet In Sheets
        workSheetNames = workSheetNames & targetWorkSheet.Name & vbCrLf
    
    Next
    
    '�N���b�v�{�[�h�ɐݒ肵�܂�
    setClipBoad (workSheetNames)

End Sub

'
'�@��������N���b�v�{�[�h�ɓ\��t���܂�
'��������
' �m�c�[���n���m�Q�Ɛݒ�n�ŁuMicrosoft Forms 2.0 Object Library�v��
' �`�F�b�N���Ďg�p����B
'�m�Q�Ɖ\�ȃ��C�u�����n�̃��X�g�ɂȂ��ꍇ�́m�Q�Ɛݒ�n
'�_�C�A���O�{�b�N�X�Łm�Q�Ɓn�{�^�����N���b�N����
'�uC:\Windows\system32\FM20.DLL�v��I������
Function setClipBoad(strValue As String)

    Dim CB As New DataObject
    With CB
        .SetText strValue
        .PutInClipboard
    End With

End Function

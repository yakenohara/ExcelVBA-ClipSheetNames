Attribute VB_Name = "SheetNamesToClipBoard"
'<License>------------------------------------------------------------
'
' Copyright (c) 2018 Shinnosuke Yakenohara
'
' This program is free software: you can redistribute it and/or modify
' it under the terms of the GNU General Public License as published by
' the Free Software Foundation, either version 3 of the License, or
' (at your option) any later version.
'
' This program is distributed in the hope that it will be useful,
' but WITHOUT ANY WARRANTY; without even the implied warranty of
' MERCHANTABILITY or FITNESS FOR A PARTICULAR PURPOSE.  See the
' GNU General Public License for more details.
'
' You should have received a copy of the GNU General Public License
' along with this program.  If not, see <http://www.gnu.org/licenses/>.
'
'-----------------------------------------------------------</License>

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
Sub SheetNamesToClipBoard()
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

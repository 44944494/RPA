On Error Resume Next ' �G���[�������J�n

' Excel�A�v���P�[�V�������쐬
Dim excelApp
Set excelApp = CreateObject("Excel.Application")

' �G���[�����������ꍇ�ɏI�����鏈��
If Err.Number <> 0 Then
    WScript.Echo "Excel�A�v���P�[�V�������쐬�ł��܂���ł����B�G���[�R�[�h: " & Err.Number
    WScript.Quit
End If

' ��\����Excel���J��
excelApp.Visible = False

' �u�b�N���J��
Dim workbookPath
workbookPath = "C:\Users\diabl\Downloads\�y�����z�V�K�o�^�\���� ��ʗpVer3.2_csv�t�@�C���o�͗p.xlsm"

Dim workbook
Set workbook = excelApp.Workbooks.Open(workbookPath)

' �G���[�����������ꍇ�ɏI�����鏈��
If Err.Number <> 0 Then
    WScript.Echo "�u�b�N���J���܂���ł����B�p�X���m�F���Ă��������B�G���[�R�[�h: " & Err.Number
    excelApp.Quit
    Set excelApp = Nothing
    WScript.Quit
End If

' �}�N�������s�i�W�����W���[���ɂ���ꍇ�j
On Error Resume Next
excelApp.Run "CopyMappedColumns"

If Err.Number <> 0 Then
    WScript.Echo "�}�N�������s�ł��܂���ł����B�G���[�R�[�h: " & Err.Number
Else
    WScript.Echo "�}�N���̎��s���������܂����B"
End If

' �K�v�ɉ����ău�b�N��ۑ����ĕ���
workbook.Close False

' Excel�A�v���P�[�V�������I��
excelApp.Quit

' COM�I�u�W�F�N�g�����
Set workbook = Nothing
Set excelApp = Nothing

WScript.Echo "�������I�����܂����B"

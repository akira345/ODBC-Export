Option Explicit
'
'ODBC�ݒ�G�N�X�|�[�g Ver1.0
'

Dim CurrentDir
Dim objShell
Dim ret

'�o�[�W�����`�F�b�N
If (GetOSVersion() <=5.0) Then
	'W2K�ȉ�
	msgbox("WindowsXP�ȍ~�łȂ��Ɠ����܂���B�c�O")
	WScript.Quit
End If

'�J�����g�f�B���N�g���擾
Set objShell = WScript.CreateObject("WScript.Shell")
CurrentDir = objShell.CurrentDirectory

'�x��
ret = objShell.Popup( _
	"�I�x���I" & vbCrlf & "�{�c�[���̓��W�X�g���𑀍삵�܂��B" & vbCrlf & _
	"�g�p�ɍۂ��ẮA���W�X�g���̃o�b�N�A�b�v���s���ȂǁA" & vbCrlf & _
	"�g�p�҂̎��ȐӔC���ōs���Ă��������B", _
	0, "���������s���܂����H", vbOKCancel+vbCritical)
If (ret = vbCancel) Then
	objShell.Popup "�L�����Z�����܂����B", ,, vbInformation
	Set objShell = Nothing
	WScript.Quit
End If
ret = objShell.Popup( _
   "ODBC�̐ݒ���J�����g�f�B���N�g���ɃG�N�X�|�[�g���܂��B" & vbCrlf & _
   "���~����ꍇ�́u�L�����Z���v�������Ă��������B", _
   0, "���������s���܂����H", vbOKCancel+vbQuestion)

 Select Case ret
   Case vbOK
	'�G�N�X�|�[�g
		objShell.Run "reg export HKEY_LOCAL_MACHINE\SOFTWARE\ODBC\ODBC.INI """ & CurrentDir & "\ODBC.reg"""
		objShell.Popup "ODBC.reg�t�@�C���ɃG�N�X�|�[�g���܂����B" & vbCrlf & _
					   "�ڍs��ɃR�s�[���āA�_�u���N���b�N���Ă��������B", ,, vbInformation
   Case vbCancel
     objShell.Popup "�L�����Z�����܂����B", ,, vbInformation
  End Select

Set objShell = Nothing

Function GetOSVersion()
'http://vbsguide.seesaa.net/article/144449959.html���

	' XP �Ȃ�� 5.1 �ƕԂ�܂�

	Dim strComputer
	Dim Wmi 
	Dim colTarget 
	Dim strWork 
	Dim objRow
	Dim aData

	strComputer = "."
	Set Wmi = GetObject("winmgmts:" _
	    & "{impersonationLevel=impersonate}!\\" & strComputer & "\root\cimv2")
	
	Set colTarget = Wmi.ExecQuery( "select Version from Win32_OperatingSystem" )
	For Each objRow in colTarget
		strWork = objRow.Version
	Next

	aData = Split( strWork, "." )
	strWork = aData(0) & "." & aData(1)

	GetOSVersion = CDbl( strWork )

End Function



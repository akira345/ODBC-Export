Option Explicit
'
'ODBC設定エクスポート Ver1.0
'

Dim CurrentDir
Dim objShell
Dim ret

'バージョンチェック
If (GetOSVersion() <=5.0) Then
	'W2K以下
	msgbox("WindowsXP以降でないと動きません。残念")
	WScript.Quit
End If

'カレントディレクトリ取得
Set objShell = WScript.CreateObject("WScript.Shell")
CurrentDir = objShell.CurrentDirectory

'警告
ret = objShell.Popup( _
	"！警告！" & vbCrlf & "本ツールはレジストリを操作します。" & vbCrlf & _
	"使用に際しては、レジストリのバックアップを行うなど、" & vbCrlf & _
	"使用者の自己責任下で行ってください。", _
	0, "処理を実行しますか？", vbOKCancel+vbCritical)
If (ret = vbCancel) Then
	objShell.Popup "キャンセルしました。", ,, vbInformation
	Set objShell = Nothing
	WScript.Quit
End If
ret = objShell.Popup( _
   "ODBCの設定をカレントディレクトリにエクスポートします。" & vbCrlf & _
   "中止する場合は「キャンセル」を押してください。", _
   0, "処理を実行しますか？", vbOKCancel+vbQuestion)

 Select Case ret
   Case vbOK
	'エクスポート
		objShell.Run "reg export HKEY_LOCAL_MACHINE\SOFTWARE\ODBC\ODBC.INI """ & CurrentDir & "\ODBC.reg"""
		objShell.Popup "ODBC.regファイルにエクスポートしました。" & vbCrlf & _
					   "移行先にコピーして、ダブルクリックしてください。", ,, vbInformation
   Case vbCancel
     objShell.Popup "キャンセルしました。", ,, vbInformation
  End Select

Set objShell = Nothing

Function GetOSVersion()
'http://vbsguide.seesaa.net/article/144449959.htmlより

	' XP ならば 5.1 と返ります

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



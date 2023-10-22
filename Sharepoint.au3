#cs ----------------------------------------------------------------------------

 AutoIt Version: 3.3.16.1
 Author:         Rubens Gomes

 Script Function:
	Template AutoIt script.

#ce ----------------------------------------------------------------------------

; Script Start - Add your code below here

ControlSetText("", "", "Scintilla2", "")

#include <ButtonConstants.au3>
#include <GUIConstantsEx.au3>
#include <ComboConstants.au3>
#include <IE.au3>

Global $app = "Sharepoint 0.2"
Global $getSystem = @OSVersion
Global $txtFile = @ScriptDir & "\Sharepoint.txt"
Global $localURL = "https://gruposigla.sharepoint.com/"
Global $finallPt = "/Documentos%20Compartilhados"
Global $finalEn = "/Shared%20Documents"
Global $localSetor[] = ["sites/", "Qualidade", "Sac_RVIMOLA", "Aéreo", "DPP", "RecursosHumanos", "Operacional", "TI", "Transportes"]


#Region ### START Koda GUI section ### Form=
Opt("GUIOnEventMode", 1) ; Enable GUI Event Mode with 1
$formGui = GUICreate($app, 250, 135, 800, 450)
GUISetBkColor(0x000)
GUISetOnEvent($GUI_EVENT_CLOSE, "ExitApp")
$Label1 = GUICtrlCreateLabel("Departamento / Setor", 16, 16, 107, 17)
GUICtrlSetColor(-1, 0xFFFBF0)
$Combo1 = GUICtrlCreateCombo("", 16, 35, 220, 15, BitOR($CBS_DROPDOWNLIST, $CBS_AUTOHSCROLL))
GUICtrlSetData(-1, "TECNOLOGIA DA INFORMAÇÃO|RECURSOS HUMANOS|SAC|AÉREO|QUALIDADE|DPP|OPERACIONAL|TRANSPORTE", "TECNOLOGIA DA INFORMAÇÃO")
$Button1 = GUICtrlCreateButton("Adicionar", 16, 65, 220, 25)
GUICtrlSetOnEvent($Button1, "StartProcess")
$Button2 = GUICtrlCreateButton("Fechar", 16, 95, 220, 25)
GUICtrlSetOnEvent($Button2, "ExitApp")
#EndRegion ### END Koda GUI section ###

StartApp()

Func ExitApp()
	FileSetAttrib($txtFile, "+H")
	If (@error) Then
		MsgBox(16, $app, "Falha ao alterar propiedades do arquivo.")
		Exit
	Else
		Exit
	EndIf
EndFunc   ;==>ExitApp

Func StartProcess()
	FileSetAttrib($txtFile, "+H")
	CheckService()
	MapAdd($localURL)
EndFunc   ;==>StartProcess

Func StartApp()
	$osVersion = StringReplace(@OSVersion, "WIN_", "")
	If ($osVersion >= 8) Then
		CheckFile($txtFile)
		If (FileExists($txtFile)) Then
			GUISetState(@SW_SHOW)
		Else
			GUISetState(@SW_HIDE)
			MsgBox(16, $app, "Falha ao encontrar o arquivo de inicialização." & @CRLF & "Por favor contate o time de TI local, obrigado.", 20)
			ExitApp()
		EndIf
	Else
		MsgBox(16, $app, "Aplicação não suportada no Windows 7 ou inferior.")
		ExitApp()
	EndIf
EndFunc   ;==>StartApp

Func CheckFile($sFile)
	FileSetAttrib($sFile, "-H")
	If (@error) Then
		ConsoleWriteError("Falha ao alterar propiedades dos arquivos.")
	EndIf

	If (FileExists($sFile)) Then
		$sRead = FileRead($sFile)
		If ($sRead >= 1) Then
			$countExec = $sRead + 1
			$fOpen = FileOpen($sFile, 2)
			FileWrite($fOpen, $countExec)
			FileClose($fOpen)
			IEAuth("https://gruposigla.sharepoint.com")
		EndIf

		If ($sRead > 2) Then
			IEAuth("https://gruposigla.sharepoint.com")
			;MsgBox(64, $app, "Autenticação realizada com sucesso." & @CRLF & "A aplicação está sendo finalizada.", 5)
			TrayTip($app, "Autenticação em Office365 realizada com sucesso." & @CRLF & "A aplicação está sendo encerrada.", 5, 1)
			ExitApp()
		EndIf

	Else
		FileWrite($sFile, "1")
		If (@error) Then
			ConsoleWriteError("Falha ao criar o arquivo de verificação." & @CRLF)
			ExitApp()
		Else
			FileSetAttrib($sFile, "+H")
			If (@error) Then
				ConsoleWriteError("Falha ao mudar as propiedades do arquivo.")
			Else
				StartApp()
			EndIf
		EndIf
	EndIf
EndFunc   ;==>CheckFile

Func CheckService()
	$webClientReg = RegRead("HKEY_LOCAL_MACHINE\SYSTEM\ControlSet001\Services\WebClient", "Start")
	If ($webClientReg = 4) Then
		RegWrite("HKEY_LOCAL_MACHINE\SYSTEM\ControlSet001\Services\WebClient", "Start", "REG_DWORD", 2)
		If (@error) Then
			ConsoleWriteError("Falha ao alterar registro do serviço webClient." & @CRLF)
		EndIf
	EndIf

	RunWait("sc start webClient", "", 1)
EndFunc   ;==>CheckService

Func CheckReg()
	If (RegRead("HKEY_LOCAL_MACHINE\SOFTWARE\Microsoft\Windows\CurrentVersion\Run", "Sharepoint")) Then
		;-
	Else
		ConsoleWriteError("Chave de registro não localizada." & @CRLF)
		RegWrite("HKEY_LOCAL_MACHINE\SOFTWARE\Microsoft\Windows\CurrentVersion\Run", "Sharepoint", "REG_SZ", '"' & @ScriptDir & '\Sharepoint.exe"')
		If (@error) Then
			ConsoleWriteError("Falha ao criar registro de inicialização.")
		Else
			;-
		EndIf
	EndIf
	If (RegRead("HKEY_CURRENT_USER\Software\Microsoft\Windows\CurrentVersion\Internet Settings\ZoneMap\Domains\sharepoint.com", "https")) Then
		;-
	Else
		ConsoleWriteError("Chave de registro não encontrada." & @CRLF)
		RegWrite("HKEY_CURRENT_USER\Software\Microsoft\Windows\CurrentVersion\Internet Settings\ZoneMap\Domains\sharepoint.com", "https", "REG_DWORD", 2)
		If (@error) Then
			ConsoleWriteError("Falha ao criar chave de registro.")
		Else
			;-
		EndIf
	EndIf
EndFunc   ;==>CheckReg

Func IEAuth($sURL)
	CheckReg()
	_IECreate($sURL, 0, 0, 1)
	If (@error = 6) Then
		ConsoleWriteError("Falha ao carrregar página por completo, por favor tente novamente.")
		Exit
	Else
		;-
		CheckProcess("iexplore.exe")
		CheckProcess("ielowutil.exe")
	EndIf
EndFunc   ;==>IEAuth

Func CheckProcess($sProcess)
	If (ProcessExists($sProcess)) Then
		ProcessClose($sProcess)
		If (@error) Then
			ConsoleWriteError("Falha ao finalizar o processo " & $sProcess)
		Else
			;-
		EndIf
	EndIf
EndFunc   ;==>CheckProcess

Func MapAdd($localURL)

	;LINHA PARA ADICIONAR O SAC
	If (GUICtrlRead($Combo1) = "SAC") Then
		DriveMapAdd("S:", $localURL & $localSetor[0] & $localSetor[2] & $finallPt, 1)
		If (@error) Then
			MsgBox(16, $app, "Falha ao mapear a unidade de rede " & $localSetor[2] & @CRLF & "Por favor contate o time de TI local, obrigado.")
			Exit
		Else
			MsgBox(64, $app, "Unidade de rede disponível.")
		EndIf

		;LINHA PARA ADICIONAR O TRANSPORTES
	ElseIf (GUICtrlRead($Combo1) = "TRANSPORTE") Then
		DriveMapAdd("T:", $localURL & $localSetor[0] & $localSetor[8] & $finalEn, 1)
		If (@error) Then
			MsgBox(16, $app, "Falha ao mapear a unidade de rede " & $localSetor[8] & @CRLF & "Por favor contate o time de TI local, obrigado.")
			ExitApp()
		Else
			MsgBox(64, $app, "Unidade de rede disponível.")
		EndIf

		;LINHA PARA ADICIONAR O TI
	ElseIf (GUICtrlRead($Combo1) = "TECNOLOGIA DA INFORMAÇÃO") Then
		DriveMapAdd("T:", $localURL & $localSetor[7], 1)
		If (@error) Then
			ConsoleWriteError($localURL & $localSetor[7])
			MsgBox(16, $app, "Falha ao mapear a unidade de rede " & $localSetor[7] & @CRLF & "Por favor contate o time de TI local, obrigado.")
			ExitApp()
		Else
			MsgBox(64, $app, "Unidade de rede disponível.")
		EndIf

		;LINHA PARA ADICIONAR A QUALIDADE
	ElseIf (GUICtrlRead($Combo1) = "QUALIDADE") Then
		DriveMapAdd("Q:", $localURL & $localSetor[0] & $localSetor[1] & $finalEn, 1)
		If (@error) Then
			MsgBox(16, $app, "Falha ao mapear a unidade de rede " & $localSetor[1] & @CRLF & "Por favor contate o time de TI local, obrigado.")
			ExitApp()
		Else
			MsgBox(64, $app, "Unidade de rede disponível.")

		EndIf

		;LINHA PARA ADICIONAR O AÉREO
	ElseIf (GUICtrlRead($Combo1) = "AÉREO") Then
		DriveMapAdd("A:", $localURL & $localSetor[0] & StringReplace($localSetor[3], "é", "e") & $finalEn, 1)
		If (@error) Then
			MsgBox(16, $app, "Falha ao mapear a unidade de rede " & $localSetor[3] & @CRLF & "Por favor contate o time de TI local, obrigado.")
			ExitApp()
		Else
			MsgBox(64, $app, "Unidade de rede disponível.")
		EndIf

		;LINHA PARA ADICIONAR RECURSOS HUMANOS
	ElseIf (GUICtrlRead($Combo1) = "RECURSOS HUMANOS") Then
		DriveMapAdd("R:", $localURL & $localSetor[0] & $localSetor[5] & $finalEn, 1)
		If (@error) Then
			MsgBox(16, $app, "Falha ao mapear a unidade de rede " & $localSetor[5] & @CRLF & "Por favor contate o time de TI local, obrigado")
			ExitApp()
		Else
			MsgBox(64, $app, "Unidade de rede disponível.")
		EndIf

		;LINHA PARA ADICIONAR O DPP
	ElseIf (GUICtrlRead($Combo1) = "DPP") Then
		DriveMapAdd("D:", $localURL & $localSetor[0] & $localSetor[4] & $finalEn, 1)
		If (@error) Then
			MsgBox(16, $app, "Falha ao mapear a unidade de rede " & $localSetor[4] & @CRLF & "Por favor contate o time de TI local, obrigado.")
		Else
			MsgBox(64, $app, "Unidade de rede disponível.")
		EndIf

		;LINHA PARA ADICIONAR O OPERACIONAL
	ElseIf (GUICtrlRead($Combo1) = "OPERACIONAL") Then
		DriveMapAdd("O:", $localURL & $localSetor[0] & $localSetor[6] & $finallPt, 1)
		If (@error) Then
			MsgBox(16, $app, "Falha ao mapear a unidade de rede " & $localSetor[6] & @CRLF & "Por favor contate o time de TI local, obrigado.")
		Else
			MsgBox(64, $app, "Unidade de rede disponível.")
		EndIf
	EndIf
EndFunc   ;==>MapAdd

While 1
	Sleep(1)
WEnd

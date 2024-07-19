#include <ScreenCapture.au3>

; Caminho para o execut�vel do Google Chrome
Local $sChromePath = "C:\Program Files\Google\Chrome\Application\chrome.exe"

; Criar pasta para armazenar as screenshots
Local $sFolderPath = @MyDocumentsDir & "\Evidencias\Bancas\" & @YEAR & @MON & @MDAY & "\" & @HOUR & "-" & @MIN & "-" & @SEC
DirCreate($sFolderPath)

; URLs do site
Local $aURLs[4] = ["https://banca.claro.com.br/", "https://banca.claro.com.br/#/login", "https://banca.claro.com.br/#/signin", "https://banca.claro.com.br/#/sms-sended"]

; Par�metros para abrir o Chrome em segundo plano
Local $sParams = " --incognito --disable-gpu --disable-infobars --no-first-run --no-default-browser-check --disable-notifications"

; Abre o Chrome em abas an�nimas para cada URL, captura o print de cada aba e salva as imagens
For $i = 0 To UBound($aURLs) - 1
    Local $sURL = $aURLs[$i]
    Local $sScreenshotPath = $sFolderPath & "\screenshot" & $i & ".png"

    ; Executa o Chrome com a URL especificada em uma nova aba an�nima
    Run($sChromePath & " --incognito --new-tab " & $sURL )

	Sleep (2000)

	; Envia a tecla F5 para atualizar a p�gina
     Send("{F5}")
	 Sleep (5000)

	; Aguardar at� que a janela do Chrome seja aberta
	Local $hChromeWindow
	While Not WinExists("[CLASS:Chrome_WidgetWin_1]")
		Sleep(500)
	WEnd

	; Aguardar at� que a p�gina esteja totalmente carregada
	While _WinAPI_GetClassName($hChromeWindow) <> "Chrome_RenderWidgetHostHWND"
    Sleep(500)
	WEnd

	;$hChromeWindow = WinGetHandle("[CLASS:Chrome_WidgetWin_1]")
    ; Aguarda novamente para permitir que a p�gina seja carregada ap�s a atualiza��o
    Sleep(2000)

    ; Captura o print da aba atual e salva a imagem
    _ScreenCapture_CaptureWnd($sScreenshotPath, $hChromeWindow, $sChromePath)

 Next

; Mensagem com o caminho onde as evid�ncias est�o salvas
	MsgBox(0, "AVISO", "Evid�ncias salvas em: " & $sFolderPath)


; Fechar o Chrome
ProcessClose("chrome.exe")

; Abre o caminho completo usando ShellExecute
ShellExecute($sFolderPath)

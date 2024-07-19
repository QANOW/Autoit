#include <ScreenCapture.au3>

; Caminho para o executável do Google Chrome
Local $sChromePath = "C:\Program Files\Google\Chrome\Application\chrome.exe"

; Criar pasta para armazenar as screenshots
Local $sFolderPath = @MyDocumentsDir & "\Evidencias\Bancas\" & @YEAR & @MON & @MDAY & "\" & @HOUR & "-" & @MIN & "-" & @SEC
DirCreate($sFolderPath)

; URLs do site
Local $aURLs[4] = ["https://banca.claro.com.br/", "https://banca.claro.com.br/#/login", "https://banca.claro.com.br/#/signin", "https://banca.claro.com.br/#/sms-sended"]

; Parâmetros para abrir o Chrome em segundo plano
Local $sParams = " --incognito --disable-gpu --disable-infobars --no-first-run --no-default-browser-check --disable-notifications"

; Abre o Chrome em abas anônimas para cada URL, captura o print de cada aba e salva as imagens
For $i = 0 To UBound($aURLs) - 1
    Local $sURL = $aURLs[$i]
    Local $sScreenshotPath = $sFolderPath & "\screenshot" & $i & ".png"

    ; Executa o Chrome com a URL especificada em uma nova aba anônima
    Run($sChromePath & " --incognito --new-tab " & $sURL )

	Sleep (2000)

	; Envia a tecla F5 para atualizar a página
     Send("{F5}")
	 Sleep (5000)

	; Aguardar até que a janela do Chrome seja aberta
	Local $hChromeWindow
	While Not WinExists("[CLASS:Chrome_WidgetWin_1]")
		Sleep(500)
	WEnd

	; Aguardar até que a página esteja totalmente carregada
	While _WinAPI_GetClassName($hChromeWindow) <> "Chrome_RenderWidgetHostHWND"
    Sleep(500)
	WEnd

	;$hChromeWindow = WinGetHandle("[CLASS:Chrome_WidgetWin_1]")
    ; Aguarda novamente para permitir que a página seja carregada após a atualização
    Sleep(2000)

    ; Captura o print da aba atual e salva a imagem
    _ScreenCapture_CaptureWnd($sScreenshotPath, $hChromeWindow, $sChromePath)

 Next

; Mensagem com o caminho onde as evidências estão salvas
	MsgBox(0, "AVISO", "Evidências salvas em: " & $sFolderPath)


; Fechar o Chrome
ProcessClose("chrome.exe")

; Abre o caminho completo usando ShellExecute
ShellExecute($sFolderPath)

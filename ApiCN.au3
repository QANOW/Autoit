#include <Inet.au3>
#include <MsgBoxConstants.au3>
#include "JSON.au3"

; Definindo a fun��o JsonDecode()
Func JsonDecode($sJson)
    Local $oJson = ObjCreate("Chilkat_9_5_0.JsonObject")

    If $oJson = 0 Then
        Return SetError(1, 0, "")
    EndIf

    If $oJson.Load($sJson) <> 1 Then
        Return SetError(2, 0, "")
    EndIf

    Return $oJson
EndFunc


; Fazendo a solicita��o GET para a API
Local $sAPIUrl = "https://hmg-banca.imusica.com.br/assets/FontManifest.json"
Local $sResponse = _INetGetSource($sAPIUrl)

If @error Then
    MsgBox($MB_OK, "Erro", "Ocorreu um erro ao acessar a API.")
    Exit
EndIf

; Decodificando a resposta JSON
Local $oJSON = JsonDecode($sResponse)

If @error Then
    MsgBox($MB_OK, "Erro", "Ocorreu um erro ao decodificar a resposta JSON.")
    Exit
EndIf

; Verificando se a resposta cont�m dados v�lidos
If Not IsObj($oJSON) Then
    MsgBox($MB_OK, "Erro", "A resposta JSON n�o cont�m dados v�lidos.")
    Exit
EndIf

; Verificando se existem destaques de not�cias
If Not $oJSON.HasKey("highlights") Then
    MsgBox($MB_OK, "Erro", "A resposta JSON n�o cont�m destaques de not�cias v�lidos.")
    Exit
EndIf

Local $oHighlights = $oJSON.Get("highlights")
If Not IsArray($oHighlights) Or UBound($oHighlights) = 0 Then
    MsgBox($MB_OK, "Erro", "A resposta JSON n�o cont�m nenhum destaque de not�cia.")
    Exit
EndIf

; Extraindo informa��es dos destaques de not�cias
For $i = 0 To UBound($oHighlights) - 1
    Local $oHighlight = $oHighlights[$i]
    Local $sTitle = $oHighlight.Get("title")
    Local $sDescription = $oHighlight.Get("description")
    Local $sAuthor = $oHighlight.Get("author")

    ; Verificando se os valores s�o v�lidos
    If @error Then
        MsgBox($MB_OK, "Erro", "Ocorreu um erro ao acessar as informa��es do destaque de not�cia #" & ($i + 1) & ".")
        ContinueLoop
    EndIf

    ; Exibindo informa��es dos destaques de not�cias
    Local $sMessage = "Destaque de Not�cia #" & ($i + 1) & @CRLF & _
                      "T�tulo: " & $sTitle & @CRLF & _
                      "Descri��o: " & $sDescription & @CRLF & _
                      "Autor: " & $sAuthor
    MsgBox($MB_OK, "Destaque de Not�cia", $sMessage)
Next

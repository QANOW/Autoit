#include <Inet.au3>
#include <MsgBoxConstants.au3>
#include "JSON.au3"

; Definindo a função JsonDecode()
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


; Fazendo a solicitação GET para a API
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

; Verificando se a resposta contém dados válidos
If Not IsObj($oJSON) Then
    MsgBox($MB_OK, "Erro", "A resposta JSON não contém dados válidos.")
    Exit
EndIf

; Verificando se existem destaques de notícias
If Not $oJSON.HasKey("highlights") Then
    MsgBox($MB_OK, "Erro", "A resposta JSON não contém destaques de notícias válidos.")
    Exit
EndIf

Local $oHighlights = $oJSON.Get("highlights")
If Not IsArray($oHighlights) Or UBound($oHighlights) = 0 Then
    MsgBox($MB_OK, "Erro", "A resposta JSON não contém nenhum destaque de notícia.")
    Exit
EndIf

; Extraindo informações dos destaques de notícias
For $i = 0 To UBound($oHighlights) - 1
    Local $oHighlight = $oHighlights[$i]
    Local $sTitle = $oHighlight.Get("title")
    Local $sDescription = $oHighlight.Get("description")
    Local $sAuthor = $oHighlight.Get("author")

    ; Verificando se os valores são válidos
    If @error Then
        MsgBox($MB_OK, "Erro", "Ocorreu um erro ao acessar as informações do destaque de notícia #" & ($i + 1) & ".")
        ContinueLoop
    EndIf

    ; Exibindo informações dos destaques de notícias
    Local $sMessage = "Destaque de Notícia #" & ($i + 1) & @CRLF & _
                      "Título: " & $sTitle & @CRLF & _
                      "Descrição: " & $sDescription & @CRLF & _
                      "Autor: " & $sAuthor
    MsgBox($MB_OK, "Destaque de Notícia", $sMessage)
Next

#include <Array.au3>
#include <File.au3>
#include <GUIConstantsEx.au3>
#include <EditConstants.au3>
#include <MsgBoxConstants.au3>
#include <WinAPIFiles.au3>
#include <Json.au3>

; Caminho do arquivo JSON
Global $sJsonFilePath = FileOpenDialog("Selecione um arquivo JSON", "", "JSON Files (*.json)", $FD_FILEMUSTEXIST)
If @error Then
    MsgBox($MB_ICONERROR, "Erro", "Nenhum arquivo selecionado.")
    Exit
EndIf

; Ler o conteúdo do arquivo JSON
Global $sJsonContent = FileRead($sJsonFilePath)
If @error Then
    MsgBox($MB_ICONERROR, "Erro", "Erro ao ler o arquivo JSON.")
    Exit
EndIf

; Decodificar o JSON
Global $aData = Json_Decode($sJsonContent)
If @error Then
    MsgBox($MB_ICONERROR, "Erro", "Falha ao decodificar o JSON.")
    Exit
EndIf

; Criar uma lista de versículos únicos
Global $aVerses[1] = ["Selecione um versículo..."]
For $i = 0 To UBound($aData) - 1
    Local $sVerse = $aData[$i]["book"] & " " & $aData[$i]["chapter"] & ":" & $aData[$i]["verse"]
    If Not _ArraySearch($aVerses, $sVerse) Then
        _ArrayAdd($aVerses, $sVerse)
    EndIf
Next

; Criar a interface gráfica
Global $hGUI = GUICreate("Bíblia", 600, 400)

Global $hEdit = GUICtrlCreateEdit("", 10, 10, 580, 260, BitOR($ES_AUTOVSCROLL, $ES_AUTOHSCROLL, $ES_READONLY))
GUICtrlSetFont($hEdit, 10, 400, Default, "Courier New")

Global $hCombo = GUICtrlCreateCombo("", 10, 280, 400, 25)
GUICtrlSetData($hCombo, $aVerses)
GUICtrlSetData($hCombo, "Selecione um versículo...", 0)

Global $hSearchButton = GUICtrlCreateButton("Buscar", 420, 280, 80, 25)

Global $hSimilarVersesLabel = GUICtrlCreateLabel("Passagens similares:", 10, 320, 200, 20)
GUICtrlSetFont($hSimilarVersesLabel, 10, 400, Default, "Arial")

Global $hSimilarVersesEdit = GUICtrlCreateEdit("", 10, 340, 580, 50, BitOR($ES_AUTOVSCROLL, $ES_AUTOHSCROLL, $ES_READONLY))
GUICtrlSetFont($hSimilarVersesEdit, 10, 400, Default, "Courier New")

GUISetState(@SW_SHOW, $hGUI)

; Função para exibir o versículo selecionado
Func DisplayVerse($sVerse)
    GUICtrlSetData($hEdit, "")
    For $i = 0 To UBound($aData) - 1
        If $aData[$i]["book"] & " " & $aData[$i]["chapter"] & ":" & $aData[$i]["verse"] = $sVerse Then
            GUICtrlSetData($hEdit, $aData[$i]["book"] & " " & $aData[$i]["chapter"] & ":" & $aData[$i]["verse"] & " - " & $aData[$i]["text"])
            ExitLoop
        EndIf
    Next
EndFunc

; Função para exibir passagens similares
Func DisplaySimilarVerses($sText)
    GUICtrlSetData($hSimilarVersesEdit, "")
    For $i = 0 To UBound($aData) - 1
        If $aData[$i]["text"] = $sText And $aData[$i]["book"] & " " & $aData[$i]["chapter"] & ":" & $aData[$i]["verse"] <> GUICtrlRead($hCombo) Then
            GUICtrlSetData($hSimilarVersesEdit, $aData[$i]["book"] & " " & $aData[$i]["chapter"] & ":" & $aData[$i]["verse"] & " - " & $aData[$i]["text"] & @CRLF, 1)
        EndIf
    Next
EndFunc

; Loop principal da GUI
While 1
    Switch GUIGetMsg()
        Case $GUI_EVENT_CLOSE
            ExitLoop
        Case $hSearchButton
            Local $sSelectedVerse = GUICtrlRead($hCombo)
            If $sSelectedVerse <> "Selecione um versículo..." Then
                DisplayVerse($sSelectedVerse)
                Local $sVerseText = StringSplit(GUICtrlRead($hEdit), @LF)
                DisplaySimilarVerses($sVerseText[2])
            EndIf
    EndSwitch
WEnd

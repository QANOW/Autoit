#include-once

#include "AutoItConstants.au3"
#include "FileConstants.au3"
#include "WinAPIError.au3"

; #INDEX# =======================================================================================================================
; Title .........: Chrome Automation UDF Library for AutoIt3
; AutoIt Version : 3.3.14.2
; Language ......: English
; Description ...: A collection of functions for creating, attaching to, reading from and manipulating Chrome.
; Author(s) .....: DaleHohm, big_daddy, jpm
; Dll ...........: user32.dll, ole32.dll, oleacc.dll
; ===============================================================================================================================

#Region Header
#cs
	Title:   Chrome Automation UDF Library for AutoIt3
	Filename:  Chrome.au3
	Description: A collection of functions for creating, attaching to, reading from and manipulating Chrome
	Author:   DaleHohm
	ModifChromed: jpm, Jon
	Version:  T3.0-1
	Last Update: 13/06/02
	Requirements: AutoIt3 3.3.9 or higher

	Update History:
	===================================================
	T3.0-2 14/8/19

	Enhancements
	- Updated  __ChromeErrorHandlerRegister to work with or without COM errors being fatal

	T3.0-1 13/6/2

	Enhancements
	- Fixed _Chrome_Introduction, _Chrome_Examples generate HTML5
	- Added check in __ChromeComErrorUnrecoverable for COM error -2147023174, "RPC server not accessible."
	- Fixed check in __ChromeComErrorUnrecoverable for COM error -2147024891, "Access is denChromed."
	- Fixed check in __ChromeComErrorUnrecoverable for COM error  -2147352567, "an exception has occurred."
	- Fixed __ChromeIsObjType() not restoring _ChromeErrorNotify()
	- Fixed $b_mustUnlock on Error in _ChromeCreate()
	- Fixed no timeout cheking if error in _ChromeLoadWait()
	- Fixed HTML5 support in _ChromeImgClick(), _ChromeFormImageClick()
	- Fixed _ChromeHeadInsertEventScript() COM error return
	- Updated _ChromeErrorNotify() default keyword support
	- Updated rename __ChromeNotify() to __ChromeConsoleWriteError() and restore calling  @error
	- Removed __ChromeInternalErrorHandler() (not used any more)
	- Updated Function Headers
	- Updated doc and splitting and checking examples

	T3.0-0 12/9/3

	Fixes
	- Removed _ChromeErrorHandlerRegister() and all internal calls to it.  Unneeded as COM errors are no longer fatal
	- Removed code deprecated in V2
	- Fixed _ChromeLoadWait check for unrecoverable COM errors
	- Removed Vcard support from _ChromePropertyGet (Chrome removed support in Chrome7)
	- Code cleanup with #AutoIt3Wrapper_Au3Check_Parameters=-d -w 1 -w 2 -w 3 -w- 4 -w 5 -w 6

	New Features
	- Added "scrollIntoVChromew" to _ChromeAction

	Enhancements
	- Added check in __ChromeComErrorUnrecoverable for COM error -2147023179, "The interface is unknown."
	- Added "Trap COM error, report and return" to functions that perform blind method calls (those without return values)

	===================================================
#ce
#EndRegion Header

; #VARIABLES# ===================================================================================================================
#Region Global Variables
Global $__g_iChromeLoadWaitTimeout = 300000 ; 5 Minutes
Global $__g_bChromeAU3Debug = False
Global $__g_bChromeErrorNotify = True
Global $__g_oChromeErrorHandler, $__g_sChromeUserErrorHandler
#EndRegion Global Variables
; ===============================================================================================================================

; #CONSTANTS# ===================================================================================================================
#Region Global Constants
Global Const $__gaChromeAU3VersionInfo[6] = ["T", 3, 0, 2, "20140819", "T3.0-2"]
Global Const $LSFW_LOCK = 1, $LSFW_UNLOCK = 2
;
; Enums
;
Global Enum _; Error Status Types
		$_ChromeSTATUS_Success = 0, _
		$_ChromeSTATUS_GeneralError, _
		$_ChromeSTATUS_ComError, _
		$_ChromeSTATUS_InvalidDataType, _
		$_ChromeSTATUS_InvalidObjectType, _
		$_ChromeSTATUS_InvalidValue, _
		$_ChromeSTATUS_LoadWaitTimeout, _
		$_ChromeSTATUS_NoMatch, _
		$_ChromeSTATUS_AccessIsDenChromed, _
		$_ChromeSTATUS_ClChromentDisconnected
;~ Global Enum Step * 2 _; NotificationLevel
;~ 		$_ChromeNotifyLevel_None = 0, _
;~ 		$_ChromeNotifyNotifyLevel_Warning = 1, _
;~ 		$_ChromeNotifyNotifyLevel_Error, _
;~ 		$_ChromeNotifyNotifyLevel_ComError
;~ Global Enum Step * 2 _; NotificationMethod
;~ 		$_ChromeNotifyMethod_Silent = 0, _
;~ 		$_ChromeNotifyMethod_Console = 1, _
;~ 		$_ChromeNotifyMethod_ToolTip, _
;~ 		$_ChromeNotifyMethod_MsgBox
#EndRegion Global Constants
; ===============================================================================================================================

; #NO_DOC_FUNCTION# =============================================================================================================
; _ChromeErrorHandlerRegister
; _ChromeErrorHandlerDeRegister
; ===============================================================================================================================

; #CURRENT# =====================================================================================================================
; _ChromeCreate
; _ChromeCreateEmbedded
; _ChromeNavigate
; _ChromeAttach
; _ChromeLoadWait
; _ChromeLoadWaitTimeout
;
; _ChromeIsFrameSet
; _ChromeFrameGetCollection
; _ChromeFrameGetObjByName
;
; _ChromeLinkClickByText
; _ChromeLinkClickByIndex
; _ChromeLinkGetCollection
;
; _ChromeImgClick
; _ChromeImgGetCollection
;
; _ChromeFormGetCollection
; _ChromeFormGetObjByName
; _ChromeFormElementGetCollection
; _ChromeFormElementGetObjByName
; _ChromeFormElementGetValue
; _ChromeFormElementSetValue
; _ChromeFormElementOptionSelect
; _ChromeFormElementCheckBoxSelect
; _ChromeFormElementRadioSelect
; _ChromeFormImageClick
; _ChromeFormSubmit
; _ChromeFormReset
;
; _ChromeTableGetCollection
; _ChromeTableWriteToArray
;
; _ChromeBodyReadHTML
; _ChromeBodyReadText
; _ChromeBodyWriteHTML
; _ChromeDocReadHTML
; _ChromeDocWriteHTML
; _ChromeDocInsertText
; _ChromeDocInsertHTML
; _ChromeHeadInsertEventScript
;
; _ChromeDocGetObj
; _ChromeTagNameGetCollection
; _ChromeTagNameAllGetCollection
; _ChromeGetObjByName
; _ChromeGetObjById
; _ChromeAction
; _ChromePropertyGet
; _ChromePropertySet
; _ChromeErrorNotify
; _ChromeQuit
;
; _Chrome_Introduction
; _Chrome_Example
; _Chrome_VersionInfo
; ===============================================================================================================================

; #INTERNAL_USE_ONLY# ===========================================================================================================
; __ChromeLockSetForegroundWindow
; __ChromeControlGetObjFromHWND
; __ChromeRegisterWindowMessage
; __ChromeSendMessageTimeout
; __ChromeIsObjType
; __ChromeConsoleWriteError
; __ChromeComErrorUnrecoverable
;
; __ChromeInternalErrorHandler
; __ChromeInternalErrorHandlerRegister
; __ChromeNavigate
; __ChromeCreateNewChrome
; __ChromeTempFile
;
; __ChromeStringToBstr
; __ChromeBstrToString
; ===============================================================================================================================

#Region Core functions
; #FUNCTION# ====================================================================================================================
; Author ........: Dale Hohm
; ModifChromed ......: jpm
; ===============================================================================================================================
Func _ChromeCreate($sUrl = "about:blank", $iTryAttach = 0, $iVisible = 1, $iWait = 1, $iTakeFocus = 1)
	If Not $iVisible Then $iTakeFocus = 0 ; Force takeFocus to 0 for hidden window

	If $iTryAttach Then
		Local $oResult = _ChromeAttach($sUrl, "url")
		If IsObj($oResult) Then
			If $iTakeFocus Then WinActivate(HWnd($oResult.hWnd))
			Return SetError($_ChromeSTATUS_Success, 1, $oResult)
		EndIf
	EndIf

	Local $iMustUnlock = 0
	If Not $iVisible And __ChromeLockSetForegroundWindow($LSFW_LOCK) Then $iMustUnlock = 1

	Local $oObject = ObjCreate("InternetExplorer.Application")
	If Not IsObj($oObject) Then
		__ChromeConsoleWriteError("Error", "_ChromeCreate", "", "Browser Object Creation Failed")
		If $iMustUnlock Then __ChromeLockSetForegroundWindow($LSFW_UNLOCK)
		Return SetError($_ChromeSTATUS_GeneralError, 0, 0)
	EndIf

	$oObject.visible = $iVisible

	; If the unlock doesn't work we may have created an unwanted modal window
	If $iMustUnlock And Not __ChromeLockSetForegroundWindow($LSFW_UNLOCK) Then __ChromeConsoleWriteError("Warning", "_ChromeCreate", "", "Foreground Window Unlock Failed!")
	_ChromeNavigate($oObject, $sUrl, $iWait)

	; Store @error after _ChromeNavigate() so that it can be returned.
	Local $ChromeError = @error

	; Chrome9 sets focus to the URL bar when an about: URI is displayed (such as about:blank).  This can cause
	; _ChromeAction(..., "focus") to work incorrectly.  It will give focus to the element (as shown by the elements's
	; appearance changing but) the input caret will not move.  The work-around for this "helpful" behavior is
	; to explicitly give focus to the document.  We should only do this for about: URIs and on successful
	; navigate.
	If Not $ChromeError And StringLeft($sUrl, 6) = "about:" Then
		Local $oDocument = $oObject.document
		_ChromeAction($oDocument, "focus")
	EndIf

	Return SetError($ChromeError, 0, $oObject)
EndFunc   ;==>_ChromeCreate

; #FUNCTION# ====================================================================================================================
; Author ........: Dale Hohm
; ===============================================================================================================================
Func _ChromeCreateEmbedded()
	Local $oObject = ObjCreate("Shell.Explorer.2")

	If Not IsObj($oObject) Then
		__ChromeConsoleWriteError("Error", "_ChromeCreateEmbedded", "", "WebBrowser Object Creation Failed")
		Return SetError($_ChromeSTATUS_GeneralError, 0, 0)
	EndIf
	;
	Return SetError($_ChromeSTATUS_Success, 0, $oObject)
EndFunc   ;==>_ChromeCreateEmbedded

; #FUNCTION# ====================================================================================================================
; Author ........: Dale Hohm
; ===============================================================================================================================
Func _ChromeNavigate(ByRef $oObject, $sUrl, $iWait = 1)
	If Not IsObj($oObject) Then
		__ChromeConsoleWriteError("Error", "_ChromeNavigate", "$_ChromeSTATUS_InvalidDataType")
		Return SetError($_ChromeSTATUS_InvalidDataType, 1, 0)
	EndIf
	;
	If Not __ChromeIsObjType($oObject, "documentContainer") Then
		__ChromeConsoleWriteError("Error", "_ChromeNavigate", "$_ChromeSTATUS_InvalidObjectType")
		Return SetError($_ChromeSTATUS_InvalidObjectType, 1, 0)
	EndIf
	;
	$oObject.navigate($sUrl)
	If @error Then ; Trap COM error, report and return
		__ChromeConsoleWriteError("Error", "_ChromeNavigate", "$_ChromeSTATUS_COMError", @error)
		Return SetError($_ChromeSTATUS_ComError, @error, 0)
	EndIf
	If $iWait Then
		_ChromeLoadWait($oObject)
		Return SetError(@error, 0, -1)
	EndIf

	Return SetError($_ChromeSTATUS_Success, 0, -1)
EndFunc   ;==>_ChromeNavigate

; #FUNCTION# ====================================================================================================================
; Author ........: Dale Hohm
; ===============================================================================================================================
Func _ChromeAttach($sString, $sMode = "title", $iInstance = 1)
	$sMode = StringLower($sMode)

	$iInstance = Int($iInstance)
	If $iInstance < 1 Then
		__ChromeConsoleWriteError("Error", "_ChromeAttach", "$_ChromeSTATUS_InvalidValue", "$iInstance < 1")
		Return SetError($_ChromeSTATUS_InvalidValue, 3, 0)
	EndIf

	If $sMode = "embedded" Or $sMode = "dialogbox" Then
		Local $iWinTitleMatchMode = Opt("WinTitleMatchMode", $OPT_MATCHANY)
		If $sMode = "dialogbox" And $iInstance > 1 Then
			If IsHWnd($sString) Then
				$iInstance = 1
				__ChromeConsoleWriteError("Warning", "_ChromeAttach", "$_ChromeSTATUS_GeneralError", "$iInstance > 1 invalid with HWnd and DialogBox.  Setting to 1.")
			Else
				Local $aWinlist = WinList($sString, "")
				If $iInstance <= $aWinlist[0][0] Then
					$sString = $aWinlist[$iInstance][1]
					$iInstance = 1
				Else
					__ChromeConsoleWriteError("Warning", "_ChromeAttach", "$_ChromeSTATUS_NoMatch")
					Opt("WinTitleMatchMode", $iWinTitleMatchMode)
					Return SetError($_ChromeSTATUS_NoMatch, 1, 0)
				EndIf
			EndIf
		EndIf
		Local $hControl = ControlGetHandle($sString, "", "[CLASS:Chrome_Server; INSTANCE:" & $iInstance & "]")
		Local $oResult = __ChromeControlGetObjFromHWND($hControl)
		Opt("WinTitleMatchMode", $iWinTitleMatchMode)
		If IsObj($oResult) Then
			Return SetError($_ChromeSTATUS_Success, 0, $oResult)
		Else
			__ChromeConsoleWriteError("Warning", "_ChromeAttach", "$_ChromeSTATUS_NoMatch")
			Return SetError($_ChromeSTATUS_NoMatch, 1, 0)
		EndIf
	EndIf

	Local $oShell = ObjCreate("Shell.Application")
	Local $oShellWindows = $oShell.Windows(); collection of all ShellWindows (Chrome and File Explorer)
	Local $iTmp = 1
	Local $iNotifyStatus, $bIsBrowser, $sTmp
	Local $bStatus
	For $oWindow In $oShellWindows
		;------------------------------------------------------------------------------------------
		; Check to verify that the window object is a valid browser, if not, skip it
		;
		; Setup internal error handler to Trap COM errors, turn off error notification,
		;     check object property validity, set a flag and reset error handler and notification
		;
		$bIsBrowser = True
		; Trap COM errors and turn off error notification
		$bStatus = __ChromeInternalErrorHandlerRegister()
		If Not $bStatus Then __ChromeConsoleWriteError("Warning", "_ChromeAttach", _
				"Cannot register internal error handler, cannot trap COM errors", _
				"Use _ChromeErrorHandlerRegister() to register a user error handler")
		; Turn off error notification for internal processing
		$iNotifyStatus = _ChromeErrorNotify() ; save current error notify status
		_ChromeErrorNotify(False)

		; Check conditions to verify that the object is a browser
		If $bIsBrowser Then
			$sTmp = $oWindow.type ; Is .type a valid property?
			If @error Then $bIsBrowser = False
		EndIf
		If $bIsBrowser Then
			$sTmp = $oWindow.document.title ; Does object have a .document and .title property?
			If @error Then $bIsBrowser = False
		EndIf

		; restore error notify
		_ChromeErrorNotify($iNotifyStatus) ; restore notification status
		__ChromeInternalErrorHandlerDeRegister()
		;------------------------------------------------------------------------------------------

		If $bIsBrowser Then
			Switch $sMode
				Case "title"
					If StringInStr($oWindow.document.title, $sString) > 0 Then
						If $iInstance = $iTmp Then
							Return SetError($_ChromeSTATUS_Success, 0, $oWindow)
						Else
							$iTmp += 1
						EndIf
					EndIf
				Case "instance"
					If $iInstance = $iTmp Then
						Return SetError($_ChromeSTATUS_Success, 0, $oWindow)
					Else
						$iTmp += 1
					EndIf
				Case "windowtitle"
					Local $bFound = False
					$sTmp = RegRead("HKEY_CURRENT_USER\Software\Microsoft\Chrome\Main\", "Window Title")
					If Not @error Then
						If StringInStr($oWindow.document.title & " - " & $sTmp, $sString) Then $bFound = True
					Else
						If StringInStr($oWindow.document.title & " - Microsoft Chrome", $sString) Then $bFound = True
						If StringInStr($oWindow.document.title & " - Windows Chrome", $sString) Then $bFound = True
					EndIf
					If $bFound Then
						If $iInstance = $iTmp Then
							Return SetError($_ChromeSTATUS_Success, 0, $oWindow)
						Else
							$iTmp += 1
						EndIf
					EndIf
				Case "url"
					If StringInStr($oWindow.LocationURL, $sString) > 0 Then
						If $iInstance = $iTmp Then
							Return SetError($_ChromeSTATUS_Success, 0, $oWindow)
						Else
							$iTmp += 1
						EndIf
					EndIf
				Case "text"
					If StringInStr($oWindow.document.body.innerText, $sString) > 0 Then
						If $iInstance = $iTmp Then
							Return SetError($_ChromeSTATUS_Success, 0, $oWindow)
						Else
							$iTmp += 1
						EndIf
					EndIf
				Case "html"
					If StringInStr($oWindow.document.body.innerHTML, $sString) > 0 Then
						If $iInstance = $iTmp Then
							Return SetError($_ChromeSTATUS_Success, 0, $oWindow)
						Else
							$iTmp += 1
						EndIf
					EndIf
				Case "hwnd"
					If $iInstance > 1 Then
						$iInstance = 1
						__ChromeConsoleWriteError("Warning", "_ChromeAttach", "$_ChromeSTATUS_GeneralError", "$iInstance > 1 invalid with HWnd.  Setting to 1.")
					EndIf
					If _ChromePropertyGet($oWindow, "hwnd") = $sString Then
						Return SetError($_ChromeSTATUS_Success, 0, $oWindow)
					EndIf
				Case Else
					; Invalid Mode
					__ChromeConsoleWriteError("Error", "_ChromeAttach", "$_ChromeSTATUS_InvalidValue", "Invalid Mode SpecifChromed")
					Return SetError($_ChromeSTATUS_InvalidValue, 2, 0)
			EndSwitch
		EndIf
	Next
	__ChromeConsoleWriteError("Warning", "_ChromeAttach", "$_ChromeSTATUS_NoMatch")
	Return SetError($_ChromeSTATUS_NoMatch, 1, 0)
EndFunc   ;==>_ChromeAttach

; #FUNCTION# ====================================================================================================================
; Author ........: Dale Hohm
; ModifChromed ......: jpm
; ===============================================================================================================================
Func _ChromeLoadWait(ByRef $oObject, $iDelay = 0, $iTimeout = -1)
	If Not IsObj($oObject) Then
		__ChromeConsoleWriteError("Error", "_ChromeLoadWait", "$_ChromeSTATUS_InvalidDataType")
		Return SetError($_ChromeSTATUS_InvalidDataType, 1, 0)
	EndIf

	If Not __ChromeIsObjType($oObject, "browserdom") Then
		__ChromeConsoleWriteError("Error", "_ChromeLoadWait", "$_ChromeSTATUS_InvalidObjectType", ObjName($oObject))
		Return SetError($_ChromeSTATUS_InvalidObjectType, 1, 0)
	EndIf

	Local $oTemp, $bAbort = False, $ChromeErrorStatusCode = $_ChromeSTATUS_Success

	; Setup internal error handler to Trap COM errors, turn off error notification
	Local $bStatus = __ChromeInternalErrorHandlerRegister()
	If Not $bStatus Then __ChromeConsoleWriteError("Warning", "_ChromeLoadWait", _
			"Cannot register internal error handler, cannot trap COM errors", _
			"Use _ChromeErrorHandlerRegister() to register a user error handler")
	Local $iNotifyStatus = _ChromeErrorNotify() ; save current error notify status
	_ChromeErrorNotify(False)

	Sleep($iDelay)
	;
	Local $ChromeError
	Local $hChromeLoadWaitTimer = TimerInit()
	If $iTimeout = -1 Then $iTimeout = $__g_iChromeLoadWaitTimeout

	Select
		Case __ChromeIsObjType($oObject, "browser", False); Chrome
			While Not (String($oObject.readyState) = "complete" Or $oObject.readyState = 4 Or $bAbort)
				; Trap unrecoverable COM errors
				If @error Then
					$ChromeError = @error
					If __ChromeComErrorUnrecoverable($ChromeError) Then
						$ChromeErrorStatusCode = __ChromeComErrorUnrecoverable($ChromeError)
						$bAbort = True
					EndIf
				ElseIf (TimerDiff($hChromeLoadWaitTimer) > $iTimeout) Then
					$ChromeErrorStatusCode = $_ChromeSTATUS_LoadWaitTimeout
					$bAbort = True
				EndIf
				Sleep(100)
			WEnd
			While Not (String($oObject.document.readyState) = "complete" Or $oObject.document.readyState = 4 Or $bAbort)
				; Trap unrecoverable COM errors
				If @error Then
					$ChromeError = @error
					If __ChromeComErrorUnrecoverable($ChromeError) Then
						$ChromeErrorStatusCode = __ChromeComErrorUnrecoverable($ChromeError)
						$bAbort = True
					EndIf
				ElseIf (TimerDiff($hChromeLoadWaitTimer) > $iTimeout) Then
					$ChromeErrorStatusCode = $_ChromeSTATUS_LoadWaitTimeout
					$bAbort = True
				EndIf
				Sleep(100)
			WEnd
		Case __ChromeIsObjType($oObject, "window", False) ; Window, Frame, iFrame
			While Not (String($oObject.document.readyState) = "complete" Or $oObject.document.readyState = 4 Or $bAbort)
				; Trap unrecoverable COM errors
				If @error Then
					$ChromeError = @error
					If __ChromeComErrorUnrecoverable($ChromeError) Then
						$ChromeErrorStatusCode = __ChromeComErrorUnrecoverable($ChromeError)
						$bAbort = True
					EndIf
				ElseIf (TimerDiff($hChromeLoadWaitTimer) > $iTimeout) Then
					$ChromeErrorStatusCode = $_ChromeSTATUS_LoadWaitTimeout
					$bAbort = True
				EndIf
				Sleep(100)
			WEnd
			While Not (String($oObject.top.document.readyState) = "complete" Or $oObject.top.document.readyState = 4 Or $bAbort)
				; Trap unrecoverable COM errors
				If @error Then
					$ChromeError = @error
					If __ChromeComErrorUnrecoverable($ChromeError) Then
						$ChromeErrorStatusCode = __ChromeComErrorUnrecoverable($ChromeError)
						$bAbort = True
					EndIf
				ElseIf (TimerDiff($hChromeLoadWaitTimer) > $iTimeout) Then
					$ChromeErrorStatusCode = $_ChromeSTATUS_LoadWaitTimeout
					$bAbort = True
				EndIf
				Sleep(100)
			WEnd
		Case __ChromeIsObjType($oObject, "document", False) ; Document
			$oTemp = $oObject.parentWindow
			While Not (String($oTemp.document.readyState) = "complete" Or $oTemp.document.readyState = 4 Or $bAbort)
				; Trap unrecoverable COM errors
				If @error Then
					$ChromeError = @error
					If __ChromeComErrorUnrecoverable($ChromeError) Then
						$ChromeErrorStatusCode = __ChromeComErrorUnrecoverable($ChromeError)
						$bAbort = True
					EndIf
				ElseIf (TimerDiff($hChromeLoadWaitTimer) > $iTimeout) Then
					$ChromeErrorStatusCode = $_ChromeSTATUS_LoadWaitTimeout
					$bAbort = True
				EndIf
				Sleep(100)
			WEnd
			While Not (String($oTemp.top.document.readyState) = "complete" Or $oTemp.top.document.readyState = 4 Or $bAbort)
				; Trap unrecoverable COM errors
				If @error Then
					$ChromeError = @error
					If __ChromeComErrorUnrecoverable($ChromeError) Then
						$ChromeErrorStatusCode = __ChromeComErrorUnrecoverable($ChromeError)
						$bAbort = True
					EndIf
				ElseIf (TimerDiff($hChromeLoadWaitTimer) > $iTimeout) Then
					$ChromeErrorStatusCode = $_ChromeSTATUS_LoadWaitTimeout
					$bAbort = True
				EndIf
				Sleep(100)
			WEnd
		Case Else ; this should work with any other DOM object
			$oTemp = $oObject.document.parentWindow
			While Not (String($oTemp.document.readyState) = "complete" Or $oTemp.document.readyState = 4 Or $bAbort)
				; Trap unrecoverable COM errors
				If @error Then
					$ChromeError = @error
					If __ChromeComErrorUnrecoverable($ChromeError) Then
						$ChromeErrorStatusCode = __ChromeComErrorUnrecoverable($ChromeError)
						$bAbort = True
					EndIf
				ElseIf (TimerDiff($hChromeLoadWaitTimer) > $iTimeout) Then
					$ChromeErrorStatusCode = $_ChromeSTATUS_LoadWaitTimeout
					$bAbort = True
				EndIf
				Sleep(100)
			WEnd
			While Not (String($oTemp.top.document.readyState) = "complete" Or $oObject.top.document.readyState = 4 Or $bAbort)
				; Trap unrecoverable COM errors
				If @error Then
					$ChromeError = @error
					If __ChromeComErrorUnrecoverable($ChromeError) Then
						$ChromeErrorStatusCode = __ChromeComErrorUnrecoverable($ChromeError)
						$bAbort = True
					EndIf
				ElseIf (TimerDiff($hChromeLoadWaitTimer) > $iTimeout) Then
					$ChromeErrorStatusCode = $_ChromeSTATUS_LoadWaitTimeout
					$bAbort = True
				EndIf
				Sleep(100)
			WEnd
	EndSelect

	; restore error notify
	_ChromeErrorNotify($iNotifyStatus) ; restore notification status
	__ChromeInternalErrorHandlerDeRegister()

	Switch $ChromeErrorStatusCode
		Case $_ChromeSTATUS_Success
			Return SetError($_ChromeSTATUS_Success, 0, 1)
		Case $_ChromeSTATUS_LoadWaitTimeout
			__ChromeConsoleWriteError("Warning", "_ChromeLoadWait", "$_ChromeSTATUS_LoadWaitTimeout")
			Return SetError($_ChromeSTATUS_LoadWaitTimeout, 3, 0)
		Case $_ChromeSTATUS_AccessIsDenChromed
			__ChromeConsoleWriteError("Warning", "_ChromeLoadWait", "$_ChromeSTATUS_AccessIsDenChromed", _
					"Cannot verify readyState.  Likely casue: cross-domain scripting security restriction. (" & $ChromeError & ")")
			Return SetError($_ChromeSTATUS_AccessIsDenChromed, 0, 0)
		Case $_ChromeSTATUS_ClChromentDisconnected
			__ChromeConsoleWriteError("Error", "_ChromeLoadWait", "$_ChromeSTATUS_ClChromentDisconnected", _
					$ChromeError & ", Browser has been deleted prior to operation.")
			Return SetError($_ChromeSTATUS_ClChromentDisconnected, 0, 0)
		Case Else
			__ChromeConsoleWriteError("Error", "_ChromeLoadWait", "$_ChromeSTATUS_GeneralError", "Invalid Error Status - Notify Chrome.au3 developer")
			Return SetError($_ChromeSTATUS_GeneralError, 0, 0)
	EndSwitch
EndFunc   ;==>_ChromeLoadWait

; #FUNCTION# ====================================================================================================================
; Author ........: Dale Hohm
; ===============================================================================================================================
Func _ChromeLoadWaitTimeout($iTimeout = -1)
	If $iTimeout = -1 Then
		Return SetError($_ChromeSTATUS_Success, 0, $__g_iChromeLoadWaitTimeout)
	Else
		$__g_iChromeLoadWaitTimeout = $iTimeout
		Return SetError($_ChromeSTATUS_Success, 0, 1)
	EndIf
EndFunc   ;==>_ChromeLoadWaitTimeout

#EndRegion Core functions

#Region Frame Functions
; Security Note on Frame functions:
; Note that security restriction in Chrome related to cross-site scripting
; between frames can cause serious problems with the frame functions.  Functions that
; work connected to one site will fail when connected to another depending on the sites
; referenced in the frames.  In general, if all the referenced pages are on the same
; webserver these functions should work as described; if not, unexpected COM failures
; can occur.
; #FUNCTION# ====================================================================================================================
; Author ........: Dale Hohm
; ===============================================================================================================================
Func _ChromeIsFrameSet(ByRef $oObject)
	; Note: this is more reliable test for a FrameSet than checking the
	; number of frames (document.frames.length) because iFrames embedded on a normal
	; page are included in the frame collection even though it is not a FrameSet
	If Not IsObj($oObject) Then
		__ChromeConsoleWriteError("Error", "_ChromeIsFrameSet", "$_ChromeSTATUS_InvalidDataType")
		Return SetError($_ChromeSTATUS_InvalidDataType, 1, 0)
	EndIf
	;
	If String($oObject.document.body.tagName) = "frameset" Then
		Return SetError($_ChromeSTATUS_Success, 0, 1)
	Else
		If @error Then ; Trap COM error, report and return
			__ChromeConsoleWriteError("Error", "_ChromeIsFrameSet", "$_ChromeSTATUS_COMError", @error)
			Return SetError($_ChromeSTATUS_ComError, @error, 0)
		EndIf
		Return SetError($_ChromeSTATUS_Success, 0, 0)
	EndIf
EndFunc   ;==>_ChromeIsFrameSet

; #FUNCTION# ====================================================================================================================
; Author ........: Dale Hohm
; ===============================================================================================================================
Func _ChromeFrameGetCollection(ByRef $oObject, $iIndex = -1)
	If Not IsObj($oObject) Then
		__ChromeConsoleWriteError("Error", "_ChromeFrameGetCollection", "$_ChromeSTATUS_InvalidDataType")
		Return SetError($_ChromeSTATUS_InvalidDataType, 1, 0)
	EndIf
	;
	$iIndex = Number($iIndex)
	Select
		Case $iIndex = -1
			Return SetError($_ChromeSTATUS_Success, $oObject.document.parentwindow.frames.length, _
					$oObject.document.parentwindow.frames)
		Case $iIndex > -1 And $iIndex < $oObject.document.parentwindow.frames.length
			Return SetError($_ChromeSTATUS_Success, $oObject.document.parentwindow.frames.length, _
					$oObject.document.parentwindow.frames.item($iIndex))
		Case $iIndex < -1
			__ChromeConsoleWriteError("Error", "_ChromeFrameGetCollection", "$_ChromeSTATUS_InvalidValue", "$iIndex < -1")
			Return SetError($_ChromeSTATUS_InvalidValue, 2, 0)
		Case Else
			__ChromeConsoleWriteError("Warning", "_ChromeFrameGetCollection", "$_ChromeSTATUS_NoMatch")
			Return SetError($_ChromeSTATUS_NoMatch, 2, 0)
	EndSelect
EndFunc   ;==>_ChromeFrameGetCollection

; #FUNCTION# ====================================================================================================================
; Author ........: Dale Hohm
; ===============================================================================================================================
Func _ChromeFrameGetObjByName(ByRef $oObject, $sName)
	If Not IsObj($oObject) Then
		__ChromeConsoleWriteError("Error", "_ChromeFrameGetObjByName", "$_ChromeSTATUS_InvalidDataType")
		Return SetError($_ChromeSTATUS_InvalidDataType, 1, 0)
	EndIf
	;
	Local $oTemp, $oFrames

	If Not __ChromeIsObjType($oObject, "browserdom") Then
		__ChromeConsoleWriteError("Error", "_ChromeFrameGetObjByName", "$_ChromeSTATUS_InvalidObjectType")
		Return SetError($_ChromeSTATUS_InvalidObjectType, 1, 0)
	EndIf

	If __ChromeIsObjType($oObject, "document") Then
		$oTemp = $oObject.parentWindow
	Else
		$oTemp = $oObject.document.parentWindow
	EndIf

	If _ChromeIsFrameSet($oTemp) Then
		$oFrames = _ChromeTagNameGetCollection($oTemp, "frame")
	Else
		$oFrames = _ChromeTagNameGetCollection($oTemp, "iframe")
	EndIf

	If $oFrames.length Then
		For $oFrame In $oFrames
			If String($oFrame.name) = $sName Then Return SetError($_ChromeSTATUS_Success, 0, $oTemp.frames($sName))
		Next
		__ChromeConsoleWriteError("Warning", "_ChromeFrameGetObjByName", "$_ChromeSTATUS_NoMatch", "No frames matching name")
		Return SetError($_ChromeSTATUS_NoMatch, 2, 0)
	Else
		__ChromeConsoleWriteError("Warning", "_ChromeFrameGetObjByName", "$_ChromeSTATUS_NoMatch", "No Frames found")
		Return SetError($_ChromeSTATUS_NoMatch, 2, 0)
	EndIf
EndFunc   ;==>_ChromeFrameGetObjByName

#EndRegion Frame Functions

#Region Link functions
; #FUNCTION# ====================================================================================================================
; Author ........: Dale Hohm
; ===============================================================================================================================
Func _ChromeLinkClickByText(ByRef $oObject, $sLinkText, $iIndex = 0, $iWait = 1)
	If Not IsObj($oObject) Then
		__ChromeConsoleWriteError("Error", "_ChromeLinkClickByText", "$_ChromeSTATUS_InvalidDataType")
		Return SetError($_ChromeSTATUS_InvalidDataType, 1, 0)
	EndIf
	;
	Local $iFound = 0, $sModeLinktext, $oLinks = $oObject.document.links
	$iIndex = Number($iIndex)
	For $oLink In $oLinks
		$sModeLinktext = String($oLink.outerText)
		If $sModeLinktext = $sLinkText Then
			If ($iFound = $iIndex) Then
				$oLink.click()
				If @error Then ; Trap COM error, report and return
					__ChromeConsoleWriteError("Error", "_ChromeLinkClickByText", "$_ChromeSTATUS_COMError", @error)
					Return SetError($_ChromeSTATUS_ComError, @error, 0)
				EndIf
				If $iWait Then
					_ChromeLoadWait($oObject)
					Return SetError(@error, 0, -1)
				EndIf
				Return SetError($_ChromeSTATUS_Success, 0, -1)
			EndIf
			$iFound = $iFound + 1
		EndIf
	Next
	__ChromeConsoleWriteError("Warning", "_ChromeLinkClickByText", "$_ChromeSTATUS_NoMatch")
	Return SetError($_ChromeSTATUS_NoMatch, 0, 0) ; Could be caused by parameter 2, 3 or both
EndFunc   ;==>_ChromeLinkClickByText

; #FUNCTION# ====================================================================================================================
; Author ........: Dale Hohm
; ===============================================================================================================================
Func _ChromeLinkClickByIndex(ByRef $oObject, $iIndex, $iWait = 1)
	If Not IsObj($oObject) Then
		__ChromeConsoleWriteError("Error", "_ChromeLinkClickByIndex", "$_ChromeSTATUS_InvalidDataType")
		Return SetError($_ChromeSTATUS_InvalidDataType, 1, 0)
	EndIf
	;
	Local $oLinks = $oObject.document.links, $oLink
	$iIndex = Number($iIndex)
	If ($iIndex >= 0) And ($iIndex <= $oLinks.length - 1) Then
		$oLink = $oLinks($iIndex)
		$oLink.click()
		If @error Then ; Trap COM error, report and return
			__ChromeConsoleWriteError("Error", "_ChromeLinkClickByIndex", "$_ChromeSTATUS_COMError", @error)
			Return SetError($_ChromeSTATUS_ComError, @error, 0)
		EndIf
		If $iWait Then
			_ChromeLoadWait($oObject)
			Return SetError(@error, 0, -1)
		EndIf
		Return SetError($_ChromeSTATUS_Success, 0, -1)
	Else
		__ChromeConsoleWriteError("Warning", "_ChromeLinkClickByIndex", "$_ChromeSTATUS_NoMatch")
		Return SetError($_ChromeSTATUS_NoMatch, 2, 0)
	EndIf
EndFunc   ;==>_ChromeLinkClickByIndex

; #FUNCTION# ====================================================================================================================
; Author ........: Dale Hohm
; ===============================================================================================================================
Func _ChromeLinkGetCollection(ByRef $oObject, $iIndex = -1)
	If Not IsObj($oObject) Then
		__ChromeConsoleWriteError("Error", "_ChromeLinkGetCollection", "$_ChromeSTATUS_InvalidDataType")
		Return SetError($_ChromeSTATUS_InvalidDataType, 1, 0)
	EndIf
	;
	$iIndex = Number($iIndex)
	Select
		Case $iIndex = -1
			Return SetError($_ChromeSTATUS_Success, $oObject.document.links.length, _
					$oObject.document.links)
		Case $iIndex > -1 And $iIndex < $oObject.document.links.length
			Return SetError($_ChromeSTATUS_Success, $oObject.document.links.length, _
					$oObject.document.links.item($iIndex))
		Case $iIndex < -1
			__ChromeConsoleWriteError("Error", "_ChromeLinkGetCollection", "$_ChromeSTATUS_InvalidValue")
			Return SetError($_ChromeSTATUS_InvalidValue, 2, 0)
		Case Else
			__ChromeConsoleWriteError("Warning", "_ChromeLinkGetCollection", "$_ChromeSTATUS_NoMatch")
			Return SetError($_ChromeSTATUS_NoMatch, 2, 0)
	EndSelect
EndFunc   ;==>_ChromeLinkGetCollection
#EndRegion Link functions

#Region Image functions
; #FUNCTION# ====================================================================================================================
; Author ........: Dale Hohm
; ModifChromed ......: jpm
; ===============================================================================================================================
Func _ChromeImgClick(ByRef $oObject, $sLinkText, $sMode = "src", $iIndex = 0, $iWait = 1)
	If Not IsObj($oObject) Then
		__ChromeConsoleWriteError("Error", "_ChromeImgClick", "$_ChromeSTATUS_InvalidDataType")
		Return SetError($_ChromeSTATUS_InvalidDataType, 1, 0)
	EndIf
	;
	Local $sModeLinktext, $iFound = 0, $oImgs = $oObject.document.images
	$sMode = StringLower($sMode)
	$iIndex = Number($iIndex)
	For $oImg In $oImgs
		Select
			Case $sMode = "alt"
				$sModeLinktext = $oImg.alt
			Case $sMode = "name"
				$sModeLinktext = $oImg.name
				If Not IsString($sModeLinktext) Then $sModeLinktext = $oImg.id ; html5 support
			Case $sMode = "id"
				$sModeLinktext = $oImg.id
			Case $sMode = "src"
				$sModeLinktext = $oImg.src
			Case Else
				__ChromeConsoleWriteError("Error", "_ChromeImgClick", "$_ChromeSTATUS_InvalidValue", "Invalid mode: " & $sMode)
				Return SetError($_ChromeSTATUS_InvalidValue, 3, 0)
		EndSelect
		If StringInStr($sModeLinktext, $sLinkText) Then
			If ($iFound = $iIndex) Then
				$oImg.click()
				If @error Then ; Trap COM error, report and return
					__ChromeConsoleWriteError("Error", "_ChromeImgClick", "$_ChromeSTATUS_COMError", @error)
					Return SetError($_ChromeSTATUS_ComError, @error, 0)
				EndIf
				If $iWait Then
					_ChromeLoadWait($oObject)
					Return SetError(@error, 0, -1)
				EndIf
				Return SetError($_ChromeSTATUS_Success, 0, -1)
			EndIf
			$iFound = $iFound + 1
		EndIf
	Next
	__ChromeConsoleWriteError("Warning", "_ChromeImgClick", "$_ChromeSTATUS_NoMatch")
	Return SetError($_ChromeSTATUS_NoMatch, 0, 0) ; Could be caused by parameter 2, 4 or both
EndFunc   ;==>_ChromeImgClick

; #FUNCTION# ====================================================================================================================
; Author ........: Dale Hohm
; ===============================================================================================================================
Func _ChromeImgGetCollection(ByRef $oObject, $iIndex = -1)
	If Not IsObj($oObject) Then
		__ChromeConsoleWriteError("Error", "_ChromeImgGetCollection", "$_ChromeSTATUS_InvalidDataType")
		Return SetError($_ChromeSTATUS_InvalidDataType, 1, 0)
	EndIf
	;
	Local $oTemp = _ChromeDocGetObj($oObject)
	$iIndex = Number($iIndex)
	Select
		Case $iIndex = -1
			Return SetError($_ChromeSTATUS_Success, $oTemp.images.length, $oTemp.images)
		Case $iIndex > -1 And $iIndex < $oTemp.images.length
			Return SetError($_ChromeSTATUS_Success, $oTemp.images.length, $oTemp.images.item($iIndex))
		Case $iIndex < -1
			__ChromeConsoleWriteError("Error", "_ChromeImgGetCollection", "$_ChromeSTATUS_InvalidValue", "$iIndex < -1")
			Return SetError($_ChromeSTATUS_InvalidValue, 2, 0)
		Case Else
			__ChromeConsoleWriteError("Warning", "_ChromeImgGetCollection", "$_ChromeSTATUS_NoMatch")
			Return SetError($_ChromeSTATUS_NoMatch, 1, 0)
	EndSelect
EndFunc   ;==>_ChromeImgGetCollection

#EndRegion Image functions

#Region Form functions
; #FUNCTION# ====================================================================================================================
; Author ........: Dale Hohm
; ===============================================================================================================================
Func _ChromeFormGetCollection(ByRef $oObject, $iIndex = -1)
	If Not IsObj($oObject) Then
		__ChromeConsoleWriteError("Error", "_ChromeFormGetCollection", "$_ChromeSTATUS_InvalidDataType")
		Return SetError($_ChromeSTATUS_InvalidDataType, 1, 0)
	EndIf
	;
	Local $oTemp = _ChromeDocGetObj($oObject)
	$iIndex = Number($iIndex)
	Select
		Case $iIndex = -1
			Return SetError($_ChromeSTATUS_Success, $oTemp.forms.length, $oTemp.forms)
		Case $iIndex > -1 And $iIndex < $oTemp.forms.length
			Return SetError($_ChromeSTATUS_Success, $oTemp.forms.length, $oTemp.forms.item($iIndex))
		Case $iIndex < -1
			__ChromeConsoleWriteError("Error", "_ChromeFormGetCollection", "$_ChromeSTATUS_InvalidValue", "$iIndex < -1")
			Return SetError($_ChromeSTATUS_InvalidValue, 2, 0)
		Case Else
			__ChromeConsoleWriteError("Warning", "_ChromeFormGetCollection", "$_ChromeSTATUS_NoMatch")
			Return SetError($_ChromeSTATUS_NoMatch, 1, 0)
	EndSelect
EndFunc   ;==>_ChromeFormGetCollection

; #FUNCTION# ====================================================================================================================
; Author ........: Dale Hohm
; ===============================================================================================================================
Func _ChromeFormGetObjByName(ByRef $oObject, $sName, $iIndex = 0)
	If Not IsObj($oObject) Then
		__ChromeConsoleWriteError("Error", "_ChromeFormGetObjByName", "$_ChromeSTATUS_InvalidDataType")
		Return SetError($_ChromeSTATUS_InvalidDataType, 1, 0)
	EndIf
	;
	;----- Determine valid collection length
	Local $iLength = 0
	Local $oCol = $oObject.document.forms.item($sName)
	If IsObj($oCol) Then
		If __ChromeIsObjType($oCol, "elementcollection") Then
			$iLength = $oCol.length
		Else
			$iLength = 1
		EndIf
	EndIf
	;-----
	$iIndex = Number($iIndex)
	If $iIndex = -1 Then
		Return SetError($_ChromeSTATUS_Success, $iLength, $oObject.document.forms.item($sName))
	Else
		If IsObj($oObject.document.forms.item($sName, $iIndex)) Then
			Return SetError($_ChromeSTATUS_Success, $iLength, $oObject.document.forms.item($sName, $iIndex))
		Else
			__ChromeConsoleWriteError("Warning", "_ChromeFormGetObjByName", "$_ChromeSTATUS_NoMatch")
			Return SetError($_ChromeSTATUS_NoMatch, 0, 0) ; Could be caused by parameter 2, 3 or both
		EndIf
	EndIf
EndFunc   ;==>_ChromeFormGetObjByName

; #FUNCTION# ====================================================================================================================
; Author ........: Dale Hohm
; ===============================================================================================================================
Func _ChromeFormElementGetCollection(ByRef $oObject, $iIndex = -1)
	If Not IsObj($oObject) Then
		__ChromeConsoleWriteError("Error", "_ChromeFormElementGetCollection", "$_ChromeSTATUS_InvalidDataType")
		Return SetError($_ChromeSTATUS_InvalidDataType, 1, 0)
	EndIf
	;
	If Not __ChromeIsObjType($oObject, "form") Then
		__ChromeConsoleWriteError("Error", "_ChromeFormElementGetCollection", "$_ChromeSTATUS_InvalidObjectType")
		Return SetError($_ChromeSTATUS_InvalidObjectType, 1, 0)
	EndIf
	;
	$iIndex = Number($iIndex)
	Select
		Case $iIndex = -1
			Return SetError($_ChromeSTATUS_Success, $oObject.elements.length, $oObject.elements)
		Case $iIndex > -1 And $iIndex < $oObject.elements.length
			Return SetError($_ChromeSTATUS_Success, $oObject.elements.length, $oObject.elements.item($iIndex))
		Case $iIndex < -1
			__ChromeConsoleWriteError("Error", "_ChromeFormElementGetCollection", "$_ChromeSTATUS_InvalidValue", "$iIndex < -1")
			Return SetError($_ChromeSTATUS_InvalidValue, 2, 0)
		Case Else
			Return SetError($_ChromeSTATUS_NoMatch, 1, 0)
	EndSelect
EndFunc   ;==>_ChromeFormElementGetCollection

; #FUNCTION# ====================================================================================================================
; Author ........: Dale Hohm
; ===============================================================================================================================
Func _ChromeFormElementGetObjByName(ByRef $oObject, $sName, $iIndex = 0)
	If Not IsObj($oObject) Then
		__ChromeConsoleWriteError("Error", "_ChromeFormElementGetObjByName", "$_ChromeSTATUS_InvalidDataType")
		Return SetError($_ChromeSTATUS_InvalidDataType, 1, 0)
	EndIf
	;
	If Not __ChromeIsObjType($oObject, "form") Then
		__ChromeConsoleWriteError("Error", "_ChromeFormElementGetObjByName", "$_ChromeSTATUS_InvalidObjectType")
		Return SetError($_ChromeSTATUS_InvalidObjectType, 1, 0)
	EndIf
	;
	;----- Determine valid collection length
	Local $iLength = 0
	Local $oCol = $oObject.elements.item($sName)
	If IsObj($oCol) Then
		If __ChromeIsObjType($oCol, "elementcollection") Then
			$iLength = $oCol.length
		Else
			$iLength = 1
		EndIf
	EndIf
	;-----
	$iIndex = Number($iIndex)
	If $iIndex = -1 Then
		Return SetError($_ChromeSTATUS_Success, $iLength, $oObject.elements.item($sName))
	Else
		If IsObj($oObject.elements.item($sName, $iIndex)) Then
			Return SetError($_ChromeSTATUS_Success, $iLength, $oObject.elements.item($sName, $iIndex))
		Else
			__ChromeConsoleWriteError("Warning", "_ChromeFormElementGetObjByName", "$_ChromeSTATUS_NoMatch")
			Return SetError($_ChromeSTATUS_NoMatch, 0, 0) ; Could be caused by parameter 2, 3 or both
		EndIf
	EndIf
EndFunc   ;==>_ChromeFormElementGetObjByName

; #FUNCTION# ====================================================================================================================
; Author ........: Dale Hohm
; ===============================================================================================================================
Func _ChromeFormElementGetValue(ByRef $oObject)
	If Not IsObj($oObject) Then
		__ChromeConsoleWriteError("Error", "_ChromeFormElementGetValue", "$_ChromeSTATUS_InvalidDataType")
		Return SetError($_ChromeSTATUS_InvalidDataType, 1, 0)
	EndIf
	;
	If Not __ChromeIsObjType($oObject, "forminputelement") Then
		__ChromeConsoleWriteError("Error", "_ChromeFormElementGetValue", "$_ChromeSTATUS_InvalidObjectType")
		Return SetError($_ChromeSTATUS_InvalidObjectType, 1, 0)
	EndIf
	;
	Local $sReturn = String($oObject.value)
	If @error Then ; Trap COM error, report and return
		__ChromeConsoleWriteError("Error", "_ChromeFormElementGetValue", "$_ChromeSTATUS_COMError", @error)
		Return SetError($_ChromeSTATUS_ComError, @error, 0)
	EndIf
	SetError($_ChromeSTATUS_Success)
	Return $sReturn
EndFunc   ;==>_ChromeFormElementGetValue

; #FUNCTION# ====================================================================================================================
; Author ........: Dale Hohm
; ===============================================================================================================================
Func _ChromeFormElementSetValue(ByRef $oObject, $sNewValue, $iFireEvent = 1)
	If Not IsObj($oObject) Then
		__ChromeConsoleWriteError("Error", "_ChromeFormElementSetValue", "$_ChromeSTATUS_InvalidDataType")
		Return SetError($_ChromeSTATUS_InvalidDataType, 1, 0)
	EndIf
	;
	If Not __ChromeIsObjType($oObject, "forminputelement") Then
		__ChromeConsoleWriteError("Error", "_ChromeFormElementSetValue", "$_ChromeSTATUS_InvalidObjectType")
		Return SetError($_ChromeSTATUS_InvalidObjectType, 1, 0)
	EndIf
	;
	If String($oObject.type) = "file" Then
		__ChromeConsoleWriteError("Error", "_ChromeFormElementSetValue", "$_ChromeSTATUS_InvalidObjectType", "Browser security prevents SetValue of TYPE=FILE")
		Return SetError($_ChromeSTATUS_InvalidObjectType, 1, 0)
	EndIf
	;
	$oObject.value = $sNewValue
	If @error Then ; Trap COM error, report and return
		__ChromeConsoleWriteError("Error", "_ChromeFormElementSetValue", "$_ChromeSTATUS_COMError", @error)
		Return SetError($_ChromeSTATUS_ComError, @error, 0)
	EndIf
	If $iFireEvent Then
		$oObject.fireEvent("OnChange")
		$oObject.fireEvent("OnClick")
	EndIf
	Return SetError($_ChromeSTATUS_Success, 0, 1)
EndFunc   ;==>_ChromeFormElementSetValue

; #FUNCTION# ====================================================================================================================
; Author ........: Dale Hohm
; ===============================================================================================================================
Func _ChromeFormElementOptionSelect(ByRef $oObject, $sString, $iSelect = 1, $sMode = "byValue", $iFireEvent = 1)
	If Not IsObj($oObject) Then
		__ChromeConsoleWriteError("Error", "_ChromeFormElementOptionSelect", "$_ChromeSTATUS_InvalidDataType")
		Return SetError($_ChromeSTATUS_InvalidDataType, 1, 0)
	EndIf
	;
	If Not __ChromeIsObjType($oObject, "formselectelement") Then
		__ChromeConsoleWriteError("Error", "_ChromeFormElementOptionSelect", "$_ChromeSTATUS_InvalidObjectType")
		Return SetError($_ChromeSTATUS_InvalidObjectType, 1, 0)
	EndIf
	;
	Local $oItem, $oItems = $oObject.options, $iNumItems = $oObject.options.length, $bIsMultiple = $oObject.multiple

	Switch $sMode
		Case "byValue"
			For $oItem In $oItems
				If $oItem.value = $sString Then
					Switch $iSelect
						Case -1
							Return SetError($_ChromeSTATUS_Success, 0, $oItem.selected)
						Case 0
							If Not $bIsMultiple Then
								__ChromeConsoleWriteError("Error", "_ChromeFormElementOptionSelect", "$_ChromeSTATUS_InvalidValue", _
										"$iSelect=0 only valid for type=select multiple")
								SetError($_ChromeSTATUS_InvalidValue, 3)
							EndIf
							If $oItem.selected Then
								$oItem.selected = False
								If $iFireEvent Then
									$oObject.fireEvent("onChange")
									$oObject.fireEvent("OnClick")
								EndIf
							EndIf
							Return SetError($_ChromeSTATUS_Success, 0, 1)
						Case 1
							If Not $oItem.selected Then
								$oItem.selected = True
								If $iFireEvent Then
									$oObject.fireEvent("onChange")
									$oObject.fireEvent("OnClick")
								EndIf
							EndIf
							Return SetError($_ChromeSTATUS_Success, 0, 1)
						Case Else
							__ChromeConsoleWriteError("Error", "_ChromeFormElementOptionSelect", "$_ChromeSTATUS_InvalidValue", "Invalid $iSelect value")
							Return SetError($_ChromeSTATUS_InvalidValue, 3, 0)
					EndSwitch
					__ChromeConsoleWriteError("Warning", "_ChromeFormElementOptionSelect", "$_ChromeSTATUS_NoMatch", "Value not matched")
					Return SetError($_ChromeSTATUS_NoMatch, 2, 0)
				EndIf
			Next
		Case "byText"
			For $oItem In $oItems
				If String($oItem.text) = $sString Then
					Switch $iSelect
						Case -1
							Return SetError($_ChromeSTATUS_Success, 0, $oItem.selected)
						Case 0
							If Not $bIsMultiple Then
								__ChromeConsoleWriteError("Error", "_ChromeFormElementOptionSelect", "$_ChromeSTATUS_InvalidValue", _
										"$iSelect=0 only valid for type=select multiple")
								SetError($_ChromeSTATUS_InvalidValue, 3)
							EndIf
							If $oItem.selected Then
								$oItem.selected = False
								If $iFireEvent Then
									$oObject.fireEvent("onChange")
									$oObject.fireEvent("OnClick")
								EndIf
							EndIf
							Return SetError($_ChromeSTATUS_Success, 0, 1)
						Case 1
							If Not $oItem.selected Then
								$oItem.selected = True
								If $iFireEvent Then
									$oObject.fireEvent("onChange")
									$oObject.fireEvent("OnClick")
								EndIf
							EndIf
							Return SetError($_ChromeSTATUS_Success, 0, 1)
						Case Else
							__ChromeConsoleWriteError("Error", "_ChromeFormElementOptionSelect", "$_ChromeSTATUS_InvalidValue", "Invalid $iSelect value")
							Return SetError($_ChromeSTATUS_InvalidValue, 3, 0)
					EndSwitch
					__ChromeConsoleWriteError("Warning", "_ChromeFormElementOptionSelect", "$_ChromeSTATUS_NoMatch", "Text not matched")
					Return SetError($_ChromeSTATUS_NoMatch, 2, 0)
				EndIf
			Next
		Case "byIndex"
			Local $iIndex = Number($sString)
			If $iIndex < 0 Or $iIndex >= $iNumItems Then
				__ChromeConsoleWriteError("Error", "_ChromeFormElementOptionSelect", "$_ChromeSTATUS_InvalidValue", "Invalid index value, " & $iIndex)
				Return SetError($_ChromeSTATUS_InvalidValue, 2, 0)
			EndIf
			$oItem = $oItems.item($iIndex)
			Switch $iSelect
				Case -1
					Return SetError($_ChromeSTATUS_Success, 0, $oItems.item($iIndex).selected)
				Case 0
					If Not $bIsMultiple Then
						__ChromeConsoleWriteError("Error", "_ChromeFormElementOptionSelect", "$_ChromeSTATUS_InvalidValue", _
								"$iSelect=0 only valid for type=select multiple")
						SetError($_ChromeSTATUS_InvalidValue, 3)
					EndIf
					If $oItem.selected Then
						$oItems.item($iIndex).selected = False
						If $iFireEvent Then
							$oObject.fireEvent("onChange")
							$oObject.fireEvent("OnClick")
						EndIf
					EndIf
					Return SetError($_ChromeSTATUS_Success, 0, 1)
				Case 1
					If Not $oItem.selected Then
						$oItems.item($iIndex).selected = True
						If $iFireEvent Then
							$oObject.fireEvent("onChange")
							$oObject.fireEvent("OnClick")
						EndIf
					EndIf
					Return SetError($_ChromeSTATUS_Success, 0, 1)
				Case Else
					__ChromeConsoleWriteError("Error", "_ChromeFormElementOptionSelect", "$_ChromeSTATUS_InvalidValue", "Invalid $iSelect value")
					Return SetError($_ChromeSTATUS_InvalidValue, 3, 0)
			EndSwitch
		Case Else
			__ChromeConsoleWriteError("Error", "_ChromeFormElementOptionSelect", "$_ChromeSTATUS_InvalidValue", "Invalid Mode")
			Return SetError($_ChromeSTATUS_InvalidValue, 4, 0)
	EndSwitch
EndFunc   ;==>_ChromeFormElementOptionSelect

; #FUNCTION# ====================================================================================================================
; Author ........: Dale Hohm
; ===============================================================================================================================
Func _ChromeFormElementCheckBoxSelect(ByRef $oObject, $sString, $sName = "", $iSelect = 1, $sMode = "byValue", $iFireEvent = 1)
	If Not IsObj($oObject) Then
		__ChromeConsoleWriteError("Error", "_ChromeFormElementCheckBoxSelect", "$_ChromeSTATUS_InvalidDataType")
		Return SetError($_ChromeSTATUS_InvalidDataType, 1, 0)
	EndIf
	;
	If Not __ChromeIsObjType($oObject, "form") Then
		__ChromeConsoleWriteError("Error", "_ChromeFormElementCheckBoxSelect", "$_ChromeSTATUS_InvalidObjectType")
		Return SetError($_ChromeSTATUS_InvalidObjectType, 1, 0)
	EndIf
	;
	$sString = String($sString)
	$sName = String($sName)

	Local $oItems
	If $sName = "" Then
		$oItems = _ChromeTagNameGetCollection($oObject, "input")
	Else
		$oItems = Execute("$oObject.elements('" & $sName & "')")
	EndIf

	If Not IsObj($oItems) Then
		__ChromeConsoleWriteError("Warning", "_ChromeFormElementCheckBoxSelect", "$_ChromeSTATUS_NoMatch")
		Return SetError($_ChromeSTATUS_NoMatch, 3, 0)
	EndIf

	Local $oItem, $bFound = False
	Switch $sMode
		Case "byValue"
			If __ChromeIsObjType($oItems, "forminputelement") Then
				$oItem = $oItems
				If String($oItem.type) = "checkbox" And String($oItem.value) = $sString Then $bFound = True
			Else
				For $oItem In $oItems
					If String($oItem.type) = "checkbox" And String($oItem.value) = $sString Then
						$bFound = True
						ExitLoop
					EndIf
				Next
			EndIf
		Case "byIndex"
			If __ChromeIsObjType($oItems, "forminputelement") Then
				$oItem = $oItems
				If String($oItem.type) = "checkbox" And Number($sString) = 0 Then $bFound = True
			Else
				Local $iCount = 0
				For $oItem In $oItems
					If String($oItem.type) = "checkbox" And Number($sString) = $iCount Then
						$bFound = True
						ExitLoop
					Else
						If String($oItem.type) = "checkbox" Then $iCount += 1
					EndIf
				Next
			EndIf
		Case Else
			__ChromeConsoleWriteError("Error", "_ChromeFormElementCheckBoxSelect", "$_ChromeSTATUS_InvalidValue", "Invalid Mode")
			Return SetError($_ChromeSTATUS_InvalidValue, 5, 0)
	EndSwitch

	If Not $bFound Then
		__ChromeConsoleWriteError("Warning", "_ChromeFormElementCheckBoxSelect", "$_ChromeSTATUS_NoMatch")
		Return SetError($_ChromeSTATUS_NoMatch, 2, 0)
	EndIf

	Switch $iSelect
		Case -1
			Return SetError($_ChromeSTATUS_Success, 0, $oItem.checked)
		Case 0
			If $oItem.checked Then
				$oItem.checked = False
				If $iFireEvent Then
					$oItem.fireEvent("onChange")
					$oItem.fireEvent("OnClick")
				EndIf
			EndIf
			Return SetError($_ChromeSTATUS_Success, 0, 1)
		Case 1
			If Not $oItem.checked Then
				$oItem.checked = True
				If $iFireEvent Then
					$oItem.fireEvent("onChange")
					$oItem.fireEvent("OnClick")
				EndIf
			EndIf
			Return SetError($_ChromeSTATUS_Success, 0, 1)
		Case Else
			__ChromeConsoleWriteError("Error", "_ChromeFormElementCheckBoxSelect", "$_ChromeSTATUS_InvalidValue", "Invalid $iSelect value")
			Return SetError($_ChromeSTATUS_InvalidValue, 3, 0)
	EndSwitch
EndFunc   ;==>_ChromeFormElementCheckBoxSelect

; #FUNCTION# ====================================================================================================================
; Author ........: Dale Hohm
; ===============================================================================================================================
Func _ChromeFormElementRadioSelect(ByRef $oObject, $sString, $sName, $iSelect = 1, $sMode = "byValue", $iFireEvent = 1)
	If Not IsObj($oObject) Then
		__ChromeConsoleWriteError("Error", "_ChromeFormElementRadioSelect", "$_ChromeSTATUS_InvalidDataType")
		Return SetError($_ChromeSTATUS_InvalidDataType, 1, 0)
	EndIf
	;
	If Not __ChromeIsObjType($oObject, "form") Then
		__ChromeConsoleWriteError("Error", "_ChromeFormElementRadioSelect", "$_ChromeSTATUS_InvalidObjectType")
		Return SetError($_ChromeSTATUS_InvalidObjectType, 1, 0)
	EndIf
	;
	$sString = String($sString)
	$sName = String($sName)

	Local $oItems = Execute("$oObject.elements('" & $sName & "')")
	If Not IsObj($oItems) Then
		__ChromeConsoleWriteError("Warning", "_ChromeFormElementRadioSelect", "$_ChromeSTATUS_NoMatch")
		Return SetError($_ChromeSTATUS_NoMatch, 3, 0)
	EndIf

	Local $oItem, $bFound = False
	Switch $sMode
		Case "byValue"
			If __ChromeIsObjType($oItems, "forminputelement") Then
				$oItem = $oItems
				If String($oItem.type) = "radio" And String($oItem.value) = $sString Then $bFound = True
			Else
				For $oItem In $oItems
					If String($oItem.type) = "radio" And String($oItem.value) = $sString Then
						$bFound = True
						ExitLoop
					EndIf
				Next
			EndIf
		Case "byIndex"
			If __ChromeIsObjType($oItems, "forminputelement") Then
				$oItem = $oItems
				If String($oItem.type) = "radio" And Number($sString) = 0 Then $bFound = True
			Else
				Local $iCount = 0
				For $oItem In $oItems
					If String($oItem.type) = "radio" And Number($sString) = $iCount Then
						$bFound = True
						ExitLoop
					Else
						$iCount += 1
					EndIf
				Next
			EndIf
		Case Else
			__ChromeConsoleWriteError("Error", "_ChromeFormElementRadioSelect", "$_ChromeSTATUS_InvalidValue", "Invalid Mode")
			Return SetError($_ChromeSTATUS_InvalidValue, 5, 0)
	EndSwitch

	If Not $bFound Then
		__ChromeConsoleWriteError("Warning", "_ChromeFormElementRadioSelect", "$_ChromeSTATUS_NoMatch")
		Return SetError($_ChromeSTATUS_NoMatch, 2, 0)
	EndIf

	Switch $iSelect
		Case -1
			Return SetError($_ChromeSTATUS_Success, 0, $oItem.checked)
		Case 0
			If $oItem.checked Then
				$oItem.checked = False
				If $iFireEvent Then
					$oItem.fireEvent("onChange")
					$oItem.fireEvent("OnClick")
				EndIf
			EndIf
			Return SetError($_ChromeSTATUS_Success, 0, 1)
		Case 1
			If Not $oItem.checked Then
				$oItem.checked = True
				If $iFireEvent Then
					$oItem.fireEvent("onChange")
					$oItem.fireEvent("OnClick")
				EndIf
			EndIf
			Return SetError($_ChromeSTATUS_Success, 0, 1)
		Case Else
			__ChromeConsoleWriteError("Error", "_ChromeFormElementRadioSelect", "$_ChromeSTATUS_InvalidValue", "$iSelect value invalid")
			Return SetError($_ChromeSTATUS_InvalidValue, 4, 0)
	EndSwitch
EndFunc   ;==>_ChromeFormElementRadioSelect

; #FUNCTION# ====================================================================================================================
; Author ........: Dale Hohm
; ModifChromed ......: jpm
; ===============================================================================================================================
Func _ChromeFormImageClick(ByRef $oObject, $sLinkText, $sMode = "src", $iIndex = 0, $iWait = 1)
	If Not IsObj($oObject) Then
		__ChromeConsoleWriteError("Error", "_ChromeFormImageClick", "$_ChromeSTATUS_InvalidDataType")
		Return SetError($_ChromeSTATUS_InvalidDataType, 1, 0)
	EndIf
	;
	Local $sModeLinktext, $iFound = 0
	Local $oTemp = _ChromeDocGetObj($oObject)
	Local $oImgs = _ChromeTagNameGetCollection($oTemp, "input")
	$sMode = StringLower($sMode)
	$iIndex = Number($iIndex)
	For $oImg In $oImgs
		If String($oImg.type) = "image" Then
			Select
				Case $sMode = "alt"
					$sModeLinktext = $oImg.alt
				Case $sMode = "name"
					$sModeLinktext = $oImg.name
					If Not IsString($sModeLinktext) Then $sModeLinktext = $oImg.id ; html5 support
				Case $sMode = "id"
					$sModeLinktext = $oImg.id
				Case $sMode = "src"
					$sModeLinktext = $oImg.src
				Case Else
					__ChromeConsoleWriteError("Error", "_ChromeFormImageClick", "$_ChromeSTATUS_InvalidValue", "Invalid mode: " & $sMode)
					Return SetError($_ChromeSTATUS_InvalidValue, 3, 0)
			EndSelect
			If StringInStr($sModeLinktext, $sLinkText) Then
				If ($iFound = $iIndex) Then
					$oImg.click()
					If @error Then ; Trap COM error, report and return
						__ChromeConsoleWriteError("Error", "_ChromeFormImageClick", "$_ChromeSTATUS_COMError", @error)
						Return SetError($_ChromeSTATUS_ComError, @error, 0)
					EndIf
					If $iWait Then
						_ChromeLoadWait($oObject)
						Return SetError(@error, 0, -1)
					EndIf
					Return SetError($_ChromeSTATUS_Success, 0, -1)
				EndIf
				$iFound = $iFound + 1
			EndIf
		EndIf
	Next
	__ChromeConsoleWriteError("Warning", "_ChromeFormImageClick", "$_ChromeSTATUS_NoMatch")
	Return SetError($_ChromeSTATUS_NoMatch, 2, 0)
EndFunc   ;==>_ChromeFormImageClick

; #FUNCTION# ====================================================================================================================
; Author ........: Dale Hohm
; ===============================================================================================================================
Func _ChromeFormSubmit(ByRef $oObject, $iWait = 1)
	If Not IsObj($oObject) Then
		__ChromeConsoleWriteError("Error", "_ChromeFormSubmit", "$_ChromeSTATUS_InvalidDataType")
		Return SetError($_ChromeSTATUS_InvalidDataType, 1, 0)
	EndIf
	;
	If Not __ChromeIsObjType($oObject, "form") Then
		__ChromeConsoleWriteError("Error", "_ChromeFormSubmit", "$_ChromeSTATUS_InvalidObjectType")
		Return SetError($_ChromeSTATUS_InvalidObjectType, 1, 0)
	EndIf
	;

	Local $oWindow = $oObject.document.parentWindow
	$oObject.submit()
	If @error Then ; Trap COM error, report and return
		__ChromeConsoleWriteError("Error", "_ChromeFormSubmit", "$_ChromeSTATUS_COMError", @error)
		Return SetError($_ChromeSTATUS_ComError, @error, 0)
	EndIf
	If $iWait Then
		_ChromeLoadWait($oWindow)
		Return SetError(@error, 0, -1)
	EndIf
	Return SetError($_ChromeSTATUS_Success, 0, -1)
EndFunc   ;==>_ChromeFormSubmit

; #FUNCTION# ====================================================================================================================
; Author ........: Dale Hohm
; ===============================================================================================================================
Func _ChromeFormReset(ByRef $oObject)
	If Not IsObj($oObject) Then
		__ChromeConsoleWriteError("Error", "_ChromeFormReset", "$_ChromeSTATUS_InvalidDataType")
		Return SetError($_ChromeSTATUS_InvalidDataType, 1, 0)
	EndIf
	;
	If Not __ChromeIsObjType($oObject, "form") Then
		__ChromeConsoleWriteError("Error", "_ChromeFormReset", "$_ChromeSTATUS_InvalidObjectType")
		Return SetError($_ChromeSTATUS_InvalidObjectType, 1, 0)
	EndIf
	;
	$oObject.reset()
	If @error Then ; Trap COM error, report and return
		__ChromeConsoleWriteError("Error", "_ChromeFormReset", "$_ChromeSTATUS_COMError", @error)
		Return SetError($_ChromeSTATUS_ComError, @error, 0)
	EndIf
	Return SetError($_ChromeSTATUS_Success, 0, 1)
EndFunc   ;==>_ChromeFormReset
#EndRegion Form functions

#Region Table functions
; #FUNCTION# ====================================================================================================================
; Author ........: Dale Hohm
; ===============================================================================================================================
Func _ChromeTableGetCollection(ByRef $oObject, $iIndex = -1)
	If Not IsObj($oObject) Then
		__ChromeConsoleWriteError("Error", "_ChromeTableGetCollection", "$_ChromeSTATUS_InvalidDataType")
		Return SetError($_ChromeSTATUS_InvalidDataType, 1, 0)
	EndIf
	;
	$iIndex = Number($iIndex)
	Select
		Case $iIndex = -1
			Return SetError($_ChromeSTATUS_Success, $oObject.document.GetElementsByTagName("table").length, _
					$oObject.document.GetElementsByTagName("table"))
		Case $iIndex > -1 And $iIndex < $oObject.document.GetElementsByTagName("table").length
			Return SetError($_ChromeSTATUS_Success, $oObject.document.GetElementsByTagName("table").length, _
					$oObject.document.GetElementsByTagName("table").item($iIndex))
		Case $iIndex < -1
			__ChromeConsoleWriteError("Error", "_ChromeTableGetCollection", "$_ChromeSTATUS_InvalidValue", "$iIndex < -1")
			Return SetError($_ChromeSTATUS_InvalidValue, 2, 0)
		Case Else
			__ChromeConsoleWriteError("Warning", "_ChromeTableGetCollection", "$_ChromeSTATUS_NoMatch")
			Return SetError($_ChromeSTATUS_NoMatch, 1, 0)
	EndSelect
EndFunc   ;==>_ChromeTableGetCollection

; #FUNCTION# ====================================================================================================================
; Author ........: Dale Hohm
; ===============================================================================================================================
Func _ChromeTableWriteToArray(ByRef $oObject, $bTranspose = False)
	If Not IsObj($oObject) Then
		__ChromeConsoleWriteError("Error", "_ChromeTableWriteToArray", "$_ChromeSTATUS_InvalidDataType")
		Return SetError($_ChromeSTATUS_InvalidDataType, 1, 0)
	EndIf
	;
	If Not __ChromeIsObjType($oObject, "table") Then
		__ChromeConsoleWriteError("Error", "_ChromeTableWriteToArray", "$_ChromeSTATUS_InvalidObjectType")
		Return SetError($_ChromeSTATUS_InvalidObjectType, 1, 0)
	EndIf
	;
	Local $iCols = 0, $oTds, $iCol
	Local $oTrs = $oObject.rows
	For $oTr In $oTrs
		$oTds = $oTr.cells
		$iCol = 0
		For $oTd In $oTds
			$iCol = $iCol + $oTd.colSpan
		Next
		If $iCol > $iCols Then $iCols = $iCol
	Next
	Local $iRows = $oTrs.length
	Local $aTableCells[$iCols][$iRows]
	Local $iRow = 0
	For $oTr In $oTrs
		$oTds = $oTr.cells
		$iCol = 0
		For $oTd In $oTds
			$aTableCells[$iCol][$iRow] = String($oTd.innerText)
			If @error Then ; Trap COM error, report and return
				__ChromeConsoleWriteError("Error", "_ChromeTableWriteToArray", "$_ChromeSTATUS_COMError", @error)
				Return SetError($_ChromeSTATUS_ComError, @error, 0)
			EndIf
			$iCol = $iCol + $oTd.colSpan
		Next
		$iRow = $iRow + 1
	Next
	If $bTranspose Then
		Local $iD1 = UBound($aTableCells, $UBOUND_ROWS), $iD2 = UBound($aTableCells, $UBOUND_COLUMNS), $aTmp[$iD2][$iD1]
		For $i = 0 To $iD2 - 1
			For $j = 0 To $iD1 - 1
				$aTmp[$i][$j] = $aTableCells[$j][$i]
			Next
		Next
		$aTableCells = $aTmp
	EndIf
	Return SetError($_ChromeSTATUS_Success, 0, $aTableCells)
EndFunc   ;==>_ChromeTableWriteToArray
#EndRegion Table functions

#Region Read/Write functions
; #FUNCTION# ====================================================================================================================
; Author ........: Dale Hohm
; ===============================================================================================================================
Func _ChromeBodyReadHTML(ByRef $oObject)
	If Not IsObj($oObject) Then
		__ChromeConsoleWriteError("Error", "_ChromeBodyReadHTML", "$_ChromeSTATUS_InvalidDataType")
		Return SetError($_ChromeSTATUS_InvalidDataType, 1, 0)
	EndIf
	;
	Return SetError($_ChromeSTATUS_Success, 0, $oObject.document.body.innerHTML)
EndFunc   ;==>_ChromeBodyReadHTML

; #FUNCTION# ====================================================================================================================
; Author ........: Dale Hohm
; ===============================================================================================================================
Func _ChromeBodyReadText(ByRef $oObject)
	If Not IsObj($oObject) Then
		__ChromeConsoleWriteError("Error", "_ChromeBodyReadText", "$_ChromeSTATUS_InvalidDataType")
		Return SetError($_ChromeSTATUS_InvalidDataType, 1, 0)
	EndIf
	If Not __ChromeIsObjType($oObject, "browserdom") Then
		__ChromeConsoleWriteError("Error", "_ChromeBodyReadText", "$_ChromeSTATUS_InvalidObjectType", "Expected document element")
		Return SetError($_ChromeSTATUS_InvalidObjectType, 1, 0)
	EndIf
	;
	Return SetError($_ChromeSTATUS_Success, 0, $oObject.document.body.innerText)
EndFunc   ;==>_ChromeBodyReadText

; #FUNCTION# ====================================================================================================================
; Author ........: Dale Hohm
; ===============================================================================================================================
Func _ChromeBodyWriteHTML(ByRef $oObject, $sHTML)
	If Not IsObj($oObject) Then
		__ChromeConsoleWriteError("Error", "_ChromeBodyWriteHTML", "$_ChromeSTATUS_InvalidDataType")
		Return SetError($_ChromeSTATUS_InvalidDataType, 1, 0)
	EndIf
	If Not __ChromeIsObjType($oObject, "browserdom") Then
		__ChromeConsoleWriteError("Error", "_ChromeBodyWriteHTML", "$_ChromeSTATUS_InvalidObjectType", "Expected document element")
		Return SetError($_ChromeSTATUS_InvalidObjectType, 1, 0)
	EndIf
	;
	$oObject.document.body.innerHTML = $sHTML
	If @error Then ; Trap COM error, report and return
		__ChromeConsoleWriteError("Error", "_ChromeBodyWriteHTML", "$_ChromeSTATUS_COMError", @error)
		Return SetError($_ChromeSTATUS_ComError, @error, 0)
	EndIf
	Local $oTemp = $oObject.document
	_ChromeLoadWait($oTemp)
	Return SetError(@error, 0, -1)
EndFunc   ;==>_ChromeBodyWriteHTML

; #FUNCTION# ====================================================================================================================
; Author ........: Dale Hohm
; ===============================================================================================================================
Func _ChromeDocReadHTML(ByRef $oObject)
	If Not IsObj($oObject) Then
		__ChromeConsoleWriteError("Error", "_ChromeDocReadHTML", "$_ChromeSTATUS_InvalidDataType")
		Return SetError($_ChromeSTATUS_InvalidDataType, 1, 0)
	EndIf
	If Not __ChromeIsObjType($oObject, "browserdom") Then
		__ChromeConsoleWriteError("Error", "_ChromeDocReadHTML", "$_ChromeSTATUS_InvalidObjectType", "Expected document element")
		Return SetError($_ChromeSTATUS_InvalidObjectType, 1, 0)
	EndIf
	;
	Return SetError($_ChromeSTATUS_Success, 0, $oObject.document.documentElement.outerHTML)
EndFunc   ;==>_ChromeDocReadHTML

; #FUNCTION# ====================================================================================================================
; Author ........: Dale Hohm
; ===============================================================================================================================
Func _ChromeDocWriteHTML(ByRef $oObject, $sHTML)
	If Not IsObj($oObject) Then
		__ChromeConsoleWriteError("Error", "_ChromeDocWriteHTML", "$_ChromeSTATUS_InvalidDataType")
		Return SetError($_ChromeSTATUS_InvalidDataType, 1, 0)
	EndIf
	If Not __ChromeIsObjType($oObject, "browserdom") Then
		__ChromeConsoleWriteError("Error", "_ChromeDocWriteHTML", "$_ChromeSTATUS_InvalidObjectType", "Expected document element")
		Return SetError($_ChromeSTATUS_InvalidObjectType, 1, 0)
	EndIf
	;
	$oObject.document.Write($sHTML)
	$oObject.document.close()
	Local $oTemp = $oObject.document
	If @error Then ; Trap COM error, report and return
		__ChromeConsoleWriteError("Error", "_ChromeDocWriteHTML", "$_ChromeSTATUS_COMError", @error)
		Return SetError($_ChromeSTATUS_ComError, @error, 0)
	EndIf
	_ChromeLoadWait($oTemp)
	Return SetError(@error, 0, -1)
EndFunc   ;==>_ChromeDocWriteHTML

; #FUNCTION# ====================================================================================================================
; Author ........: Dale Hohm
; ===============================================================================================================================
Func _ChromeDocInsertText(ByRef $oObject, $sString, $sWhere = "beforeend")
	If Not IsObj($oObject) Then
		__ChromeConsoleWriteError("Error", "_ChromeDocInsertText", "$_ChromeSTATUS_InvalidDataType")
		Return SetError($_ChromeSTATUS_InvalidDataType, 1, 0)
	EndIf
	If Not __ChromeIsObjType($oObject, "browserdom") Or __ChromeIsObjType($oObject, "documentcontainer") Or __ChromeIsObjType($oObject, "document") Then
		__ChromeConsoleWriteError("Error", "_ChromeDocInsertText", "$_ChromeSTATUS_InvalidObjectType", "Expected document element")
		Return SetError($_ChromeSTATUS_InvalidObjectType, 1, 0)
	EndIf

	$sWhere = StringLower($sWhere)
	Select
		Case $sWhere = "beforebegin"
			$oObject.insertAdjacentText($sWhere, $sString)
		Case $sWhere = "afterbegin"
			$oObject.insertAdjacentText($sWhere, $sString)
		Case $sWhere = "beforeend"
			$oObject.insertAdjacentText($sWhere, $sString)
		Case $sWhere = "afterend"
			$oObject.insertAdjacentText($sWhere, $sString)
		Case Else
			; Unsupported Where
			__ChromeConsoleWriteError("Error", "_ChromeDocInsertText", "$_ChromeSTATUS_InvalidValue", "Invalid where value")
			Return SetError($_ChromeSTATUS_InvalidValue, 3, 0)
	EndSelect

	If @error Then ; Trap COM error, report and return
		__ChromeConsoleWriteError("Error", "_ChromeDocInsertText", "$_ChromeSTATUS_COMError", @error)
		Return SetError($_ChromeSTATUS_ComError, @error, 0)
	EndIf
	Return SetError($_ChromeSTATUS_Success, 0, 1)
EndFunc   ;==>_ChromeDocInsertText

; #FUNCTION# ====================================================================================================================
; Author ........: Dale Hohm
; ===============================================================================================================================
Func _ChromeDocInsertHTML(ByRef $oObject, $sString, $sWhere = "beforeend")
	If Not IsObj($oObject) Then
		__ChromeConsoleWriteError("Error", "_ChromeDocInsertHTML", "$_ChromeSTATUS_InvalidDataType")
		Return SetError($_ChromeSTATUS_InvalidDataType, 1, 0)
	EndIf
	If Not __ChromeIsObjType($oObject, "browserdom") Or __ChromeIsObjType($oObject, "documentcontainer") Or __ChromeIsObjType($oObject, "document") Then
		__ChromeConsoleWriteError("Error", "_ChromeDocInsertHTML", "$_ChromeSTATUS_InvalidObjectType", "Expected document element")
		Return SetError($_ChromeSTATUS_InvalidObjectType, 1, 0)
	EndIf

	$sWhere = StringLower($sWhere)
	Select
		Case $sWhere = "beforebegin"
			$oObject.insertAdjacentHTML($sWhere, $sString)
		Case $sWhere = "afterbegin"
			$oObject.insertAdjacentHTML($sWhere, $sString)
		Case $sWhere = "beforeend"
			$oObject.insertAdjacentHTML($sWhere, $sString)
		Case $sWhere = "afterend"
			$oObject.insertAdjacentHTML($sWhere, $sString)
		Case Else
			; Unsupported Where
			__ChromeConsoleWriteError("Error", "_ChromeDocInsertHTML", "$_ChromeSTATUS_InvalidValue", "Invalid where value")
			Return SetError($_ChromeSTATUS_InvalidValue, 3, 0)
	EndSelect

	If @error Then ; Trap COM error, report and return
		__ChromeConsoleWriteError("Error", "_ChromeDocInsertHTML", "$_ChromeSTATUS_COMError", @error)
		Return SetError($_ChromeSTATUS_ComError, @error, 0)
	EndIf
	Return SetError($_ChromeSTATUS_Success, 0, 1)
EndFunc   ;==>_ChromeDocInsertHTML

; #FUNCTION# ====================================================================================================================
; Author ........: Dale Hohm
; ModifChromed ......: jpm
; ===============================================================================================================================
Func _ChromeHeadInsertEventScript(ByRef $oObject, $sHTMLFor, $sEvent, $sScript)
	If Not IsObj($oObject) Then
		__ChromeConsoleWriteError("Error", "_ChromeHeadInsertEventScript", "$_ChromeSTATUS_InvalidDataType")
		Return SetError($_ChromeSTATUS_InvalidDataType, 1, 0)
	EndIf

	Local $oHead = $oObject.document.all.tags("HEAD").Item(0)
	Local $oScript = $oObject.document.createElement("script")
	If @error Then ; Trap COM error, report and return
		__ChromeConsoleWriteError("Error", "_ChromeHeadInsertEventScript(script)", "$_ChromeSTATUS_COMError", @error)
		Return SetError($_ChromeSTATUS_ComError, @error, 0)
	EndIf
	With $oScript
		.defer = True
		.language = "jscript"
		.type = "text/javascript"
		.htmlFor = $sHTMLFor
		.event = $sEvent
		.text = $sScript
	EndWith
	$oHead.appendChild($oScript)
	If @error Then ; Trap COM error, report and return
		__ChromeConsoleWriteError("Error", "_ChromeHeadInsertEventScript", "$_ChromeSTATUS_COMError", @error)
		Return SetError($_ChromeSTATUS_ComError, @error, 0)
	EndIf
	Return SetError($_ChromeSTATUS_Success, 0, 1)
EndFunc   ;==>_ChromeHeadInsertEventScript
#EndRegion Read/Write functions

#Region Utility functions
; #FUNCTION# ====================================================================================================================
; Author ........: Dale Hohm
; ===============================================================================================================================
Func _ChromeDocGetObj(ByRef $oObject)
	If Not IsObj($oObject) Then
		__ChromeConsoleWriteError("Error", "_ChromeDocGetObj", "$_ChromeSTATUS_InvalidDataType")
		Return SetError($_ChromeSTATUS_InvalidDataType, 1, 0)
	EndIf
	;
	If __ChromeIsObjType($oObject, "document") Then
		Return SetError($_ChromeSTATUS_Success, 0, $oObject)
	EndIf

	Return SetError($_ChromeSTATUS_Success, 0, $oObject.document)
EndFunc   ;==>_ChromeDocGetObj

; #FUNCTION# ====================================================================================================================
; Author ........: Dale Hohm
; ===============================================================================================================================
Func _ChromeTagNameGetCollection(ByRef $oObject, $sTagName, $iIndex = -1)
	If Not IsObj($oObject) Then
		__ChromeConsoleWriteError("Error", "_ChromeTagNameGetCollection", "$_ChromeSTATUS_InvalidDataType")
		Return SetError($_ChromeSTATUS_InvalidDataType, 1, 0)
	EndIf
	;
	If Not __ChromeIsObjType($oObject, "browserdom") Then
		__ChromeConsoleWriteError("Error", "_ChromeTagNameGetCollection", "$_ChromeSTATUS_InvalidObjectType")
		Return SetError($_ChromeSTATUS_InvalidObjectType, 1, 0)
	EndIf

	Local $oTemp
	If __ChromeIsObjType($oObject, "documentcontainer") Then
		$oTemp = _ChromeDocGetObj($oObject)
	Else
		$oTemp = $oObject
	EndIf

	$iIndex = Number($iIndex)
	Select
		Case $iIndex = -1
			Return SetError($_ChromeSTATUS_Success, $oTemp.GetElementsByTagName($sTagName).length, _
					$oTemp.GetElementsByTagName($sTagName))
		Case $iIndex > -1 And $iIndex < $oTemp.GetElementsByTagName($sTagName).length
			Return SetError($_ChromeSTATUS_Success, $oTemp.GetElementsByTagName($sTagName).length, _
					$oTemp.GetElementsByTagName($sTagName).item($iIndex))
		Case $iIndex < -1
			__ChromeConsoleWriteError("Error", "_ChromeTagNameGetCollection", "$_ChromeSTATUS_InvalidValue", "$iIndex < -1")
			Return SetError($_ChromeSTATUS_InvalidValue, 3, 0)
		Case Else
			__ChromeConsoleWriteError("Error", "_ChromeTagNameGetCollection", "$_ChromeSTATUS_NoMatch")
			Return SetError($_ChromeSTATUS_NoMatch, 0, 0) ; Could be caused by parameter 2, 3 or both
	EndSelect
EndFunc   ;==>_ChromeTagNameGetCollection

; #FUNCTION# ====================================================================================================================
; Author ........: Dale Hohm
; ===============================================================================================================================
Func _ChromeTagNameAllGetCollection(ByRef $oObject, $iIndex = -1)
	If Not IsObj($oObject) Then
		__ChromeConsoleWriteError("Error", "_ChromeTagNameAllGetCollection", "$_ChromeSTATUS_InvalidDataType")
		Return SetError($_ChromeSTATUS_InvalidDataType, 1, 0)
	EndIf
	;
	If Not __ChromeIsObjType($oObject, "browserdom") Then
		__ChromeConsoleWriteError("Error", "_ChromeTagNameAllGetCollection", "$_ChromeSTATUS_InvalidObjectType")
		Return SetError($_ChromeSTATUS_InvalidObjectType, 1, 0)
	EndIf

	Local $oTemp
	If __ChromeIsObjType($oObject, "documentcontainer") Then
		$oTemp = _ChromeDocGetObj($oObject)
	Else
		$oTemp = $oObject
	EndIf

	$iIndex = Number($iIndex)
	Select
		Case $iIndex = -1
			Return SetError($_ChromeSTATUS_Success, $oTemp.all.length, $oTemp.all)
		Case $iIndex > -1 And $iIndex < $oTemp.all.length
			Return SetError($_ChromeSTATUS_Success, $oTemp.all.length, $oTemp.all.item($iIndex))
		Case $iIndex < -1
			__ChromeConsoleWriteError("Error", "_ChromeTagNameAllGetCollection", "$_ChromeSTATUS_InvalidValue", "$iIndex < -1")
			Return SetError($_ChromeSTATUS_InvalidValue, 2, 0)
		Case Else
			__ChromeConsoleWriteError("Error", "_ChromeTagNameAllGetCollection", "$_ChromeSTATUS_NoMatch")
			Return SetError($_ChromeSTATUS_NoMatch, 1, 0)
	EndSelect
EndFunc   ;==>_ChromeTagNameAllGetCollection

; #FUNCTION# ====================================================================================================================
; Author ........: Dale Hohm
; ===============================================================================================================================
Func _ChromeGetObjByName(ByRef $oObject, $sName, $iIndex = 0)
	If Not IsObj($oObject) Then
		__ChromeConsoleWriteError("Error", "_ChromeGetObjByName", "$_ChromeSTATUS_InvalidDataType")
		Return SetError($_ChromeSTATUS_InvalidDataType, 1, 0)
	EndIf
	;
	$iIndex = Number($iIndex)
	If $iIndex = -1 Then
		Return SetError($_ChromeSTATUS_Success, $oObject.document.GetElementsByName($sName).length, _
				$oObject.document.GetElementsByName($sName))
	Else
		If IsObj($oObject.document.GetElementsByName($sName).item($iIndex)) Then
			Return SetError($_ChromeSTATUS_Success, $oObject.document.GetElementsByName($sName).length, _
					$oObject.document.GetElementsByName($sName).item($iIndex))
		Else
			__ChromeConsoleWriteError("Warning", "_ChromeGetObjByName", "$_ChromeSTATUS_NoMatch", "Name: " & $sName & ", Index: " & $iIndex)
			Return SetError($_ChromeSTATUS_NoMatch, 0, 0) ; Could be caused by parameter 2, 3 or both
		EndIf
	EndIf
EndFunc   ;==>_ChromeGetObjByName

; #FUNCTION# ====================================================================================================================
; Author ........: Dale Hohm
; ===============================================================================================================================
Func _ChromeGetObjById(ByRef $oObject, $sID)
	If Not IsObj($oObject) Then
		__ChromeConsoleWriteError("Error", "_ChromeGetObjById", "$_ChromeSTATUS_InvalidDataType")
		Return SetError($_ChromeSTATUS_InvalidDataType, 1, 0)
	EndIf
	;
	If Not __ChromeIsObjType($oObject, "browserdom") Then
		__ChromeConsoleWriteError("Error", "_ChromeGetObById", "$_ChromeSTATUS_InvalidObjectType")
		Return SetError($_ChromeSTATUS_InvalidObjectType, 1, 0)
	EndIf
	;
	If IsObj($oObject.document.getElementById($sID)) Then
		Return SetError($_ChromeSTATUS_Success, 0, $oObject.document.getElementById($sID))
	Else
		__ChromeConsoleWriteError("Warning", "_ChromeGetObjById", "$_ChromeSTATUS_NoMatch", $sID)
		Return SetError($_ChromeSTATUS_NoMatch, 2, 0)
	EndIf
EndFunc   ;==>_ChromeGetObjById

; #FUNCTION# ====================================================================================================================
; Author ........: Dale Hohm
; ===============================================================================================================================
Func _ChromeAction(ByRef $oObject, $sAction)
	If Not IsObj($oObject) Then
		__ChromeConsoleWriteError("Error", "_ChromeAction(" & $sAction & ")", "$_ChromeSTATUS_InvalidDataType")
		Return SetError($_ChromeSTATUS_InvalidDataType, 1, 0)
	EndIf
	;
	$sAction = StringLower($sAction)
	Select
		; DOM objects
		Case $sAction = "click"
			If __ChromeIsObjType($oObject, "documentContainer") Then
				__ChromeConsoleWriteError("Error", "_ChromeAction(click)", " $_ChromeSTATUS_InvalidObjectType")
				Return SetError($_ChromeSTATUS_InvalidObjectType, 1, 0)
			EndIf
			$oObject.Click()
		Case $sAction = "disable"
			If __ChromeIsObjType($oObject, "documentContainer") Then
				__ChromeConsoleWriteError("Error", "_ChromeAction(disable)", " $_ChromeSTATUS_InvalidObjectType")
				Return SetError($_ChromeSTATUS_InvalidObjectType, 1, 0)
			EndIf
			$oObject.disabled = True
		Case $sAction = "enable"
			If __ChromeIsObjType($oObject, "documentContainer") Then
				__ChromeConsoleWriteError("Error", "_ChromeAction(enable)", " $_ChromeSTATUS_InvalidObjectType")
				Return SetError($_ChromeSTATUS_InvalidObjectType, 1, 0)
			EndIf
			$oObject.disabled = False
		Case $sAction = "focus"
			If __ChromeIsObjType($oObject, "documentContainer") Then
				__ChromeConsoleWriteError("Error", "_ChromeAction(focus)", "$_ChromeSTATUS_InvalidObjectType")
				Return SetError($_ChromeSTATUS_InvalidObjectType, 1, 0)
			EndIf
			$oObject.Focus()
		Case $sAction = "scrollintovChromew"
			If __ChromeIsObjType($oObject, "documentContainer") Then
				__ChromeConsoleWriteError("Error", "_ChromeAction(scrollintovChromew)", "$_ChromeSTATUS_InvalidObjectType")
				Return SetError($_ChromeSTATUS_InvalidObjectType, 1, 0)
			EndIf
			$oObject.scrollIntoVChromew()
			; Browser Object
		Case $sAction = "copy"
			$oObject.document.execCommand("Copy")
		Case $sAction = "cut"
			$oObject.document.execCommand("Cut")
		Case $sAction = "paste"
			$oObject.document.execCommand("Paste")
		Case $sAction = "delete"
			$oObject.document.execCommand("Delete")
		Case $sAction = "saveas"
			$oObject.document.execCommand("SaveAs")
		Case $sAction = "refresh"
			$oObject.document.execCommand("Refresh")
			If @error Then ; Trap COM error, report and return
				__ChromeConsoleWriteError("Error", "_ChromeAction(refresh)", "$_ChromeSTATUS_COMError", @error)
				Return SetError($_ChromeSTATUS_ComError, @error, 0)
			EndIf
			_ChromeLoadWait($oObject)
		Case $sAction = "selectall"
			$oObject.document.execCommand("SelectAll")
		Case $sAction = "unselect"
			$oObject.document.execCommand("Unselect")
		Case $sAction = "print"
			$oObject.document.parentwindow.Print()
		Case $sAction = "printdefault"
			If Not __ChromeIsObjType($oObject, "browser") Then
				__ChromeConsoleWriteError("Error", "_ChromeAction(printdefault)", "$_ChromeSTATUS_InvalidObjectType")
				Return SetError($_ChromeSTATUS_InvalidObjectType, 1, 0)
			EndIf
			$oObject.execWB(6, 2)
		Case $sAction = "back"
			If Not __ChromeIsObjType($oObject, "documentContainer") Then
				__ChromeConsoleWriteError("Error", "_ChromeAction(back)", "$_ChromeSTATUS_InvalidObjectType")
				Return SetError($_ChromeSTATUS_InvalidObjectType, 1, 0)
			EndIf
			$oObject.GoBack()
		Case $sAction = "blur"
			$oObject.Blur()
		Case $sAction = "forward"
			If Not __ChromeIsObjType($oObject, "documentContainer") Then
				__ChromeConsoleWriteError("Error", "_ChromeAction(forward)", "$_ChromeSTATUS_InvalidObjectType")
				Return SetError($_ChromeSTATUS_InvalidObjectType, 1, 0)
			EndIf
			$oObject.GoForward()
		Case $sAction = "home"
			If Not __ChromeIsObjType($oObject, "documentContainer") Then
				__ChromeConsoleWriteError("Error", "_ChromeAction(home)", "$_ChromeSTATUS_InvalidObjectType")
				Return SetError($_ChromeSTATUS_InvalidObjectType, 1, 0)
			EndIf
			$oObject.GoHome()
		Case $sAction = "invisible"
			If Not __ChromeIsObjType($oObject, "browser") Then
				__ChromeConsoleWriteError("Error", "_ChromeAction(invisible)", "$_ChromeSTATUS_InvalidObjectType")
				Return SetError($_ChromeSTATUS_InvalidObjectType, 1, 0)
			EndIf
			$oObject.visible = 0
		Case $sAction = "visible"
			If Not __ChromeIsObjType($oObject, "browser") Then
				__ChromeConsoleWriteError("Error", "_ChromeAction(visible)", "$_ChromeSTATUS_InvalidObjectType")
				Return SetError($_ChromeSTATUS_InvalidObjectType, 1, 0)
			EndIf
			$oObject.visible = 1
		Case $sAction = "search"
			If Not __ChromeIsObjType($oObject, "browser") Then
				__ChromeConsoleWriteError("Error", "_ChromeAction(search)", "$_ChromeSTATUS_InvalidObjectType")
				Return SetError($_ChromeSTATUS_InvalidObjectType, 1, 0)
			EndIf
			$oObject.GoSearch()
		Case $sAction = "stop"
			If Not __ChromeIsObjType($oObject, "documentContainer") Then
				__ChromeConsoleWriteError("Error", "_ChromeAction(stop)", "$_ChromeSTATUS_InvalidObjectType")
				Return SetError($_ChromeSTATUS_InvalidObjectType, 1, 0)
			EndIf
			$oObject.Stop()
		Case $sAction = "quit"
			If Not __ChromeIsObjType($oObject, "browser") Then
				__ChromeConsoleWriteError("Error", "_ChromeAction(quit)", "$_ChromeSTATUS_InvalidObjectType")
				Return SetError($_ChromeSTATUS_InvalidObjectType, 1, 0)
			EndIf
			$oObject.Quit()
			If @error Then ; Trap COM error, report and return
				__ChromeConsoleWriteError("Error", "_ChromeAction(" & $sAction & ")", "$_ChromeSTATUS_COMError", @error)
				Return SetError($_ChromeSTATUS_ComError, @error, 0)
			EndIf
			$oObject = 0
			Return SetError($_ChromeSTATUS_Success, 0, 1)
		Case Else
			; Unsupported Action
			__ChromeConsoleWriteError("Error", "_ChromeAction(" & $sAction & ")", "$_ChromeSTATUS_InvalidValue", "Invalid Action")
			Return SetError($_ChromeSTATUS_InvalidValue, 2, 0)
	EndSelect

	If @error Then ; Trap COM error, report and return
		__ChromeConsoleWriteError("Error", "_ChromeAction(" & $sAction & ")", "$_ChromeSTATUS_COMError", @error)
		Return SetError($_ChromeSTATUS_ComError, @error, 0)
	EndIf
	Return SetError($_ChromeSTATUS_Success, 0, 1)
EndFunc   ;==>_ChromeAction

; #FUNCTION# ====================================================================================================================
; Author ........: Dale Hohm
; ===============================================================================================================================
Func _ChromePropertyGet(ByRef $oObject, $sProperty)
	If Not IsObj($oObject) Then
		__ChromeConsoleWriteError("Error", "_ChromePropertyGet", "$_ChromeSTATUS_InvalidDataType")
		Return SetError($_ChromeSTATUS_InvalidDataType, 1, 0)
	EndIf
	If Not __ChromeIsObjType($oObject, "browserdom") Then
		__ChromeConsoleWriteError("Error", "_ChromePropertyGet", "$_ChromeSTATUS_InvalidObjectType")
		Return SetError($_ChromeSTATUS_InvalidObjectType, 1, 0)
	EndIf
	;
	Local $oTemp, $iTemp
	$sProperty = StringLower($sProperty)
	Select
		Case $sProperty = "browserx"
			If __ChromeIsObjType($oObject, "browsercontainer") Or __ChromeIsObjType($oObject, "document") Then
				__ChromeConsoleWriteError("Error", "_ChromePropertyGet", "$_ChromeSTATUS_InvalidObjectType")
				Return SetError($_ChromeSTATUS_InvalidObjectType, 1, 0)
			EndIf
			$oTemp = $oObject
			$iTemp = 0
			While IsObj($oTemp)
				$iTemp += $oTemp.offsetLeft
				$oTemp = $oTemp.offsetParent
			WEnd
			Return SetError($_ChromeSTATUS_Success, 0, $iTemp)
		Case $sProperty = "browsery"
			If __ChromeIsObjType($oObject, "browsercontainer") Or __ChromeIsObjType($oObject, "document") Then
				__ChromeConsoleWriteError("Error", "_ChromePropertyGet", "$_ChromeSTATUS_InvalidObjectType")
				Return SetError($_ChromeSTATUS_InvalidObjectType, 1, 0)
			EndIf
			$oTemp = $oObject
			$iTemp = 0
			While IsObj($oTemp)
				$iTemp += $oTemp.offsetTop
				$oTemp = $oTemp.offsetParent
			WEnd
			Return SetError($_ChromeSTATUS_Success, 0, $iTemp)
		Case $sProperty = "screenx"
			If __ChromeIsObjType($oObject, "window") Or __ChromeIsObjType($oObject, "document") Then
				__ChromeConsoleWriteError("Error", "_ChromePropertyGet", "$_ChromeSTATUS_InvalidObjectType")
				Return SetError($_ChromeSTATUS_InvalidObjectType, 1, 0)
			EndIf
			If __ChromeIsObjType($oObject, "browser") Then
				Return SetError($_ChromeSTATUS_Success, 0, $oObject.left())
			Else
				$oTemp = $oObject
				$iTemp = 0
				While IsObj($oTemp)
					$iTemp += $oTemp.offsetLeft
					$oTemp = $oTemp.offsetParent
				WEnd
			EndIf
			Return SetError($_ChromeSTATUS_Success, 0, _
					$iTemp + $oObject.document.parentWindow.screenLeft)
		Case $sProperty = "screeny"
			If __ChromeIsObjType($oObject, "window") Or __ChromeIsObjType($oObject, "document") Then
				__ChromeConsoleWriteError("Error", "_ChromePropertyGet", "$_ChromeSTATUS_InvalidObjectType")
				Return SetError($_ChromeSTATUS_InvalidObjectType, 1, 0)
			EndIf
			If __ChromeIsObjType($oObject, "browser") Then
				Return SetError($_ChromeSTATUS_Success, 0, $oObject.top())
			Else
				$oTemp = $oObject
				$iTemp = 0
				While IsObj($oTemp)
					$iTemp += $oTemp.offsetTop
					$oTemp = $oTemp.offsetParent
				WEnd
			EndIf
			Return SetError($_ChromeSTATUS_Success, 0, _
					$iTemp + $oObject.document.parentWindow.screenTop)
		Case $sProperty = "height"
			If __ChromeIsObjType($oObject, "window") Or __ChromeIsObjType($oObject, "document") Then
				__ChromeConsoleWriteError("Error", "_ChromePropertyGet", "$_ChromeSTATUS_InvalidObjectType")
				Return SetError($_ChromeSTATUS_InvalidObjectType, 1, 0)
			EndIf
			If __ChromeIsObjType($oObject, "browser") Then
				Return SetError($_ChromeSTATUS_Success, 0, $oObject.Height())
			Else
				Return SetError($_ChromeSTATUS_Success, 0, $oObject.offsetHeight)
			EndIf
		Case $sProperty = "width"
			If __ChromeIsObjType($oObject, "window") Or __ChromeIsObjType($oObject, "document") Then
				__ChromeConsoleWriteError("Error", "_ChromePropertyGet", "$_ChromeSTATUS_InvalidObjectType")
				Return SetError($_ChromeSTATUS_InvalidObjectType, 1, 0)
			EndIf
			If __ChromeIsObjType($oObject, "browser") Then
				Return SetError($_ChromeSTATUS_Success, 0, $oObject.Width())
			Else
				Return SetError($_ChromeSTATUS_Success, 0, $oObject.offsetWidth)
			EndIf
		Case $sProperty = "isdisabled"
			Return SetError($_ChromeSTATUS_Success, 0, $oObject.isDisabled())
		Case $sProperty = "addressbar"
			If Not __ChromeIsObjType($oObject, "browser") Then
				__ChromeConsoleWriteError("Error", "_ChromePropertyGet", "$_ChromeSTATUS_InvalidObjectType")
				Return SetError($_ChromeSTATUS_InvalidObjectType, 1, 0)
			EndIf
			Return SetError($_ChromeSTATUS_Success, 0, $oObject.AddressBar())
		Case $sProperty = "busy"
			If Not __ChromeIsObjType($oObject, "browser") Then
				__ChromeConsoleWriteError("Error", "_ChromePropertyGet", "$_ChromeSTATUS_InvalidObjectType")
				Return SetError($_ChromeSTATUS_InvalidObjectType, 1, 0)
			EndIf
			Return SetError($_ChromeSTATUS_Success, 0, $oObject.Busy())
		Case $sProperty = "fullscreen"
			If Not __ChromeIsObjType($oObject, "browser") Then
				__ChromeConsoleWriteError("Error", "_ChromePropertyGet", "$_ChromeSTATUS_InvalidObjectType")
				Return SetError($_ChromeSTATUS_InvalidObjectType, 1, 0)
			EndIf
			Return SetError($_ChromeSTATUS_Success, 0, $oObject.fullScreen())
		Case $sProperty = "hwnd"
			If Not __ChromeIsObjType($oObject, "browser") Then
				__ChromeConsoleWriteError("Error", "_ChromePropertyGet", "$_ChromeSTATUS_InvalidObjectType")
				Return SetError($_ChromeSTATUS_InvalidObjectType, 1, 0)
			EndIf
			Return SetError($_ChromeSTATUS_Success, 0, HWnd($oObject.HWnd()))
		Case $sProperty = "left"
			If Not __ChromeIsObjType($oObject, "browser") Then
				__ChromeConsoleWriteError("Error", "_ChromePropertyGet", "$_ChromeSTATUS_InvalidObjectType")
				Return SetError($_ChromeSTATUS_InvalidObjectType, 1, 0)
			EndIf
			Return SetError($_ChromeSTATUS_Success, 0, $oObject.Left())
		Case $sProperty = "locationname"
			If Not __ChromeIsObjType($oObject, "browser") Then
				__ChromeConsoleWriteError("Error", "_ChromePropertyGet", "$_ChromeSTATUS_InvalidObjectType")
				Return SetError($_ChromeSTATUS_InvalidObjectType, 1, 0)
			EndIf
			Return SetError($_ChromeSTATUS_Success, 0, $oObject.LocationName())
		Case $sProperty = "locationurl"
			If __ChromeIsObjType($oObject, "browser") Then
				Return SetError($_ChromeSTATUS_Success, 0, $oObject.locationURL())
			EndIf
			If __ChromeIsObjType($oObject, "window") Then
				Return SetError($_ChromeSTATUS_Success, 0, $oObject.location.href())
			EndIf
			If __ChromeIsObjType($oObject, "document") Then
				Return SetError($_ChromeSTATUS_Success, 0, $oObject.parentwindow.location.href())
			EndIf
			Return SetError($_ChromeSTATUS_Success, 0, $oObject.document.parentwindow.location.href())
		Case $sProperty = "menubar"
			If Not __ChromeIsObjType($oObject, "browser") Then
				__ChromeConsoleWriteError("Error", "_ChromePropertyGet", "$_ChromeSTATUS_InvalidObjectType")
				Return SetError($_ChromeSTATUS_InvalidObjectType, 1, 0)
			EndIf
			Return SetError($_ChromeSTATUS_Success, 0, $oObject.MenuBar())
		Case $sProperty = "offline"
			If Not __ChromeIsObjType($oObject, "browser") Then
				__ChromeConsoleWriteError("Error", "_ChromePropertyGet", "$_ChromeSTATUS_InvalidObjectType")
				Return SetError($_ChromeSTATUS_InvalidObjectType, 1, 0)
			EndIf
			Return SetError($_ChromeSTATUS_Success, 0, $oObject.OffLine())
		Case $sProperty = "readystate"
			If Not __ChromeIsObjType($oObject, "browser") Then
				__ChromeConsoleWriteError("Error", "_ChromePropertyGet", "$_ChromeSTATUS_InvalidObjectType")
				Return SetError($_ChromeSTATUS_InvalidObjectType, 1, 0)
			EndIf
			Return SetError($_ChromeSTATUS_Success, 0, $oObject.ReadyState())
		Case $sProperty = "resizable"
			If Not __ChromeIsObjType($oObject, "browser") Then
				__ChromeConsoleWriteError("Error", "_ChromePropertyGet", "$_ChromeSTATUS_InvalidObjectType")
				Return SetError($_ChromeSTATUS_InvalidObjectType, 1, 0)
			EndIf
			Return SetError($_ChromeSTATUS_Success, 0, $oObject.Resizable())
		Case $sProperty = "silent"
			If Not __ChromeIsObjType($oObject, "browser") Then
				__ChromeConsoleWriteError("Error", "_ChromePropertyGet", "$_ChromeSTATUS_InvalidObjectType")
				Return SetError($_ChromeSTATUS_InvalidObjectType, 1, 0)
			EndIf
			Return SetError($_ChromeSTATUS_Success, 0, $oObject.Silent())
		Case $sProperty = "statusbar"
			If Not __ChromeIsObjType($oObject, "browser") Then
				__ChromeConsoleWriteError("Error", "_ChromePropertyGet", "$_ChromeSTATUS_InvalidObjectType")
				Return SetError($_ChromeSTATUS_InvalidObjectType, 1, 0)
			EndIf
			Return SetError($_ChromeSTATUS_Success, 0, $oObject.StatusBar())
		Case $sProperty = "statustext"
			If Not __ChromeIsObjType($oObject, "browser") Then
				__ChromeConsoleWriteError("Error", "_ChromePropertyGet", "$_ChromeSTATUS_InvalidObjectType")
				Return SetError($_ChromeSTATUS_InvalidObjectType, 1, 0)
			EndIf
			Return SetError($_ChromeSTATUS_Success, 0, $oObject.StatusText())
		Case $sProperty = "top"
			If Not __ChromeIsObjType($oObject, "browser") Then
				__ChromeConsoleWriteError("Error", "_ChromePropertyGet", "$_ChromeSTATUS_InvalidObjectType")
				Return SetError($_ChromeSTATUS_InvalidObjectType, 1, 0)
			EndIf
			Return SetError($_ChromeSTATUS_Success, 0, $oObject.Top())
		Case $sProperty = "visible"
			If Not __ChromeIsObjType($oObject, "browser") Then
				__ChromeConsoleWriteError("Error", "_ChromePropertyGet", "$_ChromeSTATUS_InvalidObjectType")
				Return SetError($_ChromeSTATUS_InvalidObjectType, 1, 0)
			EndIf
			Return SetError($_ChromeSTATUS_Success, 0, $oObject.Visible())
		Case $sProperty = "appcodename"
			Return SetError($_ChromeSTATUS_Success, 0, $oObject.document.parentWindow.top.navigator.appCodeName())
		Case $sProperty = "appminorversion"
			Return SetError($_ChromeSTATUS_Success, 0, $oObject.document.parentWindow.top.navigator.appMinorVersion())
		Case $sProperty = "appname"
			Return SetError($_ChromeSTATUS_Success, 0, $oObject.document.parentWindow.top.navigator.appName())
		Case $sProperty = "appversion"
			Return SetError($_ChromeSTATUS_Success, 0, $oObject.document.parentWindow.top.navigator.appVersion())
		Case $sProperty = "browserlanguage"
			Return SetError($_ChromeSTATUS_Success, 0, $oObject.document.parentWindow.top.navigator.browserLanguage())
		Case $sProperty = "cookChromeenabled"
			Return SetError($_ChromeSTATUS_Success, 0, $oObject.document.parentWindow.top.navigator.cookChromeEnabled())
		Case $sProperty = "cpuclass"
			Return SetError($_ChromeSTATUS_Success, 0, $oObject.document.parentWindow.top.navigator.cpuClass())
		Case $sProperty = "javaenabled"
			Return SetError($_ChromeSTATUS_Success, 0, $oObject.document.parentWindow.top.navigator.javaEnabled())
		Case $sProperty = "online"
			Return SetError($_ChromeSTATUS_Success, 0, $oObject.document.parentWindow.top.navigator.onLine())
		Case $sProperty = "platform"
			Return SetError($_ChromeSTATUS_Success, 0, $oObject.document.parentWindow.top.navigator.platform())
		Case $sProperty = "systemlanguage"
			Return SetError($_ChromeSTATUS_Success, 0, $oObject.document.parentWindow.top.navigator.systemLanguage())
		Case $sProperty = "useragent"
			Return SetError($_ChromeSTATUS_Success, 0, $oObject.document.parentWindow.top.navigator.userAgent())
		Case $sProperty = "userlanguage"
			Return SetError($_ChromeSTATUS_Success, 0, $oObject.document.parentWindow.top.navigator.userLanguage())
		Case $sProperty = "referrer"
			Return SetError($_ChromeSTATUS_Success, 0, $oObject.document.referrer)
		Case $sProperty = "theatermode"
			If Not __ChromeIsObjType($oObject, "browser") Then
				__ChromeConsoleWriteError("Error", "_ChromePropertyGet", "$_ChromeSTATUS_InvalidObjectType")
				Return SetError($_ChromeSTATUS_InvalidObjectType, 1, 0)
			EndIf
			Return SetError($_ChromeSTATUS_Success, 0, $oObject.TheaterMode)
		Case $sProperty = "toolbar"
			If Not __ChromeIsObjType($oObject, "browser") Then
				__ChromeConsoleWriteError("Error", "_ChromePropertyGet", "$_ChromeSTATUS_InvalidObjectType")
				Return SetError($_ChromeSTATUS_InvalidObjectType, 1, 0)
			EndIf
			Return SetError($_ChromeSTATUS_Success, 0, $oObject.ToolBar)
		Case $sProperty = "contenteditable"
			If __ChromeIsObjType($oObject, "browser") Or __ChromeIsObjType($oObject, "document") Then
				$oTemp = $oObject.document.body
			Else
				$oTemp = $oObject
			EndIf
			Return SetError($_ChromeSTATUS_Success, 0, $oTemp.isContentEditable)
		Case $sProperty = "innertext"
			If __ChromeIsObjType($oObject, "documentcontainer") Or __ChromeIsObjType($oObject, "document") Then
				$oTemp = $oObject.document.body
			Else
				$oTemp = $oObject
			EndIf
			Return SetError($_ChromeSTATUS_Success, 0, $oTemp.innerText)
		Case $sProperty = "outertext"
			If __ChromeIsObjType($oObject, "documentcontainer") Or __ChromeIsObjType($oObject, "document") Then
				$oTemp = $oObject.document.body
			Else
				$oTemp = $oObject
			EndIf
			Return SetError($_ChromeSTATUS_Success, 0, $oTemp.outerText)
		Case $sProperty = "innerhtml"
			If __ChromeIsObjType($oObject, "documentcontainer") Or __ChromeIsObjType($oObject, "document") Then
				$oTemp = $oObject.document.body
			Else
				$oTemp = $oObject
			EndIf
			Return SetError($_ChromeSTATUS_Success, 0, $oTemp.innerHTML)
		Case $sProperty = "outerhtml"
			If __ChromeIsObjType($oObject, "documentcontainer") Or __ChromeIsObjType($oObject, "document") Then
				$oTemp = $oObject.document.body
			Else
				$oTemp = $oObject
			EndIf
			Return SetError($_ChromeSTATUS_Success, 0, $oTemp.outerHTML)
		Case $sProperty = "title"
			Return SetError($_ChromeSTATUS_Success, 0, $oObject.document.title)
		Case $sProperty = "uniqueid"
			If __ChromeIsObjType($oObject, "window") Then
				__ChromeConsoleWriteError("Error", "_ChromePropertyGet", "$_ChromeSTATUS_InvalidObjectType")
				Return SetError($_ChromeSTATUS_InvalidObjectType, 1, 0)
			Else
				Return SetError($_ChromeSTATUS_Success, 0, $oObject.uniqueID)
			EndIf
		Case Else
			; Unsupported Property
			__ChromeConsoleWriteError("Error", "_ChromePropertyGet", "$_ChromeSTATUS_InvalidValue", "Invalid Property")
			Return SetError($_ChromeSTATUS_InvalidValue, 2, 0)
	EndSelect
EndFunc   ;==>_ChromePropertyGet

; #FUNCTION# ====================================================================================================================
; Author ........: Dale Hohm
; ===============================================================================================================================
Func _ChromePropertySet(ByRef $oObject, $sProperty, $vValue)
	If Not IsObj($oObject) Then
		__ChromeConsoleWriteError("Error", "_ChromePropertySet", "$_ChromeSTATUS_InvalidDataType")
		Return SetError($_ChromeSTATUS_InvalidDataType, 1, 0)
	EndIf
	;
	Local $oTemp
	#forceref $oTemp
	$sProperty = StringLower($sProperty)
	Select
		Case $sProperty = "addressbar"
			If Not __ChromeIsObjType($oObject, "browser") Then
				__ChromeConsoleWriteError("Error", "_ChromePropertySet", "$_ChromeSTATUS_InvalidObjectType")
				Return SetError($_ChromeSTATUS_InvalidObjectType, 1, 0)
			EndIf
			$oObject.AddressBar = $vValue
		Case $sProperty = "height"
			If Not __ChromeIsObjType($oObject, "browser") Then
				__ChromeConsoleWriteError("Error", "_ChromePropertySet", "$_ChromeSTATUS_InvalidObjectType")
				Return SetError($_ChromeSTATUS_InvalidObjectType, 1, 0)
			EndIf
			$oObject.Height = $vValue
		Case $sProperty = "left"
			If Not __ChromeIsObjType($oObject, "browser") Then
				__ChromeConsoleWriteError("Error", "_ChromePropertySet", "$_ChromeSTATUS_InvalidObjectType")
				Return SetError($_ChromeSTATUS_InvalidObjectType, 1, 0)
			EndIf
			$oObject.Left = $vValue
		Case $sProperty = "menubar"
			If Not __ChromeIsObjType($oObject, "browser") Then
				__ChromeConsoleWriteError("Error", "_ChromePropertySet", "$_ChromeSTATUS_InvalidObjectType")
				Return SetError($_ChromeSTATUS_InvalidObjectType, 1, 0)
			EndIf
			$oObject.MenuBar = $vValue
		Case $sProperty = "offline"
			If Not __ChromeIsObjType($oObject, "browser") Then
				__ChromeConsoleWriteError("Error", "_ChromePropertySet", "$_ChromeSTATUS_InvalidObjectType")
				Return SetError($_ChromeSTATUS_InvalidObjectType, 1, 0)
			EndIf
			$oObject.OffLine = $vValue
		Case $sProperty = "resizable"
			If Not __ChromeIsObjType($oObject, "browser") Then
				__ChromeConsoleWriteError("Error", "_ChromePropertySet", "$_ChromeSTATUS_InvalidObjectType")
				Return SetError($_ChromeSTATUS_InvalidObjectType, 1, 0)
			EndIf
			$oObject.Resizable = $vValue
		Case $sProperty = "statusbar"
			If Not __ChromeIsObjType($oObject, "browser") Then
				__ChromeConsoleWriteError("Error", "_ChromePropertySet", "$_ChromeSTATUS_InvalidObjectType")
				Return SetError($_ChromeSTATUS_InvalidObjectType, 1, 0)
			EndIf
			$oObject.StatusBar = $vValue
		Case $sProperty = "statustext"
			If Not __ChromeIsObjType($oObject, "browser") Then
				__ChromeConsoleWriteError("Error", "_ChromePropertySet", "$_ChromeSTATUS_InvalidObjectType")
				Return SetError($_ChromeSTATUS_InvalidObjectType, 1, 0)
			EndIf
			$oObject.StatusText = $vValue
		Case $sProperty = "top"
			If Not __ChromeIsObjType($oObject, "browser") Then
				__ChromeConsoleWriteError("Error", "_ChromePropertySet", "$_ChromeSTATUS_InvalidObjectType")
				Return SetError($_ChromeSTATUS_InvalidObjectType, 1, 0)
			EndIf
			$oObject.Top = $vValue
		Case $sProperty = "width"
			If Not __ChromeIsObjType($oObject, "browser") Then
				__ChromeConsoleWriteError("Error", "_ChromePropertySet", "$_ChromeSTATUS_InvalidObjectType")
				Return SetError($_ChromeSTATUS_InvalidObjectType, 1, 0)
			EndIf
			$oObject.Width = $vValue
		Case $sProperty = "theatermode"
			If Not __ChromeIsObjType($oObject, "browser") Then
				__ChromeConsoleWriteError("Error", "_ChromePropertySet", "$_ChromeSTATUS_InvalidObjectType")
				Return SetError($_ChromeSTATUS_InvalidObjectType, 1, 0)
			EndIf
			If $vValue Then
				$oObject.TheaterMode = True
			Else
				$oObject.TheaterMode = False
			EndIf
		Case $sProperty = "toolbar"
			If Not __ChromeIsObjType($oObject, "browser") Then
				__ChromeConsoleWriteError("Error", "_ChromePropertySet", "$_ChromeSTATUS_InvalidObjectType")
				Return SetError($_ChromeSTATUS_InvalidObjectType, 1, 0)
			EndIf
			If $vValue Then
				$oObject.ToolBar = True
			Else
				$oObject.ToolBar = False
			EndIf
		Case $sProperty = "contenteditable"
			If __ChromeIsObjType($oObject, "browser") Or __ChromeIsObjType($oObject, "document") Then
				$oTemp = $oObject.document.body
			Else
				$oTemp = $oObject
			EndIf
			If $vValue Then
				$oTemp.contentEditable = "true"
			Else
				$oTemp.contentEditable = "false"
			EndIf
		Case $sProperty = "innertext"
			If __ChromeIsObjType($oObject, "documentcontainer") Or __ChromeIsObjType($oObject, "document") Then
				$oTemp = $oObject.document.body
			Else
				$oTemp = $oObject
			EndIf
			$oTemp.innerText = $vValue
		Case $sProperty = "outertext"
			If __ChromeIsObjType($oObject, "documentcontainer") Or __ChromeIsObjType($oObject, "document") Then
				$oTemp = $oObject.document.body
			Else
				$oTemp = $oObject
			EndIf
			$oTemp.outerText = $vValue
		Case $sProperty = "innerhtml"
			If __ChromeIsObjType($oObject, "documentcontainer") Or __ChromeIsObjType($oObject, "document") Then
				$oTemp = $oObject.document.body
			Else
				$oTemp = $oObject
			EndIf
			$oTemp.innerHTML = $vValue
		Case $sProperty = "outerhtml"
			If __ChromeIsObjType($oObject, "documentcontainer") Or __ChromeIsObjType($oObject, "document") Then
				$oTemp = $oObject.document.body
			Else
				$oTemp = $oObject
			EndIf
			$oTemp.outerHTML = $vValue
		Case $sProperty = "title"
			$oObject.document.title = $vValue
		Case $sProperty = "silent"
			If Not __ChromeIsObjType($oObject, "browser") Then
				__ChromeConsoleWriteError("Error", "_ChromePropertySet", "$_ChromeSTATUS_InvalidObjectType")
				Return SetError($_ChromeSTATUS_InvalidObjectType, 1, 0)
			EndIf
			If $vValue Then
				$oObject.silent = True
			Else
				$oObject.silent = False
			EndIf
		Case Else
			; Unsupported Property
			__ChromeConsoleWriteError("Error", "_ChromePropertySet", "$_ChromeSTATUS_InvalidValue", "Invalid Property")
			Return SetError($_ChromeSTATUS_InvalidValue, 2, 0)
	EndSelect

	If @error Then ; Trap COM error, report and return
		__ChromeConsoleWriteError("Error", "_ChromePropertySet", "$_ChromeSTATUS_COMError", @error)
		Return SetError($_ChromeSTATUS_ComError, @error, 0)
	EndIf
	Return SetError($_ChromeSTATUS_Success, 0, 0)
EndFunc   ;==>_ChromePropertySet

; #FUNCTION# ====================================================================================================================
; Author ........: Dale Hohm
; ModifChromed ......: jpm
; ===============================================================================================================================
Func _ChromeErrorNotify($vNotify = Default)
	If $vNotify = Default Then Return $__g_bChromeErrorNotify

	If $vNotify Then
		$__g_bChromeErrorNotify = True
	Else
		$__g_bChromeErrorNotify = False
	EndIf
	Return 1
EndFunc   ;==>_ChromeErrorNotify

; #NO_DOC_FUNCTION# =============================================================================================================
; Name...........: _ChromeErrorHandlerRegister
; Description ...: Register and enable a user COM error handler
; Parameters ....: $sFunctionName - String variable with the name of a user-defined COM error handler
;									  defaults to the internal COM error handler in this UDF
; Return values .: On Success 	- Returns 1
;                  On Failure	- Returns 0 and sets @error
;					@error		- 0 ($_ChromeStatus_Success) = No Error
;								- 1 ($_ChromeStatus_GeneralError) = General Error
;					@extended	- Contains invalid parameter number
; Author ........: Dale Hohm
; ===============================================================================================================================
Func _ChromeErrorHandlerRegister($sFunctionName = "__ChromeInternalErrorHandler")
	$__g_oChromeErrorHandler = ObjEvent("AutoIt.Error", $sFunctionName)
	If IsObj($__g_oChromeErrorHandler) Then
		$__g_sChromeUserErrorHandler = $sFunctionName
		Return SetError($_ChromeSTATUS_Success, 0, 1)
	Else
		$__g_oChromeErrorHandler = ""
		__ChromeConsoleWriteError("Error", "_ChromeErrorHandlerRegister", "$_ChromeStatus_GeneralError", _
				"Error Handler Not Registered - Check existance of error function")
		Return SetError($_ChromeStatus_GeneralError, 1, 0)
	EndIf
EndFunc   ;==>_ChromeErrorHandlerRegister

; #NO_DOC_FUNCTION# =============================================================================================================
; Name...........: _ChromeErrorHandlerDeRegister
; Description ...: Disable a registered user COM error handler
; Parameters ....: None
; Return values .: On Success 	- Returns 1
;                  On Failure	- None
; Author ........: Dale Hohm
; ===============================================================================================================================
Func _ChromeErrorHandlerDeRegister()
	$__g_sChromeUserErrorHandler = ""
	$__g_oChromeErrorHandler = ""
	Return SetError($_ChromeSTATUS_Success, 0, 1)
EndFunc   ;==>_ChromeErrorHandlerDeRegister

; #INTERNAL_USE_ONLY# ===========================================================================================================
; Name...........: __ChromeInternalErrorHandlerRegister
; Description ...: to be called on error
; Author ........: Dale Hohm
; ModifChromed ......:
; ===============================================================================================================================
Func __ChromeInternalErrorHandlerRegister()
	Local $sCurrentErrorHandler = ObjEvent("AutoIt.Error")
	If $sCurrentErrorHandler <> "" And Not IsObj($__g_oChromeErrorHandler) Then
		; We've got trouble... User COM Error handler assigned without using _ChromeUserErrorHandlerRegister
		Return SetError($_ChromeStatus_GeneralError, 0, False)
	EndIf
	$__g_oChromeErrorHandler = ObjEvent("AutoIt.Error", "__ChromeInternalErrorHandler")
	If IsObj($__g_oChromeErrorHandler) Then
		Return SetError($_ChromeSTATUS_Success, 0, True)
	Else
		$__g_oChromeErrorHandler = ""
		Return SetError($_ChromeStatus_GeneralError, 0, False)
	EndIf
EndFunc   ;==>__ChromeInternalErrorHandlerRegister

Func __ChromeInternalErrorHandlerDeRegister()
	$__g_oChromeErrorHandler = ""
	If $__g_sChromeUserErrorHandler <> "" Then
		$__g_oChromeErrorHandler = ObjEvent("AutoIt.Error", $__g_sChromeUserErrorHandler)
	EndIf
	Return SetError($_ChromeSTATUS_Success, 0, 1)
EndFunc   ;==>__ChromeInternalErrorHandlerDeRegister

; #INTERNAL_USE_ONLY# ===========================================================================================================
; Name...........: __ChromeInternalErrorHandler
; Description ...: to be called on error
; Author ........: Dale Hohm
; ModifChromed ......:
; ===============================================================================================================================
Func __ChromeInternalErrorHandler($oCOMError)
	If $__g_bChromeErrorNotify Or $__g_bChromeAU3Debug Then ConsoleWrite("--> " & __COMErrorFormating($oCOMError, "----> $ChromeComError") & @CRLF)
	SetError($_ChromeStatus_ComError)
	Return
EndFunc   ;==>__ChromeInternalErrorHandler

; #FUNCTION# ====================================================================================================================
; Author ........: Dale Hohm
; ===============================================================================================================================
Func _ChromeQuit(ByRef $oObject)
;~ 	Local $sName_ChromeQuit = String(ObjName($oObject))
	If Not IsObj($oObject) Then
		__ChromeConsoleWriteError("Error", "_ChromeQuit", "$_ChromeSTATUS_InvalidDataType")
		Return SetError($_ChromeSTATUS_InvalidDataType, 1, 0)
	EndIf
	;
	If Not __ChromeIsObjType($oObject, "browser") Then
		__ChromeConsoleWriteError("Error", "_ChromeQuit", "$_ChromeSTATUS_InvalidObjectType")
		Return SetError($_ChromeSTATUS_InvalidObjectType, 1, 0)
	EndIf
	;
	$oObject.quit()
	If @error Then ; Trap COM error, report and return
		__ChromeConsoleWriteError("Error", "_ChromeQuit", "$_ChromeSTATUS_COMError", @error)
		Return SetError($_ChromeSTATUS_ComError, @error, 0)
	EndIf
	$oObject = 0
	Return SetError($_ChromeSTATUS_Success, 0, 1)
EndFunc   ;==>_ChromeQuit

#EndRegion Utility functions

#Region General
; #FUNCTION# ====================================================================================================================
; Author ........: Dale Hohm
; ModifChromed ......: jpm
; ===============================================================================================================================
Func _Chrome_Introduction($sModule = "basic")
	Local $sHTML = ""
	Switch $sModule
		Case "basic"
			$sHTML &= '<!DOCTYPE html>' & @CR
			$sHTML &= '<html>' & @CR
			$sHTML &= '<head>' & @CR
			$sHTML &= '<meta content="text/html; charset=UTF-8" http-equiv="content-type">' & @CR
			$sHTML &= '<title>_Chrome_Introduction ("basic")</title>' & @CR
			$sHTML &= '<style>body {font-family: Arial}' & @CR
			$sHTML &= 'td {padding:6px}</style>' & @CR
			$sHTML &= '</head>' & @CR
			$sHTML &= '<body>' & @CR
			$sHTML &= '<table border=1 id="table1" style="width:600px;border-spacing:6px;">' & @CR
			$sHTML &= '<tr>' & @CR
			$sHTML &= '<td>' & @CR
			$sHTML &= '<h1>Welcome to Chrome.au3</h1>' & @CR
			$sHTML &= 'Chrome.au3 is a UDF (User Defined Function) library for the ' & @CR
			$sHTML &= '<a href="http://www.autoitscript.com">AutoIt</a> scripting language.' & @CR
			$sHTML &= '<br>  ' & @CR
			$sHTML &= 'Chrome.au3 allows you to either create or attach to an Chrome browser and do ' & @CR
			$sHTML &= 'just about anything you could do with it interactively with the mouse and ' & @CR
			$sHTML &= 'keyboard, but do it through script.' & @CR
			$sHTML &= '<br>' & @CR
			$sHTML &= 'You can navigate to pages, click links, fill and submit forms etc. You can ' & @CR
			$sHTML &= 'also do things you cannot do interactively like change or rewrite page ' & @CR
			$sHTML &= 'content and JavaScripts, read, parse and save page content and monitor and act ' & @CR
			$sHTML &= 'upon browser "events".<br>' & @CR
			$sHTML &= 'Chrome.au3 uses the COM interface in AutoIt to interact with the Chrome ' & @CR
			$sHTML &= 'object model and the DOM (Document Object Model) supported by the browser.' & @CR
			$sHTML &= '<br>' & @CR
			$sHTML &= 'Here are some links for more information and helpful tools:<br>' & @CR
			$sHTML &= 'Reference Material: ' & @CR
			$sHTML &= '<ul>' & @CR
			$sHTML &= '<li><a href="http://msdn1.microsoft.com/">MSDN (Microsoft Developer Network)</a></li>' & @CR
			$sHTML &= '<li><a href="http://msdn2.microsoft.com/en-us/library/aa752084.aspx" target="_blank">InternetExplorer Object</a></li>' & @CR
			$sHTML &= '<li><a href="http://msdn2.microsoft.com/en-us/library/ms531073.aspx" target="_blank">Document Object</a></li>' & @CR
			$sHTML &= '<li><a href="http://msdn2.microsoft.com/en-us/Chrome/aa740473.aspx" target="_blank">OvervChromews and Tutorials</a></li>' & @CR
			$sHTML &= '<li><a href="http://msdn2.microsoft.com/en-us/library/ms533029.aspx" target="_blank">DHTML Objects</a></li>' & @CR
			$sHTML &= '<li><a href="http://msdn2.microsoft.com/en-us/library/ms533051.aspx" target="_blank">DHTML Events</a></li>' & @CR
			$sHTML &= '</ul><br>' & @CR
			$sHTML &= 'Helpful Tools: ' & @CR
			$sHTML &= '<ul>' & @CR
			$sHTML &= '<li><a href="http://www.autoitscript.com/forum/index.php?showtopic=19368" target="_blank">AutoIt Chrome Builder</a> (build Chrome scripts interactively)</li>' & @CR
			$sHTML &= '<li><a href="http://www.debugbar.com/" target="_blank">DebugBar</a> (DOM inspector, HTTP inspector, HTML validator and more - free for personal use) Recommended</li>' & @CR
			$sHTML &= '<li><a href="http://www.microsoft.com/downloads/details.aspx?FamilyID=e59c3964-672d-4511-bb3e-2d5e1db91038&amp;displaylang=en" target="_blank">Chrome Developer Toolbar</a> (comprehensive DOM analysis tool)</li>' & @CR
			$sHTML &= '<li><a href="http://slayeroffice.com/tools/modi/v2.0/modi_help.html" target="_blank">MODIV2</a> (vChromew the DOM of a web page by mousing around)</li>' & @CR
			$sHTML &= '<li><a href="http://validator.w3.org/" target="_blank">HTML Validator</a> (verify HTML follows format rules)</li>' & @CR
			$sHTML &= '<li><a href="http://www.fiddlertool.com/fiddler/" target="_blank">Fiddler</a> (examine HTTP traffic)</li>' & @CR
			$sHTML &= '</ul>' & @CR
			$sHTML &= '</td>' & @CR
			$sHTML &= '</tr>' & @CR
			$sHTML &= '</table>' & @CR
			$sHTML &= '</body>' & @CR
			$sHTML &= '</html>'
		Case Else
			__ChromeConsoleWriteError("Error", "_Chrome_Introduction", "$_ChromeSTATUS_InvalidValue")
			Return SetError($_ChromeSTATUS_InvalidValue, 1, 0)
	EndSwitch
	Local $oObject = _ChromeCreate()
	_ChromeDocWriteHTML($oObject, $sHTML)
	Return SetError($_ChromeSTATUS_Success, 0, $oObject)
EndFunc   ;==>_Chrome_Introduction

; #FUNCTION# ====================================================================================================================
; Author ........: Dale Hohm
; ModifChromed ......: jpm
; ===============================================================================================================================
Func _Chrome_Example($sModule = "basic")
	Local $sHTML = "", $oObject
	Switch $sModule
		Case "basic"
			$sHTML &= '<!DOCTYPE html>' & @CR
			$sHTML &= '<html>' & @CR
			$sHTML &= '<head>' & @CR
			$sHTML &= '<meta content="text/html; charset=UTF-8" http-equiv="content-type">' & @CR
			$sHTML &= '<title>_Chrome_Example("basic")</title>' & @CR
			$sHTML &= '<style>body {font-family: Arial}</style>' & @CR
			$sHTML &= '</head>' & @CR
			$sHTML &= '<body>' & @CR
			$sHTML &= '<a href="http://www.autoitscript.com"><img src="http://www.autoitscript.com/images/autoit_6_240x100.jpg" id="AutoItImage" alt="AutoIt Homepage Image"></a>' & @CR
			$sHTML &= '<p></p>' & @CR
			$sHTML &= '<div id="line1">This is a simple HTML page with text, links and images.</div>' & @CR
			$sHTML &= '<br>' & @CR
			$sHTML &= '<div id="line2"><a href="http://www.autoitscript.com">AutoIt</a> is a wonderful automation scripting language.</div>' & @CR
			$sHTML &= '<br>' & @CR
			$sHTML &= '<div id="line3">It is supported by a very active and supporting <a href="http://www.autoitscript.com/forum/">user forum</a>.</div>' & @CR
			$sHTML &= '<br>' & @CR
			$sHTML &= '<div id="ChromeAu3Data"></div>' & @CR
			$sHTML &= '</body>' & @CR
			$sHTML &= '</html>'
			$oObject = _ChromeCreate()
			_ChromeDocWriteHTML($oObject, $sHTML)
		Case "table"
			$sHTML &= '<!DOCTYPE html>' & @CR
			$sHTML &= '<html>' & @CR
			$sHTML &= '<head>' & @CR
			$sHTML &= '<meta content="text/html; charset=utf-8" http-equiv="content-type">' & @CR
			$sHTML &= '<title>_Chrome_Example("table")</title>' & @CR
			$sHTML &= '<style>body {font-family: Arial}</style>' & @CR
			$sHTML &= '</head>' & @CR
			$sHTML &= '<body>' & @CR
			$sHTML &= '$oTableOne = _ChromeTableGetObjByName($oChrome, "tableOne")<br>' & @CR
			$sHTML &= '&lt;table border=1 id="tableOne"&gt;<br>' & @CR
			$sHTML &= '<table border=1 id="tableOne">' & @CR
			$sHTML &= '	<tr>' & @CR
			$sHTML &= '		<td>AutoIt</td>' & @CR
			$sHTML &= '		<td>is</td>' & @CR
			$sHTML &= '		<td>really</td>' & @CR
			$sHTML &= '		<td>great</td>' & @CR
			$sHTML &= '		<td>with</td>' & @CR
			$sHTML &= '		<td>Chrome.au3</td>' & @CR
			$sHTML &= '	</tr>' & @CR
			$sHTML &= '	<tr>' & @CR
			$sHTML &= '		<td>1</td>' & @CR
			$sHTML &= '		<td>2</td>' & @CR
			$sHTML &= '		<td>3</td>' & @CR
			$sHTML &= '		<td>4</td>' & @CR
			$sHTML &= '		<td>5</td>' & @CR
			$sHTML &= '		<td>6</td>' & @CR
			$sHTML &= '	</tr>' & @CR
			$sHTML &= '	<tr>' & @CR
			$sHTML &= '		<td>the</td>' & @CR
			$sHTML &= '		<td>quick</td>' & @CR
			$sHTML &= '		<td>red</td>' & @CR
			$sHTML &= '		<td>fox</td>' & @CR
			$sHTML &= '		<td>jumped</td>' & @CR
			$sHTML &= '		<td>over</td>' & @CR
			$sHTML &= '	</tr>' & @CR
			$sHTML &= '	<tr>' & @CR
			$sHTML &= '		<td>the</td>' & @CR
			$sHTML &= '		<td>lazy</td>' & @CR
			$sHTML &= '		<td>brown</td>' & @CR
			$sHTML &= '		<td>dog</td>' & @CR
			$sHTML &= '		<td>the</td>' & @CR
			$sHTML &= '		<td>time</td>' & @CR
			$sHTML &= '	</tr>' & @CR
			$sHTML &= '	<tr>' & @CR
			$sHTML &= '		<td>has</td>' & @CR
			$sHTML &= '		<td>come</td>' & @CR
			$sHTML &= '		<td>for</td>' & @CR
			$sHTML &= '		<td>all</td>' & @CR
			$sHTML &= '		<td>good</td>' & @CR
			$sHTML &= '		<td>men</td>' & @CR
			$sHTML &= '	</tr>' & @CR
			$sHTML &= '	<tr>' & @CR
			$sHTML &= '		<td>to</td>' & @CR
			$sHTML &= '		<td>come</td>' & @CR
			$sHTML &= '		<td>to</td>' & @CR
			$sHTML &= '		<td>the</td>' & @CR
			$sHTML &= '		<td>aid</td>' & @CR
			$sHTML &= '		<td>of</td>' & @CR
			$sHTML &= '	</tr>' & @CR
			$sHTML &= '</table>' & @CR
			$sHTML &= '<br>' & @CR
			$sHTML &= '$oTableTwo = _ChromeTableGetObjByName($oChrome, "tableTwo")<br>' & @CR
			$sHTML &= '&lt;table border="1" id="tableTwo"&gt;<br>' & @CR
			$sHTML &= '<table border=1 id="tableTwo">' & @CR
			$sHTML &= '	<tr>' & @CR
			$sHTML &= '		<td colspan="4">Table Top</td>' & @CR
			$sHTML &= '	</tr>' & @CR
			$sHTML &= '	<tr>' & @CR
			$sHTML &= '		<td>One</td>' & @CR
			$sHTML &= '		<td colspan="3">Two</td>' & @CR
			$sHTML &= '	</tr>' & @CR
			$sHTML &= '	<tr>' & @CR
			$sHTML &= '		<td>Three</td>' & @CR
			$sHTML &= '		<td>Four</td>' & @CR
			$sHTML &= '		<td colspan="2">Five</td>' & @CR
			$sHTML &= '	</tr>' & @CR
			$sHTML &= '	<tr>' & @CR
			$sHTML &= '		<td>Six</td>' & @CR
			$sHTML &= '		<td colspan="3">Seven</td>' & @CR
			$sHTML &= '	</tr>' & @CR
			$sHTML &= '	<tr>' & @CR
			$sHTML &= '		<td>Eight</td>' & @CR
			$sHTML &= '		<td>Nine</td>' & @CR
			$sHTML &= '		<td>Ten</td>' & @CR
			$sHTML &= '		<td>Eleven</td>' & @CR
			$sHTML &= '	</tr>' & @CR
			$sHTML &= '</table>' & @CR
			$sHTML &= '</body>' & @CR
			$sHTML &= '</html>'
			$oObject = _ChromeCreate()
			_ChromeDocWriteHTML($oObject, $sHTML)
		Case "form"
			$sHTML &= '<!DOCTYPE html>' & @CR
			$sHTML &= '<html>' & @CR
			$sHTML &= '<head>' & @CR
			$sHTML &= '<meta content="text/html; charset=UTF-8" http-equiv="content-type">' & @CR
			$sHTML &= '<title>_Chrome_Example("form")</title>' & @CR
			$sHTML &= '<style>body {font-family: Arial}' & @CR
			$sHTML &= 'td {padding:6px}</style>' & @CR
			$sHTML &= '</head>' & @CR
			$sHTML &= '<body>' & @CR
			$sHTML &= '<form name="ExampleForm" onSubmit="javascript:alert(''ExampleFormSubmitted'');" method="post">' & @CR
			$sHTML &= '<table style="border-spacing:6px 6px;" border=1>' & @CR
			$sHTML &= '<tr>' & @CR
			$sHTML &= '<td>ExampleForm</td>' & @CR
			$sHTML &= '<td>&lt;form name="ExampleForm" onSubmit="javascript:alert(''ExampleFormSubmitted'');" method="post"&gt;</td>' & @CR
			$sHTML &= '</tr>' & @CR
			$sHTML &= '<tr>' & @CR
			$sHTML &= '<td>Hidden Input Element<input type="hidden" name="hiddenExample" value="secret value"></td>' & @CR
			$sHTML &= '<td>&lt;input type="hidden" name="hiddenExample" value="secret value"&gt;</td>' & @CR
			$sHTML &= '</tr>' & @CR
			$sHTML &= '<tr>' & @CR
			$sHTML &= '<td>' & @CR
			$sHTML &= '<input type="text" name="textExample" value="http://" size="20" maxlength="30">' & @CR
			$sHTML &= '</td>' & @CR
			$sHTML &= '<td>&lt;input type="text" name="textExample" value="http://" size="20" maxlength="30"&gt;</td>' & @CR
			$sHTML &= '</tr>' & @CR
			$sHTML &= '<tr>' & @CR
			$sHTML &= '<td>' & @CR
			$sHTML &= '<input type="password" name="passwordExample" size="10">' & @CR
			$sHTML &= '</td>' & @CR
			$sHTML &= '<td>&lt;input type="password" name="passwordExample" size="10"&gt;</td>' & @CR
			$sHTML &= '</tr>' & @CR
			$sHTML &= '<tr>' & @CR
			$sHTML &= '<td>' & @CR
			$sHTML &= '<input type="file" name="fileExample">' & @CR
			$sHTML &= '</td>' & @CR
			$sHTML &= '<td>&lt;input type="file" name="fileExample"&gt;</td>' & @CR
			$sHTML &= '</tr>' & @CR
			$sHTML &= '<tr>' & @CR
			$sHTML &= '<td>' & @CR
			$sHTML &= '<input type="image" name="imageExample" alt="AutoIt Homepage" src="http://www.autoitscript.com/images/autoit_6_240x100.jpg">' & @CR
			$sHTML &= '</td>' & @CR
			$sHTML &= '<td>&lt;input type="image" name="imageExample" alt="AutoIt Homepage" src="http://www.autoitscript.com/images/autoit_6_240x100.jpg"&gt;</td>' & @CR
			$sHTML &= '</tr>' & @CR
			$sHTML &= '<tr>' & @CR
			$sHTML &= '<td>' & @CR
			$sHTML &= '<textarea name="textareaExample" rows="5" cols="15">Hello!</textarea>' & @CR
			$sHTML &= '</td>' & @CR
			$sHTML &= '<td>&lt;textarea name="textareaExample" rows="5" cols="15"&gt;Hello!&lt;/textarea&gt;</td>' & @CR
			$sHTML &= '</tr>' & @CR
			$sHTML &= '<tr>' & @CR
			$sHTML &= '<td>' & @CR
			$sHTML &= '<input type="checkbox" name="checkboxG1Example" value="gameBasketball">Basketball<br>' & @CR
			$sHTML &= '<input type="checkbox" name="checkboxG1Example" value="gameFootball">Football<br>' & @CR
			$sHTML &= '<input type="checkbox" name="checkboxG2Example" value="gameTennis" checked>Tennis<br>' & @CR
			$sHTML &= '<input type="checkbox" name="checkboxG2Example" value="gameBaseball">Baseball' & @CR
			$sHTML &= '</td>' & @CR
			$sHTML &= '<td>&lt;input type="checkbox" name="checkboxG1Example" value="gameBasketball"&gt;Basketball&lt;br&gt;<br>' & @CR
			$sHTML &= '&lt;input type="checkbox" name="checkboxG1Example" value="gameFootball"&gt;Football&lt;br&gt;<br>' & @CR
			$sHTML &= '&lt;input type="checkbox" name="checkboxG2Example" value="gameTennis" checked&gt;Tennis&lt;br&gt;<br>' & @CR
			$sHTML &= '&lt;input type="checkbox" name="checkboxG2Example" value="gameBaseball"&gt;Baseball</td>' & @CR
			$sHTML &= '</tr>' & @CR
			$sHTML &= '<tr>' & @CR
			$sHTML &= '<td>' & @CR
			$sHTML &= '<input type="radio" name="radioExample" value="vehicleAirplane">Airplane<br>' & @CR
			$sHTML &= '<input type="radio" name="radioExample" value="vehicleTrain" checked>Train<br>' & @CR
			$sHTML &= '<input type="radio" name="radioExample" value="vehicleBoat">Boat<br>' & @CR
			$sHTML &= '<input type="radio" name="radioExample" value="vehicleCar">Car</td>' & @CR
			$sHTML &= '<td>&lt;input type="radio" name="radioExample" value="vehicleAirplane"&gt;Airplane&lt;br&gt;<br>' & @CR
			$sHTML &= '&lt;input type="radio" name="radioExample" value="vehicleTrain" checked&gt;Train&lt;br&gt;<br>' & @CR
			$sHTML &= '&lt;input type="radio" name="radioExample" value="vehicleBoat"&gt;Boat&lt;br&gt;<br>' & @CR
			$sHTML &= '&lt;input type="radio" name="radioExample" value="vehicleCar"&gt;Car&lt;br&gt;</td>' & @CR
			$sHTML &= '</tr>' & @CR
			$sHTML &= '<tr>' & @CR
			$sHTML &= '<td>' & @CR
			$sHTML &= '<select name="selectExample">' & @CR
			$sHTML &= '<option value="homepage.html">Homepage' & @CR
			$sHTML &= '<option value="midipage.html">Midipage' & @CR
			$sHTML &= '<option value="freepage.html">Freepage' & @CR
			$sHTML &= '</select>' & @CR
			$sHTML &= '</td>' & @CR
			$sHTML &= '<td>&lt;select name="selectExample"&gt;<br>' & @CR
			$sHTML &= '&lt;option value="homepage.html"&gt;Homepage<br>' & @CR
			$sHTML &= '&lt;option value="midipage.html"&gt;Midipage<br>' & @CR
			$sHTML &= '&lt;option value="freepage.html"&gt;Freepage<br>' & @CR
			$sHTML &= '&lt;/select&gt;</td>' & @CR
			$sHTML &= '</tr>' & @CR
			$sHTML &= '<tr>' & @CR
			$sHTML &= '<td>' & @CR
			$sHTML &= '<select name="multipleSelectExample" size="6" multiple>' & @CR
			$sHTML &= '<option value="Name1">Aaron' & @CR
			$sHTML &= '<option value="Name2">Bruce' & @CR
			$sHTML &= '<option value="Name3">Carlos' & @CR
			$sHTML &= '<option value="Name4">Denis' & @CR
			$sHTML &= '<option value="Name5">Ed' & @CR
			$sHTML &= '<option value="Name6">Freddy' & @CR
			$sHTML &= '</select>' & @CR
			$sHTML &= '</td>' & @CR
			$sHTML &= '<td>&lt;select name="multipleSelectExample" size="6" multiple&gt;<br>' & @CR
			$sHTML &= '&lt;option value="Name1"&gt;Aaron<br>' & @CR
			$sHTML &= '&lt;option value="Name2"&gt;Bruce<br>' & @CR
			$sHTML &= '&lt;option value="Name3"&gt;Carlos<br>' & @CR
			$sHTML &= '&lt;option value="Name4"&gt;Denis<br>' & @CR
			$sHTML &= '&lt;option value="Name5"&gt;Ed<br>' & @CR
			$sHTML &= '&lt;option value="Name6"&gt;Freddy<br>' & @CR
			$sHTML &= '&lt;/select&gt;</td>' & @CR
			$sHTML &= '</tr>' & @CR
			$sHTML &= '<tr>' & @CR
			$sHTML &= '<td>' & @CR
			$sHTML &= '<input name="submitExample" type="submit" value="Submit">' & @CR
			$sHTML &= '<input name="resetExample" type="reset" value="Reset">' & @CR
			$sHTML &= '</td>' & @CR
			$sHTML &= '<td>&lt;input name="submitExample" type="submit" value="Submit"&gt;<br>' & @CR
			$sHTML &= '&lt;input name="resetExample" type="reset" value="Reset"&gt;</td>' & @CR
			$sHTML &= '</tr>' & @CR
			$sHTML &= '</table>' & @CR
			$sHTML &= '<input type="hidden" name="hiddenExample" value="secret value">' & @CR
			$sHTML &= '</form>' & @CR
			$sHTML &= '</body>' & @CR
			$sHTML &= '</html>'
			$oObject = _ChromeCreate()
			_ChromeDocWriteHTML($oObject, $sHTML)
		Case "frameset"
			$sHTML &= '<!DOCTYPE html>' & @CR
			$sHTML &= '<html>' & @CR
			$sHTML &= '<head>' & @CR
			$sHTML &= '<meta content="text/html; charset=UTF-8" http-equiv="content-type">' & @CR
			$sHTML &= '<title>_Chrome_Example("frameset")</title>' & @CR
			$sHTML &= '</head>' & @CR
			$sHTML &= '<frameset rows="25,200">' & @CR
			$sHTML &= '	<frame name=Top SRC=about:blank>' & @CR
			$sHTML &= '	<frameset cols="100,500">' & @CR
			$sHTML &= '		<frame name=Menu SRC=about:blank>' & @CR
			$sHTML &= '		<frame name=Main SRC=about:blank>' & @CR
			$sHTML &= '	</frameset>' & @CR
			$sHTML &= '</frameset>' & @CR
			$sHTML &= '</html>'
			$oObject = _ChromeCreate()
			_ChromeDocWriteHTML($oObject, $sHTML)
			_ChromeAction($oObject, "refresh")
			Local $oFrameTop = _ChromeFrameGetObjByName($oObject, "Top")
			Local $oFrameMenu = _ChromeFrameGetObjByName($oObject, "Menu")
			Local $oFrameMain = _ChromeFrameGetObjByName($oObject, "Main")
			_ChromeBodyWriteHTML($oFrameTop, '$oFrameTop = _ChromeFrameGetObjByName($oChrome, "Top")')
			_ChromeBodyWriteHTML($oFrameMenu, '$oFrameMenu = _ChromeFrameGetObjByName($oChrome, "Menu")')
			_ChromeBodyWriteHTML($oFrameMain, '$oFrameMain = _ChromeFrameGetObjByName($oChrome, "Main")')
		Case "iframe"
			$sHTML &= '<!DOCTYPE html>' & @CR
			$sHTML &= '<html>' & @CR
			$sHTML &= '<head>' & @CR
			$sHTML &= '<meta content="text/html; charset=UTF-8" http-equiv="content-type">' & @CR
			$sHTML &= '<title>_Chrome_Example("iframe")</title>' & @CR
			$sHTML &= '<style>td {padding:6px}</style>' & @CR
			$sHTML &= '</head>' & @CR
			$sHTML &= '<body>' & @CR
			$sHTML &= '<table style="border-spacing:6px" border=1>' & @CR
			$sHTML &= '<tr>' & @CR
			$sHTML &= '<td><iframe name="iFrameOne" src="about:blank" title="iFrameOne"></iframe></td>' & @CR
			$sHTML &= '<td>&lt;iframe name="iFrameOne" src="about:blank" title="iFrameOne"&gt;</td>' & @CR
			$sHTML &= '</tr>' & @CR
			$sHTML &= '<tr>' & @CR
			$sHTML &= '<td><iframe name="iFrameTwo" src="about:blank" title="iFrameTwo"></iframe></td>' & @CR
			$sHTML &= '<td>&lt;iframe name="iFrameTwo" src="about:blank" title="iFrameTwo"&gt;</td>' & @CR
			$sHTML &= '</tr>' & @CR
			$sHTML &= '</table>' & @CR
			$sHTML &= '</body>' & @CR
			$sHTML &= '</html>'
			$oObject = _ChromeCreate()
			_ChromeDocWriteHTML($oObject, $sHTML)
			_ChromeAction($oObject, "refresh")
			Local $oIFrameOne = _ChromeFrameGetObjByName($oObject, "iFrameOne")
			Local $oIFrameTwo = _ChromeFrameGetObjByName($oObject, "iFrameTwo")
			_ChromeBodyWriteHTML($oIFrameOne, '$oIFrameOne = _ChromeFrameGetObjByName($oChrome, "iFrameOne")')
			_ChromeBodyWriteHTML($oIFrameTwo, '$oIFrameTwo = _ChromeFrameGetObjByName($oChrome, "iFrameTwo")')
		Case Else
			__ChromeConsoleWriteError("Error", "_Chrome_Example", "$_ChromeSTATUS_InvalidValue")
			Return SetError($_ChromeSTATUS_InvalidValue, 1, 0)
	EndSwitch

	;	at least under Chrome10 some delay is needed to have functions as _ChromePropertySet() working
	;	value can depend of processor speed ...
	Sleep(500)
	Return SetError($_ChromeSTATUS_Success, 0, $oObject)
EndFunc   ;==>_Chrome_Example

; #FUNCTION# ====================================================================================================================
; Author ........: Dale Hohm
; ===============================================================================================================================
Func _Chrome_VersionInfo()
	__ChromeConsoleWriteError("Information", "_Chrome_VersionInfo", "version " & _
			$__gaChromeAU3VersionInfo[0] & _
			$__gaChromeAU3VersionInfo[1] & "." & _
			$__gaChromeAU3VersionInfo[2] & "-" & _
			$__gaChromeAU3VersionInfo[3], "Release date: " & $__gaChromeAU3VersionInfo[4])
	Return SetError($_ChromeSTATUS_Success, 0, $__gaChromeAU3VersionInfo)
EndFunc   ;==>_Chrome_VersionInfo

#EndRegion General

#Region Internal functions
;
; Internal Functions with names starting with two underscores will not be documented
; as user functions
;
; #INTERNAL_USE_ONLY# ===========================================================================================================
; Name...........: __ChromeLockSetForegroundWindow
; Description ...: Locks (and Unlocks) current Foregrouns Window focus to prevent a new window
; 					from stealing it (e.g. when creating invisible Chrome browser)
; Parameters ....: $iLockCode	- 1 Lock Foreground Window Focus, 2 Unlock Foreground Window Focus
; Return values .: On Success 	- 1
;                   On Failure 	- 0  and sets @error and @extended to non-zero values
; Author ........: Valik
; ===============================================================================================================================
Func __ChromeLockSetForegroundWindow($iLockCode)
	Local $aRet = DllCall("user32.dll", "bool", "LockSetForegroundWindow", "uint", $iLockCode)
	If @error Or Not $aRet[0] Then Return SetError(1, _WinAPI_GetLastError(), 0)
	Return $aRet[0]
EndFunc   ;==>__ChromeLockSetForegroundWindow

; #INTERNAL_USE_ONLY# ===========================================================================================================
; Name...........: __ChromeControlGetObjFromHWND
; Description ...: Returns a COM Object Window reference to an embebedded Webbrowser control
; Parameters ....: $hWin		- HWND of a Chrome_Server1 control obtained for example:
; 					$hwnd = ControlGetHandle("MyApp","","Chrome_Server1")
; Return values .: On Success 	- Returns DOM Window object
;                   On Failure 	- 0  and sets @error = 1
; Author ........: Larry with thanks to Valik
; Remarks .......:
; ===============================================================================================================================
Func __ChromeControlGetObjFromHWND(ByRef $hWin)
	; The code assumes CoInitialize() succeeded due to the number of different
	; yet successful return values it has.
	DllCall("ole32.dll", "long", "CoInitialize", "ptr", 0)
	If @error Then Return SetError(2, @error, 0)

	Local Const $WM_HTML_GETOBJECT = __ChromeRegisterWindowMessage("WM_HTML_GETOBJECT")
	Local Const $SMTO_ABORTIFHUNG = 0x0002
	Local $iResult

	__ChromeSendMessageTimeout($hWin, $WM_HTML_GETOBJECT, 0, 0, $SMTO_ABORTIFHUNG, 1000, $iResult)

	Local $tUUID = DllStructCreate("int;short;short;byte[8]")
	DllStructSetData($tUUID, 1, 0x626FC520)
	DllStructSetData($tUUID, 2, 0xA41E)
	DllStructSetData($tUUID, 3, 0x11CF)
	DllStructSetData($tUUID, 4, 0xA7, 1)
	DllStructSetData($tUUID, 4, 0x31, 2)
	DllStructSetData($tUUID, 4, 0x0, 3)
	DllStructSetData($tUUID, 4, 0xA0, 4)
	DllStructSetData($tUUID, 4, 0xC9, 5)
	DllStructSetData($tUUID, 4, 0x8, 6)
	DllStructSetData($tUUID, 4, 0x26, 7)
	DllStructSetData($tUUID, 4, 0x37, 8)

	Local $aRet = DllCall("oleacc.dll", "long", "ObjectFromLresult", "lresult", $iResult, "struct*", $tUUID, _
			"wparam", 0, "idispatch*", 0)
	If @error Then Return SetError(3, @error, 0)

	If IsObj($aRet[4]) Then
		Local $oChrome = $aRet[4].Script()
		; $oChrome is now a valid IDispatch object
		Return $oChrome.Document.parentwindow
	Else
		Return SetError(1, $aRet[0], 0)
	EndIf
EndFunc   ;==>__ChromeControlGetObjFromHWND

; #INTERNAL_USE_ONLY# ===========================================================================================================
; Name...........: __ChromeRegisterWindowMessage
; Description ...: Required by __ChromeControlGetObjFromHWND()
; Author ........: Larry with thanks to Valik
; ===============================================================================================================================
Func __ChromeRegisterWindowMessage($sMsg)
	Local $aRet = DllCall("user32.dll", "uint", "RegisterWindowMessageW", "wstr", $sMsg)
	If @error Then Return SetError(@error, @extended, 0)
	If $aRet[0] = 0 Then Return SetError(10, _WinAPI_GetLastError(), 0)
	Return $aRet[0]
EndFunc   ;==>__ChromeRegisterWindowMessage

; #INTERNAL_USE_ONLY# ===========================================================================================================
; Name...........: __ChromeSendMessageTimeout
; Description ...: Required by __ChromeControlGetObjFromHWND()
; Author ........: Larry with thanks to Valik
; ===============================================================================================================================
Func __ChromeSendMessageTimeout($hWnd, $iMsg, $wParam, $lParam, $iFlags, $iTimeout, ByRef $vOut, $r = 0, $sT1 = "int", $sT2 = "int")
	Local $aRet = DllCall("user32.dll", "lresult", "SendMessageTimeout", "hwnd", $hWnd, "uint", $iMsg, $sT1, $wParam, _
			$sT2, $lParam, "uint", $iFlags, "uint", $iTimeout, "dword_ptr*", "")
	If @error Or $aRet[0] = 0 Then
		$vOut = 0
		Return SetError(1, _WinAPI_GetLastError(), 0)
	EndIf
	$vOut = $aRet[7]
	If $r >= 0 And $r <= 4 Then Return $aRet[$r]
	Return $aRet
EndFunc   ;==>__ChromeSendMessageTimeout

; #INTERNAL_USE_ONLY# ===========================================================================================================
; Name...........: __ChromeIsObjType
; Description ...: Check to see if an object variable is of a specific type
; Author ........: Dale Hohm
; ModifChromed ......: jpm
; ===============================================================================================================================
Func __ChromeIsObjType(ByRef $oObject, $sType, $bRegister = True)
	If Not IsObj($oObject) Then
		Return SetError($_ChromeSTATUS_InvalidDataType, 1, 0)
	EndIf

	; Setup internal error handler to Trap COM errors, turn off error notification
	Local $bStatus = $bRegister
	If $bRegister Then
		$bStatus = __ChromeInternalErrorHandlerRegister()
		If Not $bStatus Then __ChromeConsoleWriteError("Warning", "internal function __ChromeIsObjType", _
				"Cannot register internal error handler, cannot trap COM errors", _
				"Use _ChromeErrorHandlerRegister() to register a user error handler")
	EndIf

	Local $iNotifyStatus = _ChromeErrorNotify() ; save current error notify status
	_ChromeErrorNotify(False)
	;
	Local $sName = String(ObjName($oObject)), $ChromeErrorStatus = $_ChromeSTATUS_InvalidObjectType

	Switch $sType
		Case "browserdom"
			If __ChromeIsObjType($oObject, "documentcontainer", False) Then
				$ChromeErrorStatus = $_ChromeSTATUS_Success
			ElseIf __ChromeIsObjType($oObject, "document", False) Then
				$ChromeErrorStatus = $_ChromeSTATUS_Success
			Else
				Local $oTemp = $oObject.document
				If __ChromeIsObjType($oTemp, "document", False) Then
					$ChromeErrorStatus = $_ChromeSTATUS_Success
				EndIf
			EndIf
		Case "browser"
			If ($sName = "IWebBrowser2") Or ($sName = "IWebBrowser") Or ($sName = "WebBrowser") Then $ChromeErrorStatus = $_ChromeSTATUS_Success
		Case "window"
			If $sName = "HTMLWindow2" Then $ChromeErrorStatus = $_ChromeSTATUS_Success
		Case "documentContainer"
			If __ChromeIsObjType($oObject, "window", False) Or __ChromeIsObjType($oObject, "browser", False) Then $ChromeErrorStatus = $_ChromeSTATUS_Success
		Case "document"
			If $sName = "HTMLDocument" Then $ChromeErrorStatus = $_ChromeSTATUS_Success
		Case "table"
			If $sName = "HTMLTable" Then $ChromeErrorStatus = $_ChromeSTATUS_Success
		Case "form"
			If $sName = "HTMLFormElement" Then $ChromeErrorStatus = $_ChromeSTATUS_Success
		Case "forminputelement"
			If ($sName = "HTMLInputElement") Or ($sName = "HTMLSelectElement") Or ($sName = "HTMLTextAreaElement") Then $ChromeErrorStatus = $_ChromeSTATUS_Success
		Case "elementcollection"
			If ($sName = "HTMLElementCollection") Then $ChromeErrorStatus = $_ChromeSTATUS_Success
		Case "formselectelement"
			If $sName = "HTMLSelectElement" Then $ChromeErrorStatus = $_ChromeSTATUS_Success
		Case Else
			; Unsupported ObjType specifChromed
			$ChromeErrorStatus = $_ChromeSTATUS_InvalidValue
	EndSwitch

	; restore error notify
	_ChromeErrorNotify($iNotifyStatus) ; restore notification status

	If $bRegister Then
		__ChromeInternalErrorHandlerDeRegister()
	EndIf

	If $ChromeErrorStatus = $_ChromeSTATUS_Success Then
		Return SetError($_ChromeSTATUS_Success, 0, 1)
	Else
		Return SetError($ChromeErrorStatus, 1, 0)
	EndIf
EndFunc   ;==>__ChromeIsObjType

; #INTERNAL_USE_ONLY# ===========================================================================================================
; Name...........: __ChromeConsoleWriteError
; Description ...: ConsoleWrite an error message if required
; Author ........: Dale Hohm
; ModifChromed ......: jpm
; ===============================================================================================================================
Func __ChromeConsoleWriteError($sSeverity, $sFunc, $sMessage = Default, $sStatus = Default)
	If $__g_bChromeErrorNotify Or $__g_bChromeAU3Debug Then
		Local $sStr = "--> Chrome.au3 " & $__gaChromeAU3VersionInfo[5] & " " & $sSeverity & " from function " & $sFunc
		If Not ($sMessage = Default) Then $sStr &= ", " & $sMessage
		If Not ($sStatus = Default) Then $sStr &= " (" & $sStatus & ")"
		ConsoleWrite($sStr & @CRLF)
	EndIf
	Return SetError($sStatus, 0, 1) ; restore calling @error
EndFunc   ;==>__ChromeConsoleWriteError

; #INTERNAL_USE_ONLY# ===========================================================================================================
; Name...........: __ChromeComErrorUnrecoverable
; Description ...: Internal function to test a COM error condition and determine if it is considered unrecoverable
; Parameters ....: Error number
; Return values .: Unrecoverable: True, Else: False
; Author ........: Dale Hohm
; ModifChromed ......: jpm
; ===============================================================================================================================
Func __ChromeComErrorUnrecoverable($ChromeError)
	Switch $ChromeError
		; Cross-domain scripting security error
		Case -2147352567 ; "an exception has occurred."
			Return $_ChromeSTATUS_AccessIsDenChromed
		Case -2147024891 ; "Access is denChromed."
			Return $_ChromeSTATUS_AccessIsDenChromed
			;
			; Browser object is destroyed before we try to operate upon it
		Case -2147417848 ; "The object invoked has disconnected from its clChroments."
			Return $_ChromeSTATUS_ClChromentDisconnected
		Case -2147023174 ; "RPC server not accessible."
			Return $_ChromeSTATUS_ClChromentDisconnected
		Case -2147023179 ; "The interface is unknown."
			Return $_ChromeSTATUS_ClChromentDisconnected
			;
		Case Else
			Return $_ChromeSTATUS_Success
	EndSwitch
EndFunc   ;==>__ChromeComErrorUnrecoverable

#EndRegion Internal functions

#Region ProtoType Functions
; #INTERNAL_USE_ONLY# ===========================================================================================================
; Name...........: __ChromeNavigate
; Description ...: ** Unsupported version of _ChromeNavigate (note second underscore in function name)
; 					** Last 4 parameters insufficChromently tested.
; 					**    - Flags and Target can create new windows and new browser object - causing confusion
; 					**    - Postdata needs SAFEARRAY and we have no way to create one
; 					Directs an existing browser window to navigate to the specifChromed URL
; Parameters ....: $oObject 		- Object variable of an InternetExplorer.Application, Window or Frame object
; 				   $sUrl 			- URL to navigate to (e.g. "http://www.autoitscript.com")
; 				   $iWait 		- Optional: specifChromes whether to wait for page to load before returning
; 										0 = Return immediately, not waiting for page to load
; 										1 = (Default) Wait for page load to complete before returning
; 				   $iFags		- URL to navigate to (e.g. "http://www.autoitscript.com")
; 				   $sTarget	- target frame
; 				   $spostdata	- data for form method="POST", non-functional - requires safearray
; 				   $sHeaders	- additional headers to be passed
; Return values .: On Success 	- Returns -1
;                  On Failure	- Returns 0 and sets @error
; 					@error		- 1 ($_ChromeSTATUS_GeneralError) = General Error
; 								- 3 ($_ChromeSTATUS_InvalidDataType) = Invalid Data Type
; 								- 4 ($_ChromeSTATUS_InvalidObjectType) = Invalid Object Type
; 								- 6 ($_ChromeSTATUS_LoadWaitTimeout) = Load Wait Timeout
; 								- 8 ($_ChromeSTATUS_AccessIsDenChromed) = Access Is DenChromed
; 								- 9 ($_ChromeSTATUS_ClChromentDisconnected) = ClChroment Disconnected
; 					@extended	- Contains invalid parameter number
; Author ........: Dale Hohm
; Remarks .......:  AutoIt3 V3.2 or higher, flags for Tabs require Chrome7 or higher
; 					Additional information on the navigate2 method here: http://msdn.microsoft.com/en-us/library/aa752134.aspx
;
; Flags:
;    navOpenInNewWindow = 0x1,
;    navNoHistory = 0x2,
;    navNoReadFromCache = 0x4,
;    navNoWriteToCache = 0x8,
;    navAllowAutosearch = 0x10,
;    navBrowserBar = 0x20,
;    navHyperlink = 0x40,
;    navEnforceRestricted = 0x80,
;    navNewWindowsManaged = 0x0100,
;    navUntrustedForDownload = 0x0200,
;    navTrustedForActiveX = 0x0400,
;    navOpenInNewTab = 0x0800,
;    navOpenInBackgroundTab = 0x1000,
;    navKeepWordWheelText = 0x2000
;
; Additional documentation on the flags can be found here:
;    http://msdn.microsoft.com/en-us/library/aa768360.aspx
; ===============================================================================================================================
Func __ChromeNavigate(ByRef $oObject, $sUrl, $iWait = 1, $iFags = 0, $sTarget = "", $sPostdata = "", $sHeaders = "")
	__ChromeConsoleWriteError("Warning", "__ChromeNavigate", "Unsupported function called. Not fully tested.")
	If Not IsObj($oObject) Then
		__ChromeConsoleWriteError("Error", "__ChromeNavigate", "$_ChromeSTATUS_InvalidDataType")
		Return SetError($_ChromeSTATUS_InvalidDataType, 1, 0)
	EndIf
	;
	If Not __ChromeIsObjType($oObject, "documentContainer") Then
		__ChromeConsoleWriteError("Error", "__ChromeNavigate", "$_ChromeSTATUS_InvalidObjectType")
		Return SetError($_ChromeSTATUS_InvalidObjectType, 1, 0)
	EndIf
	;
	$oObject.navigate($sUrl, $iFags, $sTarget, $sPostdata, $sHeaders)
	If $iWait Then
		_ChromeLoadWait($oObject)
		Return SetError(@error, 0, $oObject)
	EndIf
	Return SetError($_ChromeSTATUS_Success, 0, $oObject)
EndFunc   ;==>__ChromeNavigate

#cs
	#include <Chrome.au3>
	; Simulates the submission of the form from the page:
	;
	;    http://www.autoitscript.com/forum/index.php?act=Search
	;
	; searches for the string safearray and returns the results as posts

	$sFormAction = "http://www.autoitscript.com/forum/index.php?act=Search&CODE=01"
	$sHeader = "Content-Type: application/x-www-form-urlencoded"

	$sDataToPost = "keywords=safearray&namesearch=&forums%5B%5D=all&searchsubs=1&prune=0&prune_type=newer&sort_key=last_post&sort_order=desc&search_in=posts&result_type=posts"
	$oDataToPostBstr = __ChromeStringToBstr($sDataToPost) ; convert string to BSTR
	ConsoleWrite(__ChromeBstrToString($oDataToPostBstr) & @CRLF) ; prove we can convert it back to a string

	$oChrome = _ChromeCreate()
	$oChrome.Navigate( $sFormAction, Default, Default, $oDataToPostBstr, $sHeader)
	; or
	;__ChromeNavigate($oChrome, $sFormAction, 1, 0, "", $oDataToPostBstr, $sHeader)
#ce

Func __ChromeStringToBstr($sString, $sCharSet = "us-ascii")
	Local Const $iTypeBinary = 1, $iTypeText = 2

	Local $oStream = ObjCreate("ADODB.Stream")

	$oStream.type = $iTypeText
	$oStream.CharSet = $sCharSet
	$oStream.Open
	$oStream.WriteText($sString)
	$oStream.Position = 0

	$oStream.type = $iTypeBinary
	$oStream.Position = 0

	Return $oStream.Read()
EndFunc   ;==>__ChromeStringToBstr

Func __ChromeBstrToString($oBstr, $sCharSet = "us-ascii")
	Local Const $iTypeBinary = 1, $iTypeText = 2

	Local $oStream = ObjCreate("ADODB.Stream")

	$oStream.type = $iTypeBinary
	$oStream.Open
	$oStream.Write($oBstr)
	$oStream.Position = 0

	$oStream.type = $iTypeText
	$oStream.CharSet = $sCharSet
	$oStream.Position = 0

	Return $oStream.ReadText()
EndFunc   ;==>__ChromeBstrToString

; #INTERNAL_USE_ONLY# ===========================================================================================================
; Name...........: __ChromeCreateNewChrome
; Description ...: Create a Webbrowser in a seperate process
; Parameters ....: None
; Return values .: On Success	- Returns a Webbrowser object reference
;                  On Failure	- Returns 0 and sets @error
; 					@error		- 0 ($_ChromeSTATUS_Success) = No Error
; 								- 1 ($_ChromeSTATUS_GeneralError) = General Error
; Author ........: Dale Hohm
; ModifChromed ......: jpm
; Remarks .......: http://msdn2.microsoft.com/en-us/library/ms536471(vs.85).aspx
; ===============================================================================================================================
Func __ChromeCreateNewChrome($sTitle, $sHead = "", $sBody = "")
	Local $sTemp = __ChromeTempFile("", "~Chrome~", ".htm")
	If @error Then
		__ChromeConsoleWriteError("Error", "_ChromeCreateHTA", "", "Error creating temporary file in @TempDir or @ScriptDir")
		Return SetError($_ChromeSTATUS_GeneralError, 1, 0)
	EndIf

	Local $sHTML = ''
	$sHTML &= '<!DOCTYPE html>' & @CR
	$sHTML &= '<html>' & @CR
	$sHTML &= '<head>' & @CR
	$sHTML &= '<meta content="text/html; charset=UTF-8" http-equiv="content-type">' & @CR
	$sHTML &= '<title>' & $sTemp & '</title>' & @CR & $sHead & @CR
	$sHTML &= '</head>' & @CR
	$sHTML &= '<body>' & @CR & $sBody & @CR
	$sHTML &= '</body>' & @CR
	$sHTML &= '</html>'

	Local $hFile = FileOpen($sTemp, $FO_OVERWRITE)
	FileWrite($hFile, $sHTML)
	FileClose($hFile)
	If @error Then
		__ChromeConsoleWriteError("Error", "_ChromeCreateNewChrome", "", "Error creating temporary file in @TempDir or @ScriptDir")
		Return SetError($_ChromeSTATUS_GeneralError, 2, 0)
	EndIf
	Run(@ProgramFilesDir & "\Chrome\Chromexplore.exe " & $sTemp)

	Local $iPID
	If WinWait($sTemp, "", 60) Then
		$iPID = WinGetProcess($sTemp)
	Else
		__ChromeConsoleWriteError("Error", "_ChromeCreateNewChrome", "", "Timeout waiting for new Chrome window creation")
		Return SetError($_ChromeSTATUS_GeneralError, 3, 0)
	EndIf

	If Not FileDelete($sTemp) Then
		__ChromeConsoleWriteError("Warning", "_ChromeCreateNewChrome", "", "Could not delete temporary file " & FileGetLongName($sTemp))
	EndIf

	Local $oObject = _ChromeAttach($sTemp)
	_ChromeLoadWait($oObject)
	_ChromePropertySet($oObject, "title", $sTitle)

	Return SetError($_ChromeSTATUS_Success, $iPID, $oObject)
EndFunc   ;==>__ChromeCreateNewChrome

; #INTERNAL_USE_ONLY# ===========================================================================================================
; Name...........: __ChromeTempFile
; Description ...: Generate a name for a temporary file. The file is guaranteed not to already exist.
; Parameters ....: $sDirectoryName    optional  Name of directory for filename, defaults to @TempDir
;                  $sFilePrefix       optional  File prefixname, defaults to "~"
;                  $sFileExtension    optional  File extenstion, defaults to ".tmp"
;                  $iRandomLength     optional  Number of characters to use to generate a unique name, defaults to 7
; Return values .: Filename of a temporary file which does not exist.
; Author ........: Dale (Klaatu) Thompson
; ModifChromed.......: Hans Harder - Added Optional parameters
;
; Adapted from excellent _TempFile() in File.au3 for Chrome.au3 by Dale Hohm
; ===============================================================================================================================
Func __ChromeTempFile($sDirectoryName = @TempDir, $sFilePrefix = "~", $sFileExtension = ".tmp", $iRandomLength = 7)
	Local $sTempName, $iTmp = 0
	; Check parameters
	If Not FileExists($sDirectoryName) Then $sDirectoryName = @TempDir ; First reset to default temp dir
	If Not FileExists($sDirectoryName) Then $sDirectoryName = @ScriptDir ; Still wrong then set to Scriptdir
	; add trailing \ for directory name
	If StringRight($sDirectoryName, 1) <> "\" Then $sDirectoryName = $sDirectoryName & "\"
	;
	Do
		$sTempName = ""
		While StringLen($sTempName) < $iRandomLength
			$sTempName = $sTempName & Chr(Random(97, 122, 1))
		WEnd
		$sTempName = $sDirectoryName & $sFilePrefix & $sTempName & $sFileExtension
		$iTmp += 1
		If $iTmp > 200 Then ; If we fail over 200 times, there is something wrong
			Return SetError($_ChromeSTATUS_GeneralError, 1, 0)
		EndIf
	Until Not FileExists($sTempName)

	Return $sTempName
EndFunc   ;==>__ChromeTempFile

#EndRegion ProtoType Functions

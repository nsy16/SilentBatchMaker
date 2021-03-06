#RequireAdmin

; Silent Batch Maker - by Nate
; License GNU GPLv3.0


#cs
    - TODO -
	Add last opened location
	Add setting to open default directory - eg downloads


	Add a open folder option to rows, to open folder directly
	Add scroll bar
	Add maximum depth option, should be 2 by default, means all exe + msi files under that folder appear as array in combobox
	Add global switch changer for all, to change switch for all options, except locked
	clicking on test has to save to a temporary file (why because automatic reboot)
	Multiple regchanges need a dropdown or choice if found
    Load function should reorder rows
	Deleting rows after loading should not erase the boxes

	custom flag dir
	more output options - eg vbs, ps1
	installed directory scan and point at .exe flag

	improve loading speed
#ce

#include <StaticConstants.au3>
#include <GUIConstantsEx.au3>
#include <WindowsConstants.au3>
#include <GuiButton.au3>
#include <ComboConstants.au3>

#include <File.au3>
#include <StringConstants.au3>

#include <MsgBoxConstants.au3>
#include <GUIScrollbars_Size.au3> ; GUIScrollBars_Size by Melba23

Global $aRow[1][1]
Global $aRegistryTest[1][1]
Global $aRegistryScan[1][1]
Global $sTestRow

; Register the function MyAdLibFunc() to be called every 250ms (default).
AdlibRegister("MyAdLibFunc")

; Main GUI header
$sTitle = "Silent Batch Maker 1.3.5 @ 28/02/2018"
$hGUI = GUICreate($sTitle , 850, 550, -1, -1)
$idBrowse = GUICtrlCreateButton("Browse", 9, 15, 100, 30)
$idSave = GUICtrlCreateButton("Save", 622, 15, 100, 30)
$idLoad = GUICtrlCreateButton("Load", 728, 15, 100, 30)
GUICtrlSetState(-1,$GUI_DISABLE)
$lbFolder = GUICtrlCreateLabel("Browse to the install folder", 120, 15, 487, 28, $SS_CENTERIMAGE)

Global $OPENDIR = @TempDir
Global $pFname[4]
$pFname[0] = "install.cmd"
$pFname[1] = "install-batch-regcheck-full.cmd"
$pFname[2] = "install-batch-regcheck-multi-full.cmd"
; search for *.vbs file - may require a manuall check


;GUICtrlCreateRadio("Lock", 694, 60, 42, 21)
;GUICtrlCreateRadio("Unlock", 736, 60, 53, 21)

GUISetState(@SW_SHOW)

; Register the handler
GUIRegisterMsg($WM_VSCROLL, "_Scrollbars_WM_VSCROLL")
; Initiate and hide the vertical scrollbar
_GUIScrollBars_Init($hGUI)
_GUIScrollBars_ShowScrollBar($hGUI, $SB_VERT, False)
_GUIScrollBars_ShowScrollBar($hGUI, $SB_HORZ, False)

While True
    $nMsg = GUIGetMsg()
    Select
        Case $nMsg = $GUI_EVENT_CLOSE
            Exit
        Case $nMsg = $idBrowse
            Local $sDir = Browse()
            ;If $OPENDIR <> $sDir Then
                GUICtrlSetData($lbFolder, $sDir)
				$OPENDIR = $sDir ; update global value so next open will go to last opened folder
                ;$aFiles = _FileListToArrayRec($sDir, '*.msi;*.mst;*.exe', $FLTAR_FILES, $FLTAR_RECUR, $FLTAR_SORT, $FLTAR_RELPATH)
                $aFiles = _FileListToArrayRec($sDir, '*.exe;*.msi;*.msp', $FLTAR_FILES, $FLTAR_RECUR, $FLTAR_SORT, $FLTAR_RELPATH)
                If Not @error Then
                    _CreateControls()
					;For $i = 0 To UBound($pFname) - 1
					;	If FileExists($sDir & '\' & $pFname[$i]) Then GUICtrlSetState($idLoad,$GUI_ENABLE)
					;Next
					GUICtrlSetState($idLoad,$GUI_ENABLE)
                EndIf
            ;EndIf
        Case $nMsg = $idSave
            If IsDeclared("aFiles") Then
                Save($sDir)
				;MsgBox(0,"Success", "Your batch files have been generated at " & $sDir)
				;Opt("TrayIconHide", 0)
				TrayTip(@ScriptName, "Your batch files have been generated at " & $sDir, 3, 16)
				;Opt("TrayIconHide", 1)
            Else
                MsgBox(0, "Error", "Nothing to save!")
            EndIf
		Case $nMsg = $idLoad
			;For $i = 0 To UBound($pFname) - 1
			;	If FileExists($sDir & '\' & $pFname[$i]) Then $profile = $sDir & '\' & $pFname[$i]
			;Next
			;If IsDeclared("profile") Then
			;	Load($profile)
			;Else
            ;    MsgBox(0, "Error", "No profile was found!")
            ;EndIf
			Local $profile = FileOpenDialog("Choose a script to load", $OPENDIR, "install scripts (*.cmd;*.vbs)")
			If @error Then
				;MsgBox(4096,"","Select cancelled.")
			Else
				Load($profile)
			EndIf




        Case Else
            For $i = 1 To UBound($aRow) - 1
                If ($nMsg = $aRow[$i][1]) Then ; Look for Up button pushed
                    ConsoleWrite("Up Button in Row " & $i & " was pressed!" & @CRLF)
                    If $i <> 1 Then
                        _ArraySwap2($aFiles[$i], $aFiles[$i - 1])
                        _CreateControls()
                    EndIf
                    ExitLoop ; Exit because we found the button, no need to keep searching
                ElseIf ($nMsg = $aRow[$i][2]) Then ; Look for Down button pushed
                    ConsoleWrite("Down Button in Row " & $i & " was pressed!" & @CRLF)
                    If $i <> $aFiles[0] Then
                        _ArraySwap2($aFiles[$i], $aFiles[$i + 1])
                        _CreateControls()
                    EndIf
                    ExitLoop ; Exit because we found the button, no need to keep searching
                ElseIf ($nMsg = $aRow[$i][6]) Then ; Look for Test button pushed
                    ConsoleWrite("Test Button in Row " & $i & " was pressed!" & @CRLF)

					; update global variable, adlib function will return registry detected values to this row
					$sTestRow = $i

					; only test the FIRST time this row is pushed!
					; however if rows are deleted
					RegShot($aRegistryTest)
					;_ArrayDisplay($aRegistryTest, "$aRegistryTest")

                    $path = GUICtrlRead($aRow[$i][3])
                    $switch = GUICtrlRead($aRow[$i][5])
                    ConsoleWrite("Test(" & $sDir & ", " & $path & ", " & $switch & ")" & @CRLF)
                    Test($sDir, $path, $switch)
                    ExitLoop ; Exit because we found the button, no need to keep searching
                ElseIf ($nMsg = $aRow[$i][7]) Then ; Look for Delete button pushed
                    ConsoleWrite("Delete Button in Row " & $i & " was pressed!" & @CRLF)
                    ; blank line and remove from array, shifting values up
                    $aFiles[$i] = ""
                    _ArrayRemoveBlanks($aFiles)
                    $aFiles[0] = $aFiles[0] - 1 ; decrease count in [0]
                    ; rebuild the GUI from array
                    _CreateControls()

                    ExitLoop ; Exit because we found the button, no need to keep searching
                EndIf
            Next
    EndSelect
WEnd

Func _CreateControls()

    GUISetState(@SW_LOCK)
    DeleteAllGUIobjects()
    CreateRows($aFiles)
    GUISetState(@SW_UNLOCK)

EndFunc


Func Browse()
    ; Display an open dialog to select a file.
    Local Const $sMessage = "Select a folder"
    Local $sFileSelectFolder = FileSelectFolder($sMessage, $OPENDIR)
    If @error Then
		Return
    Else
        Return $sFileSelectFolder
    EndIf
EndFunc   ;==>Browse


Func Save($saveDir)    ; take all the exe files and switches and generate batch files in the same order
    Local $msg
	$parentDir = StringRegExpReplace($sDir, '.*\\', '')

	; Determine which batch files to generate
	For $i = 1 To UBound($aRow) - 1
		GUICtrlRead($aRow[$i][8])  ;$regK
		GUICtrlRead($aRow[$i][9])  ;$regDN
		GUICtrlRead($aRow[$i][10]) ;$regDV
    Next


	; Basic - no check
	$msg = ':: Install ' & $sDir & @CRLF
	;$msg &= ':: Generated by ' & $sTitle & @CRLF
	$msg &= '' & @CRLF
    For $i = 1 To UBound($aRow) - 1
        $path   = GUICtrlRead($aRow[$i][3])
        $switch = GUICtrlRead($aRow[$i][5])
		$msg &= '"%~dp0' & $path & '" ' & $switch & @CRLF
    Next

    ;MsgBox(0, 'Debug', $msg)

;~     Local $File = FileOpen($saveDir & '\install-batch-basic.cmd', 2) ;Use UTF8, overwrite existing data, create directory
;~     FileWrite($File, $msg)
;~     FileClose($File)


	; Basic - generate flag
	; Check registry - single file only
    $msg = '@echo off&cls' & @CRLF
    $msg &= 'for %%a in ("%~dp0\.") do set _parentdir=%%~nxa' & @CRLF
    $msg &= 'title %_parentdir%' & @CRLF
    $msg &= '' & @CRLF
	$msg &= ':: Install ' & $parentDir & @CRLF
	;$msg &= ':: Generated by ' & $sTitle & @CRLF
	$msg &= '' & @CRLF
    $msg &= 'set _flag=' & $parentDir & '.flag' & @CRLF
	$msg &= 'set _flagdir=c:\Flags' & @CRLF
	$msg &= 'if exist "%_flagdir%\%_flag%" goto :end' & @CRLF
	$msg &= '' & @CRLF
	$msg &= ':start' & @CRLF
    For $i = 1 To UBound($aRow) - 1
        $path   = GUICtrlRead($aRow[$i][3])
        $switch = GUICtrlRead($aRow[$i][5])
		$msg &= 'start /wait "" "%~dp0' & $path & '" ' & $switch & @CRLF
    Next
	$msg &= '' & @CRLF
	$msg &= ':flag' & @CRLF
	$msg &= 'md "%_flagdir%" 2>nul' & @CRLF
	$msg &= 'echo Installed %~nx0 >> "%_flagdir%\%_flag%"' & @CRLF
	$msg &= 'hostname >> "%_flagdir%\%_flag%"' & @CRLF
	$msg &= 'echo %date% - %time% >> "%_flagdir%\%_flag%"' & @CRLF
	$msg &= 'echo --- >> "%_flagdir%\%_flag%"' & @CRLF
	$msg &= '' & @CRLF
	$msg &= ':end' & @CRLF

    ;MsgBox(0, 'Debug', $msg)

;~     Local $File = FileOpen($saveDir & '\install-batch-flag.cmd', 2) ;Use UTF8, overwrite existing data, create directory
;~     FileWrite($File, $msg)
;~     FileClose($File)


	If UBound($aRow) - 1 < 2 Then ; only one entry - note need to check if multiple rows have empty reg completed entries


;~ 		; Check registry - single file only
;~ 		$msg = '@echo off&cls' & @CRLF
;~ 		$msg &= 'for %%a in ("%~dp0\.") do set _parentdir=%%~nxa' & @CRLF
;~ 		$msg &= 'title %_parentdir%' & @CRLF
;~ 		$msg &= '' & @CRLF
;~ 		$msg &= ':: Install ' & $parentDir & @CRLF
;~ 		$msg &= ':: Generated by ' & $sTitle & @CRLF
;~ 		$msg &= '' & @CRLF
;~ 		For $i = 1 To UBound($aRow) - 1
;~ 			$fPath = GUICtrlRead($aRow[$i][3])
;~ 			$swtch = GUICtrlRead($aRow[$i][5])
;~ 			$regKy = GUICtrlRead($aRow[$i][8])
;~ 			$regDN = GUICtrlRead($aRow[$i][9])
;~ 			$regDV = GUICtrlRead($aRow[$i][10])

;~ 			;ConsoleWrite("$regDN: " & $regDN & @CRLF)
;~ 			;ConsoleWrite("$regDV: " & $regDV & @CRLF)

;~ 			$msg &= 'set _fPath=%~dp0' & $fPath & @CRLF
;~ 			$msg &= 'set _swtch=' & $swtch & @CRLF
;~ 			$msg &= 'set _regKy=' & $regKy & @CRLF

;~ 			Select
;~ 				Case $regDV <> ""
;~ 					$msg &= 'set _regDN=' & $regDN & @CRLF
;~ 					$msg &= 'set _regDV=' & $regDV & @CRLF
;~ 					$msg &= '' & @CRLF
;~ 					$msg &= ':: query registry for DisplayVersion 64bit and 32bit uninstall locations' & @CRLF
;~ 					$msg &= 'for /f "tokens=3" %%a in (''reg query "HKLM\SOFTWARE\Microsoft\Windows\CurrentVersion\Uninstall\%_regKy%" /v "DisplayVersion" 2^>nul'') do if "%%a" geq "%_regDV%" goto :end' & @CRLF
;~ 					$msg &= 'for /f "tokens=3" %%a in (''reg query "HKLM\SOFTWARE\WOW6432Node\Microsoft\Windows\CurrentVersion\Uninstall\%_regKy%" /v "DisplayVersion" 2^>nul'') do if "%%a" geq "%_regDV%" goto :end' & @CRLF
;~ 					$msg &= '' & @CRLF
;~ 				Case $regDN <> ""
;~ 					$msg &= 'set _regDN=' & $regDN & @CRLF
;~ 					$msg &= '' & @CRLF
;~ 					$msg &= ':: query registry for DisplayName 64bit and 32bit uninstall locations' & @CRLF
;~ 					$msg &= 'for /f "usebackq tokens=1-2*" %%a in (`reg query "HKLM\SOFTWARE\Microsoft\Windows\CurrentVersion\Uninstall\%_regKy%" /v "DisplayName" 2^>nul`) do if "%%c" equ "%_regDN%" goto :end' & @CRLF
;~ 					$msg &= 'for /f "usebackq tokens=1-2*" %%a in (`reg query "HKLM\SOFTWARE\WOW6432Node\Microsoft\Windows\CurrentVersion\Uninstall\%_regKy%" /v "DisplayName" 2^>nul`) do if "%%c" equ "%_regDN%" goto :end' & @CRLF
;~ 					$msg &= '' & @CRLF
;~ 			EndSelect

;~ 			$msg &= ':: install software' & @CRLF
;~ 			$msg &= 'echo Installing %_regDN%' & @CRLF
;~ 			$msg &= 'start /wait "" "%_fPath%" %_swtch%' & @CRLF
;~ 			$msg &= '' & @CRLF
;~ 			$msg &= ':end' & @CRLF
;~ 		Next

;~ 		;MsgBox(0, 'Debug', $msg)

;~ 		Local $File = FileOpen($saveDir & '\install-batch-regcheck.cmd', 2) ;Use UTF8, overwrite existing data, create directory
;~ 		FileWrite($File, $msg)
;~ 		FileClose($File)



		; Check registry - single file only - full details
		$msg = '@echo off&cls' & @CRLF
		$msg &= 'for %%a in ("%~dp0\.") do set _parentdir=%%~nxa' & @CRLF
		$msg &= 'title %_parentdir%' & @CRLF
		$msg &= '' & @CRLF
		$msg &= ':: Install ' & $parentDir & @CRLF
		;$msg &= ':: Generated by ' & $sTitle & @CRLF
		$msg &= '' & @CRLF
		For $i = 1 To UBound($aRow) - 1
			$fPath = GUICtrlRead($aRow[$i][3])
			$swtch = GUICtrlRead($aRow[$i][5])
			$regKy = GUICtrlRead($aRow[$i][8])
			$regDN = GUICtrlRead($aRow[$i][9])
			$regDV = GUICtrlRead($aRow[$i][10])

			Select
				Case $regDN <> "" And $regDV = ""
					$msg &= 'set _fPath=%~dp0' & $fPath & @CRLF
					$msg &= 'set _swtch=' & $swtch & @CRLF
					$msg &= 'set _regKy=' & $regKy & @CRLF
					$msg &= 'set _regDN=' & $regDN & @CRLF
					$msg &= '' & @CRLF
					$msg &= ':: query registry for DisplayName 64bit and 32bit uninstall locations' & @CRLF
					$msg &= 'for /f "usebackq tokens=1-2*" %%a in (`reg query "HKLM\SOFTWARE\Microsoft\Windows\CurrentVersion\Uninstall\%_regKy%" /v "DisplayName" 2^>nul`) do if "%%c" equ "%_regDN%" goto :end' & @CRLF
					$msg &= 'for /f "usebackq tokens=1-2*" %%a in (`reg query "HKLM\SOFTWARE\WOW6432Node\Microsoft\Windows\CurrentVersion\Uninstall\%_regKy%" /v "DisplayName" 2^>nul`) do if "%%c" equ "%_regDN%" goto :end' & @CRLF
					$msg &= '' & @CRLF
					$msg &= ':: install software' & @CRLF
					$msg &= 'echo Installing %_regDN%' & @CRLF

;~ 				Case $regDN = "" And $regDV <> "" ; Display Name will always be set - update this section if a case is found
;~ 					$msg &= 'set _fPath=%~dp0' & $fPath & @CRLF
;~ 					$msg &= 'set _swtch=' & $swtch & @CRLF
;~ 					$msg &= 'set _regKy=' & $regKy & @CRLF
;~ 					$msg &= 'set _regDN=' & $regDN & @CRLF
;~ 					$msg &= 'set _regDV=' & $regDV & @CRLF
;~ 					$msg &= '' & @CRLF
;~ 					$msg &= ':: query registry for DisplayVersion 64bit and 32bit uninstall locations' & @CRLF
;~ 					$msg &= 'for /f "tokens=3" %%a in (''reg query "HKLM\SOFTWARE\Microsoft\Windows\CurrentVersion\Uninstall\%_regKy%" /v "DisplayVersion" 2^>nul'') do if "%%a" geq "%_regDV%" goto :end' & @CRLF
;~ 					$msg &= 'for /f "tokens=3" %%a in (''reg query "HKLM\SOFTWARE\WOW6432Node\Microsoft\Windows\CurrentVersion\Uninstall\%_regKy%" /v "DisplayVersion" 2^>nul'') do if "%%a" geq "%_regDV%" goto :end' & @CRLF
;~ 					$msg &= '' & @CRLF
;~ 					$msg &= ':: install software' & @CRLF
;~ 					$msg &= 'echo Installing ' & $parentDir & ' %_regDV%' & @CRLF

				Case $regDN <> "" And $regDV <> ""
					$msg &= 'set _fPath=%~dp0' & $fPath & @CRLF
					$msg &= 'set _swtch=' & $swtch & @CRLF
					$msg &= 'set _regKy=' & $regKy & @CRLF
					$msg &= 'set _regDN=' & $regDN & @CRLF
					$msg &= 'set _regDV=' & $regDV & @CRLF
					$msg &= '' & @CRLF
					$msg &= ':: query registry for DisplayVersion 64bit and 32bit uninstall locations' & @CRLF
					$msg &= 'for /f "tokens=3" %%a in (''reg query "HKLM\SOFTWARE\Microsoft\Windows\CurrentVersion\Uninstall\%_regKy%" /v "DisplayVersion" 2^>nul'') do if "%%a" geq "%_regDV%" goto :end' & @CRLF
					$msg &= 'for /f "tokens=3" %%a in (''reg query "HKLM\SOFTWARE\WOW6432Node\Microsoft\Windows\CurrentVersion\Uninstall\%_regKy%" /v "DisplayVersion" 2^>nul'') do if "%%a" geq "%_regDV%" goto :end' & @CRLF
					$msg &= '' & @CRLF
					$msg &= ':: install software' & @CRLF
					$msg &= 'echo Installing %_regDN% %_regDV%' & @CRLF

			EndSelect

		Next

		$msg &= 'start /wait "" "%_fPath%" %_swtch%' & @CRLF
		$msg &= '' & @CRLF
		$msg &= ':end' & @CRLF

		;MsgBox(0, 'Debug', $msg)

		;Local $File = FileOpen($saveDir & '\install-single-regcheck.cmd', 2) ;Use UTF8, overwrite existing data, create directory
		Local $File = FileOpen($saveDir & '\install.cmd', 2) ;Use UTF8, overwrite existing data, create directory
		FileWrite($File, $msg)
		FileClose($File)


	Else ; multiple entries


		; Check registry - multiple files - array - full details
		Local $regKyCount = 0
		Local $regDNCount = 0
		Local $regDVCount = 0
		Local $loop = 0

		$msg = '@echo off&cls' & @CRLF
		$msg &= 'for %%a in ("%~dp0\.") do set _parentdir=%%~nxa' & @CRLF
		$msg &= 'title %_parentdir%' & @CRLF
		$msg &= '' & @CRLF
		$msg &= ':: Install ' & $parentDir & @CRLF
		;$msg &= ':: Generated by ' & $sTitle & @CRLF
		$msg &= '' & @CRLF

		For $i = 1 To UBound($aRow) - 1
			$fPath = GUICtrlRead($aRow[$i][3])
			$swtch = GUICtrlRead($aRow[$i][5])
			$regKy = GUICtrlRead($aRow[$i][8])
			$regDN = GUICtrlRead($aRow[$i][9])
			$regDV = GUICtrlRead($aRow[$i][10])
			Select
				Case $regDN <> "" And $regDV = ""
					$loop += 1
					$msg &= 'set _fPath_['&$loop&']=%~dp0' & $fPath & @CRLF
					$msg &= 'set _swtch_['&$loop&']=' & $swtch & @CRLF
					$msg &= 'set _regKy_['&$loop&']=' & $regKy & @CRLF
					$msg &= 'set _regDN_['&$loop&']=' & $regDN & @CRLF
					$msg &= '' & @CRLF
				Case $regDN = "" And $regDV <> ""
					$loop += 1
					$msg &= 'set _fPath_['&$loop&']=%~dp0' & $fPath & @CRLF
					$msg &= 'set _swtch_['&$loop&']=' & $swtch & @CRLF
					$msg &= 'set _regKy_['&$loop&']=' & $regKy & @CRLF
					$msg &= 'set _regDN_['&$loop&']=' & $regDN & @CRLF
					$msg &= 'set _regDV_['&$loop&']=' & $regDV & @CRLF
					$msg &= '' & @CRLF
				Case $regDN <> "" And $regDV <> ""
					$loop += 1
					$msg &= 'set _fPath_['&$loop&']=%~dp0' & $fPath & @CRLF
					$msg &= 'set _swtch_['&$loop&']=' & $swtch & @CRLF
					$msg &= 'set _regKy_['&$loop&']=' & $regKy & @CRLF
					$msg &= 'set _regDN_['&$loop&']=' & $regDN & @CRLF
					$msg &= 'set _regDV_['&$loop&']=' & $regDV & @CRLF
					$msg &= '' & @CRLF
			EndSelect
			If $regKy <> "" Then $regKyCount += 1

			; build sub based on what data is available - DisplayVersion is better than DisplayName
			If $regDN <> "" And $regDV <> "" Then $regDVCount += 1
			If $regDN = "" And $regDV <> ""  Then $regDVCount += 1
			If $regDN <> "" And $regDV = ""  Then $regDNCount += 1

		Next
		If $regKyCount = 0 Then Return; Error Return

		$msg &= '' & @CRLF
		$msg &= ':: loop through installation packages' & @CRLF
		$msg &= 'set _count=' & $loop & @CRLF
		$msg &= 'for /l %%g in (1,1,%_count%) do (call :sub_install "%%g")' & @CRLF
		$msg &= '' & @CRLF
		$msg &= ':: script finished' & @CRLF
		;$msg &= 'pause' & @CRLF
		$msg &= 'goto :eof' & @CRLF
		$msg &= '' & @CRLF
		$msg &= '' & @CRLF
		$msg &= ':sub_install <index>' & @CRLF
		$msg &= '  set _index=%~1' & @CRLF
		$msg &= '  call set _fPath=%%_fPath_[%_index%]%%%' & @CRLF
		$msg &= '  if not defined _fPath goto :eof' & @CRLF
		$msg &= '  call set _swtch=%%_swtch_[%_index%]%%%' & @CRLF
		$msg &= '  call set _regKy=%%_regKy_[%_index%]%%%' & @CRLF
		$msg &= '  call set _regDN=%%_regDN_[%_index%]%%%' & @CRLF
		If $regDVCount = 0 Then
			$msg &= '' & @CRLF
			$msg &= '  :: query registry for DisplayName 64bit and 32bit uninstall locations' & @CRLF
			$msg &= '  for /f "usebackq tokens=1-2*" %%a in (`reg query "HKLM\SOFTWARE\Microsoft\Windows\CurrentVersion\Uninstall\%_regKy%" /v "DisplayName" 2^>nul`) do if "%%c" equ "%_regDN%" goto :eof' & @CRLF
			$msg &= '  for /f "usebackq tokens=1-2*" %%a in (`reg query "HKLM\SOFTWARE\WOW6432Node\Microsoft\Windows\CurrentVersion\Uninstall\%_regKy%" /v "DisplayName" 2^>nul`) do if "%%c" equ "%_regDN%" goto :eof' & @CRLF
			$msg &= '' & @CRLF
			$msg &= '  :: install software with switch' & @CRLF
			$msg &= '  echo Installing %_regDN%' & @CRLF
		EndIf
		If $regDVCount > 0 And $regDNCount = 0 Then
			$msg &= '  call set _regDV=%%_regDV_[%_index%]%%%' & @CRLF
			$msg &= '' & @CRLF
			$msg &= '  :: query registry for DisplayVersion 64bit and 32bit uninstall locations' & @CRLF
			$msg &= '  for /f "tokens=3" %%a in (''reg query "HKLM\SOFTWARE\Microsoft\Windows\CurrentVersion\Uninstall\%_regKy%" /v "DisplayVersion" 2^>nul'') do if "%%a" geq "%_regDV%" goto :eof' & @CRLF
			$msg &= '  for /f "tokens=3" %%a in (''reg query "HKLM\SOFTWARE\WOW6432Node\Microsoft\Windows\CurrentVersion\Uninstall\%_regKy%" /v "DisplayVersion" 2^>nul'') do if "%%a" geq "%_regDV%" goto :eof' & @CRLF
			$msg &= '' & @CRLF
			$msg &= '  :: install software with switch' & @CRLF
			$msg &= '  echo Installing %_regDN% %_regDV%' & @CRLF
		ElseIf $regDVCount > 0 And $regDNCount > 0 Then
			$msg &= '  call set _regDV=%%_regDV_[%_index%]%%%' & @CRLF
			$msg &= '' & @CRLF
			$msg &= '  :: multi query registry for 64bit and 32bit uninstall locations' & @CRLF
			$msg &= '  :: display version' & @CRLF
			$msg &= '  if defined _regDV for /f "tokens=3" %%a in (''reg query "HKLM\SOFTWARE\Microsoft\Windows\CurrentVersion\Uninstall\%_regKy%" /v "DisplayVersion" 2^>nul'') do if "%%a" geq "%_regDV%" goto :eof' & @CRLF
			$msg &= '  if defined _regDV for /f "tokens=3" %%a in (''reg query "HKLM\SOFTWARE\WOW6432Node\Microsoft\Windows\CurrentVersion\Uninstall\%_regKy%" /v "DisplayVersion" 2^>nul'') do if "%%a" geq "%_regDV%" goto :eof' & @CRLF
			$msg &= '  :: display name' & @CRLF
			$msg &= '  if not defined _regDV for /f "usebackq tokens=1-2*" %%a in (`reg query "HKLM\SOFTWARE\Microsoft\Windows\CurrentVersion\Uninstall\%_regKy%" /v "DisplayName" 2^>nul`) do if "%%c" equ "%_regDN%" goto :eof' & @CRLF
			$msg &= '  if not defined _regDV for /f "usebackq tokens=1-2*" %%a in (`reg query "HKLM\SOFTWARE\WOW6432Node\Microsoft\Windows\CurrentVersion\Uninstall\%_regKy%" /v "DisplayName" 2^>nul`) do if "%%c" equ "%_regDN%" goto :eof' & @CRLF
			$msg &= '' & @CRLF
			$msg &= '  :: install software with switch' & @CRLF
			$msg &= '  echo Installing %_regDN%' & @CRLF
		EndIf

		$msg &= '  start /wait "" "%_fPath%" %_swtch%' & @CRLF
		$msg &= '  goto :eof' & @CRLF
		$msg &= '' & @CRLF


		;MsgBox(0, 'Debug', $msg)

		Local $File = FileOpen($saveDir & '\' & $pFname[1], 2) ;Use UTF8, overwrite existing data, create directory
		FileWrite($File, $msg)
		FileClose($File)

	EndIf

EndFunc   ;==>Save


Func Load($profile)
	; reads a batch file for values and instates them into the boxes in SBM
	; assumes path has been browsed too
	; creates an array which is passed to subLoadValuesToGUI() function

	Dim $aLoadValues[1][5] = [["$fPath", "$swtch", "$regKy", "$regDN", "$regDV"]]
	;_ArrayDisplay($aLoadValues, "$aLoadValues")

	Local $count = 1, $loop = ""

	$readfile = FileOpen($profile, 0)
	If $readfile = -1 Then	; Check if file opened for reading OK
		MsgBox(0, @ScriptName & " Error", "Unable to open file: " & $profile)
		Exit
	EndIf
	While 1
		$line = FileReadLine($readfile)
		If @error = -1 Then
			ExitLoop
		EndIf

		;ConsoleWrite("$line - " & $line & @CRLF)
		;ConsoleWrite("regmatch - " & StringRegExp($line, "aInstall\(\d,0\)", $STR_REGEXPMATCH) & @CRLF)
		;If StringRegExp($line, "aInstall\(\d,0\)", $STR_REGEXPMATCH) = True Then ConsoleWrite("$line - " & $line & @CRLF)

		; read all values, in correct order
		If StringInStr($line, "set _fPath") or StringInStr($line, "set _exe") Then
			; increment the counter if on this first section again
			If $loop = "" Then ; Has not been past this section yet, set loop initial value
				$loop = "initialised"
				ConsoleWrite("First Loop " & $count& @CRLF)
			Else ; increase count by one
				$count += 1
				ConsoleWrite("Next Loop " & $count& @CRLF)
			EndIf

			; increase array rows if required
			If $count > UBound($aLoadValues, 1) - 1 Then ReDim $aLoadValues[UBound($aLoadValues) + 1][5]

			$a = SanitizeString($line)
			ConsoleWrite("_fPath: " & $a & @CRLF)
			$fPath = $a
			$aLoadValues[$count][0] = $fPath
		EndIf
		; vbs title splits path and file
		If StringRegExp($line, "aInstall\(\d,[0]\)", $STR_REGEXPMATCH) Then
			;ConsoleWrite($line & @CRLF)
			; increment the counter if on this first section again
			If $loop = "" Then ; Has not been past this section yet, set loop initial value
				$loop = "initialised"
				ConsoleWrite("First Loop " & $count& @CRLF)
			Else ; increase count by one
				$count += 1
				ConsoleWrite("Next Loop " & $count& @CRLF)
			EndIf

			; increase array rows if required
			If $count > UBound($aLoadValues, 1) - 1 Then ReDim $aLoadValues[UBound($aLoadValues) + 1][5]

			$a = SanitizeString($line)
			ConsoleWrite("_fPath: " & $a & @CRLF)
			$fPath = $a
			;$aLoadValues[$count][0] = $fPath
		EndIf
		If StringRegExp($line, "aInstall\(\d,[1]\)", $STR_REGEXPMATCH) Then
			$a = SanitizeString($line)
			If $fPath = "" Then
				$fPath = $a
			Else
				$fPath = $fPath & "\" & $a
			EndIf
			ConsoleWrite("_fPath: " & $fPath & @CRLF)
			$aLoadValues[$count][0] = $fPath
		EndIf

		;ConsoleWrite($line & @CRLF)
		;_ArrayDisplay($aLoadValues)

		If StringInStr($line, "set _swtch") Or StringInStr($line, "set _switch") Or StringRegExp($line, "aInstall\(\d,[2]\)", $STR_REGEXPMATCH) Then
			$a = SanitizeString($line)
			ConsoleWrite("_swtch: " & $a & @CRLF)
			$swtch = $a
			$aLoadValues[$count][1] = $swtch
		EndIf

		If StringInStr($line, "set _regKy") Or StringInStr($line, "set _regK") Or StringRegExp($line, "aInstall\(\d,[3]\)", $STR_REGEXPMATCH) Then
			$a = SanitizeString($line)
			ConsoleWrite("_regKy: " & $a & @CRLF)
			$regKy = $a
			$aLoadValues[$count][2] = $regKy
		EndIf

		If StringInStr($line, "set _regDN") Or StringRegExp($line, "aInstall\(\d,[4]\)", $STR_REGEXPMATCH) Then
			$a = SanitizeString($line)
			If $a <> "" Then
				ConsoleWrite("_regDN: " & $a & @CRLF)
				$regDN = $a
				$aLoadValues[$count][3] = $regDN
			EndIf
		EndIf

		If StringInStr($line, "set _regDV") Or StringRegExp($line, "aInstall\(\d,[5]\)", $STR_REGEXPMATCH) Then
			$a = SanitizeString($line)
			If $a <> "" Then
				ConsoleWrite("_regDV: " & $a & @CRLF)
				$regDV = $a
				$aLoadValues[$count][4] = $regDV
			EndIf

			; NB - could exit loop here if all values loaded - eg look for :: and exitloop

		EndIf
		$prevline = $line ; keep last line in buffer for OSArch comparison
	Wend

	FileClose($readfile)
;~ 	_ArrayDisplay($aLoadValues, "$aLoadValues")
	subLoadValuesToGUI($aLoadValues)

EndFunc   ;==>Load


Func SanitizeString($string)
	;$a = StringRegExpReplace($string, '(=.+)$', "")  ; Erase from first "=", to end
	$b = StringRegExpReplace($string, '(^.+?=)', "") ; Erase from beginning, to first "="
	$c = StringRegExpReplace($b, '(?i)%~dp0', "") ; Strip %~dp0 from string
	$d = StringRegExpReplace($c, '(?i)strScriptPath & "\" &', "") ; Strip strScriptPath (extra) from string
	$e = StringRegExpReplace($d, '(?i)strScriptPath', "") ; Strip strScriptPath (single) from string

	$f = StringStripWS($e, $STR_STRIPLEADING + $STR_STRIPTRAILING + $STR_STRIPSPACES) ; remove leading and trailing double spaces
	$g = StringRegExpReplace($f, '(?m)^"|"(?=[^"]*"$)|"$', ""); remove leading and trailing quote marks

	Local $cleanstring = $g
	ConsoleWrite("$sleanstring: " & $cleanstring & @CRLF)
	Return $cleanstring
EndFunc   ;==>SanitizeString


Func subLoadValuesToGUI($aLoadValues)
	; Reads array and enters values into matching boxes
	; Uses $fPath to check as this value should always be non empty

	;$lbOrder               = 0
    ;$idButtonUp            = 1
    ;$idButtonDown          = 2
    ;$idComboPath           = 3
    ;$idIcon                = 4
    ;$idComboSwitch         = 5
    ;$idButtonTest          = 6
    ;$idButtonDelete        = 7
    ;$idInputRegKey         = 8
    ;$idInputDisplayName    = 9
    ;$idInputDisplayVersion = 10

	For $i = 1 To UBound($aRow) - 1
		$comboPath = GUICtrlRead($aRow[$i][3]) ;$fPath
		For $j = 1 To UBound($aLoadValues, 1) - 1
			$arrayPath = $aLoadValues[$j][0] ;$fPath
			;ConsoleWrite($comboPath & ' = ' & $arrayPath & @CRLF)
			If $comboPath = $arrayPath Then
				GUICtrlSetData($aRow[$i][5],  $aLoadValues[$j][1], $aLoadValues[$j][1]) ; $swtch - set to value incase not in existing list
				GUICtrlSetData($aRow[$i][8],  $aLoadValues[$j][2], "") ; $regKy
				GUICtrlSetData($aRow[$i][9],  $aLoadValues[$j][3], "") ; $regDN
				GUICtrlSetData($aRow[$i][10], $aLoadValues[$j][4], "") ; $regDV
			EndIf
		Next

    Next

EndFunc   ;==>subLoadValuesToGUI




Func Test($rootdir, $path, $switch)
    ; Run test command with argument
    If FileExists($rootdir & "\" & $path) Then
        If GetFileExtension($aFiles[$i]) == ".exe" Then Run($rootdir & '\' & $path & ' ' & $switch)
        If GetFileExtension($aFiles[$i]) == ".msi" Then Run('msiexec /i "' & $rootdir & '\' & $path & '" ' & $switch)
		If GetFileExtension($aFiles[$i]) == ".msp" Then Run('msiexec /update "' & $rootdir & '\' & $path & '" ' & $switch)
        If GetFileExtension($aFiles[$i]) == ".mst" Then Run('msiexec /update "' & $rootdir & '\' & $path & '" ' & $switch)
    Else
        MsgBox(0, "Error locating file", "Cannot find file @ " & @CRLF & $rootdir & "\" & $path)
    EndIf
EndFunc   ;==>Test


Func GetFileExtension($File)
    ; return file extension, eg .exe, .msi, .mst
    For $YLoop = StringLen($File) To 1 Step -1
        If StringMid($File, $YLoop, 1) == "." Then
            $fl_Ext = StringMid($File, $YLoop)
            $YLoop = 1
        EndIf
    Next
    Return StringLower($fl_Ext)
EndFunc   ;==>GetFileExtension


Func DeleteAllGUIobjects()

    _GUIScrollBars_SetScrollInfoPos($hGUI, $SB_VERT, 0)

    For $i = 1 To UBound($aRow) - 1
        For $j = 0 To UBound($aRow, 2) - 1
            GUICtrlDelete($aRow[$i][$j])
        Next
    Next
EndFunc   ;==>DeleteAllGUIobjects


Func CreateRows(ByRef $aFiles)
    ; Takes an array of files and creates rows as required
    ;$lbOrder               = 0
    ;$idButtonUp            = 1
    ;$idButtonDown          = 2
    ;$idComboPath           = 3
    ;$idIcon                = 4
    ;$idComboSwitch         = 5
    ;$idButtonTest          = 6
    ;$idButtonDelete        = 7
    ;$idInputRegKey         = 8
    ;$idInputDisplayName    = 9
    ;$idInputDisplayVersion = 10

    Dim $aRow[($aFiles[0] + 1)][10 + 1] ; variable used in the while loop
    For $i = 1 To $aFiles[0]
        $x = $i * 30 + 4 + $i * 26 ; first value = 60
        ; Left column
        $aRow[$i][0] = GUICtrlCreateLabel("[" & $i & "]", 8, $x + 10, 26, 21, BitOR($SS_CENTER, $SS_CENTERIMAGE))
            GUICtrlSetState(-1, BitOR($GUI_CHECKED, $GUI_SHOW, $GUI_ENABLE))
            GUICtrlSetFont(-1, 10, 400, 0, "MS Sans Serif")
        $aRow[$i][1] = GUICtrlCreateButton("Up", 34, $x + 10, 39, 21)
        $aRow[$i][2] = GUICtrlCreateButton("Down", 74, $x + 10, 39, 21)
        ; Middle column
        $aRow[$i][3] = GUICtrlCreateCombo($aFiles[$i], 113, $x, 393, 21)
        $aRow[$i][4] = GUICtrlCreateIcon($sDir & "\" & $aFiles[$i], -1, 506, $x - 2, 24, 24)
            GUISetIcon(-1)
        $aRow[$i][5] = GUICtrlCreateCombo("", 527, $x, 107, 21)
            If GetFileExtension($aFiles[$i]) == ".exe" Then GUICtrlSetData(-1, '/s|/S|/q|/s /v"/passive"|/s /v"/passive /norestart"|/S/v/qn /V"/qb"|/VERYSILENT /SUPPRESSMSGBOXES /NORESTART /SP-|/silent|/silent /accepteula|/sAll|-silent|-silent -eulaAccepted|/?', '/s')
            If GetFileExtension($aFiles[$i]) == ".msi" Then GUICtrlSetData(-1, '/qb|/passive|/passive ALLUSERS=1|/?', '/passive')
			If GetFileExtension($aFiles[$i]) == ".msp" Then GUICtrlSetData(-1, '/qb|/passive|/passive ALLUSERS=1|/?', '/passive')
            If GetFileExtension($aFiles[$i]) == ".mst" Then GUICtrlSetData(-1, '/qb|/passive|/passive ALLUSERS=1|/?', '/passive')
        ; Right column
        $aRow[$i][6] = GUICtrlCreateButton("Test", 634, $x, 55, 21)
        $aRow[$i][7] = GUICtrlCreateButton("Delete", 789, $x, 39, 21)
        ; Lower column
        $aRow[$i][8] = GUICtrlCreateInput("", 113, $x + 26, 237, 21) ; Uninstall Registry Key
        $aRow[$i][9] = GUICtrlCreateInput("", 359, $x + 26, 213, 21) ; DisplayName
        $aRow[$i][10] = GUICtrlCreateInput("", 582, $x + 26, 107, 21) ; DisplayVersion
        ;_ArrayDisplay($aRow)
    Next

    ; Hide scrollbar
    _GUIScrollBars_ShowScrollBar($hGUI, $SB_VERT, False)
    If $x + 60 > 550 Then
        ; If a scrollbar is needed
        $aRet = _GUIScrollbars_Size(0, $x + 60, 850, 550)
        ; Reshow it with the correct parameters
        _GUIScrollBars_SetScrollInfoPage($hGUI, $SB_VERT, $aRet[2])
        _GUIScrollBars_SetScrollInfoMax($hGUI, $SB_VERT, $aRet[3])
        _GUIScrollBars_ShowScrollBar($hGUI, $SB_VERT, True)
    EndIf


EndFunc   ;==>CreateRows


Func _ArrayRemoveBlanks(ByRef $arr)
    $idx = 0
    For $i = 0 To UBound($arr) - 1
        If $arr[$i] <> "" Then
            $arr[$idx] = $arr[$i]
            $idx += 1
        EndIf
    Next
    ReDim $arr[$idx]
EndFunc   ;==>_ArrayRemoveBlanks


Func _ArraySwap2(ByRef $1, ByRef $2)
    Local $Tmp = $1
    $1 = $2
    $2 = $Tmp
EndFunc   ;==>_ArraySwap2

Func _Scrollbars_WM_VSCROLL($hWnd, $Msg, $wParam, $lParam)

    #forceref $Msg, $wParam, $lParam
    Local $nScrollCode = BitAND($wParam, 0x0000FFFF)
    Local $iIndex = -1, $yChar, $yPos
    Local $Min, $Max, $Page, $Pos, $TrackPos

    For $x = 0 To UBound($__g_aSB_WindowInfo) - 1
        If $__g_aSB_WindowInfo[$x][0] = $hWnd Then
            $iIndex = $x
            $yChar = $__g_aSB_WindowInfo[$iIndex][3]
            ExitLoop
        EndIf
    Next
    If $iIndex = -1 Then Return 0

    Local $tSCROLLINFO = _GUIScrollBars_GetScrollInfoEx($hWnd, $SB_VERT)
    $Min = DllStructGetData($tSCROLLINFO, "nMin")
    $Max = DllStructGetData($tSCROLLINFO, "nMax")
    $Page = DllStructGetData($tSCROLLINFO, "nPage")
    $yPos = DllStructGetData($tSCROLLINFO, "nPos")
    $Pos = $yPos
    $TrackPos = DllStructGetData($tSCROLLINFO, "nTrackPos")

    Switch $nScrollCode
        Case $SB_TOP
            DllStructSetData($tSCROLLINFO, "nPos", $Min)
        Case $SB_BOTTOM
            DllStructSetData($tSCROLLINFO, "nPos", $Max)
        Case $SB_LINEUP
            DllStructSetData($tSCROLLINFO, "nPos", $Pos - 1)
        Case $SB_LINEDOWN
            DllStructSetData($tSCROLLINFO, "nPos", $Pos + 1)
        Case $SB_PAGEUP
            DllStructSetData($tSCROLLINFO, "nPos", $Pos - $Page)
        Case $SB_PAGEDOWN
            DllStructSetData($tSCROLLINFO, "nPos", $Pos + $Page)
        Case $SB_THUMBTRACK
            DllStructSetData($tSCROLLINFO, "nPos", $TrackPos)
    EndSwitch

    DllStructSetData($tSCROLLINFO, "fMask", $SIF_POS)
    _GUIScrollBars_SetScrollInfo($hWnd, $SB_VERT, $tSCROLLINFO)
    _GUIScrollBars_GetScrollInfo($hWnd, $SB_VERT, $tSCROLLINFO)

    $Pos = DllStructGetData($tSCROLLINFO, "nPos")
    If ($Pos <> $yPos) Then
        _GUIScrollBars_ScrollWindow($hWnd, 0, $yChar * ($yPos - $Pos))
        $yPos = $Pos
    EndIf

    Return $GUI_RUNDEFMSG

EndFunc   ;==>_Scrollbars_WM_VSCROLL


Func RegShot(ByRef $array)
	Local $iTimer = TimerInit()
	; Create an array of uninstall keys in the registry
	Dim $regU[3]
	$regU[0] = "HKLM\SOFTWARE\Microsoft\Windows\CurrentVersion\Uninstall"
	$regU[1] = "HKLM\SOFTWARE\Wow6432Node\Microsoft\Windows\CurrentVersion\Uninstall"
	$regU[2] = "HKLM64\SOFTWARE\Microsoft\Windows\CurrentVersion\Uninstall"
	Dim $array[100][3]
	For $i = 0 To UBound($regU) - 1 ; Loop through uninstall version in registry
		$j = 0
		While 1 ; Loop through all keys - exit when done
			$j += 1
			If UBound($array, $UBOUND_ROWS) <= $j Then
				ReDim $array[$j + 100][3]
			EndIf
			$array[$j][$i] = $regU[$i] & "\" & RegEnumKey($regU[$i], $j)
			$array[0][$i] = $j ; so we only need to compare the first column for changes
			;ConsoleWrite("$regU - " & $j & ": " & $regU[$i] & "\" & RegEnumKey($regU[$i], $j) & @CRLF)
			;_ArrayDisplay($array)
			$var = RegEnumKey($regU[$i], $j)
			If @error Then ExitLoop ; - no more keys found, exitloop
		WEnd
	Next
	$iTimer = Round(TimerDiff($iTimer), 2) & "ms"
	;ConsoleWrite("Time = " & $iTimer & @CRLF)
	;_ArrayDisplay($array)
EndFunc   ;==>RegShot


Func MyAdLibFunc()
	Local $iTimer = TimerInit()
	; Assign a static variable to hold the number of times the function is called.
	Local Static $iCount = 0
	$iCount += 1
	RegShot($aRegistryScan)

	;_ArrayDisplay(_Separate($aRegistryTest, $aRegistryScan))
	;$avResult = _ArrayDiff($aRegistryTest, $aRegistryScan)

	; Scan all three registries
	;_ArrayDisplay($aRegistryTest)
	If $aRegistryTest[0][0] <> 0 Then
		For $h = 0 To 2
			If $aRegistryScan[0][$h] <> $aRegistryTest[0][$h] Then
				;_ArrayDisplay($aRegistryScan, "$aRegistryScan")
				;_ArrayDisplay($aRegistryTest, "$aRegistryTest")
				ConsoleWrite("Test reg " & $h & ": " & $aRegistryTest[0][$h] & @CRLF)
				ConsoleWrite("Scan reg " & $h & ": " & $aRegistryScan[0][$h] & @CRLF)
				;ConsoleWrite("Test reg " & $h & ": " & UBound($aRegistryTest, $h) & @CRLF)
				;ConsoleWrite("Scan reg " & $h & ": " & UBound($aRegistryScan, $h) & @CRLF)

				; compare arrays for change
				$in0 = _ArrayUnique($aRegistryTest, $h, Default, Default, 0)
				$in1 = _ArrayUnique($aRegistryScan, $h, Default, Default, 0)
				$aDiff = _Separate($in0, $in1)
				If Not $aDiff[1][1] = "" Then
					;MsgBox(0,"Added Key", $aDiff[1][1])
					$sRegKey = StringSplit($aDiff[1][1],"\")
					GUICtrlSetData($aRow[$sTestRow][8], $sRegKey[$sRegKey[0]])
					$sRegDN = RegRead($aDiff[1][1], "DisplayName")
					If Not @error Then GUICtrlSetData($aRow[$sTestRow][9], $sRegDN)
					$sRegDV = RegRead($aDiff[1][1], "DisplayVersion")
					If Not @error Then GUICtrlSetData($aRow[$sTestRow][10], $sRegDV)
					; set $aFile array to update when gui rebuilt
					;
					;
				EndIf

				 ; reset test array
				ReDim $aRegistryTest[1][1]
				$aRegistryTest[0][0] = 0
				ExitLoop

			EndIf
		Next
	EndIf

	$iTimer = Round(TimerDiff($iTimer), 2) & "ms"
	;_ArrayDisplay($avResult, "Time = " & $iTimer)
	;ConsoleWrite("Time = " & $iTimer)
	;ConsoleWrite(" MyAdLibFunc called " & $iCount & " time(s)" & @CRLF)
EndFunc   ;==>MyAdLibFunc


Func _Separate(ByRef $in0, ByRef $in1)
    $in0 = _ArrayUnique($in0, 0, Default, Default, 0)
    $in1 = _ArrayUnique($in1, 0, Default, Default, 0)
    Local $z[2] = [UBound($in0), UBound($in1)], $low = 1 * ($z[0] > $z[1]), $aTemp[$z[Not $low]][3], $aOut = $aTemp, $aNdx[3]
    For $i = 0 To $z[Not $low] - 1
        If $i < $z[0] Then $aTemp[$i][0] = $in0[$i]
        If $i < $z[1] Then $aTemp[$i][1] = $in1[$i]
    Next
    For $i = 0 To $z[$low] - 1
        $x = _ArrayFindAll($aTemp, $aTemp[$i][$low], 0, 0, 1, 0, Not $low)
        If Not @error Then ; both
            For $j = 0 To UBound($x) - 1
                $aTemp[$x[$j]][2] = 1
            Next
            $aOut[$aNdx[2]][2] = $aTemp[$i][$low]
            $aNdx[2] += 1
        Else ; only in $low
            $aOut[$aNdx[$low]][$low] = $aTemp[$i][$low]
            $aNdx[$low] += 1
        EndIf
    Next

	; at this point we have enough information to determine differences

	; can return just removed compare - may be needed as uninstall keys show up unattendedly

    For $i = 0 To $z[Not $low] - 1
        If $aTemp[$i][2] <> 1 Then
            $aOut[$aNdx[Not $low]][Not $low] = $aTemp[$i][Not $low]
            $aNdx[Not $low] += 1
        EndIf
    Next
    ReDim $aOut[_ArrayMax($aNdx)][3]
    Return $aOut
EndFunc   ;==>_Separate

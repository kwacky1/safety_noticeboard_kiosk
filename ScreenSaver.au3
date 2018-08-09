#Region ;**** Directives created by AutoIt3Wrapper_GUI ****
#AutoIt3Wrapper_Outfile=bin\ScreenSaver.scr
#EndRegion ;**** Directives created by AutoIt3Wrapper_GUI ****
#include <GUIConstantsEx.au3>
#include <WinAPISysWin.au3>
#include <..\Includes\Log.au3>
#include <..\Includes\_SS_UDF.au3>
#include <..\Includes\Powerpoint.au3>
$inifile = @ScriptDir & "\Presentation.txt"


Global $oPPT, $oShow
_PPT_ErrorNotify(1)
$log = OpenLog( @TempDir & "\screensaver.log" )
WriteLog( $log, "Screensaver has started" )
$sswin = _SS_GUICreate()
WriteLog( $log, "Creating IE object" )
$oIE = ObjCreate("Shell.Explorer.2")
$objPreview = GUICtrlCreateObj( $oIE, 0, 0, $_SS_WinWidth, $_SS_Winheight )
WriteLog( $log, "IE object created" )
_SS_SetMainLoop( "Showtime" )
_SS_Start()

Func Showtime()
	$x = 1
	$dispitems = IniReadSectionNames( $inifile )
	WriteLog( $log, "Total items found is " & $dispitems[0] )
	Do
		WriteLog( $log, "Checking for Powerpoint Object" )
		If IsObj( $oShow ) Then _PPT_PresentationClose( $oShow, False )
		If IsObj( $oPPT) Then _PPT_Close( $oPPT, False )
		If $x > $dispitems[0] Then $x = 1
		WriteLog( $log, "Section " & $dispitems[$x] & " begins" )
		$type = IniRead( $inifile, $dispitems[$x], "Type", "none" )
		WriteLog( $log, "Type: " & $type )
		$url = IniRead( $inifile, $dispitems[$x], "URL", "" )
		WriteLog( $log, "URL: " & $url )
		$len = IniRead( $inifile, $dispitems[$x], "Length", 0 )
		WriteLog( $log, "Length: " & $len )
		Select
			Case $type = "PowerPoint"
				WriteLog( $log, "Showing blank IE page" )
				$oIE.navigate( "about:blank" )
				WriteLog( $log, "Create Powerpoint Object" )
				$oPPT = _PPT_Open()
				$sShow = $url
				WriteLog( $log, "Opening " & $url )
				$oShow = _PPT_PresentationOpen( $oPPT, $sShow, True, False )
				GUICtrlSetState( $objPreview, $GUI_HIDE )
				WriteLog( $log, "Displaying Slideshow" )
				_PPT_SlideShow( $oShow, True, 1, Default, False, 1 )
				$hShow = WinGetHandle( "[CLASS:screenClass]" )
				_WinAPI_SetParent( $hShow, $sswin )
			Case Else
				WriteLog( $log, "Navigating to " & $url )
				GUICtrlSetState( $objPreview, $GUI_SHOW )
				$oIE.navigate( $url )
		EndSelect
		WriteLog( $log, "Starting countdown timer" )
		$start = TimerInit()
		Do
			If TimerDiff( $start ) > ($len * 1000) + 5000 Then ExitLoop
		Until _SS_ShouldExit()
		WriteLog( $log, "Section " & $dispitems[$x] & " ends" )
		$x += 1
	Until _SS_ShouldExit()
	If IsObj( $oPPT) Then _PPT_Close( $oPPT, False )
EndFunc
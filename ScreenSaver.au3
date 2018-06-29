#include <..\Includes\_SS_UDF.au3>

$inifile = @ScriptDir & "\Presentation.txt"


$sswin = _SS_GUICreate()
$oIE = ObjCreate("Shell.Explorer.2")
$objPreview = GUICtrlCreateObj( $oIE, 0, 0, $_SS_WinWidth, $_SS_Winheight )
_SS_SetMainLoop( "Showtime" )
_SS_Start()

Func Showtime()
	$x = 1
	$dispitems = IniReadSectionNames( $inifile )
	Do
		If $x > $dispitems[0] Then $x = 1
		$url = IniRead( $inifile, $dispitems[$x], "URL", "none" )
		$len = IniRead( $inifile, $dispitems[$x], "Length", 0 )
		$oIE.navigate( $url )
		$start = TimerInit()
		Do
			If TimerDiff( $start ) > ($len * 1000) + 5000 Then ExitLoop
		Until _SS_ShouldExit()
		$x += 1
	Until _SS_ShouldExit()
EndFunc
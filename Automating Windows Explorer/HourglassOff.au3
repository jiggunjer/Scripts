; If the script fails for some reason and you are left with an hour-
; glass cursor, then run HourglassOff.au3 to get the normal cursor back.

Global Const $dllUser32  = DllOpen( "user32.dll" )
#include "Includes\Hourglass.au3"
Hourglass( True )
DllClose( $dllUser32 )

#Include <APIConstants.au3>
#Include <WinAPIEx.au3>

Opt('MustDeclareVars', 1)

Global Const $FaceName = 'Arial'

ConsoleWrite($FaceName & ' Regular => ' & _WinAPI_GetFontName($FaceName, $FS_REGULAR) & @CR)
ConsoleWrite($FaceName & ' Bold => ' & _WinAPI_GetFontName($FaceName, $FS_BOLD) & @CR)
ConsoleWrite($FaceName & ' Italic => ' & _WinAPI_GetFontName($FaceName, $FS_ITALIC) & @CR)
ConsoleWrite($FaceName & ' Bold Italic => ' & _WinAPI_GetFontName($FaceName, BitOR($FS_BOLD, $FS_ITALIC)) & @CR)

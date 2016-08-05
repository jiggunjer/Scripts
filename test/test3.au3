#include <inet.au3>
#include <array.au3>
#include <GuiListView.au3>
#include <GUIConstantsEx.au3>
#include <WindowsConstants.au3>
#include <EditConstants.au3>
#include <StaticConstants.au3>
$test = GUICreate("test", 800, 800, -1, -1)
$input = GUICtrlCreateInput("", 5, 10, 550, 30)
$list = GUICtrlCreateListView("Suggestion", 5, 40, 260, 700)
_GUICtrlListView_SetColumnWidth($list, 0, 245)

$defines = GUICtrlCreateEdit("", 265, 192, 530, 550, $ES_MULTILINE, 0)
$edit = GUICtrlCreateEdit("", 265, 40, 295, 150, $ES_MULTILINE, 0)
$from = GUICtrlCreateCombo("en - English", 560, 13, 90, 20)
GUICtrlSetData(-1, "es - Spanish|sq - Albanian|ar - Arabic|bg - Bulgarian|ca - Catalan|hr - Croatian|cs - Czeh|da - Danish|nl - Dutch|et - Estonian|tl - Filipino|fi - Finish|fr - French|gl - Galacian|de - German|el - Greek|iw - Hebrew|hi - Hindi|hu - Hungarian|id - Indonesian|it - Italian|lv - Latvian|vi - Vietnamese|tr - Turkish|sv - Sweedish|ru - Russian|pt - Portuguese")
$to = GUICtrlCreateCombo("es - Spanish", 650, 13, 90, 20)
GUICtrlSetData(-1, "en - English|sq - Albanian|ar - Arabic|bg - Bulgarian|ca - Catalan|hr - Croatian|cs - Czeh|da - Danish|nl - Dutch|et - Estonian|tl - Filipino|fi - Finish|fr - French|gl - Galacian|de - German|el - Greek|iw - Hebrew|hi - Hindi|hu - Hungarian|id - Indonesian|it - Italian|lv - Latvian|vi - Vietnamese|tr - Turkish|sv - Sweedish|ru - Russian|pt - Portuguese")
GUICtrlCreateLabel("", 563, 40, 235, 150)
GUICtrlSetData(-1, "spanish = es, Albanian = sq, Arabic = ar, Bulgarian = bg,Catalan = ca, Croatian = hr, Czech = cs,Danish = da" & _
        "dutch = nl,Estonian = et,Filipino = tl, Finnish = fi, French = fr, Galician = gl,German = de,Greek = el" & _
        "Hebrew = iw,Hindi = hi - no, Hungarian = hu,Indonesian = id, Italian = it, Latvian = lv,Vietnamese = vi" & _
        "Turkish = tr,Swedish = sv,Russian = ru, Portuguese = pt, English = en")
GUISetState()
$olddata = ""

ConsoleWrite(_GUICtrlListView_GetSelectedCount($list) > 0 & @CRLF)
While 1
    $msg = GUIGetMsg()
    Switch $msg
        Case - 3
            Exit
        Case $from, $to
            $Translate = _GoogleTranslate(GUICtrlRead($input), StringLeft(GUICtrlRead($to), 2), StringLeft(GUICtrlRead($from), 2))
            GUICtrlSetData($edit, $Translate)
    EndSwitch
    If _GUICtrlListView_GetSelectedCount($list) <> False Then
        $found = _GUICtrlListView_GetItemTextString($list, -1)
        GUICtrlSetData($input, $found)
    EndIf
    If GUICtrlRead($input) <> "" Then
        $data = GUICtrlRead($input)
        If $olddata <> $data Then
            If StringLen($data) < 3 Then
            Else
                $suggestions = _GoogleSuggestions($data)
                _GUICtrlListView_DeleteAllItems($list)
                GUICtrlSetData($defines, "")
                GUICtrlSetData($edit, "")

                For $i = 1 To UBound($suggestions) - 1
                    GUICtrlCreateListViewItem($suggestions[$i], $list)
                    ConsoleWrite($suggestions[$i] & @CRLF)
                Next
                $Definitions = _GoogleDefine($data)
                For $i = 0 To UBound($Definitions) - 1
                    GUICtrlSetData($defines, $i & ": " & $Definitions[$i] & @CRLF, $defines)
                Next

                $Translate = _GoogleTranslate($data, StringLeft(GUICtrlRead($to), 2), StringLeft(GUICtrlRead($from), 2))
                GUICtrlSetData($edit, $Translate)


            EndIf
        EndIf
        $olddata = $data
    EndIf

WEnd


$suggestions = _GoogleSuggestions('autoit s')
_ArrayDisplay($suggestions, 'Suggestions')

$Definitions = _GoogleDefine("Pulp")
_ArrayDisplay($Definitions, 'Pulp ')

$Translate = _GoogleTranslate("Hello how are you?")
MsgBox(0, 'English to Spanish  ', "Hello how are you?" & @CRLF & $Translate)


Func _GoogleSuggestions($sSuggest); Retruns an Array of Suggestions
    Local $sSource, $aResults;, $aQueries
    $sSource = _INetGetSource("http://google.com/complete/search?output=toolbar&q=" & $sSuggest)
    If @error Then Return SetError(1)
    $aResults = StringRegExp($sSource, '<CompleteSuggestion><suggestion data="(.*?)"/>', 3)
;~ $sQueries = StringRegExp($source, '"/><num_queries int="(\d{0,})"/>', 3)
    Return $aResults
EndFunc   ;==>_GoogleSuggestions

Func _GoogleDefine($sKeyWord)
    Local $aDefintions, $sSource
    $sSource = _INetGetSource("http://www.google.com/search?q=define%3A" & $sKeyWord)
    If @error Then Return SetError(1)
    $aDefintions = StringRegExp($sSource, "<li>(.*?)<br>", 1)
    If IsArray($aDefintions) Then
        Return StringSplit(StringRegExpReplace($aDefintions[0], '(<li>)', '|'), '|')
    Else
        Return ""
    EndIf
EndFunc   ;==>_GoogleDefine

Func _GoogleTranslate($sText, $sTo = "es", $sFrom = "en")
    Local $sTranslation, $sUrl, $sSource
    $sUrl = StringFormat("http://ajax.googleapis.com/ajax/services/language/translate?v=1.0&q=%s&langpair=%s%%7C%s", $sText, $sFrom, $sTo)
    $sSource = _INetGetSource($sUrl)
    $sTranslation = StringRegExp(BinaryToString($sSource, 4), '"translatedText":"([^"]+)"', 1)
    If IsArray($sTranslation) Then
        Return $sTranslation[0]
    Else
        Return ""
    EndIf
EndFunc   ;==>_GoogleTranslate
;  Your   code   here
#IfWinActive, ahk_exe iexplore.exe
:*:IECLEAR::
	URLs := []
	for wb in ComObjCreate("Shell.Application").Windows
		if instr(wb.FullName,"iexplore.exe") {
			i := URLs.Push(wb.LocationURL)
			str := str . "," . wb.LocationURL
		}
		
	While Winexist("ahk_exe iexplore.exe"){
		Process,Close, iexplore.exe
	} 
	
	IE := ComObjCreate("InternetExplorer.Application")
	IE.Visible := 1
	Loop, % URLs.Length()
	{
		
		Try {
			IE.Navigate2(URLs[A_Index],2048)
		} Catch e {
			Try {
				IE := IEGetByLocationURL(URLs[1])
				IE.Navigate2(URLs[A_Index],2048)
			} Catch e {
					Msgbox, An Unknown error Occurred: `r`n%e%
			}
		}
	}
	
Return

IEGetByLocationURL(URL="") {
   for wb in ComObjCreate("Shell.Application").Windows
	{
    if wb.LocationURL=URL and InStr(wb.FullName, "iexplore.exe")
		return wb
	}
}
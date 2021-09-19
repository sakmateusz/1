DzisSystem = CStr(Replace(FormatDateTime(Date, vbShortDate), "-", "."))
arr = Split(DzisSystem, ".")
Dzis = arr(0) & "." & arr(1) & "." & arr(2) 

miesiac = CStr(Month(date))
ODPMm = CStr(Month(date)-1)
ODPMy = CStr(Year(date))

'jeżeli poprzedni miesiąc wynosi wartość to:
if ODPMm = "4" OR "6" OR "9" OR "11" then
	ODPMd = "31"
elseif ODPMm = "1" OR "3" OR "5" OR "7" OR "8" OR "10" OR "12" then
	ODPMd = "30"
elseif ODPMm = "2" then
	ODPMd = "28"
End if


msgbox("Dzis jest: " & Dzis)
msgbox("Teraz mamy miesiac: " & miesiac)
msgbox("Teraz mamy rok: " & ODPMy)
msgbox("Wczesniej byl miesiac: " & ODPMm)

msgbox("Ostatni dzien poprzedniego mc-a: " & ODPMy & "." & ODPMm & "." & ODPMd)
DzisSystem = CStr(Replace(FormatDateTime(Date, vbShortDate), "-", "."))
arr = Split(DzisSystem, ".")
Dzis = arr(0) & "." & arr(1) & "." & arr(2) 

miesiac = CStr(Month(date))
miesiac = Right("000" & miesiac, 2)

ODPMm = CStr(Month(date)-1)
ODPMm = Right("000" & ODPMm, 2)

ODPMy = CStr(Year(date))
if CStr(Month(date)) = "12" then
	ODPMy = CStr(Year(date)-1)
End if

ODPMyb = CStr(Year(date)-1)

if ODPMm = "04" OR "06" OR "09" OR "11" then
	ODPMd = "31"
elseif ODPMm = "01" OR "03" OR "05" OR "07" OR "08" OR "10" OR "12" then
	ODPMd = "30"
elseif ODPMm = "02" then
	x = (Year(date))/400
	if Len(x) = 4 then
		ODPMd = "29"
	else 
		ODPMd = "28"
	End if
End if


msgbox("Dzis jest: " & Dzis)
msgbox("Teraz mamy miesiac: " & miesiac)
msgbox("Teraz mamy rok: " & ODPMy)
msgbox("Wczesniej byl rok: " & ODPMyb)
msgbox("Wczesniej byl miesiac: " & ODPMm)

msgbox("Ostatni dzien poprzedniego mc-a: " & ODPMd & "." & ODPMm & "." & ODPMy)
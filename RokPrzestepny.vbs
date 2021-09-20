x = (Year(date))/400
if Len(x) = 4 then
	msgbox("Rok " & (Year(date)) & " to rok przestepny")
else 
	msgbox("Rok " & (Year(date)) & " to NIE jest rok przestepny")
End if
'msgbox(x)

Z = 2028
y = Z/400
if Len(y) = 4 then
	msgbox("Rok " & Z & " to rok przestepny")
else 
	msgbox("Rok " & Z & " to NIE jest rok przestepny")
End if
'msgbox(y)
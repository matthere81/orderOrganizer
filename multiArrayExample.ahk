#SingleInstance, force

data := "
(
iPhone1,1:iPhone
iPhone1,2:iPhone 3G
iPhone2,1:iPhone 3GS
iPhone3,1:iPhone 4
iPhone3,2:iPhone 4 GSM Rev A
iPhone3,3:iPhone 4 CDMA
iPhone4,1:iPhone 4S
)"

a := []

For k, v in StrSplit(data, "`n")
	For j, w in StrSplit(v, ":")
		a[k,j] := w

MsgBox, % a[3][1] " -> " a[1][2]    ; iPhone1,1 -> iPhone
return
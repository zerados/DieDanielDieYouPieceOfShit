Set WshShell = WScript.CreateObject("WScript.Shell")
dim r
randomize
r = int(rnd*1800000) + 300000
randomize
sup = int(rnd+5)+ 1
num = 1
do while (true)
	Select Case sup
		case 1
			WshShell.SendKeys "{DOWN}"
		case 2
			WshShell.SendKeys "{TAB}"
		case 3
			WshShell.SendKeys "{RIGHT}"
		case 4
			WshShell.SendKeys "{PGDN}"
		case 5
			WshShell.SendKeys "{CAPSLOCK}"
		case 6
			WshShell.SendKeys "{INSERT}"
		case 7
			WshShell.SendKeys "{PGDN}"
		case 8
			WshShell.SendKeys "{HOME}"
		Case else
			WshShell.SendKeys "{CAPSLOCK}"
	End Select
	randomize
	r = int(((rnd*1800000) + 300000)/num)
	sup = int(rnd*10)+ 1
	num = num + 1
	WScript.Sleep r
loop
if type('_screen.cThorDispatcher') = 'C'
	lcFoxBin2PRGPath = execscript(_screen.cThorDispatcher, 'Thor_Proc_GetFoxBin2PrgFolder')
else
	lcFoxBin2PRGPath = getdir('', 'Locate FoxBin2PRG')
endif type('_screen.cThorDispatcher') = 'C'
if empty(lcFoxBin2PRGPath)
	messagebox("FoxBin2Prg not found", 0, "Error...")
	return	
endif empty(lcFoxBin2PRGPath)
lcFoxBin2PRG = forcepath('FoxBin2Prg.prg', lcFoxBin2PRGPath)
if not file(lcFoxBin2PRG)
	messagebox("FoxBin2Prg app not found", 0, "Error...")
	Return
EndIf

Do (lcFoxBin2Prg) With "GoFish5.pjx", "*"
Return

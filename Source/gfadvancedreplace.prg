#Define CR Chr[13]
#Define LF Chr[10]

* GF Advanced replace
* Parameter lcText in the line of code where a match was found
* Result is the new line of code.

* As Matt would have said, careful, buddy!

*** JRN 2023-01-31 : returning 0 or .F. will not change the line
*** JRN 2024-09-19 : returning null deletes the line

Lparameters tcOrigtext

Local lcAlias, lcLeft, lcNew, lcOrder, lcParams, lcResult, lnPos

* Paste in one the examples below and modify as appropriate

Return m.lcResult



* ================================================================================ 
* Example 8: Add   &lcFont   to the end of the line
lcResult =  m.tcOrigtext + ' &lcFont'


* ================================================================================ 
* Example 7: Create two new lines, one before, one after
lcResult =  'Local lcFont' + CR + LF + m.tcOrigtext + CR + LF + 'lcFont = ThisForm.GetContextMenuFont()'


* ================================================================================ 
* Example 6: Split  Procedure SomeProc(lparam1, param2 ...) into separate lines
lcParams = Strextract(m.tcOrigtext, '(', ')', 1)
If Empty(m.lcParams)
	Return 0
Endif

lnPos	 = At('(', m.tcOrigtext)
lcResult = Left(m.tcOrigtext, m.lnPos - 1) + Chr[13] + Chr[10] + 'LParameters ' + m.lcParams


* ================================================================================ 
* Example 5:  Change   prompt = (any value)  to    AddID = 'ThorRepository'
lcLeft	 = Left(m.lcResult, At('=', m.lcResult))
lcNew	 = Strtran(m.lcLeft, 'prompt', 'AppID ', 1, 1, 1) + [ 'ThorRepository']
lcResult = m.lcResult + CRLF + m.lcNew


* ================================================================================ 
* Example 4:  Change   'Z:\Somefilename'   to    MapToUNC('Somefilename')
lcResult = m.tcOrigtext

lcResult = Fix4(m.lcResult, ['], ['])
lcResult = Fix4(m.lcResult, ["], ["])
lcResult = Fix4(m.lcResult, '[', ']')

Return m.lcResult


* ================================================================================ 
* Example 3: Insert a call to DoDefault into a procedure
* Used:
*		after Whole Word Search
*		for Procedure*Click
*		Using filter form to exclude procedures already having references to DoDefault
lcResult =  m.tcOrigtext + CR + LF + CR + LF + 'DoDefault()' + CR + LF


* ================================================================================ 
* Example 2: Add comment && To be investigated    to the end of the line
lcResult =  m.tcOrigtext + ' &' + '& To be investigated'


* ================================================================================ 
* The very first use of Advanced Replace, summer 2015 ... over 5,200 replacements
* Example 1: Map calls to this procedure:
*		OpenFile2(tcPath, tcFilename, tnWorkArea, tcAlias, tlShared, tcOrder, tlNoFailure, tcIndexes)
*   to this:
*		UseTable(tcAlias, tcOrder)

If Occurs('(', m.tcOrigtext) # 1
	Return 0
Endif

lcParams = Strextract(m.tcOrigtext, '(', ')')

lcAlias = Alltrim(Getwordnum(m.lcParams, 4, ','))
If Empty(m.lcAlias)
	Return 0
Endif

lcOrder = Alltrim(Getwordnum(m.lcParams, 6, ','))
If Empty(m.lcOrder)
	lcResult = Textmerge([UseTable(<<lcAlias>>)])
Else
	lcResult =  Textmerge([UseTable(<<lcAlias>>, <<lcOrder>>)])
Endif

lcResult = m.lcResult + '  &' + '& ' + Alltrim(m.tcOrigtext, 1, ' ', Tab)



* ================================================================================
* ================================================================================
Procedure Fix4(lcText, lcLeft, lcRight)
	lcMatch = Strextract(m.lcText, m.lcLeft + 'Z:\', m.lcRight, 1, 5)
	If Empty(m.lcMatch)
		Return m.lcText
	Else
		Return Strtran(m.lcText, m.lcMatch, ' MapToUNC(' + m.lcMatch + ')', 1, 1, 1)
	Endif
Endproc



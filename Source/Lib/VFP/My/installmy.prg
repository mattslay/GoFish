#include MyConstants.H
local lnSelect, ;
	lcAlias, ;
	loException as Exception, ;
	lcCode

* Open the IntelliSense table.

lnSelect = select()
select 0
try
	use (_foxcode) again shared
	lcAlias = alias()
catch to loException
	messagebox(ccERR_COULD_NOT_OPEN_FOXCODE_LOC + ccCR + ccCR + ;
		loException.Message, MB_ICONEXCLAMATION, ccCAP_MY_INSTALLATION_LOC)
endtry
if not empty(lcAlias)

* Add the TYPE record, used for IntelliSense on LOCAL My AS, if necessary.

	locate for TYPE = 'T' and upper(ABBREV) = padr('MY', len(ABBREV)) and ;
		not deleted()
	if not found()
		insert into (lcAlias) ;
				(TYPE, ;
				ABBREV, ;
				CMD, ;
				DATA) ;
			values ;
				('T', ;
				'My', ;
				'{myscript}', ;
				'My')
	endif not found()

* Create the code for the script record.

	lcCode = GetScriptCode()

* Add the script record if necessary.

	locate for TYPE = 'S' and ;
		upper(ABBREV) = padr('MYSCRIPT', len(ABBREV)) and not deleted()
	if not found()
		insert into (lcAlias) ;
				(TYPE, ;
				ABBREV) ;
			values ;
				('S', ;
				'MyScript')
	endif not found()
	replace DATA with lcCode

* Clean up and exit.

	use in (lcAlias)
	messagebox(ccMSG_INSTALL_SUCCESS_LOC, MB_ICONINFORMATION, ;
		ccCAP_MY_INSTALLATION_LOC)
endif not empty(lcAlias)
select (lnSelect)

function GetScriptCode
local lcDirectory, ;
	lcCode
lcDirectory = sys(16, 1)
lcDirectory = addbs(justpath(substr(lcDirectory, at(' ', lcDirectory, 2) + 1)))
text to lcCode noshow textmerge
lparameters toFoxcode
local loFoxCodeLoader, ;
	luReturn
if file(_codesense)
	set procedure to (_codesense) additive
	loFoxCodeLoader = createobject('FoxCodeLoader')
	luReturn        = loFoxCodeLoader.Start(toFoxcode)
	loFoxCodeLoader = .NULL.
	if atc(_codesense, set('PROCEDURE')) > 0
		release procedure (_codesense)
	endif atc(_codesense, set('PROCEDURE')) > 0
else
	luReturn = ''
endif file(_codesense)
return luReturn

define class FoxCodeLoader as FoxCodeScript
	cProxyClass    = 'MyFoxCode'
	cProxyClasslib = '<<lcDirectory>>my.vcx'

	procedure Main
		local loFoxCode, ;
			luReturn
		loFoxCode = newobject(This.cProxyClass, This.cProxyClasslib)
		if vartype(loFoxCode) = 'O'
			luReturn = loFoxCode.Main(This.oFoxCode)
		else
			luReturn = ''
		endif vartype(loFoxCode) = 'O'
		return luReturn
	endproc
enddefine
endtext
return lcCode


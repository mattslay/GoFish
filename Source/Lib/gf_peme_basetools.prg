#Include GoFish.h		&& Include constants so these, for example "CR", can be used here ...

Define Class gf_peme_basetools As Custom

	lReleaseOnDestroy = .F.

	*----------------------------------------------------------------------------------
	Procedure AddMRUFile(lcFileName, lcClassName, lcMRU_ID)

		#Define DELIMITERCHAR  Chr(0)
		#Define MAXITEMS       24
		#Define ResourceAlias  crsr_MRU_Resource_Add

		Local lcData, lcLine25, lcNewData, lcSearchString, lcSourceFile, lnPos, lnSelect

		If 'ON' # Set ('Resource')
			Return
		Endif

		If lcFileName # '\-'
			lcFileName = This.DiskFileName(FullPath(lcFileName))
		EndIf

		lcSourceFile = Set ('Resource', 1)

		If Empty (lcMRU_ID)
			lcMRU_ID = This.GetMRUID (lcFileName)
			If '?' $ lcMRU_ID
				Return
			Endif
		Endif

		If lcMRU_ID = 'MRUI'
			If Empty (lcClassName) && Class library (artificial)
				lcMRU_ID	   = 'MRU2'
				lcSearchString = lcFileName + DELIMITERCHAR
			Else
				lcSearchString = lcFileName + '|' + Lower (lcClassName) + DELIMITERCHAR
			Endif
		Else
			lcSearchString = lcFileName + DELIMITERCHAR
		Endif

		lnSelect = Select()
		Select 0
		Use (lcSourceFile) Again Shared Alias ResourceAlias

		Locate For Id = lcMRU_ID
		If Found()
			lcData = DELIMITERCHAR + Substr (Data, 3)
			lnPos  = Atcc (DELIMITERCHAR + lcSearchString, lcData)
			Do Case
				Case lnPos = 1
					* already tops of the list
					lcNewData = Data
				Case lnPos = 0
					* must add to list
					lcNewData = Stuff (Data, 3, 0, lcSearchString)
					* note that GetWordNum won't accept CHR(0) as a delimiter
					lcLine25  = Getwordnum (Chrtran (Substr (lcNewData, 3), DELIMITERCHAR, CR), MAXITEMS + 1, CR)
					If Not Empty (lcLine25)
						lcNewData = Strtran (lcNewData, DELIMITERCHAR + lcLine25 + DELIMITERCHAR, DELIMITERCHAR, 1, 1, 1)
					Endif
				Otherwise
					lcNewData = Stuff (Data, lnPos + 1, Len (lcSearchString), '')
					lcNewData = Stuff (lcNewData, 3, 0, lcSearchString)
			Endcase
			Replace																;
					Data	 With  lcNewData									;
					CkVal	 With  Val (Sys(2007, Substr (lcNewData, 3)))		;
					Updated	 With  Date()
		Else
			lcNewData = Chr(4) + DELIMITERCHAR + lcSearchString
			Insert Into ResourceAlias					;
				(Type, Id, CkVal, Data, Updated)		;
				Values									;
				('PREFW', lcMRU_ID, Val (Sys(2007, Substr (lcNewData, 3))), lcNewData, Date())

		Endif

		Use
		Select (lnSelect)

		Return
	
	EndProc


	*----------------------------------------------------------------------------------
	Procedure BrowseTable(tcFileName, tnRecno)

		Local lcCursor, lnShiftX, lnShiftY, loForm, loTempForm, luSBReturn, lnDataSession

		lcCursor = 'GoFishTableBrowse'
		*-- Create a temp form so we can control the Browse window size and location. 
		*-- Don't want it to fit hte same size and location as the GoFish Window, becuase
		*-- it looks weird when that happens.
		loTempForm = .null.

		If Vartype(_Screen.ActiveForm) = 'O'
			lnShiftX = 50
			lnShiftY = 125
			loTempForm = Createobject('Form')
			loForm = _Screen.ActiveForm
			loTempForm.Left = loForm.Left + lnShiftX
			loTempForm.Width = loForm.Width - lnShiftX
			loTempForm.Top = loForm.Top + lnShiftY
			loTempForm.Height = loForm.Height - lnShiftY
			loTempForm.Show()
		Endif

		If Used(lcCursor)
			Select (lcCursor)
		Else
			Select 0
		Endif

		luSBReturn = .null.

		*-- Call SuperBrowse, if Thor is present
		If Type('_Screen.cThorDispatcher') = 'C'
			lnDataSession = Set('DataSession')
			Set DataSession To 1
			Use (tcFileName) Again Alias &lcCursor

			If !Empty(tnRecno)
				Goto tnRecno
			Endif

			luSBReturn = Execscript(_Screen.cThorDispatcher, 'Thor_Proc_SuperBrowse', lcCursor) != .null.
			Set DataSession To &lnDataSession
		EndIf

		*-- Regular brows is Thor not present, or SuperBrowse call failed
		If luSBReturn = .null.
			Use (tcFileName) Again Alias &lcCursor

			If !Empty(tnRecno)
				Goto tnRecno
			Endif

			Browse Last Nowait Title (tcFileName) Normal
		Endif

		If Vartype(loTempForm) = 'O'
			loTempForm.Release()
		Endif

	EndProc


	*----------------------------------------------------------------------------------
	Procedure CheckOutSCC(lcFileName)

		Local lnSelect

		*** JRN 11/17/2010 : Bhavbhuti: Source Control
		* Select 0 -- used because it appears that CheckOut may kill current work area
		If 'O' = Type ('This.oPrefs') And This.oPrefs.lCheckOutSCC And 0 # _vfp.Projects.Count
			lnSelect = Select()
			Select 0
			Try
				If Not Inlist (_vfp.ActiveProject.Files (m.lcFileName).SCCStatus, 0, 2)
					If Not _vfp.ActiveProject.Files (m.lcFileName).CheckOut()

					Endif
				Endif
		    Catch to loException
		        Do Case
		            Case loException.ErrorNo = 1943 && if not found in this project
		            Case loException.ErrorNo = 1426 && OLE error code 0x85ff012e: Unknown COM status code
		                _vfp.ActiveProject.Files (m.lcFileName).CheckOut()
		        Otherwise
		            This.ShowErrorMsg(loException)
		        EndCase
		    EndTry
		    Select (lnSelect)
		Endif

	EndProc


	*----------------------------------------------------------------------------------
	* Clear all the objects I've created
	Procedure Destroy()
	
		Local laMembers[1], lcMember
		Amembers (laMembers, This, 0)
		For Each lcMember In laMembers
		    lcMember = Upper (lcMember)
		    If lcMember = 'O' And lcMember # 'OBJECTS' And "O" = Vartype (Getpem (This, lcMember))
		        This.&lcMember. = .Null.
		    Endif
		Endfor

		This.Release()
		
	EndProc


	*----------------------------------------------------------------------------------
	Procedure DiskFileName(lcFileName)

		#Define MAX_PATH 260

		Local lnFindFileData, lnHandle, lcXXX
		Declare Integer FindFirstFile In win32api String @, String @
		Declare Integer FindNextFile In win32api Integer, String @
		Declare Integer FindClose In win32api Integer

		Do Case
			Case ( Right (lcFileName, 1) == '\' )
				Return Addbs (This.DiskFileName (Left (lcFileName, Len (lcFileName) - 1)))

			Case Empty (lcFileName)
				Return ''

			Case ( Len (lcFileName) == 2 ) And ( Right (lcFileName, 1) == ':' )
				Return Upper (lcFileName)	&& win2k gives curdir() for C:
		Endcase

		lnFindFileData = Space(4 + 8 + 8 + 8 + 4 + 4 + 4 + 4 + MAX_PATH + 14)
		lnHandle		 = FindFirstFile (@lcFileName, @lnFindFileData)

		If ( lnHandle < 0 )
			If ( Not Empty (Justfname (lcFileName)) )
				lcXXX = Justfname (lcFileName)
			Else
				lcXXX = lcFileName
			Endif
		Else
			= FindClose (lnHandle)
			lcXXX	= Substr (lnFindFileData, 45, MAX_PATH)
			lcXXX	= Left (lcXXX, At (Chr(0), lcXXX) - 1)
		Endif


		Do Case
			Case Empty (Justpath (lcFileName))
				Return lcXXX
			Case ( Justpath (lcFileName) == '\' ) And (Left (lcFileName, 2) == '\\')	&& unc
				Return '\\' + lcXXX
			Otherwise
				Return Addbs (This.DiskFileName (Justpath (lcFileName))) + lcXXX
		Endcase

		Return
		
	EndProc


	*----------------------------------------------------------------------------------
	Procedure EditSourceX(tcFileName, tcClass, tnStartRange, tnEndRange, tcMethod, tnRecno)

		Local lcClass, lcCursor, lcExt, lcFileName, lcMethod, lnEndRange, lnStartRange, loException, loTools, lnSuccess

		lcExt = Upper(Justext(tcFileName))

		*** From JRN 11/21/2011 : Use EditSourceX from IDE Tools, if available,
		*** which provides for source control and maintains MRU lists
		If Type('_Screen.cThorDispatcher') = 'C' and (lcExt <> 'DBF')
			loTools = Execscript(_Screen.cThorDispatcher, "Class= tools from pemeditor")
			If Not IsNull(loTools)
				loTools.EditSourceX(tcFileName, tcClass, tcMethod, tnStartRange, tnEndRange)
				Return
			Endif
		EndIf

		*-- If Thor Tool above was not available, then we will handle it locally...
		*-- Note: This means we will not get CheckOutSCC support, since that requires Thor/Peme from above.
		lcFileName 		= This.DiskFileName(Trim(tcFileName))
		lcClass	   		= Trim(Evl(tcClass, ''))
		lnStartRange	= Evl(tnStartRange, 1)
		lnEndRange		= Evl(tnEndRange, 1)
		lcMethod 		= Evl(tcMethod, '')

		This.AddMRUFile(lcFileName, lcClass)

		Try
			Do Case
				Case lcExt = 'PJX'
					Modify Project(lcFileName) Nowait

				Case lcExt = 'VCX' And Empty(lcClass)
					Do(_Browser) With (lcFileName)

				Case lcExt = 'DBC'
					Modify Database (lcFileName)

				Case lcExt $ ' FRX LBX MNX '
					Local lnDataSession
					lnDataSession = Set('Datasession')
					Set DataSession To 1
					lnSuccess = EditSource(lcFileName)
					IF lnSuccess > 0
						*** Error handling for wrong object reference call or already/still used files.
						MESSAGEBOX("There was an error opening the file:" + CHR(13) + ;
									lcFileName + CHR(13) + CHR(13) + ;
									"Error-Code: " + ALLTRIM(STR(lnSuccess)) + CHR(13) + ;
									ICASE(INLIST(lnSuccess,132,705), "File in use. Cannot be opened.", ;
											lnSuccess=200, "File not opened due to invalid object reference. Verify the presence of cMethodName in the object referenced by the cClassName parameter.", ;
											lnSuccess=901, "File opened but invalid object reference in cMethodName. Check the reference in the cMethodName parameter. Use a reference such as MyForm.MyList.CLICK.", ;
											"<Unknown Error Code>"))
					ENDIF 
					Set DataSession To &lnDataSession

				Case lcExt $ ' VCX SCX '
					lnSuccess = Editsource(lcFileName, lnStartRange, lcClass, lcMethod)
					IF lnSuccess > 0
						*** Error handling for wrong object reference call or already/still used files.
						MESSAGEBOX("There was an error opening the file:" + CHR(13) + ;
									lcFileName + CHR(13) + CHR(13) + ;
									"Error-Code: " + ALLTRIM(STR(lnSuccess)) + CHR(13) + ;
									ICASE(INLIST(lnSuccess,132,705), "File in use. Cannot be opened.", ;
											lnSuccess=200, "File not opened due to invalid object reference. Verify the presence of cMethodName in the object referenced by the cClassName parameter.", ;
											lnSuccess=925, "File opened but invalid object reference in cMethodName. Check the reference in the cMethodName parameter. Use a reference such as MyForm.MyList.CLICK.", ;
											"<Unknown Error Code>") + CHR(13) + CHR(13) + ;
									"lnStartRange: " + ALLTRIM(STR(lnStartRange)) + CHR(13) + ;
									"lcClass : " + lcClass + CHR(13) + ;
									"lcMethod: " + lcMethod)
					ENDIF 

				Case lcExt = 'PRG'
					Modify Command(lcFileName) Range lnStartRange, lnEndRange Nowait

				Case lcExt $ ' MPR QPR TXT H INI '
					Modify File(lcFileName) Range lnStartRange, lnEndRange Nowait

				Case lcExt = 'DBF'
					This.BrowseTable(lcFileName, tnRecno)

				Otherwise
					This.OpenURL(lcFileName) && Throw it off to the OS to handle opening the file.

			Endcase

		Catch To loException

			This.ShowErrorMsg(loException)

		Endtry
	EndProc


	*----------------------------------------------------------------------------------
	Procedure GetControlCount(loObject)

		Local lnCount

		With loObject
			Do Case
				Case Not This.GetPEMStatus(loObject, 'Objects', 5)
					lnCount = 0
				Case This.GetPEMStatus(loObject, 'ControlCount', 5)
					lnCount = .ControlCount
				Case Inlist(Lower(.BaseClass), [pageframe])
					lnCount = .PageCount
				Case Inlist(Lower(.BaseClass), [grid])
					lnCount = .ColumnCount
				Case Inlist(Lower(.BaseClass), [optiongroup], [commandgroup])
					lnCount = .ButtonCount
				Case Inlist(Lower(.BaseClass), [formset])
					lnCount = .FormCount
				Case Inlist(Lower(.BaseClass), [dataenvironment])
					lnCount = 0
					Do While 'O' = Type(".Objects(lnCount + 1)")
						lnCount = lnCount + 1
					Enddo
				Otherwise
					lnCount = 0
			Endcase
		Endwith

		Return lnCount
		
	EndProc


	*----------------------------------------------------------------------------------
	Procedure GetMRUID(lcFileName)

		Local lcExt, lcList, lcMRU_ID, lnPos
		lcExt = Upper (Justext ('.' + lcFileName))

		lcList = ',VCX=MRUI,PRG=MRUB,MPR=MRUB,QPR=MRUB,SCX=MRUH,MNX=MRUE,FRX=MRUG,DBF=MRUS,DBC=???,LBX=???,PJX=MRUL'
		lnPos  = At (',' + lcExt + '=', lcList)
		If lnPos = 0
			lcMRU_ID = 'MRUC'
		Else
			lcMRU_ID = Substr (lcList, lnPos + 5, 4)
		Endif

		Return lcMRU_ID
		
	EndProc


	*----------------------------------------------------------------------------------
	Procedure GetMRUList(lcMRU_ID)

		#Define DELIMITERCHAR Chr(0)

		Local loCollection As 'Collection'
		Local laItems(1), lcData, lcList, lcSourceFile, lnI, lnPos, lnSelect

		loCollection = Createobject ('Collection')

		If 'ON' # Set ('Resource')
			Return loCollection
		Endif

		lnSelect	 = Select()
		lcSourceFile = Set ('Resource', 1)
		Select 0
		Use (lcSourceFile) Again Shared Alias MRU_Resource

		If lcMRU_ID # 'MRU'
			lcMRU_ID = This.GetMRUID (lcMRU_ID)
			If '?' $ lcMRU_ID
				Return
			Endif
		Endif

		Locate For Id = lcMRU_ID
		If Found()
			lcData = Data
			Alines (laItems, Substr (lcData, 3), 0, DELIMITERCHAR)
			For lnI = 1 To Alen (laItems)
				lcItem = laItems(lnI)
				If Not Empty(lcItem) and ":" $ lcItem
					loCollection.Add (lcItem)
				Endif
			Endfor
		Endif

		Use
		Select (lnSelect)
		Return loCollection
		
	EndProc


	*----------------------------------------------------------------------------------
	Procedure GetPEMStatus(loObject, lcPEM, nAttribute)

		If Upper(loObject.BaseClass) = 'OLE'

			Local laMembers(1), lnRow
			Amembers(laMembers, loObject, 1, 'PHG#')
			lnRow = Ascan(laMembers, lcPEM, -1, -1, 1, 15)

			Do Case
				Case lnRow = 0
					Return .F.
				Case nAttribute = 0 && changed
					Return 'C' $ laMembers(lnRow, 3)
				Case nAttribute = 1 && readonly
					Return 'R' $ laMembers(lnRow, 3)
				Case nAttribute = 2 && protected
					Return 'P' $ laMembers(lnRow, 3)
				Case nAttribute = 3 && type
					Return laMembers(lnRow, 2)
				Case nAttribute = 4 && user-defined
					Return 'U' $ laMembers(lnRow, 3)
				Case nAttribute = 5 && defined
					Return .T.
				Case nAttribute = 6 && inherited
					Return 'I' $ laMembers(lnRow, 3)
			Endcase

		Else

			Return Pemstatus (loObject, lcPEM, nAttribute)

		EndIf
		
	EndProc


	*----------------------------------------------------------------------------------
	Procedure GetRelativePath(lcName, lcPath)
	
		Local lcNew, lnPos
		If Empty (lcPath)
			lcNew = Sys(2014, lcName)
		Else
			lcNew = Sys(2014, lcName, lcPath)
		Endif

		If Len (lcNew) < Len (lcName)
			lnPos = Rat ('..\', lcNew)
			If lnPos # 0
				lnPos = lnPos + 2
			Endif
			Return Left (lcNew, lnPos) + Right (lcName, Len (lcNew) - lnPos)
		Else
			Return lcName
		EndIf
		
	EndProc


	*----------------------------------------------------------------------------------
	Procedure IsNameChar(tcChar)

		Return Isalpha (tcChar) Or Isdigit (tcChar) Or tcChar = '_'
		
	EndProc


	*----------------------------------------------------------------------------------
	******************
	***    Author: Rick Strahl
	***            (c) West Wind Technologies, 1996
	***   Contact: rstrahl@west-wind.com
	***  Modified: 03/14/96
	***  Function: Starts associated Web Browser
	***            and goes to the specified URL.
	***            If Browser is already open it
	***            reloads the page.
	***    Assume: Works only on Win95 and NT 4.0
	***      Pass: tcUrl  - The URL of the site or
	***                     HTML page to bring up
	***                     in the Browser
	***    Return: 2  - Bad Association (invalid URL)
	***            31 - No application association
	***            29 - Failure to load application
	***            30 - Application is busy 
	***
	***            Values over 32 indicate success
	***            and return an instance handle for
	***            the application started (the browser) 
	****************************************************
	Procedure OpenURL(tcUrl, tcAction, tcDirectory, tcParms)

		If Empty(tcUrl)
			Return - 1
		Endif
		If Empty(tcAction)
			tcAction = "OPEN"
		Endif
		If Empty(tcDirectory)
			tcDirectory = Sys(2023)
		Endif

		Declare Integer ShellExecute ;
			IN SHELL32.Dll ;
			INTEGER nWinHandle, ;
			STRING cOperation, ;
			STRING cFileName, ;
			STRING cParameters, ;
			STRING cDirectory, ;
			INTEGER nShowWindow
		If Empty(tcParms)
			tcParms = ""
		Endif

		Declare Integer FindWindow ;
			IN WIN32API ;
			STRING cNull, String cWinName

		Return ShellExecute(0, ;
			tcAction, tcUrl, ;
			tcParms, tcDirectory, 1)
			
	EndProc


	*----------------------------------------------------------------------------------
	Procedure Release
	
		Local laMembers(1), lcMember
		
		Amembers (laMembers, This, 0)
		
		For Each lcMember In laMembers
			lcMember = Upper (lcMember)
			If lcMember = 'O' And Pemstatus(This, lcMember, 4) And 'O' = Vartype (Getpem (This, lcMember))
				If This.lReleaseOnDestroy
					Try
						This.&lcMember..Release()
					Catch
					Endtry
				Endif
				This.&lcMember. = .Null.
			Endif
		EndFor
		
	EndProc


	*----------------------------------------------------------------------------------
	Procedure ShowErrorMsg(loException, lcTitleBar, lcPRGName, lcAddlInfo)

		Messagebox ('Error: ' + Transform (loException.ErrorNo) 	+ ccCRLF +							;
			  		'Message: ' + loException.Message 					+ ccCRLF +							;
			  		'Procedure: ' + IIf (Empty (lcPRGName), loException.Procedure, Justfname (lcPRGName)) + ccCRLF + ;
			  		'Line: ' + Transform (loException.Lineno) 			+ ccCRLF +							;
			  		'Code: ' + loException.LineContents														;
			  		+ IIf (Empty (lcAddlInfo), '', ccCRLF + 'NOTES: ' + lcAddlInfo)							;
			  		, MB_OK + MB_ICONEXCLAMATION, Evl (lcTitleBar, 'Error'))
	EndProc


	*----------------------------------------------------------------------------------
	Procedure ShowHelp(lnHelpID)

		Local lcCurrentHelpFile, lcHelpFile, lcPath

		lcCurrentHelpFile = Set("Help", 1)

		lcPath = This.cApplicationPath
		lcHelpFile = lcPath + "PemEditor.CHM"

		*** JRN 2010-05-06 : Remove security warning; from http://www.foxpert.com/knowlbits_200906_1.htm
		If File(lcHelpFile)
			*PEME_RemSecurityWarning(lcHelpFile)

			Set Help To (lcHelpFile)

			If Empty (lnHelpID)
				Help
			Else
				Help Id (lnHelpID)
			Endif
		Endif

		Set Help To (lcCurrentHelpFile)
		
	EndProc


	*----------------------------------------------------------------------------------
	Procedure StripTabs(tcAbstract)

		* Abstract:
		*   Replace all tabs with spaces; also removes leading / trailing blanks
		*
		* Parameters:
		*	<cAbstract> = string to strip tabs/spaces from
		Return Alltrim (Chrtran (m.tcAbstract, ccTAB, ' '))
		
	EndProc

EndDefine

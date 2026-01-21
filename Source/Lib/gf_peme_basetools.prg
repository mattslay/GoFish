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
	Procedure BrowseTable(tcFileName, tnRecno, tlOnlyNormalBrowse)
	
		Local lcCursor, llSuperBrowsed, lnDataSession, lnShiftX, lnShiftY, loException, loForm, loTempForm
	
		*-- Create a temp form so we can control the Browse window size and location. 
		*-- Don't want it to fit hte same size and location as the GoFish Window, becuase
		*-- it looks weird when that happens.
		loTempForm = .Null.
	
		llSuperBrowsed = .F.
	
		lnDataSession = Set('DataSession')
		Set DataSession To 1
		Select 0
	
		lcCursor = Chrtran(Juststem(m.tcFileName), ' ', '_') + Sys(2015)
		Try
			Use (m.tcFileName) Again Alias &lcCursor
		Catch To m.loException
		Endtry
	
		If Used(m.lcCursor)
	
			If Not Empty(m.tnRecno)
				Goto m.tnRecno
			Endif
	
			*-- Call SuperBrowse, if Thor is present
			If Type('_Screen.cThorDispatcher') = 'C' And m.tlOnlyNormalBrowse = .F.
				llSuperBrowsed = Not Isnull(Execscript(_Screen.cThorDispatcher, 'Thor_Proc_SuperBrowse', m.lcCursor))
			Endif
	
			*-- Regular browse if Thor not present, or SuperBrowse call failed
			If Not m.llSuperBrowsed
				*!* ******** JRN Removed 2024-04-14 ********
				*!* If Vartype(_Screen.ActiveForm) = 'O'
				*!* 	lnShiftX   = 50
				*!* 	lnShiftY   = 125
				*!* 	If Type('_Screen._GoFish.oResultsForm') = 'O' And Not Isnull(_Screen._GoFish.oResultsForm)
				*!* 		loForm			  = _Screen._GoFish.oResultsForm
				*!* 	Else
				*!* 		loForm			  = _Screen.ActiveForm
				*!* 	Endif

				*!* 	loTempForm = Createobject('Form')
				*!* 	Try
				*!* 		loTempForm.Left	  = m.loForm.Left   + m.lnShiftX
				*!* 		loTempForm.Width  = m.loForm.Width  - m.lnShiftX
				*!* 		loTempForm.Top	  = m.loForm.Top    + m.lnShiftY
				*!* 		loTempForm.Height = m.loForm.Height - m.lnShiftY
				*!* 	Catch To m.loException

				*!* 	EndTry

				*!* 	m.loTempForm.Show()
				*!* Endif

				Browse Last Nowait Title (m.tcFileName) Normal
			Endif
	
			If Vartype(m.loTempForm) = 'O'
				m.loTempForm.Release()
			Endif
		Endif && Used(m.lcCursor)
	
		Set DataSession To &lnDataSession
	
	Endproc
			

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
	
		Local lcClass, lcExt, lcFileName, lcMethod, llEdited, lnDataSession, lnEndRange, lnStartRange
		Local lnSuccess, loException, loPE, loTools
	
		Set Message to 
		
		lcExt = Upper(Justext(m.tcFileName))
	
		*** From JRN 11/21/2011 : Use EditSourceX from IDE Tools, if available,
		*** which provides for source control and maintains MRU lists
		If Type('_Screen.cThorDispatcher') = 'C' And (m.lcExt # 'DBF')
			loTools = Execscript(_Screen.cThorDispatcher, 'Class= tools from pemeditor')
			If Not Isnull(m.loTools)
				m.loTools.EditSourceX(m.tcFileName, m.tcClass, m.tcMethod, m.tnStartRange, m.tnEndRange)
				Set Message to 
				Return
			Endif
		Endif
	
		*-- If Thor Tool above was not available, then we will handle it locally...
		*-- Note: This means we will not get CheckOutSCC support, since that requires Thor/Peme from above.
		lcFileName	 = This.DiskFileName(Trim(m.tcFileName))
		lcClass		 = Trim(Evl(m.tcClass, ''))
		lnStartRange = Evl(m.tnStartRange, 1)
		lnEndRange	 = Evl(m.tnEndRange, 1)
		lcMethod	 = Evl(m.tcMethod, '')
	
		This.AddMRUFile(m.lcFileName, m.lcClass)
	
		Try
			Do Case
				Case m.lcExt = 'PJX'
					Modify Project(m.lcFileName) Nowait
	
				Case m.lcExt = 'VCX' And Empty(m.lcClass)
					Do(_Browser) With (m.lcFileName)
	
				Case m.lcExt = 'DBC'
					Modify Database (m.lcFileName)
	
				Case m.lcExt $ ' FRX LBX MNX '
					lnDataSession = Set('Datasession')
					Set DataSession To 1
	
					llEdited  = .F.
					lnSuccess = 0 && oddly, zero on success, else error code
					*** JRN 2024-02-29 : Per Doug, use Project Explorer to open file if it's around
					If Type('_screen.oProjectExplorers') = 'O' And _Screen.oProjectExplorers.Count > 0
						loPE	 = _Screen.oProjectExplorers[1]
						llEdited = m.loPE.EditFile(m.lcFileName)
					Endif
					If Not m.llEdited
						lnSuccess = Editsource(m.lcFileName)
					EndIf
					
					If m.lnSuccess > 0
						*** Error handling for wrong object reference call or already/still used files.
						Messagebox('There was an error opening the file:' + Chr(13) +				;
							  m.lcFileName + Chr(13) + Chr(13) +									;
							  'Error-Code: ' + Alltrim(Str(m.lnSuccess)) + Chr(13) +				;
							  Icase(Inlist(m.lnSuccess, 132, 705), 'File in use. Cannot be opened.', ;
								m.lnSuccess = 200, 'File not opened due to invalid object reference. Verify the presence of cMethodName in the object referenced by the cClassName parameter.', ;
								m.lnSuccess = 901, 'File opened but invalid object reference in cMethodName. Check the reference in the cMethodName parameter. Use a reference such as MyForm.MyList.CLICK.', ;
								'<Unknown Error Code>'))
					Endif
					Set DataSession To &lnDataSession
	
				Case m.lcExt $ ' VCX SCX '

					llEdited  = .F.
					lnSuccess = 0 && oddly, zero on success, else error code					
					*** JRN 2024-02-29 : Per Doug, use Project Explorer to open file if it's around
					If Type('_screen.oProjectExplorers') = 'O' And _Screen.oProjectExplorers.Count > 0
						loPE	 = _Screen.oProjectExplorers[1]
						llEdited = m.loPE.EditFile(m.lcFileName, m.lnStartRange, m.lcClass, m.lcMethod)
					Endif
					If Not m.llEdited
						lnSuccess = Editsource(m.lcFileName, m.lnStartRange, m.lcClass, m.lcMethod)
					Endif

					If m.lnSuccess > 0
						*** Error handling for wrong object reference call or already/still used files.
						Messagebox('There was an error opening the file:' + Chr(13) +				;
							  m.lcFileName + Chr(13) + Chr(13) +									;
							  'Error-Code: ' + Alltrim(Str(m.lnSuccess)) + Chr(13) +				;
							  Icase(Inlist(m.lnSuccess, 132, 705), 'File in use. Cannot be opened.', ;
								m.lnSuccess = 200, 'File not opened due to invalid object reference. Verify the presence of cMethodName in the object referenced by the cClassName parameter.', ;
								m.lnSuccess = 925, 'File opened but invalid object reference in cMethodName. Check the reference in the cMethodName parameter. Use a reference such as MyForm.MyList.CLICK.', ;
								'<Unknown Error Code>') + Chr(13) + Chr(13) +						;
							  'lnStartRange: ' + Alltrim(Str(m.lnStartRange)) + Chr(13) +			;
							  'lcClass : ' + m.lcClass + Chr(13) +									;
							  'lcMethod: ' + m.lcMethod)
					Endif
	
				Case m.lcExt $ 'PRG SPR '
					Modify Command(m.lcFileName) Range m.lnStartRange, m.lnEndRange Nowait
	
				Case m.lcExt $ ' MPR QPR TXT H INI '
					Modify File(m.lcFileName) Range m.lnStartRange, m.lnEndRange Nowait
	
				Case m.lcExt = 'DBF'
					This.BrowseTable(m.lcFileName, m.tnRecno)
	
				Otherwise
					This.OpenURL(m.lcFileName) && Throw it off to the OS to handle opening the file.
	
			Endcase
	
		Catch To m.loException
	
			This.ShowErrorMsg(m.loException)
	
		Endtry
	Endproc
	

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

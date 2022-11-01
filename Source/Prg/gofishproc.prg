* Change log:
* 2021-03-19	See RemoveFolder() function.
*=======================================================================


*---------------------------------------------------------------------------
Define Class GF_FileResult as Custom

	Process = .f.
	FileName = ''
	FilePath = ''
	_Name = ''
	_Class = ''
	_Baseclass = ''
	_ParentClass = ''
	Classloc = ''
	FileType = ''
	MatchType = ''
	MatchLine = ''
	TrimmedMatchLine = ''
	Recno = 0
	IsText = .f.
	Column = ''
	TimeStamp =  {// :: AM}
	ContainingClass = ''
		
EndDefine

*---------------------------------------------------------------------------
Define Class GF_SearchResult as Custom

	Type = ''
	MethodName = ''
	*ContainingClass = ''
	MatchLine = ''
	TrimmedMatchLine = ''
	MatchType = ''
	ProcStart = 0
	ProcEnd = 0
	ProcCode = ''
	Statement = ''
	StatementStart = 0
	MatchStart = 0
	MatchLen = 0
	Code = '' && Stores the entire method code that was passed in for searching
	UserField = .null.
	oProcedure = .null.
	oMatch = .null.
	
	*---------------------------------------------------------------------------------------
	Procedure Init
		This.oProcedure = CreateObject('GF_Procedure')
	EndProc
		
EndDefine

*---------------------------------------------------------------------------
Define Class GF_Procedure As Custom

	Type		 = ''
	StartByte	 = 0
	EndByte		 = 0
	_Name		 = ''
	_ClassName	 = ''
	_ParentClass = ''
	_BaseClass	 = ''

Enddefine

*---------------------------------------------------------------------------
Define Class GF_SearchResultsFilter as Custom

	FileName = ''
	FilePath = ''
	ObjectName = ''
	ParentName = ''
	
	MatchType_Baseclass = .f.
	MatchType_ClassDef = .f.
	MatchType_ContainingClass = .f.
	MatchType_ParentClass = .f.
	MatchType_Filename = .f.
	MatchType_Function= .f.
	MatchType_Method = .f.
	MatchType_MethodDef = .f.
	MatchType_MethodDesc = .f.
	MatchType_Class = .f.
	MatchType_Name = .f.
	MatchType_Parent = .f.
	MatchType_Procedure = .f.
	MatchType_PropertyDef = .f.
	MatchType_PropertyName = .f.
	MatchType_PropertyValue = .f.
	MatchType_PropertyDesc = .f.
	MatchType_Code = .f.
	MatchType_Constant = .f.
	MatchType_Comment = .f.

	MatchType_FileDate = .f.
	MatchType_TimeStamp = .f.
	
	*-- Reports --------------------
	MatchType_Expr = .f.
	MatchType_SupExpr = .f.
	MatchType_Name = .f.
	MatchType_Tag = .f.
	MatchType_Tag2 = .f.
	MatchType_Picture = .f.
	
	FileType_SCX = .f.
	FileType_VCX = .f.
	FileType_FRX = .f.
	FileType_LBX = .f.
	FileType_MNX = .f.
	FileType_PJX = .f.
	FileType_DBC = .f.
	FileType_PRG = .f.
	FileType_MPR = .f.
	FileType_SPR = .f.
	FileType_INI = .f.
	FileType_H = .f.
	FileType_HTML = .f.
	FileType_XML = .f.
	FileType_TXT = .f.
	FileType_ASP = .f.
	FileType_JAVA = .f.
	FileType_JSP = .f.

	*Type_Procedure = .f.
	*Type_Method = .f.
	*Type_Class = .f.
	*Type_Blank = .f.
	
	Filename_Filter = ''
	FilePath_Filter = ''
	BaseClass_Filter = ''
	ParentClass_Filter = ''
	MethodName_Filter = ''
	Name_Filter = ''
	Class_Filter = ''
	Baseclass_Filter = ''
	MatchLine_Filter = ''
	Statement_Filter = ''
	ProcCode_Filter = ''
	ContainingClass_Filter = ''
	

	TextFiles = .f.
	TableFiles = .t.
	
	Timestamp_FilterFrom = {//}
	Timestamp_FilterTo = {//}
	Timestamp_Filter = .f.	
	
	FilterNot = .f.
	FilterLike = .f.
	FilterExactMatch = .f.
	
	OnlyFirstMatchInStatement = .F.
	OnlyFirstMatchInProcedure = .F.

	*---------------------------------------------------------------------------------------
	Procedure LoadFromFile(tcFile)

		Local loMy as 'My' OF 'My.vcx'
		Local laProperties[1], lcProperty

		If !File(tcFile)
			Return
		EndIf

		loMy = Newobject('My', 'My.vcx')
		AMembers(laProperties, This, 0, 'U')
		loMy.Settings.Load(tcFile)

		With loMy.Settings
		 
		 For x = 1 to Alen(laProperties)
		 	lcProperty = laProperties[x]
			Try
			 	Store Evaluate('.' + lcProperty) to ('This.' + lcProperty)
			Catch
			EndTry
		 Endfor

		Endwith

	Endproc
	

EndDefine


* --------------------------------------------------------------------------------
Procedure OpenExplorerWindow(lcPath)

	Do Case
		Case IsThorThere()
			Execscript(_Screen.cThorDispatcher, 'Thor_Proc_OpenExplorer', m.lcPath)
		Case File(lcPath)
			lcPath = '/select, ' + Fullpath(m.lcPath)
			Run /N "explorer" &lcPath
		Otherwise 
			Run /N "explorer" &lcPath
	Endcase

Endproc


* --------------------------------------------------------------------------------
Procedure Shell

	Lparameters lcFileURL

	Local oShell As 'wscript.shell'
	oShell = Createobject ('wscript.shell')
	m.oShell.Run (m.lcFileURL)

Endproc


* --------------------------------------------------------------------------------
Procedure IsThorThere

	Return Type('_Screen.cThorDispatcher') = 'C'

Endproc


*=======================================================================================
Function PropNvl (toObject, tcProperty, tuDefaultValue, tlAddPropertyIfNotPresent)

	If PemStatus(toObject, Alltrim(tcProperty), 5)
		If (Type('toObject.' + Alltrim(tcProperty)) <> Vartype(tuDefaultValue)) and Pcount() >= 3
			Return tuDefaultValue
		Else
			Return Evaluate('toObject.' + Alltrim(tcProperty))
		Endif
	Else
		If tlAddPropertyIfNotPresent
			AddProperty(toObject, tcProperty, tuDefaultValue)
		EndIf
		Return tuDefaultValue
	EndIf

EndFunc


* --------------------------------------------------------------------------------
*-- This method is used by the Delete button on the Search History Form.
*-- Revised 2021-03-19: 
*--   As as safety measure, it will only delete a folder path which contains "gf_saved_search_results".
* --------------------------------------------------------------------------------
Procedure RemoveFolder(lcFolderName)
	Local laFiles[1], lcFileName, lcFileNameWithPath, llFailure, lnFileCount, lnI, loException

	Declare Integer SetFileAttributes In kernel32 String, Integer
	lcFolderName = Trim(m.lcFolderName)
	
	If Empty(lcFolderName) Or !("gf_saved_search_results" $ Lower(lcFolderName))
		Return .f.
	Endif

	Try
		lnFileCount = Adir(laFiles, m.lcFolderName + '\*', 'DH')
		For lnI = 1 To m.lnFileCount
			lcFileName = m.laFiles[m.lnI, 1]
			If Left(m.lcFileName, 1) # '.'
				lcFileNameWithPath = m.lcFolderName + '\' + m.lcFileName
				SetFileAttributes(m.lcFileNameWithPath, 0)
				If 'D' $ m.laFiles[m.lnI, 5] && directory?
					RemoveFolder(m.lcFileNameWithPath)
				Else
					Delete File(m.lcFileNameWithPath)
				Endif
			Endif
		Endfor
		Rmdir(m.lcFolderName)

	Catch To m.loException
		llFailure = .T.
	Endtry

	Return m.llFailure = .F.

Endproc




* --------------------------------------------------------------------------------
* --------------------------------------------------------------------------------
*** JRN 10/14/2015 : Added to process context menus
Procedure CreateContextMenu(lcMenuName)
	Local loPosition

	loPosition = CalculateShortcutMenuPosition()

	*** JRN 2010-11-10 : Following is an attempt to solve the problem
	* when there is another form already open; apparently, if the 
	* focus is on the screen, the positioning of the popup still works OK
	_Screen.Show()

	Define Popup (m.lcMenuName)			;
		shortcut						;
		Relative						;
		From m.loPosition.Row, m.loPosition.Column

Endproc


Procedure CalculateShortcutMenuPosition

	Local lcPOINT, lnSMCol, lnSMRow, loResult, loWas

	Declare Long GetCursorPos In WIN32API String @lpPoint
	Declare Long ScreenToClient In WIN32API Long HWnd, String @lpPoint

	lcPOINT = Replicate (Chr(0), 8)
	&& Get mouse location in Windows desktop coordinates (pixels)
	= GetCursorPos (@m.lcPOINT)
	&& Convert to VFP Desktop (_Screen) coordinates
	= ScreenToClient (_Screen.HWnd, @m.lcPOINT)
	&& Covert the coordinates to foxels

	lnSMCol	= Pix2Fox (Long2Num (Left (m.lcPOINT, 4)), .F., _Screen.FontName, _Screen.FontSize)
	lnSMRow	= Pix2Fox (Long2Num (Right (m.lcPOINT, 4)), .T., _Screen.FontName, _Screen.FontSize)

	loResult = Createobject ('Empty')
	AddProperty (m.loResult, 'Column', m.lnSMCol )
	AddProperty (m.loResult, 'Row', m.lnSMRow )
	Return m.loResult

Endproc


Procedure Pix2Fox

	Lparameter tnPixels, tlVertical, tcFontName, tnFontSize
	&& tnPixels - pixels to convert
	&& tlVertical - .F./.T. convert horizontal/vertical coordinates
	&& tcFontName, tnFontSize - use specified font/size 
	&&         or current form (active output window) font/size, if not specified 
	Local lnFoxels

	If Pcount() > 2
		lnFoxels = m.tnPixels / Fontmetric(Iif(m.tlVertical, 1, 6), m.tcFontName, m.tnFontSize)
	Else
		lnFoxels = m.tnPixels / Fontmetric(Iif(m.tlVertical, 1, 6))
	Endif

	Return m.lnFoxels
Endproc


Function Long2Num(tcLong)
	Local lnNum
	lnNum = 0
	= RtlS2PL(@m.lnNum, m.tcLong, 4)
	Return m.lnNum
Endfunc


Function RtlS2PL(tnDest, tcSrc, tnLen)

	Declare RtlMoveMemory In WIN32API As RtlS2PL Long @Dest, String Source, Long Length
	Return 	RtlS2PL(@m.tnDest, @m.tcSrc, m.tnLen)

Endfunc


Define Class CreateExports As Session


	Procedure Init(lnDataSession)
		Set DataSession To (Evl(m.lnDataSession, 1))
	Endproc


	Procedure ExportToCursor(lcSourceFile, lcCursorName)
		Local lcDestFile
		Use (m.lcSourceFile) In 0
		lcDestFile = This.GetCursorName(m.lcCursorName)
		Select * From (Juststem(m.lcSourceFile)) Into Cursor (Juststem(m.lcDestFile)) Readwrite
		Use In (Juststem(m.lcSourceFile))
		Erase (Forceext(m.lcSourceFile, '*')) 
		Return lcDestFile
	Endproc


	Procedure GetCursorName(lcCursorName)
		Local lcDestFile, lnSuffix
		lnSuffix = 0
		Do While .T.
			lcDestFile = m.lcCursorName + Iif(m.lnSuffix = 0, '', Transform(m.lnSuffix))
			If Used(m.lcDestFile)
				lnSuffix = m.lnSuffix + 1
			Else
				Return m.lcDestFile
			Endif
		Enddo
	Endproc


	Procedure ExportToExcel(lcAlias, lcDestFile)

	Local loExcel As 'Excel.Application'
	Local lcClipText, lnI, lnRecords

	loExcel = Createobject('Excel.Application')

	With m.loExcel
		.Application.Visible	   = .F.
		.Application.DisplayAlerts = .F. && for now, no alerts
		.WorkBooks.Add()

		* keep only the first worksheet
		For lnI = 2 To .WorkSheets.Count
			.WorkSheets(2).Delete()
		Next m.lnI

		lcClipText = _Cliptext
		lnRecords  = _vfp.DataToClip(m.lcAlias, Reccount(m.lcAlias), 3)
		.Range('A1').Select()
		.ActiveSheet.Paste()
		_Cliptext = m.lcClipText

		.Range('A1').Select()
		.Application.Visible = .T.

		.Application.DisplayAlerts = .T.
		.ActiveWorkBook.SaveAs(m.lcDestFile)

	Endwith

	This.ForceForegroundWindow(m.loExcel.HWnd)


	Procedure ForceForegroundWindow
		Lparameters lnHWND

	    LOCAL nForeThread, nAppThread
	    
	    DECLARE Long BringWindowToTop In Win32API Long

	    DECLARE Long ShowWindow In Win32API Long, Long

	    DECLARE INTEGER GetCurrentThreadId; 
	        IN kernel32
	     
	    DECLARE INTEGER GetWindowThreadProcessId IN user32; 
	        INTEGER   hWnd,; 
	        INTEGER @ lpdwProcId  
	     
	    DECLARE INTEGER GetCurrentThreadId; 
	        IN kernel32  
	        
	    DECLARE INTEGER AttachThreadInput IN user32 ;
	        INTEGER idAttach, ;
	        INTEGER idAttachTo, ;
	        INTEGER fAttach

	    DECLARE INTEGER GetForegroundWindow IN user32  
	 
	    nForeThread = GetWindowThreadProcessId(GetForegroundWindow(), 0)
	    nAppThread = GetCurrentThreadId()

	    IF nForeThread != nAppThread
	        AttachThreadInput(nForeThread, nAppThread, .T.)
	        BringWindowToTop(lnHWND)
	        ShowWindow(lnHWND,3)
	        AttachThreadInput(nForeThread, nAppThread, .F.)
	    ELSE
	        BringWindowToTop(lnHWND)
	        ShowWindow(lnHWND,3)
	    ENDIF
	    

	EndProc 
	
	
Enddefine

* Change log:
* 2021-03-19	See RemoveFolder() function.
*=======================================================================


*---------------------------------------------------------------------------
Define Class GF_FileResult As Custom

	Process = .F.
	FileName = ''
	FilePath = ''
	_Name = ''
	_Class = ''
	_Baseclass = ''
	_ParentClass = ''
	Classloc = ''
	FileType = ''
	MatchType = ''
	MatchLine = ''
	TrimmedMatchLine = ''
	Recno = 0
	IsText = .F.
	Column = ''
	Timestamp =  {// :: AM}
	ContainingClass = ''

Enddefine

*---------------------------------------------------------------------------
Define Class GF_SearchResult As Custom

	Type = ''
	MethodName = ''
*ContainingClass = ''
	MatchLine = ''
	TrimmedMatchLine = ''
	MatchType = ''
	ProcStart = 0
	ProcEnd = 0
	ProcCode = ''
	Statement = ''
	StatementStart = 0
	MatchStart = 0
	MatchLen = 0
	Code = '' && Stores the entire method code that was passed in for searching
	UserField = .Null.
	oProcedure = .Null.
	oMatch = .Null.

*---------------------------------------------------------------------------------------
	Procedure Init
		This.oProcedure = Createobject('GF_Procedure')
	Endproc

Enddefine

*---------------------------------------------------------------------------
Define Class GF_Procedure As Custom

	Type		 = ''
	StartByte	 = 0
	EndByte		 = 0
	_Name		 = ''
	_ClassName	 = ''
	_ParentClass = ''
	_Baseclass	 = ''

Enddefine

*---------------------------------------------------------------------------
Define Class GF_SearchResultsFilter As Custom

	FileName = ''
	FilePath = ''
	ObjectName = ''
	ParentName = ''

	MatchType_Baseclass = .F.
	MatchType_ClassDef = .F.
	MatchType_ContainingClass = .F.
	MatchType_ParentClass = .F.
	MatchType_Filename = .F.
	MatchType_Function= .F.
	MatchType_Method = .F.
	MatchType_MethodDef = .F.
	MatchType_MethodDesc = .F.
	MatchType_Class = .F.
	MatchType_Name = .F.
	MatchType_Parent = .F.
	MatchType_Procedure = .F.
	MatchType_PropertyDef = .F.
	MatchType_PropertyName = .F.
	MatchType_PropertyValue = .F.
	MatchType_PropertyDesc = .F.
	MatchType_Code = .F.
	MatchType_Constant = .F.
	MatchType_Comment = .F.

	MatchType_FileDate = .F.
	MatchType_TimeStamp = .F.

*-- Reports --------------------
	MatchType_Expr = .F.
	MatchType_SupExpr = .F.
	MatchType_Name = .F.
	MatchType_Tag = .F.
	MatchType_Tag2 = .F.
	MatchType_Picture = .F.

	FileType_SCX = .F.
	FileType_VCX = .F.
	FileType_FRX = .F.
	FileType_LBX = .F.
	FileType_MNX = .F.
	FileType_PJX = .F.
	FileType_DBC = .F.
	FileType_PRG = .F.
	FileType_MPR = .F.
	FileType_SPR = .F.
	FileType_INI = .F.
	FileType_H = .F.
	FileType_HTML = .F.
	FileType_XML = .F.
	FileType_TXT = .F.
	FileType_ASP = .F.
	FileType_JAVA = .F.
	FileType_JSP = .F.

*Type_Procedure = .f.
*Type_Method = .f.
*Type_Class = .f.
*Type_Blank = .f.

	Filename_Filter = ''
	FilePath_Filter = ''
	BaseClass_Filter = ''
	ParentClass_Filter = ''
	MethodName_Filter = ''
	Name_Filter = ''
	Class_Filter = ''
	BaseClass_Filter = ''
	MatchLine_Filter = ''
	Statement_Filter = ''
	ProcCode_Filter = ''
	ContainingClass_Filter = ''


	TextFiles = .F.
	TableFiles = .T.

	Timestamp_FilterFrom = {//}
	Timestamp_FilterTo = {//}
	Timestamp_Filter = .F.

	FilterNot = .F.
	FilterLike = .F.
	FilterExactMatch = .F.

	OnlyFirstMatchInStatement = .F.
	OnlyFirstMatchInProcedure = .F.

*---------------------------------------------------------------------------------------
	Procedure LoadFromFile(tcFile)

		Local loMy As 'My' Of 'My.vcx'
		Local laProperties[1], lcProperty

		If !File(tcFile)
			Return
		Endif

		loMy = Newobject('My', 'My.vcx')
		Amembers(laProperties, This, 0, 'U')
		loMy.Settings.Load(tcFile)

		With loMy.Settings

			For x = 1 To Alen(laProperties)
				lcProperty = laProperties[x]
				Try
						Store Evaluate('.' + lcProperty) To ('This.' + lcProperty)
					Catch
				Endtry
			Endfor

		Endwith

	Endproc


Enddefine


* --------------------------------------------------------------------------------
Procedure OpenExplorerWindow(lcPath)

	Do Case
		Case IsThorThere()
			Execscript(_Screen.cThorDispatcher, 'Thor_Proc_OpenExplorer', m.lcPath)
		Case File(lcPath)
			lcPath = '/select, ' + Fullpath(m.lcPath)
			Run /N "explorer" &lcPath
		Otherwise
			Run /N "explorer" &lcPath
	Endcase

Endproc


* --------------------------------------------------------------------------------
Procedure Shell

	Lparameters lcFileURL

	Local oShell As 'wscript.shell'
	oShell = Createobject ('wscript.shell')
	m.oShell.Run (m.lcFileURL)

Endproc


* --------------------------------------------------------------------------------
Procedure IsThorThere

	Return Type('_Screen.cThorDispatcher') = 'C'

Endproc


*=======================================================================================
Function PropNvl (toObject, tcProperty, tuDefaultValue, tlAddPropertyIfNotPresent)

	If Pemstatus(toObject, Alltrim(tcProperty), 5)
		If (Type('toObject.' + Alltrim(tcProperty)) <> Vartype(tuDefaultValue)) And Pcount() >= 3
			Return tuDefaultValue
		Else
			Return Evaluate('toObject.' + Alltrim(tcProperty))
		Endif
	Else
		If tlAddPropertyIfNotPresent
			AddProperty(toObject, tcProperty, tuDefaultValue)
		Endif
		Return tuDefaultValue
	Endif

Endfunc


* --------------------------------------------------------------------------------
*-- This method is used by the Delete button on the Search History Form.
*-- Revised 2021-03-19:
*--   As as safety measure, it will only delete a folder path which contains "gf_saved_search_results".
* --------------------------------------------------------------------------------
Procedure RemoveFolder(lcFolderName)
	Local laFiles[1], lcFileName, lcFileNameWithPath, llFailure, lnFileCount, lnI, loException

	Declare Integer SetFileAttributes In kernel32 String, Integer
	lcFolderName = Trim(m.lcFolderName)

	If Empty(lcFolderName) Or !("gf_saved_search_results" $ Lower(lcFolderName))
		Return .F.
	Endif

	Try
			lnFileCount = Adir(laFiles, m.lcFolderName + '\*', 'DH')
			For lnI = 1 To m.lnFileCount
				lcFileName = m.laFiles[m.lnI, 1]
				If Left(m.lcFileName, 1) # '.'
					lcFileNameWithPath = m.lcFolderName + '\' + m.lcFileName
					SetFileAttributes(m.lcFileNameWithPath, 0)
					If 'D' $ m.laFiles[m.lnI, 5] && directory?
						RemoveFolder(m.lcFileNameWithPath)
					Else
						Delete File(m.lcFileNameWithPath)
					Endif
				Endif
			Endfor
			Rmdir(m.lcFolderName)

		Catch To m.loException
			llFailure = .T.
	Endtry

	Return m.llFailure = .F.

Endproc




* --------------------------------------------------------------------------------
* --------------------------------------------------------------------------------
*** JRN 10/14/2015 : Added to process context menus
Procedure CreateContextMenu(lcMenuName)
	Local loPosition

	loPosition = CalculateShortcutMenuPosition()

*** JRN 2010-11-10 : Following is an attempt to solve the problem
* when there is another form already open; apparently, if the
* focus is on the screen, the positioning of the popup still works OK
	_Screen.Show()

	Define Popup (m.lcMenuName)			;
		shortcut						;
		Relative						;
		From m.loPosition.Row, m.loPosition.Column

Endproc


Procedure CalculateShortcutMenuPosition

	Local lcPOINT, lnSMCol, lnSMRow, loResult, loWas

	Declare Long GetCursorPos In WIN32API String @lpPoint
	Declare Long ScreenToClient In WIN32API Long HWnd, String @lpPoint

	lcPOINT = Replicate (Chr(0), 8)
&& Get mouse location in Windows desktop coordinates (pixels)
	= GetCursorPos (@m.lcPOINT)
&& Convert to VFP Desktop (_Screen) coordinates
	= ScreenToClient (_Screen.HWnd, @m.lcPOINT)
&& Covert the coordinates to foxels

	lnSMCol	= Pix2Fox (Long2Num (Left (m.lcPOINT, 4)), .F., _Screen.FontName, _Screen.FontSize)
	lnSMRow	= Pix2Fox (Long2Num (Right (m.lcPOINT, 4)), .T., _Screen.FontName, _Screen.FontSize)

	loResult = Createobject ('Empty')
	AddProperty (m.loResult, 'Column', m.lnSMCol )
	AddProperty (m.loResult, 'Row', m.lnSMRow )
	Return m.loResult

Endproc


Procedure Pix2Fox

	Lparameter tnPixels, tlVertical, tcFontName, tnFontSize
&& tnPixels - pixels to convert
&& tlVertical - .F./.T. convert horizontal/vertical coordinates
&& tcFontName, tnFontSize - use specified font/size
&&         or current form (active output window) font/size, if not specified
	Local lnFoxels

	If Pcount() > 2
		lnFoxels = m.tnPixels / Fontmetric(Iif(m.tlVertical, 1, 6), m.tcFontName, m.tnFontSize)
	Else
		lnFoxels = m.tnPixels / Fontmetric(Iif(m.tlVertical, 1, 6))
	Endif

	Return m.lnFoxels
Endproc


Function Long2Num(tcLong)
	Local lnNum
	lnNum = 0
	= RtlS2PL(@m.lnNum, m.tcLong, 4)
	Return m.lnNum
Endfunc


Function RtlS2PL(tnDest, tcSrc, tnLen)

	Declare RtlMoveMemory In WIN32API As RtlS2PL Long @Dest, String Source, Long Length
	Return 	RtlS2PL(@m.tnDest, @m.tcSrc, m.tnLen)

Endfunc


Define Class CreateExports As Session


	Procedure Init(lnDataSession)
		Set DataSession To (Evl(m.lnDataSession, 1))
	Endproc


	Procedure ExportToCursor(lcSourceFile, lcCursorName)
		Local lcDestFile
		Use (m.lcSourceFile) In 0
		lcDestFile = This.GetCursorName(m.lcCursorName)
		Select * From (Juststem(m.lcSourceFile)) Into Cursor (Juststem(m.lcDestFile)) Readwrite
		Use In (Juststem(m.lcSourceFile))
		Erase (Forceext(m.lcSourceFile, '*'))
		Return lcDestFile
	Endproc


	Procedure GetCursorName(lcCursorName)
		Local lcDestFile, lnSuffix
		lnSuffix = 0
		Do While .T.
			lcDestFile = m.lcCursorName + Iif(m.lnSuffix = 0, '', Transform(m.lnSuffix))
			If Used(m.lcDestFile)
				lnSuffix = m.lnSuffix + 1
			Else
				Return m.lcDestFile
			Endif
		Enddo
	Endproc


	Procedure ExportToExcel(lcAlias, lcDestFile)

		Local loExcel As 'Excel.Application'
		Local lcClipText, lnI, lnRecords

		loExcel = Createobject('Excel.Application')

		With m.loExcel
			.Application.Visible	   = .F.
			.Application.DisplayAlerts = .F. && for now, no alerts
			.WorkBooks.Add()

* keep only the first worksheet
			For lnI = 2 To .WorkSheets.Count
				.WorkSheets(2).Delete()
			Next m.lnI

			lcClipText = _Cliptext
			lnRecords  = _vfp.DataToClip(m.lcAlias, Reccount(m.lcAlias), 3)
			.Range('A1').Select()
			.ActiveSheet.Paste()
			_Cliptext = m.lcClipText

			.Range('A1').Select()
			.Application.Visible = .T.

			.Application.DisplayAlerts = .T.
			.ActiveWorkBook.SaveAs(m.lcDestFile)

		Endwith

		This.ForceForegroundWindow(m.loExcel.HWnd)


	Procedure ForceForegroundWindow
		Lparameters lnHWND

		Local nForeThread, nAppThread

		Declare Long BringWindowToTop In Win32API Long

		Declare Long ShowWindow In Win32API Long, Long

		Declare Integer GetCurrentThreadId;
			IN kernel32

		Declare Integer GetWindowThreadProcessId In user32;
			INTEGER   HWnd,;
			INTEGER @ lpdwProcId

		Declare Integer GetCurrentThreadId;
			IN kernel32

		Declare Integer AttachThreadInput In user32 ;
			INTEGER idAttach, ;
			INTEGER idAttachTo, ;
			INTEGER fAttach

		Declare Integer GetForegroundWindow In user32

		nForeThread = GetWindowThreadProcessId(GetForegroundWindow(), 0)
		nAppThread = GetCurrentThreadId()

		If nForeThread != nAppThread
			AttachThreadInput(nForeThread, nAppThread, .T.)
			BringWindowToTop(lnHWND)
			ShowWindow(lnHWND,3)
			AttachThreadInput(nForeThread, nAppThread, .F.)
		Else
			BringWindowToTop(lnHWND)
			ShowWindow(lnHWND,3)
		Endif


	Endproc


Enddefine

Procedure Get_LocalSettings	&& Determine storage place global / local
	Lparameters;
		toSettings,;
		tcSettingsFile

	Local;
		lcFolder     As String,;
		lcSourceFile As String,;
		llFound      As Boolean

*only if resource file is oon and used
	If m.toSettings.Exists('lCR_AllowEd') And m.toSettings.lCR_Allow Then
		toSettings.lCR_AllowEd = Set("Resource")=="ON" And File(Set("Resource",1))

	Else  &&m.toSettings.Exists('lCR_AllowEd') And m.toSettings.lCR_Allow
		toSettings.Add('lCR_AllowEd',Set("Resource")=="ON" And File(Set("Resource",1)))

	Endif &&m.toSettings.Exists('lCR_AllowEd') And m.toSettings.lCR_Allow

	If m.toSettings.lCR_AllowEd And m.toSettings.lCR_Local Then
*get location for GoFish from ResourceFile
		llFound = Get_LocalPath(@lcFolder)
		llFound = Create_LocalPath(@lcFolder,m.llFound,toSettings.lCR_Local_Default,toSettings.lCR_Local_Default)

		If Empty(m.lcFolder) Then
*No folder selected, we keep normal GoFish mode of operation
			toSettings.lCR_AllowEd = .F.

		Endif &&EMPTY(m.lcFolder)

		If !Isnull(m.llFound) Then
* if changed
			Put_LocalPath(m.lcFolder)
		Endif &&!Isnull(m.llFound)

	Endif &&m.toSettings.lCR_AllowEd AND m.toSettings.lCR_Local

*still allowed to use local
	If m.toSettings.lCR_AllowEd And m.toSettings.lCR_Local Then
*see if we can locate local settings file
		tcSettingsFile = Addbs(m.lcFolder)+Justfname(m.tcSettingsFile)
		llFound        = File(m.tcSettingsFile)

		If m.llFound Then
*Get settings
			toSettings.Load(m.tcSettingsFile)

		Endif &&m.llFound
	Endif &&m.toSettings.lCR_AllowEd AND m.toSettings.lCR_Local
Endproc &&Get_LocalSettings

Procedure Get_LocalPath	&&Get the local storage path from resource file
	Lparameters;
		tcFolder

	Local;
		lcSourceFile As String,;
		lnSelected   As Integer,;
		llFound      As Boolean

	lnSelected = Select()

	lcSourceFile = Set ("Resource", 1)
	Use (lcSourceFile) Again Shared Alias ResourceAlias In Select('ResourceAlias')
	Select ResourceAlias
	Locate For Type='GoFish  ' And Id="DirLoc  "
	llFound = Found()

	If m.llFound Then
		tcFolder = Data
	Endif &&m.llFound

	Use In Select('ResourceAlias')
	Select (m.lnSelected)

	Return m.llFound

Endproc &&Get_LocalPath

Procedure Put_LocalPath	&&Put the local storage path to resource file
	Lparameters;
		tcFolder

	Local;
		lcSourceFile As String,;
		lnSelected   As Integer,;
		llFound      As Boolean

	lnSelected = Select()

	lcSourceFile = Set ("Resource", 1)
	Use (lcSourceFile) Again Shared Alias ResourceAlias In Select('ResourceAlias')
	Select ResourceAlias
	Locate For Type='GoFish  ' And Id="DirLoc  "
	llFound = Found()

	If m.llFound Then
*found, needs update
		Replace;
			CkVal   With Val (Sys(2007, m.tcFolder)),;
			Data    With m.tcFolder,;
			Updated With Date()

	Else  &&m.llFound
*not found, needs insert
		Insert Into ResourceAlias					;
			(Type, Id, CkVal, Data, Updated)		;
			Values									;
			('GoFish', "DirLoc", Val (Sys(2007, m.tcFolder)), m.tcFolder, Date())

	Endif &&m.llFound

	Use In Select('ResourceAlias')
	Select (m.lnSelected)

Endproc &&Put_LocalPath

Procedure Create_LocalPath	&&Create the local storage path, with user interface if bot given
	Lparameters;
		tcFolder,;
		tlFound,;
		tlDefault,;
		tlCreate

	Local;
		lcOldDir As String,;
		llFill   As Boolean,;
		lnFile   As Number,;
		lnFiles  As Number

	Local Array;
		laDir(1)

	lcOldDir = Fullpath("", "")

	If Empty(m.tcFolder) Then
		tcFolder = Justpath(Set("Resource",1))
		If m.tlDefault Then
			tcFolder = m.tcFolder+"\GoFish_"

		Endif &&m.tlDefault
	Endif &&Empty(m.tcFolder)

	Do Case
		Case Directory(m.tcFolder)
*folder found, do we need to set resource?
			tlFound = Iif(m.tlFound,.Null.,.F.)
			llFill  = Empty(Adir(laDir, Addbs(m.tcFolder) + "GF_*.xml", "", 1))

		Case m.tlCreate
*just create
			Mkdir (m.tcFolder)
			tlFound = .F.
			llFill  = .T.

		Otherwise
*Folder not found
			tcFolder = Getdir(m.tcFolder,'Local GoFish config folder not found. Please pick one.','',64+1+32+2+8)
			If Empty(m.tcFolder) Or !Directory(m.tcFolder) THEN 
*No folder selected, we keep normal GoFish mode of operation
				tlFound = NIL
			Else  &&Empty(m.tcFolder) Or !Directory(m.tcFolder)
				llFill  = Empty(Adir(laDir, Addbs(m.tcFolder) + "GF_*.xml", "", 1))
			Endif &&Empty(m.tcFolder) Or !Directory(m.tcFolder)

	Endcase

	If !Empty(m.tcFolder) Then
*create .gitignore, if neeeded
		If !File(Addbs(m.tcFolder)+'.gitignore') Then
			Strtofile('*.*'+0h0D0A,Addbs(m.tcFolder)+'.gitignore')
		Endif &&!FILE(ADDBS(m.tcFolder)+'.gitignore')

		If m.llFill Then
			lnFiles = Adir(laDir, Addbs(Home(7)) + "GF_*.xml", "", 1)
			For lnFile = 1 To m.lnFiles
				Copy File (Addbs(Home(7)) + laDir(m.lnFile,1)) To (Addbs(m.tcFolder) + laDir(m.lnFile,1))
			Endfor &&lnFile
		Endif &&m.llFill
	Endif &&!EMPTY(m.tcFolder)

	Cd (m.lcOldDir)

	Return m.tlFound

Endproc &&Create_LocalPath

Procedure Change_TableStruct	&&Update structur of storage tables from version pre 6.*.*
	Lparameters;
		toResultForm,;
		tcRoot,;
		tcSavedSearchResults

	Local;
		lcDatabase As String,;
		lcDBF      As String,;
		lcDbc      As String,;
		lcDir      As String,;
		lcOldDir   As String,;
		lnSelect   As Integer,;
		lnReturn   As Integer,;
		lnResult   As Integer,;
		lnResults  As Integer

	lcDBF = m.toResultForm.cUISettingsFile

	If !File(m.lcDBF) Then
		Return 0
	Endif &&!File(m.lcDBF)

	lcDbc = m.tcRoot + m.toResultForm.cSaveDBC
	If File(m.lcDbc) Then
		Return 0
	Endif &&File(m.lcDbc)

	lcDBF = Forceext(m.lcDBF, "DBF")
	lnReturn = 1

	lcOldDir   = Fullpath("","")
	lnSelect   = Select()
	lcDatabase = Set("Database")

	Cd (m.tcRoot)
	Select 0
	Assert .F.
*- Create the DBC
	Try
			lcDir = Addbs(m.tcRoot) + m.tcSavedSearchResults
			If !Directory(m.lcDir) Then
				Mkdir (m.lcDir)
			Endif &&!Directory(m.lcDir)
			lcDir = Addbs(m.lcDir)

			If !File(Addbs(m.lcDir) + 'README.md') Then
				Strtofile(Get_Readme_Text(2),Addbs(m.lcDir) + 'README.md')
			Endif &&!FILE(ADDBS(m.lcDir) + 'README.md')

			Create Database (m.lcDbc)

*- Add existing tables to DBC
			Add Table GF_Search_Expression_History.Dbf
			Use GF_Search_Expression_History Exclusive
			Pack
			Index On Item Tag _Item
			Use

			Add Table GF_Search_Scope_History.Dbf
			Use GF_Search_Scope_History Exclusive
			Pack
			Index On Item Tag _Item
			Use

*- Add search history table to DBC
			= m.toResultForm.oSearchEngine.ClearResultsCursor()

*-- Create the table to save the search results
			Select (m.toResultForm.oSearchEngine.cSearchResultsAlias)
			Copy To (m.lcDBF) Database (m.lcDbc)
			Use (m.lcDBF) In Select(Juststem(m.lcDBF))

*-- Create search history mother table
			lnResults = m.toResultForm.BuildSearchHistoryCursor(.T., .T.)

			Assert .F.
			If m.lnResults > 0 Then
				Messagebox("Updating search history structure." + 0h0D0A + "Please be patient.", 0, "GoFish", 5000)
				?"Total history jobs",m.lnResults
				?""

				Select GF_SearchHistory
				lnResult = 0

				Scan
***
					lnResult = m.lnResult+1
					??''+0h0d+' '+0h0d+"Processing ",Justfname(Justpath(SearchHistoryFolder))," No.",m.lnResult
					lcUni    = cUni
					lcDate   = Datetime

					lcFilePath = Trim(SearchHistoryFolder)
					llReturn   = m.toResultForm.LoadSavedResults(m.lcFilePath, .T.)

					lcAlias2   = Alias()
					lnReccount = Reccount(m.lcAlias2)
					If m.lnReccount>0 Then
						??", item count:",m.lnReccount
						?"/"
						lcMakro = ""
						If Empty(Field("cUni")) Then
							lcMakro = m.lcMakro + "gf_SearchHistory.cUni AS cUni," +;
								"SPACE(23) AS cUni_File,"
						Endif &&Empty(Field("cUni"))
						If Empty(Field("Datetime")) Then
							lcMakro = m.lcMakro + "gf_SearchHistory.Datetime AS Datetime,"
						Endif &&Empty(Field("Datetime"))
						If Empty(Field("Search")) Then
							lcMakro = m.lcMakro + "gf_SearchHistory.Search_Expression AS Search,"
						Endif &&Empty(Field("Search"))
						If Empty(Field("Scope")) Then
							lcMakro = m.lcMakro + "gf_SearchHistory.Scope AS Scope,"
						Endif &&Empty(Field("Scope"))
						If !Empty(m.lcMakro) Then
							Select;
								&lcMakro;
								Cur1.*;
								From (m.lcAlias2) As Cur1;
								Into Cursor (m.lcAlias2) NoFilter Readwrite

						Endif &&!Empty(m.lcMakro)

						If Datetime#m.lcDate Then
							Replace All;
								cUni      With m.lcUni,;
								cUni_File With m.lcUni + "_" + Sys(2007, Trim(Padl(Id, 11)), 0 ,1) + "_",;
								Datetime  With m.lcDate
							Go Top

						Endif &&Datetime#m.lcDate
*						m.toResultForm.ApplyFilter()
*						??''+0h0d+' '+0h0d+"/"
						llReturn = m.toResultForm.FillSearchResultsCursor(.T.) && Pulls records from the search engine's results cursor.
						If !m.llReturn Then
							??''+0h0d+' '+0h0d+"failed"
							?''
							Select GF_SearchHistory
							Delete
							Loop
						Endif &&!m.llReturn
						??''+0h0d+' '+0h0d+"-"

						lnCount = 0
						Scan
							Strtofile(Code    , m.lcDir + Trim(cUni_File)+"Code.txt")
							Strtofile(ProcCode, m.lcDir + Trim(cUni_File)+"ProcCode.txt")
							Replace;
								ProcCode With "",;
								Code     With ""
							Do Case
								Case m.lnCount%100=0
									??''+0h0d+' '+0h0d+"\"
								Case m.lnCount%100=25
									??''+0h0d+' '+0h0d+"|"
								Case m.lnCount%100=50
									??''+0h0d+' '+0h0d+"/"
								Case m.lnCount%100=75
									??''+0h0d+' '+0h0d+"-"
							Endcase
							lnCount = m.lnCount+1
						Endscan &&All

						Select Juststem(m.lcDBF)
						Append From Dbf(m.lcAlias2)
					Else  &&m.lnReccount>0
						??", nothing to do"
						?""

					Endif &&m.lnReccount>0
					Use In Select(m.lcAlias2)

					Delete File (m.lcFilePath + "\*.*")
					Rmdir (m.lcFilePath)

				Endscan &&All

				toResultForm.ProgressBar.Stop()

			Endif &&m.lnResults > 0

		Catch To loException
*			Set Step On
			lnReturn = 2

	Endtry

*-- Create the table to save the search results main info
	lcDBF = Addbs(m.tcRoot) + "GF_Search_History"
	Select GF_SearchHistory
	Copy To (m.lcDBF) Database (m.lcDbc)

	Alter Table GF_Search_History Drop Column SearchHistoryFolder

	Use In Select("GF_Search_History")
	Use In Select(m.toResultForm.oSearchEngine.cSearchResultsAlias)
	Use In GF_SearchHistory

	Close Database
	Set Database To (m.lcDatabase)
	Select(m.lnSelect)


	Cd (m.lcOldDir)

	Return m.lnReturn
Endproc &&Change_TableStruct

Procedure Get_Readme_Text  	&&Text for README.md Files
	Lparameters;
		tnFile

	Local;
		lcText As String

	Do Case
		Case m.tnFile=1
			TEXT To m.lcText Noshow
## GoFish settings folder
This folder contains the settings and history table for GoFish! code search tool.

- If GoFish is NOT working, try to delete the files AND subfolders as a whole.
- If the folder grows to large **Do not delete single files or tables**
  - delete *all files and folders*, but this will reset all options
  - delete history files through GoFish
    - Right click nodes, choose clear (depends on tree mode)
    - Delete searches via *History* button

GoFish is available in VFP9 SP2 or VFPA from Thor or from https://github.com/VFPX/GoFish

			ENDTEXT &&lcText

		Case m.tnFile=2
			TEXT To m.lcText Noshow
## GoFish history files folder
This folder contains the history files for GoFish! code search tool.
The files are not in a memo, so that the size might grow above 2GiB and faster access for the history is possible.

If the folder grows to large **Do not delete files**
- delete all GoFish files and folders, but this will reset all options
  - delete all files `..\GF_*.*`
  - delete all files in this folder
- delete history files through GoFish
  - Right click nodes, choose clear (depends on tree mode)
  - Delete searches via *History* button

GoFish is available in VFP9 SP2 or VFPA from Thor or from https://github.com/VFPX/GoFish

			ENDTEXT &&lcText
		Otherwise
			lcText = ''
	Endcase

	Return m.lcText

Endproc &&Get_Readme_Text

Procedure GetMonitorStatistics
************************************************************************
* GetMonitorStatistics
****************************************
***  Function: Returns information about the desktop screen
***            Can be used to check for desktop width and size
***            and determine when a second monitor is disabled
***            and layout needs to be adjusted to keep the
***            window visible.
***      Pass:
***    Return:  Monitor Object
***
*** From https://west-wind.com/wconnect/weblog/ShowEntry.blog?id=836
*** Handling Multiple Screens in Visual FoxPro Desktop Applications
***
************************************************************************

	#Define SM_CXFULLSCREEN		16
	#Define SM_CYFULLSCREEN		17

	#Define SM_XVIRTUALSCREEN	76
	#Define SM_YVIRTUALSCREEN	77
	#Define SM_CXVIRTUALSCREEN	78
	#Define SM_CYVIRTUALSCREEN	79
	#Define SM_CMONITORS		80

	#Define SM_XVIRTUALSCREEN	76
	#Define SM_YVIRTUALSCREEN	77
	#Define SM_CXVIRTUALSCREEN	78
	#Define SM_CYVIRTUALSCREEN	79


	Declare Integer GetSystemMetrics In user32 Integer nIndex

	loMonitor = Createobject("EMPTY")
*	AddProperty(loMonitor, "gnMonitors",      GetSystemMetrics(SM_CMONITORS))

	AddProperty(loMonitor, "gnVirtualLeft",   GetSystemMetrics(SM_XVIRTUALSCREEN))
	AddProperty(loMonitor, "gnVirtualTop",    GetSystemMetrics(SM_YVIRTUALSCREEN))

	AddProperty(loMonitor, "gnVirtualWidth",  GetSystemMetrics(SM_CXVIRTUALSCREEN))
	AddProperty(loMonitor, "gnVirtualHeight", GetSystemMetrics(SM_CYVIRTUALSCREEN))

	AddProperty(loMonitor, "gnVirtualRight",  m.loMonitor.gnVirtualWidth -  Abs(m.loMonitor.gnVirtualLeft) - 10)
	AddProperty(loMonitor, "gnVirtualBottom", m.loMonitor.gnVirtualHeight - Abs(m.loMonitor.gnVirtualTop)  -  5)

*ADDPROPERTY(loMonitor, "gnScreenHeight",  GetSystemMetrics(SM_CYFULLSCREEN))
*ADDPROPERTY(loMonitor, "gnScreenWidth",   GetSystemMetrics(SM_CXFULLSCREEN))

	Return m.loMonitor

Endproc &&GetMonitorStatistics

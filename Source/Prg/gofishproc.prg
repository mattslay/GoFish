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
Procedure RemoveFolder(lcFolderName)
	Local laFiles[1], lcFileName, lcFileNameWithPath, llFailure, lnFileCount, lnI, loException

	Declare Integer SetFileAttributes In kernel32 String, Integer
	lcFolderName = Trim(m.lcFolderName)

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




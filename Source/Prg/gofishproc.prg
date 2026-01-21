* Change log:
* 2021-03-19	See GF_RemoveFolder() function.
* see git history
*=======================================================================


*---------------------------------------------------------------------------
Define Class GF_FileResult As Custom

	Process          = .F.
	FileName         = ""
	FilePath         = ""
	_Name            = ""
	_Class           = ""
	_Baseclass       = ""
	_ParentClass     = ""
	Classloc         = ""
	FileType         = ""
	MatchType        = ""
	MatchLine        = ""
	TrimmedMatchLine = ""
	Recno            = 0
	IsText           = .F.
	Column           = ""
	Timestamp        =  {// :: AM}
	ContainingClass  = ""

Enddefine

*---------------------------------------------------------------------------
Define Class GF_SearchResult As Custom

	Type       = ""
	MethodName = ""
*ContainingClass = ""
	MatchLine        = ""
	TrimmedMatchLine = ""
	MatchType        = ""
	ProcStart        = 0
	ProcEnd          = 0
	ProcCode         = ""
	Statement        = ""
	StatementStart   = 0
	MatchStart       = 0
	MatchLen         = 0
	Code             = "" && Stores the entire method code that was passed in for searching
	UserField        = .Null.
	oProcedure       = .Null.
	oMatch           = .Null.

*---------------------------------------------------------------------------------------
	Procedure Init
		This.oProcedure = Createobject("GF_Procedure")
	Endproc

Enddefine

*---------------------------------------------------------------------------
Define Class GF_Procedure As Custom

	Type         = ""
	StartByte    = 0
	EndByte      = 0
	_Name        = ""
	_ClassName   = ""
	_ParentClass = ""
	_Baseclass   = ""

Enddefine

*---------------------------------------------------------------------------
Define Class GF_SearchResultsFilter As Custom

	FileName   = ""
	FilePath   = ""
	ObjectName = ""
	ParentName = ""

	MatchType_Baseclass       = .F.
	MatchType_ClassDef        = .F.
	MatchType_ContainingClass = .F.
	MatchType_ParentClass     = .F.
	MatchType_Filename        = .F.
	MatchType_Function        = .F.
	MatchType_Method          = .F.
	MatchType_MethodDef       = .F.
	MatchType_MethodDesc      = .F.
	MatchType_Class           = .F.
	MatchType_Name            = .F.
	MatchType_Parent          = .F.
	MatchType_Procedure       = .F.
	MatchType_PropertyDef     = .F.
	MatchType_PropertyName    = .F.
	MatchType_PropertyValue   = .F.
	MatchType_PropertyDesc    = .F.
	MatchType_Code            = .F.
	MatchType_Constant        = .F.
	MatchType_Comment         = .F.

	MatchType_FileDate  = .F.
	MatchType_TimeStamp = .F.

*-- Reports --------------------
	MatchType_Expr    = .F.
	MatchType_SupExpr = .F.
	MatchType_Name    = .F.
	MatchType_Tag     = .F.
	MatchType_Tag2    = .F.
	MatchType_Picture = .F.

	FileType_SCX  = .F.
	FileType_VCX  = .F.
	FileType_FRX  = .F.
	FileType_LBX  = .F.
	FileType_MNX  = .F.
	FileType_PJX  = .F.
	FileType_DBC  = .F.
	FileType_PRG  = .F.
	FileType_MPR  = .F.
	FileType_SPR  = .F.
	FileType_INI  = .F.
	FileType_H    = .F.
	FileType_HTML = .F.
	FileType_XML  = .F.
	FileType_TXT  = .F.
	FileType_ASP  = .F.
	FileType_JAVA = .F.
	FileType_JSP  = .F.

*Type_Procedure = .f.
*Type_Method = .f.
*Type_Class = .f.
*Type_Blank = .f.

	Filename_Filter        = ""
	FilePath_Filter        = ""
	BaseClass_Filter       = ""
	ParentClass_Filter     = ""
	MethodName_Filter      = ""
	Name_Filter            = ""
	Class_Filter           = ""
	BaseClass_Filter       = ""
	MatchLine_Filter       = ""
	Statement_Filter       = ""
	ProcCode_Filter        = ""
	ContainingClass_Filter = ""


	TextFiles  = .F.
	TableFiles = .T.

	Timestamp_FilterFrom = {//}
	Timestamp_FilterTo   = {//}
	Timestamp_Filter     = .F.

	FilterNot        = .F.
	FilterLike       = .F.
	FilterExactMatch = .F.

	OnlyFirstMatchInStatement = .F.
	OnlyFirstMatchInProcedure = .F.

*---------------------------------------------------------------------------------------
	Procedure LoadFromFile(tcFile)

		Local;
			lcProperty As String,;
			loMy    As "My" Of "My.vcx"

		Local Array;
			laProperties(1)

*:Global;
x

		If !File(m.tcFile)
			Return
		Endif

		loMy = Newobject("My", "My.vcx")
		Amembers(laProperties, This, 0, "U")
		loMy.Settings.Load(m.tcFile)

		With m.loMy.Settings

			For x = 1 To Alen(m.laProperties)
				lcProperty = laProperties[x]
				Try
						Store Evaluate("." + m.lcProperty) To ("This." + m.lcProperty)
					Catch
				Endtry
			Endfor

		Endwith

	Endproc


Enddefine

Define Class CreateExports As Session


	Procedure Init(lnDataSession)
		Set DataSession To (Evl(m.lnDataSession, 1))
	Endproc


	Procedure ExportToCursor(lcSourceFile, lcCursorName)
		Local;
			lcDestFile As String

		Use (m.lcSourceFile) In 0
		lcDestFile = This.GetCursorName(m.lcCursorName)
		Select *;
			From (Juststem(m.lcSourceFile));
			Into Cursor (Juststem(m.lcDestFile)) Readwrite
		Use In (Juststem(m.lcSourceFile))
		Erase (Forceext(m.lcSourceFile, "*"))
		Return m.lcDestFile
	Endproc


	Procedure GetCursorName(lcCursorName)
		Local;
			lcDestFile As String,;
			lnSuffix As Number

		lnSuffix = 0
		Do While .T.
			lcDestFile = m.lcCursorName + Iif(m.lnSuffix = 0, "", Transform(m.lnSuffix))
			If Used(m.lcDestFile)
				lnSuffix = m.lnSuffix + 1
			Else
				Return m.lcDestFile
			Endif
		Enddo
	Endproc


	Procedure ExportToExcel(lcAlias, lcDestFile)

		Local;
			lcClipText As String,;
			lnI     As Number,;
			lnRecords As Number,;
			loExcel As "Excel.Application"

		loExcel = Createobject("Excel.Application")

		With m.loExcel
			.Application.Visible       = .F.
			.Application.DisplayAlerts = .F. && for now, no alerts
			.WorkBooks.Add()

* keep only the first worksheet
			For lnI = 2 To .WorkSheets.Count
				.WorkSheets(2).Delete()
			Next m.lnI

			lcClipText = _Cliptext
			lnRecords  = _vfp.DataToClip(m.lcAlias, Reccount(m.lcAlias), 3)
			.Range("A1").Select()
			.ActiveSheet.Paste()
			_Cliptext = m.lcClipText

			.Range("A1").Select()
			.Application.Visible = .T.

			.Application.DisplayAlerts = .T.
			.ActiveWorkBook.SaveAs(m.lcDestFile)

		Endwith

		This.ForceForegroundWindow(m.loExcel.HWnd)


	Procedure ForceForegroundWindow
		Lparameters lnHWND

*:Global;
nAppThread,;
nForeThread

		Declare Long BringWindowToTop In Win32API Long

		Declare Long ShowWindow In Win32API Long, Long

		Declare Integer GetCurrentThreadId;
			In kernel32

		Declare Integer GetWindowThreadProcessId In user32;
			Integer   HWnd,;
			Integer @ lpdwProcId

		Declare Integer GetCurrentThreadId;
			In kernel32

		Declare Integer AttachThreadInput In user32 ;
			Integer idAttach, ;
			Integer idAttachTo, ;
			Integer fAttach

		Declare Integer GetForegroundWindow In user32

		nForeThread = GetWindowThreadProcessId(GetForegroundWindow(), 0)
		nAppThread  = GetCurrentThreadId()

		If nForeThread != nAppThread
			AttachThreadInput(nForeThread, nAppThread, .T.)
			BringWindowToTop(m.lnHWND)
			ShowWindow(m.lnHWND,3)
			AttachThreadInput(nForeThread, nAppThread, .F.)
		Else
			BringWindowToTop(m.lnHWND)
			ShowWindow(m.lnHWND,3)
		Endif


	Endproc


Enddefine

* --------------------------------------------------------------------------------
Procedure GF_OpenExplorerWindow(lcPath)

	Do Case
		Case GF_IsThorThere()
			Execscript(_Screen.cThorDispatcher, "Thor_Proc_OpenExplorer", m.lcPath)
		Case File(m.lcPath)
			lcPath = "/select, " + Fullpath(m.lcPath)
			Run /N "explorer" &lcPath
		Otherwise
			Run /N "explorer" &lcPath
	Endcase

Endproc


* --------------------------------------------------------------------------------
Procedure GF_Shell

	Lparameters lcFileURL

*:Global;
oShell As "wscript.shell"

	oShell = Createobject ("wscript.shell")
	oShell.Run (m.lcFileURL, 7, .F.)

Endproc


* --------------------------------------------------------------------------------
Procedure GF_IsThorThere

	Return Type("_Screen.cThorDispatcher") = "C"

Endproc


*=======================================================================================
Function GF_PropNvl (toObject, tcProperty, tuDefaultValue, tlAddPropertyIfNotPresent)

	If Pemstatus(m.toObject, Alltrim(m.tcProperty), 5)
		If (Type("toObject." + Alltrim(m.tcProperty)) <> Vartype(m.tuDefaultValue)) And Pcount() >= 3
			Return m.tuDefaultValue
		Else
			Return Evaluate("toObject." + Alltrim(m.tcProperty))
		Endif
	Else
		If m.tlAddPropertyIfNotPresent
			AddProperty(m.toObject, m.tcProperty, m.tuDefaultValue)
		Endif
		Return m.tuDefaultValue
	Endif

Endfunc


* --------------------------------------------------------------------------------
*-- This method is used by the Delete button on the Search History Form.
*-- Revised 2021-03-19:
*--   As as safety measure, it will only delete a folder path which contains "gf_saved_search_results".
*-- Revised 2023-03-04:
*--   Skip the safety measure, we can rename our folders as we like
*-- use something more sophisticated to remove the folder, check migrate
* --------------------------------------------------------------------------------
Procedure GF_RemoveFolder(tcFolderName,tlJustAttributes)
	Local;
		lcFileName      As String,;
		lcFileNameWithPath As String,;
		llFailure       As Boolean,;
		lnFileCount     As Number,;
		lnI             As Number,;
		loException     As Object,;
		loFSO           As "Scripting.FileSystemObject"

	Local Array;
		laFiles(1)

	tcFolderName = Trim(m.tcFolderName)

	If Empty(m.tcFolderName)
		Return .F.
	Endif

	If Right(m.tcFolderName,1) == "\"
		tcFolderName = Justpath(m.tcFolderName)
	Endif &&Right(m.tcFolderName,1) == "\"
	If !Directory(m.tcFolderName) Then
		Return .F.
	Endif &&!Directory(m.tcFolderName)

	If !m.tlJustAttributes Then
		loFSO = Createobject("Scripting.FileSystemObject")
		loFSO.DeleteFolder(m.tcFolderName, .T.)
	Endif &&!m.tlJustAttributes

	If Directory(m.tcFolderName) Then
		Declare Integer SetFileAttributes In kernel32 String, Integer
		Try
				lnFileCount = Adir(laFiles, m.tcFolderName + "\*", "DH")
				For lnI = 1 To m.lnFileCount
					lcFileName = laFiles[m.lnI, 1]
					If Left(m.lcFileName, 1) # "."
						lcFileNameWithPath = m.tcFolderName + "\" + m.lcFileName
						SetFileAttributes(m.lcFileNameWithPath, 0)
						If "D" $ laFiles[m.lnI, 5] && directory?
							GF_RemoveFolder(m.lcFileNameWithPath, .T.)
						Endif
					Endif
				Endfor
				If !m.tlJustAttributes Then
					loFSO = Createobject("Scripting.FileSystemObject")
					loFSO.DeleteFolder(m.tcFolderName, .T.)
				Endif &&!m.tlJustAttributes

			Catch To m.loException
				llFailure = .T.
		Endtry
	Endif &&Directory(m.tcFolderName)

	Return !m.llFailure

Endproc




* --------------------------------------------------------------------------------
* --------------------------------------------------------------------------------
*** JRN 10/14/2015 : Added to process context menus
Procedure GF_CreateContextMenu(lcMenuName, toResultsForm)
	Local loPosition As Object
	Local lcFont

	loPosition = GF_CalculateShortcutMenuPosition()

	lcFont = m.toResultsForm.GetContextMenuFont()

	Define Popup (m.lcMenuName)							;
		ShortCut										;
		Relative										;
		From m.loPosition.Row, m.loPosition.Column		;
		&lcFont											;
		In Screen

Endproc


Procedure GF_CalculateShortcutMenuPosition

	Local;
		lcPOINT As String,;
		lnSMCol As Number,;
		lnSMRow As Number,;
		loResult As Object,;
		loWas As Object

	Declare Long GetCursorPos In Win32API String @lpPoint
	Declare Long ScreenToClient In Win32API Long HWnd, String @lpPoint

	lcPOINT = Replicate (Chr(0), 8)
&& Get mouse location in Windows desktop coordinates (pixels)
	= GetCursorPos (@lcPOINT)
&& Convert to VFP Desktop (_Screen) coordinates
	= ScreenToClient (_Screen.HWnd, @lcPOINT)
&& Covert the coordinates to foxels

	lnSMCol = GF_Pix2Fox (GF_Long2Num (Left (m.lcPOINT, 4)), .F., _Screen.FontName, _Screen.FontSize)
	lnSMRow = GF_Pix2Fox (GF_Long2Num (Right (m.lcPOINT, 4)), .T., _Screen.FontName, _Screen.FontSize)

	loResult = Createobject ("Empty")
	AddProperty (m.loResult, "Column", m.lnSMCol )
	AddProperty (m.loResult, "Row", m.lnSMRow )
	Return m.loResult

Endproc


Procedure GF_Pix2Fox

	Lparameter tnPixels, tlVertical, tcFontName, tnFontSize
	Local;
		lnFoxels As Number

	If Pcount() > 2
		lnFoxels = m.tnPixels / Fontmetric(Iif(m.tlVertical, 1, 6), m.tcFontName, m.tnFontSize)
	Else
		lnFoxels = m.tnPixels / Fontmetric(Iif(m.tlVertical, 1, 6))
	Endif

	Return m.lnFoxels
Endproc


Function GF_Long2Num(tcLong)
	Local;
		lnNum As Number

	lnNum = 0
	= GF_RtlS2PL(@lnNum, m.tcLong, 4)
	Return m.lnNum
Endfunc


Function GF_RtlS2PL(tnDest, tcSrc, tnLen)

	Declare RtlMoveMemory In Win32API As GF_RtlS2PL Long @Dest, String Source, Long Length
	Return 	GF_RtlS2PL(@tnDest, @tcSrc, m.tnLen)

Endfunc

Procedure GF_Get_LocalSettings	&& Determine storage place global / local
	Lparameters;
		toSettings,;
		tcSettingsFile,;
		tlClear

	Local;
		lcFolder  As String,;
		lcSourceFile As String,;
		llFound   As Boolean

*only if resource file is on and used
	If toSettings.Exists("lCR_AllowEd") And m.toSettings.lCR_Allow Then
		toSettings.lCR_AllowEd = Set("Resource")=="ON" And File(Set("Resource",1))

	Else  &&toSettings.Exists("lCR_AllowEd") And m.toSettings.lCR_Allow
		toSettings.Add("lCR_AllowEd",Set("Resource")=="ON" And File(Set("Resource",1)))

	Endif &&toSettings.Exists("lCR_AllowEd") And m.toSettings.lCR_Allow

	If m.toSettings.lCR_AllowEd And m.toSettings.lCR_Local Then
*get location for GoFish from ResourceFile
		llFound = GF_Get_LocalPath(@lcFolder)
		llFound = GF_Create_LocalPath(@lcFolder,m.llFound,m.toSettings.lCR_Local_Default,m.toSettings.lCR_Local_Default,m.tlClear)

		If Empty(m.lcFolder) Then
*No folder selected, we keep normal GoFish mode of operation
			toSettings.lCR_AllowEd = .F.

		Endif &&EMPTY(m.lcFolder)

		If !Isnull(m.llFound) Then
* if changed
			GF_Put_LocalPath(m.lcFolder)
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
Endproc &&GF_Get_LocalSettings

Procedure GF_Get_LocalPath	&&Get the local storage path from resource file
	Lparameters;
		tcFolder

	Local;
		lcSourceFile As String,;
		llFound   As Boolean,;
		lnSelected As Integer

	lnSelected = Select()

	lcSourceFile = Set ("Resource", 1)
	Use (m.lcSourceFile) Again Shared Alias ResourceAlias In Select("ResourceAlias")
	Select ResourceAlias
	Locate For Type="GoFish  " And Id="DirLoc  "
	llFound = Found()

	If m.llFound Then
		tcFolder = Data
	Endif &&m.llFound

	Use In Select("ResourceAlias")
	Select (m.lnSelected)

	Return m.llFound

Endproc &&GF_Get_LocalPath

Procedure GF_Put_LocalPath	&&Put the local storage path to resource file
	Lparameters;
		tcFolder

	Local;
		lcSourceFile As String,;
		llFound   As Boolean,;
		lnSelected As Integer

	lnSelected = Select()

	lcSourceFile = Set ("Resource", 1)
	Use (m.lcSourceFile) Again Shared Alias ResourceAlias In Select("ResourceAlias")
	Select ResourceAlias
	Locate For Type="GoFish  " And Id="DirLoc  "
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
			("GoFish", "DirLoc", Val (Sys(2007, m.tcFolder)), m.tcFolder, Date())

	Endif &&m.llFound

	Use In Select("ResourceAlias")
	Select (m.lnSelected)

Endproc &&GF_Put_LocalPath

Procedure GF_Create_LocalPath	&&Create the local storage path, with user interface if bot given
	Lparameters;
		tcFolder,;
		tlFound,;
		tlDefault,;
		tlCreate,;
		tlClear

	Local;
		lcOldDir As String,;
		llFill As Boolean,;
		lnFile As Number,;
		lnFiles As Number

	Local Array;
		laDir(1)

	lcOldDir = Fullpath("", "")

	If Empty(m.tcFolder) Then
		tcFolder = Justpath(Set("Resource",1))
		If m.tlDefault Then
			tcFolder = m.tcFolder+"\GoFish_"

		Endif &&m.tlDefault
	Endif &&Empty(m.tcFolder)

	If Directory(m.tcFolder) And m.tlClear
*folder found, we need to clear
		GF_RemoveFolder(m.tcFolder)
		tlCreate = .T.
	Endif &&Directory(m.tcFolder) AND m.tlClear

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
			tcFolder = Getdir(m.tcFolder,"Local GoFish config folder not found. Please pick one.","",64+1+32+2+8)
			If Empty(m.tcFolder) Or !Directory(m.tcFolder) Then
*No folder selected, we keep normal GoFish mode of operation
				tlFound = NIL
			Else  &&Empty(m.tcFolder) Or !Directory(m.tcFolder)
				llFill  = Empty(Adir(laDir, Addbs(m.tcFolder) + "GF_*.xml", "", 1))
			Endif &&Empty(m.tcFolder) Or !Directory(m.tcFolder)

	Endcase

	If !Empty(m.tcFolder) Then
*create .gitignore, if neeeded
		If !File(Addbs(m.tcFolder)+".gitignore") Then
			Strtofile("#Set by GoFish."+0h0D0A+"*.*"+0h0D0A,Addbs(m.tcFolder)+".gitignore")
		Endif &&!FILE(ADDBS(m.tcFolder)+".gitignore")
*create .FoxBin2Prg_Ignore, if neeeded
		If !File(Addbs(m.tcFolder)+".FoxBin2Prg_Ignore") Then
			Strtofile("#Set by GoFish",Addbs(m.tcFolder)+".FoxBin2Prg_Ignore")
		Endif &&!FILE(ADDBS(m.tcFolder)+".FoxBin2Prg_Ignore")

		If m.llFill Then
			lnFiles = Adir(laDir, Addbs(Home(7)) + "GoFish_\GF_*.xml", "", 1)
			For lnFile = 1 To m.lnFiles
				Copy File (Addbs(Home(7) + "GoFish_") + laDir(m.lnFile,1)) To (Addbs(m.tcFolder) + laDir(m.lnFile,1))
			Endfor &&lnFile
			GF_Write_Readme_Text(1, Addbs(m.tcFolder) + "README.md", .T.)

		Endif &&m.llFill
	Endif &&!EMPTY(m.tcFolder)

	Cd (m.lcOldDir)

	Return m.tlFound

Endproc &&GF_Create_LocalPath

Procedure GF_Change_TableStruct	&&Update structure of storage tables from version pre version 6.*.*
	Lparameters;
		toResultForm,;
		tcRoot,;
		tcSavedSearchResults

	Local;
		lcAlias2 As String,;
		lcComment As String,;
		lcDBF    As String,;
		lcDBF_H  As String,;
		lcDatabase As String,;
		lcDate   As String,;
		lcDbc    As String,;
		lcDir    As String,;
		lcFilePath As String,;
		lcMakro  As String,;
		lcOldDir As String,;
		lcUni    As String,;
		llReturn As Boolean,;
		lnCompare As Integer,;
		lnCount  As Number,;
		lnReccount As Number,;
		lnResult As Integer,;
		lnResults As Integer,;
		lnReturn As Integer,;
		lnSelect As Integer,;
		lnVerNo  As Integer,;
		loException As Object

	lcDBF   = Addbs(m.tcRoot) + m.toResultForm.cUISettingsFile
	lcDBF_H = Addbs(m.tcRoot) + "GF_Search_History"

	If File(ForceExt(m.lcdbf_h, 'dbf')) Then
		Return 0
	Endif &&!File(m.lcDBF)

	Text To m.lcMsg Noshow Textmerge
Converting to Version 6/7 table structures.

Search and replace history from Version 5 have not been imported.

However, they are still available in GoFish 5 and GoFish Classic.
	Endtext
	MessageBox(lcmsg, 64, 'Table structure changes')

	lcDbc = m.tcRoot + m.toResultForm.cSaveDBC
	If File(m.lcDbc) Then
*Version control
        llReturn = .T.
		Open Database (m.lcDbc)
		lcComment = DBGetProp(Justfname(m.lcDbc),'DATABASE','COMMENT')
		lnCompare = Compare_VerNo(, m.lcComment, @lnVerNo)
		Do Case
			Case m.lnCompare=0
* Database fits to app
			Case m.lnCompare=2
* database newer then GoFish
				Close Databases
				Return 3
			Otherwise
*update
				llReturn = NewVersion(m.lnVerNo)

				If m.llReturn Then
					DBSetProp(Justfname(m.lcDbc),'DATABASE','COMMENT',_Screen._GoFish.cVersion)
				Endif &&m.llReturn
				*/update
		Endcase
		Close Databases

*just check that the folder exists
		lcDir = Addbs(m.tcRoot) + m.tcSavedSearchResults
		If !Directory(m.lcDir) Then
			Mkdir (m.lcDir)
			lcDir = Addbs(m.lcDir)

			GF_Write_Readme_Text(3, m.lcDir + "README.md")
		Endif &&!Directory(m.lcDir)

		Return IIF(m.llReturn, 0, 4)
	Endif &&File(m.lcDbc)

	lcDBF    = Forceext(m.lcDBF, "DBF")
	lnReturn = 1

	lcOldDir   = Fullpath("","")
	lnSelect   = Select()
	lcDatabase = Set("Database")

	Cd (m.tcRoot)
	Select 0

*- Create the DBC
	Try
*temp use of var ...
			lcDir    = "README.md"
			lnResult = 1
			If Upper(m.tcRoot)==Upper(Home(7) + "GoFish_\") Then
				lcDir    = "GF_" + m.lcDir
				lnResult = 2
			Endif &&Upper(m.tcRoot)==Upper(Home(7) + "GoFish_\")

			GF_Write_Readme_Text(m.lnResult, m.tcRoot + m.lcDir)

			lcDir = Addbs(m.tcRoot) + m.tcSavedSearchResults
			If !Directory(m.lcDir) Then
				Mkdir (m.lcDir)
			Endif &&!Directory(m.lcDir)
			lcDir = Addbs(m.lcDir)

			GF_Write_Readme_Text(3, m.lcDir + "README.md")

			Create Database (m.lcDbc)
			DBSetProp(Justfname(m.lcDbc),'DATABASE','COMMENT',_Screen._GoFish.cVersion)

*- Add existing tables to DBC
			If File("GF_Search_Expression_History.Dbf") Then
				Add Table GF_Search_Expression_History.Dbf
				Use GF_Search_Expression_History Exclusive
				Delete all
				Pack
				Index On Item Tag _Item
				Use
			Endif &&File("GF_Search_Expression_History.Dbf")

			If File("GF_Search_Scope_History.Dbf") Then
				Add Table GF_Search_Scope_History.Dbf
				Use GF_Search_Scope_History Exclusive
				Pack
				Index On Item Tag _Item
				Use
			Endif &&File("GF_Search_Scope_History.Dbf")


*- Add search history table to DBC
			= toResultForm.oSearchEngine.ClearResultsCursor()

*-- Create the table to save the search results
			Select (m.toResultForm.oSearchEngine.cSearchResultsAlias)
			Copy To (m.lcDBF) Database (m.lcDbc)
			Use (m.lcDBF) In Select(Juststem(m.lcDBF))

*-- Create search history mother table
			lnResults = toResultForm.BuildSearchHistoryCursor(.T., .T.)
			
			*** JRN 2024-02-26 : Act like we found NO records, since we are not converting history
			lnResults = 0

			If m.lnResults > 0 Then
				Messagebox("Updating search history structure." + 0h0D0A + "Please be patient.", 0, "GoFish", 5000)
				?"Total history jobs",m.lnResults
				?""

				Select GF_SearchHistory
				lnResult = 0
				Scan
***
					lnResult = m.lnResult+1
					??""+0h0d+" "+0h0d+"Processing ",Justfname(Justpath(SearchHistoryFolder))," No.",m.lnResult
					lcUni  = cUni
					lcDate = Datetime

					lcFilePath = Trim(SearchHistoryFolder)

					llReturn   = toResultForm.LoadSavedResults(m.lcFilePath, .T., , .T.)

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
							lcMakro = m.lcMakro + "gf_SearchHistory.Search AS Search,"
						Endif &&Empty(Field("Search"))
						If Empty(Field("Scope")) Then
							lcMakro = m.lcMakro + "gf_SearchHistory.Scope AS Scope,"
						Endif &&Empty(Field("Scope"))

						If !Empty(m.lcMakro) Then
							Select;
								&lcMakro;
								Cur1.*,;
								.F.          As lMemLoaded,;
								Cast(0 As I) As iReplaceFolder,;
								.F.          As lJustReplaced,;
								.T.          As lSaved;
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
*						??""+0h0d+" "+0h0d+"/"
						llReturn = toResultForm.FillSearchResultsCursor(.T.) && Pulls records from the search engine's results cursor.
						If !m.llReturn Then
							??""+0h0d+" "+0h0d+"failed"
							?""
							Select GF_SearchHistory
							Delete
							Loop
						Endif &&!m.llReturn
						??""+0h0d+" "+0h0d+"-"

						lnCount = 0
						Scan
							Strtofile(ProcCode, m.lcDir + Trim(cUni_File) + "ProcCode.txt")

							Strtofile(Code    , m.lcDir + Trim(cUni_File) + "Code.txt")
							Replace;
								ProcCode   With "",;
								Code       With "",;
								lMemLoaded With .F.
							Do Case
								Case m.lnCount%1000=0
									??""+0h0d+" "+0h0d+"\"
								Case m.lnCount%1000=250
									??""+0h0d+" "+0h0d+"|"
								Case m.lnCount%1000=500
									??""+0h0d+" "+0h0d+"/"
								Case m.lnCount%1000=750
									??""+0h0d+" "+0h0d+"-"
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

					Select GF_SearchHistory

					If File(m.lcFilePath + "GF_Saved_Search_Results.txt") Then
						Rename (m.lcFilePath + "GF_Saved_Search_Results.txt") To (m.lcDir + Trim(cUni)+"_Saved_Search_Results.txt"))
					Endif &&File(m.lcFilePath + "GF_Saved_Search_Results.txt")

					If File(m.lcFilePath + "GF_Search_Settings.xml") Then
						Rename (m.lcFilePath + "GF_Search_Settings.xml") To (m.lcDir + Trim(cUni)+"_Search_Settings.xml"))
					Endif &&File(m.lcFilePath + "GF_Search_Settings.xml")

					If File(m.lcFilePath + "GF_Results_Form_Settings.xml") Then
						Rename (m.lcFilePath + "GF_Results_Form_Settings.xml") To (m.lcDir + Trim(cUni)+"_Results_Form_Settings.xml"))
					Endif &&File(m.lcFilePath + "GF_Results_Form_Settings.xml")

					Delete File (m.lcFilePath + "\*.*")
					Rmdir (m.lcFilePath)

				Endscan &&All

				toResultForm.ProgressBar.Stop()

			Endif &&m.lnResults > 0

*-- Create the table to save the search results main info
			If Used("GF_SearchHistory") Then
				Select GF_SearchHistory
				*** JRN 2024-02-26 : Empty it ... we are not converting search history
				Copy To (m.lcDBF_H) Database (m.lcDbc) For .F.

				Alter Table GF_Search_History Drop Column SearchHistoryFolder
*		Alter Table GF_Search_History Add Column lSaved L
*		Alter Table GF_Search_History Add Column lReplace L

				Update Cur1 Set;
					lSaved = .T.;
					From GF_Search_History As Cur1

				Use In GF_SearchHistory
			Endif &&USED("GF_SearchHistory")

*-- Create the table to save the running ID of replace backup folder
			Create Table (Addbs(m.tcRoot) + "GF_ReplaceID");
				(iID I)

			Append Blank
* Todo: Get the ID of existing folder from GF_Replace_History.dbf
			Replace;
				iID With 1000

*Replace History

			*** JRN 2024-02-26 : We are not converting history
			If .F. ;
					And Directory(Addbs(m.tcRoot) + "GF_ReplaceBackups");
					And File(Addbs(m.tcRoot) + "GF_REPLACE_DETAILV5.DBC");
					And File(Addbs(m.tcRoot) + "GF_REPLACE_DETAILV5.DBF");
					And File(Addbs(m.tcRoot) + "GF_Replace_History.DBF") Then
* only if we found stored Backups
*GF_REPLACE_DETAILV5.DBC
*GF_Replace_DetailV5.dbf
*GF_Replace_History.dbf
				Open Database (Addbs(m.tcRoot) + "GF_REPLACE_DETAILV5.DBC")
				Select;
					"_" + Sys(2007, Ttoc(Date_time), 0, 1)                  As cUni,;
					"_" + Sys(2007, Ttoc(Date_time), 0, 1);
					+ "_" + Sys(2007, Trim(Padl(Cur2.PK, 11)), 0 ,1) + "_" As cUni_File,;
					Ttoc(Date_time)                                         As Datetime,;
					.F.                                                     As lMemLoaded,;
					Cur2.historyFK                                          As iReplaceFolder,;
					;
					Cur1.scope, ;
					Cur1.searchstr                                          As Search, ;
					Cur1.replacestr                                         As Replace_String, ;
					Cur2.* ;
					From       (Addbs(m.tcRoot) + "GF_Replace_History.DBF")  As Cur1 ;
					Inner Join (Addbs(m.tcRoot) + "GF_REPLACE_DETAILV5.DBF") As Cur2;
					On Cur1.Id = Cur2.historyFK ;
					Into Cursor GoFishReplaceHistory NoFilter Readwrite ;
					Order By Cur1.Id  Desc, Cur2.PK

				Use In GF_Replace_History
				Use In GF_REPLACE_DETAILV5
				Close Databases
				Set Database To (m.lcDbc)

				??""+0h0d+" "+0h0d+"Replace backup:"
				??", item count:",Reccount()
				?"-"
				lnCount = 0
				Scan
					Strtofile(ProcCode, m.lcDir + Trim(cUni_File) + "ProcCode.txt")
					Strtofile(Code    , m.lcDir + Trim(cUni_File) + "Code.txt")

					Do Case
						Case m.lnCount%1000=0
							??""+0h0d+" "+0h0d+"\"
						Case m.lnCount%1000=250
							??""+0h0d+" "+0h0d+"|"
						Case m.lnCount%1000=500
							??""+0h0d+" "+0h0d+"/"
						Case m.lnCount%1000=750
							??""+0h0d+" "+0h0d+"-"
					Endcase
					lnCount = m.lnCount+1
				Endscan &&All

*backup folder number
				Select;
					Max(iReplaceFolder) As iReplaceFolder;
					From GoFishReplaceHistory As Cur1;
					Into Cursor GoFishReplace_History

				If Reccount()=1
					Select GF_ReplaceID
					Replace;
						iID With GoFishReplace_History.iReplaceFolder+1
				Endif &&Reccount()=1
*/backup folder number

*history
				Select;
					Cur1.cUni,;
					Cur1.Datetime, ;
					Cur1.Search, ;
					Cur1.scope, ;
					Count (*) As RESULTS;
					From GoFishReplaceHistory As Cur1;
					Into Cursor GoFishReplace_History NoFilter;
					Group By 1,2,3,4
*append
				Insert Into (m.lcDBF_H);
					(cUni,;
					Datetime,;
					Search,;
					RESULTS,;
					scope,;
					lSaved,;
					lReplace);
					Select;
					Cur1.cUni,;
					Cur1.Datetime, ;
					Cur1.Search, ;
					Cur1.RESULTS, ;
					Cur1.scope,;
					.F. As lSaved,;
					.T. As lReplace;
					From GoFishReplace_History As Cur1

				Use In GoFishReplace_History
*/history

*search result
*save, if
				Insert Into (m.lcDBF);
					(cUni,cUni_File,Datetime,;
					scope,Search,lMemLoaded,;
					Process,FilePath,FileName,;
					TrimmedMatchLine,BaseClass,ParentClass,;
					Class,Name,MethodName,;
					ContainingClass,Classloc,MatchType,;
					Timestamp,FileType,Type,;
					Recno,ProcStart,ProcEnd,;
					Statement,StatementStart,firstmatchinstatement,;
					firstmatchinprocedure,MatchStart,MatchLen,;
					lIsText,Column,Id,;
					MatchLine,Replaced,TrimmedReplaceLine,;
					ReplaceLine,ReplaceRisk,Replace_DT,;
					iReplaceFolder,;
					lJustReplace;
					);
					Select;
					cUni,;
					cUni_File,;
					Datetime,;
					scope,;
					Search,;
					.F.,;
					Process,;
					FilePath,;
					FileName,;
					TrimmedMatchLine,;
					BaseClass,;
					ParentClass,;
					Class,;
					Name,;
					MethodName,;
					ContainingClass,;
					Classloc,;
					MatchType,;
					Timestamp,;
					FileType,;
					Type,;
					Recno,;
					ProcStart,;
					ProcEnd,;
					Statement,;
					StatementStart,;
					firstmatchinstatement,;
					firstmatchinprocedure,;
					MatchStart,;
					MatchLen,;
					lIsText,;
					Column,;
					Id,;
					MatchLine,;
					Replaced,;
					TrimmedReplaceLine,;
					ReplaceLine,;
					ReplaceRisk,;
					Replace_DT,;
					iReplaceFolder,;
					.T.;
					From GoFishReplaceHistory As Cur1
*search result

				Use In GoFishReplaceHistory

*+we need cUni, somehow
*+cUni_File, empty
*+date_time (t)-> DateTime (c)
*+lMemLoaded = .f.
*proccode -> nach GF_Saved_Search_Results, wie oben
*code -> nach GF_Saved_Search_Results, wie oben
*replacerisk = 100
**historyFK -> iReplaceFolder
*lJustReplaced = .t.
*lSaved = .f.

*parent
*ALTE daten via id/HistoryFK verknüft, siehe oben
*date_time (t)-> DateTime (c)
*search aus cursor
*results berechnen
*lsaved = .f.
*lreplace = .t.

*sum up max folder
*write folder nr

			Endif &&DIRECTORY(Addbs(m.tcRoot) + "GF_ReplaceBackups") AND FILE(Addbs(m.tcRoot) + "GF_REPLACE_DETAILV5.DBC") AND ...

*remove replace files
			Delete File (Addbs(m.tcRoot) + "GF_REPLACE_DETAILV5.*")
			Delete File (Addbs(m.tcRoot) + "GF_Replace_History.*")
*/Replace History


		Catch To m.loException
*			Set Step On
			lnReturn = 2

	Endtry

	Use In GF_ReplaceID
	Use In Select(m.toResultForm.oSearchEngine.cSearchResultsAlias)
	Use In Select(Juststem(m.lcDBF))
	Use In Select(Juststem(m.lcDBF_H))

	Close Database
	Set Database To (m.lcDatabase)
	Select(m.lnSelect)


	Cd (m.lcOldDir)

	Return m.lnReturn
Endproc &&GF_Change_TableStruct

Procedure GF_Move_GlobalPath  		&&Move GF settings from Home(7) to HOME(7)+"GoFish\"
	Local;
		lcSource As String,;
		lcTarget As String,;
		loFSO As "Scripting.FileSystemObject"

	loFSO = Createobject("Scripting.FileSystemObject")

*	lcSource = '"' + Home(7) + "GF_*.*" + '"'
*	lcTarget = '"' + Home(7) + "GoFish_ + '"'
*	Rename (m.lcSource) To (m.lcTarget)

	lcSource = Home(7) + "GF_*.*"
	lcTarget = Home(7) + "GoFish_"
	loFSO.MoveFile(m.lcSource, m.lcTarget)

	lcSource = Home(7) + "GF_Saved_Search_Results"
	If Directory(m.lcSource) Then
		lcTarget = Home(7) + "GoFish_\GF_Saved_Search_Results"

		loFSO.MoveFolder(m.lcSource,m.lcTarget)
	Endif &&DIRECTORY(lcSource)

	lcSource = Home(7) + "GoFishBackups"
	If Directory(m.lcSource) Then
		lcTarget = Home(7) + "GoFish_\GF_ReplaceBackups"

		loFSO.MoveFolder(m.lcSource,m.lcTarget)
	Endif &&DIRECTORY(lcSource)

Endproc &&GF_Move_GlobalPath

Procedure GF_Backup_GlobalPath  		&&Move GF settings from Home(7) to HOME(7)+"GoFish\"
	Local;
		lcSource As String,;
		lcTarget As String,;
		lcTarget2 As String,;
		lnLoop As Integer,;
		loFSO  As "Scripting.FileSystemObject"

	loFSO = Createobject("Scripting.FileSystemObject")

	lcSource = Home(7) + "GF_*.*"
	lcTarget = Home(7) + "GoFish_Backup"
	lnLoop   = 0
	Do While Directory(m.lcTarget)
		lnLoop   = m.lnLoop+1
		lcTarget = Home(7) + "GoFish_Backup_" +  Ltrim(Str(m.lnLoop,11,0))

	Enddo &&DIRECTORY(m.lcTarget)
	Mkdir (m.lcTarget)

	lcTarget2 = '"' + m.lcTarget + "\*.*" + '"'
	loFSO.CopyFile(m.lcSource, m.lcTarget)

	lcSource = Home(7) + "GF_Saved_Search_Results"
	If Directory(m.lcSource) Then
		lcTarget2 = m.lcTarget + "\GF_Saved_Search_Results"

		loFSO.CopyFolder(m.lcSource, m.lcTarget2)
	Endif &&DIRECTORY(lcSource)

	lcSource = Home(7) + "GoFishBackups"
	If Directory(m.lcSource) Then
		lcTarget2 = m.lcTarget + "\GoFishBackups"

		loFSO.CopyFolder(m.lcSource, m.lcTarget2)
	Endif &&DIRECTORY(lcSource)

Endproc &&GF_Backup_GlobalPath


* ================================================================================ 
Procedure GF_Copy_GlobalPath  		&& Copy GF settings from Home(7) to HOME(7)+"GoFish\"
	Local lcSource As String
	Local lcTarget As String
	Local loFSO As 'Scripting.FileSystemObject'

	loFSO = Createobject('Scripting.FileSystemObject')

	lcSource = Home(7) + 'GF_*.*'
	lcTarget = Home(7) + 'GoFish_'
	m.loFSO.CopyFile(m.lcSource, m.lcTarget)

	Mkdir(m.lcTarget + '\GF_Saved_Search_Results')

Endproc &&GF_Copy_GlobalPath



* ================================================================================

Procedure GF_Write_Readme_Text  	&&Create for README.md files
	Lparameters;
		tnFile,;
		tcFile,;
		tlForce

	Local;
		lcText As String

	If !File(m.tcFile) Or m.tlForce Then
		Do Case
			Case m.tnFile=1
*base folder
				TEXT To m.lcText Noshow
## GoFish settings folder
This folder contains the settings and history table for GoFish! code search tool.

- If GoFish is NOT working, try to delete the files AND subfolders as a whole.
- If the folder grows to large **Do not delete single files or tables**
  - delete *all files and folders*, but this will reset all options
  - delete history files through GoFish!
    - Right click nodes, choose clear (depends on tree mode)
    - Delete searches via *History* button
    - Wipe in Options form

GoFish! is available in VFP9 SP2 or VFPA from Thor or from https://github.com/VFPX/GoFish

				ENDTEXT &&lcText

			Case m.tnFile=2
*settings folder
				TEXT To m.lcText Noshow
## VFP settings folder with GoFish settings
This folder contains the settings and history table for GoFish! code search tool.

- If GoFish is NOT working, try to delete the `GF_*.*` files and the `GF_Saved_Search_Results` subfolder.
- If the GoFish data grows to large **Do not delete single files or tables**
  - delete files as mentioned above, but this will reset all options
  - delete history files through GoFish!
    - Right click nodes, choose clear (depends on tree mode)
    - Delete searches via *History* button
    - Wipe in Options form

GoFish! is available in VFP9 SP2 or VFPA from Thor or from https://github.com/VFPX/GoFish

				ENDTEXT &&lcText

			Case m.tnFile=3
*history folder
				TEXT To m.lcText Noshow
## GoFish history files folder
This folder contains the history files for GoFish! code search tool.
The files are not in a memo, so that the size might grow above 2GiB and faster access for the history is possible.

If the folder grows to large **Do not delete single files**.
- delete all GoFish files and folders, but this will reset all options
  - delete all files `..\GF_*.*`
  - delete all files in this folder
- delete history files through GoFish!
  - Right click nodes, choose clear (depends on tree mode)
  - Delete searches via *History* button
  - Wipe in Options form

GoFish! is available in VFP9 SP2 or VFPA from Thor or from https://github.com/VFPX/GoFish

				ENDTEXT &&lcText

			Case m.tnFile=4
*Replace Backup folder
				TEXT To m.lcText Noshow
## GoFish replace backup directories folder
This folder contains the backups of files replaced by GoFish! code search tool.

**Do not delete single files**.
- delete all GoFish files and folders, but this will reset all options
  - delete all files `..\GF_*.*`
  - delete all files in this folder
- delete backup files through GoFish!
  - Switch to the _Show replace history_ mode
  - Right click nodes, choose clear (depends on tree mode)
  - Wipe in Options form

GoFish! is available in VFP9 SP2 or VFPA from Thor or from https://github.com/VFPX/GoFish

				ENDTEXT &&lcText

			Case m.tnFile=5
*a Replace Backup
				TEXT To m.lcText Noshow
## GoFish replace backup directories folder
This folder contains the backups of a replace by GoFish! code search tool. See bottom.

**Do not delete single files**.
- delete all GoFish files and folders, but this will reset all options
  - delete all files `..\GF_*.*`
  - delete all files in this folder
- delete backup files through GoFish!
  - Switch to the _Show replace history_ mode
  - Right click nodes, choose clear (depends on tree mode)
  - Wipe in Options form

GoFish! is available in VFP9 SP2 or VFPA from Thor or from https://github.com/VFPX/GoFish

				ENDTEXT &&lcText

			Otherwise
				lcText = ""
		Endcase

		Strtofile(m.lcText, m.tcFile)

	Endif &&!File(m.tcFile) Or m.tlForce
Endproc &&GF_Write_Readme_Text

Procedure GF_GetMonitorStatistics
************************************************************************
* GF_GetMonitorStatistics
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


	Local;
		loMonitor As "EMPTY"

	Declare Integer GetSystemMetrics In user32 Integer nIndex

	loMonitor = Createobject("EMPTY")
*	AddProperty(loMonitor, "gnMonitors",      GetSystemMetrics(SM_CMONITORS))

	AddProperty(m.loMonitor, "gnVirtualLeft",   GetSystemMetrics(SM_XVIRTUALSCREEN))
	AddProperty(m.loMonitor, "gnVirtualTop",    GetSystemMetrics(SM_YVIRTUALSCREEN))

	AddProperty(m.loMonitor, "gnVirtualWidth",  GetSystemMetrics(SM_CXVIRTUALSCREEN))
	AddProperty(m.loMonitor, "gnVirtualHeight", GetSystemMetrics(SM_CYVIRTUALSCREEN))

	AddProperty(m.loMonitor, "gnVirtualRight",  m.loMonitor.gnVirtualWidth -  Abs(m.loMonitor.gnVirtualLeft) - 10)
	AddProperty(m.loMonitor, "gnVirtualBottom", m.loMonitor.gnVirtualHeight - Abs(m.loMonitor.gnVirtualTop)  -  5)

*ADDPROPERTY(loMonitor, "gnScreenHeight",  GetSystemMetrics(SM_CYFULLSCREEN))
*ADDPROPERTY(loMonitor, "gnScreenWidth",   GetSystemMetrics(SM_CXFULLSCREEN))

	Return m.loMonitor

Endproc &&GF_GetMonitorStatistics

Procedure Compare_VerNo
************************************************************************
* Compare_VerNo
****************************************
***  Function: Compares Version numbers in the form aaaa[.bbbb[.cccc[.ddddd]]]
***      Pass:
***       tcVerNoSys
***          System side of the comparision, usually this of the app.
***          Type 	Character
***          Direction 	Input
***          Optional
***          If empty, _Screen._GoFish.cVersion is used.
***      tcVerNoDBC
***          File side of the comparison
***          Typ 	Character
***          Direction 	Input
***          Typicaly the version of a database.
***      tnVerNoDBC
***          The return of tcVerNoDBC as numerical value, depending on tnPos
***          Typ 	Numeric
***          By refernce
***          Direction 	Output
***          Optional
***      tnPos
***          Sets how many levels of the versions are compared
***          Typ 	Numeric
***          By refernce
***          Direction 	Input
***			 Value 	Description
***          1 	only the aaaa part
***          2 	10^4*aaaa+bbbb (this is the default)
***          3 	10^4*(10^4*aaaa+bbbb)+cccc
***          4 	10^4*(10^4*(10^4*aaaa+bbbb)+cccc)+dddd
***
***    Return:  numeric
***    Value 	Description
***    0 	tcVerNoSys=tcVerNoDBC
***    1 	tcVerNoSys>tcVerNoDBC
***    2 	tcVerNoSys<tcVerNoDBC
***
*** SF 20230414
***
************************************************************************

	Lparameters ;
		tcVerNoSys,;
		tcVerNoDBC,;
		tnVerNoDBC,;
		tnPos

	Local;
		lnVerNoSys As Number

	If Empty(m.tcVerNoSys) Then
		tcVerNoSys = _Screen._GoFish.cVersion
	Endif &&EMPTY(tcVerNoSys)

	If Vartype(m.tnPos)#"N" Then
		tnPos = 2
	Endif &&VARTYPE(tnPos)#"N"


	lnVerNoSys = Val(Getwordnum(m.tcVerNoSys, 1, '.'))
	tnVerNoDBC = Val(Getwordnum(m.tcVerNoDBC, 1, '.'))

	If m.tnPos>1 Then
		lnVerNoSys = m.lnVerNoSys*10^4+Val(Getwordnum(m.tcVerNoSys, 2, '.'))
		tnVerNoDBC = m.tnVerNoDBC*10^4+Val(Getwordnum(m.tcVerNoDBC, 2, '.'))
		If m.tnPos>2 Then
			lnVerNoSys = m.lnVerNoSys*10^4+Val(Getwordnum(m.tcVerNoSys, 3, '.'))
			tnVerNoDBC = m.tnVerNoDBC*10^4+Val(Getwordnum(m.tcVerNoDBC, 3, '.'))
			If m.tnPos>3 Then
				lnVerNoSys = m.lnVerNoSys*10^4+Val(Getwordnum(m.tcVerNoSys, 4, '.'))
				tnVerNoDBC = m.tnVerNoDBC*10^4+Val(Getwordnum(m.tcVerNoDBC, 4, '.'))
			Endif &&tnPos>3
		Endif &&tnPos>2
	Endif &&tnPos>1

	Do Case
		Case m.lnVerNoSys=m.tnVerNoDBC
			Return 0
		Case m.lnVerNoSys>m.tnVerNoDBC
			Return 1
		Case m.lnVerNoSys<m.tnVerNoDBC
			Return 2
	Endcase
Endproc &&Compare_VerNo

Procedure NewVersion
************************************************************************
* NewVersion
****************************************
***  Function: The programm is newer then the data, change the data
***      Pass:
***      tnVerNoDBC
***          The version of the DBC as numerical value
***
*** SF 20230415
***
************************************************************************

	Lparameters ;
		tnVerNoDBC
	
	Local;
		llReturn As Boolean
	
	llReturn = .T.
	Do Case
		Case Val(Getwordnum(_Screen._GoFish.cVersion, 1, "."))=6;
				And Val(Getwordnum(_Screen._GoFish.cVersion, 2, "."))=1
	*update to 6.1.*
			Do Case
				Case m.tnVerNoDBC=0
	*from older version (i.e. 6.0)
					llReturn = Update_to_6_1()
	
				Otherwise
	
			Endcase
		Case Val(Getwordnum(_Screen._GoFish.cVersion, 1, "."))=6;
				And Val(Getwordnum(_Screen._GoFish.cVersion, 2, "."))=2
	*update to 6.2.*
			Do Case
				Case m.tnVerNoDBC=0
	*from older version (i.e. 6.0)
					llReturn = Update_to_6_2()
	
				Case m.tnVerNoDBC=60001
	*from version 6.1
					llReturn = Update_to_6_2()
	
				Otherwise
	
			Endcase
	
		Otherwise
	
	Endcase
	Return m.llReturn
	Endproc &&NewVersion
	
Procedure Update_to_6_1
************************************************************************
* Update_to_6_1
****************************************
***  Update to version 6.1 (from lower versions)
***
*** field FileName to short, change size, rewrite
*** SF 20230415
***
*** would fail if no search was running ever
*** SF 20230415
***
************************************************************************
	If Indbc("GF_Results_Form_Settings", "Table") Then
		Use GF_Results_Form_Settings Exclusive
		Alter Table GF_Results_Form_Settings;
			Alter Column FileName c(100)

		Update GF_Results_Form_Settings Set;
		 FileName = Justfname(FilePath)
		Use
	Endif &&Indbc("GF_Results_Form_Settings", "Table")

	RETURN .T.
Endproc &&Update_to_6_1

Procedure Update_to_6_2
************************************************************************
* Update_to_6_2
****************************************
***  Update to from version 6.1 to version 6.2
***
*** Tables might not be part of the DBC
***
*** SF 20230419
***
************************************************************************
	Local;
		lcPath As String,;
		lnTable As Number

	Local Array;
		laTables(5, 2)

	laTables = .T.

	laTables[1, 1] = "GF_ReplaceID"
	laTables[2, 1] = "GF_Search_History"
	laTables[3, 1] = "GF_Search_Scope_History"
	laTables[4, 1] = "GF_Results_Form_Settings"
	laTables[4, 2] = .F.
	laTables[5, 1] = "GF_Search_Expression_History"
	lcPath         = Addbs(Justpath(Dbc()))

	For lnTable = 1 To 5
		If !Indbc(laTables[m.lnTable, 1], "Table") Then
			laTables[m.lnTable, 1] = m.lcPath + laTables[m.lnTable, 1] + ".DBF"
			If File(laTables[m.lnTable, 1]) Then
				Add Table (laTables[m.lnTable, 1] + ".DBF")

			Else &&File(laTables[m.lnTable, 1])
				If laTables[m.lnTable, 2] Then
					Return .F.
				Endif &&laTables[m.lnTable, 2]

			Endif &&File(laTables[m.lnTable, 1])
		Endif &&!Indbc(laTables[1, 1], "Table")
	Endfor &&lnTable
	
	llReturn = Update_to_6_1()

	RETURN .T.
Endproc &&Update_to_6_2




* ================================================================================

Procedure GF_RelativePath
	Lparameters tcFileName, toForm, toOptions

	Local lcScope, lcPath

	If m.toForm.lShowRelativePath
		lcScope = m.toOptions.cRecentScope
		Do Case
			Case m.toOptions.nSearchScope = 5 && all open projects.
				lcPath = Lower(Sys(2014, m.tcFileName))
			Case File(m.lcScope)
				lcPath = Lower(Sys(2014, m.tcFileName, m.lcScope))
			Otherwise
				lcPath = Lower(Sys(2014, m.tcFileName, Addbs(m.lcScope)))
		Endcase

		If Len(m.lcPath) < Len(m.tcFileName)
			Return m.lcPath
		Else
			Return m.tcFileName
		EndIf
		
	Else
		Return m.tcFileName
	Endif && toOptions.lRelativePath

Endproc


*----------------------------------------------------------------------------------
Procedure GF_GetRegExp

	Local loRegEx As 'VBScript.RegExp'
	Local loException

	Try
		loRegEx = Createobject ('VBScript.RegExp')
	Catch To m.loException

	Endtry

	If Vartype(m.loRegEx) # 'O'
		loRegEx = SF_RegExp()
	Endif

	Return m.loRegEx

Endproc


*----------------------------------------------------------------------------------
Procedure GF_toRGB
	Lparameters tnColor

	Local lcResult, lnColor

	lnColor	 = m.tnColor
	lcResult = 'rgb(' + Transform(m.lnColor % 256) + ','
	lnColor	 = Int(m.lnColor / 256)
	lcResult = m.lcResult + Transform(m.lnColor % 256) + ','
	lnColor	 = Int(m.lnColor / 256)
	lcResult = m.lcResult + Transform(m.lnColor % 256) + ')'

	Return m.lcResult

Endproc


* ================================================================================ 
Procedure GF_FileNameToClipboard(tcFileName, tcType)
	Local lcPrompt

	Do Case
		Case m.tcType = 'Name'
			_Cliptext = Justfname(Trim(m.tcFileName))
			lcPrompt  = 'File name copied to clipboard'
		Case m.tcType = 'Stem'
			_Cliptext = Juststem(Trim(m.tcFileName))
			lcPrompt  = 'File stem copied to clipboard'
		Case m.tcType = 'Path'
			_Cliptext = Trim(m.tcFileName)
			lcPrompt  = 'File name (full path) copied to clipboard'
	Endcase
	Wait (m.lcPrompt)					;
		Window At Mrow(), Mcol()		;
		Nowait Timeout 2
Endproc


* ================================================================================ 
Procedure GF_InternalTimeStamp
	Set Date To American
	Set Hours To 12
Endproc


* ================================================================================ 
Procedure GF_ExternalTimeStamp
	Set Date To (_Screen._GoFish.cSetDate)
	Set Hours To (_Screen._GoFish.csethours)
Endproc


* ================================================================================ 
Procedure GF_GetHackCX

	#Define		ccThorKey			'File name'
	#Define		ccThorTool			'HackCX'

	Local lcExecutable

	If GF_IsThorThere()
		lcExecutable = Execscript(_Screen.cThorDispatcher, 'Get Option=', ccThorKey, ccThorTool)
		If File(Nvl(m.lcExecutable, ''))
			Return m.lcExecutable
		Endif
	Endif
	Return ''
Endproc


* ================================================================================ 
Procedure GF_FullPath(tcFileName, tcProjectPath)

	Local lcFileName, lcProjectPath

	lcFileName	  = Trim(m.tcFileName, 1, Chr[0])
	lcProjectPath = Addbs(m.tcProjectPath)
	Do Case
		Case ':' $ m.lcFileName
			Return m.lcFileName
		Case Isalpha(m.lcFileName)
			Return m.lcProjectPath + m.lcFileName
		Case m.lcFileName = '..\'
			Return GF_FullPath(Substr(m.lcFileName, 4), Justpath(Justpath(m.lcProjectPath)))
		Otherwise
			Return Fullpath(m.lcFileName, m.lcProjectPath)
	Endcase
Endproc


* ================================================================================ 
Function GF_GetMixedCaseCombination(m.tcString)
	Local lcResult, lnI

	lcResult = ''
	For lnI = 1 To Len(m.tcString)
		lcResult = m.lcResult + '[' + Upper(Substr(m.tcString, m.lnI, 1)) + Lower(Substr(m.tcString, m.lnI, 1)) + ']'
	Endfor

	Return m.lcResult
Endfunc


* ================================================================================ 
Procedure Sleep
	Lparameters lnMilliseconds
	Declare Sleep In kernel32 Integer dwMilliseconds
	Sleep(m.lnMilliseconds)
Endproc


* ================================================================================
Procedure GF_ModifyCommand(tcFile)
	Modify Command (m.tcFile) Nowait Save
EndProc 

* ================================================================================ 
Procedure GF_FixFontSize
	Lparameters toObject, tnOldFontSize, tnNewFontSize

	Local loFontsize As 'FixFontSize'
	
	If m.tnOldFontSize >= m.tnNewFontSize
		Return
	EndIf 

	loFontsize = Createobject('FixFontSize', m.toObject, m.tnOldFontSize, m.tnNewFontSize)
	
	toObject.FontSize = m.tnNewFontSize
	
	If PemStatus(toObject, 'nFontSizeMultiplier', 5)
		toObject.nFontSizeMultiplier = loFontsize.nFontSizeMultiplier
	endif

	Return
	
EndProc


Procedure GF_EditSkipList(toSearchEngine)

	Local loPBT As 'GF_PEME_BaseTools'
	Local lcFileName, lcText

	lcFileName = m.toSearchEngine.cFilesToSkipFile

	If Not File(m.lcFileName)

		Text To m.lcText Pretext 3 Noshow
		** Contains names of files and folders to be skipped: one per line, case insensitive, wildcards allowed
		**  -- File name (no path)
		**        Main.PRG
		**        FOO*.TXT
		**  -- Folders (with leading and trailing backslash)
		**        \Purchased\WestWind\
		**        \Temp*\
		**  -- Or combined
		**        \Temp*\*.txt

		** These files and folders will be excluded if the "Exclude Files option" is checked on the Advanced form.
		** The default name and location for this file is: (Home(7) + 'GoFish_\GF_Files_To_Skip.txt')
		** Might be set to a local directory

		** Blank lines and lines beginning with ** are not processed.
		
		Endtext

		Strtofile(m.lcText, m.lcFileName)
	Endif

	Modify File (m.lcFileName)
	loPBT = Createobject('GF_PEME_BaseTools')
	m.loPBT.AddMRUFile(m.lcFileName)

Endproc


Procedure GF_EditIncludeList(toSearchEngine)

	Local loPBT As 'GF_PEME_BaseTools'
	Local lcFileName, lcText

	lcFileName = m.toSearchEngine.cFilesToIncludeFile

	If Not File(m.lcFileName)

		Text To m.lcText Pretext 3 Noshow
		** Contains names of files or folders to be included (in addition to the selected scope): 
		**   - one per line, case insensitive, no wildcards
		**   - for directories, does not search sub-directories

		** These files and folders will be included if the "Include Files option" is checked on the Advanced form.
		** The default name and location for this file is: (Home(7) + 'GoFish_\GF_Files_To_Include.txt')

		** Blank lines and lines beginning with ** are not processed.
		
		Endtext

		Strtofile(m.lcText, m.lcFileName)
	Endif

	Modify File (m.lcFileName)
	loPBT = Createobject('GF_PEME_BaseTools')
	m.loPBT.AddMRUFile(m.lcFileName)

Endproc


* ================================================================================

Define Class FixFontSize As Custom

	oForm		= Null
	nFontSizeMultiplier	= 0

	Dimension aAnchors(1, 2)


	Procedure Init(toObject, tnOldFontSize, tnNewFontSize)

		This.oForm		 = m.toObject
		This.nFontSizeMultiplier = m.tnNewFontSize / m.tnOldFontSize

		This.SaveAnchors(m.toObject)

		This.ResizeObject(m.toObject)

		This.RestoreAnchors()

		Return

	Endproc


	Procedure SaveAnchors(toObject)

		Local lnAnchor, lnCount, loChildObject, loException

		If Pemstatus(m.toObject, 'Objects', 5)

			Try
				For Each m.loChildObject In m.toObject.Objects FoxObject
					If Pemstatus(m.loChildObject, 'Anchor', 5)
						lnAnchor = Getpem(m.loChildObject, 'Anchor')
						If m.lnAnchor > 0
							lnCount = Alen(This.aAnchors, 1) + 1
							Dimension This.aAnchors(m.lnCount, 2)
							This.aAnchors(m.lnCount, 1)	= m.loChildObject
							This.aAnchors(m.lnCount, 2)	= m.lnAnchor
							loChildObject.Anchor		= 0
						Endif
					Endif
					This.SaveAnchors(m.loChildObject)

				Endfor
			Catch To m.loException

			Endtry

		Endif && Pemstatus(m.toObject, 'Objects', 5)

	Endproc


	Procedure ResizeObject(toObject)

		Local lcSaveName, loChildObject, loException

		lcSaveName = 'Anchor' + Sys(2015)
		If Pemstatus(m.toObject, 'Objects', 5)

			Try
				For Each m.loChildObject In m.toObject.Objects FoxObject
					This.ResizeObject(m.loChildObject)
				Endfor
			Catch To m.loException

			Endtry

		Endif

		This.SetFontSize(m.toObject)

	Endproc


	Procedure SetFontSize(toObject)
		If m.toObject.BaseClass = 'Form'
			This.SetSize(m.toObject, 'MaxHeight')
			This.SetSize(m.toObject, 'MaxWidth')
		Else
			This.SetSize(m.toObject, 'Top')
			This.SetSize(m.toObject, 'Left')
		Endif

		This.SetSize(m.toObject, 'Height')
		This.SetSize(m.toObject, 'Width')
		This.SetSize(m.toObject, 'FontSize')

	Endproc


	Procedure SetSize(toObject, tcProperty)
		Local lnOldValue

		If Pemstatus(m.toObject, m.tcProperty, 5)
			lnOldValue			 = Getpem(m.toObject, m.tcProperty)
			toObject.&tcProperty = Round(m.lnOldValue * This.nFontSizeMultiplier, 0)
		Endif
	Endproc


	Procedure RestoreAnchors()

		Local lnAnchor, lnI, loObject

		For lnI = 2 To Alen(This.aAnchors, 1)
			loObject				= This.aAnchors[m.lnI, 1]
			lnAnchor				= This.aAnchors[m.lnI, 2]
			loObject.Anchor			= m.lnAnchor
			This.aAnchors[m.lnI, 1]	= Null
		Endfor

	EndProc
	
	
Enddefine
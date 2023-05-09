#INCLUDE BuildGoFish.h

*---------------------------------------------------------------------------------------------
Lparameters;
	tcInitialSource,;
	tlSuppressRegisteredWithThorDialog,;
	tv03,tv04,tv05,tv06,tv07,tv08,;
	tv09,tv10,tv11,tv12,tv13,tv14,tv15,tv16,;
	tv17,tv18,tv19,tv20,tv21,tv22,tv23,tv24


Local;
	lcInitialSource   As String,;
	lcOlSafety        As String,;
	lcPath            As String,;
	lcSettingsFile    As String,;
	llResetCommonStorage As Boolean,;
	llResetLocalStorage As Boolean,;
	llReturn          As Boolean,;
	loMy              As "My" Of "My.vcx",;
	loRegexp          As Object,;
	loSettings        As Object

lcInitialSource = Evl(m.tcInitialSource, "")

SetupEnvironment()

*!* ******************** Removed 11/03/2015 *****************
*!* CreateVersionFile() && Version file is re-built every time

If !Empty(m.tcInitialSource) And Upper(m.tcInitialSource) = "THOR"
	llReturn = TryRegisterWithThor(.T., !m.tlSuppressRegisteredWithThorDialog)
	Return m.llReturn
Else
	TryRegisterWithThor() && GoFish is re-registered with Thor every time
Endif

*-- Only allow one instance of the form to be running
If Type("_Screen._GoFish.oResultsForm") = "O" And !Isnull(_Screen._GoFish.oResultsForm)
	If _Screen._GoFish.oResultsForm.WindowState = 1
		_Screen._GoFish.oResultsForm.WindowState = 0
	Else
		_Screen._GoFish.oResultsForm.Show()
	Endif

	Return
Endif

Set Procedure To "GoFishProc" Additive
Set Procedure To "mhHtmlCode" Additive
Set Procedure To "GoFishSearchEngine.prg" Additive
Set Procedure To "GoFishSearchOptions.prg" Additive
Set Procedure To "GF_PEME_BaseTools.prg" Additive

* SF 20221123 use subfolder of Home(7) by default
lcSettingsFile = Addbs(Home(7)) + "GoFish_"

If !Directory(m.lcSettingsFile) Then
	Mkdir (Addbs(Home(7)) + "GoFish_")
	Do Form gf_migrate_6.scx
	Read Events
	If File(Addbs(Home(7)) + "GF_Results_Form_Settings.xml") Then
		lcOlSafety = Set("Safety")
		Set Safety Off
		GF_Backup_GlobalPath()
		GF_Move_GlobalPath()
		Set Safety &lcOlSafety
	Endif &&File(Addbs(Home(7)) + "GF_Results_Form_Settings.xml")
Endif &&DIRECTORY(m.lcSettingsFile)
* /SF 20221123 use subfolder of Home(7) by default

If Inlist(Upper(Alltrim(m.lcInitialSource)), "/?", "-?", "/H", "-H","HELP")
	HelpScreen()
	Return
Endif &&INLIST(Upper(Alltrim(m.lcInitialSource)), "/?", "-?", "/H", "-H","HELP")

*-- Erase all existing GF XML files to reset to default
If Upper(Alltrim(m.lcInitialSource)) == "-RESETLOCAL"
	llResetLocalStorage = .T.
	lcInitialSource     = ""
Endif &&Upper(Alltrim(lcInitialSource)) == "-RESETLOCAL"

If Upper(Alltrim(m.lcInitialSource)) == "-RESET"
	llResetCommonStorage = .T.
	lcInitialSource      = ""

Endif &&Upper(Alltrim(lcInitialSource)) == "-RESET"

If m.llResetCommonStorage
	GF_RemoveFolder(m.lcSettingsFile)
Endif

*-- Erase all existing GF XML files to reset to default

loMy           = Newobject("My", "My.vcx")
loSettings     = m.loMy.Settings
lcSettingsFile = m.lcSettingsFile + "\GF_Results_Form_Settings.xml"
* SF 20221017 test for file
If File(m.lcSettingsFile)
	loSettings.Load(m.lcSettingsFile)
Endif
* /SF 20221017 test for file

* SF 20221017
* special local settings
If loSettings.Exists("lCR_Allow") And m.loSettings.lCR_Allow Then
	GF_Get_LocalSettings(@loSettings, m.lcSettingsFile, m.llResetLocalStorage)

Endif &&loSettings.EXISTS("lCR_Allow") And m.loSettings.lCR_Allow
*/ SF 20221017

lnSys16 = 1
DO WHILE !LOWER(JUSTFNAME(SYS(16,m.lnSys16)))=="gofish5.app"
   lnSys16 = m.lnSys16+1
ENDDO

lcPath = Sys(16,m.lnSys16)

lcPath = Right(m.lcPath,Len(m.lcPath)-At(" ",m.lcPath,2))
lcPath = Justpath(m.lcPath)+'\SF_RegExp'

loRegexp = SF_RegExp(m.lcPath)

If Pemstatus(m.loSettings, "lDesktop", 5) And m.loSettings.lDesktop
	Do Form GoFish_Results_Desktop With m.lcInitialSource
Else
	Do Form GoFish_Results With m.lcInitialSource
Endif Pemstatus(m.loSettings, "lDesktop", 5) ...

If Version(2) = 0 && If running as an .EXE, setup Read Events loop
	On Shutdown Clear Events
	Read Events
	On Shutdown

Endif

*----------------------------------------------------------------------------
Procedure SetupEnvironment


	Local;
		lcAppName As String,;
		lcAppPath As String,;
		loGoFish As Object

*:Global;
x

*-- 4.2.003 - Need to clear out these class definitions which may be cached by VFP from an
*-- older version of GoFish
	Clear Class "GoFishSearchEngine"
	Clear Class "GoFishSearchOptions"
	Clear Class "GF_PEME_BaseTools"

	For x = Program(-1) To 1 Step -1 && Look up through run stack to find the name of the running .APP file
		lcAppName = Sys(16, x)
		If ".APP" $ Upper(m.lcAppName)
			Exit
		Else
			lcAppName = ""
		Endif
	Endfor

	lcAppName = Evl(m.lcAppName, Sys(16, 1))
	lcAppPath = Addbs(Justpath (Getwordnum(m.lcAppName, 3)))
	If Empty(m.lcAppPath)
		lcAppPath = Addbs(Justpath(m.lcAppName))
	Endif
	lcAppName = Justfname(m.lcAppName)

	If ".FXP" $ m.lcAppName
		SetPathsForDevelopmentMode(m.lcAppPath)
		lcAppName = GOFISH_APP_FILE
	Endif

	If Type("_Screen._GoFish.oResultsForm") = "O" And !Isnull(_Screen._GoFish.oResultsForm)
		Return
	Else

*-- 4.2.003 - Create this object locally, rather than getting if from the VCX
*-- This is better since finding the VCX during a Thor installation of GoFish can
*-- be tricky.
		loGoFish = Createobject("Empty")
		AddProperty(m.loGoFish, "cAppPath", m.lcAppPath)
		AddProperty(m.loGoFish, "cAppName", m.lcAppName)
		AddProperty(m.loGoFish, "cVersion", GOFISH_VERSION)
		AddProperty(m.loGoFish, "cBuildDate", GOFISH_BUILDDATE)
		AddProperty(m.loGoFish, "dBuildDate", GOFISH_dBUILDDATE)
		AddProperty(m.loGoFish, "oResultsForm", .Null.)

		_Screen.AddProperty("_GoFish", m.loGoFish) && Add this object onto _Screen, so it can accessed from GoFish forms

	Endif

Endproc

*--------------------------------------------------------------------------------
Procedure SetPathsForDevelopmentMode(tcAppPath)

*-- If running this bootstrap in dev mode,we need to setup these paths
*--  Note: This is not required when the compiled .app file is running
	Set Path To (m.tcAppPath) Additive
	Set Path To (m.tcAppPath + "Prg") Additive
	Set Path To (m.tcAppPath + "Forms") Additive
	Set Path To (m.tcAppPath + "Lib") Additive
	Set Path To (m.tcAppPath + "Lib\VFP\My") Additive
	Set Path To (m.tcAppPath + "Lib\VFP\FFC") Additive
	Set Path To (m.tcAppPath + "Images") Additive
	Set Path To (m.tcAppPath + "Menus") Additive

Endproc

*----------------------------------------------------------------------------
Procedure TryRegisterWithThor(tlForcedRegister, tlShowConfirmationDialog)

	Local;
		lcThorVersion As String,;
		llReturn   As Boolean,;
		llThorPresent As Boolean

	Try && See if Thor is running
			lcThorVersion = Execscript(_Screen.cThorDispatcher, "Version=")
			llThorPresent = .T.
		Catch
			If m.tlShowConfirmationDialog
				Messagebox("Thor is not presently running on your system. Please install/run it first.", 0, "Thor required.")
			Endif
			llThorPresent = .F.
		Finally
	Endtry

	If !m.llThorPresent
		Return .F.
	Endif

	llReturn = RegisterWithThor()

	If !m.tlForcedRegister
		Return
	Endif

	If m.tlShowConfirmationDialog
		If m.llReturn
			Messagebox("Successfully registered with Thor.", 0, GOFISH_APP_NAME)
		Else
			Messagebox("Error attempting to register with Thor.", 0, GOFISH_APP_NAME)
		Endif
	Endif

	Return m.llReturn

Endproc

*------------------------------------------------------------------------------------------------------
Procedure RegisterWithThor

	Local;
		llRegisterApp  As Boolean,;
		llRegisterUpdater As Boolean

	llRegisterApp = RegisterAppWithThor()

	llRegisterUpdater = RegisterUpdaterWithThor()

	RegisterVfpxLinkWithThor()
	RegisterDiscussionGroupWithThor()

	Return (m.llRegisterApp And m.llRegisterUpdater)

Endproc


*--------------------------------------------------------------------------------------
Procedure RegisterAppWithThor()

	Local;
		lcCode     As String,;
		lcFolderName As String,;
		lcFullAppName As String,;
		lcPlugIn   As String,;
		llRegister As Boolean,;
		loThorInfo As Object

	Try
			loThorInfo = Execscript (_Screen.cThorDispatcher, "Thor Register=")
		Catch
			loThorInfo = .Null.
	Endtry

	If Isnull (m.loThorInfo)
		Return .F.
	Endif


	With m.loThorInfo
* Required
		.Prompt      = GOFISH_APP_NAME && used when tool appears in a menu
		.Description = "Advanced Code Search Tool" && may be lengthy, including CRs, etc
		.Author      = "Matt Slay"
		.PRGName     = THOR_TOOL_NAME  && a unique name for the tool; note the required prefix
		.Category    = "Applications|GoFish"

		.AppName     = GOFISH_APP_FILE  && no path, but include the extension; for example, GoFish5.App
* Note that this creates This.FullAppName, which determines the full path of the APP
* and also This.FolderName

*	.FolderName = "C:\Visual FoxPro\Programs\GoFish All Versions\GoFish5\Source"
*	.FullAppName = .FolderName + "\" + GOFISH_APP_FILE

		lcFolderName  = .FolderName
		lcFullAppName = .FullAppName

		.PlugInClasses = "clsGoFishFormatGrid"
		.PlugIns       = "GoFish Results Grid"

		Text To m.lcCode Noshow Textmerge
  		Do "<<lcFullAppName>>"

		Endtext

		Text To m.lcPlugIn Noshow

EndProc


Define Class clsGoFishFormatGrid As Custom

	Source				= "GoFish5"
	PlugIn				= "GoFish Results Grid"
	Description			= "Provides access to GoFish results grid to set colors and other dynamic properties."
	Tools				= "GoFish5"
	FileNames			= "Thor_Proc_GoFish_FormatGrid.PRG"
	DefaultFileName		= "*Thor_Proc_GoFish_FormatGrid.PRG"
	DefaultFileContents	= ""

	Procedure Init
		****************************************************************
		****************************************************************
		***TEXT*** To This.DefaultFileContents Noshow
Lparameters toGrid, tcResultsCursor

*-- Sample 1: Dynamic row coloring as used by GF
Local lcFileNameColor, lcPRG, lcPRGColor, lcSCX, lcSCXColor, lcVCX, lcVCXColor

lcSCX	   = "Upper(" + tcResultsCursor + '.filetype) = "SCX"'
lcSCXColor = "RGB(0,0,128)"

lcVCX	   = "Upper(" + tcResultsCursor + '.filetype) = "VCX"'
lcVCXColor = "RGB(0,128,0)"

lcPRG	   = "Upper(" + tcResultsCursor + '.filetype) $ "PRG TXT H INI XML HTM HTML ASP ASPX"'
lcPRGColor = "RGB(255,0,0)"

toGrid.SetAll("DynamicForeColor", "Iif(" + m.lcSCX + ", " + m.lcSCXColor + ", " +		;
	  "Iif(" + m.lcVCX + ", " + m.lcVCXColor + ", " +									;
	  "Iif(" + m.lcPRG + "," + m.lcPRGColor + ", RGB(0,0,0))" +							;
	  ")" +																				;
	  ")", "COLUMN")
Return

*-- Sample 2: Alternative provided by Jim R. Nelson
*-- Assigns row colors based on field MatchType
#Define ccBolds "<Method>", "<Procedure>", "<Function>", "<Constant>", "<<Class Def>>", "<<Method Def>>", "<<Property Def>>"
#Define ccPropertyName "<Property Name>"
#Define ccPropertyValue "<Property Value>"

Local lcBolds, lcCode, lcCodeColor, lcComments, lcCommentsColor, lcFileName, lcOthersColor
Local lcPropNameColor, lcPropertyName, llRegister

lcPropertyName = '0 # Atc("<Property", # + m.tcResultsCursor + ".MatchType)"
lcComments	   = '0 # Atc("<Comment", ' + m.tcResultsCursor + ".MatchType)"
lcFileName	   = '0 # Atc("<File", ' + m.tcResultsCursor + ".MatchType)"
lcCode		   = "Left(" + m.tcResultsCursor + '.MatchType, 1) # "<"'

* ForeColor
lcPropNameColor	= "RGB(0,128,0)"
lcCommentsColor	= "RGB(0,0,0)"
lcCodeColor		= "RGB(0,0,0)"
lcOthersColor	= "RGB(0,0,255)"

m.toGrid.SetAll ("DynamicForeColor", "ICase(" +					;
	  m.lcPropertyName + ", " + m.lcPropNameColor + ", " +		;
	  m.lcComments + ", " + m.lcCommentsColor + ", " +			;
	  m.lcCode + ", " + m.lcCodeColor + ", " +					;
	  m.lcOthersColor + ")")

* BackColor
lcCommentsColor	= "RGB(192,192,192)"
lcFileNameColor	= "Rgb(176, 224, 230)"
m.toGrid.SetAll ("DynamicBackColor", "ICase(" +				;
	  m.lcComments + ", " + m.lcCommentsColor + ", " +		;
	  m.lcFileName + ", " + m.lcFileNameColor + ", " +		;
	  " Rgb(255,255,255))")

* Bold
lcBolds		 = "Inlist(" + m.tcResultsCursor + ".MatchType, ccBolds)"
m.toGrid.SetAll ("DynamicFontBold", m.lcBolds)

Return
		End***TEXT***
		endproc
	enddefine
		Endtext
		.Code = Strtran(lcCode + Evl(lcPlugIn, ""), "***TEXT***", "Text")

* Optional
		.StatusBarText = GOFISH_APP_NAME
		.Summary       = "Code Search Tool" && if empty, first line of .Description is used
		.Classes       = "loGoFish = gofishsearchengine of lib\gofishsearchengine.vcx"

* For public tools, such as PEM Editor, etc.
		.Source  = "GoFish" && e.g., "PEM Editor"
		.Version = GOFISH_VERSION  && e.g., "Version 7, May 18, 2011"
		.Sort    = Iif("BETA" $ Upper(GOFISH_APP_NAME), 5, 1) && the sort order for all items from the same .Source
		.Link    = GOFISH_HOME_PAGE

		llRegister = .Register()

	Endwith

	Return llRegister
Endproc

*--------------------------------------------------------------------------------------
Procedure RegisterUpdaterWithThor()

*-- ToDo: 2011-10-18: Jim Nelson needs to add support for this call in Thor Framework.
*-- See: https://bitbucket.org/JimRNelson/thor/issue/3/need-a-thor-api-call-to-create-an-updater

	Return


Endproc

*--------------------------------------------------------------------------------------
Procedure RegisterVfpxLinkWithThor()

	Local;
		lcCommonText As String,;
		lcFolderName As String,;
		lcFullAppName As String,;
		llRegister As Boolean,;
		loThorInfo As Object

	Try
			loThorInfo = Execscript (_Screen.cThorDispatcher, "Thor Register=")
		Catch
			loThorInfo = .Null.
	Endtry

	If Isnull (m.loThorInfo)
		Return .F.
	Endif

	lcCommonText = "GoFish page on VFPx"

	With m.loThorInfo
* Required
		.Prompt      = m.lcCommonText && used when tool appears in a menu
		.Description = m.lcCommonText && may be lengthy, including CRs, etc
		.Author      = "Matt Slay"
		.PRGName     = "Thor_Tool_GoFish_VFPx_Page"  && a unique name for the tool; note the required prefix
		.Category    = "Applications|GoFish"

		.AppName      = GOFISH_APP_FILE && no path, but include the extension; for example, GoFish5.App
		lcFolderName  = .FolderName
		lcFullAppName = .FullAppName

		Text To .Code Noshow Textmerge
			oShell = Createobject("wscript.shell")
			oShell.Run("https://github.com/VFPX/GoFish")
		Endtext

* Optional
		.StatusBarText = m.lcCommonText
		.Summary       = m.lcCommonText && if empty, first line of .Description is used

* For public tools, such as PEM Editor, etc.
		.Source  = "GoFish" && e.g., "PEM Editor"
		.Version = "" && e.g., "Version 7, May 18, 2011"
		.Sort    = 3 && the sort order for all items from the same .Source
		.Link    = "https://github.com/VFPX/GoFish"

		llRegister = .Register()

	Endwith

	Return m.llRegister
Endproc

*--------------------------------------------------------------------------------------
Procedure RegisterDiscussionGroupWithThor()

	Local;
		lcCommonText As String,;
		lcFolderName As String,;
		lcFullAppName As String,;
		llRegister As Boolean,;
		loThorInfo As Object

	Try
			loThorInfo = Execscript (_Screen.cThorDispatcher, "Thor Register=")
		Catch
			loThorInfo = .Null.
	Endtry

	If Isnull (m.loThorInfo)
		Return .F.
	Endif

*	lcCommonText = "GoFish discussion group"
	lcCommonText = "GoFish issues and discussions"

	With m.loThorInfo
* Required
		.Prompt      = m.lcCommonText && used when tool appears in a menu
		.Description = m.lcCommonText && may be lengthy, including CRs, etc
*		.Author      = "Matt Slay"
*		.PRGName     = "Thor_Tool_GoFish_Discussion_Group"  && a unique name for the tool; note the required prefix
		.Author   = "Lutz Scheffler"
		.PRGName  = "Issues"  && a unique name for the tool; note the required prefix
		.Category = "Applications|GoFish"

		.AppName      = GOFISH_APP_FILE && no path, but include the extension; for example, GoFish4.App
		lcFolderName  = .FolderName
		lcFullAppName = .FullAppName

		Text To .Code Noshow Textmerge
			oShell = Createobject("wscript.shell")
*			oShell.Run("http://groups.google.com/group/foxprogofish")
			oShell.Run("https://github.com/VFPX/GoFish/issues")
		Endtext

* Optional
		.StatusBarText = m.lcCommonText
		.Summary       = m.lcCommonText && if empty, first line of .Description is used

* For public tools, such as PEM Editor, etc.
		.Source  = "GoFish" && e.g., "PEM Editor"
		.Version = "" && e.g., "Version 7, May 18, 2011"
		.Sort    = 2 && the sort order for all items from the same .Source
		.Link    = "https://github.com/VFPX/GoFish"

		llRegister = .Register()

	Endwith

	Return m.llRegister
Endproc

*-----------------------------------------------------------------------------------------
Procedure CreateVersionFile

	Local;
		lcFileData    As String,;
		lcOldVersionFile As String,;
		lcVersionFile As String,;
		loGoFish      As Object

*-- Rebuild Version File ------------------
	loGoFish = _Screen._GoFish

*-- Delete an older verison filename that is no longer used.
	lcOldVersionFile = Addbs(m.loGoFish.cAppPath) + "GoFishVersion.txt"
	Delete File (m.lcOldVersionFile)

*== Write the current version to a file ========================
	lcVersionFile = Addbs(m.loGoFish.cAppPath) + VERSION_LOCAL_FILE

	lcFileData = GOFISH_APP_NAME + Chr(13) + Chr(10) + ;
		GOFISH_VERSION_STRING_FOR_VERSION_FILE + Chr(13) + Chr(10) + ;
		GOFISH_DOWNLOAD_URL

*-- Delete current file
	Delete File (m.lcVersionFile)

*-- Re-create local version file
	Strtofile(m.lcFileData, m.lcVersionFile, 0)

Endproc

Procedure HelpScreen
	Local;
		lcText As String

	Text To m.lcText Noshow
DO GoFish5.APP WITH ["/?"]|["-Reset"]|["-ResetLocal"]|["-Clear"]
An advanced code search tool for MS Visual Foxpro 9

PARAMETERS (only one parameter)
/?                  This help screen
-Reset           Reset the common settings in HOME(7)+"\GoFish_"
-ResetLocal  Reset the settings local to the ressource file, if used
-Clear           Delete stored searches and replace data

DO GoFish5.APP WITH ["-P"]|["-F"]|["cFolder"]|["cProject"]
Initial scope for the search, or the restore history:
-P             Use the active project as scope
-F             Use the active folder as scope
cProject   A project as scope
cFolder    A folder as scope

IF no parameter is given, the scope on startup will depend on settings.
	Endtext &&lcText

	Messagebox(m.lcText, 0, "GoFish parameters")
Endproc &&HelpScreen

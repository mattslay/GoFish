#INCLUDE BuildGoFish.h
#Define dcGoFishName	GOFISH
#Define dlAllowDevMode	.F.


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
	loSettings        As Object

lcInitialSource = Evl(m.tcInitialSource, "")

If !SetupEnvironment() Then
	Return .F.
Endif &&!SetupEnvironment()

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
		lcApp     As String,;
		lcAppName As String,;
		lcAppPath As String,;
		lnSys16   As Integer,;
		loGoFish  As Object,;
		loRegexp  As Object

*-- 4.2.003 - Need to clear out these class definitions which may be cached by VFP from an
*-- older version of GoFish
	Clear Class "GoFishSearchEngine"
	Clear Class "GoFishSearchOptions"
	Clear Class "GF_PEME_BaseTools"

	lnSys16   = 1
	lcApp     = Sys(16, m.lnSys16)
	lcAppPath = Getwordnum(m.lcApp, 3)
	lcAppName = Justfname(m.lcAppPath)
	lcAppPath = Addbs(Justpath(m.lcAppPath))
	Do While !Empty(m.lcApp) And !m.lcAppName==[dcGoFishName.APP] And !m.lcAppName==[GOFISH.FXP]
		lnSys16   = m.lnSys16+1
		lcApp     = Sys(16, m.lnSys16)
		lcAppPath = Getwordnum(m.lcApp, 3)
		lcAppName = Justfname(m.lcAppPath)
		lcAppPath = Addbs(Justpath(m.lcAppPath))
	Enddo &&!Empty(m.lcApp) And !m.lcAppName==[dcGoFishName.APP] And !m.lcAppName==[GOFISH.FXP]

	Do Case
		Case Empty(m.lcApp)
* nothing found
			Messagebox("Error starting.", 0, GOFISH_APP_NAME)
			Return .F.
		Case Justext(m.lcAppName)=="APP"
* all fine, app
		Case !dlAllowDevMode
*we can not go into DevelopmentMode
			Messagebox("Error starting. DevelopmentMode not allowed.", 0, GOFISH_APP_NAME)
			Return .F.
		Case Justext(m.lcAppName)=="FXP"
*DevelopmentMode
			SetPathsForDevelopmentMode(m.lcAppPath)
			lcAppName = GOFISH_APP_FILE
		Otherwise
*some error	  
			Messagebox("Error starting.", 0, GOFISH_APP_NAME)
			Return .F.
	Endcase

	loRegexp = SF_RegExp(m.lcAppPath+'SF_RegExp')

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
	Set Path To (m.tcAppPath + "SF_RegExp") Additive

Endproc

*----------------------------------------------------------------------------
Procedure HelpScreen
	Local;
		lcText As String

	Text To m.lcText Noshow
DO GoFish.APP WITH ["/?"]|["-Reset"]|["-ResetLocal"]|["-Clear"]
An advanced code search tool for MS Visual Foxpro 9

PARAMETERS (only one parameter)
/?                  This help screen
-Reset           Reset the common settings in HOME(7)+"\GoFish_"
-ResetLocal  Reset the settings local to the ressource file, if used
-Clear           Delete stored searches and replace data

DO GoFish.APP WITH ["-P"]|["-F"]|["cFolder"]|["cProject"]
Initial scope for the search, or the restore history:
-P             Use the active project as scope
-F             Use the active folder as scope
cProject   A project as scope
cFolder    A folder as scope

IF no parameter is given, the scope on startup will depend on settings.
	Endtext &&lcText

	Messagebox(m.lcText, 0, "GoFish parameters")
Endproc &&HelpScreen

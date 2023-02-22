#INCLUDE BuildGoFish.h

*---------------------------------------------------------------------------------------------
Lparameters tcInitialSource, tlSuppressRegisteredWithThorDialog


Local lcInitialSource, lcXMLFiles, llReturn
Local;
	lcSettingsFile As String,;
	lcSettingsFile As String,;
	loMy           As 'My' Of 'My.vcx',;
	loSettings


lcInitialSource = Evl(tcInitialSource, '')

SetupEnvironment()

*!* ******************** Removed 11/03/2015 *****************
*!* CreateVersionFile() && Version file is re-built every time

If !Empty(tcInitialSource) And Upper(tcInitialSource) = 'THOR'
	llReturn = TryRegisterWithThor(.T., !tlSuppressRegisteredWithThorDialog)
	Return llReturn
Else
	TryRegisterWithThor() && GoFish is re-registered with Thor every time
Endif

*-- Only allow one instance of the form to be running
If Type('_Screen._GoFish.oResultsForm') = 'O' And !Isnull(_Screen._GoFish.oResultsForm)
	If _Screen._GoFish.oResultsForm.WindowState = 1
		_Screen._GoFish.oResultsForm.WindowState = 0
	Else
		_Screen._GoFish.oResultsForm.Show()
	Endif

	Return
Endif

*-- Erase all existing GF XML files to reset to default
If Upper(Alltrim(lcInitialSource)) = 'RESET'
	lcXMLFiles = '"' + Home(7) + 'GF_*.xml' + '"'
	Delete File (lcXMLFiles)
	lcInitialSource = ''
Endif

Set Procedure To 'GoFishProc' Additive
Set Procedure To 'mhHtmlCode' Additive
Set Procedure To 'GoFishSearchEngine.prg' Additive
Set Procedure To 'GoFishSearchOptions.prg' Additive
Set Procedure To 'GF_PEME_BaseTools.prg' Additive

loMy = Newobject('My', 'My.vcx')
loSettings     = loMy.Settings
lcSettingsFile = Addbs(Home(7)) + 'GF_Results_Form_Settings.xml'
* SF 20221017 test for file
If File(m.lcSettingsFile)
	m.loSettings.Load(m.lcSettingsFile)
Endif
* /SF 20221017 test for file

* SF 20221017
* special local settings
If m.loSettings.Exists('lCR_Allow') And m.loSettings.lCR_Allow Then
	GF_Get_LocalSettings(@loSettings,m.lcSettingsFile)

Endif &&m.loSettings.EXISTS('lCR_Allow') And m.loSettings.lCR_Allow
*/ SF 20221017

If Pemstatus(m.loSettings, 'lDesktop', 5) And m.loSettings.lDesktop
	Do Form GoFish_Results_Desktop With lcInitialSource
Else
	Do Form GoFish_Results With lcInitialSource
Endif Pemstatus(m.loSettings, 'lDesktop', 5) ...

If Version(2) = 0 && If running as an .EXE, setup Read Events loop
	On Shutdown Clear Events
	Read Events
	On Shutdown

Endif

*----------------------------------------------------------------------------
Procedure SetupEnvironment


Local loGoFish
Local lcAppName, lcAppPath

*-- 4.2.003 - Need to clear out these class definitions which may be cached by VFP from an
*-- older version of GoFish
Clear Class 'GoFishSearchEngine'
Clear Class 'GoFishSearchOptions'
Clear Class 'GF_PEME_BaseTools'

For x = Program(-1) To 1 Step -1 && Look up through run stack to find the name of the running .APP file
	lcAppName = Sys(16, x)
	If '.APP' $ Upper(lcAppName)
		Exit
	Else
		lcAppName = ''
	Endif
Endfor

lcAppName = Evl(lcAppName, Sys(16, 1))
lcAppPath = Addbs(Justpath (Getwordnum(lcAppName, 3)))
If Empty(lcAppPath)
	lcAppPath = Addbs(Justpath(lcAppName))
Endif
lcAppName = Justfname(lcAppName)

If '.FXP' $ lcAppName
	SetPathsForDevelopmentMode(lcAppPath)
	lcAppName = GOFISH_APP_FILE
Endif

If Type('_Screen._GoFish.oResultsForm') = 'O' And !Isnull(_Screen._GoFish.oResultsForm)
	Return
Else

*-- 4.2.003 - Create this object locally, rather than getting if from the VCX
*-- This is better since finding the VCX during a Thor installation of GoFish can
*-- be tricky.
	loGoFish = Createobject('Empty')
	AddProperty(loGoFish, 'cAppPath', lcAppPath)
	AddProperty(loGoFish, 'cAppName', lcAppName)
	AddProperty(loGoFish, 'cVersion', GOFISH_VERSION)
	AddProperty(loGoFish, 'cBuildDate', GOFISH_BUILDDATE)
	AddProperty(loGoFish, 'dBuildDate', GOFISH_dBUILDDATE)
	AddProperty(loGoFish, 'oResultsForm', .Null.)

	_Screen.AddProperty('_GoFish', loGoFish) && Add this object onto _Screen, so it can accessed from GoFish forms

Endif

Endproc

*--------------------------------------------------------------------------------
Procedure SetPathsForDevelopmentMode(tcAppPath)

*-- If running this bootstrap in dev mode,we need to setup these paths
*--  Note: This is not required when the compiled .app file is running
Set Path To (tcAppPath) Additive
Set Path To (tcAppPath + 'Prg') Additive
Set Path To (tcAppPath + 'Forms') Additive
Set Path To (tcAppPath + 'Lib') Additive
Set Path To (tcAppPath + 'Lib\VFP\My') Additive
Set Path To (tcAppPath + 'Lib\VFP\FFC') Additive
Set Path To (tcAppPath + 'Images') Additive
Set Path To (tcAppPath + 'Menus') Additive

Endproc

*----------------------------------------------------------------------------
Procedure TryRegisterWithThor(tlForcedRegister, tlShowConfirmationDialog)

Local lcThorVersion, llReturn, llThorPresent


Try && See if Thor is running
		lcThorVersion = Execscript(_Screen.cThorDispatcher, "Version=")
		llThorPresent = .T.
	Catch
		If tlShowConfirmationDialog
			Messagebox('Thor is not presently running on your system. Please install/run it first.', 0, 'Thor required.')
		Endif
		llThorPresent = .F.
	Finally
Endtry

If !llThorPresent
	Return .F.
Endif

llReturn = RegisterWithThor()

If !tlForcedRegister
	Return
Endif

If tlShowConfirmationDialog
	If llReturn
		Messagebox('Successfully registered with Thor.', 0, GOFISH_APP_NAME)
	Else
		Messagebox('Error attempting to register with Thor.', 0, GOFISH_APP_NAME)
	Endif
Endif

Return llReturn

Endproc

*------------------------------------------------------------------------------------------------------
Procedure RegisterWithThor

Local llRegisterApp, llRegisterUpdater

llRegisterApp = RegisterAppWithThor()

llRegisterUpdater = RegisterUpdaterWithThor()

RegisterVfpxLinkWithThor()
RegisterDiscussionGroupWithThor()

Return (llRegisterApp And llRegisterUpdater)

Endproc


*--------------------------------------------------------------------------------------
Procedure RegisterAppWithThor()

Local lcFolderName, lcFullAppName, loThorInfo, lcCode, lcPlugIn

Try
		loThorInfo = Execscript (_Screen.cThorDispatcher, 'Thor Register=')
	Catch
		loThorInfo = .Null.
Endtry

If Isnull (loThorInfo)
	Return .F.
Endif


With loThorInfo
* Required
	.Prompt		 = GOFISH_APP_NAME && used when tool appears in a menu
	.Description = 'Advanced Code Search Tool' && may be lengthy, including CRs, etc
	.Author 	 = 'Matt Slay'
	.PRGName	 = THOR_TOOL_NAME  && a unique name for the tool; note the required prefix
	.Category 	 = 'Applications|GoFish'

	.AppName     = GOFISH_APP_FILE  && no path, but include the extension; for example, GoFish5.App
* Note that this creates This.FullAppName, which determines the full path of the APP
* and also This.FolderName

*	.FolderName = 'C:\Visual FoxPro\Programs\GoFish All Versions\GoFish5\Source'
*	.FullAppName = .FolderName + '\' + GOFISH_APP_FILE

	lcFolderName = .FolderName
	lcFullAppName = .FullAppName

	.PlugInClasses   = 'clsGoFishFormatGrid'
	.PlugIns		 = 'GoFish Results Grid'

	TEXT To lcCode Noshow Textmerge
  		Do '<<lcFullAppName>>'

	ENDTEXT

	TEXT To lcPlugIn Noshow

EndProc


Define Class clsGoFishFormatGrid As Custom

	Source				= 'GoFish5'
	PlugIn				= 'GoFish Results Grid'
	Description			= 'Provides access to GoFish results grid to set colors and other dynamic properties.'
	Tools				= 'GoFish5'
	FileNames			= 'Thor_Proc_GoFish_FormatGrid.PRG'
	DefaultFileName		= '*Thor_Proc_GoFish_FormatGrid.PRG'
	DefaultFileContents	= ''

	Procedure Init
		****************************************************************
		****************************************************************
		***TEXT*** To This.DefaultFileContents Noshow
Lparameters toGrid, tcResultsCursor

*-- Sample 1: Dynamic row coloring as used by GF
Local lcFileNameColor, lcPRG, lcPRGColor, lcSCX, lcSCXColor, lcVCX, lcVCXColor

lcSCX	   = 'Upper(' + tcResultsCursor + '.filetype) = "SCX"'
lcSCXColor = 'RGB(0,0,128)'

lcVCX	   = 'Upper(' + tcResultsCursor + '.filetype) = "VCX"'
lcVCXColor = 'RGB(0,128,0)'

lcPRG	   = 'Upper(' + tcResultsCursor + '.filetype) $ "PRG TXT H INI XML HTM HTML ASP ASPX"'
lcPRGColor = 'RGB(255,0,0)'

toGrid.SetAll('DynamicForeColor', 'Iif(' + m.lcSCX + ', ' + m.lcSCXColor + ', ' +		;
	  'Iif(' + m.lcVCX + ', ' + m.lcVCXColor + ', ' +									;
	  'Iif(' + m.lcPRG + ',' + m.lcPRGColor + ', RGB(0,0,0))' +							;
	  ')' +																				;
	  ')', 'COLUMN')
Return

*-- Sample 2: Alternative provided by Jim R. Nelson
*-- Assigns row colors based on field MatchType
#Define ccBolds "<Method>", "<Procedure>", "<Function>", "<Constant>", "<<Class Def>>", "<<Method Def>>", "<<Property Def>>"
#Define ccPropertyName "<Property Name>"
#Define ccPropertyValue "<Property Value>"

Local lcBolds, lcCode, lcCodeColor, lcComments, lcCommentsColor, lcFileName, lcOthersColor
Local lcPropNameColor, lcPropertyName, llRegister

lcPropertyName = [0 # Atc("<Property", ] + m.tcResultsCursor + [.MatchType)]
lcComments	   = [0 # Atc("<Comment", ] + m.tcResultsCursor + [.MatchType)]
lcFileName	   = [0 # Atc("<File", ] + m.tcResultsCursor + [.MatchType)]
lcCode		   = 'Left(' + m.tcResultsCursor + '.MatchType, 1) # "<"'

* ForeColor
lcPropNameColor	= 'RGB(0,128,0)'
lcCommentsColor	= 'RGB(0,0,0)'
lcCodeColor		= 'RGB(0,0,0)'
lcOthersColor	= 'RGB(0,0,255)'

m.toGrid.SetAll ('DynamicForeColor', 'ICase(' +					;
	  m.lcPropertyName + ', ' + m.lcPropNameColor + ', ' +		;
	  m.lcComments + ', ' + m.lcCommentsColor + ', ' +			;
	  m.lcCode + ', ' + m.lcCodeColor + ', ' +					;
	  m.lcOthersColor + ')')

* BackColor
lcCommentsColor	= 'RGB(192,192,192)'
lcFileNameColor	= 'Rgb(176, 224, 230)'
m.toGrid.SetAll ('DynamicBackColor', 'ICase(' +				;
	  m.lcComments + ', ' + m.lcCommentsColor + ', ' +		;
	  m.lcFileName + ', ' + m.lcFileNameColor + ', ' +		;
	  ' Rgb(255,255,255))')

* Bold
lcBolds		 = [Inlist(] + m.tcResultsCursor + [.MatchType, ccBolds)]
m.toGrid.SetAll ('DynamicFontBold', m.lcBolds)

Return
		End***TEXT***
		endproc
	enddefine
	ENDTEXT
	.Code = Strtran(lcCode + Evl(lcPlugIn, ''), '***TEXT***', 'Text')

* Optional
	.StatusBarText = GOFISH_APP_NAME
	.Summary	   = 'Code Search Tool' && if empty, first line of .Description is used
	.Classes	   = 'loGoFish = gofishsearchengine of lib\gofishsearchengine.vcx'

* For public tools, such as PEM Editor, etc.
	.Source	 = 'GoFish' && e.g., 'PEM Editor'
	.Version = GOFISH_VERSION  && e.g., 'Version 7, May 18, 2011'
	.Sort	 = Iif('BETA' $ Upper(GOFISH_APP_NAME), 5, 1) && the sort order for all items from the same .Source
	.Link	 = GOFISH_HOME_PAGE

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

Local lcFolderName, lcFullAppName, loThorInfo, lcCommonText, llRegister

Try
		loThorInfo = Execscript (_Screen.cThorDispatcher, 'Thor Register=')
	Catch
		loThorInfo = .Null.
Endtry

If Isnull (loThorInfo)
	Return .F.
Endif

lcCommonText = 'GoFish page on VFPx'

With loThorInfo
* Required
	.Prompt		 = lcCommonText && used when tool appears in a menu
	.Description = lcCommonText && may be lengthy, including CRs, etc
	.Author 	 = 'Matt Slay'
	.PRGName	 = 'Thor_Tool_GoFish_VFPx_Page'  && a unique name for the tool; note the required prefix
	.Category 	 = 'Applications|GoFish'

	.AppName     = GOFISH_APP_FILE && no path, but include the extension; for example, GoFish5.App
	lcFolderName = .FolderName
	lcFullAppName = .FullAppName

	TEXT To .Code Noshow Textmerge
			oShell = Createobject("wscript.shell")
			oShell.Run('https://github.com/VFPX/GoFish')
	ENDTEXT

* Optional
	.StatusBarText = lcCommonText
	.Summary	   = lcCommonText && if empty, first line of .Description is used

* For public tools, such as PEM Editor, etc.
	.Source	 = 'GoFish' && e.g., 'PEM Editor'
	.Version = '' && e.g., 'Version 7, May 18, 2011'
	.Sort	 = 3 && the sort order for all items from the same .Source
	.Link	 = 'https://github.com/VFPX/GoFish'

	llRegister = .Register()

Endwith

Return llRegister
Endproc

*--------------------------------------------------------------------------------------
Procedure RegisterDiscussionGroupWithThor()

Local lcFolderName, lcFullAppName, loThorInfo, llRegister, lcCommonText

Try
		loThorInfo = Execscript (_Screen.cThorDispatcher, 'Thor Register=')
	Catch
		loThorInfo = .Null.
Endtry

If Isnull (loThorInfo)
	Return .F.
Endif

lcCommonText = 'GoFish discussion group'

With loThorInfo
* Required
	.Prompt		 = lcCommonText && used when tool appears in a menu
	.Description = lcCommonText && may be lengthy, including CRs, etc
	.Author 	 = 'Matt Slay'
	.PRGName	 = 'Thor_Tool_GoFish_Discussion_Group'  && a unique name for the tool; note the required prefix
	.Category 	 = 'Applications|GoFish'

	.AppName     = GOFISH_APP_FILE && no path, but include the extension; for example, GoFish4.App
	lcFolderName = .FolderName
	lcFullAppName = .FullAppName

	TEXT To .Code Noshow Textmerge
			oShell = Createobject("wscript.shell")
			oShell.Run('http://groups.google.com/group/foxprogofish')
	ENDTEXT

* Optional
	.StatusBarText = lcCommonText
	.Summary	   = lcCommonText && if empty, first line of .Description is used

* For public tools, such as PEM Editor, etc.
	.Source	 = 'GoFish' && e.g., 'PEM Editor'
	.Version = '' && e.g., 'Version 7, May 18, 2011'
	.Sort	 = 2 && the sort order for all items from the same .Source
	.Link	 = 'https://github.com/VFPX/GoFish'

	llRegister = .Register()

Endwith

Return llRegister
Endproc

*-----------------------------------------------------------------------------------------
Procedure CreateVersionFile

Local lcFileData, lcOldVersionFile, lcVersionFile, loGoFish

*-- Rebuild Version File ------------------
loGoFish = _Screen._GoFish

*-- Delete an older verison filename that is no longer used.
lcOldVersionFile = Addbs(loGoFish.cAppPath) + 'GoFishVersion.txt'
Delete File (lcOldVersionFile)

*== Write the current version to a file ========================
lcVersionFile = Addbs(loGoFish.cAppPath) + VERSION_LOCAL_FILE

lcFileData = GOFISH_APP_NAME + Chr(13) + Chr(10) + ;
	GOFISH_VERSION_STRING_FOR_VERSION_FILE + Chr(13) + Chr(10) + ;
	GOFISH_DOWNLOAD_URL

*-- Delete current file
Delete File (lcVersionFile)

*-- Re-create local version file
Strtofile(lcFileData, lcVersionFile, 0)

Endproc

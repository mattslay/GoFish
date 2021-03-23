#INCLUDE BuildGoFish.h

Local lcApplicationName, lcBetaFolder, lcBootStrap, lcBuild, lcBuildDate, lcBuildGoFishHeaderFile
Local lcCloudVersionFileContents, lcDate, lcDownloadUrl, lcErrorFile, lcGoFishHomePage, lcHgData
Local lcIncludeFileText, lcMercurialVersionFile, lcMonthDay, lcProject, lcSourceLoc, lcVersion
Local lcVersionFileContents, lcVersionFileURL, lcVersionLocalFile, lcVersionStringForVersionFile
Local lcAppFile, lcWindowTitle, lcYear, llBuildAsBeta, lnCurrentVersion, lcSafety, lcCloudVersionFile

*!* ******************** Removed 10/27/2015 *****************
*!* Release All
*!* Clear All

lcSafety = Set('Safety')
Set Safety Off

lcSourceLoc =  Addbs(JustPath(Sys(16))) && Change this to location of source files
Cd (lcSourceLoc)

*== Most of these local variables will make their way into BuildGoFish.h
lcVersion = '5.0'
lcBuild = '167' && <---- Set your desired version level of the tool here
								&& This will be used on the GoFish form and will live on the
								&& the _GoFish object that gets attached to _Screen at run time.

lcApplicationName	= 'GoFish5'
lcVersionLocalFile	= 'GoFishVersionFile.txt'
lcVersionFileURL	= 'http://foxpro.mattslay.com/GoFish5/_GoFishVersionFile.txt'
lcDownloadUrl		= 'http://bit.ly/GoFish5'
lcGoFishHomePage	= 'https://github.com/mattslay/GoFish'

*==== Provision for building as a Beta version ============================
llBuildAsBeta = .f.

If llBuildAsBeta = .t.
	lcApplicationName = lcApplicationName + '_Beta'
	lcBetaFolder = 'http://foxpro.mattslay.com/GoFish4/Beta/' 
	lcVersionFileURL	= lcBetaFolder + '_GoFishVersionFile.txt'
	lcDownloadUrl		= lcBetaFolder + lcApplicationName + '.zip'
EndIf

* This will appear as: '5.X.YYY' on the form, and in the cVersion property on the GoFishSeaerch Engine class


lcWindowTitle = 'Building ' + lcApplicationName
*!* ******************** Removed 11/01/2015 *****************
*!* Define Window talkWindow From 1, 1 To 75, 100 In Screen Float Close Grow Zoom Title lcWindowTitle
*!* Activate Window talkWindow
*!* ******************** Removed 11/01/2015 *****************
*!* Set Talk On Window talkWindow
? 'Building ' + lcApplicationName + '  Ver ' + lcVersion + '  Build ' + lcBuild + '...'
? ' '

*------------------------------------------------------------------------------------
*=================================================================================================

lcVersion = lcVersion + '.' + lcBuild

*-- Create a file in the source file which contains the Version number
lcDate = Transform(Date(), '@YL')
lcMonthDay = Alltrim(GetWordNum(lcDate, 2, ','))
lcYear = Alltrim(GetWordNum(lcDate, 3, ','))
lcBuildDate = lcMonthDay + ', ' + lcYear

*lcVersionStringForVersionFile = lcApplicationName + ' ' + lcVersion + ' ' + lcBuildDate 

lcVersionStringForVersionFile = lcApplicationName + ' - ' + lcVersion + ' - ' + lcBuildDate + ' - ' + Dtoc (Date(), 1)

Text to lcIncludeFileText NoShow TextMerge 

	#DEFINE GOFISH_APP_NAME '<<lcApplicationName>>'
	#DEFINE GOFISH_APP_FILE  '<<lcApplicationName + ".app">>'
	#DEFINE GOFISH_VERSION '<<lcVersion>>'
	#DEFINE GOFISH_BUILDDATE '<<lcBuildDate>>'
	#DEFINE GOFISH_DBUILDDATE Date(<<Year(Date())>>,<<Month(Date())>>, <<Day(Date())>>)
	#DEFINE GOFISH_DOWNLOAD_URL '<<lcDownloadUrl>>'
	#DEFINE GOFISH_VERSION_STRING_FOR_VERSION_FILE '<<lcVersionStringForVersionFile>>'
	#DEFINE THOR_TOOL_NAME	'<<"Thor_Tool_" + lcApplicationName>>'
	#DEFINE VERSON_FILE_URL	'<<lcVersionFileURL>>' 
	#DEFINE VERSION_LOCAL_FILE '<<lcVersionLocalFile>>'
	#DEFINE GOFISH_HOME_PAGE '<<lcGoFishHomePage>>'

EndText

*=== Create the local version File ======================================================
lcVersionFileContents = lcApplicationName + Chr(13) + Chr(10) + ;
						lcVersionStringForVersionFile + Chr(13) + Chr(10) + ;
						lcDownloadUrl 
						
*Delete File (lcVersionLocalFile) 
? "Writing local version file: " + lcVersionLocalFile
StrToFile(lcVersionFileContents, lcVersionLocalFile)


*=== Create the Cloud version file ===========================================
Text to lcCloudVersionFileContents NoShow TextMerge PRETEXT 3

	Lparameters toUpdateInfo

		###Text to lcNote NoShow
			<<FileToStr('Changelog.txt')>>
		###EndText 

		AddProperty(toUpdateInfo, 'AvailableVersion', '<<lcVersionStringForVersionFile>>')
		AddProperty(toUpdateInfo, 'SourceFileUrl', '<<lcDownloadUrl>>')
		AddProperty(toUpdateInfo, 'LinkPrompt', 'GoFish Home Page')
		AddProperty(toUpdateInfo, 'Link', '<<lcGoFishHomePage>>')
		AddProperty(toUpdateInfo, 'Notes', lcNote)

	Execscript (_Screen.cThorDispatcher, 'Result=', toUpdateInfo)
	Return toUpdateInfo

EndText

lcCode = Strtran(lcCloudVersionFileContents, '###', '')
lcCloudVersionFile = '_' + lcVersionLocalFile
? "Writing cloud version file: " + lcCloudVersionFile 
StrToFile(lcCode, lcCloudVersionFile)


*-- Write out the Include file with the current version and build date as contants
lcBuildGoFishHeaderFile = 'BuildGoFish.h'
Delete File (lcBuildGoFishHeaderFile)  
? "Writing include file: " + lcBuildGoFishHeaderFile
StrToFile(lcIncludeFileText, lcBuildGoFishHeaderFile) 

lcBootStrap = 'GoFish.prg'
Compile (lcBootStrap)

lcProject = lcSourceLoc + 'GoFish5.Pjx'

If File(lcProject)
	Build App (lcApplicationName) From (lcProject)
	?
	lcAppFile =  lcApplicationName + '.app'
	Dir (lcAppFile) 
	? Fdate(lcAppFile, 1)

	lcErrorFile = lcApplicationName + '.err'
	If File(lcErrorFile)
		?
		? ' *** ERRORS ****' 
		? FileToStr(lcErrorFile)
	EndIf
	
Else
	MessageBox('Error locating Project file [' + lcProject + ']',16, lcApplicationName + ' - Building App')
Endif
    
Set Safety &lcSafety


           
If Version(4) != "09.00.0000.7423" Then
	Messagebox("Run with VFP9 SP2!")
	Return .F.
Endif &&Version(4)!= "09.00.0000.7423"

CLEAR ALL

Local lcApplicationName, lcBetaFolder, lcBootStrap, lcBuild, lcBuildDate, lcBuildGoFishHeaderFile
Local lcCloudVersionFileContents, lcDate, lcDownloadUrl, lcErrorFile, lcGoFishHomePage, lcHgData
Local lcIncludeFileText, lcMercurialVersionFile, lcMonthDay, lcProject, lcSourceLoc, lcVersion
Local lcVersionFileContents, lcVersionFileURL, lcVersionLocalFile, lcVersionStringForVersionFile
Local lcAppFile, lcWindowTitle, lcYear, llBuildAsBeta, lnCurrentVersion, lcSafety, lcCloudVersionFile

lcSafety = Set('Safety')
Set Safety Off

lcSourceLoc =  Addbs(Justpath(Sys(16))) && Change this to location of source files
Cd (lcSourceLoc)

*== Most of these local variables will make their way into BuildGoFish.h
lcVersion = Alltrim(FileToStr('VersionNumber.txt'), ' ', CHR[9], CHR[13], chr[10]) 
lcBuild   = Alltrim(FileToStr('BuildNumber.txt'), ' ', CHR[9], CHR[13], chr[10]) 
&& This will be used on the GoFish form and will live on the
&& the _GoFish object that gets attached to _Screen at run time.

lcApplicationName	= 'GoFish'
lcVersionLocalFile	= 'GoFishVersionFile.txt'
lcVersionFileURL	= 'https://github.com/VFPX/GoFish/_GoFishVersionFile.txt'
lcDownloadUrl		= 'https://raw.githubusercontent.com/VFPX/GoFish/master/Source/Source.zip'
lcGoFishHomePage	= 'https://github.com/VFPX/GoFish'

* This will appear as: '5.X.YYY' on the form, and in the cVersion property on the GoFishSeaerch Engine class


lcWindowTitle = 'Building ' + lcApplicationName
? 'Building ' + lcApplicationName + '  Ver ' + lcVersion + '  Build ' + lcBuild
?? " (" + lcVersion + '.' + lcBuild + ")..."
? ' '

*------------------------------------------------------------------------------------
*=================================================================================================

lcVersion = lcVersion + '.' + lcBuild

*-- Create a file in the source file which contains the Version number
lcDate = Transform(Date(), '@YL')
lcMonthDay = Alltrim(Getwordnum(lcDate, 2, ','))
lcYear = Alltrim(Getwordnum(lcDate, 3, ','))
lcBuildDate = lcMonthDay + ', ' + lcYear

lcVersionStringForVersionFile = lcApplicationName + ' - ' + lcVersion + ' - ' + lcBuildDate + ' - ' + Dtoc (Date(), 1)

TEXT to lcIncludeFileText NoShow TextMerge

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

ENDTEXT

*=== Create the local version File ======================================================
lcVersionFileContents = lcApplicationName + Chr(13) + Chr(10) + ;
	lcVersionStringForVersionFile + Chr(13) + Chr(10) + ;
	lcDownloadUrl

*=== Create the Cloud version file ===========================================
TEXT to lcCloudVersionFileContents NoShow TextMerge PRETEXT 3

	Lparameters toUpdateInfo

		###Text to lcNote NoShow
			<<Left(Strtran(FileToStr('..\docs\Changelog.md'), Chr[10], Chr[13] + chr[10]), 8000)>>
		###EndText
		
		* For GF, strip off links at top of change log
		lcNote = Substr(lcNote, At('--', lcNote))
		lcNote = Alltrim(lcNote, 1, ' ', '-', Chr(13), Chr(10)) 

		AddProperty(toUpdateInfo, 'AvailableVersion', '<<lcVersionStringForVersionFile>>')
		AddProperty(toUpdateInfo, 'SourceFileUrl', '<<lcDownloadUrl>>')
		AddProperty(toUpdateInfo, 'LinkPrompt', 'GoFish Home Page')
		AddProperty(toUpdateInfo, 'Link', '<<lcGoFishHomePage>>')
		AddProperty(toUpdateInfo, 'Notes', lcNote)

	Execscript (_Screen.cThorDispatcher, 'Result=', toUpdateInfo)
	Return toUpdateInfo

ENDTEXT

lcCode = Strtran(lcCloudVersionFileContents, '###', '')
lcCloudVersionFile = '..\_' + lcVersionLocalFile
? "Writing cloud version file: " + lcCloudVersionFile
Strtofile(lcCode, lcCloudVersionFile)


*-- Write out the Include file with the current version and build date as contants
lcBuildGoFishHeaderFile = 'BuildGoFish.h'
Delete File (lcBuildGoFishHeaderFile)
? "Writing include file: " + lcBuildGoFishHeaderFile
Strtofile(lcIncludeFileText, lcBuildGoFishHeaderFile)

lcBootStrap = 'GoFish.prg'
Compile (lcBootStrap)

lcProject = lcSourceLoc + 'GoFish.Pjx'

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
		? Filetostr(lcErrorFile)
	Endif

Else
	Messagebox('Error locating Project file [' + lcProject + ']',16, lcApplicationName + ' - Building App')
Endif

Set Safety &lcSafety

? ""
? "Generation source code text files with FoxBin2Prg."
? "  Working..."
Do GoFishGenerateSCC.prg
? "Done."


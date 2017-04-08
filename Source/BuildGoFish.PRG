#INCLUDE BuildGoFish.h

Local lcApplicationName, lcBetaFolder, lcBootStrap, lcBuild, lcBuildDate, lcBuildGoFishHeaderFile
Local lcCloudVersionFileContents, lcDate, lcDownloadUrl, lcErrorFile, lcGoFishHomePage, lcHgData
Local lcIncludeFileText, lcMercurialVersionFile, lcMonthDay, lcProject, lcSourceLoc, lcVersion
Local lcVersionFileContents, lcVersionFileURL, lcVersionLocalFile, lcVersionStringForVersionFile
Local lcAppFile, lcWindowTitle, lcYear, llBuildAsBeta, lnCurrentVersion, lcSafety

*!* ******************** Removed 10/27/2015 *****************
*!* Release All
*!* Clear All

lcSafety = Set('Safety')
Set Safety Off

lcSourceLoc =  Addbs(JustPath(Sys(16))) && Change this to location of source files
Cd (lcSourceLoc)

*== Most of these local variables will make their way into BuildGoFish.h
lcVersion = '5.0'
lcBuild = '001' && <---- Set your desired version level of the tool here
								&& This will be used on the GoFish form and will live on the
								&& the _GoFish object that gets attached to _Screen at run time.

*** JRN 10/27/2015 : increment build number
If File('BuildCount.txt')
	lcBuild = Evl(Filetostr('BuildCount.txt'), '001')
EndIf

Strtofile(Transform(1 + Val(m.lcBuild)), 'BuildCount.txt')

lcApplicationName	= 'GoFish5'
lcVersionLocalFile	= 'GoFishVersionFile.txt'
lcVersionFileURL	= 'http://bit.ly/_GoFishVersionFile'
lcDownloadUrl		= 'http://bit.ly/GoFish4Zip'
lcGoFishHomePage	= 'http://vfpx.codeplex.com/wikipage?title=GoFish'

*==== Provision for building as a Beta version ============================
llBuildAsBeta = .f.

If llBuildAsBeta = .t.
	lcApplicationName = lcApplicationName + '_Beta'
	lcBetaFolder = 'http://foxpro.mattslay.com/GoFish4/Beta/' 
	lcVersionFileURL	= lcBetaFolder + '_GoFishVersionFile.txt'
	lcDownloadUrl		= lcBetaFolder + lcApplicationName + '.zip'
EndIf

* This will appear as: '4.0.001' on the form, and in the cVersion property on the GoFishSeaerch Engine class

*=================================================================================================
*------------------------------------------------------------------------------------
* Message from GoFish 4 developer, Matt Slay, 2011-05-13:
* During the development of GoFish4, I used Mercurial Verion Control repository to 
* track changes to the source code and provide backup at each state of development.
* (You can find the public repo here: https://bitbucket.org/mattslay/gofish4)
*
* So, each time I compiled and distribute the app, I wanted to syncronize the version number
* displayed on the forms with the revision level of the source code in the repositoty.
* I discovered where I could read the verision number from the Mercurial store on my machine
* and include this information in the compiled app automatically.
*
* However, if you are working with this source code at some time in the future, you do not
* have to use Mercurial Version Control to edit the source code. It will not affect the use
* or editing of these FoxPro files at all. 
*
* The code below looks for the Mercurial repository within in the source code path,
* and if it does not find it, it just uses default strings set at the top for the version level
* of the build. This way, you can manually set the version to whatever you want it to say.
*
*If it does find Mercurial, the current version level from the repo will be used.
*-------------------------------------------------------------------------------------------
lcMercurialVersionFile = lcSourceLoc + '.hg\cache\tags'

If File(lcMercurialVersionFile)
	lcHgData = FileToStr(lcMercurialVersionFile)
	lnOffsetFromCurrentBuild = 0
	lnCurrentVersion = Int(Val(GetWordNum(lcHgData , 1))) + lnOffsetFromCurrentBuild 
	lcBuild = Padl(lnCurrentVersion + 1, 3, '0')
EndIf

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
		AddProperty(toUpdateInfo, 'Link', 'http://vfpx.codeplex.com/wikipage?title=GoFish')
		AddProperty(toUpdateInfo, 'Notes', lcNote)

	Execscript (_Screen.cThorDispatcher, 'Result=', toUpdateInfo)
	Return toUpdateInfo

EndText

lcCode = Strtran(lcCloudVersionFileContents, '###', '') 

StrToFile(lcCode, '_' + lcVersionLocalFile)


*-- Write out the Include file with the current version and build date as contants
lcBuildGoFishHeaderFile = 'BuildGoFish.h'
Delete File (lcBuildGoFishHeaderFile)  
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
	
	*!* ******************** Removed 11/01/2015 *****************
	*!* ? ''
	*!* ? 'Version file contents:'
	*!* ? '============================================='
	*!* ? FileToStr(lcVersionLocalFile)
	*!* ? ''
	*!* ? 'Cloud Version file contents:'
	*!* ? '============================================='
	*!* ? FileToStr('_' + lcVersionLocalFile)
	

Else
	MessageBox('Error locating Project file [' + lcProject + ']',16, lcApplicationName + ' - Building App')
Endif
    
Set Safety &lcSafety


           
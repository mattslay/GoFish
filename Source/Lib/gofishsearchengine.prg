#Include GoFish.h

Define Class GoFishSearchEngine As Custom

	cCR_StoreLocal                   = Addbs(Home(7) + 'GoFish_')

	cBackupPRG   = 'GoFishBackup.prg'
	cFilesToSkip = ''

* This string contains a list of files to be skipped during the search.
* One filname on each line. This list is only skipped if lSkipFiles is .t.
	cFilesToSkipFile                 = 'GF_Files_To_Skip.txt'

	cGraphicsExtensions = 'PNG ICO JPG JPEG TIF TIFF GIF BMP MSK CUR ANI'
	cInitialDefaultDir  = ''

* A text list of projects that matches oProjects. Makes looking for existing projects fast than analyzing
* the oProjects collection. This property is only to be used by the class. Please don't touch it.
	cProjects                        = ''

* Holds the code for a UDF to be used on Replace operations if nReplaceMode is the Advanced Replace mode
	cReplaceUDFCode                  = ''

* The default Search Options class to be used. Can be overriden by  passing a string to the Init() method.
	cSearchOptionsClass              = 'GoFishSearchOptions'

* This is the name of the cursor where the results rows ill be stored.
	cSearchResultsAlias              = 'GFSE_SearchResults'

* These are the filetypes that will be handled by the SearchInTable() method. All other filetypes are
* assumed to be plain text files and will be handled by the SearchInTextFile() method.
	cTableExtensions                 = 'SCX VCX FRX MNX LBX DBC PJX DBF'

	cVersion                         = ''

* Indicates if the ESC key was pressed by the user during on of the lower processing loops of a Search.
	lEscPress                        = .F.

* This flag is set any time a MODIFY operation is launched from the grid. We asume they made changes
* to the file, requiring a new search before they can do a Replace. This prevents the REPLACE button from
* being available until the search is run again.
	lFileHasBeenEdited               = .F.

* Will indicate if there were any files in a ProcessProject() or
* ProcessPath() that were not found. You can check this flag after a earch call.
	lFileNotFound                    = .F.

	lReadyToReplace                  = .F.

* Indicates if the max search results limit was reach during a search. See nMaxResults property on the Search Options class.
	lResultsLimitReached             = .F.

	lTimeStampDataProvided           = .F.
* This values indicates how many files were found to have matches in them.
	nFileCount                       = 0

* This values indicates the total number of files processed that matched the file filter, whether they had matches or not.
	nFilesProcessed                  = 0

* How many matched lines found in the last search. Note: this counts lines that had a match, note each match.
* It's possible that one line could have multiple matches.
	nMatchLines                      = 0

	nReplaceCount                    = 0

	nReplaceFileCount                = 0

* Store the current ID number of the Replace History record.
	nReplaceHistoryId                = 0

* 1 = Regular replace, 2 = Advanced Replace (UDF Replace)
	nReplaceMode                     = 1

* Tells how long the last search took. It is only reset by SearchInPath() or SerachInProject() took,
* and not by the lower level search like SearchInFile, or SearchInTextFile, or SearchInTable.
	nSearchTime                      = 0

	nWildCardFilesToSkip             = 0

* Internal DateTime of Search
	tRunTime                         = Datetime()
* Internal unique ID of Search
	cUni                             = "_"


* Internally used by the SearcthInPath() to build a collection of directories to be searched.
	oDirectories                     = .Null.

* An FFC class used to generate a TimeStamp so the TimeStamp field can be updated when replacing code in a table based file.
	oFrxCursor                       = .Null.

	oFSO                             = .Null.

	oProgressBar                     = .Null.

* Internally created to show a collection of recently used Projects, or projects found in the current path folder.
* This is built so the GoFish Advanced form can allow user to choose a Project.
	oProjects                        = .Null.

	oRegExForProcedureStartPositions = .Null.

	oRegExForSearch                  = .Null.

	*** JRN 2024-02-14 : Separate RegEx for searching in code:
	* for plain searches, identical to oRegExForSearch
	* for wild card searches, only searches for everything up to the *
	oRegExForSearchInCode            = .Null.

	oRegExForCommentSearch           = .Null.

	oReplaceErrors                   = .Null.

* This is a collection of match objects from the last search. Must set  lCreateResultsCollection if you want this collection to be built.
	oResults                         = .Null.

* A collection of any errors that happened during the last search.
	oSearchErrors                    = .Null.

* An object instance of the Search Options class that holds properties  to controll how the search is performed.
	oSearchOptions                   = .Null.

	Dimension aMenuStartPositions[1]
	Dimension aWildcardFiles[1]

	*** JRN 2024-02-14 : Used if wild cards using whole word search
	cWholeWordSearch = ''

*----------------------------------------------------------------------------------
	Procedure AddColumn(tcTable, tcColumnName, tcColumnDetails)

		Local;
			lcAlias As String

		lcAlias = Juststem(m.tcTable)
		Try
				Alter Table (m.tcTable) Add Column &tcColumnName &tcColumnDetails
			Catch
		Endtry

	Endproc


*----------------------------------------------------------------------------------
	Procedure AddFieldToReplaceTable(lcTable, lcCsr, lcFieldName, lcDataType)

		Local;
			llSuccess As Boolean

		If Empty(Field(m.lcFieldName, m.lcCsr))
			Use In (Juststem(m.lcTable)) && Close main table, so Alter Table in next called method can get Exclusive use
			Try
					Select 0
					Use (m.lcTable) Exclusive
					llSuccess = .T.
				Catch
					llSuccess = .F.
			Endtry

*-- Migrate up to version 4.3.022 (circa 2012-06-30 ---------------------------
			If m.llSuccess
				This.AddColumn(m.lcTable, m.lcFieldName, m.lcDataType)
			Endif
		Endif

	Endproc


*----------------------------------------------------------------------------------
	Procedure AddProject(tcProject)

		Local;
			llAlreadyInCollection As Boolean

		llAlreadyInCollection = Atline(Upper(m.tcProject), Upper(This.cProjects)) <> 0

		If !m.llAlreadyInCollection
			This.oProjects.Add(Lower(m.tcProject))
			This.cProjects = This.cProjects + m.tcProject + Chr(13)
		Endif

	Endproc


*----------------------------------------------------------------------------------
	Procedure AssignMatchType(toObject)

		Local;
			lcFileType           As String,;
			lcMatchType          As String,;
			lcName               As String,;
			lcTrimmedMatchLine   As String,;
			lcValue              As String,;
			llError              As Boolean,;
			llNameHasDot         As Boolean,;
			llWorkingOnClassFromVCX As Boolean,;
			loLineMatches        As Object,;
			loNameMatches        As Object,;
			loValueMatches       As Object

		lcFileType = Upper(m.toObject.UserField.FileType)

		lcTrimmedMatchLine        = This.TrimWhiteSpace(m.toObject.MatchLine)&& Trimmed version for display in the grid
		toObject.TrimmedMatchLine = m.lcTrimmedMatchLine

*-- We read MatchType of UserField, but from here on, until the result row is created, we will
*-- move this value to toObject.MatchType, and do some tweaking on it to make it the right value.
*-- We'll never change the value that was passed in on toObject.UserField.MatchType
		lcMatchType        = m.toObject.UserField.MatchType
		toObject.MatchType = m.lcMatchType
*=============================================================================================
* This area contains a few overrides that I've had to build in to make final tweeks on columns
*=============================================================================================
*-- Sometimes in a VCX/SCX the MethodName will be empty and MatchLine will contain the PROCEDURE name
		If Empty(m.toObject.MethodName) And Upper(Getwordnum(m.lcTrimmedMatchLine, 1)) = 'PROCEDURE'  and GetWordNum(lcTrimmedMatchLine, 2) # '='
			toObject.MethodName = Getwordnum(m.lcTrimmedMatchLine, 2)
		Endif

		If !Empty(m.toObject.MethodName)
			With m.toObject.UserField
				If '.' $ m.toObject.MethodName
					._Name              = Alltrim(._Name + '.' + This.ExtractObjectName(m.toObject.MethodName), 1, '.')
					toObject.MethodName = Justext(m.toObject.MethodName)

					If ._ParentClass <> ._BaseClass
						._ParentClass = ''
						._BaseClass   = ''
					Else
						.ContainingClass = ''
					Endif
				Else
					.ContainingClass = ''
				Endif

				If !Empty(._Class) && Trim Class name off of the front (only affects VCX results)
					._Name = Strtran(._Name, ._Class + '.', '', 1, 1)
				Endif
			Endwith
		Else
			With m.toObject.UserField
				If !Empty(._Class) && Trim Class name off of the front (only affects VCX results)
					._Name = Strtran(._Name, ._Class + '.', '', 1, 1)
				Endif
			Endwith
		Endif

		If Empty(m.toObject.UserField.ClassLoc)
			toObject.UserField._ParentClass = '' && Affects VCXs. PRGs will be address in these next lines
		Endif

		If m.lcFileType = 'PRG'
			With m.toObject.UserField
				.ContainingClass = ''
				._Class          = m.toObject.oProcedure._ClassName
				._ParentClass    = m.toObject.oProcedure._ParentClass
				._BaseClass      = m.toObject.oProcedure._BaseClass
				If ._Name = ._Class
					._Name = ''
				Endif
				If  Upper('ENDDEFINE') $ Upper(m.toObject.MethodName)
					toObject.MethodName = ''
					._Name              = ''
				Endif
			Endwith

		Endif

		If Upper(m.lcMatchType) # 'RESERVED3' And This.IsFullLineComment(m.lcTrimmedMatchLine)
			toObject.MatchType = MATCHTYPE_COMMENT
			This.CreateResult(m.toObject)
			Return .Null. && Exit out, we're done with this record!
		Endif

*=============================================================================================
* Handle a few tweaks on MatchType assignments
*=============================================================================================
		This.ProcessInlineComments(m.toObject)

		Do Case

*-- A TimeStamp only search, with no search expression...
			Case Isnull(m.toObject.oMatch)
				If This.oSearchOptions.lTimeStamp And Empty(This.oSearchOptions.cSearchExpression)
					If Empty(m.toObject.UserField._Name) And Empty(m.toObject.UserField.ContainingClass) And Empty(m.toObject.UserField._Class)
						toObject.MatchType =  MATCHTYPE_FILEDATE
					Else
						toObject.MatchType =  MATCHTYPE_TIMESTAMP
					Endif
				Else
					toObject.MatchType = MATCHTYPE_FILENAME
				Endif

			Case Inlist(m.lcFileType, 'SCX', 'VCX', 'FRX')&& And lcMatchType # MATCHTYPE_FILENAME
				This.AssignMatchTypeForScxVcx(m.toObject)

			Case m.lcFileType = 'PRG'
				This.AssignMatchTypeForPrg(m.toObject)

		Endcase

*-- Read MatchType back off toObject for a final bit of tweaking...
		lcMatchType = m.toObject.MatchType

		Do Case
			Case Inlist(m.lcFileType, 'MPR') && not good to replace
				lcMatchType = MATCHTYPE_MPR
			Case Empty(m.lcMatchType)
				lcMatchType = MATCHTYPE_CODE

			Case Upper(Getwordnum(m.lcTrimmedMatchLine, 1)) = '#DEFINE'
				lcMatchType = MATCHTYPE_CONSTANT

			Case m.lcMatchType = MATCHTYPE_PROPERTY_DESC Or m.lcMatchType = MATCHTYPE_PROPERTY_DEF
				toObject.UserField.ContainingClass = ''
				toObject.UserField._Name           = ''
				toObject.MethodName                = Getwordnum(m.toObject.MatchLine, 1, ' ')

			Case m.lcMatchType = MATCHTYPE_PROPERTY

				If Atc('=', m.lcTrimmedMatchLine) = 0
					toObject.MatchType = MATCHTYPE_CODE
					Return m.toObject
				Endif

				lcName              = Getwordnum(m.lcTrimmedMatchLine, 1, ' =') && The Property Name only
				toObject.MethodName = m.lcName

				Try
						If Atc('.', m.lcName) > 0 && Could be ObjectName.ObjectName.ObjectName.PropertyName
							lcName       = Justext(m.lcName) && Need to pick off just the property name, and make sure that's where the match is.
							llNameHasDot = .T.
						Else
							llNameHasDot = .F.
						Endif

*	toObject.UserField.MethodName = lcName
						lcName = m.lcName + ' =' && Need to construct property name like this example:   Caption =

						lcValue        = Alltrim(Substr(m.lcTrimmedMatchLine, 1 + At('=', m.lcTrimmedMatchLine))) && GetWordNum(lcTrimmedMatchLine, 2, '=')
						loNameMatches  = This.oRegExForSearch.Execute(m.lcName)
						loValueMatches = This.oRegExForSearch.Execute(m.lcValue)
* loLineMatches = This.oRegExForSearch.Execute(lcTrimmedMatchLine)
						loLineMatches = This.oRegExForSearch.Execute(m.lcName + m.lcValue)

						With m.toObject.UserField
							If m.llNameHasDot
								If ._ParentClass <> ._BaseClass
									._ParentClass = ''
									._BaseClass   = ''
								Else
									.ContainingClass = ''
								Endif
							Else
								.ContainingClass = ''
							Endif
							If Empty(.ClassLoc)
								._ParentClass = ''
							Endif
						Endwith

						Do Case
							Case m.loNameMatches.Count > 0 And m.loValueMatches.Count > 0 && If match on both sides, make an extra call here for the Name
								toObject.MatchType = MATCHTYPE_PROPERTY_NAME
								This.CreateResult(m.toObject)
								lcMatchType = MATCHTYPE_PROPERTY_VALUE
							Case m.loNameMatches.Count > 0 Or m.loValueMatches.Count > 0 && Only matched on one side
								If m.loValueMatches.Count > 0 And This.oSearchOptions.lIgnoreMemberData And Lower(m.lcName) = '_memberdata ='
									llError = .T. && so this is skipped
								Else
									lcMatchType = Iif(m.loNameMatches.Count > 0, MATCHTYPE_PROPERTY_NAME, MATCHTYPE_PROPERTY_VALUE)
								Endif
							Case m.loLineMatches.Count > 0 && Matched SOMEWHERE on the line. Can span " = " this way
*-- No modification to matchtype required. Will record as MATCHTYPE_PROPERTY
							Case m.loNameMatches.Count = 0 And m.loValueMatches.Count = 0 && Possible that there is not match at all, so we record nothing
								llError = .T.
							Otherwise
* lcMatchType = Iif(loNameMatches.count > 0, MATCHTYPE_PROPERTY_NAME, MATCHTYPE_PROPERTY_VALUE)
						Endcase
					Catch
						lcMatchType = MATCHTYPE_CODE && IF anything above failed, then just consider this a regular code match
				Endtry

		Endcase

		If m.llError = .T.
			Return .Null.
		Endif

*-- Wrap MatchType in brackets (if not already set), and if it's not MATCHTYPE_CODE ...
		If m.lcMatchType # MATCHTYPE_CODE And Left(m.lcMatchType, 1) # '<'
			lcMatchType = '<' + m.lcMatchType + '>'
		Endif

		toObject.MatchType = m.lcMatchType

		Return m.toObject

	Endproc


*----------------------------------------------------------------------------------
	Procedure AssignMatchTypeForPrg(toObject)

		Local;
			lcFirstWord     As String,;
			lcMatchType     As String,;
			lcName          As String,;
			lcParams        As String,;
			lcProcedureType As String,;
			lcTrimmedMatchLine As String,;
			lnMatchStart    As Number,;
			lnProcedureStart As Number,;
			loMatch         As Object,;
			loMatches       As Object,;
			loNameMatches   As Object,;
			loParamMatches  As Object,;
			loProcedure     As Object

		loProcedure        = m.toObject.oProcedure
		loMatch            = m.toObject.oMatch
		lcMatchType        = Upper(m.toObject.MatchType)
		lnProcedureStart   = m.loProcedure.StartByte
		lnMatchStart       = Iif(Vartype(m.loMatch) = 'O', m.loMatch.FirstIndex, 0)
		lcTrimmedMatchLine = m.toObject.TrimmedMatchLine

		Do Case
			Case m.lcMatchType = 'CLASS' && Note, this case also handles Properties on a Class...

				lcFirstWord = Upper(Getwordnum(m.lcTrimmedMatchLine, 1))
				If m.lcFirstWord $ 'PROCEDURE'
					lcMatchType = MatchType_Procedure
				Else
					lcMatchType = Iif(m.lnMatchStart = m.lnProcedureStart, MATCHTYPE_CLASS_DEF, MATCHTYPE_PROPERTY)
				Endif
*toObject.MethodName = ''

			Case Inlist(m.lcMatchType, 'METHOD', 'PROCEDURE', 'FUNCTION')

*-- This test looks for matches in on the Procedure Name versus possible parameters:
*-- Ex: PROCEDURE ProcessJob(lcJobNo). )
				If m.lnMatchStart = m.lnProcedureStart
					lcName   = Getwordnum(m.lcTrimmedMatchLine, 1, '(')
					lcParams = Getwordnum(m.lcTrimmedMatchLine, 2, '(')

					loNameMatches  = This.oRegExForSearch.Execute(m.lcName)
					loParamMatches = This.oRegExForSearch.Execute(m.lcParams)

					If m.loNameMatches.Count > 0 And m.loParamMatches.Count > 0 && If match on both sides, make an extra call here for the Name
						toObject.UserField.MatchType = '<' + Proper(m.lcMatchType) + '>'
						This.CreateResult(m.toObject)
						lcMatchType = MATCHTYPE_CODE
					Else
						lcMatchType = Iif(m.loParamMatches.Count > 0, MATCHTYPE_CODE, Proper(m.lcMatchType))
					Endif
				Else
					lcMatchType = MATCHTYPE_CODE
				Endif

			Otherwise
				lcMatchType = m.toObject.MatchType && Restore it back

		Endcase

		toObject.MatchType = m.lcMatchType

	Endproc


*----------------------------------------------------------------------------------
	Procedure AssignMatchTypeForScxVcx(toObject)

		Local;
			lcClass         As String,;
			lcContainingClass As String,;
			lcMatchType     As String,;
			lcMethodName    As String,;
			lcName          As String,;
			lcProcedureType As String,;
			lcPropertyName  As String,;
			lcTrimmedMatchLine As String,;
			lnMatchStart    As Number,;
			lnProcedureStart As Number,;
			loMatches       As Object

		lcMethodName       = m.toObject.MethodName
		lcTrimmedMatchLine = m.toObject.TrimmedMatchLine

		lcProcedureType  = m.toObject.oProcedure.Type
		lnProcedureStart = m.toObject.oProcedure.StartByte

		lnMatchStart = m.toObject.oMatch.FirstIndex
		lcMatchType  = Upper(m.toObject.MatchType)

		With m.toObject.UserField
			lcClass           = ._Class
			lcContainingClass = .ContainingClass
			lcName            = ._Name
		Endwith

		Do Case

			Case Alltrim(m.lcClass) == Alltrim(m.lcTrimmedMatchLine) And !Empty(m.lcClass) And Empty(m.lcName)
				lcMatchType = MATCHTYPE_CLASS_DEF

			Case m.lcMatchType = 'PROCEDURE'
				If m.lnMatchStart = m.lnProcedureStart And !Empty(m.toObject.oProcedure.ParentClass)
					lcMatchType = MatchType_Method
				Else
					lcMatchType = MATCHTYPE_CODE
				Endif

			Case m.lcMatchType = 'PROPERTIES'
				lcMatchType = MATCHTYPE_PROPERTY

			Case m.lcMatchType = 'OBJNAME'
				lcMatchType = Iif(Empty(m.lcName), MatchType_Class, MatchType_Name)

			Case m.lcMatchType = 'CLASS'
				lcMatchType = MatchType_Class

			Case m.lcMatchType = 'RESERVED3'
				If Left(m.lcTrimmedMatchLine, 1) = '*' && A Method Definition line
					lcMethodName        = Substr(m.lcTrimmedMatchLine, 2, Len(Getwordnum(m.lcTrimmedMatchLine, 1)) - 1)
					loMatches           = This.oRegExForSearch.Execute(m.lcMethodName)
					lcMatchType         = Iif(m.loMatches.Count > 0, MATCHTYPE_METHOD_DEF, MATCHTYPE_METHOD_DESC)
					toObject.MethodName = Iif(m.loMatches.Count > 0, m.lcMethodName, '')
				Else && A Property Definition line
					lcPropertyName = Getwordnum(m.lcTrimmedMatchLine, 1)
					If Atc('.', m.lcPropertyName) > 0
						lcPropertyName = Justext(m.lcPropertyName)
					Endif
					loMatches   = This.oRegExForSearch.Execute(m.lcPropertyName)
					lcMatchType = Iif(m.loMatches.Count > 0, MATCHTYPE_PROPERTY_DEF, MATCHTYPE_PROPERTY_DESC)
				Endif

			Case m.lcMatchType = 'RESERVED7'
				lcMatchType = MATCHTYPE_CLASS_DESC

			Case m.lcMatchType = 'RESERVED8'
				lcMatchType = MATCHTYPE_INCLUDE_FILE

			Otherwise
				lcMatchType = m.toObject.MatchType && Restore it back

		Endcase

		toObject.MatchType = m.lcMatchType

	Endproc


*----------------------------------------------------------------------------------
	Procedure BackupFile(tcFilePath, tnReplaceHistoryId)

*SF 20221018 -> local storage
*#Define ccBACKUPFOLDER Addbs(Home(7) + 'GoFishBackups')
*SF 20230131 -> issue #41
*Thisform was not a good idea here
		Local;
			lcBackupPRG     As String,;
			lcDestFile      As String,;
			lcExt           As String,;
			lcExtensions    As String,;
			lcSourceFile    As String,;
			lcThisBackupFolder As String,;
			llCopyError     As Boolean,;
			lnI             As Number,;
			lcBACKUPFOLDER

		Local Array;
			laExtensions(1)
			


*		ccBACKUPFOLDER = Addbs(This.cCR_StoreLocal + 'GoFishBackups')
		lcBACKUPFOLDER = Addbs(This.cCR_StoreLocal + 'GF_ReplaceBackups')
*/SF 20230131 -> issue #41
*/SF 20221018 -> local storage


		If This.oSearchOptions.lPreviewReplace
			Return
		Endif

		llCopyError = .F.

*-- If the user has created a custom backup PRG, and placed it in their path, then call it instead
		lcBackupPRG = 'GoFish_Backup.prg'

		If File(m.lcBackupPRG)
			Do &lcBackupPRG With m.tcFilePath, m.tnReplaceHistoryId
			Return
		Endif

		If Not Directory (lcBACKUPFOLDER) && Create main folder for backups, if necessary
			Mkdir (lcBACKUPFOLDER)
			GF_Write_Readme_Text(4, Addbs(m.lcBACKUPFOLDER) + 'README.md', .T.)

		Endif

* Create folder for this ReplaceHistorrID, if necessary
		lcThisBackupFolder = Addbs (lcBACKUPFOLDER + Transform (m.tnReplaceHistoryId))

		If Not Directory (m.lcThisBackupFolder)
			Mkdir (m.lcThisBackupFolder)

			GF_Write_Readme_Text(5, Addbs(m.lcThisBackupFolder) + 'README.md', .T.)
			Strtofile(Ttoc(Datetime()), Addbs(m.lcThisBackupFolder) + 'README.md', .T.)
		Endif

* Determine the extensions we need to consider
		lcExt = Upper (Justext (m.tcFilePath))

		Do Case
			Case m.lcExt = 'SCX'
				lcExtensions = 'SCX,SCT'
			Case m.lcExt = 'VCX'
				lcExtensions = 'VCX,VCT'
			Case m.lcExt = 'FRX'
				lcExtensions = 'FRX,FRT'
			Case m.lcExt = 'MNX'
				lcExtensions = 'MNX,MNT,MPR,MPX'
			Case m.lcExt = 'DBC'
				lcExtensions = 'DBC,DCT,DCX'
			Case m.lcExt = 'LBX'
				lcExtensions = 'LBX,LBT'
			Otherwise
				lcExtensions = m.lcExt
		Endcase

*-- Copy each file into the destination folder, if its not already there
		Alines (laExtensions, m.lcExtensions, 0, ',')

		For lnI = 1 To Alen (m.laExtensions)
			lcSourceFile = Forceext (m.tcFilePath, laExtensions (m.lnI))
			lcDestFile   = m.lcThisBackupFolder + Justfname (m.lcSourceFile)
			If Not File (m.lcDestFile)
				Try
						Copy File (m.lcSourceFile) To (m.lcDestFile)
					Catch
						If !m.llCopyError
							This.SetReplaceError('Error creating backup of file.', m.tcFilePath, m.tnReplaceHistoryId)
						Endif
						llCopyError = .T.
				Endtry
			Endif
		Endfor

		Return !m.llCopyError

	Endproc


*----------------------------------------------------------------------------------
	Procedure BuildDirectoriesCollection(tcDir, tlWithRepo, tcRepo)

*-- Note: This method is called recursively on itself if subfolders are found. See the For loop at the bottom...
*-- For more good info on recursive processing of directories, see this page: http://fox.wikis.com/wc.dll?Wiki~RecursiveDirectoryProcessing

		Local;
			lcCurrentDirectory As String,;
			lcDriveAndDirectory As String,;
			llChanged        As Boolean,;
			lnDirCount       As Number,;
			lnFileCount      As Number,;
			lnPtr            As Number

		Local Array;
			laDirList(1),;
			laFileList(1)

*!* ** { JRN -- 07/11/2016 08:11 AM - Begin
*!* If Lastkey() = 27 or Inkey() = 27
		If Inkey() = 27
*!* ** } JRN -- 07/11/2016 08:11 AM - End
			This.lEscPress = .T.
			Clear Typeahead
			Return 0
		Endif

*get the toplevel folder, if we are in a repo
		If !m.tlWithRepo Then
			lcCommand = 'git rev-parse --git-dir>git_x.tmp'
			Run &lcCommand

			If File('git_x.tmp') Then
*the result is either the git base folder or empty for no git repo
				tcRepo = Upper(Fullpath(Chrtran(Filetostr('git_x.tmp'), '/' + Chr(13) + Chr(10), '\')))
				Delete File git_x.tmp
				tlWithRepo = .T.
			Else &&file('git_x.tmp')
* no file, no git
				tcRepo = ''
			Endif &&file('git_x.tmp')

		Endif &&!m.tlWithRepo

*SF 20221216 special folders to skip
		If Upper(Fullpath(m.tcDir)) == m.tcRepo
*git toplevel folder. We do not look this up
			Return 0
		Endif
		If Directory(Fullpath(Addbs(m.tcDir) + "GF_Saved_Search_Results" ))
*			;
Or "GF_SAVED_SEARCH_RESULTS" $ Upper(m.tcDir) THEN
*GoFish storage folder. Do not touch. Might create havoc
*One still might search the GF_Saved_Search_Results folder itself
			Return 0
		Endif &&Directory(Fullpath(Addbs(m.tcDir) + "GF_Saved_Search_Results" ))
*/SF 20221216 special folders to skip

		Try
				Chdir (m.tcDir)
				llChanged = .T.
			Catch
				llChanged = .F.
		Endtry

		If !m.llChanged
			Return .F.
		Endif

		This.ShowWaitMessage('Scanning directory ' + m.tcDir)

		lcCurrentDirectory  = Curdir()
		lcDriveAndDirectory = Addbs(Sys(5) + Sys(2003))

		This.oDirectories.Add(m.lcDriveAndDirectory)

		lnDirCount = Adir(laDirList, '*.*', 'D')

		If Vartype(This.oProgressBar) = 'O'
			lnFileCount                 = Adir(laFileList, m.lcDriveAndDirectory + '*.*')
			This.oProgressBar.nMaxValue = This.oProgressBar.nMaxValue + m.lnFileCount
		Endif

		For lnPtr = 1 To m.lnDirCount
			If 'D' $ laDirList(m.lnPtr, 5) && If we have found another dir, then we need to work through it also
				If Vartype(This.oProgressBar) = 'O'
					This.oProgressBar.nMaxValue = This.oProgressBar.nMaxValue - 0 && Subtract off directories from file count
				Endif
				lcCurrentDirectory = laDirList(m.lnPtr, 1)
				If m.lcCurrentDirectory <> '.' And m.lcCurrentDirectory <> '..'
					This.BuildDirectoriesCollection(m.lcCurrentDirectory, m.tlWithRepo, m.tcRepo)
				Endif
			Endif
		Endfor

		Cd ..

	Endproc


*----------------------------------------------------------------------------------
	Procedure BuildProjectsCollection

		Local;
			lcCurrentDir  As String,;
			lcProject     As String,;
			lnX           As Number,;
			loMRU_Project As Object,;
			loMRU_Projects As Object,;
			loPEME_BaseTools As 'GF_PEME_BaseTools' Of 'Lib\GF_PEME_BaseTools.prg',;
			loProject     As Object

		Local Array;
			laProjects(1)

		lcCurrentDir = Addbs(Sys(5) + Sys(2003)) && Current Default Drive and path

*-- Blank out current Projects collecitons. Will rebuild below...
		This.oProjects = Createobject('Collection')
		This.cProjects = ''

		If Version(2) = 0 && If we are running from an .EXE file then exit (No projects will be open)
			Return
		Endif

*-- Add all open Projects in _VFP to the Collection
		For Each m.loProject In _vfp.Projects
			lcProject = Lower(m.loProject.Name)
			This.AddProject(m.lcProject)
			This.cProjects = This.cProjects + m.lcProject + Chr(13)
		Endfor

*-- Add any Projects in the current folder
		Adir(laProjects, m.lcCurrentDir + '*.pjx')

		For lnX = 1 To Alen(m.laProjects) / 5
			lcProject = Lower(Fullpath(laProjects(m.lnX, 1)))
			This.AddProject(m.lcProject)
			This.cProjects = This.cProjects + m.lcProject + Chr(13)
		Endfor

*-- Add MRU Projects to the Collection...
		loPEME_BaseTools = Createobject('GF_PEME_BaseTools')

		loMRU_Projects = loPEME_BaseTools.GetMRUList('PJX')

		For Each m.loMRU_Project In m.loMRU_Projects
			lcProject = Lower(m.loMRU_Project)
			This.AddProject(m.lcProject)
			This.cProjects = This.cProjects + m.lcProject + Chr(13)
		Endfor

	Endproc


*----------------------------------------------------------------------------------
	Procedure ChangeCurrentDir(tcDir)

		Local;
			lcCurrentDirectory As String,;
			lcDefaultDrive  As String,;
			lcPath          As String,;
			llReturn        As Boolean

*-- Attempt to change current dir to passed in location -------
		If !Empty(m.tcDir)
			Try
					Cd (m.tcDir)
					llReturn = .T.
				Catch
					This.SetSearchError('Invalid path [' + m.tcDir + '] passed to ChangeCurrentDir() method.')
					llReturn = .F.
			Endtry
		Else
			llReturn = .T.
		Endif

		This.BuildProjectsCollection()

		Return m.llReturn

	Endproc


*----------------------------------------------------------------------------------
*SF 20230620 unused
*!*		Procedure CheckFileExtTemplate(tcFile)

*!*			Local;
*!*				lcFileExtTemplate As String,;
*!*				lcFileName     As String,;
*!*				lcFilenameMask As String,;
*!*				llFilenameMatch As Boolean,;
*!*				llReturn       As Boolean

*!*			lcFileExtTemplate = Justext(This.oSearchOptions.cFileTemplate)

*!*			llReturn = This.MatchTemplate(m.tcFile, m.lcFileExtTemplate)

*!*		Endproc


*----------------------------------------------------------------------------------
*SF 20230620 unused
*!*		Procedure CheckFilenameTemplate(tcFile)

*!*			Local;
*!*				lcFileName  As String,;
*!*				lcFilenameMask As String,;
*!*				llMatch     As Boolean,;
*!*				lnLength    As Number

*!*			If Empty(Juststem(This.oSearchOptions.cFileTemplate))
*!*				Return .T.
*!*			Endif

*!*			lcFilenameMask = Upper(Juststem(This.oSearchOptions.cFileTemplate))
*!*			lcFileName     = Upper(Juststem(m.tcFile))

*!*			Do Case
*!*				Case m.lcFilenameMask = '*'
*!*					llMatch = .T.
*!*				Case (Left(m.lcFilenameMask, 1) = '*' And Right(m.lcFilenameMask, 1) = '*') Or Atc('*', m.lcFilenameMask) = 0
*!*					llMatch = m.lcFilenameMask $ m.lcFileName
*!*				Case Right(m.lcFilenameMask, 1) = '*'
*!*					lnLength       = Len(cFilenameMask) - 1
*!*					lcFilenameMask = Left(m.lcFilenameMask, m.lnLength)
*!*					llMatch        = Left(m.lcFileName, m.lnLength) = m.lcFilenameMask
*!*				Case Left(m.lcFilenameMask, 1) = '*'
*!*					lnLength       = Len(cFilenameMask) - 1
*!*					lcFilenameMask = Right(m.lcFilenameMask, m.lnLength)
*!*					llMatch        = Right(m.lcFileName, m.lnLength) = m.lcFilenameMask
*!*			Endcase

*!*			Return m.llMatch

*!*		Endproc


*----------------------------------------------------------------------------------
	Procedure CleanUpBinaryString(tcString, llClipAtChr8)

		Local;
			lnStart As Number

*:Global;
x

		If m.llClipAtChr8 && The Select statement from a DBC View needs to be clipped at the Chr(8) near the end of the statement
			lnStart  = Atc(Chr(8), m.tcString)
			tcString = Left(m.tcString, m.lnStart)
		Endif

*-- Replace junk characters with a space
		For x = 0 To 31
			tcString = Strtran(m.tcString, Chr(x), ' ')
		Endfor

		Return m.tcString

	Endproc


*----------------------------------------------------------------------------------
	Procedure ClearReplaceErrorMessage

		This.oSearchOptions.cReplaceErrorMessage = ''

	Endproc


*----------------------------------------------------------------------------------
	Procedure ClearReplaceSettings

		This.oSearchOptions.lAllowBlankReplace = .F.
		This.oSearchOptions.cReplaceExpression = ''

	Endproc


*----------------------------------------------------------------------------------
	Procedure ClearResultsCollection

		This.oResults = Createobject('Collection')

	Endproc


*----------------------------------------------------------------------------------
	Procedure ClearResultsCursor()

		Local;
			lcSearchResultsAlias As String,;
			lnSelect          As Number

		lnSelect = Select()

		lcSearchResultsAlias = This.cSearchResultsAlias

		Create Cursor (m.lcSearchResultsAlias)( ;
			cUni      c(11), ;
			cUni_File c(23), ;
			Datetime c(24), ;
			Scope v(254), ;
			Search v(254), ;
			lMemLoaded L,;
			lMemSaved  L,;
			Process L, ;
			FilePath c(254), ;
			FileName c(100), ;
			TrimmedMatchLine c(254), ;
			BaseClass c(254), ;
			ParentClass c(254), ;
			Class c(254), ;
			Name c(254), ;
			MethodName c(80), ;
			ContainingClass c(254), ;
			ClassLoc c(254), ;
			MatchType c(25), ;
			Timestamp T, ;
			FileType c(4), ;
			Type c(12), ;
			Recno N(6, 0), ;
			ProcStart I, ;
			procend I, ;
			proccode M, ;
			statement M, ;
			statementstart I, ;
			firstmatchinstatement L, ;
			firstmatchinprocedure L, ;
			MatchStart I, ;
			MatchLen I, ;
			lIsText L, ;
			Column c(10), ;
			Code M, ;
			Id I, ;
			MatchLine M, ;
			Replaced L, ;
			TrimmedReplaceLine c(254), ;
			ReplaceLine c(254), ;
			ReplaceRisk I, ;
			Replace_DT T, ;
			iReplaceFolder I, ;
			lJustReplace L, ;
			lSaved L ;
			)

		Select (m.lnSelect)

	Endproc


*----------------------------------------------------------------------------------
	Procedure Compile(tcFile)

		Local;
			lcExt As String

		If This.oSearchOptions.lPreviewReplace
			Return
		Endif

		lcExt = Alltrim(Upper(Justext(m.tcFile)))

		Do Case
			Case m.lcExt = 'VCX'
				Compile Classlib (m.tcFile)

			Case m.lcExt = 'SCX'
				Compile Form (m.tcFile)

			Case m.lcExt = 'LBX'
				Compile Label (m.tcFile)

			Case m.lcExt = 'FRX'
				Compile Report (m.tcFile)
		Endcase

	Endproc


*----------------------------------------------------------------------------------
	Procedure CreateMenuDisplay(tcMenu)

		*** JRN 2024-02-05 : Change the layout, slightly, of the menu display
		
		#Define SPACING 3
		#Define PREFIX ''

		Local laLevels[1], lcIndent, lcPrompt, lcResult, lnLevel, lnSelect

		lnSelect = Select()
		Select (tcMenu)

		lcResult = ''
		lnLevel = 1
		Dimension This.aMenuStartPositions[Reccount(tcmenu)]

		Scan
			This.aMenuStartPositions[Recno(tcmenu)] = Len(lcResult)
			lcIndent = Replicate(Tab, m.lnLevel - 1)
			Do Case
				Case objCode = 22

				Case objCode = 1
					laLevels[1]	= Name

				Case objCode = 77
					lcPrompt = Prompt
					lnLevel	 = Ascan(m.laLevels, Trim(LevelName))

				Case objCode = 0
					lcResult = m.lcResult + m.lcIndent + Strtran(m.lcPrompt, '\-', Replicate('-', 8)) + CR
					lnLevel	 = m.lnLevel + 1
					Dimension m.laLevels[m.lnLevel]
					laLevels[m.lnLevel]	= Name

				Otherwise
					lnLevel	 = Ascan(m.laLevels, Trim(LevelName))
					lcResult = m.lcResult + m.lcIndent + Strtran(Prompt, '\-', Replicate('-', 8)) + CR 
			Endcase
		Endscan

		Select(m.lnSelect)

		Return m.lcResult
		
	EndProc


*----------------------------------------------------------------------------------
	Procedure CreateResult(toObject)
		Local;
			llReturn As Boolean

		llReturn         = .T.
		This.nMatchLines = This.nMatchLines + 1

		If This.oSearchOptions.lCreateResultsCursor
			llReturn = This.CreateResultsRow(m.toObject)
		Endif

		If This.oSearchOptions.lCreateResultsCollection
			This.oResults.Add(m.toObject)
		Endif

		Return m.llReturn
	Endproc


*----------------------------------------------------------------------------------
	Procedure CreateResultsRow(toObject)

*-- This set of mem vars is required to insert a new row into the local results cursor.
*-- The passed in toObject must be an object which has the reference properties on it, so
*-- that a complete record can be created.

		Local;
			lIsText               As Integer,;
			lMemLoaded,;
			lcObjectNameFromProperty As String,;
			lcProperty            As String,;
			lcResultsAlias        As String,;
			llReturn              As Boolean,;
			lnWords               As Number,;
			loException           As Exception

*:Global;
BaseClass,;
Class,;
ClassLoc,;
Code,;
Column,;
ContainingClass,;
Datetime,;
FileName,;
FilePath,;
FileType,;
Id,;
MatchLen,;
MatchLine,;
MatchStart,;
MatchType,;
MethodName,;
Name,;
ParentClass,;
ProcStart,;
Process,;
Recno,;
ReplaceRisk,;
Scope,;
Search,;
Timestamp,;
TrimmedMatchLine,;
cUni,;
cUni_File,;
proccode,;
procend,;
statement,;
statementstart

		llReturn       = .T.
		lcResultsAlias = This.cSearchResultsAlias

		Set Hours To 24
		Datetime = Ttoc(This.tRunTime)
		Set Hours To &lnHours

*		cUni     = "_" + Sys(2007, m.Datetime, 0, 1)
		cUni       = This.cUni
		Scope      = This.oSearchOptions.cRecentScope
		Search     = This.oSearchOptions.cSearchExpression
		lMemLoaded = .T.

		With m.toObject
			MethodName       = This.FixPropertyName(.MethodName)
			MatchLine        = .MatchLine
			TrimmedMatchLine = .TrimmedMatchLine
			ProcStart        = .ProcStart
			procend          = .procend
			proccode         = Evl(.proccode, .Code)
			statement        = Evl(.statement, .MatchLine)
			statementstart   = .statementstart
			MatchStart       = .MatchStart
			MatchLen         = .MatchLen
			MatchType        = .MatchType
			Code             = .Code
		Endwith

		With m.toObject.UserField
			Process         = .F.
			FilePath        = Lower(.FilePath)
			FileName        = Lower(.FileName)
			FileType        = .FileType
			lIsText         = .IsText
			BaseClass       = ._BaseClass
			ParentClass     = ._ParentClass
			ContainingClass = .ContainingClass
			Name            = ._Name
			Class           = ._Class
			ClassLoc        = .ClassLoc
			Recno           = .Recno && from the VCX, SCX, VCX, etc.
			Timestamp       = .Timestamp
			Column          = .Column
		Endwith

* *-- Removed 07/07/2012
* *--- Clean up / doctor up the Object Name
* If 'scx' $ Lower(m.filetype)  && trim off the form name from front of object name
* 	m.name = Substr(m.name, Atc('.', m.name) + 1)
* EndIf

*--- Sometimes, part of the object name may live on the match line
*--- So, we need to append it to the end of the object name
		If m.MatchType $ (MATCHTYPE_PROPERTY_NAME + MATCHTYPE_PROPERTY_VALUE)
			lcObjectNameFromProperty = ''
			lcProperty               = Getwordnum(m.TrimmedMatchLine, 1)
			lnWords                  = Getwordcount(m.lcProperty, '.')

			If m.lnWords > 1
				lcObjectNameFromProperty = Left(m.lcProperty, Atc('.', m.lcProperty, m.lnWords - 1) - 1)
			Endif

			Name = Alltrim(m.name + '.' + m.lcObjectNameFromProperty, '.')
		Endif

*--------------------------------------------------------------------------------

		Id = Reccount(m.lcResultsAlias) + 1 && A unique key for each record

		cUni_File = m.cUni + "_" + Sys(2007,Trim(Padr(m.Id, 11)), 0 ,1) + "_"

		ReplaceRisk = This.GetReplaceRiskLevel(m.toObject)

		Try
*Was zu testen ist
				Insert Into &lcResultsAlias From Memvar

			Catch To m.loException When m.loException.ErrorNo=1190
* Abzufangender Fehler
				If Messagebox('File too large "' + m.FilePath + '"',1) = 2
					llReturn = .F.
					Assert .F.
				Endif &&Messagebox('File too large "' + m.FilePath + '"',1) = 2

			Catch To m.loException
* andere Fehler, Standardhandler rufen
				Throw
			Finally
*
		Endtry

		Return m.llReturn
	Endproc


*----------------------------------------------------------------------------------
	Procedure Destroy

		This.oRegExForProcedureStartPositions = .Null.
		This.oRegExForSearch                  = .Null.
		This.oRegExForSearchInCode 			  = .Null.
		This.oResults                         = .Null.
		This.oSearchOptions                   = .Null.
		This.oFrxCursor                       = .Null.
		This.oProjects                        = .Null.
		This.oSearchErrors                    = .Null.
		This.oReplaceErrors                   = .Null.
		This.oDirectories                     = .Null.
		This.oProgressBar                     = .Null.
		This.oFSO                             = .Null.

	Endproc


*----------------------------------------------------------------------------------
	Procedure DropColumn(tcTable, tcColumnName)

		Local;
			lcAlias As String

		lcAlias = Juststem(m.tcTable)

		Try
				Alter Table (m.tcTable)  Drop Column &tcColumnName
			Catch
		Endtry

	Endproc


*----------------------------------------------------------------------------------
	Procedure EditFromCurrentRow(tcCursor, tlSelectObjectOnly, tlMoveToTopleft)

		Local;
			lcClass     As String,;
			lcCodeBLock As String,;
			lcExt       As String,;
			lcFileToEdit As String,;
			lcMatchType As String,;
			lcMethod    As String,;
			lcMethodString As String,;
			lcName      As String,;
			lcProperty  As String,;
			lnMatchStart As Number,;
			lnProcStart As Number,;
			lnRecNo     As Number,;
			lnStart     As Number,;
			lnWords     As Number,;
			loPBT       As 'GF_PEME_BaseTools',;
			loTools     As Object

		lcExt        = Alltrim(Upper(&tcCursor..FileType))
		lcFileToEdit = Upper(Alltrim(&tcCursor..FilePath))
		lcClass      = Alltrim(&tcCursor..Class)
		lcName       = Alltrim(&tcCursor..Name)
		lcMethod     = Alltrim(&tcCursor..MethodName)
		lcMatchType  = Alltrim(&tcCursor..MatchType)
		lnRecNo      = &tcCursor..Recno
		lnProcStart  = &tcCursor..ProcStart
		lnMatchStart = &tcCursor..MatchStart

*!*	Changed by: nmpetkov 27.3.2023
*!*	<pdm>
*!*	<change date="{^2023-03-27,15:45:00}">Changed by: nmpetkov<br />
*!*	Changes to  Highlight searched text in opened window #75
*!*	</change>
*!*	</pdm>
*If lcExt # 'PRG' And (Empty(m.lcMethod) Or 0 # Atc('<Property', m.lcMatchType))
* here any file that is a text file should be accepted to position the cursor when it is opened
		If !Inlist(m.lcExt, 'PRG', 'SPR', 'MPR', 'QPR', 'H', 'INI', 'TXT', 'XML', 'HTM');
				And (Empty(m.lcMethod) Or 0 # Atc('<Property', m.lcMatchType))
*!*	/Changed by: nmpetkov 27.3.2023
			lcMethodString = ''
			lnStart        = 1
		Else
			lcMethodString = Alltrim(m.lcName + '.' + m.lcMethod, 1, '.')

			If m.lcExt $ ' SCX VCX '
*-- Calculate Line No from procstart and matchstart postitions...
				lcCodeBLock = Substr(&tcCursor..Code, m.lnProcStart + 1, m.lnMatchStart - m.lnProcStart)
				lnStart     = Getwordcount(m.lcCodeBLock, Chr(13)) - 1 && The LINE NUMBER that match in on within the method
				lnStart     = Iif(m.lnStart > 0, m.lnStart, 1)
				Do Case
					Case m.lcExt = 'SCX'
						lcClass = ''

					Case m.lcExt = 'VCX'
						If m.lcName = m.lcClass
							lcMethodString 	= m.lcMethod
						Endif
				Endcase
			Else
				lnStart = (&tcCursor..MatchStart) + 1 && The CHARACTER position of the line where the match is on
			Endif

		Endif

		loPBT = Createobject('GF_PEME_BaseTools')

*** JRN 2021-03-21 : If match is to a name of a file in a Project, open that file
		If &tcCursor..FileType = 'PJX' And &tcCursor..MatchType = MatchType_Name
			lcFileToEdit = Fullpath(Upper(Addbs(Justpath(Trim(&tcCursor..FilePath))) + Trim(&tcCursor..TrimmedMatchLine)))
			loPBT.EditSourceX(m.lcFileToEdit)
			Return
		Else
			lcFileToEdit = Upper(Alltrim(&tcCursor..FilePath))
		Endif
* --------------------------------------------------------------------------------

*-- 2011-12-28 (As requested by JRN) -------------
*-- The following code will automatically select the actual Object on the form or class, or select the Property name.
*-- This will also select it in the PEM Editor main form.
		If Type('_Screen.cThorDispatcher') = 'C'

			loTools = Execscript(_Screen.cThorDispatcher, 'Class= tools from pemeditor')

			If Vartype(m.loTools) = 'O'

				If m.lcExt = 'SCX' And &tcCursor..BaseClass = 'form' && Must trim off form name from front of object name
					lcName = ''
				Endif

				Do Case
					Case m.lcMatchType = MatchType_Name
						loPBT.EditSourceX(m.lcFileToEdit, m.lcClass)
						loTools.SelectObject(m.lcName)
						Return

					Case m.lcMatchType $ (MATCHTYPE_PROPERTY_NAME + MATCHTYPE_PROPERTY_VALUE + MATCHTYPE_PROPERTY_DEF )
*-- Pull out the Property name from the MatchLine (it can be preceded by an object name)
						lcProperty = Getwordnum(&tcCursor..TrimmedMatchLine, 1)
						lnWords    = Getwordcount(m.lcProperty, '.')
						lcProperty = Getwordnum(m.lcProperty, m.lnWords, '. ')
						lcProperty = This.FixPropertyName(m.lcProperty)

						loPBT.EditSourceX(m.lcFileToEdit, m.lcClass)
						loTools.SelectObject(m.lcName, m.lcProperty)
						Return
				Endcase
			Endif

		Endif

		If m.lcExt = 'MNX'
			*** JRN 2024-02-05 : special handling for MNXs - if possible, using the keyboard buffer
			* to navigate to the actual record and procedure
			lcKeyStrokes = This.GetMenuKeystrokes(m.lcFileToEdit, m.lnRecNo, m.lcMatchType)
			m.loPBT.EditSourceX(m.lcFileToEdit, m.lcClass, m.lnStart, m.lnStart, m.lcMethodString, m.lnRecNo)
			If not Empty(m.lcKeyStrokes)
				Keyboard(m.lcKeyStrokes)
			Endif
		Else
			m.loPBT.EditSourceX(m.lcFileToEdit, m.lcClass, m.lnStart, m.lnStart, m.lcMethodString, m.lnRecNo)
		EndIf


*!*	Changed by: nmpetkov 27.3.2023
*!*	<pdm>
*!*	<change date="{^2023-03-27,15:45:00}">Changed by: nmpetkov<br />
*!*	Changes to  Highlight searched text in opened window #75
*!*	</change>
*!*	</pdm>
		lcMatchType = Alltrim(&tcCursor..MatchType)
*	Try to select searched text if found in normal windows only - exclude internal for VFP places
		If !Inlist(m.lcMatchType, MATCHTYPE_FILENAME, MATCHTYPE_CLASS_DEF, MATCHTYPE_CLASS_DESC, MATCHTYPE_METHOD_DEF, MATCHTYPE_PROPERTY_DEF, ;
				MATCHTYPE_CONTAINING_CLASS, MATCHTYPE_PARENTCLASS, MATCHTYPE_BASECLASS, MATCHTYPE_METHOD_DESC, MATCHTYPE_PROPERTY, ;
				MATCHTYPE_PROPERTY_DESC, MATCHTYPE_PROPERTY_NAME, MATCHTYPE_PROPERTY_VALUE)
			This.SelectSearchedText(&tcCursor..MatchStart,&tcCursor..MatchLen, Trim(&tcCursor..Search))
		Endif
*!*	/Changed by: nmpetkov 27.3.2023

		If m.tlMoveToTopleft And (m.lcExt = 'PRG' Or Not Empty(m.lcMethodString))
			This.ThorMoveWindow()
		Endif

	Endproc

*!*	Changed by: nmpetkov 27.3.2023
*!*	<pdm>
*!*	<change date="{^2023-03-27,15:45:00}">Changed by: nmpetkov<br />
*!*	Changes to  Highlight searched text in opened window #75
*!*	</change>
*!*	</pdm>
*----------------------------------------------------------------------------------
*	Highlight searched text in opened window
*		tnRangeStart - start of the line where the search is found
*		tnRangelen - length of the line where the search is found - optional, can reduce the length of the text to be searched
*		tcSearch - searched text
*
*nmpetkov 27.3.2023
*----------------------------------------------------------------------------------
	Procedure SelectSearchedText(tnRangeStart, tnRangelen, tcSearch)
		Local;
			lLibrRelease As Boolean,;
			lcFoxtoolsFll As String,;
			lcLine     As String,;
			llMatchCase As Boolean,;
			lnPos      As Number,;
			lnRangeEnd As Number,;
			lnRangeStart As Number,;
			lnRetCode  As Number,;
			lnSelEnd   As Number,;
			lnSelStart As Number,;
			lnWhandle  As Number,;
			loMatch    As Object,;
			loMatches  As Object

		Local Array;
			aEdEnv(25)

		If Atc("foxtools.fll", Set("LIBRARY")) = 0
			lcFoxtoolsFll = Sys(2004) + "foxtools.fll"
			If File(m.lcFoxtoolsFll)
				lLibrRelease = .T.
				Set Library To (m.lcFoxtoolsFll) Additive
			Endif
		Endif

		If Atc("foxtools.fll", Set("LIBRARY")) > 0
			lnWhandle = _WOnTop()
			lnRetCode = _EdGetEnv(m.lnWhandle, @aEdEnv) && aEdEnv: 1 - filename, 2 - size, 12 - readonly?, 17 - selected start, 18 selected  end
			If m.lnRetCode = 1 And aEdEnv[2] > 0 && content size is > 0
* determine the range in which to be searched
				If aEdEnv[17] > 0 Or Empty(m.tnRangeStart) Or m.tnRangeStart >= aEdEnv[2]
* defaults to current cursor position, if is set, otherwise the given as parameter
*tnRangeStart = _EdGetPos(m.lnWhandle) && this value is allready available in aEdEnv
					tnRangeStart = aEdEnv[17]
				Endif
				lnRangeStart = m.tnRangeStart
				If Empty(m.tnRangelen)
					tnRangelen = aEdEnv[2] - m.lnRangeStart + 1
				Endif
				lnRangeEnd = m.lnRangeStart + m.tnRangelen && determine where the search to be searched :-)
				If m.lnRangeEnd > aEdEnv[2] && check we are not beyond the end, will throw error
					lnRangeEnd = aEdEnv[2]
				Endif
				lcLine = _EdGetStr(m.lnWhandle, m.lnRangeStart, m.lnRangeEnd)
* determine real string to search in case pattern or Regex
				llMatchCase = This.oSearchOptions.lMatchCase
				If This.oSearchOptions.nSearchMode > 1
					This.PrepareRegExForSearch()
					This.PrepareRegExForReplace()
					loMatches = This.oRegExForSearch.Execute(m.lcLine)
					If m.loMatches.Count > 0
						loMatch     = loMatches.Item(0)
						tcSearch    = m.loMatch.Value
						llMatchCase = .T.
					Endif
				Endif
* search what to be selected in the range
				If m.llMatchCase
					lnPos = At(m.tcSearch, m.lcLine)
				Else
					lnPos = Atc(m.tcSearch, m.lcLine)
				Endif
				If m.lnPos > 0
					lnSelStart = m.lnRangeStart + m.lnPos - 1
					lnSelEnd   = m.lnSelStart + Len(m.tcSearch)
				Else
* In case the search fails (match word or regular expressions), select whole the line
					lnSelStart = m.tnRangeStart
					lnSelEnd   = m.tnRangeStart + m.tnRangelen
				Endif
* select at the end
				If m.lnSelEnd > m.lnSelStart
					_EdSelect(m.lnWhandle, m.lnSelStart, m.lnSelEnd)
				Endif
			Endif
		Endif

		If m.lLibrRelease And Atc(m.lcFoxtoolsFll, Set("LIBRARY")) > 0
			Release Library (m.lcFoxtoolsFll)
		Endif
	Endproc
*!*	/Changed by: nmpetkov 27.3.2023

*----------------------------------------------------------------------------------
	Procedure EditMenuFromCurrentRow(tcCursor)

		Local;
			lcFileToEdit As String,;
			lcMenuAlias As String,;
			lcMenuDisplay As String,;
			lcTempFile As String,;
			llSuccess  As Boolean,;
			lnEndPos   As Number,;
			lnRecNo    As Number,;
			lnStartPos As Number,;
			loEditorWin As Editorwin Of 'c:\visual foxpro\programs\mythor\thor\tools\apps\pem editor\source\peme_editorwin.vcx'

		lcMenuAlias  = 'Menu' + Sys(2015)
		lcFileToEdit = Upper(Alltrim(&tcCursor..FilePath))

		Try
				Use (m.lcFileToEdit) Shared Again In 0 Alias &lcMenuAlias
				llSuccess = .T.
			Catch
				llSuccess = .F.
		Endtry

		If m.llSuccess = .F.
			Return m.llSuccess
		Endif

		lcMenuDisplay = This.CreateMenuDisplay(m.lcMenuAlias)
		lcTempFile    = Addbs(Sys(2023)) + Chrtran(Justfname(m.lcFileToEdit), '.', '_')  + Sys(2015) + '.txt'
		Strtofile(m.lcMenuDisplay, m.lcTempFile)
		Modify File (m.lcTempFile) Nowait

		loEditorWin = Execscript(_Screen.cThorDispatcher, 'Class= editorwin from pemeditor')
		loEditorWin.ResizeWindow(600, 800)
		loEditorWin.SetTitle(m.lcTempFile)

		lnRecNo = &tcCursor..Recno

		If Between(m.lnRecNo, 1, Reccount(m.lcMenuAlias))
			lnStartPos = This.aMenuStartPositions[m.lnRecNo]
			If m.lnRecNo < Reccount(m.lcMenuAlias)
				lnEndPos = This.aMenuStartPositions[m.lnRecNo + 1]
			Else
				lnEndPos = 1000000
			Endif

			loEditorWin.EnsureVisible(0)
			loEditorWin.Select(m.lnStartPos, m.lnEndPos)
			loEditorWin.EnsureVisible(m.lnStartPos)
		Endif

		Use In (m.lcMenuAlias)

	Endproc


*----------------------------------------------------------------------------------
	Procedure EditObjectFromCurrentRow(tcCursor)

		Local;
			lcClass   As String,;
			lcExt     As String,;
			lcFileToEdit As String,;
			lcMatchType As String,;
			lcName    As String,;
			lcProperty As String,;
			lnWords   As Number,;
			loPBT     As 'GF_PEME_BaseTools',;
			loTools   As Object

		lcExt        = Alltrim(Upper(&tcCursor..FileType))
		lcMatchType  = Alltrim(&tcCursor..MatchType)
		lcFileToEdit = Upper(Alltrim(&tcCursor..FilePath))
		lcClass      = Alltrim(&tcCursor..Class)
		lcName       = Alltrim(&tcCursor..Name)

		loPBT = Createobject('GF_PEME_BaseTools')
		loPBT.EditSourceX(m.lcFileToEdit, m.lcClass)

		If Type('_Screen.cThorDispatcher') = 'C'

			loTools = Execscript(_Screen.cThorDispatcher, 'Class= tools from pemeditor')

			If Vartype(m.loTools) = 'O'

				If m.lcExt = 'SCX' And &tcCursor..BaseClass = 'form' && Must trim off form name from front of object name
					lcName = ''
				Endif

				If m.lcMatchType $ (MATCHTYPE_PROPERTY_NAME + MATCHTYPE_PROPERTY_VALUE + MATCHTYPE_PROPERTY_DEF )
*-- Pull out the Property name from the MatchLine (it can be preceded by an object name)
					lcProperty = Getwordnum(&tcCursor..TrimmedMatchLine, 1)
					lnWords    = Getwordcount(m.lcProperty, '.')
					lcProperty = Getwordnum(m.lcProperty, m.lnWords, '. ')
					lcProperty = This.FixPropertyName(m.lcProperty)

					loTools.SelectObject(m.lcName, m.lcProperty)
				Else
					loTools.SelectObject(m.lcName)
				Endif
			Endif

		Endif

	Endproc


*----------------------------------------------------------------------------------
	Procedure EndTimer()

		This.nSearchTime = Seconds() - This.nSearchTime

	Endproc


*----------------------------------------------------------------------------------
	Procedure EscapeSearchExpression(tcString,tnMode)

		Local;
			lcString As String

		If Empty(m.tnMode)
			tnMode = 0
		Endif &&Empty(m.tnMode)

		lcString = m.tcString

		lcString = Strtran(m.tcString, '\', '\\')
		lcString = Strtran(m.lcString, '+', '\+')
		lcString = Strtran(m.lcString, '.', '\.')
		lcString = Strtran(m.lcString, '|', '\|')
		lcString = Strtran(m.lcString, '{', '\{')
		lcString = Strtran(m.lcString, '}', '\}')
		lcString = Strtran(m.lcString, '[', '\[')
		lcString = Strtran(m.lcString, ']', '\]')
		lcString = Strtran(m.lcString, '(', '\(')
		lcString = Strtran(m.lcString, ')', '\)')
		lcString = Strtran(m.lcString, '$', '\$')

		lcString = Strtran(m.lcString, '^', '\^')
		lcString = Strtran(m.lcString, ':', '\:')
		lcString = Strtran(m.lcString, ';', '\;')
		lcString = Strtran(m.lcString, '-', '\-')

		Do Case
			Case (Empty(m.tnMode) And This.oSearchOptions.nSearchMode = GF_SEARCH_MODE_LIKE) Or m.tnMode=1
				lcString = Strtran(m.lcString, '?', '.')
				lcString = Strtran(m.lcString, '*', '.*')
			Case Empty(m.tnMode)  Or m.tnMode=2
				lcString = Strtran(m.lcString, '?', '\?')
				lcString = Strtran(m.lcString, '*', '\*')
			Otherwise

		Endcase

		Return m.lcString

* http://stackoverflow.com/questions/280793/case-insensitive-string-replacement-in-javascript

*!*	RegExp.escape = function(str) {
*!*	var specials = new RegExp("[.*+?|()\\[\\]{}\\\\]", "g"); // .*+?|()[]{}\
*!*	return str.replace(specials, "\\$&");
*!*	}

	Endproc


*----------------------------------------------------------------------------------
	Procedure ExtractMethodName(tcReference)

		If !Empty(m.tcReference)
			Return Justext(m.tcReference)
		Else
			Return ''
		Endif

	Endproc


*----------------------------------------------------------------------------------
	Procedure ExtractObjectName(tcReference)

		If !Empty(m.tcReference)
			Return Juststem(m.tcReference)
		Else
			Return ''
		Endif

	Endproc


*----------------------------------------------------------------------------------
	Procedure FilesToSkip(tcFile)

		Local;
			lcFileName As String,;
			lnI     As Number

		lcFileName = Upper(Justfname(m.tcFile))

		If (Chr(13) + m.lcFileName + Chr(13)) $ This.cFilesToSkip
			Return .T.
		Endif

		For lnI = 1 To This.nWildCardFilesToSkip
			If Like(This.aWildcardFiles[m.lni], m.tcFile)
				Return .T.
			Endif
		Endfor

		Return .F.

	Endproc


*----------------------------------------------------------------------------------
	Procedure FindProcedureForMatch(toProcedureStartPositions, tnStartByte)

		Local;
			llClassDef As Boolean,;
			lnX      As Number,;
			loClassDef As 'GF_Procedure',;
			loNextMatch As Object,;
			loReturn As 'GF_Procedure'

*:Global;
Result

		loReturn = Createobject('GF_Procedure')

		If Isnull(m.toProcedureStartPositions)
			Return m.loReturn
		Endif

		lnX = 1

		For Each Result In m.toProcedureStartPositions

			If Result.StartByte > m.tnStartByte
				Exit
			Else
				loReturn = Result
			Endif

			Do Case
				Case 'END CLASS' $ Upper(Result.Type)
					llClassDef = .F.
					loClassDef = Createobject('GF_Procedure') && An empty result
				Case 'CLASS' $ Upper(Result.Type)
					llClassDef = .T.
					loClassDef = Result
			Endcase

			lnX = m.lnX + 1
		Endfor

*-- This code attempted to identify matches that we INSIDE of a CLASS, but not inside of a Proc.
*-- This would catch wildly located code in class that is between Proc definitions.
* Removed 10/04/2012
* If lnX < toProcedureStartPositions.count and llClassDef
* 	loNextMatch = toProcedureStartPositions.Item(lnX + 1))
* 	If tnStartByte < loNextMatch.StartByte
* 		loReturn = loClassDef
* 	Endif
* Endif

		Return m.loReturn

	Endproc


*----------------------------------------------------------------------------------
	Procedure FindStatement(loObject)

		Local;
			lcLastLine As String,;
			lcMatchLine As String,;
			lcPreceding As String,;
			lcProcCode As String,;
			lcResult As String,;
			lnCRPos  As Number,;
			lnLen    As Number,;
			lnLength As Number,;
			lnStart  As Number,;
			lnTextStart As Number

		lnStart	 = m.loObject.MatchStart - m.loObject.ProcStart + 1

*!* ** { JRN -- 08/05/2016 07:16 AM - Begin
*!* lnLength = m.loObject.MatchLen - 1
		lnLength = m.loObject.MatchLen
*!* ** } JRN -- 08/05/2016 07:16 AM - End

		lcProcCode = Evl(m.loObject.proccode, m.loObject.Code)

* previously, assumed trailing CR, but this dropped off last character if not found
*!* ** { JRN -- 08/05/2016 07:16 AM - Begin
*!* lcMatchLine	= Substr(m.lcProcCode, m.lnStart, m.lnLength)
		lcMatchLine	= Trim(Substr(m.lcProcCode, m.lnStart, m.lnLength), 1, Chr[13], Chr[10])
*!* ** } JRN -- 08/05/2016 07:16 AM - End

		lcResult	= m.lcMatchLine
*** JRN 12/02/2015 : Add in leading lines, if any
		Do While .T.
			lcPreceding = Left(m.lcProcCode, m.lnStart - 1)
			lnCRPos     = Rat(CR, m.lcPreceding, 2)
			If m.lnCRPos > 0
				lcPreceding = Substr(m.lcPreceding, m.lnCRPos + 1)
			Endif
			If This.IsContinuation(m.lcPreceding)
				lcResult = m.lcPreceding + m.lcResult
				lnStart  = m.lnStart - Len(m.lcPreceding)
				lnLength = Len(m.lcResult)
			Else
				Exit
			Endif
		Enddo

*** JRN 12/02/2015 : Add in following lines, if any
		lcLastLine = m.lcMatchLine
		Do While This.IsContinuation(m.lcLastLine)
			lcLastLine = Substr(m.lcProcCode, m.lnStart + m.lnLength)
			lnLen      = At(CR, m.lcLastLine, 2)
			If m.lnLen > 0
				lcLastLine = Left(m.lcLastLine, m.lnLen - 1)
			Endif
			lcResult = m.lcResult + m.lcLastLine
			lnLength = Len(m.lcResult)
		Enddo

*** JRN 12/05/2015 : within Text / Endtext
		If Atc('text', Left(m.lcProcCode, m.lnStart)) != 0 And ;
				Atc('endtext', Substr(m.lcProcCode, m.lnStart + m.lnLength)) != 0

			lnTextStart = m.lnStart
			Do While m.lnTextStart > 1
				lcPreceding = Left(m.lcProcCode, m.lnTextStart - 1)
				lnCRPos     = Rat(CR, m.lcPreceding, 2)
				If m.lnCRPos > 0
					lcPreceding = Substr(m.lcPreceding, m.lnCRPos + 1)
				Endif
				Do Case
					Case 'text' = Lower(Getwordnum(m.lcPreceding, 1, ' ' + Tab + CR + lf))
						lnTextStart = m.lnTextStart - Len(m.lcPreceding)
						lnLength    = m.lnLength + m.lnStart - m.lnTextStart
						lnStart     = m.lnTextStart
						lcResult    = Substr(m.lcProcCode, m.lnStart, m.lnLength)
						Do While Len(m.lcProcCode) > m.lnStart + m.lnLength
							lcLastLine = Substr(m.lcProcCode, m.lnStart + m.lnLength)
							lnLen      = At(CR, m.lcLastLine, 2)
							If m.lnLen > 0
								lcLastLine = Left(m.lcLastLine, m.lnLen - 1)
							Endif
							lcResult = m.lcResult + m.lcLastLine
							lnLength = Len(m.lcResult)
							If 'endtext' = Lower(Getwordnum(m.lcLastLine, 1, ' ' + Tab + CR + lf))
								Exit
							Endif
						Enddo
						Exit
					Case 'endtext' = Lower(Getwordnum(m.lcPreceding, 1, ' ' + Tab + CR + lf))
						Exit
					Otherwise
						lnTextStart = m.lnTextStart - Len(m.lcPreceding)
				Endcase
			Enddo
		Endif


		loObject.statement      = m.lcResult
		loObject.statementstart = m.lnStart

	Endproc


*----------------------------------------------------------------------------------
	Procedure FixPropertyName(lcProperty)

* Gets rid of dimensions and leading '*^'

		Return Chrtran(Getwordnum(m.lcProperty, 1, '(['), '*^', '')

	Endproc


*----------------------------------------------------------------------------------
	Procedure GenerateHTMLCode(tcCode, tcMatchLine, tnMatchStart, tcCss, tcJavaScript, tcReplaceLine, tlAlreadyReplaced, tnTabsToSpaces, ;
			tcSearch, tcStatementFilter, tcProcFilter)

		Local;
			lcBr             As String,;
			lcColorizedCode  As String,;
			lcCss            As String,;
			lcHTML           As String,;
			lcHtmlBody       As String,;
			lcInitialBr      As String,;
			lcJavaScript     As String,;
			lcLeft           As String,;
			lcMatchLine      As String,;
			lcMatchLinePrefix As String,;
			lcMatchLineSuffix As String,;
			lcMatchPrefix    As String,;
			lcMatchSuffix    As String,;
			lcMatchWordPrefix As String,;
			lcMatchWordSuffix As String,;
			lcProcFilterPrefix As String,;
			lcProcFilterSuffix As String,;
			lcReplaceExpression As String,;
			lcReplaceLine    As String,;
			lcReplaceLinePrefix As String,;
			lcReplaceLineSuffix As String,;
			lcRight          As String,;
			lcRightCode      As String,;
			lcStateFilterPrefix As String,;
			lcStateFilterSuffix As String,;
			lnEndProc        As Number,;
			lnMatchLineLength As Number,;
			lnReplaceLineLength As Number

		lcCss        = Evl(m.tcCss, '')
		lcJavaScript = Evl(m.tcJavaScript, '')

		lcMatchLinePrefix   = '<div id="matchline" class="matchline">'
		lcMatchLineSuffix   = '</div>'
		lcReplaceLinePrefix = '<div id="repalceline" class="replaceline">'
		lcReplaceLineSuffix = '</div>'

		lcMatchWordPrefix = '<span id="matchword" class="matchword">'
		lcMatchWordSuffix = '</span>'

		lcStateFilterPrefix = '<span id="statefilter" class="statefilter">'
		lcStateFilterSuffix = '</span>'

		lcProcFilterPrefix = '<span id="procfilter" class="procfilter">'
		lcProcFilterSuffix = '</span>'

		If !Empty(m.tcMatchLine)

*-- Dress up the code that comes before the match line...
			lcBr   = '<br />'
			lcLeft = Left(m.tcCode, m.tnMatchStart)
			lcLeft = Evl(This.HtmlEncode(m.lcLeft), m.lcBr)

*-- Dress up the matchline...
			lnMatchLineLength   = Len(m.tcMatchLine)
			lnReplaceLineLength = Len(Rtrim(m.tcReplaceLine))

*===================== Colorize the Replace Preview line, if passed ==============================
			If !Empty(m.tcReplaceLine)
				lcColorizedCode   = This.HtmlEncode(m.tcReplaceLine)
				lcReplaceLine     = m.lcReplaceLinePrefix + m.lcColorizedCode + m.lcReplaceLineSuffix
				lcMatchLinePrefix = '<div id="matchline" class="strikeout">'
*SF 20230507 tags where ordered odd
*				lcMatchLinePrefix = m.lcMatchLinePrefix + '<del>'
				lcMatchLinePrefix = '<del>'+m.lcMatchLinePrefix
				lcMatchLineSuffix = m.lcMatchLineSuffix + '</del>'
			Else
				lcReplaceLine = ''
			Endif

*===================== Colorize the match line ====================================================
*-- Mark the match WORD(s), so I can find them after the VFP code is colorized...
			lcReplaceExpression = '[:GOFISHMATCHWORDSTART:] + lcMatch + [:GOFISHMATCHWORDEND:]'

			*** JRN 2024-02-17 : highlighting search words
			Local lcColorizedCode, lcSearch, lnI
			Do Case
				*** JRN 2024-02-17 : old case, no wildcards
				Case Not This.IsWildCardStatementSearch()
					lcColorizedCode = This.RegExReplace(tcMatchLine, '', lcReplaceExpression, .T.)
				*** JRN 2024-02-17 : whole word match
				Case This.oSearchOptions.lMatchWholeWord
					lcColorizedCode = m.tcMatchLine
					For lnI = 1 To Getwordcount(This.oSearchOptions.cWholeWordSearch, '.*')
						lcSearch = Getwordnum(This.oSearchOptions.cWholeWordSearch, m.lnI, '.*')
						If Not Empty(m.lcSearch)
							lcColorizedCode = This.RegExReplace(m.lcColorizedCode, m.lcSearch, lcReplaceExpression, .T.)
						Endif
					EndFor
				*** JRN 2024-02-17 : wildcards, not whole word
				Otherwise
					lcColorizedCode = m.tcMatchLine
					For lnI = 1 To Getwordcount(This.oSearchOptions.cSearchExpression, '*')
						lcSearch = Getwordnum(This.oSearchOptions.cSearchExpression, m.lnI, '*')
						If Not Empty(m.lcSearch)
							lcColorizedCode = This.RegExReplace(m.lcColorizedCode, m.lcSearch, lcReplaceExpression, .T.)
						Endif
					Endfor
			Endcase

*!*	/Changed by: LScheffler 18.3.2023

			lcColorizedCode = This.HtmlEncode(m.lcColorizedCode)

*-- Next, add <span> tags around previously marked match Word(s)
			lcColorizedCode = Strtran(m.lcColorizedCode, ':GOFISHMATCHWORDSTART:', m.lcMatchWordPrefix)
			lcColorizedCode = Strtran(m.lcColorizedCode, ':GOFISHMATCHWORDEND:', m.lcMatchWordSuffix)

*-- Finally, add <div> tags around the entire Matched Line -------------------
			lcMatchLine = m.lcMatchLinePrefix + m.lcColorizedCode + m.lcMatchLineSuffix
*=================================================================================================

*-- Dress up the code that comes after the match line...
*-- (Look for EndProc to know where to end the code)---
			If m.tlAlreadyReplaced
				lcRightCode = Substr(m.tcCode, m.tnMatchStart + 1 + m.lnReplaceLineLength)
			Else
				lcRightCode = Substr(m.tcCode, m.tnMatchStart + 1 + m.lnMatchLineLength)
			Endif

			lnEndProc = Atc('EndProc', m.lcRightCode)

			If m.lnEndProc > 0
				lcRightCode = Substr(m.lcRightCode, 1, m.lnEndProc + 6) && It ends at "E" of "EndProc", so add 6 to get the rest of the word
			Endif

			lcRight = This.HtmlEncode(m.lcRightCode)

*!* ******************** Removed 12/02/2015 *****************
*!* *** JRN 11/14/2015 : Highlight sub-search (filter within same procedure) if it is supplied
*!* If 'C' = Vartype(tcProcFilter) and not Empty(tcProcFilter)
*!* 	lcLeft = This.HighlightProcFilter(lcLeft, tcProcFilter, lcMatchWordPrefix, lcMatchWordSuffix)
*!* 	lcRight = This.HighlightProcFilter(lcRight, tcProcFilter, lcMatchWordPrefix, lcMatchWordSuffix)
*!* EndIf

			lcHtmlBody = m.lcLeft + m.lcMatchLine + m.lcReplaceLine + m.lcRight &&Build the body

*!* ******************** Removed 08/14/2022 *****************
*!* Highlighting for the search results is already done in "RegExReplace()"
*!* If 'C' = Vartype(tcSearch) And Not Empty(tcSearch)
*!* 	lcHtmlBody = This.HighlightProcFilter(lcHtmlBody, tcSearch, lcMatchWordPrefix, lcMatchWordSuffix)
*!* Endif

			*!* ******** JRN Removed 2024-02-24 ******** Buggy, may be re-investigated at a later time
			*!* If 'C' = Vartype(m.tcStatementFilter) And Not Empty(m.tcStatementFilter)
*!* *for statement
			*!* 	lcHtmlBody = This.HighlightStatementFilter(m.lcHtmlBody, m.tcStatementFilter, m.lcStateFilterPrefix, m.lcStateFilterSuffix,;
			*!* 		M.lcMatchLinePrefix, m.lcMatchLineSuffix)
*!* *for replace (if)
*!* *				If !Empty(m.tcReplaceLine) THEN
*!* *						lcHtmlBody = This.HighlightStatementFilter(lcHtmlBody, tcStatementFilter, lcStateFilterPrefix, lcStateFilterSuffix,;
*!* m.lcReplaceLinePrefix, m.lcReplaceLineSuffix)
*!* *				ENDIF &&!Empty(m.tcReplaceLine)
			*!* Endif

			*!* If 'C' = Vartype(m.tcProcFilter) And Not Empty(m.tcProcFilter)
*!* *for proc
			*!* 	lcHtmlBody = This.HighlightProcFilter(m.lcHtmlBody, m.tcProcFilter, m.lcProcFilterPrefix, m.lcProcFilterSuffix)
			*!* Endif

			If Not Empty(m.tnTabsToSpaces)
				lcHtmlBody = Strtran(m.lcHtmlBody, Chr[9], Space(m.tnTabsToSpaces))
			Endif
			lcHtmlBody = Alltrim(m.lcHtmlBody, 1, Chr[13], Chr[10])
		Else

*-- Just a plain blob of VFP code, with no match lines or match words...
*-- Need an empty MatchLine Divs so the JavaScript on the page will find it to scroll the page
			lcHtmlBody = '<div id="matchline"></div>' + This.HtmlEncode(m.tcCode)

		Endif

*-- Build the whole Html by combining the html parts defined above -------------
		TEXT To m.lcHTML Noshow Textmerge Pretext 3
<html>
 <head>
  <title>GoFish code snippet</title>
  <<m.lcCss>>
 </head>

 <body>
  <<m.lcHtmlBody>>
  <br /><br /><br />
  <<m.lcJavaScript>>
 </body>
</html>
		ENDTEXT


		Return m.lcHTML

	Endproc


*----------------------------------------------------------------------------------
	Procedure GetActiveProject()

		Local;
			lcCurrentProject As String

		If Type('_VFP.ActiveProject.Name') = 'C'
			lcCurrentProject = _vfp.ActiveProject.Name
		Else
			lcCurrentProject = ''
		Endif

		Return m.lcCurrentProject

	Endproc


*----------------------------------------------------------------------------------
	Procedure GetCurrentDirectory

		Return Addbs(Sys(5) + Sys(2003))

	Endproc


*----------------------------------------------------------------------------------
	Procedure GetDirectories(tcPath, tlIncludeSubDirectories)

		Local;
			lnFiles As Number

		Local Array;
			laFiles(1)

		This.oDirectories = Createobject('Collection')

		If m.tlIncludeSubDirectories
			This.BuildDirectoriesCollection(m.tcPath)
		Else
			This.oDirectories.Add(m.tcPath)
			If Vartype(This.oProgressBar) = 'O'
				lnFiles                     = Adir(laFiles, '*.*')
				This.oProgressBar.nMaxValue = m.lnFiles
			Endif
		Endif

		Return This.oDirectories

	Endproc


*----------------------------------------------------------------------------------
	Procedure GetFileDateTime(tcFile)

		Local;
			lcExt   As String,;
			ldFileDate As Date,;
			loFile  As Object

		Local Array;
			laMaxDateTime(1)

		ldFileDate = {// ::}
		lcExt      = Upper(Justext(m.tcFile))

		If Inlist(m.lcExt, 'SCX', 'VCX', 'FRX', 'MNX', 'LBX')
			Try
					Use (m.tcFile) Again In 0 Alias 'GF_GetMaxTimeStamp' Shared
					Select Max(Timestamp);
						From GF_GetMaxTimeStamp;
						Into Array laMaxDateTime
					ldFileDate = Ctot(This.TimeStampToDate(m.laMaxDateTime))
				Catch
				Finally
					If Used('GF_GetMaxTimeStamp')
						Use In ('GF_GetMaxTimeStamp')
					Endif
			Endtry
		Endif

		If Empty(m.ldFileDate)
			Try
					ldFileDate = Fdate(m.tcFile, 1)
				Catch
					loFile     = This.oFSO.Getfile(m.tcFile)
					ldFileDate = m.loFile.DateLastModified
			Endtry
		Endif

		Return m.ldFileDate

	Endproc


*----------------------------------------------------------------------------------
	Procedure GetFrxObjectType(tnObjType, tnObjCode)

		Local;
			lcObjectType As String

*-- Details from: http://www.dbmonster.com/Uwe/Forum.aspx/foxpro/4719/Code-meanings-for-Report-Format-ObjType-field

		lcObjectType = ''

		Do Case
			Case m.tnObjType = 1
				lcObjectType = 'Report'
			Case m.tnObjType = 2
				lcObjectType = 'Workarea'
			Case m.tnObjType = 3
				lcObjectType = 'Index'
			Case m.tnObjType = 4
				lcObjectType = 'Relation'
			Case m.tnObjType = 5
				lcObjectType = 'Text'
			Case m.tnObjType = 6
				lcObjectType = 'Line'
			Case m.tnObjType = 7
				lcObjectType = 'Box'
			Case m.tnObjType = 8
				lcObjectType = 'Field'
			Case m.tnObjType = 9 && Band Info
				Do Case
					Case m.tnObjCode = 0
						lcObjectType = 'Title'
					Case m.tnObjCode = 1
						lcObjectType = 'PageHeader'
					Case m.tnObjCode = 2
						lcObjectType = 'Column Header'
					Case m.tnObjCode = 3
						lcObjectType = 'Group Header'
					Case m.tnObjCode = 4
						lcObjectType = 'Detail Band'
					Case m.tnObjCode = 5
						lcObjectType = 'Group Footer'
					Case m.tnObjCode = 6
						lcObjectType = 'Column Footer'
					Case m.tnObjCode = 7
						lcObjectType = 'Page Footer'
					Case m.tnObjCode = 8
						lcObjectType = 'Summary'
				Endcase
			Case m.tnObjType = 10
				lcObjectType = 'Group'
			Case m.tnObjType = 17
				lcObjectType = 'Picture/OLE'
			Case m.tnObjType = 18
				lcObjectType = 'Variable'
			Case m.tnObjType = 21
				lcObjectType = 'Print Drv Setup'
			Case m.tnObjType = 25
				lcObjectType = 'Data Env'
			Case m.tnObjType = 26
				lcObjectType = 'Cursor Obj'
		Endcase

		Return m.lcObjectType

	Endproc


	*----------------------------------------------------------------------------------
	Procedure GetFullMenuPrompt
		*** JRN 2024-02-05 : Get the Full menu prompt (includes prompts for parent sub-menus)
		* Assumes current record in current table; written this way to avoid modifying record pointer
		Local laField[1], laParent[1], lcDBF, lcPrompt, lnLevelName, lnRecNo
	
		lcPrompt = Alltrim(Prompt)
		lcDBF	 = Dbf()
		lnRecNo	 = Recno()
	
		Select  LevelName,					;
				Prompt						;
			From (m.lcDBF)					;
			Where Recno() = m.lnRecNo		;
			Into Array laField
	
		Do While m.laField[1] # '_MSYSMENU'
			lnLevelName = m.laField[1]
			Select  Recno()										;
				From (m.lcDBF)									;
				Where objCode = 0								;
					And Trim(Name) = Trim(m.lnLevelName)		;
				Into Array laParent
			If _Tally = 0
				Exit
			Endif
			lnRecNo = m.laParent[1] - 1
			Select  LevelName,					;
					Prompt						;
				From (m.lcDBF)					;
				Where Recno() = m.lnRecNo		;
				Into Array laField
			lcPrompt = Alltrim(m.laField[2]) + ' => ' + m.lcPrompt
		Enddo
	
		Return m.lcPrompt
	Endproc
	
	*----------------------------------------------------------------------------------
	Procedure GetMenuKeystrokes(lcFileToEdit, lnRecNo, lcMatchType)
	
		*** JRN 2024-02-05 : Retrieves keystrokes to navigate an MNX down to the
		*   record for <lnRecno>
	
		*** JRN 2024-02-04 : Apparently, pausing briefly between keystrokes is necessary
		#Define ccDownArrow '{Pause .2}{DnArrow}'
		#Define ccTab		'{Pause .2}{Tab}'
		#Define ccEnter		'{Pause .2}{Enter}'
	
		Local laField[1], laParent[1], lcKeystrokes, llSuccess, lnLevelName, lnSelect
	
		lcKeystrokes = ''
		lnSelect	 = Select()
	
		*** JRN 2024-02-04 : Only works if we can open the file
		Select 0
		Try
			Use (m.lcFileToEdit) In 0 Alias GF_Menu
			llSuccess = .T.
		Catch
			llSuccess = .F.
		Endtry
	
		If m.llSuccess
	
			Select  LevelName,					;
					Prompt,						;
					Int(Val(ItemNum))			;
				From GF_Menu					;
				Where Recno() = m.lnRecNo		;
				Into Array laField
	
			If _Tally = 0
				lcKeystrokes = ''
			Endif
	
			* down arrow to get to our record
			If m.laField[3] > 1
				lcKeystrokes = Replicate(ccDownArrow, m.laField[3] - 1)
			Endif
	
			* and if a procedure or command, tab over to it
			Do Case
				Case Upper(m.lcMatchType) = '<COMMAND>'
					lcKeystrokes = m.lcKeystrokes + ccTab + ccTab
				Case Upper(m.lcMatchType) = '<PROCEDURE>'
					lcKeystrokes = m.lcKeystrokes + ccTab + ccTab + ccEnter
			Endcase
	
			Do While m.laField[1] # '_MSYSMENU'
				lnLevelName = m.laField[1]
				Select Recno() From GF_Menu Where objCode = 0 And Trim(Name) = Trim(m.lnLevelName) Into Array laParent
				If _Tally = 0
					lcKeystrokes = ''
					Exit
				Endif
				lnRecNo = m.laParent[1] - 1
				Select  LevelName,					;
						Prompt,						;
						Int(Val(ItemNum))			;
					From GF_Menu					;
					Where Recno() = m.lnRecNo		;
					Into Array laField
				If _Tally = 0
					lcKeystrokes = ''
					Exit
				Endif
	
				* tab over to the submenu definition
				lcKeystrokes = ccTab + ccTab + ccEnter + m.lcKeystrokes
	
				* down arrow to get to our record
				If m.laField[3] > 1
					lcKeystrokes = Replicate(ccDownArrow, m.laField[3] - 1) + m.lcKeystrokes
				Endif
	
			Enddo
	
		Else
	
			Return ''
	
		Endif
	
		Use
		Select (m.lnSelect)
	
		Return m.lcKeystrokes
	Endproc
			
*----------------------------------------------------------------------------------
	Procedure GetProcedureStartPositions(tcCode, tcName)

		Local;
			lcBaseClass As String,;
			lcClassName As String,;
			lcMatch    As String,;
			lcName     As String,;
			lcParentClass As String,;
			lcType     As String,;
			lcWord1    As String,;
			llClassDef As Boolean,;
			llTextEndText As Boolean,;
			lnEndByte  As Number,;
			lnI        As Number,;
			lnLFs      As Number,;
			lnStartByte As Number,;
			lnX        As Number,;
			loException As Object,;
			loMatch    As Object,;
			loMatches  As Object,;
			loObject   As 'Empty',;
			loRegExp   As Object,;
			loResult   As 'Collection'

*
* Original code provided by Jim R Nelson circa March 2011
* Returns a collection indicating the beginning of each procedure / class / etc
* Each member in the collection has these properties:
*   .Type == 'Procedure'(Procedures and Functions))
*         == 'Class'    (Class Definition)
*         == 'End Class'(End of Class Definition)
*         == 'Method'   (Procedures and Functions within a class)
*   .StartByte == starts at zero; thus, # of chars preceding start position
*   .Name
*   .containingclass

****************************************************************
		loRegExp = This.oRegExForProcedureStartPositions

		loMatches = loRegExp.Execute(m.tcCode)

		loResult = Createobject('Collection')

		llClassDef    = .F. && currently within a class?
		llTextEndText = .F. && currently within a Text/EndText block?
		lcClassName   = ''
		lcParentClass = ''
		lcBaseClass   = ''

		For lnI = 1 To m.loMatches.Count

			loMatch = loMatches.Item(m.lnI - 1)

			With m.loMatch
				lnStartByte = .FirstIndex
				lcMatch     = Chrtran(.Value, CR + lf, '  ')
				lcName      = Getwordnum(m.lcMatch, Getwordcount(m.lcMatch))
				lcWord1     = Upper(Getwordnum(m.lcMatch, Max(1, Getwordcount(m.lcMatch) - 1)))
			Endwith

			Do Case
				Case m.llTextEndText
					If 'ENDTEXT' = m.lcWord1
						llTextEndText = .F.
					Endif
					Loop

				Case m.llClassDef
					If 'ENDDEFINE' = m.lcWord1
						llClassDef    = .F.
						lcType        = 'End Class'
						lcName        = m.lcClassName + '.-EndDefine'
						lcClassName   = ''
						lcParentClass = ''
						lcBaseClass   = ''
					Else
						lcType = 'Method'
						lcName = m.lcClassName + '.' + m.lcName
					Endif

				Case ' CLASS ' $ Upper(m.lcMatch) && Notice the spaces in ' CLASS '
					llClassDef    = .T.
					lcType        = 'Class'
					lcClassName   = Getwordnum(m.lcMatch, 3)
					lcParentClass = Getwordnum(m.lcMatch, 5)
					lcName        = ''
					lcBaseClass   = ''
					If This.IsBaseclass(m.lcParentClass)
						lcBaseClass   = Lower(m.lcParentClass)
						lcParentClass = ''
					Endif

				Case 'FUNCTION' = m.lcWord1
					lcType = 'Function'

				Otherwise
					lcType = 'Procedure'

			Endcase

			lnLFs = Occurs(Chr(10), m.loMatch.Value)
			lnX   = 0
* ignore leading CRLF's, and [spaces and tabs, except on the matched line]
			Do While Substr(m.tcCode, m.lnStartByte + 1, 1) $ Chr(10) + Chr(13) + Chr(32) + Chr(9) And m.lnX < m.lnLFs
				If Substr(m.tcCode, m.lnStartByte + 1, 1) = Chr(10)
					lnX = m.lnX + 1
				Endif
				lnStartByte = m.lnStartByte + 1
			Enddo

			loObject = Createobject('GF_Procedure')

			With m.loObject
				.Type         = m.lcType
				.StartByte    = m.lnStartByte
				._Name        = m.lcName
				._ClassName   = m.lcClassName
				._ParentClass = m.lcParentClass
				._BaseClass   = m.lcBaseClass
			Endwith

			Try
					loResult.Add(m.loObject, m.lcName)
				Catch To m.loException When m.loException.ErrorNo = 2062 Or m.loException.ErrorNo = 11
*loResult.Add(loObject, lcName + ' ' + Transform(lnStartByte))
					loResult.Add(m.loObject, m.lcName + Sys(2015))
				Catch To m.loException
					This.ShowErrorMsg(m.loException)
			Endtry


		Endfor

*** JRN 11/09/2015 : determine ending byte for each entry
		lnEndByte = Len(m.tcCode)
		For lnI = m.loResult.Count To 1 Step - 1
			loResult[m.lnI].EndByte = m.lnEndByte
			lnEndByte = 	loResult[m.lnI].StartByte
		Endfor

		Return m.loResult

	Endproc


*----------------------------------------------------------------------------------
	Procedure GetRegExForProcedureStartPositions()

		Local;
			lcPattern As String,;
			loRegExp As 'VBScript.RegExp'

		loRegExp = GF_GetRegExp()

		With m.loRegExp
			.IgnoreCase = .T.
			.Global     = .T.
			.MultiLine  = .T.
		Endwith

		lcPattern = 'PROC(|E|ED|EDU|EDUR|EDURE)\s+(\w|\.)+'
		lcPattern = m.lcPattern + '|' + 'FUNC(|T|TI|TIO|TION)\s+(\w|\.)+'
		lcPattern = m.lcPattern + '|' + 'DEFINE\s+CLASS\s+\w+\s+\w+\s+\w+'
		lcPattern = m.lcPattern + '|' + 'DEFI\s+CLAS\s+\w+'
		lcPattern = m.lcPattern + '|' + 'ENDD(E|EF|EFI|EFIN|EFINE)\s+'
		lcPattern = m.lcPattern + '|' + 'PROT(|E|EC|ECT|ECTE|ECTED)\s+\w+\s+\w+'
		lcPattern = m.lcPattern + '|' + 'HIDD(|E|EN)\s+\w+\s+\w+'

		With m.loRegExp
			.Pattern	= '^\s*(' + m.lcPattern + ')'
		Endwith

		Return m.loRegExp

	Endproc


*----------------------------------------------------------------------------------
	Procedure GetRegExForSearch
	
		Local loRegEx As 'VBScript.RegExp'
	
		loRegEx = GF_GetRegExp()
	
		Return m.loRegEx
	
	Endproc
	

*----------------------------------------------------------------------------------
	Procedure GetReplaceResultObject

		Local;
			loResult As 'Empty'

		loResult = Createobject('Empty')
		AddProperty(m.loResult, 'lError', .F.)
		AddProperty(m.loResult, 'nErrorCode', GF_REPLACE_NOTTTOUCHED)
		AddProperty(m.loResult, 'nChangeLength', 0)
		AddProperty(m.loResult, 'cNewCode', '')
		AddProperty(m.loResult, 'cReplaceLine', '')
		AddProperty(m.loResult, 'cTrimmedReplaceLine', '')
		AddProperty(m.loResult, 'lReplaced', .F.)
		Return m.loResult

	Endproc


*----------------------------------------------------------------------------------
	Procedure GetReplaceRiskLevel(toObject)

		Local;
			lcMatchType As String,;
			lnReturn As Number

		lcMatchType = m.toObject.MatchType

		lnReturn = 4 && Assume everything is very risky to start with !!!

		Do Case

			Case Inlist(m.lcMatchType, MatchType_Name, MATCHTYPE_CONSTANT, '<Parent>', ;
					MATCHTYPE_PROPERTY_DEF, MATCHTYPE_PROPERTY_DESC, MATCHTYPE_PROPERTY_NAME, ;
					MATCHTYPE_PROPERTY, MATCHTYPE_PROPERTY_VALUE, ;
					MATCHTYPE_METHOD_DEF, MATCHTYPE_METHOD_DESC, MatchType_Method, ;
					MATCHTYPE_MPR )

				lnReturn = 3

			Case Inlist(m.lcMatchType, MATCHTYPE_INCLUDE_FILE, '<Expr>', '<Supexpr>', '<Picture>', '<Prompt>', '<Procedure>', ;
					'<Skipfor>', '<Message>', '<Tag>', '<Tag2>');
					Or ;
					M.toObject.UserField.FileType = 'DBF'

				lnReturn = 2

			Case Inlist(m.lcMatchType, MATCHTYPE_CODE, MATCHTYPE_COMMENT) Or;
					(m.toObject.UserField.IsText And !Inlist(m.lcMatchType, MATCHTYPE_FILENAME, MATCHTYPE_TIMESTAMP))

				lnReturn = 1

		Endcase

		Return m.lnReturn

	Endproc


*----------------------------------------------------------------------------------
	Procedure HighlightStatementFilter(tcCode, tcProcFilter, tcMatchWordPrefix, tcMatchWordSuffix, tcMatchLinePrefix, tcMatchLineSuffix)
*just for now
*to do
*get statement snippet
*run HighlightProcFilter for the value of the snippet
*stuff into tcCode

		Local;
			lcReturn As String,;
			loMatch As Object,;
			loRegExp As Object

		loRegExp = GF_GetRegExp()

		loRegExp.IgnoreCase       = .T.
		loRegExp.MultiLine        = .T.
		loRegExp.ReturnFoxObjects = .T.
*		loRegExp.AutoExpandGroups = .T.
		loRegExp.Singleline       = .T.
		loRegExp.Pattern          = loRegExp.Escape(m.tcMatchLinePrefix) + "(.*)" + loRegExp.Escape(m.tcMatchLineSuffix)
		loMatch                   = loRegExp.Match(m.tcCode)

		lcReturn = m.tcCode
		If loMatch.Groups(2).Success Then
			lcReturn = This.HighlightProcFilter(loMatch.Groups(2).Value, m.tcProcFilter, m.tcMatchWordPrefix, m.tcMatchWordSuffix)
			lcReturn = Stuff(m.tcCode, loMatch.Groups(2).Index, loMatch.Groups(2).Length, m.lcReturn)
		Endif &&loMatch.Groups(1).Success

		Return m.lcReturn

	Endproc &&HighlightStatementFilter(tcCode, tcProcFilter, tcMatchWordPrefix, tcMatchWordSuffix, tcMatchLinePrefix, ...

	Procedure HighlightProcFilter(tcCode, tcProcFilter, tcMatchWordPrefix, tcMatchWordSuffix)

		#Define VISIBLE_AND   '|and|'
		#Define VISIBLE_OR    '|or|'

		#Define AND_DELIMITER Chr[255]
		#Define OR_DELIMITER  Chr[254]

		Local;
			lcCode     As String,;
			lcMatch    As String,;
			lcPattern  As String,;
			lcProcFilter As String,;
			lcValue    As String,;
			lcxx       As String,;
			lnATC      As Number,;
			lnCount    As Number,;
			lnFilterCount As Number,;
			lnI        As Number,;
			lnJ        As Number,;
			lnMatch    As Number,;
			loMatch    As Object,;
			loMatches  As Object,;
			loRegExp   As Object

		Local Array;
			laList(1),;
			laValues(1)

		lcValue = Strtran(m.tcProcFilter, '&', '&amp;')
		lcValue = Strtran(m.lcValue, '<', '&lt;')
		lcValue = Strtran(m.lcValue, '>', '&gt;')

		If '|' $ m.lcValue
			lcValue = Alltrim(Upper(m.lcValue))
			lcValue = Strtran(m.lcValue, VISIBLE_AND, AND_DELIMITER, 1, 100, 1)
			lcValue = Strtran(m.lcValue, VISIBLE_OR, OR_DELIMITER, 1, 100, 1)
			lcValue = Strtran(m.lcValue, '|', OR_DELIMITER, 1, 100, 1)
		Else
			lcValue = m.lcValue
		Endif

		lnFilterCount = Alines(laValues, m.lcValue, 0, OR_DELIMITER, AND_DELIMITER)
		lcCode        = m.tcCode

		loRegExp = GF_GetRegExp()
		If m.lnFilterCount = 1 Then
			lcPattern = "|" + loRegExp.Escape(m.lcValue)

		Else  &&m.lnFilterCount = 1
			lcPattern  = ""

			For lnJ = 1 To m.lnFilterCount
*				lcIgnore  = m.lcIgnore  + "|(?:" + m.laValues[m.lnJ] + ")"
*				lcIgnore  = m.lcIgnore  + "|" + m.laValues[m.lnJ]
*				lcInclude = m.lcInclude + "|(" + m.laValues[m.lnJ] + ")"
				lcPattern = m.lcPattern + "|" + loRegExp.Escape(laValues[m.lnJ])
			Endfor &&lnJ
		Endif &&m.lnFilterCount = 1

*		lcIgnore = Substr(m.lcIgnore,2)
*		lcInclude = "|(" + Substr(m.lcInclude,2) + ")"
		lcPattern = Substr(m.lcPattern ,2)

		loRegExp.IgnoreCase       = .T.
		loRegExp.MultiLine        = .T.
		loRegExp.ReturnFoxObjects = .T.
*		loRegExp.AutoExpandGroups = .T.
		loRegExp.Pattern          = "\<[^\>]*?(?:" + m.lcPattern + ")+?.*?\>|(" + m.lcPattern + ")"
		loMatches                 = loRegExp.Matches(m.tcCode)
*		_cliptext = loRegExp.Show_Unwind(m.loMatches)

		lcxx = m.tcCode
		For lnMatch =  m.loMatches.Count To 1 Step -1
			loMatch = loMatches.Item(m.lnMatch)
			If loMatch.Groups(2).Success Then
				lcxx = Stuff(m.lcxx,loMatch.Groups(2).Index,loMatch.Groups(2).Length,m.tcMatchWordPrefix + loMatch.Groups(2).Value + m.tcMatchWordSuffix)
			Endif &&loMatch.Groups(2).Success
		Endfor &&lnMatch

		Return m.lcxx

*escape the pattern(s) (this means, we need to know if this pattern is meant as like or regexp?)
*if there is more then one pattern we need the tag pattern with alternating patterns and many groups
*matches (no lines, caseinsensitive)
*trough all matches backwards
* if a group exists
* get index and lenght
* stuff the place with tcMatchWordPrefix+ value+ tcMatchWordSuffix (so we keep case)


		For lnJ = 1 To m.lnFilterCount
			lcProcFilter = laValues[m.lnJ]
			lnCount      = 0
			For lnI = 1 To 10000
				lnATC = Atc(m.lcProcFilter, m.tcCode, m.lnI)
				If m.lnATC = 0
					Exit
				Endif
				lcMatch = Substr(m.tcCode, m.lnATC, Len(m.lcProcFilter))
* items in this array are all case-sensitive combinations found
* so that the STRTRAN farther down keeps the original case
				If m.lnCount = 0 Or Ascan(m.laList, m.lcMatch, 1) = 0
					lnCount = m.lnCount + 1
					Dimension laList[m.lnCount]
					laList[m.lnCount] = m.lcMatch
				Endif
			Endfor

			For lnI = 1 To m.lnCount
				lcMatch = laList[m.lnI]
				lcCode  = Strtran(m.lcCode, m.lcMatch, m.tcMatchWordPrefix + m.lcMatch + m.tcMatchWordSuffix, 1, 1000)
			Endfor
		Endfor

		Return m.lcCode

	Endproc


*----------------------------------------------------------------------------------
	Procedure HtmlEncode(tcCode)

		Local;
			lcHTML As String,;
			loHTML As 'htmlcode' Of 'mhhtmlcode.prg'

*-- See: http://www.universalthread.com/ViewPageNewDownload.aspx?ID=9679
*-- From: Michael Helland - mobydikc@gmail.com

		loHTML = Newobject('htmlcode', 'mhhtmlcode.prg')
		lcHTML = loHTML.PRGToHTML(m.tcCode)

		Return m.lcHTML

	Endproc


*----------------------------------------------------------------------------------
	Procedure IncrementProgressBar(tnAmount)

		If Vartype(This.oProgressBar) = 'O'
			This.oProgressBar.nValue = This.oProgressBar.nValue + m.tnAmount
		Endif

	Endproc


*----------------------------------------------------------------------------------
	Procedure Init(tlPreserveExistingResults, tcCR_StoreLocal)

		#Include ..\BuildGoFish.h

*SF 20221018 -> local storage
		If !Empty(m.tcCR_StoreLocal) Then
			This.cCR_StoreLocal       = m.tcCR_StoreLocal
		Endif &&!Empty(m.tcCR_StoreLocal) Then

		This.cFilesToSkipFile     = This.cCR_StoreLocal + 'GF_Files_To_Skip.txt'
*/SF 20221018 -> local storage

		This.cVersion        = GOFISH_VERSION  && Comes from include file above
		This.oFSO            = Createobject("Scripting.FileSystemObject")
		This.oRegExForSearch = This.GetRegExForSearch()

		If Isnull(This.oRegExForSearch)
			Messagebox('Error creating oRegExForSearch')
			Return .F.
		Endif

		This.oRegExForSearchInCode = This.GetRegExForSearch()
		If Isnull(This.oRegExForSearchInCode)
			Messagebox('Error creating oRegExForSearchInCode')
			Return .F.
		Endif
		
		This.oRegExForCommentSearch = This.GetRegExForSearch()
		If Isnull(This.oRegExForCommentSearch)
			Messagebox('Error creating oRegExForCommentSearch')
			Return .F.
		Endif
		This.oRegExForCommentSearch.Pattern = '^\s*(\*|NOTE|&' + '&)'	&& Set default-pattern for searching comments

		This.oRegExForProcedureStartPositions = This.GetRegExForProcedureStartPositions()
		If Isnull(This.oRegExForProcedureStartPositions)
			Messagebox('Error creating oRegExForProcedureStartPositions')
			Return .F.
		Endif

		This.BuildProjectsCollection()

		This.oSearchOptions = Createobject(This.cSearchOptionsClass)

		This.oSearchErrors  = Createobject('Collection')
		This.oReplaceErrors = Createobject('Collection')

*-- An FFC class used to generate a TimeStamp so the TimeStamp field can be updated when replacing code in a table based file.
		This.oFrxCursor = Newobject('FrxCursor', Home() + '\ffc\_FrxCursor')

		This.PrepareForSearch()

	Endproc


*----------------------------------------------------------------------------------
	Procedure IsBaseclass(tcString)

		Local;
			lcBaseclasses As String

*-- Note: Each word below contains a space at the beginning and end of the word so the final match test
*-- wil not return .t. for partial matches.

		TEXT To m.lcBaseclasses Noshow
			 CheckBox
			 Collection
			 Column
			 ComboBox
			 CommandButton
			 CommandGroup
			 Container
			 Control
			 Cursor
			 CursorAdapter
			 Custom
			 DataEnvironment
			 EditBox
			 Empty
			 Exception
			 Form
			 FormSet
			 Grid
			 Header
			 Hyperlink
			 Image
			 Label
			 Line
			 ListBox
			 OLEBound
			 OLEContainer
			 OptionButton
			 OptionGroup
			 Page
			 PageFrame
			 ProjectHook
			 Relation
			 ReportListener
			 Separator
			 SessionObject
			 Shape
			 Spinner
			 TextBox
			 Timer
			 ToolBar
			 XMLAdapter
			 XMLField
			 XMLTable
		ENDTEXT

		Return  Upper((' ' + Alltrim(m.tcString) + ' ')) $ Upper(m.lcBaseclasses)


	Endproc


*----------------------------------------------------------------------------------
	Procedure IsComment(tcLine)

		Local;
			lcLine As String,;
			llReturn As Boolean,;
			lnCount As Number,;
			loMatches As Object,;
			loRegEx As Object

		llReturn = This.IsFullLineComment(m.tcLine)

		If m.llReturn
			Return .T.
		Endif

*-- Look for a match BEFORE any && comment characters
		lnCount = Atc('&' + '&', m.tcLine)

		If m.lnCount > 0
			lcLine    = Left(m.tcLine, m.lnCount - 1)
			loMatches = This.oRegExForSearch.Execute(m.lcLine)

			If m.loMatches.Count > 0
				Return .F.
			Endif
		Else
			Return .F.
		Endif

	Endproc


*----------------------------------------------------------------------------------
	Procedure IsContinuation(lcLine)

		Local;
			lnAT As Number

		If ';' = Right(Alltrim(m.lcLine, 1, ' ', Tab, CR, lf), 1)
			Return .T.
		Else
			lnAT = Rat('&' + '&', m.lcLine)
			If m.lnAT > 0 And  Right(Alltrim(Left(m.lcLine, m.lnAT - 1), 1, ' ', Tab), 1) = ';'
				Return .T.
			Endif
		Endif
		Return .F.

	Endproc


*----------------------------------------------------------------------------------
	Procedure IsFileIncluded(tcFile)

*SF Depricated
*!*			If This.oSearchOptions.lIncludeAllFileTypes
*!*				Return .T.
*!*			Endif

* Two (i.e.three) ways to include all, set cOtherIncludes, this is text below to file types in advanced, to * -> all
* or set the Templates (the textbox above) to all -> *.* or old style *
		If Empty(This.oSearchOptions.cFileTemplate)
*only if no template, because this disables extensions

			If !Empty(This.oSearchOptions.cOtherIncludes) And;
					"*" == This.oSearchOptions.cOtherIncludes
*check for all-in-cOtherIncludes
				Return .T.
			Endif

			lcFileType = Upper(Justext(m.tcFile))

*!*			If !Empty(Justext(This.oSearchOptions.cFileTemplate))
*!*				Return This.MatchTemplate(m.lcFileType, Justext(This.oSearchOptions.cFileTemplate))
*!*			Endif

*SF 20230620, it is with extension turned on or not no need for the group
*-- Table-based Files --------------------------------------
*!*			If Inlist(m.lcFileType, 'SCX', 'VCX', 'FRX', 'DBC', 'MNX', 'LBX', 'PJX')
*!*				Return Icase(m.lcFileType = 'SCX', This.oSearchOptions.lIncludeSCX, ;
*!*					M.lcFileType = 'VCX', This.oSearchOptions.lIncludeVCX, ;
*!*					M.lcFileType = 'FRX', This.oSearchOptions.lIncludeFRX, ;
*!*					M.lcFileType = 'DBC', This.oSearchOptions.lIncludeDBC, ;
*!*					M.lcFileType = 'MNX', This.oSearchOptions.lIncludeMNX, ;
*!*					M.lcFileType = 'LBX', This.oSearchOptions.lIncludeLBX, ;
*!*					M.lcFileType = 'PJX', This.oSearchOptions.lIncludePJX, ;
*!*					.F.)		&& Last ".F." is a default value, in case one check is missing in this ICASE()
*!*			Endif

*!*	*-- Code based files ----------------------------------
*!*			If Inlist(m.lcFileType, 'PRG', 'MPR', 'TXT', 'INI', 'H', 'XML', 'SPR', 'ASP', 'JSP', 'JAVA')
*!*				Return Icase(m.lcFileType = 'PRG', This.oSearchOptions.lIncludePRG, ;
*!*					M.lcFileType = 'MPR', This.oSearchOptions.lIncludeMPR, ;
*!*					M.lcFileType = 'TXT', This.oSearchOptions.lIncludeTXT, ;
*!*					M.lcFileType = 'INI', This.oSearchOptions.lIncludeINI, ;
*!*					M.lcFileType = 'H', This.oSearchOptions.lIncludeH, ;
*!*					M.lcFileType = 'XML', This.oSearchOptions.lIncludeXML, ;
*!*					M.lcFileType = 'SPR', This.oSearchOptions.lIncludeSPR, ;
*!*					M.lcFileType = 'ASP', This.oSearchOptions.lIncludeASP, ;
*!*					M.lcFileType = 'JSP', This.oSearchOptions.lIncludeJSP, ;
*!*					M.lcFileType = 'JAVA', This.oSearchOptions.lIncludeJAVA, ;
*!*					.F.)		&& Last ".F." is a default value, in case one check is missing in this ICASE()
*!*			Endif

*-- Table-based Files --------------------------------------
			If (This.oSearchOptions.lIncludeSCX And m.lcFileType = 'SCX');
					Or (This.oSearchOptions.lIncludeVCX  And m.lcFileType = 'VCX' );
					Or (This.oSearchOptions.lIncludeFRX  And m.lcFileType = 'FRX' );
					Or (This.oSearchOptions.lIncludeDBC  And m.lcFileType = 'DBC' );
					Or (This.oSearchOptions.lIncludeMNX  And m.lcFileType = 'MNX' );
					Or (This.oSearchOptions.lIncludeLBX  And m.lcFileType = 'LBX' );
					Or (This.oSearchOptions.lIncludePJX  And m.lcFileType = 'PJX' );
					Or (This.oSearchOptions.lIncludePRG  And m.lcFileType = 'PRG' );
					Or (This.oSearchOptions.lIncludeMPR  And m.lcFileType = 'MPR' );
					Or (This.oSearchOptions.lIncludeTXT  And m.lcFileType = 'TXT' );
					Or (This.oSearchOptions.lIncludeINI  And m.lcFileType = 'INI' );
					Or (This.oSearchOptions.lIncludeH    And m.lcFileType = 'H'   );
					Or (This.oSearchOptions.lIncludeXML  And m.lcFileType = 'XML' );
					Or (This.oSearchOptions.lIncludeSPR  And m.lcFileType = 'SPR' );
					Or (This.oSearchOptions.lIncludeASP  And m.lcFileType = 'ASP' );
					Or (This.oSearchOptions.lIncludeJSP  And m.lcFileType = 'JSP' );
					Or (This.oSearchOptions.lIncludeJAVA And m.lcFileType = 'JAVA')
				Return .T.
			Endif
*/SF 20230620, it is with extension turned on or not no need for the group

*-- Code based files (any HTM* file) ----------------------------------
			If This.oSearchOptions.lIncludeHTML And 'HTM' $ m.lcFileType
				Return .T.
			Endif

*-- Lastly, is it match with other includes??? (but only if no template, because this disables extensions)
			If !Empty(This.oSearchOptions.cOtherIncludes) And m.lcFileType $ Upper(This.oSearchOptions.cOtherIncludes)
				Return .T.
			Endif

		Else  &&Empty(This.oSearchOptions.cFileTemplate)
*just template

*check for all-in-template (no regexp, faster)
			If "*.*" == This.oSearchOptions.cFileTemplate Or "*" == This.oSearchOptions.cFileTemplate
				Return .T.
			Endif
*SF 20230619
** try for full file patterns in cFileTemplate
			If !Empty(This.oSearchOptions.cFileTemplate)
*			 SET STEP ON 
				Return This.oSearchOptions.oRegExpFileTemplate.IsMatch(Justfname(m.tcFile))
			Endif

		Endif &&Empty(This.oSearchOptions.cFileTemplate)

*** No matching filetype found => so don't include in this search!
		Return .F.
	Endproc


*----------------------------------------------------------------------------------
	Procedure IsFullLineComment(tcLine)

*-- See if the entire line is a comment
		Local;
			loMatches As Object

		loMatches = This.oRegExForCommentSearch.Execute(m.tcLine)

		If m.loMatches.Count > 0
			Return .T.
		Else
			Return .F.
		Endif

	Endproc


*----------------------------------------------------------------------------------
	Procedure IsTextFile(tcFile)

		Local;
			lcExt     As String,;
			llIsTextFile As Boolean

		If Empty(m.tcFile)
			Return .F.
		Endif

		lcExt = Upper(Justext(m.tcFile))

		llIsTextFile = !(m.lcExt $ This.cTableExtensions)

		Return m.llIsTextFile

	Endproc


	*----------------------------------------------------------------------------------
	Procedure IsWildCardStatementSearch
		Return This.oSearchOptions.nSearchMode = GF_SEARCH_MODE_LIKE and '*' $ This.oSearchOptions.cSearchExpression
	EndProc 	


*----------------------------------------------------------------------------------
	Procedure LoadOptions(tcFile)

		Local;
			lcProperty As String,;
			loMy    As 'My' Of 'My.vcx'

		Local Array;
			laProperties(1)

*:Global;
x

		If !File(m.tcFile)
			Return .F.
		Endif

*-- Get an array of properties that are on the SearchOptions object
		Amembers(laProperties, This.oSearchOptions, 0, 'U')

*-- Load settings from file...
		loMy = Newobject('My', 'My.vcx')
		loMy.Settings.Load(m.tcFile)

*--- Scan over Object properties, and look for a corresponding props on the My Settings object (if present)
		With m.loMy.Settings
			For x = 1 To Alen(m.laProperties)
				lcProperty = laProperties[x]
				If Type('.' + m.lcProperty) <> 'U'
					Store Evaluate('.' + m.lcProperty) To ('This.oSearchOptions.' + m.lcProperty)
				Endif
			Endfor
		Endwith

*-- My.Settings stores Dates as DateTimes, so I need to convert them to just Date datatypes
		Try
				This.oSearchOptions.dTimeStampFrom = Ttod(This.oSearchOptions.dTimeStampFrom)
			Catch
				This.oSearchOptions.dTimeStampFrom = {}
		Endtry

		Try
				This.oSearchOptions.dTimeStampTo = Ttod(This.oSearchOptions.dTimeStampTo)
			Catch
				This.oSearchOptions.dTimeStampTo = {}
		Endtry

		Return .T.

	Endproc


*----------------------------------------------------------------------------------
	Procedure lReadyToReplace_Access

		Local;
			llReturn As Boolean

		llReturn = This.nMatchLines > 0 ;
			And (!Empty(This.oSearchOptions.cReplaceExpression) Or This.oSearchOptions.lAllowBlankReplace) ;
			And !This.lFileHasBeenEdited

		Return m.llReturn

	Endproc


*----------------------------------------------------------------------------------
	Procedure lTimeStampDataProvided_Access

		If This.oSearchOptions.lTimeStamp And !Empty(This.oSearchOptions.dTimeStampFrom) And !Empty(This.oSearchOptions.dTimeStampTo) && If both dates are supplied
			Return .T.
		Else
			Return .F.
		Endif

	Endproc

*----------------------------------------------------------------------------------
	Procedure MatchTemplate(tcString, tcTemplate)

*-- Supports normal wildcard matching with * and ?, just like old DOS matching.

		Local;
			lcString As String,;
			lcTemplate As String,;
			llMatch As Boolean,;
			lnLength As Number

		If Empty(m.tcTemplate) Or m.tcTemplate = '*'
			Return .T.
		Endif

		lcString   = Upper(Alltrim(Juststem(m.tcString)))
		lcTemplate = Upper(Alltrim((m.tcTemplate)))

		llMatch = Like(m.lcTemplate, m.lcString)

		Return m.llMatch


* * Removed 04/08/2012

* 	Do Case
* 		Case (Left(lcTemplate, 1) = '*' and Right(lcTemplate, 1) = '*')
* 			lcTemplate = Alltrim(lcTemplate, '*')
* 			llMatch = lcTemplate $ lcString
* 		Case Atc('*', lcTemplate) = 0
* 			llMatch = (lcTemplate == lcString)
* 		Case Right(lcTemplate, 1) = '*'
* 			lnLength = Len(lcTemplate) - 1
* 			lcTemplate = Left(lcTemplate, lnLength)
* 			llMatch = Left(lcString, lnLength) = lcTemplate
* 		Case Left(lcTemplate, 1) = '*'
* 			lnLength = Len(lcTemplate) - 1
* 			lcTemplate= Right(lcTemplate, lnLength)
* 			llMatch = Right(lcString, lnLength) = lcTemplate
* 	Endcase

* Return llMatch

	Endproc


*SF 20230228 -we call this oboslete
*!*	*----------------------------------------------------------------------------------
*!*	*-- Migrate any exitisting Replace Detail Table up to ver 4.3.022 ----
*!*		Procedure MigrateReplaceDetailTable

*!*			Local lcCsr, lcDataType, lcFieldName, lcTable, llSuccess, lnSelect

*!*			lcTable = This.cReplaceDetailTable
*!*			lcCsr = 'csrGF_ReplaceSchemaTest'

*!*			If File(lcTable)
*!*				lnSelect = Select()
*!*				Select * From (lcTable) Where 0 = 1 Into Cursor &lcCsr

*!*	*** JRN 11/09/2015 : add field ProcEnd if not already there
*!*				This.AddFieldToReplaceTable(lcTable, lcCsr, 'ProcEnd', 'I')
*!*				This.AddFieldToReplaceTable(lcTable, lcCsr, 'ProcCode', 'M')
*!*				This.AddFieldToReplaceTable(lcTable, lcCsr, 'Statement', 'M')
*!*				This.AddFieldToReplaceTable(lcTable, lcCsr, 'StatementStart', 'I')
*!*				This.AddFieldToReplaceTable(lcTable, lcCsr, 'FirstMatchInStatement', 'L')
*!*				This.AddFieldToReplaceTable(lcTable, lcCsr, 'FirstMatchInProcedure', 'L')

*!*				Use In &lcCsr
*!*				Select (lnSelect)
*!*			Endif

*!*		Endproc


*----------------------------------------------------------------------------------
	Procedure OpenTableForReplace(tcFileToOpen, tcCursor, tnResultId)

		Local;
			llReturn As Boolean,;
			lnSelect As Number

		lnSelect = Select()

		If Used(m.tcCursor)
			Use In (m.tcCursor)
		Endif

		Select 0

		Try
				Use (m.tcFileToOpen) Exclusive Alias (m.tcCursor)
				llReturn = .T.
			Catch
				This.SetReplaceError('Cannot open file for exclusive use: ' + Chr(13) + Chr(13), m.tcFileToOpen, m.tnResultId)
				Select (m.lnSelect)
				llReturn = .F.
		Endtry

		Return m.llReturn

	Endproc


*----------------------------------------------------------------------------------
	Procedure PrepareForSearch

		Clear Typeahead

		This.lEscPress            = .F.
		This.lFileNotFound        = .F.
		This.nMatchLines          = 0
		This.nFileCount           = 0
		This.nFilesProcessed      = 0
		This.nSearchTime          = 0
		This.lResultsLimitReached = .F.

		This.PrepareRegExForSearch()

		This.ClearResultsCursor()
		This.ClearResultsCollection()

		This.ClearReplaceSettings()

		This.oSearchErrors  = Createobject('Collection')
		This.oReplaceErrors = Createobject('Collection')
		This.oDirectories   = Createobject('Collection')

		This.SetIncludePattern()

		This.SetFilesToSkip()

	Endproc


	* ================================================================================
	Procedure PrepareForWholeWordSearch(lcText)
	
		Local lInWord, lcLetter, lcResult, lnPos
	
		lcResult = ''
		lcText = this.EscapeSearchExpression(lcText)	

		lInWord	 = .F.
		For lnPos = 1 To Len(m.lcText)
			lcLetter = Substr(m.lcText, m.lnPos, 1)
			If Isalpha(m.lcLetter) Or Isdigit(m.lcLetter) Or m.lcLetter = '_'
				If Not m.lInWord
					lcResult = m.lcResult + '\b'
					lInWord	 = .T.
				Endif
			Else
				If m.lInWord
					lcResult = m.lcResult + '\b'
					lInWord	 = .F.
				Endif
			Endif
			lcResult = m.lcResult + m.lcLetter
		Endfor
	
		If m.lInWord
			lcResult = m.lcResult + '\b'
		EndIf
		
		Return m.lcResult
	Endproc
		


*----------------------------------------------------------------------------------
	Procedure PrepareRegExForReplace

		Local;
			lcPattern As String

		lcPattern = This.oSearchOptions.cEscapedSearchExpression

*If !this.oSearchOptions.lRegularExpression
*-- Need to trim off the pre- and post- wild card characters so we can get back to just the search phrase
		If Left(m.lcPattern, 2) = '.*'
			lcPattern = Substr(m.lcPattern, 3)
		Endif

		If Right(m.lcPattern, 2) = '.*'
			lcPattern = Left(m.lcPattern, Len(m.lcPattern) - 2)
		Endif

*EndIf

		This.oRegExForSearch.Pattern = m.lcPattern

	Endproc


	*----------------------------------------------------------------------------------
	Procedure PrepareRegExForSearch
		Local lcSearchExpression
	
		lcSearchExpression = This.oSearchOptions.cSearchExpression
		*** JRN 2024-02-14 : "Normal" regex for non wild-card searches
		This.PrepareRegExForSearchV2(This.oRegExForSearch, m.lcSearchExpression, .T.)
	
		If This.IsWildCardStatementSearch()
			*** JRN 2024-02-14 : for wild card searches, only search for up to the first *
			If This.oSearchOptions.lMatchWholeWord
				This.oSearchOptions.cWholeWordSearch = This.PrepareForWholeWordSearch(m.lcSearchExpression)
			Endif
			lcSearchExpression		= Left(m.lcSearchExpression, Atc('*', m.lcSearchExpression) - 1)
		Endif
		This.PrepareRegExForSearchV2(This.oRegExForSearchInCode, m.lcSearchExpression, .F.)
	
	Endproc
			
	
	Procedure PrepareRegExForSearchV2(loRegEx, lcSearchExpression, llMain)
	
		Local lcPattern, lcRegexPattern, lcSearchExpression, loRegEx

		With loRegEx

			.IgnoreCase = ! This.oSearchOptions.lMatchCase
			.Global = .T.
			.MultiLine = .T.

			If This.oSearchOptions.lRegularExpression

				If Left(lcSearchExpression, 1) != '^'
					lcSearchExpression = '.*' + lcSearchExpression
				Endif

				If Right(lcSearchExpression, 1) != '$'
					lcSearchExpression = lcSearchExpression + '.*'
				Endif

				lcPattern = lcSearchExpression

			Else

				lcPattern = This.EscapeSearchExpression(lcSearchExpression)

				If This.oSearchOptions.lMatchWholeWord
					lcPattern = '.*\b' + lcPattern + '\b.*'
				Else
					lcPattern = '.*' + lcPattern + '.*'
				Endif

			Endif

			If llMain
				This.oSearchOptions.cEscapedSearchExpression = lcPattern	
			EndIf 

			*-- Need to add some extra markings around lcPattern to use it as the lcRegExpression
			lcRegexPattern = lcPattern

			If Left(lcRegexPattern, 1) != '^'
				lcRegexPattern = '^' + lcRegexPattern
			Endif

			If Right(lcRegexPattern, 1) != '$'
				lcRegexPattern = lcRegexPattern + '$'
			Endif

			.Pattern = lcRegexPattern

		Endwith
		
	EndProc


*----------------------------------------------------------------------------------
	Procedure ProcessInlineComments(toObject)

		Local;
			lcCode          As String,;
			lcComment       As String,;
			lcMatchType     As String,;
			lcTrimmedMatchLine As String,;
			lnCount         As Number,;
			loCodeMatches   As Object,;
			loCommentMatches As Object

		lcTrimmedMatchLine = m.toObject.TrimmedMatchLine
		lcMatchType        = m.toObject.MatchType

		lnCount = Atc('&' + '&', m.lcTrimmedMatchLine)

		If m.lnCount > 0 And This.oSearchOptions.lSearchInComments
			lcCode           = Left(m.lcTrimmedMatchLine, m.lnCount - 1)
			lcComment        = Substr(m.lcTrimmedMatchLine, m.lnCount)
			loCodeMatches    = This.oRegExForSearch.Execute(m.lcCode)
			loCommentMatches = This.oRegExForSearch.Execute(m.lcComment)

			If m.loCodeMatches.Count > 0 And m.loCommentMatches.Count > 0
				toObject.MatchType = MATCHTYPE_COMMENT
				This.CreateResult(m.toObject)
				lcMatchType = m.toObject.UserField.MatchType && Restore to UserField MatchType for further
			Else
				lcMatchType = Iif(m.loCommentMatches.Count > 0, MATCHTYPE_COMMENT, m.toObject.MatchType)
			Endif
		Endif

		toObject.MatchType = m.lcMatchType

	Endproc


*----------------------------------------------------------------------------------
	Procedure ProcessSearchResult(toObject)

		Local;
			lcBaseClass    As String,;
			lcContainingClass As String,;
			lcMatchType    As String,;
			lcMethodName   As String,;
			lcParentClass  As String,;
			lcSaveObjectName As String,;
			lcSave_Baseclass As String,;
			llReturn       As Boolean,;
			loObject       As Object

		lcMatchType = m.toObject.UserField.MatchType

*-- Store these so we can revert back after processing, becuase it's important to reset back
*-- so any further matches in the code can be processed correctly
		With m.toObject.UserField
			lcSaveObjectName  = ._Name
			lcSave_Baseclass  = ._BaseClass
			lcBaseClass       = ._BaseClass
			lcMethodName      = m.toObject.MethodName
			lcParentClass     = ._ParentClass
			lcContainingClass = .ContainingClass
		Endwith

		If m.lcMatchType # MATCHTYPE_FILENAME
			loObject = This.AssignMatchType(m.toObject)
		Else
			loObject = m.toObject
		Endif

		If !Isnull(m.loObject)
			llReturn = This.CreateResult(m.loObject)
		Else
			llReturn = .T.
		Endif

		With m.toObject.UserField
			._Name              = m.lcSaveObjectName
			._BaseClass         = m.lcSave_Baseclass
			._BaseClass         = m.lcBaseClass
			toObject.MethodName = m.lcMethodName
			._ParentClass       = m.lcParentClass
			.ContainingClass    = m.lcContainingClass
		Endwith

		Return m.llReturn
	Endproc

*----------------------------------------------------------------------------------
	Procedure ReduceProgressBarMaxValue(tnReduction)

		Try
				This.oProgressBar.nMaxValue = This.oProgressBar.nMaxValue - m.tnReduction
			Catch
		Endtry

	Endproc

*----------------------------------------------------------------------------------
* See: http://www.west-wind.com/wconnect/weblog/ShowEntry.blog?id=605
************************************************************************
* wwUtils ::  Replace
****************************************
***  Function: Replaces the replace string or expression for
***            any RegEx matches found in a source string
***    Assume: NOTE: very different from native REplace method
***      Pass: lcSource
***            lcRegEx
***            lcReplace   -   String or Expression to replace with
***            llIsExpression - if .T. lcReplace is EVAL()'d
***
***            Expression can use a value of lcMatch to get the
***            current matched string value.
***    Return: updated string
************************************************************************
	Procedure RegExReplace(lcSource, lcRegEx, lcReplace, llIsExpression)

		Local;
			lcMatch As String,;
			lcRepl As String,;
			lnCount As Number,;
			lnX    As Number,;
			loMatch As Object,;
			loMatches As Object,;
			loRegEx As Object

		*** JRN 2024-02-23 : There are some conditions where the RegEx fails, presumably
		*   because there are reserved characters in the search.
		*	Since code is merely to "pretty things up", no harm in just exiting
		Try 
			This.PrepareRegExForSearch()
			This.PrepareRegExForReplace()

			loRegEx = This.oRegExForSearch

			If !Empty(m.lcRegEx)
				loRegEx.Pattern = m.lcRegEx
			Endif

			loMatches = loRegEx.Execute(m.lcSource)

			lnCount = m.loMatches.Count
		Catch to loException
			m.lnCount = 0
		EndTry

		If m.lnCount = 0
			Return m.lcSource
		Endif

		lcRepl = m.lcReplace

*** Note we have to go last to first to not hose relative string indexes of the match
		For lnX = m.lnCount - 1 To 0 Step - 1
			loMatch = loMatches.Item(m.lnX)
			lcMatch = m.loMatch.Value
			If m.llIsExpression
				lcRepl = Eval(m.lcReplace) &&Evaluate dynamic expression each time
			Endif
			lcSource = Stuff(m.lcSource, m.loMatch.FirstIndex + 1, m.loMatch.Length, m.lcRepl)
		Endfor

		Return m.lcSource

	Endproc


*----------------------------------------------------------------------------------
	Procedure RenameColumn(tcTable, tcOldFieldName, tcNewFieldName)

		Local;
			lcAlias As String

		lcAlias = Juststem(m.tcTable)

		If Empty(Field(m.tcNewFieldName, m.lcAlias)) And !Empty(Field(m.tcOldFieldName, m.lcAlias))
			Try
					Alter Table (m.lcAlias) Rename Column (m.tcOldFieldName) To (m.tcNewFieldName)
				Catch
			Endtry
		Endif

	Endproc


*----------------------------------------------------------------------------------
	Procedure ReplaceFromCurrentRow(tcCursor, tcReplaceLine, tnReplaceId)

		Local;
			lcColumn    As String,;
			lcFileToModify As String,;
			lcUni       As String,;
			llBackedUp  As Boolean,;
			lnCurrentRecno As Number,;
			lnMatchStart As Number,;
			lnProcStart As Number,;
			lnResultRecno As Number,;
			lnReturn    As Number,;
			lnSelect    As Number,;
			loReplace   As Object,;
			loResult    As Object

		lnSelect = Select()
		Select (m.tcCursor)
		lnCurrentRecno = Recno()

		If Replaced && If it's already been processed
			Select(m.lnSelect)
			Return GF_REPLACE_FILE_HAS_ALREADY_BEEN_PROCESSED
		Endif

		If !Process And !Empty(ReplaceLine) && Could be that the row was previous marked for replace, and now it has been cleared.
			Replace ReplaceLine With '' In (m.tcCursor)
			Replace TrimmedReplaceLine With '' In (m.tcCursor)
			Select(m.lnSelect)
			Return GF_REPLACE_RECORD_IS_NOT_MARKED_FOR_REPLACE
		Endif

		If !Process && Just not touched
			Select(m.lnSelect)
			Return GF_REPLACE_RECORD_IS_NOT_MARKED_FOR_REPLACE
		Endif

		If !File(FilePath)
			This.SetReplaceError('File not found:', FilePath, Id)
			Select(m.lnSelect)
			Return GF_REPLACE_FILE_NOT_FOUND
		Endif

		If This.oSearchOptions.lBackup
			llBackedUp = This.BackupFile(FilePath, m.tnReplaceId)
			If !m.llBackedUp
				Select(m.lnSelect)
				Return GF_REPLACE_BACKUP_ERROR
			Endif
		Endif

		This.PrepareRegExForSearch() && This will setup the Search part of the RegEx
		This.PrepareRegExForReplace() && This will setup the Replace part of the RegEx

		Scatter Name m.loReplace Memo

		If This.IsTextFile(FilePath)
			loResult = This.ReplaceInTextFile(m.loReplace, m.tcReplaceLine)
		Else
			loResult = This.ReplaceInTable(m.loReplace, m.tcReplaceLine)
		Endif

		If !m.loResult.lError
*-- We must update all match result rows that are of the same source line as this row.
*-- The reason is that search matches can result in multiple rows, and we can't process them again.
			lcFileToModify = FilePath
			lnResultRecno  = Recno
			lnProcStart    = ProcStart
			lnMatchStart   = MatchStart
			lcColumn       = Column
			lcUni          = cUni

			If This.oSearchOptions.lPreviewReplace
				tnReplaceId = 0
			Endif

			Update (m.tcCursor) ;
				Set TrimmedReplaceLine = m.loResult.cTrimmedReplaceLine, ;
				ReplaceLine = m.loResult.cReplaceLine,;
				iReplaceFolder = m.tnReplaceId;
				Where cUni == m.lcUni And ;
				FilePath == m.lcFileToModify And ;
				Recno = m.lnResultRecno And ;
				Column = m.lcColumn And ;
				MatchStart = m.lnMatchStart

			Try
					Goto (m.lnCurrentRecno)
				Catch
			Endtry

			If m.loResult.lReplaced
				This.nReplaceCount = This.nReplaceCount + 1
			Endif

*-- Removed this in 4.3.014. We *do* need to re-compile.
*If !Empty(tcReplaceLine)
			This.Compile(FilePath)
*Endif

			This.UpdateCursorAfterReplace(m.tcCursor, m.loResult)

			lnReturn = GF_REPLACE_SUCCESS

		Else

			lnReturn = 	m.loResult.nErrorCode

		Endif

		Select (m.lnSelect)

		Return m.lnReturn

	Endproc


*----------------------------------------------------------------------------------
	Procedure ReplaceInCode(toReplace, tcReplaceLine)

* tcReplaceLine, if passed, will be used to replace the entire oringinal match line,
* rather than using the RexEx replace with cReplaceExpression on the original line.

* Notes:
* For the Replace, the pattern on the regex must already be set (use PrepareRegExForReplace)
* Note: Unless a full replacement line is passed in tcReplaceLine, ALL instances of the pattern will be replaced on the tcMatchLine

		Local;
			lcCode            As String,;
			lcLeft            As String,;
			lcLineFromFile    As String,;
			lcMatchLine       As String,;
			lcNewCode         As String,;
			lcReplaceExpression As String,;
			lcReplaceLine     As String,;
			lcRight           As String,;
			lnLineToChangeLength As Number,;
			lnMatchStart      As Number,;
			loRegEx           As Object,;
			loResult          As Object

		loResult = This.GetReplaceResultObject()
		lcCode   = m.toReplace.Code

		lcMatchLine = Left(m.toReplace.MatchLine, m.toReplace.MatchLen)

		lnMatchStart         = m.toReplace.MatchStart
		lnLineToChangeLength = Len(m.lcMatchLine)

		lcLineFromFile = Substr(m.lcCode, m.lnMatchStart + 1, m.lnLineToChangeLength)

		If m.lcLineFromFile != m.lcMatchLine && Ensure that line from file still matches the passed in line from the orginal search!!
			This.SetReplaceError('Source file has changed since original search:', Alltrim(m.toReplace.FilePath), m.toReplace.Id)
			loResult.lError = .T.
			Return m.loResult
		Endif

		lcLeft = Left(m.lcCode, m.lnMatchStart)

*-- IMPORTANT CODE HERE... Revised code line is determined here!!!! -------------
		If Empty(m.tcReplaceLine)
			loRegEx             = This.oRegExForSearch
			lcReplaceExpression = This.oSearchOptions.cReplaceExpression
			Do Case
				Case This.nReplaceMode = 1
					lcReplaceLine = loRegEx.Replace(m.lcMatchLine, m.lcReplaceExpression)
				Case This.nReplaceMode = 2
					lcReplaceLine = ''
				Case This.nReplaceMode = 3 And !Empty(This.cReplaceUDFCode)
					lcReplaceLine = This.ReplaceLineWithUDF(m.lcMatchLine)
				Otherwise
					lcReplaceLine = m.lcMatchLine
			Endcase
		Else
			lcReplaceLine = m.tcReplaceLine
		Endif

		lcRight = Substr(m.lcCode, m.lnMatchStart + 1 + m.lnLineToChangeLength)

*--Added this in 4.3.014 to handle case of deleting the entire line
		If Empty(m.lcReplaceLine)
			lcRight = Ltrim(m.lcRight, 0, Chr(10)) && Need to strip off initial Chr(10) of Right hand code block
		Endif

		lcNewCode = m.lcLeft + m.lcReplaceLine + m.lcRight

		With m.loResult
			.nChangeLength = Len(m.lcReplaceLine) - Len(m.lcMatchLine)
*--Added this in 4.3.014 to handle case of deleting the entire line
			If Empty(m.lcReplaceLine)
				.nChangeLength = .nChangeLength - 1 && to account for the Chr(10) we stripped off above
			Endif
			.cNewCode            = m.lcNewCode
			.cReplaceLine        = m.lcReplaceLine
			.cTrimmedReplaceLine = This.TrimWhiteSpace(.cReplaceLine)
		Endwith

		toReplace.ReplaceLine        = m.loResult.cReplaceLine
		toReplace.TrimmedReplaceLine = m.loResult.cTrimmedReplaceLine

		Return m.loResult

	Endproc


*----------------------------------------------------------------------------------
	Procedure ReplaceInTable(toReplace, tcReplaceLine)

		Local;
			lcColumn      As String,;
			lcFileToModify As String,;
			lcMatchLine   As String,;
			lcReplaceCursor As String,;
			llTableWasOpened As Boolean,;
			lnMatchStart  As Number,;
			lnRecNo       As Number,;
			lnResultId    As Number,;
			lnSelect      As Number,;
			loResult      As Object

		lcFileToModify = Alltrim(m.toReplace.FilePath)
		lcMatchLine    = Left(m.toReplace.MatchLine, m.toReplace.MatchLen)
		lnMatchStart   = m.toReplace.MatchStart
		lnResultId     = m.toReplace.Id
		lcColumn       = Alltrim(m.toReplace.Column)
		lnRecNo        = m.toReplace.Recno

		lcReplaceCursor = 'ReplaceCursor'
		lnSelect        = Select()

		loResult = This.GetReplaceResultObject()

*!*	If !File(lcFileToModify)
*!*		This.SetReplaceError('File not found:', lcFileToModify, lnResultId)
*!*		loResult.lError = .t.
*!*		loResult.nErrorCode = GF_REPLACE_FILE_NOT_FOUND
*!*	Endif

		If !This.OpenTableForReplace(m.lcFileToModify, m.lcReplaceCursor, m.lnResultId)
			loResult.lError     = .T.
			loResult.nErrorCode = GF_REPLACE_UNABLE_TO_USE_TABLE_FOR_REPLACE
		Else
			llTableWasOpened = .T.
		Endif

		If !m.loResult.lError
			Try
					Goto m.lnRecNo
				Catch
					This.SetReplaceError('Error locating record in file:', m.lcFileToModify, m.lnResultId)
					loResult.lError     = .T.
					loResult.nErrorCode = GF_REPLACE_ERROR_LOCATING_RECORD_IN_FILE
			Endtry
		Endif

		If !m.loResult.lError
			toReplace.Code = Evaluate(m.lcReplaceCursor + '.' + m.lcColumn)
			loResult       = This.ReplaceInCode(m.toReplace, m.tcReplaceLine)
		Endif

*-- Big step here... Replace code in actual record!!! (If not in Preview Mode)
		If !m.loResult.lError And This.oSearchOptions.lPreviewReplace = .F.
			Replace (m.lcColumn) With m.loResult.cNewCode In (m.lcReplaceCursor) && Update code in table
			If Type('timestamp') != 'U'
				Replace Timestamp With This.oFrxCursor.getFrxTimeStamp() In (m.lcReplaceCursor)
			Endif
			loResult.lReplaced  = .T.

		Endif

		If m.llTableWasOpened
			Use && Close the table based file we opened above
		Endif

		Select (m.lnSelect)

		Return m.loResult

	Endproc


*----------------------------------------------------------------------------------
	Procedure ReplaceInTextFile(toReplace, tcReplaceLine)

		Local;
			lcFileToModify As String,;
			lcOldCode   As String,;
			loReseult   As Object,;
			loResult    As Object

		lcFileToModify = Alltrim(m.toReplace.FilePath)

*!*	If !File(lcFileToModify)
*!*		This.SetReplaceError('File not found:', lcFileToModify, lnResultId)
*!*		loResult = This.GetReplaceResultObject()
*!*		loResult.lError = .t.
*!*		Return loResult
*!*	EndIf

		toReplace.Code = Filetostr(m.lcFileToModify)
		loResult       = This.ReplaceInCode(m.toReplace, m.tcReplaceLine)

		If m.loResult.lError Or This.oSearchOptions.lPreviewReplace
			Return m.loResult
		Endif

*== Big step here... About to replace old file with the new code!!!
		Try
				If !Empty(m.loResult.cNewCode) && Do not dare replace the file with and empty string. Something must be wrong!
					Strtofile(m.loResult.cNewCode, m.lcFileToModify, 0)
					loResult.lReplaced  = .T.

				Endif
			Catch
				This.SetReplaceError('Error saving file: ', m.lcFileToModify, m.toReplace.Id)
		Endtry

		Return m.loResult

	Endproc


*----------------------------------------------------------------------------------
	Procedure ReplaceLine(tcCursor, tnID, tcReplaceLine, tnReplaceId)

		Local;
			lcReplaceLine As String,;
			llReturn   As Boolean,;
			lnLastChar As Number,;
			lnReturn   As Number,;
			lnSelect   As Number

		lnSelect = Select()

		lcReplaceLine = m.tcReplaceLine
		lnLastChar    = Asc(Right(m.lcReplaceLine, 1))

		If m.lnLastChar = 10 && Editbox will add a Chr(10) so this has to be stripped off
			lcReplaceLine = Left(m.lcReplaceLine, Len(m.lcReplaceLine) - 1)
		Endif

		lnLastChar = Asc(Right(m.lcReplaceLine, 1))

		If m.lnLastChar <> 13 And !Empty(m.lcReplaceLine) && Make sure user has not stripped of the Chr(13) that came with the MatchLine
			lcReplaceLine = m.lcReplaceLine + Chr(13)
		Endif

		Select(m.tcCursor)
		Locate For Id = m.tnID

		If Found()

			If Replaced
				Return .T.
			Endif

			Replace Process With .T. In (m.tcCursor)

			lnReturn = This.ReplaceFromCurrentRow(m.tcCursor, m.lcReplaceLine, m.tnReplaceId)

			If m.lnReturn >= 0
				llReturn = .T.
			Else
				llReturn = .F.
			Endif

		Else

			This.SetReplaceError('Error locating record in call to ReplaceLine() method.', '', m.tnID)
			llReturn = .F.

		Endif

		Return m.llReturn

	Endproc


*----------------------------------------------------------------------------------
	Procedure ReplaceLineWithUDF(tcMatchLine)

		Local;
			lcMatchLine As String,;
			lcReplaceLine As String,;
			llCR       As Boolean

		lcMatchLine = m.tcMatchLine

*-- If there is a CR at the end, pull it off before calling the UDF. Will add back later...
		If Right(m.tcMatchLine, 1) = Chr(13)
			llCR        = .T.
			lcMatchLine = Left(m.tcMatchLine, Len(m.tcMatchLine) - 1)
		Endif

*-- Call the UDF ---------------
		Try
				lcReplaceLine = Execscript(This.cReplaceUDFCode, m.lcMatchLine)
			Catch
				lcReplaceLine = m.lcMatchLine && Keep the line the same if UDF failed
			Finally
		Endtry

		If Vartype(m.lcReplaceLine) <> 'C'
			lcReplaceLine = m.lcMatchLine
		Endif

		If m.llCR
			lcReplaceLine = m.lcReplaceLine + Chr(13)
		Endif

		Return m.lcReplaceLine

	Endproc


*----------------------------------------------------------------------------------
	Procedure ReplaceMarkedRows(tcCursor, tnReplaceId)

		Local;
			lcFile           As String,;
			lcFileList       As String,;
			lcLastFile       As String,;
			lcReplaceExpression As String,;
			lcSearchExpression As String,;
			lnResult         As Number,;
			lnSelect         As Number,;
			lnShift          As Number

		This.nReplaceCount     = 0
		This.nReplaceFileCount = 0

		lcSearchExpression  = Alltrim(This.oSearchOptions.cSearchExpression)
		lcReplaceExpression = Alltrim(This.oSearchOptions.cReplaceExpression)
		lnShift             = Len(m.lcReplaceExpression) - Len(m.lcSearchExpression)

		This.oReplaceErrors = Createobject('Collection')

		If Empty(This.oSearchOptions.cReplaceExpression) And !This.oSearchOptions.lAllowBlankReplace
			This.SetReplaceError('Replace expression is blank, but ALLOW BLANK flag is not set.')
			Return .F.
		Endif

		lnSelect = Select()
		Select (m.tcCursor)

		lcLastFile = ''

		Scan

			If Vartype(This.oProgressBar) = 'O'
				This.oProgressBar.nValue = This.oProgressBar.nValue + 1
			Endif

			lnResult = This.ReplaceFromCurrentRow(m.tcCursor, , m.tnReplaceId)

*-- Skip to next file if have had any of there errors:
			If m.lnResult = GF_REPLACE_BACKUP_ERROR Or;
					M.lnResult = GF_REPLACE_UNABLE_TO_USE_TABLE_FOR_REPLACE Or ;
					M.lnResult = GF_REPLACE_FILE_NOT_FOUND
				lcFile = FilePath
				Locate For FilePath <> m.lcFile Rest
				If !Bof()
					Skip - 1
				Endif
			Endif

			If FilePath <> m.lcLastFile And !Empty(m.lcLastFile) && If we are on a new file, then compile the previous file
				This.Compile(m.lcLastFile)
				lcLastFile = ''
			Endif

			If m.lnResult = GF_REPLACE_SUCCESS
				lcLastFile = FilePath
			Endif

		Endscan

*-- Must look at compiling one last time now that loop has ended.
*SF 20230301 -> only file is not empty, we are on the end of the scan anyway
*!*			If FilePath <> lcLastFile And !Empty(lcLastFile) && If we are on a new file, then compile the previous file
*!*				This.Compile(lcLastFile)
*!*			Endif
		If !Empty(m.lcLastFile) && We named the last file in cursor as replaced, compile it
			This.Compile(m.lcLastFile)
		Endif

		Select (m.lnSelect)

		This.ShowWaitMessage('Replace Done.')

	Endproc


*----------------------------------------------------------------------------------
	Procedure RestoreDefaultDir

		Cd (This.cInitialDefaultDir)

	Endproc


*----------------------------------------------------------------------------------
	Procedure SaveOptions(tcFile)

		Local;
			lcProperty As String,;
			loMy    As 'My' Of 'My.vcx'

		Local Array;
			laProperties(1)

*:Global;
x

		loMy = Newobject('My', 'My.vcx')

		Amembers(laProperties, This.oSearchOptions, 0, 'U')

		With m.loMy.Settings

			For x = 1 To Alen(m.laProperties)
				lcProperty = laProperties[x]
				If !Inlist(m.lcProperty, '_MEMBERDATA', 'CPROJECTS', 'OREGEXPFILETEMPLATE')
					.Add(m.lcProperty, Evaluate('This.oSearchOptions.' + m.lcProperty))
				Endif
			Endfor

			.Save(m.tcFile)

		Endwith

	Endproc


*----------------------------------------------------------------------------------
	Procedure SearchFinished(tnSelect)

		This.lResultsLimitReached = (This.nMatchLines >= This.oSearchOptions.nMaxResults)

		This.EndTimer()

		Inkey(.10) && Delay is needed to allow Progress Bar to fully update before it disappears.
		This.StopProgressBar()

		If This.nMatchLines = 0 And This.oSearchOptions.lShowNoMatchesMessage And Not This.lEscPress
			Messagebox('No matches found', 64, 'GoFish')
		Endif

*** JRN 07/10/2016 : ascertain the first match in each statement based on FilePath, Class, MethodName, StatementStart
*SF 20221111 (Helau!)
* Change suggested by Chen, see issue #34
* splitted lon string comaprision into single ones
		Update Results ;
			Set firstmatchinstatement = .T. ;
			From (This.cSearchResultsAlias) As Results ;
			Join (Select FilePath, ;
			Class, ;
			Name, ;
			MethodName, ;
			statementstart, ;
			Min(MatchStart) As MatchStart ;
			From (This.cSearchResultsAlias)            ;
			Group By FilePath, Class, Name, MethodName, statementstart) ;
			As FirstMatch ;
			On Results.FilePath = FirstMatch.FilePath And ;
			Results.Class = FirstMatch.Class And ;
			Results.Name = FirstMatch.Name And ;
			Results.MethodName = FirstMatch.MethodName ;
			And Results.statementstart = FirstMatch.statementstart ;
			And Results.MatchStart = FirstMatch.MatchStart
*!*			Update  Results ;
*!*				Set firstmatchinstatement = .T. ;
*!*				From (This.cSearchResultsAlias)    As  Results ;
*!*				Join (Select  FilePath, ;
*!*						   Class, ;
*!*						   Name, ;
*!*						   MethodName, ;
*!*						   statementstart, ;
*!*						   Min(MatchStart) As  MatchStart ;
*!*					   From (This.cSearchResultsAlias)            ;
*!*					   Group By FilePath, Class, Name, MethodName, statementstart) ;
*!*				 As  FirstMatch ;
*!*				 On Results.FilePath + Results.Class + Results.Name + Results.MethodName ;
*!*				 	= FirstMatch.FilePath + FirstMatch.Class + FirstMatch.Name + FirstMatch.MethodName ;
*!*				 And Results.statementstart = FirstMatch.statementstart ;
*!*				 And Results.MatchStart = FirstMatch.MatchStart
*/SF 20221111 (Helau!)

		Update Results ;
			Set firstmatchinprocedure = .T. ;
			From (This.cSearchResultsAlias) As Results ;
			Join (Select FilePath, ;
			Class, ;
			Name, ;
			MethodName, ;
			Min(MatchStart) As MatchStart ;
			From (This.cSearchResultsAlias) ;
			Group By FilePath, Class, Name, MethodName)           As FirstMatch ;
			On Results.FilePath + Results.Class + Results.Name + Results.MethodName ;
			= FirstMatch.FilePath + FirstMatch.Class + FirstMatch.Name + FirstMatch.MethodName ;
			And Results.MatchStart = FirstMatch.MatchStart

		Select (m.tnSelect)

	Endproc


*----------------------------------------------------------------------------------
	Procedure SearchInCode(tcCode, tuUserField, tlHasProcedures)

		Local;
			lcErrorMessage         As String,;
			lcMatchType            As String,;
			llScxVcx               As Boolean,;
			lnMatchCount           As Number,;
			lnSelect               As Number,;
			loMatch                As Object,;
			loMatches              As Object,;
			loObject               As 'GF_SearchResult',;
			loProcedure            As Object,;
			loProcedureStartPositions As Object

		lnSelect = Select()

		If Empty(m.tcCode)
			Return 0
		Endif
*-- Be sure that oRegExForSearch has been setup... Use This.PrepareRegExForSearch() or roll-your-own
		Try
				loMatches = This.oRegExForSearchInCode.Execute(tcCode)
			Catch
		Endtry

		If Type('loMatches') = 'O'
			lnMatchCount = m.loMatches.Count
		Else
			lcErrorMessage = 'Error processing regular expression.    ' + This.oRegExForSearch.Pattern
			This.SetSearchError(m.lcErrorMessage)
			Return - 1
		Endif

		If m.lnMatchCount > 0

			loProcedureStartPositions = Iif(m.tlHasProcedures, This.GetProcedureStartPositions(m.tcCode), .Null.)
			lnCount = 0

			For Each m.loMatch In m.loMatches FoxObject

				If m.tlHasProcedures And !This.oSearchOptions.lSearchInComments And This.IsComment(m.loMatch.Value)
					Loop
				Endif
				loProcedure = This.FindProcedureForMatch(m.loProcedureStartPositions, m.loMatch.FirstIndex)
				loObject    = Createobject('GF_SearchResult')

				With m.loObject
					.UserField  = m.tuUserField
					.oMatch     = m.loMatch
					.oProcedure = m.loProcedure

					.Type = Proper(.oProcedure.Type)
* .ContainingClass =	.oProcedure._ClassName	&& Not used on this object. This line to be deleted after testing. (2012-07-11))
					.MethodName = .oProcedure._Name
					.ProcStart  = .oProcedure.StartByte
					.procend    = .oProcedure.EndByte
					.proccode   = Substr(m.tcCode, .ProcStart + 1, Max(0, .procend - .ProcStart))

					.MatchLine  = .oMatch.Value
					.MatchStart = .oMatch.FirstIndex
					.MatchLen   = Len(.oMatch.Value)

					If m.tlHasProcedures
						.MatchType            = m.loProcedure.Type && Use what was determined by call to FindProcedureForMatch())
						tuUserField.MatchType = m.loProcedure.Type
					Else
						.MatchType = m.tuUserField.MatchType && Use what was passed.
					Endif

					.Code = Iif(This.oSearchOptions.lStoreCode, m.tcCode, '')
				Endwith

*	Assert Upper(JustExt(Trim(loobject.uSERFIELD.FILENAME)))  # 'PRG'

				This.FindStatement(m.loObject)

				*** JRN 2024-02-14 : for wild card searches, we have only found matches to first "word" (preceding the *)
				* since the remainder of the matches may be on continuation lines and the original search does not
				* search continuation lines, we find the entire statement and search the remainder of the statement
				If This.IsWildCardStatementSearch()
					Local lcOldPattern, lcStatement, lnStartPos, loMatches, loRegEx
					
					lcStatement	= loObject.Statement
					lnStartPos	= Atc(loObject.MatchLine, m.lcStatement)
					If m.lnStartPos > 0
						lcStatement = Substr(m.lcStatement, m.lnStartPos)
					Endif
					Do Case
						Case Not Like('*' + Upper(This.oSearchOptions.cSearchExpression) + '*', Upper(m.lcStatement))
							Loop
						Case This.oSearchOptions.lMatchWholeWord
							loRegEx			= This.oRegExForSearchInCode
							lcOldPattern	= m.loRegEx.Pattern
							loRegEx.Pattern	= This.oSearchOptions.cWholeWordSearch
							loMatches		= m.loRegEx.Execute(Chrtran(m.lcStatement, CR + LF + Tab, '   '))
							loRegEx.Pattern	= m.lcOldPattern
					
							If m.loMatches.Count = 0
								Loop
							Endif
					Endcase
				Endif

				lnCount = lnCount + 1

				If !This.ProcessSearchResult(m.loObject)
					Exit
				Endif &&!This.ProcessSearchResult(loObject)

			Endfor

		Endif

		Select(m.lnSelect)

		Return m.lnMatchCount

	Endproc


*----------------------------------------------------------------------------------
	Procedure SearchInFile(tcFile, tlForce)

*-- Only searches passed file if its file ext is marked for inclusion (i.e. lIncludeSCX)
*-- Optionally, pass tlForce = .t. to force the file to be searched.
		Local;
			llTextFile        As Boolean,;
			lnFileNameMatchCount As Number,;
			lnMatchCount      As Number

*!* ** { JRN -- 11/18/2015 12:16 PM - Begin
*!* If Lastkey() = 27 or Inkey() = 27
		If Inkey() = 27
*!* ** } JRN -- 11/18/2015 12:16 PM - End
			This.lEscPress = .T.
			Clear Typeahead
			Return 0
		Endif

*SF 20230620 will be called in IsFileIncluded anyway
*!*	*-- See if the filename matches the File template filter (if one is set) ----
*!*			If !Empty(This.oSearchOptions.cFileTemplate)
*!*				If !This.MatchTemplate(Justfname(m.tcFile), Juststem(Justfname(This.oSearchOptions.cFileTemplate)))
*!*					This.ReduceProgressBarMaxValue(1)
*!*					Return 0
*!*				Endif
*!*			Endif

		If This.FilesToSkip(m.tcFile)
			Return 0
		Endif

		If !This.IsFileIncluded(m.tcFile) And !m.tlForce
*This.ReduceProgressBarMaxValue(1)
			Return 0
		Endif

		If !File(m.tcFile)
			This.lFileNotFound = .T.
			This.SetSearchError('File not found: ' + m.tcFile)
*This.ReduceProgressBarMaxValue(1)
			Return 0
		Endif

		This.ShowWaitMessage('Processing file: ' + m.tcFile)

*-- Look for a match on the file name ----------------------
		lnFileNameMatchCount = This.SearchInFileName(m.tcFile)

		If m.lnFileNameMatchCount < 0
			Return m.lnFileNameMatchCount
		Endif

		llTextFile = This.IsTextFile(m.tcFile)

*-- Do not search inside of file if we are only looking at timestamps and have and empty string
		If m.llTextFile And This.oSearchOptions.lTimeStamp And Empty(This.oSearchOptions.cSearchExpression)
			This.nFilesProcessed = This.nFilesProcessed + 1
			This.nFileCount      = This.nFileCount + 1
			Return m.lnFileNameMatchCount
		Endif

*-- Look for a match within the file contents ----------------------
		If m.llTextFile
			lnMatchCount = This.SearchInTextFile(m.tcFile)
		Else
			lnMatchCount = This.SearchInTable(m.tcFile)
		Endif

		This.nFilesProcessed = This.nFilesProcessed + 1

		If m.lnMatchCount < 0
			Return m.lnMatchCount
		Endif

*-- Count number of files that had a match by either search above ---
		If m.lnMatchCount > 0 Or m.lnFileNameMatchCount > 0
			This.nFileCount = This.nFileCount + 1
		Endif

		Return m.lnMatchCount + m.lnFileNameMatchCount

	Endproc


*----------------------------------------------------------------------------------
	Procedure SearchInFileName(tcFile)

		Local;
			lcCode            As String,;
			lcErrorMessage    As String,;
			ldFileDate        As Date,;
			ldFromDate        As Date,;
			ldToDate          As Date,;
			llHasMethods      As Boolean,;
			lnMatchCount      As Number,;
			lnSelect          As Number,;
			loFileResultObject As 'GF_FileResult',;
			loMatches         As Object,;
			loSearchResultObject As 'GF_SearchResult'

		lnSelect = Select()

		If !File(m.tcFile)
			This.lFileNotFound = .T.
			This.SetSearchError('File not found: ' + m.tcFile)
			Return 0
		Endif

*-- Be sure that oRegExForSearch has been setup... Use This.PrepareRegExForSearch() or roll-your-own
		Try
				loMatches = This.oRegExForSearch.Execute(Justfname(m.tcFile))
			Catch
		Endtry

		If Type('loMatches') = 'O'
			lnMatchCount = m.loMatches.Count
		Else
			lcErrorMessage = 'Error processing regular expression.    ' + This.oRegExForSearch.Pattern
			This.SetSearchError(m.lcErrorMessage)
			Return - 1
		Endif

		If m.lnMatchCount = 0 And !Empty(This.oSearchOptions.cSearchExpression)
			Return 0
		Endif

		ldFileDate = This.GetFileDateTime(m.tcFile)

		ldFromDate = Evl(This.oSearchOptions.dTimeStampFrom, {^1900-01-01})
		ldToDate   = Evl(This.oSearchOptions.dTimeStampTo, {^9999-01-01})
		ldToDate   = m.ldToDate + 1 &&86400 && Must bump into to next day, since TimeStamp from table has time on it

		If This.oSearchOptions.lTimeStamp And !Between(m.ldFileDate, m.ldFromDate, m.ldToDate)
			Return 0
		Endif

		loFileResultObject = Createobject('GF_FileResult')	&& This custom class has all the properties that must be populated if you want to
&& have a cursor created
		With m.loFileResultObject
			.FileName  = Justfname(m.tcFile)
			.FilePath  = m.tcFile
			.MatchType = MATCHTYPE_FILENAME
			.FileType  = Upper(Justext(m.tcFile))
			.Timestamp = m.ldFileDate

			.MatchLine        = 'File name = "' + .FileName + '"'
			.TrimmedMatchLine = .MatchLine
		Endwith

		loSearchResultObject = Createobject('GF_SearchResult')
		With m.loSearchResultObject
			.UserField        = m.loFileResultObject
			.MatchType        = MATCHTYPE_FILENAME
			.MatchLine        = 'File name = "' + m.loFileResultObject.FileName + '"'
			.TrimmedMatchLine = 'File name = "' + m.loFileResultObject.FileName + '"'
		Endwith

		If This.IsTextFile(m.tcFile)
			*** JRN 2024-02-05 : Correction to show the entire file
			* 	(Previously, leading characters got lost)
			loSearchResultObject.Code = loSearchResultObject.MatchLine 		;
				+ Replicate('=', Max(Len(loSearchResultObject.MatchLine), 60)) + Chr[13]			;
				+ Filetostr(tcFile)
		Endif

		This.ProcessSearchResult(m.loSearchResultObject)

		Select (m.lnSelect)

		Return 1

	Endproc


*----------------------------------------------------------------------------------
	Procedure SearchInOpenProjects(tcProject, ttTime, tcUni)
	
		Local lcFile As String
		Local lcProjectAlias As String
		Local lcProjectPath As String
		Local lnReturn As Number
		Local lnSelect As Number
		Local lnX As Number
		Local laProjectFiles[1], lcProject, lnI
	
		lnSelect	  = Select()
		This.tRunTime = Evl(m.ttTime, Datetime())
		This.cUni	  = Evl(m.tcUni, '_' + Sys(2007, Ttoc(This.tRunTime), 0, 1))
	
		Create Cursor ProjectFiles (FileName C(200), Type C(1))
	
		For lnI = 1 To _vfp.Projects.Count
			lcProject = _vfp.Projects[m.lnI].Name
	
			lcProjectPath  = Addbs(Justpath(Alltrim(m.lcProject)))
			lcProjectAlias = 'GF_ProjectSearch'
	
			This.oSearchOptions.cProject = m.lcProject
	
			Use (m.lcProject) Again Shared Alias (m.lcProjectAlias) In 0
	
			Insert Into ProjectFiles													;
				Select  Name,															;
						Type															;
					From (m.lcProjectAlias)												;
					Where Type $ 'EHKMPRVBdTxD' And										;
						Not Deleted()													;
						And Not (Upper(Justext(Name)) $ This.cGraphicsExtensions)		;
					Order By Type
	
			Use In Alias (m.lcProjectAlias)
	
		Endfor
	
		Select Distinct * From ProjectFiles Order By Type Into Array laProjectFiles
	
		If Type('laProjectFiles') = 'L'
			This.SearchFinished(m.lnSelect)
			Return 1
		Endif
	
		This.PrepareForSearch()
		This.StartTimer()
		This.StartProgressBar(Alen(m.laProjectFiles) / 2.0)
	
		*** Uncomment this code to track the execution time for the result
		*!*			LOCAL nSeconds
		*!*			nSeconds = SECONDS()
	
		For lnX = 1 To Alen(m.laProjectFiles) Step 2
	
			lcFile = m.laProjectFiles(m.lnX)
			lcFile = Fullpath(m.lcFile, m.lcProjectPath)
			lcFile = Strtran(m.lcFile, Chr(0), '') && Strip out junk char from the end
	
			If This.oSearchOptions.lLimitToProjectFolder
				If Not (Upper(m.lcProjectPath) $ Upper(Addbs(Justpath(m.lcFile))))
					Loop
				Endif
			Endif
	
			lnReturn = This.SearchInFile(m.lcFile)
	
			This.UpdateProgressBar(This.nFilesProcessed)
	
			If (m.lnReturn < 0) Or This.lEscPress Or This.nMatchLines >= This.oSearchOptions.nMaxResults
				Exit
			Endif
	
		Endfor
	
		*** Uncomment this code to show used execution time for the result
		*!*			MESSAGEBOX(ALLTRIM(STR((SECONDS()-nSeconds)*1000)) + " ms")
	
		This.SearchFinished(m.lnSelect)
	
		If m.lnReturn >= 0
			Return 1
		Else
			Return m.lnReturn
		Endif
	
	Endproc
		

*----------------------------------------------------------------------------------
	Procedure SearchInPath(tcPath, ttTime, tcUni)

		Local;
			lcDirectory   As String,;
			lcDirectory2  As String,;
			lcFile        As String,;
			lcFileFilter  As String,;
			lcFileName    As String,;
			lnFileCount   As Number,;
			lnReturn      As Number,;
			lnSelect      As Number,;
			lnTotalFileCount As Number

		Local Array;
			laTemp(1)

*:Global;
j

		lnSelect = Select()

		If Empty(m.tcPath)
			This.SetSearchError('Path parameter [' + m.tcPath + '] is empty in call to SearchInPath()')
			Return 0
		Endif

		This.tRunTime             = Evl(m.ttTime, Datetime())
		This.cUni                 = Evl(m.tcUni, "_" + Sys(2007, Ttoc(This.tRunTime), 0, 1))
		This.oSearchOptions.cPath = m.tcPath

		This.StoreInitialDefaultDir()

		If !This.ChangeCurrentDir(m.tcPath) && If there was a problem CD-ing into the starting path
			This.RestoreDefaultDir()
			Return - 1
		Endif

		This.PrepareForSearch()
		This.StartTimer()

		lnReturn = 1 && Assume success, testing below will set negative if there are errors

		This.ShowWaitMessage('Scanning directory...')

		If Vartype(This.oProgressBar) = 'O'
			This.oProgressBar.nMaxValue = 0
		Endif

		This.StartProgressBar(0)

		This.oDirectories = This.GetDirectories(m.tcPath, This.oSearchOptions.lIncludeSubdirectories)

		If This.lEscPress
			This.SearchFinished(m.lnSelect)
			Return 0
		Endif

		Chdir (m.tcPath) && Must go back, since above call to BuildDirList prolly changed our directory!

		This.StartProgressBar(This.oProgressBar.nMaxValue)
		lnTotalFileCount = 0

		For Each m.lcDirectory In This.oDirectories
			lcDirectory2 = Lower(Justfname(Justpath(m.lcDirectory)))
			If This.FilesToSkip(Upper(m.lcDirectory + '\-'))
				Loop
			Endif

			lcFileFilter = Addbs(m.lcDirectory) + '*.*'

			If Adir(laTemp, m.lcFileFilter) = 0 && 0 means no files in the Dir
				Loop
			Endif

			Asort(m.laTemp)

			lnFileCount = Alen(m.laTemp) / 5 && The number of files that match the filter criteria for this pass

			For j = 1 To m.lnFileCount
				lcFileName = laTemp(j, 1) && Just the name and ext, no path info
				lcFile     = Addbs(m.lcDirectory) + m.lcFileName && path + filename
				lnReturn   = This.SearchInFile(m.lcFile)

				lnTotalFileCount = m.lnTotalFileCount + 1
				This.UpdateProgressBar(m.lnTotalFileCount)

				If m.lnReturn < 0 Or This.lEscPress Or This.nMatchLines >= This.oSearchOptions.nMaxResults
					Exit
				Endif
			Endfor

			If m.lnReturn < 0 Or This.lEscPress Or This.nMatchLines >= This.oSearchOptions.nMaxResults
				Exit
			Endif
		Endfor

		This.SearchFinished(m.lnSelect)

		This.RestoreDefaultDir()

		If m.lnReturn >= 0
			Return 1
		Else
			Return m.lnReturn
		Endif

	Endproc


*----------------------------------------------------------------------------------
	Procedure SearchInProject(tcProject, ttTime, tcUni)

		Local;
			lcFile      As String,;
			lcProjectAlias As String,;
			lcProjectPath As String,;
			lnReturn    As Number,;
			lnSelect    As Number,;
			lnX         As Number

		Local Array;
			laProjectFiles(1)

		lnSelect = Select()

		lcProjectPath  = Addbs(Justpath(Alltrim(m.tcProject)))
		lcProjectAlias = 'GF_ProjectSearch'

		This.tRunTime                = Evl(m.ttTime, Datetime())
		This.cUni                    = Evl(m.tcUni, "_" + Sys(2007, Ttoc(This.tRunTime), 0, 1))
		This.oSearchOptions.cProject = m.tcProject

		If Empty(m.tcProject)
			This.SetSearchError('Project parameter [' + m.tcProject + '] is empty in call to SearchInProject().')
			Return 0
		Endif

		If !File(m.tcProject)
			This.SetSearchError('Project file [' + m.tcProject + '] not found in call to SearchInProject().')
			Return 0
		Endif

		Try && Attempt to open Project.PJX in a cursor...
				Select 0
				Use (m.tcProject) Again Shared Alias &lcProjectAlias
				lnReturn = 1
			Catch
				lnReturn = -2
		Endtry

		If m.lnReturn = -2
			This.SetSearchError('Cannot open project file[' + m.tcProject + ']')
			This.SearchFinished(m.lnSelect)
			Return m.lnReturn
		Endif

		Select Name,;
			Type ;
			From (m.lcProjectAlias) ;
			Where Type $ 'EHKMPRVBdTxD' And ;
			Not Deleted() ;
			And !(Upper(Justext(Name)) $ This.cGraphicsExtensions) ;
			Order By Type ;
			Into Array laProjectFiles

		If Type('laProjectFiles') = 'L'
			This.SearchFinished(m.lnSelect)
			Return 1
		Endif

		Use In Alias (m.lcProjectAlias)

		This.PrepareForSearch()
		This.StartTimer()
		This.StartProgressBar(Alen(m.laProjectFiles) / 2.0)

*** Uncomment this code to track the execution time for the result
*!*			LOCAL nSeconds
*!*			nSeconds = SECONDS()

		For lnX = 1 To Alen(m.laProjectFiles) Step 2

			lcFile = laProjectFiles(m.lnX)
			lcFile = Fullpath(m.lcFile, m.lcProjectPath)
			lcFile = Strtran(m.lcFile, Chr(0), '') && Strip out junk char from the end

			If This.oSearchOptions.lLimitToProjectFolder
				If !(Upper(m.lcProjectPath) $ Upper(Addbs(Justpath(m.lcFile))))
					Loop
				Endif
			Endif

			lnReturn = This.SearchInFile(m.lcFile)

			This.UpdateProgressBar(This.nFilesProcessed)

			If (m.lnReturn < 0) Or This.lEscPress Or This.nMatchLines >= This.oSearchOptions.nMaxResults
				Exit
			Endif

		Endfor

*** Uncomment this code to show used execution time for the result
*!*			MESSAGEBOX(ALLTRIM(STR((SECONDS()-nSeconds)*1000)) + " ms")

		This.SearchFinished(m.lnSelect)

		If m.lnReturn >= 0
			Return 1
		Else
			Return m.lnReturn
		Endif

	Endproc


*----------------------------------------------------------------------------------
	Procedure SearchInTable(tcFile)

		Local;
			lcClass              As String,;
			lcCode               As String,;
			lcDataType           As String,;
			lcDeleted            As String,;
			lcExt                As String,;
			lcField              As String,;
			lcFieldSource        As String,;
			lcFormClass          As String,;
			lcFormClassloc       As String,;
			lcFormName           As String,;
			lcName               As String,;
			lcObjectType         As String,;
			lcParent             As String,;
			lcParentName         As String,;
			lcProject            As String,;
			lcSearchExpression   As String,;
			ldFromDate           As Date,;
			ldMaxTimeStamp       As Date,;
			ldToDate             As Date,;
			llContinueError      As Boolean,;
			llHasMethods         As Boolean,;
			llIgnorePropertiesField As Boolean,;
			llLocateError        As Boolean,;
			llProcessThisMatch   As Boolean,;
			llScxVcx             As Boolean,;
			lnEndColumn          As Number,;
			lnMatchCount         As Number,;
			lnParentId           As Number,;
			lnSelect             As Number,;
			lnStart              As Number,;
			lnStartColumn        As Number,;
			lnTotalMatches       As Number,;
			loException          As Object,;
			loFileResultObject   As 'GF_FileResult',;
			loSearchResultObject As 'GF_SearchResult'

		Local Array;
			laMaxTimeStamp(1),;
			laParent(1)

*:Global;
ii

		lnSelect = Select()

		lnMatchCount   = 0
		lnTotalMatches = 0
		lcExt          = Upper(Justext(m.tcFile))
		lcProject      = This.oSearchOptions.cProject

		lcSearchExpression = Upper(Alltrim(This.oSearchOptions.cSearchExpression))

		Try
				Use (m.tcFile) Again Shared Alias 'GF_TableSearch' In Select('GF_TableSearch')
			Catch To m.loException

		Endtry

		If Not Used('GF_TableSearch')
			This.SetSearchError('Cannot open file: ' + Alltrim(m.tcFile) + CR + Space(5) + m.loException.Message, 16, 'File Error')
			Return 0
		Else
			Select('GF_TableSearch')

			If m.lcExt = 'SCX'
				If Empty(Field('BaseClass'))
					lcFormName     = ''
					lcFormClass    = ''
					lcFormClassloc = ''
				Else
					Locate For BaseClass = 'form'
					lcFormName     = ObjName
					lcFormClass    = Class
					lcFormClassloc = ClassLoc
				Endif
			Endif

		Endif

		If This.oSearchOptions.lTimeStamp And Type('timestamp') = 'U'
			Use In 'GF_TableSearch'
			Return 0
		Endif

		This.ShowWaitMessage('Searching File: ' + m.tcFile)

		lnEndColumn             = 255
		llIgnorePropertiesField = .F.

		Do Case
			Case 'VCX' $ m.lcExt
				lnStartColumn           = 4
				llIgnorePropertiesField = This.oSearchOptions.lIgnorePropertiesField
			Case 'SCX' $ m.lcExt
				lnStartColumn           = 4
				llIgnorePropertiesField = This.oSearchOptions.lIgnorePropertiesField
* lnEndColumn = 12
			Case 'FRX' = m.lcExt
				lnStartColumn = 3 && Newer reports could start at col 6, but older reports can have text data starting in column 3
* lnEndColumn = 21
				If Len(Field('timestamp', 'GF_TableSearch')) > 0 && Some really old reports may not have this field.
					Select Max(Timestamp);
						From 'GF_TableSearch';
						Into Array laMaxTimeStamp
				Else
					laMaxTimeStamp = {}
				Endif
				ldMaxTimeStamp = Ctot(This.TimeStampToDate(m.laMaxTimeStamp))
			Case 'DBC' = m.lcExt
				lnStartColumn = 3
* lnEndColumn = 6
			Case 	'MNX' = m.lcExt
				lnStartColumn = 1
			Otherwise
				lnStartColumn = 1
		Endcase

		lcDeleted = Set('Deleted')
		Set Deleted On

*-- Scan across all table columns looking for matches on each row
*--	See: http://fox.wikis.com/wc.dll?Wiki~VFPVcxStructure for details about scx/vcx columns
*-- See: http://mattslay.com/foxpro-report-frx-table-structure for details about FRX structure

		For ii = m.lnStartColumn To m.lnEndColumn Step 1
			Goto Top
			lcField       = Upper(Field(m.ii))
			llLocateError = .F.

			If Empty(m.lcField)
				Exit
			Endif

			If  Not Type(m.lcField) $ 'CM' Or					; && If not a character or Memo field
				('TAG' $ m.lcField And m.lcExt # 'FRX') Or		;
					Inlist(m.lcField, 'OBJCODE', 'OBJECT', 'SYMBOLS')
				Loop
			Endif

			If m.llIgnorePropertiesField And m.lcField == 'RESERVED3'
				Loop
			Endif

			If m.lcExt = 'DBC'
				lcObjectType = Alltrim(Upper(ObjectType))
				If Type('objectname') = 'U'
					Loop
				Endif
			Endif

*-- This is an important speed part of GoFish... If the user is not using a regular expression, then we
*-- can use the Locate command to make a quick look for a match anywhere in this column. We will handle the
*-- whole word part later on in the code, but a quick partial match hit helps us skips rows that have not
*-- match at all.
*-- If we find a match, we process it futher and then call Continue to look for the next partial and repeat.
*-- This logic is not used if we are doing a Reg Ex search.
			If Not This.oSearchOptions.lRegularExpression
				If This.oSearchOptions.lTimeStamp
					ldFromDate = Evl(This.oSearchOptions.dTimeStampFrom, {^1900-01-01})
					ldToDate   = Evl(This.oSearchOptions.dTimeStampTo, {^9999-01-01})
					ldToDate   = m.ldToDate + 1 &&86400 && Must bump into to next day, since TimeStamp from table has time on it

					Locate For Between(Ctot(This.TimeStampToDate(Timestamp)), m.ldFromDate, m.ldToDate) Nooptimize

					If Not Found() && If doing a TimeStmp search and we did not find a match, we can skip out of this file
						Exit
					Endif
				Else
					Try
							If This.oSearchOptions.nSearchMode = GF_SEARCH_MODE_LIKE
								Locate For Likec('*' + Upper(m.lcSearchExpression) + '*', Upper(Evaluate(m.lcField))) Nooptimize
							Else
								Locate For Upper(m.lcSearchExpression) $ Upper(Evaluate(m.lcField)) Nooptimize
							Endif
						Catch
							This.SetSearchError('Error scanning through table [' + m.tcFile + ']. File may be corrupt.')
							llLocateError = .T.
						Finally
					Endtry
				Endif

				If Not Found() Or m.llLocateError
					Loop && Loop to next column
				Endif

			Endif

			Do While Not Eof()

				lnMatchCount       = 0
				loFileResultObject = Createobject('GF_FileResult')	&& This custom class has all the properties that must be populated if you want to
				llProcessThisMatch = .T.														&& have a cursor created
				llScxVcx           = Inlist(m.lcExt, 'VCX', 'SCX')
				lcCode             = Evaluate(m.lcField)

				With m.loFileResultObject
					.Process   = .F.
					.FileName  = Justfname(m.tcFile)
					.FilePath  = m.tcFile
					.MatchType = Proper(m.lcField)
					.FileType  = Upper(m.lcExt)
					.Column    = m.lcField
					.IsText    = .F.
					.Recno     = Recno()
					.Timestamp = Iif(Type('timestamp') # 'U', Ctot(This.TimeStampToDate(Timestamp)), {// :: AM})

					lcClass          = Iif(Type('class') # 'U', Class, '')
					.ContainingClass = m.lcClass
					._ParentClass    = m.lcClass
					._BaseClass      = Iif(Type('baseclass') # 'U', BaseClass, '')
					.ClassLoc        = Iif(Type('classloc') # 'U', ClassLoc, '')

					lcParent = Iif(Type('parent') # 'U', Parent, '')
					lcName   = Iif(Type('objname') # 'U', ObjName, '')
					._Name   = Alltrim(m.lcParent + '.' + m.lcName, '.')

					Do Case
						Case m.lcExt = 'SCX'

							._Class = ''

						Case m.lcExt = 'VCX'

							If Not Empty(m.lcParent)
								._Class = Getwordnum(m.lcParent, 1, '.')
							Else
								._Class = Alltrim(ObjName)
								._Name  = ''
							Endif

						Case m.lcExt = 'FRX'
							._Name  = Name
							._Class = This.GetFrxObjectType(ObjType, objCode)
							If Empty(.Timestamp)
								.Timestamp = m.ldMaxTimeStamp
							Endif

						Case m.lcExt = 'DBF'
							.MatchType = '<Field>'
							._Name     = Proper(m.lcField)

						*** JRN 2024-02-05 : For some MNX matches, show more info from the same record
						Case m.lcExt = 'MNX' and InList(Upper(m.lcField), 'PROMPT', 'COMMAND', 'PROCEDURE', 'SKIPFOR')
							lcCode = ;
								Iif(Empty(Prompt), 	  '' , 'Prompt    = "' + This.GetFullMenuPrompt() + '"' + CRLF) + ;
								Iif(Empty(Command),   '' , 'Command   = ' + Command + CRLF) + ;
								Iif(Empty(Procedure), '' , 'Procedure = ' + Iif(CR $ Trim(Procedure, 1, CR, LF, Tab,' '), CRLF, '') + Procedure + CRLF) + ;
								Iif(Empty(SkipFor),   '' , 'SkipFor   = ' + SkipFor + CRLF) 

						Case m.lcExt = 'DBC'
							._Name  = Alltrim(ObjectName)
							._Class = Alltrim(ObjectType)

							Do Case

								Case ._Class = 'Database' And m.lcField = 'OBJECTNAME'
*lcCode = '' && Will cause this match to be skipped. Don't want to record these matches.

								Case ._Class = 'Table'
*lcCode = ._Class + '.dbf' && The name of the Table attached to the DBC
									lcCode = This.CleanUpBinaryString(m.lcCode)  && The SQL statement that makes up the View

								Case ._Class = 'View'
									lnStart = Atc('Select', m.lcCode)
									lcCode  = Substr(m.lcCode, m.lnStart)
									lcCode  = This.CleanUpBinaryString(m.lcCode, .T.)  && The SQL statement that makes up the View

								Case ._Class = 'Field' && Fields can be part of Tables or Views
*-- Get some info about the parent of this field
									lnParentId = parentId
									Select ObjectType,;
										ObjectName;
										From (m.tcFile);
										Where objectid = m.lnParentId;
										Into Array laParent
									lcParentName = Alltrim(laParent[2])

*-- Parse the field into a field name and field source
									lnStart = Atc('#', m.lcCode)
									lcCode  = Substr(m.lcCode, m.lnStart + 1)
									lcCode  = This.CleanUpBinaryString(m.lcCode)

									lcFieldSource = Alltrim(Getwordnum(m.lcCode, 1))
									lcDataType    = Substr(Alltrim(Getwordnum(m.lcCode, 2)), 2)

									If m.lcFieldSource = '0'
										lcFieldSource = '[Table alias in query]'
										lcDataType    = ''
									Endif

									If Not Empty(m.lcDataType)
										lcCode     = m.lcParentName + ' references ' + m.lcFieldSource + ' (data type: ' + m.lcDataType + ')'
										.MatchType = 'Field Source'
									Else
										lcCode     = m.lcParentName + '.' + m.lcFieldSource
										.MatchType = Alltrim(laParent[1]) + ' Field'
									Endif

									._Class = .MatchType

							Endcase
					Endcase

*-- Here is where we can skip the processing of certain records that we want to ignore, even though we found a match in them.
					If (m.lcExt = 'VCX' And Empty(m.lcClass)) Or ;					 	&& This is the ending row of a Class def in a vcx. Need to skip over it.
						(m.lcExt = 'FRX' And m.lcField = 'TAG2' And Recno() = 1) Or ;	&& Tag2 on first record in a FRX is binary and I want to skip it.
						(m.lcExt = 'PJX' And m.lcField = 'KEY') 					  		&& Added this filter on 2021-03-24, as requested by Jim Nelson.

						llProcessThisMatch = .F.
					Endif

				Endwith

				If This.oSearchOptions.lTimeStamp And Not Between(Ctot(This.TimeStampToDate(Timestamp)), m.ldFromDate, m.ldToDate)
					llProcessThisMatch = .F.
				Endif

				If m.llProcessThisMatch
					If Not Empty(This.oSearchOptions.cSearchExpression)
*lcCode = Evaluate(lcField)
						llHasMethods = Upper(m.lcField) = 'METHODS' Or		;
							M.lcExt = 'FRX' And Upper(m.lcField) = 'TAG' And Upper(Name) = 'DATAENVIRONMENT'
						lnMatchCount = This.SearchInCode(m.lcCode, m.loFileResultObject, m.llHasMethods)
					Else
* Can't search since there is no cSearchExpression, so we just log the file as a result.
* This handles TimeStamp searches, where the cSearchExpression is empty
						loSearchResultObject           = Createobject('GF_SearchResult')
						loSearchResultObject.Code      = Iif(Type('properties') # 'U', Properties, '')
						loSearchResultObject.Code      = m.loSearchResultObject.Code + CR + Iif(Type('methods') # 'U', Methods, '')
						loSearchResultObject.UserField = m.loFileResultObject

						If m.lcExt = 'FRX'
							loSearchResultObject.MatchLine        = Expr
							loSearchResultObject.TrimmedMatchLine = Expr
						Endif

						This.ProcessSearchResult(m.loSearchResultObject)

						ii           = 1000 && To end the outer for loop when the Do loop ends
						lnMatchCount = m.lnMatchCount + 1
					Endif
				Endif

				If m.lnMatchCount < 0 && There was an error in above call, need to exit
					Exit
				Else
					lnTotalMatches = m.lnTotalMatches + m.lnMatchCount
				Endif


				If Not This.oSearchOptions.lRegularExpression
					Try
							Continue
						Catch
							This.SetSearchError('Error scanning through table [' + m.tcFile + ']. File may be corrupt.')
							llContinueError = .T.
						Finally
					Endtry

					If m.llContinueError
						Exit
					Endif


				Else
					Skip 1
				Endif

			Enddo

			If m.lnMatchCount < 0 && There was an error in above call
				Exit
			Endif

		Endfor

		Set Deleted &lcDeleted

		Use In 'GF_TableSearch'

		Select (m.lnSelect)

		If m.lnMatchCount < 0
			Return m.lnMatchCount
		Else
			Return m.lnTotalMatches
		Endif

	Endproc


*----------------------------------------------------------------------------------
	Procedure SearchInTextFile(tcFile)

		Local;
			lcCode            As String,;
			ldFileDate        As Date,;
			ldFromDate        As Date,;
			ldToDate          As Date,;
			llHasMethods      As Boolean,;
			llReadFile        As Boolean,;
			lnMatchCount      As Number,;
			lnSelect          As Number,;
			loFileResultObject As 'GF_FileResult',;
			loSearchResultObject As 'GF_SearchResult'

		lnSelect = Select()

		If !File(m.tcFile)
			This.lFileNotFound = .T.
			This.SetSearchError('File not found: ' + m.tcFile)
			Return 0
		Endif

		ldFileDate = This.GetFileDateTime(m.tcFile)

		ldFromDate = Evl(This.oSearchOptions.dTimeStampFrom, {^1900-01-01})
		ldToDate   = Evl(This.oSearchOptions.dTimeStampTo, {^9999-01-01})
		ldToDate   = m.ldToDate + 1&&86400 && Must bump into to next day, since TimeStamp from table has time on it

		If This.oSearchOptions.lTimeStamp And !Between(m.ldFileDate, m.ldFromDate, m.ldToDate)
			Return 0
		Endif

		loFileResultObject = Createobject('GF_FileResult')	&& This custom class has all the properties that must be populated if you want to
&& have a cursor created
		With m.loFileResultObject
			.FileName  = Justfname(m.tcFile)
			.FilePath  = m.tcFile
			.MatchType = Proper(Justext(m.tcFile))
			.FileType  = Upper(Justext(m.tcFile))
			.IsText    = .T.
			.Timestamp = m.ldFileDate
		Endwith

		If !Empty(This.oSearchOptions.cSearchExpression)
			Try
					lcCode     = Filetostr(m.tcFile) && File could be in use by some other app and can't be read in
					llReadFile = .T.
				Catch
					This.SetSearchError('Could not open file [' + m.tcFile + '] for reading.')
					llReadFile = .F.
			Endtry
			If !m.llReadFile
				Select (m.lnSelect)
				Return 0
			Endif

			llHasMethods = Inlist(Upper(m.loFileResultObject.MatchType) + ' ', 'PRG ', 'MPR ', 'H ')
			lnMatchCount = This.SearchInCode(m.lcCode, m.loFileResultObject, m.llHasMethods)
		Else
* Can't search since there is no cSearchExpression, so we just log the file as a result.
* This handles TimeStamp searches, where the cSearchExpression is empty
			loSearchResultObject           = Createobject('GF_SearchResult')
			loSearchResultObject.UserField = m.loFileResultObject

			This.ProcessSearchResult(m.loSearchResultObject)
			lnMatchCount = 1
		Endif

		Select (m.lnSelect)

		Return m.lnMatchCount

	Endproc

*----------------------------------------------------------------------------------
*-- Read file patterns to include in the search
* SF 20230619
	Procedure SetIncludePattern()

		This.oSearchOptions.cFileTemplate  = Alltrim(This.oSearchOptions.cFileTemplate)
		This.oSearchOptions.cOtherIncludes = Alltrim(This.oSearchOptions.cOtherIncludes)

		If !Empty(This.oSearchOptions.cFileTemplate) Then
			If Isnull(This.oSearchOptions.oRegExpFileTemplate) Then
				loRegExp = GF_GetRegExp()
				loRegExp.IgnoreCase       = .T.
				loRegExp.MultiLine        = .T.
				loRegExp.ReturnFoxObjects = .T.
*				loRegExp.AutoExpandGroups  = .T.
				loRegExp.Singleline       = .T.
				This.oSearchOptions.oRegExpFileTemplate = loRegExp

			Else &&ISNULL(This.oSearchOptions.oRegExpFileTemplate)
				loRegExp = This.oSearchOptions.oRegExpFileTemplate

			Endif &&ISNULL(This.oSearchOptions.oRegExpFileTemplate)

			lnPatterns = Alines(laPattern, This.oSearchOptions.cFileTemplate, 1, ",", ";")
			If m.lnPatterns = 1 Then
*!*						If Justext(laPattern(1))=="" Then
*!*							If Juststem(laPattern(1))=="" Then
*!*								lcPattern = ""
*!*							Endif &&JUSTSTEM(laPattern(1))==""
*!*	*filename without extension
*SET STEP ON 
*!*							lcPattern = This.EscapePattern(m.loRegExp, Justext(laPattern(1))) + "(?<!\..+)$"
*!*						Else  &&JUSTEXT(laPattern(1))==""
						lcPattern = This.EscapePattern(m.loRegExp, laPattern(1))
*!*						Endif &&JUSTEXT(laPattern(1))==""

*				lcPattern = This.EscapePattern(m.loRegExp, laPattern(1))

			Else  &&m.lnPatterns = 1
				lcPattern = ""

				For lnPattern = 1 To m.lnPatterns
					If Justext(laPattern(m.lnPattern))=="" Then
						If Juststem(laPattern(m.lnPattern))=="" Then
							Loop
						Endif &&JUSTSTEM(laPattern(m.lnPattern))==""
*filename without extension
						lcPattern = m.lcPattern + "|(" + This.EscapePattern(m.loRegExp, Justext(laPattern(m.lnPattern))) + "(?<!\..+)$)"
					Else  &&JUSTEXT(laPattern(m.lnPattern))==""
						lcPattern = m.lcPattern + "|(" + This.EscapePattern(m.loRegExp, laPattern(m.lnPattern)) + ")"
					Endif &&JUSTEXT(laPattern(m.lnPattern))==""

*					IIF("." $ laPattern(m.lnPattern),;
loRegExp.Escape_Like(JUSTSTEM(laPattern(m.lnPattern))) + "\." + loRegExp.Escape_Like(justext(laPattern(m.lnPattern))),;
".*\." + loRegExp.Escape_Like(laPattern(m.lnPattern))) +;
")"
				Endfor &&lnPattern
				lcPattern = Substr(m.lcPattern, 2)
			Endif &&m.lnPatterns
			loRegExp.Pattern = m.lcPattern

		Endif &&!Empty(This.oSearchOptions.cFileTemplate)
	Endproc

*----------------------------------------------------------------------------------
*-- Escape a file pattern
	Procedure EscapePattern()
		Lparameters;
			toRegExp,;
			tcPattern

		Return Icase(;
			Left(tcPattern, 1)  = "." , ".*\." + toRegExp.Escape_Like(Justext(m.tcPattern)),;
			Right(tcPattern, 1) = "." , toRegExp.Escape_Like(Juststem(m.tcPattern)) + "(?<!\..+)$",;
			"." $ m.tcPattern         , loRegExp.Escape_Like(Juststem(m.tcPattern)) + "\." + loRegExp.Escape_Like(Justext(m.tcPattern)),;
			EMPTY(tcPattern)          , "",;
			toRegExp.Escape_Like(m.tcPattern) + "\..*")

	Endproc

*----------------------------------------------------------------------------------
*-- Read a user file set the the cFilesToSkip property
	Procedure SetFilesToSkip

		Local;
			lcExclusionFile As String,;
			lcFilesToSkip As String,;
			lcLeft       As String,;
			lcLine       As String,;
			lcRight      As String,;
			lnI          As Number

		Local Array;
			laLines(1)

		lcFilesToSkip   = ''
		lcExclusionFile = This.cFilesToSkipFile

		If File(m.lcExclusionFile) And This.oSearchOptions.lSkipFiles
			lcFilesToSkip = Filetostr(m.lcExclusionFile)
		Endif

		This.cFilesToSkip         = CR
		This.nWildCardFilesToSkip = 0

		For lnI = 1 To Alines(laLines, m.lcFilesToSkip + Chr(13) + '_command.prg', 5)
			lcLine  = Upper(laLines[m.lni])
			lcLeft  = Left(m.lcLine, 1)
			lcRight = Right(m.lcLine, 1)

			Do Case
				Case Left(m.lcLine, 2) = '**'

				Case m.lcLeft = '\' And m.lcRight = '\'
					This.nWildCardFilesToSkip = This.nWildCardFilesToSkip + 1
					Dimension This.aWildcardFiles[This.nWildCardFilesToSkip]
					This.aWildcardFiles[This.nWildCardFilesToSkip] = '*' + m.lcLine + '*'

				Case '\' $ m.lcLine
					This.nWildCardFilesToSkip = This.nWildCardFilesToSkip + 1
					Dimension This.aWildcardFiles[This.nWildCardFilesToSkip]
					This.aWildcardFiles[This.nWildCardFilesToSkip] = Icase(m.lcLeft = '*', '', m.lcLeft = '\', '*', '*\')  + m.lcLine

				Case '*' $ m.lcLine Or '?' $ m.lcLine
					This.nWildCardFilesToSkip = This.nWildCardFilesToSkip + 1
					Dimension This.aWildcardFiles[This.nWildCardFilesToSkip]
					This.aWildcardFiles[This.nWildCardFilesToSkip] = '*\' + m.lcLine

				Otherwise
					This.cFilesToSkip = This.cFilesToSkip + m.lcLine + CR
			Endcase
		Endfor

	Endproc


*----------------------------------------------------------------------------------
	Procedure SetProject(tcProject)

		Local;
			lcProject As String,;
			llReturn As Boolean

		lcProject = Lower(Evl(m.tcProject, ''))

		If Empty(m.lcProject)
			Return .T.
		Endif

		If !('.pjx' $ m.lcProject)
			lcProject = m.lcProject + '.pjx'
		Endif

		If File(m.lcProject)
			This.AddProject(m.lcProject)
			This.oSearchOptions.cProject = m.lcProject
			llReturn                     = .T.
		Else
			This.oSearchOptions.cProject = ''
			This.SetSearchError('Project not found [' + m.lcProject + '] in call to SetProject() method.')
			llReturn = .F.
		Endif

		Return m.llReturn

	Endproc


*----------------------------------------------------------------------------------
	Procedure SetReplaceError(tcErrorMessage, tcFile, tnResultId, tnDialogBoxType, tcTitle)

		Local;
			lcErrorMessage As String,;
			lcFile      As String,;
			lcResultId  As String,;
			lnResultId  As Number

		lcFile     = Alltrim(Evl(m.tcFile, 'None'))
		lnResultId = Evl(m.tnResultId, 0)

		lcResultId = Iif(m.lnResultId = 0, 'None', Alltrim(Str(m.lnResultId)))

		lcErrorMessage = m.tcErrorMessage + Space(4) + ;
			'[File: ' + m.lcFile + ']' + Space(4) + ;
			'[Result Id: ' + m.lcResultId + ']'

		This.ShowError(m.lcErrorMessage, m.tnDialogBoxType, m.tcTitle)

		This.oReplaceErrors.Add(m.lcErrorMessage)

	Endproc


*----------------------------------------------------------------------------------
	Procedure SetSearchError(tcErrorMessage, tnDialogBoxType, tcTitle)

		This.oSearchErrors.Add(m.tcErrorMessage)

	Endproc


*----------------------------------------------------------------------------------
	Procedure ShowError(tcErrorMessage, tnDialogBoxType, tcTitle)

		Local;
			lcTitle      As String,;
			lnDialogBoxType As Number

		If Empty(m.tcErrorMessage) Or !This.oSearchOptions.lShowErrorMessages
			Return
		Endif

		lnDialogBoxType = Evl(m.tnDialogBoxType, 0)
		lcTitle         = Evl(m.tcTitle, 'GoFishSearchEngine Error:')

	Endproc


*----------------------------------------------------------------------------------
	Procedure ShowWaitMessage(tcMessage)

		If This.oSearchOptions.lShowWaitMessages
			Wait Window At 5, Wcols() / 2 Nowait m.tcMessage
		Endif

	Endproc


*----------------------------------------------------------------------------------
	Procedure StartProgressBar(tnMaxValue)

		If Vartype(This.oProgressBar) = 'O'
			This.oProgressBar.Start(m.tnMaxValue)
		Endif

	Endproc


*----------------------------------------------------------------------------------
	Procedure StartTimer

		This.nSearchTime = Seconds()

	Endproc


*----------------------------------------------------------------------------------
	Procedure StopProgressBar

		If Vartype(This.oProgressBar) = 'O'
			This.oProgressBar.Stop()
		Endif

	Endproc


*----------------------------------------------------------------------------------
	Procedure StoreInitialDefaultDir

		This.cInitialDefaultDir = Sys(5) + Sys(2003)

	Endproc


*----------------------------------------------------------------------------------
	Procedure ThorMoveWindow

		If Type('_Screen.cThorDispatcher') = 'C'
			Execscript (_Screen.cThorDispatcher, 'PEMEditor_StartIDETools')
			_oPEMEditor.Outils.oIDEx.MoveWindow()
		Endif

	Endproc


*----------------------------------------------------------------------------------
*  METHOD: TimeStamp2Date()
*
*  AUTHOR: Richard A. Schummer            September 1994
*
*  COPYRIGHT (c) 1994-2001    Richard A. Schummer
*     42759 Flis Dr
*     Sterling Heights, MI  48314-2850
*     RSchummer@CompuServe.com
*
*  METHOD DESCRIPTION:
*     This procedure handles the conversion of the FoxPro TIMESTAMP field to
*     a readable (and understandable) date and time.  The procedure will return
*     the date/time in three formats based on the cStyle parameter.  Timestamp
*     field is a 32-bit (numeric compressed) system that the FoxPro development
*     team created to save on file space in the projects, screens, reports, and
*     label databases.  This field is used to determine if objects need to be
*     recompiled (project manager), or syncronized across platforms (screens,
*     reports, and labels).
*
*  CALLING SYNTAX:
*     <variable> = ctrMetaDecode.TimeStamp2Date(<nTimeStamp>,<cStyle>)
*
*     Sample:
*     ltDateTime = ctrMetaDecode.TimeStamp2Date(TimeStamp,"DATETIME")
*
*  INPUT PARAMETERS:
*     nTimeStamp = Required field, must be numeric, no check to verify the
*                  data passed is valid FoxPro Timestamp, just be sure it is
*     cStyle     = Not required (defaults to "DATETIME"), must be character,
*                  and must be one of the following:
*                   "DATETIME" will return the date/time in MM/DD/YY HH:MM:SS
*                   "DATE"     will return the date in MM/DD/YY format
*                   "TIME"     will return the time in HH:MM:SS format
*
*  OUTPUT PARAMETERS:
*     lcRetval    = The date/time (in requested format) is returned in
*                   character type.  Must be converted and parsed to be
*                   used as date type.
*

	Procedure TimeStampToDate(tnTimeStamp, tcStyle)

*=============================================================
* Tried to use this FFC class, but it sometimes gave an error:
* This.oFrxCursor.GetTimeStampString(timestamp)
*=============================================================

		Local;
			lcRetVal As String,;
			lnDay    As Number,;
			lnHour   As Number,;
			lnMinute As Number,;
			lnMonth  As Number,;
			lnSecond As Number,;
			lnYear   As Number,;
			loException As Object

		If Type('tnTimeStamp') != "N"          &&  Timestamp must be numeric
* Wait Window "Time stamp passed is not numeric" NoWait
			Return ""
		Endif

		If m.tnTimeStamp = 0                     &&  Timestamp is zero until built in project
			Return "Not built into App"
		Endif

		If Type('tcStyle') != "C"              &&  Default return style to both date and time
			tcStyle = "DATETIME"
		Endif

		If !Inlist(Upper(m.tcStyle), "DATE", "TIME", "DATETIME")
			Wait Window "Style parameter must be DATE, TIME, or DATETIME"
			Return ""
		Endif

		lnYear  = ((m.tnTimeStamp / (2 ** 25) + 1980))
		lnMonth = ((m.lnYear - Int(m.lnYear)    ) * (2 ** 25)) / (2 ** 21)
		lnDay   = ((m.lnMonth - Int(m.lnMonth)  ) * (2 ** 21)) / (2 ** 16)

		lnHour   = ((m.lnDay - Int(m.lnDay)      ) * (2 ** 16)) / (2 ** 11)
		lnMinute = ((m.lnHour - Int(m.lnHour)    ) * (2 ** 11)) / (2 ** 05)
		lnSecond = ((m.lnMinute - Int(m.lnMinute)) * (2 ** 05)) * 2       &&  Multiply by two to correct
&&  truncation problem built in
&&  to the creation algorithm
&&  (Source: Microsoft Tech Support)

		lcRetVal = ""

		If "DATE" $ Upper(m.tcStyle)
*< 4-Feb-2001 Fixed to display date in machine designated format (Regional Settings)
*< lcRetVal = lcRetVal + RIGHT("0"+ALLTRIM(STR(INT(lnMonth))),2) + "/" + ;
*<                       RIGHT("0"+ALLTRIM(STR(INT(lnDay))),2)   + "/" + ;
*<                       RIGHT("0"+ALLTRIM(STR(INT(lnYear))), IIF(SET("CENTURY") = "ON", 4, 2))

*< RAS 23-Nov-2004, change to work around behavior change in VFP 9.
*< lcRetVal = lcRetVal + DTOC(DATE(lnYear, lnMonth, lnDay))
			Try
					lcRetVal = m.lcRetVal + Dtoc(Date(Int(m.lnYear), Int(m.lnMonth), Int(m.lnDay)))
				Catch To m.loException
					lcRetVal = m.lcRetVal + Dtoc(Date(1901, 1, 1))
			Endtry
		Endif

		If "TIME" $ Upper(m.tcStyle)
			lcRetVal = m.lcRetVal + Iif("DATE" $ Upper(m.tcStyle), " ", "")
			lcRetVal = m.lcRetVal + Right("0" + Alltrim(Str(Int(m.lnHour))), 2)   + ":" + ;
				Right("0" + Alltrim(Str(Int(m.lnMinute))), 2) + ":" + ;
				Right("0" + Alltrim(Str(Int(m.lnSecond))), 2)
		Endif

		Return m.lcRetVal

	Endproc


*----------------------------------------------------------------------------------
	Procedure TrimWhiteSpace(tcString)

		Local;
			lcTrimmedString As String

		lcTrimmedString = Alltrim(m.tcString, 1, Chr(32), Chr(9), Chr(10), Chr(13), Chr(0))
		lcTrimmedString = Strtran(m.lcTrimmedString, Chr(9), Chr(32))

		Return m.lcTrimmedString

	Endproc


*----------------------------------------------------------------------------------
	Procedure UpdateCursorAfterReplace(tcCursor, toResult)

		Local;
			lcColumn    As String,;
			lcFileToModify As String,;
			lnChangeLength As Number,;
			lnCurrentRecno As Number,;
			lnMatchStart As Number,;
			lnProcStart As Number,;
			lnResultRecno As Number,;
			lnSelect    As Number

		If This.oSearchOptions.lPreviewReplace
			Return
		Endif

		lnChangeLength = m.toResult.nChangeLength

		lnSelect = Select()
		Select (m.tcCursor)
		lnCurrentRecno = Recno()

*-- Create local vars of certain fields from the current row that we need to use below
		lcFileToModify = Alltrim(FilePath)
		lnResultRecno  = Recno
		lnProcStart    = ProcStart
		lnMatchStart   = MatchStart
		lcColumn       = Column

*-- Cannot process same source code line more than once, so mark this and all other rows of
*-- the same oringal source line with replacd = .t., and also update the matchlen

		Update &tcCursor ;
			Set Replaced = .T., ;
			Replace_DT = Datetime(),;
			MatchLen = Max(MatchLen + m.lnChangeLength, 0) ;
			Where Alltrim(FilePath) == m.lcFileToModify And ;
			Recno = m.lnResultRecno And ;
			Column = m.lcColumn And ;
			MatchStart = m.lnMatchStart

*-- Update the stored code with the new code for all records of the same original source
		Update &tcCursor ;
			Set Code = m.toResult.cNewCode;
			Where Alltrim(FilePath) == m.lcFileToModify And ;
			Recno = m.lnResultRecno And ;
			Column = m.lcColumn

*-- Update matchstart values on remaining records of same file, recno, and column type
		Update &tcCursor ;
			Set MatchStart = (MatchStart + m.lnChangeLength) ;
			Where Alltrim(FilePath) == m.lcFileToModify And ;
			Recno = m.lnResultRecno And ;
			Column = m.lcColumn And ;
			MatchStart > m.lnMatchStart

*-- Update procstart values on remaining records of same file, recno, and column type
		Update &tcCursor ;
			Set ProcStart = (ProcStart + m.lnChangeLength) ;
			Where Alltrim(FilePath) == m.lcFileToModify And ;
			Recno = m.lnResultRecno And ;
			Column = m.lcColumn And ;
			ProcStart > m.lnProcStart

		Goto (m.lnCurrentRecno)

		Select (m.lnSelect)

	Endproc


*----------------------------------------------------------------------------------
	Procedure UpdateProgressBar(tnValue)

		If Vartype(This.oProgressBar) = 'O'
			This.oProgressBar.nValue = m.tnValue
		Endif

	Endproc
Enddefine

*!*	Changed by: nmpetkov 27.3.2023
*!*	<pdm>
*!*	<change date="{^2023-03-27,15:45:00}">Changed by: nmpetkov<br />
*!*	Changes to  Highlight searched text in opened window #75
*!*	</change>
*!*	</pdm>
Function _WOnTop
Function _EdGetEnv
Function _EdGetStr
Function _EdSelect
Function _EDGETPOS
*!*	/Changed by: nmpetkov 27.3.2023

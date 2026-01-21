#Include GoFish.h

#Define ccMatchWordStart	Chr[2] + chr[4]
#Define ccMatchWordEnd 		Chr[3] + chr[4]
#Define	ccBinaries			' SCX VCX MNX PJX FRX LBX '

Define Class GoFishSearchEngine As Custom

	cCR_StoreLocal                   = Addbs(Home(7) + 'GoFish_')

	cBackupPRG   = 'GoFishBackup.prg'
	cFilesToSkip = ''

* This string contains a list of files to be skipped during the search.
* One filname on each line. This list is only skipped if lSkipFiles is .t.
	cFilesToSkipFile                 = ''  && 'GF_Files_To_Skip.txt'    -- See Init
	cFilesToIncludeFile              = ''  && 'GF_Files_To_Include.txt' -- See Init

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
	nFilesSearched					 = 0

* How many matched lines found in the last search. Note: this counts lines that had a match, note each match.
* It's possible that one line could have multiple matches.
	nMatchLines                      = 0

	nReplaceCount                    = 0

	nReplaceFileCount                = 0

* Store the current ID number of the Replace History record.
	nReplaceHistoryId                = 0

* 1 = Regular replace, 2 = Advanced Replace (UDF Replace)
	nReplaceMode                     = 1

* Tells how long the last search took. It is only reset by SearchInPath() or SearchInProject() took,
* and not by the lower level search like SearchInFile, or SearchInTextFile, or SearchInTable.
	nSearchStartTime                 = 0
	nSearchTime                      = 0
	nADirTime		 				 = 0

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
	
	*** JRN 2024-03-10 : In Code View pane, # lines to precede the match
	* provided for those where the code view does not auto-align
	nCodeViewLinesBefore = 1000
	
	dSearchFromDate 	= {}
	dSearchToDate	 	= {}
	
	cWildCardSearch = ''

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
			This.cProjects = This.cProjects + m.tcProject + CR
		Endif

	Endproc


* ================================================================================
	Procedure AnyRegExMatches
		Lparameters tcCode
		Local loMatches
	
		loMatches = This.oRegExForSearchInCode.Execute(m.tcCode)
		Return m.loMatches.Count > 0
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

		If m.lcFileType # 'EXT'
			If Upper(m.lcMatchType) # 'RESERVED3' And This.IsFullLineComment(m.lcTrimmedMatchLine)
				toObject.MatchType = MatchType_Comment
				This.CreateResult(m.toObject)
				Return .Null. && Exit out, we're done with this record!
			Endif
		
			*=============================================================================================
			* Handle a few tweaks on MatchType assignments
			*=============================================================================================
			This.ProcessInlineComments(m.toObject)		
		Endif
		
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
		Local lcBackupPRG As String
		Local lcDestFile As String
		Local lcExt As String
		Local lcExtensions As String
		Local lcSourceFile As String
		Local lcThisBackupFolder As String
		Local llCopyError As Boolean
		Local lnI As Number
		Local laExtensions[1], lcBackupFolder, loException
	
		*		ccBACKUPFOLDER = Addbs(This.cCR_StoreLocal + 'GoFishBackups')
		lcBackupFolder = Addbs(This.cCR_StoreLocal + 'GF_ReplaceBackups')
		*/SF 20230131 -> issue #41
		*/SF 20221018 -> local storage
	
	
		If This.oSearchOptions.lpreviewreplace
			Return
		Endif
	
		llCopyError = .F.
	
		*-- If the user has created a custom backup PRG, and placed it in their path, then call it instead
		lcBackupPRG = 'GoFish_Backup.prg'
	
		If File(m.lcBackupPRG)
			Do &lcBackupPRG With m.tcFilePath, m.tnReplaceHistoryId
			Return
		Endif
	
		If Not Directory (m.lcBackupFolder) && Create main folder for backups, if necessary
			Mkdir (m.lcBackupFolder)
			GF_Write_Readme_Text(4, Addbs(m.lcBackupFolder) + 'README.md', .T.)
	
		Endif
	
		* Create folder for this ReplaceHistorrID, if necessary
		lcThisBackupFolder = Addbs (m.lcBackupFolder + Transform (m.tnReplaceHistoryId))
	
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
			lcSourceFile = Forceext (m.tcFilePath, m.laExtensions (m.lnI))
			lcDestFile	 = m.lcThisBackupFolder + Justfname (m.lcSourceFile)
			If Not File (m.lcDestFile) && no need to repeat if already backed up
				*** JRN 2024-08-22 : MPR and MPX files need not exist
				If File(m.lcSourceFile) or not Upper(JustExt(m.lcSourceFile)) $ 'MPR,MPX'
					Try
						Copy File (m.lcSourceFile) To (m.lcDestFile)
					Catch To m.loException
						If Not m.llCopyError
							This.SetReplaceError('Error creating backup of file - ' + m.loException.Message, m.tcFilePath, m.tnReplaceHistoryId)
						Endif
						llCopyError = .T.
					Endtry
				EndIf
			Endif
		Endfor
	
		Return Not m.llCopyError
	
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

		lnSeconds = Seconds()
		* ================================================================================

		*!*  Apparently this bothers to skip the .git folder, if any, but causes
		*!*    screen flashing because of the call to Run

			*!* ******** JRN Removed 2024-03-03 ********
			*!* *get the toplevel folder, if we are in a repo
			*!* 		If !m.tlWithRepo Then
			*!* 			lcCommand = 'git rev-parse --git-dir>git_x.tmp'
			*!* 			Run &lcCommand

			*!* 			If File('git_x.tmp') Then
			*!* *the result is either the git base folder or empty for no git repo
			*!* 				tcRepo = Upper(Fullpath(Chrtran(Filetostr('git_x.tmp'), '/' + CR + LF, '\')))
			*!* 				Delete File git_x.tmp
			*!* 				tlWithRepo = .T.
			*!* 			Else &&file('git_x.tmp')
			*!* * no file, no git
			*!* 				tcRepo = ''
			*!* 			Endif &&file('git_x.tmp')

			*!* 		Endif &&!m.tlWithRepo

			*!* *SF 20221216 special folders to skip
			*!* 		If Upper(Fullpath(m.tcDir)) == m.tcRepo
			*!* *git toplevel folder. We do not look this up
			*!* 			Return 0
			*!* 		EndIf
		
		* ================================================================================
		
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
		EndFor
		
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
			This.cProjects = This.cProjects + m.lcProject + CR
		Endfor

*-- Add any Projects in the current folder
		Adir(laProjects, m.lcCurrentDir + '*.pjx')

		For lnX = 1 To Alen(m.laProjects) / 5
			lcProject = Lower(Fullpath(laProjects(m.lnX, 1)))
			This.AddProject(m.lcProject)
			This.cProjects = This.cProjects + m.lcProject + CR
		Endfor

*-- Add MRU Projects to the Collection...
		loPEME_BaseTools = Createobject('GF_PEME_BaseTools')

		loMRU_Projects = loPEME_BaseTools.GetMRUList('PJX')

		For Each m.loMRU_Project In m.loMRU_Projects
			lcProject = Lower(m.loMRU_Project)
			This.AddProject(m.lcProject)
			This.cProjects = This.cProjects + m.lcProject + CR
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
			TrimmedReplaceLine c(254) Null, ;
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
			Code             = '' && .Code && No longer any need to capture this in the memo field.
			Type			 = .Type

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
	Procedure EditFromCurrentRow(tcCursor, tlSelectControl, tlMoveToTopleft)
	
		Local lcClass As String
		Local lcCodeBlock As String
		Local lcExt As String
		Local lcFileToEdit As String
		Local lcMatchType As String
		Local lcMethod As String
		Local lcMethodString As String
		Local lcName As String
		Local lcProperty As String
		Local lnMatchStart As Number
		Local lnProcStart As Number
		Local lnRecNo As Number
		Local lnStart As Number
		Local lnWords As Number
		Local loPBT As 'GF_PEME_BaseTools'
		Local loTools As Object
		Local lcKeystrokes, lnProcedureLineOffset, lnSelect, loResults, loDefaultButton
	
		lnSelect = Select()
	
		Select (m.tcCursor)
		Scatter Name m.loResults Memo
		Select (m.lnSelect)
		
		If This.oSearchOptions.lPreVFP9
			This.EditPreVFP9(m.loResults)
			Return
		EndIf 

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
		If Not Inlist(m.lcExt, 'PRG', 'SPR', 'MPR', 'QPR', 'H', 'INI', 'TXT', 'XML', 'HTM')			;
				And (Empty(m.lcMethod) Or 0 # Atc('<Property', m.lcMatchType))
			*!*	/Changed by: nmpetkov 27.3.2023
			lcMethodString = ''
			lnStart		   = 1
		Else
			lcMethodString = Alltrim(m.lcName + '.' + m.lcMethod, 1, '.')
	
			If m.lcExt $ ' SCX VCX '
				*-- Calculate Line No from procstart and matchstart postitions...
				lcCodeBlock	= &tcCursor..ProcCode
				lnStart		= Getwordcount(m.lcCodeBlock, CR) - 1 && The LINE NUMBER that match in on within the method
				lnStart		= Iif(m.lnStart > 0, m.lnStart, 1)
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
			m.loPBT.EditSourceX(m.lcFileToEdit)
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
	
				Do Case
					Case m.lcExt = 'SCX' And '.' $ m.lcName  && Must trim off form name from front of object name
						lcName = Substr(lcName, 1 + At('.', m.lcName)) 
					Case m.lcExt = 'SCX'  && Must trim off form name from front of object name
						lcName = ''
				Endcase
	
				Do Case
					Case m.lcMatchType = MatchType_Name
						m.loPBT.EditSourceX(m.lcFileToEdit, m.lcClass)
						m.loTools.SelectObject(m.lcName, , .T.)
						Return
	
					Case m.lcMatchType $ (MATCHTYPE_PROPERTY_NAME + MATCHTYPE_PROPERTY_VALUE + MATCHTYPE_PROPERTY_DEF )
						*-- Pull out the Property name from the MatchLine (it can be preceded by an object name)
						lcProperty = Getwordnum(&tcCursor..TrimmedMatchLine, 1)
						lnWords	   = Getwordcount(m.lcProperty, '.')
						lcProperty = Getwordnum(m.lcProperty, m.lnWords, '. ')
						lcProperty = This.FixPropertyName(m.lcProperty)
	
						m.loPBT.EditSourceX(m.lcFileToEdit, m.lcClass)
						If m.tlSelectControl 
							This.SelectControl(loTools, lcName, lcProperty)	
						EndIf 
						Return

					Case InList(m.lcExt, 'SCX', 'VCX') and not Empty(m.lcName) 
						Local loDefaultButton, loThisObject, loTopMostParent
						
						m.loPBT.EditSourceX(m.lcFileToEdit, m.lcClass)
						If m.tlSelectControl
							This.SelectControl(loTools, lcName)	
						Endif
												
				Endcase
			Endif
	
		Endif
	
		If m.lcExt = 'MNX'
			*** JRN 2024-02-05 : special handling for MNXs - if possible, using the keyboard buffer
			* to navigate to the actual record and procedure
			lcKeystrokes = This.GetMenuKeystrokes(m.lcFileToEdit, m.lnRecNo, m.lcMatchType)
			m.loPBT.EditSourceX(m.lcFileToEdit, m.lcClass, m.lnStart, m.lnStart, m.lcMethodString, m.lnRecNo)
			If Not Empty(m.lcKeystrokes)
				Keyboard(m.lcKeystrokes)
			Endif
		Else
			m.loPBT.EditSourceX(m.lcFileToEdit, m.lcClass, m.lnStart, m.lnStart, m.lcMethodString, m.lnRecNo)
		Endif
	

		*!*	Changed by: nmpetkov 27.3.2023
		*!*	<pdm>
		*!*	<change date="{^2023-03-27,15:45:00}">Changed by: nmpetkov<br />
		*!*	Changes to  Highlight searched text in opened window #75
		*!*	</change>
		*!*	</pdm>
		lcMatchType = Alltrim(&tcCursor..MatchType)
		*	Try to select searched text if found in normal windows only - exclude internal for VFP places
		Do Case
			Case Inlist(m.lcMatchType, MatchType_Filename, MATCHTYPE_CLASS_DEF, MATCHTYPE_CLASS_DESC, MATCHTYPE_METHOD_DEF, MATCHTYPE_PROPERTY_DEF, ;
					  MATCHTYPE_CONTAINING_CLASS, MatchType_ParentClass, MatchType_BaseClass, MATCHTYPE_METHOD_DESC, MATCHTYPE_PROPERTY, ;
					  MATCHTYPE_PROPERTY_DESC, MATCHTYPE_PROPERTY_NAME, MATCHTYPE_PROPERTY_VALUE)
			Case Inlist(m.lcExt, 'SCX', 'VCX')
				lcProcCode = Alltrim(loResults.ProcCode, 1, LF) 
				lnProcedureLineOffset = Atc(LF, m.lcProcCode)
				lnStart = &tcCursor..MatchStart - &tcCursor..ProcStart - m.lnProcedureLineOffset
				This.SelectSearchedText(m.lnStart, &tcCursor..MatchLen, Trim(&tcCursor..Search), &tcCursor..MatchLine)
			Otherwise
				This.SelectSearchedText(&tcCursor..MatchStart, &tcCursor..MatchLen, Trim(&tcCursor..Search), &tcCursor..MatchLine)
		Endcase
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
	Procedure SelectSearchedText(tnRangeStart, tnRangelen, tcSearch, tcMatchLine)
		Local lLibrRelease As Boolean
		Local lcFoxtoolsFll As String
		Local lcLine As String
		Local llMatchCase As Boolean
		Local lnPos As Number
		Local lnRangeEnd As Number
		Local lnRangeStart As Number
		Local lnRetCode As Number
		Local lnSelEnd As Number
		Local lnSelStart As Number
		Local lnWHandle As Number
		Local loMatch As Object
		Local loMatches As Object
		Local lcSearch, lcText
		Local aEdEnv[25]
			
		If Atc('foxtools.fll', Set('LIBRARY')) = 0
			lcFoxtoolsFll = Sys(2004) + 'foxtools.fll'
			If File(m.lcFoxtoolsFll)
				lLibrRelease = .T.
				Set Library To (m.lcFoxtoolsFll) Additive
			Endif
		Endif
	
		If Atc('foxtools.fll', Set('LIBRARY')) > 0
			lnWHandle = _WOnTop()
			lnRetCode = _EDGetenv(m.lnWHandle, @m.aEdEnv)
			* aEdEnv: 
			*	1 - filename
			*	2 - size
			*  12 - readonly?
			*  17 - selected start
			*  18 - selected end
			
			If m.lnRetCode = 1 And m.aEdEnv[2] > 0 && content size is > 0
					
				*** JRN 2024-04-19 : Per issue #221, problem if last line does not have following CR
				lcText = _EdGetStr(m.lnWHandle, m.tnRangeStart, Min(aEdEnv[2], m.tnRangeStart + m.tnRangelen))
	
				If Atc(m.tcMatchLine, m.lcText) > 0 And This.oSearchOptions.nSearchMode # 3
					If This.oSearchOptions.nSearchMode = 2 And '*' $ m.tcSearch
						lcSearch = Left(m.tcSearch, At('*', m.tcSearch) - 1)
					Else
						lcSearch = m.tcSearch
					Endif
					lnSelStart = m.tnRangeStart + Atc(m.lcSearch, m.lcText) - 1
					lnSelEnd   = m.lnSelStart + Len(m.lcSearch)
				Else
	
					* ================================================================================ 
					lnRangeStart = m.tnRangeStart
					If Empty(m.tnRangelen)
						tnRangelen = m.aEdEnv[2] - m.lnRangeStart + 1
					Endif
					lnRangeEnd = m.lnRangeStart + m.tnRangelen && determine where the search to be searched :-)
					If m.lnRangeEnd > m.aEdEnv[2] && check we are not beyond the end, will throw error
						lnRangeEnd = m.aEdEnv[2]
					Endif
					lcLine = _EdGetStr(m.lnWHandle, m.lnRangeStart, m.lnRangeEnd)
					* determine real string to search in case pattern or Regex
					llMatchCase = This.oSearchOptions.lMatchCase
					If This.oSearchOptions.nSearchMode > 1
						This.PrepareRegExForSearch()
						This.PrepareRegExForReplace()
						loMatches = This.oRegExForSearch.Execute(m.lcLine)
						If m.loMatches.Count > 0
							loMatch		= m.loMatches.Item(0)
							tcSearch	= m.loMatch.Value
							llMatchCase	= .T.
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
					* ================================================================================ 
				Endif
	
				* select at the end
				If m.lnSelEnd > m.lnSelStart
					_EdSelect(m.lnWHandle, m.lnSelStart, m.lnSelEnd)
					_EdStoPos(m.lnWHandle, m.lnSelStart, .T.)
				Endif
			Endif
		Endif
	
		If m.lLibrRelease And Atc(m.lcFoxtoolsFll, Set('LIBRARY')) > 0
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


	Procedure EditPreVFP9(toResults)
		Local lcClass, lcClipText, lcEditSource, lcExt, lcFileText, lcFileToEdit, lcKeyStrokes, lcMatchType
		Local lcMethod, lcMethodString, lcName, lcProcCode, lnLineNumber, lnMatchStart, lnProcStart, lnRecNo
		Local loException
	
		lcExt		 = Alltrim(Upper(m.toResults.FileType))
		lcFileToEdit = Upper(Alltrim(m.toResults.FilePath))
		lcClass		 = Alltrim(m.toResults.Class)
		lcName		 = Alltrim(m.toResults.Name)
		lcMethod	 = Alltrim(m.toResults.MethodName)
		lcMatchType	 = Alltrim(m.toResults.MatchType)
		lnRecNo		 = m.toResults.Recno
		lnProcStart	 = m.toResults.ProcStart
		lnMatchStart = m.toResults.MatchStart
		lcProcCode	 = m.toResults.ProcCode
	
		lcClipText	 = ["] + m.lcFileToEdit + ["]
		lnLineNumber = 0
	
		Do Case
			Case m.lcExt $ 'VCX SCX' And Not Empty(m.lcMethod)
				lnLineNumber   = Occurs(CR, Left(m.lcProcCode, m.lnMatchStart - m.lnProcStart))
				lcMethodString = Alltrim(m.lcName + '.' + m.lcMethod, 1, '.')
				lcClipText	   = m.lcClipText + Textmerge([, <<m.lnLineNumber>>, "<<m.lcClass>>", "<<m.lcMethodString>>"])
	
			Case m.lcExt $ 'VCX SCX' And Not Empty(m.lcClass)
				lcClipText = m.lcClipText + Textmerge([, 0, "<<m.lcClass>>"])
	
			Case m.lcExt $ 'PRG' And m.lcMatchType  = 'Code'
				Try
					lcFileText = Filetostr(m.lcFileToEdit)
				Catch To m.loException
	
				Endtry
				lnLineNumber = Occurs(CR, Left(m.lcFileText, m.lnMatchStart)) + 1
				lcClipText	 = m.lcClipText + Textmerge([, <<m.lnLineNumber>>])
	
			Case m.lcExt $ 'MNX'
				lcKeyStrokes = This.GetMenuKeystrokes(m.lcFileToEdit, m.lnRecNo, m.lcMatchType)
				lcClipText	 = m.lcClipText + Textmerge([, , , , "<<m.lcKeystrokes>>"])
	
		Endcase
	
		lcEditSource = This.oSearchOptions.lPreVFP9EditSource
		Do Case
			Case Empty(m.lcEditSource)
				_Cliptext = m.lcClipText
			Case Upper(Getwordnum(m.lcEditSource, Getwordcount(m.lcEditSource))) == 'WITH'
				_Cliptext = m.lcEditSource + ' ' + m.lcClipText
			Otherwise
				_Cliptext = m.lcEditSource + [(] + m.lcClipText + [)]
		Endcase
	
		Wait (Left(_Cliptext, 150)) Window At Mrow(), Mcol() Nowait
	
	Endproc
										
*----------------------------------------------------------------------------------
	Procedure EndTimer()

		This.nSearchTime = This.ElapsedTimeSinceStart()

	Endproc


*----------------------------------------------------------------------------------
	Procedure ElapsedTimeSinceStart()

		Return Seconds() - This.nSearchStartTime

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

		If (CR + m.lcFileName + CR) $ This.cFilesToSkip
			Return .T.
		Endif

		For lnI = 1 To This.nWildCardFilesToSkip
			If Like(This.aWildcardFiles[m.lni], Upper(m.tcFile))
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
				If lnX = 1
					loReturn.StartByte = 0
					loReturn.EndByte   = Result.StartByte - 1
				EndIf 
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
	
		Local lcLastLine As String
		Local lcMatchLine As String
		Local lcPreceding As String
		Local lcProcCode As String
		Local lcResult As String
		Local lnCRPos As Number
		Local lnLen As Number
		Local lnLength As Number
		Local lnStart As Number
		Local lnTextStart As Number
		Local lcFirstWord
	
		lnStart	 = m.loObject.MatchStart - m.loObject.ProcStart + 1
	
		*!* ** { JRN -- 08/05/2016 07:16 AM - Begin
		*!* lnLength = m.loObject.MatchLen - 1
		lnLength = m.loObject.MatchLen
		*!* ** } JRN -- 08/05/2016 07:16 AM - End
	
		lcProcCode = Evl(m.loObject.ProcCode, m.loObject.Code)
	
		* previously, assumed trailing CR, but this dropped off last character if not found
		*!* ** { JRN -- 08/05/2016 07:16 AM - Begin
		*!* lcMatchLine	= Substr(m.lcProcCode, m.lnStart, m.lnLength)
		lcMatchLine	= Trim(Substr(m.lcProcCode, m.lnStart, m.lnLength), 1, CR, LF)
		*!* ** } JRN -- 08/05/2016 07:16 AM - End
		lcResult	= m.lcMatchLine
	
		lcFirstWord = Lower(Getwordnum(m.lcMatchLine, 1, ' ' + Tab + CR + LF))
		Do Case
			Case 'text' = m.lcFirstWord
	
				&& everything up to and including EndText
				lnLength = Len(m.lcResult)
				Do While Len(m.lcProcCode) > m.lnStart + m.lnLength
					lcLastLine = Substr(m.lcProcCode, m.lnStart + m.lnLength)
					lnLen	   = At(CR, m.lcLastLine, 2)
					If m.lnLen > 0
						lcLastLine = Left(m.lcLastLine, m.lnLen - 1)
					Endif
					lcResult = m.lcResult + m.lcLastLine
					lnLength = Len(m.lcResult)
					If 'endtext' = Lower(Getwordnum(m.lcLastLine, 1, ' ' + Tab + CR + LF))
						Exit
					Endif
				Enddo
	
			Case 'endtext' = m.lcFirstWord
				&& Nothing to do here, this is always left as is
	
			Otherwise
	
				*** JRN 12/02/2015 : Add in leading lines, if any
				Do While .T.
					lcPreceding	= Left(m.lcProcCode, m.lnStart - 1)
					lnCRPos		= Rat(CR, m.lcPreceding, 2)
					If m.lnCRPos > 0
						lcPreceding = Substr(m.lcPreceding, m.lnCRPos + 1)
					Endif
					If This.IsContinuation(m.lcPreceding)
						lcResult = m.lcPreceding + m.lcResult
						lnStart	 = m.lnStart - Len(m.lcPreceding)
						lnLength = Len(m.lcResult)
					Else
						Exit
					Endif
				Enddo
	
				*** JRN 12/02/2015 : Add in following lines, if any
				lcLastLine = m.lcMatchLine
				Do While This.IsContinuation(m.lcLastLine)
					lcLastLine = Substr(m.lcProcCode, m.lnStart + m.lnLength)
					lnLen	   = At(CR, m.lcLastLine, 2)
					If m.lnLen > 0
						lcLastLine = Left(m.lcLastLine, m.lnLen - 1)
					Endif
					lcResult = m.lcResult + m.lcLastLine
					lnLength = Len(m.lcResult)
				Enddo
	
				*** JRN 12/05/2015 : within Text / Endtext
				If Atc('text', Left(m.lcProcCode, m.lnStart)) # 0 And		;
						Atc('endtext', Substr(m.lcProcCode, m.lnStart + m.lnLength)) # 0
	
					lnTextStart = m.lnStart
					Do While m.lnTextStart > 1
						lcPreceding	= Left(m.lcProcCode, m.lnTextStart - 1)
						lnCRPos		= Rat(CR, m.lcPreceding, 2)
						If m.lnCRPos > 0
							lcPreceding = Substr(m.lcPreceding, m.lnCRPos + 1)
						Endif
						Do Case
							Case 'text' = Lower(Getwordnum(m.lcPreceding, 1, ' ' + Tab + CR + LF))
								lnTextStart	= m.lnTextStart - Len(m.lcPreceding)
								lnLength	= m.lnLength + m.lnStart - m.lnTextStart
								lnStart		= m.lnTextStart
								lcResult	= Substr(m.lcProcCode, m.lnStart, m.lnLength)
								Do While Len(m.lcProcCode) > m.lnStart + m.lnLength
									lcLastLine = Substr(m.lcProcCode, m.lnStart + m.lnLength)
									lnLen	   = At(CR, m.lcLastLine, 2)
									If m.lnLen > 0
										lcLastLine = Left(m.lcLastLine, m.lnLen - 1)
									Endif
									lcResult = m.lcResult + m.lcLastLine
									lnLength = Len(m.lcResult)
									If 'endtext' = Lower(Getwordnum(m.lcLastLine, 1, ' ' + Tab + CR + LF))
										Exit
									Endif
								Enddo
								Exit
							Case 'endtext' = Lower(Getwordnum(m.lcPreceding, 1, ' ' + Tab + CR + LF))
								Exit
							Otherwise
								lnTextStart = m.lnTextStart - Len(m.lcPreceding)
						Endcase
					Enddo
				Endif
	
		Endcase
	
		loObject.Statement		= m.lcResult
		loObject.StatementStart	= m.lnStart
	
	Endproc
		

*----------------------------------------------------------------------------------
	Procedure FixPropertyName(lcProperty)

* Gets rid of dimensions and leading '*^'

		Return Chrtran(Getwordnum(m.lcProperty, 1, '(['), '*^', '')

	Endproc


*----------------------------------------------------------------------------------
	Procedure GenerateHTMLCode(tcCode, tcMatchLine, tnMatchStart, tcCss, tcJavaScript, tcReplaceLine, ;
			tlAlreadyReplaced, tnTabsToSpaces, tlDarkMode)
	
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
			
			* ================================================================================
			*** JRN 2024-03-10 : truncate lines shown before the match
			Local lnMaxLines			
			lnMaxLines = This.nCodeViewLinesBefore
			Do Case
				Case m.lnMaxLines <= 0
					lcLeft = ''
				Case Occurs(LF, m.lcLeft) > m.lnMaxLines
					lcLeft = Substr(m.lcLeft, Rat(LF, m.lcLeft, m.lnMaxLines + 1) + 1)
			Endcase
						* ================================================================================ 
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
**** this odd next line is created as something to be executed
			lcReplaceExpression = ['] + ccMatchWordStart + [' + lcMatch + '] + ccMatchWordEnd + ['] 

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
			lcColorizedCode = Strtran(m.lcColorizedCode, ccMatchWordStart, m.lcMatchWordPrefix)
			lcColorizedCode = Strtran(m.lcColorizedCode, ccMatchWordEnd,   m.lcMatchWordSuffix)

*-- Finally, add <div> tags around the entire Matched Line -------------------
			lcMatchLine = m.lcMatchLinePrefix + m.lcColorizedCode + m.lcMatchLineSuffix
*=================================================================================================

*-- Dress up the code that comes after the match line...
*-- (Look for EndProc to know where to end the code)---
			*!* ******** JRN Removed 2024-08-20 ********
			*!* If m.tlAlreadyReplaced
			*!* 	lcRightCode = Substr(m.tcCode, m.tnMatchStart + 1 + m.lnReplaceLineLength)
			*!* Else
			*!* 	lcRightCode = Substr(m.tcCode, m.tnMatchStart + 1 + m.lnMatchLineLength)
			*!* Endif

			lcRightCode = Substr(m.tcCode, m.tnMatchStart + 1 + m.lnMatchLineLength)

			*!* ******** JRN Removed 2024-03-16 ******** redundant
			*!* lnEndProc = Atc('EndProc', m.lcRightCode)

			*!* If m.lnEndProc > 0
			*!* 	lcRightCode = Substr(m.lcRightCode, 1, m.lnEndProc + 6) && It ends at "E" of "EndProc", so add 6 to get the rest of the word
			*!* Endif

			lcRight = This.HtmlEncode(m.lcRightCode)

			lcHtmlBody = m.lcLeft + m.lcMatchLine + m.lcReplaceLine + m.lcRight &&Build the body
			If Not Empty(m.tnTabsToSpaces)
				lcHtmlBody = Strtran(m.lcHtmlBody, Chr[9], Space(m.tnTabsToSpaces))
			Endif
			lcHtmlBody = Alltrim(m.lcHtmlBody, 1, CR, LF)
		Else

*-- Just a plain blob of VFP code, with no match lines or match words...
*-- Need an empty MatchLine Divs so the JavaScript on the page will find it to scroll the page
			lcHtmlBody = '<div id="matchline"></div>' + This.HtmlEncode(m.tcCode)

		Endif

*-- Build the whole Html by combining the html parts defined above -------------
		TEXT To m.lcHTML Noshow Textmerge Pretext 3
<html>
 <head>
  <<Iif(m.tlDarkMode, [<body style="background-color:Rgb(30,30,30);">], [])>>
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


* Returns the child object, a command button, with Default = .T.
* ================================================================================
	Procedure GetDefaultButton(loParent)
		Local laObjects[1], loDefault, loException, loObject
	
		For Each m.loObject In m.loParent.Objects FoxObject
			If Pemstatus(m.loObject, 'Default', 5) And m.loObject.Default
				Return m.loObject
			Endif
			If Pemstatus(m.loObject, 'Objects', 5)
				Try
					loDefault = This.GetDefaultButton(m.loObject)
				Catch To m.loException
					loDefault = Null
				Endtry
				If Not Isnull(m.loDefault)
					Return m.loDefault
				Endif
			Endif
		Endfor
		Return Null
	Endproc
						
*----------------------------------------------------------------------------------
	Procedure GetDirectories(tcPath, tlIncludeSubdirectories)
	
		Local;
			lnFiles As Number,;
			lnSeconds  

		Local Array;
			laFiles(1)

		lnSeconds = Seconds() 
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

		This.nADirTime = This.ElapsedTimeSinceStart()

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
		Local lShowMenuBarNumbers, laField[1], laParent[1], lcDBF, lcPrompt, lnLevelName, lnRecNo
	
		lShowMenuBarNumbers	= This.oSearchOptions.lShowMenuBarNumbers
		lcDBF				= Dbf()
		lnRecNo				= Recno()
	
		Select  LevelName,					;
				Prompt,						;
				Int(Val(ItemNum))			;
			From (m.lcDBF)					;
			Where Recno() = m.lnRecNo		;
			Into Array laField
	
		lcPrompt = ["] + m.laField[2] + '"' + Iif(m.lShowMenuBarNumbers, ' (Bar ' + Transform(m.laField[3]) + ')', '')
	
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
					Prompt,						;
					Int(Val(ItemNum))			;
				From (m.lcDBF)					;
				Where Recno() = m.lnRecNo		;
				Into Array laField
	
			lcPrompt = ["] + m.laField[2] + '"' + Iif(m.lShowMenuBarNumbers, ' (Bar ' + Transform(m.laField[3]) + ')', '') + ' => ' + m.lcPrompt
	
		Enddo
	
	
		Return Strtran(m.lcPrompt, '\<', '')
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
				*** JRN 2024-09-10 : The "Comment" field shows up in all caps
				* to distinguish it from comments in a procedure
				Case m.lcMatchType = '<Comment>'
					lcKeystrokes = m.lcKeystrokes + ccTab + ccTab + ccEnter
				Case Upper(m.lcMatchType) = '<PROCEDURE>'
					lcKeystrokes = m.lcKeystrokes + ccTab + ccTab + ccEnter
				Case Upper(m.lcMatchType) # '<PROMPT>' 
					lcKeystrokes = m.lcKeystrokes + ccTab + ccTab + ccTab + ccEnter
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

			lnLFs = Occurs(LF, m.loMatch.Value)
			lnX   = 0
* ignore leading CRLF's, and [spaces and tabs, except on the matched line]
			Do While Substr(m.tcCode, m.lnStartByte + 1, 1) $ LF + CR + Chr(32) + Chr(9) And m.lnX < m.lnLFs
				If Substr(m.tcCode, m.lnStartByte + 1, 1) = LF
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

			Case Inlist(m.lcMatchType, MATCHTYPE_INCLUDE_FILE, '<Expr>', '<Supexpr>', '<Picture>', '<Prompt>', '<Procedure>', '<Command>', ;
					'<Skipfor>', '<Message>', '<Tag>', '<Tag2>');
					Or ;
					M.toObject.UserField.FileType = 'DBF'

				lnReturn = 2

			Case Inlist(m.lcMatchType, MATCHTYPE_CODE, MATCHTYPE_COMMENT, '<Cleanup>') Or;
					(m.toObject.UserField.IsText And !Inlist(m.lcMatchType, MATCHTYPE_FILENAME, MATCHTYPE_TIMESTAMP))

				lnReturn = 1

		Endcase

		Return m.lnReturn

	Endproc


*----------------------------------------------------------------------------------


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
	Procedure IncludeExtension(tcSelected, tcExtension)
	
		If m.tcSelected
			This.oSearchOptions.cSearchExtensions = This.oSearchOptions.cSearchExtensions + m.tcExtension + ' '
		Endif
	
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
		This.cFilesToIncludeFile     = This.cCR_StoreLocal + 'GF_Files_To_Include.txt'
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
	
		Local lcFileType, loMatches
	
		* If file template supplied, use it
		If Not Empty(This.oSearchOptions.cFileTemplate)
			loMatches = This.oSearchOptions.oRegExpFileTemplate.Execute(Justfname(m.tcFile))
			* failure if no match
			If m.loMatches.Count = 0
				Return .F.
			Endif
			* success if match and no extension supplied 
			* if extension IS supplied, use the normal list of extensions
			If '.' $ This.oSearchOptions.cFileTemplate
				Return .T.
			Endif
		Endif
	
		If '*' == This.oSearchOptions.cOtherIncludes
			*check for all-in-cOtherIncludes
			Return .T.
		Endif
	
		lcFileType = Upper(Justext(m.tcFile))
	
		*-- Table-based Files --------------------------------------
		If (This.oSearchOptions.lIncludeSCX And m.lcFileType = 'SCX')					;
				Or (This.oSearchOptions.lIncludeVCX  And m.lcFileType = 'VCX' )			;
				Or (This.oSearchOptions.lIncludeFRX  And m.lcFileType = 'FRX' )			;
				Or (This.oSearchOptions.lIncludeDBC  And m.lcFileType = 'DBC' )			;
				Or (This.oSearchOptions.lIncludeMNX  And m.lcFileType = 'MNX' )			;
				Or (This.oSearchOptions.lIncludeLBX  And m.lcFileType = 'LBX' )			;
				Or (This.oSearchOptions.lIncludePJX  And m.lcFileType = 'PJX' )			;
				Or (This.oSearchOptions.lIncludePRG  And m.lcFileType = 'PRG' )			;
				Or (This.oSearchOptions.lIncludeMPR  And m.lcFileType = 'MPR' )			;
				Or (This.oSearchOptions.lIncludeTXT  And m.lcFileType = 'TXT' )			;
				Or (This.oSearchOptions.lIncludeINI  And m.lcFileType = 'INI' )			;
				Or (This.oSearchOptions.lIncludeH    And m.lcFileType = 'H'   )			;
				Or (This.oSearchOptions.lIncludeXML  And m.lcFileType = 'XML' )			;
				Or (This.oSearchOptions.lIncludeSPR  And m.lcFileType = 'SPR' )			;
				Or (This.oSearchOptions.lIncludeASP  And m.lcFileType = 'ASP' )			;
				Or (This.oSearchOptions.lIncludeJSP  And m.lcFileType = 'JSP' )			;
				Or (This.oSearchOptions.lIncludeJAVA And m.lcFileType = 'JAVA')
			Return .T.
		Endif
	
		*-- Code based files (any HTM* file) ----------------------------------
		If This.oSearchOptions.lIncludeHTML And 'HTM' $ m.lcFileType
			Return .T.
		Endif
	
		*-- Lastly, is it match with other includes??? (but only if no template, because this disables extensions)
		If Not Empty(This.oSearchOptions.cOtherIncludes) And m.lcFileType $ Upper(This.oSearchOptions.cOtherIncludes)
			Return .T.
		Endif
	
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
		Return This.oSearchOptions.nSearchMode = GF_SEARCH_MODE_LIKE and '*' $ This.oSearchOptions.cSearchExpression and This.oSearchOptions.lWildCardContinuationLines
	EndProc 	


*----------------------------------------------------------------------------------
	Procedure LoadOptions(tcFile, tlInit)
	
		Local lcProperty As String
		Local loMy As 'My' Of 'My.vcx'
		Local laProperties[1], lnI
	
		If Not File(m.tcFile)
			Return .F.
		Endif
	
		*-- Get an array of properties that are on the SearchOptions object
		Amembers(laProperties, This.oSearchOptions, 0, 'U')
	
		*-- Load settings from file...
		loMy = Newobject('My', 'My.vcx')
		m.loMy.Settings.Load(m.tcFile)
	
		*--- Scan over Object properties, and look for a corresponding props on the My Settings object (if present)
		With m.loMy.Settings
			For lnI = 1 To Alen(m.laProperties)
				lcProperty = m.laProperties[m.lnI]
				If Type('.' + m.lcProperty) # 'U'
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
	
		If m.tlInit and This.oSearchOptions.nSearchScope = 6 && Results
			This.oSearchOptions.nSearchScope = This.oSearchOptions.nPreviousSearchScope
		Endif
	
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

* ================================================================================
    Procedure MarkUnchangeableRows(tcGridCursor)
    	Local lcResultsAlias, lcTemp, lnSelect
    
    	lnSelect = Select()
    
    	lcResultsAlias = This.cSearchResultsAlias
    	Select  Distinct FilePath,								;
    			.F.    As  InUse								;
    		From (m.tcGridCursor )								;
    		Where Not (This.IsTextFile(Trim(FileName)))			;
    		Into Cursor crsrFiles Readwrite
    	lcTemp = Sys(2015)
    
    	Scan
    		If This.OpenTableForReplace(Trim(FilePath), m.lcTemp, 0)
    			Use
    		Else
    			Replace InUse With .T.
    		Endif
    	Endscan
    
    	Update  Target												;
    		Set TrimmedReplaceLine = Null							;
    		From (m.tcGridCursor)    As  Target						;
    			Join crsrFiles										;
    				On Target.FilePath = crsrFiles.FilePath			;
    		Where crsrFiles.InUse
    
    	Select(m.lnSelect)
    
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
				This.SetReplaceError('Cannot open file for exclusive use: ', m.tcFileToOpen, m.tnResultId)
				Select (m.lnSelect)
				llReturn = .F.
		Endtry

		Return m.llReturn

	Endproc


*----------------------------------------------------------------------------------
	Procedure PrepareForSearch

		Local ldToDate
	
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

		This.SetSearchedExtensions()
	
		This.dSearchFromDate = Evl(This.oSearchOptions.dTimeStampFrom, {^1900-01-01})
		ldToDate			 = Evl(This.oSearchOptions.dTimeStampTo, {^9999-01-01})
		This.dSearchToDate	 = m.ldToDate + 1 &&86400 && Must bump into to next day, since TimeStamp from table has time on it
	
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
		Local lcSearchExpression, lcWord, lnI
	
		lcSearchExpression = This.oSearchOptions.cSearchExpression
		*** JRN 2024-02-14 : "Normal" regex for non wild-card searches
		This.PrepareRegExForSearchV2(This.oRegExForSearch, m.lcSearchExpression, .T.)
	
		If This.IsWildCardStatementSearch()
			*** JRN 2024-06-13 : nah, this fails later, must only search for first word
			*** JRN 2024-02-14 : for wild card searches, only search for the longest text between *s (NOPE!)
			If This.oSearchOptions.lMatchWholeWord
				This.oSearchOptions.cWholeWordSearch = This.PrepareForWholeWordSearch(m.lcSearchExpression)
			EndIf
			lcSearchExpression = Getwordnum(This.oSearchOptions.cSearchExpression, 1, '*')
			*** JRN 2024-09-23 : If first expression is too short, choose longest
			If Len(m.lcSearchExpression) < 4
				For lnI = 2 To Getwordcount(This.oSearchOptions.cSearchExpression, '*')
					lcWord = Getwordnum(This.oSearchOptions.cSearchExpression, m.lnI, '*')
					If Len(m.lcWord) > Len(m.lcSearchExpression)
						lcSearchExpression = m.lcWord
					Endif
				EndFor
			EndIf 
		Endif
		This.PrepareRegExForSearchV2(This.oRegExForSearchInCode, m.lcSearchExpression, .F.)
		This.cWildCardSearch = lcSearchExpression
	
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
			Replace TrimmedReplaceLine with Null in (m.tcCursor)

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
	
		Local lcCode, lcLeft, lcLineFromFile, lcMatchLine, lcNewCode, lcReplaceExpression, lcReplaceLine
		Local lcRight, lnLength, lnLineToChangeLength, lnMNXOffset, lnMatchStart, lnStart, loRegEx, loResult
	
		loResult = This.GetReplaceResultObject()
		lcCode	 = m.toReplace.Code
	
		lcMatchLine			 = Left(m.toReplace.MatchLine, m.toReplace.MatchLen)
		lnMatchStart		 = m.toReplace.MatchStart
		lnLineToChangeLength = Len(m.lcMatchLine)
		* ================================================================================ 
		*** JRN 2024-08-20 : special handling for MNXs, as field "proccode" does not only have
		* code, but the other fields (so that it appears in Code View pane appropriately)
		Do Case
			Case Empty(m.toReplace.Type)
				lcLineFromFile = Substr(m.lcCode, m.lnMatchStart + 1, m.lnLineToChangeLength)
				lcLeft		   = Left(m.lcCode, m.lnMatchStart)
				lcRight		   = Substr(m.lcCode, m.lnMatchStart + 1 + m.lnLineToChangeLength)
			Case m.toReplace.MatchType = MATCHTYPE_PROCEDURE or m.toReplace.MatchType = '<Comment>'
				lnStart		   = Val(Getwordnum(m.toReplace.Type, 1))
				lnLength	   = Val(Getwordnum(m.toReplace.Type, 2))
				lnMNXOffset	   = m.lnStart - 1
				
				lcLineFromFile = Substr(m.lcCode, m.lnMatchStart + 1 - m.lnMNXOffset, m.lnLineToChangeLength)
				lcLeft		   = Left(m.lcCode, m.lnMatchStart - m.lnMNXOffset)
				lcRight		   = Substr(m.lcCode, m.lnMatchStart + 1 + m.lnLineToChangeLength - m.lnMNXOffset)
			Otherwise
						
				lnStart		= Val(Getwordnum(m.toReplace.Type, 1))
				lnLength	= Val(Getwordnum(m.toReplace.Type, 2))
				lnMNXOffset	= m.lnStart - 1
				
				lcMatchLine    = Substr(toReplace.ProcCode, lnStart, lnLength)
				lcLineFromFile = lcCode
				lcLeft 		   = ''
				lcRight        = ''
				
		EndCase
		* ================================================================================ 
	
		If m.lcLineFromFile # m.lcMatchLine && Ensure that line from file still matches the passed in line from the orginal search!!
			*** JRN 2024-03-13 : For MNXs, MatchLine may have irrelevant trailing CR
			If m.lcLineFromFile # Trim(m.lcMatchLine, 1, CR, LF)
				This.SetReplaceError('Source file has changed since original search:', Alltrim(m.toReplace.FilePath), m.toReplace.Id)
				loResult.lError = .T.
				Return m.loResult
			Else
				lcMatchLine = m.lcLineFromFile
			Endif
		Endif
	
		*-- IMPORTANT CODE HERE... Revised code line is determined here!!!! -------------
		If Empty(m.tcReplaceLine)
			loRegEx				= This.oRegExForSearch
			lcReplaceExpression	= This.oSearchOptions.cReplaceExpression
			Do Case
				Case This.nReplaceMode = 1
					lcReplaceLine = m.loRegEx.Replace(m.lcMatchLine, m.lcReplaceExpression)
				Case This.nReplaceMode = 2
					lcReplaceLine = ''
				Case This.nReplaceMode = 3 And Not Empty(This.cReplaceUDFCode)
					lcReplaceLine = This.ReplaceLineWithUDF(m.lcMatchLine)
				Otherwise
					lcReplaceLine = m.lcMatchLine
			Endcase
		Else
			lcReplaceLine = m.tcReplaceLine
		Endif
	
		*--Added this in 4.3.014 to handle case of deleting the entire line
		If Empty(m.lcReplaceLine)
			lcRight = Ltrim(m.lcRight, 0, LF) && Need to strip off initial LF of Right hand code block
		Endif
	
		lcNewCode = m.lcLeft + m.lcReplaceLine + m.lcRight
	
		With m.loResult
			.nChangeLength = Len(m.lcReplaceLine) - Len(m.lcMatchLine)
			*--Added this in 4.3.014 to handle case of deleting the entire line
			If Empty(m.lcReplaceLine)
				.nChangeLength = .nChangeLength - 1 && to account for the LF we stripped off above
			Endif
			.cNewCode			 = m.lcNewCode
			.cReplaceLine		 = m.lcReplaceLine
			.cTrimmedReplaceLine = This.TrimWhiteSpace(.cReplaceLine)
		Endwith
	
		toReplace.ReplaceLine		 = m.loResult.cReplaceLine
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

		If m.lnLastChar = 10 && Editbox will add a LF so this has to be stripped off
			lcReplaceLine = Left(m.lcReplaceLine, Len(m.lcReplaceLine) - 1)
		Endif

		lnLastChar = Asc(Right(m.lcReplaceLine, 1))

		If m.lnLastChar <> 13 And !Empty(m.lcReplaceLine) && Make sure user has not stripped of the CR that came with the MatchLine
			lcReplaceLine = m.lcReplaceLine + CR
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
	
		Local lcMatchLine As String
		Local lcReplaceLine As String
		Local llCR As Boolean
	
		lcMatchLine = m.tcMatchLine
	
		*-- If there is a CR at the end, pull it off before calling the UDF. Will add back later...
		If Right(m.tcMatchLine, 1) = CR
			llCR		= .T.
			lcMatchLine	= Left(m.tcMatchLine, Len(m.tcMatchLine) - 1)
		Endif
	
		*-- Call the UDF ---------------
		Try
			lcReplaceLine = Execscript(This.cReplaceUDFCode, m.lcMatchLine)
		Catch
			lcReplaceLine = m.lcMatchLine && Keep the line the same if UDF failed
		Finally
		Endtry
	
		Do Case
			Case Isnull(m.lcReplaceLine)
				lcReplaceLine = ''
			Case Vartype(m.lcReplaceLine) # 'C'
				lcReplaceLine = m.tcMatchLine
			Case m.llCR
				lcReplaceLine = m.lcReplaceLine + CR
		Endcase
	
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
				
				Local lnRecNo				
				lnRecNo = Recno(m.tcCursor)
				Update Target Set TrimmedReplaceLine = Null From (m.tcCursor) As Target Where FilePath = lcFile
				Goto (m.lnRecNo) In (m.tcCursor)
								
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
	Procedure SearchInSingleProject(toProject, ttTime, tcUni)
	
		Local lcFile As String
		Local lnReturn As Number
		Local lnSelect As Number
		Local lnX As Number
		Local laProjectFiles[1], lcProjectPath, lnSeconds, loFile, loProject
	
		This.PrepareForSearch()
		This.StartTimer()
		This.StartProgressBar(_vfp.Projects.Count)
	
		lnSeconds	  = Seconds()
		lnSelect	  = Select()
		This.tRunTime = Evl(m.ttTime, Datetime())
		This.cUni	  = Evl(m.tcUni, '_' + Sys(2007, Ttoc(This.tRunTime), 0, 1))
	
		Create Cursor ProjectFiles (FileName C(200))
	
		loProject					 = m.toProject
		This.oSearchOptions.cProject = m.loProject.Name
		lcProjectPath				 = Addbs(Justpath(Alltrim(m.loProject.Name)))
	
		Insert Into ProjectFiles (FileName) Values (Lower(m.loProject.Name))
	
		For Each m.loFile In m.loProject.Files FoxObject
			If m.loFile.Type $ 'EHKMPRVBdTxD'
				Insert Into ProjectFiles (FileName) Values (m.loFile.Name)
			Endif
		Endfor
		loFile	  = Null
	
		Select  Distinct *															;
			From ProjectFiles														;
			Where Not (Upper(Justext(FileName)) $ This.cGraphicsExtensions)			;
			Into Array laProjectFiles
	
		This.nADirTime = This.ElapsedTimeSinceStart()
	
		If Type('laProjectFiles') = 'L'
			This.SearchFinished(m.lnSelect)
			Return 1
		Endif
	
		This.StartProgressBar(Alen(m.laProjectFiles))
	
		*** Uncomment this code to track the execution time for the result
		*!*			LOCAL nSeconds
		*!*			nSeconds = SECONDS()
	
		For lnX = 1 To Alen(m.laProjectFiles)
	
			lcFile = m.laProjectFiles(m.lnX)
	
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
	
	
* ================================================================================ 
	Procedure SearchIncludedFiles
		Local laFiles[1], laLines[1], lcFile, lcFolder, lcIncludeFile, lnCount, lnFileCount, lnI, lnJ
		Local lnReturn, lnSelect
	
		If Not This.oSearchOptions.lIncludeFiles
			Return 1
		Endif
	
		If This.lEscPress Or This.nMatchLines >= This.oSearchOptions.nMaxResults
			Return 1
		Endif
	
		lcIncludeFile = This.cFilesToIncludeFile
		If Not File(m.lcIncludeFile)
			Return 1
		Endif
	
		Create Cursor IncludeFiles(FileName M)
		lnCount = Alines(laLines, Filetostr(m.lcIncludeFile), 5)
		For lnI = 1 To m.lnCount
			lcFile = m.laLines[m.lni]
			If Not '*' $ m.lcFile
				lcFile = Fullpath(m.lcFile)
				Do Case
					Case File(m.lcFile)
						Insert Into IncludeFiles Values (Lower(m.lcFile))
					Case Directory(m.lcFile)
						lcFolder	= Addbs(m.lcFile)
						lnFileCount	= Adir(laFiles, m.lcFolder + '*.*')
						For lnJ = 1 To m.lnFileCount
							Insert Into IncludeFiles Values (Lower(m.lcFolder + m.laFiles[m.lnJ, 1]))
						Endfor
				Endcase
			Endif
		Endfor
	
		If Reccount('IncludeFiles') > 0
	
			lnSelect = Select()
			Select IncludeFiles
	
			This.StartProgressBar(Reccount())
	
			Scan
	
				lcFile = Trim(FileName)
	
				lnReturn = This.SearchInFile(m.lcFile)
	
				This.UpdateProgressBar(Recno())
	
				If (m.lnReturn < 0) Or This.lEscPress Or This.nMatchLines >= This.oSearchOptions.nMaxResults
					Exit
				Endif
	
			Endscan
	
			This.SearchFinished(m.lnSelect)
	
		EndIf
		
		Return 0
	
	Endproc
														

*----------------------------------------------------------------------------------
	Procedure SearchInCode(tcCode, tuUserField, tlHasProcedures, lnMinMatchStart, lnMaxMatchStart, tcType)

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
			lcErrorMessage = 'Error processing regular expression. ' + This.oRegExForSearch.Pattern
			lcErrorMessage = m.lcErrorMessage + CRLF + Space(5) + 'Search In Code'
			If Vartype(tuUserField) = 'O'
				lcErrorMessage = m.lcErrorMessage + CRLF + Space(10) + 'File: ' + m.tuUserField.FilePath
				lcErrorMessage = m.lcErrorMessage + CRLF + Space(10) + 'Column: ' + m.tuUserField.Column
				lcErrorMessage = m.lcErrorMessage + CRLF + Space(10) + 'Record: ' + Transform(m.tuUserField.Recno)
				lcErrorMessage = m.lcErrorMessage + CRLF + Space(10) + 'Code: ' + Left(Transform(tcCode), 50)
			Endif
			This.SetSearchError(m.lcErrorMessage)
			Return - 1
		Endif

		If m.lnMatchCount > 0

			loProcedureStartPositions = Iif(m.tlHasProcedures, This.GetProcedureStartPositions(m.tcCode), .Null.)
			lnMatchCount = 0

			For Each m.loMatch In m.loMatches FoxObject

				If m.tlHasProcedures And !This.oSearchOptions.lSearchInComments And This.IsComment(m.loMatch.Value)
					Loop
				Endif
				If m.loMatch.FirstIndex < Evl(lnMinMatchStart, 0) or m.loMatch.FirstIndex >= Evl(lnMaxMatchStart, 1E9) 
					Loop
				EndIf
				
				loProcedure = This.FindProcedureForMatch(m.loProcedureStartPositions, m.loMatch.FirstIndex)
				loObject    = Createobject('GF_SearchResult')

				With m.loObject
					.UserField  = m.tuUserField
					.oMatch     = m.loMatch
					.oProcedure = m.loProcedure

					.Type = Evl(tcType, '')

* .ContainingClass =	.oProcedure._ClassName	&& Not used on this object. This line to be deleted after testing. (2012-07-11))
					.MethodName = .oProcedure._Name
					.ProcStart  = .oProcedure.StartByte
					*!* ******** JRN Removed 2024-11-18 ********
					*!* .procend    = Min(Evl(.oProcedure.EndByte, Len(m.tccode)), .oMatch.FirstIndex + MEMOFIELDMINSIZE)
					.procend    = Evl(.oProcedure.EndByte, Len(m.tccode))
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

				lnMatchCount = lnMatchCount + 1

				If !This.ProcessSearchResult(m.loObject)
					Exit
				Endif &&!This.ProcessSearchResult(loObject)

			Endfor

		Endif

		Select(m.lnSelect)

		Return m.lnMatchCount

	Endproc


*----------------------------------------------------------------------------------
	Procedure SearchInCustomUDFCursor(tcCustomScopeUDFFileName, m.ttTime, m.tcUni)
	
		Local lcCustomAlias, lnReturn, lnSeconds, lnSelect
	
		If Empty(m.tcCustomScopeUDFFileName) Or Not File(m.tcCustomScopeUDFFileName)
			This.SetSearchError('Custom Scope UDF not defined' + CRLF + CRLF + [See "New in V7" page in the Options form], 16)
			Return 0
		Endif
	
		lnSeconds = Seconds()
		lnSelect  = Select()
	
		This.tRunTime = Evl(m.ttTime, Datetime())
		This.cUni	  = Evl(m.tcUni, '_' + Sys(2007, Ttoc(This.tRunTime), 0, 1))
	
		This.PrepareForSearch()
		This.StartTimer()
		This.oProgressBar.Start(100, 'Creating list of files')
	
		lcCustomAlias = 'GF_CustomScope' + Sys(2015)
		Create Cursor (m.lcCustomAlias) (FileName C(240))
	
		Do (m.tcCustomScopeUDFFileName) With This.oSearchOptions, m.lcCustomAlias
		
		*!* ******** JRN Removed 2024-06-24 ******** 
		*!* * This was not enough of an improvement
		*!* This.PreprocessFileList(m.lcCustomAlias, m.lcCustomAlias)
	
		This.nADirTime = This.ElapsedTimeSinceStart()
	
		Select Distinct Lower(FileName) As FileName From (m.lcCustomAlias) Into Cursor (m.lcCustomAlias)
		This.StartProgressBar(Reccount())
	
		Scan
	
			lnReturn = This.SearchInFile(Alltrim(FileName))
	
			This.UpdateProgressBar(This.nFilesProcessed)
	
			If (m.lnReturn < 0) Or This.lEscPress Or This.nMatchLines >= This.oSearchOptions.nMaxResults
				Exit
			Endif
	
		Endscan
	
		This.SearchFinished(m.lnSelect)
	
		Use In (m.lcCustomAlias)
	
		If Evl(m.lnReturn, 0) >= 0
			Return 1
		Else
			Return m.lnReturn
		Endif
	
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
			If This.CanUseGrepForFileList()
				Return 0
			EndIf 
			This.lFileNotFound = .T.
			This.SetSearchError('File not found: ' + m.tcFile)
*This.ReduceProgressBarMaxValue(1)
			Return 0
		Endif

		This.ShowWaitMessage(m.tcFile)

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
	
		Local lcErrorMessage As String
		Local ldFileDate As Date
		Local ldFromDate As Date
		Local ldToDate As Date
		Local lnMatchCount As Number
		Local lnSelect As Number
		Local loFileResultObject As 'GF_FileResult'
		Local loMatches As Object
		Local loSearchResultObject As 'GF_SearchResult'
		Local lcFileToStr, loException
	
		lnSelect = Select()
	
		If Not File(m.tcFile)
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
			lcErrorMessage = 'Error processing regular expression.  ' + This.oRegExForSearch.Pattern
			lcErrorMessage = lcErrorMessage + CRLF + '  Search In File Name - ' + Transform(tcFile)
			This.SetSearchError(m.lcErrorMessage)
			Return - 1
		Endif
	
		If m.lnMatchCount = 0 And Not Empty(This.oSearchOptions.cSearchExpression)
			Return 0
		Endif
	
		ldFileDate = This.GetFileDateTime(m.tcFile)
	
		ldFromDate = Evl(This.oSearchOptions.dTimeStampFrom, {^1900-01-01})
		ldToDate   = Evl(This.oSearchOptions.dTimeStampTo, {^9999-01-01})
		ldToDate   = m.ldToDate + 1 &&86400 && Must bump into to next day, since TimeStamp from table has time on it
	
		If This.oSearchOptions.lTimeStamp And Not Between(m.ldFileDate, m.ldFromDate, m.ldToDate)
			Return 0
		Endif
	
		loFileResultObject = Createobject('GF_FileResult')	&& This custom class has all the properties that must be populated if you want to
		&& have a cursor created
		With m.loFileResultObject
			.FileName  = Justfname(m.tcFile)
			.FilePath  = m.tcFile
			.MatchType = MatchType_Filename
			.FileType  = Upper(Justext(m.tcFile))
			.Timestamp = m.ldFileDate
	
			.MatchLine		  = 'File name = "' + .FileName + '"'
			.TrimmedMatchLine = .MatchLine
		Endwith
	
		loSearchResultObject = Createobject('GF_SearchResult')
		With m.loSearchResultObject
			.UserField		  = m.loFileResultObject
			.MatchType		  = MatchType_Filename
			.MatchLine		  = 'File name = "' + m.loFileResultObject.FileName + '"'
			.TrimmedMatchLine = 'File name = "' + m.loFileResultObject.FileName + '"'
		Endwith
	
		If This.IsTextFile(m.tcFile)
			*** JRN 2024-02-05 : Correction to show the entire file
			* 	(Previously, leading characters got lost)
			Try
				lcFileToStr = Filetostr(m.tcFile)
			Catch To m.loException
				lcFileToStr = ''
			Endtry
			loSearchResultObject.Code = m.loSearchResultObject.MatchLine						;
				+ Replicate('=', Max(Len(m.loSearchResultObject.MatchLine), 60)) + CR		;
				+ Left(m.lcFileToStr, 7000) 
		Endif
	
		This.ProcessSearchResult(m.loSearchResultObject)
	
		Select (m.lnSelect)
	
		Return 1
	
	Endproc
	

*----------------------------------------------------------------------------------
	Procedure SearchInOpenProjects(tcProject, ttTime, tcUni)
	
		Local laProjectFiles[1], lcFile, lnCount, lnI, lnReturn, lnSeconds, lnSelect, lnX, loFile, loProject
	
		This.PrepareForSearch()
		This.StartTimer()
	
		lnSeconds	  = Seconds()
		lnSelect	  = Select()
		This.tRunTime = Evl(m.ttTime, Datetime())
		This.cUni	  = Evl(m.tcUni, '_' + Sys(2007, Ttoc(This.tRunTime), 0, 1))
	
		Create Cursor ProjectFiles (FileName C(200))
		lnCount =  _vfp.Projects.Count
		This.oProgressBar.Start(m.lnCount, 'Creating list of files from ' + Transform(m.lnCount) + ' project(s).')

		For lnI = 1 To m.lnCount
			loProject					 = _vfp.Projects[m.lnI]
			This.oSearchOptions.cProject = m.loProject.Name
			Insert Into ProjectFiles (FileName) Values (Lower(m.loProject.Name))
	
			For Each m.loFile In m.loProject.Files FoxObject
				If m.loFile.Type $ 'EHKMPRVBdTxD'
					Insert Into ProjectFiles (FileName) Values (m.loFile.Name)
				Endif
			Endfor
			loFile					 = Null
			loProject				 = Null
			This.oProgressBar.nValue = This.oProgressBar.nValue + 1
		Endfor
	
		Select  Distinct *															;
			From ProjectFiles														;
			Where Not (Upper(Justext(FileName)) $ This.cGraphicsExtensions)			;
			Into Array laProjectFiles
	
		This.nADirTime = This.ElapsedTimeSinceStart()
	
		If Type('laProjectFiles') = 'L'
			This.SearchFinished(m.lnSelect)
			Return 1
		Endif
	
		This.StartProgressBar(Alen(m.laProjectFiles))
	
		For lnX = 1 To Alen(m.laProjectFiles)
	
			lcFile = m.laProjectFiles(m.lnX)
	
			lnReturn = This.SearchInFile(m.lcFile)
	
			This.UpdateProgressBar(This.nFilesProcessed)
	
			If (m.lnReturn < 0) Or This.lEscPress Or This.nMatchLines >= This.oSearchOptions.nMaxResults
				Exit
			Endif
	
		Endfor
	
		This.SearchFinished(m.lnSelect)
	
		If m.lnReturn >= 0
			Return 1
		Else
			Return m.lnReturn
		Endif
	
	Endproc
			
	
*----------------------------------------------------------------------------------
	Procedure SearchProjectsInCurDir(tcProject, ttTime, tcUni)
	
		Local laProjectFiles[1], laProjects[1], lcCurDir, lcFile, lcFileName, lcScope, llAlreadyOpen
		Local llOpened, lnI, lnJ, lnProjects, lnReturn, lnSeconds, lnSelect, lnX, loException, loFile
		Local loProject
	
		This.PrepareForSearch()
		This.StartTimer()
	
		lnSeconds	  = Seconds()
		lnSelect	  = Select()
		This.tRunTime = Evl(m.ttTime, Datetime())
		This.cUni	  = Evl(m.tcUni, '_' + Sys(2007, Ttoc(This.tRunTime), 0, 1))
	
		lcCurDir   = Addbs(Curdir())
		lnProjects = Adir(laProjects, m.lcCurDir + '*.pjx')
	
		If m.lnProjects = 0
			lcScope	 = Alltrim(Lower(This.GetCurrentDirectory()), '\')
			lnReturn = This.SearchInPath(m.lcScope, m.ttTime, m.tcUni)
			Return m.lnReturn
		Endif
	
		Create Cursor ProjectFiles (FileName C(200))
		This.oProgressBar.Start(m.lnProjects, 'Creating list of files from ' + Transform(m.lnProjects) + ' project(s).')
	
		For lnI = 1 To m.lnProjects
			lcFileName	  = m.lcCurDir + m.laProjects[m.lnI, 1]
			llAlreadyOpen = .F.
			For lnJ = 1 To _vfp.Projects.Count
				If Lower(m.lcFileName) $ Lower(_vfp.Projects[m.lnJ].Name)
					loProject	  = _vfp.Projects[m.lnJ]
					llAlreadyOpen = .T.
				Endif
			Endfor
	
			If Not m.llAlreadyOpen
				Try
					Modify Project (m.lcFileName) Nowait Noshow
					loProject = _vfp.ActiveProject
					llOpened  = .T.
				Catch To m.loException
					llOpened = .F.
				Endtry
	
				If Not m.llOpened
					Loop
				Endif
			Endif
	
			This.oSearchOptions.cProject = m.loProject.Name
			Insert Into ProjectFiles (FileName) Values (Lower(m.loProject.Name))
	
			For Each m.loFile In m.loProject.Files FoxObject
				If m.loFile.Type $ 'EHKMPRVBdTxD'
					Insert Into ProjectFiles (FileName) Values (m.loFile.Name)
				Endif
			Endfor
	
			If Not m.llAlreadyOpen
				m.loProject.Close()
			Endif
	
			loFile	  = Null
			loProject = Null
	
			This.oProgressBar.nValue = This.oProgressBar.nValue + 1
	
		Endfor
	
	
		Select  Distinct *															;
			From ProjectFiles														;
			Where Not (Upper(Justext(FileName)) $ This.cGraphicsExtensions)			;
			Into Array laProjectFiles
	
		This.nADirTime = This.ElapsedTimeSinceStart()
	
		If Type('laProjectFiles') = 'L'
			This.SearchFinished(m.lnSelect)
			Return 1
		Endif
	
		This.StartProgressBar(Alen(m.laProjectFiles))
	
		For lnX = 1 To Alen(m.laProjectFiles)
	
			lcFile = m.laProjectFiles(m.lnX)
	
			lnReturn = This.SearchInFile(m.lcFile)
	
			This.UpdateProgressBar(This.nFilesProcessed)
	
			If (m.lnReturn < 0) Or This.lEscPress Or This.nMatchLines >= This.oSearchOptions.nMaxResults
				Exit
			Endif
	
		Endfor
	
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

		This.PrepareForSearch()
		This.StartTimer()

		If This.CanUseGrepForFileList()
			This.SearchUsingGrep(m.tcPath, m.ttTime, m.tcUni)
			Return 1
		EndIf 

		*** JRN 2024-05-31 : old (slow) style
		This.tRunTime             = Evl(m.ttTime, Datetime())
		This.cUni                 = Evl(m.tcUni, "_" + Sys(2007, Ttoc(This.tRunTime), 0, 1))
		This.oSearchOptions.cPath = m.tcPath
		
		This.StoreInitialDefaultDir()

		If !This.ChangeCurrentDir(m.tcPath) && If there was a problem CD-ing into the starting path
			This.RestoreDefaultDir()
			Return - 1
		Endif

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
			lnX         As Number,;
			lnSeconds

		Local Array;
			laProjectFiles(1)

		lnSeconds = Seconds() 
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

		This.nADirTime = This.ElapsedTimeSinceStart()

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
			lcFile = GF_Fullpath(m.lcFile, m.lcProjectPath)
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

		This.SearchFinished(m.lnSelect)

		If m.lnReturn >= 0
			Return 1
		Else
			Return m.lnReturn
		Endif

	Endproc


*----------------------------------------------------------------------------------
	Procedure SearchInResults(tcResultsCursor, m.ttTime, m.tcUni)
	
		Local lcCustomAlias, lcFilter, lnReturn, lnSeconds, lnSelect
	
		This.PrepareForSearch()
		This.StartTimer()
	
		lnSelect  = Select()
	
		This.tRunTime = Evl(m.ttTime, Datetime())
		This.cUni	  = Evl(m.tcUni, '_' + Sys(2007, Ttoc(This.tRunTime), 0, 1))
	
		This.nADirTime = This.ElapsedTimeSinceStart()
	
		Select (m.tcResultsCursor)
		This.StartProgressBar(Reccount())
	
		Scan
	
			lnReturn = This.SearchInFile(Alltrim(FileName))
	
			This.UpdateProgressBar(This.nFilesProcessed)
	
			If (m.lnReturn < 0) Or This.lEscPress Or This.nMatchLines >= This.oSearchOptions.nMaxResults
				Exit
			Endif
	
		Endscan
	
		This.SearchFinished(m.lnSelect)
	
		Use In (m.tcResultsCursor)
	
		If m.lnReturn >= 0
			Return 1
		Else
			Return m.lnReturn
		Endif
	
	Endproc
	

*----------------------------------------------------------------------------------
	Procedure SearchInTable(tcFile)
	
		Local lcClass As String
		Local lcCode As String
		Local lcDataType As String
		Local lcDeleted As String
		Local lcExt As String
		Local lcField As String
		Local lcFieldSource As String
		Local lcFormClass As String
		Local lcFormClassloc As String
		Local lcFormName As String
		Local lcName As String
		Local lcObjectType As String
		Local lcParent As String
		Local lcParentName As String
		Local lcProject As String
		Local lcSearchExpression As String
		Local ldFromDate As Date
		Local ldMaxTimeStamp As Date
		Local ldToDate As Date
		Local llContinueError As Boolean
		Local llHasMethods As Boolean
		Local llIgnorePropertiesField As Boolean
		Local llLocateError As Boolean
		Local llProcessThisMatch As Boolean
		Local llScxVcx As Boolean
		Local lnEndColumn As Number
		Local lnMatchCount As Number
		Local lnParentID As Number
		Local lnSelect As Number
		Local lnStart As Number
		Local lnStartColumn As Number
		Local lnTotalMatches As Number
		Local loException As Object
		Local loFileResultObject As 'GF_FileResult'
		Local loSearchResultObject As 'GF_SearchResult'
		Local laMaxTimeStamp[1], laParent[1], lcText, lcType, llThisField, lnMaxMatchStart, lnMinMatchStart
	
		*:Global		;
		ii
	
		lnSelect = Select()
	
		lnMatchCount   = 0
		lnTotalMatches = 0
		lcExt		   = Upper(Justext(m.tcFile))
		lcProject	   = This.oSearchOptions.cProject
	
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
					lcFormName	   = ''
					lcFormClass	   = ''
					lcFormClassloc = ''
				Else
					Locate For BaseClass = 'form'
					lcFormName	   = ObjName
					lcFormClass	   = Class
					lcFormClassloc = CLASSLoc
				Endif
			Endif
	
		Endif
	
		If This.oSearchOptions.lTimeStamp And Type('timestamp') = 'U'
			Use In 'GF_TableSearch'
			Return 0
		Endif
	
		* This.ShowWaitMessage(m.tcFile)
	
		lnEndColumn				= 255
		llIgnorePropertiesField	= .F.
	
		Do Case
			Case 'VCX' $ m.lcExt
				lnStartColumn			= 4
				llIgnorePropertiesField	= This.oSearchOptions.lIgnorePropertiesField
			Case 'SCX' $ m.lcExt
				lnStartColumn			= 4
				llIgnorePropertiesField	= This.oSearchOptions.lIgnorePropertiesField
				* lnEndColumn = 12
			Case 'FRX' = m.lcExt
				lnStartColumn = 3 && Newer reports could start at col 6, but older reports can have text data starting in column 3
				* lnEndColumn = 21
				If Len(Field('timestamp', 'GF_TableSearch')) > 0 && Some really old reports may not have this field.
					Select  Max(Timestamp)			;
						From 'GF_TableSearch'		;
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
			lcField		  = Upper(Field(m.ii))
			llLocateError = .F.
	
			If Empty(m.lcField)
				Exit
			Endif
	
			If  Not Type(m.lcField) $ 'CM' Or					; && If not a character or Memo field
				('TAG' $ m.lcField And m.lcExt # 'FRX') Or		;
					Inlist(m.lcField, 'OBJCODE', 'OBJECT', 'SYMBOLS', 'LEVELNAME')
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
			If This.oSearchOptions.lRegularExpression
				Locate For This.AnyRegExMatches(Evaluate(m.lcField)) Nooptimize
	
			Else
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
	
				lnMatchCount	   = 0
				loFileResultObject = Createobject('GF_FileResult')	&& This custom class has all the properties that must be populated if you want to
				llProcessThisMatch = .T.														&& have a cursor created
				llScxVcx		   = Inlist(m.lcExt, 'VCX', 'SCX')
				lcCode			   = Evaluate(m.lcField)
				lnMinMatchStart	   = 0
				lnMaxMatchStart	   = 1E9
				*** JRN 2024-08-18 : re-purposing this empty field to contain information
				* required to do replacments in MNXs; this occurs because of the
				* way that matches to MNXs appear in code view pane
				lcType			   = ''
	
				With m.loFileResultObject
					.Process   = .F.
					.FileName  = Justfname(m.tcFile)
					.FilePath  = m.tcFile
					.MatchType = iif(m.lcField = 'COMMENT', m.lcfield, Proper(m.lcField))
					.FileType  = Upper(m.lcExt)
					.Column	   = m.lcField
					.IsText	   = .F.
					.Recno	   = Recno()
					.Timestamp = Iif(Type('timestamp') # 'U', Ctot(This.TimeStampToDate(Timestamp)), {// :: AM})
	
					lcClass			 = Iif(Type('class') # 'U', Class, '')
					.ContainingClass = m.lcClass
					._ParentClass	 = m.lcClass
					._BaseClass		 = Iif(Type('baseclass') # 'U', BaseClass, '')
					.CLASSLoc		 = Iif(Type('classloc') # 'U', CLASSLoc, '')
	
					lcParent = Iif(Type('parent') # 'U', Parent, '')
					lcName	 = Iif(Type('objname') # 'U', ObjName, '')
					._Name	 = Alltrim(m.lcParent + '.' + m.lcName, '.')
	
					Do Case
						Case m.lcExt = 'SCX'
	
							._Class = ''
	
						Case m.lcExt = 'VCX'
	
							If Not Empty(m.lcParent)
								._Class = Getwordnum(m.lcParent, 1, '.')
							Else
								._Class	= Alltrim(ObjName)
								._Name	= ''
							Endif
	
						Case m.lcExt = 'FRX'
							._Name	= Name
							._Class	= This.GetFrxObjectType(ObjType, objCode)
							If Empty(.Timestamp)
								.Timestamp = m.ldMaxTimeStamp
							Endif
	
						Case m.lcExt = 'DBF'
							.MatchType = '<Field>'
							._Name	   = Proper(m.lcField)
	
							*** JRN 2024-02-05 : For some MNX matches, show more info from the same record
						Case m.lcExt = 'MNX' And m.lcField # 'LEVELNAME'
	
							lnMinMatchStart	= 0
							lnMaxMatchStart	= 1E9
							lcCode			= ''
	
							If Not Empty(Prompt)
								llThisField		= m.lcField = 'PROMPT'
								lnMinMatchStart	= Iif(m.llThisField, Len(m.lcCode), m.lnMinMatchStart)
								lcCode			= m.lcCode + 'Prompt    = '	+ This.GetFullMenuPrompt()
								If m.llThisField
									lcText = ["] + Trim(Prompt) + ["]
									lcType = Transform(Rat(m.lcText, m.lcCode) + 1) + ' ' + Transform(Len(Trim(Prompt)))
								Endif
								lcCode			= m.lcCode + Iif(objCode = 77, ' &' + '& Submenu', '')	+ CRLF
								lnMaxMatchStart	= Iif(m.llThisField, Len(m.lcCode), m.lnMaxMatchStart)
							Endif
	
							If Not Empty(Command)
								llThisField		= m.lcField = 'COMMAND'
								lnMinMatchStart	= Iif(m.llThisField, Len(m.lcCode), m.lnMinMatchStart)
								lcType			= Iif(m.llThisField, Transform(Len(m.lcCode) + 13) + ' ' + Transform(Len(Command)), m.lcType)
								lcCode			= m.lcCode + 'Command   = ' + Command + CRLF
								lnMaxMatchStart	= Iif(m.llThisField, Len(m.lcCode), m.lnMaxMatchStart)
							Endif
	
							If Not Empty(KeyName)
								llThisField		= m.lcField = 'KEYNAME'
								lnMinMatchStart	= Iif(m.llThisField, Len(m.lcCode), m.lnMinMatchStart)
								lcType			= Iif(m.llThisField, Transform(Len(m.lcCode) + 13) + ' ' + Transform(Len(KeyName)), m.lcType)
								lcCode			= m.lcCode + 'Key Label = ' + KeyName + CRLF
								lnMaxMatchStart	= Iif(m.llThisField, Len(m.lcCode), m.lnMaxMatchStart)
							Endif
	
							If Not Empty(KeyLabel)
								llThisField		= m.lcField = 'KEYLABEL'
								lnMinMatchStart	= Iif(m.llThisField, Len(m.lcCode), m.lnMinMatchStart)
								lcType			= Iif(m.llThisField, Transform(Len(m.lcCode) + 13) + ' ' + Transform(Len(KeyLabel)), m.lcType)
								lcCode			= m.lcCode + 'Key Text  = ' + KeyLabel + CRLF
								lnMaxMatchStart	= Iif(m.llThisField, Len(m.lcCode), m.lnMaxMatchStart)
							Endif
	
							If Not Empty(SkipFor)
								llThisField		= m.lcField = 'SKIPFOR'
								lnMinMatchStart	= Iif(m.llThisField, Len(m.lcCode), m.lnMinMatchStart)
								lcType			= Iif(m.llThisField, Transform(Len(m.lcCode) + 13) + ' ' + Transform(Len(SkipFor)), m.lcType)
								lcCode			= m.lcCode + 'Skipfor   = ' + SkipFor + CRLF
								lnMaxMatchStart	= Iif(m.llThisField, Len(m.lcCode), m.lnMaxMatchStart)
							Endif
	
							If Not Empty(Message)
								llThisField		= m.lcField = 'MESSAGE'
								lnMinMatchStart	= Iif(m.llThisField, Len(m.lcCode), m.lnMinMatchStart)
								lcType			= Iif(m.llThisField, Transform(Len(m.lcCode) + 13) + ' ' + Transform(Len(Message)), m.lcType)
								lcCode			= m.lcCode + 'Message   = ' + Message + CRLF
								lnMaxMatchStart	= Iif(m.llThisField, Len(m.lcCode), m.lnMaxMatchStart)
							Endif
	
							If Not Empty(ResName)
								llThisField		= m.lcField = 'RESNAME'
								lnMinMatchStart	= Iif(m.llThisField, Len(m.lcCode), m.lnMinMatchStart)
								lcType			= Iif(m.llThisField, Transform(Len(m.lcCode) + 13) + ' ' + Transform(Len(ResName)), m.lcType)
								lcCode			= m.lcCode + 'Resource  = ' + ResName + CRLF
								lnMaxMatchStart	= Iif(m.llThisField, Len(m.lcCode), m.lnMaxMatchStart)
							Endif
	
							If Not Empty(Comment)
								llThisField		= m.lcField = 'COMMENT'
								lnMinMatchStart	= Iif(m.llThisField, Len(m.lcCode), m.lnMinMatchStart)
								lcType			= Iif(m.llThisField, Transform(Len(m.lcCode) + 13) + ' ' + Transform(Len(Comment)), m.lcType)
								lcCode			= m.lcCode + 'Comment   = ' + Comment + CRLF
								lnMaxMatchStart	= Iif(m.llThisField, Len(m.lcCode), m.lnMaxMatchStart)
							Endif
	
							If Not Empty(Setup)
								lcCode			= m.lcCode + CRLF + Replicate('*', 60) + CRLF + CRLF
								llThisField		= m.lcField = 'SETUP'
								lnMinMatchStart	= Iif(m.llThisField, Len(m.lcCode), m.lnMinMatchStart)
								lcType			= Iif(m.llThisField, Transform(Len(m.lcCode) + 1) + ' ' + Transform(Len(Setup)), m.lcType)
								lcCode			= m.lcCode + Setup + CRLF
								lnMaxMatchStart	= Iif(m.llThisField, Len(m.lcCode), m.lnMaxMatchStart)
							Endif
	
							If Not Empty(Cleanup)
								lcCode			= m.lcCode + CRLF + Replicate('*', 60) + CRLF + CRLF
								llThisField		= m.lcField = 'CLEANUP'
								lnMinMatchStart	= Iif(m.llThisField, Len(m.lcCode), m.lnMinMatchStart)
								lcType			= Iif(m.llThisField, Transform(Len(m.lcCode) + 1) + ' ' + Transform(Len(Cleanup)), m.lcType)
								lcCode			= m.lcCode + Cleanup + CRLF
								lnMaxMatchStart	= Iif(m.llThisField, Len(m.lcCode), m.lnMaxMatchStart)
							Endif
	
							If Not Empty(Procedure)
								lcCode			= m.lcCode + CRLF + Replicate('*', 60) + CRLF + CRLF
								llThisField		= m.lcField = 'PROCEDURE'
								lnMinMatchStart	= Iif(m.llThisField, Len(m.lcCode), m.lnMinMatchStart)
								lcType			= Iif(m.llThisField, Transform(Len(m.lcCode) + 1) + ' ' + Transform(Len(Procedure)), m.lcType)
								lcCode			= m.lcCode + Procedure + CRLF
								lnMaxMatchStart	= Iif(m.llThisField, Len(m.lcCode), m.lnMaxMatchStart)
							Endif
		
						Case m.lcExt = 'DBC'
							._Name	= Alltrim(ObjectName)
							._Class	= Alltrim(ObjectType)
	
							Do Case
	
								Case ._Class = 'Database' And m.lcField = 'OBJECTNAME'
									*lcCode = '' && Will cause this match to be skipped. Don't want to record these matches.
	
								Case ._Class = 'Table'
									*lcCode = ._Class + '.dbf' && The name of the Table attached to the DBC
									lcCode = This.CleanUpBinaryString(m.lcCode)  && The SQL statement that makes up the View
	
								Case ._Class = 'View'
									lnStart	= Atc('Select', m.lcCode)
									lcCode	= Substr(m.lcCode, m.lnStart)
									lcCode	= This.CleanUpBinaryString(m.lcCode, .T.)  && The SQL statement that makes up the View
	
								Case ._Class = 'Field' && Fields can be part of Tables or Views
									*-- Get some info about the parent of this field
									lnParentID = ParentID
									Select  ObjectType,						;
											ObjectName						;
										From (m.tcFile)						;
										Where ObjectID = m.lnParentID		;
										Into Array laParent
									lcParentName = Alltrim(m.laParent[2])
	
									*-- Parse the field into a field name and field source
									lnStart	= Atc('#', m.lcCode)
									lcCode	= Substr(m.lcCode, m.lnStart + 1)
									lcCode	= This.CleanUpBinaryString(m.lcCode)
	
									lcFieldSource = Alltrim(Getwordnum(m.lcCode, 1))
									lcDataType	  = Substr(Alltrim(Getwordnum(m.lcCode, 2)), 2)
	
									If m.lcFieldSource = '0'
										lcFieldSource = '[Table alias in query]'
										lcDataType	  = ''
									Endif
	
									If Not Empty(m.lcDataType)
										lcCode	   = m.lcParentName + ' references ' + m.lcFieldSource + ' (data type: ' + m.lcDataType + ')'
										.MatchType = 'Field Source'
									Else
										lcCode	   = m.lcParentName + '.' + m.lcFieldSource
										.MatchType = Alltrim(m.laParent[1]) + ' Field'
									Endif
	
									._Class = .MatchType
	
							Endcase
					Endcase
	
					*-- Here is where we can skip the processing of certain records that we want to ignore, even though we found a match in them.
					If (m.lcExt = 'VCX' And Empty(m.lcClass)) Or							;					 	&& This is the ending row of a Class def in a vcx. Need to skip over it.
						(m.lcExt = 'FRX' And m.lcField = 'TAG2' And Recno() = 1) Or			;	&& Tag2 on first record in a FRX is binary and I want to skip it.
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
						lnMatchCount = This.SearchInCode(m.lcCode, m.loFileResultObject, m.llHasMethods, m.lnMinMatchStart, m.lnMaxMatchStart, m.lcType)
					Else
						* Can't search since there is no cSearchExpression, so we just log the file as a result.
						* This handles TimeStamp searches, where the cSearchExpression is empty
						loSearchResultObject		   = Createobject('GF_SearchResult')
						loSearchResultObject.Code	   = Iif(Type('properties') # 'U', Properties, '')
						loSearchResultObject.Code	   = m.loSearchResultObject.Code + CR + Iif(Type('methods') # 'U', Methods, '')
						loSearchResultObject.UserField = m.loFileResultObject
	
						If m.lcExt = 'FRX'
							loSearchResultObject.MatchLine		  = Expr
							loSearchResultObject.TrimmedMatchLine = Expr
						Endif
	
						This.ProcessSearchResult(m.loSearchResultObject)
	
						ii			 = 1000 && To end the outer for loop when the Do loop ends
						lnMatchCount = m.lnMatchCount + 1
					Endif
				Endif
	
				If m.lnMatchCount < 0 && There was an error in above call, need to exit
					Exit
				Else
					lnTotalMatches = m.lnTotalMatches + m.lnMatchCount
				Endif
		
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


* ================================================================================
	Procedure SelectControl(loTools, lcName, lcProperty)
		Local loDefaultButton, loException, loTopMostParent
	
		Try
			loTopMostParent	= m.loTools.oUtils.FindTopMostParent()
			loDefaultButton	= This.GetDefaultButton(m.loTopMostParent) && is there a button with Default = .T.
	
			m.loTools.SelectObject(m.lcName, m.lcProperty, .T.)
	
			* did the Default button get changed?  .. if so, change it back
			If Vartype(m.loDefaultButton) = 'O' And m.loDefaultButton.Default = .F.
				loDefaultButton.Default = .T.
			Endif
		Catch To m.loException
	
		Endtry
	
	Endproc
				

*----------------------------------------------------------------------------------
*-- Read file patterns to include in the search
* SF 20230619
	Procedure SetIncludePattern()
	
		Local laPattern[1], lcPattern, lnPattern, lnPatterns, loRegExp
	
		This.oSearchOptions.cFileTemplate  = Alltrim(This.oSearchOptions.cFileTemplate)
		This.oSearchOptions.cOtherIncludes = Alltrim(This.oSearchOptions.cOtherIncludes)
	
		If Not Empty(This.oSearchOptions.cFileTemplate) Then
			If Isnull(This.oSearchOptions.oRegExpFileTemplate) Then
				loRegExp			= GF_GetRegExp()
				loRegExp.IgnoreCase	= .T.
				loRegExp.MultiLine	= .T.
				*				loRegExp.ReturnFoxObjects = .T.
				*				loRegExp.AutoExpandGroups  = .T.
				*				loRegExp.AutoExpandGroups  = .T.
				*				loRegExp.Singleline						= .T.
				This.oSearchOptions.oRegExpFileTemplate	= m.loRegExp
	
			Else &&ISNULL(This.oSearchOptions.oRegExpFileTemplate)
				loRegExp = This.oSearchOptions.oRegExpFileTemplate
	
			Endif &&ISNULL(This.oSearchOptions.oRegExpFileTemplate)
	
			lnPatterns = Alines(laPattern, This.oSearchOptions.cFileTemplate, 1, ',', ';')
			If m.lnPatterns = 1 Then
				lcPattern = This.EscapePattern(m.loRegExp, m.laPattern(1))
			Else  &&m.lnPatterns = 1
				lcPattern = ''
	
				For lnPattern = 1 To m.lnPatterns
					If Justext(m.laPattern(m.lnPattern)) == '' Then
						If Juststem(m.laPattern(m.lnPattern)) == '' Then
							Loop
						Endif &&JUSTSTEM(laPattern(m.lnPattern))==""
						*filename without extension
						lcPattern = m.lcPattern + '|(' + This.EscapePattern(m.loRegExp, Justext(m.laPattern(m.lnPattern))) + '(?<!\..+)$)'
					Else  &&JUSTEXT(laPattern(m.lnPattern))==""
						lcPattern = m.lcPattern + '|(' + This.EscapePattern(m.loRegExp, m.laPattern(m.lnPattern)) + ')'
					Endif &&JUSTEXT(laPattern(m.lnPattern))==""
				Endfor &&lnPattern
				lcPattern = Substr(m.lcPattern, 2)
			Endif &&m.lnPatterns
			loRegExp.Pattern = m.lcPattern
	
		Endif &&!Empty(This.oSearchOptions.cFileTemplate)
	Endproc
	
*----------------------------------------------------------------------------------
*-- Escape a file pattern
	Procedure EscapePattern()
		Lparameters			;
			toRegExp,		;
			tcPattern
	
		Local lcPattern
	
		lcPattern = Iif(Left(m.tcPattern, 1) = '.', '*', '')		;
			+ m.tcPattern											;
			+ Iif(Right(m.tcPattern, 1) = '.', '*', '')				;
			+ Iif('.' $ m.tcPattern, '', '.*')
		
		Return '^' + This.EscapeSearchExpression(m.lcPattern, 1) + '$'
		
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

		For lnI = 1 To Alines(laLines, m.lcFilesToSkip + CR + '_command.prg', 5)
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
	
		Local lcErrorMessage, lcFile
	
		lcFile     = Alltrim(Evl(m.tcFile, 'None'))
	
		lcErrorMessage = m.tcErrorMessage + CR														;
			+ Space(4) + '[File: ' + m.lcFile + ']'													;
			+ Iif(Empty(m.tnResultId), '',  Space(4) + '[Record Id: ' + Transform(m.tnResultId) + ']') ;
			+ CR
	
		This.ShowError(m.lcErrorMessage, m.tnDialogBoxType, m.tcTitle)
	
		This.oReplaceErrors.Add(m.lcErrorMessage)
	
	Endproc
			

*----------------------------------------------------------------------------------
	Procedure SetSearchError(tcErrorMessage, tnDialogBoxType, tcTitle)

		This.oSearchErrors.Add(m.tcErrorMessage)

	Endproc


*----------------------------------------------------------------------------------

	Proc SetSearchedExtensions
		This.oSearchOptions.cSearchExtensions = ''
	
		This.IncludeExtension(This.oSearchOptions.lIncludeSCX, 'SCX')
		This.IncludeExtension(This.oSearchOptions.lIncludeVCX, 'VCX' )
		This.IncludeExtension(This.oSearchOptions.lIncludeFRX, 'FRX' )
		This.IncludeExtension(This.oSearchOptions.lIncludeDBC, 'DBC' )
		This.IncludeExtension(This.oSearchOptions.lIncludeMNX, 'MNX' )
		This.IncludeExtension(This.oSearchOptions.lIncludeLBX, 'LBX' )
		This.IncludeExtension(This.oSearchOptions.lIncludePJX, 'PJX' )
		This.IncludeExtension(This.oSearchOptions.lIncludePRG, 'PRG' )
		This.IncludeExtension(This.oSearchOptions.lIncludeMPR, 'MPR' )
		This.IncludeExtension(This.oSearchOptions.lIncludeTXT, 'TXT' )
		This.IncludeExtension(This.oSearchOptions.lIncludeINI, 'INI' )
		This.IncludeExtension(This.oSearchOptions.lIncludeH,   'H'   )
		This.IncludeExtension(This.oSearchOptions.lIncludeXML, 'XML' )
		This.IncludeExtension(This.oSearchOptions.lIncludeSPR, 'SPR' )
		This.IncludeExtension(This.oSearchOptions.lIncludeASP, 'ASP' )
		This.IncludeExtension(This.oSearchOptions.lIncludeJSP, 'JSP' )
		This.IncludeExtension(This.oSearchOptions.lIncludeJAVA, 'JAVA')
		This.IncludeExtension(This.oSearchOptions.lIncludeHTML, 'HTM HTML')
		This.IncludeExtension(.T., Upper(Chrtran(This.oSearchOptions.cOtherIncludes, ',;', '  ')))
	
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
			*	Wait Window At 5, Wcols() / 2 Nowait m.tcMessage
			If Vartype(This.oProgressBar) = 'O'
				This.oProgressBar.UpdateWait(m.tcMessage)
			Endif
	
		Endif
	
	Endproc
		

*----------------------------------------------------------------------------------
	Procedure StartProgressBar(tnMaxValue)
	
		This.nFilesSearched = m.tnMaxValue
	
		If Vartype(This.oProgressBar) = 'O'
			This.oProgressBar.Start(m.tnMaxValue)
		Endif
	
	Endproc
		

*----------------------------------------------------------------------------------
	Procedure StartTimer

		This.nSearchStartTime = Seconds()

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

		lcTrimmedString = Alltrim(m.tcString, 1, Chr(32), Chr(9), LF, CR, Chr(0))
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

	EndProc
	
	Procedure CanUseGrepForFileList
		*** JRN 2024-05-31 : With V7.1, use grep to pre-filter files to be searched
		* However, cannot do so if there is a file template (rare)
		* or " in the search expression (calls to grep use ")
		* or any characters before chr[32] or after chr[126]
		* or has a trailing backslash
		* or, for some users, wildcard * along with ( or )
		
		Local lcChar, lcEscapedSearchExpression, lcSearchExpression, lnI
	
		If Not This.oSearchOptions.lOptimizeWithGrep
			Return .F.
		Endif
	
		If Not Empty(This.oSearchOptions.cFileTemplate)
			Return .F.
		Endif
	
		lcSearchExpression = This.oSearchOptions.cSearchExpression
		If ["] $ lcSearchExpression
			Return .F.
		Endif
	
		For lnI = 1 To Len(lcSearchExpression)
			lcChar = Substr(lcSearchExpression, lnI)
			If Not Between(Asc(lcChar), 32, 126)
				Return .F.
			Endif
		Endfor
	
		If Right(lcSearchExpression, 1) = '\'
			Return .F.
		Endif
	
		If This.oSearchOptions.lUnmatchedParenError
			*** JRN 2024-11-07 : A rare case for some users: cannot use grep if wild card * along with either of ( or )
			lcEscapedSearchExpression = This.oSearchOptions.cEscapedSearchExpression
			* ignore leading and trailing .*
			lcEscapedSearchExpression = Substr(m.lcEscapedSearchExpression, 3, Len(m.lcEscapedSearchExpression) - 4)
		
			If '.*' $ m.lcEscapedSearchExpression And ('\(' $ m.lcEscapedSearchExpression Or '\)' $ m.lcEscapedSearchExpression)
				Return .F.
			Endif
		Endif
			
		Return .T.
	
	Endproc
			
* ================================================================================ 
* ================================================================================ 
*** JRN 2024-05-31 : Optimized section, using grep to create (shorter) list of files

	Procedure SearchUsingGrep (tcPath, ttTime, tcUni)
	
		Local lcCustomAlias, lnReturn, lnSeconds, lnSelect
	
		lnSelect  = Select()
	
		This.tRunTime = Evl(m.ttTime, Datetime())
		This.cUni	  = Evl(m.tcUni, '_' + Sys(2007, Ttoc(This.tRunTime), 0, 1))
	
		This.oProgressBar.Start(100, 'Creating list of files (using grep.exe)')
	
		lcCustomAlias = 'GF_CustomScope' + Sys(2015)
		Create Cursor (m.lcCustomAlias) (FileName C(240))
	
		This.CreateFileListUsingGrep(This.oSearchOptions, m.lcCustomAlias, m.tcPath)
	
		This.nADirTime = This.ElapsedTimeSinceStart()
	
		Select Distinct Lower(FileName) As FileName From (m.lcCustomAlias) Into Cursor (m.lcCustomAlias)
		This.StartProgressBar(Reccount())
	
		Scan
	
			lnReturn = This.SearchInFile(Alltrim(FileName))
	
			This.UpdateProgressBar(This.nFilesProcessed)
	
			If (m.lnReturn < 0) Or This.lEscPress Or This.nMatchLines >= This.oSearchOptions.nMaxResults
				Exit
			Endif
	
		Endscan
	
		This.SearchFinished(m.lnSelect)
	
		Use In (m.lcCustomAlias)
	
		If Evl(m.lnReturn, 0) >= 0
			Return 1
		Else
			Return m.lnReturn
		Endif
	
	Endproc
			

* ================================================================================

	Procedure CreateFileListUsingGrep
		Lparameters toSettings, tcCursor, tcPath
	
		*** JRN 2024-05-22 : Concept and coding suggestions from Mike Yearwood
		* C:\mygrep\grep.exe -r -l -i -P --include=.{??a,prg} '^(?!\s*).*THERE\s+IS\s+NO\s+SUCH\s+APPRO' c:\mysource
		*  -r recurse folders, -l list file name only, -i case insensitive, -P use Perl regex
	
		Local lcBATFile, lcCommand, lcDoneFile, lcErrFile, lcExtensions, lcFolder, lcOutFile, lcScope
		Local lcSearchExpression, lcSearchPattern, lcStem, lcSubDirs, lcgrep, lnSearchMode
	
		* Search Mode: 1 = Plain, 2 = WildCard, 3 = Regex
		lnSearchMode = m.toSettings.nSearchMode
	
		* Search Expression (as entered)
		lcSearchExpression = m.toSettings.cSearchExpression
		Do Case
			Case m.lnSearchMode = 3
				lcSearchPattern	   = '-P'
			Case m.lnSearchMode = 2 And ('*' $ m.lcSearchExpression Or '?' $ m.lcSearchExpression)
				lcSearchExpression = m.toSettings.cEscapedSearchExpression
				*** JRN 2024-06-09 : strip off leading and trailing .*
				lcSearchExpression = Substr(m.lcSearchExpression, 3, Len(m.lcSearchExpression ) - 4)
				lcSearchPattern	   = '-G'
			Otherwise
				*** JRN 2024-06-09 : oddly enough, if a trailing '\', needs another
				lcSearchExpression = m.lcSearchExpression + Iif(Right(m.lcSearchExpression, 1) = '\', '\', '')
				lcSearchPattern	   = '-F'
		Endcase
	
		* Extensions searched
		lcExtensions = This.GetExtensions(m.toSettings.cSearchExtensions)
	
		* ================================================================================  
	
		lcStem	   = Addbs(Sys(2023)) + 'GF_Files' + Sys(2015)
		lcBATFile  = m.lcStem + '.bat'
		lcOutFile  = m.lcStem + '.txt'
		lcErrFile  = m.lcStem + '.err'
		lcDoneFile = m.lcStem + '.done'
	
		Strtofile('@echo off' + CRLF, m.lcBATFile, 1)
	
		* Scope (project or path) 
		lcScope	 = m.tcPath && m.toSettings.cRecentScope
	
		If Directory(m.lcScope)

			* ================================================================================ 
			lcFolder = Addbs(m.lcScope)
	
			If ':' $ m.lcScope
				Strtofile(Left(m.lcScope, 2)  + CRLF, m.lcBATFile, 1)
			Endif
			Strtofile('cd "' + m.lcScope + '"' + CRLF, m.lcBATFile, 1)
	
			*** JRN 2024-05-14 : file contents
			lcgrep	= ["] + Addbs(_Screen._GoFish.cAppPath) + 'grep\grep.exe' + ["]
	
			lcScope = ["] + m.lcScope + ["]
	
			If m.toSettings.lIncludeSubDirectories
				lcCommand = m.lcgrep + [ -r -l -i -z ] + m.lcSearchPattern
			Else
				lcCommand = m.lcgrep + [    -l -i -z ] + m.lcSearchPattern
				lcScope	  = m.lcScope + [\*.*]
			Endif
			lcScope	= Chrtran(m.lcScope, '\', '/')
	
			lcCommand = m.lcCommand + Textmerge([ --include=*.{<<m.lcExtensions>>} -- "<<m.lcSearchExpression>>" <<m.lcScope>>])
	
			lcCommand = m.lcCommand + [ 1> "] + m.lcOutFile + [" 2>> "] + m.lcErrFile + ["]
			Strtofile(m.lcCommand + CRLF, m.lcBATFile, 1)
	
			* ================================================================================
			*** JRN 2024-05-30 : search for file names
			*   DIR *blank*.* /s /b  gives a simple file listing
			lcSubDirs = Iif(m.toSettings.lIncludeSubDirectories, '/s', '')
			If m.lnSearchMode # 3 And Not '\' $ m.toSettings.cSearchExpression
				lcCommand = Textmerge([DIR *"<<m.toSettings.cSearchExpression>>"*.* <<lcSubDirs>> /b 1>> "<<m.lcOutFile>>" 2>> "<<m.lcErrFile>>"])
			Endif
			Strtofile(m.lcCommand + CRLF, m.lcBATFile, 1)

		Else
	
			* ================================================================================
			
		Endif
	
		Strtofile('Echo Done >> "' + m.lcDoneFile  + '"' + CRLF, m.lcBATFile, 1)
	
		This.CallShell(m.lcBATFile, m.lcDoneFile)
	
		* ================================================================================
		
		This.CheckForUnmatchedParenError(m.lcErrFile)
	
		This.AppendtoCursor(m.tcCursor, m.lcOutFile, m.lcErrFile)
	
		This.CleanUpGrepFiles(m.lcStem)
	
		Return
	
	Endproc
		

	* ===============================================================================
	Procedure GetExtensions(lcSearchExtensions)
		Local lcExt, lcExtensions, lnI

		lcExtensions = ''
		For lnI = 1 To Getwordcount(m.lcSearchExtensions)
			lcExt = Getwordnum(m.lcSearchExtensions, m.lnI)
			If ' ' + m.lcExt + ' ' $ ccBinaries
				lcExt = Left(m.lcExt, 2) + 'T'
			Endif
			lcExtensions = m.lcExtensions + Iif(m.lnI = 1, '', ',') + GF_GetMixedCaseCombination(m.lcExt)
		Endfor

		Return m.lcExtensions

	EndProc
	

	* ================================================================================
	Procedure CallShell(tcBATFile, tcDoneFile)
		*!* ******** JRN Removed 2024-06-07 ********
		*!* * Shell, no command window, wait for result
		*!* Local oShell As 'wscript.shell'
	
		*!* oShell = Createobject ('wscript.shell')
		*!* m.oShell.Run (m.tcBATFile, 7, .T.)
		Local lnElapsed, lnSeconds
	
		GF_Shell(m.tcBATFile)
		This.oProgressBar.nMaxValue	= 10
		lnSeconds					= Seconds()
		Do While Not File(m.tcDoneFile)
			lnElapsed = Seconds() - m.lnSeconds
			* recalibrate during really big searches
			If m.lnElapsed > This.oProgressBar.nMaxValue
				This.oProgressBar.nMaxValue = 2 * This.oProgressBar.nMaxValue
			Endif
			This.UpdateProgressBar(m.lnElapsed)
			Sleep (250) && quarter second
			
			If Inkey() = 27
				This.lEscPress = .T.
				Clear Typeahead
				Return 
			Endif

		EndDo
	Endproc
		

	* ================================================================================
	Procedure AppendtoCursor(tcCursor, tcOutFile, tcErrFile)
		Local lcTexts

		Select (m.tcCursor)
		Append From (m.tcOutFile) Sdf
		If not Empty(m.tcErrFile)
			This.AppendErrorFiles(m.tcErrFile)
		Endif

		Replace All FileName With Chrtran(FileName, '/', '\')

		* All those T extensions that should be 'X' 
		lcTexts	 = Chrtran(ccBinaries, 'X', 'T')
		Replace	FileName  With	Strtran(FileName, 't    ', 'x', 1, 1, 1)		;
			For ' ' + Upper(Justext(FileName)) + ' ' $ m.lcTexts

	Endproc


	* ================================================================================ 
	Procedure CheckForUnmatchedParenError(lcErrFile)
		If This.oSearchOptions.nSearchMode = 2		;
				And File(m.lcErrFile)				;
				And Atc('grep: unmatched', Filetostr(m.lcErrFile)) # 0
			This.StopProgressBar()
			This.oSearchOptions.lUnmatchedParenError = .T.
			MessageBox('GoFish: Internal error (grep.exe)' + CR + CR + 'Please try again', 16, 'GoFish - but try again')
		Endif
	Endproc
	
	* ================================================================================
	Procedure AppendErrorFiles(lcErrFile)
		Local laErrors[1], lcText, lnCount, lnI

		lnCount = Alines(laErrors, Filetostr(m.lcErrFile))
		For lnI = 1 To m.lnCount
			lcText = m.laErrors[m.lnI]
			lcText = Substr(m.lcText, At(':', m.lcText) + 1)
			lcText = Left(m.lcText, Rat(':', m.lcText) - 1)
			Insert Into (Alias()) Values (Trim(m.lcText))
		Endfor
	Endproc


	* ================================================================================
	Procedure CleanUpGrepFiles(tcStem)
		If This.oSearchOptions.lCleanUpGrepFiles
			Erase (m.tcStem + '.*')
		EndIf 
	EndProc

	* ================================================================================
	*!* ******** JRN Removed 2024-07-08 ********
	*!* Procedure PreprocessFileList(tcFilesCursor, tcDestCursor)

	*!* 	Local lcExtensions, lcTexts

	*!* 	*** JRN 2024-07-03 : list of all files, only needed extensions, and 't' files for binaries
	*!* 	lcExtensions = Lower(' ' + This.oSearchOptions.cSearchExtensions + ' ')
	*!* 	Select  Distinct Lower(FileName)    As  FileName						;
	*!* 		From (m.tcFilesCursor)												;
	*!* 		Where ' ' + Lower(Justext(FileName)) + ' ' $ m.lcExtensions			;
	*!* 		Into Cursor FilesCursor Readwrite

	*!* 	Replace	FileName  With	Strtran(FileName, 'x    ', 't', 1, 1, 1)		;
	*!* 		For ' ' + Upper(Justext(FileName)) + ' ' $ ccBinaries

	*!* 	*** JRN 2024-07-03 : create txt file from this cursor
	*!* 	lcText = FilesToCursor('FileName')

	*!* 	Select  *											;
	*!* 		From FilesCursor								;
	*!* 		Where This.AnyRegExMatchesInFile(FileName)		;
	*!* 		Into Cursor (m.tcDestCursor) readwrite

	*!* 	lcTexts	 = Chrtran(ccBinaries, 'X', 'T')
	*!* 	Replace	FileName  With	Strtran(FileName, 't    ', 'x', 1, 1, 1)		;
	*!* 		For ' ' + Upper(Justext(FileName)) + ' ' $ m.lcTexts

	*!* 	*RegExpFileList input.txt THERE\s+IS\s+ output.txt

	*!* Endproc


 	* ================================================================================ 
 	Procedure FilesCursorToText
 		Lparameters lcAlias
 	
 		ccFileDelimiter = CRLF
 		Local lcClipText, lcFiles, lnI
 	
 		lcClipText = _Cliptext
 		_vfp.DataToClip(m.lcAlias)
 		* strip off first line, which contains the field name
 		lcFiles	  = Substr(_Cliptext, 1 + At(CR, _Cliptext))
 		_Cliptext = m.lcClipText

 		* fix the delimiter to what xargs needs
 		lcFiles = Strtran(m.lcFiles, CR, ccFileDelimiter)
 	
 		* alltrim on all records
 		For lnI = 7 To 0 Step - 1
 			lcFiles = Strtran(m.lcFiles, Replicate(' ', 2 ** m.lnI) + ccFileDelimiter, ccFileDelimiter)
 		Endfor
 	
 		* and change .??x to .??t for grep
 		lcFiles = Strtran(m.lcFiles, 'x' + ccFileDelimiter, 't' + ccFileDelimiter)
 	
 		Return lcFiles
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
Function _EdStoPos
*!*	/Changed by: nmpetkov 27.3.2023


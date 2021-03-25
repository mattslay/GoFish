#Include GoFish.h


Define Class GoFishSearchEngine As Custom

	cBackupPRG                       = 'GoFishBackup.prg'
	cFilesToSkip                     = ''

	* This string contains a list of files to be skipped during the search.
	* One filname on each line. This list is only skipped if lSkipFiles is .t.
	cFilesToSkipFile                 = 'c:\users\matt\appdata\roaming\microsoft\visual foxpro 9\GF_Files_To_Skip.txt'

	cGraphicsExtensions              = 'PNG ICO JPG JPEG TIF TIFF GIF BMP MSK CUR ANI'
	cInitialDefaultDir               = ''

	* A text list of projects that matches oProjects. Makes looking for existing projects fast than analyzing
	* the oProjects collection. This property is only to be used by the class. Please don't touch it.
	cProjects                        = ''

	* A detail record is stored here for every single match line that is replaced.
	cReplaceDetailTable              = 'c:\users\matt\appdata\roaming\microsoft\visual foxpro 9\GF_Replace_DetailV5.dbf'

	* A single header record is stored here for each time the ReplaceFromMarkedRows() method is called.
	cReplaceHistoryTable             = 'c:\users\matt\appdata\roaming\microsoft\visual foxpro 9\GF_Replace_History.dbf'

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

	oReplaceErrors                   = .Null.

	* This is a collection of match objects from the last search. Must set  lCreateResultsCollection if you want this collection to be built.
	oResults                         = .Null.

	* A collection of any errors that happened during the last search.
	oSearchErrors                    = .Null.

	* An object instance of the Search Options class that holds properties  to controll how the search is performed.
	oSearchOptions                   = .Null.
	Dimension aMenuStartPositions[1]
	Dimension aWildcardFiles[1]

	*----------------------------------------------------------------------------------
	Procedure AddColumn(tcTable, tcColumnName, tcColumnDetails)

		Local lcAlias

		lcAlias = Juststem(tcTable)
		Try
			Alter Table (tcTable) Add Column &tcColumnName &tcColumnDetails
		Catch
		Endtry
		
	EndProc


	*----------------------------------------------------------------------------------
	Procedure AddFieldToReplaceTable(lcTable, lcCsr, lcFieldName, lcDataType)

		Local llSuccess
		
		If Empty(Field(m.lcFieldName, lcCsr))
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
		
	EndProc


	*----------------------------------------------------------------------------------
	Procedure AddProject(tcProject)

		Local llAlreadyInCollection

		llAlreadyInCollection = Atline(Upper(tcProject), Upper(This.cProjects)) <> 0

		If !llAlreadyInCollection
			This.oProjects.Add(Lower(tcProject))
			This.cProjects = This.cProjects + tcProject + Chr(13)
		Endif
		
	EndProc


	*----------------------------------------------------------------------------------
	Procedure AssignMatchType(toObject)

		Local lcFileType, lcMatchType, lcName, lcTrimmedMatchLine, lcValue, loNameMatches, loValueMatches
		Local llError, llWorkingOnClassFromVCX
		Local llNameHasDot, loLineMatches

		lcFileType = Upper(toObject.UserField.FileType)

		lcTrimmedMatchLine = This.TrimWhiteSpace(toObject.MatchLine)&& Trimmed version for display in the grid
		toObject.TrimmedMatchLine = lcTrimmedMatchLine

		*-- We read MatchType of UserField, but from here on, until the result row is created, we will
		*-- move this value to toObject.MatchType, and do some tweaking on it to make it the right value.
		*-- We'll never change the value that was passed in on toObject.UserField.MatchType
		lcMatchType = toObject.UserField.MatchType
		toObject.MatchType = lcMatchType
		*=============================================================================================
		* This area contains a few overrides that I've had to build in to make final tweeks on columns
		*=============================================================================================
		*-- Sometimes in a VCX/SCX the MethodName will be empty and MatchLine will contain the PROCEDURE name
		If Empty(toObject.MethodName) And Upper(Getwordnum(lcTrimmedMatchLine, 1)) = 'PROCEDURE'
			toObject.MethodName = Getwordnum(lcTrimmedMatchLine, 2)
		Endif

		If !Empty(toObject.MethodName)
			With toObject.UserField
				If '.' $ toObject.MethodName
					._Name = Alltrim(._Name + '.' + This.ExtractObjectName(toObject.MethodName), 1, '.')
					toObject.MethodName = Justext(toObject.MethodName)

					If ._ParentClass <> ._BaseClass
						._ParentClass = ''
						._BaseClass = ''
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
			With toObject.UserField
				If !Empty(._Class) && Trim Class name off of the front (only affects VCX results)
					._Name = Strtran(._Name, ._Class + '.', '', 1, 1)
				Endif
			Endwith
		Endif

		If Empty(toObject.UserField.ClassLoc)
			toObject.UserField._ParentClass = '' && Affects VCXs. PRGs will be address in these next lines
		Endif

		If lcFileType = 'PRG'
			With toObject.UserField
				.ContainingClass = ''
				._Class = toObject.oProcedure._ClassName
				._ParentClass = toObject.oProcedure._ParentClass
				._BaseClass = toObject.oProcedure._BaseClass
				If ._Name = ._Class
					._Name = ''
				Endif
				If  Upper('ENDDEFINE') $ Upper(toObject.MethodName)
					toObject.MethodName = ''
					._Name = ''
				Endif
			Endwith

		Endif

		If Upper(lcMatchType) # 'RESERVED3' And This.IsFullLineComment(lcTrimmedMatchLine)
			toObject.MatchType = MATCHTYPE_COMMENT
			This.CreateResult(toObject)
			Return .Null. && Exit out, we're done with this record!
		Endif

		*=============================================================================================
		* Handle a few tweaks on MatchType assignments
		*=============================================================================================
		This.ProcessInlineComments(toObject)

		Do Case

			*-- A TimeStamp only search, with no search expression...
			Case Isnull(toObject.oMatch)
				If This.oSearchOptions.lTimeStamp And Empty(This.oSearchOptions.cSearchExpression)
					If Empty(toObject.UserField._Name) And Empty(toObject.UserField.ContainingClass) And Empty(toObject.UserField._Class)
						toObject.MatchType =  MATCHTYPE_FILEDATE
					Else
						toObject.MatchType =  MATCHTYPE_TIMESTAMP
					Endif
				Else
					toObject.MatchType = MATCHTYPE_FILENAME
				Endif

			Case Inlist(lcFileType, 'SCX', 'VCX', 'FRX')&& And lcMatchType # MATCHTYPE_FILENAME
				This.AssignMatchTypeForScxVcx(toObject)

			Case lcFileType = 'PRG'
				This.AssignMatchTypeForPrg(toObject)

		Endcase

		*-- Read MatchType back off toObject for a final bit of tweaking...
		lcMatchType = toObject.MatchType

		Do Case
			Case Empty(lcMatchType)
				lcMatchType = MATCHTYPE_CODE

			Case Upper(Getwordnum(lcTrimmedMatchLine, 1)) = '#DEFINE'
				lcMatchType = MATCHTYPE_CONSTANT

			Case lcMatchType = MATCHTYPE_PROPERTY_DESC Or lcMatchType = MATCHTYPE_PROPERTY_DEF
				toObject.UserField.ContainingClass = ''
				toObject.UserField._Name = ''
				toObject.MethodName = Getwordnum(toObject.MatchLine, 1, ' ')

			Case lcMatchType = MATCHTYPE_PROPERTY

				If Atc('=', lcTrimmedMatchLine) = 0
					toObject.MatchType = MATCHTYPE_CODE
					Return toObject
				Endif

				lcName = Getwordnum(lcTrimmedMatchLine, 1, ' =') && The Property Name only
				toObject.MethodName = lcName

				Try
					If Atc('.', lcName) > 0 && Could be ObjectName.ObjectName.ObjectName.PropertyName
						lcName = Justext(lcName) && Need to pick off just the property name, and make sure that's where the match is.
						llNameHasDot = .T.
					Else
						llNameHasDot = .F.
					Endif

					*	toObject.UserField.MethodName = lcName
					lcName = lcName + ' =' && Need to construct property name like this example:   Caption =

					lcValue = Alltrim(Substr(lcTrimmedMatchLine, 1 + At('=', lcTrimmedMatchLine))) && GetWordNum(lcTrimmedMatchLine, 2, '=')
					loNameMatches = This.oRegExForSearch.Execute(lcName)
					loValueMatches = This.oRegExForSearch.Execute(lcValue)
					* loLineMatches = This.oRegExForSearch.Execute(lcTrimmedMatchLine)
					loLineMatches = This.oRegExForSearch.Execute(lcName + lcValue)

					With toObject.UserField
						If llNameHasDot
							If ._ParentClass <> ._BaseClass
								._ParentClass = ''
								._BaseClass = ''
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
						Case loNameMatches.Count > 0 And loValueMatches.Count > 0 && If match on both sides, make an extra call here for the Name
							toObject.MatchType = MATCHTYPE_PROPERTY_NAME
							This.CreateResult(toObject)
							lcMatchType = MATCHTYPE_PROPERTY_VALUE
						Case loNameMatches.Count > 0 Or loValueMatches.Count > 0 && Only matched on one side
							If loValueMatches.Count > 0 And This.oSearchOptions.lIgnoreMemberData And Lower(lcName) = '_memberdata ='
								llError = .T. && so this is skipped
							Else
								lcMatchType = Iif(loNameMatches.Count > 0, MATCHTYPE_PROPERTY_NAME, MATCHTYPE_PROPERTY_VALUE)
							Endif
						Case loLineMatches.Count > 0 && Matched SOMEWHERE on the line. Can span " = " this way
						*-- No modification to matchtype required. Will record as MATCHTYPE_PROPERTY
						Case loNameMatches.Count = 0 And loValueMatches.Count = 0 && Possible that there is not match at all, so we record nothing
							llError = .T.
						Otherwise
							* lcMatchType = Iif(loNameMatches.count > 0, MATCHTYPE_PROPERTY_NAME, MATCHTYPE_PROPERTY_VALUE)
					Endcase
				Catch
					lcMatchType = MATCHTYPE_CODE && IF anything above failed, then just consider this a regular code match
				Endtry

		Endcase

		If llError = .T.
			Return .Null.
		Endif

		*-- Wrap MatchType in brackets (if not already set), and if it's not MATCHTYPE_CODE ...
		If lcMatchType # MATCHTYPE_CODE And Left(lcMatchType, 1) # '<'
			lcMatchType = '<' + lcMatchType + '>'
		Endif

		toObject.MatchType = lcMatchType

		Return toObject
		
	EndProc


	*----------------------------------------------------------------------------------
	Procedure AssignMatchTypeForPrg(toObject)

		Local lcMatchType, lcParams, lcProcedureType, lnMatchStart, lnProcedureStart, loMatch, loMatches
		Local loProcedure
		Local lcName, lcTrimmedMatchLine, loNameMatches, loParamMatches

		loProcedure = toObject.oProcedure
		loMatch = toObject.oMatch
		lcMatchType = Upper(toObject.MatchType)
		lnProcedureStart = loProcedure.StartByte
		lnMatchStart = Iif(Vartype(loMatch) = 'O', loMatch.FirstIndex, 0)
		lcTrimmedMatchLine = toObject.TrimmedMatchLine

		Do Case
		Case lcMatchType = 'CLASS' && Note, this case also handles Properties on a Class...

			lcFirstWord = Upper(Getwordnum(lcTrimmedMatchLine, 1))
			If lcFirstWord $ 'PROCEDURE'
				lcMatchType = MatchType_Procedure
			Else
				lcMatchType = Iif(lnMatchStart = lnProcedureStart, MATCHTYPE_CLASS_DEF, MATCHTYPE_PROPERTY)
			Endif
		*toObject.MethodName = ''

		Case Inlist(lcMatchType, 'METHOD', 'PROCEDURE', 'FUNCTION')

			*-- This test looks for matches in on the Procedure Name versus possible parameters:
			*-- Ex: PROCEDURE ProcessJob(lcJobNo). )
			If lnMatchStart = lnProcedureStart
				lcName = Getwordnum(lcTrimmedMatchLine, 1, '(')
				lcParams = Getwordnum(lcTrimmedMatchLine, 2, '(')

				loNameMatches = This.oRegExForSearch.Execute(lcName)
				loParamMatches = This.oRegExForSearch.Execute(lcParams)

				If loNameMatches.Count > 0 And loParamMatches.Count > 0 && If match on both sides, make an extra call here for the Name
					toObject.UserField.MatchType = '<' + Proper(lcMatchType) + '>'
					This.CreateResult(toObject)
					lcMatchType = MatchType_Code
				Else
					lcMatchType = Iif(loParamMatches.Count > 0, MatchType_Code, Proper(lcMatchType))
				Endif
			Else
				lcMatchType = MatchType_Code
			Endif

		Otherwise
			lcMatchType = toObject.MatchType && Restore it back

		Endcase

		toObject.MatchType = lcMatchType
		
	EndProc


	*----------------------------------------------------------------------------------
	Procedure AssignMatchTypeForScxVcx(toObject)

		Local lcClass, lcContainingClass, lcMatchType, lcMethodName, lcName, lcProcedureType, lcPropertyName
		Local lcTrimmedMatchLine, lnMatchStart, lnProcedureStart, loMatches

		lcMethodName = toObject.MethodName
		lcTrimmedMatchLine = toObject.TrimmedMatchLine

		lcProcedureType = toObject.oProcedure.Type
		lnProcedureStart = toObject.oProcedure.StartByte

		lnMatchStart = toObject.oMatch.FirstIndex
		lcMatchType = Upper(toObject.MatchType)

		With toObject.UserField
			lcClass = ._Class
			lcContainingClass = .ContainingClass
			lcName = ._Name
		Endwith

		Do Case

			Case Alltrim(lcClass) == Alltrim(lcTrimmedMatchLine) And !Empty(lcClass) And Empty(lcName)
				lcMatchType = MATCHTYPE_CLASS_DEF

			Case lcMatchType = 'RESERVED3'
				If Left(lcTrimmedMatchLine, 1) = '*' && A Method Definition line
					lcMethodName = Substr(lcTrimmedMatchLine, 2, Len(Getwordnum(lcTrimmedMatchLine, 1)) - 1)
					loMatches = This.oRegExForSearch.Execute(lcMethodName)
					lcMatchType = Iif(loMatches.Count > 0, MATCHTYPE_METHOD_DEF, MATCHTYPE_METHOD_DESC)
					toObject.MethodName = Iif(loMatches.Count > 0, lcMethodName, '')
				Else && A Property Definition line
					lcPropertyName = Getwordnum(lcTrimmedMatchLine, 1)
					If Atc('.', lcPropertyName) > 0
						lcPropertyName = Justext(lcPropertyName)
					Endif
					loMatches = This.oRegExForSearch.Execute(lcPropertyName)
					lcMatchType = Iif(loMatches.Count > 0, MATCHTYPE_PROPERTY_DEF, MATCHTYPE_PROPERTY_DESC)
				Endif

			Case lcMatchType = 'RESERVED7'
				lcMatchType = MATCHTYPE_CLASS_DESC

			Case lcMatchType = 'RESERVED8'
				lcMatchType = MATCHTYPE_INCLUDE_FILE

			Case lcMatchType = 'OBJNAME'
				lcMatchType = Iif(Empty(lcName), MatchType_Class, MatchType_Name)

			Case lcMatchType = 'PROCEDURE'
				If lnMatchStart = lnProcedureStart And !Empty(toObject.oProcedure.ParentClass)
					lcMatchType = MatchType_Method
				Else
					lcMatchType = MatchType_Code
				Endif

			Case lcMatchType = 'CLASS'
				lcMatchType = MatchType_Class

			Case lcMatchType = 'PROPERTIES'
				lcMatchType = MATCHTYPE_PROPERTY

			Otherwise
				lcMatchType = toObject.MatchType && Restore it back

		Endcase

		toObject.MatchType = lcMatchType
		
	EndProc


	*----------------------------------------------------------------------------------
	Procedure BackupFile(tcFilePath, tnReplaceHistoryId)

		#Define ccBACKUPFOLDER Addbs(Home(7) + 'GoFishBackups')

		Local lcBackupPRG, llCopyError
		Local laExtensions[1], lcDestFile, lcExt, lcExtensions, lcSourceFile, lcThisBackupFolder, lnI

		If This.oSearchOptions.lPreviewReplace = .T.
			Return
		Endif

		llCopyError = .F.

		*-- If the user has created a custom backup PRG, and placed it in their path, then call it instead
		lcBackupPRG = 'GoFish_Backup.prg'

		If File(lcBackupPRG)
			Do &lcBackupPRG With tcFilePath, tnReplaceHistoryId
			Return
		Endif

		If Not Directory (ccBACKUPFOLDER) && Create main folder for backups, if necessary
			Mkdir (ccBACKUPFOLDER)
		Endif

		* Create folder for this ReplaceHistorrID, if necessary
		lcThisBackupFolder = Addbs (ccBACKUPFOLDER + Transform (tnReplaceHistoryId))

		If Not Directory (lcThisBackupFolder)
			Mkdir (lcThisBackupFolder)
		Endif

		* Determine the extensions we need to consider
		lcExt = Upper (Justext (tcFilePath))

		Do Case
			Case lcExt = 'SCX'
				lcExtensions = 'SCX,SCT'
			Case lcExt = 'VCX'
				lcExtensions = 'VCX,VCT'
			Case lcExt = 'FRX'
				lcExtensions = 'FRX,FRT'
			Case lcExt = 'MNX'
				lcExtensions = 'MNX,MNT,MPR,MPX'
			Case lcExt = 'DBC'
				lcExtensions = 'DBC,DCT,DCX'
			Case lcExt = 'LBX'
				lcExtensions = 'LBX,LBT'
			Otherwise
				lcExtensions = lcExt
		Endcase

		*-- Copy each file into the destination folder, if its not already there
		Alines (laExtensions, lcExtensions, 0, ',')

		For lnI = 1 To Alen (laExtensions)
			lcSourceFile = Forceext (tcFilePath, laExtensions (lnI))
			lcDestFile	 = lcThisBackupFolder + Justfname (lcSourceFile)
			If Not File (lcDestFile)
				Try
					Copy File (lcSourceFile) To (lcDestFile)
				Catch
					If !llCopyError
						This.SetReplaceError('Error creating backup of file.', tcFilePath, tnReplaceHistoryId)
					Endif
					llCopyError = .T.
				Endtry
			Endif
		Endfor

		Return !llCopyError
		
	EndProc


	*----------------------------------------------------------------------------------
	Procedure BuildDirectoriesCollection(tcDir)

		*-- Note: This method is called recursively on itself if subfolders are found. See the For loop at the bottom...
		*-- For more good info on recursive processing of directories, see this page: http://fox.wikis.com/wc.dll?Wiki~RecursiveDirectoryProcessing

		Local laDirList[1], laFileList[1], lcCurrentDirectory, lcDriveAndDirectory, lnDirCount, lnPtr, lnFileCount

		*!* ** { JRN -- 07/11/2016 08:11 AM - Begin
		*!* If Lastkey() = 27 or Inkey() = 27
		If Inkey() = 27
		*!* ** } JRN -- 07/11/2016 08:11 AM - End
			This.lEscPress = .T.
			Clear Typeahead
			Return 0
		Endif

		Try
			Chdir (tcDir)
			llChanged = .T.
		Catch
			llChanged = .F.
		Endtry

		If !llChanged
			Return .F.
		Endif

		This.ShowWaitMessage('Scanning directory ' + tcDir)

		lcCurrentDirectory = Curdir()
		lcDriveAndDirectory = Addbs(Sys(5) + Sys(2003))

		This.oDirectories.Add(lcDriveAndDirectory)

		lnDirCount = Adir(laDirList, '*.*', 'D')

		If Vartype(This.oProgressBar) = 'O'
			lnFileCount = Adir(laFileList, lcDriveAndDirectory + '*.*')
			This.oProgressBar.nMaxValue = This.oProgressBar.nMaxValue + lnFileCount
		Endif

		For lnPtr = 1 To lnDirCount
			If 'D' $ laDirList(lnPtr, 5) && If we have found another dir, then we need to work through it also
				If Vartype(This.oProgressBar) = 'O'
					This.oProgressBar.nMaxValue = This.oProgressBar.nMaxValue - 0 && Subtract off directories from file count
				Endif
				lcCurrentDirectory = laDirList(lnPtr, 1)
				If lcCurrentDirectory <> '.' And lcCurrentDirectory <> '..'
					This.BuildDirectoriesCollection(lcCurrentDirectory)
				Endif
			Endif
		Endfor

		Cd ..
		
	EndProc


	*----------------------------------------------------------------------------------
	Procedure BuildProjectsCollection
	
		Local loPEME_BaseTools As 'GF_PEME_BaseTools' Of 'Lib\GF_PEME_BaseTools.prg'
		Local laProjects[1], lcCurrentDir, lcProject, loMRU_Project, loMRU_Projects, loProject, lnX

		lcCurrentDir = Addbs(Sys(5) + Sys(2003)) && Current Default Drive and path

		*-- Blank out current Projects collecitons. Will rebuild below...
		This.oProjects = Createobject('Collection')
		This.cProjects = ''

		If Version(2) = 0 && If we are running from an .EXE file then exit (No projects will be open)
			Return
		Endif

		*-- Add all open Projects in _VFP to the Collection
		For Each loProject In _vfp.Projects
			lcProject = Lower(loProject.Name)
			This.AddProject(lcProject)
			This.cProjects = This.cProjects + lcProject + Chr(13)
		Endfor

		*-- Add any Projects in the current folder
		Adir(laProjects, lcCurrentDir + '*.pjx')

		For lnX = 1 To Alen(laProjects) / 5
			lcProject = Lower(Fullpath(laProjects(lnX, 1)))
			This.AddProject(lcProject)
			This.cProjects = This.cProjects + lcProject + Chr(13)
		Endfor

		*-- Add MRU Projects to the Collection...
		loPEME_BaseTools = CreateObject('GF_PEME_BaseTools')
		
		loMRU_Projects = loPEME_BaseTools.GetMRUList('PJX')

		For Each loMRU_Project In loMRU_Projects
			lcProject = Lower(loMRU_Project)
			This.AddProject(lcProject)
			This.cProjects = This.cProjects + lcProject + Chr(13)
		EndFor
		
	EndProc


	*----------------------------------------------------------------------------------
	Procedure ChangeCurrentDir(tcDir)

		Local lcCurrentDirectory, lcDefaultDrive, lcPath, llReturn

		*-- Attempt to change current dir to passed in location -------
		If !Empty(tcDir)
			Try
				Cd (tcDir)
				llReturn = .T.
			Catch
				This.SetSearchError('Invalid path [' + tcDir + '] passed to ChangeCurrentDir() method.')
				llReturn = .F.
			Endtry
		Else
			llReturn = .T.
		Endif

		This.BuildProjectsCollection()

		Return llReturn
		
	EndProc


	*----------------------------------------------------------------------------------
	Procedure CheckFileExtTemplate(tcFile)

		Local lcFileName, lcFilenameMask, llFilenameMatch, llReturn

		lcFileExtTemplate = Justext(This.oSearchOptions.cFileTemplate)

		llReturn = This.MatchTemplate(tcFile, lcFileExtTemplate)
		
	EndProc


	*----------------------------------------------------------------------------------
	Procedure CheckFilenameTemplate(tcFile)

		Local lcFileName, lcFilenameMask, llMatch, lnLength

		If Empty(Juststem(This.oSearchOptions.cFileTemplate))
			Return .T.
		Endif

		lcFilenameMask = Upper(Juststem(This.oSearchOptions.cFileTemplate))
		lcFileName = Upper(Juststem(tcFile))

		Do Case
			Case lcFilenameMask = '*'
				llMatch = .T.
			Case (Left(lcFilenameMask, 1) = '*' And Right(lcFilenameMask, 1) = '*') Or Atc('*', lcFilenameMask) = 0
				llMatch = lcFilenameMask $ lcFileName
			Case Right(lcFilenameMask, 1) = '*'
				lnLength = Len(cFilenameMask) - 1
				lcFilenameMask = Left(lcFilenameMask, lnLength)
				llMatch = Left(lcFileName, lnLength) = lcFilenameMask
			Case Left(lcFilenameMask, 1) = '*'
				lnLength = Len(cFilenameMask) - 1
				lcFilenameMask = Right(lcFilenameMask, lnLength)
				llMatch = Right(lcFileName, lnLength) = lcFilenameMask
		Endcase

		Return llMatch
		
	EndProc


	*----------------------------------------------------------------------------------
	Procedure CleanUpBinaryString(tcString, llClipAtChr8)

		If llClipAtChr8 && The Select statement from a DBC View needs to be clipped at the Chr(8) near the end of the statement
			lnStart = Atc(Chr(8), tcString)
			tcString = Left(tcString, lnStart)
		Endif

		*-- Replace junk characters with a space
		For x = 0 To 31
			tcString = Strtran(tcString, Chr(x), ' ')
		Endfor

		Return tcString
		
	EndProc


	*----------------------------------------------------------------------------------
	Procedure ClearReplaceErrorMessage
	
		This.oSearchOptions.cReplaceErrorMessage = ''
		
	EndProc


	*----------------------------------------------------------------------------------
	Procedure ClearReplaceSettings
	
		This.oSearchOptions.lAllowBlankReplace = .F.
		This.oSearchOptions.cReplaceExpression = ''
		
	EndProc


	*----------------------------------------------------------------------------------
	Procedure ClearResultsCollection
	
		This.oResults = Createobject('Collection')
		
	EndProc


	*----------------------------------------------------------------------------------
	Procedure ClearResultsCursor()
		
		Local lcSearchResultsAlias, lnSelect

		lnSelect = Select()

		lcSearchResultsAlias = This.cSearchResultsAlias

		Create Cursor (lcSearchResultsAlias)( ;
					Process L, ;
					FilePath C(254), ;
					FileName C(50), ;
					TrimmedMatchLine C(254), ;
					BaseClass C(254), ;
					ParentClass C(254), ;
					Class C(254), ;
					Name C(254), ;
					MethodName C(80), ;
					ContainingClass C(254), ;
					ClassLoc C(254), ;
					MatchType C(25), ;
					Timestamp T, ;
					FileType C(4), ;
					Type C(12), ;
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
					Column C(10), ;
					Code M, ;
					Id I, ;
					MatchLine M, ;
					Replaced L, ;
					TrimmedReplaceLine C(254), ;
					ReplaceLine C(254), ;
					ReplaceRisk I, ;
					Replace_DT T;
					)

		Select (lnSelect)
		
	EndProc


	*----------------------------------------------------------------------------------
	Procedure Compile(tcFile)

		Local lcExt

		If This.oSearchOptions.lPreviewReplace = .T.
			Return
		Endif

		lcExt = Alltrim(Upper(Justext(tcFile)))

		Do Case
			Case lcExt = 'VCX'
				Compile Classlib (tcFile)

			Case lcExt = 'SCX'
				Compile Form (tcFile)

			Case lcExt = 'LBX'
				Compile Label (tcFile)

			Case lcExt = 'FRX'
				Compile Report (tcFile)
		Endcase
		
	EndProc


	*----------------------------------------------------------------------------------
	Procedure CreateMenuDisplay(tcMenu)

		#Define SPACING 3
		#Define PREFIX '*'

		Local laLevels[1], lcPrompt, lcResult, lnLevel, lnSelect

		lnSelect = Select()
		Select (tcMenu)

		lcResult = ''
		lnLevel = 1
		Dimension This.aMenuStartPositions[Reccount(tcmenu)]

		Scan
			This.aMenuStartPositions[Recno(tcmenu)] = Len(lcResult)
			Do Case
				Case objCode = 22

				Case objCode = 1
					laLevels[1]	= Name

				Case objCode = 77
					lcPrompt = Prompt
					lnLevel	 = Ascan(m.laLevels, Trim(LevelName))

				Case objCode = 0
					lcResult = m.lcResult + PREFIX + Space(SPACING * m.lnLevel) + Strtran(m.lcPrompt, '\-', '-----') + CR
					lnLevel	 = m.lnLevel + 1
					Dimension m.laLevels[m.lnLevel]
					laLevels[m.lnLevel]	= Name

				Otherwise
					lnLevel	 = Ascan(m.laLevels, Trim(LevelName))
					lcResult = m.lcResult + PREFIX + Space(SPACING * m.lnLevel) + Strtran(Prompt, '\-', '-----') + CR
			Endcase
		Endscan

		Select(m.lnSelect)

		Return m.lcResult
		
	EndProc


	*----------------------------------------------------------------------------------
	Procedure CreateReplaceDetailRecord(toReplace)

		Local lcDBC, lnID, lnSelect

		If This.oSearchOptions.lPreviewReplace = .T. && Do nothing if we are only in Preview mode
			Return
		Endif

		lnSelect = Select()

		lcDBC = Strtran(Upper(This.cReplaceDetailTable), '.DBF', '.DBC')

		If !File(This.cReplaceDetailTable) Or !File(lcDBC)
			Set Safety Off
			Delete File (lcDBC)
			Set Safety On
			Create Database (lcDBC)
			Select (This.cSearchResultsAlias)
			Copy Structure To (This.cReplaceDetailTable) Database (lcDBC)

			Alter Table (This.cReplaceDetailTable) Add Column Pk I Autoinc
			Alter Table (This.cReplaceDetailTable) Add Column HistoryFK I

			Use In (This.cReplaceDetailTable)
		Endif

		This.MigrateReplaceDetailTable()

		toReplace.Replace_DT = Datetime()
		toReplace.Replaced = .T.
		AddProperty(toReplace, 'HistoryFK', This.nReplaceHistoryId)

		Insert Into (This.cReplaceDetailTable) From Name toReplace

		Select(lnSelect)
		
	EndProc


	*----------------------------------------------------------------------------------
	Procedure CreateReplaceHistoryRecord
		
		Local lcScope, lnSelect

		If This.oSearchOptions.lPreviewReplace = .T. && Do nothing if we are only in Replace Preview mode
			Return
		Endif

		lnSelect = Select()

		If !File(This.cReplaceHistoryTable)
			Create Table (This.cReplaceHistoryTable) Free (;
						Id I Autoinc Nextvalue 1001, ;
						Date_Time T, ;
						replaces I, ;
						Scope C(254), ;
						searchstr C(254), ;
						replacestr C(254) ;
						)
		Endif

		If Inlist(This.oSearchOptions.nSearchScope, 1, 2)
			lcScope = This.oSearchOptions.cProject
		Else
			lcScope = This.oSearchOptions.cPath
		Endif

		Insert Into (This.cReplaceHistoryTable) (Scope, Date_Time, searchstr, replacestr) ;
			Values (lcScope, Datetime(), This.oSearchOptions.cSearchExpression, This.oSearchOptions.cReplaceExpression)

		This.nReplaceHistoryId = Evaluate(Juststem(Justfname(This.cReplaceHistoryTable)) + '.Id')

		Select (lnSelect)
		
	EndProc


	*----------------------------------------------------------------------------------
	Procedure CreateResult(toObject)

		This.nMatchLines = This.nMatchLines + 1

		If This.oSearchOptions.lCreateResultsCursor
			This.CreateResultsRow(toObject)
		Endif

		If This.oSearchOptions.lCreateResultsCollection
			This.oResults.Add(toObject)
		Endif
		
	EndProc


	*----------------------------------------------------------------------------------
	Procedure CreateResultsRow(toObject)

		*-- This set of mem vars is required to insert a new row into the local results cursor.
		*-- The passed in toObject must be an object which has the reference properties on it, so
		*-- that a complete record can be created.

		Local lIsText, lcObjectNameFromProperty, lcProperty, lcResultsAlias, lnWords
		Local Timestamp, BaseClass, Class, ClassLoc, Code, Column, ContainingClass, FileName, FilePath
		Local FileType, Id, MatchLen, MatchLine, MatchStart, MatchType, MethodName, Name, ParentClass
		Local proccode, procend, Process, ProcStart, Recno, ReplaceRisk, TrimmedMatchLine
		Local statement, statementstart

		lcResultsAlias = This.cSearchResultsAlias

		With m.toObject
			MethodName		 = This.FixPropertyName(.MethodName)
			MatchLine		 = .MatchLine
			TrimmedMatchLine = .TrimmedMatchLine
			ProcStart		 = .ProcStart
			procend			 = .procend
			proccode		 = Evl(.proccode, .Code)
			statement		 = Evl(.statement, .MatchLine)
			statementstart   = .statementstart
			MatchStart		 = .MatchStart
			MatchLen		 = .MatchLen
			MatchType		 = .MatchType
			Code			 = .Code
		Endwith

		With m.toObject.UserField
			Process			= .F.
			FilePath		= Lower(.FilePath)
			FileName		= Lower(.FileName)
			FileType		= .FileType
			lIsText			= .IsText
			BaseClass		= ._BaseClass
			ParentClass		= ._ParentClass
			ContainingClass	= .ContainingClass
			Name			= ._Name
			Class			= ._Class
			ClassLoc		= .ClassLoc
			Recno			= .Recno && from the VCX, SCX, VCX, etc.
			Timestamp		= .Timestamp
			Column			= .Column
		Endwith

		If "key" $ Lower(m.MatchType)
			Suspend
		Endif

		* *-- Removed 07/07/2012
		* *--- Clean up / doctor up the Object Name
		* If 'scx' $ Lower(m.filetype)  && trim off the form name from front of object name
		* 	m.name = Substr(m.name, Atc('.', m.name) + 1)
		* EndIf

		*--- Sometimes, part of the object name may live on the match line
		*--- So, we need to append it to the end of the object name
		If m.MatchType $ (MATCHTYPE_PROPERTY_NAME + MATCHTYPE_PROPERTY_VALUE)
			lcObjectNameFromProperty = ''
			lcProperty				 = Getwordnum(m.TrimmedMatchLine, 1)
			lnWords					 = Getwordcount(m.lcProperty, '.')

			If m.lnWords > 1
				lcObjectNameFromProperty = Left(m.lcProperty, Atc('.', m.lcProperty, m.lnWords - 1) - 1)
			Endif

			Name = Alltrim(m.name + '.' + m.lcObjectNameFromProperty, '.')
		Endif

		*--------------------------------------------------------------------------------

		Id = Reccount(m.lcResultsAlias) + 1 && A unique key for each record

		ReplaceRisk = This.GetReplaceRiskLevel(m.toObject)

		Insert Into &lcResultsAlias From Memvar
		
	EndProc


	*----------------------------------------------------------------------------------
	Procedure Destroy
	
		This.oRegExForProcedureStartPositions = .Null.
		This.oRegExForSearch = .Null.
		This.oResults = .Null.
		This.oSearchOptions = .Null.
		This.oFrxCursor = .Null.
		This.oProjects = .Null.
		This.oSearchErrors = .Null.
		This.oReplaceErrors = .Null.
		This.oDirectories = .Null.
		This.oProgressBar = .Null.
		This.oFSO = .Null.
		
	EndProc


	*----------------------------------------------------------------------------------
	Procedure DropColumn(tcTable, tcColumnName)

		Local lcAlias

		lcAlias = Juststem(tcTable)

		Try
			Alter Table (tcTable)  Drop Column &tcColumnName
		Catch
		Endtry
		
	EndProc


	*----------------------------------------------------------------------------------
	Procedure EditFromCurrentRow(tcCursor, tlSelectObjectOnly)

		Local loPBT As 'GF_PEME_BaseTools'
		Local lcClass, lcCodeBLock, lcExt, lcFileToEdit, lcMatchType, lcMethod, lcMethodString, lcName
		Local lcProperty, lnMatchStart, lnProcStart, lnRecNo, lnStart, lnWords, loTools

		lcExt		 = Alltrim(Upper(&tcCursor..FileType))
		lcFileToEdit = Upper(Alltrim(&tcCursor..FilePath))
		lcClass		 = Alltrim(&tcCursor..Class)
		lcName		 = Alltrim(&tcCursor..Name)
		lcMethod	 = Alltrim(&tcCursor..MethodName)
		lcMatchType	 = Alltrim(&tcCursor..MatchType)
		lnRecNo		 = &tcCursor..Recno
		lnProcStart	 = &tcCursor..ProcStart
		lnMatchStart = &tcCursor..MatchStart

		If lcExt # 'PRG' And (Empty(m.lcMethod) Or 0 # Atc('<Property', m.lcMatchType))
			lcMethodString = ''
			lnStart		   = 1
		Else
			lcMethodString = Alltrim(m.lcName + '.' + m.lcMethod, 1, '.')

			If m.lcExt $ ' SCX VCX '
				*-- Calculate Line No from procstart and matchstart postitions...
				lcCodeBLock	= Substr(&tcCursor..Code, m.lnProcStart + 1, m.lnMatchStart - m.lnProcStart)
				lnStart		= Getwordcount(m.lcCodeBLock, Chr(13)) - 1 && The LINE NUMBER that match in on within the method
				lnStart		= Iif(m.lnStart > 0, m.lnStart, 1)
				Do Case
					Case m.lcExt = 'SCX'
						lcClass = ''
						*If Lower(Alltrim(&tcCursor..Baseclass)) <> 'form'
						lcMethodString = 'x.' + m.lcMethodString
						*EndIf

					Case m.lcExt = 'VCX'
						If m.lcName = m.lcClass
							lcMethodString 	= m.lcMethod
						Endif
				Endcase
			Else
				lnStart = (&tcCursor..MatchStart) + 1 && The CHARACTER position of the line where the match is on
			Endif

		EndIf
		
		loPBT = CreateObject('GF_PEME_BaseTools')

		*** JRN 2021-03-21 : If match is to a name of a file in a Project, open that file
		If &tcCursor..FileType = 'PJX' And &tcCursor..MatchType = MATCHTYPE_NAME
			lcFileToEdit = FullPath(Upper(Addbs(JustPath(Trim(&tcCursor..FilePath))) + Trim(&tcCursor..TrimmedMatchLine)))
			m.loPBT.EditSourceX(m.lcFileToEdit)
			Return
		Else
			lcFileToEdit = Upper(Alltrim(&tcCursor..FilePath))
		EndIf
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
					m.loPBT.EditSourceX(m.lcFileToEdit, m.lcClass)
					m.loTools.SelectObject(m.lcName)
					Return

				Case m.lcMatchType $ (MATCHTYPE_PROPERTY_NAME + MATCHTYPE_PROPERTY_VALUE + MATCHTYPE_PROPERTY_DEF )
					*-- Pull out the Property name from the MatchLine (it can be preceded by an object name)
					lcProperty = Getwordnum(&tcCursor..TrimmedMatchLine, 1)
					lnWords	   = Getwordcount(m.lcProperty, '.')
					lcProperty = Getwordnum(m.lcProperty, m.lnWords, '. ')
					lcProperty = This.FixPropertyName(m.lcProperty)

					m.loPBT.EditSourceX(m.lcFileToEdit, m.lcClass)
					m.loTools.SelectObject(m.lcName, m.lcProperty)
					Return
				Endcase
			Endif

		Endif

		m.loPBT.EditSourceX(m.lcFileToEdit, m.lcClass, m.lnStart, m.lnStart, m.lcMethodString, m.lnRecNo)

		If m.lcExt = 'PRG' Or Not Empty(m.lcMethodString)
			This.ThorMoveWindow()
		Endif
		
	EndProc


	*----------------------------------------------------------------------------------
	Procedure EditMenuFromCurrentRow(tcCursor)

		Local loEditorWin As Editorwin Of 'c:\visual foxpro\programs\mythor\thor\tools\apps\pem editor\source\peme_editorwin.vcx'
		Local lcFileToEdit, lcMenuAlias, lcMenuDisplay, lcTempFile, llSuccess, lnEndPos, lnRecNo, lnStartPos

		lcMenuAlias	 = 'Menu' + Sys(2015)
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
		lcTempFile	  = Addbs(Sys(2023)) + Chrtran(Justfname(m.lcFileToEdit), '.', '_')  + Sys(2015) + '.txt'
		Strtofile(m.lcMenuDisplay, m.lcTempFile)
		Modify File (m.lcTempFile) Nowait

		loEditorWin = Execscript(_Screen.cThorDispatcher, 'Class= editorwin from pemeditor')
		m.loEditorWin.ResizeWindow(600, 800)
		m.loEditorWin.SetTitle(m.lcTempFile)

		lnRecNo = &tcCursor..Recno
		
		If Between(m.lnRecNo, 1, Reccount(m.lcMenuAlias))
			lnStartPos = This.aMenuStartPositions[m.lnRecNo]
			If m.lnRecNo < Reccount(m.lcMenuAlias)
				lnEndPos = This.aMenuStartPositions[m.lnRecNo + 1]
			Else
				lnEndPos = 1000000
			Endif

			m.loEditorWin.EnsureVisible(0)
			m.loEditorWin.Select(m.lnStartPos, m.lnEndPos)
			m.loEditorWin.EnsureVisible(m.lnStartPos)
		Endif

		Use In (m.lcMenuAlias)
		
	EndProc


	*----------------------------------------------------------------------------------
	Procedure EditObjectFromCurrentRow(tcCursor)

		Local loPBT As 'GF_PEME_BaseTools'
		Local lcClass, lcExt, lcFileToEdit, lcMatchType, lcName, lcProperty, lnWords, loTools

		lcExt		 = Alltrim(Upper(&tcCursor..FileType))
		lcMatchType	 = Alltrim(&tcCursor..MatchType)
		lcFileToEdit = Upper(Alltrim(&tcCursor..FilePath))
		lcClass		 = Alltrim(&tcCursor..Class)
		lcName		 = Alltrim(&tcCursor..Name)

		loPBT = CreateObject('GF_PEME_BaseTools')
		m.loPBT.EditSourceX(m.lcFileToEdit, m.lcClass)

		If Type('_Screen.cThorDispatcher') = 'C'

			loTools = Execscript(_Screen.cThorDispatcher, 'Class= tools from pemeditor')

			If Vartype(m.loTools) = 'O'

				If m.lcExt = 'SCX' And &tcCursor..BaseClass = 'form' && Must trim off form name from front of object name
					lcName = ''
				Endif

				If m.lcMatchType $ (MATCHTYPE_PROPERTY_NAME + MATCHTYPE_PROPERTY_VALUE + MATCHTYPE_PROPERTY_DEF )
					*-- Pull out the Property name from the MatchLine (it can be preceded by an object name)
					lcProperty = Getwordnum(&tcCursor..TrimmedMatchLine, 1)
					lnWords	   = Getwordcount(m.lcProperty, '.')
					lcProperty = Getwordnum(m.lcProperty, m.lnWords, '. ')
					lcProperty = This.FixPropertyName(m.lcProperty)

					m.loTools.SelectObject(m.lcName, m.lcProperty)
				Else
					m.loTools.SelectObject(m.lcName)
				Endif
			Endif

		Endif
		
	EndProc


	*----------------------------------------------------------------------------------
	Procedure EndTimer()
	
		This.nSearchTime = Seconds() - This.nSearchTime
		
	EndProc


	*----------------------------------------------------------------------------------
	Procedure EscapeSearchExpression(tcString)

		Local lcString

		lcString = tcString

		lcString = Strtran(tcString, '\', '\\')
		lcString = Strtran(lcString, '+', '\+')
		lcString = Strtran(lcString, '.', '\.')
		lcString = Strtran(lcString, '|', '\|')
		lcString = Strtran(lcString, '{', '\{')
		lcString = Strtran(lcString, '}', '\}')
		lcString = Strtran(lcString, '[', '\[')
		lcString = Strtran(lcString, ']', '\]')
		lcString = Strtran(lcString, '(', '\(')
		lcString = Strtran(lcString, ')', '\)')
		lcString = Strtran(lcString, '$', '\$')

		lcString = Strtran(lcString, '^', '\^')
		lcString = Strtran(lcString, ':', '\:')
		lcString = Strtran(lcString, ';', '\;')
		lcString = Strtran(lcString, '-', '\-')

		If This.oSearchOptions.nSearchMode = GF_SEARCH_MODE_LIKE
			lcString = Strtran(lcString, '?', '.')
			lcString = Strtran(lcString, '*', '.*')
		Else
			lcString = Strtran(lcString, '?', '\?')
			lcString = Strtran(lcString, '*', '\*')
		Endif

		Return lcString

		* http://stackoverflow.com/questions/280793/case-insensitive-string-replacement-in-javascript

		*!*	RegExp.escape = function(str) {
		*!*	var specials = new RegExp("[.*+?|()\\[\\]{}\\\\]", "g"); // .*+?|()[]{}\
		*!*	return str.replace(specials, "\\$&");
		*!*	}
		
	EndProc


	*----------------------------------------------------------------------------------
	Procedure ExtractMethodName(tcReference)

		If !Empty(tcReference)
			Return Justext(tcReference)
		Else
			Return ''
		Endif
		
	EndProc


	*----------------------------------------------------------------------------------
	Procedure ExtractObjectName(tcReference)

		If !Empty(tcReference)
			Return Juststem(tcReference)
		Else
			Return ''
		Endif
		
	EndProc


	*----------------------------------------------------------------------------------
	Procedure FilesToSkip(tcFile)

		Local lcFileName, lnI

		lcFileName = Upper(Justfname(m.tcFile))

		If (Chr(13) + m.lcFileName + Chr(13)) $ This.cFilesToSkip
			Return .T.
		Endif

		For lnI = 1 To This.nWildCardFilesToSkip
			If Like(This.aWildcardFiles[lni], m.tcFile)
				Return .T.
			Endif
		Endfor

		Return .F.
		
	EndProc


	*----------------------------------------------------------------------------------
	Procedure FindProcedureForMatch(toProcedureStartPositions, tnStartByte)

		Local loReturn As 'GF_Procedure'
		Local loClassDef As 'GF_Procedure'
		Local llClassDef, lnX, loNextMatch

		loReturn = Createobject('GF_Procedure')

		If Isnull(toProcedureStartPositions)
			Return loReturn
		Endif

		lnX = 1

		For Each Result In toProcedureStartPositions

			If Result.StartByte > tnStartByte
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

			lnX = lnX + 1
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

		Return loReturn
		
	EndProc


	*----------------------------------------------------------------------------------
	Procedure FindStatement(loObject)

		Local lcLastLine, lcMatchLine, lcPreceding, lcProcCode, lcResult, lnCRPos, lnLen, lnLength, lnStart
		Local lnTextStart

		lnStart	 = m.loObject.MatchStart - m.loObject.ProcStart + 1

		*!* ** { JRN -- 08/05/2016 07:16 AM - Begin
		*!* lnLength = m.loObject.MatchLen - 1
		lnLength = m.loObject.MatchLen
		*!* ** } JRN -- 08/05/2016 07:16 AM - End

		lcProcCode = Evl(loObject.proccode, loObject.Code)

		* previously, assumed trailing CR, but this dropped off last character if not found
		*!* ** { JRN -- 08/05/2016 07:16 AM - Begin
		*!* lcMatchLine	= Substr(m.lcProcCode, m.lnStart, m.lnLength)
		lcMatchLine	= Trim(Substr(m.lcProcCode, m.lnStart, m.lnLength), 1, Chr[13], Chr[10])
		*!* ** } JRN -- 08/05/2016 07:16 AM - End

		lcResult	= m.lcMatchLine
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
		If Atc('text', Left(m.lcProcCode, m.lnStart)) != 0 And ;
				Atc('endtext', Substr(m.lcProcCode, m.lnStart + m.lnLength)) != 0

			lnTextStart = m.lnStart
			Do While m.lnTextStart > 1
				lcPreceding	= Left(m.lcProcCode, m.lnTextStart - 1)
				lnCRPos		= Rat(CR, m.lcPreceding, 2)
				If m.lnCRPos > 0
					lcPreceding = Substr(m.lcPreceding, m.lnCRPos + 1)
				Endif
				Do Case
				Case 'text' = Lower(Getwordnum(m.lcPreceding, 1, ' ' + Tab + CR + lf))
					lnTextStart = m.lnTextStart - Len(m.lcPreceding)
					lnLength = m.lnLength + m.lnStart - m.lnTextStart
					lnStart	 = m.lnTextStart
					lcResult = Substr(m.lcProcCode, m.lnStart, m.lnLength)
					Do While Len(m.lcProcCode) > m.lnStart + m.lnLength
						lcLastLine = Substr(m.lcProcCode, m.lnStart + m.lnLength)
						lnLen	   = At(CR, m.lcLastLine, 2)
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


		loObject.statement		= m.lcResult
		loObject.statementstart	= m.lnStart
		
	EndProc


	*----------------------------------------------------------------------------------
	Procedure FixPropertyName(lcProperty)

		* Gets rid of dimensions and leading '*^'

		Return ChrTran(Getwordnum(m.lcProperty, 1, '(['), '*^', '')
		
	EndProc


	*----------------------------------------------------------------------------------
	Procedure GenerateHTMLCode(tcCode, tcMatchLine, tnMatchStart, tcCss, tcJavaScript, tcReplaceLine, tlAlreadyReplaced, tnTabsToSpaces, ;
							   tcSearch, tcStatementFilter, tcProcFilter)

		Local lcColorizedCode, lcCss, lcHTML, lcHtmlBody, lcInitialBr, lcJavaScript, lcLeft, lcMatchLine
		Local lcMatchPrefix, lcMatchSuffix, lcMatchWordPrefix, lcMatchWordSuffix, lcReplaceExpression
		Local lcRight, lcRightCode, lnEndProc, lnMatchLineLength, lcBr
		Local lcMatchLinePrefix, lcMatchLineSuffix, lcReplaceLine
		Local lcReplaceLinePrefix, lcReplaceLineSuffix, lnReplaceLineLength

		lcCss = Evl(tcCss, '')
		lcJavaScript = Evl(tcJavaScript, '')

		lcMatchLinePrefix = '<div id="matchline" class="matchline">'
		lcMatchLineSuffix = '</div>'
		lcReplaceLinePrefix = '<div id="repalceline" class="replaceline">'
		lcReplaceLineSuffix = '</div>'

		lcMatchWordPrefix = '<span id="matchword" class="matchword">'
		lcMatchWordSuffix = '</span>'

		If !Empty(tcMatchLine)

			*-- Dress up the code that comes before the match line...
			lcBr = '<br />'
			lcLeft = Left(tcCode, tnMatchStart)
			lcLeft = Evl(This.HtmlEncode(lcLeft), lcBr)

			*-- Dress up the matchline...
			lnMatchLineLength = Len(tcMatchLine)
			lnReplaceLineLength = Len(Rtrim(tcReplaceLine))

			*===================== Colorize the Replace Preview line, if passed ==============================
			If !Empty(tcReplaceLine)
				lcColorizedCode = This.HtmlEncode(tcReplaceLine)
				lcReplaceLine = lcReplaceLinePrefix + lcColorizedCode + lcReplaceLineSuffix
				lcMatchLinePrefix = '<div id="matchline" class="strikeout">'
				lcMatchLinePrefix = lcMatchLinePrefix + '<del>'
				lcMatchLineSuffix = lcMatchLineSuffix + '</del>'
			Else
				lcReplaceLine = ''
			Endif

			*===================== Colorize the match line ====================================================
			*-- Mark the match WORD(s), so I can find them after the VFP code is colorized...
			lcReplaceExpression = '[:GOFISHMATCHWORDSTART:] + lcMatch + [:GOFISHMATCHWORDEND:]'
			lcColorizedCode = This.RegExReplace(tcMatchLine, '', lcReplaceExpression, .T.)

			lcColorizedCode = This.HtmlEncode(lcColorizedCode)

			*-- Next, add <span> tags around previously marked match Word(s)
			lcColorizedCode = Strtran(lcColorizedCode, ':GOFISHMATCHWORDSTART:', lcMatchWordPrefix)
			lcColorizedCode = Strtran(lcColorizedCode, ':GOFISHMATCHWORDEND:', lcMatchWordSuffix)

			*-- Finally, add <div> tags around the entire Matched Line -------------------
			lcMatchLine = lcMatchLinePrefix + lcColorizedCode + lcMatchLineSuffix
			*=================================================================================================

			*-- Dress up the code that comes after the match line...
			*-- (Look for EndProc to know where to end the code)---
			If tlAlreadyReplaced = .T.
				lcRightCode = Substr(tcCode, tnMatchStart + 1 + lnReplaceLineLength)
			Else
				lcRightCode = Substr(tcCode, tnMatchStart + 1 + lnMatchLineLength)
			Endif

			lnEndProc = Atc('EndProc', lcRightCode)

			If lnEndProc > 0
				lcRightCode = Substr(lcRightCode, 1, lnEndProc + 6) && It ends at "E" of "EndProc", so add 6 to get the rest of the word
			Endif

			lcRight = This.HtmlEncode(lcRightCode)

			*!* ******************** Removed 12/02/2015 *****************
			*!* *** JRN 11/14/2015 : Highlight sub-search (filter within same procedure) if it is supplied
			*!* If 'C' = Vartype(tcProcFilter) and not Empty(tcProcFilter)
			*!* 	lcLeft = This.HighlightProcFilter(lcLeft, tcProcFilter, lcMatchWordPrefix, lcMatchWordSuffix) 
			*!* 	lcRight = This.HighlightProcFilter(lcRight, tcProcFilter, lcMatchWordPrefix, lcMatchWordSuffix) 
			*!* EndIf 		

			lcHtmlBody = lcLeft + lcMatchLine + lcReplaceLine + lcRight &&Build the body

			If 'C' = Vartype(tcSearch) And Not Empty(tcSearch)
				lcHtmlBody = This.HighlightProcFilter(lcHtmlBody, tcSearch, lcMatchWordPrefix, lcMatchWordSuffix)
			Endif

			If 'C' = Vartype(tcStatementFilter) And Not Empty(tcStatementFilter)
				lcHtmlBody = This.HighlightProcFilter(lcHtmlBody, tcStatementFilter, lcMatchWordPrefix, lcMatchWordSuffix)
			Endif

			If 'C' = Vartype(tcProcFilter) And Not Empty(tcProcFilter)
				lcHtmlBody = This.HighlightProcFilter(lcHtmlBody, tcProcFilter, lcMatchWordPrefix, lcMatchWordSuffix)
			Endif

			If Not Empty(tnTabsToSpaces)
				lcHtmlBody = Strtran(lcHtmlBody, Chr[9], Space(tnTabsToSpaces))
			Endif
			lcHtmlBody = Alltrim(lcHtmlBody, 1, Chr[13], Chr[10])
		Else

			*-- Just a plain blob of VFP code, with no match lines or match words...
			*-- Need an empty MatchLine Divs so the JavaScript on the page will find it to scroll the page
			lcHtmlBody = '<div id="matchline"></div>' + This.HtmlEncode(tcCode)

		Endif

		*-- Build the whole Html by combining the html parts defined above -------------
		Text To lcHtml Noshow Textmerge Pretext 3
			<html>
				<head>
					<<lcCss>>
				</head>
		
				<body>
					<<lcHtmlBody>>
					<br /><br /><br />
					<<lcJavaScript>>
				</body>
			</html>
		ENDTEXT


		Return lcHTML
		
	EndProc


	*----------------------------------------------------------------------------------
	Procedure GetActiveProject()
	
		Local lcCurrentProject

		If Type('_VFP.ActiveProject.Name') = 'C'
			lcCurrentProject = _vfp.ActiveProject.Name
		Else
			lcCurrentProject = ''
		Endif

		Return lcCurrentProject
		
	EndProc


	*----------------------------------------------------------------------------------
	Procedure GetCurrentDirectory

		Return Addbs(Sys(5) + Sys(2003))
		
	EndProc


	*----------------------------------------------------------------------------------
	Procedure GetDirectories(tcPath, tlIncludeSubDirectories)

		Local laFiles[1], lnFiles

		This.oDirectories = Createobject('Collection')

		If tlIncludeSubDirectories
			This.BuildDirectoriesCollection(tcPath)
		Else
			This.oDirectories.Add(tcPath)
			If Vartype(This.oProgressBar) = 'O'
				lnFiles = Adir(laFiles, '*.*')
				This.oProgressBar.nMaxValue = lnFiles
			Endif
		Endif

		Return This.oDirectories
		
	EndProc


	*----------------------------------------------------------------------------------
	Procedure GetFileDateTime(tcFile)

		Local lcFileName, loFile, lcExt

		ldFileDate = {// ::}
		lcExt = Upper(Justext(tcFile))

		If Inlist(lcExt, 'SCX', 'VCX', 'FRX', 'MNX', 'LBX')
			Try
				Use (tcFile) Again In 0 Alias 'GF_GetMaxTimeStamp' Shared
				Select Max(Timestamp) From GF_GetMaxTimeStamp Into Array laMaxDateTime
				ldFileDate = Ctot(This.TimeStampToDate(laMaxDateTime))
			Catch
			Finally
				If Used('GF_GetMaxTimeStamp')
					Use In ('GF_GetMaxTimeStamp')
				Endif
			Endtry
		Endif

		If Empty(ldFileDate)
			Try
				ldFileDate = Fdate(tcFile, 1)
			Catch
				loFile = This.oFSO.Getfile(tcFile)
				ldFileDate = loFile.DateLastModified
			Endtry
		Endif

		Return ldFileDate
		
	EndProc


	*----------------------------------------------------------------------------------
	Procedure GetFrxObjectType(tnObjType, tnObjCode)

		Local lcObjectType

		*-- Details from: http://www.dbmonster.com/Uwe/Forum.aspx/foxpro/4719/Code-meanings-for-Report-Format-ObjType-field

		lcObjectType = ''

		Do Case
			Case tnObjType = 1
				lcObjectType = 'Report'
			Case tnObjType = 2
				lcObjectType = 'Workarea'
			Case tnObjType = 3
				lcObjectType = 'Index'
			Case tnObjType = 4
				lcObjectType = 'Relation'
			Case tnObjType = 5
				lcObjectType = 'Text'
			Case tnObjType = 6
				lcObjectType = 'Line'
			Case tnObjType = 7
				lcObjectType = 'Box'
			Case tnObjType = 8
				lcObjectType = 'Field'
			Case tnObjType = 9 && Band Info
				Do Case
					Case tnObjCode = 0
						lcObjectType = 'Title'
					Case tnObjCode = 1
						lcObjectType = 'PageHeader'
					Case tnObjCode = 2
						lcObjectType = 'Column Header'
					Case tnObjCode = 3
						lcObjectType = 'Group Header'
					Case tnObjCode = 4
						lcObjectType = 'Detail Band'
					Case tnObjCode = 5
						lcObjectType = 'Group Footer'
					Case tnObjCode = 6
						lcObjectType = 'Column Footer'
					Case tnObjCode = 7
						lcObjectType = 'Page Footer'
					Case tnObjCode = 8
						lcObjectType = 'Summary'
				Endcase
			Case tnObjType = 10
				lcObjectType = 'Group'
			Case tnObjType = 17
				lcObjectType = 'Picture/OLE'
			Case tnObjType = 18
				lcObjectType = 'Variable'
			Case tnObjType = 21
				lcObjectType = 'Print Drv Setup'
			Case tnObjType = 25
				lcObjectType = 'Data Env'
			Case tnObjType = 26
				lcObjectType = 'Cursor Obj'
		Endcase

		Return lcObjectType
		
	EndProc


	*----------------------------------------------------------------------------------
	Procedure GetProcedureStartPositions(tcCode, tcName)

		Local loObject As 'Empty'
		Local loRegExp As 'VBScript.RegExp'
		Local loResult As 'Collection'
		Local lcBaseClass, lcClassName, lcMatch, lcName, lcParentClass, lcType, lcWord1, llClassDef
		Local llTextEndText, lnEndByte, lnI, lnLFs, lnStartByte, lnX, loException, loMatch, loMatches

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

		loMatches = m.loRegExp.Execute(m.tcCode)

		loResult = Createobject('Collection')

		llClassDef	  = .F. && currently within a class?
		llTextEndText = .F. && currently within a Text/EndText block?
		lcClassName	  = ''
		lcParentClass = ''
		lcBaseClass	  = ''

		For lnI = 1 To m.loMatches.Count

			loMatch = m.loMatches.Item(m.lnI - 1)

			With m.loMatch
				lnStartByte	= .FirstIndex
				lcMatch		= Chrtran(.Value, CR + lf, '  ')
				lcName		= Getwordnum(m.lcMatch, Getwordcount(m.lcMatch))
				lcWord1		= Upper(Getwordnum(m.lcMatch, Max(1, Getwordcount(m.lcMatch) - 1)))
			Endwith

			Do Case
			Case m.llTextEndText
				If 'ENDTEXT' = m.lcWord1
					llTextEndText = .F.
				Endif
				Loop

			Case m.llClassDef
				If 'ENDDEFINE' = m.lcWord1
					llClassDef	  = .F.
					lcType		  = 'End Class'
					lcName		  = m.lcClassName + '.-EndDefine'
					lcClassName	  = ''
					lcParentClass = ''
					lcBaseClass	  = ''
				Else
					lcType = 'Method'
					lcName = m.lcClassName + '.' + m.lcName
				Endif

			Case ' CLASS ' $ Upper(m.lcMatch) && Notice the spaces in ' CLASS '
				llClassDef	  = .T.
				lcType		  = 'Class'
				lcClassName	  = Getwordnum(m.lcMatch, 3)
				lcParentClass = Getwordnum(m.lcMatch, 5)
				lcName		  = ''
				lcBaseClass	  = ''
				If This.IsBaseclass(m.lcParentClass)
					lcBaseClass	  = Lower(m.lcParentClass)
					lcParentClass = ''
				Endif

			Case 'FUNCTION' = m.lcWord1
				lcType = 'Function'

			Otherwise
				lcType = 'Procedure'

			Endcase

			lnLFs = Occurs(Chr(10), m.loMatch.Value)
			lnX	  = 0
			* ignore leading CRLF's, and [spaces and tabs, except on the matched line]
			Do While Substr(m.tcCode, m.lnStartByte + 1, 1) $ Chr(10) + Chr(13) + Chr(32) + Chr(9) And m.lnX < m.lnLFs
				If Substr(m.tcCode, m.lnStartByte + 1, 1) = Chr(10)
					lnX = m.lnX + 1
				Endif
				lnStartByte = m.lnStartByte + 1
			Enddo

			loObject = Createobject('GF_Procedure')

			With m.loObject
				.Type		  = m.lcType
				.StartByte	  = m.lnStartByte
				._Name		  = m.lcName
				._ClassName	  = m.lcClassName
				._ParentClass = m.lcParentClass
				._BaseClass	  = m.lcBaseClass
			Endwith

			Try
				m.loResult.Add(m.loObject, m.lcName)
			Catch To m.loException When m.loException.ErrorNo = 2062 Or m.loException.ErrorNo = 11
				*loResult.Add(loObject, lcName + ' ' + Transform(lnStartByte))
				m.loResult.Add(m.loObject, m.lcName + Sys(2015))
			Catch To m.loException
				This.ShowErrorMsg(m.loException)
			Endtry


		Endfor

		*** JRN 11/09/2015 : determine ending byte for each entry
		lnEndByte = Len(m.tcCode)
		For lnI = m.loResult.Count To 1 Step - 1
			loResult[m.lnI].EndByte = m.lnEndByte
			lnEndByte = 	m.loResult[m.lnI].StartByte
		Endfor

		Return m.loResult
		
	EndProc


	*----------------------------------------------------------------------------------
	Procedure GetRegExForProcedureStartPositions()
	
		Local loRegExp As 'VBScript.RegExp'
		Local lcPattern

		loRegExp = Createobject('VBScript.RegExp')

		With loRegExp
			.IgnoreCase	= .T.
			.Global		= .T.
			.MultiLine	= .T.
		Endwith

		lcPattern = 'PROC(|E|ED|EDU|EDUR|EDURE)\s+(\w|\.)+'
		lcPattern = lcPattern + '|' + 'FUNC(|T|TI|TIO|TION)\s+(\w|\.)+'
		lcPattern = lcPattern + '|' + 'DEFINE\s+CLASS\s+\w+\s+\w+\s+\w+'
		lcPattern = lcPattern + '|' + 'DEFI\s+CLAS\s+\w+'
		lcPattern = lcPattern + '|' + 'ENDD(E|EF|EFI|EFIN|EFINE)\s+'
		lcPattern = lcPattern + '|' + 'PROT(|E|EC|ECT|ECTE|ECTED)\s+\w+\s+\w+'
		lcPattern = lcPattern + '|' + 'HIDD(|E|EN)\s+\w+\s+\w+'

		With loRegExp
			.Pattern	= '^\s*(' + lcPattern + ')'
		Endwith

		Return loRegExp
		
	EndProc


	*----------------------------------------------------------------------------------
	Procedure GetRegExForSearch
	
		Local loRegEx As 'VBScript.RegExp'

		loRegEx = Createobject ('VBScript.RegExp')

		Return loRegEx
		
	EndProc


	*----------------------------------------------------------------------------------
	Procedure GetReplaceResultObject
	
		Local loResult As 'Empty'

		loResult = Createobject('Empty')
		AddProperty(loResult, 'lError', .F.)
		AddProperty(loResult, 'nErrorCode', 0)
		AddProperty(loResult, 'nChangeLength', 0)
		AddProperty(loResult, 'cNewCode', '')
		AddProperty(loResult, 'cReplaceLine', '')
		AddProperty(loResult, 'cTrimmedReplaceLine', '')
		AddProperty(loResult, 'lReplaced', .F.)
		Return loResult
		
	EndProc


	*----------------------------------------------------------------------------------
	Procedure GetReplaceRiskLevel(toObject)

		Local lcMatchType, lnReturn

		lcMatchType = toObject.MatchType

		lnReturn = 4 && Assume everything is very risky to start with !!!

		Do Case

			Case Inlist(lcMatchType, MatchType_Name, MatchType_Constant, '<Parent>', ;
							MATCHTYPE_PROPERTY_DEF, MATCHTYPE_PROPERTY_DESC, MATCHTYPE_PROPERTY_NAME, ;
							MATCHTYPE_PROPERTY, MATCHTYPE_PROPERTY_VALUE, ;
							MATCHTYPE_METHOD_DEF, MATCHTYPE_METHOD_DESC, MatchType_Method ;
							)

				lnReturn = 3

			Case Inlist(lcMatchType, MATCHTYPE_INCLUDE_FILE, '<Expr>', '<Supexpr>', '<Picture>', '<Prompt>', '<Procedure>', ;
							'<Skipfor>', '<Message>', '<Tag>', '<Tag2>');
					Or ;
					toObject.UserField.FileType = 'DBF'

				lnReturn = 2

			Case Inlist(lcMatchType, MatchType_Code, MatchType_Comment) Or;
					(toObject.UserField.IsText And !Inlist(lcMatchType, MatchType_Filename, MatchType_TimeStamp))

				lnReturn = 1

		Endcase

		Return lnReturn
		
	EndProc


	*----------------------------------------------------------------------------------
	Procedure HighlightProcFilter(tcCode, tcProcFilter, tcMatchWordPrefix, tcMatchWordSuffix)

		#Define VISIBLE_AND   '|and|'
		#Define VISIBLE_OR    '|or|'

		#Define AND_DELIMITER Chr[255]
		#Define OR_DELIMITER  Chr[254]

		Local laList[1], laValues[1], lcCode, lcMatch, lcProcFilter, lcValue, lnATC, lnCount, lnFilterCount
		Local lnI, lnJ
		If '|' $ m.tcProcFilter
			lcValue	= Alltrim(Upper(m.tcProcFilter))
			lcValue	= Strtran(m.lcValue, VISIBLE_AND, AND_DELIMITER, 1, 100, 1)
			lcValue	= Strtran(m.lcValue, VISIBLE_OR, OR_DELIMITER, 1, 100, 1)
			lcValue	= Strtran(m.lcValue, '|', OR_DELIMITER, 1, 100, 1)
		Else
			lcValue = m.tcProcFilter
		Endif

		lnFilterCount = Alines(laValues, m.lcValue, 0, OR_DELIMITER, AND_DELIMITER)
		lcCode		  = m.tcCode

		For lnJ = 1 To m.lnFilterCount
			lcProcFilter = m.laValues[m.lnJ]
			lnCount		 = 0
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
					Dimension m.laList[m.lnCount]
					laList[m.lnCount] = m.lcMatch
				Endif
			Endfor

			For lnI = 1 To m.lnCount
				lcMatch	= m.laList[m.lnI]
				lcCode	= Strtran(m.lcCode, m.lcMatch, m.tcMatchWordPrefix + m.lcMatch + m.tcMatchWordSuffix, 1, 1000)
			Endfor
		Endfor

		Return m.lcCode
		
	EndProc


	*----------------------------------------------------------------------------------
	Procedure HtmlEncode(tcCode)

		Local loHTML As 'htmlcode' Of 'mhhtmlcode.prg'
		Local lcHTML

		*-- See: http://www.universalthread.com/ViewPageNewDownload.aspx?ID=9679
		*-- From: Michael Helland - mobydikc@gmail.com

		loHTML = NewObject('htmlcode', 'mhhtmlcode.prg')
		lcHTML = loHTML.PRGToHTML(tcCode)

		Return lcHTML
		
	EndProc


	*----------------------------------------------------------------------------------
	Procedure IncrementProgressBar(tnAmount)

		If Vartype(This.oProgressBar) = 'O'
			This.oProgressBar.nValue = This.oProgressBar.nValue + tnAmount
		Endif
		
	EndProc


	*----------------------------------------------------------------------------------
	Procedure Init(tlPreserveExistingResults)
		
		#Include ..\BuildGoFish.h

		This.cVersion = GOFISH_VERSION  && Comes from include file above
		This.oFSO = Createobject("Scripting.FileSystemObject")
		This.oRegExForSearch = This.GetRegExForSearch()

		If Isnull(This.oRegExForSearch)
			Messagebox('Error creating oRegExForSearch')
			Return .F.
		Endif

		This.oRegExForProcedureStartPositions = This.GetRegExForProcedureStartPositions()
		If Isnull(This.oRegExForProcedureStartPositions)
			Messagebox('Error creating oRegExForProcedureStartPositions')
			Return .F.
		Endif

		This.BuildProjectsCollection()

		This.oSearchOptions = CreateObject(This.cSearchOptionsClass)

		This.oSearchErrors = CreateObject('Collection')
		This.oReplaceErrors = CreateObject('Collection')

		*-- An FFC class used to generate a TimeStamp so the TimeStamp field can be updated when replacing code in a table based file.
		This.oFrxCursor = NewObject('FrxCursor', Home() + '\ffc\_FrxCursor')

		This.PrepareForSearch()
		
	EndProc


	*----------------------------------------------------------------------------------
	Procedure IsBaseclass(tcString)

		Local lcBaseclasses

		*-- Note: Each word below contains a space at the beginning and end of the word so the final match test
		*-- wil not return .t. for partial matches.

		Text To lcBaseclasses Noshow
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

		Return  Upper((' ' + Alltrim(tcString) + ' ')) $ Upper(lcBaseclasses)
		
		
	EndProc


	*----------------------------------------------------------------------------------
	Procedure IsComment(tcLine)

		Local lcLine, lnCount, loMatches, loRegEx, llReturn

		llReturn = This.IsFullLineComment(tcLine)

		If llReturn = .T.
			Return .T.
		Endif

		*-- Look for a match BEFORE any && comment characters
		lnCount = Atc('&' + '&', tcLine)

		If lnCount > 0
			lcLine = Left(tcLine, lnCount - 1)
			loMatches = This.oRegExForSearch.Execute(lcLine)

			If loMatches.Count > 0
				Return .F.
			Endif
		Else
			Return .F.
		Endif
		
	EndProc


	*----------------------------------------------------------------------------------
	Procedure IsContinuation(lcLine)

		Local lnAT
		
		If ';' = Right(Alltrim(m.lcLine, 1, ' ', Tab, CR, lf), 1)
			Return .T.
		Else
			lnAT = Rat('&' + '&', m.lcLine)
			If m.lnAT > 0 And  Right(Alltrim(Left(m.lcLine, m.lnAT - 1), 1, ' ', Tab), 1) = ';'
				Return .T.
			Endif
		Endif
		Return .F.
		
	EndProc


	*----------------------------------------------------------------------------------
	Procedure IsFileTypeIncluded(tcFileType)

		Local lcFileType, llReturn, loOptions

		lcFileType = Upper(tcFileType)
		loOptions = This.oSearchOptions

		If loOptions.lIncludeAllFileTypes
			Return .T.
		Endif

		If !Empty(Justext(This.oSearchOptions.cFileTemplate))
			Return This.MatchTemplate(tcFileType, Justext(This.oSearchOptions.cFileTemplate))
		Endif

		Do Case
			*-- Table-based Files --------------------------------------
			Case lcFileType = 'SCX' And loOptions.lIncludeSCX
				llReturn = .T.
			Case lcFileType = 'VCX' And loOptions.lIncludeVCX
				llReturn = .T.
			Case lcFileType = 'FRX' And loOptions.lIncludeFRX
				llReturn = .T.
			Case lcFileType = 'DBC'And loOptions.lIncludeDBC
				llReturn = .T.
			Case lcFileType = 'MNX' And loOptions.lIncludeMNX
				llReturn = .T.
			Case lcFileType = 'LBX' And loOptions.lIncludeLBX
				llReturn = .T.
			Case lcFileType = 'PJX' And loOptions.lIncludePJX
				llReturn = .T.

			*-- Code based files ----------------------------------
			Case lcFileType = 'PRG' And loOptions.lIncludePRG
				llReturn = .T.
			Case lcFileType = 'SPR' And loOptions.lIncludeSPR
				llReturn = .T.
			Case lcFileType = 'MPR' And loOptions.lIncludeMPR
				llReturn = .T.
			Case 'HTM' $ lcFileType And loOptions.lIncludeHTML
				llReturn = .T.
			Case lcFileType = 'H' And loOptions.lIncludeH
				llReturn = .T.
			Case lcFileType = 'ASP' And loOptions.lIncludeASP
				llReturn = .T.
			Case lcFileType = 'INI' And loOptions.lIncludeINI
				llReturn = .T.
			Case lcFileType = 'JAVA' And loOptions.lIncludeJAVA
				llReturn = .T.
			Case lcFileType = 'JSP' And loOptions.lIncludeJSP
				llReturn = .T.
			Case lcFileType = 'XML' And loOptions.lIncludeXML
				llReturn = .T.
			Case lcFileType = 'TXT' And loOptions.lIncludeTXT
				llReturn = .T.

			*-- Lastly, is it match with other includes???
			Case (lcFileType $ Upper(loOptions.cOtherIncludes)) And !Empty(loOptions.cOtherIncludes)
				llReturn = .T.
			Otherwise
				llReturn = .F.
		Endcase

		Return llReturn
		
	EndProc


	*----------------------------------------------------------------------------------
	Procedure IsFullLineComment(tcLine)

		*-- See if the entire line is a comment
		Local loMatches, loRegEx

		loRegEx = This.GetRegExForSearch()
		loRegEx.Pattern = '^\s*(\*|NOTE|&' + '&)'

		loMatches = loRegEx.Execute(tcLine)

		If loMatches.Count > 0
			Return .T.
		Else
			Return .F.
		Endif
		
	EndProc


	*----------------------------------------------------------------------------------
	Procedure IsTextFile(tcFile)

		Local lcExt, llIsTextFile

		If Empty(tcFile)
			Return .F.
		Endif

		lcExt = Upper(Justext(tcFile))

		llIsTextFile = !(lcExt $ This.cTableExtensions)

		Return llIsTextFile
		
	EndProc


	*----------------------------------------------------------------------------------
	Procedure LoadOptions(tcFile)

		Local loMy As 'My' Of 'My.vcx'
		Local laProperties[1], lcProperty

		If !File(tcFile)
			Return .F.
		Endif

		*-- Get an array of properties that are on the SearchOptions object
		Amembers(laProperties, This.oSearchOptions, 0, 'U')

		*-- Load settings from file...
		loMy = NewObject('My', 'My.vcx')
		loMy.Settings.Load(tcFile)

		*--- Scan over Object properties, and look for a corresponding props on the My Settings object (if present)
		With loMy.Settings
			For x = 1 To Alen(laProperties)
				lcProperty = laProperties[x]
				If Type('.' + lcProperty) <> 'U'
					Store Evaluate('.' + lcProperty) To ('This.oSearchOptions.' + lcProperty)
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
		
	EndProc


	*----------------------------------------------------------------------------------
	Procedure lReadyToReplace_Access
		
		Local llReturn
	
		llReturn = This.nMatchLines > 0 ;
			And (!Empty(This.oSearchOptions.cReplaceExpression) Or This.oSearchOptions.lAllowBlankReplace) ;
			And !This.lFileHasBeenEdited

		Return llReturn
		
	EndProc


	*----------------------------------------------------------------------------------
	Procedure lTimeStampDataProvided_Access

		If This.oSearchOptions.lTimeStamp And !Empty(This.oSearchOptions.dTimeStampFrom) And !Empty(This.oSearchOptions.dTimeStampTo) && If both dates are supplied
			Return .T.
		Else
			Return .F.
		Endif
		
	EndProc

	*----------------------------------------------------------------------------------
	Procedure MatchTemplate(tcString, tcTemplate)

		*-- Supports normal wildcard matching with * and ?, just like old DOS matching.

		Local lcString, lcTemplate, llMatch, lnLength

		If Empty(tcTemplate) Or tcTemplate = '*'
			Return .T.
		Endif

		lcString = Upper(Alltrim(Juststem(tcString)))
		lcTemplate = Upper(Alltrim((tcTemplate)))

		llMatch = Like(lcTemplate, lcString)

		Return llMatch


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
		
	EndProc


	*----------------------------------------------------------------------------------
	*-- Migrate any exitisting Replace Detail Table up to ver 4.3.022 ----
	Procedure MigrateReplaceDetailTable
	
		Local lcCsr, lcDataType, lcFieldName, lcTable, llSuccess, lnSelect

		lcTable = This.cReplaceDetailTable
		lcCsr = 'csrGF_ReplaceSchemaTest'

		If File(lcTable)
			lnSelect = Select()
			Select * From (lcTable) Where 0 = 1 Into Cursor &lcCsr

			*** JRN 11/09/2015 : add field ProcEnd if not already there
			This.addfieldtoreplacetable(lcTable, lcCsr, 'ProcEnd', 'I')
			This.addfieldtoreplacetable(lcTable, lcCsr, 'ProcCode', 'M')
			This.addfieldtoreplacetable(lcTable, lcCsr, 'Statement', 'M')
			This.addfieldtoreplacetable(lcTable, lcCsr, 'StatementStart', 'I')
			This.addfieldtoreplacetable(lcTable, lcCsr, 'FirstMatchInStatement', 'L')
			This.addfieldtoreplacetable(lcTable, lcCsr, 'FirstMatchInProcedure', 'L')

			Use In &lcCsr
			Select (lnSelect)
		Endif
		
	EndProc


	*----------------------------------------------------------------------------------
	Procedure OpenTableForReplace(tcFileToOpen, tcCursor, tnResultId)

		Local llReturn, lnSelect

		lnSelect = Select()

		If Used(tcCursor)
			Use In (tcCursor)
		Endif

		Select 0

		Try
			Use (tcFileToOpen) Exclusive Alias (tcCursor)
			llReturn = .T.
		Catch
			This.SetReplaceError('Cannot open file for exclusive use: ' + Chr(13) + Chr(13), tcFileToOpen, tnResultId)
			Select (lnSelect)
			llReturn = .F.
		Endtry

		Return llReturn
		
	EndProc


	*----------------------------------------------------------------------------------
	Procedure PrepareForSearch
	
		Clear Typeahead

		This.lEscPress = .F.
		This.lFileNotFound = .F.
		This.nMatchLines = 0
		This.nFileCount = 0
		This.nFilesProcessed = 0
		This.nSearchTime = 0
		This.lResultsLimitReached = .F.

		This.PrepareRegExForSearch()

		This.ClearResultsCursor()
		This.ClearResultsCollection()

		This.ClearReplaceSettings()

		This.oSearchErrors = Createobject('Collection')
		This.oReplaceErrors = Createobject('Collection')
		This.oDirectories = Createobject('Collection')

		This.SetFilesToSkip()
		
	EndProc


	*----------------------------------------------------------------------------------
	Procedure PrepareRegExForReplace
	
		Local lcPattern

		lcPattern = This.oSearchOptions.cEscapedSearchExpression

		*If !this.oSearchOptions.lRegularExpression
			*-- Need to trim off the pre- and post- wild card characters so we can get back to just the search phrase
			If Left(lcPattern, 2) = '.*'
				lcPattern = Substr(lcPattern, 3)
			Endif

			If Right(lcPattern, 2) = '.*'
				lcPattern = Left(lcPattern, Len(lcPattern) - 2)
			Endif

		*EndIf

		This.oRegExForSearch.Pattern = lcPattern
		
	EndProc


	*----------------------------------------------------------------------------------
	Procedure PrepareRegExForSearch
	
		Local lcPattern, lcRegexPattern, lcSearchExpression, loRegEx

		loRegEx = This.oRegExForSearch
		lcSearchExpression = This.oSearchOptions.cSearchExpression

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

			This.oSearchOptions.cEscapedSearchExpression = lcPattern

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

		Local lcCode, lcComment, lcMatchType, lcTrimmedMatchLine, lnCount, loCodeMatches, loCommentMatches

		lcTrimmedMatchLine = toObject.TrimmedMatchLine
		lcMatchType = toObject.MatchType

		lnCount = Atc('&' + '&', lcTrimmedMatchLine)

		If lnCount > 0 And This.oSearchOptions.lSearchInComments
			lcCode = Left(lcTrimmedMatchLine, lnCount - 1)
			lcComment = Substr(lcTrimmedMatchLine, lnCount)
			loCodeMatches = This.oRegExForSearch.Execute(lcCode)
			loCommentMatches = This.oRegExForSearch.Execute(lcComment)

			If loCodeMatches.Count > 0 And loCommentMatches.Count > 0
				toObject.MatchType = MatchType_Comment
				This.CreateResult(toObject)
				lcMatchType = toObject.UserField.MatchType && Restore to UserField MatchType for further
			Else
				lcMatchType = Iif(loCommentMatches.Count > 0, MatchType_Comment, toObject.MatchType)
			Endif
		Endif

		toObject.MatchType = lcMatchType
		
	EndProc


	*----------------------------------------------------------------------------------
	Procedure ProcessSearchResult(toObject)

		Local lcMatchType, lcSaveObjectName, loObject
		Local lcBaseClass, lcContainingClass, lcMethodName, lcParentClass, lcSave_Baseclass

		lcMatchType = toObject.UserField.MatchType

		*-- Store these so we can revert back after processing, becuase it's important to reset back
		*-- so any further matches in the code can be processed correctly
		With toObject.UserField
			lcSaveObjectName = ._Name
			lcSave_Baseclass = ._BaseClass
			lcBaseClass = ._BaseClass
			lcMethodName = toObject.MethodName
			lcParentClass = ._ParentClass
			lcContainingClass = .ContainingClass
		Endwith

		If lcMatchType # MatchType_Filename
			loObject = This.AssignMatchType(toObject)
		Else
			loObject = toObject
		Endif

		If !Isnull(loObject)
			This.CreateResult(loObject)
		Endif

		With toObject.UserField
			._Name = lcSaveObjectName
			._BaseClass = lcSave_Baseclass
			._BaseClass = lcBaseClass
			toObject.MethodName = lcMethodName
			._ParentClass = lcParentClass
			.ContainingClass = lcContainingClass
		Endwith
		
	EndProc

	*----------------------------------------------------------------------------------
	Procedure ReduceProgressBarMaxValue(tnReduction)

		Try
			This.oProgressBar.nMaxValue = This.oProgressBar.nMaxValue - tnReduction
		Catch
		Endtry
		
	EndProc

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
	
		Local loMatches, lnX, loMatch, lcRepl

		This.PrepareRegExForSearch()
		This.PrepareRegExForReplace()

		loRegEx = This.oRegExForSearch

		If !Empty(lcRegEx)
			loRegEx.Pattern = lcRegEx
		Endif

		loMatches = loRegEx.Execute(lcSource)

		lnCount = loMatches.Count

		If lnCount = 0
			Return lcSource
		Endif

		lcRepl = lcReplace

		*** Note we have to go last to first to not hose relative string indexes of the match
		For lnX = lnCount - 1 To 0 Step - 1
			loMatch = loMatches.Item(lnX)
			lcMatch = loMatch.Value
			If llIsExpression
				lcRepl = Eval(lcReplace) &&Evaluate dynamic expression each time
			Endif
			lcSource = Stuff(lcSource, loMatch.FirstIndex + 1, loMatch.Length, lcRepl)
		Endfor

		Return lcSource
		
	EndProc


	*----------------------------------------------------------------------------------
	Procedure RenameColumn(tcTable, tcOldFieldName, tcNewFieldName)

		Local lcAlias

		lcAlias = Juststem(tcTable)

		If Empty(Field(tcNewFieldName, lcAlias)) And !Empty(Field(tcOldFieldName, lcAlias))
			Try
				Alter Table (lcAlias) Rename Column (tcOldFieldName) To (tcNewFieldName)
			Catch
			Endtry
		Endif
		
	EndProc


	*----------------------------------------------------------------------------------
	Procedure ReplaceFromCurrentRow(tcCursor, tcReplaceLine)

		Local lcColumn, lcFileToModify, lnCurrentRecno, lnMatchStart, lnProcStart, lnResultRecno, lnSelect
		Local loReplace, loResult

		lnSelect = Select()
		Select (tcCursor)
		lnCurrentRecno = Recno()

		If Process = .F. && Could be that the row was previous marked for replace, and now it has been cleared.
			Replace ReplaceLine With '' In (tcCursor)
			Replace TrimmedReplaceLine With '' In (tcCursor)
			Select(lnSelect)
			Return GF_REPLACE_RECORD_IS_NOT_MARKED_FOR_REPLACE
		Endif

		If Replaced = .T. && If it's already been processed
			Select(lnSelect)
			Return GF_REPLACE_FILE_HAS_ALREADY_BEEN_PROCESSED
		Endif

		If !File(FilePath)
			This.SetReplaceError('File not found:', FilePath, Id)
			Select(lnSelect)
			Return GF_REPLACE_FILE_NOT_FOUND
		Endif

		*!*	If !Empty(tcReplaceLine) && If doing a "Replace Line", then the regular History
		*!*	 This.CreateReplaceHistoryRecord()
		*!*	Endif

		If This.oSearchOptions.lBackup
			llBackedUp = This.BackupFile(FilePath, This.nReplaceHistoryId)
			If !llBackedUp
				Select(lnSelect)
				Return GF_REPLACE_BACKUP_ERROR
			Endif
		Endif

		This.PrepareRegExForSearch() && This will setup the Search part of the RegEx
		This.PrepareRegExForReplace() && This will setup the Replace part of the RegEx

		Scatter Name loReplace Memo

		If This.IsTextFile(FilePath)
			loResult = This.ReplaceInTextFile(loReplace, tcReplaceLine)
		Else
			loResult = This.ReplaceInTable(loReplace, tcReplaceLine)
		Endif

		If !loResult.lError
		*-- We must update all match result rows that are of the same source line as this row.
		*-- The reason is that search matches can result in multiple rows, and we can't process them again.
			lcFileToModify = FilePath
			lnResultRecno = Recno
			lnProcStart = ProcStart
			lnMatchStart = MatchStart
			lcColumn = Column

	  Update  (tcCursor) ;
		  Set TrimmedReplaceLine = loResult.cTrimmedReplaceLine, ;
			  ReplaceLine = loResult.cReplaceLine ;
		  Where FilePath == lcFileToModify And ;
			  Recno = lnResultRecno And ;
			  Column = lcColumn And ;
			  MatchStart = lnMatchStart

			Try
				Goto (lnCurrentRecno)
			Catch
			Endtry

			If loResult.lReplaced  = .T.
				This.nReplaceCount = This.nReplaceCount + 1
			Endif

			*-- Removed this in 4.3.014. We *do* need to re-compile.
			*If !Empty(tcReplaceLine) 
				This.Compile(FilePath)
			*Endif

			This.UpdateCursorAfterReplace(tcCursor, loResult)

			lnReturn = GF_REPLACE_SUCCESS

		Else

			lnReturn = 	loResult.nErrorCode

		Endif

		Select (lnSelect)

		Return lnReturn
		
	EndProc


	*----------------------------------------------------------------------------------
	Procedure ReplaceInCode(toReplace, tcReplaceLine)

		* tcReplaceLine, if passed, will be used to replace the entire oringinal match line,
		* rather than using the RexEx replace with cReplaceExpression on the original line.

		* Notes:
		* For the Replace, the pattern on the regex must already be set (use PrepareRegExForReplace)
		* Note: Unless a full replacement line is passed in tcReplaceLine, ALL instances of the pattern will be replaced on the tcMatchLine

		Local lcLeft, lcLineFromFile, lcNewCode, lcReplaceExpression, lcReplaceLine, lcRight
		Local lcCode, lcMatchLine, lnLineToChangeLength, lnMatchStart, loRegEx, loResult

		loResult = This.GetReplaceResultObject()
		lcCode = toReplace.Code

		lcMatchLine = Left(toReplace.MatchLine, toReplace.MatchLen)

		lnMatchStart = toReplace.MatchStart
		lnLineToChangeLength = Len(lcMatchLine)

		lcLineFromFile = Substr(lcCode, lnMatchStart + 1, lnLineToChangeLength)

		If lcLineFromFile != lcMatchLine && Ensure that line from file still matches the passed in line from the orginal search!!
			This.SetReplaceError('Source file has changed since original search:', Alltrim(toReplace.FilePath), toReplace.Id)
			loResult.lError = .T.
			Return loResult
		Endif

		lcLeft = Left(lcCode, lnMatchStart)

		*-- IMPORTANT CODE HERE... Revised code line is determined here!!!! -------------
		If Empty(tcReplaceLine)
			loRegEx = This.oRegExForSearch
			lcReplaceExpression = This.oSearchOptions.cReplaceExpression
			Do Case
			Case This.nReplaceMode = 1
				lcReplaceLine = loRegEx.Replace(lcMatchLine, lcReplaceExpression)
			Case This.nReplaceMode = 2
				lcReplaceLine = ''
			Case This.nReplaceMode = 3 And !Empty(This.cReplaceUDFCode)
				lcReplaceLine = This.ReplaceLineWithUDF(lcMatchLine)
			Otherwise
				lcReplaceLine = lcMatchLine
			Endcase
		Else
			lcReplaceLine = tcReplaceLine
		Endif

		lcRight = Substr(lcCode, lnMatchStart + 1 + lnLineToChangeLength)

		*--Added this in 4.3.014 to handle case of deleting the entire line
		If Empty(lcReplaceLine)
			lcRight = Ltrim(lcRight, 0, Chr(10)) && Need to strip off initial Chr(10) of Right hand code block
		Endif

		lcNewCode = lcLeft + lcReplaceLine + lcRight

		With loResult
			.nChangeLength = Len(lcReplaceLine) - Len(lcMatchLine)
		*--Added this in 4.3.014 to handle case of deleting the entire line
			If Empty(lcReplaceLine)
				.nChangeLength = .nChangeLength - 1 && to account for the Chr(10) we stripped off above
			Endif
			.cNewCode = lcNewCode
			.cReplaceLine = lcReplaceLine
			.cTrimmedReplaceLine = This.TrimWhiteSpace(.cReplaceLine)
		Endwith

		toReplace.ReplaceLine = loResult.cReplaceLine
		toReplace.TrimmedReplaceLine = loResult.cTrimmedReplaceLine

		Return loResult
		
	EndProc


	*----------------------------------------------------------------------------------
	Procedure ReplaceInTable(toReplace, tcReplaceLine)

		Local lcColumn, lcFileToModify, lcMatchLine, lcReplaceCursor, llTableWasOpened, lnMatchStart
		Local lnRecNo, lnResultId, lnSelect, loResult

		lcFileToModify = Alltrim(toReplace.FilePath)
		lcMatchLine = Left(toReplace.MatchLine, toReplace.MatchLen)
		lnMatchStart = toReplace.MatchStart
		lnResultId = toReplace.Id
		lcColumn = Alltrim(toReplace.Column)
		lnRecNo = toReplace.Recno

		lcReplaceCursor = 'ReplaceCursor'
		lnSelect = Select()

		loResult = This.GetReplaceResultObject()

		*!*	If !File(lcFileToModify)
		*!*		This.SetReplaceError('File not found:', lcFileToModify, lnResultId)
		*!*		loResult.lError = .t.
		*!*		loResult.nErrorCode = GF_REPLACE_FILE_NOT_FOUND
		*!*	Endif

		If !This.OpenTableForReplace(lcFileToModify, lcReplaceCursor, lnResultId)
			loResult.lError = .T.
			loResult.nErrorCode = GF_REPLACE_UNABLE_TO_USE_TABLE_FOR_REPLACE
		Else
			llTableWasOpened = .T.
		Endif

		If !loResult.lError
			Try
				Goto lnRecNo
			Catch
				This.SetReplaceError('Error locating record in file:', lcFileToModify, lnResultId)
				loResult.lError = .T.
				loResult.nErrorCode = GF_REPLACE_ERROR_LOCATING_RECORD_IN_FILE
			Endtry
		Endif

		If !loResult.lError
			toReplace.Code = Evaluate(lcReplaceCursor + '.' + lcColumn)
			loResult = This.ReplaceInCode(toReplace, tcReplaceLine)
		Endif

		*-- Big step here... Replace code in actual record!!! (If not in Preview Mode)
		If !loResult.lError And This.oSearchOptions.lPreviewReplace = .F.
			Replace (lcColumn) With loResult.cNewCode In (lcReplaceCursor) && Update code in table
			If Type('timestamp') != 'U'
				Replace Timestamp With This.oFrxCursor.getFrxTimeStamp() In (lcReplaceCursor)
			Endif
			loResult.lReplaced  = .T.
			This.CreateReplaceDetailRecord(toReplace)
		Endif

		If llTableWasOpened
			Use && Close the table based file we opened above
		Endif

		Select (lnSelect)

		Return loResult
		
	EndProc


	*----------------------------------------------------------------------------------
	Procedure ReplaceInTextFile(toReplace, tcReplaceLine)

		Local lcFileToModify

		lcFileToModify = Alltrim(toReplace.FilePath)

		Local lcOldCode, loReseult, loResult

		*!*	If !File(lcFileToModify)
		*!*		This.SetReplaceError('File not found:', lcFileToModify, lnResultId)
		*!*		loResult = This.GetReplaceResultObject()
		*!*		loResult.lError = .t.
		*!*		Return loResult
		*!*	EndIf

		toReplace.Code = Filetostr(lcFileToModify)
		loResult = This.ReplaceInCode(toReplace, tcReplaceLine)

		If loResult.lError Or This.oSearchOptions.lPreviewReplace = .T.
			Return loResult
		Endif

		*== Big step here... About to replace old file with the new code!!!
		Try
			If !Empty(loResult.cNewCode) && Do not dare replace the file with and empty string. Something must be wrong!
				Strtofile(loResult.cNewCode, lcFileToModify, 0)
				loResult.lReplaced  = .T.
				This.CreateReplaceDetailRecord(toReplace)
			Endif
		Catch
			This.SetReplaceError('Error saving file: ', lcFileToModify, toReplace.Id)
		Endtry

		Return loResult
		
	EndProc


	*----------------------------------------------------------------------------------
	Procedure ReplaceLine(tcCursor, tnID, tcReplaceLine)

		Local lnSelect, lnLastChar, llReturn

		lnSelect = Select()

		lcReplaceLine = tcReplaceLine
		lnLastChar = Asc(Right(lcReplaceLine, 1))

		If lnLastChar = 10 && Editbox will add a Chr(10) so this has to be stripped off
			lcReplaceLine = Left(lcReplaceLine, Len(lcReplaceLine) - 1)
		Endif

		lnLastChar = Asc(Right(lcReplaceLine, 1))

		If lnLastChar <> 13 And !Empty(lcReplaceLine) && Make sure user has not stripped of the Chr(13) that came with the MatchLine
			lcReplaceLine = lcReplaceLine + Chr(13)
		Endif

		Select(tcCursor)
		Locate For Id = tnID

		If Found()

			If Replaced = .T.
				Return .T.
			Endif

			Replace Process With .T. In (tcCursor)

			This.CreateReplaceHistoryRecord()

			lnReturn = This.ReplaceFromCurrentRow(tcCursor, lcReplaceLine)

			If lnReturn >= 0
				This.UpdateReplaceHistoryRecord()
				llReturn = .T.
			Else
				llReturn = .F.
			Endif

		Else

			This.SetReplaceError('Error locating record in call to ReplaceLine() method.', '', tnID)
			llReturn = .F.

		Endif

		Return llReturn
		
	EndProc


	*----------------------------------------------------------------------------------
	Procedure ReplaceLineWithUDF(tcMatchLine)

		Local lcMatchLine, lcReplaceLine, llCR

		lcMatchLine = tcMatchLine

		*-- If there is a CR at the end, pull it off before calling the UDF. Will add back later...
		If Right(tcMatchLine, 1) = Chr(13)
			llCR = .T.
			lcMatchLine = Left(tcMatchLine, Len(tcMatchLine) - 1)
		Endif

		*-- Call the UDF ---------------
		Try
			lcReplaceLine = Execscript(This.cReplaceUDFCode, lcMatchLine)
		Catch
			lcReplaceLine = lcMatchLine && Keep the line the same if UDF failed
		Finally
		Endtry

		If Vartype(lcReplaceLine) <> 'C'
			lcReplaceLine = lcMatchLine
		Endif

		If llCR
			lcReplaceLine = lcReplaceLine + Chr(13)
		Endif

		Return lcReplaceLine
		
	EndProc


	*----------------------------------------------------------------------------------
	Procedure ReplaceMarkedRows(tcCursor)

		Local lcFile, lcFileList, lcLastFile, lcReplaceExpression, lcSearchExpression, lnResult, lnSelect
		Local lnShift

		This.nReplaceCount = 0
		This.nReplaceFileCount = 0

		lcSearchExpression = Alltrim(This.oSearchOptions.cSearchExpression)
		lcReplaceExpression = Alltrim(This.oSearchOptions.cReplaceExpression)
		lnShift = Len(lcReplaceExpression) - Len(lcSearchExpression)

		This.oReplaceErrors = Createobject('Collection')

		If Empty(This.oSearchOptions.cReplaceExpression) And !This.oSearchOptions.lAllowBlankReplace
			This.SetReplaceError('Replace expression is blank, but ALLOW BLANK flag is not set.')
			Return .F.
		Endif

		lnSelect = Select()
		Select (tcCursor)

		This.CreateReplaceHistoryRecord()

		lcLastFile = ''

		Scan

			If Vartype(This.oProgressBar) = 'O'
				This.oProgressBar.nValue = This.oProgressBar.nValue + 1
			Endif

			lnResult = This.ReplaceFromCurrentRow(tcCursor)

		*-- Skip to next file if have had any of there errors:
			If lnResult = GF_REPLACE_BACKUP_ERROR Or;
					lnResult = GF_REPLACE_UNABLE_TO_USE_TABLE_FOR_REPLACE Or ;
					lnResult = GF_REPLACE_FILE_NOT_FOUND
				lcFile = FilePath
				Locate For FilePath <> lcFile Rest
				If !Bof()
					Skip - 1
				Endif
			Endif

			If FilePath <> lcLastFile And !Empty(lcLastFile) && If we are on a new file, then compile the previous file
				This.Compile(lcLastFile)
				lcLastFile = ''
			Endif

			If lnResult = GF_REPLACE_SUCCESS
				lcLastFile = FilePath
			Endif

		Endscan

		*-- Must look at compiling one last time now that loop has ended.
		If FilePath <> lcLastFile And !Empty(lcLastFile) && If we are on a new file, then compile the previous file
			This.Compile(lcLastFile)
		Endif


		This.UpdateReplaceHistoryRecord()

		Select (lnSelect)

		This.ShowWaitMessage('Replace Done.')
		
	EndProc


	*----------------------------------------------------------------------------------
	Procedure RestoreDefaultDir
		
		Cd (This.cInitialDefaultDir)
		
	EndProc


	*----------------------------------------------------------------------------------
	Procedure SaveOptions(tcFile)

		Local loMy As 'My' Of 'My.vcx'
		Local laProperties[1], lcProperty

		loMy = NewObject('My', 'My.vcx')

		Amembers(laProperties, This.oSearchOptions, 0, 'U')

		With loMy.Settings

			For x = 1 To Alen(laProperties)
				lcProperty = laProperties[x]
				If !Inlist(lcProperty, '_MEMBERDATA', 'CPROJECTS')
					.Add(lcProperty, Evaluate('This.oSearchOptions.' + lcProperty))
				Endif
			Endfor

			.Save(tcFile)

		Endwith
		
	EndProc


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
		Update  Results ;
			Set firstmatchinstatement = .T. ;
			From (This.cSearchResultsAlias)    As  Results ;
			Join (Select  FilePath, ;
					   Class, ;
					   Name, ;
					   MethodName, ;
					   statementstart, ;
					   Min(MatchStart) As  MatchStart ;
				   From (This.cSearchResultsAlias)            ;
				   Group By FilePath, Class, Name, MethodName, statementstart) ;
			 As  FirstMatch ;
			 On Results.FilePath + Results.Class + Results.Name + Results.MethodName ;
			 	= FirstMatch.FilePath + FirstMatch.Class + FirstMatch.Name + FirstMatch.MethodName ;
			 And Results.statementstart = FirstMatch.statementstart ;
			 And Results.MatchStart = FirstMatch.MatchStart

		 Update  Results ;
			 Set firstmatchinprocedure = .T. ;
			 From (This.cSearchResultsAlias) As  Results ;
			 Join (Select  FilePath, ;
						   Class, ;
						   Name, ;
						   MethodName, ;
						   Min(MatchStart) As  MatchStart ;
					   From (This.cSearchResultsAlias) ;
					   Group By FilePath, Class, Name, MethodName) As  FirstMatch ;
				 On Results.FilePath + Results.Class + Results.Name + Results.MethodName ;
				 	= FirstMatch.FilePath + FirstMatch.Class + FirstMatch.Name + FirstMatch.MethodName ;
				 And Results.MatchStart = FirstMatch.MatchStart

		Select (m.tnSelect)
		
	EndProc


	*----------------------------------------------------------------------------------
	Procedure SearchInCode(tcCode, tuUserField, tlHasProcedures)

		Local loObject As 'GF_SearchResult'
		Local lcErrorMessage, llScxVcx, lnMatchCount, lnSelect, loMatch, loMatches, loProcedure
		Local loProcedureStartPositions, lcMatchType

		lnSelect = Select()

		If Empty(tcCode)
			Return 0
		Endif
		*-- Be sure that oRegExForSearch has been setup... Use This.PrepareRegExForSearch() or roll-your-own
		Try
			loMatches = This.oRegExForSearch.Execute(tcCode)
		Catch
		Endtry

		If Type('loMatches') = 'O'
			lnMatchCount = loMatches.Count
		Else
			lcErrorMessage = 'Error processing regular expression.    ' + This.oRegExForSearch.Pattern
			This.SetSearchError(lcErrorMessage)
			Return - 1
		Endif

		If lnMatchCount > 0

			loProcedureStartPositions = Iif(tlHasProcedures, This.GetProcedureStartPositions(tcCode), .Null.)

			For Each loMatch In loMatches FoxObject

				If tlHasProcedures And !This.oSearchOptions.lSearchInComments And This.IsComment(loMatch.Value)
					Loop
				Endif
				loProcedure = This.FindProcedureForMatch(loProcedureStartPositions, loMatch.FirstIndex)
				loObject = Createobject('GF_SearchResult')

				With loObject
					.UserField = tuUserField
					.oMatch = loMatch
					.oProcedure = loProcedure

					.Type = Proper(.oProcedure.Type)
					* .ContainingClass =	.oProcedure._ClassName	&& Not used on this object. This line to be deleted after testing. (2012-07-11))
					.MethodName = .oProcedure._Name
					.ProcStart = .oProcedure.StartByte
					.procend = .oProcedure.EndByte
					.proccode = Substr(tcCode, .ProcStart + 1, Max(0, .procend - .ProcStart))

					.MatchLine = .oMatch.Value
					.MatchStart = .oMatch.FirstIndex
					.MatchLen = Len(.oMatch.Value)

					If tlHasProcedures
						.MatchType = loProcedure.Type && Use what was determined by call to FindProcedureForMatch())
						tuUserField.MatchType = loProcedure.Type
					Else
						.MatchType = tuUserField.MatchType && Use what was passed.
					Endif

					.Code = Iif(This.oSearchOptions.lStoreCode, tcCode, '')
				Endwith

				*	Assert Upper(JustExt(Trim(loobject.uSERFIELD.FILENAME)))  # 'PRG' 

				This.FindStatement(loObject)

				This.ProcessSearchResult(loObject)

			Endfor

		Endif

		Select(lnSelect)

		Return lnMatchCount
		
	EndProc


	*----------------------------------------------------------------------------------
	Procedure SearchInFile(tcFile, tlForce)

		*-- Only searches passed file if its file ext is marked for inclusion (i.e. lIncludeSCX)
		*-- Optionally, pass tlForce = .t. to force the file to be searched.
		Local lnMatchCount

		*!* ** { JRN -- 11/18/2015 12:16 PM - Begin
		*!* If Lastkey() = 27 or Inkey() = 27
		If Inkey() = 27
		*!* ** } JRN -- 11/18/2015 12:16 PM - End
			This.lEscPress = .T.
			Clear Typeahead
			Return 0
		Endif

		*-- See if the filename matches the File template filter (if one is set) ----
		If !Empty(This.oSearchOptions.cFileTemplate)
			If !This.MatchTemplate(Justfname(tcFile), Juststem(Justfname(This.oSearchOptions.cFileTemplate)))
				This.ReduceProgressBarMaxValue(1)
				Return 0
			Endif
		Endif

		If This.FilesToSkip(tcFile)
			Return 0
		Endif

		If !This.IsFileTypeIncluded(Justext(tcFile)) And !tlForce
			*This.ReduceProgressBarMaxValue(1)
			Return 0
		Endif

		If !File(tcFile)
			This.lFileNotFound = .T.
			This.SetSearchError('File not found: ' + tcFile)
			*This.ReduceProgressBarMaxValue(1)
			Return 0
		Endif

		This.ShowWaitMessage('Processing file: ' + tcFile)

		*-- Look for a match on the file name ----------------------
		lnFileNameMatchCount = This.SearchInFileName(tcFile)

		If lnFileNameMatchCount < 0
			Return lnFileNameMatchCount
		Endif

		llTextFile = This.IsTextFile(tcFile)

		*-- Do not search inside of file if we are only looking at timestamps and have and empty string
		If llTextFile And This.oSearchOptions.lTimeStamp And Empty(This.oSearchOptions.cSearchExpression)
			This.nFilesProcessed = This.nFilesProcessed + 1
			This.nFileCount = This.nFileCount + 1
			Return lnFileNameMatchCount
		Endif

		*-- Look for a match within the file contents ----------------------
		If llTextFile
			lnMatchCount = This.SearchInTextFile(tcFile)
		Else
			lnMatchCount = This.SearchInTable(tcFile)
		Endif

		This.nFilesProcessed = This.nFilesProcessed + 1

		If lnMatchCount < 0
			Return lnMatchCount
		Endif

		*-- Count number of files that had a match by either search above ---
		If lnMatchCount > 0 Or lnFileNameMatchCount > 0
			This.nFileCount = This.nFileCount + 1
		Endif

		Return lnMatchCount + lnFileNameMatchCount
		
	EndProc


	*----------------------------------------------------------------------------------
	Procedure SearchInFileName(tcFile)

		Local loFileResultObject As 'GF_FileResult'
		Local loSearchResultObject As 'GF_SearchResult'
		Local lcCode, ldFileDate, lnMatchCount, lnSelect, llHasMethods

		lnSelect = Select()

		If !File(tcFile)
			This.lFileNotFound = .T.
			This.SetSearchError('File not found: ' + tcFile)
			Return 0
		Endif

		ldFileDate = This.GetFileDateTime(tcFile)

		ldFromDate = Evl(This.oSearchOptions.dTimeStampFrom, {^1900-01-01})
		ldToDate = Evl(This.oSearchOptions.dTimeStampTo, {^9999-01-01})
		ldToDate = ldToDate + 1 &&86400 && Must bump into to next day, since TimeStamp from table has time on it

		If This.oSearchOptions.lTimeStamp And !Between(ldFileDate, ldFromDate, ldToDate)
			Return 0
		Endif

		*-- Be sure that oRegExForSearch has been setup... Use This.PrepareRegExForSearch() or roll-your-own
		Try
			loMatches = This.oRegExForSearch.Execute(Justfname(tcFile))
		Catch
		Endtry

		If Type('loMatches') = 'O'
			lnMatchCount = loMatches.Count
		Else
			lcErrorMessage = 'Error processing regular expression.    ' + This.oRegExForSearch.Pattern
			This.SetSearchError(lcErrorMessage)
			Return - 1
		Endif

		If lnMatchCount = 0 And !Empty(This.oSearchOptions.cSearchExpression)
			Return 0
		Endif

		loFileResultObject = Createobject('GF_FileResult')	&& This custom class has all the properties that must be populated if you want to
															&& have a cursor created
		With loFileResultObject
			.FileName = Justfname(tcFile)
			.FilePath = tcFile
			.MatchType = MatchType_Filename
			.FileType = Upper(Justext(tcFile))
			.Timestamp = ldFileDate

			.MatchLine = 'File name = "' + .FileName + '"'
			.TrimmedMatchLine = .MatchLine
		Endwith

		loSearchResultObject = Createobject('GF_SearchResult')
		With loSearchResultObject
			.UserField		  = loFileResultObject
			.MatchType		  = MatchType_Filename
			.MatchLine		  = 'File name = "' + loFileResultObject.FileName + '"'
			.TrimmedMatchLine = 'File name = "' + loFileResultObject.FileName + '"'
		Endwith

		If This.IsTextFile(tcFile)
			loSearchResultObject.Code = Filetostr(tcFile)
		Endif

		This.ProcessSearchResult(loSearchResultObject)

		Select (lnSelect)

		Return 1
		
	EndProc


	*----------------------------------------------------------------------------------
	Procedure SearchInPath(tcPath)

		Local lcDirectory, lcFile, lcFileFilter, lcFileName, lnFileCount, lnReturn, lnSelect, j

		lnSelect = Select()

		If Empty(tcPath)
			This.SetSearchError('Path parameter [' + tcPath + '] is empty in call to SearchInPath()')
			Return 0
		Endif

		This.oSearchOptions.cPath = tcPath

		This.StoreInitialDefaultDir()

		If !This.ChangeCurrentDir(tcPath) && If there was a problem CD-ing into the starting path
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

		This.oDirectories = This.GetDirectories(tcPath, This.oSearchOptions.lIncludeSubdirectories)

		If This.lEscPress = .T.
			This.SearchFinished(lnSelect)
			Return 0
		Endif

		Chdir (tcPath) && Must go back, since above call to BuildDirList prolly changed our directory!

		This.StartProgressBar(This.oProgressBar.nMaxValue)
		lnTotalFileCount = 0

		For Each lcDirectory In This.oDirectories

			If This.FilesToSkip(Upper(lcDirectory + '\-'))
				Loop
			Endif


			lcFileFilter = Addbs(lcDirectory) + '*.*'

			If Adir(laTemp, lcFileFilter) = 0 && 0 means no files in the Dir
				Loop
			Endif

			Asort(laTemp)

			lnFileCount = Alen(laTemp) / 5 && The number of files that match the filter criteria for this pass

			For j = 1 To lnFileCount
				lcFileName = laTemp(j, 1) && Just the name and ext, no path info
				lcFile = Addbs(lcDirectory) + lcFileName && path + filename
				lnReturn = This.SearchInFile(lcFile)

				lnTotalFileCount = lnTotalFileCount + 1
				This.UpdateProgressBar(lnTotalFileCount)

				If lnReturn < 0 Or This.lEscPress = .T. Or This.nMatchLines >= This.oSearchOptions.nMaxResults
					Exit
				Endif
			Endfor

			If lnReturn < 0 Or This.lEscPress = .T. Or This.nMatchLines >= This.oSearchOptions.nMaxResults
				Exit
			Endif
		Endfor

		This.SearchFinished(lnSelect)

		This.RestoreDefaultDir()

		If lnReturn >= 0
			Return 1
		Else
			Return lnReturn
		Endif
		
	EndProc


	*----------------------------------------------------------------------------------
	Procedure SearchInProject(tcProject)

		Local laProjectFiles[1], lcFile, lcProjectAlias, lcProjectPath, lnReturn, lnSelect, lnX

		lnSelect = Select()

		lcProjectPath = Addbs(Justpath(Alltrim(tcProject)))
		lcProjectAlias = 'GF_ProjectSearch'

		This.oSearchOptions.cProject = tcProject

		If Empty(tcProject)
			This.SetSearchError('Project parameter [' + tcProject + '] is empty in call to SearchInProject().')
			Return 0
		Endif

		If !File(tcProject)
			This.SetSearchError('Project file [' + tcProject + '] not found in call to SearchInProject().')
			Return 0
		Endif

		Try && Attempt to open Project.PJX in a cursor...
			Select 0
			Use (tcProject) Again Shared Alias &lcProjectAlias
			lnReturn = 1
		Catch
			lnReturn = -2
		Endtry

		If lnReturn = -2
			This.SetSearchError('Cannot open project file[' + tcProject + ']')
			This.SearchFinished(lnSelect)
			Return lnReturn
		Endif

	 	Select  Name, Type ;
			From (lcProjectAlias) ;
			Where Type $ 'EHKMPRVBdTxD' And ;
				 Not Deleted() ;
				 And !(Upper(Justext(Name)) $ This.cGraphicsExtensions) ;
			 Order By Type ;
			 Into Array laProjectFiles

		If Type('laProjectFiles') = 'L'
			This.SearchFinished(lnSelect)
			Return 1
		Endif

		Use In Alias (lcProjectAlias)

		This.PrepareForSearch()
		This.StartTimer()
		This.StartProgressBar(Alen(laProjectFiles) / 2.0)

		For lnX = 1 To Alen(laProjectFiles) Step 2

			lcFile = laProjectFiles(lnX)
			lcFile = Fullpath(lcFile, lcProjectPath)
			lcFile = Strtran(lcFile, Chr(0), '') && Strip out junk char from the end

			If This.oSearchOptions.lLimitToProjectFolder
				If !(Upper(lcProjectPath) $ Upper(Addbs(Justpath(lcFile))))
					Loop
				Endif
			Endif

			lnReturn = This.SearchInFile(lcFile)

			This.UpdateProgressBar(This.nFilesProcessed)

			If (lnReturn < 0) Or This.lEscPress = .T. Or This.nMatchLines >= This.oSearchOptions.nMaxResults
				Exit
			Endif

		Endfor

		This.SearchFinished(lnSelect)

		If lnReturn >= 0
			Return 1
		Else
			Return lnReturn
		Endif
		
	EndProc


	*----------------------------------------------------------------------------------
	Procedure SearchInTable(tcFile)

		Local loFileResultObject As 'GF_FileResult'
		Local loSearchResultObject As 'GF_SearchResult'
		Local laMaxTimeStamp[1], laParent[1], lcClass, lcCode, lcDataType, lcDeleted, lcExt, lcField
		Local lcFieldSource, lcFormClass, lcFormClassloc, lcFormName, lcName, lcObjectType, lcParent
		Local lcParentName, lcProject, lcSearchExpression, ldFromDate, ldMaxTimeStamp, ldToDate
		Local llContinueError, llHasMethods, llLocateError, llProcessThisMatch, llScxVcx, lnEndColumn
		Local lnMatchCount, lnParentId, lnSelect, lnStart, lnStartColumn, lnTotalMatches, loException
		Local ii, llIgnorePropertiesField

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
					lcFormClassloc = ClassLoc
				Endif
			Endif

		Endif

		If This.oSearchOptions.lTimeStamp And Type('timestamp') = 'U'
			Use In 'GF_TableSearch'
			Return 0
		Endif

		This.ShowWaitMessage('Searching File: ' + m.tcFile)

		lnEndColumn = 255
		llIgnorePropertiesField = .F.

		Do Case
			Case 'VCX' $ m.lcExt
				lnStartColumn = 4
				llIgnorePropertiesField = This.oSearchOptions.lIgnorePropertiesField
			Case 'SCX' $ m.lcExt
				lnStartColumn = 4
				llIgnorePropertiesField = This.oSearchOptions.lIgnorePropertiesField
				* lnEndColumn = 12
			Case 'FRX' = m.lcExt
				lnStartColumn = 3 && Newer reports could start at col 6, but older reports can have text data starting in column 3
				* lnEndColumn = 21
				If Len(Field('timestamp', 'GF_TableSearch')) > 0 && Some really old reports may not have this field.
					Select Max(Timestamp) From 'GF_TableSearch' Into Array laMaxTimeStamp
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
					Inlist(m.lcField, 'OBJCODE', 'OBJECT', 'SYMBOLS')
				Loop
			Endif

			If llIgnorePropertiesField And lcField == 'RESERVED3'
				Loop
			Endif

			If m.lcExt = 'DBC'
				lcObjectType = Alltrim(Upper(ObjectType))
				If Type('objectname') = 'U'&& or;
					* 'OBJECT' $ Upper(objectname) or		;
					InList(lcObjectType, 'FIELD', 'VIEW', 'TABLE')
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

				If Not Found() Or m.llLocateError = .T.
					Loop && Loop to next column
				Endif

			Endif

			Do While Not Eof()

				lnMatchCount	   = 0
				loFileResultObject = Createobject('GF_FileResult')	&& This custom class has all the properties that must be populated if you want to
				llProcessThisMatch = .T.														&& have a cursor created
				llScxVcx		   = Inlist(m.lcExt, 'VCX', 'SCX')
				lcCode			   = Evaluate(m.lcField)

				With m.loFileResultObject
					.Process   = .F.
					.FileName  = Justfname(m.tcFile)
					.FilePath  = m.tcFile
					.MatchType = Proper(m.lcField)
					.FileType  = Upper(m.lcExt)
					.Column	   = m.lcField
					.IsText	   = .F.
					.Recno	   = Recno()
					.Timestamp = Iif(Type('timestamp') # 'U', Ctot(This.TimeStampToDate(Timestamp)), {// :: AM})

					lcClass			 = Iif(Type('class') # 'U', Class, '')
					.ContainingClass = m.lcClass
					._ParentClass	 = m.lcClass
					._BaseClass		 = Iif(Type('baseclass') # 'U', BaseClass, '')
					.ClassLoc		 = Iif(Type('classloc') # 'U', ClassLoc, '')

					lcParent = Iif(Type('parent') # 'U', Parent, '')
					lcName	 = Iif(Type('objname') # 'U', ObjName, '')
					._Name	 = Alltrim(m.lcParent + '.' + m.lcName, '.')

					Do Case
						Case m.lcExt = 'SCX'

							._Class = ''

							If Not Empty(m.lcParent)
								._Name = Strtran(._Name, m.lcFormName + '.', '', 1, 1) && Trim off Form name from the beginning of object name
							Else
								._Name = ''
							Endif

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
									lnParentId = parentId
									Select ObjectType, ObjectName From (m.tcFile) Where objectid = m.lnParentId Into Array laParent
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
					If (m.lcExt = 'VCX' And Empty(m.lcClass)) Or ;					 	&& This is the ending row of a Class def in a vcx. Need to skip over it.
					   (m.lcExt = 'FRX' And m.lcField = 'TAG2' And Recno() = 1) Or ;	&& Tag2 on first record in a FRX is binary and I want to skip it.
					   (m.lcExt = 'PJX' and m.lcField = 'KEY') 					  		&& Added this filter on 2021-03-24, as requested by Jim Nelson.

						llProcessThisMatch = .F.
					Endif

				Endwith

				If This.oSearchOptions.lTimeStamp And Not Between(Ctot(This.TimeStampToDate(Timestamp)), m.ldFromDate, m.ldToDate)
					llProcessThisMatch = .F.
				Endif

				If m.llProcessThisMatch = .T.
					If Not Empty(This.oSearchOptions.cSearchExpression)
						*lcCode = Evaluate(lcField)
						llHasMethods = Upper(m.lcField) = 'METHODS' Or		;
							m.lcExt = 'FRX' And Upper(m.lcField) = 'TAG' And Upper(Name) = 'DATAENVIRONMENT'
						lnMatchCount = This.SearchInCode(m.lcCode, m.loFileResultObject, m.llHasMethods)
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


				If Not This.oSearchOptions.lRegularExpression
					Try
						Continue
					Catch
						This.SetSearchError('Error scanning through table [' + m.tcFile + ']. File may be corrupt.')
						llContinueError = .T.
					Finally
					Endtry

					If m.llContinueError = .T.
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
		
	EndProc


	*----------------------------------------------------------------------------------
	Procedure SearchInTextFile(tcFile)

		Local loFileResultObject As 'GF_FileResult'
		Local loSearchResultObject As 'GF_SearchResult'
		Local lcCode, ldFileDate, lnMatchCount, lnSelect, llHasMethods

		lnSelect = Select()

		If !File(tcFile)
			This.lFileNotFound = .T.
			This.SetSearchError('File not found: ' + tcFile)
			Return 0
		Endif

		ldFileDate = This.GetFileDateTime(tcFile)

		ldFromDate = Evl(This.oSearchOptions.dTimeStampFrom, {^1900-01-01})
		ldToDate = Evl(This.oSearchOptions.dTimeStampTo, {^9999-01-01})
		ldToDate = ldToDate + 1&&86400 && Must bump into to next day, since TimeStamp from table has time on it

		If This.oSearchOptions.lTimeStamp And !Between(ldFileDate, ldFromDate, ldToDate)
			Return 0
		Endif

		loFileResultObject = Createobject('GF_FileResult')	&& This custom class has all the properties that must be populated if you want to
															&& have a cursor created
		With loFileResultObject
			.FileName = Justfname(tcFile)
			.FilePath = tcFile
			.MatchType = Proper(Justext(tcFile))
			.FileType = Upper(Justext(tcFile))
			.IsText = .T.
			.Timestamp = ldFileDate
		Endwith

		If !Empty(This.oSearchOptions.cSearchExpression)
			Try
				lcCode = Filetostr(tcFile) && File could be in use by some other app and can't be read in
				llReadFile = .T.
			Catch
				This.SetSearchError('Could not open file [' + tcFile + '] for reading.')
				llReadFile = .F.
			Endtry
			If !llReadFile
				Select (lnSelect)
				Return 0
			Endif

			llHasMethods = Inlist(Upper(loFileResultObject.MatchType) + ' ', 'PRG ', 'MPR ', 'H ')
			lnMatchCount = This.SearchInCode(lcCode, loFileResultObject, llHasMethods)
		Else
			* Can't search since there is no cSearchExpression, so we just log the file as a result.
			* This handles TimeStamp searches, where the cSearchExpression is empty
			loSearchResultObject = Createobject('GF_SearchResult')
			loSearchResultObject.UserField = loFileResultObject

			This.ProcessSearchResult(loSearchResultObject)
			lnMatchCount = 1
		Endif

		Select (lnSelect)

		Return lnMatchCount
		
	EndProc


	*----------------------------------------------------------------------------------
	*-- Read a user file set the the cFilesToSkip property
	Procedure SetFilesToSkip

		Local lcExclusionFile, lcFilesToSkip, lcLeft, lcLine, lcRight, lnI

		lcFilesToSkip	= ''
		lcExclusionFile	= This.cFilesToSkipFile

		If File(m.lcExclusionFile) And This.oSearchOptions.lSkipFiles
			lcFilesToSkip = Filetostr(m.lcExclusionFile)
		Endif

		This.cFilesToSkip		  = CR
		This.nWildCardFilesToSkip = 0

		For lnI = 1 To Alines(laLines, m.lcFilesToSkip + Chr(13) + '_command.prg', 5)
			lcLine	= Upper(laLines[m.lni])
			lcLeft	= Left(m.lcLine, 1)
			lcRight	= Right(m.lcLine, 1)

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
		
	EndProc


	*----------------------------------------------------------------------------------
	Procedure SetProject(tcProject)

		Local lcProject, llReturn

		lcProject = Lower(Evl(tcProject, ''))

		If Empty(lcProject)
			Return .T.
		Endif

		If !('.pjx' $ lcProject)
			lcProject = lcProject + '.pjx'
		Endif

		If File(lcProject)
			This.AddProject(lcProject)
			This.oSearchOptions.cProject = lcProject
			llReturn = .T.
		Else
			This.oSearchOptions.cProject = ''
			This.SetSearchError('Project not found [' + lcProject + '] in call to SetProject() method.')
			llReturn = .F.
		Endif

		Return llReturn
		
	EndProc


	*----------------------------------------------------------------------------------
	Procedure SetReplaceError(tcErrorMessage, tcFile, tnResultId, tnDialogBoxType, tcTitle)

		Local lcErrorMessage, lcFile, lnResultId

		lcFile = Alltrim(Evl(tcFile, 'None'))
		lnResultId = Evl(tnResultId, 0)

		lcResultId = Iif(lnResultId = 0, 'None', Alltrim(Str(lnResultId)))

		lcErrorMessage = tcErrorMessage + Space(4) + ;
			'[File: ' + lcFile + ']' + Space(4) + ;
			'[Result Id: ' + lcResultId + ']'

		This.ShowError(lcErrorMessage, tnDialogBoxType, tcTitle)

		This.oReplaceErrors.Add(lcErrorMessage)
		
	EndProc


	*----------------------------------------------------------------------------------
	Procedure SetSearchError(tcErrorMessage, tnDialogBoxType, tcTitle)

		* This.ShowError(tcErrorMessage, tnDialogBoxType, tcTitle)

		This.oSearchErrors.Add(tcErrorMessage)
		
	EndProc


	*----------------------------------------------------------------------------------
	Procedure ShowError(tcErrorMessage, tnDialogBoxType, tcTitle)

		Local lcTitle, lnDialogBoxType

		If Empty(tcErrorMessage) Or !This.oSearchOptions.lShowErrorMessages
			Return
		Endif

		lnDialogBoxType	= Evl(tnDialogBoxType, 0)
		lcTitle					= Evl(tcTitle, 'GoFishSearchEngine Error:')

		*!* ******************** Removed 11/10/2015 *****************
		*!* MessageBox(tcErrorMessage, lnDialogBoxType, lcTitle)
		
	EndProc


	*----------------------------------------------------------------------------------
	Procedure ShowWaitMessage(tcMessage)

		If This.oSearchOptions.lShowWaitMessages
			Wait Window At 5, Wcols() / 2 Nowait tcMessage
		Endif
		
	EndProc


	*----------------------------------------------------------------------------------
	Procedure StartProgressBar(tnMaxValue)

		If Vartype(This.oProgressBar) = 'O'
			This.oProgressBar.Start(tnMaxValue)
		Endif
		
	EndProc


	*----------------------------------------------------------------------------------
	Procedure StartTimer
	
		This.nSearchTime = Seconds()
		
	EndProc


	*----------------------------------------------------------------------------------
	Procedure StopProgressBar
	
		If Vartype(This.oProgressBar) = 'O'
			This.oProgressBar.Stop()
		Endif
		
	EndProc


	*----------------------------------------------------------------------------------
	Procedure StoreInitialDefaultDir
	
		This.cInitialDefaultDir = Sys(5) + Sys(2003)
		
	EndProc


	*----------------------------------------------------------------------------------
	Procedure ThorMoveWindow
	
		If Type('_Screen.cThorDispatcher') = 'C'
			Execscript (_Screen.cThorDispatcher, 'PEMEditor_StartIDETools')
			_oPEMEditor.Outils.oIDEx.MoveWindow()
		Endif
		
	EndProc


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

		Local lcRetVal                         &&  Requested data returned from procedure, ;
			lnYear, ;
			lnMonth, ;
			lnDay, ;
			lnHour, ;
			lnMinute, ;
			lnSecond, ;
			loException

		If Type('tnTimeStamp') != "N"          &&  Timestamp must be numeric
		* Wait Window "Time stamp passed is not numeric" NoWait
			Return ""
		Endif

		If tnTimeStamp = 0                     &&  Timestamp is zero until built in project
			Return "Not built into App"
		Endif

		If Type('tcStyle') != "C"              &&  Default return style to both date and time
			tcStyle = "DATETIME"
		Endif

		If !Inlist(Upper(tcStyle), "DATE", "TIME", "DATETIME")
			Wait Window "Style parameter must be DATE, TIME, or DATETIME"
			Return ""
		Endif

		lnYear   = ((tnTimeStamp / (2 ** 25) + 1980))
		lnMonth  = ((lnYear - Int(lnYear)    ) * (2 ** 25)) / (2 ** 21)
		lnDay    = ((lnMonth - Int(lnMonth)  ) * (2 ** 21)) / (2 ** 16)

		lnHour   = ((lnDay - Int(lnDay)      ) * (2 ** 16)) / (2 ** 11)
		lnMinute = ((lnHour - Int(lnHour)    ) * (2 ** 11)) / (2 ** 05)
		lnSecond = ((lnMinute - Int(lnMinute)) * (2 ** 05)) * 2       &&  Multiply by two to correct
		&&  truncation problem built in
		&&  to the creation algorithm
		&&  (Source: Microsoft Tech Support)

		lcRetVal = ""

		If "DATE" $ Upper(tcStyle)
		*< 4-Feb-2001 Fixed to display date in machine designated format (Regional Settings)
		*< lcRetVal = lcRetVal + RIGHT("0"+ALLTRIM(STR(INT(lnMonth))),2) + "/" + ;
		*<                       RIGHT("0"+ALLTRIM(STR(INT(lnDay))),2)   + "/" + ;
		*<                       RIGHT("0"+ALLTRIM(STR(INT(lnYear))), IIF(SET("CENTURY") = "ON", 4, 2))

		*< RAS 23-Nov-2004, change to work around behavior change in VFP 9.
		*< lcRetVal = lcRetVal + DTOC(DATE(lnYear, lnMonth, lnDay))
			Try
				lcRetVal = lcRetVal + Dtoc(Date(Int(lnYear), Int(lnMonth), Int(lnDay)))
			Catch To loException
				lcRetVal = lcRetVal + Dtoc(Date(1901, 1, 1))
			Endtry
		Endif

		If "TIME" $ Upper(tcStyle)
			lcRetVal = lcRetVal + Iif("DATE" $ Upper(tcStyle), " ", "")
			lcRetVal = lcRetVal + Right("0" + Alltrim(Str(Int(lnHour))), 2)   + ":" + ;
				Right("0" + Alltrim(Str(Int(lnMinute))), 2) + ":" + ;
				Right("0" + Alltrim(Str(Int(lnSecond))), 2)
		Endif

		Return lcRetVal
		
	EndProc


	*----------------------------------------------------------------------------------
	Procedure TrimWhiteSpace(tcString)

		Local lcTrimmedString

		lcTrimmedString = Alltrim(tcString, 1, Chr(32), Chr(9), Chr(10), Chr(13), Chr(0))
		lcTrimmedString = Strtran(lcTrimmedString, Chr(9), Chr(32))

		Return lcTrimmedString
		
	EndProc


	*----------------------------------------------------------------------------------
	Procedure UpdateCursorAfterReplace(tcCursor, toResult)

		Local lcColumn, lcFileToModify, lnChangeLength, lnCurrentRecno, lnMatchStart, lnProcStart
		Local lnResultRecno, lnSelect

		If This.oSearchOptions.lPreviewReplace = .T.
			Return
		Endif

		lnChangeLength = toResult.nChangeLength

		lnSelect = Select()
		Select (tcCursor)
		lnCurrentRecno = Recno()

		*-- Create local vars of certain fields from the current row that we need to use below
		lcFileToModify = Alltrim(FilePath)
		lnResultRecno = Recno
		lnProcStart = ProcStart
		lnMatchStart = MatchStart
		lcColumn = Column

		*-- Cannot process same source code line more than once, so mark this and all other rows of
		*-- the same oringal source line with replacd = .t., and also update the matchlen
		 Update  &tcCursor ;
			 Set Replaced = .T., ;
				 MatchLen = Max(MatchLen + lnChangeLength, 0) ;
			 Where Alltrim(FilePath) == lcFileToModify And ;
				 Recno = lnResultRecno And ;
				 Column = lcColumn And ;
				 MatchStart = lnMatchStart

			*-- Update the stored code with the new code for all records of the same original source
		 Update  &tcCursor ;
			 Set Code = toResult.cNewCode;
			 Where Alltrim(FilePath) == lcFileToModify And ;
				 Recno = lnResultRecno And ;
				 Column = lcColumn

			*-- Update matchstart values on remaining records of same file, recno, and column type
		 Update  &tcCursor ;
			 Set MatchStart = (MatchStart + lnChangeLength) ;
			 Where Alltrim(FilePath) == lcFileToModify And ;
				 Recno = lnResultRecno And ;
				 Column = lcColumn And ;
				 MatchStart > lnMatchStart

			*-- Update procstart values on remaining records of same file, recno, and column type
		 Update  &tcCursor ;
			 Set ProcStart = (ProcStart + lnChangeLength) ;
			 Where Alltrim(FilePath) == lcFileToModify And ;
				 Recno = lnResultRecno And ;
				 Column = lcColumn And ;
				 ProcStart > lnProcStart

		Goto (lnCurrentRecno)

		Select (lnSelect)
		
	EndProc


	*----------------------------------------------------------------------------------
	Procedure UpdateProgressBar(tnValue)

		If Vartype(This.oProgressBar) = 'O'
			This.oProgressBar.nValue = tnValue
		Endif
		
	EndProc


	*----------------------------------------------------------------------------------
	*-- If we are in Replace Preview mode, do not attempt to update the Replace Detail record.
	Procedure UpdateReplaceHistoryRecord

		If This.oSearchOptions.lPreviewReplace = .F.
			Update (This.cReplaceHistoryTable) Set replaces = This.nReplaceCount Where Id = This.nReplaceHistoryId
		Endif
		
	EndProc


EndDefine
  
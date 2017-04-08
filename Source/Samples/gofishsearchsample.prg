*-------------------------------------------------------------------------------------------------------
* See http://vfpx.codeplex.com/wikipage?title=GoFish for the lastest GoFish4 Code Search sample programs:
*-------------------------------------------------------------------------------------------------------
*-- GoFish4 Search Engine Sample Usage     (Updated 2011-05-18)
* 
*-- This code serves as a test suite and coding examples of how to use the GoFish Search Engine
*-------------------------------------------------------------------------------------------------------
 
Local loSearchEngine as 'GoFishSearchEngine'
Local lcPath, lcProject

*-- Change these refereces to actual files/paths to test the search engine
lcPath = 'C:\Temp\Source\'
lcProject = lcPath  + 'TestProject.pjx'
lcCodeFile = lcPath + 'GoFish_PrgTest.prg'
lcFormFile = lcPath + 'Forms\About.scx'

loSearchEngine = CreateObject('GoFishSearchEngine')
loSearchEngine.cSearchResultsAlias = 'MyCursor'          
 
With loSearchEngine.oSearchOptions
	.cSearchExpression = 'test'
	.lRegularExpression = .f.
	.lMatchWholeWord = .t.
	.lMatchCase = .f.
	.lSearchInComments = .t.
	 
	.lCreateResultsCursor = .t.
	.lCreateResultsCollection = .t.

	.lShowErrorMessages = .f. && Will pop up MessageBox to show any errors
	.lShowWaitMessages = .f.
	.lShowNoMatchesMessage = .f.
EndWith


Clear 
 
*-- Sample 1 - Searching a path -----------------
		? 'Test 1 - Searching Path: ' + lcPath
    loSearchEngine.oSearchOptions.lIncludeSubdirectories = .t.
    loSearchEngine.SearchInPath(lcPath)
		ShowResults(loSearchEngine)
 
 
*-- Sample 2 - Search a project ---------------------------
		? 'Test 2 - Searching Project: ' + lcProject
		loSearchEngine.oSearchOptions.lLimitToProjectFolder = .f.
		loSearchEngine.SearchInProject(lcProject) 
		ShowResults(loSearchEngine)
 
*-- Sample 3 - Search for changed objects or files in project based on TimeStamps ---------------------------
		? 'Test 3 - Searching Project with TimeStamp: ' + lcProject
     With loSearchEngine.oSearchOptions
      		.cSearchExpression = ''
          .lTimeStamp = .t.
          .dTimeStampFrom = {^2011-04-01}
          .dTimeStampTo = {}
     EndWith
 
    loSearchEngine.SearchInProject(lcProject)  
		ShowResults(loSearchEngine)
 
 
*-- Sample 4 - Searching individual PRG file -------------------------
		? 'Test 4 - Searching a PRG file: ' + lcCodeFile
		loSearchEngine.PrepareForSearch()
    loSearchEngine.SearchInFile(lcCodeFile)
		ShowResults(loSearchEngine)
		
 
*-- Sample 5 - Searching individual SCX file -------------------------
		? 'Test 5 - Searching a SCX file: ' + lcFormFile
		loSearchEngine.PrepareForSearch()
    loSearchEngine.SearchInFile(lcFormFile)
		ShowResults(loSearchEngine)

 
*-- Sample 6 - Search a blob of text (Typcailly from a PRG, but could be any text at all) ---------------------
		? 'Test 6 - Searching a misc blob of text/code'
		loRefData = CreateObject('GF_FileResult')
		lcSomeCode = FileToStr(lcCodeFile) && Code Text can come from anywhere. Doesn't even have to be FoxPro Code
		loSearchEngine.PrepareForSearch()
		loSearchEngine.SearchInCode(lcSomeCode, loRefData) && lxYourRefData is optional. Could be a string, int, or any object you wish to be affiliated with any match that is found
 		ShowResults(loSearchEngine)

 
 
*-- Notes --------------------------------------------------------------------------------------
 
*-- Before any given search, you can clear previous results like this, else results will be appended.
*     loSearchEngine.ClearResultsCursor()
*     loSearchEngine.ClearResultsCollection()


*------------------------------------------------------------------------------
Procedure ShowResults

Lparameters toSearchEngine

lcIndent = Space(20)

 ? Space(12) + 'Error count: ' + Alltrim(Str(toSearchEngine.oErrors.Count))

 For Each loError In toSearchEngine.oErrors
 	? lcIndent  + loError
 Endfor

 ? Space(1)

*-- After a search, you can read these properties:
 *(you'll also have a cursor, if you set the properties accordingly.)


 ? lcIndent  + 'MatchLines:'  + Str(toSearchEngine.nMatchLines) && How many lines in the code/file matches.
 ? lcIndent  + 'Files with matches:' + Str(toSearchEngine.nFileCount)   && How many files had matches. (Doesn't apply if SearchInCode() called directly)
 ? lcIndent  + 'Time: ' + Alltrim(Str(toSearchEngine.nSearchTime)) + ' seconds' && How long did the search take
 ? lcIndent  + 'Cursor Rows: ' + Alltrim(Str(Reccount(toSearchEngine.cSearchResultsAlias)))
 ? lcIndent  + 'oResults Count: ' + Alltrim(Str(toSearchEngine.oResults.Count))  && A collection (if lCreateResultsCollection was set)
     
 ? Space(1)     

 
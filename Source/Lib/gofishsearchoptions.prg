#Include GoFish.h

Define Class GoFishSearchOptions As Custom

	* This is the escaped version of the cSearchExpression that is used for the actual
	* Reg Ex search. The class maintains this, so the class user should never touch this.
	cEscapedSearchExpression               = ''
	
	* A file mask template, like Job*, or *Job, or *job*, or Job*.scx, or *.scx, etc.
	cFileTemplate                          = ''
	
	cIndexString                           = ''
	
	* An additional list of file extensions to be included in the search.
	* Just separate each extension by a space. No dot is needed. Case does not matter.
	cOtherIncludes                         = ''
	
	* This property can be used to store the name of a Path on the object.
	cPath                                  = ''
	
	* This property can be used to store the name of a Project on the object.
	cProject                               = ''
	cRecentScope                           = ''
	cReplaceExpression                     = ''
	
	* Search expression to be sought. Can be a regex. Set lRegularExpression to .t. if it is.
	cSearchExpression                      = ''
	
	dTimeStampFrom                         =  {}
	dTimeStampTo                           =  {}
	
	* This flag must be set in order to do a replace operation where the
	* replacement string is a blank string.
	lAllowBlankReplace                     = .F.
	
	lBackup                                = .T.
	lColorizeVFPCode                       = .F.
	
	* Indicates if you want to store each match result object into the
	* oResults colleciton on the Search Engine class.
	lCreateResultsCollection               = .F.
	
	* Inidcates if you want to store each search result in a local cursor. See cSearchResultsAlias
	* propertyy on the Search Class for the namee of the Cursor.
	lCreateResultsCursor                   = .T.
	
	* This will supress the warning dialog that pops up every time you perform a replace.
	lDoNotShowReplaceWarning               = .F.
	
	lEnableReplaceMode                     = .F.
	lIgnoreMemberData                      = .F.
	lIgnorePropertiesField                 = .F.

	* Not used.
	lIncludeAllFileTypes                   = .F.

	lIncludeASP                            = .T.
	lIncludeDBC                            = .F.
	lIncludeFRX                            = .T.
	lIncludeH                              = .T.
	lIncludeHTML                           = .T.
	lIncludeINI                            = .T.
	lIncludeJAVA                           = .T.
	lIncludeJSP                            = .T.
	lIncludeLBX                            = .T.
	lIncludeMNX                            = .T.
	lIncludeMPR                            = .T.
	lIncludePJX                            = .F.
	lIncludePRG                            = .T.
	lIncludeSCX                            = .T.
	lIncludeSPR                            = .T.
	lIncludeSubdirectories                 = .T.
	lIncludeTXT                            = .T.
	lIncludeVCX                            = .T.
	lIncludeXML                            = .T.
	
	* When searching a PJX file, this flag will skip over searches on any
	* files in the project which are not located in or below the Project's home path.
	lLimitToProjectFolder                  = .F.
	
	lMatchCase                             = .F.
	lMatchWholeWord                        = .F.
	lPreviewReplace                        = .T.
	
	* Indicates if the cSearchExpression is intended to be used as a Regular Expression.
	lRegularExpression                     = .F.
	
	lSearchInComments                      = .T.
	lShowAdvancedFormOnStartup             = .F.
	
	* Determines if a messagebox will pop up any time there is an error on
	* the Search or Replace family of methods.
	lShowErrorMessages                     = .T.
	
	* Display a MessageBox at the end of a search if there were no matches found.
	lShowNoMatchesMessage                  = .T.
	
	* Will display the name of each file in a wait window as they are being
	* processed. Caution: using this feature slows down the search a good bit.
	lShowWaitMessages                      = .F.
	
	* A flag to indicate if the user wants to process the files in
	* cFilesToSkipFile files to skip over certain files during the search.
	lSkipFiles                             = .F.

	* A flag to indicate if the user wants to process the files in
	* cFilesToIncludeFile files to cnclude over certain files during the search.
	lIncludeFiles                          = .F.

	* Indicates if you want a copy of the code block stored on each search
	* reesult record (for ResultsCursor) or collection node (for ResultsCollection).
	lStoreCode                             = .F.

	* Indicates if the search will be limited to filedates and object
	* TimeStamps that fall on or between the dTimeStampForm and dTimeStampTo values.
	lTimeStamp                             = .F.

	lWarnWhenUnableToOpenFilesDuringSearch = .T.
	nHtmlMatchLineColor                    = 8454143

	* Limits the amx number of results created. Search will stop at this
	* limit. Collection restutls over 50,000 with lStoreCode = .t. could
	* lead to memory problems.
	nMaxResults                            = 10000

	* Indicates which search mode to use: Plain, LIKE, or RegEx.
	* See GoFish.h constants file for values.
	nSearchMode                            = 1
	lWildCardContinuationLines			   = .T.
	lOptimizeWithGrep					   = .T.

	* 1 = Active Project, 2=Browse Project, 3 = Current Dir, 4= Browse Directory, 5 = All Open Projects
	nSearchScope                           = 3
	nPreviousSearchScope				   = 3
	
*!* - rk - 2023-10-22 - added property to stop problem in 
*!* - gofishSearchEngine.prg|SetIncludePattern 
	oRegExpFileTemplate              = .Null.

	* Indicates whether column 'File Name with Path' shows a Relative Path
	lShowRelativePath                      = .F.

	* Indicates whether to filter history form by the current Scope
	lFilterHistoryByScope                  = .F.

	* Name of the field used to filter history form by the current Scope
	cSearchHistorySort			           = ''

	* Indicates where the projects/folders in the scope dropdown are to be sorted.
	lSortScopeDropdown			           = .F.

	*** JRN 2024-02-14 : Used if wild cards use whole word search
	cWholeWordSearch = ''

	*** JRN 2024-05-04 : File extensions searched in current search
	cSearchExtensions = ''
	
	lShowMenuBarNumbers = .F.
		
	lCleanUpGrepFiles = .T.
	
	lPreVFP9		   = .F.
	lPreVFP9EditSource = 'EditSource'
	lUnmatchedParenError = .F.

	*---------------------------------------------------------------------

	Procedure lRegularExpression_Access
	
		Return (This.nSearchMode = GF_SEARCH_MODE_REGEX)
		
	EndProc


EndDefine

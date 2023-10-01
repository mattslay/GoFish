# Changes
Change list of GoFish - An advanced code search tool for MS Visual Foxpro 9

## Purpose of this document
This document list the changes of GoFish.   
Replaces the lists in
- [Readme.md](/readme.md)
- ChangeLog_Ver_4.txt
- ChangeLog_Ver_5.txt

----
## Ver 6.2.005
**Released 2023-09-07**
### Changes
- Added file .FoxBin2Prg_Ignore in local storage directory, to stop FoxBin2Prg after version [1.20.07](https://github.com/lscheffler/foxbin2prg/releases/tag/1.20.07) from processing this directories.

## Ver 6.2.004
**Released 2023-09-04**
### Changes
- fixed problem with replace backup folder missing
- fixed an issue where whole history was populated in tree, while only that of the recent session was turned on
- fixed an issue with history tree turned on and search not returning result
- fixed: Scope combo could not be changed with keyboard, if CodeReferences option where turned on.
- added option to enter * or *.* in Advanced Search \-\> Filetypes \-\> Others to express **ALL**
- added option to enter multiple file skeletons to file templates
- better UI to set/unset file types
- optimized speed of folder search
- ignore git toplevel folder instead of .git
- ignore GoFish's settings folder, but allow to search the GF_Saved_Search_Results by explicitly selecting it (on your own risk)
- fixed Button to show menu tree should not appear for MPR files; fixed #102
- MPR are now in the "Dangerous risk" replace group; closing #109
- DynamicBackColor for Replacerisk 3 - Dangerous was wrong
- files not in filter visible in tree; fixed issue #118
- wrong property used of SF_RegExp
- searching for file without extension is now allowed in "Advanced Search", "File template". #84 
  - Note: "\*" is all files WITH extension, without extension is like "\*."
- fixed: Error "invalid path or file name" on startup; fixed #120 (PatrickvonDeetzen)
- forced to use of APP over PRG
- Renamed GoFish5 to GoFish

----
## Ver 6.2.001
**Released 2023-05-09**
### Changes
- Fixed: Code - window shows info for files without hit in filter #94  
- Fixed: If search is restored via "History", restore search settings; fixed #92
- Fixed: Sort for full file path and date of search has higher order then the column clicked; fixed #91
- Fixed:  Filter Builder errors when filter string is equal sign #8 
- Improved: Messagebox text
- Improved: Different colours for filter in search result #88
  - Note: Filter expressions in the *Like* style are not highlighted (and never will be)
- Improved: CSS for "Code Window" in settings folder to alter colours
- Improved: DotNet-RegExp for some operations (not search)

## Ver 6.2.000
**Released 2023-04-20**
### Changes
- Fixed: Update failed when no search was started after first run
- Improved: First run sets now version to settings DBC
- Improved: Update messages give info for storage path

## Ver 6.1.000
**Released 2023-04-15**
### Changes
- Improved: Update of storage structure by version change
- Improved: Highlight searched text in opened window; merged #75
- Improved: Max length of field FileName set to 100, issue #78
- Fixed: Files with to long filename where sorted to wrong node, fixing #78
- Fixed: Error while exporting, fixing #79
- Fixed: Scope disappears after a project search, fixing #69 

## Ver 6.0.010
**Released 2023-03-23**
### Changes
- Improved: Files of type SPR will be modified like PRG. (again)

## Ver 6.0.008
**Released 2023-03-18**
### Changes
- Fixed: Some searches produce OLE error code 0x800a139c: Unknown COM status code #68
- Fixed: GoFish 6.0 is not active by default #66
- Fixed: issue with filter results

## Ver 6.0.007
**Released 2023-03-18**
### Changes
- Improved: Pre-update screen for bug search.

## Ver 6.0.006
**Released 2023-03-12**
### Changes
- Fixed: Suppress error with odd ORDER BY from old version #53
- Fixed: Check box *Whole Word* was not framed when non default value #7

## Ver 6.0.005
**Released 2023-03-11**
### Changes
- Improved: Screen to help with the duplicate issue #53
  SET SAFETY ON to help to find the problem

## Ver 6.0.004
**Released 2023-03-11**
### Changes
- Improved: `SET SAFETY OFF` to circumvent a problem with double copy. (issue #53) Until real fix.

## Ver 6.0.003
**Released 2023-03-10**
### Changes
- New: Migrating will create a backup of old GoFish data in `Home(7) + "GoFish_Backup"`.

## Ver 6.0.000
**Released 2023-03-08**
### Changes
For more details check [Changes on V6](./New_6.md).
- New: Mode to display history in tree
- New: Modified structure to store history to speed up display of history in tree while keeeping common history table small.
- New: Option to swap buttons in search history. Send OK / Cancel buttom left.
- New: Option to show search history by active scope
- New: Option to set scope to Active Project / Active Path as default
- New: Additional parameters for active scope
- New: Additional parameters to clear settings / storage
- Improved: Alternative button layout for *Load History Form*
- Improved: Load history faster
- Improved: Storage of replace history
- Improved: Display of replace history
- Improved: Sort per filepath rather then filename

## Ver 5.1.010
**Released 2023-02-23**
### Fixes
- Fixed: Options Dialog Desktop checkbox crashes on InteractiveChange code, fix issue #47

## Ver 5.1.010
**Released 2023-02-22**
### Fixes
- Fixed: Renamed a buch of procedures to pseudo-unique names, try to fix issue #45

## Ver 5.1.009
**Released 2023-01-31**
### Fixes
- Fixed: Fixed problem with wrong object addressed in backup issue #41

## Ver 5.1.008
**Released 2022-11-23**
### Fixes
- Fixed: Problem with filter settings storage place from filter form #38

## Ver 5.1.007
**Released 2022-11-22**
### Fixes
- Fixed: Problem with Filter form issue #36

## Ver 5.1.006
**Released 2022-11-11**
### Fixes
- Fixed: Problem with Collate and expression length issue #34

## Ver 5.1.005
**Released 2022-11-09**
### Fixes
- Fixed: Highlighted match not visible in code view for SCX/VCX issue #27 
- Fixed: Due to unlucky git settings, some text files where line end with LF instead of CRLF
### Changes
- Improved: git line end logic forces checkin-as-is checkout-as-is for this repo

## Ver 5.1.004
**Released 2022-11-01**
### Fixes
- Fixed: GF_Search_History_Form_Settings.xml written to default folder. issue #31 
- Fixed: User preference settings ignored and overwritten. issue #29 
- Fixed: Startup parameter ignored. issue #21
- Fixed: Load history form should provide keyboard access to grid. issue #10
### Changes
- Improved: buildgofish.prg forces to use of VFP9 SP2

## Ver 5.1.003
**Released 2022-10-29**
### Fixes
- Fixed: User preference settings ignored and overwritten. issue #29
- Fixed: Opening Filter dialog fails with class not found. issue #28

## Ver 5.1.002
**Released 2022-10-23**
### Fixes
- Fixed: problem with deleted settings file on startup. issue #25
### Changes
- Improved: Issues bug template.

## Ver 5.1.001
**Released 2022-10-21**
### Fixes
- Fixed: Problem registring with Thor  issue #23

## Ver 5.1.000
**Released 2022-10-19**
### Fixes
- Fixed: Setting Desktop will engage if form is not closed with *Ok*  issue #19
- Fixed: problem with deleted settings file on startup. issue #17
### Changes
- New: Option for local settings. Allows local settings and history. Select storage location if local.
- Improved: Storage of some options separated from search settings.

## Ver 5.0.209
**Released 2022-10-17**
### Stuff
- Just to force right verison of the app

## Ver 5.0.208
**Released 2022-10-16**
### Changes
- Improved: Added options to allow control of several colours

## Ver 5.0.206
**Released 2022-10-16**
### Fixes
- Fixed: The View Replace History button wasn't working because of a renamed file.

## Ver 5.0.207
**Released 2022-10-16**
### Changes
- Improved: Added option to allow code window to be opened at stored position. (Option: Open code windows at top left corner)

## Ver 5.0.206 
**Released 2022-09-23**
### Fixes
- Fixed: The View Replace History button wasn't working because of a renamed file.

## Ver 5.0.205
**Released 2022-08-16**
### Fixes
- Fixed: When the results are shown and the order of the results is changed by clicking the header of column "Proccode", "Statement" or "Code" then the error "SQL: ORDER BY clause in invalid" was thrown. This was fixed by also excluding these columns in the functions "SortColumn" and "SortColumSecondary".
- Fixed: When the search option "Whole word" was active then it could happen that too many higlightings where shown in the HTML preview. This affected only lines with an appropiate result (matching "Whole word") and having another inappropiate result (not matching "Whole word).
- Fixed: On clicking "Edit" to open the file of the selected result could lead to error "Variable 'CR' is not found.".
- Fixed: For results that are part of a class-container and used in a form with a formset GF showed a wrong "name" (or you may call it "path") in the results. 
### Changes
- Improved: Some changes were made to "gofishsearchengine.prg" so that GF can wrap up the results even faster.
- Improved: When clicking "Edit" to open a form or class of the selected result and this file could not be opened then nothing happened. From now on GF will tell you that the file can´t be opened and why not.

## Ver 5.0.202
**Released 2022-06-24**
### Changes
- Made the setting of Desktop for the main GoFish form an option

## Ver 5.0.201
**Released 2022-05-19**
### Changes
- Set Desktop to .T. for the main GoFish form so it can be moved outside the VFP window

## Ver 5.0.200 
**Released 2022-05-14**
### Fixes
- Fixed "constant already defined issue" when building app
- Fixed an error displaying the options dialog
### Changes
- Changed CodePlex references to VFPX and removed reference to BitBucket
- Removed unneeded files from the repository
- Improved the build process

## Ver 5.0.170
**Released 2021-03-24**
### Changes
- GoFishSearchEngine class: Fixed folder paths for cFilesToSkipFile, cReplaceDetailTable, cReplaceDetailTable.
- Options Form-> Replace tab: Added read-only fields to show values for cReplaceDetailTable and cReplaceDetailTable.

## Ver 5.0.169 
**Released 2021-03-24**
### Changes
- Removed matches from a Project PJX on the "Key" field (i.e. Match Type = "< Key >" in the results grid.)

## Ver 5.0.168
**Released 2021-03-23**
### Changes
- In the search results grid, if **Match Type** is `\<\<Filename>>` (i.e. matching a filename in a Project), then double-clicking on the row will now open the file with correct editor. (Contributed by Jim Nelson)
- In the search results treeview, the results are now ordered by Class, Filename, FilePath. (Contributed by Jim Nelson)
- `GoFishSearchEngine.vcx` has been converted to PRGs file. The 2 classes in the VCX are replaced by `GoFishSearchEngine.prg` and `GoFishSearchOptions.prg`. This was done to remove binaries (VCX) from the code base, in hopes of making it easier to maintain the source code in a more source control friendly style.
- Converted `GF_PEME_BaseTools.vcx` to prg.

## Previous releases
- Ver 5.0.163 - Released 2017-02-12
- Ver 5.0.162 - Released 2016-12-06

## What's new in version 5.0:
### Changes
- Saved Search History – You can now save the history of your searches, either automatically (for all searches) or selectively, and can restore the search parameters and results grid from these saved searches.  You can also selectively delete your saved searches or use the GF Janitor to automatically delete old ones.
- More powerful filter options – You can filter on secondary matches in the code "in the neighborhood" (that is, in the same statement or same procedure) of the original match; the filter form has been re-organized to provide clarity; and filters can be combined using logical operators AND or OR.
- Handling of PRG-based classes – Special attention has been paid to matches in PRG-based classes, which now are treated as much as possible like VCX-based classes. Their names appear in the "Class", "Base Class", and "Parent Class" columns in the grid and also in the new category "Classes" in the Treeview.
- Treeview changes – There are a few new categories: "Classes", "Menus", and "Projects". The "Classes" category, unlike all the other categories which have files as sub-nodes, has classes as sub-nodes. 
- Column Changes  – There are two new columns, "Parent Class VCX" and "Containing Class"; the column headings for some of the other columns have been reworded; and the contents of some of the class-related columns have been enhanced.
- Plug-In to control the grid display – You can set the use the plug-in to set the Dynamic* (or any other) properties of the grid. The sample provided changes the colors used for each row.
- Other UI changes – There are many, many other UI changes, including the ability to select the location of the Code View window relative to the grid.

## Build 4.4.446 BETA
**2015-04-28**
- Added checkbox on Options form to disable saving of search results history
- Changed timestamp of search results history to 24-hour format
  
## Build 4.4.445 BETA
**2015-04-28**
- Previous build incorrectly used older version of Sedna My.
- Updated Sedna My to latest VFPx version of 2015-01-23
- Recompiled to fix bug in using the Filter form.

## Build 4.4.444 BETA
**2015-02-02**
- Restoring previous search will no longer leave Search History DBC open.
- Restored warning message if search phrase contains leaded or trailing spaces.
- Updated Thor Check For Updates menu image on Options-> Update page.

## Build 4.4.443 BETA
**2015-01-26**
- Fixed anchor on folder icon on Search History form.
	
## Build 4.4.442 BETA
**2015-01-26**
- Fixed sound being made in some cases when starting a new search.

## Build 4.4.441 BETA
**2015-01-26**
- Fixed bug where LockScreen was not getting cleared.
	
## Build 4.4.440 BETA
**2015-01-24**
- Fixed startup sequence if Saved Search Folder does not exist.

## Build 4.4.438 BETA
**2015-01-24**
- Now making a copy of the Filter form settings with each saved search.
- Added folder icon button on Search History form to open the folder, if needed.

## Build 4.4.437 BETA
**2015-01-24**
- Added short key for Restore button
- Can now Right-Click on Restore button to restore last search.
- Can now double-click on row in Search History click to select and close form.
- Search History form now remembers form size and positions between use.

## Build 4.4.436 BETA
**2015-01-24**
- Speed improvements when restoring a previous search result.
- Made the "Load" button default button on Search History form.

## Build 4.4.434 BETA
**2015-01-24**
- Added UI form to select and restore earlier search results.

## Build 4.4.431 BETA
**2015-01-23**
- Added Restore Previous Searches feature.

## Build 4.4.429 BETA
**2015-01-21**
- Added feature to restore previous search results when GoFish starts up.
- Converted SSC format to use FoxBin2Prg

## Build 4.3.055 BETA
**2013-06-06**
- Added code to browse DBFs with SuperBrowse if Thor is present

## Build 4.3.054 BETA
**2013-06-06**
- Fixed bug where DBFs could not be searched in Project scope.

## Build 4.3.053 BETA
**2012-11-15**
- Added handling for files with corrupted memo fields.
  
## Build 4.3.052 BETA
**2012-10-06**
- Fixed bug when searching old reports from FoxPro 2.6 
- Added support for exporting to Excel.
	
## Build 4.3.050 BETA
**2012-10-06**
- Improved sorting capabilities for grid. Click columns headers as follows:
  1. Left click always selects primary sort, pushes previous primary and secondary sorts down to second and third.
  2. Right click selects secondary, leaves primary untouched, pushes current secondary down to third.

## Build 4.3.049 BETA
**2012-10-05**
- Added Right-Click on Header for Secondary Sorting
- Sorting in columns is now case insensitive

## Build 4.3.048 BETA
**2012-10-04**
- Fixed bug which prevented matches from being found in PRG files in certain cases. (Reported by JRN 2012-10-04)

## Build 4.3.047 BETA
**2012-09-10**
- Fixed bug which prevented double-clicking from opening methods on Forms.
  
## Build 4.3.046 BETA
**2012-09-03**
- Now using BindEvent() to handle double clicks from main grid

## Build 4.3.045 BETA
**2012-08-30**
- Added handler for 'dot' so GoFish will work properly with 'IntellisenseX by Dot' from Thor

## Build 4.3.027 BETA
**2012-07-01**
- Fixed call in main form Init() that was resetting grid to defaults.
  
**Switched it from FormatGrid() to FormatGridForReplaceMode(), which also fixed the case where the form was closed while Replace mode columns were visible)
- Extracted some code from FormatGridForReplaceMode() to ShowGridColumn()

## Build 4.3.026 BETA
**2012-07-01** 
- Fixed column ordering on grid to better handle hidden columns. Columns on options form are now sorted by visible first.
- Fixed bug that caused columns to get jumbled up when showing columns related to Replace Mode.

## Build 4.3.024 BETA
**2012-07-01**
- Added code to migrate the Replace Detail table to the new structure of the Results cursor.
- Added Options spinner for setting the html code view zoom factor.
- Now show all Replace Detail fields when browsing replace history.
 
## Build 4.3.022 BETA
**2012-06-29**
- Added ClassLoc column to search results grid
- Rename several columns in the search results related to class, baseclass, parent, and object names
- Added dialog to notify user if there were any Error during Replace operation
- Updated image for Replace errors to a larger image so it's earlier to see.

## Build 4.3.014
**2012-06-18**
- Fixed bug in Edit Line Replace Mode when deleting the entire line.
- Fixed bug which prevented GF from opening the correct method when editing file from GF.

## Build 4.3.012
**2012-06-08**
- Fixed Editing of SCX and VCX from within GoFish so it will open to the correct method and object.\\
 (This bug only affected non-Thor users)

## Build 4.3.011
**2012-06-03**
- Filter Form: Added support for wildcards (* and ?) and the Not operator (!) in string match fields.
- Added option for confirming 'Whole Word' search if it is marked when starting a new search.
- Added option for clearing the 'Apply Filter' checkbox with each new search
- Added option to specify number of MRU entries to display in Scope/Search Expression dropdowns.
- Fixed Report so that it will correctly use the same filter that is in place on the main grid.

 ## Build 4.3.002
**2012-05-07**
- Added new Search Mode: Uses wildcard matching (uses the FoxPro LIKE() function)
- Added radio buttons on Advanced form to select Search Mode
- Restored match word highlighting in Html code view
- Enhanced File Template filter to accept '?' in filename or ext.
- Fixed some other misc html code rendering things.
- Fixed issue where Browser Zoom factor can get corrupted.
- Big-A and Little-A buttons will now be hidden when Html code browser panel is collapsed.
- Enhanced PropNvl() function to return default value if data type of stored value is does not match default value data type
- Added initial code to prepare for support of Google Chrome Frame to render html view.


## Ver 4.3
**Released 2012-05-07**
### Fixes
- Fixed Report so that it will correctly use the same filter that is in place on the main grid.
### Changes
- New  Wildcard Search Mode allows `*` and `?` in Search Expression.
- Wildcards (`*` and `?`) and the NOT operator (`!`) can be now used on the Filter Form.
- Added option for confirming "Whole Word" search if it is marked when starting a new search.
- Added option for clearing the "Apply Filter" checkbox with each new search.
- Added option to specify number of MRU entries to display in Scope/Search Expression dropdowns.

## Build 4.2.068
**2012-03-30**
- Fixed bug when double-clicking to view a match from a dbf (Browse is called) (Fixed in TreeView and in Grid)
 
## Build 4.2.067
**2012-03-29**
- Fixed searching so it will find matches in the .h include filename in SCX forms 
- Can now be compiled to a free-standing .exe for use outside of VFP IDE
- Adjusted some label captions on main and Advanced form

## Build 4.2.066
**2012-02-24**
- Added necessary file extensions to backup DBCs and Labels during Replace

## Build 4.2.065
**2012-02-21**
- Fixed Search Scope to work with UNC paths i.e.  \\server_name\folder_name

## Build 4.2.064
**2012-02-21**
- Fixed Search Scope to work with mapped drive letters
 
## Build 4.2.063
**2012-02-21**
- Fixed SetScope() method to properly allow Search Scope to be set from an external call

## Build 4.2.062
**2012-02-20**
- Fixed Search Scope History combo to remember `\<\<Active Project>>` and `\<\<Current Dir>>`

## Build 4.2.061
**2012-02-14**
- Fixed bug in Search Scope combo when choosing `\<\<Active Project>>` or `\<\<Current Dir>>`
- Added `<Include File>` to the list of MatchTypes that can be replaced

## Build 4.2.060
**2012-02-10**
- Fixed Search Scope combo to better handle clicking, editing, pasting
- Added Try/Catch SearchInTable() method to guard against corrupt dbf/fpt files
 
## Build 4.2.059
**2012-02-08**
- You can now click to set focus into the Search Scope combo box
 
## Build 4.2.057
**2012-02-02**
- Set AllowOutput = .f. on main GoFish for to prevent from appearing in form.
- Added .Category setting to Thor_Tool_GoFish.prg to work with Thor 1.20 'Thor Tools' menu
  
## Build 4.2.056
**2012-02-02**
- Bug fixed in constructing Replace line text
 
## Build 4.2.055
**2012-02-01**
- Public release - See notes for 4.2.054 and 4.2.052

## Build 4.2.054
**2012-01-30**
- Corrected error message that is displayed when RegEx expression fails.
 
## Build 4.2.052
**2012-01-25**
- Corrected the way matches are found on Property Names and Property Values
- Added support to Browse a dbf when you double click a row in the grid. [GF_PEME_BaseTools.EditSourceX()]

## Previous releases
- Added "Data" node to the TreeView to host matches from DBF tables. 
- Fixed Progress Bar UI control so the meter bar would not expand beyond themframe.
- Added Up Arrow image to Project so it will build into the .app file. 
- Ver 4.2 - Released 2012-01-21
- Ver 4.1 Beta 1 - Released 2011-11-17
- Ver 4.1 Alpha 1  - Released 2011-08-11  (474 download)
- Ver 4.0 Beta 1 - Released 2011-06-07 GoFish 4 Beta 1 released on VFPx. (698 downloads)
- Ver 4.0 Beta - Released 2011-05-14
- Ver 4.0  Alpha - Released 2010-11-25 - By Matt Slay - Initial work on Ver 4.0 begins, 6 years after last release by Peter Diotte.
- Ver 3.1b - Released 2005-07-23
- Ver 3.0 - Released 2002-12-04
- Ver 2.0 - Released 2002-03-05 - Ability to do a REPLACE
- Ver 1.0 - Released 2001-03-01 - By Peter Diotte 
- Code named "Thong" (find the string...) 
- Park Ave Marketing firm suggests rename - "Go Fish"

----
Last changed: _2023/09/07_  ![Picture](./pictures/vfpxpoweredby_alternative.gif)
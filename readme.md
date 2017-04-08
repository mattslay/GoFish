# GoFish 5.0 
_ver 5.0.163   released 2017-02-12_

## GoFish is an advanced code search tool for fast searching and replacing of Visual FoxPro source code.

![Screenshot](Screenshots/GoFishScreenshot01.png?raw=true "Title")

What’s new in version 5:

* Saved Search History – You can now save the history of your searches, either automatically (for all searches) or selectively, and can restore the search parameters and results grid from these saved searches.  You can also selectively delete your saved searches or use the GF Janitor to automatically delete old ones.

* More powerful filter options – You can filter on secondary matches in the code “in the neighborhood” (that is, in the same statement or same procedure) of the original match; the filter form has been re-organized to provide clarity; and filters can be combined using logical operators AND or OR.

* Handling of PRG-based classes – Special attention has been paid to matches in PRG-based classes, which now are treated as much as possible like VCX-based classes. Their names appear in the “Class”, “Base Class”, and “Parent Class” columns in the grid and also in the new category “Classes” in the Treeview.

* Treeview changes – There are a few new categories: “Classes”, “Menus”, and “Projects”. The “Classes” category, unlike all the other categories which have files as sub-nodes, has classes as sub-nodes. 

* Column Changes  – There are two new columns, “Parent Class VCX” and “Containing Class”; the column headings for some of the other columns have been reworded; and the contents of some of the class-related columns have been enhanced.

* Plug-In to control the grid display – You can set the use the plug-in to set the Dynamic* (or any other) properties of the grid. The sample provided changes the colors used for each row.

* Other UI changes – There are many, many other UI changes, including the ability to select the location of the Code View window relative to the grid.

## Discussions

Post questions, bug reports, discussions in the <a href="http://groups.google.com/group/foxprogofish" target="_blank">GoFish Discussion Group</a>

## Videos

<a href="https://www.youtube.com/watch?v=0MdpWyPnfus" target="_blank">Getting started with GoFish</a>

## Audio Podcast
<a href="http://bit.ly/GoFish-On-TheFoxShow-Podcast-72" target="_blank">Hear The Fox Show podcast interview with GoFish author Matt Slay</a>

## Search Results Form

_{TODO: Insert screenshot link here}_

## New features in version 4.3…

* New  Wildcard Search Mode allows `*` and `?` in Search Expression.
* Wildcards (`*` and `?`) and the NOT operator (`!`) can be now used on the Filter Form.
* Added option for confirming "Whole Word" search if it is marked when starting a new search.
* Added option for clearing the "Apply Filter" checkbox with each new search.
* Added option to specify number of MRU entries to display in Scope/Search Expression dropdowns.
* Fixed Report so that it will correctly use the same filter that is in place on the main grid.

## Feature List
* Super FAST code searches are powered by the new GoFish Search Engine class.
* Search scope can be a Project or a Path.
* Recent Search Expressions and Search Scopes are maintained in dropdown combos.
* Can maintain a list of files to skip.
* Uses XML files to store settings and options between sessions. See GoFish Config Files Help for more info.
* Code Replacements - See GoFish Replace Help for more info.
* Performs backups during Replace operations  - See GoFish Backup Help for more info.
* Colorized code view and highlighting of matched line.
* TreeView provides quick filtering on the results grid by filetype or specific file.
* Filter Button provides advanced filtering across multiple columns after initial search if performed.
* Search Expression supports Regular Expressions (use Advanced button on main form).
* Form is resizable, and dockable (if desired).
* Column order and width can be changed.
* Columns are sortable.
* Window panes are resizable using the vertical and horizontal splitter controls.
* Double-click a row to open the file and edit the code in native FoxPro method windows.
* Integrates with PEM Editor IDE Tools to open files through Source Code Control Checkout.


## Learn more about...

* Searching in Reports: _{Link and page needed: GoFish 4 Searching in FoxPro Reports}_

* Search and Replace: _{Link and page needed: GoFish Replace Help}_

* Updating GoFish: _{Link and page needed: How to automatically update to the latest version}_

* Advanced _{Link needed to page: Searching with Wildcards and Regular Expressions}_

## Further options can be accessed on the Advanced Search form: 

* Supports searching Active Project, user selected Project, Current Directory, or user selected Directory.
* Allows searching on TimeStamp of changed objects on forms and classes, as well as file dates.
* Initial file type filtering limits which files are searched.
* Whole word/partial match, match case, include/exclude comments.
* Supports Regular Expressions.
* All user settings are saved between sessions, including column display, width, and order, as well as form size and position on the screen.
Can specify a filename temple (i.e.  `job*` )

## Technical References

* Search results cursor schema _{Link and page needed}_
* Using the GoFish Search Engine class to perform searches programatically _{Link and page needed}_

## Advanced Search Dialog

    _{TODO: Insert screenshot here}_
    
    
### Filename template

* Supports wildcard matching (`*` and `?`) on filename pattern to be searched.

    _{TODO: Insert screenshot here}_
    
## Filter Form

* Allows post-search filtering across multiple grid columns.
* Wildcards (* and ?) and the NOT operator (!) can be used on the Filter Form.

    _{TODO: Insert screenshot here }_
    
## Options Form 

* Column Selection - Allows choosing from over 20 fields of data to show in the Results Grid.
* Preferences - Basic user preferences for font size, docking, and message display and more.
* Advanced - View and manage the XML Config Files _{Line and page needed}_ used by GoFish to save settings.
* Replace – To enable Replace Mode _{Line and page needed}_ and learn how it works.
* Backups – To enable GoFish Backups _{Line and page needed}_ and learn how they work.
* Thor – Explains that GoFish will self-register with Thor to create a launch tool and Thor menu for GoFish.
* Update – Explains that GoFish can be updated using Thor “Check for Updates” feature.

    _{TODO: Insert screeshot here}_
    
## About screen contains additional links to the GoFish project

    _{TODO: Insert screenshot here}_

## Acknowledgments

* Thanks to original GoFish author, Peter Diotte, for granting me permission to create the version 4 update.
* Thanks to Jim Nelson for features and enhancements in Version 5.0

## Release history

* Ver 5.0.163 - released 2017-02-12
* Ver 5.0.162 - 2016-12-06
* Ver 4.3.012 – 2012-06-11
* Ver 4.3.011 – 2012-06-03
* Ver 4.3 - 2012-05-07
* Ver 4.2 - 2012-01-21
* Ver 4.1 Beta 1 - 2011-11-17
* Ver 4.1 Alpha 1  - 2011-08-11  (474 download)
* Ver 4.0 Beta 1 - 2011-06-07 GoFish 4 Beta 1 released on VFPx. (698 downloads)
* Ver 4.0 Beta - 2011-05-14
* Ver 4.0  Alpha - 2010-11-25 - By Matt Slay - Initial work on Ver 4.0 begins, 6 years after last release by Peter Diotte.
* Ver 3.1b - 2005-07-23
* Ver 3.0 - 2002-12-04
* Ver 2.0 - 2002-03-05 - Ability to do a REPLACE
* Ver 1.0 - 2001-03-01 - By Peter Diotte 
* Code named "Thong" (find the string...) 
* Park Ave Marketing firm suggests rename - "Go Fish"



    


    


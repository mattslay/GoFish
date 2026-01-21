# Home
Documentation of GoFish - An advanced code search tool for MS Visual Foxpro 9

## Purpose of this document
This document introduce the basic use of GoFish.

Basically this is the moved docu out of roots [readme.md](/readme.md), with all the missing document links.
In many parts this is a stub inherted.

See [changes](./Changes.md) for additional information.

There is a stub about using the [engine](/Screenshots/GoFishSearchEngine.md) without the UI.

----
## Table of contents
- [Feature List](#feature-list)
- [Learn more about](#learn-more-about)
- [Advanced Search form](#advanced-search-form)
  - [Advanced Search Dialog](#advanced-search-dialog)
- [Technical References](#technical-references)
- [Filename template](#filename-template)
- [Filter Form](#filter-form)
- [Options Form](#options-form)
- [About screen](#about-screen)
- [Parameter](#parameter)
- [Acknowledgments](#acknowledgments)

## Feature List
- Super FAST code searches are powered by the new GoFish Search Engine class.
- Search scope can be a Project or a Path.
- Recent Search Expressions and Search Scopes are maintained in dropdown combos.
- Can maintain a list of files to skip.
- Uses XML files to store settings and options between sessions. See GoFish Config Files Help for more info.
- Code Replacements - See GoFish Replace Help for more info.
- Performs backups during Replace operations  - See GoFish Backup Help for more info.
- Colorized code view and highlighting of matched line.
- TreeView provides quick filtering on the results grid by filetype or specific file.
- Filter Button provides advanced filtering across multiple columns after initial search if performed.
- Search Expression supports Regular Expressions (use Advanced button on main form).
- Form is resizable, and dockable (if desired).
- Column order and width can be changed.
- Columns are sortable.
- Window panes are resizable using the vertical and horizontal splitter controls.
- Double-click a row to open the file and edit the code in native FoxPro method windows.
- Integrates with PEM Editor IDE Tools to open files through Source Code Control Checkout.

## Learn more about
- Searching in Reports: _{Link and page needed: GoFish 4 Searching in FoxPro Reports}_

- Search and Replace: _{Link and page needed: GoFish Replace Help}_

- Updating GoFish: _{Link and page needed: How to automatically update to the latest version}_

- Advanced _{Link needed to page: Searching with Wildcards and Regular Expressions}_

## Advanced Search form 
Further options can be accessed on the Advanced Search form: 
- Supports searching Active Project, user selected Project, Current Directory, or user selected Directory.
- Allows searching on TimeStamp of changed objects on forms and classes, as well as file dates.
- Initial file type filtering limits which files are searched.
- Whole word/partial match, match case, include/exclude comments.
- Supports Regular Expressions.
- All user settings are saved between sessions, including column display, width, and order, as well as form size and position on the screen.
Can specify a filename temple (i.e.  `job*` )
### Advanced Search Dialog
    _{TODO: Insert screenshot here}_

## Technical References
- Search results cursor schema _{Link and page needed}_
- Using the GoFish Search Engine class to perform searches programatically _{Link and page needed}_

### Filename template
- Supports wildcard matching (`*` and `?`) on filename pattern to be searched.

    _{TODO: Insert screenshot here}_
    
## Filter Form
- Allows post-search filtering across multiple grid columns.
- Wildcards (* and ?) and the NOT operator (!) can be used on the Filter Form.

    _{TODO: Insert screenshot here }_
    
## Options Form 
- Column Selection - Allows choosing from over 20 fields of data to show in the Results Grid.
- Preferences - Basic user preferences for font size, docking, and message display and more.
- Advanced - View and manage the XML Config Files _{Line and page needed}_ used by GoFish to save settings.
- Replace – To enable Replace Mode _{Line and page needed}_ and learn how it works.
- Backups – To enable GoFish Backups _{Line and page needed}_ and learn how they work.
- Thor – Explains that GoFish will self-register with Thor to create a launch tool and Thor menu for GoFish.
- Update – Explains that GoFish can be updated using Thor "Check for Updates" feature.

    _{TODO: Insert screenshot here}_
    
## About screen
About screen contains additional links to the GoFish project
    _{TODO: Insert screenshot here}_

## Parameter
The GoFish app accepts some parameters, one at a time.
### /?
Help Screen, listing the parameters.
### -Reset
Reset the settings and storage in `Home(7)`  
This will not clear local storages. Local storages will be reused as soon as the option local is turned on and the app is restarted.
### -Resetlocal
Reset the settings and storage of the local ressource file. If local is not activated, no function.   
Note that the location of a local store is stored in the resource file, so the location of the local storage will be kept.
### -Clear
Clear the stored search and replace data and backup.
This clears the storage used by the project, or if a common storage is used, the common storage.
### -P
Load the active project as active scope.
### -F
Load the active folder as active scope.
### cProject
Load the project named as active scope.
### cFolder
Load the folder named as active scope.

For the *-P*, *-F*, *cProject* and *cFolder* option, the active scope will determine the history to load on startup,
if options *Restore previous search results on startup* and *Show history by scope* are activated.
## Acknowledgments
- Thanks to original GoFish author, Peter Diotte, for granting me permission to create the version 4 update.
- Thanks to Jim Nelson for features and enhancements in Version 5.0
- Thanks to late Matt Slay for creating the tool we see know


----
Last changed: _2023/03/08_  ![Picture](./pictures/vfpxpoweredby_alternative.gif)
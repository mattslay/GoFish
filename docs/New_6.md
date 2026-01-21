# Changes in 6.0
Documentation of GoFish - An advanced code search tool for MS Visual Foxpro 9

## Purpose of this document
This document list new function in Version 6.0.3.

----
## Table of contents
- [Basics](#basics)
- [Stored settings](#stored-settings)
  - [User storage](#user-storage)
  - [Project storage](#project-storage)
  - [Storage structure](#storage-structure)
- [Treeview](#treeview)
  - [History](#history)
    - [Show history](#show-history)
    - [Previous sessions](#previous-sessions)
  - [Simple tree](#simple-tree)
- [Replace](#replace)
- [Using](#using)
  - [Open on active scope](#open-on-active-scope)
  - [Autofilter by scope](#autofilter-by-scope)
    - [Search history by scope](#search-history-by-scope)
    - [Search combo](#search-combo)
  - [Pick folder as scope](#pick-folder-as-scope)
  - [Pick project as scope](#pick-project-as-scope)
  - [Refresh from tree](#refresh-from-tree)
  - [Delete from tree](#delete-from-tree)
  - [Grid](#grid)
    - [Visibility](#visibility)
    - [Replace back colour](#replace-back-colour)

## Basics
The changes of this version try add functions to GoFish, that are uesfull in *Code References*.
Mainly settings pre ressource and history in tree.   
A lot of boolean expressions in the form `lVar = .T.` or `lVar = .F.` are replaced with the short form.

## Stored settings
### User storage
The storage structure in version pre-dating 6.0 used a couple of files directly stored to HOME(7), and a subdirectory to store history.
For better clearance, the User storage is now moved to `HOME(7)+"\GoFish_"`.
Deleting the folder or it's contents as a whole resets GoFish to factory values.   
This could also be achieved using a startup parameter `"-Reset"`.   
This will not clear local storages. Local storages will be reused as soon as the option local is turned on and the app is restarted.
### Project storage
Depending on the personal style of development some prefer to have distinct IDE per project.
On this, it's very odd to mix up search history between projects.
A new option (Page: Advanced / *Local settings / storage*) allows to store depending of the resource file (FoxUser.dbf) loaded.
By default, it stores to `Justpath(Set("Resource",1))+"\GoFish_"`. The location might be altered on the same page.   
This could also be achieved using a startup parameter `"-Resetlocal"`.   
Since the location of a local store is stored in the resource file, so the location of the local storage will be kept.
#### Data need to use Project storage
The settings in `HOME(7)+"\GoFish_"` have a flag set to notify GoFish on startup to look for a record in the resource file.
The record is `Type="GoFish  " And Id="DirLoc  "`, the `Data` fields holds the folder storing the GoFish data.
### Storage structure
The storage  of historical data is now in a database rather then in a bunch of subdirectories.
There is still a bunch of files for each history record in the *GF_Saved_Search_Results* subfolder.
This is because the size of data stored in a memo will fast reach the 2GiB limit for files.   
In difference to *Code References* app, GoFish stores whole procedures per hit, and this might grow fast on failed search expressions.   
Backup of Replace operation will be stored to *GF_ReplaceBackups* folder, in a subfolder per replace operation.   
**Do not delete files out of these folders, delete via the history form, Right Click in the Treeview or use the Janitor to remove files.  
The storage data could be cleared by calling `Do GoFish5.app WITH "-Clear"` this will clear both search and replace storage in the folder used.**
#### Migration
GoFish will determine if folder `HOME(7)+"\GoFish_"` exists.
If not, it will start to transform and move settings from `HOME(7)` to `HOME(7)+"\GoFish_"`, deleting the old files.   
The new style will show replace information inline of historical search hits.
The migration can not merge replace history/ and search history,
so the replaces from the old version are available throug *View Replace History* only.
#### Backup
(Version 6.0.3)   
The files pre Version 6 will be duplicated to backup folder `Home(7) + "GoFish_Backup"`.
Deleting the folder `Home(7) + "GoFish_"` and moving these files back to `Home(7)` will allow to work with old versions.
Multiple updates will create numbered backup folders.
#### Parallel use of old versions
Since v6.0 GoFish will not migrate old data again as long as the folder `HOME(7)+"\GoFish_"` exists,
the use of older versions does not interfer with GoFish once migrated.

## Treeview
Some changes are made to allow the Treeview better fit to personal style.
While the GoFish style offers good access, the style of the *Code References* app has it's good points as well.
This is in special (All settings on Page: Code References)
### Context menu
The tree offers a context menu to clear or refresh searches, or to change the sort order.   
It is possible to refresh all previous (and visible) searches, but keep in mind this could take a while.
Any way, I found it very usefull while reworking projects.   
If the grid is showing the replace history, the function will clear replace backups.   
Clearing search or clearing replace will not interfer with the other,
records that are both saved and replaced will only be filtered from the respective view.
### History
Allows the tree to show old searches. Note that the result set might be limited by active scope, see [below](#search-history).
#### Show history
Option *Show history in tree*. Like in *Code References*, the tree will show old searches.   
Just recent session or stored history, see below. Set the *Show History in Tree* option. There are two sub options.
- *Fill all searches with data* The history is stored without the code snippets, the snippets will be loaded on need.
This might slow down the actual response of the app. The switch will speed this up on coast of initial load time.
**Keep in mind that this could violate the 2GiB limit.**
- *Sort alphanumerical* By default the results are sorted by search time ascending.
#### Previous sessions
For full history as in *Code References*, additionally the following options must be active:
- Page: Preferences / *Save each search result to search history folder*, otherwise only manually stored searches
- Page: Preferences / *Restore previous search results on startup*   
### Simple tree
Option *Simple tree, ...* to reduce the search nodes to *Code References* level.
A sub option *Sort Tree by extension first* allows to group the results a bit.

## Replace
The replace function is only active for the recent (Not: most recent) search.   
New: The restore history will follow the search history, so one can see previous replaces on history searches.
### Altered functions
The *Replace View* Checkbox in the bar above the grid opens a replace area. This is independend of enableing the Replace mode.
If the view is visible, previously replaced records in the active or in in restored searches are highlighted inside the grid.
Setting the Radio Box to *View Replace History* switches the whole form to display previous replaces.   
The *Preview* allows to see the result of the *Replace Text* function before replacing.
The hold version used this as the only way, but it was a bit slow on large result sets.
### Grid context menu
If a record in the grid is replaced and a backup is created, the backup-folder of the replace is available via the contextmenu of the grid.
This is independend of the state of *Replace View*.

## Using
### Pick a scope
Right of the scope combo are two objects to pick a folder or a project as scope.
### Open on active scope
Like in *Code References*, GoFish might open with the active project or the the recent folder as scope if no project is open.
The option (Page: Code References / *Open GoFish like Code References: Active project / Recent folder*) will turn this on.
### Autofilter by scope
Two functions allow to filter by the active scope
#### Search history by scope
The option (Page: Code References / *Show search history (button) per scope*) limits the results shown in the History form to those of the recent scope.   
Additionally the results in the Treeview will be limited to the same set.
#### Search combo
The option (Page: Code References / *Show history in search combo per scope*) limits the previous searches shown in the search combo those used with the recent scope.   
### Pick folder as scope
A new folder as search scope can be picked on clicking the folder icon right of the scope combo.
### Pick project as scope
A project as search scope can be picked on clicking the project icon right of the scope combo.
### Refresh from tree
A single search or all searches can be rerun by context menu of a Treeview-node.
### Delete from tree
A single search or all searches can be deleted from history by context menu (*Clear*) of a Treeview-node.
The *Clear All* function respects the set of [Search history by scope](#search-history-by-scope) option.   
Note that *Clear* and *Clear All* only clear the data of the active mode, "Search History" or "Replace History".
### Grid
The columns of the grid are reordered. 
The previous way to create the grid was very uncomfortable when changing the table structure,
this is changed from visual to code based.   
Due to this, the column information stored will not survive version change.   
#### Visibility   

The page *Column Selection* on options form is reordered and suppresses fields with internal function.
They could be turned on by a change in source code, there is a define `dlDebug` in `GF_ResultsForm.FormatGrid`.   
A way to set the grid to default is added. Fields belonging to the replace function are highlighted in *Replaced* back colour.   
If you think you need a certain field, please post an [issue](https://github.com/VFPX/GoFish/issues).   
The full list of fields, cyan is turned off by default, reds are just internal data, normaly turned off:   
![File list](./Screenshots/GF6_AllFields.png)

----
Last changed: _2023/03/12_  ![Picture](./pictures/vfpxpoweredby_alternative.gif)
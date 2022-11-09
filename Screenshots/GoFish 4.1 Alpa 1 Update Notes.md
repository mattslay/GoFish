**GoFish 4.1 Alpha 1 Update Notes:**

Because this build includes a new, powerful feature for REPLACING code on your source files, it is considered an Alpha version at this time. Please help by testing out this build and reporting any problems.

Note: To perform Replace actions with GoFish you must enable this feature on the Replace tab of the Options page. Also, [BE SURE TO READ ALL THE NOTES ABOUT THE RISKS OF DOING REPLACES](GoFish-Replace-Help), for important information about using GoFish to perform replaces on your code.

**New features in 4.1 Alpha 1**
 1. Added new feature to perform mass Replaces on code search matches.
 2. Added new feature to perform single line edits on code search matches.
 3. Added a Progress Bar to the main form to show progress of search.
 4. The Search Expression can now contain preceding and trailing spaces.
 5. Enhanced search results are now gathered when searching FRX Report files.


**Changes / Fixes from Ver 4 Beta 1:**

 1. Fixed FileDate info on text-based files.
 2. Fixed Ascending/Descending sorting when clicking on column headers.
 3. Now using BindEvent to bind grid headers to SortColumn() method.
 4. Fixed small bug in file template matching on Advanced page.
 5. The Results form and the Search form now use an EditBox for the Search text, rather than TextBoxes.
 6. The My.Settings.Load() method has a custom code line from the publicly released version to prevent the Load() method from stripping of whitespaces. 
 7. Added new MatchType to replace Reserved7 with <Class Desc>
 8. Added new MatchType to replace Reserved8 with <Include File>
 9. Fixed a bug to prevent crashing when attempting to search a text file that is locked by another app.
10. Fixed Regular Expression searching to automatically include 'beginning of line' and 'end of line' to the entered search expression.
 



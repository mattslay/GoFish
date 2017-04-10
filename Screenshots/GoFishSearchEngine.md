# GoFish Search Engine Class
_This class is available in the source code download of [Release: GoFish](67852)._

Most people don't realize that GoFish 4 is really just a pretty UI form wrapped around a powerful Search Engine Class. 

With the GoFish Search Engine Class, you can programmatically perform searches on code paths or projects, or any blob of text from anywhere you can dig it up. So, using the GoFish Search Engine Class, you can add complex searching to any business app or other FoxPro IDE tools, without ever using the default UI form that we use from the FoxPro IDE.

It basically works like this:

{{
*-- Create and object instance of the GoFish Search Engine class--
loSearchEngine = CreateObject('GoFishSearchEngine')
loSearchEngine.cSearchResultsAlias = 'MyCursor'          
 
*-- Configure it--
With loSearchEngine.oSearchOptions
     .cSearchExpression = 'test'
     .lRegularExpression = .f.
     .lMatchWholeWord = .t.
     .lMatchCase = .f.
     .lSearchInComments = .t.
     
     .lCreateResultsCursor = .t.
     .lCreateResultsCollection = .t.

     .lShowErrorMessages = .f. && Controls displaying errors in a MessageBox
     .lShowWaitMessages = .f.
     .lShowNoMatchesMessage = .f.
EndWith

*-- Perform a search -----------------
 loSearchEngine.oSearchOptions.lIncludeSubdirectories = .t.
 loSearchEngine.SearchInPath(‘C:\Source\Project1\’)
}}

**What next?**
Well, all the search results from that search now live in the cursor ‘MyCursor’ that **you** specified. Now you can do whatever you want with that cursor. And, if you want to perform another search (for a different word), all you've got to to is change the property _loSearchEngine.oSearchOptions.cSearchExpression_ to a new value, and run the search again. The cursor will be rebuilt under the same name:

{{ 
loSearchEngine.oSearchOptions.cSearchExpression = 'test2'
loSearchEngine.SearchInPath(‘C:\Source\Project1\’)
}}

The GFSE can also create a Collection object with the results if you prefer to work with collections rather than cursors.

**See this link for more samples:**  [http://codepaste.net/8srdb8](http://codepaste.net/8srdb8)


**Here is a complete list of the Properties on GFSE:**

![](GoFishSearchEngine_ GFSE_Properties.jpg)



**Here is a complete list of the Methods on GFSE:**

![](GoFishSearchEngine_ GFSE_Methods.jpg)


**Finally, here is the properties on the Options class used by GFSE to fine tune the search:**

![](GoFishSearchEngine_ GFSE_Options.jpg)

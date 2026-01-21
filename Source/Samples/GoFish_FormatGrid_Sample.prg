*-- Sample to format GoFish grid color. Provided by Jim R. Nelson
*-- Assigns row colors based on field MatchType:
*-- Certain color for Comments, bold for Filenames and Class Names, red for Procedures, Functions, Methods, etc.


Lparameters toGrid, tcResultsCursor

Local lcComments, lcCommentsColor, lcDefs, lcDefsColor, lcOthers, lcOthersColor, lcProcs
Local lcProcsColor

lcProcs         = 'Inlist(' + tcResultsCursor + '.MatchType, "<Method>", "<Procedure>", "<Function>")'
lcProcsColor = 'RGB(255,0,0)'

lcDefs        = 'Left(' + tcResultsCursor + '.MatchType, 2) = "<<"'
lcDefsColor    = 'RGB(0,0,0)'

lcOthers      = 'Left(' + tcResultsCursor + '.MatchType, 1) = "<"'
lcOthersColor = 'RGB(0,128,0)'

toGrid.SetAll ('DynamicForeColor', 'ICase(' +    ;
      lcProcs + ', ' + lcProcsColor + ', ' +    ;
      lcDefs + ', ' + lcDefsColor + ', ' +        ;
      lcOthers + ', ' + lcOthersColor + ', ' +    ;
      ' Rgb(0,0,0))')

lcComments        = tcResultsCursor + '.MatchType = "<Comment>"'
lcCommentsColor    = 'RGB(224,224,224)'

toGrid.SetAll ('DynamicBackColor', 'ICase(' +        ;
      lcComments + ', ' + lcCommentsColor + ', ' +    ;
      ' Rgb(255,255,255))')

toGrid.SetAll ('DynamicFontBold', 'Left(' + tcResultsCursor + '.MatchType, 2) = "<<"')


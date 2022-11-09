********************************************************************
* mhHtmlCode by Mike Helland  - mobydikc@gmail.com
*
*-- See: http://www.universalthread.com/ViewPageNewDownload.aspx?ID=9679
*
* This takes code and makes it turn colors
* You can use it for your private use
*
* Whats special about this version?
* The performance has gotten better. On my PC, 3420 lines in 1.07 seconds
* I used alot of Compile time stuff this time
* I don't know if it helps or not, but it does look cool
*
* Its also OO (will, not really, since the bulk of it can't be easily
* broken into seperate methods) but it is based on an object, so it can
* be compiled as a COM or a Web Service.
*
* And, as an object, I can open and close the words table on once
* but run the functionality multiple times.
*
* There are a couple core functionality things fixed to
* But enough about me...

* This is the mode we run in
#DEFINE cnMODE "CSS"

* Tag Constants
#IF cnMODE = "CSS"
	#DEFINE ccROOT		"<pre><span class='vfpcode'>"
	#DEFINE ccCOMMENT	"<span class='vfpcomment'>"
	#DEFINE ccRESERVED	"<span class='vfpreserved'>"
	#DEFINE ccLITERAL	"<span class='vfpliteral'>"
	#DEFINE ccSTRING	"<span class='vfpstring'>"
	#DEFINE ccVARIABLE	""
	#DEFINE ccROOT_CLOSE		"</span></pre>"
	#DEFINE ccCOMMENT_CLOSE		"</span>"
	#DEFINE ccRESERVED_CLOSE	"</span>"
	#DEFINE ccLITERAL_CLOSE		"</span>"
	#DEFINE ccSTRING_CLOSE		"</span>"
	#DEFINE ccVARIABLE_CLOSE	""
	#DEFINE cnCOMMENT_LEN	25
	#DEFINE cnRESERVED_LEN	26
	#DEFINE cnLITERAL_LEN	25
	#DEFINE cnSTRING_LEN	24
	#DEFINE cnVARIABLE_LEN	0
	#DEFINE cnCOMMENT_CLOSE_LEN		7
	#DEFINE cnRESERVED_CLOSE_LEN	7
	#DEFINE cnLITERAL_CLOSE_LEN		7
	#DEFINE cnSTRING_CLOSE_LEN		7
	#DEFINE cnVARIABLE_CLOSE_LEN	0
#ELSE
	#DEFINE ccROOT		"<pre>"
	#DEFINE ccCOMMENT	"<font color='green'>"
	#DEFINE ccRESERVED	"<font color='blue'>"
	#DEFINE ccLITERAL	"<font color='red'>"
	#DEFINE ccSTRING	"<font color='red'>"
	#DEFINE ccVARIABLE	""
	#DEFINE ccROOT_CLOSE		"</pre>"
	#DEFINE ccCOMMENT_CLOSE		"</font>"
	#DEFINE ccRESERVED_CLOSE	"</font>"
	#DEFINE ccLITERAL_CLOSE		"</font>"
	#DEFINE ccSTRING_CLOSE		"</font>"
	#DEFINE ccVARIABLE_CLOSE	""
	#DEFINE cnCOMMENT_LEN	20
	#DEFINE cnRESERVED_LEN	19
	#DEFINE cnLITERAL_LEN	18
	#DEFINE cnSTRING_LEN	18
	#DEFINE cnVARIABLE_LEN	0
	#DEFINE cnCOMMENT_CLOSE_LEN		7
	#DEFINE cnRESERVED_CLOSE_LEN	7
	#DEFINE cnLITERAL_CLOSE_LEN		7
	#DEFINE cnSTRING_CLOSE_LEN		7
	#DEFINE cnVARIABLE_CLOSE_LEN	0
#ENDIF

* Other ones
#DEFINE ccCRLF		CHR(13) + CHR(10)
#DEFINE ccVERSION = "4"

DEFINE CLASS htmlcode AS Session OLEPUBLIC

	nSeconds = 0
	nLines = 0

	PROCEDURE PRGToHTML(tcCode AS String) AS String 

		* Performance checker
		this.nSeconds = SECONDS()

		local lcCode, ;
			lcCRLF, ;
			lnLines, ;
			laLines[1], ;
			lcReturn, ;
			lnI, ;
			lcLine, ;
			lcNoTabLine, ;
			lnStartComment, ;
			lcInlineComment, ;
			lnStuffedChars, ;
			llString, ;
			lcTempLine, ;
			lnOffset, ;
			lnWords, ;
			laWords[1], ;
			lnWord, ;
			lnWordStart, ;
			lcWord, ;
			lnWordLen, ;
			lcEndString

		* Do the CRLF now for small performance reasons
		lcCRLF = ccCRLF

		* if no code was passed, return blank
		lcCode = tcCode
		if empty(lcCode)
			return ''
		endif

		* Coming in as a web service, there shoudl be no CRs and lots of LFs, change them
		*Give Alines() something consitant
		IF AT(CHR(13), lcCode) = 0 AND AT(CHR(10), lcCode) > 0
			lcCode = strtran(lcCode, CHR(10), CHR(13) + CHR(10))
		ELSE 
			lcCode = strtran(chrtran(lcCode, CHR(10), ''), CHR(13), CHR(13) + CHR(10))
		ENDIF 

		*Do some basic HTML intializing
		lcCode = strtran(lcCode, '&', '&amp;')
		lcCode = strtran(lcCode, '<', '&lt;')
		lcCode = strtran(lcCode, '>', '&gt;')

		*This is one is diffent, so its recognized as a word
		lcCode = strtran(lcCode, '[', ' [')

		*Create an array element for each line
		lnLines = alines(laLines, lcCode)
		this.nLines = lnLines
		lcReturn = ''

		*Proccess each line
		for lnI = 1 to lnLines
			lcLine = laLines[lnI]

			*Don't even bother the blank ones
			if empty(alltrim(lcLine))
				lcReturn = lcReturn + lcCRLF
				loop
			endif

			* Full Line Comments, first are the first test
			lcNoTabLine = LEFT(upper(ltrim(chrtran(lcLine, CHR(9), ''))), 1)
			if lcNoTabLine = '*' or ;
				lcNoTabLine = 'NOTE'
				#IF cnCOMMENT_LEN > 0
					lcLine = ccCOMMENT + lcLine + ccCOMMENT_CLOSE
				#ENDIF 

				*And move on
				lcReturn = lcReturn + lcCRLF + lcLine
				loop
			endif

			* Now end of the line comments
			lnStartComment = at('&amp;&amp;', lcLine)
			IF lnStartComment > 0
				lcInLineComment = ccCOMMENT + SUBSTR(lcLine, lnStartComment) + ccCOMMENT_CLOSE
				lcLine = substr(lcLine, 1, lnStartComment - 1)
			ELSE 
				lcInLineComment = ''
			ENDIF

			* Prerun fun
			lnStuffedChars = 0
			llString = .F.

			*Lets loop through every word in our line. Break the line into words
			lcTempLine = ' ' + ltrim(chrtran(upper(lcLine), ;
				'~!@#$%^&*()-+=|\{}:;,./\<>?' + CHR(9), space(28))) + ' '

			* Offset it the number of tabs and spaces leading the line
			lnOffset = len(lcLine) - len(lcTempLine) + 1

			* Run through each words
			lnWords = occurs(' ', lcTempLine) - 1
			dimension laWords[lnWords, 2]
			for lnWord = 1 to lnWords

				* Deteremine the word and where it starts and ends
				lnWordStart = at(' ', lcTempLine, lnWord) + 1 + lnOffset
				lcWord = substr(lcTempLine, ;
					lnWordStart - lnOffSet, ;
					at(' ', lcTempLine, lnWord + 1) - lnWordStart + lnOffSet)
				lnWordLen = len(lcWord)

				* See what type of word we got
				DO CASE
					CASE EMPTY(lcWord)

					* If we're in a string, don't color the words
					CASE llString

						* If this is the end of the string, set it
						IF RIGHT(lcWord, 1) = lcEndString
							llString = .F.
							#IF cnSTRING_LEN > 0
								lcLine = stuff(lcLine, lnWordStart + lnStuffedChars + lnWordLen, ;
									0, ccSTRING_CLOSE)
								lnStuffedChars = lnStuffedChars + cnSTRING_CLOSE_LEN
							#ENDIF 
						ENDIF 

					* See if the word is the beginning of a string using quotes (first one)
					* or brackets (second case)
					CASE INLIST(lcWord, '"', "'")

						* Turn the string flag on, so we don't color words in strings
						llString = .T.
						lcEndString = LEFT(lcWord, 1)

						#IF cnSTRING_LEN > 0

							* Add the string tag
							lcLine = stuff(lcLine, lnWordStart + lnStuffedChars, 0, ccSTRING)
							lnStuffedChars = lnStuffedChars + cnSTRING_LEN

							* If a string is the beginning and end of a word, end it now it
							IF lnWordLen > 1 and RIGHT(lcWord, 1) = lcEndString
								llString = .F.
								lcLine = stuff(lcLine, lnWordStart + lnStuffedChars + lnWordLen, ;
									0, ccSTRING_CLOSE)
								lnStuffedChars = lnStuffedChars + cnSTRING_CLOSE_LEN
							ENDIF 
						#ENDIF 
					CASE lcWord = '['
						llString = .T.
						lcEndString = ']'
						* This one is special, we need to insert it before the preceding space
						#IF cnSTRING_LEN > 0
							lcLine = stuff(lcLine, lnWordStart + lnStuffedChars - 1, 0, ccSTRING)
							lnStuffedChars = lnStuffedChars + cnSTRING_LEN
							IF lnWordLen > 1 and RIGHT(lcWord, 1) = lcEndString
								llString = .F.
								lcLine = stuff(lcLine, lnWordStart + lnStuffedChars + lnWordLen, ;
									0, ccSTRING_CLOSE)
								lnStuffedChars = lnStuffedChars + cnSTRING_CLOSE_LEN
							ENDIF 
						#ENDIF 

					* The word is a literal
					CASE ISDIGIT(lcWord) and TYPE(lcWord) = 'N'
						#IF cnLITERAL_LEN > 0
							* Insert the tags and bump the counter
							lcLine = stuff(lcLine, lnWordStart + lnStuffedChars, 0, ccLITERAL)
							lcLine = stuff(lcLine, lnWordStart + lnStuffedChars + ;
								cnLITERAL_LEN + lnWordLen, 0, ccLITERAL_CLOSE)
							lnStuffedChars = lnStuffedChars + cnLITERAL_LEN + cnLITERAL_CLOSE_LEN
						#ENDIF

					* This is a non colored operator
					CASE INLIST(UPPER(lcWord) + SPACE(1), 'AND', 'OR', 'NOT', 'NULL')
						* Don't do anything
					
					* Our word is a reserved word
					CASE seek(iif(lnWordLen < 4, padr(lcWord, 4), lcWord), 'words', 'revword')

						#IF cnRESERVED_LEN > 0
							* Insert the tags and bump the counter
							lcLine = stuff(lcLine, lnWordStart + lnStuffedChars, 0, ccRESERVED)
							lcLine = stuff(lcLine, lnWordStart + lnStuffedChars + ;
								cnRESERVED_LEN + lnWordLen, 0, ccRESERVED_CLOSE)
							lnStuffedChars = lnStuffedChars + cnRESERVED_LEN + cnRESERVED_CLOSE_LEN
						#ENDIF 

					* Must be a variable
					OTHERWISE
						#IF cnVARIABLE_LEN > 0
							* Insert the tags and bump the counter
							lcLine = stuff(lcLine, lnWordStart + lnStuffedChars, 0, ccVARIABLE)
							lcLine = stuff(lcLine, lnWordStart + lnStuffedChars + ;
								cnVARIABLE_LEN + lnWordLen, 0, ccVARIABLE_CLOSE)
							lnStuffedChars = lnStuffedChars + cnVARIABLE_LEN + ccVARIABLE_CLOSE_LEN
						#ENDIF

				ENDCASE 
			endfor	

			* If a string was left open, close it
			IF llString
				lcLine = lcLine + ccSTRING_CLOSE
			ENDIF 

			* Finish out the line
			lcReturn = lcReturn + lcCRLF + lcLine + lcInLineComment

		endfor

		* Revert this strange 
		lcReturn = strtran(lcReturn, ' [', '[')

		* These really slow down the process, choosing a blank Variable color is best for performance
		#IF cnVARIABLE_LEN > 0
			lcReturn = STRTRAN(lcReturn, '&' + ccVARIABLE + 'lt' + ccVARIABLE_CLOSE + ';', '&lt;')
			lcReturn = STRTRAN(lcReturn, '&' + ccVARIABLE + 'gt' + ccVARIABLE_CLOSE + ';', '&gt;')
			lcReturn = STRTRAN(lcReturn, '&' + ccVARIABLE + 'amp' + ccVARIABLE_CLOSE + ';', '&amp;')
			IF lnLiteralLen > 0
				lcReturn = STRTRAN(lcReturn, '.' + ccVARIABLE + 'T' + ccVARIABLE_CLOSE + '.', ;
					ccLITERAL + '.T.' + ccLITERAL_CLOSE)
				lcReturn = STRTRAN(lcReturn, '.' + ccVARIABLE + 'F' + ccVARIABLE_CLOSE + '.', ;
					ccLITERAL + '.F.' + ccLITERAL_CLOSE)
				lcReturn = STRTRAN(lcReturn, '.' + ccVARIABLE + 't' + ccVARIABLE_CLOSE + '.', ;
					ccLITERAL + '.t.' + ccLITERAL_CLOSE)
				lcReturn = STRTRAN(lcReturn, '.' + ccVARIABLE + 'f' + ccVARIABLE_CLOSE + '.', ;
					ccLITERAL + '.f.' + ccLITERAL_CLOSE)
			ELSE 
				lcReturn = STRTRAN(lcReturn, '.' + ccVARIABLE + 'T' + ccVARIABLE_CLOSE + '.', '.T.')
				lcReturn = STRTRAN(lcReturn, '.' + ccVARIABLE + 'F' + ccVARIABLE_CLOSE + '.', '.F.')
				lcReturn = STRTRAN(lcReturn, '.' + ccVARIABLE + 't' + ccVARIABLE_CLOSE + '.', '.t.')
				lcReturn = STRTRAN(lcReturn, '.' + ccVARIABLE + 'f' + ccVARIABLE_CLOSE + '.', '.f.')
			ENDIF 	
		#ENDIF 

		*Return the orginal value so we can pass by ref if we want
		tcCode = ccROOT + lcReturn + ccROOT_CLOSE

		this.nSeconds = seconds() - this.nSeconds

	RETURN tcCode

	PROCEDURE GetVersion AS String 
	RETURN ccVERSION

	PROCEDURE Init

		* The words table should be in the same directory as this program
		LOCAL lcWordsTable
		lcWordsTable = IIF(INLIST(_vfp.StartMode, 2, 3, 5), JUSTPATH(_vfp.ServerName) + '\', '') + 'WORDS.DBF'
		IF FILE(lcWordsTable)
			use (lcWordsTable) in 0 shared ALIAS words
		ENDIF
		IF not USED('words')
			RETURN .F.
		ENDIF 

	ENDPROC 

	PROCEDURE Destroy
		USE IN SELECT('words')
	ENDPROC 

ENDDEFINE 
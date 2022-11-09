* Include other include files.

#include FoxPro.H
#include MyConstants_Loc.H

* Control character constants.

#define ccCR                    	chr(13)
#define ccNULL                  	chr(0)

* VFP error numbers:

#define cnERR_FILE_NOT_FOUND		   1
	* File does not exist
#define cnERR_ARGUMENT_INVALID		  11
	* Function argument value, type, or count is invalid
#define cnERR_ALIAS_NOTFOUND		  13
	* Alias is not found
#define cnERR_INVALID_PATH_OR_FILE   202
	* Invalid path or filename
#define cnERR_PROPERTY_READ_ONLY	1743
	* Property is read-only
#define cnERR_KEY_EXISTS			2062
	* The specified key already exists

* .NET System Data Types

#define ccSTRING					'System.String'
#define ccINTEGER					'System.Integer'
#define ccDOUBLE					'System.Double'
#define ccCURRENCY					'System.Decimal'
#define ccBOOLEAN					'System.Boolean'
#define ccDATETIME					'System.DateTime'

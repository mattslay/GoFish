#DEFINE dcRegExpVerNo "1.2.0"
LPARAMETERS;
 tcPathTo_wwDotnetBridge

RETURN CREATEOBJECT("SF_RegExp",m.tcPathTo_wwDotnetBridge)

DEFINE CLASS SF_RegExp AS SESSION
*Internal object
*PRIVATE
 PROTECTED rnOffset  			&& Internal, string offset to a start string VFP like
 PROTECTED roRegExp 			&& Internal, reference to the C# DotNet wrapper.
 rnOffset = 1					
 roRegExp = .NULL.

*Return FoxObjects instead of DotNet ones (slower)
 ReturnFoxObjects   = .F.		&& Controls type of objects returned
 AutoExpandGroup    = .F.		&& Controls type of objects returned
 AutoExpandCaptures = .F.		&& Controls type of objects returned
 DotNetOffset       = .F.		&& Strings should start with 0 or 1

*Pattern
 PATTERN = ""					&& The DotNet Pattern property
*RegExp Options
 Compiled                = .F.	&& The Compiled part of the DotNet Options property   8
 CultureInvariant        = .F.	&& The CultureInvariant part of the DotNet Options property 512
 ECMAScript              = .F.	&& The ECMAScript part of the DotNet Options property 256
 ExplicitCapture         = .F.	&& The ExplicitCapture part of the DotNet Options property   4
 IgnoreCase              = .F.	&& The IgnoreCase part of the DotNet Options property   1
 IgnorePatternWhitespace = .F.	&& The IgnorePatternWhitespace part of the DotNet Options property  32
 MultiLine               = .F.	&& The MultiLine part of the DotNet Options property   2
 RIGHTTOLEFT             = .F.	&& The RightToLeft part of the DotNet Options property  64
 SingleLine              = .F.	&& The SingleLine part of the DotNet Options property  16

*Others
 CacheSize    = 0
 MatchTimeout = 0

 PROCEDURE INIT
  LPARAMETERS;
   tcPathTo_SF_RegExp
  LOCAL;
   lcOldPath   AS STRING,;
   lcPath      AS STRING,;
   llError     AS BOOLEAN,;
   loBridge    AS OBJECT,;
   loException AS EXCEPTION

  lcOldPath = FULLPATH("","")

  lcPath = SYS(16)
  lcPath = RIGHT(m.lcPath,LEN(m.lcPath)-AT(" ",m.lcPath,2))
  lcPath = JUSTPATH(m.lcPath)


  IF VARTYPE(m.__SF_REGEXPATH)="C" THEN
   IF EMPTY(m.tcPathTo_SF_RegExp) THEN
    tcPathTo_SF_RegExp = m.__SF_REGEXPATH
   ELSE  &&Vartype(__SF_REGEXPATH)="C"
    __SF_REGEXPATH = m.tcPathTo_SF_RegExp
   ENDIF &&Vartype(__SF_REGEXPATH)="C"

  ELSE  &&Vartype(__SF_REGEXPATH)="C"
   IF EMPTY(m.tcPathTo_SF_RegExp) THEN
    tcPathTo_SF_RegExp = m.lcPath
   ENDIF &&EMPTY(m.tcPathTo_SF_RegExp)
   PUBLIC __SF_REGEXPATH
   __SF_REGEXPATH = m.tcPathTo_SF_RegExp

  ENDIF &&Vartype(__SF_REGEXPATH)="C"

  CD (m.tcPathTo_SF_RegExp)

  TRY
    DO wwDotnetBridge

   CATCH
    TRY
      SET PATH TO ('"'+m.tcPathTo_SF_RegExp+'"') ADDITIVE
      DO wwDotnetBridge

     CATCH
      llError = .T.
    ENDTRY
  ENDTRY

  IF m.llError THEN
   CD (m.lcOldPath)
   RETURN .F.
  ENDIF &&m.llError

  loBridge = GetwwDotnetBridge()

  TRY
*Was zu testen ist
    loBridge.LoadAssembly("SF_RegExp.dll")

    THIS.roRegExp = loBridge.CreateInstance("SF_RegExp.SF_RegExp")

   CATCH
    TRY
      loBridge.LoadAssembly(ADDBS(m.tcPathTo_SF_RegExp)+"SF_RegExp.dll")

      THIS.roRegExp = loBridge.CreateInstance("SF_RegExp.SF_RegExp")

     CATCH
      llError = .T.
    ENDTRY
  ENDTRY

  IF m.llError THEN
   CD (m.lcOldPath)
   RETURN .F.
  ENDIF &&m.llError

  IF ISNULL(THIS.roRegExp) THEN
   CD (m.lcOldPath)
   RETURN .F.
  ENDIF &&ISNULL(THIS.roRegExp)

  CD (m.lcOldPath)
 ENDPROC &&Init

* Assign Access
 PROCEDURE DotNetOffset_Assign	&& Write if strings should start with 0 or 1
  LPARAMETERS;
   tvValue

  THIS.roRegExp.DotNetOffset = m.tvValue
  THIS.rnOffset              = IIF(m.tvValue,0,1)
 ENDPROC &&DotNetOffset_Assign

 PROCEDURE VERSION &&Return version
  RETURN dcRegExpVerNo
 ENDPROC &&Version

 PROCEDURE Pattern_Access		&& Access the DotNet Pattern property
  RETURN THIS.roRegExp.PATTERN
 ENDPROC &&Pattern_Access
 PROCEDURE Pattern_Assign		&& Assign the DotNet Pattern property
  LPARAMETERS;
   tvValue

  THIS.roRegExp.PATTERN = m.tvValue
 ENDPROC &&Pattern_Assign

 PROCEDURE CacheSize_Access		&& Access the CacheSize part of the DotNet Options property
  RETURN THIS.roRegExp.CacheSize
 ENDPROC &&CacheSize_Access
 PROCEDURE CacheSize_Assign		&& Assign the CacheSize part of the DotNet Options property
  LPARAMETERS;
   tvValue

  THIS.roRegExp.CacheSize = m.tvValue
 ENDPROC &&CacheSize_Assign

 PROCEDURE MatchTimeout_Access		&& Access the MatchTimeout part of the DotNet Options property
  RETURN THIS.roRegExp.MatchTimeout
 ENDPROC &&MatchTimeout_Access
 PROCEDURE MatchTimeout_Assign		&& Assign the MatchTimeout part of the DotNet Options property
  LPARAMETERS;
   tvValue

  THIS.roRegExp.MatchTimeout = m.tvValue
 ENDPROC &&MatchTimeout_Assign

*RegExp Options
 PROCEDURE Compiled_Access		&& Access the Compiled part of the DotNet Options property
  RETURN THIS.roRegExp.Compiled
 ENDPROC &&Compiled_Access
 PROCEDURE Compiled_Assign		&& Assign the Compiled part of the DotNet Options property
  LPARAMETERS;
   tvValue

  THIS.roRegExp.Compiled = m.tvValue
 ENDPROC &&Compiled_Assign

 PROCEDURE CultureInvariant_Access		&& Access the CultureInvariant part of the DotNet Options property
  RETURN THIS.roRegExp.CultureInvariant
 ENDPROC &&CultureInvariant_Access
 PROCEDURE CultureInvariant_Assign		&& Assign the CultureInvariant part of the DotNet Options property
  LPARAMETERS;
   tvValue

  THIS.roRegExp.CultureInvariant = m.tvValue
 ENDPROC &&CultureInvariant_Assign

 PROCEDURE ECMAScript_Access		&& Access the ECMAScript part of the DotNet Options property
  RETURN THIS.roRegExp.ECMAScript
 ENDPROC &&ECMAScript_Access
 PROCEDURE ECMAScript_Assign		&& Assign the ECMAScript part of the DotNet Options property
  LPARAMETERS;
   tvValue

  THIS.roRegExp.ECMAScript = m.tvValue
 ENDPROC &&ECMAScript_Assign

 PROCEDURE ExplicitCapture_Access		&& Access the ExplicitCapture part of the DotNet Options property
  RETURN THIS.roRegExp.ExplicitCapture
 ENDPROC &&ExplicitCapture_Access
 PROCEDURE ExplicitCapture_Assign		&& Assign the ExplicitCapture part of the DotNet Options property
  LPARAMETERS;
   tvValue

  THIS.roRegExp.ExplicitCapture = m.tvValue
 ENDPROC &&ExplicitCapture_Assign

 PROCEDURE IgnoreCase_Access		&& Access the IgnoreCase part of the DotNet Options property
  RETURN THIS.roRegExp.IgnoreCase
 ENDPROC &&IgnoreCase_Access
 PROCEDURE IgnoreCase_Assign		&& Assign the IgnoreCase part of the DotNet Options property
  LPARAMETERS;
   tvValue

  THIS.roRegExp.IgnoreCase = m.tvValue
 ENDPROC &&IgnoreCase_Assign

 PROCEDURE IgnorePatternWhitespace_Access		&& Access the IgnorePatternWhitespace part of the DotNet Options property
  RETURN THIS.roRegExp.IgnorePatternWhitespace
 ENDPROC &&IgnorePatternWhitespace_Access
 PROCEDURE IgnorePatternWhitespace_Assign		&& Assign the IgnorePatternWhitespace part of the DotNet Options property
  LPARAMETERS;
   tvValue

  THIS.roRegExp.IgnorePatternWhitespace = m.tvValue
 ENDPROC &&IgnorePatternWhitespace_Assign

 PROCEDURE MultiLine_Access		&& Access the MultiLine part of the DotNet Options property
  RETURN THIS.roRegExp.MultiLine
 ENDPROC &&Multiline_Access
 PROCEDURE MultiLine_Assign		&& Assign the MultiLine part of the DotNet Options property
  LPARAMETERS;
   tvValue

  THIS.roRegExp.MultiLine = m.tvValue
 ENDPROC &&Multiline_Assign

 PROCEDURE RightToLeft_Access		&& Access the RightToLeft part of the DotNet Options property
  RETURN THIS.roRegExp.RIGHTTOLEFT
 ENDPROC &&RightToLeft_Access
 PROCEDURE RightToLeft_Assign		&& Assign the RightToLeft part of the DotNet Options property
  LPARAMETERS;
   tvValue

  THIS.roRegExp.RIGHTTOLEFT = m.tvValue
 ENDPROC &&RightToLeft_Assign

 PROCEDURE Singleline_Access		&& Access the Singleline part of the DotNet Options property
  RETURN THIS.roRegExp.SingleLine
 ENDPROC &&Singleline_Access
 PROCEDURE Singleline_Assign		&& Assign the Singleline part of the DotNet Options property
  LPARAMETERS;
   tvValue

  THIS.roRegExp.SingleLine = m.tvValue
 ENDPROC &&Singleline_Assign
*/ RegExp Options
*/ Assign Access

*VFP specific stuff
 PROCEDURE Escape_Like		&& Interface to the DotNet Escape method, with special operation to keep LIKE function
  LPARAMETERS;
   tvValue

  LOCAL;
   lvReturn AS STRING

  lvReturn = THIS.roRegExp.ESCAPE(m.tvValue)
  lvReturn = STRTRAN(STRTRAN(m.lvReturn,"\*",".*"),"\?",".")

  RETURN m.lvReturn
 ENDPROC &&Escape_Like

 PROCEDURE UnEscape_Like		&& Interface to the DotNet UnEscape method, with special operation to keep LIKE function
  LPARAMETERS;
   tvValue

  LOCAL;
   lvReturn AS STRING

  lvReturn = ""+STRTRAN(STRTRAN(STRTRAN(STRTRAN(m.tvValue,".*","\*"),"\.",0h00),".","\?"),0h00,"\.")
  RETURN THIS.roRegExp.UnEscape(m.lvReturn)
 ENDPROC &&UnEscape_Like

 PROCEDURE Fill_Matches		&& Rebuild the Matches object returned by DotNet to VFP Collections object
  LPARAMETERS;
   toMatches,;
   tlAutoExpandGroup,;
   tlAutoExpandCaptures

  LOCAL;
   lnMatch   AS INTEGER,;
   loMatches AS COLLECTION

  loMatches = CREATEOBJECT("Collection")
  FOR lnMatch = 0 TO m.toMatches.COUNT-1
*   loMatches.ADD(THIS.Fill_Match(toMatches.ITEM(m.lnMatch),m.tlAutoExpandGroup,m.tlAutoExpandCaptures,m.lnMatch=0))
   loMatches.ADD(THIS.Fill_Match(toMatches.ITEM(m.lnMatch),m.tlAutoExpandGroup,m.tlAutoExpandCaptures,.t.))
  ENDFOR &&lnMatch

  RETURN m.loMatches
 ENDPROC &&Fill_Matches

 PROCEDURE Fill_Match		&& Rebuild the Match object returned by DotNet to VFP Empty object
  LPARAMETERS;
   toMatch,;
   tlAutoExpandGroup,;
   tlAutoExpandCaptures,;
   tlAddCaptures

  LOCAL;
   loMatch AS EMPTY

  loMatch = CREATEOBJECT("Empty")
  ADDPROPERTY(m.loMatch,"Value",  THIS.Get_Value  (m.toMatch))
  ADDPROPERTY(m.loMatch,"Index",  THIS.Get_Index  (m.toMatch))
  ADDPROPERTY(m.loMatch,"Length", THIS.Get_Length (m.toMatch))
  ADDPROPERTY(m.loMatch,"Success",THIS.Get_Success(m.toMatch))
  ADDPROPERTY(m.loMatch,"Name",   THIS.Get_Name   (m.toMatch))
  IF m.tlAddCaptures THEN
   IF m.tlAutoExpandCaptures THEN
    ADDPROPERTY(m.loMatch,"Captures",THIS.Fill_Captures(THIS.Get_Captures(m.toMatch),.T.))
   ELSE  &&m.tlAutoExpandCaptures
    ADDPROPERTY(m.loMatch,"Captures",THIS.Get_Captures(m.toMatch))
   ENDIF &&m.tlAutoExpandCaptures
  ENDIF &&m.tlAddCaptures

  IF m.tlAutoExpandGroup THEN
   ADDPROPERTY(m.loMatch,"Groups",THIS.Fill_Groups(THIS.Get_Groups(m.toMatch),m.tlAutoExpandCaptures))
  ELSE  &&m.tlAutoExpandGroup
   ADDPROPERTY(m.loMatch,"Groups",THIS.Get_Groups(m.toMatch))
  ENDIF &&m.tlAutoExpandGroup

  RETURN m.loMatch
 ENDPROC &&Fill_Matches

 PROCEDURE Fill_Groups		&& Rebuild the Groups object returned by DotNet to VFP Collections object
  LPARAMETERS;
   toGroups,;
   tlAutoExpandCaptures

  LOCAL;
   lnGroup  AS INTEGER,;
   loGroups AS COLLECTION

  loGroups = CREATEOBJECT("Collection")
  FOR lnGroup = 0 TO m.toGroups.COUNT-1
   loGroups.ADD(THIS.Fill_Group(toGroups.ITEM(m.lnGroup),m.tlAutoExpandCaptures))
  ENDFOR &&lnGroup

  RETURN m.loGroups
 ENDPROC &&Fill_Groups

 PROCEDURE Fill_Group		&& Rebuild the Group object returned by DotNet to VFP Empty object
  LPARAMETERS;
   toGroup,;
   tlAutoExpandCaptures

  LOCAL;
   loGroup AS EMPTY

  loGroup = CREATEOBJECT("Empty")
  ADDPROPERTY(m.loGroup,"Value",  THIS.Get_Value  (m.toGroup))
  ADDPROPERTY(m.loGroup,"Index",  THIS.Get_Index  (m.toGroup))
  ADDPROPERTY(m.loGroup,"Length", THIS.Get_Length (m.toGroup))
  ADDPROPERTY(m.loGroup,"Name",   THIS.Get_Name   (m.toGroup))
  ADDPROPERTY(m.loGroup,"Success",THIS.Get_Success(m.toGroup))

  IF m.tlAutoExpandCaptures THEN
   ADDPROPERTY(m.loGroup,"Captures",THIS.Fill_Captures(THIS.Get_Captures(m.toGroup)))
  ELSE  &&m.tlAutoExpandCaptures
   ADDPROPERTY(m.loGroup,"Captures",THIS.Get_Captures(m.toGroup))
  ENDIF &&m.tlAutoExpandCaptures

  RETURN m.loGroup
 ENDPROC &&Fill_Groups

 PROCEDURE Fill_Captures		&& Rebuild the Captures object returned by DotNet to VFP Collections object
  LPARAMETERS;
   toCaptures,;
   tlWithName

  LOCAL;
   lnCapture  AS INTEGER,;
   loCaptures AS COLLECTION

  loCaptures = CREATEOBJECT("Collection")
  FOR lnCapture = 0 TO m.toCaptures.COUNT-1
   loCaptures.ADD(THIS.Fill_Capture(toCaptures.ITEM(m.lnCapture),m.tlWithName))
  ENDFOR &&lnCapture

  RETURN m.loCaptures
 ENDPROC &&Fill_Captures

 PROCEDURE Fill_Capture		&& Rebuild the Capture object returned by DotNet to VFP Empty object
  LPARAMETERS;
   toCapture,;
   tlWithName

  LOCAL;
   loCapture AS EMPTY

  loCapture = CREATEOBJECT("Empty")
  ADDPROPERTY(m.loCapture,"Value", THIS.Get_Value (m.toCapture))
  ADDPROPERTY(m.loCapture,"Index", THIS.Get_Index (m.toCapture))
  ADDPROPERTY(m.loCapture,"Length",THIS.Get_Length(m.toCapture))
  IF m.tlWithName THEN
   ADDPROPERTY(m.loCapture,"Name",   THIS.Get_Name   (m.toCapture))
   ADDPROPERTY(m.loCapture,"Success",THIS.Get_Success(m.toCapture))
  ENDIF &&m.tlWithName

  RETURN m.loCapture
 ENDPROC &&Fill_Capture

*/VFP specific stuff

*Wrapper,simple calls
 PROCEDURE ESCAPE		&& Interface to the DotNet Escape method
  LPARAMETERS;
   tvValue

  RETURN THIS.roRegExp.ESCAPE(m.tvValue)
 ENDPROC &&Escape

 PROCEDURE UnEscape		&& Interface to the DotNet UnEscape method
  LPARAMETERS;
   tvValue

  RETURN THIS.roRegExp.UnEscape(m.tvValue)
 ENDPROC &&UnEscape

 PROCEDURE GetGroupNames		&& Interface to the DotNet GetGroupNames method
  RETURN THIS.roRegExp.GetGroupNames()
 ENDPROC &&m.tvValue

 PROCEDURE GetGroupNumbers		&& Interface to the DotNet GetGroupNumbers method
  RETURN THIS.roRegExp.GetGroupNumbers()
 ENDPROC &&m.tvValue

 PROCEDURE GroupNameFromNumber		&& Interface to the DotNet GroupNameFromNumber method
  LPARAMETERS;
   tvValue

  RETURN THIS.roRegExp.GroupNameFromNumber(m.tvValue)
 ENDPROC &&GroupNameFromNumber


 PROCEDURE GroupNumberFromName		&& Interface to the DotNet GroupNumberFromName method
  LPARAMETERS;
   tvValue

  RETURN THIS.roRegExp.GroupNumberFromName(m.tvValue)
 ENDPROC &&GroupNumberFromName
*/Wrapper,simple calls

*Wrapper,RegExp action
 PROCEDURE IsMatch		&& Interface to the DotNet IsMatch method
  LPARAMETERS;
   tcInput,;
   tvValue1
  LOCAL;
   loBridge AS OBJECT,;
   lvReturn AS VARIANT

  loBridge = GetwwDotnetBridge()

  IF VARTYPE(m.tvValue1)='N' THEN
*StartAt like VFP,starts with 1
   tvValue1 = m.tvValue1-THIS.rnOffset
  ENDIF &&VARTYPE(m.tvValue1)='N'

  DO CASE
   CASE PCOUNT()=1
    lvReturn = loBridge.InvokeMethod(THIS.roRegExp,"IsMatch",m.tcInput)
   CASE PCOUNT()=2
    lvReturn = loBridge.InvokeMethod(THIS.roRegExp,"IsMatch",m.tcInput,m.tvValue1)
   OTHERWISE
    lvReturn = .NULL.
  ENDCASE

  RETURN m.lvReturn
 ENDPROC &&IsMatch

 PROCEDURE Match		&& Interface to the DotNet Match method
  LPARAMETERS;
   tcInput,;
   tvValue1,;
   tvValue2
  LOCAL;
   loBridge AS OBJECT,;
   lvReturn AS VARIANT

  loBridge = GetwwDotnetBridge()

  IF VARTYPE(m.tvValue1)='N' THEN
*StartAt like VFP,starts with 1
   tvValue1 = m.tvValue1-THIS.rnOffset
  ENDIF &&VARTYPE(m.tvValue1)='N'

  DO CASE
   CASE PCOUNT()=1
    lvReturn = loBridge.InvokeMethod(THIS.roRegExp,"Match",m.tcInput)
   CASE PCOUNT()=2
    lvReturn = loBridge.InvokeMethod(THIS.roRegExp,"Match",m.tcInput,m.tvValue1)
   CASE PCOUNT()=3
    lvReturn = loBridge.InvokeMethod(THIS.roRegExp,"Match",m.tcInput,m.tvValue1,m.tvValue2)
   OTHERWISE
    lvReturn = .NULL.
  ENDCASE

  IF THIS.ReturnFoxObjects THEN
   lvReturn = THIS.Fill_Match(m.lvReturn,THIS.AutoExpandGroup,THIS.AutoExpandCaptures,.T.)
  ENDIF &&THIS.ReturnFoxObjects

  RETURN m.lvReturn
 ENDPROC &&Match

 PROCEDURE Matches		&& Interface to the DotNet Matches method
  LPARAMETERS;
   tcInput,;
   tvValue1
  LOCAL;
   loBridge AS OBJECT,;
   lvReturn AS VARIANT

  loBridge = GetwwDotnetBridge()

  IF VARTYPE(m.tvValue1)='N' THEN
*StartAt like VFP,starts with 1
   tvValue1 = m.tvValue1-THIS.rnOffset
  ENDIF &&VARTYPE(m.tvValue1)='N'

  DO CASE
   CASE PCOUNT()=1
    lvReturn = loBridge.InvokeMethod(THIS.roRegExp,"Matches",m.tcInput)
   CASE PCOUNT()=2
    lvReturn = loBridge.InvokeMethod(THIS.roRegExp,"Matches",m.tcInput,m.tvValue1)
   OTHERWISE
    lvReturn = .NULL.
  ENDCASE


  IF THIS.ReturnFoxObjects THEN
   lvReturn = THIS.Fill_Matches(m.lvReturn,THIS.AutoExpandGroup,THIS.AutoExpandCaptures)
  ENDIF &&THIS.ReturnFoxObjects

  RETURN m.lvReturn
 ENDPROC &&Matches

 PROCEDURE REPLACE		&& Interface to the DotNet Replace method
  LPARAMETERS;
   tcInput,;
   tvValue1,;
   tvValue2,;
   tvValue3
  LOCAL;
   loBridge AS OBJECT,;
   lvReturn AS VARIANT

  loBridge = GetwwDotnetBridge()

  IF VARTYPE(m.tvValue3)='N' THEN
*StartAt like VFP, starts with 1
   tvValue3 = m.tvValue3-THIS.rnOffset
  ENDIF &&VARTYPE(m.tvValue3)='N'

  DO CASE
   CASE PCOUNT()=1
    lvReturn = loBridge.InvokeMethod(THIS.roRegExp,"Replace",m.tcInput)
   CASE PCOUNT()=2
    lvReturn = loBridge.InvokeMethod(THIS.roRegExp,"Replace",m.tcInput,m.tvValue1)
   CASE PCOUNT()=3
    lvReturn = loBridge.InvokeMethod(THIS.roRegExp,"Replace",m.tcInput,m.tvValue1,m.tvValue2)
   CASE PCOUNT()=4
    lvReturn = loBridge.InvokeMethod(THIS.roRegExp,"Replace",m.tcInput,m.tvValue1,m.tvValue2,m.tvValue3)
   OTHERWISE
    lvReturn = .NULL.
  ENDCASE

  RETURN m.lvReturn
 ENDPROC &&Replace

 PROCEDURE SPLIT		&& Interface to the DotNet Split method
  LPARAMETERS;
   tcInput,;
   tvValue1,;
   tvValue2
  LOCAL;
   loBridge AS OBJECT,;
   lvReturn AS VARIANT

  loBridge = GetwwDotnetBridge()

  IF VARTYPE(m.tvValue1)='N' THEN
*StartAt like VFP, starts with 1
   tvValue1 = m.tvValue1-THIS.rnOffset
  ENDIF &&VARTYPE(m.tvValue1)='N'

  DO CASE
   CASE PCOUNT()=1
    lvReturn = loBridge.InvokeMethod(THIS.roRegExp,"Split",m.tcInput)
   CASE PCOUNT()=2
    lvReturn = loBridge.InvokeMethod(THIS.roRegExp,"Split",m.tcInput,m.tvValue1)
   CASE PCOUNT()=3
    lvReturn = loBridge.InvokeMethod(THIS.roRegExp,"Split",m.tcInput,m.tvValue1,m.tvValue2)
   OTHERWISE
    lvReturn = .NULL.
  ENDCASE

  RETURN m.lvReturn
 ENDPROC &&Split
*/Wrapper,RegExp action

*Wrapper, subobjects
 PROCEDURE Get_Value		&& Get the Value property from DotNet Match, Group, Capture objects
  LPARAMETERS;
   loMatch

  LOCAL;
   loBridge AS OBJECT,;
   lvReturn AS VARIANT

  loBridge = GetwwDotnetBridge()
  lvReturn = loBridge.GetProperty(m.loMatch,"Value")

  RETURN m.lvReturn

 ENDPROC &&Get_Value

 PROCEDURE Get_Index		&& Get the Index property from DotNet Match, Group, Capture objects
  LPARAMETERS;
   loMatch

  LOCAL;
   loBridge AS OBJECT,;
   lvReturn AS VARIANT

  loBridge = GetwwDotnetBridge()
  lvReturn = loBridge.GetProperty(m.loMatch,"Index")

  RETURN m.lvReturn+THIS.rnOffset
 ENDPROC &&Get_Index

 PROCEDURE Get_Length		&& Get the Length property from DotNet Match, Group, Capture objects
  LPARAMETERS;
   loMatch

  LOCAL;
   loBridge AS OBJECT,;
   lvReturn AS VARIANT

  loBridge = GetwwDotnetBridge()
  lvReturn = loBridge.GetProperty(m.loMatch,"Length")

  RETURN m.lvReturn
 ENDPROC &&Get_Length

 PROCEDURE Get_Name		&& Get the Name property from DotNet Match, Group, Capture objects
  LPARAMETERS;
   loMatch

  LOCAL;
   loBridge AS OBJECT,;
   lvReturn AS VARIANT

  loBridge = GetwwDotnetBridge()
  lvReturn = loBridge.GetProperty(m.loMatch,"Name")
  RETURN m.lvReturn
 ENDPROC &&Get_Name

 PROCEDURE Get_Groups		&& Get the Groups property from DotNet Match objects
  LPARAMETERS;
   loMatch

  LOCAL;
   loBridge AS OBJECT,;
   lvReturn AS VARIANT

  loBridge = GetwwDotnetBridge()
  lvReturn = loBridge.GetProperty(m.loMatch,"Groups")
  RETURN m.lvReturn
 ENDPROC &&Get_Groups

 PROCEDURE Get_Captures		&& Get the Captures property from DotNet Match, Group objects
  LPARAMETERS;
   loMatch

  LOCAL;
   loBridge AS OBJECT,;
   lvReturn AS VARIANT

  loBridge = GetwwDotnetBridge()
  lvReturn = loBridge.GetProperty(m.loMatch,"Captures")
  RETURN m.lvReturn
 ENDPROC &&Get_Captures

 PROCEDURE Get_Success		&& Get the Success property from DotNet Match, Capture objects
  LPARAMETERS;
   loMatch

  LOCAL;
   loBridge AS OBJECT,;
   lvReturn AS VARIANT

  loBridge = GetwwDotnetBridge()
  lvReturn = loBridge.GetProperty(m.loMatch,"Success")
  RETURN m.lvReturn
 ENDPROC &&Get_Success

*starting with .Net 6
*!*		Procedure Get_ValueSpan		&& Get the Get_ValueSpan property from DotNet Match, Group, Capture objects
*!*			Lparameters;
*!*				loMatch

*!*			Local;
*!*				loBridge As Object,;
*!*				lvReturn As Variant

*!*			loBridge = GetwwDotnetBridge()
*!*			lvReturn = loBridge.GetProperty(m.loMatch,"ValueSpan")
*!*			Return m.lvReturn
*!*		Endproc &&Get_ValueSpan

*or we need additional objects to wrap this
*/Wrapper, subobjects

*Helpers to show the contents of a Match or Matches
 PROCEDURE Show_Unwind		&& Expand to a Matches or Match object for to show all data in the object
  LPARAMETERS;
   toMatches
 
  LOCAL;
   lcReturn  AS STRING,;
   llMatches AS INTEGER
 
 *  SET STEP ON
 
  llMatches = .T.
  TRY
    =m.toMatches.COUNT
   CATCH
    llMatches = .F.
  ENDTRY
 
  IF m.llMatches THEN
 *Matches
   lcReturn = "Matches "
   IF PEMSTATUS(m.toMatches,"tag",5) THEN
 *VFP object
    lcReturn =  "Expanding VFP " + m.lcReturn + 0h0D0A + THIS.Unwind_Matches_V(m.toMatches)
   ELSE  &&PEMSTATUS(m.toMatches,"tag",5)
 *DotNet object
    lcReturn =  "Expanding DotNet " + m.lcReturn + "Matches" + 0h0D0A + THIS.Unwind_Matches(m.toMatches)
   ENDIF &&PEMSTATUS(m.toMatches,"tag",5)   ELSE  &&m.llMatches
 
  ELSE  &&m.llMatches
 *Matches 
   lcReturn = "Match "
   IF PEMSTATUS(m.toMatches,"value",5) THEN
 *VFP object
    lcReturn =  "Expanding VFP " + m.lcReturn + 0h0D0A + THIS.Unwind_Match_V(m.toMatches)
   ELSE  &&PEMSTATUS(m.toMatches,"value",5)
 *DotNet object
    lcReturn =  "Expanding DotNet " + m.lcReturn + 0h0D0A + THIS.Unwind_Match(m.toMatches)
   ENDIF &&PEMSTATUS(m.toMatches,"value",5)
  ENDIF &&m.llMatches
 
  RETURN m.lcReturn
 
 ENDPROC &&Show_Unwind
   
 PROTECTED PROCEDURE Unwind_Matches   && Show all data for a matches object, DotNet object
  LPARAMETERS;
   toMatches

  LOCAL;
   lcReturn   AS STRING,;
   lnMatch    AS NUMBER,;
   loMatch    AS OBJECT

  lcReturn = "Matches:  " + TRIM(PADR(m.toMatches.COUNT,11)) + 0h0D0A
  FOR lnMatch = 0 TO m.toMatches.COUNT-1
   loMatch = toMatches.ITEM(m.lnMatch)

   lcReturn = m.lcReturn + " Match  " + TRIM(PADR(m.lnMatch,11)) + ";"
   lcReturn = m.lcReturn + THIS.Unwind_Match(m.loMatch) + 0h0D0A
  ENDFOR &&lnMatch

  RETURN m.lcReturn
 ENDPROC &&Unwind_Matches

 PROTECTED PROCEDURE Unwind_Matches_V   && Show all data for a matches object, VFP object
  LPARAMETERS;
   toMatches

  LOCAL;
   lcReturn   AS STRING,;
   lnMatch    AS NUMBER,;
   loMatch    AS OBJECT

   lcReturn = "Matches:  " + TRIM(PADR(m.toMatches.COUNT,11)) + 0h0D0A
  FOR lnMatch = 1 TO m.toMatches.COUNT
   loMatch = toMatches.ITEM(m.lnMatch)

   lcReturn = m.lcReturn + " Match  " + TRIM(PADR(m.lnMatch,11)) + ";"
   lcReturn = m.lcReturn + THIS.Unwind_Match_V(m.loMatch) + 0h0D0A
  ENDFOR &&lnMatch

  RETURN m.lcReturn
 ENDPROC &&Unwind_Matches_V

 PROTECTED PROCEDURE Unwind_Match   && Show all data for a match object, DotNet object
  LPARAMETERS;
   toMatch

  LOCAL;
   lcReturn   AS STRING,;
   lnCapture  AS NUMBER,;
   lnGroup    AS NUMBER,;
   loCapture  AS OBJECT,;
   loCaptures AS OBJECT,;
   loGroup    AS OBJECT,;
   loGroups   AS OBJECT

  lcReturn = ""
  SET TEXTMERGE ON NOSHOW
  SET TEXTMERGE TO MEMVAR m.lcReturn
   \\ value:  "<<THIS.Get_Value(m.toMatch)>>"
   \\,at: <<THIS.Get_Index(m.toMatch)>>
   \\,len: <<THIS.Get_Length(m.toMatch)>>
   \\, success  <<THIS.Get_Success(m.toMatch)>>
   \\, name: <<THIS.Get_Name(m.toMatch)>>

  loCaptures = THIS.Get_Captures(m.toMatch)
   \  MCaptures:  <<m.loCaptures.COUNT>>
  FOR lnCapture = 0 TO m.loCaptures.COUNT-1
   loCapture = loCaptures.ITEM(m.lnCapture)
    \   MCapture  <<m.lnCapture>>; value: "<<THIS.Get_Value(m.loCapture)>>"
    \\,at: <<THIS.Get_Index(m.loCapture)>>
    \\,len: <<THIS.Get_Length(m.loCapture)>>
    \\, success  <<THIS.Get_Success(m.loCapture)>>
    \\, Name  <<THIS.Get_Name(m.loCapture)>>
  ENDFOR &&lnCapture

  loGroups = THIS.Get_Groups(m.toMatch)
   \  Groups:  <<m.loGroups.COUNT>>
  FOR lnGroup = 0 TO m.loGroups.COUNT-1
   loGroup = loGroups.ITEM(m.lnGroup)
    \   Group  <<m.lnGroup>>; value: "<<THIS.Get_Value(m.loGroup)>>"
    \\,at: <<THIS.Get_Index(m.loGroup)>>
    \\,len: <<THIS.Get_Length(m.loGroup)>>
    \\, success  <<THIS.Get_Success(m.loGroup)>>
    \\, name  <<THIS.Get_Name(m.loGroup)>>
   loCaptures = THIS.Get_Captures(m.loGroup)
    \    Captures:  <<m.loCaptures.COUNT>>
   FOR lnCapture = 0 TO m.loCaptures.COUNT-1
    loCapture = loCaptures.ITEM(m.lnCapture)
     \     Capture  <<m.lnCapture>>; value: "<<THIS.Get_Value(m.loCapture)>>"
     \\,at: <<THIS.Get_Index(m.loCapture)>>
     \\,len: <<THIS.Get_Length(m.loCapture)>>
    IF m.lnGroup =0 THEN
      \\, success  <<THIS.Get_Success(m.loCapture)>>
      \\, name  <<THIS.Get_Name(m.loCapture)>>
    ENDIF &&m.lnGroup =0
   ENDFOR &&lnCapture
  ENDFOR &&lnGroup
  SET TEXTMERGE TO
  SET TEXTMERGE OFF

  RETURN m.lcReturn

 ENDPROC &&Unwind_Match

 PROTECTED PROCEDURE Unwind_Match_V   && Show all data for a match object, VFP object
  LPARAMETERS;
   toMatch
 
  LOCAL;
   lcReturn   AS STRING,;
   lnCapture  AS NUMBER,;
   lnGroup    AS NUMBER,;
   loCapture  AS OBJECT,;
   loCaptures AS OBJECT,;
   loGroup    AS OBJECT,;
   loGroups   AS OBJECT
  LOCAL llWitchCaptures,llWitchGroups,llWitchMatchCaptures
 
 
 *SET STEP ON 
  llWitchGroups        = .T.
  llWitchCaptures      = .T.
  llWitchMatchCaptures = .T.
  TRY
    = m.toMatch.Groups.TAG
   CATCH
    llWitchGroups = .F.
  ENDTRY
 
  TRY
    = m.toMatch.Captures.TAG
   CATCH
    llWitchMatchCaptures = .F.
    IF m.llWitchGroups THEN
     TRY
       = toMatch.Groups.ITEM(1).Captures.TAG
      CATCH
       llWitchCaptures = .F.
     ENDTRY
    ELSE  &&m.llWitchGroups
     llWitchCaptures = .F.
    ENDIF &&m.llWitchGroups
  ENDTRY
 
  lcReturn = ""
  SET TEXTMERGE ON NOSHOW
  SET TEXTMERGE TO MEMVAR m.lcReturn
     \\ value:  "<<m.toMatch.Value>>"
     \\,at: <<m.toMatch.Index>>
     \\,len: <<m.toMatch.Length>>
     \\, success  <<m.toMatch.Success>>
     \\, name: <<m.toMatch.Name>>
 
  IF m.llWitchMatchCaptures THEN
   loCaptures = m.toMatch.Captures
     \  MCaptures:  <<m.loCaptures.COUNT>>
   FOR lnCapture = 1TO m.loCaptures.COUNT
    loCapture = loCaptures.ITEM(m.lnCapture)
      \   MCapture  <<m.lnCapture>>; value: "<<m.loCapture.Value>>"
      \\,at: <<m.loCapture.Index>>
      \\,len: <<m.loCapture.Length>>
      \\, success  <<m.loCapture.Success>>
      \\, Name  <<m.loCapture.Name>>
   ENDFOR &&lnCapture
 
  ELSE  &&m.llWitchMatchCaptures
     \  MCaptures:  not expanded
  ENDIF &&m.llWitchMatchCaptures
 
  IF m.llWitchGroups THEN
   loGroups = m.toMatch.Groups
     \  Groups:  <<m.loGroups.COUNT>>
   FOR lnGroup = 1TO m.loGroups.COUNT
    loGroup = loGroups.ITEM(m.lnGroup)
      \   Group  <<m.lnGroup>>; value: "<<m.loGroup.Value>>"
      \\,at: <<m.loGroup.Index>>
      \\,len: <<m.loGroup.Length>>
      \\, success  <<m.loGroup.Success>>
      \\, name  <<m.loGroup.Name>>
 
    IF m.llWitchCaptures THEN
     loCaptures = m.loGroup.Captures
      \    Captures:  <<m.loCaptures.COUNT>>
     FOR lnCapture = 1 TO m.loCaptures.COUNT
      loCapture = loCaptures.ITEM(m.lnCapture)
       \     Capture  <<m.lnCapture>>; value: "<<m.loCapture.Value>>"
       \\,at: <<m.loCapture.Index>>
       \\,len: <<m.loCapture.Length>>
      IF m.lnGroup =0 THEN
        \\, success  <<.loCapture.Success>>
        \\, name  <<m.loCapture.Name>>
      ENDIF &&m.lnGroup =0
     ENDFOR &&lnCapture
 
    ELSE  &&m.llWitchCaptures
      \    Captures:  not expanded
    ENDIF &&m.llWitchCaptures
   ENDFOR &&lnGroup
 
  ELSE  &&m.llWitchGroups
    \  Groups:  not expanded
  ENDIF &&m.llWitchGroups
 
  SET TEXTMERGE TO
  SET TEXTMERGE OFF
 
  RETURN m.lcReturn
 
 ENDPROC &&Unwind_Match_V
*/Helpers to show the contents of a Match or Matches
 
ENDDEFINE &&SF_RegExp As Session

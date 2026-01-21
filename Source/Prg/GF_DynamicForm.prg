*=======================================================================================================
* Dynamic Form - 1.9.1
*---------------------------------------------------------------------------------------
* By: Matt Slay 
*-------------------------------------------------------------------------------------------------------
*--
*-- Web Site: https://github.com/mattslay/DynamicForms
*--
*-- 
*-- You can automatically download this Component and its updates through "Thor - Check for Updates".
*--        To learn more about Thor and its Tools and Updaters for FoxPro, see: 
*--        https://github.com/VFPX/Thor/blob/master/Docs/Thor_one-click_update.md
*--
*-- User Discussions Group: https://groups.google.com/forum/#!forum/foxprodynamicforms
*--
*--	 after JOINING the Group above, you can post to the web forum or 
*--    email to: foxprodynamicforms@googlegroups.com 
*--
*--======================================================================================= 
*--
*-- This source code PRG contains the class definitions needed to create Dynamic Forms in your apps,
*-- and is also a ready-to-run sample to show you a rendered form sample.
*--
*-- Just run this PRG to see a Dynamic Form sample from the code below. The sample creates a Modal form
*-- which is bound to a few Private variables and an simple data object (the data object is created in the code
*-- at run time simply to mock a real data object and to prevent the demo from having to ship with a sample dbf.)
*-- You can also easily bind to local cursors via the cAlias property, rather than data objects. Please
*-- see the documentation link below for more details on using Dynamic Forms with cursors, and advanced
*-- uses like creating Modeless forms or working with Business Objects to save the data when binding to 
*-- an oDataObject.
*--
*--======================================================================================= 
*-- DOCUMENTAION - Visit this link to see FULL DOCUMENTAION pages:
*--
*--   https://github.com/mattslay/DynamicForms
*--
*-----------------------------------------------------------------------------------------
*-- VIDEOS - 
*--
*--   Video #1 ? Introduction and Demos (8:15) View here:  http://bit.ly/DynamicForms-Video-1
*--
*--   Video #2 ? Exploring the class PRG and code sample (9:09) View here: http://bit.ly/Dynamic-Forms-Video-2
*--
*----------------------------------------------------------------------------------------- 
*--
*-- Version History
*--
*-- 1.9.1 2017-08-30 Fixed spelling, capitalization, formatting (no functional changes).
*-- 1.9.0 Production release - August 27, 2017 - 20170827 (Migrated from VFPx/CodePlex to GitHub)
*-- 1.8.2 Beta - September 29, 2014 - 2014-09-29
*-- 1.7.0 Alpha - October 31, 2013 - 2013-10-31
*-- 2012-10-18 Alpha 1.5.0 released.
*-- 2012-10-26 Alpha 1.4.1 released.
*-- 2012-10-23 Alpha 1.4.0 released.
*-- 2012-10-08 Alpha 1.3.0 released.
*-- 2012-09-25 Alpha 1.2.0 released.
*-- 2012-09-18 Alpha 1.0.0 released.
*-- 2012-09-04 Public alpha release 0.9.0 released on VFPx.
*-- 2012-08-31 Dynamic Forms accepted as an official VFPx Project.
*-- 2012-08-24 Project proposal submitted to VFPx admins.
*-- 2012-05-08 Initial concept class created and the first ever Dynamic Form was generated from a few lines of code.
*--
*---------------------------------------------------------------------------------------

*-- Example usage:

Private lnPrice, laOptions[1], lnOption
Private loForm as 'DynamicForm'
Local loObject, lcBodyMarkup

*---- Step 1: Prepare the data...
*--   As noted in the Documentation you can bind the form to table aliases and cursors, private or public variables, or Data Objects.
*--   See https://github.com/mattslay/DynamicForms#binding
*--
*--   In this example, we'll build a Data Object in code so that we do not have to distribute a sample dbf with this project package.
*--   Often, "Data Objects" come from table rows, via the Scatter command, or other object-building techniques. You then pass that
*--   object into this Dynamic Form per this code sample. Watch the video series to see examples of working with cursors.
*--
	loObject = CreateObject('Empty')
	AddProperty(loObject , 'id', '12345')
	AddProperty(loObject , 'first_name', 'Joe')
	AddProperty(loObject , 'mid_init', 'N.')
	AddProperty(loObject , 'last_name', 'Coderman')
	AddProperty(loObject , 'ad_type', 'Banner')
	AddProperty(loObject , 'notes', 'This man came here and wrote some codez.')
	AddProperty(loObject , 'still_here', .f.)
	AddProperty(loObject , 'has_laptop', .t.)
	AddProperty(loObject , 'bool_1', .t.)
	AddProperty(loObject , 'bool_2', .t.)
	AddProperty(loObject , 'bool_3', .t.)
	AddProperty(loObject , 'weight', 185)

*-- Step 2: Define the UI layout string (similar to HTML / XAML markup) --------------------------------------------
*-- See https://github.com/mattslay/DynamicForms#markup_syntax

Text to lcBodyMarkup NoShow

 	.lLabelsAbove = .t. |

	id			.enabled = .f.	
				.fontbold = .t.
				.label.FontBold = .t. |
	
	ad_type 	.class = 'combobox'
				.RowSource = 'laOptions'
				.RowSourceType = 5 
				.row-increment = 0 |

	bool_3		.caption = 'You can specify BOLD captions.'
				.FontBold = .t. 
				.width = 400 |

	
	first_name	.set-focus = .t. |
	mid_init	.row-increment = 0 |
	last_name	.row-increment = 0 |

	notes		.class = 'editbox'
				.width = 400
				.height = 80
				.anchor = 10|

	lnPrice		.label.caption = 'List Price'
				.label.alignment = 1 |
				
	weight		.row-increment  = 0
				.label.alignment = 1 |
	
	lnOption	.class = 'optiongroup'
				.caption = 'Color options.'
				.buttoncount = 2
				.width = 200
				.height = 60
				.option1.caption = 'Red with orange stripes'
				.option1.autosize = .t.
				.option2.caption = 'Purple with black dots'
				.option2.autosize = .t. |

	.class = 'label' 	.caption = 'Thank you for trying DynamicForm.'
						.autosize = .t.
						.render-if = (Day(Date()) > 1)

EndText

*-- Example of a Private variables that can be bound to also
lnPrice = 107.15
lnOption = 2

*-- Array of options/values to display in the ComboBox defined in cMarkup above (note that it's declared Private above ---
Dimension laOptions[3]
laOptions[1] = 'Banner'
laOptions[2] = 'Placard'
laOptions[3] = 'Name Tag'

*-- Step 3. Create an instance of DynamicForm class
*-- See https://github.com/mattslay/DynamicForms#step3
loForm = CreateObject('DynamicForm')

*-- Step 4. Set a few properties on loForm to wire everything up and set rendering options...
*-- See https://github.com/mattslay/DynamicForms#step4
loForm.oDataObject = loObject && Set the data object that the form fields bind to

loForm.Caption = 'Dynamic Forms'
loForm.oRenderEngine.lLabelsAbove = .t. && Generate field labels above each control. Default is .F., meaning "inline with controls", to the left.

*-- Setup Header area (or disable it)-------------
loForm.cHeading = 'Sample Form'
loForm.nHeadingFontSize = 14 && You can set heading label font size as desired

*loForm.cHeaderMarkup = ''	&& Set to empty to disable automatic Header markup. Or, 
							&& assign custom markup string to customize the header area. See http://vfpx.codeplex.com/wikipage?title=Dynamic%20Form%20Main%20Form%20Layout

*-- Set the main body area markup ---------
loForm.cBodyMarkup = lcBodyMarkup

* loForm.cFooterMarkup = ''	&& Set to empty to disable use of automatic Footer markup. Or, 
							&& assign custom markup string to customize the footer area. See http://vfpx.codeplex.com/wikipage?title=Dynamic%20Form%20Main%20Form%20Layout

*-- Step 5. Call Render method to create the controls in the Form
*--			See https://github.com/mattslay/DynamicForms#step5
llResult = loForm.Render()

loForm.MinHeight = loForm.cntMain.Height
loForm.MinWidth = loForm.cntMain.Width

*-- Step 6. Show the form to the user
*-- See https://github.com/mattslay/DynamicForms#step6
If llResult = .t.
	*-- Note. You have a chance here to programmatically change anything on the form or controls
	*-- in any way needed before showing the form to the user...
	loForm.Show()
	*loForm.Show(1, '300,10')	&& 0 = Modeless, 1 = Modal. See http://vfpx.codeplex.com/wikipage?title=Dynamic%20Forms%20Properties#show
								&& You can also pass a 'left,top' pair to position the form at a fixed point.
		Else
	MessageBox(loForm.oRenderEngine.GetErrorsAsString() , 0, 'Notice.')
	*-- If there were any rendering errors (llResult = .f.), then you can read loForm.oRenderEngine.nErrorCount property
	*-- and loForm.oRenderEngine.oErrors collection for a detail of each error. Or call loForm.oRenderEngine.GetErrorsAsString().
	loForm.Show(1)
Endif

*-- At this point, the user is interacting with the form, and it will eventually be closed when they click
*-- Save, Cancel, or the [X] button. At that time, flow will return here, and we can then read any property
*-- on loForm and  loForm.oRenderEngine, and even access the rendered controls.

*-- Step 7. Proceed with program flow based on whether user clicked Save or Cancel/closed the form.
*-- See https://github.com/mattslay/DynamicForms#step7
If Vartype(loForm) = 'O' and Lower(loForm.cReturn) = 'save'
	*-- If Save is clicked, the controlsources are already updated with the new values from the UI.
	*-- Do whatever local processing you need following the Save click by the user...
	Release loForm
Else
	*-- Do whatever processing for Close/Cancel user action...
	*-- If using the Button Bar or and instance of DF_CancelButton on the form and Cancel was clicked,
	*-- and the property loForm.lRestoreDataOnCancel = .t. (default), then the controlsources will already
	*-- be restored to their original value by the Form class.
EndIf

*-- After the preceding Save/Cancel processing, we can now Release the loForm object.




*-- Class definitions follow:


*======================================================================================= 
#DEFINE CR Chr(13)
#DEFINE CRCR Chr(13)

Define Class DynamicForm as Form

  	cVersion = '1.9.1'
  	cVersionFull = '1.9.1 Production Release - August 30, 2017'
  	Caption = ''
	*-- Binding form fields to a cursor/alias...
	cAlias = ''                 && The cursor/alias that your form fields bind to. Make sure this alias is opened and positioned to the correct record.

	*-- Binding form fields to a DataObject for data, and optionally setting an oBusinessObject for Saving the data.
	oDataObject = .null.        && The object which has the data properties that you form fields bind to.
	oBusinessObject = .null.	&& oBusinessObject often has a Save() method to save the values from oDataObject back to its table.
	cBusinessObjectSaveMethod = 'Save()' && This method on the oBusinessObject will be called when the Save button is clicked.
	cDataObjectRef = 'Thisform.oDataObject'

	lRestoreDataOnCancel = .t.	&& When .T., changes to oDataObject or cAlias will be restored to their original values if form is Cancelled by user.
	lClearEventsOnClose = .f.	&& Only use if you want to call Clear Events when this form is closed.
	oRenderEngine = .null. 		&& Will be populated in Thisform.Init() event. You can override with your own instance of a Render Engine after this form in Initlialized.
	
	cHeading = .null.			&& This text is displayed in the default Header area of the form.
	cSaveButtonCaption = .null.
	cCancelButtonCaption = .null.
	nHeadingFontSize = 14		&& The font size for the label in the Header area.

	cHeaderMarkup = .null.
	cBodyMarkup = .null.
	cFooterMarkup = .null.
	cPopupFormBodyMarkup = .null.
	
	MinWidth = 400
	MinHeight = 250
	
	*-- Consider these as ReadOnly after the form has been Hidden by one of the form Buttons --------
	lSaveClicked = .f.
	lCancelClicked 	= .f.
	cReturn = ''
	cHandle = '' && A unique string used to store a a reference to this form on _screen. This is used in handling Modeless forms.
	
	Width = 10000 && This high values will be set to actual rendered size in the Show() method
	Height = 10000 && This high values will be set to actual rendered size in the Show() method
	DataSession = 1
	
	Add Object cntMain as Container With ;
		Top = 0, ;
		Left = 0, ;
		Width = 1, ; 	&& Will be resized by RenderEngine to size required to hold all controls
		Height = 1, ;
		BorderWidth = 0, ;
		Anchor = 15, ;
		margin_left = 0, ;
		margin_right = 0, ;
		margin_top = 0, ;
		margin_bottom = 0
	
	*---------------------------------------------------------------------------------------
	Procedure cSaveButtonCaption_Assign()tcCaption

		This.oRenderEngine.cSaveButtonCaption = tcCaption

	Endproc
	
	*---------------------------------------------------------------------------------------
	Procedure cCancelButtonCaption_Assign(tcCaption)

		This.oRenderEngine.cCancelButtonCaption = tcCaption

	Endproc
	
	*--------------------------------------------------------------------------------------- 
	Procedure cHeading_assign(tcCaption)

		This.oRenderEngine.cHeading = tcCaption

	Endproc
	
	*--------------------------------------------------------------------------------------- 
	Procedure nHeadingFontSize_assign(tnFontSize)

		This.oRenderEngine.nHeadingFontSize = tnFontSize

	Endproc
	
	*--------------------------------------------------------------------------------------- 
	Procedure cHeaderMarkup_assign(tcMarkup)

		This.oRenderEngine.cHeaderMarkup = tcMarkup

	EndProc
	
	*--------------------------------------------------------------------------------------- 
	Procedure cBodyMarkup_assign(tcMarkup)
		
		This.cBodyMarkup = tcMarkup
		This.oRenderEngine.cBodyMarkup = tcMarkup

	Endproc
	
	*--------------------------------------------------------------------------------------- 
	Procedure cFooterMarkup_assign(tcMarkup)

		This.oRenderEngine.cFooterMarkup = tcMarkup

	EndProc

	*--------------------------------------------------------------------------------------- 
	Procedure Init()

		This.cHandle = 'DF_' + Sys(2015) && Used to keep a ref to modeless forms alive
		This.oRenderEngine = CreateObject('DynamicFormRenderEngine')

	Endproc
	
	*--------------------------------------------------------------------------------------- 
	Procedure Save
		&& Add any code to be called when the Save() button is clicked.
		&& You can also set a BindEvent() call to this method to react to the Save button click.
	EndProc
	
	*---------------------------------------------------------------------------------------
	Procedure Activate

		This.Refresh()

	Endproc
	
 	*--------------------------------------------------------------------------------------- 
	Procedure Destroy
	
		If Vartype(This.oRenderEngine) = 'O'
			This.oRenderEngine.Destroy() && This will force objects on RE to get released
		EndIf
 		
		This.oRenderEngine = .null.
		This.oBusinessObject = .null.
		This.oDataObject = .null.

		Store .null. to (This.cHandle)

		RemoveProperty(_screen, This.cHandle)
		 
		If This.lClearEventsOnClose 
			Clear Events
		Endif

	EndProc
	
	*---------------------------------------------------------------------------------------
	Procedure QueryUnload
	
		*-- This Event is triggered when the Form's close button [X] is clicked.
		If This.lRestoreDataOnCancel = .t.
			This.RestoreData()
		Endif	

	EndProc

	*--------------------------------------------------------------------------------------- 
	Procedure Show(tnStyle, toHostForm  )

		*-- Params:
		*-- tnStyle. 1 = Modal (Default), 0 = Modeless
		*-- toHostForm. (Optional) The form this was called from. If passed, we will center this form in the center of host form.

		Local lnAnchor, lnStyle, lnX, loControl
	
		*--		lnAnchor = This.cntMain.Anchor
		*--		This.cntMain.Anchor = 0
		*--		This.cntMain.Width = Max(This.MinWidth, This.Width)
		*--		This.cntMain.Height = Max(This.MinHeight, This.Height)
		*--		This.cntMain.Anchor = lnAnchor


		If This.oRenderEngine.lRendered = .f.
			This.Render()
		EndIf

		If Pcount() < 2 and (Version(2) <> 0) && No host passed, and working in dev mode
			This.Left = Max(Int(Min(_vfp.Width,1600) - This.Width) / 2, 0)
			This.Top  = Max(Int((_vfp.Height - This.Height) / 2) - Sysmetric(9) - 0, 0)
		Endif

		*-- If  a reference to the calling form was passed, then center this form in host form
		If Vartype(toHostForm) = 'O'
			Do Case
				Case PemStatus(toHostForm, 'ShowWindow', 5) and toHostForm.ShowWindow = 1
				This.Left = Max(Int(toHostForm.Width - This.Width) / 2 + toHostForm.Left, 0)
				This.Top  = Max(Int(toHostForm.Height - This.Height - Sysmetric(9)) / 2  + toHostForm.Top, 0)
				Case PemStatus(toHostForm, 'ShowWindow', 5) and toHostForm.ShowWindow = 2
				This.Left = Max(Int(toHostForm.Width - This.Width) / 2, 0)
				This.Top  = Max(Int(toHostForm.Height - This.Height - Sysmetric(9)) / 2, 0)
				Otherwise
				This.Left = Max(Int(toHostForm.Width - This.Width) / 2 + toHostForm.Left, 0)
				This.Top  = Max(Int(toHostForm.Height - This.Height) / 2 + toHostForm.Top, 0)
			Endcase
		Endif
	
		If Vartype(toHostForm) = 'C'
			This.Top = Val(GetWordNum(toHostForm, 2, ','))
			This.Left = Val(GetWordNum(toHostForm, 1, ','))
		Endif
			
		If Vartype(tnStyle) # 'N' or tnStyle < 0 or tnStyle > 1
			lnStyle = 1
		Else
			lnStyle = tnStyle
		Endif
		
		If lnStyle = 0 && Modeless
			AddProperty(_screen, This.cHandle, This)
		Endif
		
		This.WindowType = lnStyle
		DoDefault(lnStyle)
	
		*-- Set Focus handling...	
		If Type('This.cntMain.DF_oSetFocus') = 'O'
			This.cntMain.DF_oSetFocus.SetFocus()
		Else && Set focus to first enabled control
			For lnX = 1 to This.cntMain.ControlCount
				loControl = This.cntMain.Controls(lnX)
				If PemStatus(loControl, 'Enabled', 5) and PemStatus(loControl, 'SetFocus', 5) and loControl.Enabled = .t.
					loControl.Setfocus()
				Exit
				Endif
			EndFor
		Endif
		
		*-- For some reason the Save button will not appear unless the form is resized. Crazy. So, I jiggle the size around and it appears!!!
		Thisform.Width = Thisform.Width + 1
		Thisform.Width = Thisform.Width - 1

	EndProc	
	
	*--------------------------------------------------------------------------------------- 
	Procedure Hide
	
		DoDefault()
		
		If This.lClearEventsOnClose 
			Clear Events
		Endif
		
	Endproc
	
	
	*--------------------------------------------------------------------------------------- 
	Procedure Render(tcBodyMarkup)

		Local lcRenderSizeMessage, llReturn, lnAnchor, lnRenderHeight, lnRenderWidth

		If Vartype(tcBodyMarkup) = 'C'
			This.cBodyMarkup = tcBodyMarkup
		EndIf
		
		If This.oRenderEngine.lRendered = .f.
			This.SetupRenderEngine()
			llReturn = This.oRenderEngine.Render()
		EndIf

		lnAnchor = This.cntMain.Anchor
		This.cntMain.Anchor = 0
		
		*-- Move continer for any margin-top or margin-bottom that was set
		This.cntMain.Left = This.cntMain.Left + This.cntMain.margin_left
		This.cntMain.Top = This.cntMain.Top + This.cntMain.margin_top
		
		*-- If form it still at its default size, then resize to fit the size of cntMain, which now has all its controls in it.
		If This.Width = 10000
			With This.cntMain
				This.Width = Max(.Width + .Left + .margin_right, This.MinWidth)
			Endwith
		EndIf

		If This.Height = 10000
			With This.cntMain
				This.Height = Max(.Top + .Height + .margin_bottom , This.MinHeight)
			Endwith
		EndIf		

		*-- Make sure container Width and Height fills up the entire width of the form.
		If !PemStatus(This.cntMain, 'container_width', 5)
			This.cntMain.Width = Thisform.Width - This.cntMain.Left - This.cntMain.margin_right
		EndIf
		If !PemStatus(This.cntMain, 'container_height', 5)
			This.cntMain.Height = Thisform.Height - This.cntMain.Top - This.cntMain.margin_bottom
		Endif
		
		lcRenderSizeMessage = ''
		lnRenderWidth = This.cntMain.Left + This.cntMain.Width + This.cntMain.margin_right
		If lnRenderWidth > This.Width
			lcRenderSizeMessage = 'Warning: Rendered control area is wider than form width.' + CRCR + ;
										'[' + Transform(lnRenderWidth) + ' vs.' + Transform(This.Width) + ']'
								 
			lnAnchor = lnAnchor - 8
		Endif
			
		lnRenderHeight = This.cntMain.Top + This.cntMain.Height + This.cntMain.margin_bottom
		If lnRenderHeight > This.Height
			lcRenderSizeMessage = Iif(!Empty(lcRenderSizeMessage), CRCR + lcRenderSizeMessage, '') + ;
								'Rendered control area it taller than form height. '+ CRCR +;
								'[' + Transform(lnRenderHeight) + ' vs.' + Transform(This.height) + ']'
			lnAnchor = lnAnchor - 4
		Endif
		
		If !Empty(lcRenderSizeMessage)
			MessageBox(lcRenderSizeMessage, 64, 'Render size warning:')
			This.oRenderEngine.AddError(lcRenderSizeMessage, .null.)
		EndIf
			
		This.cntMain.Anchor = Iif(lnAnchor > 0, lnAnchor, 0)
		
		Return llReturn
	
	EndProc
	
	*---------------------------------------------------------------------------------------
	Procedure RestoreData
	
		If Vartype(This.oRenderEngine) = 'O'
			This.oRenderEngine.RestoreData()
		Endif
	
	EndProc 
	
	*--------------------------------------------------------------------------------------- 
	Procedure SetupRenderEngine

		With This.oRenderEngine
			.cAlias = This.cAlias
			.oBusinessObject = This.oBusinessObject
			.oDataObject = This.oDataObject
			.cDataObjectRef = This.cDataObjectRef
			.cBusinessObjectSaveMethod = This.cBusinessObjectSaveMethod
			.oContainer = This.cntMain
			.lResizeContainer = .t.
		EndWith			

	EndProc
	
		*--------------------------------------------------------------------------------------- 
		*-- This method allows you to pass in toBusinessObject and toDataObject all at once.
		Procedure BindBusinessAndDataObjects(toBusinessObject, toDataObject)

		
		This.oBusinessObject = Evl(toBusinessObject, .null.)
		This.oDataObject = Evl(toDataObject, .null.)
		
	Endproc
	
	
	
EndDefine

*======================================================================================= 
Define Class DynamicFormRenderEngine as Custom

	*-- See website for complete documentation.
	*-- https://github.com/mattslay/DynamicForms
	
  	cVersion = '1.9.1'
  	cVersionFull = '1.9.1 Production Release - August 30, 2017'

	cAlias = '' 		&& The name of a cursor or alias to which the cMarkup controls are bound

	*-- These properties deal with the data object and properties/fields on it
	oBusinessObject = .null.
	oDataObject = .null.	&& The Object to which the cMarkup controls are bound.
	cDataObjectRef = '' 	&& The reference to the oDataObject to be used in the ControlSource property of each control that gets generated.. 
							&& I.e.  'Thisform.oBusObj.oData' 
	cBusinessObjectSaveMethod = ''
						
	cSkipFields = ''
	*cDisabledFields = '' 2012-09-26 Support for this feature has been removed.

	cAttributeNameDelimiterPattern 	= ':'
	cAttributeValueDelimiterPattern 	= '=>'
	cFieldDelimiterPattern			= '|'		 && Caution. Don't use a comma as delimiter. It will likely break things!

	oContainer = .null.	&& The container in which controls will be rendered. Set by the calling form.
	
	*-- These properties are only used when a BusinessObject is configured to handle the Save button click.
	lShowSaveErrors = .t.	&& Determines if an error dialog will appear if the call to the Business Object Save() method returns .f.
	cSaveErrorMsg = 'Could not save data in Business Object.'
	cSaveErrorCaption = 'Warning...'

	cHeading = ''
	cSaveButtonCaption = .null.
	cCancelButtonCaption = .null.
	nHeadingFontSize = 14
	cHeaderMarkup = .null.			&& See GetHeaderMarkup() for default markup string
	cBodyMarkup = .null.			&& The markup/field list for the "body" of the form.
	cFooterMarkup = .null.			&& See GetFooterMarkup() for default markup string
	cPopupFormBodyMarkup = .null.	&& See GetPopupFormBodyMarkupMarkup() for default markup string

	*-- These properties control the visual layout and flow of the UI controls
	nControlLeft = .null. 		&& See Render method for calculation of default value
	nFirstControlTop = .null.	&& See Render method for calculation of default value
	nVerticalSpacing = .null.	&& See Render method for calculation of default value
	nVerticalSpacingNonControlSourceControls = .null.
	nHorizontalSpacing = .null.	&& Only used when rendering on the same row with .row=increment = '0'
	nHorizontalLineLeft = 10
	nControlHeight = 24 		&& I.e. the Height of Textboxes
	nCheckboxHeight = 24		&& Default Height for Checkboxes
	nCommandButtonHeight = This.nControlHeight
	nHorizontalLabelGap = 8 	&& The horizontal spacing between the label and the input control (when label are NOT above the inputs)
	lLabelsAbove = .f.			&& Default position if for labels to be inline with the input control, to its left. Set this property to .t. to have the labels placed ABOVE the input control.
	lAutoAdjustVerticalPositionAndHeight = .f.	&& Forces the .Top and .Height of each control to 'snap'? to a grid system based on increments on
													&& nControlHeight and nVerticalSpacing. The helps keeps control vertically aligned when form spans two
													&& columns or more. When enabling this feature, any .Top and .Height values specified in attributes may
													&& be adjusted to �snap� to the grid system at its incremental points.
	lResizeContainer = .f.	&&Indicates if engine should resize (enlarge) oContainer to fit controls as they are added.
	lGenerateEditButtonForEditBoxes = .t.	&& If the field is form a cursor or table and it is a mem data type
											&& DF can render a small command button beside the editbox which can be used
											&& to pop-up a larger editbox for the memo field. This pop-up can also be activated
											&& by double-clicking in the editbox.

	*-- Properties related to the popup editbox form feature
	cPopupFormEditboxClass = 'editbox'
	nPopupFormEditboxWidth = 500
	nPopupFormEditboxHeight = 300
	*-- Fields related to columns. See. http.//vfpx.codeplex.com/wikipage?title=Dynamic%20Forms#columns
	nColumnWidth = 200
	nColumnHeight = 800 	&& The host container can grow to this height before engine will switch to next column

	nTextBoxWidth = 100
	nEditBoxWidth = 200
	nNumericFieldTextboxWidth = 100
	nDateFieldTextboxWidth = 100
	nDateTimeFieldTextboxWidth = 150
	nCheckBoxWidth = 100
	nControlWidth = 100		&& For any other controls besides the specific ones above

	nCheckBoxAlignment = 0	&& 0 = Middle Left (Default in VFP) - Places caption to the right of checkbox.
							&& 1 = Middle Right - Places caption to the left of checkbox.
	*nWidth = 0				&& The _Assign method for this property will set oContianer Width to this value
	*nHeight = 0				&& The _Assign method for this property will set oContianer Height to this value
	
	*-- Default classes used to create UI controls (You can override these at run time to use your own custom classes.)
	cLabelClass = 'DF_Label'
	cLabelClassLib = ''
	cTextboxClass = 'textbox'
	cTextboxClassLib = ''
	cEditboxClass = 'DF_MemoFieldEditBox'
	cEditboxCLassLib = ''
	cCommandButtonClass = 'DF_ResultButton'
	cCommandButtonClassLib = ''
	cOptionGroupClass = 'optiongroup'
	cOptionGroupClassLib = ''
	cCheckboxClass = 'DF_Checkbox'
	cCheckboxClassLib = ''
	cComboboxClass = 'combobox'
	cComboboxClassLib = ''
	cListboxClass = 'listbox'
	cListboxClassLib = ''
	cSpinnerClass = 'spinner'
	cSpinnerClassLib = ''
	cGridClass = 'grid'
	cGridClassLib = ''
	cImageClass = 'image'
	cImageClassLib = ''
	cTimerClass = 'timer'
	cTimerClassLib = ''
	cPageframeGroupClass = 'pageframe'
	cPageframeClassLib = ''
	cLineClass = 'line'
	cLineClassLib = ''
	cShapeClass = 'shape'
	cShapeClassLib = ''
	cContainerClass = 'container'
	cContainerClassLib = ''
	cClassLib = '' && General classlib where controls can be found, if not specified above.
	
	*--------------------------------------------------------------------------------------- 
	* Control classes based on data types. If specified, will override the default classes.
	cCharacterClass = ''
	cCharacterClassLib = ''
	cNumericClass = ''
	cNumericClassLib = ''
	cDateClass = ''
	cDateClassLib = ''
	cDateTimeClass = ''
	cDateTimeClassLib = ''
	
	*-- Consider these read only ---
	nErrorCount = 0
	oErrors = .null.
	oRegex = .null.

	*======================================================================================= 
	*-- Private properties used/maintained by this class only!!
	*Hidden nFieldCount
	*Hidden nColumnCount
	nNextControlTop = 0
	
	nColumnCount = 1
	nFieldsInCurrentColumn = 1
	oFieldList = .null.
	nLastControlTop = 0
	nLastControlBottom = 0
	nLastControlLeft = 0
	nLastControlRight = 0
	nControlCount = 0
	lInHeader = .f.
	lInBody = .f.
	lInFooter = .f.
	lLastControlRendered  = .f.
	lRendered = .f.
	cMarkup = ''
	
	Dimension aBackup[1]
	Dimension aColumnWidths[1]
	
	*--------------------------------------------------------------------------------------- 
	Procedure Init()
	
		This.oFieldList = CreateObject('Collection')
		This.oErrors = CreateObject('Collection')
	
	EndProc
	
	*--------------------------------------------------------------------------------------- 
	Procedure Destroy
	
		This.oContainer = .null.
		This.oBusinessObject = .null.
		This.oDataObject = .null.
		This.oErrors = .null.
		This.oFieldList = .null.
		This.oRegex = .null.
		
	Endproc
	
	*--------------------------------------------------------------------------------------- 
	Procedure Render(toContainer)

		Local lcCode, lcControlSource, loField
		
		*-- Open table if cAlias points to something not already open
		If !Empty(This.cAlias) and !Used(JustStem(This.cAlias))
			llReturn = This.OpenTable()
			If !lLReturn
				Return .f.
			Endif
		Endif
		
		This.cAlias = JustStem(This.cAlias)
		
		This.oContainer = Iif(Vartype(toContainer) = 'O', toContainer, This.oContainer)
		If Vartype(This.oContainer) = 'U'
			MessageBox('Must pass container object into Render() method, or set .oContainer property.', 0, 'Warning.')
			Return -1
		Else
			AddProperty(This.oContainer, 'oRenderEngine', This) && Temporary. This reference will be cleared out at the end of this method.
		EndIf	

		This.cBodyMarkup = Nvl(This.cBodyMarkup, This.GetBodyMarkupForAll())

		If Vartype(This.oRegex) <> 'O'
			This.oRegex = CreateObject('VBScript.RegExp')
			This.PrepareRegex()
		Endif

		This.PreProcessBodyMarkup()
		This.BuildMarkup()		&& Merges Header, Body, and Footer marker, and adds some special formatting to help with rendering.
		This.BuildFieldList()		&& Build a collection of controls to be rendered by parsing the cMarkup built in BuildMarkup().
			
		*-- Set default values for various class properties, if the user has not set any values to them...
		This.nControlLeft = Nvl(This.nControlLeft , Iif(This.lLabelsAbove = .t., 20, 120))
		This.nFirstControlTop = Nvl(This.nFirstControlTop , Iif(This.lLabelsAbove = .t., 30, 10))
		This.nHorizontalSpacing = Nvl(This.nHorizontalSpacing, 15) && Only used when rendering on the same row with .row=increment = '0'

		If This.lAutoAdjustVerticalPositionAndHeight = .t.
			If This.lLabelsAbove = .t.
				This.nVerticalSpacing = Nvl(This.nVerticalSpacing, 50) &&	Value is distance from the .Top of the last control to the the .Top of the next control
			Else
			This.nVerticalSpacing = Nvl(This.nVerticalSpacing, 30) 
			Endif
		Else
			If This.lLabelsAbove = .t.
				This.nVerticalSpacing = Nvl(This.nVerticalSpacing, 30) &&	Value is distance from the BOTTOM of the last control to the the .Top of the next control
			Else
				This.nVerticalSpacing = Nvl(This.nVerticalSpacing, 15) 
			Endif
		EndIf
		
		This.nVerticalSpacingNonControlSourceControls = Nvl(This.nVerticalSpacingNonControlSourceControls, 10)

		This.nNextControlTop = This.nFirstControlTop
		*This.nLastControlTop = This.nFirstControlTop
		This.nLastControlRight = This.nControlLeft - This.nHorizontalSpacing
		This.aColumnWidths[1] = This.nColumnWidth

		This.cSkipFields = ' ' + Strtran(This.cSkipFields, ',', ' ') + ' '

		*-- Loop over the FieldList collection to render each control, or execute embedded code...
		For Each loField in This.oFieldList FOXOBJECT
			lcControlSource = loField.ControlSource
				If Left(lcControlSource, 1) + Right(lcControlSource, 1)= '()' && If ControlSource element is wrapped in (), then it's to be executed as a VFP code block, Execute it!!
					lcCode = Substr(lcControlSource, 2, Len(lcControlSource) - 2)
					Try
						&lcCode
					Catch
						This.AddError('Error executing code block.', loField)
					EndTry
				Else
					If Empty(lcControlSource) or !(' ' + Upper(lcControlSource) + ' ' $ Upper(This.cSkipFields))
						This.GenerateControl(loField)
					EndIf
				Endif
		EndFor
		
		This.lRendered = .t.  && This indicates that the Render method has been called and has completed.
		
		*-- Remove the reference to this Render Engine from the oContainer
		This.oContainer.oRenderEngine = .null.
		RemoveProperty(This.oContainer, 'oRenderEngine')
		
		Return (This.nErrorCount = 0)
	
	EndProc
	
	*---------------------------------------------------------------------------------------
	*-- If cAlias is specified, but not open, then attempt to open it...
	Procedure OpenTable

		Local lcAlias, loException

		Try
			Use (This.cAlias) Again In 0
			This.cAlias = JustStem(This.cAlias) && Now that it's open, trim off any path and extension
		Catch to loException
			This.oErrors.Add(loException)
			Return .F.
		Endtry

	Endproc
	
	*--------------------------------------------------------------------------------------- 
	*-- This method combines the Header, Body, and Footer markup together, and mixes in a
	*-- little extra markup between each section that will help the rendering process keep
	*-- track of where it is working.
	Procedure BuildMarkup

		
		This.cHeaderMarkup = Nvl(This.cHeaderMarkup, This.GetHeaderMarkup())
		*-- See PreProcessBodyMarkup() method to see how it is preparied for use here
		This.cFooterMarkup = Nvl(This.cFooterMarkup, This.GetFooterMarkup())
		
		Text to This.cMarkup NoShow TextMerge
		
						(This.lInHeader = .t.) |
						<<This.cHeaderMarkup>> |
						(This.lInHeader = .f.) |
						
						(This.nFirstControlTop = This.nLastControlTop + This.nFirstControlTop) |
						(This.nFieldsInCurrentColumn = 1) |
						(This.lInBody = .t.) |
						<<This.cBodyMarkup>> |
						(This.lInBody = .f.) |
						
						(This.nLastControlBottom = This.oContainer.Height) |
						(This.lInFooter = .t.) |
						<<This.cFooterMarkup>> |
						(This.lInFooter = .f.) |
		EndText
		
		This.cMarkup = Chrtran(This.cMarkup, Chr(13) + Chr(10), '  ')	

	EndProc
	
	
	*--------------------------------------------------------------------------------------- 
	Procedure PropertyMatch(tcField, tcList)
	
		Local llSkip, x

		For x = 1 to GetWordCount(tcList, ', ')
			If Like(Upper(Alltrim(GetWordNum(tcList, x, ', '))), Upper(tcField))
			Return .t.
			Endif
		Endfor

		Return .f.
	
	EndProc
	
	*--------------------------------------------------------------------------------------- 
	Procedure AddControl(tcClass, tcClassLib, tcControlSourceField, tcDataType)
		
		Local lcBaseClass, lcClass, lcClassLib, lcControlName, lcPrefix, llNewObject, loControl

		Do Case
			Case Lower(tcClass) == 'label'
				lcClass = This.cLabelClass
				lcClassLib = Evl(This.cLabelClasslib, This.cClassLib)
			Case Lower(tcClass) == 'textbox'
				lcClass = This.cTextBoxClass
				lcClassLib = Evl(This.cTextboxClasslib, This.cClassLib)
				
				*-- These properties, if set, override the default class determined
				If Vartype(tcDataType) = 'C'
					Do Case
						Case tcDataType = 'C'
							lcClass = Evl(This.cCharacterClass, lcClass)
							lcClassLib = Evl(This.cCharacterClassLib, lcClassLib)
						Case tcDataType = 'N'
							lcClass = Evl(This.cNumericClass, lcClass)
							lcClassLib = Evl(This.cNumericClassLib, lcClassLib)
						Case tcDataType = 'D'
							lcClass = Evl(This.cDateClass, lcClass)
							lcClassLib = Evl(This.cDateTimeClassLib, lcClassLib)
						Case tcDataType = 'T'
							lcClass = Evl(This.cDateTimeClass, lcClass)
							lcClassLib = Evl(This.cDateTimeClassLib, lcClassLib)
					Endcase
				Endif
			Case Lower(tcClass) == 'editbox'
				lcClass = This.cEditboxClass
				lcClassLib = Evl(This.cEditboxClasslib, This.cClassLib)
			Case Lower(tcClass) == 'commandbutton'
				lcClass = This.cCommandButtonClass
				lcClassLib = Evl(This.cCommandButtonClasslib, This.cClassLib)
			Case Lower(tcClass) == 'optiongroup'
				lcClass = This.cOptionGroupClass
				lcClassLib = Evl(This.cOptionGroupClasslib, This.cClassLib)
			Case Lower(tcClass) == 'checkbox'
				lcClass = This.cCheckboxClass
				lcClassLib = Evl(This.cCheckboxClasslib, This.cClassLib)
			Case Lower(tcClass) == 'combobox'
				lcClass = This.cComboboxClass
				lcClassLib = Evl(This.cComboboxClasslib, This.cClassLib)
			Case Lower(tcClass) == 'listbox'
				lcClass = This.cListboxClass
				lcClassLib = Evl(This.cListboxClasslib, This.cClassLib)
			Case Lower(tcClass) == 'spinner'
				lcClass = This.cSpinnerClass
				lcClassLib = Evl(This.cSpinnerClasslib, This.cClassLib)
			Case Lower(tcClass) == 'grid'
				lcClass = This.cGridClass
				lcClassLib = Evl(This.cGridClasslib, This.cClassLib)
			Case Lower(tcClass) == 'image'
				lcClass = This.cImageClass
				lcClassLib = Evl(This.cImageClasslib, This.cClassLib)
			Case Lower(tcClass) == 'timer'
				lcClass = This.cTimerClass
				lcClassLib = Evl(This.cTimerClasslib, This.cClassLib)
			Case Lower(tcClass) == 'pageframe'
				lcClass = This.cPageframeClass
				lcClassLib = Evl(This.cPageframeClasslib, This.cClassLib)
			Case Lower(tcClass) == 'line'
				lcClass = This.cLineClass
				lcClassLib = Evl(This.cLineClasslib, This.cClassLib)
			Case Lower(tcClass) == 'shape'
				lcClass = This.cShapeClass
				lcClassLib = Evl(This.cShapeClasslib, This.cClassLib)
			Case Lower(tcClass) == 'container'
				lcClass = This.cContainerClass
				lcClassLib = Evl(This.cContainerClasslib, This.cClassLib)
			Otherwise
				lcClass = tcClass
				lcClassLib = Evl(tcClassLib, This.cClassLib)
		EndCase
		
		Try
			llNewObject = This.oContainer.NewObject(Sys(2015), lcClass, lcClassLib) && Sys(2015) = Random name for object. Will rename below...
			This.nControlCount = This.nControlCount + 1
		Catch
			llNewObject = .f.
		Endtry

		
		If llNewObject = .t.
			loControl = This.oContainer.Controls(This.oContainer.ControlCount) && The last control added. See above.
			lcBaseClass = Lower(loControl.baseclass)
			Do Case
				Case lcBaseclass = 'label'
					lcPrefix = 'lbl'
				Case lcBaseclass = 'textbox'
					lcPrefix = 'txt'
				Case lcBaseclass = 'editbox'
					lcPrefix = 'edit'
				Case lcBaseclass = 'commandbutton'
					lcPrefix = 'cmd'
				Case lcBaseclass = 'optiongroup'
					lcPrefix = 'opt'
				Case lcBaseclass = 'checkbox'
					lcPrefix = 'chk'
				Case lcBaseclass = 'combobox'
					lcPrefix = 'cbo'
				Case lcBaseclass = 'listbox'
					lcPrefix = 'list'
				Case lcBaseclass = 'spinner'
					lcPrefix = 'spinner'
				Case lcBaseclass = 'grid'
					lcPrefix = 'grid'
				Case lcBaseclass = 'image'
					lcPrefix = 'img'
				Case lcBaseclass = 'timer'
					lcPrefix = 'timer'
				Case lcBaseclass = 'pageframe'
					lcPrefix = 'pageframe'
				Case lcBaseclass = 'line'
					lcPrefix = 'line'
				Case lcBaseclass = 'shape'
					lcPrefix = 'shape'
				Case lcBaseclass = 'container'
					lcPrefix = 'cnt'
				Otherwise
					lcPrefix = lcClass
			EndCase		
			
			lcControlName = lcPrefix + Iif(!Empty(tcControlSourceField), Strtran(tcControlSourceField, '.', '_'), '') + '_' + Transform(This.nControlCount)
			
			Try
				loControl.Name = lcControlName
			Catch
			Endtry
		Else
			loControl = .null.
		Endif
		
		Return loControl
	
	Endproc

	*--------------------------------------------------------------------------------------- 
	Procedure GenerateControl(toField)
			
		Local laProperties[1], lcAttribute, lcBackupType, lcClassLib, lcControlClass, lcControlSource
		Local lcDataSource, lcDataType, lcErrorMessage, llSuccess, lnControlMultiplier, lnX, loControl
		Local loLabel, luData, llRender, llIsMemoField
		
		*-- Handle/Test the .render-if clause --------------------
		If PemStatus(toField, 'render_if', 5)
			Try
				llRender = This.GetValue(toField.render_if)
			Catch
				lcErrorMessage = 'Error in .render-if clause for [' + toField.ControlSource + '].'
				This.AddError(lcErrorMessage, toField)
				AddProperty(toField, 'class', 'DF_ErrorContainer')
				AddProperty(toField, 'cErrorMsg', lcErrorMessage)
				If !PemStatus(toField, 'Width', 5)
					AddProperty(toField, 'width', 200)
				Endif
				AddProperty(toField, 'controlsource', toField.ControlSource)
				loControl = This.AddControl('DF_ErrorContainer', '', toField.ControlSource)		
				This.StyleControl(loControl, toField)
			EndTry
			
			If !llRender
				If This.lLastControlRendered = .t. and This.GetValue(toField.row_increment) <> 0
					This.nLastControlTop = This.nNextControlTop + ((This.GetValue(toField.row_increment) - 1) * This.nVerticalSpacing)
					This.nNextControlTop = This.nLastControlTop
					This.nLastControlRight = This.nControlLeft - This.nHorizontalSpacing
				EndIf
				This.lLastControlRendered = .f.
				Return
			Endif
		Endif
		
		
		If !Empty(toField.ControlSource)
		
			Do Case
				*-- Does the referenced controlsource exist on the oDataObject?
				Case Vartype(This.oDataObject) = 'O' and PemStatus(This.oDataObject, toField.ControlSource, 5) && Binding on an Object
					lcControlSource = This.cDataObjectRef + '.' + toField.ControlSource
					lcDataSource = 'This.oDataObject.' + toField.ControlSource
				*--  Does the referenced controlsource exist in cAlias?
				Case !Empty(This.cAlias) and Used(This.cAlias) and !Empty(Field(toField.ControlSource, This.cAlias))
					lcControlSource = This.cAlias  + '.' + toField.ControlSource
					lcDataSource = lcControlSource
					llIsMemoField = This.IsMemoField(This.cAlias, toField.ControlSource)
				*--  Does the referenced controlsource exist in the current work area ?
				Case !Empty(Alias()) and !Empty(Field(toField.ControlSource, Alias()))
					lcControlSource = Alias() + '.' + toField.ControlSource
					lcDataSource = lcControlSource 
					llIsMemoField = This.IsMemoField(Alias(), toField.ControlSource)
				*-- If none of the above were true, let's see if the referenced controlsource is a defined variable, or a Cursor.Field
				Otherwise 
					*-- Probably a Cursor.Field controlsource
					llIsMemoField = This.IsMemoField(GetWordNum(toField.ControlSource, 1, '.'), GetWordNum(toField.ControlSource, 2, '.'))
					lcControlSource = toField.ControlSource
					lcDataSource = lcControlSource
			EndCase

			Try
				luData = Evaluate(lcDataSource)
				lcDataType = Vartype(luData)
				llSuccess = .t.
			Catch 
				*-- If all of the above failed to give us any data, then render an error label in this spot.
				lcErrorMessage = 'Controlsource [' + toField.ControlSource + '] not found.'
				This.AddError(lcErrorMessage, toField)
				AddProperty(toField, 'class', 'DF_ErrorContainer')
				AddProperty(toField, 'cErrorMsg', lcErrorMessage)
				If !PemStatus(toField, 'Width', 5)
					AddProperty(toField, 'width', 200)
				Endif
				AddProperty(toField, 'controlsource', lcControlSource)				
			Endtry
		EndIf

		*-- If we were able to resolve the Controlsource, let's analyze what the original source was
		If llSuccess
			Do Case
			Case Type(JustStem(lcDataSource)) = 'O'
				lcBackupType = 'Object'
			Case '.' $ lcDataSource and Used(JustStem(lcDataSource))
				lcBackupType = 'Alias'
			Case !('.' $ lcDataSource and Used(JustStem(lcDataSource)))
				lcBackupType = 'Property'
			Otherwise
				lcBackupType = ''
			EndCase
		Else
			*lcControlSource = ''
			lcDataType = ''
			lcBackupType = ''
		EndIf
		
		If !Empty(toField.ControlSource)
			*-- Note: These baseclasses will be re-mapped to specific classes from the RenderEngine properties in the AddControl() method call below.
			Do Case
				Case llIsMemoField
					lcControlClass = 'editbox'
				Case lcDataType = 'L'
					lcControlClass = 'checkbox'
				Otherwise
					lcControlClass = 'textbox'
			Endcase
		Else
			lcControlClass = ''
		Endif
		
		*-- Override default class, if one is specified in attributes
		If PemStatus(toField, 'Class', 5)
			lcControlClass = toField.class
		Endif

		*-- Handle ClassLib assignment in the attribute list...
		If PemStatus(toField, 'ClassLibrary', 5)
			lcClassLib = toField.classlibrary
		Else
			lcClassLib = This.cClassLib
		Endif

		*-- Create the control, (also adds it to the continer)
		If !Empty(lcControlClass)
			loControl = This.AddControl(lcControlClass, lcClassLib, toField.ControlSource, lcDataType)
		Else
			Return
		EndIf
		
		If Type('loControl') = 'O' and PemStatus(loControl, 'oRenderEngine', 5)
			loControl.oRenderEngine = this
		Endif

		*-- If control could not be created, show an error container...
		If Vartype(loControl) # 'O'
			llSuccess = .f.
			lcErrorMessage = 'Error in Class/ClassLibrary settings for [' + toField.ControlSource + '].'
			This.AddError(lcErrorMessage, toField)
			AddProperty(toField, 'class', 'DF_ErrorContainer')
			AddProperty(toField, 'cErrorMsg', lcErrorMessage)
			If !PemStatus(toField, 'Width', 5)
				AddProperty(toField, 'width', 200)
			Endif
			AddProperty(toField, 'controlsource', lcControlSource)				
			loControl = This.AddControl('DF_ErrorContainer', '', toField.ControlSource)		
		EndIf

		AddProperty(loControl, 'DataType', lcDataType)

		*-- Set ControlSource on loControl
		Try
			loControl.ControlSource = lcControlSource
			If llSuccess
				This.BackupData(lcDataSource, luData, lcBackupType)
			Endif
		Catch
		EndTry		

		This.StyleControl(loControl, toField) && (Will also add the label)
		
		If This.lGenerateEditButtonForEditBoxes
			If (Lower(loControl.baseclass) = 'editbox' or (PemStatus(toField, 'ShowEditButton',5) and This.GetValue(toField.ShowEditButton)))
				If !(PemStatus(toField, 'ShowEditButton',5) and This.GetValue(toField.ShowEditButton) = .f.)
					ln = loControl.Anchor
					loControl.Anchor = 0
					loControl.Width = loControl.Width - 20
					loControl.Anchor = ln
					loEditButton = This.AddControl('DF_EditButton', '', '')
					AddProperty(loEditButton, 'oEditBox', loControl)
					loEditButton.Visible = .t.
					loEditButton.Top = This.nLastControlTop + 2
		 			loEditButton.Left = This.nLastControlRight - 18 
		 			Do case
		 				Case InList(ln, 4, 6)
		 					loEditButton.Anchor = 4
		 				Case InList(ln, 8, 9, 10, 11, 13, 15)
		 					loEditButton.Anchor = 8
		 				Case InList(ln, 12, 14)
		 					loEditButton.Anchor = 12
		 			EndCase 
				Endif
			Endif
		Endif
		
		
	EndProc
	
	
	*--------------------------------------------------------------------------------------- 
	Procedure IsMemoField(tcCursor, tcField)
		
		Local laFieldsFromAlias[1], llIsMemoField, lnFieldFromArray
		
		If Empty(tcCursor) or Empty(tcField) or not Used(tcCursor)
			Return .F.
		Endif


		Try
			AFields(laFieldsFromAlias, tcCursor)
			lnFieldFromArray = Ascan(laFieldsFromAlias, Upper(tcField))
			If laFieldsFromAlias[lnFieldFromArray + 1] = 'M'
				llIsMemoField = .t.
			EndIf
		Catch
			llIsMemoField = .f.
		Endtry

		Return llIsMemoField

	Endproc

	*--------------------------------------------------------------------------------------- 
	*-- This procedure will set the top, left, width, height, and apply the attributes to the control, 
	Procedure StyleControl(toControl, toField)

		Local laProperties[1], lcAttribute, lnColumnTest, lnControlMultiplier, lnX, loLabel

		*-- Format the control and update a few class properties to manage flow
		With toControl
		
			This.AssignHeight(toControl, toField)
			This.ManageColumn(toControl, toField)
			This.AssignTop(toControl, toField)

			If Lower(toControl.Class) <> 'df_horizontalline' && Special handling for this control in Do Case below. Can't apply Left or Width here.
				This.AssignWidth(toControl, toField)
				This.AssignLeft(toControl, toField)
			Endif

			Do Case
				Case  Lower(.baseclass) = 'commandbutton'
					This.StyleCommandButton(toControl, toField)
				Case Lower(.baseclass) = 'checkbox'
					This.StyleCheckbox(toControl, toField)
				Case Lower(toControl.Class) = 'df_horizontalline'
					This.StyleHorizontalLine(toControl, toField)
			EndCase

			.Visible = .T.		
			
			This.ApplyAttributes(toControl, toField)
			
			If Lower(toControl.class) = 'df_errorcontainer'
				AddProperty(toControl, 'Enabled', .t.) && This is to ensure that error container is rendered as ENABLED
			EndIf
			
			*-- For Checkboxes, may need to change the Left property if Alignment is set for caption on left side
			If Lower(.baseclass) = 'checkbox'
				If .Alignment = 1
					.Left = .Left - .Width + 14
				EndIf
			EndIf		
			
			*-- Create a label for any control that has a ControlSource property (skip Checkboxes, they are already handled above)
			If PemStatus(toControl, 'ControlSource', 5) and !Empty(toField.ControlSource) and !(Lower(toControl.baseclass) = 'checkbox')
				loLabel = This.GenerateLabel(toControl, toField)
				*-- Setting .row-increment = '0' will force this control to be generated on same row as last control, so
				*-- we need to shift the label and control over to the right by the width of the label
				Try
					If This.GetValue(toField.row_increment) = 0 and !This.lLabelsAbove
						toControl.Left = toControl.Left + loLabel.Width
						loLabel.Left = loLabel.Left + loLabel.Width
					EndIf
				Catch
				EndTry
				
				AddProperty(toControl, 'cLabelcaption', loLabel.caption) && Added this in 1.6.3
				
			EndIf

			If This.lResizeContainer = .t.
				This.ResizeContainer(toControl)
			Endif
			
			*-- Test is we need to widen the current column
			lnColumnTest = (toControl.Left + toControl.Width) - (This.GetColumnLeft() + This.aColumnWidths[This.nColumnCount])

			If lnColumnTest > 0
				This.aColumnWidths[This.nColumnCount] = This.aColumnWidths[This.nColumnCount] + lnColumnTest
			Endif

			*-- Store some values to class properties so the rendering flow can continue from here for the next control
			*--This.nFieldsInCurrentColumn = This.nFieldsInCurrentColumn + lnControHeightMultiplier 
			This.nFieldsInCurrentColumn = This.nFieldsInCurrentColumn + 1
			This.nLastControlTop = toControl.Top
			This.nLastControlBottom = toControl.Top + toControl.Height + This.GetValue(toField.margin_bottom)
			This.nLastControlLeft = toControl.Left
			This.nLastControlRight = toControl.Left + toControl.Width + Iif(PemStatus(toField, 'margin_right', 5), This.GetValue(toField.margin_right), 0)
			This.lLastControlRendered = .t.	
		
		EndWith

	
	EndProc
	
	*---------------------------------------------------------------------------------------
	*-- Apply any attributes that were set by the user in cMarkup for this control
	Procedure ApplyAttributes(toControl, toField)
		

		Local laProperties[1], lcAttribute, lcObjectName, lcObjectProperty, llApplied, llValue, lnAnchor
		Local lnX, lxValue

		    AMembers(laProperties, toField) 
			
			* For lnX = 1 to Alen(laProperties)
			* 	lcAttribute = Lower(laProperties[lnX])
			* 	lxValue	  = Getpem(toField, lcAttribute)

			lnX = 0			
			For Each loProperty in toField.oRenderOrder FOXOBJECT
				lnX = lnX + 1
				lcAttribute = Lower(loProperty)
				lxValue	  = Getpem(toField, lcAttribute)

				Do Case
					Case lcAttribute $ 'controlsource top left'
						Loop && Do not apply these attribute, as we've already dealt with them in the AssignXXXXX() methods.
					Case lcAttribute = 'set_focus'
						Try
							llValue = This.GetValue(lxValue)
							If llValue = .t.
								AddProperty(This.oContainer, 'DF_oSetFocus', toControl)
							EndIf
						Catch
						Endtry
						Loop
				Endcase
				
				lnAnchor = toControl.Anchor
				toControl.Anchor = 0

				If Lower(toControl.baseclass) = 'optiongroup' and ('__' $ lcAttribute)
					lcObjectName = JustStem(Strtran(lcAttribute, '__', '.'))
					lcObjectProperty = JustExt(Strtran(lcAttribute, '__', '.'))
					Try
						AddProperty(toControl.&lcObjectName, lcObjectProperty, Evaluate(toField.&lcAttribute.))
					Catch
						Try
							AddProperty(toControl.&lcObjectName, lcObjectProperty, toField.&lcAttribute.)
						Catch
						Endtry
					EndTry
					Loop
				EndIf

				*-- JRN 9/13/2012 . Only use EVAL if the attribute value begins with ' ('
				*-- the special text put in place by ParseField
				*-- Todo: Per JRN, need to add code to handle these special ones:
				*-- Here's the list of properties that PEME uses that can look numeric but must be character.
				*-- 'CAPTION', 'COLUMNWIDTHS', 'COMMENT', 'DISPLAYVALUE', 'FORMAT', 'INPUTMASK', 'TAG', 'TOOLTIPTEXT', 'VALUE'

				llApplied = .F.
				If Vartype(lxValue) = 'C' and Left(lxValue, 2) = ' ('
					Try
						toControl.&lcAttribute.	= Evaluate(lxValue)
						llApplied	= .T.
					Catch
						lxValue = Substr(lxValue, 3, Len(lxValue) - 3)
					EndTry
				Endif
				If !llApplied
					Try
						toControl.&lcAttribute. = lxValue
						llApplied = .t.
					Catch
					Endtry
				EndIf
				
						
				If Lower(lcAttribute) # 'anchor'
					toControl.Anchor = lnAnchor
				Endif

			EndFor	
	Endproc
	
	
	*--------------------------------------------------------------------------------------- 
	Procedure StyleHorizontalLine(toControl, toField)

		Local lnAnchor

		*-- Special handling for DF_HorizontalLine Class ---------------
		With toControl
			lnAnchor = .Anchor
			.Anchor = 0
			AddProperty(toField, 'Height', 0)
			toField.oRenderOrder.Add('Height')
			
			If .Left = -1 && If still at the default value (from its class definition)
				.Left = This.nHorizontalLineLeft
			EndIf
			If .Width = 10000 && If still at the default value (from its class definition)
				.Width = Max(This.oContainer.Width - 2 * toControl.Left, 1)
			EndIf
			.Anchor = lnAnchor
		EndWith
	
	Endproc
	
	*---------------------------------------------------------------------------------------
	Procedure StyleCheckbox(toControl, toField)
		
		Local lnAnchor

		*-- Slightly special handling for Checkboxes, mostly dealing with whether caption is on the left vs. right of the checkbox
		With toControl
			lnAnchor = .Anchor
			.Anchor = 0		
			.Caption = This.GetLabelCaption(toControl, toField)
			.WordWrap = .t.
			.Height = This.nCheckboxHeight
			.Alignment = This.nCheckBoxAlignment && Read default value from class property. Passed attribute may override
			.Anchor = lnAnchor
		Endwith
	Endproc
	
	*---------------------------------------------------------------------------------------
	Procedure StyleCommandButton(toControl, toField)
		
		Local lnAnchor

		With toControl
			lnAnchor = .Anchor
			.Anchor = 0
			.Height = This.nCommandButtonHeight
			.Anchor = lnAnchor
		Endwith
	Endproc
	

	*--------------------------------------------------------------------------------------- 
	Procedure AssignTop(toControl, toField, tnSpacing)
		
		Local luValue
		
		*-- For controls like commandbutton, line, checkbox, picture, etc, we only need a small amount of vertical space between this control and the previous control
		If !PemStatus(toControl, 'ControlSource', 5) or (Lower(toControl.baseclass) = 'checkbox' and This.lLabelsAbove)
			lnSpacing = This.nVerticalSpacingNonControlSourceControls 
		Else
			lnSpacing = This.nVerticalSpacing
		EndIf	

		With toControl
			*-- Setting .row-increment = '0' will force this control to be generated on same row as last control.
			If PemStatus(toField, 'row', 5)
				Try
					.Top = This.nFirstControlTop + ((This.GetValue(toField.row) - 1) * (lnSpacing  + This.nControlHeight))
				Catch
					This.AddError('Error in row attribute value.', toField)
				Endtry
			Else
				Try
					If This.nFieldsInCurrentColumn = 1 and !This.lInFooter && If we are working on the first control in the column.
						.Top = This.nFirstControlTop
					Else
						*-- See what row-increment is. Default is 1. User might have requested 0, or more than 1...
						luValue = This.GetValue(toField.row_increment)
						If luValue = 0
							.Top = This.nLastControlTop
						Else
							.Top = This.nLastControlBottom + lnSpacing + ((luValue -1) * (lnSpacing  + This.nControlHeight))
						EndIf
					Endif
				Catch
					This.AddError('Error in row-increment attribute value.', toField)
				Endtry
			Endif	

			*-- Override with setting from markup attribute, if specified.
			If PemStatus(toField, 'top', 5)
				Try
					.Top = This.GetValue(toField.top)
				Catch
					This.AddError('Error in Top attribute value.', toField)
				Endtry			
			EndIf
			
			*-- Add any .margin-top spacing that was set in attributes
			If Vartype(tnSpacing) = 'L'
				Try
					.Top = .Top + This.GetValue(toField.margin_top)
				Catch
					This.AddError('Error in margin-top attribute value.', toField)
				Endtry
			EndIf

		Endwith
	EndProc
	
	*---------------------------------------------------------------------------------------
	Procedure AssignLeft(toControl, toField)

		*-- First, assign a default Left value
		toControl.Left = This.nControlLeft + This.GetColumnLeft()
		
		*-- Setting .row = "0" will force this control to be generated on same row as last control, and off the the right of the last control
		Try
			If This.GetValue(toField.row_increment) = 0
				toControl.Left = This.nLastControlRight + This.nHorizontalSpacing
			EndIf
			Catch
		EndTry

		*-- Override with setting from markup attribute, if specified.
		If PemStatus(toField, 'left', 5)
			Try
				.Left = This.GetValue(toField.left)
			Catch
				This.AddError('Error in Left attribute value.', toField)
			Endtry			
		Endif
		
		*-- Adjust for .margin-left as specified in markup attributes
		Try
			.Left = .Left + This.GetValue(toField.margin_left)
		Catch
			This.AddError('Error in margin-left attribute value.', toField)
		Endtry			
		
		If PemStatus(toField, 'centered', 5)
			toControl.Left = (This.oContainer.Width - toControl.Width)/2
		Endif
		
	Endproc
	
	*--------------------------------------------------------------------------------------- 
	*-- This method assigns a default Width value, from Render Engine logic and properties, but,
	*-- this value may be overridden by the user in their markup syntax. If so, it will applied in 
	*-- the ApplyAfftributes method() call which occurs later.
	Procedure AssignWidth(toControl, toField)


		Local lcNothing, lnAnchor

		lnAnchor = toControl.Anchor
		toControl.Anchor = 0
	
		If PemStatus(toControl, 'controlsource', 5)
			Do Case
				Case PemStatus(toControl, 'ControlSource', 5) and Empty(toControl.ControlSource)
					*-- Do nothing
				Case toControl.DataType  $ 'N'
					toControl.Width = Int(This.nNumericFieldTextboxWidth)			
				Case toControl.DataType  $ 'D'
					toControl.Width = Int(This.nDateFieldTextboxWidth)			
				Case toControl.DataType  $ 'T'
					toControl.Width = Int(This.nDateTimeFieldTextboxWidth)			
				Case Lower(toControl.baseclass) = 'checkbox'
					toControl.Width = Int(This.nCheckboxWidth)			
				Case Lower(toControl.baseclass) = 'textbox'
					toControl.Width = This.nTextBoxWidth
				Case Lower(toControl.baseclass) = 'editbox'
					toControl.Width = This.nEditBoxWidth
				Otherwise
					toControl.Width = This.nControlWidth
			Endcase
		EndIf
			
		toControl.Anchor = lnAnchor
			
	Endproc

	*--------------------------------------------------------------------------------------- 
	*-- This method assigns a default Height value, from Render Engine logic and properties, but,
	*-- this value may be overridden by the user in their markup syntax. If so, it will applied in 
	*-- the ApplyAfftributes method() call which occurs later.
	Procedure AssignHeight(toControl, toField)

		Local lnControHeightMultiplier


		With toControl
			*-- Slight override to the height if the user passed in a height attribute...
			*-- (We need it to be an integer increment of This.nVerticalSpacing, so let's do a little rounding...)
			lnControHeightMultiplier = 1
			If PemStatus(toField, 'height', 5) and This.lAutoAdjustVerticalPositionAndHeight = .t.
				Try
					lnControHeightMultiplier = Int((This.GetValue(toField.Height) - This.nControlHeight) / This.nVerticalSpacing) + 1
					.Height = This.nControlHeight + (lnControHeightMultiplier - 1) * This.nVerticalSpacing
				Catch
					This.AddError('Error in height attribute value.', toField)
				Endtry
			Else
				If Lower(toControl.baseclass) = 'checkbox'
					.Height = This.nCheckboxHeight
				Else
					.Height = This.nControlHeight && + (lnControHeightMultiplier - 1) * This.nVerticalSpacing
				Endif
			EndIf

		EndWith
		
	Endproc
  

	*---------------------------------------------------------------------------------------
	Procedure GetColumnLeft

		Local lnColumnLeft
		
		*-- Calculate Left position, considering which "column" we are in...
		lnColumnLeft = 0
		For lnX = 1 to Alen(This.aColumnWidths) - 1
			lnColumnLeft = lnColumnLeft + This.aColumnWidths[lnX]
		EndFor
		
		Return lnColumnLeft

	Endproc
	
	
	*--------------------------------------------------------------------------------------- 
	*-- Manage which column we are working in ----------------
	Procedure ManageColumn(toControl, toField)

		Local lnMarginBottom, lnMarginTop, lnX

		If PemStatus(toField, 'column', 5)
			Try
				If This.GetValue(toField.column) > This.nColumnCount
					This.nColumnCount = This.GetValue(toField.column)
					This.nFieldsInCurrentColumn = 1
					*-- This.nNextControlTop = This.nFirstControlTop
				Endif
			Catch
				This.AddError('Error in column attribute value.', toField)
			Endtry	
		Else
			lnMarginTop = This.GetValue(toField.margin_top)
			lnMarginBottom = This.GetValue(toField.margin_bottom)
			If This.GetValue(toField.row_increment) > 0 and ;
				(This.nLastControlBottom + This.nVerticalSpacing + toControl.Height + lnMarginTop + lnMarginBottom) > This.nColumnHeight
				This.nColumnCount = This.nColumnCount + 1
				This.nFieldsInCurrentColumn = 1
				*-- This.nNextControlTop = This.nFirstControlTop
			EndIf
		EndIf

		*-- Update column widths array, and set default column width for any new columns.
		Dimension This.aColumnWidths[This.nColumnCount]
		For lnX = 1 to Alen(This.aColumnWidths)
			If Vartype(This.aColumnWidths[lnX]) = 'L'
			This.aColumnWidths[lnX] = This.nColumnWidth
			Endif
		Endfor	
		
	Endproc
	

	*--------------------------------------------------------------------------------------- 
	Procedure ResizeContainer(toControl)

		Local laControls[1], lnX, loControl, loFixList
		
		lnAnchor = toControl.Anchor
		toControl.Anchor = 0
		
		loFixList = CreateObject('Collection')

		*-- Store each control who use value 4 (bottom) in its anchor setting. Need to clear this, then reset it
		*-- after the container is resized.
		For each loControl in This.oContainer.controls
			If InList(loControl.Anchor, 4, 5, 6, 7, 12, 14, 15)
				loFixList.Add(loControl)
				loControl.Anchor = loControl.Anchor - 4
			Endif
		Endfor
	
		*-- Adjust container Width and Height if the current control size and positions falls off the container size
		With toControl
			If PemStatus(This.oContainer, 'Width', 5)
				This.oContainer.Width = Max(This.oContainer.Width, .left + .width + 10)
			EndIf
			If PemStatus(This.oContainer, 'Height', 5)
				This.oContainer.Height = Max(This.oContainer.Height, .top + .height + 10)
			Endif
		EndWith
		
		For each loControl in loFixList FOXOBJECT
			loControl.Anchor = loControl.Anchor + 4
		Endfor
		
		toControl.Anchor = lnAnchor
	
	Endproc
	
	
	
	*--------------------------------------------------------------------------------------- 
	Procedure GenerateLabel(toControl, toField)

		Local laProperties[1], lcAttribute, lcibute, lnX, loLabel

		loLabel = This.AddControl(This.cLabelClass, This.cLabelClassLib, toField.ControlSource)
	
		With loLabel
		
			If This.lLabelsAbove = .f.
				.Top = toControl.Top + 4
				.Alignment = 1
			Else
				.Top = toControl.Top - 18
			EndIf
			
			.AutoSize = .t.
			.Caption = This.GetLabelCaption(toControl, toField)
			.Visible = .t.

			If This.lLabelsAbove = .f.
				.Left = toControl.Left - This.nHorizontalLabelGap - .Width && The locates it properly in case the Caption was set in an attribute
			Else
				If loLabel.Alignment = 1
					.Left = toControl.Left + toControl.Width - loLabel.Width
				Else
					.Left = toControl.Left
				Endif
			Endif			
			
			*-- Apply any attributes that were set by the user for this field
			AMembers(laProperties, toField) 
			For lnX = 1 to Alen(laProperties)
				lcFullAttribute = laProperties[lnX]
				If Lower(GetWordNum(lcFullAttribute, 1, '_')) = 'label'
					lcAttribute = GetWordNum(lcFullAttribute, 2, '_')
					Try
						.&lcAttribute. = Evaluate(toField.&lcFullAttribute.)
					Catch
						Try
							.&lcAttribute. = toField.&lcFullAttribute.
						Catch
							loLabel.Caption = 'Error in cMarkup for this label'
							loLabel.ForeColor = Rgb(255,0,0)
						Endtry
					EndTry
				Endif
			EndFor
			
			*If PemStatus(toField, 'left', 5) && Re-apply Left if it was hard set in an attribute.
			*	.Left = This.GetValue(toField.Left)
			*Endif
		
		EndWith	
		
		Return loLabel
	
	EndProc
	
	
	*--------------------------------------------------------------------------------------- 
	Procedure GetLabelCaption(toControl, toField)

		Local lcCaption

		lcCaption = ''

		If PemStatus(toField, 'caption', 5)
			lcCaption = toField.Caption
		Else
			If PemStatus(toControl, 'ControlSource', 5)
				lcCaption = toControl.ControlSource
			Else
				If PemStatus(toField, 'ControlSource', 5)
				lcCaption = toField.ControlSource
				Endif
			EndIf

			If '.' $ lcCaption
				lcCaption = JustExt(lcCaption)
			Endif

			lcCaption = Strtran(Proper(lcCaption), '_', ' ')
		EndIf
		
		Return lcCaption

	Endproc
	

	*--------------------------------------------------------------------------------------- 
	*-- This method parses the field definition string passed in (i.e one of the items from cMarkup)
	*-- to return an oField object containing properties of each attribute in the item
	Procedure ParseField(tcParam)

		Local lnX, lcAttribute, lcValue, lcControlSource
		Local loMatch, loMatches, loField, laPositions[1], llHasCodeLine
		Local loRegEx as 'VBScript.RegExp'
		Local lcErrorMsg, lcProperty, loException

		tcParam = This.TrimIt(tcParam)
		loRegEx = This.oRegex
		
		Do Case
			Case Left(tcParam,1) = This.cAttributeNameDelimiterPattern
				lcControlSource = ''
			Case Left(tcParam,1) + Right(tcParam,1)= '()' && This allows FoxPro code to be executed, if it's store as a string in the ControlSource area. No attributes should follow  on this line!!
				llHasCodeLine = .t.
				*--lcControlSource = Substr(tcParam, 2, Len(tcParam) - 2)
				lcControlSource = tcParam
			Otherwise
				* (here is the 1.5.0 version):  lcControlSource = GetWordNum(tcParam, 1, ' .,' + Chr(9)) && There might be no whitespace after property name and before first colon
				* New in 1.6.0:
				lcControlSource = GetWordNum(tcParam, 1, ' ' + Chr(9) + Chr(13)) && There must be at least one whitespace or newline after controlsource and before first attribute
		Endcase
		
		loField = CreateObject('Empty')
		AddProperty(loField, 'ControlSource', lcControlSource)
		AddProperty(loField, 'row_increment', 1) && Go ahead and add this default to every object. Mayb be overridden in the attributes
		AddProperty(loField, 'margin_top', 0) && etc
		AddProperty(loField, 'margin_bottom', 0) && etc
		AddProperty(loField, 'margin_left', 0) && etc
		AddProperty(loField, 'margin_right', 0) && etc
		
		oRenderOrder = CreateObject('Collection')
		AddProperty(loField, 'oRenderOrder', oRenderOrder)
		
		loMatches = loRegEx.Execute(tcParam)

		If Type('loMatches') = 'O' and !llHasCodeLine
			Dimension laPositions[loMatches.Count + 1] && Create and array of position matches ------------
			lnX = 1
			For Each loMatch In loMatches FOXOBJECT
				laPositions[lnX] = loMatch.firstindex
				lnX = lnX + 1
			EndFor
			laPositions[lnX] = Len(tcParam)
			*-- Loop over regex matches to pull out attribute/value pairs
			lnX = 1
			
			For Each loMatch In loMatches FOXOBJECT
				lcAttribute = Alltrim(loMatch.value, 1, Chr(32), Chr(9), Chr(10), Chr(13))
				lcAttribute = Substr(lcAttribute, Len(This.cAttributeNameDelimiterPattern) + 1, Len(lcAttribute) - Len(This.cAttributeNameDelimiterPattern + This.cAttributeValueDelimiterPattern))
				lcAttribute = Alltrim(lcAttribute, 1, Chr(32), Chr(9), Chr(10), Chr(13))
				lcAttribute = Strtran(lcAttribute, '-', '_') && Must convert dashes to underscores, as dashes are not allowed in Property names		

				lcValue = Rtrim(Left(tcParam, laPositions[lnX + 1])) && Trim off everything to the right, starting at the NEXT match pos.
				lcValue = Substr(lcValue, loMatch.firstindex + Len(loMatch.value) + 1)  && Now pull out the value
				lcValue = Alltrim(lcValue, ' ', ',', Chr(9), Chr(10), Chr(13), This.cFieldDelimiterPattern) && Trim off delimiters and white space

				*** JRN 9/13/2012 Properties to be EVAL'd are wrapped in ' (' + lcValue + ')'
				*** note that they can be EVAL'd directly; the parentheses do not hurt
				Do Case
					Case Left(lcValue, 1) = '('
						lcValue = ' ' + lcValue
					Case Left(lcValue, 1) $ '.0123456789-' && allows numbers and logicals
						lcValue = ' (' + lcValue + ')'
					*-- Removed support for this older syntax as of 2014-09-29
					*Case Substr(lcValue, 2, 1)  $ '.0123456789-' && old style numbers and logicals in quotes
					*	lcValue = ' (' + Substr(lcValue, 2, Len(lcValue) - 2) + ')'
					Otherwise
						lcValue = Substr(lcValue, 2, Len(lcValue) - 2) && Trim off first and last char, which would be some kind of string symbol
				Endcase
				
				*-- If a dot is used in the attribute, we'll flag it will 2 underscores, as this is the format for a property assignment
				*-- on a child object within the current control (i.e. optiongroup)
				lcAttribute = Strtran(lcAttribute, '.', '__')
				
				*-- If this attribute name matches a property name on the RenderEngine class, then apply it to the RenderEngine.
				Try
					If PemStatus(this, lcAttribute, 5)
						AddProperty(This, lcAttribute, This.GetValue(lcValue))
					EndIf
				Catch
				EndTry
				
				*-- If this attribute name begins with "form_" and matches a property on the oConatiner, then apply it to the oContainer's form.
				If 'form_' $ Lower(lcAttribute)
					lcProperty = Strtran(lcAttribute, 'form_', '', -1, -1, 1)
					Try
						AddProperty(This.oContainer.parent, lcProperty, This.GetValue(lcValue))
					Catch
					Endtry
				EndIf

				*-- If this attribute name begins with "container_" and matches a property on the oConatiner, then apply it to the oContainer.
				If 'container_' $ Lower(lcAttribute)
					lcProperty = Strtran(lcAttribute, 'container_', '', -1, -1, 1)
					Try
						AddProperty(This.oContainer, lcProperty, This.GetValue(lcValue))
					Catch
					Endtry
				EndIf
				
				Try
					AddProperty(loField, lcAttribute, lcValue) 
					loField.oRenderOrder.Add(lcAttribute) && Store this property in the oRenderOrder collection. Used in ApplyAttributes to apply in the same order as they appear in the markup.
					lnX = lnX + 1
				Catch to loException
					lcErrorMsg = 'Error on [' + lcControlSource + ']. Attribute: ' + lcAttribute + ' Value: ' + lcValue
					AddProperty(loField, 'class', 'label')
					AddProperty(loField, 'caption', 'Error in cMarkup.    ' + tcParam)
					AddProperty(loField, 'forecolor', ' (Rgb(255,0,0))')
					AddProperty(loField, 'fontbold', .t.)
					AddProperty(loField, 'wordwrap', .t.)
					AddProperty(loField, 'width', 500)
					AddProperty(loField, 'tooltiptext', lcErrorMsg)
					This.AddError(lcErrorMsg, toField, loException)
				Endtry
			Endfor
		EndIf
	
		Return loField
		
	EndProc
	
	*======================================================================================= 
	Procedure EscapeForRegex(tcString)

		Local lcString

		lcString = tcString

		lcString = Strtran(tcString, '\', '\\')
		lcString = Strtran(lcString, '+', '\+')
		lcString = Strtran(lcString, '.', '\.')
		lcString = Strtran(lcString, '|', '\|')
		lcString = Strtran(lcString, '{', '\{')
		lcString = Strtran(lcString, '}', '\}')
		lcString = Strtran(lcString, '[', '\[')
		lcString = Strtran(lcString, ']', '\]')
		lcString = Strtran(lcString, '(', '\(')
		lcString = Strtran(lcString, ')', '\)')
		lcString = Strtran(lcString, '$', '\$')

		lcString = Strtran(lcString, '^', '\^')
		*lcString = Strtran(lcString, ':', '\:')
		lcString = Strtran(lcString, ';', '\;')
		lcString = Strtran(lcString, '-', '\-')
		lcString = Strtran(lcString, '?', '\?')
		lcString = Strtran(lcString, '*', '\*')

		Return lcString	

	Endproc
	
	*--------------------------------------------------------------------------------------- 
	Procedure TrimIt(tcString)
		
		Return Alltrim(tcString, 1, ' ', Chr(9))

	Endproc	
	
	*--------------------------------------------------------------------------------------- 
	Procedure GetValue(tuExpression)
	
		If Vartype(tuExpression) = 'C' and Left(tuExpression, 2) = ' ('
			Return Evaluate(tuExpression)
		Else
			Return tuExpression
		Endif

	EndProc
	
	
	*--------------------------------------------------------------------------------------- 
	*-- This method iterates over the fields in just the cBodyMarkup, and calls off the ParseField()
	*-- method to do a little trick before the real rendering happens... The purpose of this is that we want
	*-- to process any field definitions from the cBodyMarkup which are attempting to set RenderEngine class
	*-- properties by using identically named attributes. 
	Procedure PreProcessBodyMarkup
	

		Local lcField, lcMarkup, lnCount, lnLastPos, lnPos, lnX, loField

		lnCount = 1
		lnLastPos = 1	

		*-- The attribute :import-header, if present, will pull in header from the GetHeaderMarkup()
		This.cBodyMarkup = Strtran(This.cBodyMarkup, This.cAttributeNameDelimiterPattern + 'import-header', This.GetHeaderMarkup(), -1, -1 ,1)

		lcMarkup = This.cBodyMarkup

		*-- Read notes at the top of this mehtod to see what this loop does...
		For lnX = 1 to Len(lcMarkup)
			lnPos = Atc(This.cFieldDelimiterPattern, lcMarkup, lnCount)
			If lnPos = 0
				Exit
			Endif
				lcField = Substr(lcMarkup, lnLastPos, lnPos - lnLastPos)
			If !Empty(lcField)
				loField = This.ParseField(lcField)
			Endif
				lnLastPos = lnPos + Len(This.cFieldDelimiterPattern)
				lnX = lnLastPos
				lnCount = lnCount + 1 
		EndFor

	Endproc
  

	*--------------------------------------------------------------------------------------- 
	*-- This method parses the cMarkup items to convert each item into an object in This.oFieldList collection.
	*-- To omit any property from processing, include it in the cSkipFields property.
	Procedure BuildFieldList

		Local lcField, lcMarkup, lnCount, lnLastPos, lnPos, lnX, loField

		lcMarkup = This.cMarkup
		lnCount = 1
		lnLastPos = 1	

		For lnX = 1 to Len(lcMarkup)
			lnPos = Atc(This.cFieldDelimiterPattern, lcMarkup, lnCount)
			If lnPos = 0
				Exit
			Endif
			lcField = Substr(lcMarkup, lnLastPos, lnPos - lnLastPos)
			If !Empty(lcField)
				loField = This.ParseField(lcField)
				This.oFieldList.Add(loField)
			Endif
			lnLastPos = lnPos + Len(This.cFieldDelimiterPattern)
			lnX = lnLastPos
			lnCount = lnCount + 1 
		EndFor
		
	EndProc
	
	*--------------------------------------------------------------------------------------- 
	*-- From. http.//www.berezniker.com/content/pages/visual-foxpro/shallow-copy-object
	Procedure CopyObject(toObject)

		Local  laProps[1], lnI, loNewObject, lcPropName

		If Vartype(toObject) = 'X' && or Type('toObject.Class') <> 'U'  && For Empty class only and Not Null only
			Return Null
		Endif

		loNewObject = Createobject('Empty')

		For lnI = 1 To Amembers(laProps, toObject, 0)
			lcPropName = Lower(laProps[lnI])
			If Type([toObject.] + lcPropName ,1) = "A"
				AddProperty(loNewObject, lcPropName + "[1]", Null)
				= Acopy(toObject.&lcPropName, loNewObject.&lcPropName)
			Else
				AddProperty(loNewObject, lcPropName, Evaluate("toObject." + lcPropName))
			Endif
		Endfor

	Return loNewObject

	Endproc
	
	*--------------------------------------------------------------------------------------- 
	*-- The original value of each control was saved into This.aBackup[] array so it nca be restored,
	*-- if this method is called by the consumer of this class.
	Procedure RestoreData


	For lnX = 1 to Alen(This.aBackup, 1)
		
		If !Empty(This.aBackup[lnX, 1])
			lcControlSource = This.aBackup[lnX, 1]
			luData = This.aBackup[lnX, 2]
			lcType = This.aBackup[lnX, 3]
			
			If lcType = 'Object' or lcType = 'Property'
				Store m.luData to &lcControlSource
			Else
				lcField = JustExt(lcControlSource)
				lcCursor = JustStem(lcControlSource)
				Replace &lcField  With m.luData In &lcCursor
			Endif
		Endif
		
	Endfor
	
	EndProc
	
	*---------------------------------------------------------------------------------------	
	*-- Each time a control is added to the container, we will make a copy of the original value
	*-- so the value can be restored later if the consumer of this class calls RestoreData(). 
	Procedure BackupData(tcControlSource, tuValue, tcType)
		
		
		Dimension This.aBackup[This.nControlCount, 3]
		
		This.aBackup[This.nControlCount, 1] = tcControlSource
		This.aBackup[This.nControlCount, 2] = tuValue
		This.aBackup[This.nControlCount, 3] = tcType

	Endproc

	*---------------------------------------------------------------------------------------
	Procedure AddError(tcMessage, toField, toException)
		
		This.nErrorCount = This.nErrorCount + 1
		
		loError = CreateObject('Empty')
		AddProperty(loError, 'cMsg', tcMessage)
		AddProperty(loError, 'oField', This.CopyObject(toField))
		This.oErrors.Add(loError)
		
	EndProc
	
	*--------------------------------------------------------------------------------------- 
	Procedure GetErrorsAsString
	
		lcString = 'There were ' + Transform(This.nErrorCount) + ' Error(s) rendering the controls.' + Chr(13) + Chr(13)
		
		For each loError in This.oErrors
			lcString = lcString + '[' + Iif(!IsNull(loError.oField), loError.oField.ControlSource, '') + '].  ' + loError.cMsg + Chr(13)
		Endfor
		
		Return lcString
	EndProc
	
	*--------------------------------------------------------------------------------------- 
	Procedure GetHeaderMarkup
	
		Local lcMarkup, lc1, lc2
		
		lc1 = This.cAttributeNameDelimiterPattern
		lc2 = This.cAttributeValueDelimiterPattern

		Text to lcMarkup NoShow TextMerge

			<<lc1>>class <<lc2>> 'label' 
			<<lc1>>render-if 	<<lc2>> (!Empty(This.cHeading))
			<<lc1>>caption 	<<lc2>> (This.cHeading)
			<<lc1>>name 		<<lc2>> 'lblHeading'
			<<lc1>>top 		<<lc2>> 10
			<<lc1>>left 		<<lc2>> 10
			<<lc1>>fontsize 	<<lc2>> (This.nHeadingFontSize)
			<<lc1>>fontbold 	<<lc2>> .f.
			<<lc1>>autosize 	<<lc2>> .t.  |
			
			<<lc1>>class <<lc2>> 'DF_HorizontalLine'
			<<lc1>>render-if <<lc2>> (!Empty(This.cHeading))
			<<lc1>>margin-top <<lc2>> -10
			<<lc1>>left <<lc2>> 10
			
		EndText
		
		Return lcMarkup
	
	EndProc
	
	*--------------------------------------------------------------------------------------- 
	*-- Loop through all properties on oDataObject and all fields on cALias to build BodyMarkup for all fields,
	*-- skipping any fields that are listed in cSksipFields
	Procedure GetBodyMarkupForAll
	
	
		Local laProperties[1], lcBodyMarkup, lcPropertyFromObject, lnX

		lcBodyMarkup = ''

		If !IsNull(This.oDataObject)
			AMembers(laProperties, This.oDataObject)
			For lnX = 1 to ALen(laProperties)
				lcPropertyFromObject = laProperties[lnX]
				If !Upper(lcPropertyFromObject) $ Upper(This.cSkipFields)
					lcBodyMarkup = lcBodyMarkup + lcPropertyFromObject + '|'
				Endif
			EndFor
		Endif

		If !Empty(This.cAlias)
			AFields(laProperties, This.cAlias)
			For lnX = 1 to ALen(laProperties) Step 18
				lcPropertyFromObject = laProperties[lnX]
				If !Upper(lcPropertyFromObject) $ Upper(This.cSkipFields)
					lcBodyMarkup = lcBodyMarkup + lcPropertyFromObject + '|'
				Endif
			EndFor
		Endif

		Return lcBodyMarkup
				
	Endproc
	

	*--------------------------------------------------------------------------------------- 
	Procedure GetFooterMarkup
	
		Local lcMarkup, lc1, lc2
		
		lc1 = This.cAttributeNameDelimiterPattern
		lc2 = This.cAttributeValueDelimiterPattern

		Text to lcMarkup NoShow TextMerge
		
			<<lc1>>class <<lc2>> 'DF_HorizontalLine'
			<<lc1>>name <<lc2>> 'lineFooter' 
			<<lc1>>left <<lc2>> 10 
			<<lc1>>anchor <<lc2>> 14 |
			
			<<lc1>>class <<lc2>> 'DF_SaveButton'
			<<lc1>>name <<lc2>> 'cmdSave'
			<<lc1>>width <<lc2>> 80
			<<lc1>>left <<lc2>> (This.oContainer.Width - 200)
			<<lc1>>anchor <<lc2>> 12 |

			<<lc1>>class <<lc2>> 'DF_CancelButton'
			<<lc1>>name <<lc2>> 'cmdCancel'
			<<lc1>>width <<lc2>> 80
			<<lc1>>left <<lc2>> (This.oContainer.Width - 100)
			<<lc1>>row-increment <<lc2>> 0
			<<lc1>>anchor <<lc2>> 12 |
			
		EndText
		
		Return lcMarkup
	
	Endproc
	
	
	*---------------------------------------------------------------------------------------
	Procedure GetPopupFormBodyMarkup
	
		Local lcMarkup, lc1, lc2
		
		lc1 = This.cAttributeNameDelimiterPattern
		lc2 = This.cAttributeValueDelimiterPattern

		Text to lcMarkup NoShow TextMerge

			__ControlSource__ 
				<<lc1>>class <<lc2>> '<<This.cPopupFormEditboxClass>>'
				<<lc1>>left <<lc2>> 10
				<<lc1>>width <<lc2>> <<This.nPopupFormEditboxWidth>>
				<<lc1>>height <<lc2>> <<This.nPopupFormEditboxHeight>>
				<<lc1>>anchor <<lc2>> 15
				<<lc1>>label.caption <<lc2>> '' |
		EndText
		
		Return lcMarkup	
		
	Endproc
	
	*--------------------------------------------------------------------------------------- 
	Procedure PrepareRegex
	
		Local lcPattern, loMathes

		*-- This method will analyze the cBodyMarkup block and attempt to assess wheter it uses the origial : and => delimiters
		*-- for attributes and values, or the newer . and = delimiters. If neither of these appear to be used, it will use the 
		*-- properties for cAttributeNameDelimiterPattern and cAttributeValueDelimiterPattern from the class properties.
	
			*--	'\w*\.?-?\w*\s*' + ;&& This one only supported one dash. New one supports mulitples dashes in attribute names
		
		lcPattern = '\w*\.?[-?\w*]*\s*' && .Pattern allows for '-' in attribute name. Will be converted to underscore, when stored as a poperty on toField object		
		
		With This.oRegEx  && .Pattern allows for '-' in attribute name. Will be converted to underscore, when stored as a poperty on toField object
			.IgnoreCase	= .T.
			.Global	= .T.
			.MultiLine = .T.
		EndWith		

		This.oRegEx.Pattern	= '(:' + lcPattern + '=>\s*)'
		loMathes = This.oRegEx.Execute(This.cBodyMarkup)
	
		If loMathes.count > 0 && First, look for the original markup delimieters (: and =>)
			This.cAttributeNameDelimiterPattern = ':'
			This.cAttributeValueDelimiterPattern = '=>'
		Else
			This.oRegEx.Pattern	= '(.' + lcPattern + '=\s*)'
			loMathes = This.oRegEx.Execute(This.cBodyMarkup) && First, look for the newer markup delimieters (. and =)
			If loMathes.count > 0
				This.cAttributeNameDelimiterPattern = '.'
				This.cAttributeValueDelimiterPattern = '='
			Endif
		Endif
		
		*-- If about two test had not mathes, then use whatever properties are on the class
		If loMathes.count = 0
			This.oRegEx.Pattern	= '(' + This.EscapeForRegex(This.cAttributeNameDelimiterPattern) + ;
							 							lcPattern + ;
														This.EscapeForRegex(This.cAttributeValueDelimiterPattern) + ;
														'\s*)'
		Endif

	EndProc
	

EndDefine 


*======================================================================================= 
Define Class DF_ErrorContainer as Container

	BorderStyle = 1
	BackStyle = 0	
	BorderWidth = 1
	BorderColor = Rgb(255,0,0)
	
	cErrorMsg = ''
	DataType = ''
	row_increment = 1
	controlsource = ''


	Add Object lblError as label with;
		Top = 5, ;
		Left = 5, ;
		AutoSize = .t., ;
		FontName = 'Arial', ;
		FontSize = 9, ;
		Name = 'lblError', ;
		Forecolor = Rgb(255,0,0), ;
		Caption = ''
	
	*--------------------------------------------------------------------------------------- 
	Procedure cErrorMsg_Assign(tcMessage)
		
		This.lblError.Caption = tcMessage
		This.Refresh()

	Endproc
	
Enddefine

*======================================================================================= 
Define Class DF_ResultButton as CommandButton 

	*---------------------------------------------------------------------------------------
	Procedure Click

		AddProperty(Thisform, 'cReturn', This.Tag)
		
		If thisform.WindowType = 0 && If Modeless, Release the form. 
			Thisform.Release()
		Else
			Thisform.Hide() && If Modal, just hide the form. Calling code should release it when ready to do so.
		EndIf
		
	EndProc
	
	*--------------------------------------------------------------------------------------- 
	Procedure Init

		This.Tag = This.Caption

	Endproc
	
	
	*---------------------------------------------------------------------------------------
	Procedure Caption_Assign(tcCaption)
		
		This.Caption = tcCaption
		This.Tag = Strtran(tcCaption, '\<', '')

	Endproc
	

Enddefine  

*======================================================================================= 
Define Class DF_SaveButton as DF_ResultButton

	Width = 50
	Height = 30
	Default = .t.
	Caption = 'Save'

	*--------------------------------------------------------------------------------------- 
	Procedure Init

		AddProperty(Thisform, 'lSaveClicked', .f.)
		This.Caption = Nvl(This.parent.oRenderEngine.cSaveButtonCaption, This.Caption)
		DoDefault()

	Endproc

	*--------------------------------------------------------------------------------------- 
	Procedure Click
	
		Local llReturn, loRenderEngine
		loRenderEngine = Thisform.oRenderEngine
		
		*-- If a BusinessObject and Save method is assigned, call the BO to save the data
		If Vartype(loRenderEngine.oBusinessObject) = 'O' and !Empty(loRenderEngine.cBusinessObjectSaveMethod)
			lcSaveCommand = 'llReturn = Thisform.oRenderEngine.oBusinessObject.' + loRenderEngine.cBusinessObjectSaveMethod
			Try
				&lcSaveCommand
			Catch
				MessageBox('Error calling ' + loRenderEngine.cBusinessObjectSaveMethod + ' on oBusinessObject.', 0, 'Error:')
			EndTry
		Else
			llReturn = .t.
		EndIf
		
		If llReturn = .t.
			Thisform.lSaveClicked = .t.
			Thisform.Save() && Call this in case the Form classs has code in it, and so any BindEvent can be triggered.
			DoDefault()
		Else
			This.SaveError()
		EndIf
		
	EndProc
	
	
	*---------------------------------------------------------------------------------------
	Procedure SaveError	
	
		Local loRenderEngine

		loRenderEngine = Thisform.oRenderEngine
			
		If loRenderEngine.lShowSaveErrors = .t.
			MessageBox(loRenderEngine.cSaveErrorMsg, 0, loRenderEngine.cSaveErrorCaption)
		Endif
	
	EndProc
	
Enddefine


*=======================================================================================
Define Class DF_EditButton as CommandButton

	Caption = '...'
	Width = 20
	Height = 20

	*---------------------------------------------------------------------------------------
	Procedure Click

		This.oEditBox.SetFocus()
		This.oEditBox.EditData()

	Endproc

Enddefine

	
*======================================================================================= 
Define Class DF_CancelButton as DF_ResultButton

	Width = 50
	Height = 30
	Cancel = .t.
	Caption = 'Cancel'

	*--------------------------------------------------------------------------------------- 
	Procedure Init

		This.Caption = Nvl(This.parent.oRenderEngine.cCancelButtonCaption, This.Caption)
		AddProperty(Thisform, 'lCancelClicked', .f.)
		DoDefault()

	Endproc

	*--------------------------------------------------------------------------------------- 
	Procedure Click

		If PemStatus(Thisform, 'lRestoreDataOnCancel', 5) and Thisform.lRestoreDataOnCancel = .t.
			Thisform.RestoreData()
		Endif

		Thisform.lCancelClicked = .t.

		DoDefault()

	Endproc

Enddefine

*=======================================================================================
Define Class DF_Label as Label

	BackStyle = 0 && Transparent

EndDefine


*=======================================================================================
Define Class DF_Checkbox as Checkbox

	BackStyle = 0 && Transparent

Enddefine

*=======================================================================================
Define Class DF_HorizontalLine as Line
	
	Height = 0
	Left = -1
	Width = 10000
	Anchor = 10

Enddefine
  
 
*=======================================================================================
Define Class DF_MemoFieldEditBox as EditBox
	
	cDF_Class = 'DynamicForm'
	cEditboxClass = 'Editbox'
	cSaveButtonCaption = 'OK'
	nEditboxWidth = 500
	nEditboxHeight = 300
	nEditboxAnchor = 15
	oRenderEngine = .null. && Will be set by GenerateControl() method in DF Render Engine
	
	*--------------------------------------------------------------------------------------- 
	Procedure DblClick

		This.EditData()

	Endproc

	*--------------------------------------------------------------------------------------- 
	Procedure EditData

		Local lcBodyMarkup, lcDF_Class, loForm, loParentForm, loParentRenderEngine

		lcDF_Class = This.cDF_Class
		Do Case
			Case ! Upper(lcDF_Class) == Upper('DynamicForm')

			Case ThisForm.DeskTop = .t. 
				lcDF_Class = Thisform.Class
			Case ThisForm.ShowWindow # 0
				lcDF_Class = Thisform.Class
		Endcase
		loForm = CreateObject(lcDF_Class)

		*-- Theses setting control the visual layout of the pop-up form...
		loForm.Caption = 'Edit ' + This.cLabelCaption
		
		loParentRenderEngine = This.oRenderEngine
		
		*-- Pass properties from original Render Engine to new form...
		With loParentRenderEngine
			loForm.oDataObject 		= .oDataObject
			loForm.oBusinessObject	= .oBusinessObject
			loForm.cDataObjectRef	= .cDataObjectRef
			loForm.cBusinessObjectSaveMethod = .cBusinessObjectSaveMethod
			loForm.cAlias			= .cAlias
		EndWith
		
		lcBodyMarkup = Nvl(loParentRenderEngine.cPopupFormBodyMarkup, loParentRenderEngine.GetPopupFormBodyMarkup())
		lcBodyMarkup = Strtran(lcBodyMarkup, '__ControlSource__', This.ControlSource, 1, 1 ,1)
		lcBodyMarkup = Strtran(lcBodyMarkup, 'Thisform.oDataObject.', '', 1, 1000 ,1)
		loForm.cBodyMarkup = lcBodyMarkup
		loForm.oRenderEngine.lGenerateEditButtonForEditBoxes = .f.
	
		*-- Find the parent form so we can center new form in parent form
		loParentForm = This.Parent
		Do While Lower(loParentForm.baseclass) # 'form'
			loParentForm = loParentForm.Parent
		EndDo
		
		loForm.Show(Thisform.WindowType, loParentForm)
		
	EndProc
	
	*---------------------------------------------------------------------------------------
	Procedure Destroy

		This.oRenderEngine = .null.

	Endproc

Enddefine

*=======================================================================================
Define Class DynamicFormDeskTop As DynamicForm

	Desktop = .T.
	
Enddefine

*=======================================================================================
Define Class DynamicFormShowWindow As DynamicForm

	ShowWindow = 1
	
Enddefine   
 
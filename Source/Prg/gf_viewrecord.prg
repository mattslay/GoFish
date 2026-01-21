Lparameters toParentForm, loGrid, lcTitle

Local loForm As 'DynamicForm' Of 'C:\Visual FoxPro\VFPX\GoFish\Source\Prg\DynamicForm.prg'
Local laColumns[1, 4], laEvents[1], lcAlias, lcCaption, lcControlSource, lcExtras, lcFieldList
Local lnColumnCount, lnI, loColumn

Private poData

lcAlias			 = Alias()
Scatter Name m.poData Memo

loForm							   = Newobject('DynamicForm', 'PRG\GF_DynamicForm.prg')
loForm.oRenderEngine.nColumnHeight = _Screen.Height - Sysmetric(9) - 100

With m.loForm
	.Caption	 = m.lcTitle
	.cAlias		 = m.lcAlias
	.MinHeight	 = 150
	.BorderStyle = 3

	With .oRenderEngine
		.nVerticalSpacing = -1
		.nTextBoxWidth	  = 600
		.nEditBoxWidth	  = 600

		.nPopupFormEditboxWidth	 = 800
		.nPopupFormEditboxHeight = 50
	Endwith
Endwith && loForm

lnColumnCount = 0
For Each m.loColumn In m.loGrid.Columns FoxObject
	If m.loColumn.Visible Or m.loColumn.gnlevel <= 2
		lnColumnCount = m.lnColumnCount + 1
		Dimension m.laColumns[m.lncolumncount, 4]
		laColumns[m.lncolumncount, 1] = m.loColumn.ColumnOrder
		laColumns[m.lncolumncount, 2] = Trim(Strtran(m.loColumn.ControlSource, m.lcAlias, 'poData', 1, -1, 1))
		laColumns[m.lncolumncount, 3] = m.loColumn.Header1.Caption
	Endif
Endfor

Asort(m.laColumns, 1, -1, 0)

lcFieldList	= ''
lcExtras	= ' .left = 150 ' && .enabled = .f. .DisabledForeColor = 0'
For lnI = 1 To m.lnColumnCount
	lcControlSource	= m.laColumns[m.lnI, 2]
	lcCaption		= m.laColumns[m.lnI, 3]
	If Atc('filepath', m.lcControlSource) # 0
		lcControlSource	= 'poData.FilePath'
		lcCaption		= 'Full Path'
	Endif

	If m.lnI = 1
		lcFieldList = m.lcControlSource
	Else
		lcFieldList = m.lcFieldList + '| ' + m.lcControlSource
	Endif
	lcFieldList = m.lcFieldList + ' .caption = [' + m.lcCaption	+ ']' + m.lcExtras
	If Atc('timestamp', m.lcControlSource) = 0
		lcFieldList = m.lcFieldList + ' .anchor = 10'
	Endif
Endfor

lcFieldList = m.lcFieldList + '|podata.Statement .caption = [Statement] .class = [DF_MemoFieldEditBox] .anchor = 10 .height = 60' + m.lcExtras

lcFieldList = m.lcFieldList + '|podata.ProcCode .caption = [Procedure] .class = [DF_MemoFieldEditBox] .anchor = 10 .height = 120' + m.lcExtras

lcFieldList = m.lcFieldList + '|podata.ID      .caption = [Record ID] .width = 40' + m.lcExtras

loForm.cBodyMarkup	= m.lcFieldList

m.loForm.Render()

GF_FixFontSize(loForm, 9, m.toParentForm.nFormFontSize)

loForm.cntMain.cmdSAVE.Visible = .F.

m.loForm.Show(1, m.toParentForm)


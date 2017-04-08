Local lcProject, lcSafety, llSCC, llTimeStamps, lnResponse

*!*	lnResponse = MessageBox('Run SSCText to generate ascii code files?', 3, 'Building GoFish...')

*!*	If lnResponse <> 6
*!*		? ' '
*!*		? 'Done.'
*!*		Return
*!*	EndIf

Try
	Clear Classlib 'Lib\GoFishSearchEngine'
Catch
Endtry

Set ClassLib to && Must clear them out, cause we're about to generate ascii files of them

lcProject = 'H:\work\repos\GoFish4_Source\GoFish4.pjx'
? lcProject

?'Updating TimeStamps...'
llTimeStamps = ExecScript(_screen.cThorDispatcher, 'Thor_Proc_UpdateTimeStampsOnProjectFiles', lcProject)

Try
	? ' '
	? 'Thor_Proc_GenerateSccFilesOnProject...'
	llGenereateSccFiles = ExecScript(_Screen.cThorDispatcher, 'Thor_Proc_GenerateSccFilesOnProject', lcProject) 
Catch
 ? 'Error calling [Thor_Proc_GenerateSccFilesOnProject]'
Finally
Endtry


? ' '
? 'Done.'
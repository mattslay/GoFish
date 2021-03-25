*-- These paths are the ones on Matt Slay's development machine. Adjust as needed to your local machine


#DEFINE FOXBIN2PRG "H:\work\util\Thor\Thor\Tools\Components\FoxBIN2PRG\FoxBin2Prg.prg"
#DEFINE SOURCE_FOLDER "H:\work\repos\GoFish\Source"

If !File(FOXBIN2PRG)
	MessageBox("FoxBin2Prg app not found at:" + FOXBIN2PRG, 0, "Error...")
	Return
EndIf


Cd SOURCE_FOLDER

lcFoxBin2Prg = FOXBIN2PRG 
Do (lcFoxBin2Prg) With "GoFish5.pjx", "*"

? "Source code files generated from FoxBin2Prg"
Return



*-- Below is the older version that created SCC, VCC files.
*-- Today, we prefer to use FoxBin2Prg. See new code above

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
Lparameters loUpdateObject

With loUpdateObject
	.AppName		  = 'GoFish4_Beta.APP'
	.ApplicationName	  = 'GoFish_Beta'
	.ToolName		  = 'Thor_Tool_GoFish4_Beta'
	.VersionFileURL		  = 'http://foxpro.mattslay.com/GoFish4/Beta/_GoFishVersionFile.txt'
	.VersionLocalFilename	= 'GoFishVersionFile.txt'
	.RegisterWithThor	  = 'Do "##InstallFolder##GoFish4_Beta.APP" with "Thor", .T.'
Endwith

Return loUpdateObject


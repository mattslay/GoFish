*---------------------------------------------------------------------------------------------------
* Sample of a GoFishBackup program. 
* Note: This is the code from the BackupFile() method GoFish Search Engine. It is shown here
* as an example in case you wich to alter the way backups are handled in GoFish.
*---------------------------------------------------------------------------------------------------

#Define ccBACKUPFOLDER Addbs(Home(7) + 'GoFishBackups')

Lparameters tcFilePath, tnReplaceHistoryId

Local lcBackupPRG, llCopyError
Local laExtensions[1], lcDestFile, lcExt, lcExtensions, lcSourceFile, lcThisBackupFolder, lnI

If This.oSearchOptions.lPreviewReplace = .t.
	Return
Endif

llCopyError = .f.

*-- If the user has created a custom backup PRG, and placed it in their path, then call it instead
lcBackupPRG = 'GoFish_Backup.prg'

If File(lcBackupPRG)
	Do &lcBackupPRG With tcFilePath, tnReplaceHistoryId 
	Return
Endif

If Not Directory (ccBACKUPFOLDER) && Create main folder for backups, if necessary
	Mkdir (ccBACKUPFOLDER)
Endif

* Create folder for this ReplaceHistorrID, if necessary
lcThisBackupFolder = Addbs (ccBACKUPFOLDER + Transform (tnReplaceHistoryId))

If Not Directory (lcThisBackupFolder)
	Mkdir (lcThisBackupFolder)
Endif

* Determine the extensions we need to consider
lcExt = Upper (Justext (tcFilePath))

Do Case
	Case lcExt = 'SCX'
		lcExtensions = 'SCX,SCT'
	Case lcExt = 'VCX'
		lcExtensions = 'VCX,VCT'
	Case lcExt = 'FRX'
		lcExtensions = 'FRX,FRT'
	Case lcExt = 'MNX'
		lcExtensions = 'MNX,MNT,MPR,MPX'
	Otherwise
		lcExtensions = lcExt
Endcase

* Copy each file into the destination folder, if its not already there
Alines (laExtensions, lcExtensions, 0, ',')

For lnI = 1 To Alen (laExtensions)
	lcSourceFile = Forceext (tcFilePath, laExtensions (lnI))
	lcDestFile	 = lcThisBackupFolder + Justfname (lcSourceFile)
	If Not File (lcDestFile)
		Try
			Copy File (lcSourceFile) To (lcDestFile)
		Catch
			If !llCopyError
				This.SetReplaceError('Error creating backup of file.', tcFilePath, tnReplaceHistoryId) 
			Endif
			llCopyError = .t.			
		Endtry
	Endif
Endfor
  
Return !llCopyError  
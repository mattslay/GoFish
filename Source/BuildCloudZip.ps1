# This script builds a ZIP file to be uploaded to the FTP site to distribute GoFish5 through
# Thor Check For Updates.
#  Github: https://github.com/mattslay/GoFish
#  email: MattSlay@jordanmachine.com for assistance with GoFish5 and assistance with distributing

Clear-Host

#=== Variable definitions =========================================
$sourceFolder = "H:\Work\Repos\GoFish\Source\"
$localVersionFile = "GoFishVersionFile.txt"
$cloudVersionFile = "_GoFishVersionFile.txt"

$beta = $false

If ($beta -eq $true) {
	$appName = "GoFish4_Beta"
	$zipFile = "GoFish4_Beta.zip"
	$appFile = "GoFish4_Beta.app"
	$ftpServerPath = "mattslay.com/VFP/GoFish4/Beta/"
}
else {
	$appName = "GoFish5"
	$zipFile = "GoFish5.zip"
	$appFile = "GoFish5.app"
	$ftpServerPath = "mattslay.com/VFP/GoFish5/"
}

Write-Host "ZIP Package builder - " + $appName
Write-Host ""

set-location $sourceFolder
If (Test-Path $zipFile){
	remove-item $zipFile
}

#--- Create fully qualified paths to filenames
$p = $PWD.path + "\"
$localZipFile = $p + $zipFile
$localAppFile = $p + $appFile
$localVersionFile = $p + $localVersionFile

#--- Create an empty ZIP file ---
set-content $zipFile ("PK" + [char]5 + [char]6 + ("$([char]0)" * 18))
(dir $zipFile).IsReadOnly = $false	

#--- Prepare to add files to the Zip  ---
"Building Zip"
$shellApplication = new-object -com shell.application
$zipPackage = $shellApplication.NameSpace($localZipFile)

#-- Add local version file to Zip ---
$zipPackage.CopyHere($localVersionFile)
Start-Sleep -Milliseconds 10

#-- Add .app file to Zip ---
$zipPackage.CopyHere($localAppFile)
Start-Sleep -Milliseconds 1000
""



#--- Display results summary ----------------------------
If ($zipPackage.Items().Count -eq 2)
{
	"ZIP File Created: " +  $zipFile
    ""
    "Contents:"
     "  " + $localAppFile
     "  " + $localVersionFile
     ""
     "You must now upload the ZIP and '_GoFishVersionFile.txt' to the FTP endpoint for distribution through Thor.."

}
else
{
	"File count in ZIP is not correct."
}
	
	



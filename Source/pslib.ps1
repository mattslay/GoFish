#-----------------------------------------------------------------------------------------------	
function AddItemToZip ($zipPackage, $item) {

# $item can be a fully qualified filename or folder path

		$fileCount = $zipPackage.Items().Count

		Write-Host "  " + $item
		$zipPackage.CopyHere($item)
		
		Do {
			Start-Sleep -Milliseconds 1	
		} While ($zipPackage.Items().Count -le $fileCount)
		
}
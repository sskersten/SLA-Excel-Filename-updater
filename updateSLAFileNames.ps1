$SLA_FILENAME_REGEXP = "\w*, \w* - \w* \w* \w* \w* \d* - (\w*|\w* \w*) \d{1,2}-\d{1,2} to \d{1,2}-\d{1,2}"
$SLA_DATE_REGEXP = "\d{1,2}-\d{1,2} to \d{1,2}-\d{1,2}"

$files = (dir *.xlsx).BaseName
foreach ($file in $files) {
	if ($file -match $SLA_FILENAME_REGEXP){
		write-output $file
		
		
		$originalFilenameDates = [regex]::matches([string]$file, "\d{1,2}-\d{1,2}")
		$newFilenameDates = 
			((get-date -Date $originalFilenameDates[0].value).AddDays(7)).ToString('M-d'),
			((get-date -Date $originalFilenameDates[1].value).AddDays(7)).ToString('M-d')
			
		$newFilename = $file -replace $SLA_DATE_REGEXP, ($newFilenameDates[0] + " to " + $newFilenameDates[1])
		
		copy-item -Path (".\" + $file + ".xlsx") -Destination (".\" + $newFilename + ".xlsx")
		
		if (-NOT (Test-Path ".\Old_SLA_Files")) {
			new-item .\Old_SLA_Files -type directory
		}
		
		move-item (".\"+$file+".xlsx") (".\Old_SLA_Files\"+$file+".xlsx") 
	}
}
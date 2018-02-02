#The entire filename used in SLA Excel files
$SLA_FILENAME_REGEXP = "\w*, \w* - \w* \w* \w* \w* \d* - (\w*|\w* \w*) \d{1,2}-\d{1,2} to \d{1,2}-\d{1,2}"
#just the date portion of the filename
$SLA_DATE_REGEXP = "\d{1,2}-\d{1,2} to \d{1,2}-\d{1,2}"

#get all of the filenames in the current directory
$files = (dir *.xlsx).BaseName
foreach ($file in $files) {
	#make sure we only mess with SLA files
	if ($file -match $SLA_FILENAME_REGEXP){
		write-output $file
		
		#get the original dates from the filename
		$originalFilenameDates = [regex]::matches([string]$file, "\d{1,2}-\d{1,2}")
		#make new dates from the old dates by adding 7 to them using the get-date object. Thanks MS!
		$newFilenameDates = 
			((get-date -Date $originalFilenameDates[0].value).AddDays(7)).ToString('M-d'),
			((get-date -Date $originalFilenameDates[1].value).AddDays(7)).ToString('M-d')
		
		#setup new filenames by replacing the old dates with new dates
		$newFilename = $file -replace $SLA_DATE_REGEXP, ($newFilenameDates[0] + " to " + $newFilenameDates[1])
		
		#create a new version of the file with the new date in the filename
		copy-item -Path (".\" + $file + ".xlsx") -Destination (".\" + $newFilename + ".xlsx")
		
		#if there's no Old_SLA_Files folder, make one
		if (-NOT (Test-Path ".\Old_SLA_Files")) {
			new-item .\Old_SLA_Files -type directory
		}
		
		#move whatever we copied to the Old_SLA_Files folder
		move-item (".\"+$file+".xlsx") (".\Old_SLA_Files\"+$file+".xlsx") 

        if ($newFilename.Contains("Lesson Plan")){
            #Modify the Excel sheets with the new dates as well
            $Excel = New-Object -ComObject Excel.Application
            $ExcelWorkBook = $Excel.workbooks.Open((Convert-Path .) + "\" + $newFilename)
            $ExcelWorkSheet = $Excel.WorkSheets.Item(1)
            #$ExcelWorkSheet.activate()

            ##modify values of a test cell
            ##Change week date and current date
            $ExcelWorkSheet.Cells.Item(4,7) = $newFilenameDates[0] + " to " + $newFilenameDates[1]
            $ExcelWorkSheet.Cells.Item(4,13) = ((get-date -Date $originalFilenameDates[1].value).AddDays(-1)).ToString('M-d')
            ##change workshop date cells
            $ExcelWorkSheet.Cells.Item(8,9) = "Workshop Date: " + ((get-date -Date $originalFilenameDates[0].value).AddDays(7)).ToString('M-d')
            $ExcelWorkSheet.Cells.Item(22,9) ="Workshop Date: " + ((get-date -Date $originalFilenameDates[1].value).AddDays(5)).ToString('M-d')

            $ExcelWorkBook.Save()
            $ExcelWorkBook.Close()
            $Excel.Quit()
            [System.Runtime.Interopservices.Marshal]::ReleaseComObject($Excel)
            Stop-Process -Name EXCEL -Force
        }
	} #end of If
} #end of foreach
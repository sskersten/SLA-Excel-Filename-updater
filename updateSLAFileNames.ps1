#just the date portion of the filename
$SINGLEDATE_REGEXP = "\d{1,2}-\d{1,2}"
#The entire filename used in SLA Excel files
$SLA_FILENAME_REGEXP = ("[\w\s]*" + $SINGLEDATE_REGEXP + " to " + $SINGLEDATE_REGEXP)
#old files folder name
$OLDFILES_FOLDER_NAME = ".\Old_SLA_Files"

#get all of the filenames in the current directory
$filenames = (dir *.xlsx).BaseName

#If a file is an SLA file, convert the old date to the new date.
foreach ($filename in $filenames) {
	#make sure we only mess with SLA files
	if ($filename -match $SLA_FILENAME_REGEXP){
		
		#get the original dates from filenames and add 7 to them.
		$originalFilenameDates = [regex]::matches([string]$filename, $SINGLEDATE_REGEXP)
		$newFilenameDates = 
			((get-date -Date $originalFilenameDates[0].value).AddDays(7)).ToString('M-d'),
			((get-date -Date $originalFilenameDates[1].value).AddDays(7)).ToString('M-d')
		
		#setup new filenames by replacing the old dates with new dates
		$newFilename = $filename -replace $SINGLEDATE_REGEXP, ($newFilenameDates[0] + " to " + $newFilenameDates[1])
		
		#create a new version of the file with the new date in the filename
		copy-item -Path (".\" + $filename + ".xlsx") -Destination (".\" + $newFilename + ".xlsx")
            
        write-output ("Updated " + $filename + ".")		

        #shove old files into a folder named after the week it's from
        #if there's no Old_SLA_Files folder, make one
		if (-NOT (Test-Path $OLDFILES_FOLDER_NAME)) {
			new-item $OLDFILES_FOLDER_NAME -type directory
		}

        $oldFilesDumpFold = $OLDFILES_FOLDER_NAME + "\" + $originalFilenameDates[0].value
        #if there's not a folder for the current week, make one
        if (-NOT (Test-Path $oldFilesDumpFold)){
            new-item $oldFilesDumpFold -type directory
        }
		
		#move whatever we copied to the Old_SLA_Files folder
		move-item (".\"+$filename+".xlsx") ($oldFilesDumpFold + "\" +$filename + ".xlsx") 


        #use e
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
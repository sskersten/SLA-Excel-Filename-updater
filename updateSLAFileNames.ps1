#just the date portion of the filename
$SINGLEDATE_REGEXP = "\d{1,2}-\d{1,2}"
#The entire filename used in SLA Excel files
$SLA_FILENAME_REGEXP = ( $SINGLEDATE_REGEXP + " to " + $SINGLEDATE_REGEXP)
#old files folder name
$OLDFILES_FOLDER_NAME = ".\Old_SLA_Files"

#get all of the filenames in the current directory
$filenames = (dir *).BaseName

#If a file is an SLA file, convert the old date to the new date.
foreach ($filename in $filenames) {
	#make sure we only mess with SLA files
	if ($filename -match $SLA_FILENAME_REGEXP){


        $extension = ".xlsx";
        #figure out extension
        if ($filename.Contains("Lesson Plan")){
        write-output ($filename + " is docx.")		
            $extension = ".docx"
        } 

        write-output ($extension)	
		
		#get the original dates from filenames and add 7 to them.
		$originalFilenameDates = [regex]::matches([string]$filename, $SINGLEDATE_REGEXP)
		$newFilenameDates = 
			((get-date -Date $originalFilenameDates[0].value).AddDays(7)).ToString('M-d'),
			((get-date -Date $originalFilenameDates[1].value).AddDays(7)).ToString('M-d')
		
		#setup new filenames by replacing the old dates with new dates
		$newFilename = $filename -replace $SLA_FILENAME_REGEXP, ($newFilenameDates[0] + " to " + $newFilenameDates[1])
		
		#create a new version of the file with the new date in the filename
		copy-item -Path (".\" + $filename + $extension) -Destination (".\" + $newFilename + $extension)
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
		move-item (".\"+$filename+$extension) ($oldFilesDumpFold + "\" +$filename + $extension) 


	} #end of If SLA file
} #end of foreach
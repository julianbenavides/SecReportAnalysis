#Set Relative Paths for scripts
$thePath=Split-Path $MyInvocation.MyCommand.Path -Parent
$thePath=Split-Path $thePath -Parent

#Document path variables
$GcalRepVar = Join-Path (Get-ScriptDirectory) 'selectPrompt.ps1'
$GcalRepVar =  & $GcalRepVar $thePath

#Initial Path to avoid an exception
$orgRepPath = "C:\"

#org Report Name
$orgRepName = "org-SecReport-$(get-date -f yyyy-MM-dd).xlsx"


# Function to determine relative path
function Get-ScriptDirectory { Split-Path $MyInvocation.ScriptName }

# generate the path to the script in the assets directory:
$geoscript = Join-Path (Get-ScriptDirectory) 'ipGeolocation.ps1'



write-host "Opening file: " $GcalRepVar
#Opens the GCAL Report
$Excel = New-Object -ComObject Excel.Application
$Excel.Visible = $false
$ExcelWordBook = $Excel.Workbooks.Open($GcalRepVar)
$ExcelWorkSheet = $Excel.WorkSheets.item(1)

write-host "Creating new org Report Spreadsheet in memory..."
#Creates org Report
$Excel2 = new-object -comobject excel.application
$ExcelWordBook2 = $Excel2.Workbooks.Add()
$ExcelWorkSheet2 = $Excel2.WorkSheets.item(1)
$Excel2.Visible = $false
$Excel2.DisplayAlerts = $false
$Excel2.ScreenUpdating = $false
$Excel2.UserControl = $false
$Excel2.Interactive = $false


#Define extra variables
$i=1
$j=1
$total=0
$Addresses = $null
$geoLoc = $null
    
write-host "Starting the data extraction process from the GCAL spreadsheet.."
do
{
    #Debug point
    #write-host "DEBUG Count: $i"
    try
    {
        $Addresses = [System.Net.Dns]::GetHostAddresses($ExcelWorkSheet.Cells.Item($i,1).Text)
        $Addresses = $Addresses.IPAddressToString
        #Debug point
        #write-host "DEBUG TRY1: $Addresses"
        
        #If IP is valid, not local, and can be resolved, let's try to get the geolocation
        if($Addresses -ne "127.0.0.1" -And $Addresses -ne "")
        {
            #Change to relative path
            $geoLoc = & $geoscript $Addresses
            #Debug point
            #write-host "DEBUG GEO Concat: " $geoscript $Addresses
        }
        else
        {
            $geoLoc = "GeoLocation not available."
        }
    }
    catch
    {
        #Debug point
        #write-host "Catch# $i : $Addresses"

        $Addresses = "Domain Name can't resolve to an IP."
        $geoLoc = "GeoLocation not needed."
    }
    #Verify that the record is not empty
    if($ExcelWorkSheet.Cells.Item($i,1).Text -ne "")
    {
        #Debug point
        #write-host "DEBUG NOTEndOfFile: " $geoscript $Addresses
        
        #Add Information to org Report
        $ExcelWorkSheet2.Cells.Item($i,$j).Value = $ExcelWorkSheet.Cells.Item($i,1).Text
        if($ExcelWorkSheet.Cells.Item($i,1).Text -eq "Indicator")
        {
            #Debug point
            #write-host "DEBUG FirsLine: Will assign headers to excel"
            $Addresses = "IP"
            $geoLoc = "IP Geolocation"
            
            $ExcelWorkSheet2.Cells.Item($i,$j+1).Value = "IP"
            $ExcelWorkSheet2.Cells.Item($i,$j+2).Value = "IP Geolocation"
            $total--
        }
        else
        {
            $ExcelWorkSheet2.Cells.Item($i,$j+1).Value = $Addresses
            $ExcelWorkSheet2.Cells.Item($i,$j+2).Value = $geoLoc
        }
        
        #Report point console
        write-host $ExcelWorkSheet.Cells.Item($i,1).Text          -          $Addresses          -          $geoLoc
    }
    #Clearing geoLoc for next round
    $geoLoc = ""
    $total++
    $i++
}
while ($ExcelWorkSheet.Cells.Item($i-1,1).Text -ne "")

#Adjusting the total counter to decrease the extra loop
$total--

write-host "Number of records processed: " $total

write-host "Closing GCAL Report"
#Closing Excel GCAL Spreasheet
$Excel.Quit()

#Prompt for user to select where to store the org Report
$orgRepPath = Join-Path (Get-ScriptDirectory) 'savePrompt.ps1'
$orgRepPath = & $orgRepPath $thePath
$orgRepVar = $orgRepPath + "\" + $orgRepName

write-host "Saving org Report at: " $orgRepVar

#Saving and closing Exce; org Report Spreadsheet
$ExcelWordBook2.SaveAs($orgRepVar)
$Excel2.Quit()

write-host "Cleaning environment"
## function to close all com objects
function Release-Ref ($ref) {
([System.Runtime.InteropServices.Marshal]::ReleaseComObject([System.__ComObject]$ref) -gt 0)
[System.GC]::Collect()
[System.GC]::WaitForPendingFinalizers()
}
## close all object references
Release-Ref($ExcelWorkSheet)
Release-Ref($ExcelWordBook)
Release-Ref($Excel)
Release-Ref($ExcelWorkSheet2)
Release-Ref($ExcelWordBook2)
Release-Ref($Excel2)

write-host "Process Finished."

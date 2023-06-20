[CmdletBinding(SupportsShouldProcess = $True)]
param (
    [Parameter(Mandatory = $False, HelpMessage = "The name of the computer to retrieve status message for")]
    [string]$ComputerName,
    [Parameter(Mandatory = $False, HelpMessage = "The number of hours past in which to retrieve status messages")]
    [int]$TimeInHours = "24",
    [Parameter(Mandatory = $False)]
    [switch]$CSV,
    [Parameter(Mandatory = $False)]
    [switch]$GridView,
    [Parameter(Mandatory = $False, HelpMessage = "The SQL server name (and instance name where appropriate)")]
    [string]$SQLServer = "atklsccm.kostweingroup.intern",
    [Parameter(Mandatory = $False, HelpMessage = "The name of the ConfigMgr database")]
    [string]$Database = "CM_KOW",
    [Parameter(Mandatory = $False, HelpMessage = "The Advertisement ID of the Task Sequence")]
    [string[]]$TSAdvertisementID = @("KOW200B3", "KOW200B4", "KOW200BB"),
    [Parameter(Mandatory = $False, HelpMessage = "The Task Sequence (package) ID of the Task Sequence")]
    [string]$TaskSequenceID = "KOW010FD",
    [Parameter(Mandatory = $False, HelpMessage = "The path to IIS folder")]
    [string]$IISPath = "$PSScriptRoot\IIS", #"C:\inetpub\OSDReporting\wwwroot",
    [Parameter(Mandatory = $False, HelpMessage = "The location of the smsmsgs directory containing the message DLLs")]
    [string]$SMSMSGSLocation = "",
    [Parameter(Mandatory = $False, HelpMessage = "Specify if using Modern Device Management")]
    [string]$MDM = $True
)


#Found that this file is in a different path based on OS/console version
If (Test-Path ((Split-Path $env:SMS_ADMIN_UI_PATH) + "\X64\system32\smsmsgs")) {
    $SMSMSGSLocation = (Split-Path $env:SMS_ADMIN_UI_PATH) + "\X64\system32\smsmsgs"
}
ElseIf (Test-Path ($ENV:SMS_ADMIN_UI_PATH + '\00000409')) {
    $SMSMSGSLocation = ($ENV:SMS_ADMIN_UI_PATH + '\00000409')
}


# Function to get the date difference
Function Get-DateDifference {
    param
    (
        [Parameter(Mandatory = $true, HelpMessage = "The start date")]
        [datetime]$StartDate,
        [Parameter(Mandatory = $true, HelpMessage = "The end date")]
        [datetime]$EndDate 
    )
    $TimeDiff = New-TimeSpan -Start $StartDate -End $EndDate
    if ($TimeDiff.Seconds -lt 0) {
        $Hrs = ($TimeDiff.Hours) + 23
        $Mins = ($TimeDiff.Minutes) + 59
        $Secs = ($TimeDiff.Seconds) + 59 
    }
    else {
        $Hrs = $TimeDiff.Hours
        $Mins = $TimeDiff.Minutes
        $Secs = $TimeDiff.Seconds 
    }
    $Difference = '{0:00}:{1:00}:{2:00}' -f $Hrs, $Mins, $Secs
    Return $Difference
}


# Function to get the status message description
function Get-StatusMessage {
    [CmdletBinding()]
    param (
        [Parameter()]
        $iMessageID,
        [Parameter()]
        [ValidateSet("srvmsgs.dll", "provmsgs.dll", "climsgs.dll")]
        $DLL,
        [Parameter()]
        [ValidateSet("Informational", "Warning", "Error")]
        $Severity,
        $InsString1,
        $InsString2,
        $InsString3,
        $InsString4,
        $InsString5,
        $InsString6,
        $InsString7,
        $InsString8,
        $InsString9,
        $InsString10
    )

    #Load DLLs. These contain the status message query text
    $Location = "C:\Users\mellunigm\OneDrive - Kostwein Maschinenbau GmbH\Dokumente\Scripts\SCCM\Enumerate Message Strings"
    $stringPathToDLL = "$Location\$DLL"

    #Load Status Message Lookup DLL into memory and get pointer to memory 
    $ptrFoo = $Win32LoadLibrary::LoadLibrary($stringPathToDLL.ToString()) 
    $ptrModule = $Win32GetModuleHandle::GetModuleHandle($stringPathToDLL.ToString()) 
    
    switch ($Severity) {
        "Informational" {
            $Code = 1073741824
            $result = $Win32FormatMessage::FormatMessage($flags, $ptrModule, 1073741824 -bor $iMessageID, 0, $stringOutput, $sizeOfBuffer, $stringArrayInput)
        }
        "Warning" {
            $Code = 2147483648
            $result = $Win32FormatMessage::FormatMessage($flags, $ptrModule, 2147483648 -bor $iMessageID, 0, $stringOutput, $sizeOfBuffer, $stringArrayInput)
        }
        "Error" {
            $Code = 3221225472
            $result = $Win32FormatMessage::FormatMessage($flags, $ptrModule, 3221225472 -bor $iMessageID, 0, $stringOutput, $sizeOfBuffer, $stringArrayInput)
        }
    }
    
    if ($result -gt 0) {
        # Add insert strings to message
        $value = $stringOutput.ToString().Replace("%11", "").Replace("%12", "").Replace("%3%4%5%6%7%8%9%10", "").Replace("%1", $InsString1).Replace("%2", $InsString2).Replace("%3", $InsString3).Replace("%4", $InsString4).Replace("%5", $InsString5).Replace("%6", $InsString6).Replace("%7", $InsString7).Replace("%8", $InsString8).Replace("%9", $InsString9).Replace("%10", $InsString10)

        $MsgStringParams = @{
            type = 'NoteProperty' 
            name = 'MessageString' 
            value = $value
        }

        Write-Verbose $value

        $objMessage = New-Object System.Object
        $objMessage | Add-Member @MsgStringParams
    }

    $objMessage
}

# Open a database connection
$connectionString = "Server=$SQLServer;Database=$database;Integrated Security=SSPI;"
$connection = New-Object System.Data.SqlClient.SqlConnection
$connection.ConnectionString = $connectionString
$connection.Open()

#Get all Advertisement IDs 
$TSAdvString = $null #This will be used in the next foreach loop
[int]$TSCount = $TSAdvertisementID.Count #If we have more than one TS advertisement we have to change up our query
ForEach ($AdvID in $TSAdvertisementID) {
    #If there is more than one, we need to add "or" in our query
    If ($TSCount -gt 1) {  
        $TSAdvstring += "v_StatMsgAttributes.AttributeValue = '" + $AdvID + "' or "
    }
    else {
        $TSAdvstring += "v_StatMsgAttributes.AttributeValue = '" + $AdvID + "'"
    }
    $TScount -= 1
}


#Let's do some logic to pull our UTC offset
[datetime]$UTCtime = (Get-Date).ToUniversalTime()
[datetime]$LocalDT = ([datetime]$UTCtime).ToLocalTime()
$HoursUTCoffset = ($LocalDT.hour - $UTCtime.hour)


# Define the SQl query
$Query = @"
select smsgs.RecordID, 
CASE smsgs.Severity
WHEN -1073741824 THEN 'Error'
WHEN 1073741824 THEN 'Informational'
WHEN -2147483648 THEN 'Warning'
ELSE 'Unknown'
END As 'SeverityName',
case smsgs.MessageType
WHEN 256 THEN 'Milestone'
WHEN 512 THEN 'Detail'
WHEN 768 THEN 'Audit'
WHEN 1024 THEN 'NT Event'
ELSE 'Unknown'
END AS 'Type',
smsgs.MessageID, smsgs.Severity, smsgs.MessageType, smsgs.ModuleName,modNames.MsgDLLName, smsgs.Component,
smsgs.MachineName, smsgs.Time, smsgs.SiteCode, smwis.InsString1,
smwis.InsString2, smwis.InsString3, smwis.InsString4, smwis.InsString5,
smwis.InsString6, smwis.InsString7, smwis.InsString8, smwis.InsString9,
smwis.InsString10, v_StatMsgAttributes.*, DATEDIFF(hour,dateadd(hh,$HoursUTCoffset,smsgs.Time),GETDATE()) as DateDiffer
from v_StatusMessage smsgs
join v_StatMsgWithInsStrings smwis on smsgs.RecordID = smwis.RecordID
join v_StatMsgModuleNames modNames on smsgs.ModuleName = modNames.ModuleName
join v_StatMsgAttributes on v_StatMsgAttributes.RecordID = smwis.RecordID
where (smsgs.Component = 'Task Sequence Engine' or smsgs.Component = 'Task Sequence Action')
and v_StatMsgAttributes.AttributeID = 401 
and ($TSAdvstring)
and DATEDIFF(hour,smsgs.Time,GETDATE()) < $TimeInHours
Order by smsgs.Time DESC
"@


# Run the query
$command = $connection.CreateCommand()
$command.CommandText = $query
$reader = $command.ExecuteReader()
$table = new-object "System.Data.DataTable"
$table.Load($reader)

# Close the connection
$connection.Close()

#Start PInvoke Code
$sigFormatMessage = @'
[DllImport("kernel32.dll")]
public static extern uint FormatMessage(uint flags, IntPtr source, uint messageId, uint langId, StringBuilder buffer, uint size, string[] arguments);
'@ 

$sigGetModuleHandle = @'
[DllImport("kernel32.dll")]
public static extern IntPtr GetModuleHandle(string lpModuleName);
'@ 

$sigLoadLibrary = @'
[DllImport("kernel32.dll")]
public static extern IntPtr LoadLibrary(string lpFileName);
'@ 

$Win32FormatMessage = Add-Type -MemberDefinition $sigFormatMessage -name "Win32FormatMessage" -namespace Win32Functions -PassThru -Using System.Text
$Win32GetModuleHandle = Add-Type -MemberDefinition $sigGetModuleHandle -name "Win32GetModuleHandle" -namespace Win32Functions -PassThru -Using System.Text
$Win32LoadLibrary = Add-Type -MemberDefinition $sigLoadLibrary -name "Win32LoadLibrary" -namespace Win32Functions -PassThru -Using System.Text
#End PInvoke Code 

$sizeOfBuffer = [int]16384
$stringArrayInput = { "%1", "%2", "%3", "%4", "%5", "%6", "%7", "%8", "%9" }
$flags = 0x00000800 -bor 0x00000200
$stringOutput = New-Object System.Text.StringBuilder $sizeOfBuffer 

# Put desired fields into an object for each result
$StatusMessages = @()

foreach ($Row in $Table.Rows) {
    $Params = @{
        iMessageID  = $Row.MessageID
        DLL         = $Row.MsgDLLName
        Severity    = $Row.SeverityName
        InsString1  = $Row.InsString1
        InsString2  = $Row.InsString2
        InsString3  = $Row.InsString3
        InsString4  = $Row.InsString4
        InsString5  = $Row.InsString5
        InsString6  = $Row.InsString6
        InsString7  = $Row.InsString7
        InsString8  = $Row.InsString8
        InsString9  = $Row.InsString9
        InsString10 = $Row.InsString10
    }
    $Message = Get-StatusMessage @params

    #Tell Powershell that our Date/Time is UTC
    $UTCDateTime = [datetime]::SpecifyKind( $Row.time, 'UTC' )
    #Convert it to Local Time
    [datetime]$LocalDateTime = ([datetime]$UTCDateTime).ToLocalTime()
    
    $StatusMessage = New-Object psobject
    Add-Member -InputObject $StatusMessage -Name Severity -MemberType NoteProperty -Value $Row.SeverityName
    Add-Member -InputObject $StatusMessage -Name Type -MemberType NoteProperty -Value $Row.Type
    Add-Member -InputObject $StatusMessage -Name SiteCode -MemberType NoteProperty -Value $Row.SiteCode
    Add-Member -InputObject $StatusMessage -Name "Date / Time" -MemberType NoteProperty -Value $LocalDateTime
    Add-Member -InputObject $StatusMessage -Name System -MemberType NoteProperty -Value $Row.MachineName
    Add-Member -InputObject $StatusMessage -Name Component -MemberType NoteProperty -Value $Row.Component
    Add-Member -InputObject $StatusMessage -Name Module -MemberType NoteProperty -Value $Row.ModuleName
    Add-Member -InputObject $StatusMessage -Name MessageID -MemberType NoteProperty -Value $Row.MessageID
    Add-Member -InputObject $StatusMessage -Name Description -MemberType NoteProperty -Value $Message.MessageString
    $StatusMessages += $StatusMessage
}

$html = @() #Create a blank array
$Messages = $StatusMessages | Sort-Object -Property "Date / Time" -descending | Group-Object -Property System #Grab our status messages, sort and group them.

#Import ConfigMgr module
Import-Module ((Split-Path $env:SMS_ADMIN_UI_PATH) + '\configurationmanager.psd1') -ErrorAction Stop

#Get Site Code. Note: Console must be on this machine for this to work.
$SiteCode = Get-PSDrive -PSProvider CMSite

#Set our drive to point to our SCCM environment
Set-Location "$($SiteCode.Name):"

#Setup some variables
#This gets all of the steps for a task sequence if they are enabled
$TSSteps = (Get-CMTaskSequenceStep -TaskSequenceID $TaskSequenceID) | Where-Object { $_.Enabled -eq 'True' } | Select-Object Name

#If not using Modern Driver Managment
If ($MDM -eq $false) {
    #This gets all of the driver install steps
    $TSDriverSteps = (Get-CMTaskSequenceStepApplyDriverPackage -TaskSequenceId $TaskSequenceID) | Select-Object Name 
    #Get the name of the first step in the list of Driver steps
    $DriverIndexStart = $TSDriverSteps[0].name 
    #Get the index (Start position) of the first driver step in the task sequence
    $index = $TSSteps.Name.IndexOf($DriverIndexStart)
    #Compare the full task sequence step list and the driver steps. Gets all that are not driver steps.
    $TSStepsNoDrivers = Compare-Object -ReferenceObject $TSSteps.Name -DifferenceObject $TSDriverSteps.Name -PassThru 
    #Rebuilds the arrays to replace the driver steps with one step lableed "Install driver"
    $TSStepsNoDrivers = $TSStepsNoDrivers[0..($index - 1)] + "Install Drivers" + $TSStepsNoDrivers[$index..($TSStepsNoDrivers.Length - 1)] 
}
elseif ($MDM -eq $True) {
    #If using Modern Driver Management, no need to process all the driver steps and can just consider all the steps good as is
    $TSStepsNoDrivers = $TSSteps.Name 
}
#RegEx used later
[regex]$ParRegex = "\((.*?)\)" 
$Script:LastLog = $null
$Script:ImageDuration = $null

#Count used when building HTML table
$tablecount = 1 
#Count used when building HTML table header
$stepscount = 1 

#Loop through each computer
ForEach ($Computer in $Messages) { 
    #Here we are resorting so that the newest statmessage comes first. We have to sort again here because the first sort puts the newest computer at top but rearranges the statmsg group
    $Computer = $Computer.Group | Sort-Object -Property "Date / Time" 
    #Green is always the same so we can declare it here.
    $green = '<img src="images/checks/greenCheckMark_round.png" alt="Green Check Mark">' 
    #hash table used for storing variables
    $varHash = [ordered]@{} 
    #Count used for cycling through driver steps
    $DriverCount = 1 

    If ($Computer.MessageID -contains "11144" -or $Computer.MessageID -eq "11140") {

        #Loop through each status message
        ForEach ($Statmsg in $Computer) { 
            $NameDuringImaging = $Statmsg.System #Get the computer name
            $VarCount = 1 #Count used for adding to name (this is used in the HTML to show what step we are on)
            $ImageDuration = $null #Null out some variables between computers so we don't carry over unwanted data
            $ImageCompleted = $null
            #$ImageStarted = $null
            $LastLog = $null
            
            #Loop through each TS Step
            ForEach ($step in $TSStepsNoDrivers) { 
                $var = "$($VarCount)" + ' - ' + $step #This sets the variable to include a number plus the name. Just for the html page

                If ($Statmsg.MessageID -eq "11144" -or $Statmsg.MessageID -eq "11140") { 
                    $ImageStarted = $statmsg."Date / Time" 
                } #MessageID 11144 is the start of a task sequence, 11140 is start if in OS/Software Center
                ElseIf (($step -eq "Install Drivers") -AND ($MDM -eq $false)) {
                    #This ElseIf block is where we process our driver steps only if not using Modern Driver Management    
                    $RegExString = $ParRegex.match($Statmsg.Description).Groups[1].value #Gets the name of the step by looking at the value between the parenthesies.

                    #Gets the correct driver step by comparing the step we are in with the list of driver steps
                    If ($TSDriverSteps -match $RegExString) { 
                        If (($Statmsg.Description -like "The task sequence execution engine successfully completed the action*$($Step)*") -AND ($statmsg.Severity -eq "Informational")) {
                            #Process if Driver step is successful
                            $var = "$($VarCount)" + ' - ' + "Install Drivers"
                            $text = $statmsg.Description.replace('"', '&quot;') #replace quotes so the html doesn't truncate
                            $varHash.Add($var, '<a href=" " title="' + $text + '"><img src="images/checks/greenCheckMark_round.png" alt="Green Check Mark"></a>')  #Adding text of the drivers just for a reference in the html
                            $LastLog = $statmsg."Date / Time"
                            $ImageDuration = Get-DateDifference -StartDate $ImageStarted -EndDate $LastLog
                        }
                        ElseIf (($Statmsg.Description -like "The task sequence execution engine failed executing the action*$($Step)*") -AND ($statmsg.Severity -eq "Error")) {
                            #Process if Driver step errors
                            $var = "$($VarCount)" + ' - ' + "Install Drivers"
                            $errortext = $statmsg.Description.replace('"', '&quot;') #replace quotes so the html doesn't truncate
                            $varHash.Add($var, '<a href=" " title="' + $errortext + '"><img src="images/checks/redCheckMark_round.png" alt="Red Check Mark"></a>') #set pic to red and include error text
                            $LastLog = $statmsg."Date / Time"
                            $ImageDuration = Get-DateDifference -StartDate $ImageStarted -EndDate $LastLog
                        }
                        ElseIf ($DriverCount -eq $TSDriverSteps.Count) {
                            #Process if there are no driver packages for this computer in the TS
                            $var = "$($VarCount)" + ' - ' + "Install Drivers"
                            $text = "There was not a driver package available in the Task Sequence for this device"
                            $varHash.Add($var, '<a href=" " title="' + $text + '"><img src="images/checks/greyCheckMark_round.png" alt="Grey Check Mark"></a>') #Adding text of the drivers just for a reference in the html
                            $LastLog = $statmsg."Date / Time"
                            $ImageDuration = Get-DateDifference -StartDate $ImageStarted -EndDate $LastLog
                        }
                        else {
                            $DriverCount += 1
                        }
                    }                                            
                }
                ElseIf (($Statmsg.Description -like "The task sequence execution engine successfully completed the action*$($Step)*") -AND ($statmsg.Severity -eq "Informational")) {
                    #Processing all successful steps
                    $varHash.Add($var, $green) 
                    $LastLog = $statmsg."Date / Time"
                    $ImageDuration = Get-DateDifference -StartDate $ImageStarted -EndDate $LastLog
                }
                ElseIf (($Statmsg.Description -like "The task sequence execution engine failed executing the action*$($Step)*") -AND ($statmsg.Severity -eq "Error")) {
                    #Processing all error steps
                    $errortext = $statmsg.Description.replace('"', '&quot;') #replace quotes so the html doesn't truncate
                    $varHash.Add($var, '<a href=" " title="' + $errortext + '"><img src="images/checks/redCheckMark_round.png" alt="Red Check Mark"></a>') #set pic to red and include error text
                    $LastLog = $statmsg."Date / Time"
                    $ImageDuration = Get-DateDifference -StartDate $ImageStarted -EndDate $LastLog
                }
                ElseIf (($Statmsg.Description -like "The task sequence execution engine skipped the action*$($Step)*") -AND ($statmsg.Severity -eq "Informational")) {
                    #Processing all skipped steps
                    $skiptext = $statmsg.Description.replace('"', '&quot;') #replace quotes so the html doesn't truncate
                    $varHash.Add($var, '<a href=" " title="' + $skiptext + '"><img src="images/checks/greyCheckMark_round.png" alt="Grey Check Mark"></a>') #set pic to grey and include error text
                    $LastLog = $statmsg."Date / Time"
                    $ImageDuration = Get-DateDifference -StartDate $ImageStarted -EndDate $LastLog
                }
                ElseIf (($Statmsg.MessageID -eq "11171") -or ($statmsg.MessageID -eq "11143")) {
                    #Task Sequence Completed Successfully
                    #Processing end of TS
                    $varHash.Add('Exit Task Sequence', $green)
                    $ImageCompleted = $statmsg."Date / Time"
                    $LastLog = $statmsg."Date / Time"
                    $ImageDuration = Get-DateDifference -StartDate $ImageStarted -EndDate $ImageCompleted
                }
                ElseIf ($Statmsg.MessageID -eq "11141") {
                    #Failed Task Sequence
                    #Processing if the TS failed (If a TS step fails and is set to NOT continue on error)
                    $index = ([array]::indexof($TSStepsNoDrivers, $step) + 1)
                    ForEach ($miniStep in $TSStepsNoDrivers) {
                        If ([array]::indexof($TSStepsNoDrivers, $ministep) -ge $index) {
                            $miniIndex = ([array]::indexof($TSStepsNoDrivers, $ministep) + 1)
                            $var = "$($miniIndex)" + ' - ' + $ministep #This sets the variable to include a number plus the name. Just for the html page
                            $varHash.Add($var, " ") #enter blank entries in our hash table. This is so our Failed TS step falls in the correct column.
                        }
                    }
                    $errortext = $statmsg.Description.replace('"', '&quot;') #replace quotes so the html doesn't truncate
                    $varHash.Add('Exit Task Sequence', '<a href=" " title="' + $errortext + '"><img src="images/checks/redCheckMark_round.png" alt="Red Check Mark"></a>') #set pic to red and include error text
                    $ImageCompleted = $statmsg."Date / Time"
                    $LastLog = $statmsg."Date / Time"
                    $ImageDuration = Get-DateDifference -StartDate $ImageStarted -EndDate $ImageCompleted
                }
                Else {
                    #We don't care about it!
                }
                $VarCount += 1 #increase our variable count
            }
                    
        
    
        }
        #Build our HTML Table
        If ($tablecount -eq 1) {
            #This ensures that our html table headers are created only on the first pass through
            $table = '
                    <thead>
                        <tr class = "row100 head">
                            <th class="column100 column2" data-column="column2">Image Started</th>
                            <th class="column100 column3" data-column="column3">Image Completed</th>
                            <th class="column100 column4" data-column="column4">Image Duration</th>
                            <th class="column100 column5" data-column="column5">Last Log</th>
                            <th class="column100 column6" data-column="column6">Name During Imaging</th>'
            $column = 7 #hardcoded number. No need to change this as it is used for building the table
            ForEach ($item in $TSStepsNoDrivers) {
                $item = "$($stepscount)" + ' - ' + $item
                $String = '<th class="column100 column' + $Column + '" data-column="Column' + $Column + '">' + $item + '</th>'
                $table += $String
                $column += 1
                $stepscount += 1
            }
            $table += '<th class="column100 column' + $Column + '" data-column="Column' + $Column + '">' + "Exit Task Sequence" + '</th>'  #Manually add this as the last step
            $table += '</tr></thead><tbody>' #Manually add this to close out the headers and start the tbody
            $tablecount += 1
        }


        #Do some logic to convert our variables into date/time for the correct culture
        If ($ImageStarted) {
            $CultureDateTimeFormat = (Get-Culture).DateTimeFormat
            $DateFormat = $CultureDateTimeFormat.ShortDatePattern
            $TimeFormat = $CultureDateTimeFormat.LongTimePattern
            $DateTimeFormat = "$DateFormat $TimeFormat"
            #$ImageStarted = [DateTime]::ParseExact($ImageStarted.ToSTring(), $DateTimeFormat, [System.Globalization.DateTimeFormatInfo]::InvariantInfo, [System.Globalization.DateTimeStyles]::None)
            $ImageStarted = $ImageStarted | Get-Date -Format "dd.MM.yyyy HH:mm:ss"
        }
        If ($ImageCompleted) { 
            $CultureDateTimeFormat = (Get-Culture).DateTimeFormat
            $DateFormat = $CultureDateTimeFormat.ShortDatePattern
            $TimeFormat = $CultureDateTimeFormat.LongTimePattern
            $DateTimeFormat = "$DateFormat $TimeFormat"
            #$ImageCompleted = [DateTime]::ParseExact($ImageCompleted.ToSTring(), $DateTimeFormat, [System.Globalization.DateTimeFormatInfo]::InvariantInfo, [System.Globalization.DateTimeStyles]::None)
            $ImageCompleted = $ImageCompleted | Get-Date -Format "dd.MM.yyyy HH:mm:ss"
        }
        If ($LastLog) {
            $CultureDateTimeFormat = (Get-Culture).DateTimeFormat
            $DateFormat = $CultureDateTimeFormat.ShortDatePattern
            $TimeFormat = $CultureDateTimeFormat.LongTimePattern
            $DateTimeFormat = "$DateFormat $TimeFormat"
            #$LastLog = [DateTime]::ParseExact($LastLog.ToSTring(), $DateTimeFormat, [System.Globalization.DateTimeFormatInfo]::InvariantInfo, [System.Globalization.DateTimeStyles]::None)
            $LastLog = $LastLog | Get-Date -Format "dd.MM.yyyy HH:mm:ss"
        }


        #Here we process each row (computer data from results above) to the table
        $table += '<tr class="row100">
                <td class="column100 column2" data-column="column2">'+ $ImageStarted + '</td>
                <td class="column100 column3" data-column="column3">'+ $ImageCompleted + '</td>
                <td class="column100 column4" data-column="column4">'+ $ImageDuration + '</td>
                <td class="column100 column5" data-column="column5">'+ $LastLog + '</td>
                <td class="column100 column6" data-column="column6">'+ $NameDuringImaging + '</td>'
        $column = 7 #hardcoded number. No need to change this as it is used for building the table
        ForEach ($item in $varHash.GetEnumerator()) {
            $String = '<td class="column100 column' + $Column + '" data-column="Column' + $Column + '">' + $item.value
            $table += $String
            $table += "`n"
            $column += 1
        }
        $table += '</tr>'
        $table += "`n"
    }
}
    
#This if statement will allow for only relevant TS status messages 
If ($null -ne $ImageStarted) {
    #Build the array. The HTML variable is used in the $template file. 
    $html = $html += $table 
}

#These if statements will process if we had an issue or if no devices were found to be imaging.
If (($html.count -eq 0) -and ($messages.Name.Count -ge 1)) {
    $html += "<h2 style='text-align: center;'><strong>At least one device was detected as starting the Task Sequence but there is an issue sorting the steps.</strong></h2>"
}
elseif (($html.count -eq 0) -and ($messages.Name.Count -eq 0)) {
    $html += "<h2 style='text-align: center;'><strong>No devices have been detected as starting the Task Sequence.</strong></h2>"
}

#Get the template file
$template = (Get-Content -Path ($IISPath + "\template.html") -raw)

#Place variables and new $html into the template file and rename it as index.html
Invoke-Expression "@`"`r`n$template`r`n`"@" | Set-Content -Path ($IISPath + "\index_$($TaskSequenceID).html") 

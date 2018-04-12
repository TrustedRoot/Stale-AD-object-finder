Function Read-HostSpecial {
    [cmdletbinding(DefaultParameterSetName="_All")]
    Param(
    [Parameter(Position = 0,Mandatory,HelpMessage = "Enter prompt text.")]
    [Alias("message")]
    [ValidateNotNullorEmpty()]
    [string]$Prompt,
    [Alias("foregroundcolor","fg")]
    [consolecolor]$PromptColor,
    [string]$Title,
    [Parameter(ParameterSetName = "SecureString")]
    [switch]$AsSecureString,
    [Parameter(ParameterSetName = "NotNull")]
    [switch]$ValidateNotNull,
    [Parameter(ParameterSetName = "Range")]
    [ValidateNotNullorEmpty()]
    [int[]]$ValidateRange,
    [Parameter(ParameterSetName = "Pattern")]
    [ValidateNotNullorEmpty()]
    [regex]$ValidatePattern,
    [Parameter(ParameterSetName = "Set")]
    [ValidateNotNullorEmpty()]
    [string[]]$ValidateSet
    )
 
    Write-Verbose "Starting: $($MyInvocation.Mycommand)"
    Write-Verbose "Parameter set = $($PSCmdlet.ParameterSetName)"
    Write-Verbose "Bound parameters $($PSBoundParameters | Out-String)"
 
 
    #combine the Title (if specified) and prompt
    $Text = @"
    $(if ($Title) {
    "$Title`n$("-" * $Title.Length)"
    })
    $Prompt : 
"@
 
    #create a hashtable of parameters to splat to Write-Host
    $paramHash = @{
    NoNewLine = $True
    Object = $Text
    }
 
    if ($PromptColor) {
        $paramHash.Add("Foregroundcolor",$PromptColor)
    }
 
    #display the prompt
    #Write-Host @paramhash
    #get the value
    if ($AsSecureString) {
        $r = $host.ui.ReadLineAsSecureString()
    }
    else {
      #read console input
      [System.Reflection.Assembly]::LoadWithPartialName('Microsoft.VisualBasic') | Out-Null
      $r = [Microsoft.VisualBasic.Interaction]::InputBox($Prompt, $Title)
      #$r = $host.ui.ReadLine() 
    }
 
    #assume the input is valid unless proved otherwise
    $Valid = $True
 
    #run validation if necessary
    if ($ValidateNotNull) {
        Write-Verbose "Validating for null or empty"
        if($r.length -eq 0 -OR $r -notmatch "\S" -OR $r -eq $Null) {
            $Valid = $False
            Write-Error "Validation test for not null or empty failed."
        }
    }
    elseif ($ValidatePattern) {
        Write-Verbose "Validating for pattern $($validatepattern.ToString())"
        If ($r -notmatch $ValidatePattern) {
            $Valid = $False
            Write-Error "Validation test for the specified pattern failed."
        }
    }
    elseif ($ValidateRange) {
        Write-Verbose "Validating for range $($ValidateRange[0])..$($ValidateRange[1]) "
        if ( -NOT ([int]$r -ge $ValidateRange[0] -AND [int]$r -le $ValidateRange[1])) {
            $Valid = $False
            Write-Error "Validation test for the specified range ($($ValidateRange[0])..$($ValidateRange[1])) failed."
        }
        else {
             #convert to an integer
            [int]$r = $r 
        }
    }
    elseif ($ValidateSet) {
        Write-Verbose "Validating for set $($validateset -join ",")"
        if ($ValidateSet -notcontains $r) {
            $Valid = $False
            Write-Error "Validation test for set $($validateset -join ",") failed."
        }
    }
    If ($Valid) {
        Write-Verbose "Writing result to the pipeline"
        #any necessary validation passed
        $r
    }
    Write-Verbose "Ending: $($MyInvocation.Mycommand)"
    }
function Get-SaveAsPath{
    Add-Type -AssemblyName System.Windows.Forms
        $dlg = New-Object System.Windows.Forms.SaveFileDialog
        $dlg.Filter = "CSV (*.csv)|*.csv|Text (*.txt)|*.txt|Excel Worksheet (*.xls)|*.xls|All Files (*.*)|*.*"
        $dlg.SupportMultiDottedExtensions = $true;
        $dlg.InitialDirectory = "$env:HOMESHARE\Desktop";

    if($dlg.ShowDialog() -eq 'Ok'){
        return $($dlg.filename)
    }
}
#$Run = 'Yes'
Do{
    while(!$Searchbase){
        #Prompt user for OU searchbase, validate input is X.500 compliant.
        try{
            $Searchbase = Read-HostSpecial -Prompt "Enter OU searchbase" -Title "OU Searchbase" -ValidatePattern '(..=)(?<Name>.*?)(?<!\\),(?<Path>.*)' -PromptColor Green -ErrorAction Ignore
        }
        catch{
            Write-Host "ERROR : Searchbase must be the Distingushed Name of the OU (X.500 Directory Specification)." -ForegroundColor Red
            $Retry = Read-Host "Retry (Y/N)?"
            $Retry = $Retry.ToUpper()
            if($Retry -eq 'Y' -or $Retry -eq 'YES'){
                continue
            }
            else{
                exit
            }
        }
    }

    #Draw a progress bar for the user, as the script will appear to hang as it gathers all user objects
    Write-Progress -Id 1 -Activity "Gathering stale user information" -Status "Getting all user object properties" -PercentComplete 25

    #Get user objects in defined searchbase
    $UserObjects = Get-ADUser -Filter * -Searchbase $Searchbase -Properties *
    #Validate we found any objects. If not, exit.
    if($UserObjects.Count -eq 0){
        Write-Host No user objects found in $Searchbase
        exit
    }
    #Initalize the output array 
    $StaleObjects = @()
    $i = 0
    #Check each user to see if they meet our staleness criteria.
    foreach($User in $UserObjects){
        #Update the progress bar with the current status.
        $i++
        $Percent = ((($i / $UserObjects.Count)*75)+25)
        Write-Progress -Id 1 -Activity "Gathering stale user information" -Status "Processing user $i of $($UserObjects.Count)" -PercentComplete $Percent 
        Start-Sleep -Milliseconds 25
        #If the user has not changed their password in the last 180 days, let's collect some info on them.
        if($User.PasswordLastSet -le ((Get-Date).AddDays(-180))){
            if($User.PasswordLastSet){
                $PasswordLastSet = ($User.PasswordLastSet).toString("MM/dd/yyyy")
            }
            if($User.LastLogonDate){
                $LastLogonDate = ($User.LastLogonDate).toString("MM/dd/yyyy")
            }
            $UserInfo = @()
            $UserInfo += New-Object psobject -Property @{
                Name=$($User.Name)
                Username=$($User.sAMAccountName)
                Email=$($User.mail)
                Phone=$($User.OfficePhone)
                Office=$($User.Office)
                Department=$($User.Department)
                Manager=$($User.Manager)
                Enabled=$($User.Enabled)
                CreateDate=($($User.Created)).toString("MM/dd/yyyy")
                LastLogonDate=$($LastLogonDate)
                PasswordNeverExpires=$($User.PasswordNeverExpires)
                PasswordLastSet=$($PasswordLastSet)
                PasswordExpired=$($User.PasswordExpired)
            }
            $StaleObjects += $UserInfo
        }
    }

    Write-Progress -Id 1 -Activity "Gathering stale user information" -Completed
    $StaleCount = $StaleObjects.Count

    #Prompt the user to export the report
    $Export = [Microsoft.VisualBasic.Interaction]::MsgBox("Found $StaleCount stale objects. Export the report?", "YesNo", "Stale objects found")
    if($Export -eq 'Yes'){
        $Path = Get-SaveAsPath
        $StaleObjects | Select-Object "Name","Username","Email","Phone","Office","Department","Manager","Enabled","CreateDate","LastLogonDate","PasswordNeverExpires","PasswordLastSet","PasswordExpired" | Export-Csv $Path -NoTypeInformation
    }

    #Prompt the user to run again
    $Run = [Microsoft.VisualBasic.Interaction]::MsgBox("Would you like to search for stale objects in another OU?", "YesNo", "Rerun")
    if($Run -eq 'Yes'){
        $SearchBase = $null
        $StaleObjects = $null
    }
}
while ($Run -eq 'Yes')

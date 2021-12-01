# cache variable for global log data
$global:content = ""

function LogMsg {
    Param (
        [string] $Type,
        [string] $Text = "",
        [string] $User
    )

    $dateTime = Get-Date -f "dd.MM.yyyy HH:mm:ss"
    
    # if is missing parameter of new line it will be switched to default blank space
    if ($newLineCharacter.Length -eq 0) {
        $newLineCharacter = " "
    }

    # replacing new line in text message
    $text = $text.replace("`n", $newLineCharacter).replace("`r", $newLineCharacter)

    # add log message separator
    if ($messageSeparator.Length -gt 0) { 
        $messageSeparator = "`r`n$messageSeparator" 
    }
    
    return "$dateTime;$Type;$User;$Text$messageSeparator"
}

<#
    .Synopsis
        Log specific message with specific setting to file or Event viewer.

    .Example
        Log -Type "ERROR" -Message "error on line 10"

    .Example
        Log -Type "ERROR" -Message "error on line 10" -Verbose

    .Example
        Log -Message "added 10 accounts" -FilePath "C:\temp\logfile.log"

    .Example
        Log -Type "INFO" -Message "added 10 accounts" -LogEventVwr "SpecLog01"
    
    .Example
        Log -Type "INFO" -Message "added 10 accounts" -User "jnovak"

#>
function Log {
    [cmdletbinding()]
    Param (
        [ValidateSet("FATAL", "ERROR", "WARN", "INFO", "DEBUG")] 
        [string] $Type = "INFO",
        [string] $User = [Environment]::UserName,
        [Parameter(Mandatory=$true)]        
        [string] $Message,
        [string] $FilePath,
        [string] $LogEventVwr,

        # if the message contains special characters `n or `r will be switched to this character in the message for better output
        [string] $NewLineCharacter = "",
        # line separator of the one message
        [string] $MessageSeparator = ""
    )

    Process {
        try {
            # separator of local variable message (content)
            if ($global:content.Length -eq 0) { $sep = "" } else { $sep = "`r`n"}
        
            $content = $global:content

            $global:content = $(LogMsg $Type $Message $User)

            Write-Verbose $global:content

            $global:content += "$sep$content"

            # if filepath param is filled save the message into this file and clear global content message cache
            if ($FilePath) { 
                if (Test-Path $FilePath) { 
                    $contentLog = Get-Content $FilePath
                } 
                else { 
                    Write-Verbose "FilePath $FilePath doesn´t exist!" 
                }

                $global:content > $FilePath
                $contentLog >> $FilePath 

                $global:content = ""
            }

            # if logEventViewer is filled save the message into this event viewer name and clear global content message cache
            if ($LogEventVwr) {
                # check if event Log exists and if no exist so create it
                if (![System.Diagnostics.EventLog]::SourceExists($LogEventVwr)) {  
                    [System.Diagnostics.EventLog]::CreateEventSource($LogEventVwr, $LogEventVwr)

                    Write-Verbose "Creating new logName: $LogEventVwr"
                }  

                # write message into log
                $mylog = New-Object System.Diagnostics.Eventlog  
                $mylog.Source = $LogEventVwr
                $mylog.WriteEntry($global:content)

                $global:content = ""
            }
        }
        catch [System.Exception] {
            Write-Host "ERROR: $_"

            return $error.Count
        }
    }
}

#-----------------------------------------------------------------------------------------------------------------

Export-ModuleMember -Function @(
    'Log'
)

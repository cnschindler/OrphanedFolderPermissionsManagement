[cmdletbinding()]
Param(
[Parameter(Mandatory=$true)]
[Parameter(ParameterSetName="Detect")]
[System.IO.DirectoryInfo]$BasePath,
[Parameter(ParameterSetName="Remove")]
[switch]$RemoveOrphanedPermissions,
[Parameter(ParameterSetName="Remove")]
[System.IO.FileInfo]$AffectedMailboxesFile
)

. 'C:\Program Files\Microsoft\Exchange Server\V15\bin\RemoteExchange.ps1'
Connect-ExchangeServer -auto -ClientApplication:ManagementShell

$LogFolderName = "Remove-OrphanedPermissions"
[string]$LogPath = Join-Path -Path $BasePath -ChildPath $LogFolderName
[string]$LogfileFullPath = Join-Path -Path $LogPath -ChildPath ("RemoveOrphanedPermissions_{0:yyyyMMdd-HHmmss}.log" -f [DateTime]::Now)
$Script:NoLogging
[string]$CSVFullPath = Join-Path -Path $LogPath -ChildPath ("AffectedMailboxes_{0:yyyyMMdd-HHmmss}.txt" -f [DateTime]::Now)

function Write-LogFile
{
    # Logging function, used for progress and error logging...
    # Uses the globally (script scoped) configured LogfileFullPath variable to identify the logfile and NoLogging to disable it.
    #
    [CmdLetBinding()]

    param
    (
        [Parameter(Mandatory = $true)]
        [string]$Message,
        [System.Management.Automation.ErrorRecord]$ErrorInfo = $null
    )
    # Prefix the string to write with the current Date and Time, add error message if present...

    if ($ErrorInfo)
    {
        $logLine = "{0:d.M.y H:mm:ss} : [Error] : {1}: {2}" -f [DateTime]::Now, $Message, $ErrorInfo.Exception.Message
    }

    else
    {
        $logLine = "{0:d.M.y H:mm:ss} : [INFO] : {1}" -f [DateTime]::Now, $Message
    }

    if (!$Script:NoLogging)
    {
        # Create the Script:Logfile and folder structure if it doesn't exist
        if (-not (Test-Path $Script:LogfileFullPath -PathType Leaf))
        {
            New-Item -ItemType File -Path $Script:LogfileFullPath -Force -Confirm:$false -WhatIf:$false | Out-Null
            Add-Content -Value "Logging started." -Path $Script:LogfileFullPath -Encoding UTF8 -WhatIf:$false -Confirm:$false
        }

        # Write to the Script:Logfile
        Add-Content -Value $logLine -Path $Script:LogfileFullPath -Encoding UTF8 -WhatIf:$false -Confirm:$false
        Write-Verbose $logLine
    }
    else
    {
        Write-Host $logLine
    }
}

if ($AffectedMailboxesFile -like "*")
{
    $mbxs = Import-Csv -Path $MailboxFile
}

else
{
    $mbxs = Get-Mailbox -resultsize unlimited    
}

Set-Content -Value "EmailAddress" -Path $CSVFullPath

foreach ($mbx in $mbxs)
{
    $AffectedMBX = $false

    $Message = "$($mbx.Name): Processing Mailbox"
    Write-Host -ForegroundColor Green -Object $Message
    Write-LogFile -Message $Message

    $folders = Get-MailboxFolderStatistics -Identity $mbx | Where-Object Containerclass -like "IPF.*"
    $folders = $folders | where-object Folderpath -ne "/Top of Information Store"
    $folders = $folders | Where-Object Containerclass -ne "IPF.Configuration"
    $folders = $folders | Where-Object Containerclass -notlike "IPF.Contact.*"
    $folders = $folders | Where-Object Containerclass -ne "IPF.Note.OutlookHomepage"
    $folders = $folders | Where-Object Containerclass -ne "IPF.Note.SocialConnector.FeedItems"
    $Address = $mbx.WindowsEmailAddress.ToString()
    
        foreach ($folder in $folders)
    {
        $parentfolder = $folder.FolderPath.Split("/")[1]
        $fperms = Get-MailboxFolderPermission -Identity ($Address + ":" + $folder.FolderId)

        foreach ($fperm in $fperms)
        {
            if ($fperm.User.DisplayName -match "NT:S-1-5-")
            {
                $AffectedMBX = $true

                $Message = "$($mbx.Name): Found Permission for SID $($fperm.user) in folder $($folder.FolderPath.Replace('/','\'))."
                Write-Host -ForegroundColor Yellow -Object $Message
                Write-LogFile -Message $Message

                if ($RemoveOrphanedPermissions)
                {
                    $Message = "$($mbx.Name): Removing permissions for SID $($fperm.user)."
                    Write-Host -ForegroundColor Yellow -Object $Message
                    Write-LogFile -Message $Message

                    try
                    {
                        $Message = "$($mbx.Name): Successfully removed permission"
                        Remove-MailboxFolderPermission -Identity $fperm.Identity -User $fperm.User.DisplayName -Confirm:$false -ErrorAction Stop
                        Write-Host -ForegroundColor Yellow -Object $Message
                        Write-LogFile -Message $Message
                    }

                    Catch
                    {
                        $Message = "$($mbx.Name): Error removing permission."
                        Write-Host -ForegroundColor Red -Object "$($Message) $_"
                        Write-LogFile -Message $Message -ErrorInfo $_
                    }
                }
            }

            elseif ($fperm.User.Displayname -like "*Administrator*")
            {
                $AffectedMBX = $true

                $Message = "$($mbx.Name): Found permissons for Administrator Account in folder $($folder.FolderPath.Replace('/','\'))."
                Write-Host -ForegroundColor Yellow -Object $Message
                Write-LogFile -Message $Message

                if ($RemoveOrphanedPermissions)
                {
                    $Message = "$($mbx.Name): Removing permissions for Administrator account..."
                    Write-Host -ForegroundColor Yellow -Object $Message
                    Write-LogFile -Message $Message

                    try
                    {
                        $Message = "$($mbx.Name): Successfully removed permission"
                        Remove-MailboxFolderPermission -Identity $fperm.Identity -User $fperm.User.DisplayName -Confirm:$false -ErrorAction Stop
                        Write-Host -ForegroundColor Yellow -Object $Message
                        Write-LogFile -Message $Message
                    }

                    Catch
                    {
                        $Message = "$($mbx.Name): Error removing permission."
                        Write-Host -ForegroundColor Red -Object "$($Message) $_"
                        Write-LogFile -Message $Message -ErrorInfo $_
                    }
                }
            }
            
        }
    }

    if ($AffectedMBX -eq $true)
    {
        Add-Content -Value $Address -Path $CSVFullPath
    }
}

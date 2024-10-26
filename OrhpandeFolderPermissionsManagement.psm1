. 'C:\Program Files\Microsoft\Exchange Server\V15\bin\RemoteExchange.ps1'
Connect-ExchangeServer -auto -ClientApplication:ManagementShell

$Script:basepath = $env:TEMP
$Script:LogFolderName = "OrphanedFolderPermissionsManagement"
$Script:OutputFileNamePrefix = "AffectedFolders"
[string]$LogPath = Join-Path -Path $Script:basepath -ChildPath $Script:LogFolderName
[string]$LogfileFullPath = Join-Path -Path $Script:LogPath -ChildPath ($Script:LogFolderName + "_{0:yyyyMMdd-HHmmss}.log" -f [DateTime]::Now)
$Script:NoLogging
[string]$CSVFullPath = Join-Path -Path $Script:LogPath -ChildPath ($Script:OutputFileNamePrefix + "_{0:yyyyMMdd-HHmmss}.txt" -f [DateTime]::Now)

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

Function Get-OrphanedFolderPermissions
{
    [cmdletbinding()]
    Param()

    $InfoMessage = "Outputfilename is $($CSVFullPath)"
    Write-Host -ForegroundColor Green -Object $InfoMessage
    Write-LogFile -Message $InfoMessage

    # Retrieve all mailboxes
    Write-LogFile -Message "Retrieving all mailboxes"
    $mbxs = Get-Mailbox -resultsize unlimited

    # Create the output file and write the header
    Set-Content -Value "Mailbox,FolderPath,FolderID,UserID" -Path $CSVFullPath -Force

    foreach ($mbx in $mbxs)
    {
        $Message = "$($mbx.Name): Processing Mailbox"
        Write-Host -ForegroundColor Green -Object $Message
        Write-LogFile -Message $Message

        $folders = Get-MailboxFolderStatistics -Identity $mbx | Where-Object Containerclass -like "IPF.*"
        #$folders = $folders | where-object Folderpath -ne "/Top of Information Store"
        $folders = $folders | Where-Object Containerclass -ne "IPF.Configuration"
        #$folders = $folders | Where-Object Containerclass -notlike "IPF.Contact.*"
        $folders = $folders | Where-Object Containerclass -ne "IPF.Note.OutlookHomepage"
        $folders = $folders | Where-Object Containerclass -ne "IPF.Note.SocialConnector.FeedItems"
        $Address = $mbx.WindowsEmailAddress.ToString()
        
        foreach ($folder in $folders)
        {
            $fperms = Get-MailboxFolderPermission -Identity ($FullFolderID)
            $Folderpath = $folder.FolderPath.Replace('/','\')

            foreach ($fperm in $fperms)
            {
                $content = $Address + "," + $Folderpath + "," + $Folder.FolderId + "," + $fperm.user.DisplayName

                if ($fperm.User.DisplayName -match "NT:S-1-5-")
                {
                    $Message = "$($mbx.Name): Found Permission for SID $($fperm.user) in folder $($Folderpath)."
                    Write-Host -ForegroundColor Yellow -Object $Message
                    Write-LogFile -Message $Message
                    Add-Content -Path $CSVFullPath -Value $content
                }
                
                elseif ($fperm.User.Displayname -like "*Administrator*")
                {
                    $Message = "$($mbx.Name): Found permissons for Administrator Account in folder $($Folderpath)."
                    Write-Host -ForegroundColor Yellow -Object $Message
                    Write-LogFile -Message $Message
                    Add-Content -Path $CSVFullPath -Value $content
                }
            }
        }
    }
}

Function Remove-OrphanedFolderPermissions
{
    [cmdletbinding()]
    Param(
    [Parameter(Mandatory=$true)]
    [System.IO.FileInfo]$FileWithAffectedFolders
    )

    # Import file with mailboxes to cleanup
    $Folders = Import-Csv -Path $FileWithFoldersAffected

    foreach ($folder in $folders)
    {
        $Message = "$($folder.Mailbox): Processing Mailbox"
        Write-Host -ForegroundColor Green -Object $Message
        Write-LogFile -Message $Message
        $FullFolderID = $Folder.Mailbox + ":" + $Folder.FolderID

        try
        {
            $Message = "$($folder.Mailbox): Successfully removed permission entry $($folder.UserID)"
            Remove-MailboxFolderPermission -Identity $FullFolderID -User $folder.UserID -Confirm:$false -ErrorAction Stop
            Write-Host -ForegroundColor Yellow -Object $Message
            Write-LogFile -Message $Message
        }

        Catch
        {
            $Message = "$($folder.Mailbox): Error removing permission entry $($folder.UserID)."
            Write-Host -ForegroundColor Red -Object "$($Message) $_"
            Write-LogFile -Message $Message -ErrorInfo $_
        }
    }
}

Export-ModuleMember -Function Get-OrphanedFolderPermissions,Remove-OrphanedFolderPermissions



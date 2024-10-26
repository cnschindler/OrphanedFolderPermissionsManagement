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

function ConnectExchange
{
    # Check if a connection to an exchange server exists and connect if necessary...
    if (-NOT (Get-PSSession | Where-Object ConfigurationName -EQ "Microsoft.Exchange"))
    {
        $LogPrefix = "ConnectExchange"

        # Test if Exchange Management Shell Module is installed - if not, exit the script
        $EMSModuleFile = (Get-ItemProperty HKLM:\SOFTWARE\Microsoft\ExchangeServer\v15\Setup -ErrorAction SilentlyContinue).MsiInstallPath + "bin\RemoteExchange.ps1"
        
        # If the EMS Module wasn't found
        if (-Not (Test-Path $EMSModuleFile))
        {
            # Write Error end exit the script
            $ErrorMessage = "Exchange Management Shell Module not found on this computer. Please run this script on a computer with Exchange Management Tools installed!"
            Write-LogFile -LogPrefix $LogPrefix -Message $ErrorMessage
            Exit
        }

        else
        {
            # Load Exchange Management Shell
            try
            {
                # Dot source the EMS Script
                . $($EMSModuleFile) -ErrorAction Stop | Out-Null
                Write-LogFile -LogPrefix $LogPrefix -Message "Successfully loaded Exchange Management Shell Module."
            }

            catch
            {
                Write-LogFile -LogPrefix $LogPrefix -Message "Unable to load Exchange Management Shell Module." -ErrorInfo $_
                Exit
            }

            # Connect to Exchange Server
            try
            {
                Connect-ExchangeServer -auto -ClientApplication:ManagementShell -ErrorAction Stop | Out-Null
                Write-LogFile -LogPrefix $LogPrefix -Message "Successfully connected to Exchange Server."
            }

            catch
            {
                Write-LogFile -LogPrefix $LogPrefix -Message "Unable to connect to Exchange Server." -ErrorInfo $_
                Exit
            }
        }
    }
}

Function Get-OrphanedFolderPermissions
{
    [cmdletbinding()]
    Param()

    ConnectExchange

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

    ConnectExchange
    
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



#requires -version 2

<#
.SYNOPSIS

  Upload home folder data to OneDrive.

.DESCRIPTION

  The Upload-OneDriveFromCSV.ps1 script will upload home folder data to OneDrive
  based on a CVS file.

.NOTES
_________  ________  _________  ________       ___    ___       ________     
|\   ___ \|\   __  \|\___   ___\\   __  \     |\  \  |\  \     |\_____  \    
\ \  \_|\ \ \  \|\  \|___ \  \_\ \  \|\  \  __\_\  \_\_\  \____\|____|\ /_   
 \ \  \ \\ \ \   __  \   \ \  \ \ \   __  \|\____    ___    ____\    \|\  \  
  \ \  \_\\ \ \  \ \  \   \ \  \ \ \  \ \  \|___| \  \__|\  \___|   __\_\  \ 
   \ \_______\ \__\ \__\   \ \__\ \ \__\ \__\  __\_\  \_\_\  \_____|\_______\
    \|_______|\|__|\|__|    \|__|  \|__|\|__| |\____    ____   ____\|_______|
                                              \|___| \  \__|\  \___|         
                                                    \ \__\ \ \__\            
                                                     \|__|  \|__|            

  This program is free software; you can redistribute it and/or 
  modify it under the terms of the GNU General Public License as 
  published by the Free Software Foundation; either version 2 of 
  the License, or (at your option) any later version. 

  This program is distributed in the hope that it will be useful, 
  but WITHOUT ANY WARRANTY; without even the implied warranty of 
  MERCHANTABILITY or FITNESS FOR A PARTICULAR PURPOSE.  
  See the GNU General Public License for more details.  

  You should have received a copy of the GNU General Public License 
  along with this program; If not, see <https://www.gnu.org/licenses/>

  Version:        1.0
  Author:         Peter Szalatnay
  Creation Date:  2020/02/25
  Purpose/Change: Initial script development

.PARAMETER TenantName
  The SharePoint tenant name.

.PARAMETER SharePointAdminAccount
  An AzureAD account with SharePoint admin access.

.PARAMETER Password
  The password for the SharePointAdminAccount.

.PARAMETER CSV
  A .csv file with UserPrincipalNames and Homefolder locations.

.PARAMETER LastWriteTime
  Only sync files from last x days.

.EXAMPLE

  PS> .\Upload-OneDriveFromCSV.ps1 -TenantName tenant_name -SharepointAdminAccount `
  user@domain.com -Password password -CSV file_location.csv
#>

[CmdletBinding()]

Param (
    [Parameter()]
    [String] $TenantName = 'changeme',

    [Parameter()]
    [String] $SharePointAdminAccount = 'changeme',
        
    [Parameter(Mandatory=$true)]
    [String] $Password,

    [Parameter()]
    [String] $CSV = 'changeme',

    [Parameter(Mandatory=$true)]
    [Int] $LastWriteTime = 0
)

Begin
{
    $Logfile = $MyInvocation.MyCommand.Path -replace '\.ps1$', ".$(get-date -f yyyy-MM-dd).log"
    Start-Transcript -Path $Logfile -Append | Out-Null

    $PasswordAsSecure = ConvertTo-SecureString $Password -AsPlainText -Force
    $Credentials = New-Object System.Management.Automation.PSCredential ($SharePointAdminAccount, $PasswordAsSecure)
    
    function Check-IllegalCharacters ($Path, $OutputFile, [switch]$Fix)
    {
        $maxCharacters = 400
        $invalidFileNames = "desktop.ini", "_vti_"

        if (test-path $outputFile)
        {
            clear-content $outputFile
        }

        Add-Content $outputFile -Value 'File/Folder Name,New Name,Comments'

        $Items = Get-ChildItem -Path $Path -Recurse -ErrorVariable ErrVar -ErrorAction SilentlyContinue
        $ErrVar | ForEach-Object {
            Add-Content $outputFile (",,$($_.Exception.Message)")
        }

        $Items | ForEach-Object {
            $valid = $true
            $comments = New-Object System.Collections.ArrayList

            if ($_.PSIsContainer) { $type = "Folder" }
            else { $type = "File" }
     
            if ($_.Name.Length -gt $maxCharacters)
            {
                [void]$comments.Add("$($type) $($_.Name) is $($_.Name.Length) characters (max is $($maxCharacters)) and will need to be truncated")
                $valid = $false
            }

            if ($invalidFileNames.Contains($_.Name))
            {
                [void]$comments.Add("$($type) $($_.Name) is not a valid filename for file sync.")
                $valid = $false
            }
          
            # Technically all of the following are illegal \ / : * ? " < > | # % 
            # However, all but the last two are already invalid Windows Filename characters, so we don't have to worry about them
            $illegalChars = '[#%]'
            filter Matches($illegalChars)
            {
                $_.Name | Select-String -AllMatches $illegalChars |
                Select-Object -ExpandProperty Matches
                Select-Object -ExpandProperty Values
            }
           
            # Replace illegal characters with legal characters where found
            $newFileName = $_.Name
            Matches $illegalChars | ForEach-Object {
                [void]$comments.Add("Illegal string '$($_.Value)' found")
                if ($_.Value -match "#") { $newFileName = ($newFileName -replace "#", "-") }
                if ($_.Value -match "%") { $newFileName = ($newFileName -replace "%", "-") }
            }

            if ($comments.Count -gt 0)
            {
                Add-Content $outputFile ("$($_.FullName),$($_.FullName -replace $([regex]::escape($_.Name)), $($newFileName)),$($comments -join '; ')")  
                
                Write-Warning "$($type) $($_.FullName): $($comments -join ', ')"
            }
               
            # Fix file and folder names if found and the Fix switch is specified
            if (($newFileName -ne $_.Name) -and ($fix -and $valid))
            {
                Rename-Item $_.FullName -NewName ($newFileName)
            }
        }
    }
}

Process
{
    Write-Verbose "Connecting to SharePoint Online"
    Connect-SPOService -Url https://$TenantName-admin.sharepoint.com -Credential $Credentials
    $AdminConnection = Connect-PnPOnline -Url https://$TenantName-admin.sharepoint.com -Credential $Credentials -ReturnConnection

    Import-Csv $CSV | ForEach-Object {
        try 
        {
            Write-Verbose "Processing user '$($_.UserPrincipalName)' home folder '$($_.homefolder)'"
            $User = $_.UserPrincipalName
            $Folder = $_.HomeFolder

            if (-not (Test-Path $Folder))
            {
                Write-Warning "Home folder path does not exist : $Folder"
                return
            }

            $SiteUrl = (Get-PnPUserProfileProperty -Account $_.UserPrincipalName -Connection $AdminConnection).PersonalUrl

            if ($SiteUrl -notlike "https://$TenantName-my.sharepoint.com/personal/*")
            {
                Write-Warning "Invalide SharePoint url : $SiteUrl"
                return
            }

            Write-Verbose "Checking home folder '$Folder' for invalide files"
            Check-IllegalCharacters -Path $Folder -OutputFile "$Folder\_RenamedFiles.csv" -Fix

            Write-Verbose "Adding '$SharePointAdminAccount' as site collection admin on OneDrive site collections"
            Set-SPOUser -Site $SiteUrl -LoginName $SharePointAdminAccount -IsSiteCollectionAdmin $true | Out-Null

            Write-Verbose "Connecting to '$($_.UserPrincipalName)' via SharePoint PNP PowerShell module"
            Connect-PnPOnline -Url $SiteUrl -Credential $Credentials

            Write-Verbose "Uploading files to OneDrive, check log for uploaded files"

            if ($LastWriteTime)
            {
                $Files = Get-ChildItem -Recurse $Folder -ErrorAction SilentlyContinue | Where-Object {-not $_.PSIsContainer -and $_.LastWriteTime -gt (Get-Date).AddDays(-$LastWriteTime)} | Select FullName, @{ n = 'Folder'; e = { (Convert-Path $_.PSParentPath).Replace($Folder, '') } } 
            }
            else
            {
                $Files = Get-ChildItem -Recurse $Folder -ErrorAction SilentlyContinue | Where-Object {-not $_.PSIsContainer } | Select FullName, @{ n = 'Folder'; e = { (Convert-Path $_.PSParentPath).Replace($Folder, '') } }
            }

            $Files | ForEach-Object {
                Write-Information "Uploading file : $($_.FullName)"
                $File = Add-PnPFile -Path $_.FullName -Folder "Documents\HomeDrive\$($_.Folder)" -ErrorVariable ErrVar -ErrorAction SilentlyContinue
                if ($ErrVar) { Write-Warning "Cannot upload file '$($_.FullName)' : $($ErrVar.Exception.Message)" }
            }

            Write-Verbose "Removing '$SharePointAdminAccount' from OneDrive site collections"
            Set-SPOUser -Site $SiteUrl -LoginName $SharePointAdminAccount -IsSiteCollectionAdmin $false | Out-Null
        }
        catch 
        {
            Write-Error "Error : $($_.Exception.Message), check log for more information"
        }
    }
}

End
{
    Stop-Transcript | Out-Null
}

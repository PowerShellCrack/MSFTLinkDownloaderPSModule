Function Get-HrefMatches {
    [CmdletBinding()]
    param(
        ## The filename to parse
        [Parameter(Mandatory = $true)]
        [string] $content,

        ## The Regular Expression pattern with which to filter
        ## the returned URLs
        [string] $Pattern = "<\s*a\s*[^>]*?href\s*=\s*[`"']*([^`"'>]+)[^>]*?>"
    )

    $returnMatches = new-object System.Collections.ArrayList

    ## Match the regular expression against the content, and
    ## add all trimmed matches to our return list
    $resultingMatches = [Regex]::Matches($content, $Pattern, "IgnoreCase")
    foreach($match in $resultingMatches)
    {
        $cleanedMatch = $match.Groups[1].Value.Trim()
        [void] $returnMatches.Add($cleanedMatch)
    }

    $returnMatches
}

Function Get-Hyperlinks {
    [CmdletBinding()]
    param(
        [Parameter(Mandatory = $true)]
        [string] $content,
        [string] $Pattern = "<A[^>]*?HREF\s*=\s*""([^""]+)""[^>]*?>([\s\S]*?)<\/A>"
    )
    $resultingMatches = [Regex]::Matches($content, $Pattern, "IgnoreCase")

    $returnMatches = @()
    foreach($match in $resultingMatches){
        $LinkObjects = New-Object -TypeName PSObject
        $LinkObjects | Add-Member -Type NoteProperty `
            -Name Text -Value $match.Groups[2].Value.Trim()
        $LinkObjects | Add-Member -Type NoteProperty `
            -Name Href -Value $match.Groups[1].Value.Trim()

        $returnMatches += $LinkObjects
    }
    $returnMatches
}

Function Get-WebContentHeader{
    #https://stackoverflow.com/questions/41602754/get-website-metadata-such-as-title-description-from-given-url-using-powershell
    [CmdletBinding()]
    param(
        [Parameter(Mandatory=$true,Position=0,ValueFromPipeline=$true,ValueFromPipelineByPropertyName=$true)]
        #[Microsoft.PowerShell.Commands.HtmlWebResponseObject]$WebContent,
        $WebContent,

        [Parameter(Mandatory=$false)]
        [ValidateSet('Keywords','Description','Title')]
        [string]$Property
    )

    ## -------- PARSE TITLE, DESCRIPTION AND KEYWORDS ----------
    $resultTable = @{}
    # Get the title
    $resultTable.title = $WebContent.ParsedHtml.title
    # Get the HTML Tag
    $HtmlTag = $WebContent.ParsedHtml.childNodes | Where-Object {$_.nodename -eq 'HTML'}
    # Get the HEAD Tag
    $HeadTag = $HtmlTag.childNodes | Where-Object {$_.nodename -eq 'HEAD'}
    # Get the Meta Tags
    $MetaTags = $HeadTag.childNodes| Where-Object {$_.nodename -eq 'META'}
    # You can view these using $metaTags | select outerhtml | fl
    # Get the value on content from the meta tag having the attribute with the name keywords
    $resultTable.keywords = $metaTags  | Where-Object {$_.name -eq 'keywords'} | Select-Object -ExpandProperty content
    # Do the same for description
    $resultTable.description = $metaTags  | Where-Object {$_.name -eq 'description'} | Select-Object -ExpandProperty content
    # Return the table we have built as an object

    switch($Property){
        'Keywords'       {Return $resultTable.keywords}
        'Description'    {Return $resultTable.description}
        'Title'          {Return $resultTable.title}
        default          {Return $resultTable}
    }
}
Function Initialize-FileDownload {
    [CmdletBinding()]
    param(
         [Parameter(Mandatory=$false)]
         [Alias("Title")]
         [string]$Name,

         [Parameter(Mandatory=$true,Position=1)]
         [string]$Url,

         [Parameter(Mandatory=$true,Position=2)]
         [Alias("TargetDest")]
         [string]$TargetFile
     )
     Begin{
         ## Get the name of this function
         [string]${CmdletName} = $PSCmdlet.MyInvocation.MyCommand.Name

         ## Check running account
         [Security.Principal.WindowsIdentity]$CurrentProcessToken = [Security.Principal.WindowsIdentity]::GetCurrent()
         [Security.Principal.SecurityIdentifier]$CurrentProcessSID = $CurrentProcessToken.User
         [boolean]$IsLocalSystemAccount = $CurrentProcessSID.IsWellKnown([Security.Principal.WellKnownSidType]'LocalSystemSid')
         [boolean]$IsLocalServiceAccount = $CurrentProcessSID.IsWellKnown([Security.Principal.WellKnownSidType]'LocalServiceSid')
         [boolean]$IsNetworkServiceAccount = $CurrentProcessSID.IsWellKnown([Security.Principal.WellKnownSidType]'NetworkServiceSid')
         [boolean]$IsServiceAccount = [boolean]($CurrentProcessToken.Groups -contains [Security.Principal.SecurityIdentifier]'S-1-5-6')
         [boolean]$IsProcessUserInteractive = [Environment]::UserInteractive
     }
     Process
     {
         $ChildURLPath = $($url.split('/') | Select-Object -Last 1)

         $uri = New-Object "System.Uri" "$url"
         $request = [System.Net.HttpWebRequest]::Create($uri)
         $request.set_Timeout(15000) #15 second timeout
         $response = $request.GetResponse()
         $totalLength = [System.Math]::Floor($response.get_ContentLength()/1024)
         $responseStream = $response.GetResponseStream()
         $targetStream = New-Object -TypeName System.IO.FileStream -ArgumentList $targetFile, Create

         $buffer = new-object byte[] 10KB
         $count = $responseStream.Read($buffer,0,$buffer.length)
         $downloadedBytes = $count

         If($Name){$Label = $Name}Else{$Label = $ChildURLPath}

         Write-Verbose ("{0} : Initializing File Download from URL: {1}" -f ${CmdletName},$Url)

         while ($count -gt 0)
         {
             $targetStream.Write($buffer, 0, $count)
             $count = $responseStream.Read($buffer,0,$buffer.length)
             $downloadedBytes = $downloadedBytes + $count

             # display progress
             #  Check if script is running with no user session or is not interactive
             If ( ($IsProcessUserInteractive -eq $false) -or $IsLocalSystemAccount -or $IsLocalServiceAccount -or $IsNetworkServiceAccount -or $IsServiceAccount) {
                 # display nothing
                 #write-host "." -NoNewline
             }
             Else{
                 Write-Progress -Activity ("Downloading {0}" -f $Name) -Status ("Downloading: {0} ($([System.Math]::Floor($downloadedBytes/1024))K of $($totalLength)K): " -f $Label) -PercentComplete ( ([System.Math]::Floor($downloadedBytes/1024) / $totalLength) * 100 ) -id 1
             }
         }

         Start-Sleep 1

         $targetStream.Flush()
         $targetStream.Close()
         $targetStream.Dispose()
         $responseStream.Dispose()
    }
    End{
        #Write-Progress -activity "Finished downloading file '$($url.split('/') | Select-Object -Last 1)'"
        If($Name){$Label = $Name}Else{$Label = $ChildURLPath}
        Write-Progress -Activity ("Finished downloading file: {0}" -f $Label) -PercentComplete 100 -Completed
        #change meta in file from internet to allow to run on system
        If(Test-Path $TargetFile){Unblock-File $TargetFile -ErrorAction SilentlyContinue | Out-Null}
    }

 }


# MICROSOFT DOWNLOAD
#==================================================
function Get-MSFTLink {
    <#
        .SYNOPSIS
        Retrieves File from Microsoft

        .DESCRIPTION
        Download files from Microsoft download site using LinkID

        .NOTES
        Created by: @PowershellCrack

        .PARAMETER LinkID
        Required. Link id from download url

        .PARAMETER Filter
        Filter to reduce files found in link

        .PARAMETER Language
        Defaults to en-US. English.

        .PARAMETER Extract
        Attempts to extract zip files or extractable exe files

        .PARAMETER Cleanup
        Available with extract; removes archive after extraction

        .PARAMETER Force
        Re-downloads file even if it exists (overwrites)

        .PARAMETER NoDownload
        Export downloaded information as object

        .PARAMETER NoProgress
        Shows no progress during download (this is useful for large file sizes and speed)

        .PARAMETER Passthru
        Export downloaded information as object

        .EXAMPLE
        Get-MSFTLink -LinkID '49117' -DestPath C:\temp\Downloads -Force

        .EXAMPLE
        Get-MSFTLink -LinkID '55319' -DestPath C:\temp\Downloads -Filter 'LGPO'

        .EXAMPLE
        Get-MSFTLink -LinkID '49117' -DestPath C:\temp\Downloads -Force -Extract -Cleanup

        .EXAMPLE
        Get-MSFTLink -LinkID '55319' -DestPath C:\temp\Downloads -Passthru

        .EXAMPLE
        49117,55319,104223 | Get-MSFTLink -DestPath C:\temp\Downloads -Passthru

        .EXAMPLE
        49117,104223 | Get-MSFTLink -DestPath C:\temp\Downloads -Passthru -NoProgress -Extract -Cleanup

        .EXAMPLE
        '55319' | Get-MSFTLink -DestPath C:\temp\Downloads -Filter 'Windows Server' -Passthru -Extract -Verbose

        .LINK
        Get-HrefMatches
        Initialize-FileDownload
    #>
    [CmdletBinding()]
    param(
        [Parameter(Mandatory=$true,Position=1,ValueFromPipeline=$true,ValueFromPipelineByPropertyName=$true)]
        [string[]]$LinkID,

        [parameter(Mandatory=$false)]
        [string]$Filter,

        [ValidateSet('en-us','en-gb','en-sg','en-au')]
        [string]$Language = "en-us",

        [parameter(Mandatory=$true)]
        [string]$DestPath,

        [Parameter(Mandatory=$false,ParameterSetName='Archive')]
        [switch]$Extract,

        [Parameter(Mandatory=$false,ParameterSetName='Archive')]
        [switch]$Cleanup,

        [switch]$Force,

        [switch]$NoProgress,

        [switch]$Passthru
	)
    Begin{
        ## Get the name of this function
        [string]${CmdletName} = $PSCmdlet.MyInvocation.MyCommand.Name

        [Net.ServicePointManager]::SecurityProtocol = [Net.SecurityProtocolType]::Tls12

        [System.Uri]$SourceURL = "https://www.microsoft.com/$Language/download"
        [string]$DownloadURL = "https://download.microsoft.com/download"

        $DownloadData = @()
    }
    Process
    {

        Try{
            ## -------- FIND FILE LINKS ----------
            $ConfirmationLink = $SourceURL.OriginalString + "/confirmation.aspx?id=$LinkID"
            Write-Verbose ("{0} : Grabbing links from [{1}]..." -f ${CmdletName},$ConfirmationLink)

            $ConfirmationContent = Invoke-WebRequest $ConfirmationLink -UseBasicParsing -ErrorAction Stop
            $OfficialDownloads = Get-HrefMatches -content [string]$ConfirmationContent  | Where-Object {$_ -match $DownloadURL} | Select-Object -Unique

            #Filter
            If($Filter){
                $OfficialDownloads = $OfficialDownloads | Where-Object {$_ -like "*$Filter*"}
                Write-Verbose ("{0} : Found {1} official downloadable links with filter [{2}]" -f ${CmdletName},$OfficialDownloads.Count,$Filter)
            }Else{
                Write-Verbose ("{0} : Found {1} official downloadable links" -f ${CmdletName},$OfficialDownloads.Count)
            }

            ## -------- BUILD ROOT FOLDER ----------
            If( !(Test-Path $DestPath)){
                New-Item $DestPath -type directory -ErrorAction SilentlyContinue | Out-Null
            }

            #TESTS $link = $OfficialDownloads[0]
            Foreach($link in $OfficialDownloads)
            {
                #Build collection object
                $Data = '' | Select LinkID,DownloadableLink,FileName,FilePath,Extracted,Removed
                $Data.LinkID = $ConfirmationLink
                $Data.DownloadableLink = $link

                $Filename = $link | Split-Path -Leaf
                $destination = Join-Path $DestPath -ChildPath $Filename
                #collect data
                $Data.FileName = $Filename
                $Data.FilePath = $destination

                ## -------- DOWNLOAD  ----------
                If( (Test-Path $destination) -and !$Force){
                    Write-Verbose ("{0} : File already exists: [{1}]..." -f ${CmdletName},$Filename)
                    #Continue
                }
                Else{
                    Try{
                        Write-Verbose ("{0} : Attempting to download: [{1}]..." -f ${CmdletName},$Filename)
                        If($PSBoundParameters.ContainsKey('NoProgress') )
                        {
                            Invoke-WebRequest -Uri $link -OutFile $destination -UseBasicParsing -ErrorAction Stop
                        }
                        Else{
                            Initialize-FileDownload -Name ("{0}" -f $Filename) -Url $link -TargetDest $destination
                        }
                        Write-Verbose ("{0} : Successfully downloaded: {1}" -f ${CmdletName},$destination)
                    }
                    Catch {
                        Write-Error ("{0} : Failed downloading [{1}]: {2}" -f ${CmdletName},$Filter,$_.Exception.Message)
                    }
                }

                ## -------- EXTRACT ----------
                If($PSBoundParameters.ContainsKey('Extract'))
                {
                    $File = Split-path $destination -Leaf
                    Try{
                        Write-Verbose ("{0} : Attempting to Extract file [{1}] to [{2}]" -f ${CmdletName},$destination,$DestPath)
                        If([System.IO.Path]::GetExtension($File) -eq '.zip'){
                            Expand-Archive -LiteralPath "$destination" -DestinationPath $DestPath -Force -ErrorAction Stop
                        }
                        #Assume if executable and extract is used; its an extractable file
                        If([System.IO.Path]::GetExtension($File) -eq '.exe'){
                            Start-Process -FilePath $destination -ArgumentList "/extract:$DestPath /quiet" -Wait -ErrorAction Stop
                        }
                        $Data.Extracted = $True
                    }
                    catch {
                        $Data.Extracted = $False
                        Write-Error ("{0} : Unable to download [{1}]. {2}" -f ${CmdletName},$Filter,$_.Exception.Message)
                    }
                }

                ## -------- REMOVE ARCHIVE ----------
                If($PSBoundParameters.ContainsKey('Cleanup') -and $Data.Extracted)
                {
                    Write-Verbose ("{0} : Removing file [{0}]" -f ${CmdletName},$destination,${CmdletName})
                    Remove-Item $destination -Force -ErrorAction Stop | Out-Null
                    $Data.Removed = $True
                }Else{
                    $Data.Removed = $False
                }

                #add data to array
                $DownloadData += $Data
            } #end loop
        }
        catch {
            Write-Error ("{0} : Unable to download [{1}]. {2}" -f ${CmdletName},$Filter,$_.Exception.Message)
        }

    }
    End{
        If($PSBoundParameters.ContainsKey('Passthru')){
            return $DownloadData
        }
    }
}

$exportModuleMemberParams = @{
    Function = @(
        'Get-MSFTLink'
    )
}

Export-ModuleMember @exportModuleMemberParams

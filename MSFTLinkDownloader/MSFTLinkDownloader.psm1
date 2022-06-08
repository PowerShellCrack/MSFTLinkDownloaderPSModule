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

         [Parameter(Mandatory=$true,Position=0)]
         [string]$Url,

         [Parameter(Mandatory=$true,Position=1)]
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
        Write-Progress -Activity ("Finished downloading file: {0}" -f $Label) -Completed
        #change meta in file from internet to allow to run on system
        If(Test-Path $TargetFile){Unblock-File $TargetFile -ErrorAction SilentlyContinue | Out-Null}
    }

}

function Get-ZipFileSize {

    param (
        [ValidateScript({Get-item $_ -Include '*.zip'})]
        $Path
    )
    $ZipSize = (Get-item $path).length/1kb

    #open zip using explorer to unzip
    $shell = New-Object -ComObject shell.application
    $zip = $shell.NameSpace($Path)
    $size = 0
    foreach ($item in $zip.items()) {
        if ($item.IsFolder) {
            $size += Get-UncompressedZipFileSize -Path $item.Path
        } else {
            $size += $item.size
        }
    }

    # It might be a good idea to dispose the COM object now explicitly, see comments below
    [System.Runtime.InteropServices.Marshal]::ReleaseComObject([System.__ComObject]$shell) | Out-Null
    [System.GC]::Collect()
    [System.GC]::WaitForPendingFinalizers()

    $Zipdata = '' | Select ZippedKbSize,UnzippedKbSize
    $Zipdata.ZippedKbSize = ("{0:F2}" -f $ZipSize)
    $Zipdata.UnzippedKbSize = ("{0:F2}" -f ($size/1kb))

    return $Zipdata
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

        .EXAMPLE
        Get-MSFTLink -LinkID '49117','104223'

        .EXAMPLE
        Get-MSFTLink -LinkID '55319' -Filter 'LGPO'

        .EXAMPLE
        49117,55319,104223 | Get-MSFTLink -Language en-gb

        .LINK
        Get-HrefMatches
    #>
    [CmdletBinding()]
    [OutputType([PSCustomObject])]
    param(
        [Parameter(Mandatory=$true,Position=0,ValueFromPipeline=$true,ValueFromPipelineByPropertyName=$true)]
        [int[]]$LinkID,

        [parameter(Mandatory=$false,Position=1)]
        [string]$Filter,

        [ValidateSet('en-us','en-gb','en-sg','en-au')]
        [string]$Language = "en-us"
	)
    Begin{
        ## Get the name of this function
        [string]${CmdletName} = $PSCmdlet.MyInvocation.MyCommand.Name

        [Net.ServicePointManager]::SecurityProtocol = [Net.SecurityProtocolType]::Tls12

        [System.Uri]$SourceURL = "https://www.microsoft.com/$Language/download"
        [string]$DownloadURL = "https://download.microsoft.com/download"

        $LinkCollection = @()
    }
    Process
    {
        Foreach($ID in $LinkID){
            Try{
                ## -------- FIND FILE LINKS ----------
                $ConfirmationLink = $SourceURL.OriginalString + "/confirmation.aspx?id=$ID"
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

                #TESTS $link = $OfficialDownloads[0]
                Foreach($link in $OfficialDownloads)
                {
                    #Build collection object
                    $Data = '' | Select LinkID,DownloadLink,FileName
                    $Data.LinkID = $ConfirmationLink
                    $Data.DownloadLink = $link

                    $Filename = $link | Split-Path -Leaf
                    $Data.FileName = $Filename

                    #add data to array
                    $LinkCollection += $Data
                } #end loop
            }
            catch {
                Write-Error ("{0} : Unable to download [{1}]. {2}" -f ${CmdletName},$Filter,$_.Exception.Message)
            }
        }
    }
    End{
        return $LinkCollection
    }
}

Function Invoke-MsftLinkDownload {
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

        .PARAMETER DownloadLink
        Required. download url )usually obtained by Get-MSFTLink

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
        Invoke-MsftLinkDownload -LinkID 49117 -DestPath C:\temp\Downloads -Force

        .EXAMPLE
        Invoke-MsftLinkDownload -LinkID 55319,104223 -Filter 'Server' -DestPath C:\temp\Downloads -Force -verbose

        .EXAMPLE
        Invoke-MsftLinkDownload -DownloadLink 'https://download.microsoft.com/download/8/5/C/85C25433-A1B0-4FFA-9429-7E023E7DA8D8/LGPO.zip' -DestPath C:\temp\Downloads -Force

        .EXAMPLE
        Invoke-MsftLinkDownload -DownloadLink 'https://download.microsoft.com/download/8/5/C/85C25433-A1B0-4FFA-9429-7E023E7DA8D8/LGPO.zip' -DestPath C:\temp\Downloads -Force -Extract -Cleanup

        .EXAMPLE
        Get-MSFTLink -LinkID 49117,104223 | Invoke-MsftLinkDownload -DestPath C:\temp\Downloads -Passthru

        .EXAMPLE
        $Links = Get-MSFTLink -LinkID 49117,55319,104223
        $Links | Invoke-MsftLinkDownload -DestPath C:\temp\Downloads -Passthru -NoProgress -Extract -Cleanup -verbose

        .LINK
        Get-MSFTLink
        Initialize-FileDownload
        Get-UncompressedZipFileSize
    #>
    [CmdletBinding(DefaultParameterSetName='URL')]
    param(
        [Parameter(Mandatory=$true,Position=0,ParameterSetName='ID')]
        [int[]]$LinkID,

        [Parameter(Mandatory=$false,ParameterSetName='ID')]
        [string]$Filter,

        [Parameter(Mandatory=$false, ParameterSetName='ID')]
        [ValidateSet('en-us','en-gb','en-sg','en-au')]
        [string]$Language = "en-us",

        [Parameter(Mandatory=$true,
                    Position=0,
                    ValueFromPipeline=$true,
                    ValueFromPipelineByPropertyName=$true,
                    ParameterSetName='URL'
        )]
        [string[]]$DownloadLink,

        [Parameter(Mandatory=$true,ParameterSetName='ID')]
        [Parameter(Mandatory=$true,ParameterSetName='URL')]
        [string]$DestPath,

        [Parameter(Mandatory=$false,ParameterSetName='ID')]
        [Parameter(Mandatory=$false,ParameterSetName='URL')]
        [switch]$Extract,

        [Parameter(Mandatory=$false,ParameterSetName='ID')]
        [Parameter(Mandatory=$false,ParameterSetName='URL')]
        [switch]$Cleanup,

        [switch]$Force,

        [switch]$NoProgress,

        [switch]$Passthru
	)
    Begin{
        ## Get the name of this function
        [string]${CmdletName} = $PSCmdlet.MyInvocation.MyCommand.Name

        If($PsCmdlet.ParameterSetName -eq 'ID'){
            $MSFTLinkParams = @{}
            If($LinkID){
               $MSFTLinkParams += @{LinkID = $LinkID}
            }
            If($Filter){
                $MSFTLinkParams += @{Filter = $Filter}
            }
            If($Language){
                $MSFTLinkParams += @{Language = $Language}
            }
            Write-Verbose ("{0} : Attempting to get links: [{1}]..." -f ${CmdletName},($LinkID -join ','))
            $DownloadLink = Get-MSFTLink @MSFTLinkParams | Select -ExpandProperty DownloadLink
        }

        $DownloadCollection = @()

        ## -------- BUILD ROOT FOLDER ----------
        If( !(Test-Path $DestPath)){
            New-Item $DestPath -type directory -ErrorAction SilentlyContinue | Out-Null
        }

    }
    Process
    {
        Write-Verbose ("{0} : Processing download link: [{1}]..." -f ${CmdletName},$DownloadLink)
        #TESTS $Link = $DownloadLink[0]
        Foreach($Link in $DownloadLink)
        {
            #Build collection object
            $Data = '' | Select DownloadURL,FileName,FilePath,Downloaded,Extracted,Removed
            $Data.DownloadURL = $Link

            $Filename = $Link | Split-Path -Leaf
            $destination = Join-Path $DestPath -ChildPath $Filename
            #collect data
            $Data.FileName = $Filename
            $Data.FilePath = $destination

            ## -------- DOWNLOAD  ----------
            $Data.Downloaded = $False
            If( (Test-Path $destination) -and !$Force){
                Write-Verbose ("{0} : File already exists: [{1}]..." -f ${CmdletName},$Filename)
                $Data.Downloaded = $True
                #Continue
            }
            Else{
                Try{
                    Write-Verbose ("{0} : Attempting to download: [{1}]..." -f ${CmdletName},$Filename)
                    If($PSBoundParameters.ContainsKey('NoProgress') )
                    {
                        Invoke-WebRequest -Uri $Data.DownloadURL -OutFile $destination -UseBasicParsing -ErrorAction Stop
                    }
                    Else{
                        Initialize-FileDownload -Name ("{0}" -f $Filename) -Url $Data.DownloadURL -TargetDest $destination
                    }
                    Write-Verbose ("{0} : Successfully downloaded: {1}" -f ${CmdletName},$destination)
                    $Data.Downloaded = $True
                }
                Catch {
                    Write-Error ("{0} : Failed downloading [{1}]: {2}" -f ${CmdletName},$Filter,$_.Exception.Message)
                }
            }


            ## -------- EXTRACT ----------
            $Data.Extracted = $False
            If($PSBoundParameters.ContainsKey('Extract'))
            {
                $File = Split-path $destination -Leaf
                Try{
                    Write-Verbose ("{0} : Attempting to Extract file [{1}] to [{2}]" -f ${CmdletName},$destination,$DestPath)
                    If([System.IO.Path]::GetExtension($File) -eq '.zip'){
                        Expand-Archive -LiteralPath "$destination" -DestinationPath $DestPath -Force -ErrorAction Stop
                        $Data.Extracted = $True
                    }
                    #Assume if executable and extract is used; its an extractable file
                    If([System.IO.Path]::GetExtension($File) -eq '.exe'){
                        $result = Start-Process -FilePath $destination -ArgumentList "/extract:$DestPath /quiet" -Wait -ErrorAction Stop -PassThru
                        If($result.ExitCode -eq 0){$Data.Extracted = $True}
                    }
                }
                catch {
                    Write-Error ("{0} : Unable to download [{1}]. {2}" -f ${CmdletName},$Filter,$_.Exception.Message)
                }
            }

            ## -------- REMOVE ARCHIVE ----------
            $Data.Removed = $False
            If($PSBoundParameters.ContainsKey('Cleanup') -and $Data.Extracted)
            {
                Write-Verbose ("{0} : Removing file [{0}]" -f ${CmdletName},$destination,${CmdletName})
                Remove-Item $destination -Force -ErrorAction SilentlyContinue | Out-Null
                $Data.Removed = $True
            }

            #add data to array
            $DownloadCollection += $Data
        } #end loop
    }
    End{
        If($PSBoundParameters.ContainsKey('Passthru')){
            return $DownloadCollection
        }
    }
}

$exportModuleMemberParams = @{
    Function = @(
        'Get-MSFTLink'
        'Invoke-MsftLinkDownload'
    )
}

Export-ModuleMember @exportModuleMemberParams

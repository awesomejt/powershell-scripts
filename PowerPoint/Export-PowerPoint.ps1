# Powershell script to export Powerpoint all slides as PNG or JPG images using the Powerpoint COM API

function Export-Deck
{
    [CmdletBinding()]
    param (
        [Parameter()]
        [string]
        $InputFile,

        [Parameter()]
        [string]
        $OutputFolder = "Exports",

        [Parameter()]
        [string]
        $OutputSubFolder,

        [Parameter()]
        [ValidateSet("JPG", "PNG")]
        [string]
        $OutputType = "PNG",

        [Parameter()]
        [ValidateSet("HD", "QHD", "4K")]
        [string]
        $OutputSize = "HD"
    )

    "Input File: $InputFile"
    "Output Folder: $OutputFolder"
    "Output SubFolder: $OutputSubFolder"
    "Output Type: $OutputType"
    "Output Size: $OutputSize"

    $currDir = (Get-Location).Path
    "Current Directory: $currDir"

    if ($OutputFolder -eq "Exports") {
        if ($OutputSubFolder) {
            $exportFolder = Join-Path $currDir (Join-Path $OutputFolder $OutputSubFolder)
        } else {
            $exportFolder = Join-Path $currDir $OutputFolder
        }
    } else {
        if ($OutputSubFolder) {
            $exportFolder = Join-Path $OutputFolder $OutputSubFolder
        } else {
            $exportFolder = $OutputFolder
        }
    }
    "Export Directory: $exportFolder"
    if (!(Test-Path -Path $exportFolder)) {
        New-Item -Path $exportFolder -ItemType Directory
    }

	# Load Powerpoint Interop Assembly
	[Reflection.Assembly]::LoadWithPartialname("Microsoft.Office.Interop.Powerpoint") > $null
	[Reflection.Assembly]::LoadWithPartialname("Office") > $null

	$msoFalse =  [Microsoft.Office.Core.MsoTristate]::msoFalse
	$msoTrue =  [Microsoft.Office.Core.MsoTristate]::msoTrue

	# start Powerpoint
	$application = New-Object "Microsoft.Office.Interop.Powerpoint.ApplicationClass" 

	# Make sure inputFile is an absolte path
	$InputFile = Resolve-Path $InputFile
   
	$presentation = $application.Presentations.Open($InputFile, $msoTrue, $msoFalse, $msoFalse)

    # default is HD / 1080p
    $outputWidth = 1920
    $outputHeigth = 1080
    if ($OutputSize -eq "QHD") {
        # 1440p
        $outputWidth = 2560
        $outputHeigth = 1440
    } elseif ($OutputSize -eq "4K") {
        # 4K / 2160p
        $outputWidth = 3840
        $outputHeigth = 2160
    }

    $presentation.Export($exportFolder, $OutputType, $outputWidth, $outputHeigth)
	
	$presentation.Close()
	$presentation = $null
	
	if($application.Windows.Count -eq 0)
	{
		$application.Quit()
	}
	
	$application = $null
	
	# Make sure references to COM objects are released, otherwise powerpoint might not close
	# (calling the methods twice is intentional, see https://msdn.microsoft.com/en-us/library/aa679807(office.11).aspx#officeinteroperabilitych2_part2_gc)
	[System.GC]::Collect();
	[System.GC]::WaitForPendingFinalizers();
	[System.GC]::Collect();
	[System.GC]::WaitForPendingFinalizers();       
}
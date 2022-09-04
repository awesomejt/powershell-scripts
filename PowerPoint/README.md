# PowerPoint Scripts

## Export Deck

This script will export all the slides within the specified PowerPoint deck into either JPG or PNG format at HD, QHD or 4K resolution (assuming 16:9 slide aspect).

Requirements:

* Recent version of MS PowerPoint (365 or v2019+) on Windows 10
* PowerPoint Deck ready for export in 16:9 aspect ratio
* PowerShell environment

How to run.

First load/run the script file: 

```powershell
# Note the space between the first period and the second one
. .\path\to\Export-Powerpoint.ps1
```

Call the Export-Deck function:

```powershell
# export deck accepting all defaults
Export-Deck -InputFile .\Path\To\Deck.pptx

# export deck into PNG format to Exports/Pings folder in 4K resolution
Export-Deck -InputFile .\Path\To\Deck.pptx -OutputSubFolder Pings -OutputSize 4K

# export deck into JPEG format to a specific export folder in 1440p (QHD) resolution
Export-Deck -InputFile .\Path\To\Deck.pptx -OutputFolder .\Path\To\Exports -OutputType JPG -OutputSize QHD
```

Parameters:

* __InputFile__: Relative or Absolute path to PowerPoint presentation to be exported
* __OutputFolder__: Relative or Absolute path to export. Default is especial "Exports" directory.
* __OutputSubFolder__: Sub-folder name for Output. Default is "Images". Default export location would be the "Exports/Images" folder within the current directory.
* __OutputType__: Image type for export. Valid options: JPG or PNG. Default is PNG.
* __OutputSize__: Image resolution size for export. Valid options: HD, QHD, or 4K. Default is HD.


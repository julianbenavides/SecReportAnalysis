param (
	[Parameter( Mandatory=$true)]
	[string]$initialDirectory
)

#Script to get the path where the org report should be created
Function RequestFilename()
{
    param($initialDirectory)

    [System.Reflection.Assembly]::LoadWithPartialName("System.windows.forms") | Out-Null
    $FolderBrowser = New-Object System.Windows.Forms.FolderBrowserDialog
    $FolderBrowser.SelectedPath=$initialDirectory
    #$FolderBrowser.RootFolder = [Environment+SpecialFolder]$initialDirectory
    $FolderBrowser.ShowDialog()| Out-Null
    $FolderBrowser.SelectedPath
}
RequestFilename $initialDirectory

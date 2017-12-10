##
## Import Commonly Used Functions
##
Import-Module ".\CommonFunctions.ps1"

<#
.SYNOPSIS
    The Get-ExcelSheetsWithoutUnderscore function returns an array containing the names of the Excel workbook sheets that doe NOT start with an underscore.

.DESCRIPTION
    This function reads in an Excel file and returns a hash table with:
        1) The hash keys being the names of the sheets of the Excel workbook
        2) The Values being an array of hash tables, where each has table is a row of the given sheet.  These hash tables have the column headers as their keys, and their respective values as their value.
    The name of the Excel file is passed into the function as a parameter.

    $ArrayVar = Get-ExcelSheetsWithoutUnderscore("ExcelFile.xls")

.NOTES
    Date           : October 17, 2017
    Filename       : LabFunctions.ps1
    Author         : Luis Chanu
    Prerequisite   : ImportExcel PowerShell module from Doug Finke (https://www.powershellgallery.com/packages/ImportExcel)
    Copyright 2017 by Luis Chanu

.LINK
    TBD

.EXAMPLE
    Example 1 to be filled in

.EXAMPLE
    Example 2 to be filled in
#>

function Get-ExcelSheetsWithoutUnderscore {
    param(
        [Parameter(Mandatory=$true)]
        [string[]]$Filename
    )

    # Initialize empty array
    $sheets = @()

    # Read Excel sheets, and remove any sheets from the pipline which begin with an underscore
    $sheets = Get-ExcelSheetInfo $filename | where-object {$_.Name -notlike "_*"} | Select-Object -ExpandProperty Name

    return $sheets
}


<#
.SYNOPSIS
    The Read-ExcelDataToHashTable function reads the contents of a given Excel file, and returns it's contents in a hash table

.DESCRIPTION
    This function reads in an Excel file and returns a hash table with:
        1) The hash keys being the names of the sheets of the Excel workbook
        2) The Values being an array of hash tables, where each has table is a row of the given sheet.  These hash tables have the column headers as their keys, and their respective values as their value.
    The name of the Excel file is passed into the function as a parameter.

    $HashTableVar = Read-ExcelDataToHashTable("ExcelFile.xls")

.NOTES
    Date           : October 17, 2017
    Filename       : LabFunctions.ps1
    Author         : Luis Chanu
    Prerequisite   : ImportExcel PowerShell module from Doug Finke (https://www.powershellgallery.com/packages/ImportExcel)
    Copyright 2017 by Luis Chanu

.LINK
    TBD

.EXAMPLE
    Example 1 to be filled in

.EXAMPLE
    Example 2 to be filled in
#>

function Read-ExcelDataToHashTable {
    param(
        [Parameter(Mandatory=$true)]
            [string]$Filename
    )

    # Obtain Workbook Sheet names
    $Sheets = @()
    $Sheets = Get-ExcelSheetsWithoutUnderscore -Filename $Filename

    # Data obtained from the various workbook sheets stored here.
    $WorkbookData = @{}

    # Iterate through all the Excel Sheets
    $Sheets | ForEach-Object {

        # Get Sheet Name
        $Sheet = $_

        # Read data from sheet
        $Data = @{}
        $Data = Import-Excel -Path (".\"+$Filename) -WorksheetName $Sheet

        # Add new data to WorkbookData hash table
        $WorkbookData.Add($Sheet, $Data)
    }

    return $WorkbookData
}



function ChooseSoftware {
    param(
        [Parameter(Mandatory=$True)]
        [String]$DialogBoxTitle,
        [Parameter(Mandatory=$True)]
        [String]$DialogBoxPrompt,
        [Parameter(Mandatory=$True)]
        [Array]$Products
    )

    # Working Script:   https://docs.microsoft.com/en-us/powershell/scripting/getting-started/cookbooks/selecting-items-from-a-list-box?view=powershell-5.1


    Add-Type -AssemblyName System.Windows.Forms
    Add-Type -AssemblyName System.Drawing

    $form = New-Object System.Windows.Forms.Form
    $form.Text = $DialogBoxTitle
    $form.Size = New-Object System.Drawing.Size(300,200)
    $form.StartPosition = "CenterScreen"

    $OKButton = New-Object System.Windows.Forms.Button
    $OKButton.Location = New-Object System.Drawing.Point(75,120)
    $OKButton.Size = New-Object System.Drawing.Size(75,23)
    $OKButton.Text = "OK"
    $OKButton.DialogResult = [System.Windows.Forms.DialogResult]::OK
    $form.AcceptButton = $OKButton
    $form.Controls.Add($OKButton)

    $CancelButton = New-Object System.Windows.Forms.Button
    $CancelButton.Location = New-Object System.Drawing.Point(150,120)
    $CancelButton.Size = New-Object System.Drawing.Size(75,23)
    $CancelButton.Text = "Cancel"
    $CancelButton.DialogResult = [System.Windows.Forms.DialogResult]::Cancel
    $form.CancelButton = $CancelButton
    $form.Controls.Add($CancelButton)

    $label = New-Object System.Windows.Forms.Label
    $label.Location = New-Object System.Drawing.Point(10,20)
    $label.Size = New-Object System.Drawing.Size(280,20)
    $label.Text = $DialogBoxPrompt
    $form.Controls.Add($label)

    $ListBox = New-Object System.Windows.Forms.ListBox
    $ListBox.Location = New-Object System.Drawing.Point(10,40)
    $ListBox.Size = New-Object System.Drawing.Size(260,20)
    $ListBox.Height = 80

    # Add the various choices to choose from to the Listbox
    $Products.GetEnumerator() | Sort-Object -Property FullName | ForEach-Object {
        [void] $ListBox.Items.Add($_.FullName)
    }

    $form.Controls.Add($ListBox)

    $form.Topmost = $True

    $result = $form.ShowDialog()

    # If user pressed OK, see what they selected
    if ($result -eq [System.Windows.Forms.DialogResult]::OK)
    {
        $Selection = $ListBox.SelectedItem

        # Iterate through the choices to see which product the user selected, and return it
        $Products.GetEnumerator() | ForEach-Object {
            if ($_.FullName -eq $Selection)
            {
                return $_
            }
        }
    }
    else
    {
        # If OK wasn't selected, then return Null
        $Selection = $null
    }
}



function Create-Switches {
    param(
        [Parameter(Mandatory=$True)]
        $Switches
    )

    # Iterate through all the switches to see if they need to be created
    $Switches | ForEach-Object {
        # Verify we really have a switch
        if ($_ -ne $Null)
        {
            # Get vCenter where this vDS should be created
            $vCenter = Get-VCSAConnection -vcsaName $_.vCenterServer

            # If existing vCenter connection exists
            if ($vCenter -ne $Null)
            {
                # Get DataCenter where switch goes
                $DataCenter = Get-Datacenter -Name $_.DataCenter -Server $vCenter

                # Create Switch
                $Name = $_.Name
                Write-Log -Message "Creating switch $Name"

                # New vDS configuration options
                $Options = @{
                    Name            = $_.Name
                    Server          = $vCenter
                    Location        = $DataCenter
                    MaxPorts        = $_.MaxPorts
                    NumUplinkPorts  = $_.NumUplinks
                    Mtu             = $_.MTU
                    Version         = $vCenter.Version
                    LinkDiscoveryProtocol           = $_.DiscoveryProtocol_Type
                    LinkDiscoveryProtocolOperation  = $_.DiscoveryProtocol_Operation
                }

                # Create New vDS Switch
                New-VDSwitch @Options
            }
        }
    }
}


##################################################################################################


# Global Variables
$DateTime = Get-Date -UFormat "%Y%m%d_%H%M%S"
$VerboseLogFile = ".\Logs\NSX-Lab-Build_LogFile_$DateTime.LOG"
$ExcelFilename  = ".\LabEnvironment.xlsx"






Write-Log -Message "Read in Excel file"
$ExcelData = @{}
$ExcelData = Read-ExcelDataToHashTable -Filename "$ExcelFilename"

# Initialize SoftwareToInstall
$SoftwareToInstall = @{}

# Have User Select vCenter Version To Install
Write-Log -Message "Prompt user for which vCenter Server version to use"
$Choices = $ExcelData.Get_Item("Software") | Where-Object -Property Family -eq "vCenter"
$Choice = ChooseSoftware -DialogBoxTitle "vCenter Server Software" -DialogBoxPrompt "Please select which vCenter Server version to use:" -Products $Choices

if ($Choice -eq $null)
{
    Write-Log -Warning -Message "vCenter Server required, but none selected"
    exit
}
else
{
    Write-Log -Message "vCenter Selected: $Choice.FullName"
    $SoftwareToInstall.Add("vCenter", $Choice)
}


# Have User Select ESXi Version To Install
Write-Log -Message "Prompt user for which ESXi version to use"
$Choices = $ExcelData.Get_Item("Software") | Where-Object -Property Family -eq "ESXi"
$Choice = ChooseSoftware -DialogBoxTitle "ESXi Software" -DialogBoxPrompt "Please select which ESXi version to use:" -Products $Choices

if ($Choice -eq $null)
{
    Write-Log -Warning -Message "ESXi required, but none selected"
    exit
}
else
{
    Write-Log -Message "ESXi Selected: $Choice.FullName"
    $SoftwareToInstall.Add("ESXi", $Choice)
}


# Have User Select NSX-v Version To Install
Write-Log -Message "Prompt user for which NSX-v version to use"
$Choices = $ExcelData.Get_Item("Software") | Where-Object -Property Family -eq "NSX-v"
$Choice = ChooseSoftware -DialogBoxTitle "NSX Software" -DialogBoxPrompt "Please select which NSX-v version to install:" -Products $Choices

if ($Choice -eq $null)
{
    Write-Log -Message "No NSX-v software selected"
}
else
{
    Write-Log -Message "NSX-v Selected: $Choice.FullName"
    $SoftwareToInstall.Add("NSX-v", $Choice)
}


#####################################################################################################
##                                Prepare Physical Environment                                     ##
#####################################################################################################


# Get all Outter vCenter servers
Write-Log -Message "Getting list of Outter vCenter Servers"
$vCenters = $ExcelData.Get_Item("Outter_Management")

# Iterate through each of them
$vCenters | ForEach-Object {
    # Get Information To Authenticate
    $vCenter  = $_.vCenterServer
    $User     = $_.vCenterAdmin
    $Password = $_.vCenterPassword

    # If vCenter contains a vCenter
    If ($vCenter -ne $Null)
    {
        # Establish connection to vCenter server, which adds it to the global VIServers table
        Write-Log -Message "Establishing connection to vCenter Server $vCenter"
        Get-VCSAConnection -vcsaName $vCenter -vcsaUser $User -vcsaPassword $Password
    }

}



# Get all Outter vDS Switches
Write-Log -Message "Getting list of Outter vDS Switches"
$Switches = $ExcelData.Get_Item("Outter_Switches")

# Create all Outter vDS Switches
Create-Switches($Switches)


Write-Host "Done"


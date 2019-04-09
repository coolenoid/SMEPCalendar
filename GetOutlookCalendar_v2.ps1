Function New-SQLQuery {
     <#
     .SYNOPSIS
        Returns data from a SQL query.
     .DESCRIPTION
        Returns data from a SQL query. Assumes integrated authentication.
     .EXAMPLE
        New-SQLQuery -Server Server1 -Instance LYNC -Database lis -Query 'Select * from lis'
     #>
    [CmdletBinding(SupportsShouldProcess = $True)] 
    param (
        [string]$Server,
        [string]$Instance = '',
        [string]$Database,
        [string]$Query
    )
 
    #Define SQL Connection String
    [string]$ServerAndInstance = $Server
    if ($Instance -ne '')
    {
        [string]$ServerAndInstance = "$ServerAndInstance\$Instance"
    }
    [string]$connstring = "server=$ServerAndInstance;database=$Database;trusted_connection=true;"
 
    #Define SQL Command
    [object]$command = New-Object System.Data.SqlClient.SqlCommand
    $command.CommandText = $Query
 
    # Define SQL connection
    [object]$connection = New-Object System.Data.SqlClient.SqlConnection
    $connection.ConnectionString = $connstring
    $connection.Open()
    $command.Connection = $connection
    
    # Create SQL data adapter and associate the query with it
    [object]$sqladapter = New-Object System.Data.SqlClient.SqlDataAdapter
    $sqladapter.SelectCommand = $command
 
    # Execute query
    [object]$results = New-Object System.Data.Dataset
    $recordcount = $sqladapter.Fill($results)
    $connection.Close()
    return $Results.Tables[0]
}

Function Get-OutlookCalendar {
<#
 .Synopsis
  This function returns items from Outlook Calendar.  You can specify the mailbox and folder.
  By: Michael Wong (modified source by Zachary Loeber)

  .OlDefaultFolders
  https://docs.microsoft.com/en-us/office/vba/api/outlook.oldefaultfolders

   .EXAMPLE
  Get-OutlookCalendar | Out-GridView

#>
 [CmdLetBinding()]
  param(
    [string]$Mailbox,
    [string]$FunctionGroup,
    [string]$FolderName = '',
    [switch]$ListMailbox,
    [switch]$UnreadOnly,
    [switch]$FilterByYear,
    [switch]$Recurse,
    [string[]]$Properties = @()
  )
  begin {
        function Release-Ref ($ref) {
            [System.Runtime.InteropServices.Marshal]::ReleaseComObject([System.__ComObject]$ref) | out-null
            [System.GC]::Collect()
            [System.GC]::WaitForPendingFinalizers()
        }

        <#$props = @("body", "header", "returnpath", "spf", "SenderId",  "antispam", "scl", 
                "pcl", "senderserver", "senderIP", "*", "subject")
        $headerProps = @("header", "returnpath", "spf", "SenderId",  "antispam", "scl", 
                "pcl", "senderserver", "senderIP", "*")#>

        $props = @("header","Recipients", "*")

        # validate properties
        if($properties) {
            foreach($prop in $properties) {
                if($props -notcontains $prop) { throw "$prop is not a valid property" }
            }
        }

        function getHeader($item) { 
            $headerScheme = "http://schemas.microsoft.com/mapi/proptag/0x007D001E"
            return $item.propertyaccessor.getproperty($headerscheme) 
        }

         function getRecipients($item) {
            # property schema
            $recipientsScheme = "http://schemas.microsoft.com/mapi/proptag/0x39FE001E"
            $recips = $item.Recipients
            $str = ''
            if ($recips) {
                foreach($recip in $recips) {
                $str += $recip.PropertyAccessor.getproperty($recipientsScheme) + "|" ;
                }
                return $(if ($str.length -gt 30000) { $str.substring(0, 30000) } else { $str })
            
            } else {
                return $str
            }
        }
    
        function Get-OutlookSubFolder($FolderSource) {
            foreach ($Folder in $FolderSource.Folders) {
                $Folder
                Get-OutlookSubFolder($Folder)
            }
        }

        function getCalendar($email, $mailboxowner, $mailboxresolvedname, $functiongroup, $foldername, $folderpath) {
            $calstuff = new-object pscustomobject
            $caldata = @{}
            $caldata.MailBoxOwner = $mailboxowner
            $caldata.FunctionGroup = $functiongroup
            $caldata.MailBoxResolvedName = $mailboxresolvedname.Name
            $caldata.Start = $email.Start -replace "(\w+),? (\w+) (\d+)\w+,? (\d+)", "$1 $3 $2 $4"
            $caldata.End = $email.End -replace "(\w+),? (\w+) (\d+)\w+,? (\d+)", "$1 $3 $2 $4"
            $caldata.Duration = $email.Duration
            $caldata.Categories = $email.Categories
            $caldata.Subject = $email.Subject
            $caldata.Location = $email.Location
            $caldata.IsRecurring = $email.IsRecurring
            $caldata.Organizer = $email.Organizer
            $caldata.Body = $(if ($email.Body.length -gt 30000) { $email.Body.substring(0, 30000) } else { $email.Body })
            $caldata.FolderName = $foldername
            $caldata.FolderPath = $folderpath
            $caldata.Recipients = getRecipients($email)
            
            return New-Object psobject -Property $caldata

        }

        Function Get-DefaultDateTime {
            # Extract the default Date/Time formatting from the local computer's "Culture" settings, and then create the format to use when parsing the date/time information.
            $CultureDateTimeFormat = (Get-Culture).DateTimeFormat
            $DateFormat = $CultureDateTimeFormat.ShortDatePattern
            $TimeFormat = $CultureDateTimeFormat.ShortTimePattern
            $DateTimeFormat = "$DateFormat $TimeFormat"
            return $DateTimeFormat
        }

    }
    Process {
        # load the required .NET types
        Add-type -assembly "Microsoft.Office.Interop.Outlook" | out-null
        # access Outlook object model
        $outlook = new-object -comobject outlook.application
        # connect to the namespace
        $mapi = $outlook.GetNameSpace("MAPI")
        if(!$mapi) { throw "Unable to create MAPI to Outlook.  Be sure Microsoft Office is installed" }

        if($ListMailbox) {
            $mapi.folders | %{ $_.fullfolderpath.trim("\\") }
            return
        }

        if(!$mailbox) {
            if ($FolderName -eq '') {
                $olFolders = "Microsoft.Office.Interop.Outlook.olDefaultFolders" -as [type]
                $FolderSource = $mapi.getDefaultFolder($olFolders::olFolderInBox)
                #$FolderSource = @($Mapi.Folders)
            }
            else {
                try {
                    #$FolderSource = $Mapi.folders.item($FolderName)
                    $olFolders = "Microsoft.Office.Interop.Outlook.olDefaultFolders" -as [type]
                    $FolderSource = $mapi.getDefaultFolder($olFolders::$FolderName)
                }
                catch {
                    throw "Can't access folder $FolderName"
                }
            }
        }
        else {
            $main = $mapi.CreateRecipient($mailbox)
            #$main = $mapi.folders.item($mailbox)

            if(!$main) {
                throw "Can't access $mailbox.  Use -ListMailbox to get valid mailboxes"
            }

            if ($FolderName -eq '') {
                #pull all folders
                #$FolderSource = $mapi.folders.item($mailbox).Folders
                $FolderSource = $mapi.folders.item($mailbox).Folders
            }
            else {
                try {
                    $olFolders = "Microsoft.Office.Interop.Outlook.olDefaultFolders" -as [type]
                    if($main.Resolve()){
                        $FolderSource = $mapi.GetSharedDefaultFolder($main, $olFolders::$FolderName)
                    }
                }
                catch {
                    throw "Can't access folder $foldername"
                }
            }
        }
        $Folders = @()
        $Folders += $FolderSource
        $Folders += if ($Recurse) { @(Get-OutlookSubFolder($FolderSource.Folders)) }
        
        Foreach ($Folder in $Folders) {
            if ($UnreadOnly) {
                Write-Verbose "Get-Outlook: Retreiving unread items from folder $($Folder.Name)"
                #$folder.items.Restrict("[UnRead] = True") | %{ getEmail $_ ($Folder).Name ($Folder).FolderPath }
                $folder.items.Restrict("[UnRead] = True") | %{ getCalendar $_ $Mailbox $main $FunctionGroup ($Folder).Name ($Folder).FolderPath }
            }
            elseif ($FilterByYear) {
                Write-Verbose "Get-Outlook: Retreiving date filtered items from folder $($Folder.Name)"
                # Change Date Filter parameters &  Trigger default Date/Time formatting
                $datefilter = "1/1/2018 12:00AM"
                $getDTFormat = Get-DefaultDateTime
                
                $folder.items.Restrict("[Start] >= '"+ (Get-Date $datefilter -Format $getDTFormat) +"'") | %{ getCalendar $_ $Mailbox $main $FunctionGroup ($Folder).Name ($Folder).FolderPath }
                #$folder.items.Restrict("[Start] >= '" + (Get-Date $datefilter -Format "MM/dd/yyyy hh:mm AMPM") + "'") | %{ getCalendar $_ $Mailbox $main $FunctionGroup ($Folder).Name ($Folder).FolderPath }
            }
            else {
                Write-Verbose "Get-Outlook: Retreiving read items from folder $($Folder.Name)"
                $folder.items | %{ getCalendar $_ $Mailbox $main $FunctionGroup ($Folder).Name ($Folder).FolderPath }
            }
        }
        Release-Ref $outlook
    }

}

Function Upload-FileWebDAV {
<#
 .Synopsis
  This function forward the XMl files generated to storage destination using WebDAV
  By: Michael Wong
  
  .EXAMPLE
  Upload-FileWebDAV -File "C:\Apps\Michael\Projects\SMEP Structured Calendar\test.xml" -URL "\\60.53.88.25\WebDAV\DavWWWRoot"

#>
[CmdletBinding(SupportsShouldProcess = $True)] 
    param (
        [string]$File,
        [string]$URL
    )

$uri = $URL
Remove-PSDrive WebDavShare -Force -ea 0
New-PSDrive -Name WebDavShare -PSProvider FileSystem -Root $uri | Out-Null
try {
    Get-ChildItem "WebDavShare:\" Select-Object -First 1 | Out-Null
    } catch {
    Remove-PSDrive WebDavShare -Force -ea 0
    $C = Get-Credential
    New-PSDrive -Name WebDavShare -Credential $C -PSProvider FileSystem -Root $uri | Out-Null
        }
Copy-Item $File -Destination $uri -force
Remove-PSDrive WebDavShare -Force -ea 0
}

Function Get-Inputs {
Add-Type -AssemblyName System.Windows.Forms
Add-Type -AssemblyName System.Drawing

$form = New-Object System.Windows.Forms.Form
$form.Text = 'SMEP Structured Calendar v01'
$form.Size = New-Object System.Drawing.Size(300,200)
$form.StartPosition = 'CenterScreen'

$OKButton = New-Object System.Windows.Forms.Button
$OKButton.Location = New-Object System.Drawing.Point(75,120)
$OKButton.Size = New-Object System.Drawing.Size(75,23)
$OKButton.Text = 'OK'
$OKButton.DialogResult = [System.Windows.Forms.DialogResult]::OK
$form.AcceptButton = $OKButton
$form.Controls.Add($OKButton)

$CancelButton = New-Object System.Windows.Forms.Button
$CancelButton.Location = New-Object System.Drawing.Point(150,120)
$CancelButton.Size = New-Object System.Drawing.Size(75,23)
$CancelButton.Text = 'Cancel'
$CancelButton.DialogResult = [System.Windows.Forms.DialogResult]::Cancel
$form.CancelButton = $CancelButton
$form.Controls.Add($CancelButton)

$label = New-Object System.Windows.Forms.Label
$label.Location = New-Object System.Drawing.Point(10,20)
$label.Size = New-Object System.Drawing.Size(280,20)
$label.Text = 'Please enter the candidate Shell Email Address:'
$form.Controls.Add($label)

$textBox = New-Object System.Windows.Forms.TextBox
$textBox.Location = New-Object System.Drawing.Point(10,40)
$textBox.Size = New-Object System.Drawing.Size(260,20)
$form.Controls.Add($textBox)

#Dropdown
[array]$DropDownArray = "SWK", "SBH", "CP", "PE", "DEV", "WELLS", "BE", "HR" ,"FIN" , "IMIT", "HSSE", "COMM", "NOV", "ER", "VP", "EXP", "GR", "LGL"

$DropDownLabel = new-object System.Windows.Forms.Label
$DropDownLabel.Location = new-object System.Drawing.Size(10,70) 
$DropDownLabel.size = new-object System.Drawing.Size(280,20) 
$DropDownLabel.Text = "Select SMEP Function Group:"
$Form.Controls.Add($DropDownLabel)

$DropDown = new-object System.Windows.Forms.ComboBox
$DropDown.Location = new-object System.Drawing.Size(10,90)
$DropDown.Size = new-object System.Drawing.Size(130,30)

ForEach ($Item in $DropDownArray) {
    [void] $DropDown.Items.Add($Item)
}
$DropDown.SelectedItem = $DropDown.Items[0]
$Form.Controls.Add($DropDown)


$form.Topmost = $true

$form.Add_Shown({$textBox.Select()})
$result = $form.ShowDialog()

if ($result -eq [System.Windows.Forms.DialogResult]::OK)
{
    $x = @{}
    $x.email = $textBox.Text
    $x.functiongroup = $DropDown.SelectedItem.ToString()
    #return $x
    return New-Object psobject -Property $x
}
}




<# Main () - Execution #>
# get Inputs & define file location
$EmailID = Get-Inputs
$FileLocation = "C:\Apps\" + $EmailID.functiongroup + "_" + ($EmailID.email -split "@")[0] + ".csv"

# Check file exist prior to cleanup
if (Test-Path $FileLocation) {
    Remove-Item $FileLocation
}


#Get-OutlookCalendar in CSV
#Get-OutlookCalendar $EmailID.email $EmailID.functiongroup olFolderCalendar | Export-Csv -NoTypeInformation $FileLocation

#Get-OutlookCalendar in CSV, filtered by year
Get-OutlookCalendar $EmailID.email $EmailID.functiongroup olFolderCalendar -FilterByYear | Export-Csv -NoTypeInformation $FileLocation

#Get-OutlookCalendar in Table-Grid view
#Get-OutlookCalendar $EmailID.email $EmailID.functiongroup olFolderCalendar | Out-GridView

#Get-OutlookCalendar in XML
#Get-OutlookCalendar $EmailID.email $EmailID.functiongroup olFolderCalendar | ConvertTo-Xml -as String -NoTypeInformation | Set-Content -Path $FileLocation

#File Export & CleanUp
Upload-FileWebDAV -File $FileLocation -URL "\\60.53.88.25\WebDAV\DavWWWRoot"
Remove-Item $FileLocation












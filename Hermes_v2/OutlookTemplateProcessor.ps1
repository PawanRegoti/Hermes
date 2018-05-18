. $PSScriptRoot\SharePointFunctions.ps1

$loadMessageForm = [System.Reflection.Assembly]::LoadWithPartialName(“System.Windows.Forms”)

#----Main-Function-------------------------------------------------------------------------------------------------

function Create-Reminders()
{
    [CmdletBinding()]
    param(
    [DeliveryItem[]] $deliveryItems,
    [string] $account)

    $outlookApp = Get-OutlookApplication    
    $resourceLocation = "$PSScriptRoot\Resource\templates.xml"

    if($outlookApp -ne $null)
    {                     
        #Processing each delivery item
        foreach($deliveryItem in $deliveryItems)
        {
            if($deliveryItem.ReleaseType -contains 'Hotfix')
            { $releaseType = 'Hotfix' }
            else
            { $releaseType = 'Prototype' }

            if(!(Test-Path -Path $resourceLocation))
            {
                Show-Message "Unable to find template.xml at $($resourceLocation)" -messageBoxIcon ([Windows.Forms.MessageBoxIcon]::Error)
            }

            <#$templateGroups = Fetch-TemplateFromXml $resourceLocation
           
            foreach($templateGroup in $templateGroups)
            {
                if($templateGroup.Name -eq 'Common' -or $templateGroup.Name -eq $releaseType)
                {
                    foreach($template in $templateGroup.Group)
                    {
                        $valueDate = $deliveryItem.CodeFreeze.Date.Add(([System.Timespan]::Parse($template.start.value)))
                        Create-Meeting $valueDate $template $outlookApp $account
                    }                    
                }
            } #>
            
            Fetch-TemplateFromXml $resourceLocation | ?{$_.Name -eq 'Common' -or $_.Name -eq $releaseType} | %{ $_.Group | %{ Create-Meeting $deliveryItem.CodeFreeze.Date.Add(([System.Timespan]::Parse($_.start.value))) $_ $outlookApp $account }}       
        }
    }
}

#-------------------------------------------------------------------------------------------------------------------

#------Helper-Functions---------------------------------------------------------------------------------------------

#--Creates Meeting Requests--#
function Create-Meeting
{
    [CmdletBinding()]
    param(
    [DateTime] $startDate,
    [object] $template,
    [object] $outlookApplication,
    [string] $account)

    $folder = Get-Folder "\\$($account)" "IPM.Appointment" $outlookApplication.Session.Folders
    $items = $outlookApplication.GetNamespace("MAPI").GetFolderFromID($folder.EntryID).Items;

    $meeting = $items.Add([Microsoft.Office.Interop.Outlook.OlItemType]::olAppointmentItem) 
    $meeting.MeetingStatus = [Microsoft.Office.Interop.Outlook.OlMeetingStatus]::olMeeting
    
    $meeting.Subject = [string]$template.subject.value
    $meeting.Subject += " ($($deliveryItem.Platform)-$($deliveryItem.N16Version))"

    $meeting.Body = [string]$template.body
    $meeting.Location = [string]$template.location.value
    $meeting.Importance = 1 
    $meeting.ResponseRequested = $template.response.value
    $meeting.BusyStatus = [Microsoft.Office.Interop.Outlook.OlBusyStatus]::olFree

    if ($template.duration.value -eq -1)
        {
          $meeting.Duration = 0;

          $meeting.Start = $startDate
          $meeting.End = $startDate.AddDays(1)

          $meeting.AllDayEvent = $true;
        }
        else
        {
          $meeting.Start = $startDate
          $meeting.Duration = $template.duration.value
          $meeting.End = $startDate.AddMinutes($template.duration.value)
        }

        if ($template.remainder.value -gt -1)
        {
          $meeting.ReminderMinutesBeforeStart = $template.remainder.value * 60; #Convert to minutes from hours
        }
        else
        {
          $meeting.ReminderSet = $false;
        }

    if([string]::IsNullOrWhitespace($template.recipients.value))
    {
        $meeting.Recipients.Add('MANGLER')
    }
    else
    {
        foreach($recipient in $template.recipients.value -split ';')
        {
            if(!([string]::IsNullOrWhitespace($recipient)))
            {
            $meeting.Recipients.Add($recipient.Trim())  
            } 
        }
    }
    $meeting.Display()
}

#--Fetches Reminder Template from Resource/Template.xml--#
function Fetch-TemplateFromXml
{
    [CmdletBinding()]
    param([string] $path)

    [xml]$xmlDocument = Get-Content -Path $path

    return $xmlDocument.templates.template | Group-Object -Property releaseType
}

#--Display's message in popup format--#
function Show-Message
{
    [CmdletBinding()]
    param(
    [string] $message,
    [Windows.Forms.MessageBoxButtons] $messageButtons = [Windows.Forms.MessageBoxButtons]::OK,
    [Windows.Forms.MessageBoxIcon] $messageBoxIcon = [Windows.Forms.MessageBoxIcon]::Information)
    
    [Windows.Forms.MessageBox]::Show($message, "Hermes", $messageButtons, $messageBoxIcon) > $null

    Write-Host $message -ForegroundColor Yellow
}

#--Get Appropriate folder from outlook folders--#
function Get-Folder
{
    [CmdletBinding()]
    [OutputType([object])]
    Param(
    [string] $fullFolderPath,
    [string] $defaultMessageClass,
    [object] $rootFolders)
      
    foreach ($folder in $rootFolders)
      {
        if ($folder.FullFolderPath.StartsWith($fullFolderPath, [System.StringComparison]::OrdinalIgnoreCase) -and $folder.DefaultMessageClass.Equals($defaultMessageClass,[System.StringComparison]::OrdinalIgnoreCase))
        {
          return $folder;
        }

        $result = Get-Folder $fullFolderPath $defaultMessageClass $folder.Folders

        if ($result -ne $null)
         { return $result;}
      }
      return $null;
}

#------------------------------------------------------------------------------------------------------------------------------------------

#--Commom-Function----(Also used by Hermes.ps1 to fill account drop down)------------------------------------------------------------------

function Get-OutlookApplication
{
    #Fetch existing or create a new Outlook object.
    $outlookApp = $null
    $outlookProcess = Get-Process outlook -ErrorAction SilentlyContinue
    if($outlookProcess -eq $null)
    {
        $outlookApp = New-Object -ComObject 'Outlook.Application'
    }
    else
    {
            try
            {
                $outlookApp = [System.Runtime.InteropServices.Marshal]::GetActiveObject('Outlook.Application') 
            }
            catch
            {
                #Powershell might be in administrator mode and outlook is not.
                Show-Message “Please open/close Outlook and try again.” -messageBoxIcon ([Windows.Forms.MessageBoxIcon]::Warning) 
            }
    }

    return $outlookApp
}

#--------------------------------------------------------------------------------------------------------------------------------------------
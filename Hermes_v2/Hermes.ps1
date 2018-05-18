. $PSScriptRoot\SharePointFunctions.ps1
. $PSScriptRoot\OutlookTemplateProcessor.ps1

#region Import the Assemblies 
[reflection.assembly]::loadwithpartialname("System.Windows.Forms") | Out-Null 
[reflection.assembly]::loadwithpartialname("System.Drawing") | Out-Null 
#endregion

#Hermes Form Function 
function Hermes { 

#region Objects 
$HermesForm = New-Object System.Windows.Forms.Form 
$FetchDeliveriesButton = New-Object System.Windows.Forms.Button 
$GenerateRemindersButton = New-Object System.Windows.Forms.Button 
$DeliveryList = New-Object System.Windows.Forms.ListBox 
$InitialFormWindowState = New-Object System.Windows.Forms.FormWindowState 
$AccountDropDown = new-object System.Windows.Forms.ComboBox

$sizeObject = New-Object System.Drawing.Size 
$locationPoint = New-Object System.Drawing.Point 
#endregion Objects
#---------------------------------------------- 

$DeliveryDict = @{}


#Event Script Blocks 
#---------------------------------------------- 
#Provide Custom Code for events specified in PrimalForms. 
$handler_FetchDeliveriesButton_Click= 
{
    #Fetch Deliveries
    $DeliveryDict.Clear()
    Get-UnbuiltDeliveries | ? { $_.CodeFreeze -ne $null} | % { $DeliveryDict.Add("[$($_.CodeFreeze.ToShortDateString())]--($($_.Platform)-$($_.N16Version))--$($_.DevBranch)",$_) }
    Write-Host "Deliveries fetched: "$DeliveryDict.Count 
    $DeliveryList.Items.Clear()
    $DeliveryList.Items.AddRange(($DeliveryDict.Keys | Sort-Object -Descending))
}

$handler_GenerateRemindersButton_Click= 
{
    #Fetch Deliveries
    $deliveriesSelected = $DeliveryList.SelectedItems
    Write-Host "Generating Reminders for "$deliveriesSelected.Count" deliveries."
    Create-Reminders ($deliveriesSelected.ForEach({$DeliveryDict[$_]})) $AccountDropDown.SelectedItem
}

$OnLoadForm_StateCorrection= 
{
    #Correct the initial state of the form to prevent the .Net maximized form issue 
    $HermesForm.WindowState = $InitialFormWindowState 

    #Assign icon if present.
    $iconPath = "$PSScriptRoot\Resource\Hermes.ico"
    if(Test-Path($iconPath))
    {
        $icon = [System.Drawing.Icon]::ExtractAssociatedIcon($iconPath)
        $HermesForm.Icon = $icon
    }

    Fill-AccountDropDown $AccountDropDown

    if($accountDropDown.Items.Count -eq 0)
    { $HermesForm.Close() }
}

#---------------------------------------------- 

#region Form Code 
$HermesForm.Text = "Hermes - Delivery Reminder" 
$HermesForm.Name = "HermesForm"
$FetchDeliveriesButton.Text = "Fetch Deliveries"
$FetchDeliveriesButton.Name = "FetchDeliveriesButton"  
$GenerateRemindersButton.Text = "Generate Reminders"
$GenerateRemindersButton.Name = "GenerateRemindersButton"

$HermesForm.DataBindings.DefaultDataSourceUpdateMode = 0 
$FetchDeliveriesButton.DataBindings.DefaultDataSourceUpdateMode = 0
$GenerateRemindersButton.DataBindings.DefaultDataSourceUpdateMode = 0

$HermesForm.AutoSize = $True
$HermesForm.FormBorderStyle = [System.Windows.Forms.FormBorderStyle]::FixedSingle
$HermesForm.Padding = 10

$FetchDeliveriesButton.TabIndex = 0 
$GenerateRemindersButton.TabIndex = 0

#Account Drop Down
$AccountDropDown_Size = $sizeObject 
$AccountDropDown_Size.Width = 350 
$AccountDropDown_Size.Height = 23 
$AccountDropDown.Size = $AccountDropDown_Size 
$FetchDeliveriesButton.UseVisualStyleBackColor = $True

$AccountDropDown_Location_Point = $locationPoint 
$AccountDropDown_Location_Point.X = 13 
$AccountDropDown_Location_Point.Y = 13 
$AccountDropDown.Location = $AccountDropDown_Location_Point 

#Fetch Delivery Button
$FetchDeliveriesButton_Size = $sizeObject 
$FetchDeliveriesButton_Size.Width = 150 
$FetchDeliveriesButton_Size.Height = 23 
$FetchDeliveriesButton.Size = $FetchDeliveriesButton_Size 
$FetchDeliveriesButton.UseVisualStyleBackColor = $True

$FetchDeliveriesButton_Location_Point = $locationPoint 
$FetchDeliveriesButton_Location_Point.X = 13 
$FetchDeliveriesButton_Location_Point.Y = 45 
$FetchDeliveriesButton.Location = $FetchDeliveriesButton_Location_Point 

#Generate Reminder Button
$GenerateRemindersButton_Size = $sizeObject 
$GenerateRemindersButton_Size.Width = 150 
$GenerateRemindersButton_Size.Height = 23 
$GenerateRemindersButton.Size = $GenerateRemindersButton_Size 
$GenerateRemindersButton.UseVisualStyleBackColor = $True

$GenerateRemindersButton_Location_Point = $locationPoint  
$GenerateRemindersButton_Location_Point.X = 13 
$GenerateRemindersButton_Location_Point.Y = 90 
$GenerateRemindersButton.Location = $GenerateRemindersButton_Location_Point 

#Delivery list
$DeliveryList_Size = $sizeObject 
$DeliveryList_Size.Width = 200 
$DeliveryList_Size.Height = 270 
$DeliveryList.Size = $DeliveryList_Size
$DeliveryList.AutoSize = $True
$DeliveryList.SelectionMode = [System.Windows.Forms.SelectionMode]::MultiSimple
$DeliveryList.HorizontalScrollbar = $True

$DeliveryList_Location_Point = $locationPoint 
$DeliveryList_Location_Point.X = 180 
$DeliveryList_Location_Point.Y = 45 
$DeliveryList.Location = $DeliveryList_Location_Point

#Button click controls
$FetchDeliveriesButton.add_Click($handler_FetchDeliveriesButton_Click)
$GenerateRemindersButton.add_Click($handler_GenerateRemindersButton_Click)

#Adding controls to forms
$HermesForm.Controls.Add($AccountDropDown)
$HermesForm.Controls.Add($FetchDeliveriesButton)
$HermesForm.Controls.Add($GenerateRemindersButton)
$HermesForm.Controls.Add($DeliveryList)
#endregion Form Code

#Save the initial state of the form 
$InitialFormWindowState = $HermesForm.WindowState 
#Init the OnLoad event to correct the initial state of the form 
$HermesForm.add_Load($OnLoadForm_StateCorrection)
#Closing event

#Show the Form 
$HermesForm.ShowDialog()| Out-Null
} #End Function

function Fill-AccountDropDown
{
    [CmdletBinding()]
    param([System.Windows.Forms.ComboBox] $accountDropDown)
    
    $outlookApp = Get-OutlookApplication
    
    if($outlookApp -ne $null)
    {
        $outlookApp.Session.Accounts | %{$accountDropDown.Items.Add($_.DisplayName)}
    
        $AccountDropDown.DropDownStyle = [System.Windows.Forms.ComboBoxStyle]::DropDownList
	    if($accountDropDown.Items.Count -gt 0) { $accountDropDown.SelectedIndex = 0 }
	}  
}
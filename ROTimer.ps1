Add-Type -AssemblyName System.Windows.Forms
Add-Type -AssemblyName System.Drawing

<#
.SYNOPSIS
  This script is a Ragnarok Online MVP timer.

.DESCRIPTION
  The script creates a Windows form with buttons which can be used to register kills against MVPs.
  The script also includes a timer that checks the current time and changes the color of the textfield depending on the spawn status.
  Green = spawn window starts in 5 minutes
  Blue = boss can spawn now
  Red = boss has already spawned
  When an MVP spawn is close, a toast notification will be shown.
  MVPs can be added by creating another PSCustomObject and adding the variable name to the MVPs array.

.NOTES
  Version:        1.0
  Author:         Jordy Groenewoud
  Creation Date:  27/08/2023

#>

# Timers
$phree = [PSCustomObject]@{
    "name" = "Phreeoni"
    "shortName" = "phree"
    "delay" = 120
    "hasTomb" = $true
}
$eddga = [PSCustomObject]@{
    "name" = "Eddga"
    "shortName" = "eddga"
    "delay" = 120
    "hasTomb" = $true
}
$drake = [PSCustomObject]@{
    "name" = "Drake"
    "shortName" = "drake"
    "delay" = 120
    "hasTomb" = $true
}
$lod = [PSCustomObject]@{
    "name" = "Lord of the Dead"
    "shortName" = "lod"
    "delay" = 133
    "hasTomb" = $false
}
$gtb = [PSCustomObject]@{
    "name" = "Golden Thief Bug"
    "shortName" = "gtb"
    "delay" = 60
    "hasTomb" = $true
}
$mis = [PSCustomObject]@{
    "name" = "Mistress"
    "shortName" = "mis"
    "delay" = 120
    "hasTomb" = $true
}
$misGD = [PSCustomObject]@{
    "name" = "Mistress (Alde GD)"
    "shortName" = "misGD"
    "delay" = 480
    "hasTomb" = $true
}
$orcL = [PSCustomObject]@{
    "name" = "Orc Lord"
    "shortName" = "orcL"
    "delay" = 120
    "hasTomb" = $true
}
$orcH = [PSCustomObject]@{
    "name" = "Orc Hero"
    "shortName" = "orcH"
    "delay" = 60
    "hasTomb" = $true
}
$phar = [PSCustomObject]@{
    "name" = "Pharaoh"
    "shortName" = "phar"
    "delay" = 60
    "hasTomb" = $true
}
$moon = [PSCustomObject]@{
    "name" = "Moonlight Flower"
    "shortName" = "moon"
    "delay" = 60
    "hasTomb" = $true
}
$amon = [PSCustomObject]@{
    "name" = "Amon Ra"
    "shortName" = "amon"
    "delay" = 60
    "hasTomb" = $true
}
$maya = [PSCustomObject]@{
    "name" = "Maya"
    "shortName" = "maya"
    "delay" = 120
    "hasTomb" = $true
}
$bloo = [PSCustomObject]@{
    "name" = "Bloody Knight"
    "shortName" = "bloo"
    "delay" = 60
    "hasTomb" = $false
}

$MVPs = @($lod,$bloo,$gtb,$amon,$orcL,$orcH,$maya,$phar,$moon,$phree,$mis,$eddga,$drake,$misGD)

$form = New-Object System.Windows.Forms.Form
$form.Text = 'RO MVP Timer'
$form.Size = New-Object System.Drawing.Size(600,900)
$form.StartPosition = 'CenterScreen'

$position = 20
$varsToRemove = @()

function Killed {
    Param(
        $mvpName
    )

    $mvpName = $mvpName -replace "KillButton",""
    $mvpToEdit = $MVPs | Where-Object {$_.shortName -eq $mvpName}
    $timer = (Get-Date).AddMinutes($mvpToEdit.delay).ToString("HH:mm")

    $timerLabel = $form.Controls | Where-Object {$_.Name -eq "$($mvpToEdit.shortName)Timer"}
    $timerLabel.ForeColor = 'Black'
    $timerLabel.Text = "Next spawn: $timer"
}

function TombFound {
    Param(
        $mvpName
    )

    $mvpName = $mvpName -replace "TombButton",""
    $mvpName = $mvpName -replace "TombField",""
    $mvpToEdit = $MVPs | Where-Object {$_.shortName -eq $mvpName}

    $timeBox = $form.Controls | Where-Object {$_.Name -eq "$($mvpToEdit.shortName)TombField"}
    $enteredTime = $timeBox.Text
    
    $timer = (Get-Date $enteredTime).AddMinutes($mvpToEdit.delay).ToString("HH:mm")
    $timerLabel = $form.Controls | Where-Object {$_.Name -eq "$($mvpToEdit.shortName)Timer"}
    $timerLabel.ForeColor = 'Black'
    $timerLabel.Text = "Next spawn: $timer"

    $timeBox.Clear()
}

function Show-Notification {
    [cmdletbinding()]
    Param (
        $ToastTitle,
        $ToastText
    )

    [Windows.UI.Notifications.ToastNotificationManager, Windows.UI.Notifications, ContentType = WindowsRuntime] > $null
    $Template = [Windows.UI.Notifications.ToastNotificationManager]::GetTemplateContent([Windows.UI.Notifications.ToastTemplateType]::ToastText02)
    $RawXml = [xml] $Template.GetXml()
    ($RawXml.toast.visual.binding.text|where {$_.id -eq "1"}).AppendChild($RawXml.CreateTextNode($ToastTitle)) > $null
    ($RawXml.toast.visual.binding.text|where {$_.id -eq "2"}).AppendChild($RawXml.CreateTextNode($ToastText)) > $null

    $SerializedXml = New-Object Windows.Data.Xml.Dom.XmlDocument
    $SerializedXml.LoadXml($RawXml.OuterXml)

    $Toast = [Windows.UI.Notifications.ToastNotification]::new($SerializedXml)
    $Toast.Tag = "RO MVP Timer"
    $Toast.Group = "RO MVP Timer"
    $Toast.ExpirationTime = [DateTimeOffset]::Now.AddSeconds(10)

    $Notifier = [Windows.UI.Notifications.ToastNotificationManager]::CreateToastNotifier("RO MVP Timer")
    $Notifier.Show($Toast);
}


foreach ($MVP in $MVPs) {
    $tempLabel = New-Object System.Windows.Forms.Label
    $tempLabel.Location = New-Object System.Drawing.Point(10,($position+4))
    $tempLabel.Size = New-Object System.Drawing.Size(150,23)
    $tempLabel.Text = $MVP.name
    $tempLabel.Name = "$($MVP.shortName)Label"

    New-Variable -Name "$($MVP.shortName)Label" -Value $tempLabel -Force
    $varsToRemove += "$($MVP.shortName)Label"
    $tempLabel = (Get-Variable -Name "$($MVP.shortName)Label").Value
    $form.Controls.Add($tempLabel)

    $tempKillButton = New-Object System.Windows.Forms.Button
    $tempKillButton.Location = New-Object System.Drawing.Point(220,$position)
    $tempKillButton.Size = New-Object System.Drawing.Size(75,25)
    $tempKillButton.Text = "Killed"
    $tempKillButton.Name = "$($MVP.shortName)KillButton"
    $tempKillButton.Add_Click({ Killed($this.Name) })

    # Add a button that registers a boss kill with the default spawn delay
    New-Variable -Name "$($MVP.shortName)KillButton" -Value $tempKillButton -Force
    $varsToRemove += "$($MVP.shortName)KillButton"
    $tempKillButton = (Get-Variable -Name "$($MVP.shortName)KillButton").Value
    $form.Controls.Add($tempKillButton)

    # Only include a button and field for tombs if the boss has one
    if ($MVP.hasTomb) {
        # Add a button that uses the value of the tomb field to calculate the next spawn time
        $tempTombButton = New-Object System.Windows.Forms.Button
        $tempTombButton.Location = New-Object System.Drawing.Point(295,$position)
        $tempTombButton.Size = New-Object System.Drawing.Size(55,25)
        $tempTombButton.Text = "Tomb:"
        $tempTombButton.Name = "$($MVP.shortName)TombButton"
        $tempTombButton.Add_Click({ TombFound($this.Name) })

        New-Variable -Name "$($MVP.shortName)TombButton" -Value $tempTombButton -Force
        $varsToRemove += "$($MVP.shortName)TombButton"
        $tempTombButton = (Get-Variable -Name "$($MVP.shortName)TombButton").Value
        $form.Controls.Add($tempTombButton)

        # Add a field where an earlier kill can be registered
        # ToDo: incorporate the server time offset so the time value on the tombstone can be entered directly
        $tempTombField = New-Object System.Windows.Forms.RichTextBox
        $tempTombField.Location = New-Object System.Drawing.Point(355,$position)
        $tempTombField.Size = New-Object System.Drawing.Size(60,25)
        $tempTombField.font = new-object system.drawing.font "Arial",12
        $tempTombField.Name = "$($MVP.shortName)TombField"
        $tempTombField.Add_KeyDown({
            if ($_.KeyCode -eq [System.Windows.Forms.Keys]::Enter) {
                TombFound($this.Name)
                $this.Clear()
                $_.Handled = $true
            }
        })
        $form.Controls.Add($tempTombField)
    }

    # Create a field to store the spawn timer
    $tempTimer = New-Object System.Windows.Forms.Label
    $tempTimer.Location = New-Object System.Drawing.Point(440,($position+4))
    $tempTimer.Size = New-Object System.Drawing.Size(150,23)
    $tempTimer.Text = 'Next spawn:'
    $tempTimer.Name = "$($MVP.shortName)Timer"

    New-Variable -Name "$($MVP.shortName)Timer" -Value $tempTimer -Force
    $varsToRemove += "$($MVP.shortName)Timer"
    $tempTimer = (Get-Variable -Name "$($MVP.shortName)Timer").Value
    $form.Controls.Add($tempTimer)

    # Move next MVP line 30 pixels down
    $position = $position + 30
}


# Run a timer that checks the current time
# and changes the color of the textfield depending on the spawn status
# Green = spawn window starts in 5 minutes
# Blue = boss can spawn now
# Red = boss has already spawned 
$global:timer = New-Object -type System.Windows.Forms.Timer
$global:timer.Interval = 29000
$global:timer.add_Tick({
    $global:textFields = $form.Controls | Where-Object {$_.Name -like "*Timer"}
    foreach ($textField in $textFields) {
        if ($textField.Text -ne "Next spawn:") {
            $global:timerValue = $textField.Text.Split(" ")[2].Trim()
            if ((Get-Date($timerValue)) -lt ((Get-Date).AddMinutes(5))) {
                $textField.ForeColor = 'Green'
                if ((Get-Date($timerValue)) -gt ((Get-Date).AddMinutes(4))) {
                    $mvpName = ($MVPs | Where-Object {$_.shortName -eq ($textField.Name.Replace("Timer",""))}).Name
                    [string]$global:noteName = $mvpName
                    [string]$global:noteText = "can spawn within 5 minutes."
                    Show-Notification($noteName,$noteText)
                }
                if ((Get-Date($timerValue)) -le (Get-Date)) {
                    $textField.ForeColor = 'Blue'
                    if ((Get-Date($timerValue)) -le ((Get-Date).AddMinutes(-10))) {
                        $textField.ForeColor = 'Red'
                    }
                }
            }
        }
    }
})

$timer.Start()
$result = $form.ShowDialog()
if ($result –eq [System.Windows.Forms.DialogResult]::Cancel)
{
    foreach ($varToRemove in $varsToRemove) {
        Remove-Variable -Name $varToRemove
    }
    $form.Dispose()
    $global:timer.Dispose()
}
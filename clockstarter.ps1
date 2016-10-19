<#
    This Script includes the windows forms to configure the settings on the main clock script
    Beware, the script closes all excel files that are open, if you don't want that, remove row nr 24 "(Get-Process excel).kill()"
#>
cls

Add-Type -AssemblyName System.Drawing, System.Windows.Forms
$xlsxpath = "C:\Path\to\Excel.xlsx" #Put the path to your excel file here

#The function for getting the time from the excel file, but you'll most likely have to change a few things for it to work
function exceldata{
    #Opens the Excel Application
    $Excel = NEW-Object –ComObject Excel.Application
    #Opens the before specified path for the Excel File
    $Workbook = $Excel.Workbooks.Open($xlsxpath)
    $Worksheet = $Workbook.sheets.item("$(Get-Date -Format "MMMM")")
    $line = ((Get-Date).AddDays(1)).Day
    $worktime = $Worksheet.Cells.Item(($line),1).Value2

    #Closes excel
    $Workbook.close()
    $Excel.Quit()
    Remove-Variable Excel
    (Get-Process excel).kill()
}

#Generates the additional "choose your own time" form
$timeform = New-Object Windows.Forms.Form
$timeform.Size = New-Object Drawing.Point 150,80
$timeform.Startposition = "CenterScreen"
$timeform.BackColor = "white"

$data = New-Object System.Windows.Forms.TextBox
$data.Location = New-Object Drawing.Point 10,10
$data.Size = New-Object Drawing.Point 120,50
$data.MaxLength = 5
$data.Add_KeyDown({
    if($data.TextLength -eq 1 -and $_.KeyCode -ne "Back"){
        $_.KeyCode >> "$env:AP_Desktop\clock\log.log"
        [System.Windows.Forms.SendKeys]::SendWait(":")
    }
    if($_.KeyCode -eq "Enter"){
        $global:var3 = $data.Text
        $timeform.Close()
    }
})
$timeform.controls.add($data)


#Generates the main form
$mainform = New-Object Windows.Forms.Form
$mainform.Text = "ClockSettings"
$mainform.Size = New-Object Drawing.Point 530,115
$mainform.Startposition = "CenterScreen"
$mainform.BackColor = "white"

#Geberates and configures the buttons for the main form
$presetbtn = New-Object Windows.Forms.Button
$win1nostat = New-Object Windows.Forms.Button
$win1stat = New-Object Windows.Forms.Button
$win1nostat = New-Object Windows.Forms.Button
$win2nostat = New-Object Windows.Forms.Button
$win2stat = New-Object Windows.Forms.Button

$presetbtn.Location = New-Object Drawing.Point 10,15
$win1nostat.Location = New-Object Drawing.Point 110,15
$win1stat.Location = New-Object Drawing.Point 210,15
$win2nostat.Location = New-Object Drawing.Point 310,15
$win2stat.Location = New-Object Drawing.Point 410,15

$presetbtn.Size = New-Object Drawing.Point 90,50
$win1nostat.Size = New-Object Drawing.Point 90,50
$win1stat.Size = New-Object Drawing.Point 90,50
$win2nostat.Size = New-Object Drawing.Point 90,50
$win2stat.Size = New-Object Drawing.Point 90,50

$presetbtn.Text = "QuickSettings"
$win1nostat.Text = "Window 1, not Static"
$win1stat.Text = "Window 1, Static"
$win2nostat.Text = "Window 2, not Static"
$win2stat.Text = "Window 2, Static"

$presetbtn.Add_Click({
    exceldata
    $global:var1 = 1
    $global:var2 = 0
    $mainform.Close()
})
$win1nostat.Add_Click({
    $timeform.ShowDialog()
    $global:var1 = 1
    $global:var2 = 0
    $mainform.Close()
})
$win1stat.Add_Click({
    $timeform.ShowDialog()
    $global:var1 = 1
    $global:var2 = 1
    $mainform.Close()
})
$win2nostat.Add_Click({
    $timeform.ShowDialog()
    $global:var1 = 2
    $global:var2 = 0
    $mainform.Close()
})
$win2stat.Add_Click({
    $timeform.ShowDialog()
    $global:var1 = 2
    $global:var2 = 1
    $mainform.Close()
})

#Add the buttons to the form
$mainform.controls.add($presetbtn)
$mainform.controls.add($win1nostat)
$mainform.controls.add($win1stat)
$mainform.controls.add($win2nostat)
$mainform.controls.add($win2stat)

$mainform.ShowDialog()

#Checks if the input from the textbox is valid
try{
    Get-Date $global:var3
}catch{
    exceldata
}finally{
    #Starts the clock script with the parameters
    ."$PSScriptRoot\clock.ps1" -monitor $global:var1 -static $global:var2 -alarmtime $global:var3
}
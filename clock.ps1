<#
These are the scripts I used as a template
The link to the original work https://gallery.technet.microsoft.com/scriptcenter/Clock-Widget-574c2988
https://gallery.technet.microsoft.com/scriptcenter/Popup-Toast-WPF-PowerShell-e228e7a3
#>

#Parameters to adjust the settings (these are just the standard ones, use the clockstarter script to change them)
Param(
    [parameter()]
    [int]$monitor = 1,
    [parameter()]
    [bool]$static = $false,
    [parameter()]
    [string]$alarmtime = "17:30"
)

#Adds the assemblies
Add-Type -AssemblyName PresentationCore,PresentationFramework,WindowsBase,System.Drawing,System.Windows.Forms

#Preparing the hashtables
$Clockhash = [hashtable]::Synchronized(@{})
$Runspacehash = [hashtable]::Synchronized(@{})

$alarm = get-date $alarmtime #changes the time for the alarm into an actual time-variable
$Clockhash.int = 1
$Clockhash.monitor = $monitor
$Clockhash.static = $static
$Clockhash.alarm = $alarm
$Runspacehash.host = $Host

#Creating the runspace
$Runspacehash.runspace = [RunspaceFactory]::CreateRunspace()
$Runspacehash.runspace.ApartmentState = “STA”
$Runspacehash.runspace.ThreadOptions = “ReuseThread”
$Runspacehash.runspace.Open()
$Runspacehash.psCmd = {Add-Type -AssemblyName PresentationCore,PresentationFramework,WindowsBase,System.Drawing,System.Windows.Forms}.GetPowerShell()
$Runspacehash.runspace.SessionStateProxy.SetVariable("Clockhash",$Clockhash)
$Runspacehash.runspace.SessionStateProxy.SetVariable("Runspacehash",$Runspacehash)
$Runspacehash.psCmd.Runspace = $Runspacehash.runspace
$Runspacehash.Handle = $Runspacehash.psCmd.AddScript({
    #Updates the variables
    $Script:Update = {
        $day,$Month, $Day_n, $Year, $Time = (Get-Date -f "dddd,MMMM,dd,yyyy,HH:mm:ss") -Split ','
        $Clockhash.time_txtbox.text = $Time.TrimStart("0")
        $Clockhash.day_txtbx.Text = $day
        $Clockhash.day_n_txtbx.text = $Day_n
        $Clockhash.month_txtbx.text = $Month
        $Clockhash.year_txtbx.text = $year

        $now = Get-Date

        #Checks if the time is past the alarm
        if($now -ge $Clockhash.alarm -and $clockhash.int -eq 0){
	        $clockhash.window.Show()
            $clockhash.Window.Activate()
            [System.Windows.Forms.MessageBox]::Show("Time to leave, the alarm was set for $(($Clockhash.alarm.AddMinutes(5)).ToShortTimeString())","It makes ring",0,[System.Windows.Forms.MessageBoxIcon]::Exclamation)
            $clockhash.int = 1
        }
    }

    #The xml data, to generate the window
    [xml]$xaml = @"
<Window 
        xmlns="http://schemas.microsoft.com/winfx/2006/xaml/presentation"
        xmlns:x="http://schemas.microsoft.com/winfx/2006/xaml"
        WindowStyle = "None" SizeToContent = "WidthAndHeight" ShowInTaskbar = "False" IsHitTestVisible="False" Topmost = "True"
        Title = "Test" Background = "Black" Opacity = "1" AllowsTransparency = "True" WindowStartupLocation = "Manual">
    <Grid x:Name = "Grid" Background = "Transparent" Margin = "15,20,15,15">
        <TextBlock x:Name = "time_txtbox" FontSize = "60" VerticalAlignment="Top" Foreground = "white"
        HorizontalAlignment="Left" Margin="0,-26,0,0">
                <TextBlock.Effect>
                    <DropShadowEffect Color = "Black" ShadowDepth = "1" BlurRadius = "5" />
                </TextBlock.Effect>
        </TextBlock>
        <TextBlock x:Name = "day_n_txtbx" FontSize = "38" Margin = "5,42,0,0" Foreground = "white" 
        HorizontalAlignment="Left">
                <TextBlock.Effect>
                    <DropShadowEffect Color = "Black" ShadowDepth = "1" BlurRadius = "2"  />
                </TextBlock.Effect>
        </TextBlock>
        <TextBlock x:Name = "month_txtbx" FontSize=  "20" Margin = "54,48,0,0" Foreground = "white" 
        HorizontalAlignment="Left">
                <TextBlock.Effect>
                    <DropShadowEffect Color = "Black" ShadowDepth = "1" BlurRadius = "2" />
                </TextBlock.Effect>
        </TextBlock>
        <TextBlock x:Name = "day_txtbx" FontSize=  "15" Margin="54,68,0,0" Foreground = "white"
        HorizontalAlignment="Left">
                <TextBlock.Effect>
                    <DropShadowEffect Color = "Black" ShadowDepth = "1" BlurRadius = "2" />
                </TextBlock.Effect>
        </TextBlock>
        <TextBlock x:Name = "year_txtbx" FontSize=  "38"  Margin="140,42,0,0" Foreground = "white"
        HorizontalAlignment="Left">
                <TextBlock.Effect>
                    <DropShadowEffect Color = "Black" ShadowDepth = "1" BlurRadius = "2" />
                </TextBlock.Effect>
        </TextBlock>
    </Grid>
</Window>
"@
    #Use the previously defined xml variable to create a window
    $reader=(New-Object System.Xml.XmlNodeReader $xaml)
    $Clockhash.Window=[Windows.Markup.XamlReader]::Load( $reader )

    $Clockhash.time_txtbox = $Clockhash.window.FindName("time_txtbox")
    $Clockhash.day_n_txtbx = $Clockhash.Window.FindName("day_n_txtbx")
    $Clockhash.month_txtbx = $Clockhash.Window.FindName("month_txtbx")
    $Clockhash.year_txtbx = $Clockhash.Window.FindName("year_txtbx")
    $Clockhash.day_txtbx = $Clockhash.Window.FindName("day_txtbx")

    #Executes when the window has been opened for the first time
    $Clockhash.Window.Add_SourceInitialized({
        #Create Timer object
        $Script:timer = new-object System.Windows.Threading.DispatcherTimer 
        #Fire off every second
        $timer.Interval = [TimeSpan]"0:0:1"
        $Clockhash.int = 0
        #Add event per tick
        $timer.Add_Tick({
            $Update.Invoke()
        })

        #Start timer
        $timer.Start()

        #Switches the monitors (only works with a screen resolution of 1920 by 1200 atm, but I will fix it)
        if($clockhash.monitor -eq 1){
            $winwidth = 1669
            $winheight = 1032
        }else{
            $winwidth = 3588
            $winheight = 1072
        }

        #Put the window into the bottom-right corner of the screen/main monitor
        $clockhash.window.Left = $winwidth
	    $clockhash.window.Top = $winheight
    }) 

    # Create notifyicon, and right-click -> Exit menu
    $clockhash.icon = [System.Drawing.Icon]::ExtractAssociatedIcon("$pshome\powershell.exe") #Gets the Icon for the notifyicon
    $clockhash.notifyicon = New-Object System.Windows.Forms.NotifyIcon
    $clockhash.notifyicon.Text = "Clock Widget"
    $clockhash.notifyicon.Icon = $clockhash.icon
    $clockhash.notifyicon.Visible = $true

    #creates the menuitem for the contextmenu
    $clockhash.menuitem = New-Object System.Windows.Forms.MenuItem
    $clockhash.menuitem.Text = "Exit"

    #Generates the contextmenu (the right-click menu on the notifyicon)
    $clockhash.contextmenu = New-Object System.Windows.Forms.ContextMenu
    $clockhash.notifyicon.ContextMenu = $clockhash.contextmenu
    $clockhash.notifyicon.ContextMenu.MenuItems.AddRange($clockhash.menuitem)

    #Close the window when it loses focus and only then, if static has been disabled
    $clockhash.window.Add_Deactivated({
        if(!$clockhash.static){
	        $clockhash.window.Hide()
        }
    })

    #Shows the window when you click on the notifyicon
    $clockhash.notifyicon.add_Click({
	    if($_.Button -eq [Windows.Forms.MouseButtons]::Left){
	        $clockhash.window.Show()
            $clockhash.Window.Activate()
        }
    })

    #The Event when the menutitem has been clicked
    $clockhash.menuitem.add_Click({
	    $clockhash.notifyicon.Visible = $false
	    $clockhash.window.Close()
        $timer.Stop()
        $Runspacehash.PowerShell.Dispose()
            
        [gc]::Collect()
        [gc]::WaitForPendingFinalizers()
    })
        
    #Activates the window for the first run
    $clockhash.window.Show()
    $clockhash.Window.Activate()

    #Runs it seperately so not only don't you need to have powershell open, but exiting works smoother too
    $appContext = New-Object System.Windows.Forms.ApplicationContext
    [void][System.Windows.Forms.Application]::Run($appContext)
}).BeginInvoke()
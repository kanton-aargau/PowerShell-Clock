# PowerShell-Clock
A simple PowerShell Clock, with a Windows 10-ish design

#Functions / What can it do
- Tell the time in a Windows 10-ish design
- Set an alarm (a popup message appears 5 minutes before your desired time)
- Run in the Background (so you don't have to keep Powershell Open)
- It displays the time in Hours:Minutes:Seconds
- It displays the Date (Day, Month, DayofWeek, Year)
- Using a pre-planned Excel-file to set the alarms

#How does it work in a Nutshell
- Runspaces and synchronized hashtables
- Timers, to update the variables every second
- xml, to generate the window
- Notifyicon, so theres no open window

The actual clock part wasn't 100% made by me, so here is a link to the original one, but I did a lot of work to upgrading it.
https://gallery.technet.microsoft.com/scriptcenter/Clock-Widget-574c2988

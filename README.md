# PowerShell-Clock
A simple PowerShell Clock, with a Windows 10-ish design

##Functions / What can it do
- Tell the time in a Windows 10-ish design
- Set an alarm (a popup message appears 5 minutes before your desired time)
- Run in the Background (so you don't have to keep Powershell Open)
- It displays the time in Hours:Minutes:Seconds
- It displays the Date (Day, Month, DayofWeek, Year)
- Using a pre-planned Excel-file to set the alarms

##How does it work in a Nutshell
- Runspaces and synchronized hashtables
- Timers, to update the variables every second
- xml, to generate the window
- Notifyicon, so theres no open window

Not everything here was made by me, so here is a link to the original sources
* https://gallery.technet.microsoft.com/scriptcenter/Clock-Widget-574c2988
* https://gallery.technet.microsoft.com/scriptcenter/Popup-Toast-WPF-PowerShell-e228e7a3

But I did a lot of work to upgrade and combine them, so they work well with each otter (pun intended)

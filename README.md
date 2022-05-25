# outlook-Auto-Out-of-Office

VBA code that will auto-set or auto-clear a rule of your choice using appointment reminders. I used my "Out of Office" rule and set appointment reminders for my first day out of the office, so Outlook turns my auto-reply on, and the day I return, so Outlook turns it off.

I put it directly in ThisOutlookSession, but you may be able to put it in a module, I haven't tried.

To use:
- Copy code from OutlookRuleToggle file
- Open Outlook and bring up the VBA Editor (Alt+F11)
- Expand Microsoft Outlook Objects and double-click ThisOutlookSession to open, then paste the code into the window on the right
- Change lines 16 and 24 to match the rule you want to toggle
- Optionally, change lines 15 and 23 - these will be the subject lines for the appointments you will use to toggle the rule

If you don't want to keep getting a popup about macro security, I recommend the following process:
- Go into File > Options > Trust Center > Trust Center Settings > Macro Settings. Set to "Notifications for digitally signed macros, all other macros disabled".
- Create a self-signed digital certificate using the SELFCERT.EXE tool that comes with Office. This will be located in the same directory where your Office applications are installed; for me (64-bit Office 365 desktop) it is in C:\Program Files\Microsoft Office\root\Office16.
- Back in Outlook, open the VBA Editor again (Alt+F11), and from the Tools menu select Digital Signature...
- Choose the digital certificate you jsut created and click OK. Save your VBA project and exit Outlook.
- Re-open Outlook. When prompted (you may have to re-open the VBA Editor first), click on Show Signature Details > View Certificate > Install Certificate, then follow the instructions given.

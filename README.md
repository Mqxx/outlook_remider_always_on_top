# Outlook Reminder Always-on-Top
Show the Outlook Reminder Window always on top.

> [!WARNING]
> This project was designed for **PowerShell 7** and above. It may work with older versions of **PowerShell** but was not tested.
> 
> Tested on **Windows 11** running **PowerShell 7.5.2**.

## Usage
1. Download [`always_on_top.ps1`](./always_on_top.ps1) and save it somewhere.
2. Edit the variables for `# Configuration` to your likings.
3. Download [`run_hidden_always_on_top.vbs`](./run_hidden_always_on_top.vbs) and save it next to your [`always_on_top.ps1`](./always_on_top.ps1).
4. Run the [`run_hidden_always_on_top.vbs`](./run_hidden_always_on_top.vbs) script to start it.

If you want to run the script in the background on startup follow this:

5. Create a Desktop Shortcut from your [`run_hidden_always_on_top.vbs`](./run_hidden_always_on_top.vbs) script and place it in your autostart folder, located at: `%appdata%\Microsoft\Windows\Start Menu\Programs\Startup`

After starting the script (either by running the script manually or by autostart) you will get a new try icon. You can quit the script any time by <kbd>rightclick</kbd> on the tray icon and <kbd>click</kbd> "Exit".

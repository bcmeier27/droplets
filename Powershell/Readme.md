Place these on your Windows Desktop to mimic the macOS "droplet" behavior:

1. **Word Link Droplet**
   This is a Windows shortcut to `WordLinkDroplet.bat`
   *(You might need to edit the shortcut properties to select the correct user's Desktop folder.)*

2. **WordLinkDroplet.bat**
   This batch file opens the `WordLinkDroplet.ps1` script:
   - Bypasses certain security parameters that prevent making a shortcut to the `.ps1` script itself
   - Enables actual droplet-like behavior

3. **WordLinkDroplet.ps1**
   The actual PowerShell script

> **Note:**
> Technically, you can place items 2 & 3 in any folder **as long as** the `.bat` file opens the `.ps1` script from its correct location.
> This means:
> - Edit the `.bat` file to include the correct path to the `.ps1` file
> - Edit the shortcut properties to point to the updated `.bat` file

#eof


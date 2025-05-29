Place these on your Windows Desktop to mimic the MacOS "droplet" behavior:

1. Word Link Droplet - this is a Windows shortcut to WordLinkDroplet.bat (might have to edit properties to select the user name's desktop folder)
2. WordLinkDroplet.bat - this batch file opens the WordLinkDroplet.ps1 script, 
		       - bypassing certain security parameters that prevent making a shortcut to the .ps1 script itself, 
		       - allowing the actual droplet behavior

3. WordLinkDroplet.ps1 - the actual PowerShell script

Technically you can place item 2 & 3 in any folder AS LONG AS the .bat file opens the .ps1 script from it's non-Desktop location.
Which means: edit .bat to include the correct path to the .ps1 file; edit the shortcut properties to set the target file to the edited .bat file.


#eof
# Artemis Law Journal Proofreading Toolbar
A Microsoft Word toolbar for proofreading and fixing common errors in law journal publications.

# Functions
The toolbar contains tools to check for and fix (1) journal name abbreviation errors or journal names not in small caps, (2) missing supra/infra cross-reference links, (3) missing permalinks or active hyperlinks, (4) un-italicized introductory signals, (5) missing periods at the end of footnotes, and (6) case name abbreviation errors. 

Note that there are edge cases in which the toolbar catches false positives or fails to catch true positives. The UI will highlight and show each error for the user to decide to fix (whether manually or using an automated fix).

# How to use
Open Microsoft Word, go to Tools -> Templates and Add-ins, and add the "Artemis Law Journal Toolbar.dotm" file as a new add-in.

Alternatively, copy "Artemis Law Journal Toolbar.dotm" to your Microsoft Word STARTUP folder. The startup folder is located at "C:\Users\[username]\AppData\Roaming\Microsoft\Word\STARTUP" in Windows Vista, 7, 8 and 10 for example.

After enabling macros, a tab named "Artemis" should appear in your Ribbon. This should only need to be done the first time using the toolbar.

The source code can be viewed and edited by opening "Artemis Law Journal Toolbar.docm" in Word and pressing Alt+F11 to open up the VBA developer interface. After editing, you need to save the docm file as a dotm, and then copy the new dotm to the startup folder.

# License
Creative Commons Attribution-NonCommercial 3.0 United States (CC BY-NC 3.0 US)

# Credits
Ian Li, Ryan Yeh, and Kyle Victor
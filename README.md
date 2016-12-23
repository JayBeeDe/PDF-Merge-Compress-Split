# PDF-Merge-Compress-Split

## GOAL

The script aims at performing the following tasks on pdf files:
- Only individually compress some pdf files (Compress Only) | Will replace the original file(s)!
- Merge some pdf files into one and compress the output file (Merge and Compress)
- Merge some pdf files into one, compress and split into several output files (Merge, Compress and Split) (NOT IMPLEMETED YET)

This powershell is a simple interface (GUI or CLI) that calls the Ghostscript software (gswin64c.exe)

## FILES

- README.md
- Main.ps1
- gsdll64.dll
- gsdll64.lib
- gswin64c.exe
- pdf.ico

## LICENSE

- Ghostscript software is under General Public License (the "AGPL"): for more information, please read the
license available at http://www.artifex.com/page/licensing-information.html.
- Otherwise the script is under GNU GPL v3.0 license and can be edited, 
distributed for commercial/private use.
BUT
Script is provided without warranty and the author/license
owner cannot be held liable for damages.
You may not grant a sublicense to modify and distribute the code to
third parties not included in the license.
See license file or http://www.gnu.org/licenses/gpl.txt for more 
information.

## INSTALLATION & USE

Download and unzip (if needed) the files.

The script can be run directly by double clicking the script, or in command line.

Note: Execution Policy must be ENABLED in order to get the PowerShell script running.
Please See https://msdn.microsoft.com/en-us/powershell/reference/5.1/microsoft.powershell.security/set-executionpolicy

From command line, it accepts the following arguments:
-inputDirectory <string>: targeted directory or file (default: the current directory)
-mode <string>: [C/m] - compress / merge and compress
-cli <boolean>: [$true/$FALSE] - enable/disable CLI instead of GUI
-translate <boolean>: [$TRUE/$false] - enable/disable translation from english to the system menu (internet connexion required)
-recurse <boolean>: [$true/$FALSE] - If inputDirectory is a directory, wether search for file also recursively
-autoRotate <boolean>: [$TRUE/$false] - try to auto rotate pages
-outName <string>: output file when merging files (default value defined by $global:defaultOutName)

## EXAMPLE

In GUI: just double click the file or run from a PowerShell terminal:
.\Main.ps1

In CLI, to compress each pdf file in the script's directory (script referential is from the Main.ps1 file), run:
.\Main.ps1 -cli $true

This script doesn't need to be run as Administrator
> 23-12-2016 | Jean-Baptiste DELON
[Issues](https://github.com/JayBeeDe/PDF-Merge-Compress-Split/issues)

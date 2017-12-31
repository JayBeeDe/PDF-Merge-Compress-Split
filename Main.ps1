<#
Application: PDF Merge Compress Split
Description: This PowerShell Script can clean your start menu and your system shell.
Author: Jean-Baptiste Delon
Third Dependency: gswin64c.exe Ghostscript software

License: the script (wihout its tird dependency) is under GNU GPL v3.0 license and can be edited, distributed for commercial/private use.
For the Ghostscript software license, see http://www.artifex.com/page/licensing-information.html.
#>

param (
   [string]$inputDirectory,
   [string]$mode,
   [boolean]$cli,
   [boolean]$translate,
   [boolean]$recurse,
   [boolean]$autoRotate,
   [string]$outName,
   [boolean]$debug
)

$global:currentLocation=Split-Path -Path $MyInvocation.MyCommand.Path
$global:applicationName="PDF Merge Compress Split"
$global:defaultOutName="output"
$global:TranslateAccountKey="a9405b496e35440882154d696d71140c"
$global:TranslateTokenURL="https://api.cognitive.microsoft.com/sts/v1.0/issueToken"
$global:TranslateURL="https://api.microsofttranslator.com/v2/Http.svc/Translate"

try{
    Import-Module "$($global:currentLocation)\Core.psm1" -Force -ErrorAction Stop -Scope Local
}catch{
    write-host "An error has occured while loading the function core module" -ForegroundColor Red
    Exit
}
try{
    $global:systemLanguage=(Get-Culture).TwoLetterISOLanguageName
}catch{
    $global:systemLanguage="en"
    write-host "Unable to detect system language - The script has been set to the default system language - English!" -ForegroundColor Yellow
}

Add-Type -Name Window -Namespace Console -MemberDefinition '
[DllImport("Kernel32.dll")]
public static extern IntPtr GetConsoleWindow();

[DllImport("user32.dll")]
public static extern bool ShowWindow(IntPtr hWnd, Int32 nCmdShow);
'
[reflection.assembly]::LoadWithPartialName("'Microsoft.VisualBasic") | Out-Null

cd /
cls
display "The script is begining!"

if($mode -eq "Merge"){
    $global:mode="m"
}else{
    $global:mode="c"
}

if($cli -eq $true){
    $global:cli=$true
}else{
    $global:cli=$false
}
if($translate -eq $false){
    $global:translate=$false
}else{
    $global:translate=$true
}
if($recurse -eq $true){
    $global:recurse=$true
}else{
    $global:recurse=$false
}
if($autoRotate -eq $false){
    $global:autoRotate=$false
}else{
    $global:autoRotate=$true
}
if($outName -eq $null -or $outName -eq ""){
    $global:outName=$global:defaultOutName
}else{
    $global:outName=$outName
}
if($debug -eq $true){
    $global:debug=$true
}else{
    $global:debug=$false
}

try{
    if(Test-Path -Path $inputDirectory -PathType Any -ErrorAction Stop){
        $global:inputDirectory=$inputDirectory
    }else{
        $global:inputDirectory=$global:currentLocation
    }
}catch{
    $global:inputDirectory=$global:currentLocation
}

if($global:cli -eq $true){
    actionFileProcess($global:inputDirectory)
}else{
    # Chargement des assemblies
    [void][System.Reflection.Assembly]::LoadWithPartialName("System.Windows.Forms")
    [void][System.Reflection.Assembly]::LoadWithPartialName("System.Drawing")
    
    # Creation de la form principale
    $global:form=New-Object Windows.Forms.Form
    $global:form.FormBorderStyle=[System.Windows.Forms.FormBorderStyle]::FixedDialog
    $global:form.MaximizeBox=$False
    $global:form.MinimizeBox=$True
    $global:form.Text=$global:applicationName
    $global:form.Size=New-Object System.Drawing.Size(400,414)
    $global:form.Icon=New-Object system.drawing.icon("$($global:currentLocation)\pdf.ico")
    $global:form.StartPosition="CenterScreen"

    $label1=New-Object System.Windows.Forms.Label
    $label1.Text=translate "Select the file or folder you want to process pdf files"
    $label1.Font="Arial,10"
    $label1.Location=New-Object System.Drawing.Point(5,5)
    $label1.MaximumSize=New-Object System.Drawing.Size(380,44)
    $label1.AutoSize=$true
    $global:form.Controls.Add($label1)

    $global:textBox=New-Object System.Windows.Forms.TextBox
    $global:textBox.Location=New-Object System.Drawing.Point(5,45)
    $global:textBox.Size=New-Object System.Drawing.Size(375,22)
    $global:textBox.AllowDrop=$true
    $global:textBox.Text=$global:inputDirectory
    $global:form.Controls.Add($global:textBox)

    $fileButton=New-Object System.Windows.Forms.Button
    $fileButton.Location=New-Object System.Drawing.Point(25,70)
    $fileButton.Size=New-Object System.Drawing.Size(150,22)
    $fileButton.Text=translate "Browse File"
    $fileButton.Add_Click({actionFileBrowse $global:textBox.Text})
    $global:form.Controls.Add($fileButton)

    $folderButton=New-Object System.Windows.Forms.Button
    $folderButton.Location=New-Object System.Drawing.Point(215,70)
    $folderButton.Size=New-Object System.Drawing.Size(150,22)
    $folderButton.Text=translate "Browse Folder"
    $folderButton.Add_Click({actionFolderBrowse $global:textBox.Text})
    $global:form.Controls.Add($folderButton)
    
    $global:recurseCheckbox=New-Object System.Windows.Forms.CheckBox
    $global:recurseCheckbox.Location=New-Object System.Drawing.Point(5,110)
    $global:recurseCheckbox.MaximumSize=New-Object System.Drawing.Size(200,22)
    $global:recurseCheckbox.Text=translate "Browse recurse folder"
    if($global:recurse -eq $true){
        $global:recurseCheckbox.Checked=$true
    }
    $global:recurseCheckbox.AutoSize=$true
    $global:form.Controls.Add($global:recurseCheckbox)
    
    $global:autoRotateCheckbox=New-Object System.Windows.Forms.CheckBox
    $global:autoRotateCheckbox.Location=New-Object System.Drawing.Point(205,110)
    $global:autoRotateCheckbox.MaximumSize=New-Object System.Drawing.Size(200,22)
    $global:autoRotateCheckbox.Text=translate "Automatic rotation of the page"
    if($global:autoRotate -eq $true){
        $global:autoRotateCheckbox.Checked=$true
    }
    $global:autoRotateCheckbox.AutoSize=$true
    $global:form.Controls.Add($global:autoRotateCheckbox)
    
    $global:modeRadio1=New-Object System.Windows.Forms.RadioButton
    $global:modeRadio1.Location=New-Object System.Drawing.Point(60,140)
    $global:modeRadio1.MaximumSize=New-Object System.Drawing.Size(300,22)
    $global:modeRadio1.Text=translate "Compress pdf files individually"
    if($global:mode -eq "c"){
        $global:modeRadio1.Checked=$true
    }
    $global:modeRadio1.AutoSize=$true
    $global:modeRadio1.Add_Click({HideShowPageSelection $false})
    $global:form.Controls.Add($global:modeRadio1)
    
    $global:modeRadio2=New-Object System.Windows.Forms.RadioButton
    $global:modeRadio2.Location=New-Object System.Drawing.Point(60,170)
    $global:modeRadio2.MaximumSize=New-Object System.Drawing.Size(300,22)
    $global:modeRadio2.Text=translate "Merge and Compress pdf files"
    if($global:mode -eq "m"){
        $global:modeRadio2.Checked=$true
    }
    $global:modeRadio2.AutoSize=$true
    $global:modeRadio2.Add_Click({HideShowPageSelection $false})
    $global:form.Controls.Add($global:modeRadio2)
    
    $global:modeRadio3=New-Object System.Windows.Forms.RadioButton
    $global:modeRadio3.Location=New-Object System.Drawing.Point(60,200)
    $global:modeRadio3.MaximumSize=New-Object System.Drawing.Size(300,22)
    $global:modeRadio3.Text=translate "Extract pages from a list of files or a unique file"
    if($global:mode -eq "s"){
        $global:modeRadio3.Checked=$true
    }
    $global:modeRadio3.AutoSize=$true
    $global:modeRadio3.Add_Click({HideShowPageSelection $true})
    $global:form.Controls.Add($global:modeRadio3)

    $global:label31=New-Object System.Windows.Forms.Label
    $global:label31.Text=translate "from page"
    $global:label31.Font="Arial,8"
    $global:label31.Location=New-Object System.Drawing.Point(70,230)
    $global:label31.MaximumSize=New-Object System.Drawing.Size(80,22)
    $global:label31.AutoSize=$true
    if($global:modeRadio3.Checked -eq $true){
        $global:label31.Enabled=$true
    }else {
        $global:label31.Enabled=$false
    }
    $global:form.Controls.Add($label31)

    $global:textBox31=New-Object System.Windows.Forms.TextBox
    $global:textBox31.Location=New-Object System.Drawing.Point(130,230)
    $global:textBox31.Size=New-Object System.Drawing.Size(22,22)
    $global:textBox31.AllowDrop=$false
    $global:textBox31.maxLength=3
    if($global:modeRadio3.Checked -eq $true){
        $global:textBox31.Enabled=$true
    }else {
        $global:textBox31.Enabled=$false
    }
    $global:form.Controls.Add($global:textBox31)

    $global:label32=New-Object System.Windows.Forms.Label
    $global:label32.Text=translate "to page"
    $global:label32.Font="Arial,8"
    $global:label32.Location=New-Object System.Drawing.Point(160,230)
    $global:label32.MaximumSize=New-Object System.Drawing.Size(80,22)
    $global:label32.AutoSize=$true
    if($global:modeRadio3.Checked -eq $true){
        $global:label32.Enabled=$true
    }else {
        $global:label32.Enabled=$false
    }
    $global:form.Controls.Add($label32)

    $global:textBox32=New-Object System.Windows.Forms.TextBox
    $global:textBox32.Location=New-Object System.Drawing.Point(210,230)
    $global:textBox32.Size=New-Object System.Drawing.Size(22,22)
    $global:textBox32.AllowDrop=$false
    $global:textBox32.maxLength=3
    if($global:modeRadio3.Checked -eq $true){
        $global:textBox32.Enabled=$true
    }else {
        $global:textBox32.Enabled=$false
    }
    $global:form.Controls.Add($global:textBox32)

    $global:successLabel=New-Object System.Windows.Forms.Label
    $global:successLabel.Visible=$false
    $global:successLabel.Font="Arial,8"
    $global:successLabel.ForeColor="Green"
    $global:successLabel.MaximumSize=New-Object System.Drawing.Size(380,66)
    $global:successLabel.Location=New-Object System.Drawing.Point(5,264)
    $global:successLabel.AutoSize=$true
    $global:form.Controls.Add($global:successLabel)

    $global:errorLabel=New-Object System.Windows.Forms.Label
    $global:errorLabel.Visible=$false
    $global:errorLabel.Font="Arial,8"
    $global:errorLabel.ForeColor="Red"
    $global:errorLabel.MaximumSize=New-Object System.Drawing.Size(380,66)
    $global:errorLabel.Location=New-Object System.Drawing.Point(5,264)
    $global:errorLabel.AutoSize=$true
    $global:form.Controls.Add($global:errorLabel)

    $global:progressStatus=New-Object System.Windows.Forms.Label
    $global:progressStatus.Visible=$false
    $global:progressStatus.Font="Arial,8"
    $global:progressStatus.MaximumSize=New-Object System.Drawing.Size(380,66)
    $global:progressStatus.Location=New-Object System.Drawing.Point(5,264)
    $global:progressStatus.AutoSize=$true
    $global:form.Controls.Add($global:progressStatus)

    $global:progressBar=New-Object System.Windows.Forms.ProgressBar
    $global:progressBar.Location=New-Object System.Drawing.Point(5,294)
    $global:progressBar.Size=New-Object System.Drawing.Size(380,22)
    $global:progressBar.Minimum=0
    $global:progressBar.Step=1
    $global:progressBar.Style="Continuous"
    $global:progressBar.Visible=$false
    $global:form.Controls.Add($global:progressBar)

    $CancelButton=New-Object System.Windows.Forms.Button
    $CancelButton.Location=New-Object System.Drawing.Point(50,344)
    $CancelButton.Size=New-Object System.Drawing.Size(75,23)
    $CancelButton.Text=translate "Cancel"
    $CancelButton.DialogResult=[System.Windows.Forms.DialogResult]::Cancel
    $global:form.CancelButton=$CancelButton
    $global:form.Controls.Add($CancelButton)

    $OKButton=New-Object System.Windows.Forms.Button
    $OKButton.Location=New-Object System.Drawing.Point(240,344)
    $OKButton.Size=New-Object System.Drawing.Size(75,23)
    $OKButton.Text=translate "Start"
    $OKButton.Add_Click({actionFileProcess $global:textBox.Text})
    $global:form.Controls.Add($OKButton)

    if($global:debug -eq $false) {
        Hide-Console
    }

    $global:form.ShowDialog()
}

display "The script has finished!"
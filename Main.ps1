param (
   [string]$inputDirectory,
   [boolean]$cli,
   [string]$outName
)


function actionFileBrowse($path){
    $path=getRootPath $path
    $fileBox=New-Object System.Windows.Forms.OpenFileDialog
    $fileBox.InitialDirectory=$path
    $fileBox.Filter="pdf files (*.pdf)|*.pdf"
    $fileBox.Multiselect=$true
    $fileBox.Title="Select the file"
    $fileBox.CheckFileExists=$true
    
    $res=$fileBox.ShowDialog()
    if($res -eq "OK"){
        $global:textBox.Text=$fileBox.FileName
    }
}
function actionFolderBrowse($path){
    $path=getRootPath $path
    $folderBox=New-Object System.Windows.Forms.FolderBrowserDialog
    $folderBox.RootFolder="MyComputer"
    $folderBox.SelectedPath=$path
    $folderBox.ShowNewFolderButton=$true
    $folderBox.Description="Select the directory the script will search for pdf files"
    
    $res=$folderBox.ShowDialog()
    if($res -eq "OK"){
        $global:textBox.Text=$folderBox.SelectedPath
    }
}
function actionFileProcess($path){
    $global:progressStatus.Visible=$false
    $global:errorLabel.Visible=$false
    $global:successLabel.Visible=$false

    if(Test-Path -Path $path -PathType Any){
        if(Test-Path -Path $path -PathType Leaf){
            $listInput=@($path)
        }else{
            try{
                if($global:recurseCheckbox.Checked -eq $true){
                    $listInput=@((Get-ChildItem -Path $path -Recurse | Where-Object{$_.PSIsContainer -ne $true -and $_.Extension -eq ".pdf"} -ErrorAction Stop).FullName)
                }else{
                    $listInput=@((Get-ChildItem -Path $path | Where-Object{$_.PSIsContainer -ne $true -and $_.Extension -eq ".pdf"} -ErrorAction Stop).FullName)
                }
            }catch{
                $global:errorLabel.Text="Unable to list the item of the requesyed path!"
                $global:errorLabel.Visible=$true
                $global:form.Refresh()
                return
            }
        }
    }
    
    if($listInput -eq $null){
        $global:progressStatus.Visible=$false
        $global:errorLabel.Text="The requested path doesn't contain any pdf files!"
        $global:errorLabel.Visible=$true
        $global:errorLabel
        $global:form.Refresh()
        return
    }

    if(!(Test-Path -Path "$($global:currentLocation)\gswin64c.exe" -PathType Leaf)){
        $global:progressStatus.Visible=$false
        $global:errorLabel.Text="The gswin64c.exe executable is not present in $($global:currentLocation)!"
        $global:errorLabel.Visible=$true
        $global:errorLabel
        $global:form.Refresh()
        return
    }

    $global:progressBar.Value=0
    $global:progressBar.Visible=$true
    $global:progressStatus.Visible=$true
    $global:progressBar.Maximum=$listInput.Length
    $global:form.Refresh()

    for($i=0;$i -lt $listInput.Length;$i++){
        $global:progressStatus.Text="Processing $($listInput[$i].FullName)..."
        $global:form.Refresh()

        $res=&"$($global:currentLocation)\gswin64c.exe" "-sDEVICE=pdfwrite" "-dCompatibilityLevel=1.4" "-dPDFSETTINGS=/ebook" "-dNOPAUSE" "-dQUIET" "-dBATCH" "-sOutputFile=$($global:currentLocation)\$($global:outName)$($i).pdf" "$($listInput[$i].FullName)"
            
        if($res -ne "" -and $res -ne $null){
            $global:progressStatus.Visible=$false
            $global:errorLabel.Text="Error while processing $($listInput[$i].FullName): $($res)"
            $global:errorLabel.Visible=$true
            $global:errorLabel
            $global:form.Refresh()
            return
        }

        $global:progressBar.Value=$i+1
        $global:form.Refresh()
    }
    $global:progressStatus.Visible=$false
    $global:successLabel.Visible=$true
    $global:form.Refresh()
}

function getRootPath($path){
    if(Test-Path -Path $path -PathType Any){
        if(Test-Path -Path $path -PathType Leaf){
            try{
                $newPath=Split-Path -Path $path -ErrorAction Stop
            }catch{
                return $global:currentLocation
            }
            return $newPath
        }else{
            return $path
        }
    }else{
        return $global:currentLocation
    }
}

$global:currentLocation=Split-Path -Path $MyInvocation.MyCommand.Path
$global:applicationName="PDF Split Merge and Compress"
$global:defaultOutName="output"

if($outName -eq $null -or $outName -eq ""){
    $global:outName=$global:defaultOutName
}else{
    $global:outName=$outName
}

if($cli -eq $true){
    $global:cli=$true
}else{
    $global:cli=$false
}#todo

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
    exit
}

# Chargement des assemblies
[void][System.Reflection.Assembly]::LoadWithPartialName("System.Windows.Forms")
[void][System.Reflection.Assembly]::LoadWithPartialName("System.Drawing")

# Creation de la form principale
$global:form=New-Object Windows.Forms.Form
$global:form.FormBorderStyle=[System.Windows.Forms.FormBorderStyle]::FixedDialog
$global:form.MaximizeBox=$False
$global:form.MinimizeBox=$True
$global:form.Text=$global:applicationName
$global:form.Size=New-Object System.Drawing.Size(400,370)
$global:form.Icon=New-Object system.drawing.icon("$($global:currentLocation)\pdf.ico")
$global:form.StartPosition = "CenterScreen"

$label1=New-Object System.Windows.Forms.Label
$label1.Text="Select the file or folder you want to process pdf files"
$label1.Font="Arial,12"
$label1.Location=New-Object System.Drawing.Point(5,5)
$label1.Size=New-Object System.Drawing.Size(0,22)
$label1.AutoSize=$true
$global:form.Controls.Add($label1)

$global:textBox=New-Object System.Windows.Forms.TextBox
$global:textBox.Location=New-Object System.Drawing.Point(5,35)
$global:textBox.Size=New-Object System.Drawing.Size(380,22)
$global:textBox.AllowDrop=$true
$global:textBox.Text=$global:inputDirectory
$global:form.Controls.Add($global:textBox)

$fileButton=New-Object System.Windows.Forms.Button
$fileButton.Location=New-Object System.Drawing.Point(50,70)
$fileButton.Size=New-Object System.Drawing.Size(100,22)
$fileButton.Text="Browse File"
$fileButton.Add_Click({actionFileBrowse $global:textBox.Text})
$global:form.Controls.Add($fileButton)

$folderButton=New-Object System.Windows.Forms.Button
$folderButton.Location=New-Object System.Drawing.Point(240,70)
$folderButton.Size=New-Object System.Drawing.Size(100,22)
$folderButton.Text="Browse Folder"
$folderButton.Add_Click({actionFolderBrowse $global:textBox.Text})
$global:form.Controls.Add($folderButton)

$global:recurseCheckbox=New-Object System.Windows.Forms.CheckBox
$global:recurseCheckbox.Location=New-Object System.Drawing.Point(5,110)
$global:recurseCheckbox.Size=New-Object System.Drawing.Size(200,22)
$global:recurseCheckbox.Text="Browse recurse folder"
$global:form.Controls.Add($global:recurseCheckbox)

$global:successLabel=New-Object System.Windows.Forms.Label
$global:successLabel.Visible=$false
$global:successLabel.Text="The requested pdf files have been successfully processed!"
$global:successLabel.Font="Arial,8"
$global:successLabel.ForeColor="Green"
$global:successLabel.MaximumSize=New-Object System.Drawing.Size(380,66)
$global:successLabel.Location=New-Object System.Drawing.Point(5,200)
$global:successLabel.AutoSize=$true
$global:form.Controls.Add($global:successLabel)

$global:errorLabel=New-Object System.Windows.Forms.Label
$global:errorLabel.Visible=$false
$global:errorLabel.Font="Arial,8"
$global:errorLabel.ForeColor="Red"
$global:errorLabel.MaximumSize=New-Object System.Drawing.Size(380,66)
$global:errorLabel.Location=New-Object System.Drawing.Point(5,200)
$global:errorLabel.AutoSize=$true
$global:form.Controls.Add($global:errorLabel)

$global:progressStatus=New-Object System.Windows.Forms.Label
$global:progressStatus.Visible=$false
$global:progressStatus.Font="Arial,8"
$global:progressStatus.MaximumSize=New-Object System.Drawing.Size(380,66)
$global:progressStatus.Location=New-Object System.Drawing.Point(5,200)
$global:progressStatus.AutoSize=$true
$global:form.Controls.Add($global:progressStatus)

$global:progressBar=New-Object System.Windows.Forms.ProgressBar
$global:progressBar.Location=New-Object System.Drawing.Point(5,250)
$global:progressBar.Size=New-Object System.Drawing.Size(380,22)
$global:progressBar.Visible=$true
$global:progressBar.Minimum=0
$global:progressBar.Step=1
$global:progressBar.Style="Continuous"
$global:progressBar.Visible=$false
$global:form.Controls.Add($global:progressBar)

$CancelButton=New-Object System.Windows.Forms.Button
$CancelButton.Location=New-Object System.Drawing.Point(50,300)
$CancelButton.Size=New-Object System.Drawing.Size(75,23)
$CancelButton.Text="Cancel"
$CancelButton.DialogResult=[System.Windows.Forms.DialogResult]::Cancel
$global:form.CancelButton=$CancelButton
$global:form.Controls.Add($CancelButton)

$OKButton=New-Object System.Windows.Forms.Button
$OKButton.Location=New-Object System.Drawing.Point(240,300)
$OKButton.Size=New-Object System.Drawing.Size(75,23)
$OKButton.Text="Start"
$OKButton.Add_Click({actionFileProcess $global:textBox.Text})
$global:form.Controls.Add($OKButton)

$global:form.ShowDialog()
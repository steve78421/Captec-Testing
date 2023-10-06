#Author : Stephen Kopparapu
Write-Host "Welcome to my glorious script. How may I help you ?"
$section1 = 2
$section2 = $section1 + 30
$section3 = $section2 + 170
$section4 = $section3 + 100
$section5 = $section4 + 30

Add-Type -AssemblyName System.Windows.Forms

# Create the form
$form = New-Object System.Windows.Forms.Form
$form.Text = "Glorious Script v2.0"
$form.Size = New-Object System.Drawing.Size(400, 440)

#-------------- Section 1 ---------------------
# TITLE
$title = New-Object System.Windows.Forms.Label
$title.Text = "CAPTEC Americas"
$title.Font = New-Object System.Drawing.Font("",18,[System.Drawing.FontStyle]::Bold)
$title.Width = 300
$title.ForeColor = "Blue"
$title.Location = New-Object System.Drawing.Point(10, $section1)
$form.Controls.Add($title)

# my name
$myname = New-Object System.Windows.Forms.Label
$myname.Text = "V2.0. Written by Stephen K"
$myname.Font = New-Object System.Drawing.Font("",8,[System.Drawing.FontStyle]::Italic)
$myname.Width = 150
$myname.Location = New-Object System.Drawing.Point(10, ($section1+30))
$form.Controls.Add($myname)

#-------------- Section 2 ---------------------

# Delete Test User
$deltest = New-Object System.Windows.Forms.Button
$deltest.Text = "Delete Test User"
$deltest.Location = New-Object System.Drawing.Point(10,($Section2+30))
$deltest.Width = 120
$deltest.Add_Click({
    Remove-LocalUser -Name "test"
    Write-Host "Test User Deleted"
})
$form.Controls.Add($deltest)

# Computer Rename title
$Crename = New-Object System.Windows.Forms.Label
$Crename.Text = "Computer Name:"
$Crename.Font = New-Object System.Drawing.Font("",9,[System.Drawing.FontStyle]::Bold)
$Crename.Width = 110
$Crename.Location = New-Object System.Drawing.Point(10, ($Section2+60))
$form.Controls.Add($Crename)

# Text input box for Computer name
$Cnamebox = New-Object System.Windows.Forms.TextBox
$Cnamebox.Location = New-Object System.Drawing.Point(120, ($Section2+60))
$form.Controls.Add($Cnamebox)

# Rename Computer button
$CrenameButton = New-Object System.Windows.Forms.Button
$CrenameButton.Text = "Rename"
$CrenameButton.Location = New-Object System.Drawing.Point(10, ($Section2+85))
$CrenameButton.Width = 70
$CrenameButton.Add_Click({
    Rename-Computer $Cnamebox.Text
    Write-Host "Computer Renamed. Restart to make changes" 
})
$form.Controls.Add($CrenameButton)

# Restart Computer
$RestartB = New-Object System.Windows.Forms.Button
$RestartB.Text = "Restart Computer"
$RestartB.Location = New-Object System.Drawing.Point(10, ($Section2+120))
$RestartB.Width = 120
$RestartB.Add_Click({
    Restart-Computer
})
$form.Controls.Add($RestartB)

#-------------- Section 3 ---------------------
# file name
$filename = New-Object System.Windows.Forms.Label
$filename.Text = "System Number:"
$filename.Font = New-Object System.Drawing.Font("",9,[System.Drawing.FontStyle]::Bold)
$filename.Width = 100
$filename.Location = New-Object System.Drawing.Point(10, $Section3)
$form.Controls.Add($filename)

# Text input box for system number
$textbox = New-Object System.Windows.Forms.TextBox
$textbox.Location = New-Object System.Drawing.Point(120, $Section3)
$form.Controls.Add($textbox)

# Copy files to USB button
$functionButton = New-Object System.Windows.Forms.Button
$functionButton.Text = "Copy files to USB"
$functionButton.Location = New-Object System.Drawing.Point(10, ($Section3+30))
$functionButton.Width = 120
$functionButton.Add_Click({
    $name = $textbox.Text
    copyfiles
    Write-Host "Files copied. Confirm Files"
})
$form.Controls.Add($functionButton)



#-------------- Section 4 ---------------------
# Uninstall files
$functionButton = New-Object System.Windows.Forms.Button
$functionButton.Text = "Uninstall Softwares"
$functionButton.Location = New-Object System.Drawing.Point(10, $Section4)
$functionButton.Width = 120
$functionButton.Add_Click({
    uninstall
    Write-Host "Uninstalled burninTest, HWinfo and performancetest"
})
$form.Controls.Add($functionButton)

# CLear Logs
$functionButton = New-Object System.Windows.Forms.Button
$functionButton.Text = "Clear logs"
$functionButton.Location = New-Object System.Drawing.Point(130, $Section4)
#$functionButton.Width = 200
$functionButton.Add_Click({
    clearlogs
})
$form.Controls.Add($functionButton)

# Shutdown
$functionButton = New-Object System.Windows.Forms.Button
$functionButton.Text = "Shutdown Computer"
$functionButton.Location = New-Object System.Drawing.Point(10, $Section5)
$functionButton.Width = 120
$functionButton.Add_Click({
    Stop-Computer -Force
})
$form.Controls.Add($functionButton)

# Check COM ports
$functionButton = New-Object System.Windows.Forms.Button
$functionButton.Text = "Check COM ports"
$functionButton.Location = New-Object System.Drawing.Point(130, $Section5)
$functionButton.Width = 120
$functionButton.Add_Click({
    comPorts
})
$form.Controls.Add($functionButton)

#------------------------------------------------------------------

function uninstall {
        $files = @("C:\Program Files\HWiNFO64\unins000.exe","C:\Program Files\PerformanceTest\unins000.exe","C:\Program Files\BurnInTest\unins000.exe")
        foreach($file in $files){
        if(Test-Path $file){
	        Start-Process $file /SILENT
            Echo "removed $file"
        }
    }
    $documents = [environment]::GetFolderPath("mydocuments")
    Remove-Item $Documents'\PassMark' -Recurse -Force
    Remove-Item 'C:\Program Files\HWiNFO64\' -Recurse -Force
     
}

#------------------------------------------------------------------

function clearlogs {
    $logs = @("Application","Security","System","Setup","ForwardedEvents")
        foreach ($log in $logs){
	    wevtutil.exe cl $log
        Write-Host "Cleared $log"
    }
   
}

#------------------------------------------------------------------

function comPorts {
    
        $comports = [System.IO.Ports.SerialPort]::GetPortNames()

    Foreach($comPort in $comports){

        $baudRate = 2400

        $serialPort = New-Object System.IO.Ports.SerialPort

        $serialPort.PortName = $comPort
        $serialPort.BaudRate = $baudRate
        $serialPort.Parity = [System.IO.Ports.Parity]::None
        $serialPort.DataBits = 8
        $serialPort.StopBits = [System.IO.Ports.StopBits]::One

        $serialPort.Open()

        $dataToSend = "man door hand hook car door`n"

        #Convert the data to bytes and send it
        $bytesToSend = [System.Text.Encoding]::ASCII.GetBytes($dataToSend)
        $serialPort.Write($bytesToSend, 0, $bytesToSend.Length)
        $serialPort.ReadTimeout = 1000 

        $err = 0

        try{
            $receivedData = $serialPort.ReadLine()
            
        }catch{
            if( $_.Exception.Message ){
                $err = 1
            }
        }
    
    #Write-Host $receivedData
    $receivedData = ""

    if($err -eq 0){
        Write-Host "Com Port '$comPort' is working" -ForegroundColor Green
        #Write-Host $receivedData
        
    }else{
        Write-Host "Com Port '$comPort' is NOT working" -ForegroundColor Red
        $receivedData = ""

    }

        $serialPort.Close()

    }
   
}
#------------------------------------------------------------------

#Description : copy paste files and edit file
function copyfiles {
    <#  ----  Variables  ----  #>
    #$name = Read-Host -Prompt 'input number'
    $hostname = hostname
    $documents = [environment]::GetFolderPath("mydocuments")

    <#  ----  get usb drive letter  ----  #>
    $volumes = Get-WmiObject Win32_Volume | Where-Object { $_.DriveType -eq 2 }
    if ($volumes) {
        # Iterate through the volumes
        foreach ($volume in $volumes) {
            # Check if the volume is a removable drive
            if ($volume.DriveLetter -and $volume.DriveType -eq 2) {
        
                $usb =  $($volume.DriveLetter)
		    echo "The usb drive letter is $usb"
            }
        }
    }
    else {
        Write-Host "Insert a USB please."
    }


    <#  ----  copy files  ----  #>
    Copy-Item $documents\PassMark\BurnInTest\BurnInTestLog.log $usb\CA${name}.log
    Copy-Item -Path "C:\Program Files\HWiNFO64\${hostname}.HTM" $usb\CA${name}.HTM
    (Get-Content $usb\CA${name}.log).Replace("file  -  https://www.passmark.com","CA${name}")| Set-Content $usb\CA${name}.log

    <#  ----  editing the file  ----  #>

    $linesToDelete = @()

    #filepath to the file in usb
    $filePath = "$usb\CA${name}.log"

    [string[]]$filearray = Get-Content -Path $usb\CA${name}.log

    # set what to delete 
    $1a = "General:"
    $1b = ""
    $2a = "Memory devices:"
    $2b = "Graphics"
    $3a = "Optical drives"
    $3b = "Network"
    $4a = "USB xHCI Compliant Host Controller"
    $4b = "RESULT SUMMARY"

    #----convert keywords to line numbers
    $1an =  ($filearray | select-string $1a).LineNumber[0] -2
    $2an =  ($filearray | select-string $2a).LineNumber[0] - 1

    if ($2an -ge ($filearray | select-string $2b).LineNumber[0]){
        $2bn =  ($filearray | select-string $2b).LineNumber[1] - 3
    }else{
        $2bn =  ($filearray | select-string $2b).LineNumber[0] - 3
    }

    $3an =  ($filearray | select-string $3a).LineNumber[0] - 1
    $3bn = ($filearray | select-string $3b).LineNumber[0] - 1       # Network
    $4an = ($filearray | select-string $4a).LineNumber - 3
    $4bn = ($filearray | select-string $4b).LineNumber - 4

    #----create array of lines to delete
    $linesToDelete += $1an
    for ($i = $2an; $i -le $2bn; $i++) {  $linesToDelete += $i  }

    if( ($3bn - $3an) -le 2){
        $linesToDelete += $3an #optical disk
    } else {}
    
    #$linesToDelete += $3bn
    for ($i = $4an; $i -le $4bn; $i++) {  $linesToDelete += $i  }


    #----Remove lines from the file log
    $updatedContent = @()
    $notInArray = @()

    for ($i = 0; $i -lt $filearray.Count; $i++) {
        $found = $false
        foreach ($num in $linesToDelete) {
            if ($i -eq $num) {
                $found = $true
                break
            }
        }
        if (-not $found) {
            $notInArray += $i
            $updatedContent += $filearray[$i]
        }
    }

    #----save the file log
    $updatedContent | Set-Content -Path $usb\CA${name}Edited.log

    Write-Host "Done"

}


#------------------------------------------------------------------

# Show the form
$form.ShowDialog()
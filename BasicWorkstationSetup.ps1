# Install Slack via MSI (Replace with the actual MSI file path)
$slackMSIPath = "https://slack.com/ssb/download-win64-msi"
#Start-Process -FilePath "msiexec.exe" -ArgumentList "/i $slackMSIPath /qn /norestart" -Wait





# Install ClickUp (Replace with the actual installer link)

# Define the URL to the ClickUp installer
$clickUpInstallerURL = "https://desktop.clickup.com/windows?_gl=1*1aoi82m*_gcl_aw*R0NMLjE2ODkwMjc3NzYuQ2p3S0NBancySzZsQmhCWEVpd0E1Ump0Q1lickJWMGdlZ21yQS1qUUJIRDV1WkpSTDZ2SDlj2z2VOSEFab2VYOVo0S3BabDdyVS1DOE5ob0NLcDRRQXZEX0J3RQ..*_gcl_au*MjA4ODA1MTczNC4xNjg5NjM1NTY2"

# Define the path where you want to save the ClickUp installer
$installerFilePath = "C:\ClickUpInstaller.exe"

# Download the ClickUp installer
#Invoke-WebRequest -Uri $clickUpInstallerURL -OutFile $installerFilePath

# Check if the download was successful
if (Test-Path $installerFilePath) {
    Write-Host "ClickUp installer downloaded successfully to: $installerFilePath"
} else {
    Write-Host "Failed to download the ClickUp installer."
}

#Start-Process -FilePath $installerFilePath -ArgumentList "/SILENT" -Wait
#Remove-Item -Path $installerFilePath






# Set Start Menu to the left
Set-ItemProperty -Path "HKCU:\Software\Microsoft\Windows\CurrentVersion\Explorer\Advanced" -Name "StartMenuAlign" -Value 0

# Turn off search
Reg add "HKCU\Software\Microsoft\Windows\CurrentVersion\Search" /v "SearchboxTaskbarMode" /t REG_DWORD /d 0 /f

# Turn off widgets
Reg add "HKCU\Software\Microsoft\Windows\CurrentVersion\Explorer\Advanced" /v "UseActionCenterExperience" /t REG_DWORD /d 0 /f

# Turn off chat
Reg add "HKCU\Software\Microsoft\Windows\CurrentVersion\Explorer\Advanced" /v "EnableXamlStartMenu" /t REG_DWORD /d 0 /f





$imageURL = "https://firebasestorage.googleapis.com/v0/b/tricks-vue.appspot.com/o/Folsom%20Wallpaper%203440x1440.png?alt=media&token=aee8dea3-54ea-46ec-8c1f-4303af481a69"

# Define the path where you want to save the ClickUp installer
$imageFilePath = "C:\TricksFolsomBackground.png"

# Download the Image
#Invoke-WebRequest -Uri $imageURL -OutFile $imageFilePath

# Check if the image file exists
if (Test-Path $imageFilePath -PathType Leaf) {







    # Set the lock screen image using the registry
    #Set-ItemProperty -Path "HKLM:\SOFTWARE\Microsoft\Windows\CurrentVersion\PersonalizationCSP" -Name "LockScreenImageUrl" -Value $image



    #$regKey = 'HKLM:\SOFTWARE\Policies\Microsoft\Windows\Personalization'
    # create the key if it doesn't already exist
    #if (!(Test-Path -Path $regKey)) {
       #$null = New-Item -Path $regKey
    #}

    # now set the registry entry
    #Set-ItemProperty -Path $regKey -Name "LockScreenImage" -value $imageFilePath





    
    # Set Lock Screen Image
    $lockScreenRegKey = "HKLM:\SOFTWARE\Policies\Microsoft\Windows\Personalization"
    if (!(Test-Path -Path $lockScreenRegKey)) {
       $null = New-Item -Path $lockScreenRegKey
    }
    Set-ItemProperty -Path $lockScreenRegKey -Name LockScreenImage -Value $imageFilePath




    
    # Set Desktop Background Image
    $wallpaperRegKey = "HKCU:\Control Panel\Desktop"
    if (!(Test-Path -Path $wallpaperRegKey)) {
       $null = New-Item -Path $wallpaperRegKey
    }
    Set-ItemProperty -Path $wallpaperRegKey -Name Wallpaper -Value $imageFilePath




    # Refresh Desktop Background
    RUNDLL32.EXE USER32.DLL,UpdatePerUserSystemParameters ,1 ,True
    
    # Trigger a background sync for the lock screen image
    Rundll32.exe C:\Windows\System32\themecpl.dll,OpenPersonalizationPage










    Write-Host "Lock screen image has been set successfully."
} else {
    Write-Host "The specified image file does not exist."
}



Write-Host "Computer setup script completed."

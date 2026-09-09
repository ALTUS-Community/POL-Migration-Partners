# Install NuGet provider if needed
Install-PackageProvider -Name NuGet -Force

# Install winget installer script
Install-Script -Name winget-install -Force

# Install WinGet
winget-install -Force

# Verify
winget --version

# Install applications
# To Do

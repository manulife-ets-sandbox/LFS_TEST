function Install-Package($packageName, $fileName, $registryPath, $isExe) 
{
    $installed = (Get-ItemProperty $registryPath | Where { $_.DisplayName -eq $packageName }) -ne $nullclear;

    Write-Host "=========================================================";
    if(-Not $installed) 
    {        
        Write-Host "'$packageName' NOT installed.";
        Write-Host "Installing '$fileName' package....";

        $installer = Get-ChildItem -Filter $fileName;
        if($isEXE) 
        {
            & ".\$installer" /S /v/qn | Out-Null;   
        }
        else 
        {            
            msiexec /i $installer -q /lvx* install.log | Out-Null;
        }
        
        if ($LASTEXITCODE -gt 0) 
        {
            Write-Host "$fileName INSTALL ERROR CODE: $LASTEXITCODE";
        }
        else 
        {
            Write-Host "'$packageName' installed successfully.";
        }
    } 
    else 
    {
        Write-Host "'$packageName' is installed.";
    }
    Write-Host "=========================================================";
}

function Install-NETFramework($netVersion, $fileName, $registryPath, $isExe) 
{
    $installed = (Get-ItemProperty $registryPath).Release -ge $netVersion;

    Write-Host "========================================================="
    if(-Not $installed) 
    {        
        Write-Host ".NET Framework NOT installed.";
        Write-Host "Installing '$fileName' package....";

        $installer = Get-ChildItem -Filter $fileName;
        Start-Process $installer -ArgumentList "/q /norestart" -Wait -Verb RunAs | Out-Null;
        if ($LASTEXITCODE -gt 0) 
        {
            Write-Host "$fileName INSTALL ERROR CODE: $LASTEXITCODE";
        }
        else 
        {
            Write-Host ".NET Framework installed successfully.";
            Write-Host "*** NOTE: .NET Framework installation requires server restart! ***";
        }
    } 
    else 
    {
        Write-Host ".NET Framework is installed.";        
    }
    Write-Host "=========================================================";
}

$registryPathX86 = "HKLM:\Software\Microsoft\Windows\CurrentVersion\Uninstall\*"
$registryPathX64 = "HKLM:\Software\WOW6432Node\Microsoft\Windows\CurrentVersion\Uninstall\*"
$registryPathNET = "HKLM:\SOFTWARE\Microsoft\NET Framework Setup\NDP\v4\Full\"

Install-Package "Crystal Reports Basic Runtime for Visual Studio 2008" "CRRedist2008_x86.msi" $registryPathX64 $false;
Install-Package "SAP Crystal Reports runtime engine for .NET Framework (32-bit)" "CRRuntime_32bit_13_0_23.msi" $registryPathX64 $false;
Install-Package "Microsoft .NET Core Host - 2.2.4 (x86)" "dotnet-hosting-2.2.4-win.exe" $registryPathX64 $true;
Install-Package "IIS URL Rewrite Module 2" "rewrite_amd64.msi" $registryPathX86 $false;
Install-Package "Microsoft Visual C++ 2017 Redistributable (x86) - 14.12.25810" "vc_redist.x86.exe" $registryPathX64 $true;
Install-NETFramework 461814 "NDP472-KB4054530-x86-x64-AllOS-ENU.exe" $registryPathNET $true;
Install-Package "Microsoft OLE DB Driver for SQL Server" "msoledbsql.msi" $registryPathX64 $false

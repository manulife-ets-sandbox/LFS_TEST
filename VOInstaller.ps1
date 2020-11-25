## --------------------------------------------------------------
## Update the below variables to match your database/user setup
$batchUser = "manulifebatch"
$batchPass = "manulifebatch"
$serviceUser = "manulifeservice"
$servicePass = "manulifeservice"
$installFolder = "C:\VitalObjects"
$database = "cwisql2017.computerworkware.com"
$databaseUser = "NOTUSED"
$databasePassword = "NOTUSED"
$primaryDB = "ManulifeQAWA_Primary"
$warehouseDB = "ManulifeQAWA_Warehouse"
$reportsDB = "ManulifeQAWA_Reports"
$domainName = "compwork"
$websiteAuthMode = "2"
$dbConnectionType = "1"
$websiteIP = "192.168.52.35"
$webSiteDesc = "Default Web Site"
$logFolder = "C:\VitalObjectsLogs"
$outputFolder = "C:\VitalObjectsOutput"
$voEFTThreads = 16
$targetVDir = "VOWeb"
## --------------------------------------------------------------
function DeleteWithRetry($dir) {
    $deleteFailed = $false
    while ((Test-Path $dir)) {
        if ($deleteFailed) {
            Start-Sleep -s 5
        }
        Write-Host "Tyring to remove folder $dir"
        Remove-Item -Recurse -Force $dir -ErrorAction SilentlyContinue| Out-Null
        $deleteFailed = $true
    }
    Write-Host "Removed folder $dir"
}

function Uninstall-Program($programName) 
{
    Write-Output "--- Looking for '$programName' installation"
    $installed = (Get-ItemProperty "HKLM:\Software\WOW6432Node\Microsoft\Windows\CurrentVersion\Uninstall\*" | Where { $_.DisplayName -eq $programName }) -ne $nullclear;
    if($installed) 
    {
        Write-Output "--- Uninstalling '$programName'"
        $app = Get-WmiObject -Class Win32_Product -Filter "Name = '$programName'"
        if ($app) {
            $app.Uninstall()
        }
    }
    else 
    {
        Write-Host "--- '$programName' not installed."
    }
}

if ($PSVersionTable.PSVersion.Major -lt 5) 
{
    Write-Host "VOInstaller script requires powershell version 5.1 or greater, please upgrade."
    exit
}
else 
{
    if ($PSVersionTable.PSVersion.Minor -lt 1) 
    {
        Write-Host "VOInstaller script requires powershell version 5.1 or greater, please upgrade."
        exit
    }
    else 
    {
        Write-Host "Running Powershell version: " + $PSVersionTable.PSVersion
    }    
}

Write-Output "--- Shut down VO"
$comAdmin = New-Object -com ("COMAdmin.COMAdminCatalog.1")
if ($comAdmin) {
	$apps =$comAdmin.GetCollection("Applications");
	$apps.Populate();
	
	if($apps | ? { $_.Name -eq "Vital Objects" }) {
		$comAdmin.ShutDownApplication("Vital Objects") | Out-Null
	}
	
	if($apps | ? { $_.Name -eq "Vital Objects Services" }) {
		$comAdmin.ShutDownApplication("Vital Objects Services") | Out-Null
	}
	
	if($apps | ? { $_.Name -eq "Vital Objects Reporting" }) {
		$comAdmin.ShutDownApplication("Vital Objects Reporting") | Out-Null
	}
}

if (Get-Service CWI* -ErrorAction SilentlyContinue)
{
    Write-Output "--- Shut down CWI services"
    Get-Service CWI* | Stop-Service | Out-Null
}
if (Get-Service TaskGroupProcessor -ErrorAction SilentlyContinue)
{
    Write-Output "--- Shut down TaskGroupProcessorNode32 service"
    Get-Service TaskGroupProcessorNode32 | Stop-Service | Out-Null
}
if (Get-Service TaskGroupProcessorNode -ErrorAction SilentlyContinue)
{
    Write-Output "--- Shut down TaskGroupProcessorNode service"
    Get-Service TaskGroupProcessorNode | Stop-Service | Out-Null
}
if (Get-Service TaskGroupProcessor -ErrorAction SilentlyContinue)
{
    Write-Output "--- Shut down TaskGroupProcessor service"
    Get-Service TaskGroupProcessor | Stop-Service | Out-Null
}
if (Get-Service Lighthouse -ErrorAction SilentlyContinue)
{
    Write-Output "--- Shut down Lighthouse service"
    Get-Service Lighthouse | Stop-Service | Out-Null
}
if (Get-Service RepMan -ErrorAction SilentlyContinue)
{
    Write-Output "--- Shut down RepMan service"
    Get-Service RepMan | Stop-Service | Out-Null
}
if (Get-Service SALoaderHostService -ErrorAction SilentlyContinue)
{
    Write-Output "--- Shut down SALoaderHostService"
    Get-Service SALoaderHostService | Stop-Service | Out-Null
}

Uninstall-Program "Vital Objects - Member History"
Uninstall-Program "Vital Objects - VOAPI"
Uninstall-Program "Vital Objects - Task Group Processor"
Uninstall-Program "Vital Objects - Lighthouse"
Uninstall-Program "Vital Objects - SALoader"
Uninstall-Program "Vital Objects - Manulife"
Uninstall-Program "Vital Objects."

DeleteWithRetry $installFolder

Write-Output "--- Creating log folder"
New-Item -ItemType directory -Path $logFolder -force

Push-Location .

Write-Output "--- Creating output folder"
New-Item -ItemType directory -Path $outputFolder -force

Push-Location .

Write-Output "--- Add environment variables"
New-ItemProperty -Path "HKLM:\SYSTEM\CurrentControlSet\Control\Session Manager\Environment" -Name "VO_EFT_THREADS" -PropertyType "String" -Value $voEFTThreads -Force

Write-Output "--- Starting Core VO Installer"

$installer = Get-ChildItem -Filter CoreInstall_*.msi
msiexec /i $installer -q `
COM_USER_NAME="$serviceUser" `
COM_USER="$domainName\$serviceUser" `
COM_PASSWORD="$servicePass" `
COM_USER_DOMAIN="$domainName" `
BATCH_USER_NAME="$batchUser" `
BATCH_USER="$domainName\$batchUser" `
BATCH_PASSWORD="$batchPass" `
BATCH_USER_DOMAIN="$domainName" `
SERVICES_USER="$domainName\$batchUser" `
INSTALLLOCATION="$installFolder" `
WEBSITE_PORT="80" `
WEBSITE_IP="$websiteIP" `
WEBSITE_HEADER=`"`" `
WEBSITE_AUTH_MODE="$websiteAuthMode" `
WEBSITE_DESCRIPTION=`"$webSiteDesc`" `
TARGETVDIR=$targetVDir `
DB_DATASOURCE="$database" `
DB_USER="$databaseUser" `
DB_PASSWORD="$databasePassword" `
DB_MAIN_CATALOG="$primaryDB" `
DB_DW_CATALOG="$warehouseDB" `
DB_REPORTS_CATALOG="$reportsDB" `
DB_CONNECTION_TYPE="$websiteAuthMode" `
RABBIT_HOST="localhost" `
RABBIT_USER="guest" `
RABBIT_PASSWORD="guest" `
RABBIT_PREFETCH="1" `
RABBIT_TIMEOUT="60" `
MSXML_INSTALLED="1" `
VO_LOG_DIR="$logFolder" `
CRYSTAL105_INSTALLED="1" `
/lvx* install.log | Out-Null

if ($LASTEXITCODE -gt 0) {
    echo "--- Core VO INSTALL ERROR: $LASTEXITCODE"
}

Write-Output "--- Starting Manulife Installer"

$installer = Get-ChildItem -Filter ManulifeInstall_*.msi
msiexec /i $installer -q `
VOCORE_INSTALLED=1 `
BATCH_USER_NAME="$batchUser" `
BATCH_USER="$domainName\$batchUser" `
BATCH_PASSWORD="$batchPass" `
BATCH_USER_DOMAIN="$domainName" `
/lvx* ManulifeInstall.log | Out-Null

if ($LASTEXITCODE -gt 0) {
    Write-Output "--- Manulife INSTALL ERROR: $LASTEXITCODE"
}


Write-Output "--- Starting SALoader Installer"

$installer = Get-ChildItem -Filter SALoaderInstall_*.msi
msiexec /i $installer -q `
BATCH_USER_NAME="$batchUser" `
BATCH_USER="$domainName\$batchUser" `
BATCH_PASSWORD="$batchPass" `
BATCH_USER_DOMAIN="$domainName" `
INSTALLLOCATION="$installFolder" `
TARGETVDIR=$targetVDir `
WEBSITE_PORT="80" `
WEBSITE_IP=`"`" `
WEBSITE_HEADER=`"`" `
WEBSITE_DESCRIPTION=`"$webSiteDesc`" `
WEBSITE_AUTH_MODE="$websiteAuthMode" `
/lvx* SALoaderInstall.log | Out-Null

if ($LASTEXITCODE -gt 0) {
    Write-Output "--- SALoader INSTALL ERROR: $LASTEXITCODE"
}

Write-Output "--- Starting Lighthouse Installer"

$installer = Get-ChildItem -Filter LighthouseInstall_*.msi
msiexec /i $installer -q `
ACTORSYSTEM="vo-taskgroup" `
CLUSTER_IP="127.0.0.1" `
CLUSTER_PORT="4053" `
CLUSTER_SEEDS="akka.tcp://vo-taskgroup@127.0.0.1:4053" `
/lvx* LighhouseInstall.log | Out-Null

if ($LASTEXITCODE -gt 0) {
    Write-Output "--- Lighthouse INSTALL ERROR: $LASTEXITCODE"
}

Write-Output "--- Starting TaskGroupProcessor Installer"

$installer = Get-ChildItem -Filter CWITaskGroupProcessorInstall_*.msi

msiexec /i $installer -q `
BATCH_USER_NAME="$batchUser" `
BATCH_USER="$domainName\$batchUser" `
BATCH_PASSWORD="$batchPass" `
BATCH_USER_DOMAIN="$domainName" `
ACTORSYSTEM="vo-taskgroup" `
CLUSTER_IP="127.0.0.1" `
CLUSTER_PORT="4053" `
CLUSTER_SEEDS="akka.tcp://vo-taskgroup@127.0.0.1:4053" `
/lvx* CWITaskGroupProcessorInstall.log | Out-Null

if ($LASTEXITCODE -gt 0) {
    Write-Output "--- TaskGroupProcessor INSTALL ERROR: $LASTEXITCODE"
}

Write-Output "--- Starting VOAPI Installer"

$installer = Get-ChildItem -Filter VOAPIInstall_*.msi

msiexec /i $installer -q `
BATCH_USER_NAME="$batchUser" `
BATCH_USER="$domainName\$batchUser" `
BATCH_PASSWORD="$batchPass" `
BATCH_USER_DOMAIN="$domainName" `
WEBSITE_PORT="80" `
WEBSITE_IP=`"`" `
WEBSITE_DESCRIPTION=`"$webSiteDesc`" `
WEBSITE_AUTH_MODE="$websiteAuthMode" `
/lvx* VOAPIInstall.log | Out-Null

if ($LASTEXITCODE -gt 0) {
    Write-Output "--- VOAPI INSTALL ERROR: $LASTEXITCODE"
}

Write-Output "--- Starting Member History Installer"

$installer = Get-ChildItem -Filter MemberHistoryInstall_*.msi

msiexec /i $installer -q `
DB_CONNECTION_TYPE="$dbConnectionType" `
TARGETVDIR="MemberHistory" `
WEBSITE_PORT="80" `
WEBSITE_IP=`"`" `
WEBSITE_DESCRIPTION=`"$webSiteDesc`" `
WEBSITE_AUTH_MODE="$websiteAuthMode" `
WEBSITE_HEADER=`"`" `
APP_POOL_USER_NAME="$batchUsr" `
APP_POOL_USER_DOMAIN="$domainName" `
APP_POOL_USER_PASSWORD="$batchPass" `
/lvx* MemberHistoryInstall.log | Out-Null

if ($LASTEXITCODE -gt 0) {
    Write-Output "--- VOAPI INSTALL ERROR: $LASTEXITCODE"
}

Write-Output "--- Starting services"
Get-Service Lighthouse | Start-Service 
Get-Service TaskGroupProcessor | Start-Service 
Get-Service TaskGroupProcessorNode | Start-Service 
Get-Service TaskGroupProcessorNode32 | Start-Service 
Get-Service CWI* | Start-Service

Write-Output "--- Restart IIS"
Restart-Service W3svc | Out-Null
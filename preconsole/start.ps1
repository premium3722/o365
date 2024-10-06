
function func_install_module 
{
    param 
    (
        [array]$modules
    )
    foreach ($module in $modules) 
    {
        if (-not (Get-Module -Name $module -ListAvailable)) 
        {
            try
            {
            Install-Module -Name $module -Force -AllowClobber
            Write-Host "Modul $module installiert"
            }
            catch
            {
            Write-Host "Fehler beim installieren vom Modul $module  $_"  
            }
        }
        else
        {
            Write-Host "Modul $module ist bereits installiert"
        }
    }
}


function func_import_module
{
    param 
    (
        [array]$modulesimport
    )
      
foreach ($moduleimport in $modulesimport) 
{
  try
  {
    Import-Module $moduleimport
    Write-Host "Modul $moduleimport importiert"
  }
  catch
  {
    Write-Host "Fehler beim Import vom Modul $moduleimport  $_"
  }
}


}

function func_connect_MSOnline 
{
    try 
    {
        Connect-MsolService -ErrorAction Stop
        Write-Host "Erfolgreich mit MSOnline verbunden."
    } catch 
    {
        Write-Host "Fehler bei der Verbindung zu MSOnline: $_"
        exit 5
    }
    
}


# Verbindung zu Exchange Online herstellen
function func_connect_ExchangeOnline
{
    try {
        Connect-ExchangeOnline -ErrorAction Stop
        Write-Host "Erfolgreich mit Exchange Online verbunden."
    } catch {
        Write-Host "Fehler bei der Verbindung zu Exchange Online: $_"
        exit 5
    }
}

# Verbindung zu Azure AD herstellen
function func_connect_AzureAD
{
    try {
        Connect-AzureAD -ErrorAction Stop
        Write-Host "Erfolgreich mit Azure AD verbunden."
    } catch {
        Write-Host "Fehler bei der Verbindung zu Azure AD: $_"
        exit 5
    }
}




func_install_module -modules @("MSOnline", "ExchangeOnlineManagement", "AzureAD")
Write-Host ""
func_import_module -modulesimport @("MSOnline", "ExchangeOnlineManagement", "AzureAD")
Write-Host ""
func_connect_MSOnline
Write-Host ""
func_connect_ExchangeOnline
Write-Host ""
func_connect_AzureAD


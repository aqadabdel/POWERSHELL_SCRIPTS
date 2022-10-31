Add-Type -AssemblyName System.Windows.Forms
Add-Type -AssemblyName System.Drawing

$script:LogFile = "c:\SCRIPTS\LOG\PrinterDirectIPRemove.log"
$script:Version = "1.2.0"

# ADD PRINTERS PARAMETERS HERE
# WITH THE DRIVER FOLDER PATH

$liste_imprimantes = @(
  @{
        name = 'PDF'
        port_name = 'PDF'
        driver_name = 'Microsoft Print To PDF'
        driver_inf = "C:\DRIVERS\MICROSOFT_PDF\*.inf"
        url = ' '
   }
,
  @{
    name = 'HP_M402'
    port_name = 'HPM402SINFOGEST'
    driver_name = 'HP LaserJet Pro M402-M403 n-dne PCL 6'
    driver_inf = "C:\DRIVERS\HP_LaserJet_Pro_M402-M403_n-dne\*.inf"
    url = 'IP of the printer'
  }
)

######################################################
Restart-Service "Spooler"

function Get-ScriptName()
{
    $tmp = $MyInvocation.ScriptName.Substring($MyInvocation.ScriptName.LastIndexOf('\') + 1)
    $tmp.Substring(0, $tmp.Length - 4)
}
 
function Write-Log($Msg, [System.Boolean]$display = $true, $foregroundColor = '')
{
    $date = Get-Date -format MM/dd/yyyy
    $time = Get-Date -format HH:mm:ss
    Add-Content -Path $LogFile -Value ($date + " " + $time + "   " + $Msg)

    if ($display)
    {
        if ($foregroundColor -eq '')
        { Write-Host "$date $time   $Msg" }
        else
        { Write-Host "$date $time   $Msg" -ForegroundColor $foregroundColor }
    }
}
 
function Initialize-LogFile([System.Boolean]$reset = $false)
{
    try
    {
        #Check if file exists
        if (Test-Path -Path $LogFile)
        {
            #Check if file should be reset
            if ($reset)
            {
                Clear-Content $LogFile -ErrorAction SilentlyContinue
            }
        }
        else
        {
            #Check if file is a local file
            if ($LogFile.Substring(1, 1) -eq ':')
            {
                #Check if drive exists
                $driveInfo = [System.IO.DriveInfo]($LogFile)
                if ($driveInfo.IsReady -eq $false)
                {
                    Write-Log -Msg ($driveInfo.Name + " not ready.")
                }
                 
                #Create folder structure if necessary
                $Dir = [System.IO.Path]::GetDirectoryName($LogFile)
                if (([System.IO.Directory]::Exists($Dir)) -eq $false)
                {
                    $objDir = [System.IO.Directory]::CreateDirectory($Dir)
                    Write-Log -Msg ($Dir + " created.")
                }
            }
        }
        #Write header
        Write-Log "************************************************************"
        Write-Log "   Version $Version"
        Write-Log "************************************************************"
    }
    catch
    {
        Write-Log $_
    }
}

function Test-IsAdmin
{
     
    ([Security.Principal.WindowsPrincipal][Security.Principal.WindowsIdentity]::GetCurrent()).IsInRole([Security.Principal.WindowsBuiltInRole] "Administrator")
     
}
 
# MAIN SCRIPT PART
Initialize-LogFile
if (!(Test-IsAdmin))
{
    Write-Log "Please run this script with admin priviliges"
    exit
}

function Del_All_printers () {

    $printers = Get-CimInstance win32_printer  | Select-Object Name

    foreach ($printer in $printers) {
        try {
            Write-Host "REMOVING printer : $($printer.Name)" -ForegroundColor Green
            Write-Log "SUPPRESION DE IMPRIMANTE $($printer.Name)"
            Remove-Printer $printer.Name
        }
        catch {
            { 1:write-host "ERROR: suppression du port $($p.Name) impossible" }
        }
       
    }

    $ports = Get-CimInstance win32_tcpipprinterport  | Select-Object Name
    foreach ($p in $ports)
    {
      try {
        Write-Log "SUPPRESION DU PORT IMPRIMANTE $p"
        Write-Host "REMOVING printer port : $($p.Name)" -ForegroundColor Green
        Remove-PrinterPort -Name $p.Name
        
       }
       catch {
        { 1: write-host "ERROR: suppression du port $($p.Name) impossible" }
       }
    }
}
 
function Show_Ports ()
{
   [array]$ports  = Get-PrinterPort  | Where-Object { $_.PortMonitor -like "TCP*" } | Select-Object NAME
    
   foreach  ( $port in $ports ) {   write-host $port.Name
    }
}
 
function Show_Printers ()
{
   [array]$printers  = Get-Printer  
    
    foreach  ( $p in $printers ) {
        write-host $p.Name
    }
}

function Ajout_imprimante ( $imprimante ) {

    $printer = Get-CimInstance win32_printer | where-object  { $_.Name -eq "$($imprimante.name)" } |Select-Object Name
  
    if (  $printer   ) {

        Write-Log "Erreur: imprimante existante PRINTER: $($imprimante.name)"
        Write-Host "Erreur: imprimante existante" -ForegroundColor Green
        [System.Windows.Forms.Messagebox]::Show(   "Imprimante existante","Erreur", [System.Windows.Forms.MessageBoxButtons]::OK )
        return
    }
    else  {
        
        pnputil.exe -i -a $imprimante.driver_inf
        Write-Host "AJOUT imprimante $($imprimante.name)" -ForegroundColor Green
        Add-PrinterDriver -Name $imprimante.driver_name
        Add-PrinterPort -Name $imprimante.port_name -PrinterHostAddress $imprimante.url 
        Add-Printer -DriverName $imprimante.driver_name -Name $imprimante.name  -PortName $imprimante.port_name
   }
} 

$form = New-Object System.Windows.Forms.Form
$form.Text = 'AJOUT IMPRIMANTE'
$form.Size = New-Object System.Drawing.Size(300,200)
$form.StartPosition = 'CenterScreen'

$OKButton = New-Object System.Windows.Forms.Button
$OKButton.Location = New-Object System.Drawing.Point(15,120)
$OKButton.Size = New-Object System.Drawing.Size(75,23)
$OKButton.Text = 'Ajouter'
$OKButton.DialogResult = [System.Windows.Forms.DialogResult]::OK
$form.AcceptButton = $OKButton
$form.Controls.Add($OKButton)

$DelAllPintersButton = New-Object System.Windows.Forms.Button
$DelAllPintersButton.Location = New-Object System.Drawing.Point(85,120)
$DelAllPintersButton.Size = New-Object System.Drawing.Size(120,23)
$DelAllPintersButton.Text = 'Delete all printers'
$DelAllPintersButton.DialogResult = [System.Windows.Forms.DialogResult]::Ignore

$form.Controls.Add($DelAllPintersButton)

$CancelButton = New-Object System.Windows.Forms.Button
$CancelButton.Location = New-Object System.Drawing.Point(200,120)
$CancelButton.Size = New-Object System.Drawing.Size(75,23)
$CancelButton.Text = 'Cancel'
$CancelButton.DialogResult = [System.Windows.Forms.DialogResult]::Cancel
$form.CancelButton = $CancelButton

$form.Controls.Add($CancelButton)

$label = New-Object System.Windows.Forms.Label
$label.Location = New-Object System.Drawing.Point(10,20)
$label.Size = New-Object System.Drawing.Size(280,20)
$label.Text = 'Choisissez le ou les copieurs à installer:'

$form.Controls.Add($label)

$listBox = New-Object System.Windows.Forms.Listbox
$listBox.Location = New-Object System.Drawing.Point(10,40)
$listBox.Size = New-Object System.Drawing.Size(260,20)

$listBox.SelectionMode = 'MultiExtended'

foreach ( $l in $liste_imprimantes ){
       [void] $listBox.Items.Add( $l.name )
}

$listBox.Height = 70
$form.Controls.Add($listBox)
$form.Topmost = $true

$result = $form.ShowDialog()

if ($result -eq [System.Windows.Forms.DialogResult]::OK)
{
   foreach ($s in $listBox.SelectedItems )
   {
     foreach ($l in $liste_imprimantes )
     {
       if( $l.name -eq $s ) 
        {
            Ajout_imprimante( $l )
         #   Write-Log "AJOUT IMPRIMANTE $($l.name) PORT $($l.port_name)"
        }
      }   
    }
}
elseif ( $result -eq [System.Windows.Forms.DialogResult]::Ignore ) 
{
    $r = [System.Windows.Forms.Messagebox]::Show(  "Etes vous sûr de supprimer toutes les imprimantes", "Suppression", [System.Windows.Forms.MessageBoxButtons]::OKCancel ) 
    if ( $r -eq [System.Windows.Forms.DialogResult]::OK ) {
        Write-Host "DELETE ALL PRINTERS" -ForegroundColor Cyan 
        Del_All_printers
        [System.Windows.Forms.Messagebox]::Show(  "Toutes les imprimantes ont été supprimées", "Suppression", [System.Windows.Forms.MessageBoxButtons]::OK )
   }
   else {
    Exit
   }
}

Restart-Service Spooler 
Start-Sleep -Seconds 2
control printers

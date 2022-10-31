## AQAD Abdelaziz
## 31/10/2022
# Script to download and install microsoft PDF printer
#

$script:LogFile = "c:\DRIVERS\LOG\AddMicrosoftPdfPrinter.log"
$script:Version = "1.0.0"

$pdf_printer = @{
        name = 'PDF'
        port_name = 'PDF'
        driver_name = 'Microsoft Print To PDF'
        driver_inf = "C:\DRIVERS\MICROSOFT_PDF\*.inf"
        url = ' '
   }

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

function Download_Expand_PDF_Drivers ()
{
    $pdf_microsoft_driver_uri = "https://catalog.s.download.windowsupdate.com/d/msdownload/update/driver/drvs/2018/12/prnms009.1550_3431c7497d8814212223c6da78e2e6f707d106f1.cab"
    Invoke-WebRequest -Uri $pdf_microsoft_driver_uri -OutFile "pdf_printer.cab"

    $Dir = "C:\DRIVERS"
    if (([System.IO.Directory]::Exists($Dir)) -eq $false)
    {
        $objDir = [System.IO.Directory]::CreateDirectory($Dir)
        Write-Log -Msg ($Dir + " created.")

    }

    $Dir += "\MS_PDF_PRINTER"
    if (([System.IO.Directory]::Exists($Dir)) -eq $false)
    {
        $objDir = [System.IO.Directory]::CreateDirectory($Dir)
        Write-Log -Msg ($Dir + " created.")

    }

    expand "pdf_printer.cab" -F:* "$Dir\"
    $driver = "$Dir\*.inf"
    return $driver
}
function Add_MS_PDF_printer ( $pdf_printer) {

# SEE IF PDF PRINTER AND PDF PRINTER PORT ARE INSTALLED
# DELETE THEM IF IT IS THE CASE 

    $printer = Get-CimInstance win32_printer  | where-object  { $_.Name -eq "$($pdfprinter.Name)" } |Select-Object Name 

    if( $printer -eq $pdfprinter.name )
   {
        try {
            $message = "REMOVING EXISTING PDF PRINTER : $($printer.Name)" 
            write-host $message -ForegroundColor Green
            Write-Log $message
            Remove-Printer $printer.Name
        }
        catch {
            { 1:write-host "ERROR: Cannot remove printer $($printer.Name)" }
        }
       
    }

    $port  = Get-CimInstance win32_tcpipprinterport  | Where-Object {  $_.Name -eq  $pdfprinter.port_name }
    
    try {
         $message =  "DELETING PDF PRINTER PORT $($port.Name)"
         Write-Host $message -ForegroundColor Green
         Write-Log  $message   
         Remove-PrinterPort -Name $port.Name
       }
       catch {
        { 1: write-host "ERROR: Cannot remove pdf printer port $($port.Name)" }
       }
    
# ADD PDF PRINTER 
    try {
        $message = "ADDING PDF PRINTER $($pdfprinter.Name) DRIVER AND PROCEED TO INSTALL IT "
        Write-Host $message -ForegroundColor Green
        Write-Log $message
        
        $pdfprinter.driver_inf = Download_Expand_PDF_Drivers
        
        pnputil.exe -i -a $pdf_printer.driver_inf
        Add-PrinterDriver -Name $pdf_printer.driver_name
        Add-PrinterPort -Name $pdf_printer.port_name -PrinterHostAddress $imprimante.url 
        Add-Printer -DriverName $pdf_printer*.driver_name -Name $pdf_printer.name  -PortName $pdf_printer.port_name
    }
    catch {
        $error_message = "CANNOT INSTALL PDF PRINTER $($pdfprinter.Name)"
        Write-Host $error_message
        Write-Log $error_message
    }
}


Add_MS_PDF_printer( $pdf_printer )
Restart-Service Spooler 
Start-Sleep -Seconds 2
control printers

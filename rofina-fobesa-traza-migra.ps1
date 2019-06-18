# requires -version 2
<#
.SYNOPSIS
	Script de envío de archivos de Trazabilidad de Fobesa a Rofina
	
.DESCRIPTION
	Script de envío de archivos de Trazabilidad de Fobesa a Rofina
	
.INPUTS
	None
	
.OUTPUTS
	Función Write-Log reemplaza llamadas a Write-Host
	Write-Host se usa únicamente para las excepciones en conjunto con la función Write-Log
		
.NOTES
	Version:		1.1
	Author:			Santiago Platero
	Creation Date:	19/01/2018
	Purpose: 		Script de envío de archivos de Trazabilidad de Fobesa a Rofina
	Changelog: 		(18/JUN/2019) Cambio destino servidor Rofina por migración
	
.EXAMPLE
	>powershell -command ".'<absolute path>\rofina-fobesa-traza.ps1'"
#>

#---------------------------------------------------------[Inicializaciones]--------------------------------------------------------

# Inicializaciones de variables
$fileLocalDestinationMTV = "\\192.168.0.130\esql\Arrakis\trazabilidad\mtv.in\*"
$fileLocalDestinationCopiedMTV = "\\192.168.0.130\esql\Arrakis\trazabilidad\mtv.out\"
$fileLocalDestinationRFO = "\\192.168.0.130\esql\Arrakis\trazabilidad\rfo.in\*"
$fileLocalDestinationCopiedRFO = "\\192.168.0.130\esql\Arrakis\trazabilidad\rfo.out\"
$fileRemoteSource = "/interfaces/sistemasmv/enviados/*"
$fileRemoteDestination = "/entrada/*"
$fileMaskMTV = "L000011*"
$fileMaskRFO = "L000012*"
$dateFormat = "dd-MMM-yyyy HH:mm:ss"

#----------------------------------------------------------[Declaraciones]----------------------------------------------------------

# Información del script
$scriptVersion = "1.1"
$scriptName = $MyInvocation.MyCommand.Name

# Información de archivos de logs
$logPath = "C:\logs"
$logName = "$($scriptName).log"
$logFile = Join-Path -Path $logPath -ChildPath $logName

#-----------------------------------------------------------[Funciones]------------------------------------------------------------

#Función para hacer algo (?) de logueo
Function Write-Log
{
	Param ([string]$logstring)	
	Add-Content $logFile -value $logstring
}

Function Write-Exception
{
	Write-Host "[$(Get-Date -format $($dateFormat))] ERROR: $($_.Exception.Message)"
	Write-Log "[$(Get-Date -format $($dateFormat))] ERROR: $($_.Exception.Message)"
	Write-Log "[$(Get-Date -format $($dateFormat))] FIN DE EJECUCION DE $($scriptName)"
	Write-Log " "
	exit 1
}

#-----------------------------------------------------------[Ejecución]------------------------------------------------------------

# Primer control de errores: falta DDL, errores del servidor remoto, etc.
Write-Log "[$(Get-Date -format $($dateFormat))] INICIO DE EJECUCION DE $($scriptName)"
try
{
	# Carga de DLL de WinSCP .NET
	Add-Type -Path "c:\git\WinSCPnet.dll"

	# Configuración de opciones de sesión
	$fobesaSessionOptions = New-Object WinSCP.SessionOptions -Property @{
		Protocol = [WinSCP.Protocol]::Sftp
		HostName = "201.216.255.169"
		PortNumber = 20022
		UserName = "sistemasmv"
		SshHostKeyFingerprint = "ssh-rsa 2048 d0:cb:1f:b6:7a:07:54:1f:e7:40:28:09:d4:f9:04:66"
		SshPrivateKeyPath = "c:\git\fobesa.ppk"
	}
	
	$rofinaSessionOptions = New-Object WinSCP.SessionOptions -Property @{
		Protocol = [WinSCP.Protocol]::Sftp
		HostName = "ftp.rofina.com.ar"
		PortNumber = 9990
		UserName = "mverde"
		SshHostKeyFingerprint = "ssh-rsa 4096 02:fb:fe:20:41:1f:01:7d:00:a1:e7:a6:01:f9:ed:7c"
		SshPrivateKeyPath = "c:\git\rofina.ppk"
	}

	$fobesaSession = New-Object WinSCP.Session
	$rofinaSession = New-Object WinSCP.Session
	# Segundo control de errores: falta archivo, ruta incorrecta, errores de transferencia, etc.
	try
	{
		# Conexión y generamos log
		$fobesaSession.Open($fobesaSessionOptions)
		Write-Log "[$(Get-Date -format $($dateFormat))] Conectando a $($fobesaSessionOptions.UserName)@$($fobesaSessionOptions.HostName):$($fobesaSessionOptions.PortNumber)"
		$rofinaSession.Open($rofinaSessionOptions)
		Write-Log "[$(Get-Date -format $($dateFormat))] Conectando a $($rofinaSessionOptions.UserName)@$($rofinaSessionOptions.HostName):$($rofinaSessionOptions.PortNumber)"
	
		# Opciones de transferencia
		$fobesaMTVTransferOptions = New-Object WinSCP.TransferOptions
		$fobesaMTVTransferOptions.FileMask = $fileMaskMTV
		$fobesaRFOTransferOptions = New-Object WinSCP.TransferOptions
		$fobesaRFOTransferOptions.FileMask = $fileMaskRFO
					
		# 
		$fobesaTransferMTVFiles = $fobesaSession.GetFiles($fileRemoteSource, $fileLocalDestinationMTV, $True, $fobesaMTVTransferOptions)		
		$fobesaTransferRFOFiles = $fobesaSession.GetFiles($fileRemoteSource, $fileLocalDestinationRFO, $True, $fobesaRFOTransferOptions)
		$rofinaTransferMTVFiles = $rofinaSession.PutFiles($fileLocalDestinationMTV, $fileRemoteDestination)
		$rofinaTransferRFOFiles = $rofinaSession.PutFiles($fileLocalDestinationRFO, $fileRemoteDestination)
			
		# Arrojar cualquier error
		$fobesaTransferMTVFiles.Check()
		$fobesaTransferRFOFiles.Check()
		$rofinaTransferMTVFiles.Check()
		$rofinaTransferRFOFiles.Check()
	
		# Loopeamos por cada archivo que se transfiera
		foreach ($fobesaMTVTransfer in $fobesaTransferMTVFiles.Transfers)
		{
			$fobesaMTVFile = $fobesaMTVTransfer.FileName
		}
		# Antes de mandar al log, verificamos que la variable no sea nula
		if (!$fobesaMTVFile)
		{
			Write-Log "[$(Get-Date -format $($dateFormat))] Ningún archivo fue encontrado/transferido"
		}
		else
		{
			Write-Log "[$(Get-Date -format $($dateFormat))] Transferencia de $($fobesaMTVTransfer.FileName) exitosa"
		}
		
		# Loopeamos por cada archivo que se transfiera
		foreach ($fobesaRFOTransfer in $fobesaTransferRFOFiles.Transfers)
		{
			$fobesaRFOFile = $fobesaRFOTransfer.FileName
		}
		# Antes de mandar al log, verificamos que la variable no sea nula
		if (!$fobesaRFOFile)
		{
			Write-Log "[$(Get-Date -format $($dateFormat))] Ningún archivo fue encontrado/transferido"
		}
		else
		{
			Write-Log "[$(Get-Date -format $($dateFormat))] Transferencia de $($fobesaRFOTransfer.FileName) exitosa"
		}
		
		# Loopeamos por cada archivo que se transfiera
		foreach ($rofinaMTVTransfer in $rofinaTransferMTVFiles.Transfers)
		{
			$rofinaMTVFile = $rofinaMTVTransfer.FileName
		}
		# Antes de mandar al log, verificamos que la variable no sea nula
		if (!$rofinaMTVFile)
		{
			Write-Log "[$(Get-Date -format $($dateFormat))] Ningún archivo fue encontrado/transferido"
		}
		else
		{
			Write-Log "[$(Get-Date -format $($dateFormat))] Transferencia de $($rofinaMTVTransfer.FileName) exitosa"
			Move-Item $rofinaMTVTransfer.FileName $fileLocalDestinationCopiedMTV
		}
		
		# Loopeamos por cada archivo que se transfiera
		foreach ($rofinaRFOTransfer in $rofinaTransferRFOFiles.Transfers)
		{
			$rofinaRFOFile = $rofinaRFOTransfer.FileName
		}
		# Antes de mandar al log, verificamos que la variable no sea nula
		if (!$rofinaRFOFile)
		{
			Write-Log "[$(Get-Date -format $($dateFormat))] Ningún archivo fue encontrado/transferido"
		}
		else
		{
			Write-Log "[$(Get-Date -format $($dateFormat))] Transferencia de $($rofinaRFOTransfer.FileName) exitosa"
			Move-Item $rofinaRFOTransfer.FileName $fileLocalDestinationCopiedRFO
		}
	}
	# Impresión en caso de error en el segundo control
	catch
	{
		Write-Exception
	}
	finally
	{
		# Desconexión, limpieza
		$fobesaSession.Dispose()
		$rofinaSession.Dispose()
	}
	Write-Log "[$(Get-Date -format $($dateFormat))] FIN DE EJECUCION DE $($scriptName)"
	Write-Log " "
	exit 0
}
# Impresión en caso de error en el primer control
catch 
{
	Write-Exception
}

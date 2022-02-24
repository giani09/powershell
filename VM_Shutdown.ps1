##################################################################
#Reads the custom attribute "USV" and shutdown the VMs
#in the definied order
#V1.0
#(c) ***
##################################################################
#Import VMware Library

# use next line to set Password
# $credential = Get-Credential
# $credential | Export-CliXml -Path 'test...\cred.xml' 

Import-Module VMware.VimAutomation.core

#Variables
$wait = 120
# $VCServer="xxx"

##################################################################
#Main
##################################################################
Set-PowerCLIConfiguration -InvalidCertificateAction Ignore -Confirm:$false
Connect-VIServer -Server xxx.xxx.local -Credential (Import-Clixml 'test...\cred.xml')
# Connect-VIServer $VCServer -Credential (Import-Clixml 'test...\cred.xml')
$AllVM = get-vm

Write-Host "Beginn Shutdown of all VMs"

foreach ($VM in $AllVM)
	{
		$attribute = $vm | Get-Annotation -CustomAttribute "USV"
		if ($attribute.Value -match "USV-Shutdown-Group01")
			{
				Shutdown-VMGuest $vm.name -Confirm:$false
				Write-Host $VM.name "stopped"
			}
	}
Start-Sleep -Seconds $wait
foreach ($VM in $AllVM)
	{
		$attribute = $vm | Get-Annotation -CustomAttribute "USV"
		if ($attribute.Value -match "USV-Shutdown-Group02")
			{
				Shutdown-VMGuest $vm.name -Confirm:$false
				Write-Host $VM.name "stopped"
				
			}
	}
Start-Sleep -Seconds $wait
foreach ($VM in $AllVM)
	{
		$attribute = $vm | Get-Annotation -CustomAttribute "USV"
		if ($attribute.Value -match "USV-Shutdown-Group03")
			{
				Shutdown-VMGuest $vm.name -Confirm:$false
				Write-Host $VM.name "stopped"
				
			}
	}
	
Write-Host "Shutdown complete"
Disconnect-VIServer -Server * -Force -Confirm:$false

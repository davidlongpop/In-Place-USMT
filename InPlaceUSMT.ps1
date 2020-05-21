<#
.Synopsis
    In Place USMT Capture Restore
.DESCRIPTION
    Used to transfer user profiles from one on-domain system to another
.EXAMPLE
    InPlaceUSMT.ps1
.NOTES
    Created:	 2020-05-19
    Version:	 0.0.1
    Author - David Long
.LINK
    https://github.com/davidlongpop
.NOTES
.Updates
    v0.0.1
    Capture all user profiles from source system and restore them to Target system

    v0.0.2
    Select specific user profiles, or all profiles to capture from source system and restore them to target system
#>

$ErrorActionPreference = 'Stop'

#Set path var to where ever the script is currently running
$Path = Split-Path -parent $MyInvocation.MyCommand.Definition

#Establish logging dir
$LogDir = $Path + "\USMTJobs\Logs"

#Make logging directory if not exists
if(!(Test-Path $LogDir)){
    mkdir $LogDir
}

Start-Transcript -OutputDirectory $LogDir -NoClobber -IncludeInvocationHeader

#Import module for SCCM. normal path is "C:\Program Files (x86)\Microsoft Configuration Manager\AdminConsole\bin"
import-module "$Path\SCCM_Module\ConfigurationManager.psd1"

#user info
Write-host "All logs and reports are written to the current running directory" -f Yellow

#Declare all variables
$SourcePCName = ""
$TargetPCName = ""
$TargetPCDN = ""
$TargetDN = ""
$SCCMSiteCode = "PDX"
$SCCMCollectionID_USMTCapture = "PDX00192" #Win 10 USMT Capture
$SCCMCollectionID_USMTRestore = "PDX00193" #Win 10 USMT Restore
$SCCMServer = "POPSC"
$SCCMPackageID_USMTCapture = "PDX0030E" #USMT Package ID for targeting USMT capture jobs that need to be re-run
$SCCMPackageID_USMTRestore = "PDX0030F" #USMT Package ID for targeting USMT restore jobs that need to be re-run


#New Deployment Function
function InPlaceUSMT {

    param(
        [string]$SourcePCName,
        [string]$TargetPCName
    )

    #Set working directory for sccm commands to work
    Set-Location "PDX:"

    #Check if Source/Target PCs exist in SCCM and exit if they do not
    $SourcePCSCCMObj = Get-CMDevice -Name $SourcePCName
    $TargetPCSCCMObj = Get-CMDevice -Name $TargetPCName

    if(!($SourcePCSCCMObj)){
        Write-host "Source PC not found in SCCM. Exiting"
        Set-Location "C:"
        Exit 1
    }

    if(!($TargetPCSCCMObj)){
        Write-host "Target PC not found in SCCM. Exiting"
        Set-Location "C:"
        Exit 1
    }

    $SourcePCResID = $SourcePCSCCMObj.ResourceID
    $TargetPCResID = $TargetPCSCCMObj.ResourceID

    #Check is Source PC is online and wait to proceed until it is
    while(!(Test-Connection -ComputerName $SourcePCName -Count 1 -Quiet)){
        write-host "$SourcePCName is offline or connected via VPN. Script will ping every 5 minutes until $SourcePCName is online."
        start-sleep 300
    }

    #Check is Target PC is online and wait to proceed until it is
    while(!(Test-Connection -ComputerName $SourcePCName -Count 1 -Quiet)){
        write-host "$TargetPCName is offline or connected via VPN. Script will ping every 5 minutes until $TargetPCName is online."
        start-sleep 300
    }

    #Verify Source PC and Destination PC aren't currently part of an association and delete if they are
    $SourcePCAssociation = Get-CMComputerAssociation -SourceComputer $SourcePCName
    while($SourcePCAssociation){
        Write-Host $SourcePCAssociation.SourceName $SourcePCAssociation.RestoreName
        Remove-CMComputerAssociation -SourceComputer $SourcePCAssociation.SourceName -DestinationComputer $SourcePCAssociation.RestoreName -Force
    }
    $TargetPCAssociation = Get-CMComputerAssociation -DestinationComputer $TargetPCName
    while($TargetPCAssociation){
        Remove-CMComputerAssociation -SourceComputer $TargetPCAssociation.SourceName -DestinationComputer $TargetPCAssociation.RestoreName -Force
    }

    #Associate Source PC to Destination PC
    New-CMComputerAssociation -SourceComputer $SourcePCName -MigrationBehavior CaptureAndRestoreAllUserAccounts -DestinationComputer $TargetPCName
    Write-Host "SOURCE: $SourcePCName, associated to TARGET:$TargetPCName"

    #Check if Source PC is in USMT capture collection, add if not, and trigger manual policy retrieval/check
    if(!(Get-CMDeviceCollectionDirectMembershipRule -CollectionId $SCCMCollectionID_USMTCapture -ResourceId $SourcePCResID)){
        Add-CMDeviceCollectionDirectMembershipRule -CollectionId $SCCMCollectionID_USMTCapture -ResourceId $SourcePCResID
        Write-host "$SourcePCName successfully added to $SCCMCollectionID_USMTCapture at $(Get-Date -Format HH:mm:ss). Triggering device collection refresh." -f Yellow
        Invoke-CMClientNotification -DeviceName $SourcePCName -ActionType ClientNotificationRequestMachinePolicyNow, ClientNotificationRequestUsersPolicyNow
    }

    #Check if there's already a record of a capture that will interfere with the current capture attempt
    $CapRecExist= Get-CMDeploymentStatus | Where-Object CollectionID -eq $SCCMCollectionID_USMTCapture | Get-CMDeploymentStatusDetails | Where-Object DeviceName -eq $SourcePCName
    If($CapRecExist){
        Write-Warning "There's already a capture job record present for $SourcePCName. Script will attempt to clear existing job and retry capture."

        #Retrieve USMT Capture Job Scheduler History Object
        $WMIObjectRemove = Get-WmiObject -ComputerName $SourcePCName -Namespace "root\ccm\scheduler" -Class ccm_scheduler_history | where {$_.ScheduleID -like "*$SCCMPackageID_USMTCapture*"}

        #Remove USMT Capture Job Scheduler History Object
        Remove-WmiObject -InputObject $WMIObjectRemove -ErrorAction Stop

        #Restart CCM on Source PC
        Get-Service -Name "CCMExec" -ComputerName $SourcePCName | Restart-Service

        #Wait 60 seconds for USMT Deployment status to reinitiate
        start-sleep 60
    }

    #Monitor USMT Capture status
    Write-host "Getting USMT Capture status..."
    Start-Sleep 20
    Do{
        Write-host "Refreshing status in 60 seconds"
        Start-sleep 60

        $DeployStatus = Get-CMDeploymentStatus | Where-Object CollectionID -eq $SCCMCollectionID_USMTCapture | Get-CMDeploymentStatusDetails | Where-Object DeviceName -eq $SourcePCName

        If($DeployStatus.StatusType -eq "5"){
            Write-host "USMT Capture job failed. Will re-initiate USMT when Source PC is back online. `nStatus: $($DeployStatus.StatusDescription)" -f Red
            #Notify Project team that Source PC USMT Capture Job has failed and that the script will attempt to reinitiate capture when source is back online
			Send-MailMessage -From "Windows10ProjectTeam@portofportland.com" -To "Windows10ProjectTeam@portofportland.com" -Subject "$SourcePCName USMT Capture has failed. Attempting to reinitiate capture job" -SmtpServer "portexlb.pop.portptld.com"

            #Check if source pc is online and wait to proceed until it is
            while(!(Test-Connection -ComputerName $SourcePCName -Count 1 -Quiet)){
	            write-host "$SourcePCName is offline or connected via VPN. Script will ping every 5 minutes and reinitiate capture job when $SourcePCName is online."
	            start-sleep 300
	            $pingStatus = Test-Connection -ComputerName $SourcePCName -Count 1 -Quiet
            }

            #Retrieve USMT Capture Job Scheduler History Object
            $WMIObjectRemove = Get-WmiObject -ComputerName $SourcePCName -Namespace "root\ccm\scheduler" -Class ccm_scheduler_history | where {$_.ScheduleID -like "*$SCCMPackageID_USMTCapture*"}

            #Remove USMT Capture Job Scheduler History Object
            Remove-WmiObject -InputObject $WMIObjectRemove -ErrorAction Stop

            #Restart CCM on Source PC
            Get-Service -Name "CCMExec" -ComputerName $SourcePCName | Restart-Service

            #Wait 60 seconds for USMT Deployment status to reinitiate
            start-sleep 60

        }
        If($DeployStatus.StatusType -eq "4"){
            Write-host "$SourcePCName is in Capture collection - `nNothing to do now but wait for SCCM. `nStatus: $($DeployStatus.StatusDescription)." -f Yellow
        }
        If($DeployStatus.StatusType -eq "2"){
            Write-host "USMT Capture in progress - Check back in a little while. `nStatus: $($DeployStatus.StatusDescription)." -f Yellow
        }
        If(!($DeployStatus.StatusType)){
            Write-Warning "Can't find a capture job associated with $SourcePCName."
        }

        $DeployStatus = Get-CMDeploymentStatus | Where-Object CollectionID -eq $SCCMCollectionID_USMTCapture | Get-CMDeploymentStatusDetails | Where-Object DeviceName -eq $SourcePCName    
        
    }Until($DeployStatus.StatusType -eq "1" )
    
    #Check if Target PC is in USMT restore collection, add if not, and trigger manual policy retrieval/check
    if(!(Get-CMDeviceCollectionDirectMembershipRule -CollectionId $SCCMCollectionID_USMTRestore -ResourceId $TargetPCResID)){
        Add-CMDeviceCollectionDirectMembershipRule -CollectionId $SCCMCollectionID_USMTRestore -ResourceId $TargetPCResID
        Write-host "$TargetPCName successfully added to $SCCMCollectionID_USMTRestore at $(Get-Date -Format HH:mm:ss). Triggering device collection refresh." -f Yellow
        Invoke-CMClientNotification -DeviceName $SourcePCName -ActionType ClientNotificationRequestMachinePolicyNow, ClientNotificationRequestUsersPolicyNow
    }

    #Check if there's already a record of a restore that will interfere with the current restore attempt
    $RestoreRecExist= Get-CMDeploymentStatus | Where-Object CollectionID -eq $SCCMCollectionID_USMTRestore | Get-CMDeploymentStatusDetails | Where-Object DeviceName -eq $TargetPCName
    If($RestoreRecExist){
        Write-Warning "There's already a restore job record present for $TargetPCName. Script will attempt to clear existing job and retry restore."

        #Retrieve USMT Restore Job Scheduler History Object
        $USMTRestoreObject = Get-WmiObject -ComputerName $TargetPCName -Namespace "root\ccm\scheduler" -Class ccm_scheduler_history | where {$_.ScheduleID -like "*$SCCMPackageID_USMTRestore*"}

        #Remove USMT Restore Job Scheduler History Object
        Remove-WmiObject -InputObject $USMTRestoreObject -ErrorAction Stop

        #Restart CCM on Target PC
        Get-Service -Name "CCMExec" -ComputerName $TargetPCName | Restart-Service

        #Wait 60 seconds for USMT Deployment status to reinitiate
        start-sleep 60
    }

    #Monitor USMT Restore status
    Write-host "Getting USMT Capture status..."
    Start-Sleep 20
    Do{
        Write-host "Refreshing status in 60 seconds"
        Start-sleep 60

        $RestoreStatus = Get-CMDeploymentStatus | Where-Object CollectionID -eq $SCCMCollectionID_USMTRestore | Get-CMDeploymentStatusDetails | Where-Object DeviceName -eq $TargetPCName

        If($RestoreStatus.StatusType -eq "5"){
            Write-host "USMT Restore job failed. Will re-initiate USMT if Target PC is back online. `nStatus: $($RestoreStatus.StatusDescription)" -f Red
            #Notify Project team that Target PC USMT Restore Job has failed and that the script will attempt to reinitiate restore when target is back online
			Send-MailMessage -From "Windows10ProjectTeam@portofportland.com" -To "Windows10ProjectTeam@portofportland.com" -Subject "$TargetPCName USMT Restore has failed. Attempting to reinitiate restore job" -SmtpServer "portexlb.pop.portptld.com"

            #Check if source pc is online and wait to proceed until it is
            while(!(Test-Connection -ComputerName $TargetPCName -Count 1 -Quiet)){
	            write-host "$TargetPCName is offline or connected via VPN. Script will ping every 5 minutes and reinitiate capture job when $TargetPCName is online."
	            start-sleep 300
	            $pingStatus = Test-Connection -ComputerName $TargetPCName -Count 1 -Quiet
            }

            #Retrieve USMT Capture Job Scheduler History Object
            $USMTRestoreObject = Get-WmiObject -ComputerName $TargetPCName -Namespace "root\ccm\scheduler" -Class ccm_scheduler_history | where {$_.ScheduleID -like "*$SCCMPackageID_USMTRestore*"}

            #Remove USMT Capture Job Scheduler History Object
            Remove-WmiObject -InputObject $USMTRestoreObject -ErrorAction Stop

            #Restart CCM on Source PC
            Get-Service -Name "CCMExec" -ComputerName $TargetPCName | Restart-Service

            #Wait 60 seconds for USMT Deployment status to reinitiate
            start-sleep 60

        }
        If($RestoreStatus.StatusType -eq "4"){
            Write-host "$TargetPCName is in Capture collection - `nNothing to do now but wait for SCCM. `nStatus: $($RestoreStatus.StatusDescription)." -f Yellow
        }
        If($RestoreStatus.StatusType -eq "2"){
            Write-host "USMT Capture in progress - Check back in a little while. `nStatus: $($RestoreStatus.StatusDescription)." -f Yellow
        }
        If(!($RestoreStatus.StatusType)){
            Write-Warning "Can't find a capture job associated with $TargetPCName."
        }

        $RestoreStatus = Get-CMDeploymentStatus | Where-Object CollectionID -eq $SCCMCollectionID_USMTRestore | Get-CMDeploymentStatusDetails | Where-Object DeviceName -eq $TargetPCName    
        
    }Until($RestoreStatus.StatusType -eq "1" )

    #Success    
    Write-host  "Job finished. Status: $JobStatus" -f Magenta
	Send-MailMessage -From "Windows10ProjectTeam@portofportland.com" -To "Windows10ProjectTeam@portofportland.com" -Subject "Successfully copied profiles to $TargetPCName" -SmtpServer "portexlb.pop.portptld.com"
    Read-host "Press any key to exit"
    Stop-Transcript
    Set-Location "C:"
    Exit

}

#Establish Source/Target PCs for Capture Job
    $SourcePCName = Read-Host "Enter the SOURCE computer name"
    $TargetPCName = Read-Host "Enter TARGET computer name"

#Verify System exists in SCCM
#else{
#    Write-host "Invalid input."    
#}


#User Profile selection for v0.0.2 (Pending...)
<#[cmdletbinding()]
param (
[parameter(ValueFromPipeline=$true,ValueFromPipelineByPropertyName=$true)]
[string[]]$ComputerName = $env:computername
)            
 
foreach ($Computer in $ComputerName) {
 $Profiles = Get-WmiObject -Class Win32_UserProfile -Computer $Computer -ea 0
 foreach ($profile in $profiles) {
  try {
      $objSID = New-Object System.Security.Principal.SecurityIdentifier($profile.sid)
      $objuser = $objsid.Translate([System.Security.Principal.NTAccount])
      $objusername = $objuser.value
  } catch {
        $objusername = $profile.sid
  }
  switch($profile.status){
   1 { $profileType="Temporary" }
   2 { $profileType="Roaming" }
   4 { $profileType="Mandatory" }
   8 { $profileType="Corrupted" }
   default { $profileType = "LOCAL" }
  }
  $User = $objUser.Value
  $ProfileLastUseTime = ([WMI]"").Converttodatetime($profile.lastusetime)
  $OutputObj = New-Object -TypeName PSobject
  $OutputObj | Add-Member -MemberType NoteProperty -Name ComputerName -Value $Computer.toUpper()
  $OutputObj | Add-Member -MemberType NoteProperty -Name ProfileName -Value $objusername
  $OutputObj | Add-Member -MemberType NoteProperty -Name ProfilePath -Value $profile.localpath
  $OutputObj | Add-Member -MemberType NoteProperty -Name ProfileType -Value $ProfileType
  $OutputObj | Add-Member -MemberType NoteProperty -Name IsinUse -Value $profile.loaded
  $OutputObj | Add-Member -MemberType NoteProperty -Name IsSystemAccount -Value $profile.special
  $OutputObj
  
 }
}

#>


#Make variables upper-case
$SourcePCName = $SourcePCName.toUpper()
$TargetPCName = $TargetPCName.toUpper()

#Call Function
InPlaceUSMT $SourcePCName $TargetPCName
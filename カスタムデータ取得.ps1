#　https://gallery.technet.microsoft.com/scriptcenter/38e9bcc0-6dfd-489f-b360-5f2e1be6240e
#　https://technet.microsoft.com/ja-jp/magazine/gg293110.aspx


 Function Set-BufferSize{

$global:bufferSize = ( Get-Host).UI.RawUI.BufferSize
$global:bufferSize.Width = 2000
$global:bufferSize.Height = 5000
(Get-Host).UI. RawUI.BufferSize = $bufferSize


}

Set-BufferSize



<#

#gwmiをリモート実行したい場合は下記変数と　コマンドの後ろに -credential $credentialをつけること


$PSuser = "system3g\administrator"
$Password = "P@ssw0rd"
$PSpassword = $Password | ConvertTo-SecureString -asPlainText -Force
$Credential = New-Object System.Management.Automation.PSCredential($PSuser,$PSpassword)
####################################################
#>



function Get-Inventory
{
   [CmdletBinding ()]
   Param(
       [Parameter(Mandatory =$true,
                 ValueFromPipeline =$true,
                 ValueFromPipelineByPropertyName =$true) ]
       [string] $computername
   )
   Process {

    #各パラメーターの入れ物を変数として宣言
      $os = gwmi win32_operatingsystem -computername $computername
      $NetAdapterConf = gwmi Win32_NetworkAdapterConfiguration -computername $computername | ? {$_.IPEnabled}  
      $PageFile = gwmi Win32_PageFile -computername $computername
      $Comp = gwmi Win32_ComputerSystem -computername $computername
      $Recover = gwmi Win32_OSRecoveryConfiguration -computername $computername

      $WinUpdate = gp "HKLM:\SOFTWARE\Microsoft\Windows\CurrentVersion\WindowsUpdate\Auto Update"
      $RDP1 = gp "HKLM:\SYSTEM\CurrentControlSet\Control\Terminal Server"
      $RDP2 = gp "HKLM:\SYSTEM\CurrentControlSet\Control\Terminal Server\WinStations\RDP-Tcp"
      $IE_ESC_Admin = gp "HKLM:\SOFTWARE\Microsoft\Active Setup\Installed Components\{A509B1A7-37EF-4b3f-8CFC-4F3A74704073}"
      $IE_ESC_User =  gp "HKLM:\SOFTWARE\Microsoft\Active Setup\Installed Components\{A509B1A8-37EF-4b3f-8CFC-4F3A74704073}"
      $Firewall = Get-NetFirewallProfile
      $UAC = gp "HKLM:\SOFTWARE\Microsoft\Windows\CurrentVersion\Policies\System"
      $Power_Settings = powercfg.exe /GETACTIVESCHEME
      $SNP1 = Get-NetOffloadGlobalSetting
      
      #w2k8R2用のみ必要（w2k12以降は規定で無効）
      #$SNP2 = gp "HKLM:SYSTEM\CurrentControlSet\Services\Tcpip\Parameters"

      $CD = gwmi Win32_CDROMDrive -computername $computername
      $Disk1 = gwmi Win32_DiskDrive -computername $computername
      $Disk2 = gwmi Win32_DiskDriveToDiskPartition -computername $computername
      $Disk3 = gwmi Win32_LogicalDiskToPartition -computername $computername
      $Service_BITS  = Get-WmiObject -Query "select * from win32_service where name='BITS'"

      


     　#Windows 2008 R2 サービスの変数宣言
      <# 
        $Service_FontCache  = Get-WmiObject -Query "select * from win32_service where name='FontCache'"
        $Service_WinRM  = Get-WmiObject -Query "select * from win32_service where name='WinRM'"
        $Service_PlugPlay  = Get-WmiObject -Query "select * from win32_service where name='PlugPlay'"
        $Service_ALG  = Get-WmiObject -Query "select * from win32_service where name='ALG'"
        $Service_AudioEndpointBuilder  = Get-WmiObject -Query "select * from win32_service where name='AudioEndpointBuilder'"
        $Service_AudioSrv  = Get-WmiObject -Query "select * from win32_service where name='AudioSrv'"
        $Service_defragsvc  = Get-WmiObject -Query "select * from win32_service where name='defragsvc'"
        $Service_dot3svc  = Get-WmiObject -Query "select * from win32_service where name='dot3svc'"
        $Service_EapHost  = Get-WmiObject -Query "select * from win32_service where name='EapHost'"
        $Service_fdPHost  = Get-WmiObject -Query "select * from win32_service where name='fdPHost'"
        $Service_hkmsvc  = Get-WmiObject -Query "select * from win32_service where name='hkmsvc'"
        $Service_iphlpsvc  = Get-WmiObject -Query "select * from win32_service where name='iphlpsvc'"
        $Service_lltdsvc  = Get-WmiObject -Query "select * from win32_service where name='lltdsvc'"
        $Service_MMCSS  = Get-WmiObject -Query "select * from win32_service where name='MMCSS'"
        $Service_MSDTC  = Get-WmiObject -Query "select * from win32_service where name='MSDTC'"
        $Service_MSiSCSI  = Get-WmiObject -Query "select * from win32_service where name='MSiSCSI'"
        $Service_PolicyAgent  = Get-WmiObject -Query "select * from win32_service where name='PolicyAgent'"
        $Service_RasAuto  = Get-WmiObject -Query "select * from win32_service where name='RasAuto'"
        $Service_SCardSvr  = Get-WmiObject -Query "select * from win32_service where name='SCardSvr'"
        $Service_SCPolicySvc  = Get-WmiObject -Query "select * from win32_service where name='SCPolicySvc'"
        $Service_Spooler  = Get-WmiObject -Query "select * from win32_service where name='Spooler'"
        $Service_SstpSvc  = Get-WmiObject -Query "select * from win32_service where name='SstpSvc'"
        $Service_TapiSrv  = Get-WmiObject -Query "select * from win32_service where name='TapiSrv'"
        $Service_TrkWks  = Get-WmiObject -Query "select * from win32_service where name='TrkWks'"
        $Service_Wecsvc  = Get-WmiObject -Query "select * from win32_service where name='Wecsvc'"
        $Service_wercplsupport  = Get-WmiObject -Query "select * from win32_service where name='wercplsupport'"
        $Service_WerSvc  = Get-WmiObject -Query "select * from win32_service where name='WerSvc'"
    　#>


      #Windows 2012 サービスの変数宣言
      <#
        $Service_ALG  = Get-WmiObject -Query "select * from win32_service where name='ALG'"
        $Service_AudioEndpointBuilder  = Get-WmiObject -Query "select * from win32_service where name='AudioEndpointBuilder'"
        $Service_Audiosrv  = Get-WmiObject -Query "select * from win32_service where name='Audiosrv'"
        $Service_defragsvc  = Get-WmiObject -Query "select * from win32_service where name='defragsvc'"
        $Service_DeviceAssociationService  = Get-WmiObject -Query "select * from win32_service where name='DeviceAssociationService'"
        $Service_dot3svc  = Get-WmiObject -Query "select * from win32_service where name='dot3svc'"
        $Service_DsmSvc  = Get-WmiObject -Query "select * from win32_service where name='DsmSvc'"
        $Service_Eaphost  = Get-WmiObject -Query "select * from win32_service where name='Eaphost'"
        $Service_fdPHost  = Get-WmiObject -Query "select * from win32_service where name='fdPHost'"
        $Service_hkmsvc  = Get-WmiObject -Query "select * from win32_service where name='hkmsvc'"
        $Service_iphlpsvc  = Get-WmiObject -Query "select * from win32_service where name='iphlpsvc'"
        $Service_KPSSVC  = Get-WmiObject -Query "select * from win32_service where name='KPSSVC'"
        $Service_lltdsvc  = Get-WmiObject -Query "select * from win32_service where name='lltdsvc'"
        $Service_MMCSS  = Get-WmiObject -Query "select * from win32_service where name='MMCSS'"
        $Service_MSDTC  = Get-WmiObject -Query "select * from win32_service where name='MSDTC'"
        $Service_MSiSCSI  = Get-WmiObject -Query "select * from win32_service where name='MSiSCSI'"
        $Service_PolicyAgent  = Get-WmiObject -Query "select * from win32_service where name='PolicyAgent'"
        $Service_PrintNotify  = Get-WmiObject -Query "select * from win32_service where name='PrintNotify'"
        $Service_RasAuto  = Get-WmiObject -Query "select * from win32_service where name='RasAuto'"
        $Service_SCPolicySvc  = Get-WmiObject -Query "select * from win32_service where name='SCPolicySvc'"
        $Service_Spooler  = Get-WmiObject -Query "select * from win32_service where name='Spooler'"
        $Service_SstpSvc  = Get-WmiObject -Query "select * from win32_service where name='SstpSvc'"
        $Service_svsvc  = Get-WmiObject -Query "select * from win32_service where name='svsvc'"
        $Service_TapiSrv  = Get-WmiObject -Query "select * from win32_service where name='TapiSrv'"
        $Service_Themes  = Get-WmiObject -Query "select * from win32_service where name='Themes'"
        $Service_TrkWks  = Get-WmiObject -Query "select * from win32_service where name='TrkWks'"
        $Service_UALSVC  = Get-WmiObject -Query "select * from win32_service where name='UALSVC'"
        $Service_vmicheartbeat  = Get-WmiObject -Query "select * from win32_service where name='vmicheartbeat'"
        $Service_vmickvpexchange  = Get-WmiObject -Query "select * from win32_service where name='vmickvpexchange'"
        $Service_vmicrdv  = Get-WmiObject -Query "select * from win32_service where name='vmicrdv'"
        $Service_vmicshutdown  = Get-WmiObject -Query "select * from win32_service where name='vmicshutdown'"
        $Service_vmictimesync  = Get-WmiObject -Query "select * from win32_service where name='vmictimesync'"
        $Service_vmicvss  = Get-WmiObject -Query "select * from win32_service where name='vmicvss'"
        $Service_Wecsvc  = Get-WmiObject -Query "select * from win32_service where name='Wecsvc'"
        $Service_wercplsupport  = Get-WmiObject -Query "select * from win32_service where name='wercplsupport'"
        $Service_WerSvc  = Get-WmiObject -Query "select * from win32_service where name='WerSvc'"
      #>


      #Windows 2012 R2 サービスの変数宣言
      <#
        $Service_BITS  = Get-WmiObject -Query "select * from win32_service where name='BITS'"
        $Service_IKEEXT  = Get-WmiObject -Query "select * from win32_service where name='IKEEXT'"
        $Service_ALG  = Get-WmiObject -Query "select * from win32_service where name='ALG'"
        $Service_AudioEndpointBuilder  = Get-WmiObject -Query "select * from win32_service where name='AudioEndpointBuilder'"
        $Service_Audiosrv  = Get-WmiObject -Query "select * from win32_service where name='Audiosrv'"
        $Service_defragsvc  = Get-WmiObject -Query "select * from win32_service where name='defragsvc'"
        $Service_DeviceAssociationService  = Get-WmiObject -Query "select * from win32_service where name='DeviceAssociationService'"
        $Service_dot3svc  = Get-WmiObject -Query "select * from win32_service where name='dot3svc'"
        $Service_DsmSvc  = Get-WmiObject -Query "select * from win32_service where name='DsmSvc'"
        $Service_Eaphost  = Get-WmiObject -Query "select * from win32_service where name='Eaphost'"
        $Service_fdPHost  = Get-WmiObject -Query "select * from win32_service where name='fdPHost'"
        $Service_hkmsvc  = Get-WmiObject -Query "select * from win32_service where name='hkmsvc'"
        $Service_IEEtwCollectorService  = Get-WmiObject -Query "select * from win32_service where name='IEEtwCollectorService'"
        $Service_iphlpsvc  = Get-WmiObject -Query "select * from win32_service where name='iphlpsvc'"
        $Service_KPSSVC  = Get-WmiObject -Query "select * from win32_service where name='KPSSVC'"
        $Service_lltdsvc  = Get-WmiObject -Query "select * from win32_service where name='lltdsvc'"
        $Service_MMCSS  = Get-WmiObject -Query "select * from win32_service where name='MMCSS'"
        $Service_MSDTC  = Get-WmiObject -Query "select * from win32_service where name='MSDTC'"
        $Service_MSiSCSI  = Get-WmiObject -Query "select * from win32_service where name='MSiSCSI'"
        $Service_PolicyAgent  = Get-WmiObject -Query "select * from win32_service where name='PolicyAgent'"
        $Service_PrintNotify  = Get-WmiObject -Query "select * from win32_service where name='PrintNotify'"
        $Service_RasAuto  = Get-WmiObject -Query "select * from win32_service where name='RasAuto'"
        $Service_ScDeviceEnum  = Get-WmiObject -Query "select * from win32_service where name='ScDeviceEnum'"
        $Service_SCPolicySvc  = Get-WmiObject -Query "select * from win32_service where name='SCPolicySvc'"
        $Service_Spooler  = Get-WmiObject -Query "select * from win32_service where name='Spooler'"
        $Service_SstpSvc  = Get-WmiObject -Query "select * from win32_service where name='SstpSvc'"
        $Service_svsvc  = Get-WmiObject -Query "select * from win32_service where name='svsvc'"
        $Service_TapiSrv  = Get-WmiObject -Query "select * from win32_service where name='TapiSrv'"
        $Service_Themes  = Get-WmiObject -Query "select * from win32_service where name='Themes'"
        $Service_TieringEngineService  = Get-WmiObject -Query "select * from win32_service where name='TieringEngineService'"
        $Service_TrkWks  = Get-WmiObject -Query "select * from win32_service where name='TrkWks'"
        $Service_UALSVC  = Get-WmiObject -Query "select * from win32_service where name='UALSVC'"
        $Service_vmicguestinterface  = Get-WmiObject -Query "select * from win32_service where name='vmicguestinterface'"
        $Service_vmicheartbeat  = Get-WmiObject -Query "select * from win32_service where name='vmicheartbeat'"
        $Service_vmickvpexchange  = Get-WmiObject -Query "select * from win32_service where name='vmickvpexchange'"
        $Service_vmicrdv  = Get-WmiObject -Query "select * from win32_service where name='vmicrdv'"
        $Service_vmicshutdown  = Get-WmiObject -Query "select * from win32_service where name='vmicshutdown'"
        $Service_vmictimesync  = Get-WmiObject -Query "select * from win32_service where name='vmictimesync'"
        $Service_vmicvss  = Get-WmiObject -Query "select * from win32_service where name='vmicvss'"
        $Service_Wecsvc  = Get-WmiObject -Query "select * from win32_service where name='Wecsvc'"
        $Service_WEPHOSTSVC  = Get-WmiObject -Query "select * from win32_service where name='WEPHOSTSVC'"
        $Service_wercplsupport  = Get-WmiObject -Query "select * from win32_service where name='wercplsupport'"
        $Service_WerSvc  = Get-WmiObject -Query "select * from win32_service where name='WerSvc'"
      #>




     #各パラメーターを行として作成
      $obj = new-object psobject
      $obj | add-member noteproperty ComputerName $computername
      $obj | add-member noteproperty Domain ($Comp.Domain)
      $obj | add-member noteproperty OSVersion/Edition ($os.Caption)
      $obj | add-member noteproperty NIC_Description ($NetAdapterConf.Description)
      $obj | add-member noteproperty IPaddress ($NetAdapterConf.IPAddress)
      $obj | add-member noteproperty NetMask ($NetAdapterConf.IPSubnet)
      $obj | add-member noteproperty DNS ($NetAdapterConf.DNSServerSearchOrder,@{Name='DNSServerSearchOrder';Expression={[string]::join(";", ($_.DNSServerSearchOrder))}} )
      $obj | add-member noteproperty Gateway ($NetAdapterConf.DefaultIPGateway)
      $obj | add-member noteproperty PageFile_Path ($PageFile.EightDotThreeFileName)
      $obj | add-member noteproperty PageFile_Min ($PageFile.InitialSize)     
      $obj | add-member noteproperty PageFile_Max ($PageFile.MaximumSize)
      $obj | add-member noteproperty PageFile_managed ($Comp.AutomaticManagedPagefile) 
      $obj | add-member noteproperty MemDump ($Recover.DebugInfoType) 
      $obj | add-member noteproperty RDP01 ($RDP1.fDenyTSConnections)
      $obj | add-member noteproperty RDP02 ($RDP2.UserAuthentication)
      $obj | add-member noteproperty WindowsUpdate ($WinUpdate.Auoptions)
      $obj | add-member noteproperty IE_ESC-Admin ($IE_ESC_Admin.IsInstalled)
      $obj | add-member noteproperty IE_ESC-User ($IE_ESC_User.IsInstalled)
      $obj | add-member noteproperty Firewall ($Firewall.enabled)
      $obj | add-member noteproperty UAC ($UAC.ConsentPromptBehaviorAdmin)
      $obj | add-member noteproperty Power_Settings ($Power_Settings)
      $obj | add-member noteproperty SNP_RSS ($SNP1.ReceiveSideScaling)
      $obj | add-member noteproperty SNP_Chimney ($SNP1.Chimney)
      
      #w2k8R2用のみ必要（w2k12以降は規定で無効）
      #$obj | add-member noteproperty SNP_NetDMA ($SNP2.EnableTCPA)

      $obj | add-member noteproperty CD_Name ($CD.Caption)
      $obj | add-member noteproperty CD_Vol_Label ($CD.Drive)
      $obj | add-member noteproperty Disk_DeviceID ($DISK1.DeviceID)
      $obj | add-member noteproperty Disk_Size_GB ($DISK1.size/1024/1024/1024) 
      $obj | add-member noteproperty PhisycalDiskToPartition ($DISK2.Dependent) 
      $obj | add-member noteproperty LogicalDiskToPartition_Antecedent ($DISK3.Antecedent)
      $obj | add-member noteproperty LogicalDiskToPartition_Dependent ($DISK3.Dependent)


      #Windows 2008 R2 サービスの表示
      <# 
        $obj | add-member noteproperty Service_FontCache ($Service_FontCache.state)
        $obj | add-member noteproperty Service_WinRM ($Service_WinRM.state)
        $obj | add-member noteproperty Service_PlugPlay ($Service_PlugPlay.state)
        $obj | add-member noteproperty Service_ALG ($Service_ALG.state)
        $obj | add-member noteproperty Service_AudioEndpointBuilder ($Service_AudioEndpointBuilder.state)
        $obj | add-member noteproperty Service_AudioSrv ($Service_AudioSrv.state)
        $obj | add-member noteproperty Service_defragsvc ($Service_defragsvc.state)
        $obj | add-member noteproperty Service_dot3svc ($Service_dot3svc.state)
        $obj | add-member noteproperty Service_EapHost ($Service_EapHost.state)
        $obj | add-member noteproperty Service_fdPHost ($Service_fdPHost.state)
        $obj | add-member noteproperty Service_hkmsvc ($Service_hkmsvc.state)
        $obj | add-member noteproperty Service_iphlpsvc ($Service_iphlpsvc.state)
        $obj | add-member noteproperty Service_lltdsvc ($Service_lltdsvc.state)
        $obj | add-member noteproperty Service_MMCSS ($Service_MMCSS.state)
        $obj | add-member noteproperty Service_MSDTC ($Service_MSDTC.state)
        $obj | add-member noteproperty Service_MSiSCSI ($Service_MSiSCSI.state)
        $obj | add-member noteproperty Service_PolicyAgent ($Service_PolicyAgent.state)
        $obj | add-member noteproperty Service_RasAuto ($Service_RasAuto.state)
        $obj | add-member noteproperty Service_SCardSvr ($Service_SCardSvr.state)
        $obj | add-member noteproperty Service_SCPolicySvc ($Service_SCPolicySvc.state)
        $obj | add-member noteproperty Service_Spooler ($Service_Spooler.state)
        $obj | add-member noteproperty Service_SstpSvc ($Service_SstpSvc.state)
        $obj | add-member noteproperty Service_TapiSrv ($Service_TapiSrv.state)
        $obj | add-member noteproperty Service_TrkWks ($Service_TrkWks.state)
        $obj | add-member noteproperty Service_Wecsvc ($Service_Wecsvc.state)
        $obj | add-member noteproperty Service_wercplsupport ($Service_wercplsupport.state)
        $obj | add-member noteproperty Service_WerSvc ($Service_WerSvc.state)
    　#>


      #Windows 2012 サービスの表示
      <#
        $obj | add-member noteproperty Service_ALG ($Service_ALG.state)
        $obj | add-member noteproperty Service_AudioEndpointBuilder ($Service_AudioEndpointBuilder.state)
        $obj | add-member noteproperty Service_Audiosrv ($Service_Audiosrv.state)
        $obj | add-member noteproperty Service_defragsvc ($Service_defragsvc.state)
        $obj | add-member noteproperty Service_DeviceAssociationService ($Service_DeviceAssociationService.state)
        $obj | add-member noteproperty Service_dot3svc ($Service_dot3svc.state)
        $obj | add-member noteproperty Service_DsmSvc ($Service_DsmSvc.state)
        $obj | add-member noteproperty Service_Eaphost ($Service_Eaphost.state)
        $obj | add-member noteproperty Service_fdPHost ($Service_fdPHost.state)
        $obj | add-member noteproperty Service_hkmsvc ($Service_hkmsvc.state)
        $obj | add-member noteproperty Service_iphlpsvc ($Service_iphlpsvc.state)
        $obj | add-member noteproperty Service_KPSSVC ($Service_KPSSVC.state)
        $obj | add-member noteproperty Service_lltdsvc ($Service_lltdsvc.state)
        $obj | add-member noteproperty Service_MMCSS ($Service_MMCSS.state)
        $obj | add-member noteproperty Service_MSDTC ($Service_MSDTC.state)
        $obj | add-member noteproperty Service_MSiSCSI ($Service_MSiSCSI.state)
        $obj | add-member noteproperty Service_PolicyAgent ($Service_PolicyAgent.state)
        $obj | add-member noteproperty Service_PrintNotify ($Service_PrintNotify.state)
        $obj | add-member noteproperty Service_RasAuto ($Service_RasAuto.state)
        $obj | add-member noteproperty Service_SCPolicySvc ($Service_SCPolicySvc.state)
        $obj | add-member noteproperty Service_Spooler ($Service_Spooler.state)
        $obj | add-member noteproperty Service_SstpSvc ($Service_SstpSvc.state)
        $obj | add-member noteproperty Service_svsvc ($Service_svsvc.state)
        $obj | add-member noteproperty Service_TapiSrv ($Service_TapiSrv.state)
        $obj | add-member noteproperty Service_Themes ($Service_Themes.state)
        $obj | add-member noteproperty Service_TrkWks ($Service_TrkWks.state)
        $obj | add-member noteproperty Service_UALSVC ($Service_UALSVC.state)
        $obj | add-member noteproperty Service_vmicheartbeat ($Service_vmicheartbeat.state)
        $obj | add-member noteproperty Service_vmickvpexchange ($Service_vmickvpexchange.state)
        $obj | add-member noteproperty Service_vmicrdv ($Service_vmicrdv.state)
        $obj | add-member noteproperty Service_vmicshutdown ($Service_vmicshutdown.state)
        $obj | add-member noteproperty Service_vmictimesync ($Service_vmictimesync.state)
        $obj | add-member noteproperty Service_vmicvss ($Service_vmicvss.state)
        $obj | add-member noteproperty Service_Wecsvc ($Service_Wecsvc.state)
        $obj | add-member noteproperty Service_wercplsupport ($Service_wercplsupport.state)
        $obj | add-member noteproperty Service_WerSvc ($Service_WerSvc.state)
      #>


      #Windows 2012 R2 サービスの表示
      <#
        $obj | add-member noteproperty Service_BITS ($Service_BITS.state)
        $obj | add-member noteproperty Service_IKEEXT ($Service_IKEEXT.state)
        $obj | add-member noteproperty Service_ALG ($Service_ALG.state)
        $obj | add-member noteproperty Service_AudioEndpointBuilder ($Service_AudioEndpointBuilder.state)
        $obj | add-member noteproperty Service_Audiosrv ($Service_Audiosrv.state)
        $obj | add-member noteproperty Service_defragsvc ($Service_defragsvc.state)
        $obj | add-member noteproperty Service_DeviceAssociationService ($Service_DeviceAssociationService.state)
        $obj | add-member noteproperty Service_dot3svc ($Service_dot3svc.state)
        $obj | add-member noteproperty Service_DsmSvc ($Service_DsmSvc.state)
        $obj | add-member noteproperty Service_Eaphost ($Service_Eaphost.state)
        $obj | add-member noteproperty Service_fdPHost ($Service_fdPHost.state)
        $obj | add-member noteproperty Service_hkmsvc ($Service_hkmsvc.state)
        $obj | add-member noteproperty Service_IEEtwCollectorService ($Service_IEEtwCollectorService.state)
        $obj | add-member noteproperty Service_iphlpsvc ($Service_iphlpsvc.state)
        $obj | add-member noteproperty Service_KPSSVC ($Service_KPSSVC.state)
        $obj | add-member noteproperty Service_lltdsvc ($Service_lltdsvc.state)
        $obj | add-member noteproperty Service_MMCSS ($Service_MMCSS.state)
        $obj | add-member noteproperty Service_MSDTC ($Service_MSDTC.state)
        $obj | add-member noteproperty Service_MSiSCSI ($Service_MSiSCSI.state)
        $obj | add-member noteproperty Service_PolicyAgent ($Service_PolicyAgent.state)
        $obj | add-member noteproperty Service_PrintNotify ($Service_PrintNotify.state)
        $obj | add-member noteproperty Service_RasAuto ($Service_RasAuto.state)
        $obj | add-member noteproperty Service_ScDeviceEnum ($Service_ScDeviceEnum.state)
        $obj | add-member noteproperty Service_SCPolicySvc ($Service_SCPolicySvc.state)
        $obj | add-member noteproperty Service_Spooler ($Service_Spooler.state)
        $obj | add-member noteproperty Service_SstpSvc ($Service_SstpSvc.state)
        $obj | add-member noteproperty Service_svsvc ($Service_svsvc.state)
        $obj | add-member noteproperty Service_TapiSrv ($Service_TapiSrv.state)
        $obj | add-member noteproperty Service_Themes ($Service_Themes.state)
        $obj | add-member noteproperty Service_TieringEngineService ($Service_TieringEngineService.state)
        $obj | add-member noteproperty Service_TrkWks ($Service_TrkWks.state)
        $obj | add-member noteproperty Service_UALSVC ($Service_UALSVC.state)
        $obj | add-member noteproperty Service_vmicguestinterface ($Service_vmicguestinterface.state)
        $obj | add-member noteproperty Service_vmicheartbeat ($Service_vmicheartbeat.state)
        $obj | add-member noteproperty Service_vmickvpexchange ($Service_vmickvpexchange.state)
        $obj | add-member noteproperty Service_vmicrdv ($Service_vmicrdv.state)
        $obj | add-member noteproperty Service_vmicshutdown ($Service_vmicshutdown.state)
        $obj | add-member noteproperty Service_vmictimesync ($Service_vmictimesync.state)
        $obj | add-member noteproperty Service_vmicvss ($Service_vmicvss.state)
        $obj | add-member noteproperty Service_Wecsvc ($Service_Wecsvc.state)
        $obj | add-member noteproperty Service_WEPHOSTSVC ($Service_WEPHOSTSVC.state)
        $obj | add-member noteproperty Service_wercplsupport ($Service_wercplsupport.state)
        $obj | add-member noteproperty Service_WerSvc ($Service_WerSvc.state)
      #>





      Write-output $obj | export-csv c:\temp\test.csv
   }
}


 "127.0.0.1" |  Get-Inventory


 #Export-Csvを問題なく実行するためのおまじない
 #gwmi Win32_NetworkAdapterConfiguration  | select DNSServerSearchOrder,@{Name='Display-DNSServerSearchOrder';Expression={[string]::join(";", ($_.DNSServerSearchOrder))}} | Export-Csv c:\temp\test.csv





 <#
 "10.131.6.2" |  Get-Inventory
 "10.132.1.16" |  Get-Inventory
 "10.132.1.17" |  Get-Inventory
 "10.132.1.18" |  Get-Inventory
 "10.132.1.19" |  Get-Inventory
 "10.132.1.20" |  Get-Inventory
 "10.132.1.21" |  Get-Inventory
 #>






# "localhost" |  Get-Inventory


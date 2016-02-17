#edit: 2016_02/17

#テスト変更01#

#  http://www.evernote.com/l/AB5Ck8L4ASNFe6vt2HdOkZaq1wB0envHRHo/

#################
##手動追加作業###
#################
#PS1, ISEからの実行では、パラメーターシートに張り付けるためのget-windowsfeature が想定通り動かない
#このため下記コマンドで別途取得してください　
#$host.ui.rawui.buffersize = new-object system.management.automation.host.size(420, 3000)
#Start-Transcript -path "ファイル保存パス"
#get-windowsfeature | ft -auto | out-file c:\WindowsFeature_$env:computername.txt
#Stop-Transcript
#

##############################################
 #<CSV変換を System.Object[] 表示じさせないための処理>
 #本処理を通さないと/通せないとCSV Exportは実質機能しないと考えてよい
 Function Convert-OutputForCSV {
    <#
        .SYNOPSIS
            Provides a way to expand collections in an object property prior
            to being sent to Export-Csv.

        .DESCRIPTION
            Provides a way to expand collections in an object property prior
            to being sent to Export-Csv. This helps to avoid the object type
            from being shown such as system.object[] in a spreadsheet.

        .PARAMETER InputObject
            The object that will be sent to Export-Csv

        .PARAMETER OutPropertyType
            This determines whether the property that has the collection will be
            shown in the CSV as a comma delimmited string or as a stacked string.

            Possible values:
            Stack
            Comma

            Default value is: Stack

        .NOTES
            Name: Convert-OutputForCSV
            Author: Boe Prox
            Created: 24 Jan 2014
            Version History:
                1.1 - 02 Feb 2014
                    -Removed OutputOrder parameter as it is no longer needed; inputobject order is now respected 
                    in the output object
                1.0 - 24 Jan 2014
                    -Initial Creation

        .EXAMPLE
            $Output = 'PSComputername','IPAddress','DNSServerSearchOrder'

            Get-WMIObject -Class Win32_NetworkAdapterConfiguration -Filter "IPEnabled='True'" |
            Select-Object $Output | Convert-OutputForCSV | 
            Export-Csv -NoTypeInformation -Path NIC.csv    
            
            Description
            -----------
            Using a predefined set of properties to display ($Output), data is collected from the 
            Win32_NetworkAdapterConfiguration class and then passed to the Convert-OutputForCSV
            funtion which expands any property with a collection so it can be read properly prior
            to being sent to Export-Csv. Properties that had a collection will be viewed as a stack
            in the spreadsheet.        
            
    #>
    #Requires -Version 3.0
    [cmdletbinding()]
    Param (
        [parameter(ValueFromPipeline)]
        [psobject]$InputObject,
        [parameter()]
        [ValidateSet('Stack','Comma')]
        [string]$OutputPropertyType = 'Stack'
    )
    Begin {
        $PSBoundParameters.GetEnumerator() | ForEach {
            Write-Verbose "$($_)"
        }
        $FirstRun = $True
    }
    Process {
        If ($FirstRun) {
            $OutputOrder = $InputObject.psobject.properties.name
            Write-Verbose "Output Order:`n $($OutputOrder -join ', ' )"
            $FirstRun = $False
            #Get properties to process
            $Properties = Get-Member -InputObject $InputObject -MemberType *Property
            #Get properties that hold a collection
            $Properties_Collection = @(($Properties | Where-Object {
                $_.Definition -match "Collection|\[\]"
            }).Name)
            #Get properties that do not hold a collection
            $Properties_NoCollection = @(($Properties | Where-Object {
                $_.Definition -notmatch "Collection|\[\]"
            }).Name)
            Write-Verbose "Properties Found that have collections:`n $(($Properties_Collection) -join ', ')"
            Write-Verbose "Properties Found that have no collections:`n $(($Properties_NoCollection) -join ', ')"
        }
 
        $InputObject | ForEach {
            $Line = $_
            $stringBuilder = New-Object Text.StringBuilder
            $Null = $stringBuilder.AppendLine("[pscustomobject] @{")

            $OutputOrder | ForEach {
                If ($OutputPropertyType -eq 'Stack') {
                    $Null = $stringBuilder.AppendLine("`"$($_)`" = `"$(($line.$($_) | Out-String).Trim())`"")
                } ElseIf ($OutputPropertyType -eq "Comma") {
                    $Null = $stringBuilder.AppendLine("`"$($_)`" = `"$($line.$($_) -join ', ')`"")                   
                }
            }
            $Null = $stringBuilder.AppendLine("}")
 
            Invoke-Expression $stringBuilder.ToString()
        }
    }
    End {}
} #End of Convert-OutputForCSV
##############################################





      
#<各パラメーターの器を変数宣言>
#gwmi は Get-WMiObject の短縮
#gp は Get-ItemProperty の短縮

#OS関連の情報
$os = gwmi win32_operatingsystem

#BIOS関連情報
$BIOS = gwmi win32_BIOS

#ネットワークアダプター（物理寄り情報）
$NetAdapter = gwmi Win32_NetworkAdapter | sort InterfaceIndex |? {$_.NetEnabled}  

#ネットワークアダプター（論理情報寄り）
$global:NetAdapterConf = gwmi Win32_NetworkAdapterConfiguration | sort InterfaceIndex |? {$_.IPEnabled}  

#1つのセルに論理ネットワーク情報を取得するための関数
Function Show-IPconfiguration{
    Foreach($global:NICtemp in $global:NetAdapterConf)
    {
    "Interface Index: " + $global:NICtemp.InterfaceIndex
    "IP Address: " + $global:NICtemp.IPaddress
    "Default Gateway: " + $global:NICtemp.DefaultIPGateway
    "DNS Server: " + $global:NICtemp.DNSServerSearchOrder
    ""
    }
} #End of Show-IPconfiguration


#ネットワークアダプター（Alias-Description-Index連携）
#w2k12以降でしか使えない
$NetworkInfo = get-netipconfiguration | sort InterfaceIndex

#ネットワークチーム
$team = Get-NetLbfoTeam

#静的ルート
$StaticRoute = gwmi -Class Win32_IP4PersistedRouteTable

#IPv6設定
$IPv6Conf = get-NetAdapterBinding -InterfaceAlias * -ComponentID ms_tcpip6 | sort name

#ページファイル設定
$PageFile = gwmi Win32_PageFile

#HW寄り情報
$Comp = gwmi Win32_ComputerSystem

#メモリダンプ情報
$Recover = gwmi Win32_OSRecoveryConfiguration

#Windows Update設定格納レジストリ
$WinUpdate = gp "HKLM:\SOFTWARE\Microsoft\Windows\CurrentVersion\WindowsUpdate\Auto Update"

#RDP 設定格納レジストリ（RDP自体の有効化）
$RDP1 = gp "HKLM:\SYSTEM\CurrentControlSet\Control\Terminal Server"

#RDP 設定格納レジストリ（以前の設定有効化）
$RDP2 = gp "HKLM:\SYSTEM\CurrentControlSet\Control\Terminal Server\WinStations\RDP-Tcp"

#IE ESC 設定格納レジストリ（管理者）
$IE_ESC_Admin = gp "HKLM:\SOFTWARE\Microsoft\Active Setup\Installed Components\{A509B1A7-37EF-4b3f-8CFC-4F3A74704073}"

#IE ESC 設定格納レジストリ（ユーザー）
$IE_ESC_User =  gp "HKLM:\SOFTWARE\Microsoft\Active Setup\Installed Components\{A509B1A8-37EF-4b3f-8CFC-4F3A74704073}"

#ファイアウォール設定
$Firewall = Get-NetFirewallProfile | sort Profile

#UAC設定
$UAC = gp "HKLM:\SOFTWARE\Microsoft\Windows\CurrentVersion\Policies\System"

#電源オプション設定
$Power_Settings = powercfg.exe /GETACTIVESCHEME

#SNP設定
$SNP1 = Get-NetOffloadGlobalSetting

#w2k8R2用のみ必要（w2k12以降は規定で無効）
#$SNP2 = gp "HKLM:SYSTEM\CurrentControlSet\Services\Tcpip\Parameters"

#CDドライブ設定
$CD = gwmi Win32_CDROMDrive

#ディスク設定（物理寄り）
$Disk1 = gwmi Win32_DiskDrive | Sort DeviceID

#物理ディスク設定とパーティション情報の紐づけ
$Disk2 = gwmi Win32_DiskDriveToDiskPartition

#論理ディスク設定とパーティション情報の紐づけ
$Disk3 = gwmi Win32_LogicalDiskToPartition

#サービス情報
$Service  = gwmi -Query "select * from win32_service" | sort displayname

#HotFix情報
$WindowsHotfix = gwmi win32_quickfixengineering | Sort HotfixID

###↓プログラムと機能取得のための変数宣言ここから↓###
$Regpath_Program_x86 = "\SOFTWARE\Microsoft\Windows\CurrentVersion\Uninstall"
$Regpath_Program_x64 = "HKLM:\SOFTWARE\Wow6432Node\Microsoft\Windows\CurrentVersion\Uninstall"
$ProgramPathArray = ("HKLM:" + $Regpath_Program_x86), ("HKCU:" + $Regpath_Program_x86), $Regpath_Program_x64
$InstalledProgram  = (Get-ChildItem -Path $ProgramPathArray -ErrorAction SilentlyContinue  | %{Get-ItemProperty $_.PsPath} |

#※下記1行を含めると更新プログラム系の表示が除外される
 ?{$_.systemcomponent -ne 1 -and $_.parentkeyname -eq $nu1111111111111111ll}  ) 
###↑プログラムと機能取得のための変数宣言ここまで↑###



 
#物理 - 論理Disk情報関連付けて情報収集
##################################################### 
Function DriveInfo {

$global:strComputer = "." 
$global:colDiskDrives = get-wmiobject -query "Select * From Win32_DiskDrive" | sort DeviceID
Foreach ($global:drive in $global:colDiskDrives) 
    { 
        "Device ID:  "+ $global:drive.DeviceID
        $global:a = $global:drive.DeviceID.Replace("\", "\\") 
        $global:colPartitions = get-wmiobject -query "Associators of {Win32_DiskDrive.DeviceID=""$global:a""} WHERE AssocClass = Win32_DiskDriveToDiskPartition"
       Foreach ($global:Partition in $global:colPartitions) 
            { 
                $global:b = $global:Partition.DeviceID 
                $global:colLogicalDisk = get-wmiobject -query "Associators of {Win32_DiskPartition.DeviceID=""$global:b""} WHERE AssocClass = Win32_LogicalDiskToPartition"
                If ($global:colLogicalDisk.Caption -ne $null) 
                    { 
                        "Drive Letter: "+ $global:colLogicalDisk.Caption 
                    } 
                Else 
                    { 
                        "Drive Letter: 無し" 
                    } 
                "Partition ID: "+ $global:Partition.DeviceID 
                $global:c = $global:colLogicalDisk.Size/1GB 
                $global:c = [math]::round($global:c, 0) 
        "Size: "+ $global:c+"GB" 
        "Volume Name: "+ $global:colLogicalDisk.VolumeName 
        " " 
            } 
    } 
    
    
    }
####################################################






#Local Groupの情報を取得 $group内の記述を変えて任意のグループを指定できる
##################################################### 
Function Show-LocalGroup {

$strComputer = “.”
$computer = [ADSI](“WinNT://” + $strComputer + “,computer”)
$Group = $computer.psbase.children.find(“Administrators”)
$members= $Group.psbase.invoke(“Members”) | %{$_.GetType().InvokeMember(“Name”, ‘GetProperty’, $null, $_, $null)}

ForEach($user in $members)

{
Write-output $user
$a = $strComputer + “!” + $user.ToString()

}


}
#####################################################







####Windows のプロダクトキーの取得####

Function Get-WindowsKey
{
$Hklm = 2147483650
$Target = $env:COMPUTERNAME
$regPath = "Software\Microsoft\Windows NT\CurrentVersion"
$DigitalID = "DigitalProductId"
$wmi = [WMIClass]"\\$Target\root\default:stdRegProv"
#Get registry value 
$Object = $wmi.GetBinaryValue($hklm,$regPath,$DigitalID)
[Array]$DigitalIDvalue = $Object.uValue 
#If get successed


#Get producnt name and product ID
#$ProductName = (Get-itemproperty -Path "HKLM:Software\Microsoft\Windows NT\CurrentVersion" -Name "ProductName").ProductName 
#$ProductID =  (Get-itemproperty -Path "HKLM:Software\Microsoft\Windows NT\CurrentVersion" -Name "ProductId").ProductId
#Convert binary value to serial number 
$Result = ConvertTokey $DigitalIDvalue
$OSInfo = (Get-WmiObject "Win32_OperatingSystem"  | select Caption).Caption



[string]$value ="Installed Key: $Result"
$value 


}

Function ConvertToKey($Key)
{
$Keyoffset = 52 
$isWin8 = [int]($Key[66]/6) -band 1
$HF7 = 0xF7
$Key[66] = ($Key[66] -band $HF7) -bOr (($isWin8 -band 2) * 4)
$i = 24
[String]$Chars = "BCDFGHJKMPQRTVWXY2346789"
do
{
$Cur = 0 
$X = 14
Do
{
$Cur = $Cur * 256    
$Cur = $Key[$X + $Keyoffset] + $Cur
$Key[$X + $Keyoffset] = [math]::Floor([double]($Cur/24))
$Cur = $Cur % 24
$X = $X - 1 
}while($X -ge 0)
$i = $i- 1
$KeyOutput = $Chars.SubString($Cur,1) + $KeyOutput
$last = $Cur
}while($i -ge 0)
$Keypart1 = $KeyOutput.SubString(1,$last)
$Keypart2 = $KeyOutput.Substring(1,$KeyOutput.length-1)
if($last -eq 0 )
{
$KeyOutput = "N" + $Keypart2
}
else
{
$KeyOutput = $Keypart2.Insert($Keypart2.IndexOf($Keypart1)+$Keypart1.length,"N")
}
$a = $KeyOutput.Substring(0,5)
$b = $KeyOutput.substring(5,5)
$c = $KeyOutput.substring(10,5)
$d = $KeyOutput.substring(15,5)
$e = $KeyOutput.substring(20,5)
$keyproduct = $a + "-" + $b + "-"+ $c + "-"+ $d + "-"+ $e
$keyproduct 
  
}

########################################







#組織と所有者名
$UserAndOrg = gp "HKLM:\Software\Microsoft\Windows NT\CurrentVersion\"  -name RegisteredOwner, RegisteredOrganization | select RegisteredOwner, RegisteredOrganization
 


#######ライセンス状態確認#######
function Get-ActivationStatus {
[CmdletBinding()]
    param(
        [Parameter(ValueFromPipeline = $true, ValueFromPipelineByPropertyName = $true)]
        [string]$DNSHostName = $Env:COMPUTERNAME
    )
    process {
        try {
            $wpa = Get-WmiObject SoftwareLicensingProduct -ComputerName $DNSHostName `
            -Filter "ApplicationID = '55c92734-d682-4d71-983e-d6ec3f16059f'" `
            -Property LicenseStatus -ErrorAction Stop
        } catch {
            $status = New-Object ComponentModel.Win32Exception ($_.Exception.ErrorCode)
            $wpa = $null    
        }
        $out = New-Object psobject -Property @{
       #     ComputerName = $DNSHostName;
            Status = [string]::Empty;
        }
        if ($wpa) {
            :outer foreach($item in $wpa) {
                switch ($item.LicenseStatus) {
                    0 {$out.Status = "Unlicensed"}
                    1 {$out.Status = "Licensed"; break outer}
                    2 {$out.Status = "Out-Of-Box Grace Period"; break outer}
                    3 {$out.Status = "Out-Of-Tolerance Grace Period"; break outer}
                    4 {$out.Status = "Non-Genuine Grace Period"; break outer}
                    5 {$out.Status = "Notification"; break outer}
                    6 {$out.Status = "Extended Grace"; break outer}
                    default {$out.Status = "Unknown value"}
                }
            }
        } else {$out.Status = $status.Message}
        $out
    }
}
###########################################




#ローカル Administratorのパスワード無期限かどうかを確認
$AdminExpiration = Get-CimInstance -ClassName Win32_UserAccount | select Caption, passwordexpires | where caption -like *administrator*  | fl *


#ローカルグループポリシーの確認
$Show_LocalPolicy =  gpresult /v /scope computer


#SNMP Communityの確認（HPサーバーでSMHを表示させるための設定）
$SNMPCommunity = Get-ItemProperty -path "HKLM:\SYSTEM\CurrentControlSet\Services\SNMP\Parameters\ValidCommunities\"  -ErrorAction SilentlyContinue | fl community 
$SNMPSecurity = Get-ItemProperty "HKLM:\SYSTEM\CurrentControlSet\Services\SNMP\Parameters\PermittedManagers\"  -ErrorAction SilentlyContinue | fl "1"





 #<各パラメーターを生成>
 #前述で宣言した器の後ろに " .パラメーター " を指定することで
 #任意の値を各クラスから呼び出している 

      $obj = new-object psobject
      $obj | add-member noteproperty "ホスト名" ($OS.__SERVER)

      $obj | add-member noteproperty "＜ドメイン情報＞ `n ドメイン名 `n[win32_ComputerSystem]" ($Comp.Domain)
      $obj | add-member noteproperty "＜OS種別＞ `n OSのバージョンとエディション `n[win32_OperatingSystem]" ($os.Caption)

      $obj | add-member noteproperty "＜BIOS情報＞ `n シリアル番号 `n [win32_BIOS] `n（仮想マシンでも何らかの値は表示される）" ($BIOS.SerialNumber)

      
      #Windows 2008向け NIC管理情報
      <#
      $obj | add-member noteproperty （NIC管理）NIC_InterfaceIndex ($NetAdapterConf.InterfaceIndex)
      $obj | add-member noteproperty （NIC管理）NIC_Label ($NetAdapter.NetConnectionID)
      $obj | add-member noteproperty （NIC管理）NIC_Description ($NetAdapterConf.Description)
      #>

      $obj | add-member noteproperty "＜NIC管理＞`n InterfaceIndex（番号） `n[get-netipconfiguration] `n（インターフェース番号、名前、デバイス名の紐づきを確認するために本情報を収集 →）" ($NetworkInfo.InterfaceIndex)
      $obj | add-member noteproperty "＜NIC管理＞`n 名前 `n [get-netipconfiguration] `n（ncpa.cpl実行時に表示される名前）" ($NetworkInfo.InterfaceAlias)
      $obj | add-member noteproperty "＜NIC管理＞`n デバイス名 `n[get-netipconfiguration]`n（←）" ($NetworkInfo.InterfaceDescription)

      $obj | add-member noteproperty "＜IPアドレス＞`n IPconfig情報 `n[get-netipconfiguration] `n（各インターフェースIDのIPアドレス情報を一覧化）" (Show-IPconfiguration)

      # NICラベル名は取得可能だが、接続されていないNICがあると目視で対応させにくいため表示させるべきか要検討 2015/8
      #Accosiatorコマンドで紐づけはおそらく可能 2015/8


      $obj | add-member noteproperty "＜NICアダプタ設定＞`n InterfaceIndex（番号） `n [Win32_NetworkAdapterConfiguration] `n （NICアダプタ寄りの情報を取得）" ($NetAdapterConf.InterfaceIndex)
      $obj | add-member noteproperty "＜NICアダプタ設定＞`n IPアドレス `n [Win32_NetworkAdapterConfiguration]" ($NetAdapterConf.IPAddress)
      $obj | add-member noteproperty "＜NICアダプタ設定＞`n サブネットマスク `n [Win32_NetworkAdapterConfiguration]" ($NetAdapterConf.IPSubnet)
      $obj | add-member noteproperty "＜NICアダプタ設定＞`n DNSサーバー `n [Win32_NetworkAdapterConfiguration]" ($NetAdapterConf.DNSServerSearchOrder)
      $obj | add-member noteproperty "＜NICアダプタ設定＞`n ゲートウェイ `n [Win32_NetworkAdapterConfiguration]" ($NetAdapterConf.DefaultIPGateway)
      $obj | add-member noteproperty "＜NICアダプタ設定＞`n Wins（1番目） `n [Win32_NetworkAdapterConfiguration] `n （設定値なければブランク）" ($NetAdapterConf.WINSPrimaryServer)
      $obj | add-member noteproperty "＜NICアダプタ設定＞`n Wins（2番目） `n [Win32_NetworkAdapterConfiguration] `n （設定値なければブランク）" ($NetAdapterConf.WINSSecondaryServer)
      $obj | add-member noteproperty "＜NICアダプタ設定＞`n この接続のアドレスをDNSに登録する `n [Win32_NetworkAdapterConfiguration] `n（TCP/IP詳細設定 - DNSタブ　に存在する項目、NIC2つ以上ある場合は要確認）" ($NetAdapterConf.FullDNSRegistrationEnabled)
      $obj | add-member noteproperty "＜NICアダプタ設定＞`n MACアドレス `n [Win32_NetworkAdapterConfiguration]" ($NetAdapterConf.MACAddress)


      $obj | add-member noteproperty "＜IPv6設定＞`n 名前 `n [get-NetAdapterBinding] `n（ncpa.cpl実行時に表示される名前）" ($IPv6Conf.name)
      $obj | add-member noteproperty "＜IPv6設定＞`n 有効/無効 `n [get-NetAdapterBinding]" ($IPv6Conf.enabled)


      $obj | add-member noteproperty "＜チーム設定＞ `n 名前 `n [Get-NetLbfoTeam] `n （無効の場合はブランク表示）" ($team.name)
      $obj | add-member noteproperty "＜チーム設定＞ `n メンバ `n [Get-NetLbfoTeam]" ($team.members)
      $obj | add-member noteproperty "＜チーム設定＞ `n モード  `n [Get-NetLbfoTeam]" ($team.TeamingMode)
      $obj | add-member noteproperty "＜チーム設定＞ `n ロードバランスアルゴリズム  `n [Get-NetLbfoTeam]" ($team.LoadBalancingAlgorithm)
      $obj | add-member noteproperty "＜チーム設定＞ `n ステータス `n [Get-NetLbfoTeam]" ($team.Status)


      $obj | add-member noteproperty "＜静的ルート＞ `n 宛先ネットワーク `n（設定していない場合は 0.0.0.0 表示）" ($StaticRoute.Destination)
      $obj | add-member noteproperty "＜静的ルート＞ `n サブネットマスク" ($StaticRoute.Mask)
      $obj | add-member noteproperty "＜静的ルート＞ `n 宛先ルーター/ネクストホップ" ($StaticRoute.NextHop)
      $obj | add-member noteproperty "＜静的ルート＞ `n メトリック" ($StaticRoute.Metric1)


      $obj | add-member noteproperty "＜ページファイル（仮想メモリ）＞ `n 保存パス `n [Win32_PageFile] `n（OSによる自動管理の場合はブランク）" ($PageFile.EightDotThreeFileName)
      $obj | add-member noteproperty "＜ページファイル（仮想メモリ）＞ `n 最小サイズ `n [Win32_PageFile] `n" ($PageFile.InitialSize)     
      $obj | add-member noteproperty "＜ページファイル（仮想メモリ）＞ `n 最大サイズ `n [Win32_PageFile] `n" ($PageFile.MaximumSize)
      $obj | add-member noteproperty "＜ページファイル（仮想メモリ）＞ `n OSによる管理 `n [Win32_PageFile] `n" ($Comp.AutomaticManagedPagefile) 
      $obj | add-member noteproperty "＜ページファイル（仮想メモリ）＞ `n MemDump  `n [Win32_PageFile] `n（7:自動、0:なし、1:最少、2:カーネル、3:完全）" ($Recover.DebugInfoType) 
      $obj | add-member noteproperty "＜リモートデスクトップ＞ `n 接続許可  `n [HKLM:\SYSTEM\CurrentControlSet\Control\Terminal Server] `n（0:拒否しない/有効、1:拒否/無効）" ($RDP1.fDenyTSConnections)
      $obj | add-member noteproperty "＜リモートデスクトップ＞ `n 以前の接続を有効化する  `n [HKLM:\SYSTEM\CurrentControlSet\Control\Terminal Server\WinStations\RDP-Tcp] `n（1:有効、0:無効）" ($RDP2.UserAuthentication)
      $obj | add-member noteproperty "＜WindowsUpdate＞ `n WindowsUpdate設定  `n [HKLM:\SOFTWARE\Microsoft\Windows\CurrentVersion\WindowsUpdate\Auto Update] `n （1:自動更新しない、2:確認のみ、3:ダウンロードのみ、4:自動適用、ブランク:設定なし）" ($WinUpdate.Auoptions)
      $obj | add-member noteproperty "＜IE_セキュリティ強化の構成＞ `n Admin  `n [HKLM:\SOFTWARE\Microsoft\Active Setup\Installed Components\{A509B1A7-37EF-4b3f-8CFC-4F3A74704073}] `n（0:無効、1:有効）" ($IE_ESC_Admin.IsInstalled)
      $obj | add-member noteproperty "＜IE_セキュリティ強化の構成＞ `n User  `n [HKLM:\SOFTWARE\Microsoft\Active Setup\Installed Components\{A509B1A8-37EF-4b3f-8CFC-4F3A74704073}] `n（0:無効、1:有効）" ($IE_ESC_User.IsInstalled)


      $obj | add-member noteproperty "＜SNMP設定＞ `n コミュニティ名 `n [HKLM:\SYSTEM\CurrentControlSet\Services\SNMP\Parameters\ValidCommunities\] `n （HPサーバーSMH対応設定:4）" ($SNMPCommunity)
      $obj | add-member noteproperty "＜SNMP設定＞ `n セキュリティ `n [HKLM:\SYSTEM\CurrentControlSet\Services\SNMP\Parameters\PermittedManagers\] `n （HPサーバーSMH対応:localhost、SNMP Service未インストール:ブランク）" ($SNMPSecurity)
      

      $obj | add-member noteproperty "＜F/W設定＞ `n Firewall_Profile `n [Get-NetFirewallProfile]" ($Firewall.Profile)
      $obj | add-member noteproperty "＜F/W設定＞ `n Firewall_Setting `n [Get-NetFirewallProfile]" ($Firewall.enabled)


      #2U系でのみ有効にすると便利
      #$obj | add-member noteproperty （F/W）ファイルとプリンターの共有（表示名） (Get-NetFirewallRule -DisplayName "ファイルとプリンターの共有 (エコー要求 - ICMPv4 受信)" | Select-Object DisplayName | fw -column 1)
      #$obj | add-member noteproperty （F/W）ファイルとプリンターの共有（有効） (Get-NetFirewallRule  -DisplayName "ファイルとプリンターの共有 (エコー要求 - ICMPv4 受信)" | Select-Object Enabled | fw -column 1)
      #$obj | add-member noteproperty （F/W）ファイルとプリンターの共有（プロファイル） (Get-NetFirewallRule  -DisplayName "ファイルとプリンターの共有 (エコー要求 - ICMPv4 受信)" | Select-Object Profile | fw -column 1)


      $obj | add-member noteproperty "＜UAC設定＞ `n UAC設定 `n [HKLM:\SOFTWARE\Microsoft\Windows\CurrentVersion\Policies\System] `n （0:無効、5:デスクトップ暗転しない、?:既定、1:常に確認）" ($UAC.ConsentPromptBehaviorAdmin)
      $obj | add-member noteproperty "＜電源オプション＞ `n 電源オプション `n [powercfg.exe /GETACTIVESCHEME] `n （規定: 381b4222-f694-41f0-9685-ff5bb260df2e  (バランス)）" ($Power_Settings)

      $obj | add-member noteproperty "＜SNP設定＞ `n ReceiveSideScaling `n [Get-NetOffloadGlobalSetting] `n " ($SNP1.ReceiveSideScaling)
      $obj | add-member noteproperty "＜SNP設定＞ `n Chimney `n [Get-NetOffloadGlobalSetting] `n " ($SNP1.Chimney)
      #w2k8R2用のみ必要（w2k12以降は規定で無効）
      #$obj | add-member noteproperty SNP_NetDMA ($SNP2.EnableTCPA)

      $obj | add-member noteproperty "＜組織と所有者＞ `n 組織と所有者 `n [HKLM:\Software\Microsoft\Windows NT\CurrentVersion\]" ($UserAndOrg | fl)

      $obj | add-member noteproperty "＜CDドライブ情報＞ `n ドライブ名 `n [Win32_CDROMDrive]" ($CD.Caption)
      $obj | add-member noteproperty "＜CDドライブ情報＞ `n ボリュームラベル `n [Win32_CDROMDrive]" ($CD.Drive)

      $obj | add-member noteproperty "＜論理ドライブ情報＞ `n 論理ドライブ情報 `n [N/A] `n （各ドライブの情報をリスト化して表示する）" (DriveInfo)



      #Administratorsグループのメンバ
      #$obj | add-member noteproperty "＜ローカルユーザ設定＞ `n Local_Administratorsの所属メンバ `n [ADSI]" (Show-LocalGroup)
      $obj | add-member noteproperty "＜ローカルユーザ設定＞ `n ローカルAdminのパスワード期限 `n [Win32_UserAccount]" ($AdminExpiration)

      #必要に応じてローカルポリシーの表示結果を調整する
      #$obj | add-member noteproperty "ローカルグループポリシー（PC） ($Show_LocalPolicy)

      $obj | add-member noteproperty "＜TimeZone＞ `n TimeZone `n [tzutil.exe]" (tzutil.exe /g)
      
      #下記はレジストリの情報をひっかけてカスタム定義した文字列に置き換えて表示しているが、煩雑なため利用を保留
      #$obj | add-member noteproperty ＜ライセンス認証＞OSライセンス状態 (Get-ActivationStatus | fl)

      #下記の方がslmgr.vbsをもとに出力しているためシンプルでよい
      $obj | add-member noteproperty "＜ライセンス認証＞ `n ライセンス認証状態 `n [C:\Windows\System32\slmgr.vbs] `n （slmgrで表示される情報から必要な文字列のみ抽出している、「ライセンスされています」と表示されていない場合は要状態確認）" (cscript "C:\Windows\System32\slmgr.vbs" -dli | select-string "ライセンスの状態:")

      $obj | add-member noteproperty "＜ライセンス認証＞ `n ライセンスキー `n [HKLM:\Microsoft\Windows NT\CurrentVersion] `n （2015/12/2 暫定運用のため、念のため正しい値かダブルチェックしてください）" (Get-WindowsKey)

      #Disk関連（物理-論理関連付け無し）2015/8/1 使い勝手が悪いので非表示
      <#
      #$obj | add-member noteproperty Disk_DeviceID ($DISK1.DeviceID)
      #$obj | add-member noteproperty Disk_Size_GB ($DISK1.size) 
      #$obj | add-member noteproperty PhisycalDiskToPartition ($DISK2.Dependent) 
      #$obj | add-member noteproperty LogicalDiskToPartition_Antecedent ($DISK3.Antecedent)
      #$obj | add-member noteproperty LogicalDiskToPartition_Dependent ($DISK3.Dependent)
      #>






      $obj | add-member noteproperty "＜サービス＞ `n 表示名 `n [win32_service]" ($Service.displayname) 
      $obj | add-member noteproperty "＜サービス＞ `n 名前 `n  [win32_service]" ($Service.name) 
      $obj | add-member noteproperty "＜サービス＞ `n 状態 `n  [win32_service]" ($Service.state) 
      $obj | add-member noteproperty "＜サービス＞ `n スタート状態 `n  [win32_service]" ($Service.startmode) 
                                      
      $obj | add-member noteproperty "＜更新プログラム＞ `n 種別 `n [win32_quickfixengineering]" ($WindowsHotfix.Description)
      $obj | add-member noteproperty "＜更新プログラム＞ `n ID `n [win32_quickfixengineering]" ($WindowsHotfix.HotfixID)
      $obj | add-member noteproperty "＜更新プログラム＞ `n インストール日 `n [win32_quickfixengineering]" ($WindowsHotfix.InstalledOn)
     
      #想定したフォーマットで取得できないため、下記は実施しない（2/17)
      #$obj | add-member noteproperty "＜役割＞ `n 一覧 `n [Get-WindowsFeature] `n （特定の情報のみではなく全情報を収集）" (Get-WindowsFeature) 
      $obj | add-member noteproperty "＜役割＞ `n 役割と機能_表示名のみ `n [Get-WindowsFeature]" ((Get-WindowsFeature).Displayname)
      $obj | add-member noteproperty "＜役割＞ `n 役割と機能_名前のみ `n [Get-WindowsFeature]" ((Get-WindowsFeature).name)
      $obj | add-member noteproperty "＜役割＞ `n 役割と機能_インストール状態のみ `n [Get-WindowsFeature]" ((Get-WindowsFeature).InstallState)


      $obj | add-member noteproperty "＜プログラム＞ `n Program_Name `n [\SOFTWARE\Microsoft\Windows\CurrentVersion\Uninstall]" ($InstalledProgram.displayname)

     
#本設定ではcドライブ直下に保存される
      Write-output $obj  | Convert-OutputForCSV | export-csv c:\inventory_$env:computername.csv -Encoding UTF8 -NoTypeInformation








<#　オプションで下記実行すること

#イベントログ出力（Applog）#本設定ではcドライブ直下に保存される
Get-EventLog -LogName "system"  | select EntryType,EventID,Source,TimeGenerated,Message | export-csv ("c:\" + $env:computername + "_syslog.csv") -encoding UTF8 -notypeinformation

#イベントログ出力（Syslog）#本設定ではcドライブ直下に保存される
Get-EventLog -LogName "application"  | select EntryType,EventID,Source,TimeGenerated,Message | export-csv ("c:\" + $env:computername + "_application.csv") -encoding UTF8 -notypeinformation

####ADの場合下記も取得####


#Directory Service
Get-EventLog -LogName "Directory Service"  | select EntryType,EventID,Source,TimeGenerated,Message | export-csv ("c:\" + $env:computername + "_Directory_Service.csv") -encoding UTF8 -notypeinformation

#DNS Server
Get-EventLog -LogName "DNS Server"  | select EntryType,EventID,Source,TimeGenerated,Message | export-csv ("c:\" + $env:computername + "_DNS_Server.csv") -encoding UTF8 -notypeinformation

#DFS Replication
Get-EventLog -LogName "DFS Replication"  | select EntryType,EventID,Source,TimeGenerated,Message | export-csv ("c:\" + $env:computername + "_DFS_Replication.csv") -encoding UTF8 -notypeinformation




#>

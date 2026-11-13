# vbs
<!-- 解説: この見出しは「vbs」を説明する節のタイトルです。 -->

## System information scripts
<!-- 解説: この見出しは「System information scripts」を説明する節のタイトルです。 -->

The following scripts query local OS, hardware, storage, and peripheral
information through WMI, Windows Script Host APIs, PowerShell, or built-in
Windows commands. Run an individual script from Command Prompt:
<!-- 解説: この本文行は、スクリプトがローカルOS、ハードウェア、記憶装置、周辺機器を照会することを説明します。 -->
<!-- 解説: この本文行は、WMI、WSH、PowerShell、Windows標準コマンドを利用することを説明します。 -->
<!-- 解説: この本文行は、個別スクリプトをコマンドプロンプトから起動する手順を案内します。 -->

```text
cscript //nologo GetOSInfo.vbs
```
<!-- 解説: この行はREADME内の実行例コードブロックの開始を示します。 -->
<!-- 解説: このコマンド例はGetOSInfo.vbsをWindows Script Hostのコンソールモードで実行します。 -->
<!-- 解説: この行は実行例コードブロックの終端です。 -->

Other scripts in this directory cover architecture, computer name, join
status, service pack, uptime, CPU and memory, page file, motherboard, BIOS,
disks, volumes, display, audio, USB devices, printers, battery, input devices,
and time-zone information. Each script prints results locally; none uploads
data.
<!-- 解説: この本文行はREADMEの使い方、対象範囲、または注意事項を説明しています: Other scripts in this directory cover architecture, computer name, join -->
<!-- 解説: この本文行はREADMEの使い方、対象範囲、または注意事項を説明しています: status, service pack, uptime, CPU and memory, page file, motherboard, BIOS, -->
<!-- 解説: この本文行はREADMEの使い方、対象範囲、または注意事項を説明しています: disks, volumes, display, audio, USB devices, printers, battery, input devices, -->
<!-- 解説: この本文行はREADMEの使い方、対象範囲、または注意事項を説明しています: and time-zone information. Each script prints results locally; none uploads -->
<!-- 解説: この本文行はREADMEの使い方、対象範囲、または注意事項を説明しています: data. -->

The original basic set does not collect clipboard contents, environment
variables, activation details, security-control status, network
identifiers/settings, installed application or update inventories, or system
logs. The additional read-only inventory scripts below cover selected network,
license, and security-status fields requested for local diagnostics; they do
not upload data. Their output can contain device identifiers or security
posture, so handle it accordingly. The computer-name script is read-only and
does not rename the machine.
<!-- 解説: この本文行は、初期の基本スクリプトがクリップボードや環境変数などを収集しない範囲を示します。 -->
<!-- 解説: この本文行は、追加スクリプトが要求されたネットワーク、ライセンス、セキュリティ状態をローカルで確認することを説明します。 -->
<!-- 解説: この本文行は、出力に含まれる端末識別情報やセキュリティ情報を適切に扱う注意を示します。 -->
<!-- 解説: この本文行は、コンピューター名のスクリプトが名前変更をせず読み取り専用であることを説明します。 -->

## Additional system and hardware inventory (No. 36-135)

The following 100 scripts are separate, read-only local diagnostics.
Availability and values depend on Windows edition, hardware, driver, WMI
provider, and permissions. Some commands require PowerShell or built-in tools;
output stays on the local console. Product-key suffixes and device identifiers
may be sensitive; handle output accordingly. The BitLocker script reports
protector IDs only, never recovery passwords.
<!-- 解説: この段落は追加された100本の読み取り専用スクリプトの制約と機微情報の取り扱いを説明します。 -->

| No. | Script | Information |
|---:|---|---|
| 36 | `036_MotherboardManufacturer.vbs` | マザーボード製造元 <!-- 解説: この行はNo.36の取得内容と対応する個別スクリプトを示します。 --> |
| 37 | `037_MotherboardProduct.vbs` | マザーボード製品型番 <!-- 解説: この行はNo.37の取得内容と対応する個別スクリプトを示します。 --> |
| 38 | `038_MotherboardSerial.vbs` | マザーボードシリアル番号 <!-- 解説: この行はNo.38の取得内容と対応する個別スクリプトを示します。 --> |
| 39 | `039_BiosVendorVersion.vbs` | BIOSベンダーとバージョン <!-- 解説: この行はNo.39の取得内容と対応する個別スクリプトを示します。 --> |
| 40 | `040_BiosReleaseDate.vbs` | BIOSリリース日時 <!-- 解説: この行はNo.40の取得内容と対応する個別スクリプトを示します。 --> |
| 41 | `041_FirmwareBootType.vbs` | UEFI/レガシーブート方式 <!-- 解説: この行はNo.41の取得内容と対応する個別スクリプトを示します。 --> |
| 42 | `042_ChassisType.vbs` | 筐体タイプ <!-- 解説: この行はNo.42の取得内容と対応する個別スクリプトを示します。 --> |
| 43 | `043_ChassisIdentifiers.vbs` | 筐体シリアル番号と資産タグ <!-- 解説: この行はNo.43の取得内容と対応する個別スクリプトを示します。 --> |
| 44 | `044_SystemFamilySku.vbs` | システムファミリとSKU <!-- 解説: この行はNo.44の取得内容と対応する個別スクリプトを示します。 --> |
| 45 | `045_PciSlots.vbs` | PCI/PCIeスロット構成 <!-- 解説: この行はNo.45の取得内容と対応する個別スクリプトを示します。 --> |
| 46 | `046_CpuProcessorName.vbs` | CPU製品名 <!-- 解説: この行はNo.46の取得内容と対応する個別スクリプトを示します。 --> |
| 47 | `047_CpuManufacturer.vbs` | CPUメーカー <!-- 解説: この行はNo.47の取得内容と対応する個別スクリプトを示します。 --> |
| 48 | `048_ProcessorCounts.vbs` | 物理/論理プロセッサ数 <!-- 解説: この行はNo.48の取得内容と対応する個別スクリプトを示します。 --> |
| 49 | `049_CpuClockFrequencies.vbs` | CPU最大/現在クロック周波数 <!-- 解説: この行はNo.49の取得内容と対応する個別スクリプトを示します。 --> |
| 50 | `050_CpuCacheSizes.vbs` | CPUキャッシュ容量 <!-- 解説: この行はNo.50の取得内容と対応する個別スクリプトを示します。 --> |
| 51 | `051_CpuSocket.vbs` | CPUソケット識別情報 <!-- 解説: この行はNo.51の取得内容と対応する個別スクリプトを示します。 --> |
| 52 | `052_CpuArchitecture.vbs` | CPUアーキテクチャ詳細 <!-- 解説: この行はNo.52の取得内容と対応する個別スクリプトを示します。 --> |
| 53 | `053_CpuVirtualization.vbs` | CPU仮想化支援状態 <!-- 解説: この行はNo.53の取得内容と対応する個別スクリプトを示します。 --> |
| 54 | `054_CpuPerCoreUsage.vbs` | 論理プロセッサごとの使用率 <!-- 解説: この行はNo.54の取得内容と対応する個別スクリプトを示します。 --> |
| 55 | `055_CpuThreadLoad.vbs` | CPUスレッド負荷カウンター <!-- 解説: この行はNo.55の取得内容と対応する個別スクリプトを示します。 --> |
| 56 | `056_PhysicalMemoryTotal.vbs` | 物理メモリ総容量 <!-- 解説: この行はNo.56の取得内容と対応する個別スクリプトを示します。 --> |
| 57 | `057_AvailableMemory.vbs` | 利用可能メモリ量 <!-- 解説: この行はNo.57の取得内容と対応する個別スクリプトを示します。 --> |
| 58 | `058_MemoryStandardSpeed.vbs` | メモリ規格と速度 <!-- 解説: この行はNo.58の取得内容と対応する個別スクリプトを示します。 --> |
| 59 | `059_MemorySlotCapacity.vbs` | DIMMごとの容量 <!-- 解説: この行はNo.59の取得内容と対応する個別スクリプトを示します。 --> |
| 60 | `060_MemoryFormFactor.vbs` | メモリフォームファクター <!-- 解説: この行はNo.60の取得内容と対応する個別スクリプトを示します。 --> |
| 61 | `061_PageFileAllocation.vbs` | ページファイル割り当てと配置 <!-- 解説: この行はNo.61の取得内容と対応する個別スクリプトを示します。 --> |
| 62 | `062_MemoryEcc.vbs` | メモリエラー訂正情報 <!-- 解説: この行はNo.62の取得内容と対応する個別スクリプトを示します。 --> |
| 63 | `063_MemoryManufacturerSerial.vbs` | メモリメーカーとシリアル番号 <!-- 解説: この行はNo.63の取得内容と対応する個別スクリプトを示します。 --> |
| 64 | `064_MemoryThresholdData.vbs` | しきい値超過判定用データ（通知なし） <!-- 解説: この行はNo.64の取得内容と対応する個別スクリプトを示します。 --> |
| 65 | `065_MemoryPerformanceCounters.vbs` | ページング等の性能カウンター（帯域幅そのものではありません） <!-- 解説: この行はNo.65の取得内容と対応する個別スクリプトを示します。 --> |
| 66 | `066_DiskModelFirmware.vbs` | ディスク型番とファームウェア <!-- 解説: この行はNo.66の取得内容と対応する個別スクリプトを示します。 --> |
| 67 | `067_DiskInterface.vbs` | ディスクインターフェース <!-- 解説: この行はNo.67の取得内容と対応する個別スクリプトを示します。 --> |
| 68 | `068_DiskCapacitySectors.vbs` | ディスク容量とセクターサイズ <!-- 解説: この行はNo.68の取得内容と対応する個別スクリプトを示します。 --> |
| 69 | `069_DriveFileSystems.vbs` | 論理ドライブのファイルシステム <!-- 解説: この行はNo.69の取得内容と対応する個別スクリプトを示します。 --> |
| 70 | `070_PartitionUsage.vbs` | パーティション総量/使用量/空き容量 <!-- 解説: この行はNo.70の取得内容と対応する個別スクリプトを示します。 --> |
| 71 | `071_VolumeLabelsSerials.vbs` | ボリュームラベルとシリアル番号 <!-- 解説: この行はNo.71の取得内容と対応する個別スクリプトを示します。 --> |
| 72 | `072_DriveTypeClassification.vbs` | 固定/リムーバブル/ネットワークドライブ判定 <!-- 解説: この行はNo.72の取得内容と対応する個別スクリプトを示します。 --> |
| 73 | `073_SmartStatus.vbs` | WMIプロバイダーが公開するS.M.A.R.T.予測状態 <!-- 解説: この行はNo.73の取得内容と対応する個別スクリプトを示します。 --> |
| 74 | `074_DiskFragmentationAnalysis.vbs` | ディスク断片化分析 <!-- 解説: この行はNo.74の取得内容と対応する個別スクリプトを示します。 --> |
| 75 | `075_BitLockerVolumeStatus.vbs` | BitLockerドライブ保護状態 <!-- 解説: この行はNo.75の取得内容と対応する個別スクリプトを示します。 --> |
| 76 | `076_NetworkAdapterStatus.vbs` | ネットワークアダプター状態 <!-- 解説: この行はNo.76の取得内容と対応する個別スクリプトを示します。 --> |
| 77 | `077_NetworkMacAddresses.vbs` | ネットワークMACアドレス <!-- 解説: この行はNo.77の取得内容と対応する個別スクリプトを示します。 --> |
| 78 | `078_IPv4Subnet.vbs` | IPv4アドレスとサブネット <!-- 解説: この行はNo.78の取得内容と対応する個別スクリプトを示します。 --> |
| 79 | `079_IPv6Addresses.vbs` | IPv6アドレス割り当て <!-- 解説: この行はNo.79の取得内容と対応する個別スクリプトを示します。 --> |
| 80 | `080_DefaultGateways.vbs` | デフォルトゲートウェイ <!-- 解説: この行はNo.80の取得内容と対応する個別スクリプトを示します。 --> |
| 81 | `081_DnsServers.vbs` | DNSサーバー設定 <!-- 解説: この行はNo.81の取得内容と対応する個別スクリプトを示します。 --> |
| 82 | `082_DhcpStatus.vbs` | DHCP割り当て状態 <!-- 解説: この行はNo.82の取得内容と対応する個別スクリプトを示します。 --> |
| 83 | `083_NetworkLinkSpeed.vbs` | ネットワークリンク速度 <!-- 解説: この行はNo.83の取得内容と対応する個別スクリプトを示します。 --> |
| 84 | `084_ActiveNetworkInterfaces.vbs` | アクティブなネットワークインターフェース <!-- 解説: この行はNo.84の取得内容と対応する個別スクリプトを示します。 --> |
| 85 | `085_DnsLookup.vbs` | ホスト名DNSルックアップ <!-- 解説: この行はNo.85の取得内容と対応する個別スクリプトを示します。 --> |
| 86 | `086_GpuModelManufacturer.vbs` | GPU型番とメーカー <!-- 解説: この行はNo.86の取得内容と対応する個別スクリプトを示します。 --> |
| 87 | `087_GpuVram.vbs` | GPUビデオメモリ容量 <!-- 解説: この行はNo.87の取得内容と対応する個別スクリプトを示します。 --> |
| 88 | `088_GpuDriverInfo.vbs` | GPUドライバ情報 <!-- 解説: この行はNo.88の取得内容と対応する個別スクリプトを示します。 --> |
| 89 | `089_DisplayResolution.vbs` | ディスプレイ解像度 <!-- 解説: この行はNo.89の取得内容と対応する個別スクリプトを示します。 --> |
| 90 | `090_DisplayColorRefresh.vbs` | 色深度とリフレッシュレート <!-- 解説: この行はNo.90の取得内容と対応する個別スクリプトを示します。 --> |
| 91 | `091_MonitorCountLayout.vbs` | モニター認識数と情報（配置座標はWMI非対応） <!-- 解説: この行はNo.91の取得内容と対応する個別スクリプトを示します。 --> |
| 92 | `092_MonitorPhysicalInfo.vbs` | モニター物理情報 <!-- 解説: この行はNo.92の取得内容と対応する個別スクリプトを示します。 --> |
| 93 | `093_HardwareAccelerationInfo.vbs` | GPU状態診断（有効/無効を断定できません） <!-- 解説: この行はNo.93の取得内容と対応する個別スクリプトを示します。 --> |
| 94 | `094_GpuTypeIdentifiers.vbs` | 統合/外部GPU識別用情報 <!-- 解説: この行はNo.94の取得内容と対応する個別スクリプトを示します。 --> |
| 95 | `095_DirectXVersion.vbs` | DirectXバージョン <!-- 解説: この行はNo.95の取得内容と対応する個別スクリプトを示します。 --> |
| 96 | `096_AudioDevices.vbs` | サウンドデバイス <!-- 解説: この行はNo.96の取得内容と対応する個別スクリプトを示します。 --> |
| 97 | `097_AudioEndpoints.vbs` | オーディオ入出力エンドポイント状態 <!-- 解説: この行はNo.97の取得内容と対応する個別スクリプトを示します。 --> |
| 98 | `098_WebCameras.vbs` | Webカメラデバイス <!-- 解説: この行はNo.98の取得内容と対応する個別スクリプトを示します。 --> |
| 99 | `099_KeyboardLayout.vbs` | キーボードとレイアウト <!-- 解説: この行はNo.99の取得内容と対応する個別スクリプトを示します。 --> |
| 100 | `100_PointingDevices.vbs` | ポインティングデバイス <!-- 解説: この行はNo.100の取得内容と対応する個別スクリプトを示します。 --> |
| 101 | `101_GameControllersHid.vbs` | HIDおよびコントローラーデバイス <!-- 解説: この行はNo.101の取得内容と対応する個別スクリプトを示します。 --> |
| 102 | `102_TouchScreenSupport.vbs` | タッチスクリーン対応デバイス <!-- 解説: この行はNo.102の取得内容と対応する個別スクリプトを示します。 --> |
| 103 | `103_BluetoothDevices.vbs` | 検出済みBluetooth PnPデバイス（ペアリング状態ではありません） <!-- 解説: この行はNo.103の取得内容と対応する個別スクリプトを示します。 --> |
| 104 | `104_SmartCardReaders.vbs` | スマートカードリーダー <!-- 解説: この行はNo.104の取得内容と対応する個別スクリプトを示します。 --> |
| 105 | `105_BiometricDevices.vbs` | 生体認証デバイス <!-- 解説: この行はNo.105の取得内容と対応する個別スクリプトを示します。 --> |
| 106 | `106_BatteryChargePercent.vbs` | バッテリー残量 <!-- 解説: この行はNo.106の取得内容と対応する個別スクリプトを示します。 --> |
| 107 | `107_BatteryChargingState.vbs` | バッテリー充電状態 <!-- 解説: この行はNo.107の取得内容と対応する個別スクリプトを示します。 --> |
| 108 | `108_BatteryCapacityCompare.vbs` | 設計容量と満充電容量 <!-- 解説: この行はNo.108の取得内容と対応する個別スクリプトを示します。 --> |
| 109 | `109_BatteryHealth.vbs` | バッテリー劣化度算出 <!-- 解説: この行はNo.109の取得内容と対応する個別スクリプトを示します。 --> |
| 110 | `110_ActivePowerPlan.vbs` | 現在の電源プラン <!-- 解説: この行はNo.110の取得内容と対応する個別スクリプトを示します。 --> |
| 111 | `111_SleepHibernateSupport.vbs` | スリープ/休止状態対応 <!-- 解説: この行はNo.111の取得内容と対応する個別スクリプトを示します。 --> |
| 112 | `112_AcpiStatus.vbs` | ACPIデバイス状態 <!-- 解説: この行はNo.112の取得内容と対応する個別スクリプトを示します。 --> |
| 113 | `113_PowerTimeouts.vbs` | 画面オフとスリープ時間設定 <!-- 解説: この行はNo.113の取得内容と対応する個別スクリプトを示します。 --> |
| 114 | `114_UpsStatus.vbs` | UPSデバイス状態 <!-- 解説: この行はNo.114の取得内容と対応する個別スクリプトを示します。 --> |
| 115 | `115_LastWakeReason.vbs` | 直近の電源復帰理由 <!-- 解説: この行はNo.115の取得内容と対応する個別スクリプトを示します。 --> |
| 116 | `116_WindowsEdition.vbs` | Windowsエディション <!-- 解説: この行はNo.116の取得内容と対応する個別スクリプトを示します。 --> |
| 117 | `117_WindowsBuildRevision.vbs` | Windowsビルドとリビジョン <!-- 解説: この行はNo.117の取得内容と対応する個別スクリプトを示します。 --> |
| 118 | `118_WindowsInstallDates.vbs` | OSインストール日時 <!-- 解説: この行はNo.118の取得内容と対応する個別スクリプトを示します。 --> |
| 119 | `119_WindowsInstallFolders.vbs` | Windowsインストール先フォルダー <!-- 解説: この行はNo.119の取得内容と対応する個別スクリプトを示します。 --> |
| 120 | `120_WindowsLocale.vbs` | 既定言語とロケール <!-- 解説: この行はNo.120の取得内容と対応する個別スクリプトを示します。 --> |
| 121 | `121_WindowsPartialProductKey.vbs` | 一部ライセンス情報の末尾5文字（取扱注意） <!-- 解説: この行はNo.121の取得内容と対応する個別スクリプトを示します。 --> |
| 122 | `122_WindowsLicenseType.vbs` | Windowsライセンス種別 <!-- 解説: この行はNo.122の取得内容と対応する個別スクリプトを示します。 --> |
| 123 | `123_WindowsArchitecture.vbs` | OSアーキテクチャ <!-- 解説: この行はNo.123の取得内容と対応する個別スクリプトを示します。 --> |
| 124 | `124_SafeModeStatus.vbs` | セーフモード起動状態 <!-- 解説: この行はNo.124の取得内容と対応する個別スクリプトを示します。 --> |
| 125 | `125_BootModeStatus.vbs` | システム起動モード <!-- 解説: この行はNo.125の取得内容と対応する個別スクリプトを示します。 --> |
| 126 | `126_TpmStatus.vbs` | TPM搭載/有効状態 <!-- 解説: この行はNo.126の取得内容と対応する個別スクリプトを示します。 --> |
| 127 | `127_SecureBootStatus.vbs` | Secure Boot状態 <!-- 解説: この行はNo.127の取得内容と対応する個別スクリプトを示します。 --> |
| 128 | `128_BitLockerProtection.vbs` | BitLocker保護状態 <!-- 解説: この行はNo.128の取得内容と対応する個別スクリプトを示します。 --> |
| 129 | `129_BitLockerRecoveryKeyIds.vbs` | BitLocker回復キーID <!-- 解説: この行はNo.129の取得内容と対応する個別スクリプトを示します。 --> |
| 130 | `130_DefenderStatus.vbs` | Defenderリアルタイム保護と定義バージョン <!-- 解説: この行はNo.130の取得内容と対応する個別スクリプトを示します。 --> |
| 131 | `131_FirewallProfiles.vbs` | ファイアウォールプロファイル状態 <!-- 解説: この行はNo.131の取得内容と対応する個別スクリプトを示します。 --> |
| 132 | `132_UacNotificationSettings.vbs` | UAC通知設定 <!-- 解説: この行はNo.132の取得内容と対応する個別スクリプトを示します。 --> |
| 133 | `133_VbsHvciStatus.vbs` | VBS/HVCI稼働状態 <!-- 解説: この行はNo.133の取得内容と対応する個別スクリプトを示します。 --> |
| 134 | `134_ParentalControlStatus.vbs` | ペアレンタルコントロール設定有無 <!-- 解説: この行はNo.134の取得内容と対応する個別スクリプトを示します。 --> |
| 135 | `135_ThirdPartyAntivirus.vbs` | サードパーティ製ウイルス対策状態 <!-- 解説: この行はNo.135の取得内容と対応する個別スクリプトを示します。 --> |

# vbs

## System information scripts

The following scripts query local OS, hardware, storage, and peripheral
information through WMI, Windows Script Host APIs, PowerShell, or built-in
Windows commands. Run an individual script from Command Prompt:

```text
cscript //nologo GetOSInfo.vbs
```

Other scripts in this directory cover architecture, computer name, join
status, service pack, uptime, CPU and memory, page file, motherboard, BIOS,
disks, volumes, display, audio, USB devices, printers, battery, input devices,
and time-zone information. Each script prints results locally; none uploads
data.

The original basic set does not collect clipboard contents, environment
variables, activation details, security-control status, network
identifiers/settings, installed application or update inventories, or system
logs. The additional read-only inventory scripts below cover selected network,
license, and security-status fields requested for local diagnostics; they do
not upload data. Their output can contain device identifiers or security
posture, so handle it accordingly. The computer-name script is read-only and
does not rename the machine.

## Additional system and hardware inventory (No. 36-135)

The following 100 scripts are separate, read-only local diagnostics.
Availability and values depend on Windows edition, hardware, driver, WMI
provider, and permissions. Some commands require PowerShell or built-in tools;
output stays on the local console. Product-key suffixes and device identifiers
may be sensitive; handle output accordingly. The BitLocker script reports
protector IDs only, never recovery passwords.

| No. | Script | Information |
|---:|---|---|
| 36 | `036_MotherboardManufacturer.vbs` | マザーボード製造元  |
| 37 | `037_MotherboardProduct.vbs` | マザーボード製品型番  |
| 38 | `038_MotherboardSerial.vbs` | マザーボードシリアル番号  |
| 39 | `039_BiosVendorVersion.vbs` | BIOSベンダーとバージョン  |
| 40 | `040_BiosReleaseDate.vbs` | BIOSリリース日時  |
| 41 | `041_FirmwareBootType.vbs` | UEFI/レガシーブート方式  |
| 42 | `042_ChassisType.vbs` | 筐体タイプ  |
| 43 | `043_ChassisIdentifiers.vbs` | 筐体シリアル番号と資産タグ  |
| 44 | `044_SystemFamilySku.vbs` | システムファミリとSKU  |
| 45 | `045_PciSlots.vbs` | PCI/PCIeスロット構成  |
| 46 | `046_CpuProcessorName.vbs` | CPU製品名  |
| 47 | `047_CpuManufacturer.vbs` | CPUメーカー  |
| 48 | `048_ProcessorCounts.vbs` | 物理/論理プロセッサ数  |
| 49 | `049_CpuClockFrequencies.vbs` | CPU最大/現在クロック周波数  |
| 50 | `050_CpuCacheSizes.vbs` | CPUキャッシュ容量  |
| 51 | `051_CpuSocket.vbs` | CPUソケット識別情報  |
| 52 | `052_CpuArchitecture.vbs` | CPUアーキテクチャ詳細  |
| 53 | `053_CpuVirtualization.vbs` | CPU仮想化支援状態  |
| 54 | `054_CpuPerCoreUsage.vbs` | 論理プロセッサごとの使用率  |
| 55 | `055_CpuThreadLoad.vbs` | CPUスレッド負荷カウンター  |
| 56 | `056_PhysicalMemoryTotal.vbs` | 物理メモリ総容量  |
| 57 | `057_AvailableMemory.vbs` | 利用可能メモリ量  |
| 58 | `058_MemoryStandardSpeed.vbs` | メモリ規格と速度  |
| 59 | `059_MemorySlotCapacity.vbs` | DIMMごとの容量  |
| 60 | `060_MemoryFormFactor.vbs` | メモリフォームファクター  |
| 61 | `061_PageFileAllocation.vbs` | ページファイル割り当てと配置  |
| 62 | `062_MemoryEcc.vbs` | メモリエラー訂正情報  |
| 63 | `063_MemoryManufacturerSerial.vbs` | メモリメーカーとシリアル番号  |
| 64 | `064_MemoryThresholdData.vbs` | しきい値超過判定用データ（通知なし）  |
| 65 | `065_MemoryPerformanceCounters.vbs` | ページング等の性能カウンター（帯域幅そのものではありません）  |
| 66 | `066_DiskModelFirmware.vbs` | ディスク型番とファームウェア  |
| 67 | `067_DiskInterface.vbs` | ディスクインターフェース  |
| 68 | `068_DiskCapacitySectors.vbs` | ディスク容量とセクターサイズ  |
| 69 | `069_DriveFileSystems.vbs` | 論理ドライブのファイルシステム  |
| 70 | `070_PartitionUsage.vbs` | パーティション総量/使用量/空き容量  |
| 71 | `071_VolumeLabelsSerials.vbs` | ボリュームラベルとシリアル番号  |
| 72 | `072_DriveTypeClassification.vbs` | 固定/リムーバブル/ネットワークドライブ判定  |
| 73 | `073_SmartStatus.vbs` | WMIプロバイダーが公開するS.M.A.R.T.予測状態  |
| 74 | `074_DiskFragmentationAnalysis.vbs` | ディスク断片化分析  |
| 75 | `075_BitLockerVolumeStatus.vbs` | BitLockerドライブ保護状態  |
| 76 | `076_NetworkAdapterStatus.vbs` | ネットワークアダプター状態  |
| 77 | `077_NetworkMacAddresses.vbs` | ネットワークMACアドレス  |
| 78 | `078_IPv4Subnet.vbs` | IPv4アドレスとサブネット  |
| 79 | `079_IPv6Addresses.vbs` | IPv6アドレス割り当て  |
| 80 | `080_DefaultGateways.vbs` | デフォルトゲートウェイ  |
| 81 | `081_DnsServers.vbs` | DNSサーバー設定  |
| 82 | `082_DhcpStatus.vbs` | DHCP割り当て状態  |
| 83 | `083_NetworkLinkSpeed.vbs` | ネットワークリンク速度  |
| 84 | `084_ActiveNetworkInterfaces.vbs` | アクティブなネットワークインターフェース  |
| 85 | `085_DnsLookup.vbs` | ホスト名DNSルックアップ  |
| 86 | `086_GpuModelManufacturer.vbs` | GPU型番とメーカー  |
| 87 | `087_GpuVram.vbs` | GPUビデオメモリ容量  |
| 88 | `088_GpuDriverInfo.vbs` | GPUドライバ情報  |
| 89 | `089_DisplayResolution.vbs` | ディスプレイ解像度  |
| 90 | `090_DisplayColorRefresh.vbs` | 色深度とリフレッシュレート  |
| 91 | `091_MonitorCountLayout.vbs` | モニター認識数と情報（配置座標はWMI非対応）  |
| 92 | `092_MonitorPhysicalInfo.vbs` | モニター物理情報  |
| 93 | `093_HardwareAccelerationInfo.vbs` | GPU状態診断（有効/無効を断定できません）  |
| 94 | `094_GpuTypeIdentifiers.vbs` | 統合/外部GPU識別用情報  |
| 95 | `095_DirectXVersion.vbs` | DirectXバージョン  |
| 96 | `096_AudioDevices.vbs` | サウンドデバイス  |
| 97 | `097_AudioEndpoints.vbs` | オーディオ入出力エンドポイント状態  |
| 98 | `098_WebCameras.vbs` | Webカメラデバイス  |
| 99 | `099_KeyboardLayout.vbs` | キーボードとレイアウト  |
| 100 | `100_PointingDevices.vbs` | ポインティングデバイス  |
| 101 | `101_GameControllersHid.vbs` | HIDおよびコントローラーデバイス  |
| 102 | `102_TouchScreenSupport.vbs` | タッチスクリーン対応デバイス  |
| 103 | `103_BluetoothDevices.vbs` | 検出済みBluetooth PnPデバイス（ペアリング状態ではありません）  |
| 104 | `104_SmartCardReaders.vbs` | スマートカードリーダー  |
| 105 | `105_BiometricDevices.vbs` | 生体認証デバイス  |
| 106 | `106_BatteryChargePercent.vbs` | バッテリー残量  |
| 107 | `107_BatteryChargingState.vbs` | バッテリー充電状態  |
| 108 | `108_BatteryCapacityCompare.vbs` | 設計容量と満充電容量  |
| 109 | `109_BatteryHealth.vbs` | バッテリー劣化度算出  |
| 110 | `110_ActivePowerPlan.vbs` | 現在の電源プラン  |
| 111 | `111_SleepHibernateSupport.vbs` | スリープ/休止状態対応  |
| 112 | `112_AcpiStatus.vbs` | ACPIデバイス状態  |
| 113 | `113_PowerTimeouts.vbs` | 画面オフとスリープ時間設定  |
| 114 | `114_UpsStatus.vbs` | UPSデバイス状態  |
| 115 | `115_LastWakeReason.vbs` | 直近の電源復帰理由  |
| 116 | `116_WindowsEdition.vbs` | Windowsエディション  |
| 117 | `117_WindowsBuildRevision.vbs` | Windowsビルドとリビジョン  |
| 118 | `118_WindowsInstallDates.vbs` | OSインストール日時  |
| 119 | `119_WindowsInstallFolders.vbs` | Windowsインストール先フォルダー  |
| 120 | `120_WindowsLocale.vbs` | 既定言語とロケール  |
| 121 | `121_WindowsPartialProductKey.vbs` | 一部ライセンス情報の末尾5文字（取扱注意）  |
| 122 | `122_WindowsLicenseType.vbs` | Windowsライセンス種別  |
| 123 | `123_WindowsArchitecture.vbs` | OSアーキテクチャ  |
| 124 | `124_SafeModeStatus.vbs` | セーフモード起動状態  |
| 125 | `125_BootModeStatus.vbs` | システム起動モード  |
| 126 | `126_TpmStatus.vbs` | TPM搭載/有効状態  |
| 127 | `127_SecureBootStatus.vbs` | Secure Boot状態  |
| 128 | `128_BitLockerProtection.vbs` | BitLocker保護状態  |
| 129 | `129_BitLockerRecoveryKeyIds.vbs` | BitLocker回復キーID  |
| 130 | `130_DefenderStatus.vbs` | Defenderリアルタイム保護と定義バージョン  |
| 131 | `131_FirewallProfiles.vbs` | ファイアウォールプロファイル状態  |
| 132 | `132_UacNotificationSettings.vbs` | UAC通知設定  |
| 133 | `133_VbsHvciStatus.vbs` | VBS/HVCI稼働状態  |
| 134 | `134_ParentalControlStatus.vbs` | ペアレンタルコントロール設定有無  |
| 135 | `135_ThirdPartyAntivirus.vbs` | サードパーティ製ウイルス対策状態  |

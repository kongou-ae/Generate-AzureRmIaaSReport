Param(
  [switch]$noLogin
)

$ErrorActionPreference = "Stop"

Function Write-Log {
  param(
    [string]$Message,
    [string]$Color = 'White'
  )

  $Date = Get-Date -Format "yyyy-MM-dd HH:mm:ss"
  Write-Host "[$Date] $Message"-ForegroundColor $Color
}

function Get-nicInfo {
  param(
    [string]$nicId,
    [System.Object[]]$nicList,
    [System.Object[]]$pipList
  )

  $nicList | ForEach-Object {
    if ( $_.Id -eq $nicId ) {
      $privateIP = $_.IpConfigurations[0].PrivateIpAddress
      $pipId = $_.IpConfigurations[0].PublicIpAddress.id
      $dns = $_.DnsSettings.DnsServers -join ","
      $pipList | Foreach-object {
        if ( $_.Id -eq $pipId ) {
          $publicIP = $_.IpAddress
        }
      }
    }
  }
  return $privateIP, $publicIP, $dns
}

function Get-vnetInfo {
  param(
    [string]$nicId,
    [Microsoft.Azure.Commands.Network.Models.PSVirtualNetwork]$vnetList
  )

  $vnetList | ForEach-Object {
    $tmpVnetName = $_.Name
    $_.Subnets | ForEach-Object {
      $tmpSubnetName = $_.Name
      $_.IpConfigurations | ForEach-Object {
        if ($_.Id -match $nicId) {
          $vnetName = $tmpVnetName
          $subnetName = $tmpSubnetName
        }
      }
    }
  }
  return $vnetName, $subnetName
}
 
function Generate-VirtualMachineReport {
  param(
    [Microsoft.Office.Interop.Excel.ApplicationClass]$excel
  )  
  $sheet = $excel.Worksheets.Item("VirtualMachine")
  $LineStyle = "microsoft.office.interop.excel.xlLineStyle" -as [type]
  $vmHeader = @(
    "Seq",
    "PowerState",
    "ResourceGroupName",
    "Name",
    "OsType"
    "Location",
    "AvailabilitySet",
    "Vmsize",
    "Vnet",
    "Subnet",
    "Private IP",
    "Public IP"
  )
  
  $i = 1
  $vmHeader | ForEach-Object {
    $sheet.Cells.Item(1, $i) = $_
    $sheet.cells.item(1, $i).borders.LineStyle = $LineStyle::xlContinuous
    $sheet.cells.item(1, $i).interior.ColorIndex = 49
    $sheet.cells.item(1, $i).Font.ColorIndex = 2
    $i = $i + 1
  }

  $i = 2
  $vmList | ForEach-Object {
    $sheet.Cells.Item($i, 1) = $i - 1 
    $sheet.Cells.Item($i, 2) = $_.PowerState
    $sheet.Cells.Item($i, 3) = $_.ResourceGroupName
    $sheet.Cells.Item($i, 4) = $_.Name
    $sheet.Cells.Item($i, 5) = [string]$_.StorageProfile.OsDisk.OsType
    $sheet.Cells.Item($i, 6) = $_.Location
    #$sheet.Cells.Item($i, 7) = $_.AvailabilitySetReference 
    $sheet.Cells.Item($i, 8) = $_.HardwareProfile.VmSize  
    $vmVnetInfo = Get-vnetInfo $_.NetworkProfile.NetworkInterfaces[0].Id $vnetList
    $vmNicInfo = Get-nicInfo $_.NetworkProfile.NetworkInterfaces[0].Id $nicList $PIPList
    $sheet.Cells.Item($i, 9) = $vmVnetInfo[0] # Vnet
    $sheet.Cells.Item($i, 10) = $vmVnetInfo[1] # subnet
    $sheet.Cells.Item($i, 11) = $vmNicInfo[0] # privateIP
    $sheet.Cells.Item($i, 12) = $vmNicInfo[1] # publicIP

    for ($j = 1; $j -lt 13; $j++) {
      $sheet.cells.item($i, $j).borders.LineStyle = $LineStyle::xlContinuous
    }

    $i = $i + 1
  }
  $sheet.Columns.AutoFit()
  $sheet.UsedRange.Font.Name = "Meiryo UI"
}

function Generate-storageAccountReport {
  param(
    [Microsoft.Office.Interop.Excel.ApplicationClass]$excel
  )  
  $sheet = $excel.Worksheets.Item("StorageAccount")
  $LineStyle = "microsoft.office.interop.excel.xlLineStyle" -as [type]

  $storageAccountHeader = @(
    "Seq",
    "ResourceGroupName",
    "StorageAccountName",
    "Location",
    "Sku",
    "PrimaryLocation",
    "StatusOfPrimary",
    "SecondaryLocation",
    "StatusOfSecondary",
    "PrimaryEndpoints"
  )

  $i = 1
  $storageAccountHeader | ForEach-Object {
    $sheet.Cells.Item(1, $i) = $_
    $sheet.cells.item(1, $i).borders.LineStyle = $LineStyle::xlContinuous
    $sheet.cells.item(1, $i).interior.ColorIndex = 49
    $sheet.cells.item(1, $i).Font.ColorIndex = 2
    $i = $i + 1
  }

  $i = 2
  $storageAccountList | ForEach-Object {
    $sheet.Cells.Item($i, 1) = $i - 1 
    $sheet.Cells.Item($i, 2) = $_.ResourceGroupName
    $sheet.Cells.Item($i, 3) = $_.StorageAccountName
    $sheet.Cells.Item($i, 4) = $_.Location
    $sheet.Cells.Item($i, 5) = [string]$_.Sku.Name
    $sheet.Cells.Item($i, 6) = $_.PrimaryLocation
    $sheet.Cells.Item($i, 7) = [string]$_.StatusOfPrimary 
    $sheet.Cells.Item($i, 8) = $_.SecondaryLocation 
    $sheet.Cells.Item($i, 9) = $_.StatusOfSecondary 
    $sheet.Cells.Item($i, 10) = $_.PrimaryEndpoints.Blob

    for ($j = 1; $j -lt 11; $j++) {
      $sheet.cells.item($i, $j).borders.LineStyle = $LineStyle::xlContinuous
    }

    $i = $i + 1
  }
  $sheet.Columns.AutoFit()
  $sheet.UsedRange.Font.Name = "Meiryo UI"
}

function Generate-diskReport {
  param(
    [Microsoft.Office.Interop.Excel.ApplicationClass]$excel
  )  
  $sheet = $excel.Worksheets.Item("Disk")
  $LineStyle = "microsoft.office.interop.excel.xlLineStyle" -as [type]

  $diskHeader = @(
    "Seq",
    "ManagedBy",
    "AttachedPoint",
    "ResourceGroupName",
    "Name",
    "Type",
    "Location",
    "Sku",
    "DiskSizeGB"
  )

  $i = 1
  $diskHeader | ForEach-Object {
    $sheet.Cells.Item(1, $i) = $_
    $sheet.cells.item(1, $i).borders.LineStyle = $LineStyle::xlContinuous
    $sheet.cells.item(1, $i).interior.ColorIndex = 49
    $sheet.cells.item(1, $i).Font.ColorIndex = 2
    $i = $i + 1
  }

  $i = 2
  $vmList | ForEach-Object {
    $vmName = $_.Name
    if ( $_.StorageProfile.OsDisk.Vhd -eq "") {
      Write-Log "VDH is not supported"
    }
    else {
      $OSDiskId = $_.StorageProfile.OsDisk.ManagedDisk.Id
      $DataDiskList = $_.StorageProfile.DataDisks

      $diskList | ForEach-Object {
        if ($_.id -eq $OSDiskId) {
          $sheet.Cells.Item($i, 1) = $i - 1 
          $sheet.Cells.Item($i, 2) = $vmName
          $sheet.Cells.Item($i, 3) = "OSDisk"
          $sheet.Cells.Item($i, 4) = $_.ResourceGroupName
          $sheet.Cells.Item($i, 5) = $_.Name
          $sheet.Cells.Item($i, 6) = $_.Type
          $sheet.Cells.Item($i, 7) = $_.Location
          $sheet.Cells.Item($i, 8) = [string]$_.Sku.Name
          $sheet.Cells.Item($i, 9) = $_.DiskSizeGB
          
          for ($j = 1; $j -lt 10; $j++) {
            $sheet.cells.item($i, $j).borders.LineStyle = $LineStyle::xlContinuous
          }

          $i = $i + 1
        }

        $tmpDisk = $_
        $DataDiskList | ForEach-Object {
          if ($tmpDisk.Id -eq $_.ManagedDisk.id) {
            $sheet.Cells.Item($i, 1) = $i - 1 
            $sheet.Cells.Item($i, 2) = $vmName
            $sheet.Cells.Item($i, 3) = "DataDisk Lun" + [string]$_.Lun
            $sheet.Cells.Item($i, 4) = $tmpDisk.ResourceGroupName
            $sheet.Cells.Item($i, 5) = $tmpDisk.Name
            $sheet.Cells.Item($i, 6) = $tmpDisk.Type
            $sheet.Cells.Item($i, 7) = $tmpDisk.Location
            $sheet.Cells.Item($i, 8) = [string]$tmpDisk.Sku.Name
            $sheet.Cells.Item($i, 9) = $tmpDisk.DiskSizeGB       

            for ($j = 1; $j -lt 10; $j++) {
              $sheet.cells.item($i, $j).borders.LineStyle = $LineStyle::xlContinuous
            }

            $i = $i + 1   
          }        
        }
      }
    }
  }
  $sheet.Columns.AutoFit()
  $sheet.UsedRange.Font.Name = "Meiryo UI"  
}

function Generate-backupReport {
  param(
    [Microsoft.Office.Interop.Excel.ApplicationClass]$excel
  )  
  $sheet = $excel.Worksheets.Item("VM Backup")
  $LineStyle = "microsoft.office.interop.excel.xlLineStyle" -as [type]

  $backupHeader = @(
    "Seq",
    "ResourceGroupName",
    "RecoveryServicesVault",
    "ProtectedVM",
    "ProtectionState",
    "LastBackupTime",
    "LastBackupStatus",
    "ProtectionPolicyName"
  )

  $i = 1
  $backupHeader | ForEach-Object {
    $sheet.Cells.Item(1, $i) = $_
    $sheet.cells.item(1, $i).borders.LineStyle = $LineStyle::xlContinuous
    $sheet.cells.item(1, $i).interior.ColorIndex = 49
    $sheet.cells.item(1, $i).Font.ColorIndex = 2
    $i = $i + 1
  }
  
  $i = 2
  $backupItemList | ForEach-Object {
    $sheet.Cells.Item($i, 1) = $i - 1 
    $_.Id -match "/resourceGroups/(.*)/providers/Microsoft.RecoveryServices"
    $sheet.Cells.Item($i, 2) = $Matches[1]
    $_.Id -match "/Microsoft.RecoveryServices/vaults/(.*)/backupFabrics/"
    $sheet.Cells.Item($i, 3) = $Matches[1]
    $_.SourceResourceId -match "Microsoft.Compute/virtualMachines/(.*)$"    
    $sheet.Cells.Item($i, 4) = $Matches[1]
    $sheet.Cells.Item($i, 5) = [string]$_.ProtectionState
    $sheet.Cells.Item($i, 6) = $_.LastBackupTime
    $sheet.Cells.Item($i, 7) = $_.LastBackupStatus
    $sheet.Cells.Item($i, 8) = $_.ProtectionPolicyName

    for ($j = 1; $j -lt 9; $j++) {
      $sheet.cells.item($i, $j).borders.LineStyle = $LineStyle::xlContinuous
    }

    $i = $i + 1
  }
  $sheet.Columns.AutoFit()
  $sheet.UsedRange.Font.Name = "Meiryo UI"
}

if ($noLogin -ne $True) {
  Write-Log "Please login to Azure Active Directory"
  Login-AzureRmAccount
  Write-Log "Please select your subscription"
  $subscription = Get-AzureRmSubscription | Out-GridView -PassThru
  Select-AzureRmSubscription -SubscriptionObject $subscription  
}

$vmList = Get-AzureRmVm -Status
$nicList = Get-AzureRmNetworkInterface
$PIPList = Get-AzureRmPublicIpAddress
$vnetList = Get-AzureRmVirtualNetwork
$storageAccountList = Get-AzureRmStorageAccount
$diskList = Get-AzureRmDisk
$backupItemList = @()
$vaultList = Get-AzureRmRecoveryServicesVault
$vaultList | ForEach-Object {
  Set-AzureRmRecoveryServicesVaultContext -Vault $_
  $containerList = Get-AzureRmRecoveryServicesBackupContainer -ContainerType AzureVM
  $containerList | ForEach-Object {
    $tmpBackupItemList = Get-AzureRmRecoveryServicesBackupItem -Container $_ -WorkloadType AzureVM
    $tmpBackupItemList | ForEach-Object {
      $backupItemList += $_
    }
  }
}

$excel = New-Object -ComObject Excel.Application
$excel.Visible = $true

Write-Log "Waiting: Generate-VirtualMachineReport"
$book = $excel.Workbooks.Add() | Out-Null
$excel.WorkSheets.item(1).name = "VirtualMachine"
Generate-VirtualMachineReport $excel
Write-Log "Success: Generate-VirtualMachineReport" -Color Green

Write-Log "Waiting: Generate-storageAccountReport"
$book = $excel.WorkSheets.Add() | Out-Null
$sheetIndex = $excel.ActiveSheet.Index
$excel.WorkSheets.item($sheetIndex).name = "StorageAccount"
Generate-storageAccountReport $excel
Write-Log "Success: Generate-storageAccountReport" -Color Green

Write-Log "Waiting: Generate-diskReport"
$book = $excel.WorkSheets.Add() | Out-Null
$sheetIndex = $excel.ActiveSheet.Index
$excel.WorkSheets.item($sheetIndex).name = "Disk"
Generate-diskReport $excel
Write-Log "Success: Generate-diskReport" -Color Green

Write-Log "Waiting: Generate-backupReport"
$book = $excel.WorkSheets.Add() | Out-Null
$sheetIndex = $excel.ActiveSheet.Index
$excel.WorkSheets.item($sheetIndex).name = "VM Backup"
Generate-backupReport $excel
Write-Log "Success: Generate-backupReport" -Color Green

$filename = Get-Date -Format "yyyy-MMdd-HHmmss"
Write-Log "Waiting: Save ${HOME}\Desktop\AzureRmIaaSReport_$filename.xlsx"
$excel.ActiveWorkbook.SaveAs("${HOME}\Desktop\AzureRmIaaSReport_$filename.xlsx")
$excel.Quit()
[GC]::Collect()
Write-Log "Success: Save ${HOME}\Desktop\AzureRmIaaSReport_$filename.xlsx" -Color Green
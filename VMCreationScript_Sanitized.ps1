###################### DO NOT MODIFY ########################
#
# Issues or questions: Please contact REDACTED
# of IT Infrastructure (ext. REDACTED)
#
#         __     ____  __ ____            _       _   
#         \ \   / /  \/  / ___|  ___ _ __(_)_ __ | |_ 
#          \ \ / /| |\/| \___ \ / __| '__| | '_ \| __|
#           \ V / | |  | |___) | (__| |  | | |_) | |_ 
#            \_/  |_|  |_|____/ \___|_|  |_| .__/ \__|
#                                          |_|         
# Step 2
# Description: Controls the creation of VMs for new users.
# 1. Retrieves CSV created from Step 1 on REDACTED,
#     c:\REDACTED\newusers.csv
# 2. Uses variables therein to select a template, name, port-
#     group and folder to place the VM.
# 3. Selects both host and datastore, by fewest running VMs
#     and most available freespace, respectively.
# *Note: Step 3 occurs on REDACTED as a scheduled task*
##############################################################

#Add snapin and connect to vCenter
add-pssnapin vm*
Connect-VIServer x.x.x.x

#Customization Specification. Current is Win7 32bit.
$Customspec = "VMScript"

#Date
$date = Get-Date -DisplayHint date

#Template
$FCRSTemplate = "REDACTED"
$GoldenImage = "REDACTED"
$GoldenImage2 = "REDACTED"

#Select host (with fewest running VMs)
$fewesthost = Get-Cluster "REDACTED" | Get-VMHost |Select Name,@{N="NumVM";E={($_ |Get-VM |where {$_.PowerState -eq "PoweredOn"}).Count}}` |Sort-Object NumVM |Select-Object -first 1
$ViewHost = $fewesthost.name

#Select datastore from cluster based on greatest amount of free space, in MB
$Datastore = Get-DatastoreCluster "REDACTED" | Get-Datastore |sort -descending FreeSpaceMB | Select-object -first 1

#Formatting and details for mail message
$smtpServer = "REDACTED"
$smtpFrom = "VM_Creation@REDACTED.com"
$smtpTo = "REDACTED@REDACTED.com"
$messageSubject = "Newly Created VMs: $date"

$message = New-Object System.Net.Mail.MailMessage $smtpfrom, $smtpto
$message.Subject = $messageSubject
$message.IsBodyHTML = $true

$style = "<style>BODY{font-family: Verdana; font-size: 10pt;}"
$style = $style + "TABLE{border: 1px solid black}"
$style = $style + "TH{border: 1px solid black; background: #0099ff; padding: 5px; border-collapse: collapse }"
$style = $style + "TD{border: 1px solid black; padding: 5px; border-collapse: collapse }"
$style = $style + "</style>"

#The magic. Import CSV file which is generated each day by VMscript step 1, running on REDACTED. It retrieves any users created within the last day, and formats them into a CSV with headers samaccountname,name,
#department,office.
Import-Csv "C:\REDACTED\newusers.csv" -UseCulture | %{
$vm = "vm"
$vmuser = $_.samaccountname

  If ($_.department -eq "REDACTED"){
  #Residential foreclosure template. NOTE: This is the only template that uses a separate image, and the portgroup is set on that image thus does not need to be specified.
    New-vm -Name "$vm$vmuser" -VMhost $ViewHost -Template $FCRSTemplate -Datastore $Datastore  -OSCustomizationspec $Customspec -Location "REDACTED"
    Start-VM -VM "$vm$vmuser" -Confirm:$false
    }
    ElseIf ($_.department -eq "IT") {
    New-vm -Name "$vm$vmuser" -VMhost $ViewHost -Template $GoldenImage2 -Datastore $Datastore  -OSCustomizationspec $Customspec -Location "IT"
    Get-NetworkAdapter -VM "$vm$vmuser" | Set-NetworkAdapter -NetworkName "pg_0035-REDACTED" -startConnected:$True -Confirm:$false
    Start-VM -VM "$vm$vmuser" -Confirm:$false
    }
    ElseIf ($_.office -eq "REDACTED") {
    New-vm -Name "$vm$vmuser" -VMhost $ViewHost -Template $GoldenImage -Datastore $Datastore  -OSCustomizationspec $Customspec -Location "REDACTED"
    Get-NetworkAdapter -VM "$vm$vmuser" | Set-NetworkAdapter -NetworkName "pg_0075-REDACTED" -startConnected:$True -Confirm:$false
    Start-VM -VM "$vm$vmuser" -Confirm:$false
    }
    ElseIf ($_.office -eq "REDACTED") {
    New-vm -Name "$vm$vmuser" -VMhost $ViewHost -Template $GoldenImage -Datastore $Datastore  -OSCustomizationspec $Customspec -Location "REDACTED"
    Get-NetworkAdapter -VM "$vm$vmuser" | Set-NetworkAdapter -NetworkName "pg_0065-REDACTED" -startConnected:$True -Confirm:$false
    }
    ElseIf ($_.office -eq "REDACTED") {
    New-vm -Name "$vm$vmuser" -VMhost $ViewHost -Template $GoldenImage -Datastore $Datastore  -OSCustomizationspec $Customspec -Location "REDACTED"
    Get-NetworkAdapter -VM "$vm$vmuser" | Set-NetworkAdapter -NetworkName "pg_0085-REDACTED" -startConnected:$True -Confirm:$false
    Start-VM -VM "$vm$vmuser" -Confirm:$false
    }
    ElseIf ($_.office -eq "REDACTED") {
    New-vm -Name "$vm$vmuser" -VMhost $ViewHost -Template $GoldenImage -Datastore $Datastore  -OSCustomizationspec $Customspec -Location "REDACTED"
    Get-NetworkAdapter -VM "$vm$vmuser" | Set-NetworkAdapter -NetworkName "pg_0085-REDACTED" -startConnected:$True -Confirm:$false
    Start-VM -VM "$vm$vmuser" -Confirm:$false
    }
    ElseIf ($_.office -eq "REDACTED") {
    $DFLPortgroup = Get-Random -InputObject "pg_REDACTED","pg_0025-REDACTED"
    New-vm -Name "$vm$vmuser" -VMhost $ViewHost -Template $GoldenImage -Datastore $Datastore  -OSCustomizationspec $Customspec -Location "REDACTED"
    Get-NetworkAdapter -VM "$vm$vmuser" | Set-NetworkAdapter -NetworkName $DFLPortgroup -startConnected:$True -Confirm:$false
    Start-VM -VM "$vm$vmuser" -Confirm:$false
    }
    ElseIf ($_.office -eq "REDACTED") {
    New-vm -Name "$vm$vmuser" -VMhost $ViewHost -Template $GoldenImage -Datastore $Datastore  -OSCustomizationspec $Customspec -Location "REDACTED"
    Get-NetworkAdapter -VM "$vm$vmuser" | Set-NetworkAdapter -NetworkName "pg_0035-REDACTED" -startConnected:$True -Confirm:$false
    Start-VM -VM "$vm$vmuser" -Confirm:$false
    }
    ElseIf ($_.office -eq "REDACTED") {
    New-vm -Name "$vm$vmuser" -VMhost $ViewHost -Template $GoldenImage -Datastore $Datastore  -OSCustomizationspec $Customspec -Location "REDACTED"
    Get-NetworkAdapter -VM "$vm$vmuser" | Set-NetworkAdapter -NetworkName "pg_0075-REDACTED-REDACTED" -startConnected:$True -Confirm:$false
    Start-VM -VM "$vm$vmuser" -Confirm:$false
    }
    ElseIf ($_.office -eq "REDACTED") {
    New-vm -Name "$vm$vmuser" -VMhost $ViewHost -Template $GoldenImage -Datastore $Datastore  -OSCustomizationspec $Customspec -Location "REDACTED"
    Get-NetworkAdapter -VM "$vm$vmuser" | Set-NetworkAdapter -NetworkName "pg_0075-REDACTED-REDACTED" -startConnected:$True -Confirm:$false
    Start-VM -VM "$vm$vmuser" -Confirm:$false
    }
    ElseIf ($_.office -eq "REDACTED") {
    New-vm -Name "$vm$vmuser" -VMhost $ViewHost -Template $GoldenImage -Datastore $Datastore  -OSCustomizationspec $Customspec -Location "REDACTED"
    Get-NetworkAdapter -VM "$vm$vmuser" | Set-NetworkAdapter -NetworkName "pg_0085-REDACTED" -startConnected:$True -Confirm:$false
    Start-VM -VM "$vm$vmuser" -Confirm:$false
    }
    ElseIf ($_.office -eq "REDACTED") {
    New-vm -Name "$vm$vmuser" -VMhost $ViewHost -Template $GoldenImage -Datastore $Datastore  -OSCustomizationspec $Customspec -Location "REDACTED"
    Get-NetworkAdapter -VM "$vm$vmuser" | Set-NetworkAdapter -NetworkName "pg_0055-REDACTED" -startConnected:$True -Confirm:$false
    Start-VM -VM "$vm$vmuser" -Confirm:$false
    }
    ElseIf ($_.office -eq "REDACTED") {
    New-vm -Name "$vm$vmuser" -VMhost $ViewHost -Template $GoldenImage -Datastore $Datastore  -OSCustomizationspec $Customspec -Location "REDACTED"
    Get-NetworkAdapter -VM "$vm$vmuser" | Set-NetworkAdapter -NetworkName "pg_0065-REDACTED-REDACTED" -startConnected:$True -Confirm:$false
    Start-VM -VM "$vm$vmuser" -Confirm:$false
    }
    ElseIf ($_.office -eq "REDACTED") {
    New-vm -Name "$vm$vmuser" -VMhost $ViewHost -Template $GoldenImage -Datastore $Datastore  -OSCustomizationspec $Customspec -Location "REDACTED"
    Get-NetworkAdapter -VM "$vm$vmuser" | Set-NetworkAdapter -NetworkName "pg_0085-REDACTED" -startConnected:$True -Confirm:$false
    Start-VM -VM "$vm$vmuser" -Confirm:$false
    }
    ElseIf ($_.office -eq "REDACTED") {
    New-vm -Name "$vm$vmuser" -VMhost $ViewHost -Template $GoldenImage -Datastore $Datastore  -OSCustomizationspec $Customspec -Location "REDACTED"
    Get-NetworkAdapter -VM "$vm$vmuser" | Set-NetworkAdapter -NetworkName "pg_0075-REDACTED-REDACTED" -startConnected:$True -Confirm:$false
    Start-VM -VM "$vm$vmuser" -Confirm:$false
    }
    ElseIf ($_.office -eq "REDACTED") {
    New-vm -Name "$vm$vmuser" -VMhost $ViewHost -Template $GoldenImage -Datastore $Datastore  -OSCustomizationspec $Customspec -Location "REDACTED"
    Get-NetworkAdapter -VM "$vm$vmuser" | Set-NetworkAdapter -NetworkName "pg_0055-REDACTED" -startConnected:$True -Confirm:$false
    Start-VM -VM "$vm$vmuser" -Confirm:$false
    }
    ElseIf ($_.office -eq "REDACTED") {
    New-vm -Name "$vm$vmuser" -VMhost $ViewHost -Template $GoldenImage -Datastore $Datastore  -OSCustomizationspec $Customspec -Location "REDACTED"
    Get-NetworkAdapter -VM "$vm$vmuser" | Set-NetworkAdapter -NetworkName "pg_0065-REDACTED-REDACTED" -startConnected:$True -Confirm:$false
    Start-VM -VM "$vm$vmuser" -Confirm:$false
    }
}

#Email message detailing created VMs.
$message.Body =(Get-VIEvent -maxsamples 10000 |where {$_.Gettype().Name-eq "VmCreatedEvent" -or $_.Gettype().Name-eq "VmBeingClonedEvent" -or $_.Gettype().Name-eq "VmBeingDeployedEvent"} | where {$_.CreatedTime -ge ((Get-Date).AddDays(-0)).Date} | Sort CreatedTime -Descending | select-object createdtime,fullformattedmessage | ConvertTo-HTML -head $style -PreContent "<hr><h4 align=`"center`" bgcolor=`"#0099ff`">VMs Created - Last Day</h4><hr>")
$smtp = New-Object Net.Mail.SmtpClient($smtpServer)
$smtp.Send($message)
exit
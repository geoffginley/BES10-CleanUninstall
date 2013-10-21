
# Variable definitions
# ====================
# $temp			: Script temporary directory
# $msilog		: Temporary path for registry backups
# $reglog		: Temporary path for registry backups
# $logname		: Log File for main script
# $prglist		: List of BlackBerry items in add/remove programs
# $pfroot		: Program Files root
# $instpath		: Install path 
# $cfpath		: Common Files path
# $ver 			: Version of the script

Function Main() {
    
	# Before doing any work check to see if a uninstaller is still a option
	# Check if still installed through add/remove and abort if so
	Write-Host "Please standby while we compile environemntal information. This can take several minutes..."
	Write-Host ''

	# Determine if OS is 32- or 64-bit and set the Program Files path
	switch (Get-WmiObject -class win32_processor | select -ExpandProperty AddressWidth)
	{
		32 { $pfroot = ${env:ProgramFiles} }
		64 { $pfroot = ${env:ProgramFiles(x86)} }
	}

	$ver = "1.0.1"
	$temp = "${env:TEMP}\bb_clean_uninstall"
	$msilog = $temp+'\msilogs'
	$reglog = $temp+'\reglogs'
	$beslog = $temp+'\beslogs'

	# Create a log file and for now append everthing to 1 log, just add date at the top when we do it
	$logname = $temp+'\cleanup.log'

	$prglist = @()
	$prglist = ('BlackBerry Enterprise Server for','BlackBerry Device Service','Universal Device Service','BlackBerry Management Studio','BlackBerry Enterprise Service','BlackBerry Device Communication Components')

	# Determine if we have components still installed through appwiz
	$prgs = @()
	$prgs = Get-ItemProperty Registry::'HKLM\SOFTWARE\Microsoft\Windows\CurrentVersion\Uninstall\*' | Where {$_.DisplayName} | Select-Object DisplayName, DisplayVersion, Publisher | %{$prgs += $_}
	$prgs += Get-ItemProperty Registry::'HKLM\SOFTWARE\Wow6432Node\Microsoft\Windows\CurrentVersion\Uninstall\*' | Where {$_.DisplayName} | Select-Object DisplayName, DisplayVersion, Publisher
	$bbprgs = @()
		foreach ($value in $prglist){
			$bbprg = @()
			$bbprg += $prgs | where {$_.displayname -match ($value)}
			$bbprgs += $bbprg
	    }

	If ($bbprgs -ne $null){
		Write-Host 'You currently have 1 or more BlackBerry products installed, please remove before running this script again'
		$bbprgs | Ft -AutoSize
		Write-Host ''
		Return
	}

	# Disclaimer	
	Disclaimer	

	# Create Temp fodlers
	WorkSpace

	Out-File -FilePath $logname -InputObject "" -Append
	Out-File -FilePath $logname -InputObject "Script version $ver started @ $(get-date)" -Append
	Out-File -FilePath $logname -InputObject "" -Append

	# Back up all the BES installer logs
	BackupBESLogs

	# MSI's
	CleanMSIs

	# Check if OS supports modules to remove sites/bindings
	$os = Get-WmiObject -class Win32_OperatingSystem | select -ExpandProperty buildnumber
	Out-File -FilePath $logname -InputObject "" -Append
	Out-File -FilePath $logname -InputObject $os -Append
	
	If($os -gt 7600){
		# Clean up UDS websites
		RemoveWebSites

		# Remove port to SSL mapping
		CleanSSLPorts
	}
	ElseIf($os -lt 7600 -and $os -ne $null){
		Out-File -FilePath $logname -InputObject "Skipping Sites and Bindings" -Append
		$legosmsg = "
It appears that this operating system doesn't support certain modules required to check/remove websites and bindings

Please check IIS for UDS.CoreModule and UDS.CommunicationModule, if found please remove along with corresponding application pools

Please also verify that there are no left over SSL ports bound to the previous sites, you can run below command for further details

	netsh http show sslcert

If there are ports listed from the above that are not in use by websites other than ours please remove manually, see below as an example

	Example: netsh http delete sslcert ipport=0.0.0.0:9081

"
	}
	Else{
		Write-Host "No OS detected, assuming 2008r1 or earlier"
		Out-File -FilePath $logname -InputObject "No OS detected, assuming 2008r1 or earlier" -Append
	}

	# Files
	CleanInstallPath
	CleanTemp

	# Registry
	CleanRegSoftware

	# Finalize
	Finalize


	return
}

Function Disclaimer() {
Write-Host ''
Write-Host '        ***********************************************************************************'
Write-Host '        *                                                       					*'
Write-Host '        *                       Disclaimer                         				*'
Write-Host '        *                                                       					*'
Write-Host '        *   The purpose of this script is to cleanly remove all BlackBerry Products.    	*'
Write-Host '        *   If you have a BlackBerry applications that you would like to keep   		*'
Write-Host '        *   This script is not for you :)                                   			*'
Write-Host '        *                                                       					*'
Write-Host '        ***********************************************************************************'

}

Function WorkSpace() {
	Write-Host "`n`n[WorkSpace] Creating temporary folders"

	# Create a tmp folders if one doesn't exist
	Foreach ($fldr in ($temp, $beslog, $msilog, $reglog)){
		If(!(Test-path $fldr)){
			Try {
				New-Item $fldr -ItemType directory | Out-Null
				Write-Host "Created $fldr"
				Out-File -FilePath $logname -InputObject "Created $fldr" -Append
			}
			Catch { 
				$_
			}
		}
		Else{ 
			Write-Host "$fldr exists" 
			Out-File -FilePath $logname -InputObject "$fldr exists" -Append
		}
	}
}

Function BackupBESLogs() {
	Write-Host "`n`n[Logs] Scanning for log files for RCA"
	Out-File -FilePath $logname -InputObject "" -Append
	Out-File -FilePath $logname -InputObject "`n`n[Logs] Scanning for log files for RCA" -Append

	# Ask if log files are set to the default path
	$title = "Confirm Log File Path"
	$message = "Was the default log file path used?`n`nEx: $pfroot\Research In Motion\<Product>\logs"

	$yes = New-Object System.Management.Automation.Host.ChoiceDescription "&Yes", `
            "BES log files are in default path"
	$no  = New-Object System.Management.Automation.Host.ChoiceDescription "&No", `
            "BES log files are in non-default path"

	$options = [System.Management.Automation.Host.ChoiceDescription[]]($yes,$no)

	$result = $Host.UI.PromptForChoice($title, $message, $options, 0)

	switch ($result)
	{ # 0 for Yes, 1 for No
	        0 { $logfldr = $pfroot+'\Research In Motion' 
		Out-File -FilePath $logname -InputObject "Default log path was selected during install" -Append
		}
		1 { 
		$browse = New-Object -ComObject shell.Application
		$folder = $browse.BrowseForFolder(0,"Please select a folder",0,17)
		$logfldr = $folder.self.Path
		Out-File -FilePath $logname -InputObject "Customer selected $logfldr for log path" -Append
		}
	}

	# We need to find the installer logs
	$installfldr = @()
	Get-ChildItem "$logfldr" -Recurse -Filter 'installer' | %{$installfldr += $_.fullname}
		If($installfldr -ne $null){
			foreach ($fldr in $installfldr){
				If(Test-Path "$fldr" -filter 'setup*.log'){
					Write-Host "Copying $fldr to workspace"
					Out-File -FilePath $logname -InputObject "Copying $fldr to workspace" -Append
					copy-item $fldr $beslog -Force -Recurse
	            		}
			}
		}
		Else{
			Write-Host "No installer logs found"
			Out-File -FilePath $logname -InputObject "No installer logs found" -Append
		}
}

Function RemoveWebSites(){
	Write-Host "`n`n[IIS] Scanning IIS for UDS sites"
	Out-File -FilePath $logname -InputObject "" -Append
	Out-File -FilePath $logname -InputObject "`n`n[IIS] Scanning IIS for UDS sites" -Append

	Import-module servermanager
	Import-Module webadministration

	# Check if IIS is installed
	$isiis = @()
	Get-WindowsFeature | where {$_.Installed -match "True" -and $_.Name -match "Web-"} | select displayname, name | %{$isiis += $_}

	If($isiis -ne $null){
		Out-File -FilePath $logname -InputObject "Installed IIS Features:" -Append
		Get-WindowsFeature | where {$_.Installed -match "True" -and $_.Name -match "Web-"} | Out-File -FilePath $logname -Append

        	# Log all the websites 
		if(Get-Website | ? {$_ -ne $null}){
			Get-Website | FT -AutoSize | Out-File -FilePath $logname -Append
		}
		Else{
			Out-File -FilePath $logname -InputObject "" -Append 
			Out-File -FilePath $logname -InputObject "No WebSites Found" -Append    
		}

		# If UDS's still has sites/pools remove
		$udssites = @()
		$udssites = ('UDS.CommunicationModule', 'UDS.CoreModule')

		If(get-website | ?{$_.name -eq $udssites[0] -or $_.name -eq $udssites[1]}){
			Foreach ($usite in $udssites){ 
				If(Test-Path IIS:\sites\$usite){
	                		If(get-website $usite | ?{$_.state -eq 'Started'}){
		                   		Write-Host "Stopping $usite Site"
						Out-File -FilePath $logname -InputObject "Stopping $usite Site" -Append 
		                    		Stop-Website $usite 
						Write-Host "Removing $usite Site"
						Out-File -FilePath $logname -InputObject "Removing $usite Site" -Append 
		                    		Remove-Website $usite 
	                		}
			                Else{
						Write-Host "Removing $usite Site"
						Out-File -FilePath $logname -InputObject "Stopping $usite Site" -Append 
			                    	Remove-Website $usite 
			                }
				}
			}
		}
		Else{
			Write-Host "No WebSites Found"
			Out-File -FilePath $logname -InputObject "No WebSites Found" -Append 
		}

			If(Get-Item iis:\apppools\* | ?{$_.name -eq $udssites[0] -or $_.name -eq $udssites[1]}){
				Foreach ($usite in $udssites){ 
					If(Get-Item iis:"\apppools\$usite"){
	                    			If(Get-Item iis:\apppools\$usite | ?{$_.state -eq 'Started'}){
	                        			Write-Host "Stopping $usite Pool"
							Out-File -FilePath $logname -InputObject "Stopping $usite Pool" -Append 
	                        			Stop-WebAppPool $usite 
							Write-Host "Removing $usite Pool"
							Out-File -FilePath $logname -InputObject "Removing $usite Pool" -Append 
	                        			Remove-WebAppPool $usite 
	                    				}		
	                    			Else{
	                        			Write-Host "Removing $usite Pool"
							Out-File -FilePath $logname -InputObject "Removing $usite Pool" -Append 
	                        			Remove-WebAppPool $usite 
	                    			}
	                		}
				}
			}
			Else{
				Write-Host "No Application Pools Found"
				Out-File -FilePath $logname -InputObject "No Application Pools Found" -Append 
			}
	}
	Else{
		Write-Host "IIS not found"
		Out-File -FilePath $logname -InputObject "IIS not found" -Append 
	}
}

Function CleanSSLPorts(){
	Out-File -FilePath $logname -InputObject "" -Append
	Out-File -FilePath $logname -InputObject "`n`n[Ports] Scanning Ports for SSL Bindings" -Append

	Import-Module webadministration
	
	$netshow = @()
	$netshow = netsh http show sslcert | ?{$_.contains('IP:port')}
	Out-File -FilePath $logname -InputObject "Current SSL bindings on " -Append
	$netshow | Out-File $logname -Append
	
	If($netshow -ne $null){

		# build list of ssl port bindings
		$sslbindings = @()
		[int[]]$sslbindings = netsh http show sslcert | ?{ $_ -match 'Ip:port'} | %{$_.split(':')[3]}

		# build list of existing bindings 
		$sitebindings = @()
		get-childItem IIS:\Sites\* | ?{$_.bindings.collection.protocol -eq 'https'}  | select -ExpandProperty Bindings | select -ExpandProperty collection | %{[int[]]$sitebindings += $_.bindingInformation.split(':')[1]}

		# Remove the site binding from the ssl port bindings
		$portlist = @()
		[int[]]$portlist = Compare-Object $sitebindings $sslbindings -PassThru

		# if ssl port bindings aren't in original binding list del...but prompt first
		If($portlist -ne $null){
		
			$title = "Delete SSL port Bindings"
			$message = "Would like to remove the SSL port mappings from the port below
				
						$portlist
							
						"

			$yes = New-Object System.Management.Automation.Host.ChoiceDescription "&Yes", `
					"remove the SSL port mappings for $portlist ports"
			$no  = New-Object System.Management.Automation.Host.ChoiceDescription "&No", `
					"leave $portlist SSL port mappings"

			$options = [System.Management.Automation.Host.ChoiceDescription[]]($yes,$no)

			$result = $Host.UI.PromptForChoice($title, $message, $options, 0)

			switch ($result)
			{ # 0 for Yes, 1 for No
				0 {
					# if ssl port bindings aren't in original binding list del
					Out-File -FilePath $logname -InputObject "" -Append
					Out-File -FilePath $logname -InputObject "Removing $portlist SSL port mappings" -Append
					foreach ($port in $portlist){
						netsh http delete sslcert ipport=0.0.0.0:$port	#left show for testing
					}
					Out-File -FilePath $logname -InputObject "" -Append
					Out-File -FilePath $logname -InputObject "Post-Clean SSL bindings" -Append
					netsh http show sslcert | Out-File $logname -Append
				}
				1 {
					write-host "You have chosen not remove the SSL port bindings"
					Out-File -FilePath $logname -InputObject "Customer chose not to remove the SSL port bindings" -Append
				}
			}
		}
		Else{
			Write-Host "Unable to find SSL port bound to sites that no longer exist"
			Out-File -FilePath $logname -InputObject "" -Append
			Out-File -FilePath $logname -InputObject "Unable to find SSL ports bound to sites that no longer exist" -Append
		}
	}
	Else{
		Write-Host "No SSL Port bindings exist"
		Out-File -FilePath $logname -InputObject "" -Append
		Out-File -FilePath $logname -InputObject "No SSL Port bindings exist" -Append	
	}
}

Function CleanMSIs() {
	Write-Host "`n`n[MSI's] Scanning Installer Files"
	Out-File -FilePath $logname -InputObject "" -Append
	Out-File -FilePath $logname -InputObject "`n`n[MSI's] Scanning Installer Files" -Append

	# Build file list to work with
	$filelist = @()
	Get-WmiObject -class Win32_Product | ?{$_.Vendor.Contains('Research In Motion')} | select packagecache | %{$filelist += $_.packagecache}
    
	# Log all installed apps
	$filelist | Out-File -FilePath $logname -Append

	If($filelist -ne $null){
	        #Work through list of files
	        foreach ($line in $filelist){
			$file = $line
 			$filename = Split-Path $line -leaf
			Write-host "Attemptimg to cleanly remove $filename"

			# Clean Uninstall - to remove ability to cancel add -! 
			Try{
				Start-Process -FilePath msiexec -ArgumentList /x, $file, PF_RIM=1, /qb, /l*v, $msilog\$filename+'.log' -Wait
			}
			catch {
				$_
			}
		}
	}
	Else{
		Write-host "No MSI's to uninstall"
		Out-File -FilePath $logname -InputObject "No MSI's to uninstall" -Append
	}
}

Function CleanInstallPath() {
	Write-Host "`n`n[Files] Checking if installation path exists"
	Out-File -FilePath $logname -InputObject "" -Append
	Out-File -FilePath $logname -InputObject "`n`n[Files] Checking if installation path exists" -Append

	$cfpath = $pfroot+'\Common Files\Research In Motion'

	# Ask if BES was installed to the default path

	$title = "Confirm Installation Path"
	$message = "Was the default installation path used?`n`nEx: $pfroot\Research In Motion"

	$yes = New-Object System.Management.Automation.Host.ChoiceDescription "&Yes", `
			"BES installed to default path"
	$no  = New-Object System.Management.Automation.Host.ChoiceDescription "&No", `
			"BES installed to non-default path"

	$options = [System.Management.Automation.Host.ChoiceDescription[]]($yes,$no)

	$result = $Host.UI.PromptForChoice($title, $message, $options, 0)

	switch ($result)
	{ # 0 for Yes, 1 for No
		0 { 
			$instpath = $pfroot+'\Research In Motion' 
			Out-File -FilePath $logname -InputObject "Default install path was selected during install" -Append
		}
		1 { 
			$browse = New-Object -ComObject shell.Application
			$folder = $browse.BrowseForFolder(0,"Please select a folder",0,17)
			$instpath = $folder.self.Path
			Out-File -FilePath $logname -InputObject "Customer selected $instpath for install path" -Append
		}
	}

	If (Test-Path $instpath){
		# if a non default root was specified its possible that it could contains other products
		# build a list to dbl check against
		$fldrvar1 = ('BlackBerry','Universal Device Service')
		If ($nondefaultpath -eq 'n'){
			Remove-Item $instpath -Recurse -Force 
			write-host "$instpath - has been removed "
			Out-File -FilePath $logname -InputObject "$instpath - was been removed " -Append
		}
		Else{
			$tfldrs = @()
			Get-ChildItem $instpath | %{$tfldrs += $_.FullName}
			If ($tfldrs -ne $null){
				foreach($path in $tfldrs){
					Remove-Item $path -Recurse -Force 
					Write-host $path ' - has been removed ' 
					Out-File -FilePath $logname -InputObject "$path - was been removed " -Append
				}
			}
			Else{
				Write-Host "Install path doesn't exist"
				Out-File -FilePath $logname -InputObject "Install path doesn't exist" -Append
			}
		}
	}

	Write-Host "`n`n[Files] Scanning for Common Files"
	Out-File -FilePath $logname -InputObject "" -Append
	Out-File -FilePath $logname -InputObject "`n`n[Files] Scanning for Common Files" -Append
	If (Test-Path $cfpath){
		Remove-Item $cfpath -Recurse -Force 
		Write-Host $cfpath ' - has been removed ' 
		Out-File -FilePath $logname -InputObject "$cfpath - was been removed " -Append
	}
	Else{
		Write-Host "Common Files path doesn't exist"
		Out-File -FilePath $logname -InputObject "Common Files path doesn't exist" -Append
	}
}

Function CleanTemp() {
	Write-Host "`n`n[Files] Clearing contents of temp folder"
	Out-File -FilePath $logname -InputObject "" -Append
	Out-File -FilePath $logname -InputObject "`n`n[Files] Clearing contents of temp folder" -Append

    	If (Test-Path -Path ${env:TEMP}'\*'){
        	Write-Host 'Purging temp folder'
			Out-File -FilePath $logname -InputObject "Purging temp folder" -Append		
        	remove-item "${env:TEMP}\*" -recurse -force -Exclude "bb_clean_uninstall" 
        	Write-Host 'Temp folder purged'
			Out-File -FilePath $logname -InputObject "Temp folder purged" -Append	
    	}
		Else{
			Write-Host 'Temp folder appears empty'
			Out-File -FilePath $logname -InputObject "Temp folder appears empty" -Append
		}
}

Function CleanRegSoftware() {
    Write-Host "`n`n[Registry] Scanning software keys"
	Out-File -FilePath $logname -InputObject "" -Append
	Out-File -FilePath $logname -InputObject "`n`n[Registry] Scanning software keys" -Append

	# Make list of registry keys to scan between HKLM and HKUsers
	$key_list = ("HKLM\Software\Wow6432node", "HKLM\Software") +
				 (Get-ChildItem registry::HKEY_USERS -ErrorAction SilentlyContinue | %{ $_.Name+'\Software'})

	# Refine to only keys that exist
	$key_list = $key_list | ?{Test-Path -Path Registry::"$_\Research In Motion"}

	# If nothing in the list, exit function
	If($key_list -ne $null){
		ForEach($reg_key in $key_list) {
			#Iterate through the list
			$logname = "$reglog\" + ($reg_key.Split('\') -join '_') + '_RIM.reg'
			$reg_key += '\Research In Motion'
			If(!(Test-Path "$logname")) {
				Reg export $reg_key "$logname" | Out-Null
			}
			If(Test-Path "$logname"){
				Try {
					Remove-Item registry::$reg_key -Recurse -Force -ea SilentlyContinue 
					Write-Host $reg_key ' - has been backed up and removed'
					Out-File -FilePath $logname -InputObject "$reg_key - has been backed up and removed" -Append
				}
				Catch {
				$_
				}
			}
		}
	}
	Else{
		Write-Host "HKLM keys don't exist"
		Out-File -FilePath $logname -InputObject "HKLM keys don't exist" -Append
	}
}

Function Finalize() {
	Write-Host "`n`n[Final] Please wait while the script finalizes..."
	Out-File -FilePath $logname -InputObject "" -Append
	Out-File -FilePath $logname -InputObject "`n`n[Final] Finalizing the script" -Append
	Out-File -FilePath $logname -InputObject "Finalizing $(get-date)" -Append
	
	Write-Host ""
	Write-Host $legosmsg
	
	# Dump WMI objects to log post uninstalls, to compare to pre-uninstalls
	$wmiobjs = @()
	Out-File -FilePath $logname -InputObject "" -Append
	Out-File -FilePath $logname -InputObject "Dumping Remaining installed apps if they exist" -Append

	Get-WmiObject -class Win32_Product | ?{$_.Vendor.Contains('Research In Motion')} | Out-File -FilePath $logname -Append
	Write-Host ""
	Write-Host "Please collect the logs from $temp and submit for review"
}

Main

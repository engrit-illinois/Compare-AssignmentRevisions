# Documentation home: https://github.com/engrit-illinois/Compare-AssignmentRevisions
# By mseng3@illinois.edu
function Compare-AssignmentRevisions {

	param(
		
		[Parameter(Position=0,Mandatory=$true,ParameterSetName="Array")]
		[string[]]$Computers,
		
		[Parameter(Position=0,Mandatory=$true,ParameterSetName="Collection")]
		[string]$Collection,
		
		[string]$SiteCode="MP0",
		
		[string]$Provider="sccmcas.ad.uillinois.edu",
		
		[string]$CMPSModulePath="$($ENV:SMS_ADMIN_UI_PATH)\..\ConfigurationManager.psd1",
		
		[switch]$DisableCaching,
		
		[switch]$DisableCIMFallbacks,
		
		[switch]$DisableCIM,
		
		[int]$CIMTimeoutSec=10,
		
		[switch]$DisableIC,
		
		[int]$ICTimeoutSec=10,
		
		[switch]$DisableWMI,
		
		[int]$WMITimeoutSec=10,
		
		[switch]$NoLog,
		
		[switch]$ComputerInfoOnly,
		
		[string]$LogPath="c:\engrit\logs\Compare-AssignmentRevisions_$(Get-Date -Format `"yyyy-MM-dd_HH-mm-ss-ffff`").log",
		
		[int]$Verbosity=0
	)
	
	$CSVPATH = $LogPath -replace "\.log",".csv"
	
	$script:CachedAppDeployments = @()
	$script:CachedTSDeployments = @()
	$script:CachedApplications = @()
	$script:CachedTaskSequences = @()
	$script:CachedCollections = @()
	
	function log {
		param (
			[string]$msg,
			[int]$l=0, # level (of indentation)
			[int]$v=0, # verbosity level
			[switch]$nots, # omit timestamp
			[switch]$nnl # No newline after output
		)
		
		# Indentation level guide
		# computers: 0
		# computer: 1
		# applications/assignments: 2
		# application/assignment: 3
		# sub-application/sub-assignment: 4
		
		if(!(Test-Path -PathType leaf -Path $LogPath)) {
			$shutup = New-Item -ItemType File -Force -Path $LogPath
		}
		
		for($i = 0; $i -lt $l; $i += 1) {
			$msg = "    $msg"
		}
		if(!$nots) {
			$ts = Get-Date -Format "yyyy-MM-dd HH:mm:ss:ffff"
			$msg = "[$ts] $msg"
		}
		
		if($v -le $Verbosity) {
			if($nnl) {
				Write-Host $msg -NoNewline
			}
			else {
				Write-Host $msg
			}
			
			if(!$NoLog) {
				if($nnl) {
					$msg | Out-File $LogPath -Append -NoNewline
				}
				else {
					$msg | Out-File $LogPath -Append
				}
			}
		}
	}
	
	function Log-Error {
		param(
			[System.Management.Automation.ErrorRecord]$e,
			[int]$v=0
		)
		
		if($v -le $Verbosity) {
			log "$($e.Exception.Message)" -l 4
			log "$($e.Exception.GetType().fullname)" -l 5
			log "$($e.InvocationInfo.PositionMessage.Split("`n")[0])" -l 5
		}
	}

	function Prep-SCCM {
		log "Preparing connection to SCCM..."
		$initParams = @{}
		if((Get-Module ConfigurationManager) -eq $null) {
			# The ConfigurationManager Powershell module switched filepaths at some point around CB 18##
			# So you may need to modify this to match your local environment
			Import-Module $CMPSModulePath @initParams -Scope Global
		}
		if((Get-PSDrive -Name $SiteCode -PSProvider CMSite -ErrorAction SilentlyContinue) -eq $null) {
			New-PSDrive -Name $SiteCode -PSProvider CMSite -Root $Provider @initParams
		}
		Set-Location "$($SiteCode):\" @initParams
		log "Done prepping connection to SCCM." -v 2
	}
	
	# Make array of objects representing computers
	function Get-CompObjects($comps) {
		log "Making array of objects to represent each computer..."
	
		# Make sure $comp is treated as an array, even if it has only one member
		# Not sure if this is necessary, but better safe than sorry
		$comps = @($comps)
		
		# Make new array to hold objects representing computers
		$compObjects = @()
		
		foreach($thisComp in $comps) {
			$thisCompHash = @{
				"name" = $thisComp
				"SCCMClientVersion" = $null
				"PSVersion" = $null
				"OSVersion" = $null
				"Make" = $null
				"Model" = $null
				"localassignments" = @()
				"localapplications" = @()
				"skip" = $false
			}
			$thisCompObject = New-Object PSObject -Property $thisCompHash
			$compObjects += @($thisCompObject)
		}
		
		log "Done making computer object array." -v 2
		$compObjects
	}
	
	function Get-CompData($comps) {
		log " " -nots
		log "Getting data for all computers..."
		$num = 1
		foreach($comp in $comps) {
			$thisCompName = $comp.name
			$count = @($comps).count
			$completion = ([math]::Round($num / $count, 2)) * 100
			log "Getting data for computer $num/$count ($completion%): `"$thisCompName`"..." -l 1
			$num += 1
			
			$comp = Get-Data "SCCMClientVersion" $comp
			
			# We could check for $comp.skip after each query, but
			# if we get even one result (SCCMClientVersion), then we should include it in the output.
			# Also if getting SCCMClientVersion works, then it's unlikely the rest will fail.
			if($comp.skip) {
				log "Computer unresponsive, or access denied. Skipping further queries." -l 2
			}
			else {
				$comp = Get-Data "PSVersion" $comp
				$comp = Get-Data "OSVersion" $comp
				$comp = Get-Data "Model" $comp
				
				if(!$ComputerInfoOnly) {
					# Never ended up doing anything with this localapplication info, so skip it to save on time and memory
					#$comp = Get-Data "LocalApplications" $comp
				}
					
				$comp = Get-Data "LocalAssignments" $comp
				
				Export-CompAssignments $comp
			}
			
			# Once we've exported the data (if necessary), remove it so it doesn't eventually eat a bunch of memory when scanning many computers
			log "Removing computer data to save memory." -l 2
			$comp = $null
			
			log " " -nots -v 1
			log "Done getting data for computer: `"$thisCompName`"." -l 1
			
			log " " -nots
			log " " -nots -v 1
			log " " -nots -v 1
		}
		log "Done getting data for all computers."
		
		$comps
	}
	
	# In some cases this still hangs, so running AsJob, so we can time Invoke-Command and WMI out
	# https://social.technet.microsoft.com/Forums/en-US/f16cb07c-10ac-458e-ae92-077f9247b436/assigning-a-timeout-to-invokecommand?forum=winserverpowershell
	function Query-AsJob {
		param(
			[string]$type,
			[string]$compName,
			[string]$scriptBlockString
		)
		
		log "Trying $type query..." -l 3 -v 1
		
		$scriptBlock = [scriptblock]::Create($scriptBlockString)
		if($type -eq "Invoke-Command") {
			$scriptBlock = {
				$scriptBlock = [scriptblock]::Create($using:scriptBlockString)
				Invoke-Command -ComputerName $using:compName -ErrorAction "Stop" -ScriptBlock $scriptBlock
			}
		}
		
		$job = Start-Job -ErrorAction "Stop" -ScriptBlock $scriptBlock
		
		# Alternate possibility for Invoke-Command
		#$job = Invoke-Command -ComputerName $compName -AsJob -JobName "ic" -ErrorAction "Stop" -ScriptBlock $scriptBlock
		
		# The point of using a job is not for asynchronicity, but merely to implement a timeout for Invoke-Command and WMI, which don't have the equivalent of Get-CIMInstance's -OperationTimeoutSec
		$count = 0
		while(
			($job.State -eq "Running") -and
			($count -lt $ICTimeoutSec)
		) {
			Start-Sleep -Seconds 1
			$count += 1
		}
		if($job.State -eq "Running") {
			$job | Stop-Job
			log "$type timed out." -l 3 -v 1
		}
		
		# Catching errors in received jobs:
		# https://jrich523.wordpress.com/2012/08/08/powershell-catching-terminating-and-non-terminating-errors-in-ps-jobs-job-pattern/
		
		$skip = $false
		try {
			$jobResult = $job | Receive-Job -ErrorAction Stop
		}
		catch {
			log "$type didn't work." -l 3 -v 1
			Log-Error $_ -v 2
			if($_.Exception.Message -like "*Access is denied.*") {
				log "Access denied! will skip further attempts." -l 3
				$skip = $true
			}
		}
		
		$job | Remove-Job
	
		log "$type query job returned:" -l 3 -v 3
		log "----------------------------" -l 4 -v 3
		log $jobResult -l 4 -v 3
		log ($jobResult | ConvertTo-Json) -l 4 -v 3
		log "----------------------------" -l 4 -v 3
		
		$result = [PSCustomObject]@{
			"result" = $jobResult
			"skip" = $skip
		}
		
		$result
	}
	
	function Get-Data($dataType, $comp) {
		$compName = $comp.name
		log " " -nots -v 1
		log "Getting $dataType for computer: `"$compName`"..." -l 2
		
		if(Test-Connection $compName -Quiet -Count 1) {
			log "Computer `"$compName`" responded." -l 3 -v 2
			
			# TODO: skip further queries if the computer doesn't respond
			# TODO: Stop sending quotes around $CIMTimeoutSec
			
			# Define variables and commands used for each dataType/queryType combo
			switch($dataType) {
				"SCCMClientVersion" {
					$namespace = "root\ccm"
					$class = "SMS_Client"
					$scriptBlockStringCIM = "Get-CIMInstance -ComputerName `"$compName`" -Namespace `"$namespace`" -Class `"$class`" -ErrorAction `"Stop`" -OperationTimeoutSec $CIMTimeoutSec"
					$scriptBlockStringIC = "Get-WMIObject -Namespace `"$namespace`" -Class `"$class`" -ErrorAction `"Stop`""
					$scriptBlockStringWMI = "Get-WMIObject -ComputerName `"$compName`" -Namespace `"$namespace`" -Class `"$class`" -ErrorAction `"Stop`""
					break
				}
				"PSVersion" {
					$scriptBlockStringIC = "`$PSVersionTable.PSVersion"
					# Fall back to some weird WMI
					# Couldn't find an easier way to do this via WMI
					# Only other way I could find would be reading powershell.exe file properties, but that seems to use a different versioning system these days
					$hklm = 2147483650 # This is some sort of unique identifier that always means "HKEY_LOCAL_MACHINE", I guess
					$key = "SOFTWARE\Microsoft\PowerShell\3\PowerShellEngine"
					$value = "PowerShellVersion"
					#$wmi = [wmiclass]"\\$compName\root\default:stdRegProv"
					#$result = ($wmi.GetStringValue($hklm,$key,$value)).svalue
					
					$scriptBlockStringWMI = "(([wmiclass]`"\\$compName\root\default:stdRegProv`").GetStringValue(`"$hklm`",`"$key`",`"$value`")).svalue"
					break
				}
				"OSVersion" {
					$class = "Win32_OperatingSystem"
					$scriptBlockStringCIM = "(Get-CIMInstance -ComputerName `"$compName`" -Class `"$class`" -ErrorAction `"Stop`" -OperationTimeoutSec $CIMTimeoutSec).Version"
					$scriptBlockStringIC = "[Environment]::OSVersion.Version"
					$scriptBlockStringWMI = "Get-WMIObject -ComputerName `"$compName`" -Class `"$class`" -ErrorAction `"Stop`""
					break
				}
				"Model" {
					$class = "Win32_ComputerSystem"
					$scriptBlockStringCIM = "Get-CIMInstance -ComputerName `"$compName`" -Class `"$class`" -ErrorAction `"Stop`" -OperationTimeoutSec $CIMTimeoutSec"
					$scriptBlockStringIC = "Get-WMIObject -Class `"$class`" -ErrorAction `"Stop`""
					$scriptBlockStringWMI = "Get-WMIObject -ComputerName `"$compName`" -Class `"$class`" -ErrorAction `"Stop`""
					break
				}
				"LocalApplications" {
					$namespace = "root\ccm\clientsdk"
					$query = "select * from CCM_Application"
					$scriptBlockStringCIM = "Get-CIMInstance -ComputerName `"$compName`" -Namespace `"$namespace`" -Query `"$query`" -ErrorAction `"Stop`" -OperationTimeoutSec $CIMTimeoutSec"
					$scriptBlockStringIC = "Get-WMIObject -Namespace `"$namespace`" -Query `"$query`" -ErrorAction `"Stop`""
					$scriptBlockStringWMI = "Get-WMIObject -ComputerName `"$compName`" -Namespace `"$namespace`" -Query `"$query`" -ErrorAction `"Stop`""
					break
				}
				"LocalAssignments" {
					$namespace = "root\ccm\policy\Machine"
					$query = "select * from CCM_ApplicationCIAssignment"
					$scriptBlockStringCIM = "Get-CIMInstance -ComputerName `"$compName`" -Namespace `"$namespace`" -Query `"$query`" -ErrorAction `"Stop`" -OperationTimeoutSec $CIMTimeoutSec"
					$scriptBlockStringIC = "Get-WMIObject -Namespace `"$namespace`" -Query `"$query`" -ErrorAction `"Stop`""
					$scriptBlockStringWMI = "Get-WMIObject -ComputerName `"$compName`" -Namespace `"$namespace`" -Query `"$query`" -ErrorAction `"Stop`""
					break
				
				}
				default {
					break
				}
			}
			
			if(
				($dataType -eq "LocalAssignments") -and
				($ComputerInfoOnly)
			) {
				log "-ComputerInfoOnly was specified. Skipping gathering assignment data." -l 3
				$dummyAssignment = [PSCustomObject]@{
					"AssignmentID" = "Dummy assignment"
				}
				$result = @($dummyAssignment)
			}
			else {
				# Try getting data via Get-CIMInstance
				# https://stackoverflow.com/questions/21558559/powershell-wmiobject-exception-handling
				if(!$DisableCIM) {
					$lastMethod = "CIM"
					$jobResult = Query-AsJob -type "CIM" -compName $compName -scriptBlockString $scriptBlockStringCIM
					$result = $jobResult.result
					$comp.skip = $jobResult.skip
				}
				# If CIM doesn't work try Invoke-Command
				# In some cases (Win7 + PSv2), CIM and remote WMI were not working, but Invoke-Command did for some reason
				if(
					(!$result) -and
					(!$DisableCIMFallbacks) -and
					(!$DisableIC) -and
					(!$comp.skip)
				) {
					$lastMethod = "Invoke-Command"
					$jobResult = Query-AsJob -type "Invoke-Command" -compName $compname -scriptBlockString $scriptBlockStringIC
					$result = $jobResult.result
					$comp.skip = $jobResult.skip
				}
		
				# If Invoke-Command doesn't work try WMI
				if(
					(!$result) -and
					(!$DisableCIMFallbacks) -and
					(!$DisableWMI) -and
					(!$comp.skip)
				) {
					$lastMethod = "WMI"
					$jobResult = Query-AsJob -type "WMI" -compName $compname -scriptBlockString $scriptBlockStringWMI
					$result = $jobResult.result
					$comp.skip = $jobResult.skip
				}
			}
	
			# If nothing worked
			if(!$result) {
				log "All methods failed or were skipped. I give up." -l 3 -v 1
			}
			# Else grab final result
			else {
				
				switch($dataType) {
					"SCCMClientVersion" {
						$target = $result.ClientVersion
						$comp.SCCMClientVersion = $target
						break
					}
					"PSVersion" {
						$target = $result
						$comp.PSVersion = $target
						break
					}
					"OSVersion" {
						if($lastMethod -eq "Invoke-Command") {
							$target = "$($result.Major).$($result.Minor).$($result.Build).$($result.Revision)"
						}
						elseif($lastMethod -eq "WMI") {
							$target = $result.Version
						}
						else {
							$target = $result
						}
						$comp.OSVersion = $target
						break
					}
					"Model" {
						$target = $result
						$comp.Make = $target.Manufacturer
						$comp.Model = $target.Model
						$target = "$($target.Manufacturer) $($target.Model)"
						break
					}
					"LocalApplications" {
						$target = $result
						$comp.localapplications = $target
						$target = @($target).count
						break
					}
					"LocalAssignments" {
						$target = $result
						$comp.localassignments = $target
						$target = @($target).count
						break
					}
					default {
						break
					}
				}
			}
	
			if($target) {
				
				switch($dataType) {
					"LocalApplications" {
						log "Retrieved $($comp.localapplications.count) local applications." -l 3
						
						#$comp = Parse-Applications $comp
						break
					}
					"LocalAssignments" {
						log "Retrieved $($comp.localassignments.count) local assignments." -l 3
						
						$comp = Parse-Assignments $comp
						break
					}
					default {
						log "$dataType is `"$target`"." -l 3
						break
					}
				}
			}
			else {
				log "$dataType not retrieved from computer: `"$compName`"!" -l 3
			}
		}
		else {
			log "Computer `"$compName`" did not respond!" -l 3
			$comp.skip = $true
		}
		log "Done getting $dataType for computer: `"$compName`"..." -l 2 -v 2
		
		$comp
	}
	
	function Parse-Applications($comp) {
		<#
		log "Parsing each local application..." -l 3
		foreach($application in $comp.localapplications) {
			log " " -nots
			log "Processing application: `"$($application.Name)`" ..." -l 4
			
			# Not sure what I want to do with this yet
			
			log "Done processing application: `"$($application.Name)`" ..." -l 4 -v 2
		}
		log "Done parsing each local application." -l 3 -v 2
		
		$comp
		#>
	}
	
	function Parse-Assignments($comp) {
		log "Parsing and getting associated deployment and application data for each assignment..." -l 3
		
		$num = 1
		foreach($assignment in $comp.localassignments) {
			$count = @($comp.localassignments).count
			$completion = ([math]::Round($num / $count, 2)) * 100
			log " " -nots -v 1
			log "Processing assignment $num/$count ($completion%) with ID: `"$($assignment.AssignmentID)`" on computer `"$($comp.name)`"..." -l 4 -v 1
			$num += 1
			
			# Add computer-specific stuff to the assignment so it can all be exported with the assignments later
			$assignment | Add-Member -NotePropertyName "_Computer" -NotePropertyValue $comp.name
			$assignment | Add-Member -NotePropertyName "_SCCMClientVersion" -NotePropertyValue $comp.SCCMClientVersion
			$assignment | Add-Member -NotePropertyName "_PSVersion" -NotePropertyValue $comp.PSVersion
			$assignment | Add-Member -NotePropertyName "_OSVersion" -NotePropertyValue $comp.OSVersion
			$assignment | Add-Member -NotePropertyName "_Make" -NotePropertyValue $comp.Make
			$assignment | Add-Member -NotePropertyName "_Model" -NotePropertyValue $comp.Model
			
			if($assignment.AssignmentID -ne "Dummy assignment") {
				$assignment = Parse-Assignment $assignment
			}
			
			log "Done processing assignment with ID: `"$($assignment.AssignmentID)`"..." -l 4 -v 2
		}
		log " " -nots -v 1
		log "Done parsing and getting associated deployment and application data for each assignment." -l 3 -v 1
		
		$comp
	}
	
	# This is actually used for both Parse-Assignment and Parse-AssignmentAppDeployment
	function Parse-DesiredConfigType($assignment) {
		# Only valid for apps, not TSes
		if($assignment._DepType -eq "ts") {
			$result = "TS"
		}
		else {
			# DesiredConfigType is whether the app is deployed to "Install" (1), or Uninstall (2)
			$configTypeNum = $assignment.DesiredConfigType
			
			# This is also encoded into the AssignmentName in the format "<app name>_<deployment collection name>_<DesiredConfigType string>"
			$nameParts = $assignment.AssignmentName.Split("_")
			# Take the last member of this array (ass opposed to the 3rd member), in case the app or collection name contain a "_"
			$configTypeNameString = $nameParts[($nameParts.count - 1)]
			
			# Check that the two match
			switch($configTypeNum) {
				1 { $configTypeNumString = "Install" }
				2 { $configTypeNumString = "Uninstall" }
				Default { $configTypeNumString = "Invalid" }
			}
			$result = $configTypeNumString
			if($configTypeNumString -ne $configTypeNameString) {
				log "Assignment/App deployment disagrees with itself on its DesiredConfigType! DesiredConfigType is `"$configTypeNum`", while AssignmentName contains `"$configTypeNameString`"." -l 6
				$result = "INVALID!"
			}
		}
		
		$assignment | Add-Member -NotePropertyName "_DesiredConfigType" -NotePropertyValue $result
		
		$assignment
	}
	
	function Parse-AssignmentCIs($assignment) {
		if($assignment.AssignedCIs) {
			# CI data is in XML format, translate into object and save
			# https://stackoverflow.com/questions/3935395/loading-xml-string-with-powershell
			$xmlString = @($assignment.AssignedCIs)[0]
			$xmlObject = New-Object -TypeName System.Xml.XmlDocument
			$xmlObject.LoadXml($xmlString)
			$assignment | Add-Member -NotePropertyName "_CI" -NotePropertyValue $xmlObject.CI
			
			# CI version (a.k.a. revision) is stored in its own CIVersion field and also as part of the CI ID string (formatted like "<ModelName>/<CIVersion>")
			# Grab the version out of the CI ID string to make sure it's the same as what's in the CIVersion field
			$ciidParts = $assignment._CI.ID.Split("/")
			$assignment | Add-Member -NotePropertyName "_Revision" -NotePropertyValue $ciidParts[2]
			# Do the same for the ModelName
			$modelName = $ciidParts[0] + "/" + $ciidParts[1]
			$assignment | Add-Member -NotePropertyName "_ModelName" -NotePropertyValue $modelName
		}
		else {
			log "Assignment is missing AssignedCIs field!" -l 6
		}
		
		$assignment
	}
	
	function Parse-AssignmentDeploymentType($assignment) {
		# Assignments with AssignmentID's of the format "DEP-MP######-<ModelName>" refer to apps in deployed task sequences
		$asNameID = $assignment.AssignmentID
		if($assignment.AssignmentID.StartsWith("DEP-MP")) {
			$type = "ts"
			log "Assignment is for an app in a task sequence." -l 5 -v 2
			log "Assignment name: TS assignments have blank assignment names." -l 5 -v 2
			# If this is for a TS app, extract the TS deployment ID
			# In this case, the assignment ID is of the format "DEP-<ts deployment id>-<ModelName>"
			# The TS deployment id is of the format "MP######"
			# In a TaskSequenceDeployment object, this is called the "AdvertisementID".
			$asIDParts = $assignment.AssignmentID.Split("-")
			$tsDepID = $asIDParts[1]
			$assignment | Add-Member -NotePropertyName "_TSDepID" -NotePropertyValue $tsDepID
			$name = "No assignment name. This is a TS assignment."
		}
		else {
			$type = "app"
			log "Assignment is for a directly deployed app." -l 5 -v 2
			log "Assignment name: `"$($assignment.AssignmentName)`"." -l 5 -v 2
			$name = $assignment.AssignmentName
		}
		$assignment | Add-Member -NotePropertyName "_DepType" -NotePropertyValue $type
		$assignment | Add-Member -NotePropertyName "_Name" -NotePropertyValue $name
		
		$assignment
	}
	
	function Parse-Assignment($assignment) {
		log "Parsing assignment..." -l 5 -v 1
		
		$assignment = Parse-AssignmentDeploymentType $assignment
		$assignment = Parse-AssignmentCIs $assignment
		$assignment = Parse-DesiredConfigType $assignment
		$assignment = Get-AssignmentDeployment $assignment
		$assignment = Get-AssignmentApplication $assignment
		
		$assignment = Compare-Revisions $assignment
		$assignment = Compare-ModelNames $assignment
		$assignment = Compare-DesiredConfigTypes $assignment
		
		log "Done parsing assignment." -l 5 -v 2
		$assignment
	}
	
	function Get-AssignmentDeployment($assignment) {
		if($assignment._DepType -eq "ts") {
			$deployment = Get-AssignmentTSDeployment $assignment
		}
		else {
			$deployment = Get-AssignmentAppDeployment $assignment
		}
		$assignment | Add-Member -NotePropertyName "_Deployment" -NotePropertyValue $deployment
		$assignment
	}
	
	function Get-AssignmentTSDeployment($assignment) {
		log "Getting TS deployment associated with assignment..." -l 5 -v 2
		# The specific deployment for this assignment can be pulled using the AdvertisementID
		
		if($DisableCaching) {
			$deployment = Get-CMTaskSequenceDeployment -AdvertisementID $assignment._TSDepID
			
			if($deployment) {
				log "Retrieved TS deployment." -l 6 -v 2
				$deployment = Parse-AssignmentTSDeployment $deployment
			}
			else {
				log "TS Deployment not retrieved from SCCM!" -l 6
			}
		}
		else {
			$deployment = Get-CachedItem "tsdep" $assignment._TSDepID
		}
		
		log "Done getting TS deployment associated with assignment." -l 5 -v 2
		$deployment
	}
	
	function Parse-AssignmentTSDeployment($deployment) {
		log "Parsing TS deployment..." -l 6 -v 2
		
		$deployment | Add-Member -NotePropertyName "_Type" -NotePropertyValue "ts"
		
		# TSDeployment objects do not store infomation about apps or revisions (that's in the TS itself)
		# The only real useful information they store is the TS name, and the relevant collection name
		# These are both stored in the AdvertisementName field in the format "<TS name>_<TS PackageID>_<collection name>"
		# However the TS and collection names here have the whitespace stripped
		# PackageID of the TS and the CollectionID are both stored as separate fields
		$deployment | Add-Member -NotePropertyName "_Name" -NotePropertyValue $deployment.AdvertisementName
		
		if($DisableCaching) {
			$ts = Get-CMTaskSequence -PackageId $deployment.PackageID
		}
		else {
			$ts = Get-CachedItem "ts" $deployment.PackageID
		}
		$deployment | Add-Member -NotePropertyName "_ContentName" -NotePropertyValue $ts.Name
		
		if($DisableCaching) {
			$tsCollection = Get-CMCollection -CollectionId $deployment.CollectionID
		}
		else {
			$tsCollection = Get-CachedItem "collection" $deployment.CollectionID
		}
		$deployment | Add-Member -NotePropertyName "_Collection" -NotePropertyValue $tsCollection.Name
		
		# These two won't be reliable if TS or collection names contain underscores
		$adNameParts = $deployment.AdvertisementName.Split("_")
		$deployment | Add-Member -NotePropertyName "_ContentNameStripped" -NotePropertyValue $adNameParts[0]
		$deployment | Add-Member -NotePropertyName "_CollectionStripped" -NotePropertyValue $adNameParts[2]
		
		log "TS deployment name: `"$($deployment._Name)`"." -l 7 -v 2
		log "TS ID: `"$($deployment.PackageID)`"." -l 7 -v 2
		log "TS name: `"$($deployment._ContentName)`"." -l 7 -v 2
		log "Collection ID: `"$($deployment.CollectionID)`"." -l 7 -v 2
		log "Collection Name: `"$($deployment._Collection)`"." -l 7 -v 2
		
		$deployment | Add-Member -NotePropertyName "_Revision" -NotePropertyValue "TS"
		$deployment | Add-Member -NotePropertyName "_ModelName" -NotePropertyValue "No ModelName. This is a TS deployment."
		
		log "Done parsing TS deployment." -l 6 -v 2
		$deployment
	}
	
	function Get-AssignmentAppDeployment($assignment) {
		log "Getting app deployment associated with assignment..." -l 5 -v 2
		# The specific deployment for this assignment can be pulled using the assignmentID
		
		if($DisableCaching) {
			$deployment = Get-CMApplicationDeployment -AssignmentUniqueID $assignment.AssignmentID
			
			if($deployment) {
				log "Retrieved app deployment." -l 6 -v 2
				$deployment = Parse-AssignmentAppDeployment $deployment
			}
			else {
				log "App deployment not retrieved from SCCM!" -l 6
			}
		}
		else {
			$deployment = Get-CachedItem "appdep" $assignment.AssignmentID
		}
		
		log "Done getting app deployment associated with assignment..." -l 5 -v 2
		$deployment
	}

	function Parse-AssignmentAppDeployment($deployment) {
		log "Parsing app deployment..." -l 6 -v 2
		
		$deployment | Add-Member -NotePropertyName "_Type" -NotePropertyValue "app"
		
		$deployment | Add-Member -NotePropertyName "_Name" -NotePropertyValue $deployment.AssignmentName
		$deployment | Add-Member -NotePropertyName "_ContentName" -NotePropertyValue $deployment.ApplicationName
		$deployment | Add-Member -NotePropertyName "_Collection" -NotePropertyValue $deployment.CollectionName
		
		log "App deployment name: `"$($deployment.AssignmentName)`"." -l 7 -v 2
		log "App ID: `"$($deployment.AssignedCI_UniqueID)`"." -l 7 -v 2
		log "App name: `"$($deployment.ApplicationName)`"." -l 7 -v 2
		log "Collection ID: `"$($deployment.TargetCollectionID)`"." -l 7 -v 2
		log "Collection Name: `"$($deployment._Collection)`"." -l 7 -v 2
		
		# The AssignedCI_UniqueID field is of the format "<ModelName>/<revision>"
		# This is the only place where a deployment natively stores this data
		$depCiid = $deployment.AssignedCI_UniqueID
		$depCiidParts = $depCiid.Split("/")
		
		$deployment | Add-Member -NotePropertyName "_Revision" -NotePropertyValue $depCiidParts[2]
		$depModelName = $depCiidParts[0] + "/" + $depCiidParts[1]
		$deployment | Add-Member -NotePropertyName "_ModelName" -NotePropertyValue $depModelName
		
		$deployment = Parse-DesiredConfigType $deployment
		
		log "Done parsing app deployment." -l 6 -v 2
		$deployment
	}

	# Get application for this assignment
	function Get-AssignmentApplication($assignment) {
		log "Getting application associated with assignment..." -l 5 -v 2
		
		$genericModelName = $assignment._ModelName -replace "RequiredApplication","Application"
		$genericModelName = $assignment._ModelName -replace "ProhibitedApplication","Application"
			
		if($DisableCaching) {
			$app = Get-CMApplication -Fast -ModelName $genericModelName
			if($app) {
				log "Retrieved application." -l 6 -v 2
				$app = Parse-AssignmentApplication $app
			}
			else {
				log "Application not retrieved from SCCM!" -l 6
			}
		}
		else {
			$app = Get-CachedItem "app" $genericModelName
		}
		
		log "Done getting application associated with assignment." -l 5 -v 2
		
		$assignment | Add-Member -NotePropertyName "_Application" -NotePropertyValue $app
		
		$assignment
	}

	# Parse application for this assignment
	function Parse-AssignmentApplication($app) {
		log "Parsing application..." -l 6 -v 2
		log "Application name: `"$($app.LocalizedDisplayName)`"..." -l 6 -v 2
		
		# CI version (a.k.a. revision) is stored in its own CIVersion field and also as part of the CI_UniqueID string (formatted like "<ModelName>/<CIVersion>")
		# Grab the version out of the CI_UniqueID string to make sure it's the same as what's in the CIVersion field
		$appCiid = $app.CI_UniqueID
		$appCiidParts = $appCiid.Split("/")
		$app | Add-Member -NotePropertyName "_Revision" -NotePropertyValue $appCiidParts[2]
		# Do the same for the ModelName
		$appModelName = $appCiidParts[0] + "/" + $appCiidParts[1]
		$app | Add-Member -NotePropertyName "_ModelName" -NotePropertyValue $appModelName
	
		log "Done parsing application." -l 6 -v 2
		$app
	}
	
	function Get-CachedItem($type, $id) {
	
		$logType = "unknown item type"
		$idType = "UnknownIdType"
		$cacheVar = "UnknownCacheVar"
		$getCmd = "log `"Unknown get command!`" -l 7"
		$parseCmd = "log `"Unknown parse command!`" -l 7"
		
		switch($type) {
			"tsdep" {
				$logType = "TS deployment"
				$idType = "AdvertisementId"
				$cacheVar = "CachedTSDeployments"
				$getCmd = "Get-CMTaskSequenceDeployment -$idType `"$id`""
				$parseCmd = "Parse-AssignmentTSDeployment `$cachedItem"
				break
			}
			"appdep" {
				$logType = "Application deployment"
				$idType = "AssignmentUniqueId"
				$cacheVar = "CachedTSDeployments"
				$getCmd = "Get-CMApplicationDeployment -$idType `"$id`""
				$parseCmd = "Parse-AssignmentAppDeployment `$cachedItem"
				break
			}
			"app" {
				$id = $id -replace "RequiredApplication","Application"
				$logType = "Application"
				$idType = "ModelName"
				$cacheVar = "CachedApplications"
				$getCmd = "Get-CMApplication -$idType `"$id`""
				$parseCmd = "Parse-AssignmentApplication `$cachedItem"
				break
			}
			"ts" {
				$logType = "Task sequence"
				$idType = "PackageId"
				$cacheVar = "CachedTaskSequences"
				$getCmd = "Get-CMTaskSequence -$idType `"$id`""
				$parseCmd = "log `"No parse command for task sequences`" -l 7 -v 2"
				break
			}
			"collection" {
				$logType = "Collection"
				$idType = "CollectionId"
				$cacheVar = "CachedCollections"
				$getCmd = "Get-CMCollection -$idType `"$id`""
				$parseCmd = "log `"No parse command for collections`" -l 7 -v 2"
				break
			}
			default {
				$logType = "unrecognized item type"
				$idType = "UnrecognizedIdType"
				$cacheVar = "UnrecognizedCacheVar"
				$getCmd = "log `"Unrecognized get command!`" -l 7"
				$parseCmd = "log `"Unrecognized parse command!`" -l 7"
			}
		}
		
		log "Searching for $logType in cache..." -l 6 -v 2
		
		log "Using `$idType: `"$idType`"" -l 7 -v 3
		log "Using `$id: `"$id`"" -l 7 -v 3
		log "Using `$cacheVar: `"$cacheVar`"" -l 7 -v 3
		log "Using `$getCmd: `"$getCmd`"" -l 7 -v 3
		log "Using `$parseCmd: `"$parseCmd`"" -l 7 -v 3
		
		$currentCache = Get-Variable -Name $cacheVar -Scope Script -ValueOnly
		$totalCount = @($currentCache).count
		log "Currently $totalCount $logType`s cached." -l 7 -v 2
		
		$cachedItems = $currentCache | Where { $_.$idType -eq $id }
		$count = @($cachedItems).count
		
		# If this item hasn't been cached yet
		if($count -eq 0) {
			log "$logType $idType not found in cache." -l 7 -v 2
			
			log "Retrieving $logType from SCCM..." -l 7 -v 2
			$cachedItem = Invoke-Expression $getCmd
			
			# If item exists
			if($cachedItem) {
				log "$logType exists in SCCM." -l 7 -v 2
				
				# Parse it
				# Don't parse TSes or Collections as it would just overwrite the item with nothing
				# I could make a parse function for these that just returns the item as-is, but meh
				if(($type -ne "ts") -and ($type -ne "collection")) {
					$cachedItem = Invoke-Expression $parseCmd
				}
				
				# Cache it and return it
				$found = $cachedItem
				log "$logType will be cached." -l 7 -v 2
				log "Will return newly retrieved and cached $logType." -l 7 -v 2
			}
			# If item doesn't exist
			else {
				log "$logType doesn't exist in SCCM." -l 7 -v 2
				# Cache that fact and return $null
				# https://ridicurious.com/2018/10/15/4-ways-to-create-powershell-objects/
				$cachedItem = [PSCustomObject]@{
					$idType = $id
					"NonExistent" = $true
				}
				log "Will cache that $logType doesn't exist." -l 7 -v 2
				log "Will return no $logType." -l 7 -v 2
				# $found = $null
			}
			
			log $cachedItem -nots -v 3
			
			$oldCache = Get-Variable -Name $cacheVar -Scope Script -ValueOnly
			$newCache = $oldCache + @($cachedItem)
			Set-Variable -Name $cacheVar -Scope Script -Value $newCache
		}
		elseif($count -eq 1) {
			log "$logType $idType found in cache." -l 7 -v 2
			
			# Technically unnecessary for single member arrays, but this is more straightforward
			# I'm always weirded out and tripped up by single member arrays in Powershell not actually being arrays -_-
			$cachedItem = @($cachedItems)[0]
			
			# If the item has been cached as non-existent
			if($cachedItem.NonExsistent) {
				log "$logType cached as non-existent in SCCM." -l 7 -v 2
				log "Will return no $logType." -l 7 -v 2
				# $found = $null
			}
			# If the item was cached
			else {
				log "$logType cached as returned from SCCM." -l 7 -v 2
				log "Will return previously retrieved and cached $logType." -l 7 -v 2
				# Return cached item
				$found = $cachedItem
			}
		}
		elseif($count -gt 1) {
			log "Error: somehow the same $logType was cached twice!" -l 7
		}
		elseif($count -lt 0) {
			log "Error: somehow this $logType was cached a negative amount of times!" -l 7
		}
		else {
			log "Unknown error retrieving cached $logType!" -l 7
		}
		
		$currentCache = Get-Variable -Name $cacheVar -Scope Script -ValueOnly
		$totalCount = @($currentCache).count
		log "Now $totalCount $logType`s cached." -l 7 -v 2
		
		log "Done searching for $logType in cache." -l 6 -v 2
		$found
	}
	
	function Compare-Revisions($assignment) {
		log "Comparing revisions of assignment, deployment, and application... " -l 5 -v 1 -nnl
		
		# Save custom property for whether the revisions match, for easy table printout
		$as1 = $assignment._Revision
		$as2 = $assignment._CI.CIVersion
		$dep = $assignment._Deployment._Revision
		$app1 = $assignment._Application._Revision
		$app2 = $assignment._Application.CIVersion
		
		$same = "MISMATCH!"
		if($assignment._DepType -eq "ts") {
			if(
				($as1 -eq $as2) -and
				($as2 -eq $app1) -and
				($app1 -eq $app2)
			) {
				$same = "yes"
				log "Revisions match." -nots -v 1
			}
			else {
				log "REVISIONS DO NOT MATCH!" -nots -v 1
			}
		}
		elseif($assignment._DepType -eq "app") {
			if(
				($as1 -eq $as2) -and
				($as2 -eq $dep) -and
				($dep -eq $app1) -and
				($app1 -eq $app2)
			) {
				$same = "yes"
				log "Revisions match." -nots -v 1
			}
			else {
				log "REVISIONS DO NOT MATCH!" -nots -v 1
			}
		}
		else {
			$same = "Deployment type error!"
			log $same -nots -v 1
		}
		
		$assignment | Add-Member -NotePropertyName "_AllRevisionsMatch" -NotePropertyValue $same
		
		log "Done comparing revisions." -l 5 -v 2
		$assignment
	}
	
	function Compare-DesiredConfigTypes($assignment) {
		log "Comparing DesiredConfigTypes of assignment and application... " -l 5 -v 1 -nnl
		
		# Save custom property for whether the revisions match, for easy table printout
		$ascf = $assignment._DesiredConfigType
		$depcf = $assignment._Deployment._DesiredConfigType
		
		$same = "MISMATCH!"
		if($ascf -eq $depcf) {
			$same = "yes"
			log "DesiredConfigTypes match." -nots -v 1
		}
		else {
			log "DesiredConfigTypes do not match!" -nots -v 1
		}
		
		$assignment | Add-Member -NotePropertyName "_DesiredConfigTypesMatch" -NotePropertyValue $same
		
		log "Done comparing DesiredConfigTypes." -l 5 -v 2
		$assignment
	}
	
	function Compare-ModelNames($assignment) {
		log "Comparing ModelNames of assignment, deployment, and application... " -l 5 -v 1 -nnl
		
		# Save custom property for whether the ModelNames match, mostly for doublechecking that we're comparing the correct apps, and also for easy table printout
		$as1 = $assignment._ModelName -replace "RequiredApplication","Application"
		$as2 = $assignment._CI.ModelName -replace "RequiredApplication","Application"
		$dep = $assignment._Deployment._ModelName -replace "RequiredApplication","Application"
		$app1 = $assignment._Application._ModelName -replace "RequiredApplication","Application"
		$app2 = $assignment._Application.ModelName -replace "RequiredApplication","Application"
		
		$same = "MISMATCH!"
		if($assignment._DepType -eq "ts") {
			if($as1 -eq $as2 -eq $app1 -eq $app2) {
				$same = "yes"
				log "ModelNames match." -nots -v 1
			}
			else {
				log "MODELNAMES DO NOT MATCH!" -nots -v 1
			}
		}
		elseif($assignment._DepType -eq "app") {
			if($as1 -eq $as2 -eq $dep -eq $app1 -eq $app2) {
				$same = "yes"
				log "ModelNames match." -nots -v 1
			}
			else {
				log "MODELNAMES DO NOT MATCH!" -nots -v 1
			}
		}
		else {
			$same = "Deployment type error!"
			log $same -nots -v 1
		}
		
		$assignment | Add-Member -NotePropertyName "_AllModelNamesMatch" -NotePropertyValue $same
		
		log "Done comparing ModelNames.$msg" -l 5 -v 2
		$assignment
	}
	
	function Export-CompAssignments($comp) {
		if($comp.localassignments) {
			log "Exporting assignments for `"$($comp.name)`" to `"$CSVPATH`"..." -l 2
			foreach($assignment in $comp.localassignments) {
				Export-Assignment $assignment
			}
		}
		else {
			log "No assignments to export to CSV for `"$($comp.name)`"." -l 2
		}
		log "Done exporting assignments." -l 2 -v 2
	}
	
	function Export-Assignment($assignment) {
		
		function line($line) {
			$line | Out-File $CSVPATH -Append -Encoding ascii
		}

		# Make CSV file and add header row if file doesn't exist
		if(!(Test-Path -PathType leaf -Path $CSVPATH)) {
			$shutup = New-Item -ItemType File -Force -Path $CSVPATH
			line "Computer,ClientVer,PSVer,OSVer,Make,Model,AssignmentID,AssignmentName,DeploymentName,DeploymentCollection,DeploymentContent,ApplicationName,AsConfigType,DepConfigType,ConfigTypesMatch,AsRev1,AsRev2,DepRev,AppRev1,AppRev2,RevsMatch,ModelsMatch,AsModel1,AsModel2,DepModel1,AppModel1,AppModel2"
		}
		
		$line = "`"" + $assignment._Computer + "`"," +
			"`"" + $assignment._SCCMClientVersion + "`"," +
			"`"" + $assignment._PSVersion + "`"," +
			"`"" + $assignment._OSVersion + "`"," +
			"`"" + $assignment._Make + "`"," +
			"`"" + $assignment._Model + "`"," +
			"`"" + $assignment.AssignmentID + "`"," +
			"`"" + $assignment._Name + "`"," +
			"`"" + $assignment._Deployment._Name + "`"," +
			"`"" + $assignment._Deployment._Collection + "`"," +
			"`"" + $assignment._Deployment._ContentName + "`"," +
			"`"" + $assignment._Application.LocalizedDisplayName + "`"," +
			"`"" + $assignment._DesiredConfigType + "`"," +
			"`"" + $assignment._Deployment._DesiredConfigType + "`"," +
			"`"" + $assignment._DesiredConfigTypesMatch + "`"," +
			"`"" + $assignment._Revision + "`"," +
			"`"" + $assignment._CI.CIVersion + "`"," +
			"`"" + $assignment._Deployment._Revision + "`"," +
			"`"" + $assignment._Application._Revision + "`"," +
			"`"" + $assignment._Application.CIVersion + "`"," +
			"`"" + $assignment._AllRevisionsMatch + "`"," +
			"`"" + $assignment._AllModelNamesMatch + "`"," +
			"`"" + $assignment._ModelName + "`"," +
			"`"" + $assignment._CI.ModelName + "`"," +
			"`"" + $assignment._Deployment._ModelName + "`"," +
			"`"" + $assignment._Application._ModelName + "`"," +
			"`"" + $assignment._Application.ModelName + "`""
		
		line $line
	}
	
	function Get-CompNameList($compNames) {
		$list = ""
		foreach($name in $compNames) {
			$list = "$list, $name"
		}
		$list = $list.Substring(2,$list.length - 2) # Remove leading ", "
		$list
	}
	
	function Get-CompNames {
		log "Getting list of computer names..."
		if($Computers) {
			log "List was given as an array." -l 1 -v 1
			$compNames = $Computers
			$list = Get-CompNameList $compNames
			log "Found $($compNames.count) computers in given array: $list." -l 1
		}
		elseif($Collection) {
			log "List was given as a collection. Getting members of collection: `"$Collection`"..." -l 1 -v 1
			$colObj = Get-CMCollection -Name $Collection
			if(!$colObj) {
				log "The given collection was not found!" -l 1
			}
			else {
				# Get comps
				$comps = Get-CMCollectionMember -CollectionName $Collection | Select Name,ClientActiveStatus
				if(!$comps) {
					log "The given collection is empty!" -l 1
				}
				else {
					# Sort by active status, with active clients first, just in case inactive clients might come online later
					# Then sort by name, just for funsies
					$comps = $comps | Sort -Property @{Expression = {$_.ClientActiveStatus}; Descending = $true}, @{Expression = {$_.Name}; Descending = $false}
					
					$compNames = $comps.Name
					$list = Get-CompNameList $compNames
					log "Found $($compNames.count) computers in `"$Collection`" collection: $list." -l 1
				}
			}
		}
		else {
			log "Somehow neither the -Computers, nor -Collection parameter was specified!" -l 1
		}
		
		log "Done getting list of computer names." -v 2
		
		$compNames
	}

	function Do-Stuff {
		log " " -nots
		
		$myPWD = $pwd.path
		Prep-SCCM
		
		$compNames = Get-CompNames
		if($compNames) {
			$comps = Get-CompObjects $compNames
			$comps = Get-CompData $comps
		}
		
		Set-Location $myPWD
		
		log "EOF"
		log " " -nots
	}
	
	Do-Stuff
}
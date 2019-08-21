$OptimizeRoute = {

	$SystemsInRoute = $SysSetBox.Rows.count

	if($SystemsInRoute -gt 1){
		
		$HasFirst = $Flase
		$HasLast = $False

		# give the use a warning
		$DoRun = [Microsoft.VisualBasic.Interaction]::MsgBox("You are about optimize you route. `n This will alter the order of you selected systems `n Do you wish to Continue?",'YesNo,Question', "Create New Route?")
		
		if($DoRun -eq 'yes'){
		
			$Systems = @()
			$ErrorCount = 0
			$SystemID = 0
			
			# loop over all rows added to route
			foreach($Row in $SysSetBox.Rows){
			
				# read system name
				$SystemName = $Row.cells[0].Value
				
				# get Coordinates
				$X, $Y, $Z = GetCoords $Row		
				
				# let's see if there is any first / last systems
				$IsFirst = $false
				$IsLast = $false
			
				if($Row.Cells[1].Value){
					$HasFirst = $True
					$IsFirst = $True
				}
				
				if($Row.Cells[2].Value){
					$HasLast = $True
					$IsLast = $True
				}
				
				# Save row data in PSObject (easier to work with later
				$SystemsObj = New-Object -TypeName psobject -Property @{
					ID = $SystemID
					Name = $SystemName
					x = $x
					y = $y
					z = $z
					First = $IsFirst
					Last = $IsLast
				}
				$Systems += $SystemsObj
				
				$SystemID += 1
			}
			
			# Once all system data have been read, we make sure we didnt encounter any errors.
			# If we do we tell the user and leav hte optimizer.
			If($ErrorCount -ne 0){
				[Microsoft.VisualBasic.Interaction]::MsgBox($('Error encountered during system metadata load' -f $SystemName),'OKonly', "Unknown System") | out-null
			}else{
				# time to optimize the route!
				
				#Build system dist array
				$SystemDistanceArray = New-Object 'object[,]' $SystemsInRoute,$SystemsInRoute

				for($i=0;$i -lt $SystemsInRoute;$i++){
					for($j=0;$j -lt $SystemsInRoute;$j++){
						$SystemDistanceArray[$i,$j] = [math]::sqrt([math]::pow(($Systems[$i].x-$Systems[$j].x),2) + [math]::pow(($Systems[$i].y-$Systems[$j].y),2) + [math]::pow(($Systems[$i].z-$Systems[$j].z),2))
					}
				}
				
				$Route = @() 
				
				# add first and last system to route
				If($HasFirst){
					$Route += $Systems | where{$_.First -eq $True}
				}
				If($HasLast){
					$Route += $Systems | where{$_.Last -eq $True}
				}
				
				#make a list of system to loop over
				$LoopSystem = $Systems | where{($_.First -eq $False) -and ($_.Last -eq $False)}
				
				foreach($System in $LoopSystem){
				
					$BestLength = 0
					
					$CurrentRouteCount = $($Route.count)		
					
					
					for($i = 0; $i -le $CurrentRouteCount; $i++){
						if( -not (($HasFirst -and $i -eq 0) -or ($HasLast -and $i -eq $CurrentRouteCount))){
							
							# build the route to test
							if($i -eq 0){
								$TestRoute = @($System) + $Route
							}elseif($i -eq $CurrentRouteCount){
								$TestRoute = $Route + $System
							}else{
								$TestRoute = $Route[0..($i-1)] + $System + $Route[$i..($CurrentRouteCount-1)]
							}
							
							$CurrentRouteLength = GetRouteLength $TestRoute $SystemDistanceArray
							
							If(($CurrentRouteLength -lt $BestLength) -or ($BestLength -eq 0)){
								$BestLength = $CurrentRouteLength
								$BestRoute = $TestRoute
							}
						}
					}
					
					$Route = $BestRoute
				}
				
				# No longer latest version
				$global:Saved = $False
				
				# clear DataGridView
				$SysSetBox.ROWS.Clear()
				
				$FirstSystem = $True
				
				foreach($System in $Route){
					#Pull data for new row
					$SystemName = $System.name
					$IsFirst = $System.First
					$IsLast = $System.Last
					
					If($IsFirst -and $IsLast){
						if($FirstSystem){
							$IsLast = $False
							
						}else{
							$IsFirst = $False
						}
					}
					
					$Coords = $('{0};{1};{2}' -f $System.x, $System.y, $System.z)
					
					# Yes yes yes i know. it's not pretty.
					if($HasFirst){
						if($IsFirst){
							$StartReadOnly = $False
						}else{				
							$StartReadOnly = $True
						}
					}else{
						$StartReadOnly = $False
					}
					
					if($HasLast){
						if($IsLast){
							$EndReadOnly = $False
						}else{				
							$EndReadOnly = $True
						}
					}else{
						$EndReadOnly = $False
					}
					
					
					#adding row to DataGridView
					$SysSetBox.Rows.Add($SystemName,$IsFirst, $IsLast)
					
					# finding row we just added
					$MaxRowIndex = $($SysSetBox.Rows.count - 1)
					$NewRow = $SysSetBox.Rows[$MaxRowIndex]
					
					#setting readonly state
					$NewRow.Cells[1].ReadOnly = $StartReadOnly
					$NewRow.Cells[2].ReadOnly = $EndReadOnly
					
					$NewRow.tag = $Coords
					
					$SysSetBox.EndEdit()
					
					$FirstSystem = $False
				}
			}
		}
	}
	UpdateRouteLength $SysSetBox $LengthLabel
}
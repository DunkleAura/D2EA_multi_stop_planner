Function UpdateRouteLength(
	$DataGridView,
	$LableBox,
	$ExclutedSystem = -1
){

	write-host $('Start UpdateRouteLength')
	
	#basic params
	$SystemsInRoute = $DataGridView.Rows.Count
	$TotalLength = 0
	
	write-host $('Systems in Route: {0}' -f $SystemsInRoute)
	write-host $('Systems Excluted: {0}' -f $ExclutedSystem)
	
	If($SystemsInRoute -gt 1){
	
		# init previous system
		$PreviousSystem = $DataGridView.Rows[0]
		
		$PreviousX, $PreviousY, $PreviousZ = GetCoords $PreviousSystem
		
		For($i = 1; $i -lt $DataGridView.Rows.Count; $i++){
			if($i -ne $ExclutedSystem){
			
				#Store Current System
				$CurrentSystem = $DataGridView.Rows[$i]
				$CurrentX, $CurrentY, $CurrentZ = GetCoords $CurrentSystem
				
				$DeltaX = $CurrentX - $PreviousX
				$DeltaY = $CurrentY - $PreviousY
				$DeltaZ = $CurrentZ - $PreviousZ
				
				$TotalLength += [math]::sqrt([math]::pow($DeltaX,2) + [math]::pow($DeltaY,2) + [math]::pow($DeltaZ,2))
				
				$PreviousSystem = $CurrentSystem
				$PreviousX = $CurrentX
				$PreviousY = $CurrentY
				$PreviousZ = $CurrentZ
			}
		}
	}
	
	
	
	$LengthLabel.text = $('Total Route Length: {0} Ly' -f $([math]::Round($TotalLength,2)))
	write-host $('Route Length update to: {0}' -f $TotalLength)
	write-host $('End UpdateRouteLength')
}

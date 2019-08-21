$WPAddClick = {
	$SystemName = $textBox.Text
	
	if($SystemName -ne ''){
		$systemMetadata = Invoke-WebRequest $('https://www.edsm.net/api-v1/system?systemName={0}&showCoordinates=1' -f $SystemName) -UseBasicParsing
		
		if($systemMetadata.content -match '{"name":"([^"]+)","coords":{"x":([\d.-]+),"y":([\d.-]+),"z":([\d.-]+)}'){
			
			$global:Saved = $False
			
			$SystemName = $matches[1]
			$Coords = $('{0};{1};{2}' -f $matches[2], $matches[3], $matches[4])
	
			$StartReadOly = $False
			$EndReadOnly = $False
			
			foreach($Row in $SysSetBox.Rows){
				if($Row.Cells[1].Value){
					$StartReadOly = $True
				}
				
				if($Row.Cells[2].Value){
					$EndReadOnly = $True
				}
			}
			
			$SysSetBox.Rows.Add($SystemName,$false, $false)
			
			$MaxRowIndex = $($SysSetBox.Rows.count - 1)
			$NewRow = $SysSetBox.Rows[$MaxRowIndex]
			
			$NewRow.Cells[1].ReadOnly = $StartReadOly
			$NewRow.Cells[2].ReadOnly = $EndReadOnly
			$NewRow.tag = $Coords
			
			$SysSetBox.EndEdit()
			UpdateRouteLength $SysSetBox $LengthLabel
		}else{
			[Microsoft.VisualBasic.Interaction]::MsgBox($('System [{0}] Unknown' -f $SystemName),'OKonly', "Unknown System") | out-null
		}
	}
}
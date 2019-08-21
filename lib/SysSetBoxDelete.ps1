$SysSetBoxDelete = {
	
	write-host 'Start SysSetBoxDelete'
	
	$SysSetBox.EndEdit()
	$rowIndex = $SysSetBox.CurrentRow.Index
	$MaxRowIndex = $($SysSetBox.Rows.count - 1)
	
	write-host $('RowIndex: {0}' -f $rowIndex)
	
	$IsStart = $SysSetBox.Rows[$rowIndex].Cells[1].Value
	$IsEnd = $SysSetBox.Rows[$rowIndex].Cells[2].Value
	
	write-host $('Row is start: {0}' -f $IsStart)
	write-host $('Row is end: {0}' -f $IsEnd)
	
	$MaxRowIndex = $($SysSetBox.Rows.count - 1)
	For($i = 0; $i -le $MaxRowIndex; $i++){
		
		$global:Saved = $False
		
		If($IsStart){
			$SysSetBox.Rows[$i].Cells[1].ReadOnly = $False
		}
		If($IsEnd){
			$SysSetBox.Rows[$i].Cells[2].ReadOnly = $False
		}
		
	}	
	$SysSetBox.EndEdit()
	
	write-host 'Updating Route Length'
	
	UpdateRouteLength $SysSetBox $LengthLabel $rowIndex
	
	write-host 'End SysSetBoxDelete'
}
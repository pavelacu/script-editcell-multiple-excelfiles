##### EDIT A CELL IN EXCEL FILE ###########
function EditCellByFile([string]$filename, [string]$sheet, [int]$rowId, [string]$colId,[string]$value )
{
    # Open the Excel document and pull in the 'Play' worksheet    
    $Excel = New-Object -Com Excel.Application
    $Workbook = $Excel.Workbooks.Open($filename)     
    $ws = $Workbook.worksheets | where-object {$_.Name -eq $page}
    $cells=$ws.Cells
    $cells.item($rowId,$colId)= $value
    ## Close the workbook and exit Excel
    $workbook.Close($true) 
    $excel.quit()
}

##### script-edit-cell-multiple-excelfiles. ###########
#   Example edit the Cell F21, value TEST
# 	parameters:
#   colId: string
#   rowId: integer
#   sheet: string
#   value: string
#
#	Example:
#	colId = "F"
#   rowId = 21
#   sheet = "sheet1"
#   value = "Hello Test"
######

# Libraries
Add-Type -AssemblyName Microsoft.VisualBasic
# set Specific folder
#$path = 'c:\...'
Write-Output("Init..")

# Set Actually folder
$path = [string](Get-Location)+"\"

$ext='*.xlsx'
# get list files with extension *.xlsx
$files = Get-ChildItem -Path $Path -Recurse -Force -Name -Include $ext
# set Column edit
$colId=[Microsoft.VisualBasic.Interaction]::InputBox('Enter a Letter Column ', 'Column', "")
# set Row edit
$rowId=[int][Microsoft.VisualBasic.Interaction]::InputBox('Enter a number Row', 'Row', "")
#set sheetName edit
$sheetName = [Microsoft.VisualBasic.Interaction]::InputBox('Enter a SheetName', 'SheetName', "")
# set specific value from inputBox
$value= [Microsoft.VisualBasic.Interaction]::InputBox('Enter a value', 'Value', "")
# set specific value
Write-Output(" Sheet: "+$sheetName+" - Cell: "+$colId+$rowId +" - value: "+$value)
foreach ($file in $files){
    $filename = $path+$file
    Write-Output("Edit: "+ $filename)
    EditCellByFile -filename $filename -sheet $sheetName -rowId $rowId -colId $colId -value $value
}
Write-Output("END.")


 #Declare the file path and sheet name
$file = "C:\Users\$env:USERNAME\Documents\SMS_Autoresponder\TargetList.xlsx"


$sheetName = "TargetSheet"
#Create an instance of Excel.Application and Open Excel file
$objExcel = New-Object -ComObject Excel.Application
$workbook = $objExcel.Workbooks.Open($file)
$sheet = $workbook.Worksheets.Item($sheetName)
$objExcel.Visible=$false



#Count max row
$rowMax = ($sheet.UsedRange.Rows).count
#Declare the starting positions


$rowName,$colName = 1,1
$rowNumber,$colNumber = 1,2



#loop to get values and store it
for ($i=1; $i -le $rowMax-1; $i++)
{
$name = $sheet.Cells.Item($rowName+$i,$colName).text
$Number = $sheet.Cells.Item($rowNumber+$i,$colNumber).text


Write-Host ("Target- "+$name )
Write-Host ("Phone Number to Send SMS" +$Number)


}


#close excel file
$objExcel.quit() 
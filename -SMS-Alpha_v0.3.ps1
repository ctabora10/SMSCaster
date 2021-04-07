<# This form was created using POSHGUI.com  a free online gui designer for PowerShell
.NAME
    Untitled
#>
Add-Type -AssemblyName System.Windows.Forms
[System.Windows.Forms.Application]::EnableVisualStyles()

$Form                            = New-Object system.Windows.Forms.Form
$Form.ClientSize                 = '822,408'
$Form.text                       = "Ctab.Tech SMS Autoresponder"
$Form.BackColor                  = "#4a4a4a"
$Form.TopMost                    = $false




$Button1                         = New-Object system.Windows.Forms.Button
$Button1.BackColor               = "#4a90e2"
$Button1.text                    = "Send Single SMS"
$Button1.width                   = 220
$Button1.height                  = 48
$Button1.location                = New-Object System.Drawing.Point(507,94)
$Button1.Font                    = 'Microsoft Sans Serif,12,style=Bold'
$Button1.ForeColor               = "#ffffff"

$Button2                         = New-Object system.Windows.Forms.Button
$Button2.BackColor               = "#4a90e2"
$Button2.text                    = "Check Message Quota"
$Button2.width                   = 220
$Button2.height                  = 48
$Button2.location                = New-Object System.Drawing.Point(507,228)
$Button2.Font                    = 'Microsoft Sans Serif,12,style=Bold'
$Button2.ForeColor               = "#ffffff"

$Label1                          = New-Object system.Windows.Forms.Label
$Label1.text                     = "-SMS- Autoresponder"
$Label1.AutoSize                 = $true
$Label1.width                    = 25
$Label1.height                   = 10
$Label1.location                 = New-Object System.Drawing.Point(45,36)
$Label1.Font                     = 'Segoe UI Symbol,30,style=Bold'
$Label1.ForeColor                = "#ffffff"

$Label2                          = New-Object system.Windows.Forms.Label
$Label2.text                     = "Version Build 0.3 -Alpha-"
$Label2.AutoSize                 = $true
$Label2.width                    = 25
$Label2.height                   = 10
$Label2.location                 = New-Object System.Drawing.Point(157,302)
$Label2.Font                     = 'Segoe UI Symbol,10,style=Bold'
$Label2.ForeColor                = "#ffffff"

$Button3                         = New-Object system.Windows.Forms.Button
$Button3.BackColor               = "#ee576b"
$Button3.text                    = "RCMH Patient Outreach"
$Button3.width                   = 220
$Button3.height                  = 48
$Button3.location                = New-Object System.Drawing.Point(507,296)
$Button3.Font                    = 'Microsoft Sans Serif,12,style=Bold'
$Button3.ForeColor               = "#ffffff"

$Button4                         = New-Object system.Windows.Forms.Button
$Button4.BackColor               = "#4a90e2"
$Button4.text                    = "Purchase More Tokens"
$Button4.width                   = 164
$Button4.height                  = 36
$Button4.location                = New-Object System.Drawing.Point(153,340)
$Button4.Font                    = 'Microsoft Sans Serif,9,style=Bold'
$Button4.ForeColor               = "#ffffff"

$Button5                         = New-Object system.Windows.Forms.Button
$Button5.BackColor               = "#4a90e2"
$Button5.text                    = "Send Mass SMS"
$Button5.width                   = 220
$Button5.height                  = 48
$Button5.location                = New-Object System.Drawing.Point(507,158)
$Button5.Font                    = 'Microsoft Sans Serif,12,style=Bold'
$Button5.ForeColor               = "#ffffff"

$PictureBox1                     = New-Object system.Windows.Forms.PictureBox
$PictureBox1.width               = 194
$PictureBox1.height              = 184
$PictureBox1.location            = New-Object System.Drawing.Point(137,108)
$PictureBox1.imageLocation       = "http://Vault.xctab.com/Picon.png"
$PictureBox1.SizeMode            = [System.Windows.Forms.PictureBoxSizeMode]::zoom
$Form.controls.AddRange(@($Button1,$Button2,$Label1,$Label2,$Button3,$Button4,$Button5,$PictureBox1))

$Button1.Add_Click({ Send_Single })
$Button2.Add_Click({ Check_Quota })
$Button3.Add_Click({ RCMH_SMS })
$Button4.Add_Click({ TextbeltURL })
$Button5.Add_Click({ MassSMS })








function TextbeltURL {start https://textbelt.com/ }


function Send_Single { 

<# This form was created using POSHGUI.com  a free online gui designer for PowerShell
.NAME
    USingleSMS
#>

Add-Type -AssemblyName System.Windows.Forms
[System.Windows.Forms.Application]::EnableVisualStyles()

$Form                            = New-Object system.Windows.Forms.Form
$Form.ClientSize                 = '512,646'
$Form.text                       = "Single SMS"
$Form.BackColor                  = "#4a4a4a"
$Form.TopMost                    = $false

$Label1                          = New-Object system.Windows.Forms.Label
$Label1.text                     = "Single SMS"
$Label1.AutoSize                 = $true
$Label1.width                    = 25
$Label1.height                   = 10
$Label1.location                 = New-Object System.Drawing.Point(107,54)
$Label1.Font                     = 'Microsoft Sans Serif,38,style=Bold'
$Label1.ForeColor                = "#ffffff"

$Label2                          = New-Object system.Windows.Forms.Label
$Label2.text                     = "Phone Number:"
$Label2.AutoSize                 = $true
$Label2.width                    = 25
$Label2.height                   = 10
$Label2.location                 = New-Object System.Drawing.Point(41,154)
$Label2.Font                     = 'Tahoma,17,style=Bold,Underline'
$Label2.ForeColor                = "#ffffff"

$TextBox1                        = New-Object system.Windows.Forms.TextBox
$TextBox1.multiline              = $false
$TextBox1.width                  = 188
$TextBox1.height                 = 20
$TextBox1.location               = New-Object System.Drawing.Point(245,150)
$TextBox1.Font                   = 'Microsoft Sans Serif,15'

$Label3                          = New-Object system.Windows.Forms.Label
$Label3.text                     = "Business Name:"
$Label3.AutoSize                 = $true
$Label3.width                    = 25
$Label3.height                   = 10
$Label3.location                 = New-Object System.Drawing.Point(41,306)
$Label3.Font                     = 'Tahoma,17,style=Bold,Underline'
$Label3.ForeColor                = "#ffffff"

$TextBox2                        = New-Object system.Windows.Forms.TextBox
$TextBox2.multiline              = $false
$TextBox2.width                  = 188
$TextBox2.height                 = 20
$TextBox2.location               = New-Object System.Drawing.Point(245,304)
$TextBox2.Font                   = 'Microsoft Sans Serif,15'

$Label4                          = New-Object system.Windows.Forms.Label
$Label4.text                     = "Coupon Percent:"
$Label4.AutoSize                 = $true
$Label4.width                    = 25
$Label4.height                   = 10
$Label4.location                 = New-Object System.Drawing.Point(49,360)
$Label4.Font                     = 'Tahoma,16,style=Bold,Underline'
$Label4.ForeColor                = "#ffffff"

$TextBox3                        = New-Object system.Windows.Forms.TextBox
$TextBox3.multiline              = $false
$TextBox3.width                  = 38
$TextBox3.height                 = 20
$TextBox3.location               = New-Object System.Drawing.Point(291,358)
$TextBox3.Font                   = 'Microsoft Sans Serif,15'

$Button1                         = New-Object system.Windows.Forms.Button
$Button1.BackColor               = "#4a90e2"
$Button1.text                    = "Send Default Promo SMS"
$Button1.width                   = 216
$Button1.height                  = 54
$Button1.location                = New-Object System.Drawing.Point(35,434)
$Button1.Font                    = 'Microsoft Sans Serif,13,style=Bold'
$Button1.ForeColor               = "#ffffff"

$Button2                         = New-Object system.Windows.Forms.Button
$Button2.BackColor               = "#4a90e2"
$Button2.text                    = "Send Custom SMS"
$Button2.width                   = 216
$Button2.height                  = 54
$Button2.location                = New-Object System.Drawing.Point(265,434)
$Button2.Font                    = 'Microsoft Sans Serif,13,style=Bold'
$Button2.ForeColor               = "#ffffff"

$Button3                         = New-Object system.Windows.Forms.Button
$Button3.BackColor               = "#d0021b"
$Button3.text                    = "Return Home"
$Button3.width                   = 140
$Button3.height                  = 32
$Button3.location                = New-Object System.Drawing.Point(189,578)
$Button3.Font                    = 'Microsoft Sans Serif,10,style=Bold'
$Button3.ForeColor               = "#ffffff"

$Label5                          = New-Object system.Windows.Forms.Label
$Label5.text                     = "Phone # Format -(XXXXXXXXXX)"
$Label5.AutoSize                 = $true
$Label5.width                    = 25
$Label5.height                   = 10
$Label5.location                 = New-Object System.Drawing.Point(111,192)
$Label5.Font                     = 'Myanmar Text,13,style=Bold'
$Label5.ForeColor                = "#ffffff"

$Label6                          = New-Object system.Windows.Forms.Label
$Label6.text                     = "%"
$Label6.AutoSize                 = $true
$Label6.width                    = 25
$Label6.height                   = 10
$Label6.location                 = New-Object System.Drawing.Point(339,362)
$Label6.Font                     = 'Tahoma,18,style=Bold'
$Label6.ForeColor                = "#ffffff"

$Button4                         = New-Object system.Windows.Forms.Button
$Button4.BackColor               = "#4a90e2"
$Button4.text                    = "Holiday Message"
$Button4.width                   = 216
$Button4.height                  = 54
$Button4.location                = New-Object System.Drawing.Point(35,500)
$Button4.Font                    = 'Microsoft Sans Serif,13,style=Bold'
$Button4.ForeColor               = "#ffffff"

$Button5                         = New-Object system.Windows.Forms.Button
$Button5.BackColor               = "#4a90e2"
$Button5.text                    = "Appointment Reminder"
$Button5.width                   = 216
$Button5.height                  = 54
$Button5.location                = New-Object System.Drawing.Point(265,500)
$Button5.Font                    = 'Microsoft Sans Serif,13,style=Bold'
$Button5.ForeColor               = "#ffffff"

$Label7                          = New-Object system.Windows.Forms.Label
$Label7.text                     = "For Promo Use Only"
$Label7.AutoSize                 = $true
$Label7.width                    = 25
$Label7.height                   = 10
$Label7.location                 = New-Object System.Drawing.Point(105,264)
$Label7.Font                     = 'Microsoft Sans Serif,22,style=Bold,Underline'
$Label7.ForeColor                = "#ffffff"

$Form.controls.AddRange(@($Label1,$Label2,$TextBox1,$Label3,$TextBox2,$Label4,$TextBox3,$Button1,$Button2,$Button3,$Label5,$Label6,$Button4,$Button5,$Label7))

$Button1.Add_Click({ SMS_S_Default })
$Button2.Add_Click({ SMS_S_Custom })
$Button3.Add_Click({ Gohome })
$Button4.Add_Click({ HolidaySMS })
$Button5.Add_Click({ App_Reminder })



function Gohome {
$Form.Close() }



function SMS_S_Custom { 

$title = Custom SMS notification
$msg2 = "Please Enter Your Message Below
Note- This is Direct input, What you type here is what will displayed to the recipient!"

$PhoneNumber = $TextBox1.Text 
$Bus_name = $Textbox2.Text
$Promo = $Textbox3.Text
$MessageContent = "($Bus_name)- Save Big with the Reward Members Program! Use Code NEWBIE at next checkout to recieve $Promo. off!"
$MessageContentCustom  = [Microsoft.VisualBasic.Interaction]::InputBox($msg2, $title)


$body = @{


  "phone"="$PhoneNumber"
  "message"="$MessageContentCustom"
  "key"="3ab03b66488923feeada54e25e5118781c190991RIHEVBwMxq0XoZuG8EsnZNJyl"
  "replyWebhookUrl" = "https://hook.io/ctabora10/textbelt"
}
$submit = Invoke-WebRequest -Uri https://textbelt.com/text -Body $body -Method Post

$wshell = New-Object -ComObject Wscript.Shell
$Output = $wshell.Popup("Text Message Sent!")


}



function SMS_S_Default { 








$PhoneNumber = $TextBox1.Text 
$Bus_name = $Textbox2.Text
$Promo = $Textbox3.Text
$MessageContent = "($Bus_name)- Save Big with the Reward Members Program! Use Code NEWBIE at next checkout to recieve $Promo% off!"



$body = @{


  "phone"="$PhoneNumber"
  "message"="$MessageContent"
  "key"="3ab03b66488923feeada54e25e5118781c190991RIHEVBwMxq0XoZuG8EsnZNJyl"
}
$submit = Invoke-WebRequest -Uri https://textbelt.com/text -Body $body -Method Post
$wshell = New-Object -ComObject Wscript.Shell
$Output = $wshell.Popup("Text Message Sent!")

$Form.Close()









}








[void]$Form.ShowDialog()









}



function RCMH_SMS {



[void][Reflection.Assembly]::LoadWithPartialName('Microsoft.VisualBasic')

$title = 'Ctab.Tech SMS Push Alpha'
$msg   = 'Enter the Phone number that you would like to send a text message to. NOTE - Format must be as follows- (XXXXXXXXXX):'


$msg2   = 'Please Enter Your Message Content'
$msg3   = 'Please Enter Patient Name'
$msg4   = 'Please Enter the Date of Patients appointment.

Note- Please format date as follows - (X/X/XX)'



$PhoneNumber = [Microsoft.VisualBasic.Interaction]::InputBox($msg, $title)
$PatientName = [Microsoft.VisualBasic.Interaction]::InputBox($msg3, $title)
$PatientAppDate = [Microsoft.VisualBasic.Interaction]::InputBox($msg4, $title)




Add-Type -AssemblyName PresentationCore,PresentationFramework
$ButtonType = [System.Windows.MessageBoxButton]::YesNoCancel
$MessageIcon = [System.Windows.MessageBoxImage]::Question
$MessageBody = "Would you like to use the default Patient Outreach Message?"
$MessageTitle = "Message Selection"
 
$Result = [System.Windows.MessageBox]::Show($MessageBody,$MessageTitle,$ButtonType,$MessageIcon)
if ($Result -eq "YES" )
{
$MessageContent = "(Ray County Memorial Hospital)

Hey $PatientName, this is just a friendly reminder for your appointment on $PatientAppDate. If you feel that you will be unable to make it, please contact us @ 8164705432 ext.180

-This is an automated message, please do not reply."
}
else
{
$MessageContent = $MessageContent = [Microsoft.VisualBasic.Interaction]::InputBox($msg2, $title)
}










$body = @{
  "phone"="$PhoneNumber"
  "message"="$MessageContent"
  "key"="3ab03b66488923feeada54e25e5118781c190991RIHEVBwMxq0XoZuG8EsnZNJyl"
}
$submit = Invoke-WebRequest -Uri https://textbelt.com/text -Body $body -Method Post























 }


function Check_Quota { 


$QuotaResultRAW = curl https://textbelt.com/quota/3ab03b66488923feeada54e25e5118781c190991RIHEVBwMxq0XoZuG8EsnZNJyl
$QuotaResult = "$QuotaResultRAW"
$QuotaFinal = $QuotaResult -replace "[^0-9]"
$QuotaDisplay = "You Have $QuotaFinal Remaining Text Tokens to use. "

Add-Type -AssemblyName PresentationCore,PresentationFramework
$ButtonType = [System.Windows.MessageBoxButton]::YesNoCancel
$MessageIcon = [System.Windows.MessageBoxImage]::Question
$MessageBody = "$QuotaDisplay"
$MessageTitle = "Remaining Tokens"
 
$Result = [System.Windows.MessageBox]::Show($MessageBody,$MessageTitle,$ButtonType,$MessageIcon)
if ($Result -eq "YES" )
{
[void]$Form.ShowDialog()
}
else
{


}










$body = @{
  "phone"="$PhoneNumber"
  "message"="$MessageContent"
  "key"="3ab03b66488923feeada54e25e5118781c190991RIHEVBwMxq0XoZuG8EsnZNJyl"
}
$submit = Invoke-WebRequest -Uri https://textbelt.com/text -Body $body -Method Post















}


function MassSMS {


<# This form was created using POSHGUI.com  a free online gui designer for PowerShell
.NAME
    UntitledMassSMS
#>

Add-Type -AssemblyName System.Windows.Forms
[System.Windows.Forms.Application]::EnableVisualStyles()

$Form                            = New-Object system.Windows.Forms.Form
$Form.ClientSize                 = '980,1036'
$Form.text                       = "Ctab.Tech -SMS- Autoresponder"
$Form.BackColor                  = "#4a4a4a"
$Form.TopMost                    = $false

$Label1                          = New-Object system.Windows.Forms.Label
$Label1.text                     = "Mass SMS "
$Label1.AutoSize                 = $true
$Label1.width                    = 25
$Label1.height                   = 10
$Label1.location                 = New-Object System.Drawing.Point(161,74)
$Label1.Font                     = 'Rockwell,40'
$Label1.ForeColor                = "#ffffff"

$Button1                         = New-Object system.Windows.Forms.Button
$Button1.BackColor               = "#ffffff"
$Button1.text                    = "Configure SMS Sheet"
$Button1.width                   = 248
$Button1.height                  = 60
$Button1.location                = New-Object System.Drawing.Point(113,660)
$Button1.Font                    = 'Microsoft Sans Serif,16,style=Bold'
$Button1.ForeColor               = "#000000"

$Button2                         = New-Object system.Windows.Forms.Button
$Button2.BackColor               = "#ffffff"
$Button2.text                    = "Print Current SMS Sheet"
$Button2.width                   = 248
$Button2.height                  = 60
$Button2.location                = New-Object System.Drawing.Point(113,744)
$Button2.Font                    = 'Microsoft Sans Serif,14,style=Bold'
$Button2.ForeColor               = "#000000"

$Button3                         = New-Object system.Windows.Forms.Button
$Button3.BackColor               = "#d0021b"
$Button3.text                    = "MASS SEND"
$Button3.width                   = 304
$Button3.height                  = 62
$Button3.location                = New-Object System.Drawing.Point(513,704)
$Button3.Font                    = 'Microsoft Sans Serif,30,style=Bold'
$Button3.ForeColor               = "#ffffff"

$PictureBox1                     = New-Object system.Windows.Forms.PictureBox
$PictureBox1.width               = 124
$PictureBox1.height              = 96
$PictureBox1.location            = New-Object System.Drawing.Point(29,44)
$PictureBox1.imageLocation       = "http://vault.xctab.com/siocn.png"
$PictureBox1.SizeMode            = [System.Windows.Forms.PictureBoxSizeMode]::zoom
$PictureBox2                     = New-Object system.Windows.Forms.PictureBox
$PictureBox2.width               = 124
$PictureBox2.height              = 96
$PictureBox2.location            = New-Object System.Drawing.Point(743,58)
$PictureBox2.imageLocation       = "http://vault.xctab.com/eiocn.jpg"
$PictureBox2.SizeMode            = [System.Windows.Forms.PictureBoxSizeMode]::zoom
$Label2                          = New-Object system.Windows.Forms.Label
$Label2.text                     = "Message Content "
$Label2.AutoSize                 = $true
$Label2.width                    = 25
$Label2.height                   = 10
$Label2.location                 = New-Object System.Drawing.Point(117,184)
$Label2.Font                     = 'Rockwell,24'
$Label2.ForeColor                = "#ffffff"

$TextBox1                        = New-Object system.Windows.Forms.TextBox
$TextBox1.multiline              = $true
$TextBox1.width                  = 756
$TextBox1.height                 = 406
$TextBox1.location               = New-Object System.Drawing.Point(113,230)
$TextBox1.Font                   = 'Microsoft Sans Serif,18'

$Label3                          = New-Object system.Windows.Forms.Label
$Label3.text                     = "- Or Use Templates-"
$Label3.AutoSize                 = $true
$Label3.width                    = 25
$Label3.height                   = 10
$Label3.location                 = New-Object System.Drawing.Point(515,798)
$Label3.Font                     = 'Rockwell,24'
$Label3.ForeColor                = "#ffffff"

$Button4                         = New-Object system.Windows.Forms.Button
$Button4.BackColor               = "#9b9b9b"
$Button4.text                    = "20% Off Promo"
$Button4.width                   = 100
$Button4.height                  = 100
$Button4.location                = New-Object System.Drawing.Point(465,862)
$Button4.Font                    = 'Microsoft Sans Serif,9,style=Bold'
$Button4.ForeColor               = "#000000"

$Button5                         = New-Object system.Windows.Forms.Button
$Button5.BackColor               = "#9b9b9b"
$Button5.text                    = "50% Off `"Crazy`""
$Button5.width                   = 100
$Button5.height                  = 100
$Button5.location                = New-Object System.Drawing.Point(615,862)
$Button5.Font                    = 'Microsoft Sans Serif,9,style=Bold'
$Button5.ForeColor               = "#000000"

$Button6                         = New-Object system.Windows.Forms.Button
$Button6.BackColor               = "#9b9b9b"
$Button6.text                    = "Limited Time"
$Button6.width                   = 100
$Button6.height                  = 100
$Button6.location                = New-Object System.Drawing.Point(769,862)
$Button6.Font                    = 'Microsoft Sans Serif,9,style=Bold'
$Button6.ForeColor               = "#000000"

$Form.controls.AddRange(@($Label1,$Button1,$Button2,$Button3,$PictureBox1,$PictureBox2,$Label2,$TextBox1,$Label3,$Button4,$Button5,$Button6))

$Button1.Add_Click({ Configure_Sheet })
$Button2.Add_Click({ PrintSheet })
$Button3.Add_Click({ MassSMS_Send })
$PictureBox2.Add_Click({ Kill_Excell })
$Button4.Add_Click({ SendPromo })
$Button5.Add_Click({ SendCrazy })
$Button6.Add_Click({ SendLimited })


$MassMessageContentCustom = $TextBox1.Text




function SendPromo{


<# This form was created using POSHGUI.com  a free online gui designer for PowerShell
.NAME
    Untitled20off
#>

Add-Type -AssemblyName System.Windows.Forms
[System.Windows.Forms.Application]::EnableVisualStyles()

$Form                            = New-Object system.Windows.Forms.Form
$Form.ClientSize                 = '472,428'
$Form.text                       = "20% Off Promo"
$Form.TopMost                    = $false

$Button1                         = New-Object system.Windows.Forms.Button
$Button1.BackColor               = "#4a90e2"
$Button1.text                    = "Send Mass SMS"
$Button1.width                   = 158
$Button1.height                  = 30
$Button1.location                = New-Object System.Drawing.Point(41,356)
$Button1.Font                    = 'Microsoft Sans Serif,13'
$Button1.ForeColor               = "#000000"

$Button2                         = New-Object system.Windows.Forms.Button
$Button2.BackColor               = "#4a90e2"
$Button2.text                    = "Nevermind"
$Button2.width                   = 158
$Button2.height                  = 30
$Button2.location                = New-Object System.Drawing.Point(259,356)
$Button2.Font                    = 'Microsoft Sans Serif,13'
$Button2.ForeColor               = "#000000"

$PictureBox1                     = New-Object system.Windows.Forms.PictureBox
$PictureBox1.width               = 468
$PictureBox1.height              = 348
$PictureBox1.location            = New-Object System.Drawing.Point(-1,0)
$PictureBox1.imageLocation       = "https://vault.xctab.com/20-Promo.png"
$PictureBox1.SizeMode            = [System.Windows.Forms.PictureBoxSizeMode]::zoom
$Form.controls.AddRange(@($Button1,$Button2,$PictureBox1))

$Button1.Add_Click({ Send20Promo_Send })
$Button2.Add_Click({ Send20Promo_Cancel })











[void]$Form.ShowDialog()









}




function Kill_Excell{


Get-Process excel | Stop-Process -Force

#Now we must release the $excel com object to ready it for garbage collection
[System.Runtime.Interopservices.Marshal]::ReleaseComObject($Excel)



}



function MassSMS_Send { 


#Are you Sure? Prompt



[System.Reflection.Assembly]::LoadWithPartialName("System.Windows.Forms")

$ConfirmationPrompt = 

 [System.Windows.Forms.MessageBox]::Show('Are you sure? This will push messages to all recipients within TargetList. - This Action Cannot be undone.','Send Confirmation','YesNoCancel','Error')
 if ($ConfirmationPrompt -eq 'YES')
  {

  ## Do something 





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




#Launch the SMS Custom Templates Below!

$MessageContentMass = "(MASS MESSAGE TEST)-

Hey $name, Its Saves Big to be part of the Family, Use Code FAMILYTREE your next time in and recieve 35% off your order!
 
Text STOP to be removed from the Reward Member notification list."





$MassMessageContentCustom = $TextBox1.Text





$body = @{


  "phone"="$Number"
  "message"="$MassMessageContentCustom"
  "key"="3ab03b66488923feeada54e25e5118781c190991RIHEVBwMxq0XoZuG8EsnZNJyl"
  "replyWebhookUrl" = "https://webhook.site/93667cde-ad54-413e-951b-a9e03e8e0240"
}
$submit = Invoke-WebRequest -Uri https://textbelt.com/text -Body $body -Method Post


[System.Windows.MessageBox]::Show("SMS are now in Transit, Due to Message limit constraints Text may be delayed slightly.


DELIVERY REPORT $name Successful Pushed to -Phone_Number- $Number")


}







  }




}








function PrintSheet {

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



Get-Process excel | Stop-Process -Force

#Now we must release the $excel com object to ready it for garbage collection
[System.Runtime.Interopservices.Marshal]::ReleaseComObject($Excel)





 }








 }





function Configure_Sheet {



$path = "C:\Users\$env:USERNAME\Documents\SMS_Autoresponder"
If(!(test-path $path))
{
      New-Item -ItemType Directory -Force -Path $path
}

if (!(Test-Path "C:\Users\$env:USERNAME\Documents\SMS_Autoresponder\TargetList.xlsx"))
{
   
$outputpathpath = "C:\Users\$env:USERNAME\Documents\SMS_Autoresponder\TargetList.xlsx"

$excel = New-Object -ComObject excel.application
$excel.visible = $True


$workbook = $excel.Workbooks.Add()

$uregwksht= $workbook.Worksheets.Item(1)
$uregwksht.Name = 'TargetSheet'


$row = 1
$Column = 1


#create the column headers
$uregwksht.Cells.Item(1,1) = 'Name'
$uregwksht.Cells.Item(1,2) = 'Number'
$uregwksht.Cells.Item(1,3) = 'Subscriber'

$excel.DisplayAlerts = 'False'
$ext=".xlsx"
$path="C:\Users\$env:USERNAME\Documents\SMS_Autoresponder\TargetList$ext"
$workbook.SaveAs($path) 
$workbook.Close
$excel.DisplayAlerts = 'False'
$excel.Quit()


#Check and you will see an excel process still exists after quiting
#Remove the excel process by piping it to stop-porcess
Get-Process excel | Stop-Process -Force

#Now we must release the $excel com object to ready it for garbage collection
[System.Runtime.Interopservices.Marshal]::ReleaseComObject($Excel)








  Invoke-Item "C:\Users\$env:USERNAME\Documents\SMS_Autoresponder\TargetList.xlsx"




}
else
{
  

  Invoke-Item "C:\Users\$env:USERNAME\Documents\SMS_Autoresponder\TargetList.xlsx"

}



[void]$Form.ShowDialog()

#Write your logic code here






}


[void]$Form.ShowDialog()



}



[void]$Form.ShowDialog()

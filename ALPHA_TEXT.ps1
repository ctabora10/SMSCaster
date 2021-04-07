<# This form was created using POSHGUI.com  a free online gui designer for PowerShell
.NAME
    Untitled
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
$Label1.location                 = New-Object System.Drawing.Point(129,128)
$Label1.Font                     = 'Microsoft Sans Serif,35,style=Bold'
$Label1.ForeColor                = "#ffffff"

$Label2                          = New-Object system.Windows.Forms.Label
$Label2.text                     = "Phone Number:"
$Label2.AutoSize                 = $true
$Label2.width                    = 25
$Label2.height                   = 10
$Label2.location                 = New-Object System.Drawing.Point(67,202)
$Label2.Font                     = 'Sylfaen,17,style=Underline'
$Label2.ForeColor                = "#ffffff"

$TextBox1                        = New-Object system.Windows.Forms.TextBox
$TextBox1.multiline              = $false
$TextBox1.width                  = 188
$TextBox1.height                 = 20
$TextBox1.location               = New-Object System.Drawing.Point(247,206)
$TextBox1.Font                   = 'Microsoft Sans Serif,10'

$Label3                          = New-Object system.Windows.Forms.Label
$Label3.text                     = "Business Name:"
$Label3.AutoSize                 = $true
$Label3.width                    = 25
$Label3.height                   = 10
$Label3.location                 = New-Object System.Drawing.Point(73,248)
$Label3.Font                     = 'Sylfaen,17,style=Underline'
$Label3.ForeColor                = "#ffffff"

$TextBox2                        = New-Object system.Windows.Forms.TextBox
$TextBox2.multiline              = $false
$TextBox2.width                  = 188
$TextBox2.height                 = 20
$TextBox2.location               = New-Object System.Drawing.Point(247,252)
$TextBox2.Font                   = 'Microsoft Sans Serif,10'

$Label4                          = New-Object system.Windows.Forms.Label
$Label4.text                     = "Business Name:"
$Label4.AutoSize                 = $true
$Label4.width                    = 25
$Label4.height                   = 10
$Label4.location                 = New-Object System.Drawing.Point(73,296)
$Label4.Font                     = 'Sylfaen,17,style=Underline'
$Label4.ForeColor                = "#ffffff"

$TextBox3                        = New-Object system.Windows.Forms.TextBox
$TextBox3.multiline              = $false
$TextBox3.width                  = 188
$TextBox3.height                 = 20
$TextBox3.location               = New-Object System.Drawing.Point(247,296)
$TextBox3.Font                   = 'Microsoft Sans Serif,10'

$Button1                         = New-Object system.Windows.Forms.Button
$Button1.BackColor               = "#4a90e2"
$Button1.text                    = "Send Default Promo SMS"
$Button1.width                   = 216
$Button1.height                  = 54
$Button1.location                = New-Object System.Drawing.Point(139,378)
$Button1.Font                    = 'Microsoft Sans Serif,13,style=Bold'
$Button1.ForeColor               = "#ffffff"

$Button2                         = New-Object system.Windows.Forms.Button
$Button2.BackColor               = "#4a90e2"
$Button2.text                    = "Send Custom SMS"
$Button2.width                   = 216
$Button2.height                  = 54
$Button2.location                = New-Object System.Drawing.Point(139,446)
$Button2.Font                    = 'Microsoft Sans Serif,13,style=Bold'
$Button2.ForeColor               = "#ffffff"

$Button3                         = New-Object system.Windows.Forms.Button
$Button3.BackColor               = "#d0021b"
$Button3.text                    = "Return Home"
$Button3.width                   = 140
$Button3.height                  = 32
$Button3.location                = New-Object System.Drawing.Point(175,514)
$Button3.Font                    = 'Microsoft Sans Serif,10,style=Bold'
$Button3.ForeColor               = "#ffffff"

$Form.controls.AddRange(@($Label1,$Label2,$TextBox1,$Label3,$TextBox2,$Label4,$TextBox3,$Button1,$Button2,$Button3))

$Button1.Add_Click({ SMS_S_Default })
$Button2.Add_Click({ SMS_S_Custom })
$Button3.Add_Click({ Gohome })





function Gohome { }



function SMS_S_Custom { 




}



function SMS_S_Default { 






$PhoneNumber = $TextBox1.Text 
$Bus_name = $Textbox2.Text
$Promo = $Textbox3.Text
$MessageContent = "($Bus_name)- Save Big with the Reward Members Program! Use Code NEWBIE at next checkout to recieve $Promo. off!"



$body = @{


  "phone"="$PhoneNumber"
  "message"="$MessageContent"
  "key"="3ab03b66488923feeada54e25e5118781c190991RIHEVBwMxq0XoZuG8EsnZNJyl"
}
$submit = Invoke-WebRequest -Uri https://textbelt.com/text -Body $body -Method Post


Add-Type -AssemblyName PresentationCore,PresentationFramework

$MessageIcon = [System.Windows.MessageBoxImage]::Question
$MessageBody = "Message Has been Sent. Press Yes to Exit."
$MessageTitle = "Message Delivery"









}








[void]$Form.ShowDialog()
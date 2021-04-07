<# This form was created using POSHGUI.com  a free online gui designer for PowerShell
.NAME
    Untitled
#>

Add-Type -AssemblyName System.Windows.Forms
[System.Windows.Forms.Application]::EnableVisualStyles()

$Form                            = New-Object system.Windows.Forms.Form
$Form.ClientSize                 = '772,400'
$Form.text                       = "Form"
$Form.BackColor                  = "#2b2626"
$Form.TopMost                    = $false

$TextBox1                        = New-Object system.Windows.Forms.TextBox
$TextBox1.multiline              = $false
$TextBox1.text                   = "-  SMS  - Autoresponder Alpha build v0.2"
$TextBox1.BackColor              = "#322f2f"
$TextBox1.width                  = 574
$TextBox1.height                 = 20
$TextBox1.location               = New-Object System.Drawing.Point(95,54)
$TextBox1.Font                   = 'Microsoft Sans Serif,25'
$TextBox1.ForeColor              = "#ffffff"

$Button1                         = New-Object system.Windows.Forms.Button
$Button1.BackColor               = "#98fffc"
$Button1.text                    = "Send SMS"
$Button1.width                   = 228
$Button1.height                  = 50
$Button1.location                = New-Object System.Drawing.Point(273,214)
$Button1.Font                    = 'Microsoft Sans Serif,20'
$Button1.ForeColor               = "#ffffff"

$Form.controls.AddRange(@($TextBox1,$Button1))



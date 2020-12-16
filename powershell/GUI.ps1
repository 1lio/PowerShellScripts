# Пример GUI


# Создание WindowsForms
Add-Type -assembly System.Windows.Forms
 
$main_form = New-Object System.Windows.Forms.Form
$main_form.Text ='Добавить пользователя в CiscoAnyConnect'
$main_form.Width = 400
$main_form.Height = 300
$main_form.AutoSize = $false
 
$Label = New-Object System.Windows.Forms.Label
$Label.Text = "ФИО пользователя"
$Label.Location  = New-Object System.Drawing.Point(10,10)
$Label.AutoSize = $true
$main_form.Controls.Add($Label)

$TextBox = New-Object System.Windows.Forms.TextBox
$TextBox.Location  = New-Object System.Drawing.Point(10,30)
$TextBox.Text = ''
$TextBox.Width = 360
$main_form.Controls.Add($TextBox)

#

$Label = New-Object System.Windows.Forms.Label
$Label.Text = "IP-адрес"
$Label.Location  = New-Object System.Drawing.Point(10,60)
$Label.AutoSize = $true
$main_form.Controls.Add($Label)

$TextBox = New-Object System.Windows.Forms.TextBox
$TextBox.Location  = New-Object System.Drawing.Point(10,80)
$TextBox.Text = '10.20.2.'
$TextBox.Width = 360
$main_form.Controls.Add($TextBox)

#

$Label = New-Object System.Windows.Forms.Label
$Label.Text = "Пароль"
$Label.Location  = New-Object System.Drawing.Point(10,110)
$Label.AutoSize = $true
$main_form.Controls.Add($Label)

$TextBox = New-Object System.Windows.Forms.TextBox
$TextBox.Location  = New-Object System.Drawing.Point(10,130)
$TextBox.Text = ''
$TextBox.Width = 360
$main_form.Controls.Add($TextBox)

#
 
$button = New-Object System.Windows.Forms.Button
$button.Text = 'Создать'
$button.Location = New-Object System.Drawing.Point(150,230)
$main_form.Controls.Add($button)


  
$main_form.ShowDialog()
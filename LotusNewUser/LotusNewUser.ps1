#####################
# LotusNewUser
# Script for generating CSV import file
# Generated file can be imported in Lotus Domino Administrator to register new user https://www.ibm.com/support/knowledgecenter/en/SSKTMJ_8.0.1/com.ibm.help.domino.admin.doc/DOC/H_ID_INFORMATION.html
# Example content of file: 
# Alexis;Catherine;;;password1;;;Marketing / Acme;;;;;;Marketing Profile 
#
# Requirements: Powershell 2.0; Remote Server Administration Tools to retrieve data from AD 
# Kosarev Albert, 2016

# Check thread is STA
if ([System.Threading.Thread]::CurrentThread.ApartmentState -eq [System.Threading.ApartmentState]::MTA) {
    powershell.exe -Sta -File $MyInvocation.MyCommand.Path
    return
}

$global:IDFolder = "E:\Lotus\Id"
$global:MailSRV = "Your mail server1"
$global:MailTemplate = ";mail85ru.ntf"
$global:CanRenewDB = $False
$global:FDBContentLoaded = $False
$global:Encoding = [System.Text.Encoding]::Default

# Departments array for cbDept
$global:ArrayDepts = @("Logistics;Логистика", "Sklad;Склад", "Guard;Охрана", 
				"Sale;Продажи", "Account;Бухгалтерия", "IT;ИТ", "HR;Персонал")

# check AD connection
if (Get-Module -Name ActiveDirectory -ListAvailable -EA 0) {
    Import-Module ActiveDirectory
	$global:ConnectAD = 0
} else {
	$global:ConnectAD = 2	
}

# Get settings from Registry
function GetRegistrySettings {
	$global:RegPath = "HKCU:\Software\LotusNewUser"
	if (!(Test-Path $global:RegPath)) {
			New-Item -Path HKCU:\Software -Name LotusNewUser | Out-Null
	} else {
		$global:FileDBLocation = (Get-ItemProperty -Path $global:RegPath -Name FileDB -EA 0).FileDB
		$global:FileImportLocation = (Get-ItemProperty -Path $global:RegPath -Name FileImport -EA 0).FileImport
	}
	$global:FileDBCheck = ((-not ([string]::IsNullOrEmpty($global:FileDBLocation))) -and (Test-Path $global:FileDBLocation)) 
	$global:FileImportCheck = ((-not ([string]::IsNullOrEmpty($global:FileImportLocation))) -and (Test-Path $global:FileImportLocation)) 
}
GetRegistrySettings

function WriteLog ($LogType, $LogMsg) {
	switch ($LogType) {
		Success {
			$rtbLog.SelectionColor = [Drawing.Color]::"Green"
			$rtbLog.AppendText($LogMsg)
		}
		Error {
			$rtbLog.SelectionColor = [Drawing.Color]::"Red"
			$rtbLog.AppendText($LogMsg)
		}
		Info {
			$rtbLog.SelectionColor = [Drawing.Color]::"Black"
			$rtbLog.AppendText($LogMsg)
		}
	}
	$rtbLog.SelectionStart = $rtbLog.Text.Length
	$rtbLog.ScrollToCaret()
}

function CheckADUser { 
	if ($global:ConnectAD -eq 0) {
		$DCName = (Get-ADDomainController -EA 0).Hostname
		if (-not ([string]::IsNullOrEmpty($DCName))) {
			WriteLog -LogType Success -LogMsg ("Подключение к "+(Get-ADDomainController).Hostname+"`tOK`n")
			$global:ConnectAD = 1
		} else {
			WriteLog -LogType Error -LogMsg ("Подключиться к домену не удалось`n")
			$global:ConnectAD = 2
		}
	} 
	if (-not ([string]::IsNullOrEmpty($tbUserAD.Text))) {
		$ADUsers = @(Get-ADUser -Filter {samAccountName -like $tbUserAD.Text} -Properties * -EA 0)
		$global:ADUser = $ADUsers[0] 
	} else {
		return
	}
	ForEach-Object { 
		$DN = $global:ADUser.DistinguishedName -split "," 
		$UserOU = $DN[$DN.Count - 4] -replace "OU="
	}

	if ($global:ADUser -ne $Null) {
		WriteLog -LogType Success -LogMsg ("Пользователь "+$global:ADUser.SamAccountName+"  найден`n")
		WriteLog -LogType Info -LogMsg ("Фамилия Имя: "+$global:ADUser.Givenname+" "+$global:ADUser.Surname+", Email: "+$global:ADUser.Mail+" подразделение: "+$global:UserDept+"`n")
	} else {
		WriteLog -LogType Error -LogMsg ("Ошибка: пользователь "+$tbUserAD.Text+" не найден`n")
	}
	#
	$tbNameRus.Text = ($global:ADUser.Surname+" "+$global:ADUser.Givenname)
	
	# Translit Name and Surname and write to textbox
	Translit
	
	$tbEmailAD.Text = $global:ADUser.Mail
	
	# reset password box
	$tbPassword.Text = ""
	
	foreach ($DeptEntry in $global:ArrayDepts) {
		if (([regex]::Match($DeptEntry,$global:UserDept)).Success) {
			$cbDept.Text = $DeptEntry
		}

	}
	
	# Set dept by AD user description
	if (([regex]::Match($global:ADUser.Description,"(?i)кадр+|персонал+")).Success) {
			$cbDept.Text = $global:ArrayDepts[[array]::IndexOf($global:ArrayDepts,"HR;Персонал")]
	}
	if (([regex]::Match($global:ADUser.Description,"(?i)бух+")).Success) {
			$cbDept.Text = $global:ArrayDepts[[array]::IndexOf($global:ArrayDepts,"Account;Бухгалтерия")]
	}
}

# Translit rules from GOST 1997-2010
function Translit{
	#param([string]$InputString)
	$TranslitTable = @{ 
	[char]'а' = "a";[char]'А' = "A";
	[char]'б' = "b";[char]'Б' = "B";
	[char]'в' = "v";[char]'В' = "V";
	[char]'г' = "g";[char]'Г' = "G";
	[char]'д' = "d";[char]'Д' = "D";
	[char]'е' = "e";[char]'Е' = "E";
	[char]'ё' = "ye";[char]'Ё' = "Ye";
	[char]'ж' = "zh";[char]'Ж' = "Zh";
	[char]'з' = "z";[char]'З' = "Z";
	[char]'и' = "i";[char]'И' = "I";
	[char]'й' = "y";[char]'Й' = "Y";
	[char]'к' = "k";[char]'К' = "K";
	[char]'л' = "l";[char]'Л' = "L";
	[char]'м' = "m";[char]'М' = "M";
	[char]'н' = "n";[char]'Н' = "N";
	[char]'о' = "o";[char]'О' = "O";
	[char]'п' = "p";[char]'П' = "P";
	[char]'р' = "r";[char]'Р' = "R";
	[char]'с' = "s";[char]'С' = "S";
	[char]'т' = "t";[char]'Т' = "T";
	[char]'у' = "u";[char]'У' = "U";
	[char]'ф' = "f";[char]'Ф' = "F";
	[char]'х' = "kh";[char]'Х' = "Kh";
	[char]'ц' = "ts";[char]'Ц' = "Ts";
	[char]'ч' = "ch";[char]'Ч' = "Ch";
	[char]'ш' = "sh";[char]'Ш' = "Sh";
	[char]'щ' = "sch";[char]'Щ' = "Shch";
	[char]'ъ' = "";[char]'Ъ' = "";
	[char]'ы' = "y";[char]'Ы' = "Y";
	[char]'ь' = "";[char]'Ь' = "";
	[char]'э' = "e";[char]'Э' = "E";
	[char]'ю' = "yu";[char]'Ю' = "Yu";
	[char]'я' = "ya";[char]'Я' = "Ya"
	}
	$outChars = ""
	$InputString = ($tbNameRus.Text).ToCharArray()
	foreach ($ch in $InputString) {
		if ($TranslitTable[$ch] -cne $Null ) {
			$outChars += $TranslitTable[$ch]
		}
		else {
			$outChars += $ch
		}
	}
	$tbNameEnglish.Text = $outChars
 }


function GenerateForm {
#region Import the Assemblies
[reflection.assembly]::loadwithpartialname("System.Windows.Forms") | Out-Null
[reflection.assembly]::loadwithpartialname("System.Drawing") | Out-Null
## visual styles
[System.Windows.Forms.Application]::EnableVisualStyles()
#endregion

#region Generated Form Objects
$form1 = New-Object System.Windows.Forms.Form
$tabControl1 = New-Object System.Windows.Forms.TabControl
$tabPage1 = New-Object System.Windows.Forms.TabPage
$labelNameRus = New-Object System.Windows.Forms.Label
$tbNameRus = New-Object System.Windows.Forms.TextBox
$labelRenewDB = New-Object System.Windows.Forms.Label
$labelCreateFileSuccess = New-Object System.Windows.Forms.Label
$btnRenewDB = New-Object System.Windows.Forms.Button
$tbPassword = New-Object System.Windows.Forms.TextBox
$labelPassword = New-Object System.Windows.Forms.Label
$btnCheckAD = New-Object System.Windows.Forms.Button
$btnCreateFile = New-Object System.Windows.Forms.Button
$tbEmailAD = New-Object System.Windows.Forms.TextBox
$tbNameEnglish = New-Object System.Windows.Forms.TextBox
$labelDept = New-Object System.Windows.Forms.Label
$labelEmailAD = New-Object System.Windows.Forms.Label
$labelNameEnglish = New-Object System.Windows.Forms.Label
$cbDept = New-Object System.Windows.Forms.ComboBox
$chbAppend = New-Object System.Windows.Forms.CheckBox
$tbUserAD = New-Object System.Windows.Forms.TextBox
$labelUserAD = New-Object System.Windows.Forms.Label
$rtbLog = New-Object System.Windows.Forms.RichTextBox
$tabPage2 = New-Object System.Windows.Forms.TabPage
$rtbSearchResult = New-Object System.Windows.Forms.RichTextBox
$textBoxSearchUser = New-Object System.Windows.Forms.TextBox
$labelSearchUser = New-Object System.Windows.Forms.Label
$tabPage3 = New-Object System.Windows.Forms.TabPage
$tbFileImport = New-Object System.Windows.Forms.TextBox
$labelFileImportSuccess = New-Object System.Windows.Forms.Label
$labelFileDBSuccess = New-Object System.Windows.Forms.Label
$tbFileDB = New-Object System.Windows.Forms.TextBox
$btnFileImportBrowse = New-Object System.Windows.Forms.Button
$btnFileDBBrowse = New-Object System.Windows.Forms.Button
$labelAbout = New-Object System.Windows.Forms.Label
$labelFileImport = New-Object System.Windows.Forms.Label
$labelFileDB = New-Object System.Windows.Forms.Label
$ofdDBFile = New-Object System.Windows.Forms.OpenFileDialog
$ofdImportFile = New-Object System.Windows.Forms.OpenFileDialog
$tooltip_btnCheckAD = New-Object System.Windows.Forms.ToolTip
$tooltip_cbDept = New-Object System.Windows.Forms.ToolTip

$InitialFormWindowState = New-Object System.Windows.Forms.FormWindowState
#endregion Generated Form Objects


$tbUserAD_KeyUp = [System.Windows.Forms.KeyEventHandler] {
    if ($_.KeyCode -eq 'Enter') {
		CheckADUser
	}
}

$textboxSearchUser_KeyUp = [System.Windows.Forms.KeyEventHandler] {
    if ($_.KeyCode -eq 'Enter') {
		Searching
	}
}

$OnLoadForm_StateCorrection=
{#Correct the initial state of the form to prevent the .Net maximized form issue
	$form1.WindowState = $InitialFormWindowState
}

#----------------------------------------------
#region Generated Form Code
$System_Drawing_Size = New-Object System.Drawing.Size
$System_Drawing_Size.Height = 500
$System_Drawing_Size.Width = 475
$form1.ClientSize = $System_Drawing_Size
$form1.DataBindings.DefaultDataSourceUpdateMode = 0
$form1.MaximizeBox = $False
$form1.StartPosition = "CenterScreen"
$form1.Name = "form1"
$form1.Text = "Создание нового пользователя Lotus"
$form1.add_Load($handler_form1_Load)
$form1.FormBorderStyle = "FixedSingle"

# icon: http://favicon.cc --> .ico file --> base64
$iconstring = (@"
AAABAAEAEBAQAAEABAAoAQAAFgAAACgAAAAQAAAAIAAAAAEABAAAAAAAgAAAAAAAAAAAAAAAEAAAAAAAAAAAAAAAAHv/AJ2msAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAIiIiIAAAAAAhERERAAAAACEREREAAAAAIREAAAAAAAAhEQAAAAAAACERAAAAAAAAIREAAAAAAAAhEQAAAAAAACERAAAAAAAAIREAAAAAAAAhEQAAAAAAACERAAAAAAAAAREAAAAAAAAAAAAAAAAAAAAAAAAAD//wAA+A8AAPgHAAD4BwAA+H8AAPh/AAD4fwAA+H8AAPh/AAD4fwAA+H8AAPh/AAD4fwAA/H8AAP//AAD//wAA
"@).replace("`n", "")
$memory = new-object System.IO.MemoryStream
$memory.write(($bytes=[System.Convert]::FromBase64String($iconstring)), 0, $bytes.length)
$form1.Icon = [System.Drawing.Icon]::FromHandle((new-object System.Drawing.Bitmap -argument $memory).GetHIcon())

$tabControl1.Anchor = 15
$tabControl1.DataBindings.DefaultDataSourceUpdateMode = 0
$System_Drawing_Point = New-Object System.Drawing.Point
$System_Drawing_Point.X = 0
$System_Drawing_Point.Y = 0
$tabControl1.Location = $System_Drawing_Point
$tabControl1.Name = "tabControl1"
$tabControl1.SelectedIndex = 0
$System_Drawing_Size = New-Object System.Drawing.Size
$System_Drawing_Size.Height = 500
$System_Drawing_Size.Width = 475
$tabControl1.Size = $System_Drawing_Size
$form1.Controls.Add($tabControl1)

#$tabPage1.BackColor = [System.Drawing.Color]::FromArgb(255,240,240,240)
$tabPage1.DataBindings.DefaultDataSourceUpdateMode = 0
$System_Drawing_Point = New-Object System.Drawing.Point
$System_Drawing_Point.X = 3
$System_Drawing_Point.Y = 22
$tabPage1.Location = $System_Drawing_Point
$tabPage1.Name = "tabPage1"
$System_Windows_Forms_Padding = New-Object System.Windows.Forms.Padding
$System_Windows_Forms_Padding.All = 3
$System_Windows_Forms_Padding.Bottom = 3
$System_Windows_Forms_Padding.Left = 3
$System_Windows_Forms_Padding.Right = 3
$System_Windows_Forms_Padding.Top = 3
$tabPage1.Padding = $System_Windows_Forms_Padding
$System_Drawing_Size = New-Object System.Drawing.Size
$System_Drawing_Size.Height = 475
$System_Drawing_Size.Width = 465
$tabPage1.Size = $System_Drawing_Size
$tabPage1.Text = "Создание"
$tabPage1.UseVisualStyleBackColor = $True
$tabControl1.Controls.Add($tabPage1)

$tbUserAD.DataBindings.DefaultDataSourceUpdateMode = 0
$System_Drawing_Point = New-Object System.Drawing.Point
$System_Drawing_Point.X = 219
$System_Drawing_Point.Y = 22
$tbUserAD.Location = $System_Drawing_Point
$tbUserAD.Name = "tbUserAD"
$System_Drawing_Size = New-Object System.Drawing.Size
$System_Drawing_Size.Height = 20
$System_Drawing_Size.Width = 171
$tbUserAD.Size = $System_Drawing_Size
$tbUserAD.TabIndex = 0
$tbUserAD.MaxLength = 18
$tbUserAD.BorderStyle = "FixedSingle"
$tbUserAD.add_KeyUp($tbUserAD_KeyUp)
$tabPage1.Controls.Add($tbUserAD)
# Disable button if there is no AD connection
if ($global:ConnectAD -eq 2) {
	$tbUserAD.Enabled = $False
	$tbUserAD.Text = "Вводите данные вручную"
}

$tbNameRus.DataBindings.DefaultDataSourceUpdateMode = 0
$tbNameRus.BackColor = "Window"
$System_Drawing_Point = New-Object System.Drawing.Point
$System_Drawing_Point.X = 219
$System_Drawing_Point.Y = 86
$tbNameRus.Location = $System_Drawing_Point
$tbNameRus.Name = "tbNameRus"
$System_Drawing_Size = New-Object System.Drawing.Size
$System_Drawing_Size.Height = 20
$System_Drawing_Size.Width = 171
$tbNameRus.Size = $System_Drawing_Size
$tbNameRus.TabIndex = 2
$tbNameRus.MaxLength = 50
$tbNameRus.BorderStyle = "FixedSingle"
$tbNameRus.add_TextChanged({Translit})
$tabPage1.Controls.Add($tbNameRus)

$labelNameRus.DataBindings.DefaultDataSourceUpdateMode = 0
$System_Drawing_Point = New-Object System.Drawing.Point
$System_Drawing_Point.X = 22
$System_Drawing_Point.Y = 86
$labelNameRus.Location = $System_Drawing_Point
$labelNameRus.Name = "labelNameRus"
$System_Drawing_Size = New-Object System.Drawing.Size
$System_Drawing_Size.Height = 23
$System_Drawing_Size.Width = 172
$labelNameRus.Size = $System_Drawing_Size
$labelNameRus.Text = "Фамилия и имя"
$labelNameRus.TextAlign = 64
$tabPage1.Controls.Add($labelNameRus)


$labelRenewDB.DataBindings.DefaultDataSourceUpdateMode = 0
$labelRenewDB.ForeColor = "Green"
$System_Drawing_Point = New-Object System.Drawing.Point
$System_Drawing_Point.X = 356
$System_Drawing_Point.Y = 334
$labelRenewDB.Location = $System_Drawing_Point
$labelRenewDB.Name = "labelRenewDB"
$System_Drawing_Size = New-Object System.Drawing.Size
$System_Drawing_Size.Height = 23
$System_Drawing_Size.Width = 71
$labelRenewDB.Size = $System_Drawing_Size
$labelRenewDB.Text = "Успех"
$labelRenewDB.TextAlign = 16
$labelRenewDB.Visible = $False
$tabPage1.Controls.Add($labelRenewDB)

$labelCreateFileSuccess.DataBindings.DefaultDataSourceUpdateMode = 0
$labelCreateFileSuccess.ForeColor = "Green"
$System_Drawing_Point = New-Object System.Drawing.Point
$System_Drawing_Point.X = 356
$System_Drawing_Point.Y = 297
$labelCreateFileSuccess.Location = $System_Drawing_Point
$labelCreateFileSuccess.Name = "labelCreateFileSuccess"
$System_Drawing_Size = New-Object System.Drawing.Size
$System_Drawing_Size.Height = 24
$System_Drawing_Size.Width = 57
$labelCreateFileSuccess.Size = $System_Drawing_Size
$labelCreateFileSuccess.Text = "Успех"
$labelCreateFileSuccess.TextAlign = 16
$labelCreateFileSuccess.Visible = $False
$tabPage1.Controls.Add($labelCreateFileSuccess)


$btnRenewDB.DataBindings.DefaultDataSourceUpdateMode = 0
$System_Drawing_Point = New-Object System.Drawing.Point
$System_Drawing_Point.X = 132
$System_Drawing_Point.Y = 334
$btnRenewDB.Location = $System_Drawing_Point
$btnRenewDB.Name = "btnRenewDB"
$System_Drawing_Size = New-Object System.Drawing.Size
$System_Drawing_Size.Height = 23
$System_Drawing_Size.Width = 201
$btnRenewDB.Size = $System_Drawing_Size
$btnRenewDB.TabIndex = 8
$btnRenewDB.Text = "Добавить пользователя в БД"
$btnRenewDB.UseVisualStyleBackColor = $True
$btnRenewDB.add_Click({FileDBRenew})
$tabPage1.Controls.Add($btnRenewDB)

$tbPassword.DataBindings.DefaultDataSourceUpdateMode = 0
$System_Drawing_Point = New-Object System.Drawing.Point
$System_Drawing_Point.X = 219
$System_Drawing_Point.Y = 206
$tbPassword.Location = $System_Drawing_Point
$tbPassword.Name = "tbPassword"
$System_Drawing_Size = New-Object System.Drawing.Size
$System_Drawing_Size.Height = 20
$System_Drawing_Size.Width = 171
$tbPassword.Size = $System_Drawing_Size
$tbPassword.TabIndex = 5
$tbPassword.MaxLength = 16
$tbPassword.BorderStyle = "FixedSingle"
$tabPage1.Controls.Add($tbPassword)

$labelPassword.DataBindings.DefaultDataSourceUpdateMode = 0
$System_Drawing_Point = New-Object System.Drawing.Point
$System_Drawing_Point.X = 94
$System_Drawing_Point.Y = 206
$labelPassword.Location = $System_Drawing_Point
$labelPassword.Name = "labelPassword"
$System_Drawing_Size = New-Object System.Drawing.Size
$System_Drawing_Size.Height = 23
$System_Drawing_Size.Width = 100
$labelPassword.Size = $System_Drawing_Size
$labelPassword.Text = "Пароль"
$labelPassword.TextAlign = 64
$tabPage1.Controls.Add($labelPassword)

$btnCheckAD.DataBindings.DefaultDataSourceUpdateMode = 0
$System_Drawing_Point = New-Object System.Drawing.Point
$System_Drawing_Point.X = 219
$System_Drawing_Point.Y = 48
$btnCheckAD.Location = $System_Drawing_Point
$btnCheckAD.Name = "btnCheckAD"
$System_Drawing_Size = New-Object System.Drawing.Size
$System_Drawing_Size.Height = 23
$System_Drawing_Size.Width = 103
$btnCheckAD.Size = $System_Drawing_Size
$btnCheckAD.TabIndex = 1
$btnCheckAD.Text = "Проверить в AD"
$btnCheckAD.UseVisualStyleBackColor = $True
$btnCheckAD.add_Click({ if (-not ([string]::IsNullOrEmpty($tbUserAD.Text))) {CheckADUser} else 	{	WriteLog -LogType Error -LogMsg "Ошибка: введите имя пользователя`n"}})
$tabPage1.Controls.Add($btnCheckAD)
# Disable button if there is no AD connection
if ($global:ConnectAD -eq 2) {
	$btnCheckAD.Enabled = $False
}

$btnCreateFile.DataBindings.DefaultDataSourceUpdateMode = 0
$System_Drawing_Point = New-Object System.Drawing.Point
$System_Drawing_Point.X = 132
$System_Drawing_Point.Y = 298
$btnCreateFile.Location = $System_Drawing_Point
$btnCreateFile.Name = "btnCreateFile"
$System_Drawing_Size = New-Object System.Drawing.Size
$System_Drawing_Size.Height = 23
$System_Drawing_Size.Width = 201
$btnCreateFile.Size = $System_Drawing_Size
$btnCreateFile.TabIndex = 7
$btnCreateFile.Text = "Создать файл импорта Lotus"
$btnCreateFile.UseVisualStyleBackColor = $True
$btnCreateFile.add_Click({FileImportCreate})
$tabPage1.Controls.Add($btnCreateFile)

$tbEmailAD.DataBindings.DefaultDataSourceUpdateMode = 0
$System_Drawing_Point = New-Object System.Drawing.Point
$System_Drawing_Point.X = 219
$System_Drawing_Point.Y = 166
$tbEmailAD.Location = $System_Drawing_Point
$tbEmailAD.Name = "tbEmailAD"
$System_Drawing_Size = New-Object System.Drawing.Size
$System_Drawing_Size.Height = 20
$System_Drawing_Size.Width = 171
$tbEmailAD.Size = $System_Drawing_Size
$tbEmailAD.TabIndex = 4
$tbEmailAD.MaxLength = 30
$tbEmailAD.BorderStyle = "FixedSingle"
$tabPage1.Controls.Add($tbEmailAD)

$tbNameEnglish.DataBindings.DefaultDataSourceUpdateMode = 0
$System_Drawing_Point = New-Object System.Drawing.Point
$System_Drawing_Point.X = 219
$System_Drawing_Point.Y = 126
$tbNameEnglish.Location = $System_Drawing_Point
$tbNameEnglish.Name = "tbNameEnglish"
$System_Drawing_Size = New-Object System.Drawing.Size
$System_Drawing_Size.Height = 20
$System_Drawing_Size.Width = 171
$tbNameEnglish.Size = $System_Drawing_Size
$tbNameEnglish.TabIndex = 3
$tbNameEnglish.MaxLength = 50
$tbNameEnglish.BorderStyle = "FixedSingle"
$tabPage1.Controls.Add($tbNameEnglish)

$labelDept.DataBindings.DefaultDataSourceUpdateMode = 0
$System_Drawing_Point = New-Object System.Drawing.Point
$System_Drawing_Point.X = 22
$System_Drawing_Point.Y = 246
$labelDept.Location = $System_Drawing_Point
$labelDept.Name = "labelDept"
$System_Drawing_Size = New-Object System.Drawing.Size
$System_Drawing_Size.Height = 23
$System_Drawing_Size.Width = 172
$labelDept.Size = $System_Drawing_Size
$labelDept.Text = "Подразделение"
$labelDept.TextAlign = 64
$tabPage1.Controls.Add($labelDept)

$labelEmailAD.DataBindings.DefaultDataSourceUpdateMode = 0
$System_Drawing_Point = New-Object System.Drawing.Point
$System_Drawing_Point.X = 22
$System_Drawing_Point.Y = 166
$labelEmailAD.Location = $System_Drawing_Point
$labelEmailAD.Name = "labelEmailAD"
$System_Drawing_Size = New-Object System.Drawing.Size
$System_Drawing_Size.Height = 23
$System_Drawing_Size.Width = 172
$labelEmailAD.Size = $System_Drawing_Size
$labelEmailAD.Text = "Email адрес"
$labelEmailAD.TextAlign = 64
$tabPage1.Controls.Add($labelEmailAD)

$labelNameEnglish.DataBindings.DefaultDataSourceUpdateMode = 0
$System_Drawing_Point = New-Object System.Drawing.Point
$System_Drawing_Point.X = 22
$System_Drawing_Point.Y = 126
$labelNameEnglish.Location = $System_Drawing_Point
$labelNameEnglish.Name = "labelNameEnglish"
$System_Drawing_Size = New-Object System.Drawing.Size
$System_Drawing_Size.Height = 23
$System_Drawing_Size.Width = 172
$labelNameEnglish.Size = $System_Drawing_Size
$labelNameEnglish.Text = "Фамилия и имя по-английски"
$labelNameEnglish.TextAlign = 64
$tabPage1.Controls.Add($labelNameEnglish)

$cbDept.DataBindings.DefaultDataSourceUpdateMode = 0
$cbDept.FormattingEnabled = $True
$System_Drawing_Point = New-Object System.Drawing.Point
$System_Drawing_Point.X = 219
$System_Drawing_Point.Y = 246
$cbDept.Location = $System_Drawing_Point
$cbDept.Name = "cbDept"
$System_Drawing_Size = New-Object System.Drawing.Size
$System_Drawing_Size.Height = 21
$System_Drawing_Size.Width = 171
$cbDept.Size = $System_Drawing_Size
$cbDept.TabIndex = 6
$cbDept.MaxLength = 50
$cbDept.MaxDropDownItems = 8
$cbDept.DropDownHeight = 200
$cbDept.Sorted = $True
$cbDept.FormattingEnabled = $True
$cbDept.ImeMode = 1
$cbDept.Items.AddRange($global:ArrayDepts)

# AutoCompletion 
$cbDept.AutoCompleteCustomSource.AddRange($global:ArrayDepts)
$cbDept.AutoCompleteMode = [System.Windows.Forms.AutoCompleteMode]::SuggestAppend;
$cbDept.AutoCompleteSource = [System.Windows.Forms.AutoCompleteSource]::CustomSource;
$tabPage1.Controls.Add($cbDept)

$chbAppend.DataBindings.DefaultDataSourceUpdateMode = 0
$System_Drawing_Point = New-Object System.Drawing.Point
$System_Drawing_Point.X = 11
$System_Drawing_Point.Y = 294
$chbAppend.Location = $System_Drawing_Point
$chbAppend.Name = "chbAppend"
$System_Drawing_Size = New-Object System.Drawing.Size
$System_Drawing_Size.Height = 34
$System_Drawing_Size.Width = 115
$chbAppend.Size = $System_Drawing_Size
$chbAppend.TabIndex = 19
$chbAppend.Text = "Дозапись файла"
$chbAppend.UseVisualStyleBackColor = $True
$chbAppend.add_CheckedChanged({FIModeChange})
$tabPage1.Controls.Add($chbAppend)

$labelUserAD.DataBindings.DefaultDataSourceUpdateMode = 0
$System_Drawing_Point = New-Object System.Drawing.Point
$System_Drawing_Point.X = 22
$System_Drawing_Point.Y = 21
$labelUserAD.Location = $System_Drawing_Point
$labelUserAD.Name = "labelUserAD"
$System_Drawing_Size = New-Object System.Drawing.Size
$System_Drawing_Size.Height = 20
$System_Drawing_Size.Width = 172
$labelUserAD.Size = $System_Drawing_Size
$labelUserAD.Text = "Учетная запись в AD"
$labelUserAD.TextAlign = 64
$tabPage1.Controls.Add($labelUserAD)

$rtbLog.Anchor = 15
$rtbLog.BackColor = [System.Drawing.Color]::FromArgb(255,255,255,225)
$rtbLog.BorderStyle = 1
$rtbLog.DataBindings.DefaultDataSourceUpdateMode = 0
$System_Drawing_Point = New-Object System.Drawing.Point
$System_Drawing_Point.X = 0
$System_Drawing_Point.Y = 373
$rtbLog.Location = $System_Drawing_Point
$rtbLog.MaxLength = 214748
$rtbLog.Name = "rtbLog"
$rtbLog.ReadOnly = $True
$rtbLog.ScrollBars = 2
$System_Drawing_Size = New-Object System.Drawing.Size
$System_Drawing_Size.Height = 102
$System_Drawing_Size.Width = 465
$rtbLog.Size = $System_Drawing_Size
$rtbLog.BorderStyle = "None"
$tabPage1.Controls.Add($rtbLog)
	if ($global:ConnectAD -eq 2) {
		WriteLog -LogType Error -LogMsg ("Подключения к домену нет. Установите RSAT Tools`n")
	}


###### Tab 3
#$tabPage3.BackColor = [System.Drawing.Color]::FromArgb(255,240,240,240)
$tabPage3.DataBindings.DefaultDataSourceUpdateMode = 0
$System_Drawing_Point = New-Object System.Drawing.Point
$System_Drawing_Point.X = 0
$System_Drawing_Point.Y = 22
$tabPage3.Location = $System_Drawing_Point
$tabPage3.Name = "tabPage3"
$System_Windows_Forms_Padding = New-Object System.Windows.Forms.Padding
$System_Windows_Forms_Padding.All = 3
$System_Windows_Forms_Padding.Bottom = 3
$System_Windows_Forms_Padding.Left = 3
$System_Windows_Forms_Padding.Right = 3
$System_Windows_Forms_Padding.Top = 3
$tabPage3.Padding = $System_Windows_Forms_Padding
$System_Drawing_Size = New-Object System.Drawing.Size
$System_Drawing_Size.Height = 475
$System_Drawing_Size.Width = 465
$tabPage3.Size = $System_Drawing_Size
$tabPage3.TabIndex = 2
$tabPage3.Text = "Настройки"
$tabPage3.add_Click($handler_TabPage:_Click)
$tabPage3.UseVisualStyleBackColor = $True
$tabControl1.Controls.Add($tabPage3)

$tbFileImport.DataBindings.DefaultDataSourceUpdateMode = 0
$System_Drawing_Point = New-Object System.Drawing.Point
$System_Drawing_Point.X = 16
$System_Drawing_Point.Y = 155
$tbFileImport.Location = $System_Drawing_Point
$tbFileImport.Name = "tbFileImport"
$System_Drawing_Size = New-Object System.Drawing.Size
$System_Drawing_Size.Height = 20
$System_Drawing_Size.Width = 269
$tbFileImport.Size = $System_Drawing_Size
$tbFileImport.TabIndex = 8
$tbFileImport.BorderStyle = "FixedSingle"
$tbFileImport.ReadOnly = $True
$tbFileImport.BackColor = "White"
$tabPage3.Controls.Add($tbFileImport)
if ($global:FileImportCheck) { 
	$tbFileImport.Text = $global:FileImportLocation
	WriteLog -LogType Success -LogMsg ("Файл импорта успешно загружен "+$global:FileImportLocation+"`n")
} else { 
	WriteLog -LogType Error -LogMsg ("Файл импорта не найден`n")
}


$labelFileImportSuccess.DataBindings.DefaultDataSourceUpdateMode = 0
$labelFileImportSuccess.ForeColor = "Green"
$System_Drawing_Point = New-Object System.Drawing.Point
$System_Drawing_Point.X = 310
$System_Drawing_Point.Y = 152
$labelFileImportSuccess.Location = $System_Drawing_Point
$labelFileImportSuccess.Name = "labelFileImportSuccess"
$System_Drawing_Size = New-Object System.Drawing.Size
$System_Drawing_Size.Height = 23
$System_Drawing_Size.Width = 100
$labelFileImportSuccess.Size = $System_Drawing_Size
$labelFileImportSuccess.Text = "ОК"
$labelFileImportSuccess.TextAlign = 16
$labelFileImportSuccess.Visible = $False
$tabPage3.Controls.Add($labelFileImportSuccess)


$labelFileDBSuccess.DataBindings.DefaultDataSourceUpdateMode = 0
$labelFileDBSuccess.ForeColor = [System.Drawing.Color]::FromArgb(255,0,128,0)
$System_Drawing_Point = New-Object System.Drawing.Point
$System_Drawing_Point.X = 310
$System_Drawing_Point.Y = 67
$labelFileDBSuccess.Location = $System_Drawing_Point
$labelFileDBSuccess.Name = "labelFileDBSuccess"
$System_Drawing_Size = New-Object System.Drawing.Size
$System_Drawing_Size.Height = 23
$System_Drawing_Size.Width = 100
$labelFileDBSuccess.Size = $System_Drawing_Size
$labelFileDBSuccess.Text = "ОК"
$labelFileDBSuccess.TextAlign = 16
$labelFileDBSuccess.Visible = $False
$labelFileDBSuccess.add_Click($handler_label8_Click)
$tabPage3.Controls.Add($labelFileDBSuccess)

$tbFileDB.DataBindings.DefaultDataSourceUpdateMode = 0
$System_Drawing_Point = New-Object System.Drawing.Point
$System_Drawing_Point.X = 16
$System_Drawing_Point.Y = 67
$tbFileDB.Location = $System_Drawing_Point
$tbFileDB.Name = "tbFileDB"
$System_Drawing_Size = New-Object System.Drawing.Size
$System_Drawing_Size.Height = 20
$System_Drawing_Size.Width = 269
$tbFileDB.Size = $System_Drawing_Size
$tbFileDB.TabIndex = 5
$tbFileDB.BorderStyle = "FixedSingle"
$tbFileDB.ReadOnly = $True
$tbFileDB.BackColor = "White"
$tabPage3.Controls.Add($tbFileDB)
if ($global:FileDBCheck) { 
	$tbFileDB.Text = $global:FileDBLocation
	WriteLog -LogType Success -LogMsg ("Файл БД успешно загружен "+$global:FileDBLocation+"`n")
} else { 
	WriteLog -LogType Error -LogMsg ("Файл БД не найден`n")
}


$btnFileImportBrowse.DataBindings.DefaultDataSourceUpdateMode = 0
$System_Drawing_Point = New-Object System.Drawing.Point
$System_Drawing_Point.X = 203
$System_Drawing_Point.Y = 113
$btnFileImportBrowse.Location = $System_Drawing_Point
$btnFileImportBrowse.Name = "btnFileImportBrowse"
$System_Drawing_Size = New-Object System.Drawing.Size
$System_Drawing_Size.Height = 23
$System_Drawing_Size.Width = 82
$btnFileImportBrowse.Size = $System_Drawing_Size
$btnFileImportBrowse.TabIndex = 1
$btnFileImportBrowse.Text = "Обзор..."
$btnFileImportBrowse.UseVisualStyleBackColor = $True
$btnFileImportBrowse.add_Click({SettingsUpdate "FileImport"})
$tabPage3.Controls.Add($btnFileImportBrowse)

$btnFileDBBrowse.DataBindings.DefaultDataSourceUpdateMode = 0
$System_Drawing_Point = New-Object System.Drawing.Point
$System_Drawing_Point.X = 203
$System_Drawing_Point.Y = 26
$btnFileDBBrowse.Location = $System_Drawing_Point
$btnFileDBBrowse.Name = "btnFileDBBrowse"
$System_Drawing_Size = New-Object System.Drawing.Size
$System_Drawing_Size.Height = 23
$System_Drawing_Size.Width = 82
$btnFileDBBrowse.Size = $System_Drawing_Size
$btnFileDBBrowse.TabIndex = 0
$btnFileDBBrowse.Text = "Обзор..."
$btnFileDBBrowse.UseVisualStyleBackColor = $True
$btnFileDBBrowse.add_Click({SettingsUpdate "FileDB"})
$tabPage3.Controls.Add($btnFileDBBrowse)

$labelAbout.DataBindings.DefaultDataSourceUpdateMode = 0
$System_Drawing_Point = New-Object System.Drawing.Point
$System_Drawing_Point.X = 5
$System_Drawing_Point.Y = 388
$labelAbout.Location = $System_Drawing_Point
$labelAbout.Name = "labelAbout"
$System_Drawing_Size = New-Object System.Drawing.Size
$System_Drawing_Size.Height = 84
$System_Drawing_Size.Width = 425
$labelAbout.Size = $System_Drawing_Size
$labelAbout.Text = "О программе:`nСоздание файла импорта для регистрации нового пользователя в Domino Administrator.`nДанные могут быть получены из AD или введены вручную.`nКосарев Альберт, 2016"
$tabPage3.Controls.Add($labelAbout)

$labelFileImport.DataBindings.DefaultDataSourceUpdateMode = 0
$System_Drawing_Point = New-Object System.Drawing.Point
$System_Drawing_Point.X = 6
$System_Drawing_Point.Y = 113
$labelFileImport.Location = $System_Drawing_Point
$labelFileImport.Name = "labelFileImport"
$System_Drawing_Size = New-Object System.Drawing.Size
$System_Drawing_Size.Height = 23
$System_Drawing_Size.Width = 186
$labelFileImport.Size = $System_Drawing_Size
$labelFileImport.Text = "Файл для импорта в Lotus Domino"
$labelFileImport.TextAlign = 32
$tabPage3.Controls.Add($labelFileImport)

$labelFileDB.DataBindings.DefaultDataSourceUpdateMode = 0
$System_Drawing_Point = New-Object System.Drawing.Point
$System_Drawing_Point.X = 16
$System_Drawing_Point.Y = 26
$labelFileDB.Location = $System_Drawing_Point
$labelFileDB.Name = "labelFileDB"
$System_Drawing_Size = New-Object System.Drawing.Size
$System_Drawing_Size.Height = 23
$System_Drawing_Size.Width = 154
$labelFileDB.Size = $System_Drawing_Size
$labelFileDB.Text = "Файл с БД пользователей"
$labelFileDB.TextAlign = 32
$tabPage3.Controls.Add($labelFileDB)


$ofdDBFile.DefaultExt = "csv"
$ofdDBFile.ShowHelp = $True
$ofdDBFile.InitialDirectory = ".\"
$ofdDBFile.Filter = "CSV файлы (*.csv)|*.csv|Текстовые файлы (*.txt)|*.txt|Все файлы|*.*"
$ofdDBFile.Title = "Выберите файл с БД пользователей"
#$ofdDBFile.add_FileOk({})

$ofdImportFile.FileName = "openFileDialog2"
$ofdImportFile.ShowHelp = $True
$ofdImportFile.InitialDirectory = ".\"
$ofdImportFile.Title = "Выберите файл для импорта в Lotus Domino"
#$ofdImportFile.add_FileOk({})

# Tooltips
$tooltip_btnCheckAD.SetToolTip($btnCheckAD, "Поиск пользователя в AD")
$tooltip_cbDept.SetToolTip($labelDept,"Формат записи: Eng;Рус")
$tooltip_cbDept.SetToolTip($cbDept,"Формат записи: Eng;Рус")


#Save the initial state of the form
$InitialFormWindowState = $form1.WindowState
#Init the OnLoad event to correct the initial state of the form
$form1.add_Load($OnLoadForm_StateCorrection)
$form1.Add_Shown({$form1.Activate();$tbUserAD.Focus()})
#Show the Form
$form1.ShowDialog()| Out-Null
}

function SettingsUpdate ($File) {
	switch ($File) {
		FileDB {
			if ($ofdDBFile.ShowDialog() -eq "OK") {
				$ofdDBFile.Filename
				Set-ItemProperty -Path $global:RegPath -Name $File -Value $ofdDBFile.Filename
				$tbFileDB.Text = $ofdDBFile.Filename
				$global:FileDBLocation = $ofdDBFile.Filename
				$global:FileDBCheck = 1
				WriteLog -LogType Success -LogMsg ("Файл БД успешно задан "+$global:FileDBLocation+"`n")
			}
		}
		FileImport {
			if ($ofdImportFile.ShowDialog() -eq "OK") {
				$ofdImportFile.Filename
				Set-ItemProperty -Path $global:RegPath -Name $File -Value $ofdImportFile.Filename
				$tbFileImport.Text = $ofdImportFile.Filename
				$global:FileImportLocation = $ofdImportFile.Filename
				$global:FileImportCheck = 1
				WriteLog -LogType Success -LogMsg ("Файл импорта успешно задан "+$global:FileImportLocation+"`n")
			}
		}
		
	}
}

function FIModeChange {
	if ($chbAppend.Checked -eq $True) {
		$btnCreateFile.Text = "Дозапись файла импорта Lotus"
	} else {
		$btnCreateFile.Text = "Создать файл импорта Lotus"
	}
}

function FileImportCreate {
	if (!$global:FileImportCheck) {
		WriteLog -LogType Error -LogMsg ("Необходимо задать настройки`n")
		return
	}
	# Input Validation
	if ([string]::IsNullOrEmpty($tbNameEnglish.Text) -or [string]::IsNullOrEmpty($tbNameRus.Text) -or [string]::IsNullOrEmpty($tbEmailAD.Text) -or [string]::IsNullOrEmpty($tbPassword.Text) -or [string]::IsNullOrEmpty($cbDept.Text)) {
		WriteLog -LogType Error -LogMsg "Все поля должны быть заполнены`n"
		return
	}
	if (!([regex]::Match($tbEmailAD.Text,"^\w+@(\w+\.){1,4}\w+$").Success)) {
		WriteLog -LogType Error -LogMsg "Поле Email-адрес должно быть задано по формату`n"
		return
	}
		
	if (!([regex]::Match($cbDept.Text,"\w;\w").Success)) {
		WriteLog -LogType Error -LogMsg "Поле подразделение должно быть задано по формату`n"
		return
	}
	
	$global:NameEmail = [regex]::Replace($tbEmailAD.Text,"@.*$","")
	$global:FDBContent = @(Get-Content $global:FileDBLocation)
	$FIContent = @(Get-Content $global:FileImportLocation)
	ForEach ($FILines in $FIContent) {
		$FIFields = $FILines -split ";"
		if ($FIFields[9] -eq ($global:NameEmail+".nsf")) {
			WriteLog -LogType Info -LogMsg ("Пользователь "+$global:NameEmail+" уже записан в файл импорта`n")
			return
		}
		$EmailToFind = $FIFields[9]
		if (![string]::IsNullOrEmpty($EmailToFind)) {
			if ((![Regex]::IsMatch($global:FDBContent,";$EmailToFind;")) -and (!$chbAppend.Checked)) {
				$Msg = [System.Windows.Forms.MessageBox]::Show("В файле импорта есть пользователь "+$FIContent[0]+", который еще на занесен в базу данных. Перезаписать файл?" , "Проблема" , 3, [System.Windows.Forms.MessageBoxIcon]::Warning)
				if ($Msg -ne "Yes") {
					return
				}
			}
		}
	}

	$NameEng = $tbNameEnglish.Text -split " "
	$Dept = $cbDept.Text -split ";"
	
	$FIString = ($NameEng[0]+";"+$NameEng[1]+";;"+$Dept[0]+";"+$tbPassword.Text+";"+$global:IDFolder+";"+$global:NameEmail+".id;"+$global:MailSRV+";;"+$global:NameEmail+".nsf;;;;;;"+$tbEmailAD.Text+";"+$global:NameEmail+ ";"+$tbNameRus.Text+";"+$Dept[1]+$global:MailTemplate+[Environment]::NewLine)
	# string for DB file without email
	$global:FDBString = ($NameEng[0]+";"+$NameEng[1]+";;"+$Dept[0]+";"+$tbPassword.Text+";;"+$global:NameEmail+";"+$global:MailSRV+";;"+$global:NameEmail+".nsf;;;;;;;"+$global:NameEmail+ ";"+$tbNameRus.Text+";"+$Dept[1]+$global:MailTemplate+[Environment]::NewLine)
	
	WriteLog -LogType Info -LogMsg ("Сформирована строка импорта: `n"+$FIString)
	$global:CanRenewDB = $True
	if ($chbAppend.Checked) {	
		try {
			[IO.File]::AppendAllText($global:FileImportLocation, $FIString, $global:Encoding)
			#Out-File $global:FileImportLocation -inputObject $FIString -Force -Append -Encoding Default
			WriteLog -LogType Success -LogMsg "Файл импорта успешно дозаписан`n"
		}
		catch {
			WriteLog -LogType Error -LogMsg "Ошибка дозаписи файла импорта`n"
			return
		}
	} else {
		try {
			[IO.File]::WriteAllText($global:FileImportLocation, $FIString, $global:Encoding)
			#Out-File $global:FileImportLocation -inputObject $FIString -Force -Encoding Default
			WriteLog -LogType Success -LogMsg "Файл импорта успешно записан`n"
		}
		catch {
			WriteLog -LogType Error -LogMsg "Ошибка записи файла импорта`n"
			return
		}
	}
}
	
function FileDBRenew {
	if (!$global:FileDBCheck) {
		WriteLog -LogType Error -LogMsg ("Необходимо задать настройки`n")
		return
	}
	if (!$global:CanRenewDB) {
		WriteLog -LogType Error -LogMsg ("Необходимо создать файл импорта`n")
		return
	}
	
	$global:FDBContent = [System.IO.File]::ReadAllText($global:FileDBLocation, $global:Encoding)
	if (![Regex]::IsMatch($global:FDBContent,"\r\n$")) {
		try {
			[IO.File]::AppendAllText($global:FileDBLocation, [Environment]::NewLine, $global:Encoding)
		}
		catch {
			WriteLog -LogType Error -LogMsg "Ошибка обновления файла БД`n"
			return
		}
	}
	
	if ([Regex]::IsMatch($global:FDBContent,";$global:NameEmail;")) {
		WriteLog -LogType Error -LogMsg ("Пользователь "+$global:NameEmail+" уже есть в БД`n")
		return
	}

	try {
		[IO.File]::AppendAllText($global:FileDBLocation, $global:FDBString, $global:Encoding)
		#Out-File $global:FileDBLocation -inputObject $global:FDBString -Force -Append -Encoding Default
		WriteLog -LogType Success -LogMsg "Файл БД успешно обновлен`n"
	}
	catch {
		WriteLog -LogType Error -LogMsg "Ошибка обновления файла БД`n"
		return
	}
	
	# Disable DB file writing without new data
	$global:CanRenewDB = $False
}
	
GenerateForm


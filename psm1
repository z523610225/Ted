```powershell

$Form2.Add_KeyDown({
		if ($_.KeyCode -eq 'Enter'){
			EmailAuth $TextBox1.Text $TextBox2.Text
		}
	}
)

$YesButton.Add_Click({
		EmailAuth $TextBox1.Text $TextBox2.Text
	}
)

$CancelButton.Add_Click({
		Get-PSSession | Remove-PSSession
		$Form2.Close()
	}
)

$Form2.Add_FormClosing({Get-PSSession | Remove-PSSession})

function EmailAuth{
	param([string]$Username,$Password1)
	$Label3.text = "Loading"
	$Password = ConvertTo-SecureString -String $Password1 -AsPlainText -Force
	$UserCredential = New-Object System.Management.Automation.PSCredential -ArgumentList $Username,$Password
	$Session = New-PSSession -ConfigurationName Microsoft.Exchange -ConnectionUri https://outlook.office365.com/powershell-liveid/ -Credential $UserCredential -Authentication Basic -AllowRedirection
	# $Session = New-PSSession -ConfigurationName Microsoft.Exchange -ConnectionUri https://outlook.office365.com/powershell-liveid/ -Credential $cred -Authentication Basic -AllowRedirection   #Will prompts for username and password
	$ss = Import-PSSession $Session -DisableNameChecking 
	# NOTE: Install-Module MSOnline
	# *********** Connect-MsolService ***********
	Connect-MsolService -credential $UserCredential 
	$UserCredential = ""
    $Label3.text = "Exchange Loaded"
	[System.Windows.MessageBox]::Show('Exchange Connected!')
	$Form2.Close()
}

$Form2.Add_Shown({$Form2.Activate()})
[void] $Form2.ShowDialog()

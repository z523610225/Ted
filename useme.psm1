# Connect to O365
function EmailAuth{
	param([string]$Username,$Password1)
	$Password = ConvertTo-SecureString -String $Password1 -AsPlainText -Force
	$UserCredential = New-Object System.Management.Automation.PSCredential -ArgumentList $Username,$Password
	$Session = New-PSSession -ConfigurationName Microsoft.Exchange -ConnectionUri https://outlook.office365.com/powershell-liveid/ -Credential $UserCredential -Authentication Basic -AllowRedirection
	$ss = Import-PSSession $Session -DisableNameChecking 
	# NOTE: Install-Module MSOnline
	# *********** Connect-MsolService ***********
	Connect-MsolService -credential $UserCredential 
	$UserCredential = ""
}


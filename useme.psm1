<# PowerShell Lib
   Author: Ted
#>

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

# AnrSearch for user
function AnrSearch{
	param([string]$Username)
	$Names = (Get-ADUser -ldapfilter "(anr=$Username)").name
	ForEach($Name in $Names){
		write-output $Name
	}
}

# Users how have not logged on for 90 days
function NotLogon {
	# CONSIDERATIONS - search base
 	Get-ADUser -Filter {enabled -eq $true} -Properties LastLogonDate | Where-Object {$_.LastLogonDate -lt (Get-Date).AddDays(-90)} | export-csv .\90days.csv -NoTypeInformation
}

# AD User information
function GetBasic{
	param([string]$Username)
	$UserObject = Get-ADUser -filter {Name -eq $Username} -Properties *
	$SAName = $UserObject.SamAccountName
	$UserAD = $UserObject | fl Enabled,LockedOut,PasswordLastSet,LastLogonDate,DisplayName,EmailAddress,Department,Created
	# $LastLogon = $UserObject | select @{n='LastLogon';e={[DateTime]::FromFileTime($_.LastLogon)}}
	$UserADGroup = Get-ADPrincipalGroupMembership -Identity $SAName | ForEach-Object{
		$_.name
	}
	
	$UserInfo = (($UserAD | out-string).trim()) + "`n`nSecurity Groups -`n" + ($UserADGroup | sort-object | out-string)
	$UserInfo
	#Load Okta status	
	$pwdTxt1 = Get-Content "C:\temp\psscripts\checkinfo\pokta.txt" 
	$securePwd1 = $pwdTxt1 | ConvertTo-SecureString 
	$Pass = (New-Object PSCredential "user",$securePwd1).GetNetworkCredential().Password
	$headers = @{
		'authorization' = "SSWS $Pass"
		'accept' = 'application/json'
		'contect-type' = 'application/json'
	}
	#check if user has an Okta account
	try{
		$OktaInfo = Invoke-RestMethod -Uri "https://pexa.okta.com/api/v1/users/$SAName" -Method Get -Headers $headers 
		$TextBox2.Text = $UserInfo + "`nOKTA Account Status - `n" + ($OktaInfo.status | out-string) + "`n" + ($OktaInfo.profile.usertype | out-string)
	}
	catch{
		$TextBox2.Text = $UserInfo + "`nOKTA Account Status - `nUser OKTA Account can not be found"
	}
}

# Check computer assignment information against SnipeIT using Rest API
function GetComputer{
	param([string]$Username)
	$TextBox3.Clear()
	$headers = @{
		'authorization' = 'Bearer eyJ0eXAiOiJKV1QiLCJhbGciOiJSUzI1NiIsImp0aSI6IjFjMDQwZWIxMWViYmIxNDYyY2IxMjNkZDRmMmFiMWVmOWNkZjI3ZjA3NWMyZTU2ZjIyM2Y4NjRmMTExOGRhZjNiMTI3YzM4MmQxNDk0NzI5In0.eyJhdWQiOiIzIiwianRpIjoiMWMwNDBlYjExZWJiYjE0NjJjYjEyM2RkNGYyYWIxZWY5Y2RmMjdmMDc1YzJlNTZmMjIzZjg2NGYxMTE4ZGFmM2IxMjdjMzgyZDE0OTQ3MjkiLCJpYXQiOjE1OTYxNTc0NjUsIm5iZiI6MTU5NjE1NzQ2NSwiZXhwIjoxNjI3NjkzNDY1LCJzdWIiOiI1MTAiLCJzY29wZXMiOltdfQ.ZJMXcw1giawt1n-Sbe3IfN6XJaUm9cD8dPyaEI_i7sIPm55B31ZgtWeb9STgBWrhU-sRMoqhSLzD1B9kk0m5jYvw823XS9uZ0MAnMSwFQWKHdR7CzmkZkwmlkSZl8deBk9O0MdLTU3GNKR6t5rj6JUDgsJPg6k3fkkNito7gjhl8t43r-kc5XAeRZhUTsSLJVpnwwKB_eHnQJWwq8ds8kvTjSEEEXbplBwv_rJ5Nc8wpLEo3ijd0lSWGO-tuXbGUxBLCKLNhxxHq1tDvyPcU6eKBoT6_VqWL_LLAjYHnJLNCxCkwoNBUnfSzbcGiIKjaO6GkA7eQNxutKAbBWpX8kBBL_caA0LQrsObsR6iI3zSus69eeJW9-e5z0OlaC4fNGDXzxU9Xz47b-wPQJ32G4pd5z69deot0nE6Prt8uaeuN3M0Gn93HABpf8sHfb_IHidjtsE2QA5nGKQNHWujClZANp3Py_IPtSqi0-SD7A9EBm0qCwWXS8fxO7oWHm8zzHfKkip15kNTD6C1MYXPcY-Lqc9SjQ77qYS9kKxjSIPmV9nwRax6apeDpc9BF-oXasCm7zaBRS1gsD3AZiWc71NI4DAZVHfF2_Bx8_x164RzlWntkueQMOkSpWSmwOfSsuYp0kA0_11X7rjGaBN6gGScnpShZjIqmlBvCXQMpVTo'
		'accept' = 'application/json'
		'contect-type' = 'application/json'
	}
	$GetUserID = Invoke-RestMethod -Uri "https://hubtools.pexa.com.au/snipe-it/api/v1/users?search=$Username" -Method Get -Headers $headers
	# if(($GetUserID.rows | measure-object -property id).count -ne 1){
		# $TextBox3.Text = "User is not unique in Snipe IT"
		# GetLicenses $Username
		# return
	# }
	#Check if the object selected is right or wrong
	$UserID = ($GetUserID.rows | select-object -first 1).id	
	if(($GetUserID.rows | select-object -first 1).name -ne $Username){
		$TextBox3.Text = "User is not unique in Snipe IT"
		GetLicenses $Username
		return
	}
	$ComputerInfo = Invoke-RestMethod -Uri "https://hubtools.pexa.com.au/snipe-it/api/v1/users/$UserID/assets" -Method Get -Headers $headers
	$ComputerName = $ComputerInfo.rows.name
	
	ForEach($Computer in $ComputerName){
		if (test-connection $Computer -count 1 -quiet){
			# $online1 = Write-host "`n`n $Computer is currently online"
			# $TextBox3.Text = $ComputerName + $online1
			$TextBox3.AppendText("$Computer is currently online`n")
			# $progressbar1.PerformStep()
		}
		else {
			# $offline1 = Write-host "`n`n $Computer is currently offline"
			# $TextBox3.Text = $ComputerName + $offline1
			$TextBox3.AppendText("$Computer is currently offline`n")
			# $progressbar1.PerformStep()
		}
	}
	if($TextBox3.text -eq ""){
		$TextBox3.Text = "No laptop assigned"
	}
} 

# O365 user license removal
function GetLicenses{
	param([string]$Username)
<# 	$cred = Get-Credential
	$Session = New-PSSession -ConfigurationName Microsoft.Exchange -ConnectionUri https://outlook.office365.com/powershell-liveid/ -Credential $cred -Authentication Basic -AllowRedirection   #Will prompts for username and password
	$ss = Import-PSSession $Session -DisableNameChecking 
	
	# NOTE: Install-Module MSOnline
	# *********** Connect-MsolService ***********
	Connect-MsolService -credential $cred #>
	$TextBox4.clear()	
	$Licenses = (Get-MsolUser -SearchString $Username -ErrorAction SilentlyContinue | where-object DisplayName -eq $Username).licenses.accountskuid 
	if (-Not $Licenses){
		$TextBox4.Text = "You have to connect to `nExchange first"
		$progressbar1.value = 100
		return
	}
	$Sku = @{
		"O365_BUSINESS_ESSENTIALS" = "Office 365 Business Essentials"
		"O365_BUSINESS_PREMIUM" = "Office 365 Business Premium"
		"DESKLESSPACK"= "Office 365 (Plan K1)"
		"DESKLESSWOFFPACK" = "Office 365 (Plan K2)"	
		"LITEPACK" = "Office 365 (Plan P1)"
		"EXCHANGESTANDARD" = "Office 365 Exchange Online Only"
		"STANDARDPACK" = "Enterprise Plan E1"
		"STANDARDWOFFPACK" = "Office 365 (Plan E2)"
		"ENTERPRISEPACK" = "Enterprise Plan E3"
		"ENTERPRISEPACKLRG" = "Enterprise Plan E3"
		"ENTERPRISEWITHSCAL" = "Enterprise Plan E4"
		"STANDARDPACK_STUDENT" = "Office 365 (Plan A1) for Students"
		"STANDARDWOFFPACKPACK_STUDENT" = "Office 365 (Plan A2) for Students"
		"ENTERPRISEPACK_STUDENT" = "Office 365 (Plan A3) for Students"
		"ENTERPRISEWITHSCAL_STUDENT" = "Office 365 (Plan A4) for Students"
		"STANDARDPACK_FACULTY" = "Office 365 (Plan A1) for Faculty"
		"STANDARDWOFFPACKPACK_FACULTY" = "Office 365 (Plan A2) for Faculty"
		"ENTERPRISEPACK_FACULTY" = "Office 365 (Plan A3) for Faculty"
		"ENTERPRISEWITHSCAL_FACULTY" = "Office 365 (Plan A4) for Faculty"
		"ENTERPRISEPACK_B_PILOT" = "Office 365 (Enterprise Preview)"
		"STANDARD_B_PILOT" = "Office 365 (Small Business Preview)"
		"VISIOCLIENT" = "Visio Pro Online"
		"POWER_BI_ADDON" = "Office 365 Power BI Addon"
		"POWER_BI_INDIVIDUAL_USE" = "Power BI Individual User"
		"POWER_BI_STANDALONE" = "Power BI Stand Alone"
		"POWER_BI_STANDARD" = "Power-BI Standard"
		"PROJECTESSENTIALS" = "Project Lite"
		"PROJECTCLIENT" = "Project Professional"
		"PROJECTONLINE_PLAN_1" = "Project Online"
		"PROJECTONLINE_PLAN_2" = "Project Online and PRO"
		"ProjectPremium" = "Project Online Premium"
		"ECAL_SERVICES" = "ECAL"
		"EMS" = "Enterprise Mobility Suite"
		"RIGHTSMANAGEMENT_ADHOC" = "Windows Azure Rights Management"
		"MCOMEETADV" = "PSTN conferencing"
		"SHAREPOINTSTORAGE" = "SharePoint storage"
		"PLANNERSTANDALONE" = "Planner Standalone"
		"CRMIUR" = "CMRIUR"
		"BI_AZURE_P1" = "Power BI Reporting and Analytics"
		"INTUNE_A" = "Windows Intune Plan A"
		"PROJECTWORKMANAGEMENT" = "Office 365 Planner Preview"
		"ATP_ENTERPRISE" = "Exchange Online Advanced Threat Protection"
		"EQUIVIO_ANALYTICS" = "Office 365 Advanced eDiscovery"
		"AAD_BASIC" = "Azure Active Directory Basic"
		"RMS_S_ENTERPRISE" = "Azure Active Directory Rights Management"
		"AAD_PREMIUM" = "Azure Active Directory Premium"
		"MFA_PREMIUM" = "Azure Multi-Factor Authentication"
		"STANDARDPACK_GOV" = "Microsoft Office 365 (Plan G1) for Government"
		"STANDARDWOFFPACK_GOV" = "Microsoft Office 365 (Plan G2) for Government"
		"ENTERPRISEPACK_GOV" = "Microsoft Office 365 (Plan G3) for Government"
		"ENTERPRISEWITHSCAL_GOV" = "Microsoft Office 365 (Plan G4) for Government"
		"DESKLESSPACK_GOV" = "Microsoft Office 365 (Plan K1) for Government"
		"ESKLESSWOFFPACK_GOV" = "Microsoft Office 365 (Plan K2) for Government"
		"EXCHANGESTANDARD_GOV" = "Microsoft Office 365 Exchange Online (Plan 1) only for Government"
		"EXCHANGEENTERPRISE_GOV" = "Microsoft Office 365 Exchange Online (Plan 2) only for Government"
		"SHAREPOINTDESKLESS_GOV" = "SharePoint Online Kiosk"
		"EXCHANGE_S_DESKLESS_GOV" = "Exchange Kiosk"
		"RMS_S_ENTERPRISE_GOV" = "Windows Azure Active Directory Rights Management"
		"OFFICESUBSCRIPTION_GOV" = "Office ProPlus"
		"MCOSTANDARD_GOV" = "Lync Plan 2G"
		"SHAREPOINTWAC_GOV" = "Office Online for Government"
		"SHAREPOINTENTERPRISE_GOV" = "SharePoint Plan 2G"
		"EXCHANGE_S_ENTERPRISE_GOV" = "Exchange Plan 2G"
		"EXCHANGE_S_ARCHIVE_ADDON_GOV" = "Exchange Online Archiving"
		"EXCHANGE_S_DESKLESS" = "Exchange Online Kiosk"
		"SHAREPOINTDESKLESS" = "SharePoint Online Kiosk"
		"SHAREPOINTWAC" = "Office Online"
		"YAMMER_ENTERPRISE" = "Yammer Enterprise"
		"EXCHANGE_L_STANDARD" = "Exchange Online (Plan 1)"
		"MCOLITE" = "Lync Online (Plan 1)"
		"SHAREPOINTLITE" = "SharePoint Online (Plan 1)"
		"OFFICE_PRO_PLUS_SUBSCRIPTION_SMBIZ" = "Office ProPlus"
		"EXCHANGE_S_STANDARD_MIDMARKET" = "Exchange Online (Plan 1)"
		"MCOSTANDARD_MIDMARKET" = "Lync Online (Plan 1)"
		"SHAREPOINTENTERPRISE_MIDMARKET" = "SharePoint Online (Plan 1)"
		"OFFICESUBSCRIPTION" = "Office ProPlus"
		"YAMMER_MIDSIZE" = "Yammer"
		"DYN365_ENTERPRISE_PLAN1" = "Dynamics 365 Customer Engagement Plan Enterprise Edition"
		"ENTERPRISEPREMIUM_NOPSTNCONF" = "Enterprise E5 (without Audio Conferencing)"
		"ENTERPRISEPREMIUM" = "Enterprise E5 (with Audio Conferencing)"
		"MCOSTANDARD" = "Skype for Business Online Standalone Plan 2"
		"PROJECT_MADEIRA_PREVIEW_IW_SKU" = "Dynamics 365 for Financials for IWs"
		"STANDARDWOFFPACK_IW_STUDENT" = "Office 365 Education for Students"
		"STANDARDWOFFPACK_IW_FACULTY"  = "Office 365 Education for Faculty"
		"EOP_ENTERPRISE_FACULTY" = "Exchange Online Protection for Faculty"
		"EXCHANGESTANDARD_STUDENT" = "Exchange Online (Plan 1) for Students"
		"OFFICESUBSCRIPTION_STUDENT" = "Office ProPlus Student Benefit"
		"STANDARDWOFFPACK_FACULTY" = "Office 365 Education E1 for Faculty"
		"STANDARDWOFFPACK_STUDENT" = "Microsoft Office 365 (Plan A2) for Students"
		"DYN365_FINANCIALS_BUSINESS_SKU" = "Dynamics 365 for Financials Business Edition"
		"DYN365_FINANCIALS_TEAM_MEMBERS_SKU" = "Dynamics 365 for Team Members Business Edition"
		"FLOW_FREE" = "Microsoft Flow Free"
		"POWER_BI_PRO" = "Power BI Pro"
		"O365_BUSINESS" = "Office 365 Business"
		"DYN365_ENTERPRISE_SALES" = "Dynamics Office 365 Enterprise Sales"
		"RIGHTSMANAGEMENT" = "Rights Management"
		"PROJECTPROFESSIONAL" = "Project Professional"
		"VISIOONLINE_PLAN1" = "Visio Online Plan 1"
		"EXCHANGEENTERPRISE" = "Exchange Online Plan 2"
		"DYN365_ENTERPRISE_P1_IW" = "Dynamics 365 P1 Trial for Information Workers"
		"DYN365_ENTERPRISE_TEAM_MEMBERS" = "Dynamics 365 For Team Members Enterprise Edition"
		"CRMSTANDARD" = "Microsoft Dynamics CRM Online Professional"
		"EXCHANGEARCHIVE_ADDON" = "Exchange Online Archiving For Exchange Online"
		"EXCHANGEDESKLESS" = "Exchange Online Kiosk"
		"SPZA_IW" = "App Connect"
		"WINDOWS_STORE" = "Windows Store for Business"
		"MCOEV" = "Microsoft Phone System"
		"VIDEO_INTEROP" = "Polycom Skype Meeting Video Interop for Skype for Business"
		"SPE_E5" = "Microsoft 365 E5"
		"SPE_E3" = "Microsoft 365 E3"
		"ATA" = "Advanced Threat Analytics"
		"MCOPSTN2" = "Domestic and International Calling Plan"
		"FLOW_P1" = "Microsoft Flow Plan 1"
		"FLOW_P2" = "Microsoft Flow Plan 2"
		"CRMSTORAGE" = "Microsoft Dynamics CRM Online Additional Storage"
		"SMB_APPS" = "Microsoft Business Apps"
		"MICROSOFT_BUSINESS_CENTER" = "Microsoft Business Center"
		"DYN365_TEAM_MEMBERS" = "Dynamics 365 Team Members"
		"STREAM" = "Microsoft Stream Trial"
		"EMSPREMIUM" = "ENTERPRISE MOBILITY + SECURITY E5"
	}
	
	Foreach($License in $Licenses){
		$LicenseItem = $License -split ":" | Select-Object -Last 1
		$UPC = $Sku.Item("$LicenseItem")
		$TextBox4.AppendText("$UPC ($License) `n")
		# write-host "$UPC ($License)"
	}
	if($TextBox4.text -eq ""){
		$TextBox4.Text = "No license assigned"
	}
	# Get-PSSession | Remove-PSSession # in case there are sessions left open
}



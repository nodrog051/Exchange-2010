
$alldistLists = Get-DistributionGroup ChubbEuropeGermanyDuesseldorf@Chubb.com

Foreach ($distList in $alldistLists){
	$distListMembers = $null
	$distListMembers = Get-DistributionGroupMember $distList

	Foreach ($distListMember in $distListMembers){
		If ($distListMember.RecipientType -eq "User"){
			$thisUser = Get-User $distListMember.Name
	
			write-host "DistListMember : " $thisUser.DisplayName ", " $thisUser.RecipientType ", " $thisUser.WindowsEmailAddress
			#See if there is a corresponding contact

			If ($thisUser.WindowsEmailAddress.length -eq "0"){
				$str = "This user has no WindowsEmailAddress "  + $thisUser.Name
				$str | Out-File c:\temp\userswithnoemailaddress.txt -append

			}

			$thisContact = Get-Contact $thisUser.WindowsEmailAddress.ToString() -ea silentlycontinue
			if ($thisContact -ne $null){
				write-Host "Adding this contact: " $thisContact.DisplayName " to distribution list : " $distList.Name
				Add-DistributionGroupMember $distlist -member $thisContact -ea silentlycontinue
				
			} else {

				$str = "This user has no corresponding contact "  + $thisUser.Name
				$str | Out-File c:\temp\userswithnoemailaddress.txt -append

			}
		}
		
	}
}
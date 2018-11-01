<#	
	.NOTES
	===========================================================================
	 Updated on:   	6/26/2018
	 Created by:   	/u/TheLazyAdministrator
     Contributors:  /u/jmn_lab, /u/nothingpersonalbro, /u/Kroucher	
	===========================================================================

        AzureAD  Module is required
            Install-Module -Name AzureAD
            https://www.powershellgallery.com/packages/azuread/
        ReportHTML Moduile is required
            Install-Module -Name ReportHTML
            https://www.powershellgallery.com/packages/ReportHTML/

	.DESCRIPTION
		Generate an interactive HTML report on your Office 365 tenant. Report on Users, Tenant information, Groups, Policies, Contacts, Mail Users, Licenses and more!
    
    .Link
        https://thelazyadministrator.com/2018/06/22/create-an-interactive-html-report-for-office-365-with-powershell/
#>
#########################################
#                                       #
#            VARIABLES                  #
#                                       #
#########################################
param([string]$dir)
#Company logo that will be displayed on the left, can be URL or UNC
$CompanyLogo = "https://irp-cdn.multiscreensite.com/6c87c673/dms3rep/multi/mobile/Future%202016%20Logo-533x133.png"

#Location the report will be saved to
$ReportSavePath = "C:\temp\ntfsreport\reports\"

#Table Variables
$FolderPermissionsTable = New-Object 'System.Collections.Generic.List[System.Object]'
$SecurityGroupsTable = New-Object 'System.Collections.Generic.List[System.Object]'

########################################

#Import Modules


Import-Module C:\Temp\ntfsreport\modules\ntfssecurity\NTFSSecurity.psd1
Import-Module C:\Temp\ntfsreport\modules\reporthtml\ReportHTML.psd1


# Folder Permissions Tab starts



$Permissions = Get-ChildItem $dir -Recurse| where {$_.Attributes -match 'Directory'} | Get-NTFSAccess -ExcludeInherited | where-object -filterscript { ($_.Account -notlike "NT AUTHORITY\SYSTEM") -and ($_.Account -notlike "CREATOR OWNER")} | select FullName,Account,AccessRights

Foreach ($Permission in $Permissions)
{

	$Name = $Permission.FullName
	$Account = $Permission.Account
	$Rights = $Permission.AccessRights


$permsobj = [PSCustomObject]@{
		'Path'	          = $Name
		'User'	 		  = $Account
		'Access Rights'	  = $Rights
	}

$FolderPermissionsTable.add($permsobj)
}
If (($FolderPermissionsTable).count -eq 0)
{
	$FolderPermissionsTable = [PSCustomObject]@{
		'Information'  = 'Information: Issue with directory path'
	}
}



#Folder Permissions tab end



 #Active Directory groups tab starts

Import-Module ActiveDirectory

$ADGroups = Get-ADGroup -Filter 'GroupCategory -eq "Security"' |where-object{$_.Name -NotMatch "Replicator|Server Operators|Account Operators|Performance Monitor Users|Performance Log Users|Pre-Windows|Certificate|^IIS|Incoming Forest|Event Log|^Cryptographic|^Distributed COM|^Network Conf|^SQL|Enterprise Admins|Schema Admins|^Enterprise Read-only|^Exchange|^Hyper|^Access Control Assistance|SBS|^Storage Replica|^Wse|^Clonable Domain|Protected Users|^DNS|Domain Controllers|Domain Computers|^Group Policy|^WSUS|RODC|Key Admins|^DHCP|Servers|Computers|^Windows Auth|^Print Op|^Backup Op|^System Man|Enterprise Key Admins|^Remote Man|^WSS_"}

foreach ($ADGroup in $ADGroups) { 

$Members = Get-ADGroup -Filter {Name -eq $ADGroup.Name}  | Get-ADGroupMember |  select Name 
         
    Foreach ($Member in $Members){

	$Name = $ADGroup.Name
	$Members = $Member.Name


$Groupsobj = [PSCustomObject]@{
		'Name'		= $Name
		'Members'   = $Members
	}

$SecurityGroupsTable.add($Groupsobj)
	}
}

If (($SecurityGroupsTable).count -eq 0)
{
	$SecurityGroupsTable = [PSCustomObject]@{
		'Information'  = 'Information: No Security Groups were found in AD'
	}
}
#Active Directory Groups tab end


#Report generation start

#No Groups
$tabarray = @('Folder Permissions', 'Security Groups')

#With Groups
#$tabarray = @('Folder Permissions', 'Active Directory Groups')




# call the function and pass the array and color expressions


$Rpt = @()
$rpt = New-Object 'System.Collections.Generic.List[System.Object]'
$rpt += get-htmlopenpage -TitleText 'Folder Permissions Report' -LeftLogoString $CompanyLogo 

$rpt += Get-HTMLTabHeader -TabNames $tabarray 
    $rpt += get-htmltabcontentopen -TabName $tabarray[0] -TabHeading ("Report: " + (Get-Date -Format MM-dd-yyyy)) 
        $rpt += Get-HTMLContentOpen -HeaderText "Folder Permissions"
        $rpt += get-htmlcontentdatatable $FolderPermissionsTable -HideFooter
        $rpt += Get-HTMLContentClose
    $rpt += get-htmltabcontentclose
        $rpt += get-htmltabcontentopen -TabName $tabarray[1] -TabHeading ("Report: " + (Get-Date -Format MM-dd-yyyy)) 
        $rpt += Get-HTMLContentOpen -HeaderText "Active Directory Groups"
            $rpt += get-htmlcontentdatatable $SecurityGroupsTable -HideFooter
        $rpt += Get-HTMLContentClose
    $rpt += get-htmltabcontentclose

$rpt += Get-HTMLClosePage

$Day = (Get-Date).Day
$Month = (Get-Date).Month
$Year = (Get-Date).Year
$ReportName = ("$Day" + "-" + "$Month" + "-" + "$Year" + "-" + "Folder Permissions Report")
Save-HTMLReport -ReportContent $rpt -ReportName $ReportName -ReportPath $ReportSavePath
$Attachment = ("$ReportSavePath" + "\" + "$ReportName" + ".html")
$Attached = ('"' + $Attachment + '"')

#Report generation end

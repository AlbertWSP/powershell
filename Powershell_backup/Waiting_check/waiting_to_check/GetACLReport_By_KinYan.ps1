$ShellVersion="2.5"
$FormTitle="ACL Report Master Pro Ultimate - V" + $ShellVersion
$PowershellPath = "C:\Support\GetACLReport"
$ReportPath = Join-Path $PowershellPath "Report"
if((Test-Path $ReportPath) -eq $False)
{
	New-Item -Path $ReportPath -ItemType Directory
}


<#
$ACLFolders=@(	
	("Beijing(CNBJS300)","MEP","\\corp.pbwan.net\cn\Transition\BJS\PROJECT"),
	("Beijing(CNBJS300)","Admin","\\corp.pbwan.net\cn\Transition\BJS\ADMIN"),
	("Beijing(CNBJS300)","BD","\\corp.pbwan.net\cn\Transition\BJS\BD"),
	("Beijing(CNBJS300)","CD","\\corp.pbwan.net\cn\Transition\BJS\CD"),
	("Beijing(CNBJS300)","Common","\\corp.pbwan.net\cn\Transition\BJS\Common"),
	("Beijing(CNBJS300)","FA","\\corp.pbwan.net\cn\Transition\BJS\FA"),
	("Beijing(CNBJS300)","HR","\\corp.pbwan.net\cn\Transition\BJS\HR"),
	("Beijing(CNBJS300)","IT","\\corp.pbwan.net\cn\Transition\BJS\IT"),
	("Beijing(CNBJS300)","Transport","\\corp.pbwan.net\cn\Transition\BJS\Transport"),
	("Chengdu(CNCDU100)","Admin","\\corp.pbwan.net\cn\transition\CDU\Admin"),
	("Chengdu(CNCDU100)","BD","\\corp.pbwan.net\cn\transition\CDU\BD"),
	("Chengdu(CNCDU100)","IT","\\cncdu100dat01\IT"),
	("Chengdu(CNCDU100)","Others","\\corp.pbwan.net\cn\transition\CDU\Others"),
	("Chengdu(CNCDU100)","MEP","\\corp.pbwan.net\cn\transition\CDU\project"),
	("Chengdu(CNCDU100)","MEP","\\corp.pbwan.net\cn\transition\CDU\project\projects"),
	("Shenzhen(CNSNZ100)","Admin","\\corp.pbwan.net\cn\Transition\SNZ\Admin"),
	("Shenzhen(CNSNZ100)","BD","\\corp.pbwan.net\cn\Transition\SNZ\BD"),
	("Shenzhen(CNSNZ100)","HR","\\corp.pbwan.net\cn\Transition\SNZ\HR"),
	("Shenzhen(CNSNZ100)","ISO","\\corp.pbwan.net\cn\Transition\SNZ\ISO"),
	("Shenzhen(CNSNZ100)","IT","\\cnsnz100dat01\IT"),
	("Shenzhen(CNSNZ100)","MEP","\\corp.pbwan.net\cn\Transition\SNZ\mep"),
	("Shenzhen(CNSNZ100)","MEP","\\corp.pbwan.net\cn\Transition\SNZ\mep\Project"),
	("Shenzhen(CNSNZ100)","Others","\\corp.pbwan.net\cn\Transition\SNZ\Others"),
	("Shenzhen(CNSNZ100)","BIM","\\corp.pbwan.net\cn\Transition\SNZ\BIM"),
	("Shenzhen(CNSNZ100)","CS","\\corp.pbwan.net\cn\Transition\SNZ\CS"),
	("Shenzhen(CNSNZ100)","CS","\\corp.pbwan.net\cn\Transition\SNZ\CS\01_Project"),
	("Shenzhen(CNSNZ100)","CSCAD","\\corp.pbwan.net\cn\Transition\SNZ\CSCAD\CS"),
	("Shenzhen(CNSNZ100)","FC","\\corp.pbwan.net\cn\Transition\SNZ\FC"),
	("Shenzhen(CNSNZ100)","Infra","\\corp.pbwan.net\cn\Transition\SNZ\Infra"),
	("Shenzhen(CNSNZ100)","SE","\\corp.pbwan.net\cn\Transition\SNZ\SE"),
	("Guangzhou(CNGGZ100)","Account","\\corp.pbwan.net\cn\Transition\GGZ\Account"),
	("Guangzhou(CNGGZ100)","Admin","\\corp.pbwan.net\cn\Transition\GGZ\Admin"),
	("Guangzhou(CNGGZ100)","BD","\\corp.pbwan.net\cn\Transition\GGZ\BD"),
	("Guangzhou(CNGGZ100)","CS","\\corp.pbwan.net\cn\Transition\GGZ\CS"),
	("Guangzhou(CNGGZ100)","CS","\\corp.pbwan.net\cn\Transition\GGZ\CS\01_Project"),
	("Guangzhou(CNGGZ100)","CSCAD","\\corp.pbwan.net\cn\Transition\GGZ\CSCAD"),
	("Guangzhou(CNGGZ100)","HR","\\corp.pbwan.net\cn\Transition\GGZ\HR"),
	("Guangzhou(CNGGZ100)","IT","\\corp.pbwan.net\cn\Transition\GGZ\IT"),
	("Guangzhou(CNGGZ100)","Others","\\corp.pbwan.net\cn\Transition\GGZ\Others"),
	("Guangzhou(CNGGZ100)","MEP","\\corp.pbwan.net\cn\Transition\GGZ\Projects"),
	("Guangzhou(CNGGZ100)","MEP","\\corp.pbwan.net\cn\Transition\GGZ\Projects\Projects"),
	("Guangzhou(CNGGZ100)","SQA","\\corp.pbwan.net\cn\Transition\GGZ\SQA"),
	("Shanghai(CNSGH300)","BIM","\\corp.pbwan.net\CN\Transition\SGH\BIM"),
	("Shanghai(CNSGH300)","BD","\\corp.pbwan.net\CN\Transition\SGH\Business Development"),
	("Shanghai(CNSGH300)","Account","\\corp.pbwan.net\CN\Transition\GGZ\Account\SH"),
	("Shanghai(CNSGH300)","Admin","\\corp.pbwan.net\cn\Transition\SGH\Admin"),
	("Shanghai(CNSGH300)","HR","\\corp.pbwan.net\cn\Transition\SGH\HR"),
	("Shanghai(CNSGH300)","PMCM","\\corp.pbwan.net\cn\Transition\SGH\PMCM"),
	("Shanghai(CNSGH300)","SQA","\\corp.pbwan.net\cn\Transition\SGH\SQA"),
	("Shanghai(CNSGH300)","BD","\\corp.pbwan.net\CN\Transition\SGH\co-work\BDShare"),
	("Shanghai(CNSGH300)","BIM","\\corp.pbwan.net\CN\Transition\SGH\co-work\BIM_Coordination"),
	("Shanghai(CNSGH300)","BIM","\\corp.pbwan.net\CN\Transition\SGH\co-work\BIM_Public"),
	("Shanghai(CNSGH300)","Account","\\corp.pbwan.net\CN\Transition\SGH\co-work\Finance Share Service"),
	("Shanghai(CNSGH300)","HR","\\corp.pbwan.net\CN\Transition\SGH\co-work\HRShare"),
	("Shanghai(CNSGH300)","BIM","\\corp.pbwan.net\CN\Transition\SGH\co-work\SHO_ShowCase"),
	("Shanghai(CNSGH300)","CS","\\corp.pbwan.net\CN\Transition\SGH\CS"),
	("Shanghai(CNSGH300)","ENV","\\corp.pbwan.net\CN\Transition\SGH\ENV"),
	("Shanghai(CNSGH300)","Facade","\\corp.pbwan.net\CN\Transition\SGH\Facade"),
	("Shanghai(CNSGH300)","MEP","\\corp.pbwan.net\CN\Transition\SGH\MEP"),
	("Shanghai(CNSGH300)","MEP","\\corp.pbwan.net\CN\Transition\SGH\MEP\Project and Admin"),
	("Shanghai(CNSGH300)","SE","\\corp.pbwan.net\CN\Transition\SGH\SE"),
	("Shanghai(CNSGH300)","TP","\\corp.pbwan.net\CN\Transition\SGH\TP"),
	("Shanghai(CNSGH300)","Common","\\corp.pbwan.net\CN\Transition\SGH\Common"),
	("Singapore(SGSIN100)","MEP","\\dcasi400dat01\SGSIN100\DAT01\MEP\MEP1-1"),
	("Singapore(SGSIN100)","MEP","\\dcasi400dat01\SGSIN100\DAT01\MEP\MEP1-2"),
	("Singapore(SGSIN100)","MEP","\\dcasi400dat01\SGSIN100\DAT01\MEP\MEP1-3"),
	("Singapore(SGSIN100)","MEP","\\dcasi400dat01\SGSIN100\DAT01\MEP\MEP1-4"),
	("Singapore(SGSIN100)","MEP","\\dcasi400dat01\SGSIN100\DAT01\MEP\MEP2-1"),
	("Singapore(SGSIN100)","MEP","\\dcasi400dat01\SGSIN100\DAT01\MEP\MEP2-2"),
	("Singapore(SGSIN100)","MEP","\\dcasi400dat01\SGSIN100\DAT02\MEP\MEP3-1"),
	("Singapore(SGSIN100)","MEP","\\dcasi400dat01\SGSIN100\DAT02\MEP\MEP4-1"),
	("Singapore(SGSIN100)","MEP","\\dcasi400dat01\SGSIN100\DAT02\MEP\MEP4-2"),
	("Singapore(SGSIN100)","MEP","\\dcasi400dat01\SGSIN100\DAT02\MEP\MEP5-1"),
	("Singapore(SGSIN100)","CS","\\dcasi400dat01\SGSIN100\DAT03\CS\CS-1"),
	("Singapore(SGSIN100)","CS","\\dcasi400dat01\SGSIN100\DAT03\CS\CS-2"),
	("Singapore(SGSIN100)","CS","\\dcasi400dat01\SGSIN100\DAT03\CS\CS-3"),
	("Singapore(SGSIN100)","CS","\\dcasi400dat01\SGSIN100\DAT04\CS\CS-4"),
	("Singapore(SGSIN100)","CS","\\dcasi400dat01\SGSIN100\DAT04\CS\CS-5"),
	("Singapore(SGSIN100)","CS","\\dcasi400dat01\SGSIN100\DAT04\CS\CS-6"),
	("Singapore(SGSIN100)","AP","\\dcasi400dat01\SGSIN100\DAT05\AP\AP-1"),
	("Singapore(SGSIN100)","AP","\\dcasi400dat01\SGSIN100\DAT05\AP\AP-2"),
	("Singapore(SGSIN100)","BD","\\dcasi400dat01\SGSIN100\DAT05\BD\BD-1"),
	("Singapore(SGSIN100)","COMMS","\\dcasi400dat01\SGSIN100\DAT05\COMMS"),
	("Singapore(SGSIN100)","PMC","\\dcasi400dat01\SGSIN100\DAT05\PMC\PMC-1"),
	("Singapore(SGSIN100)","POWER","\\dcasi400dat01\SGSIN100\DAT05\Power\Power-1"),
	("Singapore(SGSIN100)","BD","\\dcasi400dat01\SGSIN100\DAT06\DATA1\Data\BD"),
	("Singapore(SGSIN100)","Finance","\\dcasi400dat01\SGSIN100\DAT06\DATA2\DEPT\acc"),
	("Singapore(SGSIN100)","ADMIN","\\dcasi400dat01\SGSIN100\DAT06\DATA2\DEPT\adm"),
	("Singapore(SGSIN100)","AP","\\dcasi400dat01\SGSIN100\DAT06\DATA2\DEPT\ap"),
	("Singapore(SGSIN100)","BD","\\dcasi400dat01\SGSIN100\DAT06\DATA2\DEPT\bdme"),
	("Singapore(SGSIN100)","BIM","\\dcasi400dat01\SGSIN100\DAT06\DATA2\DEPT\bim"),
	("Singapore(SGSIN100)","CAD","\\dcasi400dat01\SGSIN100\DAT06\DATA2\DEPT\cad"),
	("Singapore(SGSIN100)","CFD","\\dcasi400dat01\SGSIN100\DAT06\DATA2\DEPT\cfd"),
	("Singapore(SGSIN100)","SHEQ","\\dcasi400dat01\SGSIN100\DAT06\DATA2\DEPT\compliance"),
	("Singapore(SGSIN100)","CS","\\dcasi400dat01\SGSIN100\DAT06\DATA2\DEPT\cs"),
	("Singapore(SGSIN100)","HR","\\dcasi400dat01\SGSIN100\DAT06\DATA2\DEPT\hr"),
	("Singapore(SGSIN100)","HR","\\dcasi400dat01\SGSIN100\DAT06\DATA2\DEPT\hr - advertisement"),
	("Singapore(SGSIN100)","IBT","\\dcasi400dat01\SGSIN100\DAT06\DATA2\DEPT\ibt"),
	("Singapore(SGSIN100)","OPS","\\dcasi400dat01\SGSIN100\DAT06\DATA2\DEPT\inter-division"),
	("Singapore(SGSIN100)","SHEQ","\\dcasi400dat01\SGSIN100\DAT06\DATA2\DEPT\iso"),
	("Singapore(SGSIN100)","MEP","\\dcasi400dat01\SGSIN100\DAT06\DATA2\DEPT\lighting"),
	("Singapore(SGSIN100)","MEP","\\dcasi400dat01\SGSIN100\DAT06\DATA2\DEPT\mep mgmt"),
	("Singapore(SGSIN100)","MEP","\\dcasi400dat01\SGSIN100\DAT06\DATA2\DEPT\mep1a"),
	("Singapore(SGSIN100)","MEP","\\dcasi400dat01\SGSIN100\DAT06\DATA2\DEPT\mep1b"),
	("Singapore(SGSIN100)","MEP","\\dcasi400dat01\SGSIN100\DAT06\DATA2\DEPT\mep2"),
	("Singapore(SGSIN100)","MEP","\\dcasi400dat01\SGSIN100\DAT06\DATA2\DEPT\mep3"),
	("Singapore(SGSIN100)","MEP","\\dcasi400dat01\SGSIN100\DAT06\DATA2\DEPT\mep4"),
	("Singapore(SGSIN100)","MEP","\\dcasi400dat01\SGSIN100\DAT06\DATA2\DEPT\mep5"),
	("Singapore(SGSIN100)","IT","\\dcasi400dat01\SGSIN100\DAT06\DATA2\DEPT\mis"),
	("Singapore(SGSIN100)","ADMIN","\\dcasi400dat01\SGSIN100\DAT06\DATA2\DEPT\pbploffice"),
	("Singapore(SGSIN100)","PMC","\\dcasi400dat01\SGSIN100\DAT06\DATA2\DEPT\pmc"),
	("Singapore(SGSIN100)","POWER","\\dcasi400dat01\SGSIN100\DAT06\DATA2\DEPT\power"),
	("Singapore(SGSIN100)","ADMIN","\\dcasi400dat01\SGSIN100\DAT06\DATA2\PUBLIC\0.SCAN"),
	("Singapore(SGSIN100)","ADMIN","\\dcasi400dat01\SGSIN100\DAT06\DATA2\PUBLIC\common"),
	("Singapore(SGSIN100)","IT","\\dcasi400dat01\SGSIN100\FIL05\ARCHIVE"),
	("Singapore(SGSIN100)","IT","\\dcasi400dat01\SGSIN100\FIL05\CFD_FILES"),
	("Singapore(SGSIN100)","IT","\\dcasi400dat01\SGSIN100\FIL05\SOFTWARES"),
	("Hongkong(HKKWN200)","Account","\\hkkwn200dat01\Accounts"),
	("Hongkong(HKKWN200)","Mutli-BU","\\hkkwn200dat01\AsiaCommonDrive"),
	("Hongkong(HKKWN200)","STR","\\hkkwn200dat01\CAD"),
	("Hongkong(HKKWN200)","Legal","\\hkkwn200dat01\China Region Litigation"),
	("Hongkong(HKKWN200)","Mutli-BU","\\hkkwn200dat01\Company"),
	("Hongkong(HKKWN200)","Legal","\\hkkwn200dat01\ComSecHKO"),
	("Hongkong(HKKWN200)","Mutli-BU","\\hkkwn200dat01\FujiXeroScan"),
	("Hongkong(HKKWN200)","HR","\\hkkwn200dat01\HR"),
	("Hongkong(HKKWN200)","HR","\\hkkwn200dat01\HR Payroll and Employee Benefits"),
	("Hongkong(HKKWN200)","Legal","\\hkkwn200dat01\Legal"),
	("Hongkong(HKKWN200)","Mutli-BU","\\hkkwn200dat01\ProjectMail"),
	("Hongkong(HKKWN200)","Legal","\\hkkwn200dat01\SEA Litigation"),
	("Hongkong(HKKWN200)","Mutli-BU","\\hkkwn200dat01\SupportServices"),
	("Hongkong(HKKWN200)","MEP3","\\hkkwn200dat02\BS"),
	("Hongkong(HKKWN200)","MEP3","\\hkkwn200dat02\BS\Projects"),
	("Hongkong(HKKWN200)","STR","\\Hkkwn200dat03\cs"),
	("Hongkong(HKKWN200)","STR","\\Hkkwn200dat03\cs\Projects"),
	("Hongkong(HKKWN200)","HR","\\hkkwn200dat05\ASIA_HQ"),
	("Hongkong(HKKWN200)","HR","\\hkkwn200dat05\AsiaStaffingNumber"),
	("Hongkong(HKKWN200)","IT_CorporateServices","\\hkkwn200dat05\Atos-OSS-Doc"),
	("Hongkong(HKKWN200)","IT_CorporateServices","\\hkkwn200dat05\Autodesk_Report"),
	("Hongkong(HKKWN200)","HR","\\hkkwn200dat05\colorscan"),
	("Hongkong(HKKWN200)","Mutli-BU","\\hkkwn200dat05\CV"),
	("Hongkong(HKKWN200)","F&A","\\hkkwn200dat05\FinanceDrive"),
	("Hongkong(HKKWN200)","F&A","\\hkkwn200dat05\FinanceDrive\grp"),
	("Hongkong(HKKWN200)","F&A","\\corp.pbwan.net\CN\Transition\GGZ\Account\"),
	("Hongkong(HKKWN200)","F&A","\\corp.pbwan.net\CN\Transition\BJS\FA\"),
	("Hongkong(HKKWN200)","F&A","\\corp.pbwan.net\CN\Transition\SGH\Co-Work\Finance Share Service"),
	("Hongkong(HKKWN200)","Marketing","\\hkkwn200dat05\HKGBD_BD"),
	("Hongkong(HKKWN200)","HR","\\hkkwn200dat05\HR MDA report"),
	("Hongkong(HKKWN200)","HR","\\hkkwn200dat05\HR-MKT"),
	("Hongkong(HKKWN200)","IT_CorporateServices","\\hkkwn200dat05\Infrastructure"),
	("Hongkong(HKKWN200)","IT_CorporateServices","\\hkkwn200dat05\IT Share Folder"),
	("Hongkong(HKKWN200)","HR","\\hkkwn200dat05\JDP"),
	("Hongkong(HKKWN200)","HR","\\hkkwn200dat05\Payroll"),
	("Hongkong(HKKWN200)","STR","\\hkkwn200dat05\PBIL"),
	("Hongkong(HKKWN200)","HR","\\hkkwn200dat05\personneldrive"),
	("Hongkong(HKKWN200)","Mutli-BU","\\hkkwn200dat05\PRJ_Support"),
	("Hongkong(HKKWN200)","Mutli-BU","\\hkkwn200dat05\PRJ4"),
	("Hongkong(HKKWN200)","Project-Asia","\\hkkwn200dat05\Project-Asia"),
	("Hongkong(HKKWN200)","Proposal","\\hkkwn200dat05\Proposal"),
	("Hongkong(HKKWN200)","HR","\\hkkwn200dat05\Recruitment"),
	("Hongkong(HKKWN200)","Mutli-BU","\\hkkwn200dat05\Return Company Property"),
	("Hongkong(HKKWN200)","IT_CorporateServices","\\hkkwn200dat05\SNAP"),
	("Hongkong(HKKWN200)","Mutli-BU","\\hkkwn200dat05\Staff_photo"),
	("Hongkong(HKKWN200)","BIM","\\hkkwn200dat07\2534047"),
	("Hongkong(HKKWN200)","BIM","\\hkkwn200dat07\BIM\BIM"),
	("Hongkong(HKKWN200)","BIM","\\hkkwn200dat07\BIM\BIM\Projects"),
	("Hongkong(HKKWN200)","BIM","\\hkkwn200dat07\BIM_training"),
	("Hongkong(HKKWN200)","BIM","\\hkkwn200dat07\BIM2"),
	("Hongkong(HKKWN200)","BIM","\\hkkwn200dat07\Shared"),
	("Hongkong(HKKWN200)","STR","\\hkkwn200dat08\GEO"),
	("Hongkong(HKKWN200)","PMCM","\\hkkwn200dat12\PMCM"),
	("Hongkong(HKKWN200)","PMCM","\\hkkwn200dat12\PMCM\D8"),
	("Hongkong(HKKWN200)","TnI","\\hkkwn200dat13\T&I"),
	("Hongkong(HKKWN200)","TnI","\\hkkwn200dat13\T&I\Projects"),
	("Hongkong(HKKWN200)","TnI","\\hkkwn200dat72\T&I"),
	("Hongkong(HKKWN200)","MEP1","\\hkkwn200dat14\MEP1"),
	("Hongkong(HKKWN200)","MEP1","\\hkkwn200dat14\MEP1\Projects"),
	("Hongkong(HKKWN200)","MEP1","\\hkkwn200dat71\MEP1"),
	("Hongkong(HKKWN200)","MEP2","\\Hkkwn200dat15\mep2"),
	("Hongkong(HKKWN200)","MEP2","\\Hkkwn200dat15\mep2\Projects"),
	("Hongkong(HKKWN200)","SDE","\\Hkkwn200dat16\SDE"),
	("Hongkong(HKKWN200)","SDE","\\Hkkwn200dat16\SDE\7. Projects"),
	("Hongkong(HKKWN200)","MEP4","\\hkkwn200dat17\Healthcare"),
	("Hongkong(HKKWN200)","MEP4","\\Hkkwn200dat17\MEP4"),
	("Hongkong(HKKWN200)","MEP4","\\Hkkwn200dat17\MEP4\Projects"),
	("Hongkong(HKKWN200)","Mutli-BU","\\hkkwn200dat18\Business Lines\H300 Management"),
	("Hongkong(HKKWN200)","H315QualityInspection","\\hkkwn200dat18\Business Lines\H315 Quality Inspection"),
	("Hongkong(HKKWN200)","Mutli-BU","\\hkkwn200dat18\Business Lines\MEP"),
	("Hongkong(HKKWN200)","Mutli-BU","\\hkkwn200dat18\Business Lines\Project Account"),
	("Hongkong(HKKWN200)","Mutli-BU","\\hkkwn200dat18\Business Lines\Project Sec"),
	("Hongkong(HKKWN200)","Admin_CorporateServices","\\hkkwn200dat18\Corporate Services\Admin"),
	("Hongkong(HKKWN200)","Admin_CorporateServices","\\hkkwn200dat18\Corporate Services\Admin\AD-0"),
	("Hongkong(HKKWN200)","HR_CorporateServices","\\hkkwn200dat18\Corporate Services\HR"),
	("Hongkong(HKKWN200)","IT_CorporateServices","\\hkkwn200dat18\Corporate Services\IT"),
	("Hongkong(HKKWN200)","Legal_CorporateServices","\\hkkwn200dat18\Corporate Services\Legal"),
	("Hongkong(HKKWN200)","Marketing_CorporateServices","\\hkkwn200dat18\Corporate Services\Marketing"),
	("Hongkong(HKKWN200)","Safety & Quality Assurance","\\hkkwn200dat05\PRJ_Support\QSD"),
	("Hongkong(HKKWN200)","BDRES","\\hkkwn200dat05\BDRES"),
	("Unknown","Unknown","") #don't modify this line
	)
#>

#Get-ACLReport 函數用於檢索並處理指定資料夾的 ACL 資訊，包括：
	#資料夾擁有者	
	#使用者和群組的訪問權限
	#身份是 Active Directory (AD) 帳戶還是本地帳戶
	#群組成員（針對 AD 群組）
	#使用者或群組成員的顯示名稱
function Get-ACLReport ($folder) {
        $report = @()
        try
        {
        $accessresult=(get-acl -LiteralPath $folder).access
        $folderowner=(get-acl -LiteralPath $folder).Owner

        $Identitys=$accessresult.IdentityReference
        $FileSystemRights=$accessresult.FileSystemRights
        #$accesscontroltypes=$accessresult.AccessControlType

        $Identity=""
        #$accesscontroltype=""
        for($i=0;$i -lt $Identitys.count;$i++)
            {
                $Identity=$Identitys[$i]

                $FileSystemRight=""
                for($j=0;$j -lt $FileSystemRights[$i].count;$j++)
                {
                    $FileSystemRight=$FileSystemRight+";"+$FileSystemRights[$i]
                }
                #$accesscontroltype=$accesscontroltypes[$i]
                $FileSystemRight=$FileSystemRight.Substring(1,$FileSystemRight.Length-1)

                #New-Object psobject -Property @{'Folder path'=$folder;'Identity'=$Identity;'AccessRight'=$FileSystemRight;'Control Type'=$accesscontroltype}

                $isADAccount=$false
                if($Identity.Value.Substring(0,5) -eq "CORP\")
                {
                    $IdentityName=$Identity.Value -replace "CORP\\"
                    $IdentityObject=Get-ADObject -Filter "SamAccountName -eq '$IdentityName'"

                    if(($IdentityObject.objectClass -ne "user") -and ($IdentityName -ne "Domain Users")) #security group, ignore "Domain Users"
                    {
                        $GroupMembers=Get-ADGroupMember $IdentityName -recursive | select-object SamAccountName
					}
					else #AD user account or "domain users" group
					{
						$GroupMembers=$IdentityName
					}
                    $isADAccount=$true
                }
                else #local account
                {
                    $GroupMembers = $Identity.Value
                }

				foreach($GroupMember in $GroupMembers)
				{
					$obj = New-Object -TypeName psobject
                    #$DisplayName=""
                    if($isADAccount -eq $true)
                    {
                        #$ADO=Get-ADObject -Filter "SamAccountName -eq '$IdentityName'"
                        if($IdentityObject.objectClass -ne "group") #user account
                        {
                            $GM="-"
                            $DisplayName=(get-aduser $GroupMember -Properties *).Displayname
                        }
                        else #AD group(include domain users)
                        {
                            if($GroupMember -eq "Domain Users")
                            {
                                $GM=$GroupMember
                                $DisplayName=$GroupMember
                            }
                            else
                            {
                                $GM=$GroupMember.SamAccountName
                                $DisplayName=(get-aduser $GroupMember.SamAccountName -Properties *).Displayname
                            }
                        }
					    $ID=$IdentityName
                    }
                    else #local id
                    {
					    $ID=$GroupMember
                        $GM="-"
                        $DisplayName = $GroupMember
                    }                  
					$obj | Add-Member -MemberType NoteProperty -Name FolderPath -Value $folder
					$obj | Add-Member -MemberType NoteProperty -Name Identity -Value $ID
				    $obj | Add-Member -MemberType NoteProperty -Name Member -Value $GM
				    $obj | Add-Member -MemberType NoteProperty -Name DisplayName -Value $DisplayName
					$obj | Add-Member -MemberType NoteProperty -Name AccessRight -Value $FileSystemRight
					#$obj | Add-Member -MemberType NoteProperty -Name AccessType -Value $accesscontroltype
					$obj | Add-Member -MemberType NoteProperty -Name Owner -Value $folderowner

					$report +=$obj
				}
            }
        }
        catch
        {
            $obj = New-Object -TypeName psobject

            $obj | Add-Member -MemberType NoteProperty -Name FolderPath -Value $folder
            $obj | Add-Member -MemberType NoteProperty -Name Identity -Value "error"
            $obj | Add-Member -MemberType NoteProperty -Name Member -Value "error"
            $obj | Add-Member -MemberType NoteProperty -Name AccessRight -Value "error"
            #$obj | Add-Member -MemberType NoteProperty -Name AccessType -Value "error"
            $obj | Add-Member -MemberType NoteProperty -Name Owner -Value "error"

            $report +=$obj
        }

        return $report
		}


<#
function RefreshTeam
{
    Param([string] $Office
    )
    $TeamList.Items.Clear()

	for($o=0;$o -lt $ACLFolders.Length-1;$o++)
	{
		if($ACLFolders[$o][0] -eq $Office)
		{
			$teamaddedcount=0
			foreach($teamitem in $TeamList.Items)
			{
				if($teamitem -eq $ACLFolders[$o][1])
				{
					$teamaddedcount++
				}
			}

			if($teamaddedcount -eq 0)
			{
				$TeamList.items.Add($ACLFolders[$o][1])
			}
		}

	}

}
#>

#### create GUI#######
Add-Type -AssemblyName System.Windows.Forms
$PowerShellForms = New-Object system.Windows.Forms.Form
$PowerShellForms.Text=$FormTitle
$PowerShellForms.Size = New-Object System.Drawing.Size(470,200)
$PowerShellForms.MinimizeBox = $True
$PowerShellForms.MaximizeBox = $False
$PowerShellForms.FormBorderStyle = [System.Windows.Forms.FormBorderStyle]::FixedDialog

$PowerShellForms.SizeGripStyle = "Hide"
$PowerShellForms.StartPosition = "CenterScreen"
$PowerShellForms.Font = New-Object System.Drawing.Font("Tahoma",10,[System.Drawing.FontStyle]::Regular)
<#
$Lboffice = New-Object System.Windows.Forms.Label
$Lboffice.Text = "Office(*)"
$Lboffice.AutoSize = $True
$Lboffice.BackColor = "Transparent"
$Lboffice.Font = New-Object System.Drawing.Font("Tahoma",10,[System.Drawing.FontStyle]::Regular)
$Lboffice.ForeColor = "Black"
$Lboffice.Location = New-Object System.Drawing.Point(10,25)

$OfficeList = New-Object System.Windows.Forms.ListBox
$OfficeList.Font = New-Object System.Drawing.Font("Tahoma",10,[System.Drawing.FontStyle]::Regular)
$OfficeList.Location = New-Object System.Drawing.Point(90,25)
$OfficeList.Size = New-Object System.Drawing.Size(380,20)
$OfficeList.Height = 140
for($o=0;$o -lt $ACLFolders.Length;$o++)
{
	$officeaddedcount=0
	foreach($officeitem in $OfficeList.Items)
	{
		if($officeitem -eq $ACLFolders[$o][0])
		{
			$officeaddedcount++
		}
	}

	if($officeaddedcount -eq 0)
	{
		$OfficeList.items.Add($ACLFolders[$o][0])
	}
}

$OfficeList.TabIndex = 1
$OfficeList.Add_SelectedIndexChanged(
    {
        RefreshTeam -Office $OfficeList.SelectedItem.ToString()
    }
)

$LbTeam = New-Object System.Windows.Forms.Label
$LbTeam.Text = "Team(*)"
$LbTeam.AutoSize = $True
$LbTeam.BackColor = "Transparent"
$LbTeam.Font = New-Object System.Drawing.Font("Tahoma",10,[System.Drawing.FontStyle]::Regular)
$LbTeam.ForeColor = "Black"
$LbTeam.Location = New-Object System.Drawing.Point(10,175)

$TeamList = New-Object System.Windows.Forms.ListBox
$TeamList.Font = New-Object System.Drawing.Font("Tahoma",10,[System.Drawing.FontStyle]::Regular)
$TeamList.Location = New-Object System.Drawing.Point(90,175)
$TeamList.Size = New-Object System.Drawing.Size(380,100)
$TeamList.Height=140
$TeamList.TabIndex = 2
#>

<#
$LbDepthDescription = New-Object System.Windows.Forms.Label
$LbDepthDescription.Text = "1:first level sub folders;*: all sub folders"
$LbDepthDescription.AutoSize = $True
$LbDepthDescription.BackColor = "Transparent"
$LbDepthDescription.Font = New-Object System.Drawing.Font("Tahoma",10,[System.Drawing.FontStyle]::Regular)
$LbDepthDescription.ForeColor = "Black"
$LbDepthDescription.Location = New-Object System.Drawing.Point(230,365)
#>
<#
$LbAllTeam = New-Object System.Windows.Forms.Label
$LbAllTeam.Text = "All teams in this office"
$LbAllTeam.AutoSize = $True
$LbAllTeam.BackColor = "Transparent"
$LbAllTeam.Font = New-Object System.Drawing.Font("Tahoma",10,[System.Drawing.FontStyle]::Regular)
$LbAllTeam.ForeColor = "Black"
$LbAllTeam.Location = New-Object System.Drawing.Point(315,310)

$CbAllTeam = New-Object System.Windows.Forms.CheckBox
$CbAllTeam.Font = New-Object System.Drawing.Font("Tahoma",10,[System.Drawing.FontStyle]::Regular)
$CbAllTeam.Location = New-Object System.Drawing.Point(456,310)
$CbAllTeam.Size = New-Object System.Drawing.Size(40,40)
$CbAllTeam.Height=20
#Group name length limit is 64, group name example: 
$CbAllTeam.TabIndex = 3
$CbAllTeam.Add_Click(
    {
        $TeamList.Enabled = !($CbAllTeam.Checked)
		
		#[System.Windows.Forms.Messagebox]::Show($CbAllTeam.Checked)
    }
)
#>
$LbTargetPath = New-Object System.Windows.Forms.Label
$LbTargetPath.Text = "Target folder path:"
$LbTargetPath.AutoSize = $True
$LbTargetPath.BackColor = "Transparent"
$LbTargetPath.Font = New-Object System.Drawing.Font("Tahoma",10,[System.Drawing.FontStyle]::Regular)
$LbTargetPath.ForeColor = "Black"
$LbTargetPath.Location = New-Object System.Drawing.Point(10,10)

$TbTargetPath = New-Object System.Windows.Forms.TextBox
$TbTargetPath.Font = New-Object System.Drawing.Font("Tahoma",10,[System.Drawing.FontStyle]::Regular)
$TbTargetPath.Location = New-Object System.Drawing.Point(10,40)
$TbTargetPath.Size = New-Object System.Drawing.Size(430,30)
$TbTargetPath.Height=20
$TbTargetPath.Text=""
$TbTargetPath.TabIndex = 3

$LbDepth = New-Object System.Windows.Forms.Label
$LbDepth.Text = "Sub folder level"
$LbDepth.AutoSize = $True
$LbDepth.BackColor = "Transparent"
$LbDepth.Font = New-Object System.Drawing.Font("Tahoma",10,[System.Drawing.FontStyle]::Regular)
$LbDepth.ForeColor = "Black"
$LbDepth.Location = New-Object System.Drawing.Point(300,70)

$TbDepth = New-Object System.Windows.Forms.TextBox
$TbDepth.Font = New-Object System.Drawing.Font("Tahoma",10,[System.Drawing.FontStyle]::Regular)
$TbDepth.Location = New-Object System.Drawing.Point(405,70)
$TbDepth.Size = New-Object System.Drawing.Size(30,30)
$TbDepth.Height=20
$TbDepth.Text="3"
$TbDepth.TabIndex = 3

$BtnOK = New-Object System.Windows.Forms.Button
$btnOK.Size = New-Object System.Drawing.Size(60,30)
$BtnOK.Location = New-Object System.Drawing.Size(380,110)
$BtnOK.Text = "Get it!"
$BtnOK.Font = New-Object System.Drawing.Font("Tahoma",10,[System.Drawing.FontStyle]::Regular)
$BtnOK.Add_Click(
    {
		$today=Get-Date -Format "yyyyMMddHHmm"
		<#
		$SelectedOffice=""
		$SelectedTeam=""
        if($TabControl1.SelectedIndex -eq 0)
        {
            if((($CbAllTeam.Checked -eq $true) -and ($null -ne $OfficeList.SelectedItem)) -or (($null -ne $OfficeList.SelectedItem) -and ($null -ne $TeamList.SelectedItem)))
            {
			    $SelectedOffice=$OfficeList.SelectedItem
			    $SelectedTeam=$TeamList.SelectedItem
		    }
            else
            {
                [System.Windows.Forms.Messagebox]::Show("Following items are mandatory: `nOffice`nTeam")
            }
        }

        elseif($TabControl1.SelectedIndex -eq 1)
        {

            if($TbTargetPath.Text.Length -gt 3)
            {
			    #$SelectedOffice="Unknown"
			    #$SelectedTeam="Unknown"
				$ACLFolders[$ACLFolders.Length-1][2]=$TbTargetPath.Text
            }
            else
            {
                [System.Windows.Forms.Messagebox]::Show("Target path cannot be blank.")
            }
        }
		#>
		<#
		if(($SelectedOffice -ne "") -and ($SelectedTeam -ne ""))
		{
			for($n=0;$n -lt $ACLFolders.Length;$n++)
			{
				if($ACLFolders[$n][0] -eq $SelectedOffice)
				{#>
					$foldername=$TbTargetPath.Text.split('\')
					$foldername=$foldername[$foldername.length-1]
					$reportname = "ACLReport_" + $foldername + "_" + $today + ".csv"
					$reportname = Join-Path $ReportPath $reportname

					<#
					if((($CbAllTeam.Checked -eq $false) -and ($SelectedTeam -eq $ACLFolders[$n][1])) -or ($CbAllTeam.Checked -eq $true))
					{
					#>
						$R=Get-ACLReport $TbTargetPath.Text
						$R | export-csv $reportname -NoTypeInformation -Force -Append -Encoding UTF8

						if($TbDepth.Text -eq "*")
						{
							$SubFolders=Get-ChildItem -path $TbTargetPath.Text -Directory -Force -Recurse | Select-Object FullName
						}
						else
						{
							$Depth=$TbDepth.Text-1
							if($Depth -lt 1)
							{$Depth = 1}

							$SubFolders=Get-ChildItem -path $TbTargetPath.Text -Depth $Depth -Directory -Force | Select-Object FullName
						}
						#$SubFolders=Get-ChildItem $ACLFolders[$n][2] -Directory
						foreach($SubFolder in $SubFolders)
						{
							$R=Get-ACLReport $SubFolder.FullName
							$R | export-csv $reportname -NoTypeInformation -Force -Append -Encoding UTF8
						}						
					#}
				#}
			#}
			[System.Windows.Forms.Messagebox]::Show("Done~ `nYou can find your report under:`n" + $ReportPath)
		#}
    }
)

<#
$TabControl1 = New-Object System.Windows.Forms.TabControl
$TabControl1.Location = New-Object System.Drawing.Point(0,0)
$TabControl1.Width=$PowerShellForms.Size.Width
$TabControl1.Height=$PowerShellForms.Size.Height-130

$tabPage1=New-Object System.Windows.Forms.TabPage
$tabPage1.text = "Be Happy"
$tabPage1.TabIndex=0
$tabPage1.Controls.AddRange(($LbTargetPath,$TbTargetPath))

$tabPage2=New-Object System.Windows.Forms.TabPage
$tabPage2.text = "Have Fun"
$tabPage2.TabIndex=1
$tabPage2.Controls.AddRange(($Lboffice, $OfficeList, $LbTeam, $TeamList, $LbAllTeam,$CbAllTeam))
#>

#$TabControl1.Controls.Add($tabPage1)
#$TabControl1.Controls.Add($tabPage2)
#$TabControl1.Size=$PowerShellForms.Size

$PowerShellForms.Controls.Add($LbTargetPath)
$PowerShellForms.Controls.Add($TbTargetPath)
$PowerShellForms.Controls.Add($LbDepth)
$PowerShellForms.Controls.Add($TbDepth)
$PowerShellForms.Controls.Add($BtnOK)

$PowerShellForms.ShowDialog()
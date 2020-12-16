<#
###################################################
### AD User creation in a Hybrid Environmnent  ####
###################################################

<#

.Author - Parag Manappetty
.Version - 1.1
.Date - 18 April 2019


.Requirement:

1) Active Directory Module
2) Presentation Framework

#>



param(
    [parameter(mandatory=$true)]
    [string]
    $FullName,
    
    [parameter(mandatory=$true)]
    [validateSet("Vancouver","Quebec","Onatrio","California")]
    [string]
    $Office,

    [parameter(mandatory=$true)]
    [string]
    $Manager,

    [parameter(mandatory=$true)]
    [string]
    $MembershipsFrom,

    [parameter(mandatory=$true)]
    [string]
    $Title

)

$arrayDN = $FullName.trim().tolower().Split()


<# --------------- 


                    Creates an AD Obect 
                    
              
                    
                                         -------------------- #>


if ($arrayDN.Count -eq 2){

        $fn = $arrayDN[0].Substring(0,1).ToUpper() + $arrayDN[0].Substring(1)
        $ln = $arrayDN[1].Substring(0,1).ToUpper() + $arrayDN[1].Substring(1)
        $fullName = $fn + " " + $ln
        $username = $ln.Substring(0,1).ToUpper() + $ln.Substring(1) + $fn.Substring(0,1).ToUpper()
        $email = $username + "@domain.com"
        $password = "123456" | ConvertTo-SecureString -AsPlainText -Force

    if((Get-ADUser -Filter {sAMAccountName -eq $username}) -eq $null){
 
        Switch($Office){
            Vancouver{$OU = "company.local/Users/Vancouver"}
            Quebec{$OU = "company.local/Users/Quebec"}
            Ontario{$OU = "company.local/Users/Ontario"}
            California{$OU = "company.local/Users/California"}
            Default{"Conditions unmatch"}
            }

         New-RemoteMailbox -Name $fullName -Password $password -UserPrincipalName $email -DisplayName $FullName -FirstName $fn -LastName $ln -OnPremisesOrganizationalUnit $OU -ResetPasswordOnNextLogon $true
        "`n$username does not exist in AD, and the Object is created."
    }

    else{
        
        $username = $ln.Substring(0,1).ToUpper() + $ln.Substring(1) + $fn.Substring(0,1).ToUpper() + $fn.Substring(1,1)
        $email = $username + "@domain.com"

        Switch($Office){
            Vancouver{$OU = "company.local/Users/Vancouver"}
            Quebec{$OU = "company.local/Users/Quebec"}
            Ontario{$OU = "company.local/Users/Ontario"}
            California{$OU = "company.local/Users/California"}
            Default{"Conditions unmatch"}
            }

         New-RemoteMailbox -Name $fullName -Password $password -UserPrincipalName $email -DisplayName $FullName -FirstName $fn -LastName $ln -OnPremisesOrganizationalUnit $OU -ResetPasswordOnNextLogon $true
        "`n$username does not exist in AD, and the Object is created."
    }

}

elseif ($arrayDN.Count -eq 3){

function check-name{

        Write-Host "`nSelect which one is the first name"
            "1) "+ $arrayDN[0]
            "2) "+ $arrayDN[1]
            "3) "+ $arrayDN[2]

    $fn = Read-Host
    Switch($fn){
        1{$fn = $arrayDN[0]}
        2{$fn = $arrayDN[1]}
        3{$fn = $arrayDN[2]}
    }
    
    "`nnow select last name"

    $ln = Read-Host
    Switch($ln){
        1{$ln = $arrayDN[0]}
        2{$ln = $arrayDN[1]}
        3{$ln = $arrayDN[2]}
    }

    "First Name: $fn"
    "Last Name : $ln"
         
    }
    
    $username = $ln.Substring(0,1).ToUpper() + $ln.Substring(1) + $fn.Substring(0,1).ToUpper()
    $email = $username + "@domain.com"

    $input = [System.Windows.MessageBox]::Show("`nFirst name: $fn `nLast name: $ln `nUsername: ","Confirm:","YesNo")   
    
    while($input -eq "No"){
        check-name
        $input = [System.Windows.MessageBox]::Show("`nFirst name: $fn `nLast name: $ln `nUsername: ","Confirm:","YesNo")   
    }

        Switch($Office){
            Vancouver{$OU = "company.local/Users/Vancouver"}
            Quebec{$OU = "company.local/Users/Quebec"}
            Ontario{$OU = "company.local/Users/Ontario"}
            California{$OU = "company.local/Users/California"}
            Default{"Conditions unmatch"}
            }

         New-RemoteMailbox -Name $FullName -Password $password -UserPrincipalName $email -DisplayName $FullName -FirstName $fn -LastName $ln -OnPremisesOrganizationalUnit $OU -ResetPasswordOnNextLogon $true
        "`n$username does not exist in AD, and the Object is created."


    
}

else{
    "Use normal procedure to create the AD Account"
}

Start-Sleep -Seconds 300



<# ------------ 



                    Adding details to the AD User 
                    
                    
                    
                    
                                                            ---------------- #>

# Parameter set the below details <-

$company = "Company Name"
$lscript = (Get-ADUser $MembershipsFrom -Properties *).scriptpath
$depart = (Get-ADUser $MembershipsFrom -Properties *).department

# ->

 # set by the para and if above conditions pass


Switch($Office){
            # Vancouver LOCATION
            Vancouver
            {
            
            $homefolder = "\\company.local\$Office\ShareDrives\$username"

            $officeLocation = "Office";
            $street = "123";
            $city = "Vancouver";
            $State = "BC";
            $zip = "123456";
            $country = "CA";
            break;
            
            }
            # Quebec LOCATION
            Quebec
            {

            $homefolder = "\\company.local\$Office\ShareDrives\$username"
            
            $officeLocation = "Office";
            $street = "123";
            $city = "Vancouver";
            $State = "BC";
            $zip = "123456";
            $country = "CA";
            break;
            
            }
            # Ontario LOCATION
            Ontario
            {

            $homefolder = "\\company.local\$Office\ShareDrives\$username"
            
            $officeLocation = "Office";
            $street = "123";
            $city = "Vancouver";
            $State = "BC";
            $zip = "123456";
            $country = "CA";
            break;

            }
            # California OFFICE
            California
            {

            $homefolder = "\\company.local\BPI\CaliforniaOffice\Users\$username"

            $officeLocation = "California Office";
            $street = "123";
            $city = "Vancouver";
            $State = "BC";
            $zip = "123456";
            $country = "CA";
            break;

            }

            Default{"Conditions unmatch"}
       }


Set-ADUser $username -Description $Title -Office $office -StreetAddress $street -City $city -State $State -PostalCode $zip -Country $country -ScriptPath $lscript -HomeDrive "H:" -HomeDirectory $homefolder -Title $Title -Department $depart -Company $company




# Transfering members

(Get-ADUser -Identity $MembershipsFrom -Properties memberof).memberof | %{Add-ADGroupMember -Identity $_ -Members $username -Confirm:$false}

# Create new home folder

New-Item -Path $homefolder -Type Directory -Force

# Sets the Manager

Set-ADUser $username -Manager $Manager

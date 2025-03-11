# added to github
# Description: This script will update user accounts in Active Directory based on an Excel spreadsheet.
# ERROR REPORTING ALL
Set-StrictMode -Version latest

#----------------------------------------------------------
# LOAD ASSEMBLIES AND MODULES
#----------------------------------------------------------
Try
{
  Import-Module ActiveDirectory -ErrorAction Stop       # RSAT or Domain Controller
  Import-Module ADSync -ErrorAction Stop                # Microsoft Azure Connect or similar sync tool (optional, comment out if not using cloud AD)
  Import-Module ImportExcel -ErrorAction Stop           # Found on the PowerShell Gallery, install with Install-Module ImportExcel
}
Catch
{
  Write-Error "A required module could not be imported: $($_.Exception.Message)`r`nScript unable to continue!`r`n"
  Exit 1
}

#----------------------------------------------------------
#STATIC VARIABLES
#----------------------------------------------------------
$path           = Split-Path -parent $MyInvocation.MyCommand.Definition
$userdb        = $path + ".\user-dump.xlsx"
$log            = $path + ".\update_ad_users.log"
# $addn           = (Get-ADDomain).DistinguishedName
# $dnsroot        = (Get-ADDomain).DNSRoot
$homeRoot       = "\\mnstco.net\resources\home"
$homeDriv       = "H:"
# $o365_groupname = "O365_User"
$phoneSysOU     = "OU=Phone System,DC=mnstco,DC=net"
$phoneSysGrp    = "CN=Mitel Users,OU=Phone System,DC=mnstco,DC=net"

#----------------------------------------------------------
#START FUNCTIONS
#----------------------------------------------------------
Function Start-Commands
{
  Update-Users
  Sync-Azure  
}

function New-RandomString {
  param(
      [int]$Length = 10
  )
  $characters = "abcdefghijklmnopqrstuvwxyzABCDEFGHIJKLMNOPQRSTUVWXYZ0123456789"
  $randomString = ""
  for ($i = 0; $i -lt $Length; $i++) {
      $randomIndex = Get-Random -Minimum 0 -Maximum $characters.Length
      $randomChar = $characters[$randomIndex]
      $randomString += $randomChar
  }
  return $randomString
}

function Convert-FirstLetterUpper {
  param (
      [string]$InputString
  )

  if ([string]::IsNullOrEmpty($InputString)) {
      return $InputString # Return the original string if it's null or empty
  }

  $firstChar = $InputString.Substring(0, 1).ToUpper()
  $remainingChars = $InputString.Substring(1)

  return $firstChar + $remainingChars
}

Function Sync-Azure

{
  Write-Host "[INFO]`t Force Sync Start.`r`n"
  "[INFO]`t Force Sync Start." | Out-File $log -Append
  Try
  {
    Start-ADSyncSyncCycle -PolicyType Delta
    Write-Host "[INFO]`t Force Sync Complete`r`n"
    "[INFO]`t Force Sync Complete" | Out-File $log -Append
  }
  Catch
  {
    Write-Error "Could not force AD Sync: $($_.Exception.Message)`r`n"
    "[ERROR]`t Could not force AD Sync: $($_.Exception.Message)" | Out-File $log -append
  }
}

Function Update-Users
{
  $i              = 1
  Import-Excel -Path $userdb | ForEach-Object {
    If ($_.SamAccountName -eq "") # skip this record if no account name specified 
    {
      Write-Error "SamAccountName cannot be blank, processing skipped for line $($i)`r`n"
      "[ERROR]`t SamAccountName cannot be blank, processing skipped for line $($i)" | Out-File $log -append
    }
    Else {
      $sam = $PSItem.SamAccountName
      Try {
        $exists = Get-ADUser -LDAPFilter "(sAMAccountName=$sam)"
      } Catch {
        Write-Warning "[WARNING]`t Error checking if user exists: $($_.Exception.Message)`r`n"
        "[WARNING]`t Error checking if user exists: $($_.Exception.Message)" | Out-File $log -append
      }

      # DisplayName, SamAccountName, GivenName, Initials, Surname, mail, ipPhone, employeeType, TelephoneNumber, Pager, Title, Department, StreetAddress, City, State, PostalCode, Company, Enabled, Update
      $params = @{} # Initialize an empty hashtable

      $specialAcct = $false

      # Add properties only if they are not null or empty
      If (-not [string]::IsNullOrEmpty($_.GivenName)) { $params.givenName = $_.GivenName.ToString() }
      If (-not [string]::IsNullOrEmpty($_.Initials)) { $params.initials = $_.Initials.ToString() }
      If (-not [string]::IsNullOrEmpty($_.Surname)) { $params.sn = $_.Surname.ToString() }
      If (-not [string]::IsNullOrEmpty($_.DisplayName)) { $params.displayName = $_.DisplayName.ToString() }
      If (-not [string]::IsNullOrEmpty($_.ipPhone)) { $params.ipPhone = $_.ipPhone.ToString() }
      If (-not [string]::IsNullOrEmpty($_.employeeType)) { $params.employeeType = $_.employeeType }
      If (-not [string]::IsNullOrEmpty($_.TelephoneNumber)) { $params.telephoneNumber = $_.TelephoneNumber.ToString()}
      If (-not [string]::IsNullOrEmpty($_.Pager)) { $params.pager = $_.Pager.ToString() }
      If (-not [string]::IsNullOrEmpty($_.Title)) { $params.title = $_.Title.ToString() 
                                                    $params.description = $_.Title.ToString() } # copy title to description if it exists
      If (-not [string]::IsNullOrEmpty($_.Department)) { $params.department = $_.Department.ToString() }
      If (-not [string]::IsNullOrEmpty($_.StreetAddress)) { $params.streetAddress = $_.StreetAddress.ToString() }
      If (-not [string]::IsNullOrEmpty($_.City)) { $params.l = $_.City.ToString() }
      If (-not [string]::IsNullOrEmpty($_.State)) { $params.st = $_.State.ToString() }
      If (-not [string]::IsNullOrEmpty($_.PostalCode)) { $params.postalCode = $_.PostalCode.ToString() }
      If (-not [string]::IsNullOrEmpty($_.Company)) { $params.company = $_.Company.ToString() }
      If (-not [string]::IsNullOrEmpty($_.Office)) { $params.physicalDeliveryOfficeName = $_.Office.ToString() }
      #If (-not [string]::IsNullOrEmpty($_.Enabled)) { $params.Enabled = $_.Enabled }

      if (($sam -like "phone_*") -or ($sam -like "fax_*")) { # format description differently for phone and fax accounts
        $specialAcct = $true
        $params.description = "Extension " + $params.telephoneNumber
      }

      If ((($PSItem.Update.ToLower()) -eq "yes") -and ($exists)) # Update user account information
      {
         If ($params.Count -gt 0) {
          Try {
              # Update User Account, only change info that won't impact login or email 
              #$params
              Set-ADUser -Identity $sam -Replace $params
              if (-not [string]::IsNullOrEmpty($_.TelephoneNumber)) { Add-ADGroupMember -Identity $phoneSysGrp -Members $sam }
              Write-Host "[INFO]`t updated user: $($sam)`r`n"
              "[INFO]`t Updated user: $($sam)" | Out-File $log -append
          }
          Catch {
              Write-Error "Couldn't update user $($sam): $($_.Exception.Message)`r`n"
              "[ERROR]`t Couldn't update user $($sam): $($_.Exception.Message)" | Out-File $log -append
          }
        } else {
          Write-Warning "[WARNING]`t No parameters to update for user $($sam).`r`n"
          "[WARNING]`t No parameters to update for user $($sam)." | Out-File $log -append
        }

      } # end ofupdate user
      elseif ((($_.Update.ToLower()) -eq "yes") -and (!$exists)) # user doesn't exist but import requested update
      {
        Write-Warning "[WARNING]`t User $($sam) does not exist and create not specified!`r`n"
        "[WARNING]`t User $($sam) does not exist and create not specified!" | Out-File $log -append
      }
      elseif ((($_.Update.ToLower()) -eq "create") -and ($exists)) # user exists but command is create
      {
        Write-Warning "[WARNING]`t User $($sam) already exists but create was specified!`r`n"
        "[WARNING]`t User $($sam) already exists but create was specified!" | Out-File $log -append
      }
      elseif ((($_.Update.ToLower()) -eq "create") -and (!$exists)) # user does not exist and command is create
      {
        # Create User Account, set as much info as possible

        Try 
        { 
          Write-Host "[INFO]`t Creating user : $($sam)`r`n"
          "[INFO]`t Creating user : $($sam)" | Out-File $log -append

          $params
          If (-not [string]::IsNullOrEmpty($_.Password)) { 
            $tempPass = $_.Password 
          }
          else {
            ( $tempPass = New-RandomString -Length 16 )
          }

          if ( $tempPass ) { Write-Host "[INFO] Temporary password set.`r`n"}

          $secPass = ConvertTo-SecureString -AsPlainText $tempPass -force

          # Write-Host "Password = $($_.Password) and setpass = $($setpass)"
          If (-not [string]::IsNullOrEmpty($_.Mail)) { $params.mail = $_.Mail.ToLower() }
          If (-not [string]::IsNullOrEmpty($_.Mail)) { $params.userPrincipalName = $_.Mail.ToLower() }

          if (($sam -like "phone_*") -or ($sam -like "fax_*")) { # special accounts no password expiration, no password change required
            New-ADUser $sam -AccountPassword $secPass -ChangePasswordAtLogon $false -CannotChangePassword $true -PasswordNeverExpires $true -Enabled $true
          } else {
            $params.HomeDirectory = $($homeRoot + "\" + $sam)
            $params.HomeDrive = $homeDriv

            New-ADUser $sam -ChangePasswordAtLogon $true -AccountPassword $secPass -Enabled $true
          }

          # update the rest of the attributes
          Get-ADUser -Identity $sam | Set-ADUser -Add $params
          
          # add user to the phone system group if they have a telephone number
          if (-not [string]::IsNullOrEmpty($_.TelephoneNumber)) { Add-ADGroupMember -Identity $phoneSysGrp -Members $sam }
          
          Write-Host "[INFO]`t Created user : $($sam)`r`n"
          "[INFO]`t Created user : $($sam)" | Out-File $log -append
                
        }
        Catch 
        { 
          Write-Error "Couldn't create user : $($_.Exception.Message)`r`n" 
          "[ERROR]`t Couldn't create user : $($_.Exception.Message)" | Out-File $log -Append
        }
  

        # Set only setup proxyaccount and rename the account for regular users
        if ($specialAcct -eq $false) {
          $proxyAddr = "SMTP:" + $_.Mail.ToLower()
          Write-Host "[INFO]`t Updating proxyAddr for $($sam) to $($proxyAddr)`r`n"
          "[INFO]`t Updating proxyAddr for $($sam) to $($proxyAddr)" | Out-File $log -append

          Try { Set-ADUser -Identity $sam -Add @{proxyAddresses=$proxyAddr} }
          Catch { Write-Error "Couldn't set the ProxyAddresses Attributes : $($_.Exception.Message)`r`n" 
                  "[ERROR]`t Couldn't set the ProxyAddresses Attributes : $($_.Exception.Message)" | Out-File $log -append
                }
      
          # Rename the object to a good looking name (otherwise you see
          # the 'ugly' shortened sAMAccountNames as a name in AD. This
          # can't be set right away (as sAMAccountName) due to the 20
          # character restriction
          Try { 
            $newDn = (Get-ADUser -Identity $sam).DistinguishedName
    
            Write-Host "[INFO]`t Updating name formatting in active directory for $($sam) to $($params.sn + ", " + $params.givenName)`r`n"
            "[INFO]`t Updating name formatting in active directory for $($sam) to $($params.sn + ", " + $params.givenName)" | Out-File $log -append
  
            Rename-ADObject -Identity $newDn -NewName ($params.sn + ", " + $params.givenName)  
            }
          Catch { Write-Error "Couldn't rename user : $($_.Exception.Message)`r`n"
                  "[ERROR]`t Couldn't rename user : $($_.Exception.Message)" | Out-File $log -Append
                }
  
        } else { # special account, move to special OU
          $newDn = (Get-ADUser -Identity $sam).DistinguishedName
    
          Move-ADObject -Identity $newDn -TargetPath $phoneSysOU
          Write-Host "[INFO]`t Moved special user $($sam) to Phone System OU`r`n" 
          "[INFO]`t Moved special user $($sam) to Phone System OU" | Out-File $log -append

          Try { 
            $newDn = (Get-ADUser -Identity $sam).DistinguishedName
            $subStrings = $sam.Split("_")
            $subStrings[0] = Convert-FirstLetterUpper -InputString $subStrings[0]
            $newName = $params.givenName

            if ($subStrings[0] -ne "Fax") { $newName = $params.givenName + " " + $subStrings[0] }

            Write-Host "[INFO]`t Updating name formatting in active directory for $($sam) to $($newName)`r`n"
            "[INFO]`t Updating name formatting in active directory for $($sam) to $($newName)" | Out-File $log -append
  
            Rename-ADObject -Identity $newDn -NewName $newName
            }
          Catch { Write-Error "Couldn't rename user $($sam) : $($_.Exception.Message)`r`n"
                  "[ERROR]`t Couldn't rename user $($sam) : $($_.Exception.Message)" | Out-File $log -Append
                }

        }

      } # end of create user
      elseif ((($_.Update.ToLower()) -eq "remove") -and ($exists)) # user exists and command is delete
      {
        Try {
          Remove-ADUser -Identity $sam -Confirm:$false 
          Write-Host "[INFO]`t Deleted user : $($sam)`r`n"
          "[INFO]`t Deleted user : $($sam)" | Out-File $log -append
        }
        Catch {
          Write-Error "Couldn't delete user $($sam) : $($_.Exception.Message)`r`n"
          "[ERROR]`t Couldn't delete user $($sam) : $($_.Exception.Message)" | Out-File $log -append
        }
      } # end of delete user
      else {
        Write-Host "[SKIP]`t No action specified for user $($sam)`r`n"
        "[SKIP]`t No action specified for user $($sam)" | Out-File $log -append
      }

    }
    $i++
  }
}

"Processing started (on " + $(Get-Date) + "): " | Out-File $log -append
"--------------------------------------------" | Out-File $log -append
Write-Host "STARTED SCRIPT"
Start-Commands
Write-Host "STOPPED SCRIPT"
"--------------------------------------------" | Out-File $log -append
"Processing stopped (on " + $(Get-Date) + "): " | Out-File $log -append

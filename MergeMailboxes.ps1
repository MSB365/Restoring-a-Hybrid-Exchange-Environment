#region Description
<#     
       .NOTES
       ===========================================================================
       Created on:         2019/12/27 
       Created by:         Drago Petrovic | Dominic Manning
       Organization:       MSB365.blog
       Filename:           MergeMailboxes.ps1     

       Find us on:
             * Website:           https://www.msb365.blog
             * Microsoft:         https://mvp.microsoft.com/en-us/PublicProfile/5003446?fullName=Drago%20Petrovic
             * Technet:           https://social.technet.microsoft.com/Profile/drago%20petrovic
             * MS Tech Community: https://techcommunity.microsoft.com/t5/user/viewprofilepage/user-id/219022
             * LinkedIn:          https://www.linkedin.com/in/drago-petrovic/
             * Xing:              https://www.xing.com/profile/Drago_Petrovic
       ===========================================================================
       .DESCRIPTION
             This script helps creating or re-creating a supported state for hybrid Exchange.
             If on-premise Exchange mailboxes where migrated with third party tools to Exchange Online, it can merge the "old" on premise mailboxes with the current Exchange Online
             mailboxes. The script is editing the on-premises mailboxes to become remote mailboxes, sets the routing address and adds the Exchange GUID from the online Mailbox to the on-premise one.
			In the beginning it will look for all AD accounts with remote mailboxes.
			To find Exchange Online mailboxes and AD accounts that belong together, one attribute on each side can be defined which will be used for matching. E.g. User Principal Name.

       .NOTES
             After the script is finished successfully, you need to enable the Exchange sync on the AAD and run the Exchange HCW (minimal) to finalize the merge.
             

       .COPYRIGHT
       Permission is hereby granted, free of charge, to any person obtaining a copy of this software and associated documentation files (the "Software"), 
       to deal in the Software without restriction, including without limitation the rights to use, copy, modify, merge, publish, distribute, sublicense, 
       and/or sell copies of the Software, and to permit persons to whom the Software is furnished to do so, subject to the following conditions:

       The above copyright notice and this permission notice shall be included in all copies or substantial portions of the Software.

       THE SOFTWARE IS PROVIDED "AS IS", WITHOUT WARRANTY OF ANY KIND, EXPRESS OR IMPLIED, INCLUDING BUT NOT LIMITED TO THE WARRANTIES OF MERCHANTABILITY, 
       FITNESS FOR A PARTICULAR PURPOSE AND NONINFRINGEMENT. IN NO EVENT SHALL THE AUTHORS OR COPYRIGHT HOLDERS BE LIABLE FOR ANY CLAIM, DAMAGES OR OTHER LIABILITY, 
       WHETHER IN AN ACTION OF CONTRACT, TORT OR OTHERWISE, ARISING FROM, OUT OF OR IN CONNECTION WITH THE SOFTWARE OR THE USE OR OTHER DEALINGS IN THE SOFTWARE.
       ===========================================================================
       .CHANGE LOG
             V0.10, 2019/12/27 - Initial version
             V0.20, 2019/12/30 - Bug fixes, User input
             V0.80, 2019/12/31 - Add error handling
             V1.00, 2020/01/04 - Final version V1
 

--- keep it simple, but significant ---
#>

#region functions

[cmdletbinding()]
param(
[switch]$accepteula,
[switch]$v)

###############################################################################
#Script Name variable
$Scriptname = "Restoring a Hybrid Exchange Environment"
$RKEY = "MSB365_Restoring-a-Hybrid-Exchange-Environment"
###############################################################################

[void][System.Reflection.Assembly]::Load('System.Drawing, Version=4.0.0.0, Culture=neutral, PublicKeyToken=b03f5f7f11d50a3a')
[void][System.Reflection.Assembly]::Load('System.Windows.Forms, Version=4.0.0.0, Culture=neutral, PublicKeyToken=b77a5c561934e089')

function ShowEULAPopup($mode)
{
    $EULA = New-Object -TypeName System.Windows.Forms.Form
    $richTextBox1 = New-Object System.Windows.Forms.RichTextBox
    $btnAcknowledge = New-Object System.Windows.Forms.Button
    $btnCancel = New-Object System.Windows.Forms.Button

    $EULA.SuspendLayout()
    $EULA.Name = "MIT"
    $EULA.Text = "$Scriptname - License Agreement"

    $richTextBox1.Anchor = [System.Windows.Forms.AnchorStyles]::Top -bor [System.Windows.Forms.AnchorStyles]::Bottom -bor [System.Windows.Forms.AnchorStyles]::Left -bor [System.Windows.Forms.AnchorStyles]::Right
    $richTextBox1.Location = New-Object System.Drawing.Point(12,12)
    $richTextBox1.Name = "richTextBox1"
    $richTextBox1.ScrollBars = [System.Windows.Forms.RichTextBoxScrollBars]::Vertical
    $richTextBox1.Size = New-Object System.Drawing.Size(776, 397)
    $richTextBox1.TabIndex = 0
    $richTextBox1.ReadOnly=$True
    $richTextBox1.Add_LinkClicked({Start-Process -FilePath $_.LinkText})
    $richTextBox1.Rtf = @"
{\rtf1\ansi\ansicpg1252\deff0\nouicompat{\fonttbl{\f0\fswiss\fprq2\fcharset0 Segoe UI;}{\f1\fnil\fcharset0 Calibri;}{\f2\fnil\fcharset0 Microsoft Sans Serif;}}
{\colortbl ;\red0\green0\blue255;}
{\*\generator Riched20 10.0.19041}{\*\mmathPr\mdispDef1\mwrapIndent1440 }\viewkind4\uc1
\pard\widctlpar\f0\fs19\lang1033 MSB365 SOFTWARE MIT LICENSE\par
Copyright (c) 2022 Drago Petrovic\par
$Scriptname \par
\par
{\pict{\*\picprop}\wmetafile8\picw26\pich26\picwgoal32000\pichgoal15
0100090000035000000000002700000000000400000003010800050000000b0200000000050000
000c0202000200030000001e000400000007010400040000000701040027000000410b2000cc00
010001000000000001000100000000002800000001000000010000000100010000000000000000
000000000000000000000000000000000000000000ffffff00000000ff040000002701ffff0300
00000000
}These license terms are an agreement between you and MSB365 (or one of its affiliates). IF YOU COMPLY WITH THESE LICENSE TERMS, YOU HAVE THE RIGHTS BELOW. BY USING THE SOFTWARE, YOU ACCEPT THESE TERMS.\par
\par
MIT License\par
{\pict{\*\picprop}\wmetafile8\picw26\pich26\picwgoal32000\pichgoal15
0100090000035000000000002700000000000400000003010800050000000b0200000000050000
000c0202000200030000001e000400000007010400040000000701040027000000410b2000cc00
010001000000000001000100000000002800000001000000010000000100010000000000000000
000000000000000000000000000000000000000000ffffff00000000ff040000002701ffff0300
00000000
}\par
\pard
{\pntext\f0 1.\tab}{\*\pn\pnlvlbody\pnf0\pnindent0\pnstart1\pndec{\pntxta.}}
\fi-360\li360 Permission is hereby granted, free of charge, to any person obtaining a copy of this software and associated documentation files (the "Software"), to deal in the Software without restriction, including without limitation the rights to use, copy, modify, merge, publish, distribute, sublicense copies of the Software, and to permit persons to whom the Software is furnished to do so, subject to the following conditions: \par
\pard\widctlpar\par
\pard\widctlpar\li360 The above copyright notice and this permission notice shall be included in all copies or substantial portions of the Software.\par
\par
\pard\widctlpar\fi-360\li360 2.\tab THE SOFTWARE IS PROVIDED "AS IS", WITHOUT WARRANTY OF ANY KIND, EXPRESS OR IMPLIED, INCLUDING BUT NOT LIMITED TO THE WARRANTIES OF MERCHANTABILITY, FITNESS FOR A PARTICULAR PURPOSE AND NONINFRINGEMENT. \par
\pard\widctlpar\par
\pard\widctlpar\fi-360\li360 3.\tab IN NO EVENT SHALL THE AUTHORS OR COPYRIGHT HOLDERS BE LIABLE FOR ANY CLAIM, DAMAGES OR OTHER LIABILITY, WHETHER IN AN ACTION OF CONTRACT, TORT OR OTHERWISE, ARISING FROM, OUT OF OR IN CONNECTION WITH THE SOFTWARE OR THE USE OR OTHER DEALINGS IN THE SOFTWARE. \par
\pard\widctlpar\par
\pard\widctlpar\fi-360\li360 4.\tab DISCLAIMER OF WARRANTY. THE SOFTWARE IS PROVIDED \ldblquote AS IS,\rdblquote  WITHOUT WARRANTY OF ANY KIND, EXPRESS OR IMPLIED, INCLUDING BUT NOT LIMITED TO THE WARRANTIES OF MERCHANTABILITY, FITNESS FOR A PARTICULAR PURPOSE AND NONINFRINGEMENT. IN NO EVENT SHALL MSB365 OR ITS LICENSORS BE LIABLE FOR ANY DIRECT, INDIRECT, INCIDENTAL, SPECIAL, EXEMPLARY, OR CONSEQUENTIAL DAMAGES (INCLUDING, BUT NOT LIMITED TO, PROCUREMENT OF SUBSTITUTE GOODS OR SERVICES; LOSS OF USE, DATA, OR PROFITS; OR BUSINESS INTERRUPTION) HOWEVER CAUSED AND ON ANY THEORY OF LIABILITY, WHETHER IN CONTRACT, STRICT LIABILITY, OR TORT (INCLUDING NEGLIGENCE OR OTHERWISE) ARISING IN ANY WAY OUT OF THE USE OF THE SOFTWARE, EVEN IF ADVISED OF THE POSSIBILITY OF SUCH DAMAGE.\par
\pard\widctlpar\par
\pard\widctlpar\fi-360\li360\qj 5.\tab LIMITATION ON AND EXCLUSION OF DAMAGES. IF YOU HAVE ANY BASIS FOR RECOVERING DAMAGES DESPITE THE PRECEDING DISCLAIMER OF WARRANTY, YOU CAN RECOVER FROM MICROSOFT AND ITS SUPPLIERS ONLY DIRECT DAMAGES UP TO U.S. $1.00. YOU CANNOT RECOVER ANY OTHER DAMAGES, INCLUDING CONSEQUENTIAL, LOST PROFITS, SPECIAL, INDIRECT, OR INCIDENTAL DAMAGES. This limitation applies to (i) anything related to the Software, services, content (including code) on third party Internet sites, or third party applications; and (ii) claims for breach of contract, warranty, guarantee, or condition; strict liability, negligence, or other tort; or any other claim; in each case to the extent permitted by applicable law. It also applies even if MSB365 knew or should have known about the possibility of the damages. The above limitation or exclusion may not apply to you because your state, province, or country may not allow the exclusion or limitation of incidental, consequential, or other damages.\par
\pard\widctlpar\par
\pard\widctlpar\fi-360\li360\qj 6.\tab ENTIRE AGREEMENT. This agreement, and any other terms MSB365 may provide for supplements, updates, or third-party applications, is the entire agreement for the software.\par
\pard\widctlpar\qj\par
\pard\widctlpar\fi-360\li360\qj 7.\tab A complete script documentation can be found on the website https://www.msb365.blog.\par
\pard\widctlpar\par
\pard\sa200\sl276\slmult1\f1\fs22\lang9\par
\pard\f2\fs17\lang2057\par
}
"@
    $richTextBox1.BackColor = [System.Drawing.Color]::White
    $btnAcknowledge.Anchor = [System.Windows.Forms.AnchorStyles]::Bottom -bor [System.Windows.Forms.AnchorStyles]::Right
    $btnAcknowledge.Location = New-Object System.Drawing.Point(544, 415)
    $btnAcknowledge.Name = "btnAcknowledge";
    $btnAcknowledge.Size = New-Object System.Drawing.Size(119, 23)
    $btnAcknowledge.TabIndex = 1
    $btnAcknowledge.Text = "Accept"
    $btnAcknowledge.UseVisualStyleBackColor = $True
    $btnAcknowledge.Add_Click({$EULA.DialogResult=[System.Windows.Forms.DialogResult]::Yes})

    $btnCancel.Anchor = [System.Windows.Forms.AnchorStyles]::Bottom -bor [System.Windows.Forms.AnchorStyles]::Right
    $btnCancel.Location = New-Object System.Drawing.Point(669, 415)
    $btnCancel.Name = "btnCancel"
    $btnCancel.Size = New-Object System.Drawing.Size(119, 23)
    $btnCancel.TabIndex = 2
    if($mode -ne 0)
    {
   $btnCancel.Text = "Close"
    }
    else
    {
   $btnCancel.Text = "Decline"
    }
    $btnCancel.UseVisualStyleBackColor = $True
    $btnCancel.Add_Click({$EULA.DialogResult=[System.Windows.Forms.DialogResult]::No})

    $EULA.AutoScaleDimensions = New-Object System.Drawing.SizeF(6.0, 13.0)
    $EULA.AutoScaleMode = [System.Windows.Forms.AutoScaleMode]::Font
    $EULA.ClientSize = New-Object System.Drawing.Size(800, 450)
    $EULA.Controls.Add($btnCancel)
    $EULA.Controls.Add($richTextBox1)
    if($mode -ne 0)
    {
   $EULA.AcceptButton=$btnCancel
    }
    else
    {
        $EULA.Controls.Add($btnAcknowledge)
   $EULA.AcceptButton=$btnAcknowledge
        $EULA.CancelButton=$btnCancel
    }
    $EULA.ResumeLayout($false)
    $EULA.Size = New-Object System.Drawing.Size(800, 650)

    Return ($EULA.ShowDialog())
}

function ShowEULAIfNeeded($toolName, $mode)
{
$eulaRegPath = "HKCU:Software\Microsoft\$RKEY"
$eulaAccepted = "No"
$eulaValue = $toolName + " EULA Accepted"
if(Test-Path $eulaRegPath)
{
$eulaRegKey = Get-Item $eulaRegPath
$eulaAccepted = $eulaRegKey.GetValue($eulaValue, "No")
}
else
{
$eulaRegKey = New-Item $eulaRegPath
}
if($mode -eq 2) # silent accept
{
$eulaAccepted = "Yes"
        $ignore = New-ItemProperty -Path $eulaRegPath -Name $eulaValue -Value $eulaAccepted -PropertyType String -Force
}
else
{
if($eulaAccepted -eq "No")
{
$eulaAccepted = ShowEULAPopup($mode)
if($eulaAccepted -eq [System.Windows.Forms.DialogResult]::Yes)
{
        $eulaAccepted = "Yes"
        $ignore = New-ItemProperty -Path $eulaRegPath -Name $eulaValue -Value $eulaAccepted -PropertyType String -Force
}
}
}
return $eulaAccepted
}

if ($accepteula)
    {
         ShowEULAIfNeeded "DS Authentication Scripts:" 2
         "EULA Accepted"
    }
else
    {
        $eulaAccepted = ShowEULAIfNeeded "DS Authentication Scripts:" 0
        if($eulaAccepted -ne "Yes")
            {
                "EULA Declined"
                exit
            }
         "EULA Accepted"
    }
###############################################################################




write-host "  _           __  __ ___ ___   ____  __ ___  " -ForegroundColor Yellow
write-host " | |__ _  _  |  \/  / __| _ ) |__ / / /| __| " -ForegroundColor Yellow
write-host " | '_ \ || | | |\/| \__ \ _ \  |_ \/ _ \__ \ " -ForegroundColor Yellow
write-host " |_.__/\_, | |_|  |_|___/___/ |___/\___/___/ " -ForegroundColor Yellow
write-host "       |__/                                  " -ForegroundColor Yellow
Start-Sleep -s 2
write-host ""                                                                                   
write-host ""
write-host ""
write-host ""
write-host ""

##############################################################################################################







function writeLog($message, [SWITCH]$logOnly,$state)
{
	switch ($state)
	{
		"i"{ $prefix = "Info: ";$fgc = "Magenta" }
		"w"{ $prefix = "Warning: "; $fgc = "Yellow" }
		"e"{ $prefix = "Error: "; $fgc = "Red"  }
		"s"{ $prefix = "Success: "; $fgc = "Green"  }
	}
	(Get-Date -Format G) + " " + $prefix + $message | Out-File $logfile -Append
	if (!($logOnly))
	{
		Write-Host (Get-Date -Format G) + " " + $prefix + $message -ForegroundColor $fgc
	}
	
}
#endregion
#region variables
$date = Get-Date -Format yyyyMMdd-HHmmss
$logfile = ".\MergeMailboxes_$date.log"
$matchescount = 0
$successcount = 0
$errorcount = 0
#endregion

#region get user input
#Set domain variable
$onpremDomain = Read-Host "Define the default email address domain eg. "contoso.com""
$o365Domain = Read-Host "Define the Office 365 email domain eg. "contoso.mail.onmicrosoft.com""
$O365Cred = Get-credential -Message "Enter Exchange Online credentials"


Write-Host "Select by which attributes the O365 mailboxes and On-Prem AD accounts should be matched (must have same value on both sides). Option 4 let's you enter the names for both attributes yourself" -ForegroundColor Cyan

#Create table for user selection
$tabName = "Matching criteria"

#Create Table object
$table = New-Object system.Data.DataTable "$tabName"

#Define Columns
$selectNo = New-Object system.Data.DataColumn "Number", ([string])
$col1 = New-Object system.Data.DataColumn "O365 mailbox", ([string])
$col2 = New-Object system.Data.DataColumn "AD account", ([string])

#Add the Columns
$table.columns.add($selectNo)
$table.columns.add($col1)
$table.columns.add($col2)

#Create a row
$row = $table.NewRow()

#Enter data in the row
$row.Number = 1
$row.'O365 mailbox' = "WindowsEmailAddress"
$row.'AD account' = "mail"

$row1 = $table.NewRow()
$row1.Number = 2
$row1.'O365 mailbox' = "UserPrincipalName"
$row1.'AD account' = "UserPrincipalName"

$row2 = $table.NewRow()
$row2.Number = 3
$row2.'O365 mailbox' = "Alias"
$row2.'AD account' = "SAMAccountName"

$row3 = $table.NewRow()
$row3.Number = 4
$row3.'O365 mailbox' = "Custom"
$row3.'AD account' = "Custom"

#Add the row to the table
$table.Rows.Add($row)
$table.Rows.Add($row1)
$table.Rows.Add($row2)
$table.Rows.Add($row3)
#send the table to out-grid and pass the output through
$result = $table | Out-GridView -OutputMode Single -Title "Select attributes for matching"
#Set attributes for matching, depending on user selection
switch ($result.Number)
{
	1 { $O365Attrib = "WindowsEmailAddress"; $ADattrib = "mail" }
	2{ $O365Attrib = "UserPrincipalName"; $ADattrib = "UserPrincipalName" }
	3{ $O365Attrib = "Alias"; $ADattrib = "SAMAccountName" }
	4{
		$O365Attrib = Read-Host "Enter attribute name of Exchange Online mailbox"
		$ADattrib = Read-Host "Enter attribute name of AD account"
	}
	$null{ "Cancelled"; return}
}

Write-Host "Thank you" -ForegroundColor Green
Start-Sleep -Seconds 1.5
#endregion

#region Collect on-premise mailboxes
writeLog -message "Getting on-premises AD users with mailboxes" -state i
Start-Sleep -Seconds 1.0
$users = Get-ADuser -filter * -Properties mail, proxyaddresses, targetAddress, msExchRecipientTypeDetails, homeMDB, homeMTA, msExchHomeServerName, `
					msExchVersion, msExchRecipientDisplayType, msExchRemoteRecipientType, msExchWhenMailboxCreated, msExchMailboxGuid | `
Where { $_.mail -notlike "SystemMailbox*" -and $_.mail -notlike "HealthMailbox*" -and $_.msExchRecipientTypeDetails -eq "2147483648" } -ErrorAction Stop

if ($users.count -gt 0)
{
	writeLog -message "Found $($user.count) users." -state i
}
Else
{
	writeLog "No users found. Exiting" -state w
	Return
}
Start-Sleep -Seconds 3
foreach ($u in $users) { write-host $u.mail }
Write-Host "done" -ForegroundColor Green
Start-Sleep -Seconds 1.5
#endregion

#region set Exchange attributes on AD user
#Set homeMDB, homeMTA, etc.
writeLog -message "Clearing mailbox attributes of AD users" -state i -logOnly
Start-Sleep -Seconds 1.0
foreach ($u in $users)
{
	try
	{
		Set-ADuser $u -clear homeMDB, homeMTA, msExchHomeServerName, msExchVersion, `
				   msExchRecipientDisplayType, msExchRecipientTypeDetails, msExchRemoteRecipientType, msExchWhenMailboxCreated, msExchMailboxGuid -ErrorAction Stop
		writeLog -message "Cleared attributes on user $($u.SAMAccountName)" -state s
	}
	catch
	{
		writeLog -message $($u.SAMAccountName) + " " +$_.Exception.Message -state e
	}
}


Write-Host "done" -ForegroundColor Green
Start-Sleep -Seconds 1.5

#Transforming Mailbox to Office 365 Mailbox
writeLog -message "Transforming on-premise mailboxes to Office 365 remote mailboxes" -state i
Start-Sleep -Seconds 1.0
foreach ($u in $users)
{
	try
	{
		Set-ADuser $u -add @{
			msExchVersion= "44220983382016"; msExchRecipientDisplayType = "-2147483642"; `
			msExchRecipientTypeDetails = "2147483648"; msExchRemoteRecipientType = "1"
		} -ErrorAction Stop
		writeLog -message "Set remote mailbox attributes on user $($u.SAMAccountName)" -state s
	}
	catch
	{
		writeLog -message $($u.SAMAccountName) + " " +$_.Exception.Message -state e
	}
	
}
Write-Host "done" -ForegroundColor Green
Start-Sleep -Seconds 1.5
#endregion

#Region Set Routing Mail Addresses
writeLog -message "Setting up routing mail addresses on the remote mailboxes" -state i

Start-Sleep -Seconds 1.0

foreach ($u in $users)
{
	try
	{
		
		#generate the onmicrosoft address string
		$nmail = $u.mail.Replace("@$onpremDomain", "@$o365Domain")
		$proxytoadd = "smtp:" + $nmail
		#generate the on-prem primary smtp address string
		$onpremSMTP = "SMTP:" + $u.givenname + "." + $u.surname + "@" + $onpremDomain
		#Adds the correct on-prem primary SMTP address if not already set
		if (!($u.proxyaddresses.Contains($onpremSMTP)))
		{
			$oldSMTP = $user.proxyaddresses | where{ $_ -ceq "SMTP:" }
			if ($oldSMTP)
			{
				$oldaddSMTP = $oldSMTP.replace("SMTP:", "smtp:")
				$u.proxyaddresses.Remove($oldSMTP)
				$u.proxyaddresses.Add($oldaddSMTP)
				$u.proxyaddresses.Add($onpremSMTP)
				
			}
		}
		#add the onmicrosoft e-mail address to the proxy addresses	
		$u.proxyaddresses.Add($proxytoadd)
		#set the onmicrosoft e-mail address as target address
		Set-ADUser $u -add @{ targetAddress = $nmail } -ErrorAction Stop
		writeLog -message "Set targetaddress $nmail on user $($u.SAMAccountName)" -state s
		#Set all other attributes which were set before
		Set-ADuser -Instance $u -ErrorAction stop	
		writeLog -message "Set proxyaddresses on user $($u.SAMAccountName)" -state s
	}
	Catch
	{
		writeLog -message $($u.SAMAccountName) + " " + $_.Exception.Message -state e
	}
	
}

Write-Host "done" -ForegroundColor Green
Start-Sleep -Seconds 1.5
#endregion

#region set Exchange GUID
#Preparing modifying GUID's on the remote Mailboxes
writeLog -message "Preparing the Exchange GUID maching between Remote and Cloud Mailboxes..." -ForegroundColor Magenta
Start-Sleep -Seconds 2.5

# connect to exo -credential $credential
writeLog -message "Connecting Exchange online" -state i
Start-Sleep -Seconds 1.0
try
{
	$Session = New-PSSession -ConfigurationName Microsoft.Exchange -ConnectionUri https://outlook.office365.com/powershell-liveid/ -Credential $O365Cred -Authentication Basic -AllowRedirection -ErrorAction Stop
	Import-PSSession $Session -Prefix O365 -ErrorAction Stop
	writeLog -message "Connected to Exchange Online" -state s
}
catch
{
	writeLog -message $_.Exception.Message -state e
}

#Modify GUID's on the remote Mailboxes
writeLog -message "Matching Exchange GUID's between Remote and Cloud Mailboxes" -state i
Start-Sleep -Seconds 1.0
cls
writeLog -message "Getting all Exchange Online mailboxes" -state i
$O365Mailboxes = Get-O365mailbox
if ($O365Mailboxes.count -gt 0)
{
	writeLog -message "Found $($O365Mailboxes.count) cloud mailboxes." -state s
}
Else
{
	writeLog -message "Could not find any cloud mailboxes." -state e
	Return
}

foreach ($mbx in $O365Mailboxes)
{
	#match O365 mailbox with AD user using the specified attribute pair
	$usermatch = $users | where { $_.$ADattrib -eq $mbx.$O365Attrib }
	$matchescount++
	if ($usermatch)
	{
		try
		{
			#Set the GUID of the O365 mailbox as Exhange GUID of the remote mailbox
			Set-Remotemailbox $usermatch.SamaccountName -ExchangeGUID $mbx.ExchangeGUID -erroraction Stop
			write-host "Set remote mailbox $($usermatch.SamaccountName)" -ForegroundColor Green
			$successcount++
			}
		catch
		{
			writeLog -message "Couldn't set ExchangeGUID of $($mbx.Alias) on AD user $($usermatch.SamaccountName)" -state e
			$errorcount++
		}	
	}
	Else
	{
		writeLog -message "No match for Exchange Online Mailbox: $($mbx.WindowsEmailAddress) found." -state w
	}
}
cls
"--------------------Summary--------------------"
"Number of O365 mailboxes found: $($O365Mailboxes.count)"
"Number of matches found: $matchescount"
"Number of successfully set ExchangeGUID: $successcount"
"Number of Errors while set ExchangeGUID: $errorcount"
""
""
#Disconnect Exchange online PowerShell Session
Remove-PSSession -Session $Session
Start-Sleep -Seconds 1.5

Write-Host "done" -ForegroundColor Green
""
#endregion
"Please view the log file to check for any errors and to find detailed information! Log file path: $logfile"
""
Write-Host "All configuration on-premises was done. Enable the Exchange sync in the AAD and run the Exchange HCW (minimal) to make the final step." -ForegroundColor Cyan

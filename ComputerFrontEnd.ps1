#=======================================================
# 
# Name: *.ps1
# 
# Created By: Niklas Landqvist
#
# Comment: Internal script for task sequence.
#
# Date: 2023-xx-xx
#
# Version: 1.0
#
# Source: https://www.modernendpoint.com/managed/Running-an-Azure-Automation-runbook-to-update-MEMCM-Primary-User/
# Source: https://github.com/CharlesNRU/mdm-adminservice/blob/master/Invoke-GetPackageIDFromAdminService.ps1
# Source: https://til.intrepidintegration.com/powershell/ssl-cert-bypass.html
#=======================================================


$AllTypes = @("A - Admin","E - Elev","G - Komvux",
"L - Lärare","V - Virituell")

$PublikDesktopSearchChar = "DT"

#=======================================================
$TSProgressUI = New-Object -ComObject Microsoft.SMS.TSProgressUI
$TSEnvironment = New-Object -ComObject Microsoft.SMS.TSEnvironment
$TSProgressUI.CloseProgressDialog()

$SiteServerFQDN = $TSEnvironment.Value("OSDSiteServerFQDN")
$AdminServiceUserName = $TSEnvironment.Value("OSDAdminServiceUserName")
$AdminServicePassword =$TSEnvironment.Value("OSDAdminServicePassword")

$Date = $(Get-Date -Format "yyMM")
$UUID = (Get-WmiObject -class Win32_ComputerSystemProduct).UUID
$ResultsChassis = Get-WmiObject -class Win32_SystemEnclosure
foreach ($Item in $ResultsChassis) {
    
    switch ($Item.ChassisTypes)
    {
        {$_ -in "8", "9", "10", "11", "12", "14", "18", "21", "30", "32"} {$ChassiType = "LT"}
        {$_ -in "3", "4", "5", "6", "7", "13", "15", "16"} {$ChassiType = "DT"}
        Default {$ChassiType = "XT"}
    }
}


$EncryptedPassword = ConvertTo-SecureString -String $AdminServicePassword -AsPlainText -Force
$Credential = New-Object -TypeName System.Management.Automation.PSCredential -ArgumentList @("$AdminServiceUserName", $EncryptedPassword)


Add-Type @"
using System.Net;
using System.Security.Cryptography.X509Certificates;
public class TrustAllCertsPolicy : ICertificatePolicy {
    public bool CheckValidationResult(
        ServicePoint srvPoint, X509Certificate certificate,
        WebRequest request, int certificateProblem) {
        return true;
    }
}
"@

[System.Net.ServicePointManager]::CertificatePolicy = New-Object TrustAllCertsPolicy
[Net.ServicePointManager]::SecurityProtocol = [Net.SecurityProtocolType]::Ssl3, [Net.SecurityProtocolType]::Tls, [Net.SecurityProtocolType]::Tls11, [Net.SecurityProtocolType]::Tls12
        


function GetDevicesInfo ($Filter)
{
    
    $Body = @{
        "`$filter" = $Filter
        "`$select" = "Name,SMBIOSGUID,Client"
    }

    #Fix for AdminService CB2111 and newer
    Add-Type -AssemblyName System.Web
    $BodyParameters = [System.Web.HttpUtility]::ParseQueryString([String]::Empty)
    foreach ($param in $Body.GetEnumerator()) {
        $BodyParameters.Add($param.Name, $param.Value)
    }

    $WMIPackageURI = "https://$SiteServerFQDN/AdminService/wmi/SMS_R_System"
    $Request = [System.UriBuilder]($WMIPackageURI)
    $Request.Query = $BodyParameters.ToString()
    $DecodedURI = [System.Web.HttpUtility]::UrlDecode($Request.Uri)
    $Devices = Invoke-RestMethod -Method Get -Uri $DecodedURI -Credential $Credential #| Select-Object -ExpandProperty value

    return $Devices
}


function NewDeviceCount ($Device)
{
    $Device = ($Device).Split("-")[1]
    $Numbers = ($Device -replace '[^\d]+')           
    if ($Numbers) {

        [int]$Numbers += 1
        $DeviceCount = $Numbers.ToString()

        if ($DeviceCount.Length -lt 4) {
            
            do {
                $DeviceCount = "0" + $DeviceCount
            }
            until ($DeviceCount.Length -eq 4)
        }

       return $DeviceCount
    }
    else {
        
        return $null
    }
}


function AddNewDevices ($NetbiosName,$SMBIOSGUID)
{
    $MACAddress = ""
    $OverwriteExistingRecord = $true

    [string]$MethodClass = "SMS_Site"
    [string]$MethodName = "ImportMachineEntryToMultipleCollections"
    $PostURL = "https://$SiteServerFQDN/AdminService/wmi/$MethodClass.$MethodName"

    $Headers = @{
        "Content-Type" = "Application/json"
    }

    $Body = @{
        NetbiosName = $NetbiosName
        SMBIOSGUID = $SMBIOSGUID
        MACAddress = $MACAddress
        OverwriteExistingRecord = $OverwriteExistingRecord 
    } | ConvertTo-Json

    $ReturnValue = Invoke-RestMethod -Method Post -Uri "$($PostURL)" -Body $Body -Headers $Headers -Credential $Credential
    return $ReturnValue.ResourceID
}


function GenerateForm {


    [Reflection.Assembly]::loadwithpartialname("System.Drawing") | Out-Null
    [Reflection.Assembly]::loadwithpartialname("System.Windows.Forms") | Out-Null
    [System.Windows.Forms.Application]::EnableVisualStyles()

    $MainForm = New-Object System.Windows.Forms.Form
    $labelType = New-Object System.Windows.Forms.Label
    $labelComputer = New-Object System.Windows.Forms.Label
    $buttonInstall = New-Object System.Windows.Forms.Button
    $textBoxComputerName = New-Object System.Windows.Forms.TextBox
    $buttonGetName = New-Object System.Windows.Forms.Button
    $groupBoxType = New-Object System.Windows.Forms.GroupBox
    $comboBoxType = New-Object System.Windows.Forms.ComboBox
    $groupBoxComputerName = New-Object System.Windows.Forms.GroupBox
    $InitialFormWindowState = New-Object System.Windows.Forms.FormWindowState

    $ArrayList = New-Object System.Collections.ArrayList


    #----------------------------------------------

    $handler_MainForm_Load= 
    {
        foreach ($Type in $AllTypes) {

            $comboBoxType.Items.Add($Type)|Out-Null
        }
        
        $textBoxComputerName.Enabled = $False
        $buttonInstall.Enabled = $False
        $buttonGetName.Enabled = $False

        $DevicesInfo = GetDevicesInfo -Filter "startswith(SMBIOSGUID,'$UUID')"
        if ($($DevicesInfo.value)) {

            if ($($DevicesInfo.value).Name -eq "Unknown") {

                $buttonGetName.Enabled = $True
            }
            else {

                $textBoxComputerName.Text = $($DevicesInfo.value).Name
                $comboBoxType.Enabled = $False
                $buttonInstall.Enabled = $True
            }
        }


    }



    $handler_comboBoxType_SelectionChangeCommitted= 
    {
        $SelectedText = $comboBoxType.Text
        if ($SelectedText) {

            $buttonGetName.Enabled = $True
        }
    }




    $buttonGetName_OnClick= 
    {
        
        if ($comboBoxType.Text)
        {
            
            if ($comboBoxType.Text.Split("-")[0].Trim() -like "*(*)*") {

                $PublikName = $comboBoxType.Text.Split("-")[0].Trim()
                $SearchName = $PublikName.Replace("(","$PublikDesktopSearchChar").Replace(")","")
                $SystemTypeName = $PublikName.Replace("(","$ChassiType").Replace(")","")
            }
            else {

                $SearchName = $comboBoxType.Text.Split("-")[0].Trim()
                $SystemTypeName = ($comboBoxType.Text.Split("-")[0].Trim() + $ChassiType)
            }

            $labelType.Text = $comboBoxType.Text.Split("-")[1].Trim()

            $AllTypeDevice = "startswith(Name,'%-$SearchName')"
            $DevicesInfo = GetDevicesInfo -Filter $AllTypeDevice 
            $DevicesName = $($DevicesInfo.value).Name
            if ($DevicesName) {

                if ($DevicesName.Count -eq 1) {
                
                    $ArrayList.Add($DevicesName)
                    $SortedArrayList = $ArrayList | Sort-Object
                }
                else {
                
                    $ArrayList.AddRange($DevicesName)
                    $SortedArrayList = $ArrayList | Sort-Object
                }

        
                if ($SortedArrayList.Count -eq 1) {
                    $LastDevice = $SortedArrayList
                }
                else {
                
                    $LastDevice = $SortedArrayList[($SortedArrayList.Count - 1)]
                }
            }
            else {
            
                $LastDevice = ("$Date" + "-" + $SystemTypeName + "0000")
            }
    
            $NewDeviceNr = NewDeviceCount -Device $LastDevice
            $NewDeviceName = ("$Date" + "-" + $SystemTypeName + $NewDeviceNr)

            $textBoxComputerName.Enabled = $True
            $buttonInstall.Enabled = $True
            $textBoxComputerName.Text = $NewDeviceName
        }


    }




    $buttonInstall_OnClick= 
    {
        
        if ($labelType.Text.Trim() -eq "Type") {

            $TSEnvironment.Value("OSDComputerName") = $textBoxComputerName.Text.Trim()
            $TSEnvironment.Value("OSDSystemType") = $labelType.Text.Trim() 
        }
        else {

            AddNewDevices -NetbiosName $textBoxComputerName.Text.Trim() -SMBIOSGUID $UUID
            $TSEnvironment.Value("OSDComputerName") = $textBoxComputerName.Text.Trim()
            $TSEnvironment.Value("OSDSystemType") = $labelType.Text.Trim() 
        }      
        
        $MainForm.Close()

    }

    $OnLoadForm_StateCorrection=
    {
	    $MainForm.WindowState = $InitialFormWindowState
    }

    #----------------------------------------------
    
    $MainForm.AcceptButton = $buttonInstall   
    $System_Drawing_Size = New-Object System.Drawing.Size
    $System_Drawing_Size.Height = 287
    $System_Drawing_Size.Width = 514
    $MainForm.ClientSize = $System_Drawing_Size
    $MainForm.ControlBox = $False
    $MainForm.DataBindings.DefaultDataSourceUpdateMode = 0
    $MainForm.FormBorderStyle = 1
    $MainForm.MaximizeBox = $False
    $MainForm.MinimizeBox = $False
    $MainForm.Name = "MainForm"
    $MainForm.StartPosition = 1
    $MainForm.Text = " Ale kommun - Select Computer Name"
    $MainForm.TopMost = $True
    $MainForm.add_Load($handler_MainForm_Load)

    $buttonInstall.Anchor = 0
    $buttonInstall.DataBindings.DefaultDataSourceUpdateMode = 0
    $System_Drawing_Point = New-Object System.Drawing.Point
    $System_Drawing_Point.X = 172
    $System_Drawing_Point.Y = 246
    $buttonInstall.Location = $System_Drawing_Point
    $buttonInstall.Name = "buttonInstall"
    $System_Drawing_Size = New-Object System.Drawing.Size
    $System_Drawing_Size.Height = 24
    $System_Drawing_Size.Width = 171
    $buttonInstall.Size = $System_Drawing_Size
    $buttonInstall.TabIndex = 3
    $buttonInstall.Text = "Install"
    $buttonInstall.UseVisualStyleBackColor = $True
    $buttonInstall.add_Click($buttonInstall_OnClick)
    $MainForm.Controls.Add($buttonInstall)

    $textBoxComputerName.Anchor = 0
    $textBoxComputerName.DataBindings.DefaultDataSourceUpdateMode = 0
    $System_Drawing_Point = New-Object System.Drawing.Point
    $System_Drawing_Point.X = 99
    $System_Drawing_Point.Y = 187
    $textBoxComputerName.Location = $System_Drawing_Point
    $textBoxComputerName.MaxLength = 15
    $textBoxComputerName.Name = "textBoxComputerName"
    $System_Drawing_Size = New-Object System.Drawing.Size
    $System_Drawing_Size.Height = 20
    $System_Drawing_Size.Width = 317
    $textBoxComputerName.Size = $System_Drawing_Size
    $textBoxComputerName.TabIndex = 2
    $textBoxComputerName.WordWrap = $False
    $MainForm.Controls.Add($textBoxComputerName)

    $buttonGetName.Anchor = 0
    $buttonGetName.DataBindings.DefaultDataSourceUpdateMode = 0
    $System_Drawing_Point = New-Object System.Drawing.Point
    $System_Drawing_Point.X = 172
    $System_Drawing_Point.Y = 122
    $buttonGetName.Location = $System_Drawing_Point
    $buttonGetName.Name = "buttonGetName"
    $System_Drawing_Size = New-Object System.Drawing.Size
    $System_Drawing_Size.Height = 24
    $System_Drawing_Size.Width = 171
    $buttonGetName.Size = $System_Drawing_Size
    $buttonGetName.TabIndex = 1
    $buttonGetName.Text = "Get Name"
    $buttonGetName.UseVisualStyleBackColor = $True
    $buttonGetName.add_Click($buttonGetName_OnClick)
    $MainForm.Controls.Add($buttonGetName)

    $groupBoxType.Anchor = 0
    $groupBoxType.DataBindings.DefaultDataSourceUpdateMode = 0
    $System_Drawing_Point = New-Object System.Drawing.Point
    $System_Drawing_Point.X = 76
    $System_Drawing_Point.Y = 34
    $groupBoxType.Location = $System_Drawing_Point
    $groupBoxType.Name = "groupBoxType"
    $System_Drawing_Size = New-Object System.Drawing.Size
    $System_Drawing_Size.Height = 73
    $System_Drawing_Size.Width = 362
    $groupBoxType.Size = $System_Drawing_Size
    $groupBoxType.TabIndex = 4
    $groupBoxType.TabStop = $False
    $groupBoxType.Text = " Select Type"
    $MainForm.Controls.Add($groupBoxType)

    $comboBoxType.DropDownStyle = 2
    $comboBoxType.DataBindings.DefaultDataSourceUpdateMode = 0
    $comboBoxType.FormattingEnabled = $True
    $System_Drawing_Point = New-Object System.Drawing.Point
    $System_Drawing_Point.X = 23
    $System_Drawing_Point.Y = 27
    $comboBoxType.Location = $System_Drawing_Point
    $comboBoxType.Name = "comboBoxType"
    $System_Drawing_Size = New-Object System.Drawing.Size
    $System_Drawing_Size.Height = 21
    $System_Drawing_Size.Width = 317
    $comboBoxType.Size = $System_Drawing_Size
    $comboBoxType.TabIndex = 0
    $comboBoxType.add_SelectionChangeCommitted($handler_comboBoxType_SelectionChangeCommitted)
    $groupBoxType.Controls.Add($comboBoxType)

    $labelType.DataBindings.DefaultDataSourceUpdateMode = 0
    $System_Drawing_Point = New-Object System.Drawing.Point
    $System_Drawing_Point.X = 13
    $System_Drawing_Point.Y = 263
    $labelType.Location = $System_Drawing_Point
    $labelType.Name = "labelType"
    $System_Drawing_Size = New-Object System.Drawing.Size
    $System_Drawing_Size.Height = 15
    $System_Drawing_Size.Width = 143
    $labelType.Size = $System_Drawing_Size
    $labelType.TabIndex = 7
    $labelType.Text = "Type"
    $labelType.Visible = $False
    $MainForm.Controls.Add($labelType)

    $labelComputer.DataBindings.DefaultDataSourceUpdateMode = 0
    $System_Drawing_Point = New-Object System.Drawing.Point
    $System_Drawing_Point.X = 13
    $System_Drawing_Point.Y = 246
    $labelComputer.Location = $System_Drawing_Point
    $labelComputer.Name = "labelComputer"
    $System_Drawing_Size = New-Object System.Drawing.Size
    $System_Drawing_Size.Height = 13
    $System_Drawing_Size.Width = 143
    $labelComputer.Size = $System_Drawing_Size
    $labelComputer.TabIndex = 6
    $labelComputer.Text = "Computer"
    $labelComputer.Visible = $False
    $MainForm.Controls.Add($labelComputer)


    $groupBoxComputerName.Anchor = 0
    $groupBoxComputerName.DataBindings.DefaultDataSourceUpdateMode = 0
    $System_Drawing_Point = New-Object System.Drawing.Point
    $System_Drawing_Point.X = 76
    $System_Drawing_Point.Y = 158
    $groupBoxComputerName.Location = $System_Drawing_Point
    $groupBoxComputerName.Name = "groupBoxComputerName"
    $System_Drawing_Size = New-Object System.Drawing.Size
    $System_Drawing_Size.Height = 70
    $System_Drawing_Size.Width = 362
    $groupBoxComputerName.Size = $System_Drawing_Size
    $groupBoxComputerName.TabIndex = 5
    $groupBoxComputerName.TabStop = $False
    $groupBoxComputerName.Text = " Computer Name "

    $MainForm.Controls.Add($groupBoxComputerName)
    $InitialFormWindowState = $MainForm.WindowState
    $MainForm.add_Load($OnLoadForm_StateCorrection)
    $MainForm.ShowDialog()| Out-Null

}

#Call the Function
GenerateForm


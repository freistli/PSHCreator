<# 
 
.DESCRIPTION 
   SelfSignedCertificate AutoScript 
 
.NOTES 
    Author: Freist Li
    Last Updated: 10/30/2014   
#> 
#Cert Genearation Related Functions
#********************************************************************************************************************
Add-Type -AssemblyName System.Windows.Forms 

#Create Cert, install Cert to My, install Cert to Root, Export Cert as pfx
    Function GenerateSelfSignedCert{
    Param (
            $certcn,
            $password,
            $certfilepath
           )    
        
    #Check if the certificate name was used before
    $thumbprintA=(dir cert:\localmachine\My -recurse | where {$_.Subject -match "CN=" + $certcn} | Select-Object -Last 1).thumbprint

    if ($thumbprintA.Length -gt 0)
    {
    Write-Host "Duplicated Cert Name used" -ForegroundColor Cyan
    return
    }
    else
    {
    $thumbprintA=New-SelfSignedCertificate -DnsName $certcn -CertStoreLocation cert:\LocalMachine\My |ForEach-Object{ $_.Thumbprint}
    }

    #If generated successfully
    if ($thumbprintA.Length -gt 0) 
    {
    #query the new installed cerificate again
    $thumbprintB=(dir cert:\localmachine\My -recurse | where {$_.Subject -match "CN=" + $certcn} | Select-Object -Last 1).thumbprint

    #If new cert installed sucessfully with the same thumbprint
        if($thumbprintA -eq $thumbprintB )
        {
            $message = $certcn + " installed into LocalMachine\My successfully with thumprint "+$thumbprintA
            Write-Host $message -ForegroundColor Cyan
            $mypwd = ConvertTo-SecureString -String $password -Force –AsPlainText
            Write-Host "Exporting Certificate as .pfx file" -ForegroundColor Cyan
            Export-PfxCertificate -FilePath $certfilepath -Cert cert:\localmachine\My\$thumbprintA -Password $mypwd
            Write-Host "Importing Certificate to LocalMachine\Root" -ForegroundColor Cyan
            Import-PfxCertificate -FilePath $certfilepath -Password $mypwd -CertStoreLocation cert:\LocalMachine\Root
        }
        else 
        {
            Write-Host "Thumbprint is not the same between new cert and installed cert." -ForegroundColor Cyan
        }

    }
    else
    {
        $message = $certcn + " is not created"
        Write-Host $message -ForegroundColor Cyan
    }
 }
 #AutoScript Part
 Function GenerateSelfCertCode
 {
           Param (           
            $certcn,
            $password,
            $certfilepath
           )    

      $samplecode = $samplecode + "`r`n" +"Function GenerateSelfSignedCert {"
      $samplecode = $samplecode + "`r`n" +  (Get-Command GenerateSelfSignedCert).Definition +"}"
      $samplecode = $samplecode + "`r`n" + "GenerateSelfSignedCert "+$certcn+" "+$password+" "+ $certfilepath

    return $samplecode  
 }

# Process Bar
Function PBStart {
Param ($name)
    $x = Get-Random -minimum 1 -maximum 50

    write-progress -activity "PowerShell AutoScript"  -CurrentOperation $name -PercentComplete $x
}
Function PBStop{
    write-progress -activity "PowerShell AutoScript" -status "Completed" -PercentComplete 100
    write-progress -activity "PowerShell AutoScript" -status "Completed" -Completed
}

 #Form UI
 #************************************************************************
 Function SelfCertLaunchForm{

        $Form = New-Object system.Windows.Forms.Form

        $Form.Text = "SelfCert AutoScript in PowerShell"
        $Form.MinimizeBox = $False
        $Form.MaximizeBox = $False
        $Form.width = 450

        $Labe1 = New-Object System.Windows.Forms.Label
        $Labe1.AutoSize = $True
        $Labe1.Text = "Cert Subject:"
        $Labe1.Location = New-Object System.Drawing.Size(20,15)
        $Form.Controls.Add($Labe1)

        $TextBox1 = New-Object System.Windows.Forms.TextBox
        $TextBox1.AutoSize = $True
        $TextBox1.Text = "www.test.com"
        $TextBox1.Location = New-Object System.Drawing.Size(120,12)
        $Form.Controls.Add($TextBox1)

        $Labe2 = New-Object System.Windows.Forms.Label
        $Labe2.AutoSize = $True
        $Labe2.Text = "PFX Password:"
        $Labe2.Location = New-Object System.Drawing.Size(20,45)
        $Form.Controls.Add($Labe2)

        $TextBox2 = New-Object System.Windows.Forms.TextBox
        $TextBox2.AutoSize = $True
        $TextBox2.Text = "P@ssw0rd"
        $TextBox2.Location = New-Object System.Drawing.Size(120,42)
        $Form.Controls.Add($TextBox2)

        $Labe3 = New-Object System.Windows.Forms.Label
        $Labe3.AutoSize = $True
        $Labe3.Text = "PFX Export Path:"
        $Labe3.Location = New-Object System.Drawing.Size(20,75)
        $Form.Controls.Add($Labe3)

        $TextBox3 = New-Object System.Windows.Forms.TextBox
        $TextBox3.AutoSize = $True
        $TextBox3.Text = "C:\Temp\cert.pfx"
        $TextBox3.Location = New-Object System.Drawing.Size(120,72)
        $Form.Controls.Add($TextBox3)

        $CodeButton = New-Object System.Windows.Forms.Button
        $CodeButton.Location = New-Object System.Drawing.Size(20,105)
        $CodeButton.AutoSize = $True
        $CodeButton.Text = "AutoScript"
        $CodeButton.Add_Click({
                                  $richTextBox1.Text=  GenerateSelfCertCode $TextBox1.Text $TextBox2.Text $TextBox3.Text  
                            })
        $Form.Controls.Add($CodeButton)

        $richTextBox1 = New-Object System.Windows.Forms.RichTextBox
        $richTextBox1.Text = ""
        $richTextBox1.Width = 390
        $richTextBox1.Height = 240
        $richTextBox1.Location = New-Object System.Drawing.Size(20,135)
        $Form.Controls.Add($richTextBox1)

        $RunButton = New-Object System.Windows.Forms.Button
        $RunButton.Location = New-Object System.Drawing.Size(75,380)
        $RunButton.Text = "Run"
        $RunButton.Add_Click({
        pbstart "Execute AutoScripted Code"
        $Label4.Text = "Starting"
        Invoke-Expression -Command $richTextBox1.Text|Out-Host
        $Label4.Text = "Finish"
        pbstop
        })
        $Form.Controls.Add($RunButton)

        $QuitButton = New-Object System.Windows.Forms.Button
        $QuitButton.Location = New-Object System.Drawing.Size(210,380)
        $QuitButton.Text = "Quit"
        $QuitButton.Add_Click({$Form.Close()})
        $Form.Controls.Add($QuitButton)
        
        $Label4 = New-Object System.Windows.Forms.Label
        $Label4.Location = New-Object System.Drawing.Size(20,410)
        $Label4.Text = "Status"
        $Form.Controls.Add($Label4)     

        $Icon = [system.drawing.icon]::ExtractAssociatedIcon($PSHOME + "\powershell.exe")

        $Form.Icon = $Icon

        $Form.AutoSize = $True
        $Form.StartPosition = "CenterScreen"
        $Form.ShowDialog()
  }

 SelfCertLaunchForm

#GenerateSelfSignedCert www.mytest1.com YourSecPassword c:\temp\mytest1.pfx


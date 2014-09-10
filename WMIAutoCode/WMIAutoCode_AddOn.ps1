Add-Type -AssemblyName System.Windows.Forms 
$list = New-Object 'System.Collections.Generic.List[string]'
$list2 = New-Object 'System.Collections.Generic.List[string]'

Function Get-WmiNamespace {
    Param (
        $Namespace='root'
    )

    Get-WmiObject -Namespace $Namespace -Class __NAMESPACE |Sort-Object | ForEach-Object {
            ($ns = '{0}\{1}' -f $_.__NAMESPACE,$_.Name )             
            $list.Add($ns)
            Get-WmiNamespace $ns
    }
}
Function Get-DynamicClass {
    Param (
        $Namespace='root'
    )
        $list2.Clear()
        $Label4.Text = "Loading WMI Classes"
        $a = Get-WmiObject -Namespace $Namespace -List | Select-Object -Property __class, qualifiers

        foreach ($x in $a)
        {
          foreach ($y in $x.qualifiers)
          {
          if ($y.name -match "^dynamic")
               { 
                $list2.Add( $x.__CLASS)
                $Label4.Text = "Loading "+ $x.__CLASS
               }

          }
        }
        $list2.sort()
        $Label4.Text = ""
}
 Function GenerateWMICode {
   Param (
        $Namespace='root',
        $DyanmicClass = 'Default'
    )
    
    $samplecode = "`n`$computer = `"LocalHost`" " + "`n`$namespace = `"$Namespace`" "  +  "`nGet-WmiObject -class  $DyanmicClass  -computername  `$computer  -namespace  `$namespace"
    return $samplecode         
}

Function LaunchForm{

$Form = New-Object system.Windows.Forms.Form

$Form.Text = "PowerShell Code Generator"
$Form.MinimizeBox = $False
$Form.MaximizeBox = $False
$Form.width = 410

$Label = New-Object System.Windows.Forms.Label
$Label.Text = "WMI NameSpace"
$Label.AutoSize = $True
$Label.Location = New-Object System.Drawing.Size(20,10)
$Form.Controls.Add($Label)

Get-WmiNamespace

$comboBox1 = New-Object System.Windows.Forms.ComboBox
$comboBox1.Location = New-Object System.Drawing.Point(20, 30)
$comboBox1.Size = New-Object System.Drawing.Size(350, 310)
foreach($wmispace in $list)
{
  $comboBox1.Items.add($wmispace)
}

$Form.Controls.Add($comboBox1)

$Label2 = New-Object System.Windows.Forms.Label
$Label2.Text = "WMI Dynamic Class"
$Label2.AutoSize = $True
$Label2.Location = New-Object System.Drawing.Size(20,60)
$Form.Controls.Add($Label2)


$comboBox2 = New-Object System.Windows.Forms.ComboBox
$comboBox2.Location = New-Object System.Drawing.Point(20, 80)
$comboBox2.Size = New-Object System.Drawing.Size(350, 310)
$Form.Controls.Add($comboBox2)


$Label3 = New-Object System.Windows.Forms.Label
$Label3.Text = "Output Code"
$Label3.AutoSize = $True
$Label3.Location = New-Object System.Drawing.Size(20,110)
$Form.Controls.Add($Label3)

$CheckBoxOutputTable = New-Object System.Windows.Forms.CheckBox
$CheckBoxOutputTable.Text = "Output Table Format"
$CheckBoxOutputTable.AutoSize = $True
$CheckBoxOutputTable.Location = New-Object System.Drawing.Size(100,109)

$Form.Controls.Add($CheckBoxOutputTable)

$richTextBox1 = New-Object System.Windows.Forms.RichTextBox
$richTextBox1.Text = ""
$richTextBox1.Width = 350
$richTextBox1.Location = New-Object System.Drawing.Size(20,130)

$Form.Controls.Add($richTextBox1)

$OKButton = New-Object System.Windows.Forms.Button
$OKButton.Location = New-Object System.Drawing.Size(130,230)

$OKButton.Size = New-Object System.Drawing.Size(75,23)
$OKButton.Text = "Quit"
$OKButton.Add_Click({$Form.Close()})
$Form.Controls.Add($OKButton)

$InsertButton = New-Object System.Windows.Forms.Button

$InsertButton.Location = New-Object System.Drawing.Size(20,230)
$InsertButton.Size = New-Object System.Drawing.Size(100,23)
$InsertButton.Text = "Insert Command"
$InsertButton.Add_Click({$psise.CurrentFile.Editor.InsertText($richTextBox1.Text)})
$Form.Controls.Add($InsertButton)

$RunButton = New-Object System.Windows.Forms.Button

$RunButton.Location = New-Object System.Drawing.Size(215,230)
$RunButton.Size = New-Object System.Drawing.Size(75,23)
$RunButton.Text = "Run"
$RunButton.Add_Click({
$Label4.Text = "Starting"
Invoke-Expression -Command $richTextBox1.Text|Out-Host
$Label4.Text = "Finish"
})
$Form.Controls.Add($RunButton)

$Label4 = New-Object System.Windows.Forms.Label
$Label4.AutoSize = $False
$Label4.Size = New-Object System.Drawing.Size(200,23)
$Label4.Location = New-Object System.Drawing.Size(20,260)
$Form.Controls.Add($Label4)

# -------------------------------------------------------------------------------

#                             Change triggered function

# ------------------------------------------------------------------------------- 

$ComboBox1_SelectedIndexChanged=

{

   Get-DynamicClass $comboBox1.Text
   $comboBox2.Items.Clear()
   foreach($wmiclass in $list2)
    {
        $comboBox2.Items.add($wmiclass)
    }
   $comboBox2.Update()
   $Label2.Text = "WMI Dynamic Classes count: "+ $list2.Count

}
$ComboBox2_SelectedIndexChanged=

{
$returncode = GenerateWMICode $comboBox1.Text $comboBox2.Text
$richTextBox1.Clear()
  $richTextBox1.AppendText($returncode )
   If ($CheckBoxOutputTable.Checked) {
        $richTextBox1.AppendText(' | format-table -wrap -autosize')
    } 

}
$CheckBoxOutputTable_CheckStateChanged={
    If ($CheckBoxOutputTable.Checked) {
        $richTextBox1.AppendText(' | format-table -wrap -autosize')
    } 
}

################ MUST CREATE BEFORE ASSIGN ################

$ComboBox1.add_SelectedIndexChanged($ComboBox1_SelectedIndexChanged)
$ComboBox2.add_SelectedIndexChanged($ComboBox2_SelectedIndexChanged)
$CheckBoxOutputTable.add_CheckStateChanged($CheckBoxOutputTable_CheckStateChanged)

#$Form.Controls.Add($textbox)

$Form.AutoSize = $True

#$Form.AutoSizeMode = "GrowAndShrink"

$Icon = [system.drawing.icon]::ExtractAssociatedIcon($PSHOME + "\powershell.exe")

$Form.Icon = $Icon


$Form.StartPosition = "CenterScreen"
$Form.ShowDialog()
}
function Test-Admin {
  $currentUser = New-Object Security.Principal.WindowsPrincipal $([Security.Principal.WindowsIdentity]::GetCurrent())
  $return = $currentUser.IsInRole([Security.Principal.WindowsBuiltinRole]::Administrator)
  return $return
}
Function LaunchMain{
if ((Test-Admin) -eq $false)  {

         
    $oReturn=[System.Windows.Forms.MessageBox]::Show("WMI AutoCode needs to start Powershell ISE in Admin permission","Info",[System.Windows.Forms.MessageBoxButtons]::OKCancel)  
 
    switch ($oReturn){ 
 
    "OK" { 
        write-host "You pressed OK" 
        Start-Process powershell_ise.exe -Verb RunAs
    }  
    "Cancel" { 
        write-host "You pressed Cancel" 
    }  
} 

        
    
}
else
{
     LaunchForm|Out-Null
    
}
}

$psISE.CurrentPowerShellTab.AddOnsMenu.Submenus.Add("WMI Auto Script", `
{LaunchMain},"ALT+F5") | out-Null

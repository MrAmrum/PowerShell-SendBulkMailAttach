#----------------------------------------------
# Source File Informaion
# Author: Markus Wertgen
# Version: 1.0
# License: CopyLeft
# Interface: GUI
# Purpose: Send Bulk Mails with Attachments
#----------------------------------------------

#----------------------------------------------
#region Application Functions
#----------------------------------------------

#endregion Application Functions

#----------------------------------------------
# Generated Form Function
#----------------------------------------------
function Show-tabcontrol_psf {

	#----------------------------------------------
	#region Import the Assemblies
	#----------------------------------------------
	[void][reflection.assembly]::Load('System.Windows.Forms, Version=2.0.0.0, Culture=neutral, PublicKeyToken=b77a5c561934e089')
	[void][reflection.assembly]::Load('System.Data, Version=2.0.0.0, Culture=neutral, PublicKeyToken=b77a5c561934e089')
	[void][reflection.assembly]::Load('System.Drawing, Version=2.0.0.0, Culture=neutral, PublicKeyToken=b03f5f7f11d50a3a')
	[void][reflection.assembly]::Load('System.DirectoryServices, Version=2.0.0.0, Culture=neutral, PublicKeyToken=b03f5f7f11d50a3a')
	[void][reflection.assembly]::Load('System.ServiceProcess, Version=2.0.0.0, Culture=neutral, PublicKeyToken=b03f5f7f11d50a3a')
	#endregion Import Assemblies

	#----------------------------------------------
	#region Generated Form Objects
	#----------------------------------------------
	[System.Windows.Forms.Application]::EnableVisualStyles()
	$form1 = New-Object 'System.Windows.Forms.Form'
	$tabcontrol1 = New-Object 'System.Windows.Forms.TabControl'
	$tabpage1 = New-Object 'System.Windows.Forms.TabPage'
	$tabpage2 = New-Object 'System.Windows.Forms.TabPage'
	$tabpage3 = New-Object 'System.Windows.Forms.TabPage'
	$InitialFormWindowState = New-Object 'System.Windows.Forms.FormWindowState'
	#endregion Generated Form Objects

	#----------------------------------------------
	# User Generated Script
	#----------------------------------------------
	
    # Build Filepath
    $path = "C:\Users\"+$env:UserName+"\AppData\Roaming\BulkMailScript\"
    $file = "Config.xml"
    $filepath = $path + $file

    
    $script:Config = New-Object System.XML.XMLDocument

	$form1_Load={
		
    # Test if Config file exists if not create. 
    if (!(Test-Path $filepath)){
        # Create Folder for Config File
        if (!(Test-Path $path)){
            mkdir $path
        }
        # Set The XML-Formatting
        $xmlsettings = New-Object System.Xml.XmlWriterSettings
        $xmlsettings.Indent = $true
        $xmlsettings.IndentChars = "   "
        $newXML = New-Object System.Xml.XmlTextWriter("$filepath",$Null)
        # XML-Create File
        $newXML.WriteStartDocument()
        $newXML.WriteStartElement("Content")
        $newXML.WriteStartElement("Settings")
        $newXML.WriteElementString("Attachments","NA")
        $newXML.WriteElementString("Reciever","NA")
        $newXML.WriteElementString("Topic","NA")
        $newXML.WriteElementString("Moveto","NA")
        $newXML.WriteElementString("MovetoName","NA")
        $newXML.WriteEndElement()
        $newXML.WriteStartElement("Data")
        $newXML.WriteElementString("LastBulk","NA")
        $newXML.WriteEndElement()
        $newXML.WriteEndElement()
        $newXML.WriteEndDocument()
        $newXML.Flush()
        $newXML.Close()
        $form1.refresh()
	}

	# Load Config
	if (Test-Path $filepath){
		$Config.Load($filepath)
		$objTextBoxAttachments.Text = $Config.Content.Settings.Attachments
		$objTextBoxReciever.Text = $Config.Content.Settings.Reciever
        $objTextBoxTopic.Text = $Config.Content.Settings.Topic
        $objTextBoxMoveto.Text = $Config.Content.Settings.Moveto
        $objTextBoxMovetoName.Text = $Config.Content.Settings.MovetoName
        $objLabelDate.Text = $Config.Content.Data.LastBulk
        $objLabelInfo.Text = "Ready"
        $form1.refresh()
	}
    }
        	
	# Save Config
	$button1_RunOnClick={
		$Config.Content.Settings.Attachments = $objTextBoxAttachments.Text
		$Config.Content.Settings.Reciever = $objTextBoxReciever.Text
        $Config.Content.Settings.Topic = $objTextBoxTopic.Text
        $Config.Content.Settings.Moveto = $objTextBoxMoveto.Text
        $Config.Content.Settings.MovetoName = $objTextBoxMovetoName.Text
        $Config.Content.Data.LastBulk = $objLabelDate.Text
        $Config.Save($filepath)
        $objLabelInfo.Text = "Settings Saved"
        $objLabelInfo.refresh()
	}
	
    # Send Bulk
	$button2_RunOnClick={
            
            $scount= @(Get-ChildItem $Config.Content.Settings.Attachments | where {$_.mode -notmatch "d"}).Count
            If($scount -ge 1){
                $Result = [System.Windows.Forms.MessageBox]::Show("Attention $([System.Environment]::NewLine)U sure? > $scount < ","Question to User",4,[System.Windows.Forms.MessageBoxIcon]::Question)
 
                If($Result -eq "Yes")
                {
                    [String]$Datetemp = Get-Date -UFormat "%Y%m%d"
                    $moveToPath = $Config.Content.Settings.Moveto + "\" + $Config.Content.Settings.MovetoName + " " + $Datetemp
                    # Create Folder
                    if (!(Test-Path $moveToPath)){
                        mkdir $moveToPath
                    }
                    $dcount = 0
                    gci  $Config.Content.Settings.Attachments -force | where {$_.mode -notmatch "d"} | select fullname | ForEach-Object{
                         $Outlook = New-Object -ComObject Outlook.Application
                         $Mail = $Outlook.CreateItem(0)
                         $Mail.To = $Config.Content.Settings.Reciever
                         $Mail.Subject = $Config.Content.Settings.Topic
                         $Body = (Get-Content -Path .\Body.txt)
                         $Mail.HTMLBody = "<HTML lang=""de""><meta charset=""utf-8""><BODY><p>$Body</p>" 
                         $Mail.Attachments.Add($_.FullName)
                         $Mail.Send()
                         $dcount++

                         Move-Item -Path $_.FullName -Destination $moveToPath -Force

                     }
                     Write-Host "$scount - $dcount"
 

                     [String]$Datetemp = Get-Date -UFormat "%d.%m.%Y %H:%M:%S"
  	                 $Config.Content.Data.LastBulk = $Datetemp
                     $Config.Save($filepath)
                     $objLabelInfo.Text = "$dcount Mails Away ... I Think"
                     $objLabelDate.Text = $Config.Content.Data.LastBulk
                     $objLabelInfo.refresh()
                }
                else
                {
                   # Do nothing
                }
            } else {
                $Result = [System.Windows.Forms.MessageBox]::Show("Attention $([System.Environment]::NewLine)No files -> no work!","No files found!",0,[System.Windows.Forms.MessageBoxIcon]::Error)
            }

    }

    # Delete Config/Data File
	$button3_RunOnClick={
            if (Test-Path $filepath){
                Remove-Item $filepath
                Remove-Item $path
            }
            $objLabelInfo.Text = "Files Removed, RESTART!!!"
    }



	
	# --End User Generated Script--
	#----------------------------------------------
	#region Generated Events
	#----------------------------------------------
	
	$Form_StateCorrection_Load=
	{
		#Correct the initial state of the form to prevent the .Net maximized form issue
		$form1.WindowState = $InitialFormWindowState
	}
	
	$Form_Cleanup_FormClosed=
	{
		#Remove all event handlers from the controls
		try
		{
			$form1.remove_Load($form1_Load)
			$form1.remove_Load($Form_StateCorrection_Load)
			$form1.remove_FormClosed($Form_Cleanup_FormClosed)
		}
		catch { Out-Null  }
	}
	#endregion Generated Events

	#----------------------------------------------
	#region Generated Form Code
	#----------------------------------------------
	$form1.SuspendLayout()
	$tabcontrol1.SuspendLayout()
	#
	# form1
	#
	#Add Label Error
	$objLabelInfo = New-Object System.Windows.Forms.Label
	$objLabelInfo.Location = New-Object System.Drawing.Size(200,10)  
	$objLabelInfo.Size = New-Object System.Drawing.Size(150,20)  
	$objLabelInfo.Text = $InfoLable
	$form1.Controls.Add($objLabelInfo)

	$form1.Controls.Add($tabcontrol1)
	$form1.AutoScaleDimensions = '6, 13'
	$form1.AutoScaleMode = 'Font'
	$form1.ClientSize = '400, 500'
	$form1.Name = 'form1'
	$form1.Text = 'Bulk Mail Script'
	$form1.add_Load($form1_Load)

	#
	# tabcontrol1
	#
	$tabcontrol1.Controls.Add($tabpage1)
	$tabcontrol1.Controls.Add($tabpage2)
	$tabcontrol1.Controls.Add($tabpage3)
	$tabcontrol1.Alignment = 'Top'
	$tabcontrol1.Location = '12, 12'
	$tabcontrol1.Multiline = $True
	$tabcontrol1.Name = 'tabcontrol1'
	$tabcontrol1.SelectedIndex = 0
	$tabcontrol1.Size = '380, 480'
	$tabcontrol1.TabIndex = 0
	#
	# tabpage1
	#
	$tabpage1.Location = '42, 4'
	$tabpage1.Name = 'tabpage1'
	$tabpage1.Padding = '3, 3, 3, 3'
	$tabpage1.Size = '350, 450'
	$tabpage1.TabIndex = 0
	$tabpage1.Text = 'Controls'
	$tabpage1.UseVisualStyleBackColor = $True
	
	#Add Label Last Bulk Send
	$objLabel = New-Object System.Windows.Forms.Label
	$objLabel.Location = New-Object System.Drawing.Size(10,25)  
	$objLabel.Size = New-Object System.Drawing.Size(120,20)  
	$objLabel.Text = "Last Bulk Send:"
	$tabpage1.Controls.Add($objLabel)
	
	$objLabelDate = New-Object System.Windows.Forms.Label
	$objLabelDate.Location = New-Object System.Drawing.Size(130,25)  
	$objLabelDate.Size = New-Object System.Drawing.Size(200,20)  
	$objLabelDate.Text = $Config.Content.Data.LastBulkSend
	$tabpage1.Controls.Add($objLabelDate)
	
	#Button Send Bulk
	$button2 = New-Object System.Windows.Forms.Button
	$button2.Name = "button2"
	$button2.Location = New-Object System.Drawing.Size(10,55)  
	$button2.Size = New-Object System.Drawing.Size(150,20)  
	$button2.Text = "Send Bulk"
	$button2.add_Click($button2_RunOnClick)
	$tabpage1.Controls.Add($button2)	
	
	#
	# tabpage2
	#
	$tabpage2.Location = '23, 4'
	$tabpage2.Name = 'tabpage2'
	$tabpage2.Padding = '3, 3, 3, 3'
	$tabpage2.Size = '350, 450'
	$tabpage2.TabIndex = 1
	$tabpage2.Text = 'Settings'
	$tabpage2.UseVisualStyleBackColor = $True
	
	#Add Label / TextBox Path to Attachments
	$objLabel = New-Object System.Windows.Forms.Label
	$objLabel.Location = New-Object System.Drawing.Size(10,25)  
	$objLabel.Size = New-Object System.Drawing.Size(200,20)  
	$objLabel.Text = "Path to Attachments"
	$tabpage2.Controls.Add($objLabel)
	$objTextBoxAttachments = New-Object System.Windows.Forms.TextBox 
	$objTextBoxAttachments.Location = New-Object System.Drawing.Size(10,45) 
	$objTextBoxAttachments.Size = New-Object System.Drawing.Size(200,20)
    $objTextBoxAttachments.Text = $Config.Content.Settings.Attachments
	$tabpage2.Controls.Add($objTextBoxAttachments) 
	
	#Add Label / TextBox Reciever
	$objLabel = New-Object System.Windows.Forms.Label
	$objLabel.Location = New-Object System.Drawing.Size(10,95)  
	$objLabel.Size = New-Object System.Drawing.Size(200,20)  
	$objLabel.Text = "To"
	$tabpage2.Controls.Add($objLabel)
	$objTextBoxReciever = New-Object System.Windows.Forms.TextBox 
	$objTextBoxReciever.Location = New-Object System.Drawing.Size(10,115) 
	$objTextBoxReciever.Size = New-Object System.Drawing.Size(200,20) 
    $objTextBoxReciever.Text = $Config.Content.Settings.Reciever
	$tabpage2.Controls.Add($objTextBoxReciever) 

    #Add Label / TextBox Topic
	$objLabel = New-Object System.Windows.Forms.Label
	$objLabel.Location = New-Object System.Drawing.Size(10,165)  
	$objLabel.Size = New-Object System.Drawing.Size(200,20)  
	$objLabel.Text = "Topic"
	$tabpage2.Controls.Add($objLabel)
	$objTextBoxTopic = New-Object System.Windows.Forms.TextBox 
	$objTextBoxTopic.Location = New-Object System.Drawing.Size(10,185) 
	$objTextBoxTopic.Size = New-Object System.Drawing.Size(200,20) 
    $objTextBoxTopic.Text = $Config.Content.Settings.Topic
	$tabpage2.Controls.Add($objTextBoxTopic) 

    #Add Label / TextBox Move File to
	$objLabel = New-Object System.Windows.Forms.Label
	$objLabel.Location = New-Object System.Drawing.Size(10,255)  
	$objLabel.Size = New-Object System.Drawing.Size(200,20)  
	$objLabel.Text = "Move File to after Sending"
	$tabpage2.Controls.Add($objLabel)
	$objTextBoxMoveto = New-Object System.Windows.Forms.TextBox 
	$objTextBoxMoveto.Location = New-Object System.Drawing.Size(10,275) 
	$objTextBoxMoveto.Size = New-Object System.Drawing.Size(200,20) 
    $objTextBoxMoveto.Text = $Config.Content.Settings.Moveto
	$tabpage2.Controls.Add($objTextBoxMoveto) 

    #Add Label / TextBox Move File to Name
	$objLabel = New-Object System.Windows.Forms.Label
	$objLabel.Location = New-Object System.Drawing.Size(10,325)  
	$objLabel.Size = New-Object System.Drawing.Size(200,20)  
	$objLabel.Text = "Folder Name Template"
	$tabpage2.Controls.Add($objLabel)
	$objTextBoxMovetoName = New-Object System.Windows.Forms.TextBox 
	$objTextBoxMovetoName.Location = New-Object System.Drawing.Size(10,345) 
	$objTextBoxMovetoName.Size = New-Object System.Drawing.Size(200,20) 
    $objTextBoxMovetoName.Text = $Config.Content.Settings.MovetoName
	$tabpage2.Controls.Add($objTextBoxMovetoName) 
	
	#Button Save Settings
	$button1 = New-Object System.Windows.Forms.Button
	$button1.Name = "button1"
	$button1.Location = New-Object System.Drawing.Size(220,45)  
	$button1.Size = New-Object System.Drawing.Size(150,20)  
	$button1.Text = "Save Settings"
	$button1.add_Click($button1_RunOnClick)
	$tabpage2.Controls.Add($button1)

	#
	# tabpage3
	#
	$tabpage3.Location = '23, 4'
	$tabpage3.Name = 'tabpage3'
	$tabpage3.Padding = '3, 3, 3, 3'
	$tabpage3.Size = '350, 450'
	$tabpage3.TabIndex = 2
	$tabpage3.Text = 'Help'
	$tabpage3.UseVisualStyleBackColor = $True
	#

    # Button Remove File
	$button3 = New-Object System.Windows.Forms.Button
	$button3.Name = "button3"
	$button3.Location = New-Object System.Drawing.Size(10,25)  
	$button3.Size = New-Object System.Drawing.Size(120,20)  
	$button3.Text = "Remove Files (WIP)"
    #	$button3.add_Click($button3_RunOnClick)
	$tabpage3.Controls.Add($button3)

	$tabcontrol1.ResumeLayout()
	$form1.ResumeLayout()
	#endregion Generated Form Code

	#----------------------------------------------

	#Save the initial state of the form
	$InitialFormWindowState = $form1.WindowState
	#Init the OnLoad event to correct the initial state of the form
	$form1.add_Load($Form_StateCorrection_Load)
	#Clean up the control events
	$form1.add_FormClosed($Form_Cleanup_FormClosed)
	#Show the Form
	return $form1.ShowDialog()

} #End Function

#Call the form
Show-tabcontrol_psf | Out-Null
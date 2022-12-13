##############################################################################################################################
#                                                       Chargement des classe                                                #
##############################################################################################################################

# Chargement des classe 
Add-Type -AssemblyName System.Windows.Forms
Add-Type -AssemblyName System.Drawing

##############################################################################################################################
#                                                              Function                                                      #
##############################################################################################################################

function convert {$ConvertTo}

##############################################################################################################################
#                                                          Window Settings                                                   #
##############################################################################################################################

# Création de la fenêtre pour contenir les éléments
$main_form = New-Object System.Windows.Forms.Form

# Le titre de la fenêtre
$main_form.Text ='Table Format Converter'

# Largeur de la fenêtre
$main_form.Width = 400

# Hauteur de la fenêtre
$main_form.Height = 400

# Étire automatiquement la fenêtre
$main_form.AutoSize = $true

# Couleur du fond
$main_form.BackColor = "gray"

# Icon du GUI
$main_form.Icon = [System.Drawing.Icon]::ExtractAssociatedIcon('C:\Users\cp-20ahb\Desktop\refresh.ico')


# Bloque la taille max et min
$main_form.minimumSize = New-Object System.Drawing.Size(565,365)
$main_form.maximumSize = New-Object System.Drawing.Size(565,365)

##############################################################################################
#                                              Toolbox                                       #
##############################################################################################

###############################################################
#                            Label                            #
###############################################################

###############################################################
#                            Button                           #
###############################################################

$ButtonLocation = New-Object System.Windows.Forms.Button

$ButtonLocation.Location = New-Object System.Drawing.Size(215,100)

$ButtonLocation.Size = New-Object System.Drawing.Size(120,40)

$ButtonLocation.Text = "Location"

$ButtonLocation.ForeColor = [System.Drawing.Color]::FromArgb(243,5,81) 

$ButtonLocation.BackColor = "White"

$ButtonLocation.Font = 'Bahnschrift,11'


$FolderBrowser = New-Object System.Windows.Forms.FolderBrowserDialog

# Event click
$ButtonLocation.Add_Click({
    $FolderBrowser.ShowDialog()
    })


##################################################################



$OKButton              = New-Object System.Windows.Forms.Button
$OKButton.Location     = New-Object System.Drawing.Size(128,100)
$OKButton.Size         = New-Object System.Drawing.Size(75,23)
$OKButton.Text         = "Convertir"
$OKButton.Add_Click({ 


#Output

$SelectedOutput = $ComboboxTypeOutput.SelectedItem
$script:x += $ComboboxTypeOutput.SelectedItem
if (![string]::IsNullOrWhiteSpace($SelectedOutput)) {
            switch ($SelectedOutput) {
                "csv" { $ConvertTo = 'ConvertTo-CSV' }
                "json" { $ConvertTo = "ConvertTo-Json" }
                "xml" { $ConvertTo = "ConvertTo-Xml" }
                "xls" { $ConvertTo = "" }
            }
            $ComboboxTypeOutput.SelectedIndex = -1   # reset the combobox to blank
        }


#Input
$SelectedInput = $ComboboxTypeInput.SelectedItem
$script:x += $ComboboxTypeInput.SelectedItem
            #CSV
            if ($SelectedOutput -eq "json" -And $SelectedInput -eq "csv") 
            {
                import-csv -Delimiter ";" "C:\ConverterBaseFormat\base.csv" | ConvertTo-Json | Add-Content -Path "C:\ConverterOutPutFormat\outputcsv.json"
                $LabelInfo.Text = "Reussi"
            }
            elseif ($SelectedOutput -eq "xml" -And $SelectedInput -eq "csv") 
            {
                import-csv -Delimiter ";" "C:\ConverterBaseFormat\base.csv" | Export-Clixml "C:\ConverterOutPutFormat\outputcsv.xml" 
                $LabelInfo.Text = "Reussi"
            }
            elseif ($SelectedOutput -eq "xls" -And $SelectedInput -eq "csv") 
            {
                $LabelInfo.Text = "Pas fini"
            }

            #json

            if ($SelectedOutput -eq "csv" -And $SelectedInput -eq "json") 
            {
                Get-Content "C:\ConverterBaseFormat\base.json" | ConvertFrom-Json | ConvertTo-Csv -Delimiter ";" | Out-File "C:\ConverterOutPutFormat\outputjson.csv" 
                $LabelInfo.Text = "Reussi"
            }
            elseif ($SelectedOutput -eq "xml" -And $SelectedInput -eq "json") 
            {
                Get-Content "C:\ConverterBaseFormat\base.json" | ConvertFrom-Json | Export-Clixml "C:\ConverterOutPutFormat\outputjson.xml" 
                $LabelInfo.Text = "Reussi"
            }
            elseif ($SelectedOutput -eq "xls" -And $SelectedInput -eq "json") 
            {
                $LabelInfo.Text = "Pas fini"
            }

            #xml

            if ($SelectedOutput -eq "csv" -And $SelectedInput -eq "xml") 
            {
                Import-Clixml "C:\ConverterBaseFormat\base.xml" | ConvertTo-Csv -Delimiter ";" | Add-Content -Path "C:\ConverterOutPutFormat\outputxml.csv" 
                $LabelInfo.Text = "Reussi"
            }
            elseif ($SelectedOutput -eq "json" -And $SelectedInput -eq "xml")
            {
                Import-Clixml "C:\ConverterBaseFormat\base.xml" | ConvertTo-Json | Out-File "C:\ConverterOutPutFormat\outputxml.json" 
                $LabelInfo.Text = "Reussi"
            }
            elseif ($SelectedOutput -eq "xls" -And $SelectedInput -eq "xml")
            {
                $LabelInfo.Text = "Pas fini"
            }

            #xls













            #Input = Output Error
            elseif ($SelectedOutput -eq "csv" -And $SelectedInput -eq "csv") {
                [System.Windows.Forms.MessageBox]::Show('Vous ne pouvez pas faire cela','Erreur','Ok','Error')
            }  
            elseif ($SelectedOutput -eq "json" -And $SelectedInput -eq "json") {
                [System.Windows.Forms.MessageBox]::Show('Vous ne pouvez pas faire cela','Erreur','Ok','Error')
            }
            elseif ($SelectedOutput -eq "xml" -And $SelectedInput -eq "xml") {
                [System.Windows.Forms.MessageBox]::Show('Vous ne pouvez pas faire cela','Erreur','Ok','Error')
            }  
            elseif ($SelectedOutput -eq "xls" -And $SelectedInput -eq "xls") {
                [System.Windows.Forms.MessageBox]::Show('Vous ne pouvez pas faire cela','Erreur','Ok','Error')
            }  
})

###############################################################
#                            Label                            #
###############################################################

#Label format d'entrée

$LabelFormatInput          = New-Object System.Windows.Forms.Label
$LabelFormatInput.Location = New-Object System.Drawing.Size(10,20)
$LabelFormatInput.Size     = New-Object System.Drawing.Size(100,20)
$LabelFormatInput.Text     = "Format de base"

#Label format de sorti


$LabelFormatOutput        = New-Object System.Windows.Forms.Label
$LabelFormatOutput.Location = New-Object System.Drawing.Size(200,20)
$LabelFormatOutput.Size     = New-Object System.Drawing.Size(100,20)
$LabelFormatOutput.Text     = "Format de sorti"

# Notif

$LabelInfo        = New-Object System.Windows.Forms.Label
$LabelInfo.Location = New-Object System.Drawing.Size(260,90)
$LabelInfo.Size     = New-Object System.Drawing.Size(280,20)
$LabelInfo.Text     = "Notif"



###############################################################
#                            Combobox                         #
###############################################################


#Input#

$ComboboxTypeInput          = New-Object System.Windows.Forms.Combobox
$ComboboxTypeInput.Location = New-Object System.Drawing.Size(10,40)
$ComboboxTypeInput.Size     = New-Object System.Drawing.Size(120,20)
$ComboboxTypeInput.Height   = 70


#[void] $ComboboxTypeInput.Items.Add("Item 1")
$ComboboxTypeInput.Items.Add("csv")
$ComboboxTypeInput.Items.Add("json")
$ComboboxTypeInput.Items.Add("xml")
$ComboboxTypeInput.Items.Add("xls")


#Output#

$ComboboxTypeOutput          = New-Object System.Windows.Forms.Combobox
$ComboboxTypeOutput.Location = New-Object System.Drawing.Size(200,40)
$ComboboxTypeOutput.Size     = New-Object System.Drawing.Size(120,20)
$ComboboxTypeOutput.Height   = 70



$ComboboxTypeOutput.Items.Add("csv")
$ComboboxTypeOutput.Items.Add("json")
$ComboboxTypeOutput.Items.Add("xml")
$ComboboxTypeOutput.Items.Add("xls")

##############################################################################################
#                                              Control ToolBox                               #
##############################################################################################

# Déclare les variables du ToolBox
$main_form.controls.AddRange(@(


$LabelFormatInput
$LabelFormatOutput
$LabelInfo

$OKButton
$ButtonLocation

$ComboboxTypeInput
$ComboboxTypeOutput

))

# Affiche la fenêtre
$main_form.ShowDialog()
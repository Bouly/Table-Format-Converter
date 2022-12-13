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
$LabelInfo.Location = New-Object System.Drawing.Size(355,80)
$LabelInfo.Size     = New-Object System.Drawing.Size(200,20)
$LabelInfo.Text     = "Chemin d'entrée non spécifié"
$LabelInfo.ForeColor = "Red"

$LabelInfo2        = New-Object System.Windows.Forms.Label
$LabelInfo2.Location = New-Object System.Drawing.Size(355, 180)
$LabelInfo2.Size     = New-Object System.Drawing.Size(200,20)
$LabelInfo2.Text     = "Chemin de sorti non spécifié"
$LabelInfo2.ForeColor = "Red"

###############################################################
#                            Button                           #
###############################################################

$ButtonLocation = New-Object System.Windows.Forms.Button

$ButtonLocation.Location = New-Object System.Drawing.Size(355,100)

$ButtonLocation.Size = New-Object System.Drawing.Size(75,23)

$ButtonLocation.Text = "Location"

#$ButtonLocation.ForeColor = [System.Drawing.Color]::FromArgb(243,5,81) 

#$ButtonLocation.BackColor = "White"

#$ButtonLocation.Font = 'Bahnschrift,11'


$FilePath = New-Object System.Windows.Forms.OpenFileDialog

# Event click
$ButtonLocation.Add_Click({
    $FilePath.ShowDialog()
    if ($FilePath.FileName -eq $FilePath.FileName) {
        $LabelInfo.Text = $FilePath.FileName                #Comprend pas (fait au hasard)
        $LabelInfo.ForeColor = "green"
    }   
    })



#################################################################


$ButtonLocation2 = New-Object System.Windows.Forms.Button

$ButtonLocation2.Location = New-Object System.Drawing.Size(355,200)

$ButtonLocation2.Size = New-Object System.Drawing.Size(75,23)

$ButtonLocation2.Text = "Location"

#$ButtonLocation2.ForeColor = [System.Drawing.Color]::FromArgb(243,5,81) 

#$ButtonLocation2.BackColor = "White"

#$ButtonLocation2.Font = 'Bahnschrift,11'


$FolderPath = New-Object System.Windows.Forms.FolderBrowserDialog

# Event click
$ButtonLocation2.Add_Click({
    $FolderPath.ShowDialog()
    if ($FolderPath.SelectedPath -eq $FolderPath.SelectedPath) {
        $LabelInfo2.Text = $FolderPath.SelectedPath                #Comprend pas (fait au hasard)
        $LabelInfo2.ForeColor = "green"
    }   
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
    if($FilePath.FileName -like "*csv*" -or $FilePath.FileName -like "*json*" -or $FilePath.FileName -like "*xml*")
    
        {

            $Destionation = $FolderPath.SelectedPath
            if ($SelectedOutput -eq "json" -And $SelectedInput -eq "csv") 
            {
                import-csv -Delimiter ";" $FilePath.FileName | ConvertTo-Json | Add-Content -Path "$Destionation\outputcsv.json"
            }
            elseif ($SelectedOutput -eq "xml" -And $SelectedInput -eq "csv") 
            {
                import-csv -Delimiter ";" "C:\ConverterBaseFormat\base.csv" | Export-Clixml "C:\ConverterOutPutFormat\outputcsv.xml" 
            }
            elseif ($SelectedOutput -eq "xls" -And $SelectedInput -eq "csv")
            {
            }

            #json

            elseif ($SelectedOutput -eq "csv" -And $SelectedInput -eq "json") 
            {
                Get-Content "C:\ConverterBaseFormat\base.json" | ConvertFrom-Json | ConvertTo-Csv -Delimiter ";" | Out-File "C:\ConverterOutPutFormat\outputjson.csv" 
            }
            elseif ($SelectedOutput -eq "xml" -And $SelectedInput -eq "json") 
            {
                Get-Content "C:\ConverterBaseFormat\base.json" | ConvertFrom-Json | Export-Clixml "C:\ConverterOutPutFormat\outputjson.xml" 
            }
            elseif ($SelectedOutput -eq "xls" -And $SelectedInput -eq "json") 
            {
            }

            #xml

            elseif ($SelectedOutput -eq "csv" -And $SelectedInput -eq "xml") 
            {
                Import-Clixml "C:\ConverterBaseFormat\base.xml" | ConvertTo-Csv -Delimiter ";" | Add-Content -Path "C:\ConverterOutPutFormat\outputxml.csv" 
            }
            elseif ($SelectedOutput -eq "json" -And $SelectedInput -eq "xml")
            {
                Import-Clixml "C:\ConverterBaseFormat\base.xml" | ConvertTo-Json | Out-File "C:\ConverterOutPutFormat\outputxml.json" 
            }
            elseif ($SelectedOutput -eq "xls" -And $SelectedInput -eq "xml")
            {
            }


            #Input = Output Error
            elseif ($SelectedOutput -eq "csv" -And $SelectedInput -eq "csv") 
            {
                [System.Windows.Forms.MessageBox]::Show('Vous ne pouvez pas faire cela','Erreur','Ok','Error')
            }  
            elseif ($SelectedOutput -eq "json" -And $SelectedInput -eq "json") 
            {
                [System.Windows.Forms.MessageBox]::Show('Vous ne pouvez pas faire cela','Erreur','Ok','Error')
            }
            elseif ($SelectedOutput -eq "xml" -And $SelectedInput -eq "xml") 
            {
                [System.Windows.Forms.MessageBox]::Show('Vous ne pouvez pas faire cela','Erreur','Ok','Error')
            }  
            elseif ($SelectedOutput -eq "xls" -And $SelectedInput -eq "xls") 
            {
                [System.Windows.Forms.MessageBox]::Show('Vous ne pouvez pas faire cela','Erreur','Ok','Error')
            } 
        }
    else 
    {
        $LabelInfo.Text = "Chemin non défini ou invalide"
        $LabelInfo.ForeColor = "red"
        [System.Windows.Forms.MessageBox]::Show('Chemin Input non défini ou invalide','Erreur','Ok','Error')
    } 
})





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
$LabelInfo2

$OKButton
$ButtonLocation
$ButtonLocation2

$ComboboxTypeInput
$ComboboxTypeOutput

))

# Affiche la fenêtre
$main_form.ShowDialog()
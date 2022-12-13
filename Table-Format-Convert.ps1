##############################################################################################################################
#                                                       Chargement des classe                                                #
##############################################################################################################################
# Chargement des classe l'interface GUI
Add-Type -AssemblyName System.Windows.Forms
Add-Type -AssemblyName System.Drawing
##############################################################################################################################
#                                                              Function                                                      #
##############################################################################################################################
#Funtion
##############################################################################################################################
#                                                          Window Settings                                                   #
##############################################################################################################################
# Création de la fenêtre pour contenir les éléments
$main_form                  = New-Object System.Windows.Forms.Form
# Le titre de la fenêtre
$main_form.Text             ='Table Format Converter'
# Largeur de la fenêtre
$main_form.Width            = 400
# Hauteur de la fenêtre
$main_form.Height           = 400
# Étire automatiquement la fenêtre
$main_form.AutoSize         = $true
# Couleur du fond
$main_form.BackColor        = "gray"
# Icon du GUI
$main_form.Icon             = [System.Drawing.Icon]::ExtractAssociatedIcon('C:\Users\cp-20ahb\Desktop\refresh.ico')
# Bloque la taille max et min
$main_form.minimumSize      = New-Object System.Drawing.Size(565,365)
$main_form.maximumSize      = New-Object System.Drawing.Size(565,365)
##############################################################################################
#                                              Toolbox                                       #
##############################################################################################
###############################################################
#                            Label                            #
###############################################################
##########################
#   Label Input format   #
##########################
#Création du label pour le format d'entrée
$LabelFormatInput           = New-Object System.Windows.Forms.Label
#Location du label
$LabelFormatInput.Location  = New-Object System.Drawing.Size(10,20)
#Taille du label
$LabelFormatInput.Size      = New-Object System.Drawing.Size(100,20)
#Text du Label
$LabelFormatInput.Text      = "Format d'entrée"
##########################
#   Label Output format  #
##########################
#Création du label pour le format de sortie
$LabelFormatOutput          = New-Object System.Windows.Forms.Label
#Location du label
$LabelFormatOutput.Location = New-Object System.Drawing.Size(200,20)
#Taille du label
$LabelFormatOutput.Size     = New-Object System.Drawing.Size(100,20)
#Text du label
$LabelFormatOutput.Text     = "Format de sortie"
##########################
#    Label Info Input    #
##########################
#Création du label pour informé l'état du chemin d'entrée
$LabelInfo                  = New-Object System.Windows.Forms.Label
#Location du label
$LabelInfo.Location         = New-Object System.Drawing.Size(355,80)
#Taille du label
$LabelInfo.Size             = New-Object System.Drawing.Size(200,20)
#Text du label
$LabelInfo.Text             = "Chemin d'entrée non spécifié"
#Couleur du text
$LabelInfo.ForeColor        = "Red"
##########################
#    Label Info Output   #
##########################
#Création du label pour informé l'état du chemin de sortie
$LabelInfo2                 = New-Object System.Windows.Forms.Label
#Location du label
$LabelInfo2.Location        = New-Object System.Drawing.Size(355, 180)
#Taille du label
$LabelInfo2.Size            = New-Object System.Drawing.Size(200,20)
#Text du label
$LabelInfo2.Text            = "Chemin de sorti non spécifié"
#Couleur du text
$LabelInfo2.ForeColor       = "Red"
###############################################################
#                            Button                           #
###############################################################
##########################
#    Button Input Loc    #
##########################
#Création du button pour le chemin de sorti
$ButtonLocation             = New-Object System.Windows.Forms.Button
#Location du button
$ButtonLocation.Location    = New-Object System.Drawing.Size(355,100)
#Taille du button
$ButtonLocation.Size        = New-Object System.Drawing.Size(75,23)
#Text du button
$ButtonLocation.Text        = "Location"

#Création du dialogue pour la séléction du chemin
$FilePath                   = New-Object System.Windows.Forms.OpenFileDialog

# Event click
$ButtonLocation.Add_Click #Quand le button cliqué
({
    $FilePath.ShowDialog() # Affiche la page de dialogue pour la séléction du chemin
    if ($FilePath.FileName -eq $FilePath.FileName) # Si le chemin = le chemin alors
    {
        $LabelInfo.Text = $FilePath.FileName # Le label d'information de l'état du chemin = Le chemin choisi
        $LabelInfo.ForeColor = "green" # Couleur du text du label d'information de l'état du chemin en "vert"
    }   
})
##########################
#   Button Output Loc    #
##########################
#Création du button pour le chemin de sorti
$ButtonLocation2            = New-Object System.Windows.Forms.Button
#Location du button
$ButtonLocation2.Location   = New-Object System.Drawing.Size(355,200)
#Taille du button
$ButtonLocation2.Size       = New-Object System.Drawing.Size(75,23)
#Text du button
$ButtonLocation2.Text       = "Location"
#Création du dialogue pour la séléction du chemin
$FolderPath = New-Object System.Windows.Forms.FolderBrowserDialog
# Event click
$ButtonLocation2.Add_Click
({
    $FolderPath.ShowDialog() # Affiche la page de dialogue pour la séléction du chemin
    if ($FolderPath.SelectedPath -eq $FolderPath.SelectedPath) # Si le chemin = le chemin alors
    {
        $LabelInfo2.Text = $FolderPath.SelectedPath # Le label d'information de l'état du chemin = Le chemin choisi
        $LabelInfo2.ForeColor = "green" # Couleur du text du label d'information de l'état du chemin en "vert"
    }   
    })
##########################
#    Button Conversion   #
##########################
#Création du button pour convertir
$OKButton              = New-Object System.Windows.Forms.Button
#Location du button
$OKButton.Location     = New-Object System.Drawing.Size(128,100)
#Taille du button
$OKButton.Size         = New-Object System.Drawing.Size(75,23)
#Text du button
$OKButton.Text         = "Convertir"
# Event click
$OKButton.Add_Click({ 
#Output
$SelectedOutput = $ComboboxTypeOutput.SelectedItem # On stock l'option séléctionné pour le format de sortie dans une variable
$script:x += $ComboboxTypeOutput.SelectedItem # Pas sur à chercher
if (![string]::IsNullOrWhiteSpace($SelectedOutput)) { # Pas sur à chercher
            switch ($SelectedOutput) { # list
                "csv" {}
                "json" {}
                "xml" {}
                "xls" {}
            }
            $ComboboxTypeOutput.SelectedIndex = -1   # vide la séléction du combobox
        }
#Input
$SelectedInput = $ComboboxTypeInput.SelectedItem # On stock l'option séléctionné pour le format de sortie dans une variable
$script:x += $ComboboxTypeInput.SelectedItem # Pas sur à chercher
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
                import-csv -Delimiter ";" $FilePath.FileName | Export-Clixml "$Destionation\outputcsv.xml" 
            }
            elseif ($SelectedOutput -eq "xls" -And $SelectedInput -eq "csv")
            {
            }

#json

            elseif ($SelectedOutput -eq "csv" -And $SelectedInput -eq "json") 
            {
                Get-Content $FilePath.FileName | ConvertFrom-Json | ConvertTo-Csv -Delimiter ";" | Out-File "$Destionation\outputjson.csv" 
            }
            elseif ($SelectedOutput -eq "xml" -And $SelectedInput -eq "json") 
            {
                Get-Content $FilePath.FileName | ConvertFrom-Json | Export-Clixml "$Destionation\outputjson.xml" 
            }
            elseif ($SelectedOutput -eq "xls" -And $SelectedInput -eq "json") 
            {
            }

#xml

            elseif ($SelectedOutput -eq "csv" -And $SelectedInput -eq "xml") 
            {
                Import-Clixml $FilePath.FileName | ConvertTo-Csv -Delimiter ";" | Add-Content -Path "$Destionation\outputxml.csv" 
            }
            elseif ($SelectedOutput -eq "json" -And $SelectedInput -eq "xml")
            {
                Import-Clixml $FilePath.FileName | ConvertTo-Json | Out-File "$Destionation\outputxml.json" 
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
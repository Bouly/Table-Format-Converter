##############################################################################################################################
#                                                       Chargement des classe                                                #
##############################################################################################################################
# Chargement des classe l'interface GUI
Add-Type -AssemblyName System.Windows.Forms
Add-Type -AssemblyName System.Drawing
##############################################################################################################################
#                                                              Function                                                      #
##############################################################################################################################
#Var
$Delimiter = ";"
#Funtion
function ModuleMissing_Visible {
    $LabelModuleCheck.Visible = $true
    $ButtonInstallModule.Visible = $true
}
function ModuleMissing_Invisible {
    $LabelModuleCheck.Visible = $false
    $ButtonInstallModule.Visible = $false
}
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
$main_form.minimumSize      = New-Object System.Drawing.Size(585,365)
$main_form.maximumSize      = New-Object System.Drawing.Size(585,365)
##############################################################################################
#                                              Toolbox                                       #
##############################################################################################
###############################################################
#                            TextBox                          #
###############################################################
#Création du champ de text pour le nom du fichier
$TextBoxOutPutFileName          = New-Object System.windows.Forms.TextBox
#Location du champ de text
$TextBoxOutPutFileName.Location = New-Object System.Drawing.Size(390,40)
#Taille du champ de text
$TextBoxOutPutFileName.Size     = New-Object System.Drawing.Size(137,20)
#L'entrée du champ de text
$TextBoxOutPutFileName.Text     = "Output"
###############################################################
#                            Label                            #
###############################################################
##########################
#   Label Module Check   #
##########################
#Création du label pour le nom du fichier
$LabelModuleCheck           = New-Object System.Windows.Forms.Label
#Location du label
$LabelModuleCheck.Location  = New-Object System.Drawing.Size(390,80)
#Taille du label
$LabelModuleCheck.Size      = New-Object System.Drawing.Size(200,20)
#Text du Label
$LabelModuleCheck.Text      = "Module ImportExcel non installé"
#Couleur du Label
$LabelModuleCheck.ForeColor        = "Red"
##########################
#   Label Output Name    #
##########################
#Création du label pour le nom du fichier
$LabelOutputName            = New-Object System.Windows.Forms.Label
#Location du label
$LabelOutputName.Location  = New-Object System.Drawing.Size(390,20)
#Taille du label
$LabelOutputName.Size      = New-Object System.Drawing.Size(137,20)
#Text du Label
$LabelOutputName.Text      = "Nom du fichier de sortie"
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
#Couleur du text
$LabelFormatOutput.ForeColor        = "black"
##########################
#    Label Info Input    #
##########################
#Création du label pour informé l'état du chemin d'entrée
$LabelInfo                  = New-Object System.Windows.Forms.Label
#Location du label
$LabelInfo.Location         = New-Object System.Drawing.Size(10,80)
#Taille du label
$LabelInfo.Size             = New-Object System.Drawing.Size(192,20)
#Text du label
$LabelInfo.Text             = "Chemin d'entrée non spécifié"
#Couleur du text
$LabelInfo.ForeColor        = "black"
##########################
#    Label Info Output   #
##########################
#Création du label pour informé l'état du chemin de sortie
$LabelInfo2                 = New-Object System.Windows.Forms.Label
#Location du label
$LabelInfo2.Location        = New-Object System.Drawing.Size(200, 80)
#Taille du label
$LabelInfo2.Size            = New-Object System.Drawing.Size(193,20)
#Text du label
$LabelInfo2.Text            = "Chemin de sorti non spécifié"
#Couleur du text
$LabelInfo2.ForeColor       = "black"
###############################################################
#                            Button                           #
###############################################################
##########################
# Button Install Module  #
##########################
#Création du button pour l'installation du module manquant
$ButtonInstallModule             = New-Object System.Windows.Forms.Button
#Location du button
$ButtonInstallModule.Location    = New-Object System.Drawing.Size(390,100)
#Taille du button
$ButtonInstallModule.Size        = New-Object System.Drawing.Size(75,23)
#Text du button
$ButtonInstallModule.Text        = "Install"
# Event click
$ButtonInstallModule.Add_Click({ #Quand le button cliqué
    Install-Module ImportExcel -AllowClobber -Force
    [System.Windows.Forms.MessageBox]::Show("Le module $ModuleCheck a bien était Installé",'Module installé','Ok','Information')
    ModuleMissing_Invisible
    $ModuleCheck = "true"
})
##########################
#    Button Input Loc    #
##########################
#Création du button pour le chemin de sorti
$ButtonLocation             = New-Object System.Windows.Forms.Button
#Location du button
$ButtonLocation.Location    = New-Object System.Drawing.Size(10,100)
#Taille du button
$ButtonLocation.Size        = New-Object System.Drawing.Size(75,23)
#Text du button
$ButtonLocation.Text        = "Location"
#Création du dialogue pour la séléction du chemin
$FilePath                   = New-Object System.Windows.Forms.OpenFileDialog
# Event click
$ButtonLocation.Add_Click({ #Quand le button cliqué
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
$ButtonLocation2.Location   = New-Object System.Drawing.Size(200,100)
#Taille du button
$ButtonLocation2.Size       = New-Object System.Drawing.Size(75,23)
#Text du button
$ButtonLocation2.Text       = "Location"
#Création du dialogue pour la séléction du chemin
$FolderPath = New-Object System.Windows.Forms.FolderBrowserDialog
### Event click ###
$ButtonLocation2.Add_Click({
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
$OKButton.Location     = New-Object System.Drawing.Size(10,180)
#Taille du button
$OKButton.Size         = New-Object System.Drawing.Size(75,23)
#Text du button
$OKButton.Text         = "Convertir"
### Event click ###
$OKButton.Add_Click({
# Module ImportExcel Check
    if (Get-Module -ListAvailable -Name ImportExcel) {
        $ModuleCheck = "true"
    } 
    else {
        $ModuleCheck = "false"
    }
#Output
$SelectedOutput = $ComboboxTypeOutput.SelectedItem # On stock l'option séléctionné pour le format de sortie dans une variable
$script:x += $ComboboxTypeOutput.SelectedItem # Pour qu'un seul item soit séléctionné
#Input
$SelectedInput = $ComboboxTypeInput.SelectedItem # On stock l'option séléctionné pour le format de sortie dans une variable
$script:x += $ComboboxTypeInput.SelectedItem # Pour qu'un seul item soit séléctionné
#CSV
#Debug verification si le chemin d'entrée est vide
    if($FilePath.FileName -eq "")
    {
        [System.Windows.Forms.MessageBox]::Show("Chemin d'entrée non défini",'Erreur','Ok','Error') # Message d'erreur
        $LabelInfo.ForeColor = "Red" # Changement de couleur pour le text du chemin d'entrée
        $LabelInfo.Text = "Chemin d'entrée non spécifié" # On remet le text car il disparait après avoir mis un chemin Null
    }
    else
    {
        $FileExtension = Get-Item $FilePath.FileName # on extrait l'information de l'extension du fichier et on la stock dans la variable $FileExtension
    }
#Debug verification si le chemin de sortie est vide
    if($FolderPath.SelectedPath -eq "")
    {
        [System.Windows.Forms.MessageBox]::Show("Chemin de sorti non défini",'Erreur','Ok','Error') # Message d'erreur
        $LabelInfo2.ForeColor = "red" # Changement de couleur pour le text du chemin de sortie
        $LabelInfo2.Text = "Chemin de sortie non spécifié" # On remet le text car il disparait après avoir mis un chemin Null
    }
#Debug verification si l'extension du fichier corrrespond au format d'entrée séléctionné
    if($FileExtension.Extension -like $SelectedInput)
        {
            $LabelFormatInput.ForeColor = "Green" # Changement de couleur pour le text du format d'entrée
            $OutputFileName = $TextBoxOutPutFileName.Text
            $Destionation = $FolderPath.SelectedPath # On stock le chemin séléctionné dans la variable $Destination
            if ($SelectedOutput -eq ".json" -And $SelectedInput -eq ".csv") # Si la sortie = ".y" et l'entrée = ".x" alors on convertit de la facon adéquate
            {
                import-csv -Delimiter "$Delimiter" $FilePath.FileName | ConvertTo-Json | Add-Content -Path "$Destionation\$OutputFileName.json"
            }
            elseif ($SelectedOutput -eq ".xml" -And $SelectedInput -eq ".csv") #Sinon la sortie = ".y" et l'entrée = ".x" alors on convertit de la facon adéquate
            {
                import-csv -Delimiter "$Delimiter" $FilePath.FileName | Export-Clixml "$Destionation\$OutputFileName.xml" 
            }
            elseif ($SelectedOutput -eq ".xls" -And $SelectedInput -eq ".csv") #Sinon la sortie = ".y" et l'entrée = ".x" alors on convertit de la facon adéquate
            {
                If ($ModuleCheck -eq "true")
                {
                import-csv -Delimiter "$Delimiter" $FilePath.FileName | Export-Excel "$Destionation\$OutputFileName.xlsx"
                }
                else
                {
                    ModuleMissing_Visible
                    [System.Windows.Forms.MessageBox]::Show("Module ImportExcel est manquant, cliquer sur Install",'Information','Ok','warning')
                }
            }

#json#

            elseif ($SelectedOutput -eq ".csv" -And $SelectedInput -eq ".json") #Sinon la sortie = ".y" et l'entrée = ".x" alors on convertit de la facon adéquate
            {
                Get-Content $FilePath.FileName | ConvertFrom-Json | ConvertTo-Csv -Delimiter "$Delimiter" | Out-File "$Destionation\$OutputFileName.csv" 
            }
            elseif ($SelectedOutput -eq ".xml" -And $SelectedInput -eq ".json") #Sinon la sortie = ".y" et l'entrée = ".x" alors on convertit de la facon adéquate
            {
                Get-Content $FilePath.FileName | ConvertFrom-Json | Export-Clixml "$Destionation\$OutputFileName.xml" 
            }
            elseif ($SelectedOutput -eq ".xls" -And $SelectedInput -eq ".json") #Sinon la sortie = ".y" et l'entrée = ".x" alors on convertit de la facon adéquate
            {
                [System.Windows.Forms.MessageBox]::Show("xls pas fait",'Information','Ok','warning')
            }

#xml#

            elseif ($SelectedOutput -eq ".csv" -And $SelectedInput -eq ".xml") #Sinon la sortie = ".y" et l'entrée = ".x" alors on convertit de la facon adéquate
            {
                Import-Clixml $FilePath.FileName | ConvertTo-Csv -Delimiter "$Delimiter" | Add-Content -Path "$Destionation\$OutputFileName.csv" 
            }
            elseif ($SelectedOutput -eq ".json" -And $SelectedInput -eq ".xml") #Sinon la sortie = ".y" et l'entrée = ".x" alors on convertit de la facon adéquate
            {
                Import-Clixml $FilePath.FileName | ConvertTo-Json | Out-File "$Destionation\$OutputFileName.json" 
            }
            elseif ($SelectedOutput -eq ".xls" -And $SelectedInput -eq ".xml") #Sinon la sortie = ".y" et l'entrée = ".x" alors on convertit de la facon adéquate
            {
                [System.Windows.Forms.MessageBox]::Show("xls pas fait",'Information','Ok','warning')
            }

#Error Input Output#
            elseif ($SelectedOutput -eq ".csv" -And $SelectedInput -eq ".csv") #Sinon la sortie = ".x" et entrée = ".x"
            {
                [System.Windows.Forms.MessageBox]::Show('Vous ne pouvez pas faire cela','Erreur','Ok','Error') # Message d'erreur
                
            }  
            elseif ($SelectedOutput -eq ".json" -And $SelectedInput -eq ".json") #Sinon la sortie = ".x" et entrée = ".x"
            {
                [System.Windows.Forms.MessageBox]::Show('Vous ne pouvez pas faire cela','Erreur','Ok','Error') # Message d'erreur
            }
            elseif ($SelectedOutput -eq ".xml" -And $SelectedInput -eq ".xml") #Sinon la sortie = ".x" et entrée = ".x"
            {
                [System.Windows.Forms.MessageBox]::Show('Vous ne pouvez pas faire cela','Erreur','Ok','Error') # Message d'erreur
            }  
            elseif ($SelectedOutput -eq ".xls" -And $SelectedInput -eq ".xls") #Sinon la sortie = ".x" et entrée = ".x"
            {
                [System.Windows.Forms.MessageBox]::Show('Vous ne pouvez pas faire cela','Erreur','Ok','Error') # Message d'erreur
            } 
        }
    else
    {
        if($FileExtension.Extension -NotLike $SelectedInput) # Si l'extension ne conrrespond pas au fichier séléctionné alors
        {
            $LabelFormatInput.ForeColor = "red" # Changement de couleur pour le text du format d'entrée
            [System.Windows.Forms.MessageBox]::Show("Format d'entrée non défini ou invalide",'Erreur','Ok','Error') # Message d'erreur
        }
    }   
})
###############################################################
#                            Combobox                         #
###############################################################
##########################
#     Combobox Input     #
##########################
#Création de la combobox pour le type d'entrée
$ComboboxTypeInput          = New-Object System.Windows.Forms.Combobox
#Location de la combobox
$ComboboxTypeInput.Location = New-Object System.Drawing.Size(10,40)
#Taille de la combobox
$ComboboxTypeInput.Size     = New-Object System.Drawing.Size(120,20)
#Taille de l'onglet
$ComboboxTypeInput.Height   = 70
#Ajout des item dans la combobox
$ComboboxTypeInput.Items.Add(".csv") #[void] $ComboboxTypeInput.Items.Add("csv")
$ComboboxTypeInput.Items.Add(".json")
$ComboboxTypeInput.Items.Add(".xml")
$ComboboxTypeInput.Items.Add(".xls")
#Item par défaut
$ComboboxTypeInput.SelectedIndex = 0
##########################
#     Combobox Output    #
##########################
#Création de la combobox pour le type de sortie
$ComboboxTypeOutput          = New-Object System.Windows.Forms.Combobox
#Location de la combobox
$ComboboxTypeOutput.Location = New-Object System.Drawing.Size(200,40)
#Taille de la combobox
$ComboboxTypeOutput.Size     = New-Object System.Drawing.Size(120,20)
#Taille de l'onglet
$ComboboxTypeOutput.Height   = 70
#Ajout des item dans la combobox
$ComboboxTypeOutput.Items.Add(".csv")
$ComboboxTypeOutput.Items.Add(".json")
$ComboboxTypeOutput.Items.Add(".xml")
$ComboboxTypeOutput.Items.Add(".xls")
#Item par défaut
$ComboboxTypeOutput.SelectedIndex = 1
##############################################################################################
#                                              Control ToolBox                               #
##############################################################################################
# Déclare les variables du ToolBox pour les afficher
$main_form.controls.AddRange(@(
#TextBox
$TextBoxOutPutFileName
#Label
$LabelOutputName
$LabelFormatInput
$LabelFormatOutput
$LabelInfo
$LabelInfo2
$LabelModuleCheck
#Button
$OKButton
$ButtonLocation
$ButtonLocation2
$ButtonInstallModule
#Combobox
$ComboboxTypeInput
$ComboboxTypeOutput
))
# Affiche/Cache les fenêtre
ModuleMissing_Invisible
$main_form.ShowDialog()
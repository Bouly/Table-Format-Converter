<#

.SYNOPSIS

Convertiseur de format de tableau.

.DESCRIPTION

Convertisseur de format de tableau en Powershell et Windows Form.
Converti plusieurs types de fichier avec une spécification du délimiter.

.LINK

https://github.com/Bouly/Table-Format-Converter

#>


##############################################################################################################################
#                                                       Chargement des classe                                                #
##############################################################################################################################
# Chargement des classe l'interface GUI
Add-Type -AssemblyName System.Windows.Forms
Add-Type -AssemblyName System.Drawing
##############################################################################################################################
#                                                              Function                                                      #
##############################################################################################################################

$CurrentPath = Get-Location # Stock le chemin courant dans une variable
$Delimiter = (Get-Culture).Textinfo.ListSeparator # Le délimiter = au délimiter de base du PC
$EncodingType = 'UTF8'

#Function d'affichage
function ModuleMissing_Visible { # Fonction qui rend visible les tools défini
    $LabelModuleCheck.Visible = $true
    $ButtonInstallModule.Visible = $true
    [System.Windows.Forms.MessageBox]::Show("Module ImportExcel est manquant, cliquer sur Install",'Information','Ok','warning') # Message informatif
}
function ModuleMissing_Invisible { # Fonction qui rend invisible les tools défini
    $LabelModuleCheck.Visible = $false
    $ButtonInstallModule.Visible = $false
}

function DefaultDelimiter{ # Fonction qui rend visible les tools défini
    $TextChoiceDelimiter.Visible = $false
    $LabelDelimiter.Visible = $false
    
}

function ChoiceDelimiter{ # Fonction qui rend visible les tools défini
    $TextChoiceDelimiter.Visible = $true
    $LabelDelimiter.Visible = $true
    }

function YouCant{
    [System.Windows.Forms.MessageBox]::Show('Vous ne pouvez pas faire cela','Erreur','Ok','Error') # Message d'erreur
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
# Centre la fênetre lors du lancement du script
$main_form.StartPosition= 'CenterScreen'
# Design de la fênetre
$main_form.FormBorderStyle = [System.Windows.Forms.FormBorderStyle]::FixedDialog
# Couleur du fond
$main_form.BackColor        = '38,36,49'
# Icon du GUI
$main_form.Icon             = [System.Drawing.Icon]::ExtractAssociatedIcon("$CurrentPath\logo.ico")
# Bloque la taille max et min
$main_form.minimumSize      = New-Object System.Drawing.Size(585,365)
$main_form.maximumSize      = New-Object System.Drawing.Size(585,365)
##############################################################################################
#                                              Toolbox                                       #
##############################################################################################

$RadioEncodageGroup = New-Object System.Windows.Forms.GroupBox
$RadioEncodageGroup.Location = '150,200'
$RadioEncodageGroup.size = '100,110'
$RadioEncodageGroup.text = "Encodage"
$RadioEncodageGroup.ForeColor = "white"
$RadioEncodageGroup.Font      = "Bahnschrift, 10"

$RadioDelimiterGroup = New-Object System.Windows.Forms.GroupBox
$RadioDelimiterGroup.Location = '300,200'
$RadioDelimiterGroup.size = '250,110'
$RadioDelimiterGroup.text = "Délimiter Par défault: " + '" ' + $Delimiter + ' "'
$RadioDelimiterGroup.ForeColor = "white"
$RadioDelimiterGroup.Font      = "Bahnschrift, 10"


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
# Design du champ de text
$TextBoxOutPutFileName.BorderStyle = [System.Windows.Forms.BorderStyle]::FixedSingle
##########################
#   TextBox Choice Deli  #
##########################
#Création du champ de text pour le nom du fichier
$TextChoiceDelimiter         = New-Object System.windows.Forms.TextBox
#Location du champ de text
$TextChoiceDelimiter.Location = New-Object System.Drawing.Size(170,55)
#Taille du champ de text
$TextChoiceDelimiter.Size     = New-Object System.Drawing.Size(15,20)
#L'entrée du champ de text
$TextChoiceDelimiter.Text     = ","
#Design du champ de text
$TextChoiceDelimiter.BorderStyle = [System.Windows.Forms.BorderStyle]::FixedSingle
#Limit de charactère dans le champ de text
$TextChoiceDelimiter.MaxLength = 1
#Cacher la TextBox
$TextChoiceDelimiter.Visible  = $false
###############################################################
#                            Label                            #
###############################################################
##########################
#   Label Delimiter      #
##########################
#Création du label pour le nom du fichier
$LabelDelimiter           = New-Object System.Windows.Forms.Label
#Location du label
$LabelDelimiter.Location  = New-Object System.Drawing.Size(10,60)
#Taille du label
$LabelDelimiter.Size      = New-Object System.Drawing.Size(200,20)
#Text du Label
$LabelDelimiter.Text      = "Entrée un délimiter valide"
#Police et taille du text du label
$LabelDelimiter.Font      = "Bahnschrift, 10"
#Couleur du Label
$LabelDelimiter.ForeColor = "243, 244, 247"
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
$LabelModuleCheck.ForeColor = "Red"
##########################
#   Label Output Name    #
##########################
#Création du label pour le nom du fichier
$LabelOutputName           = New-Object System.Windows.Forms.Label
#Location du label
$LabelOutputName.Location  = New-Object System.Drawing.Size(386,20)
#Taille du label
$LabelOutputName.Size      = New-Object System.Drawing.Size(167,20)
#Text du Label
$LabelOutputName.Text      = "Nom du fichier"
#Police et taille du text du label
$LabelOutputName.Font      = [System.Drawing.Font]::new("Verdana", 12)
#Couleur du label
$LabelOutputName.ForeColor = "243, 244, 247"
##########################
#   Label Input format   #
##########################
#Création du label pour le format d'entrée
$LabelFormatInput           = New-Object System.Windows.Forms.Label
#Location du label
$LabelFormatInput.Location  = New-Object System.Drawing.Size(6,20)
#Taille du label
$LabelFormatInput.Size      = New-Object System.Drawing.Size(150,20)
#Text du Label
$LabelFormatInput.Text      = "Format d'entrée"
#Police et taille du text du label
$LabelFormatInput.Font      = [System.Drawing.Font]::new("Verdana", 12)
#Couleur du label
$LabelFormatInput.ForeColor = "243, 244, 247"
##########################
#   Label Output format  #
##########################
#Création du label pour le format de sortie
$LabelFormatOutput          = New-Object System.Windows.Forms.Label
#Location du label
$LabelFormatOutput.Location = New-Object System.Drawing.Size(196,20)
#Taille du label
$LabelFormatOutput.Size     = New-Object System.Drawing.Size(150,20)
#Text du label
$LabelFormatOutput.Text     = "Format de sortie"
#Police et taille du label
$LabelFormatOutput.Font      = [System.Drawing.Font]::new("Verdana", 12)
#Couleur du text
$LabelFormatOutput.ForeColor= "243, 244, 247"
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
$LabelInfo.ForeColor        = "243, 244, 247"
#Plice et taille du label
$LabelInfo.Font      = "Bahnschrift, 10"
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
$LabelInfo2.ForeColor       = "243, 244, 247"
#Police et taille du label
$LabelInfo2.Font      = "Bahnschrift, 10"
##########################
#  Label Delimiter Alert #
##########################
#Création du label pour informé l'état du délimiter
$DelimiterAlert                 = New-Object System.Windows.Forms.Label
#Location du label
$DelimiterAlert.Location        = New-Object System.Drawing.Size(325, 180)
#Taille du label
$DelimiterAlert.Size            = New-Object System.Drawing.Size(300,20)
#Text du label
$DelimiterAlert.Text            = "Délimiter Incorrect"
#Couleur du text
$DelimiterAlert.ForeColor       = "red"
#Police et taille du label
$DelimiterAlert.Font      = "Cascadia Code, 14"
###############################################################
#                            RadioButton                      #
###############################################################
##########################
#RadioButton Default Deli#
##########################
#Création du Radio Button pour choisir le délimiter par défaut
$RadioButtonDefaultDelimiter       = New-Object System.Windows.Forms.RadioButton
#Location du champ de text
$RadioButtonDefaultDelimiter.Location = New-Object System.Drawing.Size(10,20)
#Taille du champ de text
$RadioButtonDefaultDelimiter.Size     = New-Object System.Drawing.Size(137,20)
#L'entrée du champ de text
$RadioButtonDefaultDelimiter.Text     = "Default Delimiter"
#Police et taille du radio button
$RadioButtonDefaultDelimiter.Font      = "Bahnschrift, 10"
#Design du radio button
$RadioButtonDefaultDelimiter.FlatStyle = [System.Windows.Forms.FlatStyle]::Flat
#Activer par défault
$RadioButtonDefaultDelimiter.Checked = $true
#Couleur du radio button
$RadioButtonDefaultDelimiter.ForeColor       = "243, 244, 247"
# Event Click
$RadioButtonDefaultDelimiter.Add_Click({ #Quand le button est cliqué
    DefaultDelimiter #Function d'affichage
})
##########################
#RadioButton Choice Deli #
##########################
#Création du Radio Button pour choisir le délimité à choix
$RadioButtonChoiceDelimiter          = New-Object System.Windows.Forms.RadioButton
#Location du champ de text
$RadioButtonChoiceDelimiter.Location = New-Object System.Drawing.Size(10,40)
#Taille du champ de text
$RadioButtonChoiceDelimiter.Size     = New-Object System.Drawing.Size(137,20)
#L'entrée du champ de text
$RadioButtonChoiceDelimiter.Text     = "Choice Delimiter"
#Design du radio button
$RadioButtonChoiceDelimiter.FlatStyle = [System.Windows.Forms.FlatStyle]::Flat
#Police et taille du text
$RadioButtonChoiceDelimiter.Font      = "Bahnschrift, 10"
#Couleur du text
$RadioButtonChoiceDelimiter.ForeColor= "243, 244, 247"
# Event Click
$RadioButtonChoiceDelimiter.Add_Click({ #Quand le button cliqué
        ChoiceDelimiter #Function d'affichage
})
###

##########################
#    RadioButton UTF8    #
##########################
#Création du Radio Button pour choisir le délimiter par défaut
$RadioButtonUTF8       = New-Object System.Windows.Forms.RadioButton
#Location du champ de text
$RadioButtonUTF8.Location = New-Object System.Drawing.Size(10,20)
#Taille du champ de text
$RadioButtonUTF8.Size     = New-Object System.Drawing.Size(60,20)
#L'entrée du champ de text
$RadioButtonUTF8.Text     = "UTF8"
#Police et taille du radio button
$RadioButtonUTF8.Font      = "Bahnschrift, 10"
#Design du radio button
$RadioButtonUTF8.FlatStyle = [System.Windows.Forms.FlatStyle]::Flat
#Activer par défault
$RadioButtonUTF8.Checked = $true
#Couleur du radio button
$RadioButtonUTF8.ForeColor       = "243, 244, 247"
# Event Click
$RadioButtonUTF8.Add_Click({ #Quand le button est cliqué
    #$EncodingType = 'utf8'
})
##########################
#  RadioButton UTF8BOM   #
##########################
#Création du Radio Button pour choisir le délimité à choix
$RadioButtonUTF8BOM          = New-Object System.Windows.Forms.RadioButton
#Location du champ de text
$RadioButtonUTF8BOM.Location = New-Object System.Drawing.Size(10,40)
#Taille du champ de text
$RadioButtonUTF8BOM.Size     = New-Object System.Drawing.Size(88,20)

$RadioButtonUTF8BOM.Gr

#L'entrée du champ de text
$RadioButtonUTF8BOM.Text     = "UTF8-BOM"
#Design du radio button
$RadioButtonUTF8BOM.FlatStyle = [System.Windows.Forms.FlatStyle]::Flat
#Police et taille du text
$RadioButtonUTF8BOM.Font      = "Bahnschrift, 10"
#Couleur du text
$RadioButtonUTF8BOM.ForeColor= "243, 244, 247"
# Event Click
$RadioButtonUTF8BOM.Add_Click({ #Quand le button cliqué
    #$EncodingType = 'utf8BOM'
})

##########################
#  RadioButton ANSI      #
##########################
#Création du Radio Button pour choisir le délimité à choix
$RadioButtonANSI          = New-Object System.Windows.Forms.RadioButton
#Location du champ de text
$RadioButtonANSI.Location = New-Object System.Drawing.Size(10,60)
#Taille du champ de text
$RadioButtonANSI.Size     = New-Object System.Drawing.Size(80,20)
#L'entrée du champ de text
$RadioButtonANSI.Text     = "ASCII"
#Design du radio button
$RadioButtonANSI.FlatStyle = [System.Windows.Forms.FlatStyle]::Flat
#Police et taille du text
$RadioButtonANSI.Font      = "Bahnschrift, 10"
#Couleur du text
$RadioButtonANSI.ForeColor= "243, 244, 247"
# Event Click
$RadioButtonANSI.Add_Click({ #Quand le button cliqué
    #$EncodingType = 'ANSI'
})

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
#Couleur du text
$ButtonInstallModule.BackColor = "153, 152, 246"
#Taille et police du text
$ButtonInstallModule.Font      = "Bahnschrift, 10"
#Design Style
$ButtonInstallModule.FlatStyle = [System.Windows.Forms.FlatStyle]::Flat
# Event click
$ButtonInstallModule.Add_Click({ #Quand le button cliqué
    Install-Module ImportExcel -AllowClobber -Force # Installation du module ImportExcel -AllowClobber(Persmission) -Force(Focer l'installation)
    [System.Windows.Forms.MessageBox]::Show("Le module ImportExcel a bien était Installé",'Module installé','Ok','Information') #Message informatif
    ModuleMissing_Invisible # Cacher la partie Installation du module
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
#Police et taille du text
$ButtonLocation.Font      = "Bahnschrift, 10"
#Couleur du text
$ButtonLocation.BackColor = "153, 152, 246"
#Deisgn du button
$ButtonLocation.FlatStyle = [System.Windows.Forms.FlatStyle]::Flat
#Couleur du text
$ButtonLocation.ForeColor = "243, 244, 247"
#Création du dialogue pour la séléction du chemin
$FilePath                   = New-Object System.Windows.Forms.OpenFileDialog
# Event click
$ButtonLocation.Add_Click({ #Quand le button cliqué
    $FilePath.ShowDialog() # Affiche la page de dialogue pour la séléction du chemin
    if ($FilePath.FileName -eq $FilePath.FileName) # Si le chemin = le chemin alors
    {
        $LabelInfo.Text = $FilePath.FileName # Le label d'information de l'état du chemin = Le chemin choisi
        $LabelInfo.ForeColor = "153, 152, 246" # Couleur du text du label d'information de l'état du chemin en "vert"
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
#Police et taille
$ButtonLocation2.Font      = "Bahnschrift, 10"
#Couleur
$ButtonLocation2.BackColor = "153, 152, 246"
#Couleur du text
$ButtonLocation2.ForeColor = "243, 244, 247"
#Design
$ButtonLocation2.FlatStyle = [System.Windows.Forms.FlatStyle]::Flat
#Création du dialogue pour la séléction du chemin
$FolderPath = New-Object System.Windows.Forms.FolderBrowserDialog
### Event click ###
$ButtonLocation2.Add_Click({
    $FolderPath.ShowDialog() # Affiche la page de dialogue pour la séléction du chemin
    if ($FolderPath.SelectedPath -eq $FolderPath.SelectedPath) # Si le chemin = le chemin alors
    {
        $LabelInfo2.Text = $FolderPath.SelectedPath # Le label d'information de l'état du chemin = Le chemin choisi
        $LabelInfo2.ForeColor = "153, 152, 246" # Couleur du text du label d'information de l'état du chemin en "vert"
    }   
})

$RadioDelimiterGroup.Controls.AddRange(@(
$RadioButtonDefaultDelimiter
$RadioButtonChoiceDelimiter
$TextChoiceDelimiter
$LabelDelimiter
))

$RadioEncodageGroup.Controls.AddRange(@(
$RadioButtonANSI
$RadioButtonUTF8
$RadioButtonUTF8BOM
))

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
#Police et taille du text
$OKButton.Font      = "Bahnschrift, 10"
#Couleur
$OKButton.BackColor = "153, 152, 246"
#Design
$OKButton.FlatStyle = [System.Windows.Forms.FlatStyle]::Flat
#Couleur du text
$OKButton.ForeColor = "243, 244, 247"
### Event click ###
$OKButton.Add_Click({
############################
# Module ImportExcel Check #
############################
if (Get-Module -ListAvailable -Name ImportExcel) 
{ # On va chercher "ImporterExcel" dans liste tout les modules installé
    $ModuleCheck = "true" # si il est présent alors "$ModuleCheck = true"
} 
else 
{ # Sinon
    $ModuleCheck = "false"
}

if ($RadioButtonUTF8.Checked -eq $true)
{
    $EncodingType = 'utf8'
}
elseif ($RadioButtonUTF8BOM.Checked -eq $true) 
{
    $EncodingType = 'utf8BOM'
}
else
{
    $EncodingType = 'ASCII'
}

#############
# Delimiter #
#############
    if ($RadioButtonDefaultDelimiter.Checked -eq $true) # Si la RadioButton est coché sur celui par défault alors
    {
        $Delimiter = (Get-Culture).Textinfo.ListSeparator # Le délimiter = au délimiter de base du PC
    }
    elseif ($RadioButtonChoiceDelimiter.Checked -eq $true) # Sinon Si la RadioButton est coché sur celui à choix alors
    {
        $Delimiter = $TextChoiceDelimiter.Text # Le délimiter = au text de la TextBox
    }
########
#Output#
########
$SelectedOutput = $ComboboxTypeOutput.SelectedItem # On stock l'option séléctionné pour le format de sortie dans une variable
$script:x += $ComboboxTypeOutput.SelectedItem # Pour qu'un seul item soit séléctionné
#######
#Input#
#######

$SelectedInput = $ComboboxTypeInput.SelectedItem # On stock l'option séléctionné pour le format de sortie dans une variable
$script:x += $ComboboxTypeInput.SelectedItem # Pour qu'un seul item soit séléctionné

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
            $LabelFormatInput.ForeColor = "153, 152, 246" # Changement de couleur pour le text du format d'entrée
            $OutputFileName = $TextBoxOutPutFileName.Text # On stock le nom entrée dans la text box dans $OutputFileName
            $Destionation = $FolderPath.SelectedPath # On stock le chemin séléctionné dans la variable $Destination

#####
#CSV#
#####
            if ($SelectedOutput -eq ".json" -And $SelectedInput -eq ".csv") # Si la sortie = ".y" et l'entrée = ".x" alors on convertit de la facon adéquate
            {
                #import-csv -Delimiter "$Delimiter" $FilePath.FileName | ConvertTo-Json | Add-Content -Path = "$Destionation\$OutputFileName.json"
                import-csv -Delimiter "$Delimiter" $FilePath.FileName | ConvertTo-Json | Out-File -Encoding $EncodingType -Path  "$Destionation\$OutputFileName.json"
            }
            elseif ($SelectedOutput -eq ".xml" -And $SelectedInput -eq ".csv") #Sinon la sortie = ".y" et l'entrée = ".x" alors on convertit de la facon adéquate
            {
                import-csv -Delimiter "$Delimiter" $FilePath.FileName | Export-Clixml -Encoding $EncodingType "$Destionation\$OutputFileName.xml" 
            }
            elseif ($SelectedOutput -eq ".xlsx" -And $SelectedInput -eq ".csv") #Sinon la sortie = ".y" et l'entrée = ".x" alors on convertit de la facon adéquate
            {
                If ($ModuleCheck -eq "true") # Si le module est présent alors 
                {
                import-csv -Delimiter "$Delimiter" $FilePath.FileName | Export-Excel -Encoding $EncodingType "$Destionation\$OutputFileName.xlsx"
                }
                elseif ($ModuleCheck -eq "false") # Sinon
                {
                    ModuleMissing_Visible #  On affiche la partie de l'installation du module
                }
            }
######
#json#
######
            elseif ($SelectedOutput -eq ".csv" -And $SelectedInput -eq ".json") #Sinon la sortie = ".y" et l'entrée = ".x" alors on convertit de la facon adéquate
            {
                Get-Content $FilePath.FileName | ConvertFrom-Json | ConvertTo-Csv -Delimiter "$Delimiter" | Out-File -Encoding $EncodingType -Path "$Destionation\$OutputFileName.csv" 
            }
            elseif ($SelectedOutput -eq ".xml" -And $SelectedInput -eq ".json") #Sinon la sortie = ".y" et l'entrée = ".x" alors on convertit de la facon adéquate
            {
                Get-Content $FilePath.FileName | ConvertFrom-Json | Export-Clixml -Delimiter "$Delimiter" -Encoding $EncodingType "$Destionation\$OutputFileName.xml" 
            }
            elseif ($SelectedOutput -eq ".xlsx" -And $SelectedInput -eq ".json") #Sinon la sortie = ".y" et l'entrée = ".x" alors on convertit de la facon adéquate
            {
                If ($ModuleCheck -eq "true") # Si le module est présent alors 
                {
                    Get-Content $FilePath.FileName | ConvertFrom-Json | Export-Excel -Delimiter "$Delimiter" -Encoding $EncodingType "$Destionation\$OutputFileName.xlsx"  
                }
                elseif ($ModuleCheck -eq "false") # Sinon
                {
                    ModuleMissing_Visible #  On affiche la partie de l'installation du module
                }
                
            }
#####
#xml#
#####
            elseif ($SelectedOutput -eq ".csv" -And $SelectedInput -eq ".xml") #Sinon la sortie = ".y" et l'entrée = ".x" alors on convertit de la facon adéquate
            {
                Import-Clixml $FilePath.FileName | ConvertTo-Csv -Delimiter "$Delimiter" | Add-Content -Encoding $EncodingType -Path "$Destionation\$OutputFileName.csv" 
            }
            elseif ($SelectedOutput -eq ".json" -And $SelectedInput -eq ".xml") #Sinon la sortie = ".y" et l'entrée = ".x" alors on convertit de la facon adéquate
            {
                Import-Clixml $FilePath.FileName | ConvertTo-Json | Out-File -Delimiter "$Delimiter" -Encoding $EncodingType "$Destionation\$OutputFileName.json" 
            }
            elseif ($SelectedOutput -eq ".xlsx" -And $SelectedInput -eq ".xml") #Sinon la sortie = ".y" et l'entrée = ".x" alors on convertit de la facon adéquate
            {
                If ($ModuleCheck -eq "true") # Si le module est présent alors 
                {
                Import-Clixml $FilePath.FileName | Export-Excel -Delimiter "$Delimiter" -Encoding $EncodingType "$Destionation\$OutputFileName.xlsx"
                }
                elseif ($ModuleCheck -eq "false") # Sinon
                {
                    ModuleMissing_Visible #  On affiche la partie de l'installation du module
                }
            }
######
#xlsx#
######
            elseif ($SelectedOutput -eq ".csv" -And $SelectedInput -eq ".xlsx") #Sinon la sortie = ".y" et l'entrée = ".x" alors on convertit de la facon adéquate
            {
                If ($ModuleCheck -eq "true") # Si le module est présent alors
                {
                Import-Excel $FilePath.FileName | ConvertTo-Csv -Delimiter "$Delimiter" | Add-Content -Encoding $EncodingType -Path "$Destionation\$OutputFileName.csv"
                }
                elseif ($ModuleCheck -eq "false") # Sinon
                {
                    ModuleMissing_Visible #  On affiche la partie de l'installation du module
                }
            }
            elseif ($SelectedOutput -eq ".json" -And $SelectedInput -eq ".xlsx") #Sinon la sortie = ".y" et l'entrée = ".x" alors on convertit de la facon adéquate
            {
                If ($ModuleCheck -eq "true") # Si le module est présent alors
                {
                    Import-Excel $FilePath.FileName | ConvertTo-Json | Out-File -Encoding $EncodingType "$Destionation\$OutputFileName.json"
                    #Import-Excel $FilePath.FileName | ConvertTo-Json | Out-File -Delimiter "$Delimiter" -Encoding $EncodingType "$Destionation\$OutputFileName.json"
                }
                elseif ($ModuleCheck -eq "false") # Sinon
                {
                    ModuleMissing_Visible #  On affiche la partie de l'installation du module
                } 
            }
            elseif ($SelectedOutput -eq ".xml" -And $SelectedInput -eq ".xlsx") #Sinon la sortie = ".y" et l'entrée = ".x" alors on convertit de la facon adéquate
            {
                If ($ModuleCheck -eq "true") # Si le module est présent alors 
                {
                Import-Excel $FilePath.FileName | Export-Clixml -Delimiter "$Delimiter" -Encoding $EncodingType "$Destionation\$OutputFileName.xml"
                }
                elseif ($ModuleCheck -eq "false") # Sinon
                {
                    ModuleMissing_Visible #  On affiche la partie de l'installation du module
                }
            }

#Error Input Output# A Compacté

            elseif ($SelectedOutput -eq ".csv" -And $SelectedInput -eq ".csv") #Sinon la sortie = ".x" et entrée = ".x"
            {
                YouCant # Message d'erreur
                $ComboboxTypeOutput.SelectedIndex = 1
                
            }
            elseif ($SelectedOutput -eq ".json" -And $SelectedInput -eq ".json") #Sinon la sortie = ".x" et entrée = ".x"
            {
                YouCant # Message d'erreur
            }
            elseif ($SelectedOutput -eq ".xml" -And $SelectedInput -eq ".xml") #Sinon la sortie = ".x" et entrée = ".x"
            {
                YouCant # Message d'erreur
            }
            elseif ($SelectedOutput -eq ".xlsx" -And $SelectedInput -eq ".xlsx") #Sinon la sortie = ".x" et entrée = ".x"
            {
                YouCant # Message d'erreur
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
# Design de la combobox
$ComboboxTypeInput.FlatStyle = [System.Windows.Forms.FlatStyle]::Flat
# Read Only la combobox
$ComboboxTypeInput.DropDownStyle = [System.Windows.Forms.ComboBoxStyle]::DropDownList
#Ajout des item dans la combobox
$ComboboxTypeInput.Items.Add(".csv") #[void] $ComboboxTypeInput.Items.Add("csv")
$ComboboxTypeInput.Items.Add(".json")
$ComboboxTypeInput.Items.Add(".xml")
$ComboboxTypeInput.Items.Add(".xlsx")
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
#Design Style
$ComboboxTypeOutput.FlatStyle = [System.Windows.Forms.FlatStyle]::Flat
# Read Only la combobox
$ComboboxTypeOutput.DropDownStyle = [System.Windows.Forms.ComboBoxStyle]::DropDownList
#Taille de l'onglet
$ComboboxTypeOutput.Height   = 70
#Ajout des item dans la combobox
$ComboboxTypeOutput.Items.Add(".csv")
$ComboboxTypeOutput.Items.Add(".json")
$ComboboxTypeOutput.Items.Add(".xml")
$ComboboxTypeOutput.Items.Add(".xlsx")
#Item par défaut
$ComboboxTypeOutput.SelectedIndex = 1
##############################################################################################
#                                              Control ToolBox                               #
##############################################################################################
# Déclare les variables du ToolBox pour les afficher

$main_form.controls.AddRange(@(
$RadioDelimiterGroup
$RadioEncodageGroup
#TextBox
$TextBoxOutPutFileName
#Label
$LabelOutputName
$LabelFormatInput
$LabelFormatOutput
$LabelInfo
$LabelInfo2
$LabelModuleCheck
$DelimiterAlert
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
$DelimiterAlert.Visible = $false
ModuleMissing_Invisible
$main_form.ShowDialog()
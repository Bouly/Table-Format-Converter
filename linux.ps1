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
#                                                              Function                                                      #
##############################################################################################################################

$CurrentPath = Get-Location # Stock le chemin courant dans une variable
$Delimiter = (Get-Culture).Textinfo.ListSeparator # Le délimiter = au délimiter de base du PC
$EncodingType = 'UTF8'
$FilePath = "/home/pc/Desktop/Table-Format-Converter/ConverterBaseFormat"
$Destionation = "/home/pc/Desktop"




$titre = "Converter Table"




function mainMenu {
    $mainMenu = 'X'
    while($mainMenu -ne ''){
        Clear-Host
        Write-Host "`n`t`t $titre`n"
        Write-Host -ForegroundColor Cyan "Menu Principale"
        Write-Host -ForegroundColor DarkCyan -NoNewline "`n["; Write-Host -NoNewline "1"; Write-Host -ForegroundColor DarkCyan -NoNewline "]"; `
            Write-Host -ForegroundColor DarkCyan " CSV"
        Write-Host -ForegroundColor DarkCyan -NoNewline "`n["; Write-Host -NoNewline "2"; Write-Host -ForegroundColor DarkCyan -NoNewline "]"; `
            Write-Host -ForegroundColor DarkCyan " JSON"
        Write-Host -ForegroundColor DarkCyan -NoNewline "`n["; Write-Host -NoNewline "3"; Write-Host -ForegroundColor DarkCyan -NoNewline "]"; `
            Write-Host -ForegroundColor DarkCyan " XML"
#        Write-Host -ForegroundColor DarkCyan -NoNewline "`n["; Write-Host -NoNewline "4"; Write-Host -ForegroundColor DarkCyan -NoNewline "]"; `
#           Write-Host -ForegroundColor DarkCyan " XLSX"
        $mainMenu = Read-Host "`nChoix (laisser vide pour quitter)"
        # Launch submenu1
        if($mainMenu -eq 1){
            subMenu1
        }
        # Launch submenu2
        if($mainMenu -eq 2){
            subMenu2
        }
        # Launch submenu3
        if($mainMenu -eq 3){
            subMenu3
        }
        # Launch submenu4
        if($mainMenu -eq 4){
            subMenu4
        }
    }
}

function subMenu1 {
    $subMenu1 = 'X'
    while($subMenu1 -ne ''){
        Clear-Host
        Write-Host "`n`t`t $titre`n"
        Write-Host -ForegroundColor Cyan "CSV TO"
        Write-Host -ForegroundColor DarkCyan -NoNewline "`n["; Write-Host -NoNewline "1"; Write-Host -ForegroundColor DarkCyan -NoNewline "]"; `
            Write-Host -ForegroundColor DarkCyan " JSON"
        Write-Host -ForegroundColor DarkCyan -NoNewline "`n["; Write-Host -NoNewline "2"; Write-Host -ForegroundColor DarkCyan -NoNewline "]"; `
            Write-Host -ForegroundColor DarkCyan " XML"
#        Write-Host -ForegroundColor DarkCyan -NoNewline "`n["; Write-Host -NoNewline "3"; Write-Host -ForegroundColor DarkCyan -NoNewline "]"; `
#            Write-Host -ForegroundColor DarkCyan " XLSX"
        $subMenu1 = Read-Host "`nChoix (laisser vide pour quitter)"
        $timeStamp = Get-Date -Uformat %m%d%y%H%M
        # Option 1 CSV TO JSON
        if($subMenu1 -eq 1){
            $FilePath = Read-Host -Prompt "Input File"
            $Destionation = Read-Host -Prompt "Output File"
            $OutputFileName = Read-Host -Prompt "Name File"
            import-csv -Delimiter "$Delimiter" "$FilePath" | ConvertTo-Json | Out-File -Encoding $EncodingType -Path  "$Destionation\$OutputFileName.json"
            # Pause and wait for input before going back to the menu
            Write-Host -ForegroundColor DarkCyan "`nCommande exécuter."
            Write-Host "`nAppuyez sur n'importe quelle touche pour revenir au menu précédent"
            [void][System.Console]::ReadKey($true)
        }
        # Option 2 CSV TO XML
        if($subMenu1 -eq 2){
            $FilePath = Read-Host -Prompt "Input File"
            $Destionation = Read-Host -Prompt "Output File"
            $OutputFileName = Read-Host -Prompt "Name File"
            import-csv -Delimiter "$Delimiter" $FilePath | Export-Clixml -Encoding $EncodingType "$Destionation\$OutputFileName.xml" 
            # Pause and wait for input before going back to the menu
            Write-Host -ForegroundColor DarkCyan "`nCommande exécuter."
            Write-Host "`nAppuyez sur n'importe quelle touche pour revenir au menu précédent"
            [void][System.Console]::ReadKey($true)
        }
        # Option 3 CSV TO XLSX
        if($subMenu1 -eq 3){
            $FilePath = Read-Host -Prompt "Input File"
            $Destionation = Read-Host -Prompt "Output File"
            $OutputFileName = Read-Host -Prompt "Name File"
            import-csv -Delimiter "$Delimiter" $FilePath | Export-Excel -Encoding $EncodingType "$Destionation\$OutputFileName.xlsx"
            # Pause and wait for input before going back to the menu
            Write-Host -ForegroundColor DarkCyan "`nCommande exécuter."
            Write-Host "`nAppuyez sur n'importe quelle touche pour revenir au menu précédent"
            [void][System.Console]::ReadKey($true)
        }
    }
}

function subMenu2 {
    $subMenu2 = 'X'
    while($subMenu2 -ne ''){
        Clear-Host
        Write-Host "`n`t`t $titre`n"
        Write-Host -ForegroundColor Cyan "JSON TO"
        Write-Host -ForegroundColor DarkCyan -NoNewline "`n["; Write-Host -NoNewline "1"; Write-Host -ForegroundColor DarkCyan -NoNewline "]"; `
            Write-Host -ForegroundColor DarkCyan " CSV"
        Write-Host -ForegroundColor DarkCyan -NoNewline "`n["; Write-Host -NoNewline "2"; Write-Host -ForegroundColor DarkCyan -NoNewline "]"; `
            Write-Host -ForegroundColor DarkCyan " XML"
#        Write-Host -ForegroundColor DarkCyan -NoNewline "`n["; Write-Host -NoNewline "3"; Write-Host -ForegroundColor DarkCyan -NoNewline "]"; `
#            Write-Host -ForegroundColor DarkCyan " XLSX"
        $subMenu2 = Read-Host "`nChoix (laisser vide pour quitter)"
        $timeStamp = Get-Date -Uformat %m%d%y%H%M
        # Option 1
        if($subMenu2 -eq 1){
        #json to csv
            $FilePath = Read-Host -Prompt "Input File"
            $Destionation = Read-Host -Prompt "Output File"
            $OutputFileName = Read-Host -Prompt "Name File"
            Get-Content $FilePath | ConvertFrom-Json | ConvertTo-Csv -Delimiter "$Delimiter" | Out-File -Encoding $EncodingType -Path "$Destionation\$OutputFileName.csv" 
            # Pause and wait for input before going back to the menu
            Write-Host -ForegroundColor DarkCyan "`nCommande exécuter."
            Write-Host "`nAppuyez sur n'importe quelle touche pour revenir au menu précédent"
            [void][System.Console]::ReadKey($true)
        }
        # Option 2
        if($subMenu2 -eq 2){
        #json to xml
            $FilePath = Read-Host -Prompt "Input File"
            $Destionation = Read-Host -Prompt "Output File"
            $OutputFileName = Read-Host -Prompt "Name File"
            Get-Content $FilePath | ConvertFrom-Json | Export-Clixml -Delimiter "$Delimiter" -Encoding $EncodingType "$Destionation\$OutputFileName.xml" 
            # Pause and wait for input before going back to the menu
            Write-Host -ForegroundColor DarkCyan "`nCommande exécuter."
            Write-Host "`nAppuyez sur n'importe quelle touche pour revenir au menu précédent"
            [void][System.Console]::ReadKey($true)
        }
        # Option 3
        if($subMenu2 -eq 3){
        #json to xlsx
            $FilePath = Read-Host -Prompt "Input File"
            $Destionation = Read-Host -Prompt "Output File"
            $OutputFileName = Read-Host -Prompt "Name File"
            Get-Content $FilePath | ConvertFrom-Json | Export-Excel -Delimiter "$Delimiter" -Encoding $EncodingType "$Destionation\$OutputFileName.xlsx"
            # Pause and wait for input before going back to the menu
            Write-Host -ForegroundColor DarkCyan "`nCommande exécuter."
            Write-Host "`nAppuyez sur n'importe quelle touche pour revenir au menu précédent"
            [void][System.Console]::ReadKey($true)
        }
    }
}

function subMenu3 {
    $subMenu3 = 'X'
    while($subMenu3 -ne ''){
        Clear-Host
        Write-Host "`n`t`t $titre`n"
        Write-Host -ForegroundColor Cyan "JSON TO"
        Write-Host -ForegroundColor DarkCyan -NoNewline "`n["; Write-Host -NoNewline "1"; Write-Host -ForegroundColor DarkCyan -NoNewline "]"; `
            Write-Host -ForegroundColor DarkCyan " CSV"
        Write-Host -ForegroundColor DarkCyan -NoNewline "`n["; Write-Host -NoNewline "2"; Write-Host -ForegroundColor DarkCyan -NoNewline "]"; `
            Write-Host -ForegroundColor DarkCyan " XML"
#        Write-Host -ForegroundColor DarkCyan -NoNewline "`n["; Write-Host -NoNewline "3"; Write-Host -ForegroundColor DarkCyan -NoNewline "]"; `
#            Write-Host -ForegroundColor DarkCyan " XLSX"
        $subMenu3 = Read-Host "`nChoix (laisser vide pour quitter)"
        $timeStamp = Get-Date -Uformat %m%d%y%H%M
        # Option 1
        if($subMenu3 -eq 1){
 #xml to csv
 $FilePath = Read-Host -Prompt "Input File"
 $Destionation = Read-Host -Prompt "Output File"
 $OutputFileName = Read-Host -Prompt "Name File"
     Import-Clixml $FilePath | ConvertTo-Csv -Delimiter "$Delimiter" | Add-Content -Encoding $EncodingType -Path "$Destionation\$OutputFileName.csv"  
            # Pause and wait for input before going back to the menu
            Write-Host -ForegroundColor DarkCyan "`nCommande exécuter."
            Write-Host "`nAppuyez sur n'importe quelle touche pour revenir au menu précédent"
            [void][System.Console]::ReadKey($true)
        }
        # Option 2
        if($subMenu3 -eq 2){
        #xml to json
            $FilePath = Read-Host -Prompt "Input File"
            $Destionation = Read-Host -Prompt "Output File"
            $OutputFileName = Read-Host -Prompt "Name File"
            Import-Clixml $FilePath | ConvertTo-Json | Out-File -Delimiter "$Delimiter" -Encoding $EncodingType "$Destionation\$OutputFileName.json" 
            # Pause and wait for input before going back to the menu
            Write-Host -ForegroundColor DarkCyan "`nCommande exécuter."
            Write-Host "`nAppuyez sur n'importe quelle touche pour revenir au menu précédent"
            [void][System.Console]::ReadKey($true)
        }
        # Option 3
        if($subMenu3 -eq 3){
        #xml to xlsx
            $FilePath = Read-Host -Prompt "Input File"
            $Destionation = Read-Host -Prompt "Output File"
            $OutputFileName = Read-Host -Prompt "Name File"
            Import-Clixml $FilePath | Export-Excel -Delimiter "$Delimiter" -Encoding $EncodingType "$Destionation\$OutputFileName.xlsx"
            # Pause and wait for input before going back to the menu
            Write-Host -ForegroundColor DarkCyan "`nCommande exécuter."
            Write-Host "`nAppuyez sur n'importe quelle touche pour revenir au menu précédent"
            [void][System.Console]::ReadKey($true)
        }
    }
}

mainMenu


            #xlsx to csv
#            $FilePath = Read-Host -Prompt "Input File"
#            $Destionation = Read-Host -Prompt "Output File"
#            $OutputFileName = Read-Host -Prompt "Name File"
#                Import-Excel $FilePath | ConvertTo-Csv -Delimiter "$Delimiter" | Add-Content -Encoding $EncodingType -Path "$Destionation\$OutputFileName.csv"
            #xlsx to json
#            $FilePath = Read-Host -Prompt "Input File"
#            $Destionation = Read-Host -Prompt "Output File"
#            $OutputFileName = Read-Host -Prompt "Name File"
#                Import-Excel $FilePath | ConvertTo-Json | Out-File -Encoding $EncodingType "$Destionation\$OutputFileName.json"
            #xlsx to xml
#            $FilePath = Read-Host -Prompt "Input File"
#            $Destionation = Read-Host -Prompt "Output File"
#            $OutputFileName = Read-Host -Prompt "Name File"
#                Import-Excel $FilePath| Export-Clixml -Delimiter "$Delimiter" -Encoding $EncodingType "$Destionation\$OutputFileName.xml"


<#
    .Synopsis
    Quitter la session en cours (local/distant)

    .Description
    Quitte la session PowerShell en cours (locale ou distante) en sauvegardant l'historique de la console
    ainsi que l'historique PSReadLine dans les documents de l'utilisateur
    Documents -> logs -> cmd_history

    .Example
    Vma-Exit
#>
function Vma-Exit {
    $date = Get-Date
    $dateStr = $date.toString("yyyyMMdd")

    $logPath = $Env:USERPROFILE + "\Documents\logs\cmd_history\" + $dateStr + ".txt"
    $detailLogPath = $Env:USERPROFILE + "\Documents\logs\cmd_history\" + $dateStr + " detail.txt"
    $logDirPath = $Env:USERPROFILE + "\Documents\logs\cmd_history"

    New-Item -ItemType Directory -Path $logDirPath -Force

    Get-History >> $logPath
    Get-History | Format-List -Property * >> $detailLogPath

    try { Copy-Item -Path (Get-PSReadLineOption).HistorySavePath -Destination $logDirPath -ErrorAction "Stop" } 
    catch { Write-Host "Could not copy PSReadLine History :" $_.FullyQualifiedErrorId }

    Exit
}





<#
    .Synopsis
    Générateur de mot de passe

    .Description
    Génère un mot de passe de longueur <= 8 réglable contenant majuscules minuscules chiffres caractères spéciaux
    Peut accepter une chaine de caractères servant de base (nom de famille)

    .Parameter Base
    (Chaîne de caractères) Sert de base au mot de passe - optionnel
    La Base sera tronquée ou complétée aléatoirement en fonction de la longueur
    Si non valide/renseigné, une chaîne aléatoire est utilisée

    .Parameter Length
    (Entier >= 8) Longueur du mot de passe généré - optionnel
    Si non valide/renseigné, une longueur de 8 est utilisée

    .Example
    # Arguments par défaut
    Vma-Generate-Password                                   #Lrib&q21

    # Base renseignée, longueur par défaut
    Vma-Generate-Password -Base "Pinocchio"                 #pIno!c45

    # Base et longueur renseignée
    Vma-Generate-Password -Base "Pinocchio" -Length 10      #piN*Occh96
#>
function Vma-Generate-Password {
    param ([String]$Base="", [Int]$Length=8)

    if (($Length.GetType().Name -ne "Int32") -or ($Length -le 8)) { $Length = 8 }

    $specialChars = [int](($Length*14)/100)
    $digits = [int](($Length*25)/100)
    $chars = $Length - $specialChars - $digits
    $caps = [int](($chars*25)/100)

    if ($Base.GetType().Name -ne "String") { $Base = -join ((65..90) | Get-Random -Count $chars | % {[char]$_}) }

    $Base = $Base.Normalize("FormD") -replace '\p{M}',''
    $Base = $Base -replace '\W','' -replace '\s','' -replace '\d',''

    if ($Base.Length -lt $chars) { $Base = $Base + -join ((65..90) | Get-Random -Count ($chars-$Base.Length) | % {[char]$_}) }
    if ($Base.Length -gt $chars) { $Base = $Base.Substring(0,$chars)}

    $Base = $Base.ToLower()
    
    $draw = (0..($chars-1)) | Get-Random -Shuffle
    for (($i=0); $i -lt $caps; $i++) {    
        $split = $Base -Csplit $Base[$draw[$i]],2
        $Base = $split[0] + ([String]$Base[$draw[$i]]).ToUpper() + $split[1]
    }

    $toDraw = '%','!','*','?','-','_','&','='
    for (($i=0); $i -lt $specialChars; $i++) {
        $draw = (0..($Base.Length)) | Get-Random
        $Base = $Base.Insert($draw, ($toDraw | Get-Random))
    }

    $toReturn = $Base + -join((0..9) | Get-Random -Count $digits)
    Write-Output $toReturn
}





<#
    .Synopsis
    Recherche d'utilisateur AD

    .Description
    Recherche d'utilisateur AD par nom, prénom et id SAM combinés
    Chaque critère est facultatif, au moins un doit être renseigné
    Si la recherche ne retourne qu'un utilisateur, il est exporté dans la variable $user (globale)
    Si la recherche retourne plus d'un utilisateur, la fonction retourne une table des résultats (nom, prénom, id SAM)

    .Parameter Nom
    Nom de famille : chaîne de caractères, optionnel, longueur > 2

    .Parameter Prenom
    Prénom : chaîne de caractères, optionnel, longueur > 2

    .Parameter Sam
    Id SAM : chaîne de caractères, optionnel, longueur >2

    .Parameter Exact
    (Switch) Force une correspondance exacte sur le SamAccountName

    .Example
    Vma-ADUser-Search -Prenom "Pilette"
    Vma-ADUser-Search -Nom "Cadeau"
    Vma-ADUser-Search -Sam "capi"
    Vma-ADUser-Search -Nom "Dupont" -Prenom "Clovis" -Sam "ducl"
    Vma-ADUser-Search -Nom "Placo" -Sam "plqu"
    Vma-ADUser-Search -Sam "plqu01" -Exact
#>
function Vma-ADUser-Search {
    param ([String]$Nom="", [String]$Prenom="", [String]$Sam="", [Switch]$Exact)

    $NomTest = ($Nom.GetType().Name -ne "String") -or ($Nom.Length -le 2)
    $PrenomTest = ($Prenom.GetType().Name -ne "String") -or ($Prenom.Length -le 2)
    $SamTest = ($Sam.GetType().Name -ne "String") -or ($Sam.Length -le 2)

    if ($NomTest -and $PrenomTest -and $SamTest) { return "Input problem" }

    $filter = ""

    if (!$NomTest -and !$PrenomTest) { $filter += "Name -like '*$Nom $Prenom*'" }
    elseif (!$NomTest) { $filter += "Name -like '*$Nom*'" }
    elseif (!$PrenomTest) { $filter += "Name -like '*$Prenom*'" }

    if ($Exact) { $samFilter = "-eq '$Sam'"} else { $samFilter = "-like '*$Sam*'" }

    if ((!$NomTest -or !$PrenomTest) -and (!$SamTest)) { $filter += " -and SamAccountName $samFilter" }
    elseif (!$SamTest) { $filter += "SamAccountName $samFilter" }

    $users = Get-ADUser -Filter $filter -Properties Name,SamAccountName

    if ($users.Length -eq 0) { return "No results" }
    
    elseif ($users.Length -gt 0) {
        Return $($users | Format-Table Name,SamAccountName -A)
    }
    
    elseif ($users.GetType().Name -eq "ADUser") {
        $Global:user = Get-ADUser -Identity $users.SamAccountName -Properties *
        Return "`$user : $($user.Name) - $($user.employeeNumber) - $($user.SamAccountName)"
    }
    
    else {
        Return "Results error"
    }
}





<#
    .Synopsis
    Actualisation de la variable globale $user

    .Description
    Si la variable globale $user est disponible et du type nommé ADUser (comme produit par la fonction Vma-ADUser-Search)
    Cette variable est actualisée
    Pour vérification/rappel ou actualisation de certain attributs pendant une série d'opérations

    .Example
    Vma-ADUser-Reload
#>
function Vma-ADUser-Reload {
    Try { $userType = $user.GetType().Name }
    Catch { Return "`$user/Type Error" }

    if ($userType -eq "ADUser") {
        $Global:user = Get-ADUser -Identity $user.SamAccountName -Properties *
        Return "`$user : $($user.Name) - $($user.employeeNumber) - $($user.SamAccountName)"
    }

    Return "Type Error"
}





<#
    .Synopsis
    Recherche de logs utilisateur

    .Description
    Recherche d'informations sur un utilisateur dans une liste de logs
    Inclut un système de logs datés pour chercher un log datant de la création de l'utilisateur par exemple

    .Parameter User
    (Utilisateur AD) à rechercher

    .Parameter LogPathTable
    (Hashtable) chemins de logs à rechercher 

    - Clés : type "N(D)" avec N entier positif et "D" à ajouter si le chemin de log contient une date
    Exemple de clés : 0, 1D, 2, 3D...

    - Valeurs : type chemin de fichier "C:\folder\mylog.txt"
    Pour les logs datés, utiliser le séparateur (défaut %) ainsi que les spécificateurs de format
    https://learn.microsoft.com/en-us/dotnet/standard/base-types/custom-date-and-time-format-strings
    "C:\folder2\mylog-date-%dd%-%mm%-%YYYY%.txt"

    .Parameter UserProperty
    (String) attribut de l'utilisateur AD à utiliser pour la recherche dans chaque log (defaut SamAccountName)

    .Parameter UserDate
    (String) attribut de l'utilisateur AD à utiliser pour les logs datés (defaut whenCreated)

    .Parameter DateSep
    (String) séparateur de date pour les logs datés (defaut %)

    .Example
    $myLogs = @{
        "0"="C:\folder\mylog.txt";
        "1D"="C:\folder2\mylog-date-%dd%-%mm%-%YYYY%.txt"
    }

    Vma-ADUser-Logs -LogPathTable $myLogs -User $myUser
#>
function Vma-ADUser-Logs {
    param(
        [Microsoft.ActiveDirectory.Management.ADAccount]$User,
        [Hashtable]$LogPathTable,
        [String]$UserProperty="SamAccountName",
        [String]$UserDate="whenCreated",
        [String]$DateSep="%"
    )

    if ($User.GetType().Name -ne "ADUser") { Return "Invalid user input" }
    if (($LogPathTable.GetType().Name -ne "Hashtable") -or ($LogPathTable.Count -lt 1)) { Return "Invalid log paths input" }
    if ($UserProperty.GetType().Name -ne "String") { Return "Invalid user property input" }
    if ($UserDate.GetType().Name -ne "String") { Return "Invalid user date input" }
    if ($($user.$UserDate).GetType().Name -ne "DateTime") { Return "Invalid user date input" }
    if ($DateSep.GetType().Name -ne "String") { Return "Invalid date sep input" }
    if ($DateSep.Length -lt 1) { Return "Invalid date sep input" }

    [regex]$keyRegex = '\d+D?'
    [regex]$dateKeyRegex = '\d+D'
    [regex]$pathDateRegex = "($DateSep\w+$DateSep)"

    $toReturn = ""

    ForEach ($h in $LogPathTable.GetEnumerator()) {
        $key = $($h.Name)
        $path = $($h.Value)

        if (!$keyRegex.Match($key).Success) {
            Write-Host $key ": Invalid key format"
            Continue 
        }    
        if ($keyRegex.Match($key).Success -and $dateKeyRegex.Match($key).Success) {
            $dateMatches = @()

            $path | 
                Select-String -Pattern $pathDateRegex -AllMatches |
                ForEach-Object { $dateMatches += $_.Matches.Value }
            
            for ($i = 0; $i -lt $dateMatches.Count; $i++) { 
                $path = $path.Replace($dateMatches[$i], $user.$UserDate.ToString($dateMatches[$i].Replace("%", "")))
            }
        }

        if (Test-Path $path) {
            $logSearch = Get-Content $path | Select-String -Pattern $user.$UserProperty
            if (($logSearch -ne $null) -and ($logSearch[0] -ne $null)) { $toReturn += "`n$path`n$($logSearch -join "`n")`n" }
        }
    }

    Write-Host $toReturn
}


#-----------------------
# FONCTIONS
#-----------------------

# Affiche le menu
function Menu(){
    Write-Host "`nVeuillez choisir une des options suivantes :`n"
    Write-Host "    1- Consulter le log"
    Write-Host "    2- Tester une connexion"
    Write-Host "    3- Réinitialiser des mots de passe"
    Write-Host "    4- Exporter des Unités Organisationnelles"
    Write-Host "    5- Exporter des groupes"
    Write-Host "    6- Exporter des utilisateurs"
    Write-Host "    7- Connexion VPN`n"
    Write-Host "    Q- Quitter le script`n"
    do{
        $Option = Read-Host "Votre choix "
        $Options = @(1, 2, 3, 4, 5, 6, 7, "Q")
        $ChoixCorrect = $true
        if(!($Option -in $Options)){
            Write-Host "Option invalide. Veuillez réessayez"
            $ChoixCorrect = $false
        }
    }While( !$ChoixCorrect )
    Return $Option
}
# Ajoute une information au log
function Journaliser($message, $resultat){
    $date = Get-Date -Format dd-MM-yyyy-HHmm
    Add-Content -Value "$date $message $resultat" -Path ".\journal.log"
}
# Afficher le log
function Afficher-Journal(){
    $journal = "journal.log"
    Get-Content ".\$journal" -ErrorAction Ignore
    if(!$?){
        Write-Host "Impossible d'ouvrir $journal"
    }
}
# Test la connexion à un ordinateur distant
function Tester-Connection($Ordinateur, $Port){
    if($Port -eq 0){
        $Test  = Test-NetConnection -ComputerName $Ordinateur
        return $Test.PingSucceeded
    }else{
        $Test = Test-NetConnection -ComputerName  $Ordinateur -Port $Port
        return $Test.TcpTestSucceeded
    }
}
# Lire le type d'objet : Utilisateur, Groupe ou Unité Organisationnelle
function Lire-TypeObjetAD(){
    $Types = "U", "G", "O"
    do{
        Write-Host " Type objet :" 
        Write-Host " U. Utilisateur" 
        Write-Host " G. Groupe" 
        Write-Host " O. Unité organisationnelle" 
        $Type = Read-Host "`n Votre choix"
        
    }while(!($Type -in  $Types))
    return $Type
}
# Réinitialise les mots de passe d’un utilisateur
function Reinitialiser-MotDePasse($Utilisateurs, $Password){
    foreach($Utilisateur in $Utilisateurs){
        Set-ADAccountPassword -Identity $Utilisateur.SamAccountName -Reset -NewPassword (ConvertTo-SecureString -AsPlainText $Password -Force) -Verbose
        Set-ADUser -Identity $Utilisateur.SamAccountName -ChangePasswordAtLogon $true
    }
}
# Réinitialise les mots de passe d’un utilisateur, d’un groupe ou d’une unité organisationnelle
function Recuperer-Utilisateurs($Type, $Nom){
    $Utilisateurs = @()
    Switch($Type){
        "U"{
            $Utilisateurs = Get-ADUser -Filter {SamAccountName -eq $Nom} | Select-Object SamAccountName
        }
        "G"{
            $Utilisateurs = Get-ADGroupMember -Identity $Nom | Select-Object SamAccountName
        }
        "O"{
            $AD = (Get-ADDomain | Select-Object DistinguishedName).DistinguishedName
            $Utilisateurs = Get-ADUser -Filter * -SearchBase "OU=$Nom,$AD" | Select-Object SamAccountName -ErrorAction Ignore
        }
    }
    return $Utilisateurs 
}
# Exporter les membres d'un groupe
function Exporter-Utilisateurs($Utilisateurs, $Nom){
    $CSV = @()
    foreach($Utilisateur in $Utilisateurs){
        $SamAccountName = $Utilisateur.SamAccountName
        $ADUser = Get-ADUser -Filter {SamAccountName -eq $SamAccountName}
        $CSV += $ADUser
    }
    $Date = Get-Date -Format ddMMyyyyHHmm
    $MachineDistante = (MachineDistante)
    $Fichier = "\\$MachineDistante\Share\export-$Nom-$Date.csv"
    $CSV | Export-Csv -Path $Fichier -NoTypeInformation -Verbose
    if($?){
        return $Fichier
    }else{
        return $false
    }

}
function MachineDistante{
    return "100.100.100.2"
}
# Permet la saisie du credential de la machine distante
function GetCredential{
    $MachineDistante = (MachineDistante)
    Write-Host "Connexion a $MachineDistante :"
    $Username = Read-Host "Utilisateur"
    $Password = Read-Host "Mot de passe" -MaskInput

    $User = "$MachineDistante\$Username"
    $PWord = ConvertTo-SecureString -String $Password -AsPlainText -Force
    $Credential = New-Object -TypeName System.Management.Automation.PSCredential -ArgumentList $User, $PWord -ErrorAction Ignore
    return $Credential
}  
# Importer les utilisateurs à partir d'un fichier CSV et les enregistrer sur la machine distante
function Importer-Utilisateurs($Nom, $Type, $Fichier, $Credential){
    # Connexion à la machine distante
    $Arguments = @{Nom=$Nom; Fichier=$Fichier; Type=$Type}
    $SessionOption = New-PSSessionOption -ApplicationArguments  $Arguments
    $Session = New-PSSession -ComputerName (MachineDistante) -SessionOption $SessionOption -Credential $Credential -ErrorAction Ignore
    # Session distante
    if($?){
        Invoke-Command -Session $Session -ScriptBlock {
            $Arguments = $PSSenderInfo.ApplicationArguments
            $Nom=$Arguments.Nom
            $Fichier=$Arguments.Fichier
            $Type=$Arguments.Type
            $AD = (Get-ADDomain | Select-Object DistinguishedName).DistinguishedName
            $DNS = (Get-ADDomain | Select-Object DNSRoot).DNSRoot
            #Creation groupe/OU
            if($Type -eq "O"){
                Get-ADOrganizationalUnit -Identity "OU=$Nom,$AD" -ErrorAction Ignore
                if(!($?)){
                    New-ADOrganizationalUnit -Name $Nom -Path $AD -ProtectedFromAccidentalDeletion $False
                }
            }elseif($Type -eq "G"){
                if(!(Get-ADGroup -Filter {Name -eq $Nom})){
                    New-ADGroup -Name $Nom -GroupCategory Security -GroupScope Global -Verbose
                }
            }
            # Importation des utilisateurs
            $CSV = Import-CSV -Path $Fichier
            $MotDePasse = "Passw0rd"
            foreach($Ligne in $CSV){
                $SamAccountName = $Ligne.SamAccountName
                $UserPrincipalName = "$SamAccountName@$DNS"
                if(!(Get-ADUser -Filter {SamAccountName -eq $SamAccountName})){
                    New-ADUser -Name $Ligne.Name -GivenName $Ligne.GivenName -Surname $Ligne.Surname -SamAccountName $Ligne.SamAccountName -UserPrincipalName $UserPrincipalName -Enabled $true -AccountPassword (ConvertTo-SecureString $MotDePasse -AsPlainText -Force) -Verbose
                    # Ajout de l'utilisateur au groupe/OU
                    $ADUser = Get-ADUser -Filter {SamAccountName -eq $SamAccountName}
                    if($Type -eq "O"){
                        Move-ADObject -Identity $ADUser.DistinguishedName -TargetPath "OU=$Nom,$AD"
                    }elseif($Type -eq "G"){
                        Add-ADGroupMember -Identity $Nom -Members $SamAccountName
                    }
                }else{
                    Write-Host "L'utilisateur $SamAccountName existe déjà"
                }
            }
        }
        Exit-PSSession
    }else{
        Write-Host "Erreur d'authentification"
    } 
}
function Ajouter-UtilisateurGroupe ($NomUtilisateur, $Groupe){
    $UtilisateurExiste = $true
    if(!(Get-ADUser -Filter {SamAccountName -eq $NomUtilisateur})){
        Write-Host "L'utilisateur n'existe pas"
        $UtilisateurExiste = $false
    }
    $GroupeExiste = $true
    if(!(Get-ADGroup -Filter {Name -eq $Groupe})){
        Write-Host "Le groupe n'existe pas"
        $GroupeExiste = $false
    }
    if($UtilisateurExiste -and $GroupeExiste){
        Add-ADGroupMember -Identity $Groupe -Members $NomUtilisateur -Verbose
    }
}


#-----------------------
# PARTIE PRINCIPALE
#-----------------------
do{
    $Choix = Menu
    Switch($Choix){
        1
        {
            Afficher-Journal
        }
        2
        {
            $Ordinateur= Read-Host "Ordinateur"
            $Port = Read-Host "Port (Entrer 0 si aucun port)"
            $Resultat = Tester-Connection $Ordinateur $Port
            Journaliser "[Standard | Test connexion] $Ordinateur $Port $Resultat" 
        }
        3
        {
            $TypeObjet = Lire-TypeObjetAD
            $NomObjet = Read-Host " Nom objet"
            $Password = Read-Host " Mot de passe" -MaskInput
            $Utilisateurs = Recuperer-Utilisateurs -Type $TypeObjet -Nom $NomObjet
            Reinitialiser-MotDePasse $Utilisateurs $Password
            Journaliser "[Standard | Reset mot de passe] $TypeObjet $NomObjet"
        }
        4
        {
            $Credential = GetCredential
            $TypeObjet = "O"
            $ObjetsString = Read-Host "Entrer les OU séparées par des espaces"
            $Objets = $ObjetsString -Split ' '
            foreach($NomObjet in $Objets){
                $Utilisateurs = Recuperer-Utilisateurs -Type $TypeObjet -Nom $NomObjet
                $Fichier = Exporter-Utilisateurs -Utilisateurs $Utilisateurs -Nom $NomObjet
                Importer-Utilisateurs -Nom $NomObjet -Type $TypeObjet -Fichier $Fichier -Credential $Credential
            }
            Journaliser "[Standard | Exportation OU] $TypeObjet $NomObjet"
        }
        5
        {
            $Credential = GetCredential
            $TypeObjet = "G"
            $ObjetsString = Read-Host "Entrer les groupes séparés par des espaces"
            $Objets = $ObjetsString -Split ' '
            foreach($NomObjet in $Objets){
                $Utilisateurs = Recuperer-Utilisateurs -Type $TypeObjet -Nom $NomObjet
                $Fichier = Exporter-Utilisateurs -Utilisateurs $Utilisateurs -Nom $NomObjet
                Importer-Utilisateurs -Nom $NomObjet -Type $TypeObjet -Fichier $Fichier -Credential $Credential
            }
            Journaliser "[Standard | Exportation Groupe(s)] $TypeObjet $NomObjet"
        }
        6
        {
            $Credential = GetCredential
            $TypeObjet = "U"
            $ObjetsString = Read-Host "Entrer les utilisateurs séparés par des espaces"
            $Objets = $ObjetsString -Split ' '
            foreach($NomObjet in $Objets){
                $Utilisateurs = Recuperer-Utilisateurs -Type $TypeObjet -Nom $NomObjet
                $Fichier = Exporter-Utilisateurs -Utilisateurs $Utilisateurs -Nom $NomObjet
                Importer-Utilisateurs -Nom $NomObjet -Type $TypeObjet -Fichier $Fichier -Credential $Credential
            }
            Journaliser "[Standard | Exportation Utilisateurs] $TypeObjet $NomObjet"
        }
        7
        {
     
        }
    }
}While(!($Choix -eq "Q"))
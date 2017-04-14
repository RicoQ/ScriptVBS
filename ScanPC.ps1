#Script Microsoft PowerShell

Function ScanPC { 
	<#
     .SYNOPSIS
Effectue un Ping sur tous les PC dans l'AD, et donne le nom de l'utilisateur connecter.
         
     .DESCRIPTION
La fonction ScanPC Effectue un Ping sur tous les ordinateur distant compris dans l'AD, et récupéré le nom des utilisateur connecter.
Il est à noter que l'ordinateur distant doit être configuré pour accepter les requêtes.
         
     .PARAMETER PC
Pour effectuer un ping sur un ou plusieurs ordinateurs distants spécifiques. Si ce paramètre n'est pas spécifié, alors la fonction ce lancera normalement.
     
     .PARAMETER Logs
         Ce paramètre créer 5 Log différent.
  - Log 1 = "List.txt" crée une liste de tous les ordinateur Scaner. 
  - Log 2 = "Résultat_Scan" crée la liste des noms de l'ordinateur et le nom
    de l'utilisateur connecter.
  - Log 3 = "PC_Sans_Reponse.txt" crée la liste des ordinateur sans réponse
    (uniquement). 
  - Log 4 = "Resultat_Complet.txt" combine les log 2, 3, et 5 dans un seul
    fichier. 
  - Log 5 = "Erreur_Scan.txt" Cree la liste des Erreur. Comprend le nom de
    L’ordinateur et l'intituler de l'erreur.
         
     .EXAMPLE
        PS > ScanPC -PC PC02*
        
     PC020 ; Pas de réponse
     PC021 ; Erreur sur la machine =>  : Accès refusé. (Exception de HRESULT 
             : 0x80070005 (E_ACCESSDENIED))
     PC022 ; Pas de réponse
     PC023 ; LeDomaine/Marie Smith
     PC024 ; Erreur sur la machine =>  :Exception lors de l'appel de « Send » avec
             « 1 » argument(s) : « Une exception s'est produite lors d'une demande
             PING. »
     PC025 ; Pas de réponse
     PC026 ; Erreur sur la machine =>  : Accès refusé. (Exception de HRESULT 
             : 0x80070005 (E_ACCESSDENIED))
     PC027 ; LeDomaine/Joe Black
     PC028 ; Pas de réponse
     PC029 ; LeDomaine/Mickael LeBlond
   
    .EXAMPLE
        PS > ScanPC -Logs
        
     Mode                LastWriteTime     Length Name                                                                                                                                   
     ----                -------------     ------ ----                                                                                                                                   
     -a---        12/08/2015     10:27          0 List.txt                                                                                                                            
     -a---        12/08/2015     10:27          0 Resultat_Scan.txt                                                                                                                     
     -a---        12/08/2015     10:27          0 PC_Sans_Reponse.txt                                                                                                                      
     -a---        12/08/2015     10:27          0 Resultat_Complet.txt                                                                                                                 
     -a---        12/08/2015     10:27          0 Erreur_Scan.txt

     PC001 ; Pas de réponse
     PC002 ; LeDomaine/Joe Black
     PC003 ; LeDomaine/Marie Smith
     PC004 ; Erreur sur la machine => : Exception lors de l'appel de « Send » avec
             « 1 » argument(s) : « Une exception s'est produite lors d'une demande
             PING. »
     PC005 ; Pas de réponse
     PC006 ; Erreur sur la machine => : Accès refusé. (Exception de HRESULT
      : 0x80070005 (E_ACCESSDENIED))
     PC007 ; Erreur sur la machine => : Accès refusé. (Exception de HRESULT 
      : 0x80070005 (E_ACCESSDENIED))
     PC008 ; Pas de réponse
     PC009 ; Pas de réponse
     PC010 ; LeDomaine/Mickale LeBlond
     
     .NOTES
 ScanPC créer 1 Log quelque soit le paramètre choisi dans le répertoire  
       ScanPC.
     
         Author:   Eric Quercia
		Assister Par, Mr Garçon Gislain (Tuteur de stage)
         Version:  4.0
	#>

    Param([String]$PC = '',[Switch]$Logs)
    $Path = [Environment]::GetFolderPath("Desktop")
    cd $Path
	    
    $ScanFolder = ".\ScanPC"
    $testFolder = Test-Path $ScanFolder
	    If ($testFolder -ne "True") {New-Item -ItemType directory -Path $ScanFolder -Force}
    $fichier4 = ".\ScanPC\Resultat_Complet.txt"
    $test4 = Test-Path $fichier4
        If ($test4 -ne "True") {New-Item -ItemType file -Path $fichier4 -Force}

    If($Logs) {     
       $fichier1 = ".\ScanPC\List.txt"
       $test1 = Test-Path $fichier1
           If ($test1 -ne "True") {New-Item -ItemType file -Path $fichier1 -Force}  
       $fichier2 = ".\ScanPC\Resultat_Scan.txt"            
       $test2 = Test-Path $fichier2
           If ($test2 -ne "True") {New-Item -ItemType file -Path $fichier2 -Force}
       $fichier3 = ".\ScanPC\PC_Sans_Reponse.txt"
       $test3 = Test-Path $fichier3
           If ($test3 -ne "True") {New-Item -ItemType file -Path $fichier3 -Force}
       $fichier5 = ".\ScanPC\Erreur_Scan.txt"
       $test5 = Test-Path $fichier5
           If ($test5 -ne "True") {New-Item -ItemType file -Path $fichier5 -Force}

       If (0 -ne $PC.Length) {
           If($PC -Match "\d{1,3}.\d{1,3}.\d{1,3}.\d{1,3}") {$HostNames = $PC}
           Else  {$HostNames = Get-ADComputer -Filter { cn -Like $PC }| Select -Expand Name}
       }
       ElseIf ($test1 -eq "true") {$HostNames = Get-Content ".\ScanPC\List.txt"}
       Else {$HostNames = Get-ADComputer -Filter { cn -Like '*' }| Select –Expand Name}
    }
    ElseIf (0 -ne $PC.Length) {
       If($PC -Match "\d{1,3}.\d{1,3}.\d{1,3}.\d{1,3}") {$HostNames = $PC}
       Else {$HostNames = Get-ADComputer -Filter { cn -Like $PC }| Select -Expand Name}
    }
    Else  {$HostNames = Get-ADComputer -Filter { cn -Like '*' }| Select -Expand Name}
              
    foreach($HostName in $HostNames) {
      Try {
	     $ping = new-object System.Net.NetworkInformation.Ping
	     $reply = $ping.send($HostName)
	     If ($reply.status –eq [System.Net.NetworkInformation.IPStatus]::Success) {
		     $connectedUser = (Get-WMIObject -class Win32_ComputerSystem  -ComputerName $HostName  | select username).username
		     If ($Logs){
			     If ($test1 -ne "true") {Add-Content -Path $fichier1 -Value "$HostName"}
        	     Add-Content -Path $fichier2 -Value "$HostName; $connectedUser"
             }
             Add-Content -Path $fichier4 -Value "$HostName ; $connectedUser"
	         write-host "[$HostName] est connecte => $connectedUser"
	     } 
         Else {
	         If ($Logs) {
	             If ($test1 -ne "true") {Add-Content -Path $fichier1 -Value "$HostName"}
		         Add-Content -Path $fichier3 -Value "$HostName ; Pas de réponse"
		     }
	         Add-Content -Path $fichier4 -Value "$HostName ; Pas de réponse"
	         write-host "[$HostName] Pas de réponse"
	     }
	  }
      Catch {
         $ErrorMessage = $_.Exception.Message
         $FailedItem = $_.Exception.ItemName
         If ($Logs) {
             If ($test1 -ne "true") {Add-Content -Path $fichier1 -Value "$HostName"}
             Add-Content -Path $fichier5 -Value "$HostName ; Erreur sur la machine => $FailedItem : $ErrorMessage"
         }
         Add-Content -Path $fichier4 -Value "$HostName ; Erreur sur la machine => $FailedItem : $ErrorMessage"
         write-host "[$HostName]  Erreur sur la machine => $FailedItem : $ErrorMessage"
      }  
      $reply  = $null
    }

} 


'on Error Resume Next

'Définition des variables locales
Dim fso, fCsv, tb
Const ForReading = 1
Const CheminAD = "ou=N°2,ou=N°1,dc=N°2,dc=N°1"
Set fso = CreateObject("scripting.filesystemobject")
Set NewServInfo  = fso.CreateTextFile("c:\test\NewServInfo.txt",TRUE)
Set fCsv = fso.OpenTextFile("PCListePrincipal.csv", ForReading)
	
'Parourir tous les Serveur du Fichier CSV
If Not fCsv.AtEndOfStream Then fCsv.ReadLine 		'Lecture de toutes les lignes du fichier CSV
	Do While Not fCsv.AtEndOfStream					'La ligne d'entête du fichier CSV est exclue
		tb = Split(fCsv.ReadLine, ";")				'Split les donnée du CSV pour traitement separée 
		If UBound(tb) = 4 Then
			NomPC   = tb(0)
			NomDns	= tb(1)
			AddIPv4	= tb(2)
			OpSys	= tb(3)
			Remarque= tb(4)	
			Wscript.echo "Checking = Nom: " & NomPC & " / Type d'OS: " & OpSys & " / Nom DNS : " & NomDns
		End If
		
		'Reinisialization des variables pour les tests
		Test = "Null"
		
		'Test que le PC n'est pas Dans l'AD
		SearchAD NomPC, Test
		If (Test <> "Vrai") Then
			Wscript.echo "Serveur : " & NomPC & " n'existe pas"
		
			'Creation du Serveur dans L'active Directory
			SrvObject NomPC , OpSys , NomDns
		End IF
	Loop	

Function SearchAD(NomPC, Test)

	'Creation d'une connection a l'Active Directory
	Const ADS_SCOPE_SUBTREE = 2
	Set objConnection = CreateObject("ADODB.Connection")
	Set objCommand =   CreateObject("ADODB.Command")
	objConnection.Provider = "ADsDSOObject"
	objConnection.Open "Active Directory Provider"
	Set objCommand.ActiveConnection = objConnection
	objCommand.Properties("Page Size") = 1000
	objCommand.Properties("Searchscope") = ADS_SCOPE_SUBTREE 
	objCommand.CommandText = _
		"SELECT 'sAMAccountName','cn' FROM 'LDAP://"& CheminAD &"' WHERE objectCategory='Computer'"
	Set objRecordSet = objCommand.Execute
	objRecordSet.MoveFirst

	'Parourir tous les PC de LDAP 		
	Do Until objRecordSet.EOF
    
		'On affiche a l'ecran le Nom du PC en cour de traitement
		PCName = objRecordSet.Fields("cn").Value
		'Wscript.Echo "PC Name= " & PCName
		
		'Test que le Serveur de la liste (NomPC) n'est pas Dans l'AD
		If (Ucase(NomPC) = Ucase(PCName)) Then
			Wscript.echo "Serveur : " & NomPC & " existe deja dans l'Active Directory"
			Wscript.echo " "
			Test = "Vrai"
			Exit Do
		End If
		
		'On pass au Prochain PC de LDAP
		Test = "Faux"
		objRecordSet.MoveNext
	Loop
End Function

Function SrvObject(NomPC , OpSys , NomDns)
	
	'On ce Place dans l'AD (à l'endroit ou les Serveurs doivent être crées)
	Set obj = GetObject("LDAP://" & CheminAD)
	
	'On Crée Le serveur avec toutes ces Info
	WScript.Echo " "
	WScript.Echo "Creation du Serveur : " & NomPC 
	Set NewServ = obj.create ( "Computer", "cn="&NomPC )
		NewServ.cn 				= NomPC
		NewServ.dNSHostName		= NomDns
		NewServ.operatingSystem	= OpSys
		NewServ.sAMAccountName  = NomPC & "$"
		'NewServ.Description 	= AddIPv4 & Remarque
		
		'On L'active et on valide la Creation
		NewServ.userAccountControl= 4096 '(4096 = WORKSTATION_TRUST_ACCOUNT)
		NewServ.setinfo
  
		'WScript.Echo "Serveur : " & NomPC & " a ete crée dans l'Active Directory"
	
	'Verification de la Creation
	NewServAD = "cn=" & NomPC & "," & CheminAD
	Set obj = GetObject("LDAP://" & NewServAD)
		
		'Collect des Info du Nouveau Serveur
		SamName		= obj.sAMAccountName
		ServName 	= obj.Name
		ServNote	= obj.Description
		DnsName 	= obj.dNSHostName
		OSType 		= obj.operatingSystem
		CheminServ	= obj.distinguishedName
			
		'Verification des info a l'ecrean
		Wscript.Echo " ==================== Serveur Name = * " & SamName & " * Debut ==================== "
		Wscript.Echo " " 
		Wscript.Echo "SamAccountName       = " & SamName
		Wscript.Echo "Nom du Serveur       = " & ServName
		Wscript.Echo "Nom DNS du Serveur   = " & DnsName
		Wscript.Echo "OS du Serveur	       = " & OSType
		Wscript.Echo "Description          = " & ServNote
		Wscript.Echo "Chemin AD du Serveur = " & CheminServ
		Wscript.Echo " "
		Wscript.Echo " ==================== Serveur Name = * " & SamName & " * Fin ====================== "
		Wscript.Echo " " 
			
		'Collect des info dans un fichier txt
		Data = ServName & ";" & DnsName & ";" & OSType & ";" & CheminServ
		NewServInfo.WriteLine(Data)
		'Wscript.Echo Data
		'Wscript.Echo " "
	
		Data  = " "
End Function

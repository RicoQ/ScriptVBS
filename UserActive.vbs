on Error Resume Next

Dim fso, fCsv 
Const ForReading = 1
Set fso = CreateObject("scripting.filesystemobject")
Set UserNotFound = fso.CreateTextFile("c:\test\UserNotFound.txt",TRUE)
Set UserInfo  = fso.CreateTextFile("c:\test\UserInfo.txt",TRUE)

Const ADS_SCOPE_SUBTREE = 2
Set objConnection = CreateObject("ADODB.Connection")
Set objCommand =   CreateObject("ADODB.Command")
objConnection.Provider = "ADsDSOObject"
objConnection.Open "Active Directory Provider"
Set objCommand.ActiveConnection = objConnection
objCommand.Properties("Page Size") = 1000
objCommand.Properties("Searchscope") = ADS_SCOPE_SUBTREE 
objCommand.CommandText = _
	"SELECT 'sAMAccountName','distinguishedName' FROM 'LDAP://ou=Racine,dc=ALES,dc=local' WHERE objectCategory='User'"
Set objRecordSet = objCommand.Execute
objRecordSet.MoveFirst

'Parourir tous les Utillisateur de LDAP 		
Do Until objRecordSet.EOF
    
	'Preparations des element pour les test
	Test = 0
	Data = " "
	
	'On affiche a l'ecran le Nom de l'Utilisateur en cour de traitement
	User = objRecordSet.Fields("sAMAccountName").Value
	'Wscript.Echo User
		
	'On vas Chercher l'utilisateur
	UserObject objRecordSet, User, Test
		
	'On pass au Prochain Utillisateur LDAP
	objRecordSet.MoveNext
Loop

Function UserObject(objRecordSet, User, Test)
	
	'Infos du chemin de l'utilisateur à traiter dans LDAP
	CheminUser = objRecordSet.Fields("distinguishedName").Value
	'Wscript.Echo CheminUser
	
	'On vas Chercher le "CN" l'Utilisateur 
	Set obj = GetObject("LDAP://" & CheminUser)
		
	'Collection des infos de l'Utilisateur à traiter dans LDAP
	SurNom 		= obj.sn
	Prenoms 	= obj.givenName
	Matricule	= obj.employeeID
	BadgeID		= obj.employeeNumber
	LastLogon = LargeIntegerToDate(obj.lastLogon)
	LastLogTime = LargeIntegerToDate(obj.lastLogonTimestamp)
		
		
	'Verification des info a l'ecrean
	'Wscript.Echo " ==================== User = * " & User & " * Debut ==================== "
	'Wscript.Echo " " 
	'Wscript.Echo User
	'Wscript.Echo SurNom
	'Wscript.Echo Prenoms
	'Wscript.Echo LastLogon
	'Wscript.Echo LastLogTime
	'Wscript.Echo CheminUser
	'Wscript.Echo " "
	'Wscript.Echo " ==================== User = * " & User & " * Fin ====================== "
	'Wscript.Echo " " 
			
	'Collect des info dans un fichier txt
	Info = User & ";" & SurNom & ";" & Prenoms & ";" & Matricule & ";" & BadgeID & ";" & LastLogon & ";" & LastLogonTime & ";" & CheminUser
	UserInfo.WriteLine(Info)
	'Wscript.Echo Info
	'Wscript.Echo " "
	
	Do While Test = 0
		
		'Test pour trouver les Utillisateurs a Modifier
		Data = "Faux"
		
		'Test de l' Utilisateur contre la List Principal (List.csv)
		TestCSV1 = ListCSV(obj, Data, SurNom, Prenoms, User, Matricule, BadgeID, CheminUser)
		If (Data = "Vrai") Then
			Wscript.echo "Utilisateur Trouver dans la List Principal: " & User & " / Nom: " & SurNom & " / Prenom: " & Prenoms
			Exit Do
		Else
			'Test de l' Utilisateur contre la List Secondaire (UserError.csv)
			TestCSV2 = UserErrorCSV(obj, Data, SurNom, Prenoms, User, Matricule, BadgeID, CheminUser)
			If (Data = "Aussi Vrai") Then
				Wscript.echo "Utilisateur Trouver dans la List Secondaire: " & User & " / Nom: " & SurNom & " / Prenom: " & Prenoms
				Exit Do
			Else
				'Si L'utilisateur n'est pas dans une des deux list on fait une entrer dans un Fichier txt (UserNotFound.txt)
				Wscript.echo "Utilisateur Introuvable : " & User & " / Nom: " & SurNom & " / Prenom: " & Prenoms & " / Last Logon: " & LastLogon & " / Chemin AD: " & CheminUser
				UserNotFound.WriteLine(Info)   'on ecrie dans le Fichier
				Exit Do
			End If
		End IF
		
		'On Reset les Variables de test pour le prochain Utillisateur
		'Test  = " "
		Data  = " "
	Loop		
End Function

Function LargeIntegerToDate(value) 
'takes Microsoft LargeInteger value (Integer8) and returns according the date and time 

	'first determine the local time from the timezone bias in the registry 
	Set sho = CreateObject("Wscript.Shell") 
	timeShiftValue = sho.RegRead("HKLM\System\CurrentControlSet\Control\TimeZoneInformation\ActiveTimeBias") 
	If IsArray(timeShiftValue) Then 
		timeShift = 0 
		For i = 0 To UBound(timeShiftValue) 
			timeShift = timeShift + (timeShiftValue(i) * 256^i) 
		Next 
	Else 
		timeShift = timeShiftValue 
	End If 
	
	'get the large integer into two long values (high part and low part) 
	i8High = value.HighPart 
	i8Low = value.LowPart 
	If (i8Low < 0) Then    
		i8High = i8High + 1 
	End If 
	
	'calculate the date and time: 100-nanosecond-steps since 12:00 AM, 1/1/1601 
	If (i8High = 0) And (i8Low = 0) Then 
		LargeIntegerToDate = #1/1/1601# 
	Else 
		LargeIntegerToDate = #1/1/1601# + (((i8High * 2^32) + i8Low)/600000000 - timeShift)/1440 
	End If 
End Function

Function ListCSV(obj, Data, SurNom, Prenoms, User, Matricule, BadgeID, CheminUser)
	
	'Définition des variables locales
	Dim fso, fCsv, tb 
	Const ForReading = 1
	Set fso = CreateObject("scripting.filesystemobject")
	Set fCsv = fso.OpenTextFile("ListePrincipal.csv", ForReading)
	
    'Parourir tous les Utillisateur du Fichier CSV
	If Not fCsv.AtEndOfStream Then fCsv.ReadLine 		'Lecture de toutes les lignes du fichier CSV
		Do While Not fCsv.AtEndOfStream					'La ligne d'entête du fichier CSV est exclue
			tb = Split(fCsv.ReadLine, ";")				'Split les donnée du CSV pour traitement separée 
			If UBound(tb) = 5 Then
				ID      = tb(0)
				Name 	= tb(1)
				SurName	= tb(2)
				NomNess = tb(3)
				Matri 	= tb(4)
				Badge 	= tb(5)
				'Wscript.echo " ========== List.csv ========== "
				'Wscript.echo "Checking = Nom: " & Name & " / Prenom: " & SurName & " / Matricule: " & Matri & " / Badge N°: " & Badge 
			End If
			
			'Creation des Variable de test
			TestName     = SansAccents(Name)    & "-" & SansAccents(SurName)
			TestPrenom   = SansAccents(SurName) & "-" & SansAccents(Name)
			TestNomAD    = SansAccents(SurNom)  & "-" & SansAccents(Prenoms)
			TestPrenomAD = SansAccents(Prenoms) & "-" & SansAccents(SurNom)
			'Wscript.Echo "Testing = " & TestPrenom & "/" & TestName & "/" & TestNomAD & "/" & TestPrenomAD
			
			'Test si l'utilisateur ce trouve dans la List Principal (List.csv)
			If ((Ucase(User) = Ucase(TestName)) or (Ucase(User) = Ucase(TestPrenom)) or (Ucase(TestNomAD) = Ucase(TestName)) or (Ucase(TestPrenomAD) = Ucase(TestName)) or (Ucase(TestNomAD) = Ucase(TestPrenom)) or (Ucase(TestPrenomAD) = Ucase(TestPrenom)) or (Matricule = Matri)) Then
							
				'Modification des Champ LDAP pour les Utillisateurs Trouver
				'Wscript.echo " "
				'Wscript.echo " **************** Initialisation des Modification pour " & User & " **************** "
				obj.Put "employeeID", Matri
				'Wscript.echo Matri
				obj.Put "employeeNumber", Badge
				'Wscript.echo Badge
				'Wscript.echo " **************** Modification Fini **************** "
				'Wscript.echo " "
		
				'Validation des Modification
				obj.SetInfo
							
				'Verification des Modification
				Data = "Vrai"
			
				'Sortie de la loop si un Utillisateur est trouver
				Exit Do 
			End IF
		Loop 
End Function

Function UserErrorCSV(obj, Data, SurNom, Prenoms, User, Matricule, BadgeID, CheminUser)
	
	'Définition des variables locales
	Dim fso, fCsv, tb 
	Const ForReading = 1
	Set fso = CreateObject("scripting.filesystemobject")
	Set fCsv = fso.OpenTextFile("ListeSecondaire.csv", ForReading)
	
    'Parourir tous les Utillisateur du Fichier CSV
	If Not fCsv.AtEndOfStream Then fCsv.ReadLine 		'Lecture de toutes les lignes du fichier CSV
		Do While Not fCsv.AtEndOfStream					'La ligne d'entête du fichier CSV est exclue
			tb = Split(fCsv.ReadLine, ";")				'Split les donnée du CSV pour traitement separée 
			If UBound(tb) = 5 Then
				SamName	= tb(0)
				Name 	= tb(1)
				SurName	= tb(2)
				Matri 	= tb(3)
				Badge 	= tb(4)
				CnChemin= tb(5)
				'Wscript.echo " ========== UserError.csv ========== "
				'Wscript.echo "NameAD: " & SamName & " / Nom: " & Name & " / Prenom: " & SurName & " / Matricule: " & Matri & " / Badge N°: " & Badge & " / Chemin AD: " & CnChemin
			End IF
			
			'Test si l'utilisateur ce trouve dans la List Secondaire (UserError.csv)
			If ((Ucase(User) = Ucase(SamName)) or (Matricule = Matri) Or (Ucase(CnChemin) = Ucase(CheminUser))) Then
				Wscript.Echo " Chemin Utilisateur Trouver =========== Pour: " & User
				
				'Modification des Champ LDAP pour les Utillisateurs Trouver
				'Wscript.echo " "
				'Wscript.echo " **************** Initialisation des Modification pour " & User & " **************** "
				obj.Put "employeeID", Matri
				'Wscript.echo Matri
				obj.Put "employeeNumber", Badge
				'Wscript.echo Badge
				'Wscript.echo " **************** Modification Fini **************** "
				'Wscript.echo " "
		
				'Validation des Modification
				obj.SetInfo
							
				'Verification des Modification
				Data = "Aussi Vrai"
			
				'Sortie de la loop si un Utillisateur est trouver
				Exit Do 
			End IF
		Loop 
End Function

Function SansAccents(strAvecAccents)

	Const ACCENT = "ÀÁÂÃÄÅàáâãäåÒÓÔÕÖØòóôõöøÈÉÊËèéêëÌÍÎÏìíîïÙÚÛÜùúûüÿÑñÇç"
	Const NOACCENT = "AAAAAAaaaaaaOOOOOOooooooEEEEeeeeIIIIiiiiUUUUuuuuyNnCc"
 
	' Définition des variables locales
	Dim i
	Dim lettre
	Dim strSansAccents
 
	strSansAccents = strAvecAccents
	  For i = 1 To Len(ACCENT)
	    lettre = Mid(ACCENT, i, 1)
	    If InStr(strSansAccents, lettre) > 0 Then
	       strSansAccents = Replace(strSansAccents, lettre, Mid(NOACCENT, i, 1))
	    End If
	  Next
	SansAccents = strSansAccents
 
	' Libération des variables locales
	Set i = Nothing
	Set lettre = Nothing
	Set strSansAccents = Nothing
 
End Function 

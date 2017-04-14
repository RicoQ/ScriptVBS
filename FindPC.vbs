'On Error Resume Next

Dim fso, fCsv 
Const ForReading = 1
Set fso = CreateObject("scripting.filesystemobject")
Set InfoPcAD = fso.CreateTextFile("c:\test\InfoPcAD.txt",TRUE)


'Preparations des element pour parcourire Le LDAP
Const ADS_SCOPE_SUBTREE = 2
Set objConnection = CreateObject("ADODB.Connection")
Set objCommand =   CreateObject("ADODB.Command")
objConnection.Provider = "ADsDSOObject"
objConnection.Open "Active Directory Provider"
Set objCommand.ActiveConnection = objConnection
objCommand.Properties("Page Size") = 1000
objCommand.Properties("Searchscope") = ADS_SCOPE_SUBTREE 
objCommand.CommandText = _
	"SELECT 'sAMAccountName','distinguishedName','Name','cn','dNSHostName','operatingSystem' FROM 'LDAP://ou=Serveurs,dc=ALES,dc=local' WHERE objectCategory='Computer'"
Set objRecordSet = objCommand.Execute
objRecordSet.MoveFirst

'Parourir tous les Ordinateur du LDAP 		
Do Until objRecordSet.EOF
    
	'Preparations des element pour les test
	Test = 1
	
	'Collection des infos pour chaque Ordinateur a traiter dans le LDAP
	CheminPC 	= objRecordSet.Fields("distinguishedName").Value
	SamName 	= objRecordSet.Fields("sAMAccountName").Value
	PCName 		= objRecordSet.Fields("Name").Value
	CnName 		= objRecordSet.Fields("cn").Value
	DsnName 	= objRecordSet.Fields("dNSHostName").Value
	OpSys  		= objRecordSet.Fields("operatingSystem").Value
			
	'Verification les info a l'ecrean
	Wscript.Echo " ====== Found ======== PC : * " & SamName & " * Debut ==================== "
	Wscript.Echo " " 
	Wscript.Echo "Sam Account Name= " & SamName
	Wscript.Echo "Nom CN= " & CnName
	Wscript.Echo "Nom PC= " & PCName
	Wscript.Echo "Nom DNS= " & DsnName
	Wscript.Echo "OS= " & objRecordSet.Fields("operatingSystem").Value
	Wscript.Echo "Chemin AD=" & CheminPC
	Wscript.Echo " "
	Wscript.Echo " ==================== PC : * " & SamName & " * Fin ====================== "
	Wscript.Echo " " 
			
	'Collect des info pour insersion dans un fichier txt
	Info = SamName & ";" & CnName & ";" & PCName & ";" & DsnName & ";" & OpSys & ";" & CheminPC
	
	'On ecrie dans un Fichier TXT
	InfoPcAD.WriteLine(Info)
   
	'On pass au prochain Ordinateur LDAP
	objRecordSet.MoveNext
	
Loop 'Fin des Ordinateur du LDAP

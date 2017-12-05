' List all the signals of the project with their associated devices

' v0: D.KOZIEL (SOGETI HT) 19/12/2011
'				-> Création du script
' v1: D.KOZIEL (SOGETI HT) 27/02/2013
'				-> Ajout des colonnes arborescence, fonction, folio et section
'				-> Si pas de composant le nom d'appareil est renseigné dans la colonne composant
' v2: D.KOZIEL (SOGETI HT) 21/10/2013
'				-> Ajout de la colonne Variable contenant les noms d'entrées/sorties RIOM
' v3: D.KOZIEL (SOGETI HT) 13/11/2013
'				-> Ajout de la colonne "Titre du Folio"
' v4: D.KOZIEL (SOGETI HT) 10/12/2013
'				-> Correction de l'affichage de l'arborescence, affiche multi-projet
'				-> Ajout de la colonne calibre
' v5: D.KOZIEL (SOGETI HT) 24/04/2014
'				-> Ajout des noms de connecteurs
' v6: D.KOZIEL (SOGETI HT) 11/05/2014
'				-> Sortie de la subdivision la plus discriminante (entre feuille et appareil)
' v7: D.KOZIEL (SOGETI HT) 13/06/2014
'				-> Prise en compte des réseaux possédant plusieurs signaux
' v8: D.KOZIEL (SOGETI HT) 19/06/2014
'				-> Prise en compte des signaux presents sur les pins sans réseaux 

' a0: A.LEFOSA (GIBELA) 05/12/2017
'				-> Code translation French - English(RSA) 

'----------------------------------------------------------------------------------------------------------
' Definition of variables
'----------------------------------------------------------------------------------------------------------

Option Explicit

Const xlContinuous       	= 1					' Excel definitions
Const xlDiagonalDown     	= 5
Const xlDiagonalUp       	= 6
Const xlEdgeBottom       	= 9
Const xlEdgeLeft         	= 7
Const xlEdgeRight        	= 10
Const xlEdgeTop          	= 8
Const xlInsideHorizontal 	= 12
Const xlInsideVertical   	= 11
Const xlNone            	= -4142
Const xlThin             	= 2
Const xlDouble           	= -4119
Const xlAutomatic        	= -4105
Const xlThick            	= 4
Const xlLandscape        	= 2
Const xlCenter		  		= -4108
Const xlLeft			  	= -4131

Dim E3s_Application,E3s_Projet
Dim E3s_Appareil,E3s_composant,E3s_Signal,E3s_Pin,E3s_Pin_2,E3s_Net,E3s_NetSegment,E3s_Noeud,E3s_Feuille,E3s_Symbole,E3s_Texte
Dim nb_signaux,ID_signaux,nb_netsegments,ID_netsegments,nb_pins,ID_pins,nb_appareils,ID_appareils,nb_text,ID_text
Dim i,j,k,Arborescence,Location,Section,x,y,grid,num_pin,nom_pin,variables_RIOM,erreur,Destination_Connexion,Assembly_Device_Destination
Dim dico,combinaison

Dim xl_Application,nline


Const nb_colonnes = 19
ReDim Tab_Excel(150000,nb_colonnes)

'----------------------------------------------------------------------------------------------------------
' Creating of objects
'----------------------------------------------------------------------------------------------------------
Set E3s_Appareil = CreateObject("CT.Device")
Set E3s_composant = CreateObject("CT.Component")
Set E3s_Signal = CreateObject("CT.Signal")
Set E3s_Pin = CreateObject("CT.Pin")
Set E3s_Pin_2 = CreateObject("CT.Pin")
Set E3s_Net = CreateObject("CT.Net")
Set E3s_NetSegment = CreateObject("CT.NetSegment")
Set E3s_Feuille = CreateObject("CT.Sheet")
Set E3s_Symbole = CreateObject("CT.Symbol")
Set E3s_Texte = CreateObject("CT.Text")

Set xl_Application = CreateObject("Excel.Application")

Set dico = CreateObject("Scripting.Dictionary")

'----------------------------------------------------------------------------------------------------------
' Establish the connection
'----------------------------------------------------------------------------------------------------------

If Not (Ouvrir_connexion()) Then
    WScript.Quit
End If

'----------------------------------------------------------------------------------------------------------
' Extraction of signals from E3_series
'----------------------------------------------------------------------------------------------------------
nb_signaux = E3s_Projet.GetSignalIds(ID_signaux)
E3s_Application.PutMessage "Lancement liste des potentiels"
For i=1 To nb_signaux
	E3s_Signal.SetId ID_signaux(i)
	nb_pins = E3s_Signal.GetPinIds(ID_pins)
	For j=1 To nb_pins
		E3s_Pin.SetId ID_pins(j)
		E3s_Appareil.SetId E3s_Pin.GetId
		E3s_composant.SetId  E3s_Appareil.GetId
		nb_netsegments = E3s_Pin.GetNetSegmentIds(ID_netsegments)
		If nb_netsegments = 0 Then
			E3s_Pin_2.SetId E3s_Pin.GetConnectedPinId
			nb_netsegments = E3s_Pin_2.GetNetSegmentIds(ID_netsegments)
		End If
		For k=1 To nb_netsegments
			E3s_NetSegment.SetId ID_netsegments(k)
			E3s_Net.SetId E3s_NetSegment.GetNetId
		Next
		If E3s_Appareil.GetName <> "" Then
			combinaison = E3s_Appareil.GetAssignment & E3s_Appareil.GetLocation & E3s_Appareil.GetName & E3s_Pin.GetId
			If Not dico.Exists(combinaison) Then
				Destination_Connexion = ""
				Assembly_Device_Destination = ""
				dico.Add combinaison,""
				Tab_Excel(nline,0) = E3s_Projet.GetGidOfId(E3s_Signal.GetId)
				Tab_Excel(nline,1) = E3s_Signal.GetName
				If E3s_Net.GetAttributeValue("CLASSE_N")<>"" Then
					Tab_Excel(nline,2) = E3s_Net.GetAttributeValue("CLASSE_N")
				Else
					Tab_Excel(nline,2) = "-"
				End If
				
				' We recover the tree in which the sheet is placed
				Arborescence = ""
				Set E3s_Noeud = E3s_Projet.CreateStructureNodeObject
				E3s_Symbole.SetId E3s_Pin.GetId
				E3s_Feuille.SetId E3s_Symbole.GetId
				E3s_Noeud.SetId E3s_Feuille.GetId
				Do
					Arborescence =  E3s_Noeud.GetName & "/" & Arborescence
					E3s_Noeud.SetID E3s_Noeud.GetParentId
				Loop While E3s_Noeud.GetID <>0
				Set E3s_Noeud = Nothing
				If InStr(Arborescence, "/SPC/") = 1 Then
					Arborescence = Replace(Arborescence, "/SPC/=", "")
				Else
					Arborescence = Mid(Arborescence, 2)
				End If
				Location = InStr(Arborescence, "/")
				Tab_Excel(nline, 3) = Mid(Arborescence, 1, Location)
				Arborescence = Mid(Arborescence, Location + 2)
				Location = InStr(Arborescence, "/")
				Tab_Excel(nline, 4) = Mid(Arborescence, 1, Location)
				Tab_Excel(nline, 5) = E3s_Feuille.Getname
				Location = E3s_Pin.GetSchemaLocation(x, y, grid)
				Section = Mid(grid,Instr(1, grid, ".")+1,50)
				Tab_Excel(nline, 6) = Section
				
				
				If E3s_Feuille.GetAssignment<>"" Then
					If InStr(E3s_Feuille.GetAssignment, "INFO") <> 0 Or E3s_Appareil.GetAssignment="" Then
						Tab_Excel(nline,7) = Mid(E3s_Feuille.GetAssignment,2,100)
					Else
						If Len(E3s_Feuille.GetAssignment) < Len(E3s_Appareil.GetAssignment) Then
							Tab_Excel(nline,7) = Mid(E3s_Feuille.GetAssignment,2,100)
						Else
							Tab_Excel(nline,7) = Mid(E3s_Appareil.GetAssignment,2,100)
						End If
					End If
				Else
					Tab_Excel(nline,7) = "-"
				End If
				If E3s_Appareil.GetLocation<>"" Then
					Tab_Excel(nline,8) = Mid(E3s_Appareil.GetLocation,2,100)
				Else
					Tab_Excel(nline,8) = "-"
				End If
				Tab_Excel(nline,9) = Mid(E3s_Appareil.GetName,2,100)
				
				'We recover the attributes of pins
				nb_text = E3s_Pin.GetTextIds (ID_Text)
				If nb_text = 2 And E3s_Symbole.IsDynamic Then
					For k=1 To nb_text
						E3s_Texte.SetId ID_text(k)
						Select Case E3s_Texte.GetTypeId	
							Case 3 num_pin = E3s_Texte.GetText
							Case 820 nom_pin = E3s_Texte.GetText
						End Select
					Next
					Tab_Excel(nline,10) = num_pin
					Tab_Excel(nline,11) = nom_pin
				ElseIf E3s_Pin.GetName<>"" Then
					Tab_Excel(nline,10) = E3s_Pin.GetName
					Tab_Excel(nline,11) = "-"
				Else
					Tab_Excel(nline,10) = "-"
					Tab_Excel(nline,11) = "-"
				End If
				If E3s_composant.GetName<>"" Then
					Tab_Excel(nline,12) = E3s_composant.GetName
				Else
					Tab_Excel(nline,12) = Mid(E3s_Appareil.GetName,2,100)
				End If
				If E3s_Symbole.IsDynamic Then
					Tab_Excel(nline,13) = "Dynamic"
				ElseIf E3s_composant.GetAttributeValue("Class")<>"" Then
					Tab_Excel(nline,13) = E3s_composant.GetAttributeValue("Class")
				Else
					Tab_Excel(nline,13) = "-"
				End If
				
				'We get the RIOM input / output names
				nb_text = E3s_Pin.GetTextIds (ID_Text)
				erreur = 0
				For k=1 To nb_text
					E3s_Texte.SetId ID_text(k)
					Select Case E3s_Texte.GetTypeId	
						Case 100 variables_RIOM = E3s_Texte.GetText
							erreur = erreur + 1
						Case 1002 Destination_Connexion = E3s_Texte.GetText
						Case 1079 Assembly_Device_Destination = E3s_Texte.GetText
					End Select
				Next
				If erreur > 1 Then
					E3s_Application.PutMessage "ERREUR - Incohérence E/S RIOM : Plusieurs noms de variables sur la même E/S"
					E3s_Application.PutMessage E3s_Appareil.GetName & "." & num_pin, E3s_Symbole.GetId
					Tab_Excel(nline,14) = "Erreur sur le symbole"
				ElseIf erreur = 1 Then
					Tab_Excel(nline,14) = variables_RIOM
				End If
				Tab_Excel(nline, 15) = E3s_Feuille.GetAttributeValue("SheetName1")
				' We recover the caliber
				If E3s_composant.GetAttributeValue("CALIBRE")<>"" Then
					Tab_Excel(nline,16) = E3s_composant.GetAttributeValue("CALIBRE")
				Else
					Tab_Excel(nline,16) = "-"
				End If
				Tab_Excel(nline,17) = Destination_Connexion
				If Destination_Connexion <> "" Then
					Destination_Connexion = Mid(Destination_Connexion, InStr(Destination_Connexion, "-") + 1)
					If InStr(Destination_Connexion, ":") <> 0 Then
						Destination_Connexion = Mid(Destination_Connexion, 1, InStr(Destination_Connexion, ":") - 1)
					End If
					If Assembly_Device_Destination = "" Then
						Tab_Excel(nline,18) = Mid(E3s_Appareil.GetName,2) & "_" & Destination_Connexion
					Else
						Tab_Excel(nline,18) = Assembly_Device_Destination & "_" & Destination_Connexion
					End If
				End If
				nline = nline + 1
			End if
		End If
		E3s_Net.SetId 0
	Next
Next


xl_application.Workbooks.add
xl_application.Visible = True
xl_application.sheets("Feuil1").Select
xl_Application.Cells.NumberFormat = "@"
xl_Application.Range(xl_Application.Cells(2, 1), xl_Application.Cells(nline+1, nb_colonnes+1))=Tab_Excel
Mise_en_forme_Excel xl_Application


E3s_Application.PutMessage "*** Terminé ***"

'----------------------------------------------------------------------------------------------------------
' Release of objects
'----------------------------------------------------------------------------------------------------------

Set E3s_Application = Nothing
Set E3s_Projet = Nothing
Set E3s_Appareil = Nothing
Set E3s_Signal = Nothing
Set E3s_Pin = Nothing
Set E3s_NetSegment = Nothing
Set E3s_Net = Nothing
Set E3s_Feuille = Nothing
Set E3s_Symbole = Nothing

Set xl_Application = Nothing

Set dico = Nothing

'----------------------------------------------------------------------------------------------------------
' Excel file formatting
'----------------------------------------------------------------------------------------------------------
Sub Mise_en_forme_Excel(ByRef Excel)
	With Excel
		.Cells(1,1) = "ID"
		.Cells(1,2) = "Signal"
		.Cells(1,3) = "Classe de Tension"
		.Cells(1,4) = "Arbo"
		.Cells(1,5) = "Fonction"
		.Cells(1,6) = "Folio"
		.Cells(1,7) = "Section"
		.Cells(1,8) = "Subdivision"
		.Cells(1,9) = "Loca"
		.Cells(1,10) = "Appareil"
		.Cells(1,11) = "Num Borne"
		.Cells(1,12) = "Nom Borne"
		.Cells(1,13) = "Composant"
		.Cells(1,14) = "Type Appareil"
		.Cells(1,15) = "Nom Variable"
		.Cells(1,16) = "Titre du Folio"
		.Cells(1,17) = "Calibre"
		.Cells(1,18) = "DESTINATION"
		.Cells(1,19) = "Nom Connecteur"
        .Range(xl_Application.Cells(1,1), xl_Application.Cells(1,nb_colonnes)).Select
		.Selection.Interior.ColorIndex = 42
		.Selection.Interior.Pattern = 1
		With .Selection.Borders(xlEdgeLeft)
			.LineStyle = xlContinuous
			.Weight = 3
			.ColorIndex = xlAutomatic
		End With
		With .Selection.Borders(xlEdgeTop)
			.LineStyle = xlContinuous
			.Weight = 3
			.ColorIndex = xlAutomatic
		End With
		With .Selection.Borders(xlEdgeBottom)
			.LineStyle = xlContinuous
			.Weight = 3
			.ColorIndex = xlAutomatic
		End With
		With .Selection.Borders(xlEdgeRight)
			.LineStyle = xlContinuous
			.Weight = 3
			.ColorIndex = xlAutomatic
		End With
		With .Selection.Borders(xlInsideVertical)
			.LineStyle = xlContinuous
			.Weight = 3
			.ColorIndex = xlAutomatic
		End With
		.Selection.Font.Size = 10							'font size of the text
		.Selection.Font.Bold = True
		.Selection.EntireColumn.Select
		.Selection.HorizontalAlignment = xlCenter
		.Selection.VerticalAlignment = xlCenter
		.Selection.AutoFilter
		.Selection.Columns.AutoFit
		.Range("A2").Select
		.ActiveWindow.FreezePanes = True
		.Cells(1,1).Select
	End With
End Sub

'----------------------------------------------------------------------------------------------------------
' Function to establish the connection with E3s
'----------------------------------------------------------------------------------------------------------
' Returns True if the connection is successful, False if it failed
' NB : We can run the script internally (launched from E3S) or externally (by clicking on the script from the explorer for example)
'
Function Ouvrir_connexion()
	Dim shell,reponse,Tableau_projets(),disp, nb_projets, Chaine, objWMIService, colitems, objitem, lst,strComputer
	If InStr(WScript.FullName, "E³") Then                                                           	' Internal case
		Ouvrir_connexion = True	
		Set E3s_Application = WScript
		Set E3s_Projet = E3s_Application.CreateJobObject
		E3s_Application.PutMessage "Project executé en interne"

	Else                                                                                            	' If executed externally
		On Error Resume Next																			' ignore error if dispatcher not installed
		Set disp   = CreateObject("CT.Dispatcher")        												'
		Set viewer = CreateObject("CT.DispatcherViewer")
		On Error GoTo 0
		Set E3s_Application = Nothing
		If IsObject(disp) Then																			' We check if the dispatcher is installed
			nb_projets = disp.GetE3Applications(lst)
			' If more than one project is open, the user is asked to choose the one he wants to open
			If nb_projets > 1 then
				If viewer.ShowViewer(e3Obj) = True Then													' View the list of loaded projects
					Set E3s_Application = e3Obj															' Take the selected
					Set E3s_Projet = E3s_Application.CreateJobObject									' the corresponding project
					Ouvrir_connexion = true
				Else
					Ouvrir_connexion = False   
				End If
			ElseIf nb_projets = 1 Then																		' If not the choice, we open the first project
				Set E3s_Application = CreateObject("CT.Application")
				Set E3s_Projet = E3s_Application.CreateJobObject
				Ouvrir_connexion = true
			Else																							' Otherwise if no projects open, we leave
				MsgBox ("Aucun projet n'est ouvert")
				Ouvrir_connexion = False   
			End If
		Else																								' dispatcher not installed
			strComputer = "."																	
			set objWMIService = GetObject("winmgmts:\\" & strComputer & "\root\cimv2")
			set colItems= objWMIService.ExecQuery("Select * from Win32_Process",,48)
			nb_projets = 0
			for each objItem in colItems
				if InStr(objItem.Caption, "E3.series") then nb_projets = nb_projets + 1
			next
			set objWMIService = Nothing
			set colItems      = Nothing
			If nb_projets>1 Then
				MsgBox  "Plus d'un projet est ouvert. Veuillez n'en laisser qu'un seul ouvert"
				Ouvrir_connexion = False
			Else
				Set E3s_Application = CreateObject("CT.Application")
				Set E3s_Projet = E3s_Application.CreateJobObject
				Ouvrir_connexion = true
			End If
		End If
	End If
End Function
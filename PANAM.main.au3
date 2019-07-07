#Region ;**** Directives created by AutoIt3Wrapper_GUI ****
#AutoIt3Wrapper_Outfile=PANAM_DUPLIQUERPF.Exe
#EndRegion ;**** Directives created by AutoIt3Wrapper_GUI ****
;----------------------------------------------------------------------------------------------------
#include-once
#include "Fonctions Metiers\VDS.au3"
;~ #include "Fonctions Metiers\VDT.au3"
#include "Fonctions Metiers\PF.au3"
;~ #include "Fonctions Metiers\PFOH.au3"
#include "Fonctions Metiers\Devis_Avenants.au3"
;~ #include "Fonctions Metiers\DEMANDES_ACTIVITES.au3"

;~ Opt("GUIOnEventMode", 1)

Dim $aTitle_Scr = [ _
		["Comptes", "Profils de facturation", "Interlocuteurs", "Demandes", "Accueil", "Opportunité", "Devis"], _
		["Compte", "Profil de facturation", "Interlocuteur", "Demande", "Accueil", "Opportunité", "Devis"], _
		["Réf. du Compte", "Réf. Profil de Factu.:", "E-mail:", "Réf. Affaire", "Accueil", "Réf. Affaire", "# Devis:"]]

;~ CLOERechercher("Comptes", "1-45*", true, True)
;~ Exit
;~ CLOERechercher("Compte")
;~ exit 0
#Region Tests ;************************Zone tests**************************************
;~  ChangerEnreMaJPrix()
;~  Terminer()
;~  Verif()
;~  Terminer()
;~ CLOERechercher("Demandes", "70501347")
;~ Exit 25

;~ Traiter("CreerVDSModifPSFTA", "CreerVDSModifPSFTA", "Comptes")
;~ Terminer()

;~ TraiterRapide("Avenant_ModifierInfosSites", "Modifier infos sites desservis", "Devis")
;~ Terminer() ;

;~ TraiterRapide("TraiterProchaineEtapeDevis", "Traitement  prochaine  Etaper Devis" , "Devis")
;~ Terminer() ;

;~ TraiterRapide("CloturerDemande", "Cloture Automatiques des Demandes OR", "Demandes", "Source.xlsm")
;~ Terminer() ;

;~ $LabelChampsRef = "def"
;~ Traiter("ResilierDevis", "Résiliation Devis en masse", "Devis")
;~ Terminer() ;

;~ Copier()
;~ DemandesOR()
;~ Terminer()

;~ $LabelChampsRef = "def"
;~ Traiter("InitierAvenant", "Initialiser Avenant", "Devis")
;~ Terminer() ;

;~ Traiter("DemandeORDevis", "DemandeORDevis", "Devis")
;~ Terminer() ;

;~ Traiter("QualifierDonneesPaiement", "Qualification des Donnees Paiement", "Profils de facturation")
;~ Terminer() ;

;~ TraiterClaudor("CloturerDemande", "Cloture Automatiques des Demandes OR", "Demandes", "Claudor.xls")
;~ Terminer()

TraiterDuplicate("DupliquerPF", "DupliquerPF", "Profils de facturation")
Terminer();

;~ Traiter("DemanderCLOEOR_MES", "DemanderCLOEOR_MES", "Comptes")
;~ Terminer()
;~ FermerDialog()
;~ ConsoleWrite('@@ Debug(' & @ScriptLineNumber & ') : FermerDialog() = ' & FermerDialog() & @CRLF & '>Error code: ' & @error & @CRLF) ;### Debug Console

;~ Traiter("SelectPF_OH", "SelectPF_OH", "Comptes") ; , "Claudor.xls")
;~ Terminer()


#EndRegion Tests ;************************Zone tests**************************************

Dim $ValeursChamps
Dim $NomFonction, $TitreApplication, $sNomEcranCLOE, $LabelChampsRef
Local $Commandes = StringSplit($cmdLineraw, "$")

If UBound($Commandes) > 3 Then
	$NomFonction = $Commandes[1]
	$TitreApplication = $Commandes[2]
	$sNomEcranCLOE = Int($Commandes[3])
EndIf

If StringInStr($TitreApplication, "chaine") > 0 Then
Else
	Traiter($NomFonction, $TitreApplication, $sNomEcranCLOE)
EndIf

;Gestion erreur COM
$oMyError = ObjEvent("AutoIt.Error", "ErrFunc")

Func ErrFunc()
	FermerIE()
	Traiter($NomFonction, $TitreApplication, $sNomEcranCLOE)
EndFunc   ;==>ErrFunc

Func TraiterDuplicate_Old($NomFonctionTraitement, $_TypeTraitement, $NomEcran = "Comptes")
	If @YEAR > 2019 Then
		MsgBox(0, "Erreur", "Erreur, Contacter l'administrateur")
		Exit 2
	EndIf

	$NNISesame = "C21373"
	$mdpSesame = "Lazare1!!"
	$Adresse = "https://pcyyyxp5.pcy.edfgdf.fr/cloe1/start.swe?SWECmd=Login&SWEFullRefresh=1&TglPrtclRfrsh=1"
	$NomAppl = "PANAM - "
	$TypeTraitement = $_TypeTraitement
	$NomFichierLog = "Journal (" & $_TypeTraitement & "_" & @YEAR & @MON & @MDAY & ").log"
	$NomFichierLogDialog = "Journal MsgDialog(" & $_TypeTraitement & "_" & @YEAR & @MON & @MDAY & ").log"
	Send("{CAPSLOCK OFF}")
	UpdateEvent("Merci de patienter. Recupération des données de traitement en cours")
	$FichierModelExcel = _Excel_BookAttach("Source.xlsm", "FileName")
	If IsObj($FichierModelExcel) = 0 Then
		MsgBox(0, "Erreur ficchier 'Source.xlsm'", "Fichier 'Source.xlsm' n'est pas  disponible", 5)
		Return 0
	EndIf

	Local $ArraySource = _Excel_RangeRead($FichierModelExcel, 1, "A5:AZ1000")
	Local $FichierModelNbreTotalLignes = Int($ArraySource[1][6])

	If InputBox("CONFIRMATION COPIE", "Souhaites-tu vraiment créer " & $FichierModelNbreTotalLignes & " copie(s)" & @CRLF & "Saisis ""CONFIRMER"" pour lancer le traitement") <> "CONFIRMER" Then
		Exit 1
	EndIf

	For $FichierModelLigneCourante = 1 To $FichierModelNbreTotalLignes + 1 ; int($ArraySource[1][6]) + 1   ; Ubound($ArraySource, 1) - 1
		$RefCompte = $ArraySource[1][0]
		UpdateEvent("Traitement en cours - Ref : " & $RefCompte)

		Local $StatutCLOEOpenObjet = 1
		If $StatutCLOEOpenObjet <> 0 Then
			UpdateEvent("Ref. Compte: " & $RefCompte & " Copie numèro " & $FichierModelLigneCourante & " en cours", $FichierModelLigneCourante, $FichierModelNbreTotalLignes)
			$ValeursChamps = _ArrayExtract($ArraySource, $FichierModelLigneCourante, $FichierModelLigneCourante)
			$ValeursChamps = Call($NomFonctionTraitement, $ValeursChamps)
			If IsArray($ValeursChamps) Then
				$ArraySource[$FichierModelLigneCourante][2] = SetRapport($ValeursChamps[0][2])
				$ArraySource[$FichierModelLigneCourante][3] = SetRapport($ValeursChamps[0][3])
				$ArraySource[$FichierModelLigneCourante][4] = SetRapport($ValeursChamps[0][4])
				$ArraySource[$FichierModelLigneCourante][5] = SetRapport($ValeursChamps[0][5])
			Else
				$ArraySource[$FichierModelLigneCourante][2] = $ValeursChamps
			EndIf

			If @error > 0 Then
				$CompteurErreur += 1
			Else
				$CompteurSucces += 1
				$CompteurErreur = 0
			EndIf
		Else
			$CompteurErreur += 1
			$ArraySource[$FichierModelLigneCourante][1] = SetRapport("La ref.compte: " & $ArraySource[$FichierModelLigneCourante][0] & " est introuvable!  :(")
		EndIf
		$FichierModelNbreLignesFaites += 1

		If @error = $ErrFermetureBoiteDialog Then
			$ArraySource[$FichierModelLigneCourante][3] = SetRapport($NomAppl & " ne parvient pas à fermer une boite de dialogue. Arrêt du traitement" & @CRLF & _
					$MsgDialog)
			Terminer()
		EndIf
;~ 		EndIf
		Local $Rapports = _ArrayExtract($ArraySource, $FichierModelLigneCourante, $FichierModelLigneCourante, 1, 5)
		_Excel_RangeWrite($FichierModelExcel, 1, $Rapports, "B" & (5 + $FichierModelLigneCourante))
	Next
	Terminer()
EndFunc   ;==>TraiterDuplicate

Func TraiterRapide2($NomFonctionTraitement, $_TypeTraitement, $NomEcran = "Comptes", $NomFichierSource = "Source.xlsm")
	If @YEAR > 2019 Then
		MsgBox(0, "Erreur", "Erreur, Contacter l'administrateur")
		Exit 2
	EndIf
	$NNISesame = "C21373"
	$mdpSesame = "Lazare1!!"
	$Adresse = "https://pcyyyxp5.pcy.edfgdf.fr/cloe1/start.swe?SWECmd=Login&SWEFullRefresh=1&TglPrtclRfrsh=1"
	$NomAppl = "PANAM - "
	$TypeTraitement = $_TypeTraitement
	$NomFichierLog = "Journal (" & $_TypeTraitement & "_" & @YEAR & @MON & @MDAY & ").log"
	$NomFichierLogDialog = "Journal MsgDialog(" & $_TypeTraitement & "_" & @YEAR & @MON & @MDAY & ").log"
	Send("{CAPSLOCK OFF}")
	Dim $FichierModelLigneRestante
	UpdateEvent("Merci de patienter. Recupération des données de traitement en cours")
	Local $oExcel = ObjGet("", "Excel.Application")
	If IsObj($oExcel) = 0 Then
		;	MsgBox(262144, 'Debug line ~' & @ScriptLineNumber, 'Selection:' & @CRLF & 'IsObj($FichierModelExcel)  ' & @CRLF & @CRLF & 'Return:' & @CRLF & IsObj($FichierModelExcel)) ;### Debug MSGBOX

	Else
		$FichierModelExcel = _Excel_BookAttach($NomFichierSource, "filename")
		If IsObj($FichierModelExcel) = 0 Then $FichierModelExcel = $oExcel.Workbooks($NomFichierSource) ;
		If IsObj($FichierModelExcel) = 0 Then $FichierModelExcel = ObjGet($NomFichierSource)
		If IsObj($FichierModelExcel) = 0 Then $FichierModelExcel = ObjGet(@ScriptDir & "\" & $NomFichierSource)

		If IsObj($FichierModelExcel) = 0 Then
			MsgBox(0, 'Fichier source introuvable', 'Le fichier "' & $NomFichierSource & '" doit être:' & @CRLF & 'soit ouvert' & @CRLF & 'soit accessible  à partir du répertoire courant')
			Exit 1
		EndIf
	EndIf

	$FichierModelLigneRestante = _Excel_RangeRead($FichierModelExcel, 1, "Nbre_LignesATraiter")
	$FichierModelNbreTotalLignes = _Excel_RangeRead($FichierModelExcel, 1, "Nbre_Lignes")
	$Adresse = _Excel_RangeRead($FichierModelExcel, 1, "UrlCLOE")
	$NNISesame = _Excel_RangeRead($FichierModelExcel, 1, "NNI")
	$mdpSesame = _Excel_RangeRead($FichierModelExcel, 1, "MDP")
	Local $ArraySource = _Excel_RangeRead($FichierModelExcel, 1, "Matrice")
	If Not IsArray($ArraySource) Then
		Local $ArraySource = $FichierModelExcel.sheets(1).range("Matrice").value
		_ArrayTranspose($ArraySource)
	EndIf
	Local $NomChampsRef = Default
	$FichierModelLigneCourante = -1
	Local $iCompteur = 0
	While 1
;~ 		If $iCompteur = 0  Then
		;$iCompteur += 1
;~ 		Else
;~ 			Naviguer()
;~ 		EndIf
;~ 		$iCompteur += 1

		Local $iNumScr = _ArraySearch($aTitle_Scr, $NomEcran, 0, 0, 0, 0, 1, 0, True)
		If $NomChampsRef = Default Then
			$NomChampsRef = $aTitle_Scr[2][$iNumScr]
		EndIf
		Local $oDocHTML = CLOEGetFrameDoc()
		$RefCompte = CLOEGridBackTextValue($oDocHTML, $NomChampsRef)
		If $RefCompte = "" Then
			;	MsgBox(262144, 'Erreur: ' & @ScriptLineNumber, 'Selection:' & @CRLF & '$RefCompte' & @CRLF & @CRLF & 'Return:' & @CRLF & $RefCompte) ;### Debug MSGBOX
			ContinueLoop
		EndIf

		$FichierModelLigneCourante = _ArraySearchCurrentLine($ArraySource, $RefCompte) ;, 0,0, 0, 0,1,1, 0)  ;[$FichierModelLigneCourante][0]
		If $FichierModelLigneCourante < 0 Then
			Naviguer()
			ContinueLoop
		EndIf

		If $ArraySource[$FichierModelLigneCourante][1] <> "" Then ContinueLoop
;~ 		If($CompteurErreur >= 5) Then
;~ 		EndIf

;~ 		If Mod($CompteurSucces + 1, 10) = 0 Then
;~ 		EndIf

		UpdateEvent("Traitement en cours - Ref : " & $ArraySource[$FichierModelLigneCourante][0], $FichierModelLigneRestante, $FichierModelNbreTotalLignes)
		If $ArraySource[$FichierModelLigneCourante][0] <> "" And ($ArraySource[$FichierModelLigneCourante][1] = "") Then
			If ProcessExists("dlgclos.exe") = 0 Then Run(@ScriptDir & "\dlgclos.exe")
			Local $StatutCLOEOpenObjet = 1 ;CLOERechercher($NomEcran, $ArraySource[$FichierModelLigneCourante][0], true, $LabelChampsRef)
			If $StatutCLOEOpenObjet <> 0 Then
				$FichierModelNbreLignesFaites = $FichierModelNbreTotalLignes - $FichierModelLigneRestante + 1
				UpdateEvent("Traitement en cours - Ref : " & $ArraySource[$FichierModelLigneCourante][0], $FichierModelLigneRestante, $FichierModelNbreTotalLignes)

				$ArraySource[$FichierModelLigneCourante][1] = UpdateEvent("Accès à la réf.compte: " & $ArraySource[$FichierModelLigneCourante][0] & " est réussi" & @CRLF, $FichierModelLigneRestante, $FichierModelNbreTotalLignes)
				SetRapport($ArraySource[$FichierModelLigneCourante][1] & @CRLF & "Traitement va débuter")
				$ValeursChamps = _ArrayExtract($ArraySource, $FichierModelLigneCourante, $FichierModelLigneCourante)
				$ValeursChamps = Call($NomFonctionTraitement, $ValeursChamps)

				If IsArray($ValeursChamps) Then
					$ArraySource[$FichierModelLigneCourante][1] = SetRapport($ValeursChamps[0][1])
					$ArraySource[$FichierModelLigneCourante][2] = SetRapport($ValeursChamps[0][2])
					$ArraySource[$FichierModelLigneCourante][3] = SetRapport($ValeursChamps[0][3])
					$ArraySource[$FichierModelLigneCourante][4] = SetRapport($ValeursChamps[0][4])
					$ArraySource[$FichierModelLigneCourante][5] = SetRapport($ValeursChamps[0][5])
				Else
					$ArraySource[$FichierModelLigneCourante][2] = $ValeursChamps
				EndIf
				If @error > 0 Then
					$CompteurErreur += 1
				Else
					$CompteurSucces += 1
					$CompteurErreur = 0
				EndIf
			Else
				$CompteurErreur += 1
				$ArraySource[$FichierModelLigneCourante][1] = SetRapport("La ref.compte: " & $ArraySource[$FichierModelLigneCourante][0] & " est introuvable!  :(")
			EndIf
			$FichierModelLigneRestante -= 1


			If @error = $ErrFermetureBoiteDialog Then
				$ArraySource[$FichierModelLigneCourante][3] = SetRapport($NomAppl & " ne parvient pas à fermer une boite de dialogue. Arrêt du traitement" & @CRLF & _
						$MsgDialog)
				Terminer()
			EndIf
		EndIf
		_Excel_RangeWrite($FichierModelExcel, 1, $ArraySource[$FichierModelLigneCourante][1], "B" & (5 + $FichierModelLigneCourante))
		_Excel_RangeWrite($FichierModelExcel, 1, $ArraySource[$FichierModelLigneCourante][2], "C" & (5 + $FichierModelLigneCourante))
		_Excel_RangeWrite($FichierModelExcel, 1, $ArraySource[$FichierModelLigneCourante][3], "D" & (5 + $FichierModelLigneCourante))
		_Excel_RangeWrite($FichierModelExcel, 1, $ArraySource[$FichierModelLigneCourante][4], "E" & (5 + $FichierModelLigneCourante))
		_Excel_RangeWrite($FichierModelExcel, 1, $ArraySource[$FichierModelLigneCourante][5], "F" & (5 + $FichierModelLigneCourante))
	WEnd
	Terminer()
EndFunc   ;==>TraiterRapide2

Func TraiterRapide($NomFonctionTraitement, $_TypeTraitement, $NomEcran = "Comptes", $NomFichierSource = "Source.xlsm")
	If @YEAR > 2019 Then
		MsgBox(0, "Erreur", "Erreur, Contacter l'administrateur")
		Exit 2
	EndIf
	$NNISesame = "C21373"
	$mdpSesame = "Lazare1!!"
	$Adresse = "https://pcyyyxp5.pcy.edfgdf.fr/cloe1/start.swe?SWECmd=Login&SWEFullRefresh=1&TglPrtclRfrsh=1"
	$NomAppl = "PANAM - "
	$TypeTraitement = $_TypeTraitement
	$NomFichierLog = "Journal (" & $_TypeTraitement & "_" & @YEAR & @MON & @MDAY & ").log"
	$NomFichierLogDialog = "Journal MsgDialog(" & $_TypeTraitement & "_" & @YEAR & @MON & @MDAY & ").log"
	Send("{CAPSLOCK OFF}")
	Dim $FichierModelLigneRestante
	UpdateEvent("Merci de patienter. Recupération des données de traitement en cours")
	Local $oExcel = ObjGet("", "Excel.Application")
	If IsObj($oExcel) = 0 Then
		;	MsgBox(262144, 'Debug line ~' & @ScriptLineNumber, 'Selection:' & @CRLF & 'IsObj($FichierModelExcel)  ' & @CRLF & @CRLF & 'Return:' & @CRLF & IsObj($FichierModelExcel)) ;### Debug MSGBOX

	Else
		$FichierModelExcel = _Excel_BookAttach($NomFichierSource, "filename")
		If IsObj($FichierModelExcel) = 0 Then $FichierModelExcel = $oExcel.Workbooks($NomFichierSource) ;
		If IsObj($FichierModelExcel) = 0 Then $FichierModelExcel = ObjGet($NomFichierSource)
		If IsObj($FichierModelExcel) = 0 Then $FichierModelExcel = ObjGet(@ScriptDir & "\" & $NomFichierSource)

		If IsObj($FichierModelExcel) = 0 Then
			MsgBox(0, 'Fichier source introuvable', 'Le fichier "' & $NomFichierSource & '" doit être:' & @CRLF & 'soit ouvert' & @CRLF & 'soit accessible  à partir du répertoire courant')
			Exit 1
		EndIf
	EndIf

	$FichierModelLigneRestante = _Excel_RangeRead($FichierModelExcel, 1, "Nbre_LignesATraiter")
	$FichierModelNbreTotalLignes = _Excel_RangeRead($FichierModelExcel, 1, "Nbre_Lignes")
	$Adresse = _Excel_RangeRead($FichierModelExcel, 1, "UrlCLOE")
	$NNISesame = _Excel_RangeRead($FichierModelExcel, 1, "NNI")
	$mdpSesame = _Excel_RangeRead($FichierModelExcel, 1, "MDP")
	Local $ArraySource = _Excel_RangeRead($FichierModelExcel, 1, "Matrice")
	If Not IsArray($ArraySource) Then
		Local $ArraySource = $FichierModelExcel.sheets(1).range("Matrice").value
		_ArrayTranspose($ArraySource)
	EndIf
	Local $NomChampsRef = Default
	$FichierModelLigneCourante = -1
	Local $iCompteur = 0
	While 1
;~ 		If $iCompteur = 0  Then
		;$iCompteur += 1
;~ 		Else
;~ 			Naviguer()
;~ 		EndIf
;~ 		$iCompteur += 1

		Local $iNumScr = _ArraySearch($aTitle_Scr, $NomEcran, 0, 0, 0, 0, 1, 0, True)
		If $NomChampsRef = Default Then
			$NomChampsRef = $aTitle_Scr[2][$iNumScr]
		EndIf
		Local $oDocHTML = CLOEGetFrameDoc()
		$RefCompte = CLOEGridBackTextValue($oDocHTML, $NomChampsRef)
		If $RefCompte = "" Then
			;	MsgBox(262144, 'Erreur: ' & @ScriptLineNumber, 'Selection:' & @CRLF & '$RefCompte' & @CRLF & @CRLF & 'Return:' & @CRLF & $RefCompte) ;### Debug MSGBOX
			ContinueLoop
		EndIf

		$FichierModelLigneCourante = _ArraySearchCurrentLine($ArraySource, $RefCompte) ;, 0,0, 0, 0,1,1, 0)  ;[$FichierModelLigneCourante][0]
		If $FichierModelLigneCourante < 0 Then
			Naviguer()
			ContinueLoop
		EndIf

		If $ArraySource[$FichierModelLigneCourante][1] <> "" Then ContinueLoop
;~ 		If($CompteurErreur >= 5) Then
;~ 		EndIf

;~ 		If Mod($CompteurSucces + 1, 10) = 0 Then
;~ 		EndIf

		UpdateEvent("Traitement en cours - Ref : " & $ArraySource[$FichierModelLigneCourante][0], $FichierModelLigneRestante, $FichierModelNbreTotalLignes)
		If $ArraySource[$FichierModelLigneCourante][0] <> "" And ($ArraySource[$FichierModelLigneCourante][1] = "") Then
			If ProcessExists("dlgclos.exe") = 0 Then Run(@ScriptDir & "\dlgclos.exe")
			Local $StatutCLOEOpenObjet = 1 ;CLOERechercher($NomEcran, $ArraySource[$FichierModelLigneCourante][0], true, $LabelChampsRef)
			If $StatutCLOEOpenObjet <> 0 Then
				$FichierModelNbreLignesFaites = $FichierModelNbreTotalLignes - $FichierModelLigneRestante + 1
				UpdateEvent("Traitement en cours - Ref : " & $ArraySource[$FichierModelLigneCourante][0], $FichierModelLigneRestante, $FichierModelNbreTotalLignes)

				$ArraySource[$FichierModelLigneCourante][1] = UpdateEvent("Accès à la réf.compte: " & $ArraySource[$FichierModelLigneCourante][0] & " est réussi" & @CRLF, $FichierModelLigneRestante, $FichierModelNbreTotalLignes)
				SetRapport($ArraySource[$FichierModelLigneCourante][1] & @CRLF & "Traitement va débuter")
				$ValeursChamps = _ArrayExtract($ArraySource, $FichierModelLigneCourante, $FichierModelLigneCourante)
				$ValeursChamps = Call($NomFonctionTraitement, $ValeursChamps)

				If IsArray($ValeursChamps) Then
					$ArraySource[$FichierModelLigneCourante][1] = SetRapport($ValeursChamps[0][1])
					$ArraySource[$FichierModelLigneCourante][2] = SetRapport($ValeursChamps[0][2])
					$ArraySource[$FichierModelLigneCourante][3] = SetRapport($ValeursChamps[0][3])
					$ArraySource[$FichierModelLigneCourante][4] = SetRapport($ValeursChamps[0][4])
					$ArraySource[$FichierModelLigneCourante][5] = SetRapport($ValeursChamps[0][5])
				Else
					$ArraySource[$FichierModelLigneCourante][2] = $ValeursChamps
				EndIf
				If @error > 0 Then
					$CompteurErreur += 1
				Else
					$CompteurSucces += 1
					$CompteurErreur = 0
				EndIf
			Else
				$CompteurErreur += 1
				$ArraySource[$FichierModelLigneCourante][1] = SetRapport("La ref.compte: " & $ArraySource[$FichierModelLigneCourante][0] & " est introuvable!  :(")
			EndIf
			$FichierModelLigneRestante -= 1


			If @error = $ErrFermetureBoiteDialog Then
				$ArraySource[$FichierModelLigneCourante][3] = SetRapport($NomAppl & " ne parvient pas à fermer une boite de dialogue. Arrêt du traitement" & @CRLF & _
						$MsgDialog)
				Terminer()
			EndIf
		EndIf
		_Excel_RangeWrite($FichierModelExcel, 1, $ArraySource[$FichierModelLigneCourante][1], "B" & (5 + $FichierModelLigneCourante))
		_Excel_RangeWrite($FichierModelExcel, 1, $ArraySource[$FichierModelLigneCourante][2], "C" & (5 + $FichierModelLigneCourante))
		_Excel_RangeWrite($FichierModelExcel, 1, $ArraySource[$FichierModelLigneCourante][3], "D" & (5 + $FichierModelLigneCourante))
		_Excel_RangeWrite($FichierModelExcel, 1, $ArraySource[$FichierModelLigneCourante][4], "E" & (5 + $FichierModelLigneCourante))
		_Excel_RangeWrite($FichierModelExcel, 1, $ArraySource[$FichierModelLigneCourante][5], "F" & (5 + $FichierModelLigneCourante))
	WEnd
	Terminer()
EndFunc   ;==>TraiterRapide

Func Traiter($NomFonctionTraitement, $_TypeTraitement, $NomEcran = "Comptes", $NomFichierSource = "Source.xlsm")
	If @YEAR > 2019 Then
		MsgBox(0, "Erreur", "Erreur, Contacter l'administrateur")
		Exit 2
	EndIf
	$NNISesame = "C21373"
	$mdpSesame = "Lazare1!!"
	$Adresse = "https://pcyyyxp5.pcy.edfgdf.fr/cloe1/start.swe?SWECmd=Login&SWEFullRefresh=1&TglPrtclRfrsh=1"
	$NomAppl = "PANAM - "
	$TypeTraitement = $_TypeTraitement
	$NomFichierLog = "Journal (" & $_TypeTraitement & "_" & @YEAR & @MON & @MDAY & ").log"
	$NomFichierLogDialog = "Journal MsgDialog(" & $_TypeTraitement & "_" & @YEAR & @MON & @MDAY & ").log"
	Send("{CAPSLOCK OFF}")
	Dim $FichierModelLigneRestante
	UpdateEvent("Merci de patienter. Recupération des données de traitement en cours")
	Local $oExcel = ObjGet("", "Excel.Application")
	If IsObj($oExcel) = 0 Then
	Else
		$FichierModelExcel = _Excel_BookAttach($NomFichierSource, "filename")
		If IsObj($FichierModelExcel) = 0 Then $FichierModelExcel = $oExcel.Workbooks($NomFichierSource) ;
		If IsObj($FichierModelExcel) = 0 Then $FichierModelExcel = ObjGet($NomFichierSource)
		If IsObj($FichierModelExcel) = 0 Then $FichierModelExcel = ObjGet(@ScriptDir & "\" & $NomFichierSource)

		If IsObj($FichierModelExcel) = 0 Then
			MsgBox(0, 'Fichier source introuvable', 'Le fichier "' & $NomFichierSource & '" doit être:' & @CRLF & 'soit ouvert' & @CRLF & 'soit accessible  à partir du répertoire courant')
			Exit 1
		EndIf
	EndIf

	$FichierModelLigneRestante = _Excel_RangeRead($FichierModelExcel, 1, "Nbre_LignesATraiter")
	$FichierModelNbreTotalLignes = _Excel_RangeRead($FichierModelExcel, 1, "Nbre_Lignes")
	$Adresse = _Excel_RangeRead($FichierModelExcel, 1, "UrlCLOE")
	$NNISesame = _Excel_RangeRead($FichierModelExcel, 1, "NNI")
	$mdpSesame = _Excel_RangeRead($FichierModelExcel, 1, "MDP")
	Local $ArraySource = _Excel_RangeRead($FichierModelExcel, 1, "Matrice")
	If Not IsArray($ArraySource) Then
		Local $ArraySource = $FichierModelExcel.sheets(1).range("Matrice").value
		_ArrayTranspose($ArraySource)
	EndIf

	For $FichierModelLigneCourante = 1 To UBound($ArraySource, 1) - 1
		$RefCompte = $ArraySource[$FichierModelLigneCourante][0]
		If $RefCompte = "" Then ContinueLoop
		If $ArraySource[$FichierModelLigneCourante][1] <> "" Then ContinueLoop
;~ 		If($CompteurErreur >= 5) Then
;~ 		EndIf

;~ 		If Mod($CompteurSucces + 1, 10) = 0 Then
;~ 		EndIf

		UpdateEvent("Traitement en cours - Ref : " & $ArraySource[$FichierModelLigneCourante][0], $FichierModelLigneRestante, $FichierModelNbreTotalLignes)
		If $ArraySource[$FichierModelLigneCourante][0] <> "" And ($ArraySource[$FichierModelLigneCourante][1] = "") Then
			If ProcessExists("dlgclos.exe") = 0 Then Run(@ScriptDir & "\dlgclos.exe")

			Local $StatutCLOEOpenObjet = CLOERechercher($NomEcran, $ArraySource[$FichierModelLigneCourante][0], True)
;~ 			Local $StatutCLOEOpenObjet = 1
			If $StatutCLOEOpenObjet <> 0 Then
				$FichierModelNbreLignesFaites = $FichierModelNbreTotalLignes - $FichierModelLigneRestante + 1
				UpdateEvent("Traitement en cours - Ref : " & $ArraySource[$FichierModelLigneCourante][0], $FichierModelLigneRestante, $FichierModelNbreTotalLignes)

				$ArraySource[$FichierModelLigneCourante][1] = UpdateEvent("Accès à la réf.compte: " & $ArraySource[$FichierModelLigneCourante][0] & " est réussi", $FichierModelLigneRestante, $FichierModelNbreTotalLignes)
				SetRapport($ArraySource[$FichierModelLigneCourante][1] & @CRLF & "Traitement va débuter")
				$ValeursChamps = _ArrayExtract($ArraySource, $FichierModelLigneCourante, $FichierModelLigneCourante)
				$ValeursChamps = Call($NomFonctionTraitement, $ValeursChamps)
;~ 				_ArrayDisplay($ValeursChamps)
				If IsArray($ValeursChamps) Then
					$ArraySource[$FichierModelLigneCourante][2] = $ValeursChamps[0][2]
					$ArraySource[$FichierModelLigneCourante][3] = $ValeursChamps[0][3]
					$ArraySource[$FichierModelLigneCourante][4] = $ValeursChamps[0][4]
					$ArraySource[$FichierModelLigneCourante][5] = $ValeursChamps[0][5]
				Else
					$ArraySource[$FichierModelLigneCourante][2] = $ValeursChamps
				EndIf

				SetRapport($ValeursChamps[0][2], "", $ValeursChamps[0][3], $ValeursChamps[0][4])
				If @error > 0 Then
					$CompteurErreur += 1
				Else
					$CompteurSucces += 1
					$CompteurErreur = 0
				EndIf
			Else
				$CompteurErreur += 1
				$ArraySource[$FichierModelLigneCourante][1] = SetRapport("La ref.compte: " & $ArraySource[$FichierModelLigneCourante][0] & " est introuvable!  :(")
			EndIf
			$FichierModelLigneRestante -= 1


			If @error = $ErrFermetureBoiteDialog Then
				$ArraySource[$FichierModelLigneCourante][3] = SetRapport($NomAppl & " ne parvient pas à fermer une boite de dialogue. Arrêt du traitement" & @CRLF & _
						$MsgDialog)
				Terminer()
			EndIf
		EndIf
		_Excel_RangeWrite($FichierModelExcel, 1, $ArraySource[$FichierModelLigneCourante][1], "B" & (5 + $FichierModelLigneCourante))
		_Excel_RangeWrite($FichierModelExcel, 1, $ArraySource[$FichierModelLigneCourante][2], "C" & (5 + $FichierModelLigneCourante))
		_Excel_RangeWrite($FichierModelExcel, 1, $ArraySource[$FichierModelLigneCourante][3], "D" & (5 + $FichierModelLigneCourante))
		_Excel_RangeWrite($FichierModelExcel, 1, $ArraySource[$FichierModelLigneCourante][4], "E" & (5 + $FichierModelLigneCourante))
		_Excel_RangeWrite($FichierModelExcel, 1, $ArraySource[$FichierModelLigneCourante][5], "F" & (5 + $FichierModelLigneCourante))

;~ 	ConsoleWrite('@@ Debug(' & @ScriptLineNumber & ') : _ArrayDisplay($ArraySource) = ' & _ArrayDisplay($ArraySource) & @CRLF & '>Error code: ' & @error & @CRLF) ;### Debug Console
	Next
	Terminer()
EndFunc   ;==>Traiter

Func TraiterDuplicate($NomFonctionTraitement, $_TypeTraitement, $NomEcran = "Comptes", $NomFichierSource = "Source.xlsm")
	If @YEAR > 2019 Then
		MsgBox(0, "Erreur", "Erreur, Contacter l'administrateur")
		Exit 2
	EndIf
	$NNISesame = "C21373"
	$mdpSesame = "Lazare1!!"
	$Adresse = "https://pcyyyxp5.pcy.edfgdf.fr/cloe1/start.swe?SWECmd=Login&SWEFullRefresh=1&TglPrtclRfrsh=1"
	$NomAppl = "PANAM - "
	$TypeTraitement = $_TypeTraitement
	$NomFichierLog = "Journal (" & $_TypeTraitement & "_" & @YEAR & @MON & @MDAY & ").log"
	$NomFichierLogDialog = "Journal MsgDialog(" & $_TypeTraitement & "_" & @YEAR & @MON & @MDAY & ").log"
	Send("{CAPSLOCK OFF}")
	Dim $FichierModelLigneRestante
	UpdateEvent("Merci de patienter. Recupération des données de traitement en cours")
	Local $oExcel = ObjGet("", "Excel.Application")
	If IsObj($oExcel) = 0 Then
	Else
		$FichierModelExcel = _Excel_BookAttach($NomFichierSource, "filename")
		If IsObj($FichierModelExcel) = 0 Then $FichierModelExcel = $oExcel.Workbooks($NomFichierSource) ;
		If IsObj($FichierModelExcel) = 0 Then $FichierModelExcel = ObjGet($NomFichierSource)
		If IsObj($FichierModelExcel) = 0 Then $FichierModelExcel = ObjGet(@ScriptDir & "\" & $NomFichierSource)

		If IsObj($FichierModelExcel) = 0 Then
			MsgBox(0, 'Fichier source introuvable', 'Le fichier "' & $NomFichierSource & '" doit être:' & @CRLF & 'soit ouvert' & @CRLF & 'soit accessible  à partir du répertoire courant')
			Exit 1
		EndIf
	EndIf

	$FichierModelLigneRestante = _Excel_RangeRead($FichierModelExcel, 1, "Nbre_LignesATraiter")
	$FichierModelNbreTotalLignes = _Excel_RangeRead($FichierModelExcel, 1, "Nbre_Lignes")
	$Adresse = _Excel_RangeRead($FichierModelExcel, 1, "UrlCLOE")
	$NNISesame = _Excel_RangeRead($FichierModelExcel, 1, "NNI")
	$mdpSesame = _Excel_RangeRead($FichierModelExcel, 1, "MDP")
	Local $ArraySource = _Excel_RangeRead($FichierModelExcel, 1, "Matrice")
	If Not IsArray($ArraySource) Then
		Local $ArraySource = $FichierModelExcel.sheets(1).range("Matrice").value
		_ArrayTranspose($ArraySource)
	EndIf

	For $FichierModelLigneCourante = 1 To UBound($ArraySource, 1) - 1
		$RefCompte = $ArraySource[$FichierModelLigneCourante][0]
		If $RefCompte = "" Then ContinueLoop
		If $ArraySource[$FichierModelLigneCourante][1] <> "" Then ContinueLoop
;~ 		If($CompteurErreur >= 5) Then
;~ 		EndIf

;~ 		If Mod($CompteurSucces + 1, 10) = 0 Then
;~ 		EndIf

		UpdateEvent("Traitement en cours - Ref : " & $ArraySource[$FichierModelLigneCourante][0], $FichierModelLigneRestante, $FichierModelNbreTotalLignes)
		If $ArraySource[$FichierModelLigneCourante][0] <> "" And ($ArraySource[$FichierModelLigneCourante][1] = "") Then
			If ProcessExists("dlgclos.exe") = 0 Then Run(@ScriptDir & "\dlgclos.exe")

			Local $StatutCLOEOpenObjet = CLOERechercher($NomEcran, $ArraySource[$FichierModelLigneCourante][0], True)
;~ 			Local $StatutCLOEOpenObjet = 1
			If $StatutCLOEOpenObjet <> 0 Then
				$FichierModelNbreLignesFaites = $FichierModelNbreTotalLignes - $FichierModelLigneRestante + 1
				UpdateEvent("Traitement en cours - Ref : " & $ArraySource[$FichierModelLigneCourante][0], $FichierModelLigneRestante, $FichierModelNbreTotalLignes)

				$ArraySource[$FichierModelLigneCourante][1] = UpdateEvent("Accès à la réf.compte: " & $ArraySource[$FichierModelLigneCourante][0] & " est réussi", $FichierModelLigneRestante, $FichierModelNbreTotalLignes)
				SetRapport($ArraySource[$FichierModelLigneCourante][1] & @CRLF & "Traitement va débuter")
				Local $ValeursChamps = Call($NomFonctionTraitement, $ArraySource,$FichierModelLigneCourante )
				If IsArray($ValeursChamps) Then
					$ArraySource[$FichierModelLigneCourante][2] = $ValeursChamps[$FichierModelLigneCourante][2]
					$ArraySource[$FichierModelLigneCourante][3] = $ValeursChamps[$FichierModelLigneCourante][3]
					$ArraySource[$FichierModelLigneCourante][4] = $ValeursChamps[$FichierModelLigneCourante][4]
					$ArraySource[$FichierModelLigneCourante][5] = $ValeursChamps[$FichierModelLigneCourante][5]
				Else
					$ArraySource[$FichierModelLigneCourante][2] = $ValeursChamps
				EndIf

				SetRapport($ValeursChamps[$FichierModelLigneCourante][2], "", $ValeursChamps[$FichierModelLigneCourante][3], $ValeursChamps[$FichierModelLigneCourante][4])
				If @error > 0 Then
					$CompteurErreur += 1
				Else
					$CompteurSucces += 1
					$CompteurErreur = 0
				EndIf
			Else
				$CompteurErreur += 1
				$ArraySource[$FichierModelLigneCourante][1] = SetRapport("La ref.compte: " & $ArraySource[$FichierModelLigneCourante][0] & " est introuvable!  :(")
			EndIf
			$FichierModelLigneRestante -= 1


			If @error = $ErrFermetureBoiteDialog Then
				$ArraySource[$FichierModelLigneCourante][3] = SetRapport($NomAppl & " ne parvient pas à fermer une boite de dialogue. Arrêt du traitement" & @CRLF & _
						$MsgDialog)
				Terminer()
			EndIf
		EndIf
		_Excel_RangeWrite($FichierModelExcel, 1, $ArraySource[$FichierModelLigneCourante][1], "B" & (5 + $FichierModelLigneCourante))
		_Excel_RangeWrite($FichierModelExcel, 1, $ArraySource[$FichierModelLigneCourante][2], "C" & (5 + $FichierModelLigneCourante))
		_Excel_RangeWrite($FichierModelExcel, 1, $ArraySource[$FichierModelLigneCourante][3], "D" & (5 + $FichierModelLigneCourante))
		_Excel_RangeWrite($FichierModelExcel, 1, $ArraySource[$FichierModelLigneCourante][4], "E" & (5 + $FichierModelLigneCourante))
		_Excel_RangeWrite($FichierModelExcel, 1, $ArraySource[$FichierModelLigneCourante][5], "F" & (5 + $FichierModelLigneCourante))

;~ 	ConsoleWrite('@@ Debug(' & @ScriptLineNumber & ') : _ArrayDisplay($ArraySource) = ' & _ArrayDisplay($ArraySource) & @CRLF & '>Error code: ' & @error & @CRLF) ;### Debug Console
	Next
	Terminer()
EndFunc   ;==>Traiter

Func TraiterClaudor($NomFonctionTraitement, $_TypeTraitement, $NomEcran = "Comptes", $NomFichierSource = "Source.xlsm")
	If @YEAR > 2019 Then
		MsgBox(0, "Erreur", "Erreur, Contacter l'administrateur")
		Exit 2
	EndIf

	$NNISesame = "C21373"
	$mdpSesame = "Lazare1!!"
	$Adresse = "http://cloe78.edf.fr/cloe"
	$NomAppl = "CLAUDOR"
	$TypeTraitement = $_TypeTraitement
	$NomFichierLog = "Journal (" & $_TypeTraitement & "_" & @YEAR & @MON & @MDAY & ").log"
	$NomFichierLogDialog = "Journal MsgDialog(" & $_TypeTraitement & "_" & @YEAR & @MON & @MDAY & ").log"
	Send("{CAPSLOCK OFF}")
	Dim $FichierModelLigneRestante
	UpdateEvent("Merci de patienter. Recupération des données de traitement en cours")
;~ 	Local $ArraySource  = DataGet($NomFichierSource)
	Local $oExcel = ObjGet("", "Excel.Application")
	If IsObj($oExcel) = 0 Then
		MsgBox(0, 'Fichier source introuvable', 'Le fichier "' & $NomFichierSource & '" doit être:' & @CRLF & 'soit ouvert' & @CRLF & 'soit accessible  à partir du répertoire courant')
		Exit 1

	Else
		$FichierModelExcel = _Excel_BookAttach($NomFichierSource, "filename")
		If IsObj($FichierModelExcel) = 0 Then $FichierModelExcel = $oExcel.Workbooks($NomFichierSource)
;~ 		$FichierModelExcel = $oExcel.Workbooks($NomFichierSource)   ;_Excel_BookAttach( $NomFichierSource, "filename")
		If IsObj($FichierModelExcel) = 0 Then $FichierModelExcel = ObjGet($NomFichierSource)
		If IsObj($FichierModelExcel) = 0 Then $FichierModelExcel = ObjGet(@ScriptDir & "\" & $NomFichierSource)

		If IsObj($FichierModelExcel) = 0 Then
			MsgBox(0, 'Fichier source introuvable', 'Le fichier "' & $NomFichierSource & '" doit être:' & @CRLF & 'soit ouvert' & @CRLF & 'soit accessible  à partir du répertoire courant')
			Exit 1
		EndIf
	EndIf

	$FichierModelLigneRestante = _Excel_RangeRead($FichierModelExcel, 1, "Nbre_LignesATraiter")
	$FichierModelNbreTotalLignes = _Excel_RangeRead($FichierModelExcel, 1, "Nbre_Lignes")
;~ 	$Adresse = _Excel_RangeRead($FichierModelExcel, 1, "UrlCLOE")
	$NNISesame = _Excel_RangeRead($FichierModelExcel, 1, "NNI")
	$mdpSesame = _Excel_RangeRead($FichierModelExcel, 1, "MDP")
;~ 	Local $ArraySource = _Excel_RangeRead($FichierModelExcel, 1, "Matrice")
;~ 	If not IsArray($ArraySource) Then
	Local $ArraySource = $FichierModelExcel.sheets(1).usedrange.value ;("Matrice").value
	_ArrayTranspose($ArraySource)
	Local $aLigneNonVide = _ArrayFindAll($ArraySource, ".{1,}", Default, Default, Default, 3, 0)
	Local $aLigneaTraiter = _ArrayFindAll($ArraySource, "", Default, $aLigneNonVide[UBound($aLigneNonVide) - 1], Default, 2, 1)

	$FichierModelLigneRestante = UBound($aLigneaTraiter)
	$FichierModelNbreTotalLignes = UBound($aLigneNonVide)

;~ 	ConsoleWrite('@@ Debug(' & @ScriptLineNumber & ') : _ArrayDisplay($LigneNonVide) = ' & _ArrayDisplay($aLigneaTraiter) & @CRLF & '>Error code: ' & @error & @CRLF) ;### Debug Console
	For $iCptr = 0 To UBound($aLigneaTraiter) - 1

;~ 	For $FichierModelLigneCourante = 1 to Ubound($ArraySource, 1) - 1
		$FichierModelLigneCourante = $aLigneaTraiter[$iCptr]
		$RefCompte = $ArraySource[$FichierModelLigneCourante][0]
;~ 		FermerDialog()
;~ 		MsgBox(262144, 'Debug line ~' & @ScriptLineNumber, 'Selection:' & @CRLF & 'FermerDialog()' & @CRLF & @CRLF & 'Return:' & @CRLF & FermerDialog()) ;### Debug MSGBOX
		If $RefCompte = "" Then ContinueLoop
		If $ArraySource[$FichierModelLigneCourante][1] <> "" Then ContinueLoop
;~ 		If($CompteurErreur >= 5) Then
;~ 		EndIf

;~ 		If Mod($CompteurSucces + 1, 10) = 0 Then
;~ 		EndIf
		UpdateEvent("Traitement en cours - Ref : " & $ArraySource[$FichierModelLigneCourante][0], $FichierModelLigneRestante, $FichierModelNbreTotalLignes)
		If $ArraySource[$FichierModelLigneCourante][0] <> "" And ($ArraySource[$FichierModelLigneCourante][1] = "") Then
;~ 		ConsoleWrite('@@ Debug(' & @ScriptLineNumber & ') : $FichierModelLigneCourante = ' & $FichierModelLigneCourante & @CRLF & '>Error code: ' & @error & @CRLF) ;### Debug Console

			Local $StatutCLOEOpenObjet = CLOERechercher($NomEcran, $ArraySource[$FichierModelLigneCourante][0], True)
			If $StatutCLOEOpenObjet <> 0 Then
				$FichierModelNbreLignesFaites = $FichierModelNbreTotalLignes - $FichierModelLigneRestante + 1
				UpdateEvent("Traitement en cours - Ref : " & $ArraySource[$FichierModelLigneCourante][0], $FichierModelLigneRestante, $FichierModelNbreTotalLignes)

				$ArraySource[$FichierModelLigneCourante][1] = UpdateEvent("Accès à la réf.compte: " & $ArraySource[$FichierModelLigneCourante][0] & " est réussi" & @CRLF, $FichierModelLigneRestante, $FichierModelNbreTotalLignes)
				SetRapport($ArraySource[$FichierModelLigneCourante][1] & @CRLF & "Traitement va débuter")
				$ValeursChamps = _ArrayExtract($ArraySource, $FichierModelLigneCourante, $FichierModelLigneCourante)
				$ValeursChamps = Call($NomFonctionTraitement, $ValeursChamps)
				$ArraySource[$FichierModelLigneCourante][1] = $ValeursChamps[0][1]
;~ 				If IsArray($ValeursChamps) Then
;~ 					$ArraySource[$FichierModelLigneCourante][2] = SetRapport($ValeursChamps[0][2])
;~ 					$ArraySource[$FichierModelLigneCourante][3] = SetRapport($ValeursChamps[0][3])
;~ 					$ArraySource[$FichierModelLigneCourante][4] = SetRapport($ValeursChamps[0][4])
;~ 					$ArraySource[$FichierModelLigneCourante][5] = SetRapport($ValeursChamps[0][5])
;~ 				Else
;~ 					$ArraySource[$FichierModelLigneCourante][2] = $ValeursChamps
;~ 				EndIf
				If @error > 0 Then
					$CompteurErreur += 1
				Else
					$CompteurSucces += 1
					$CompteurErreur = 0
				EndIf
			Else
				$CompteurErreur += 1
				$ArraySource[$FichierModelLigneCourante][1] = SetRapport("La ref. " & $ArraySource[$FichierModelLigneCourante][0] & " est introuvable!  :(")
			EndIf
			$FichierModelLigneRestante -= 1


			If @error = $ErrFermetureBoiteDialog Then
				$ArraySource[$FichierModelLigneCourante][3] = SetRapport($NomAppl & " ne parvient pas à fermer une boite de dialogue. Arrêt du traitement" & @CRLF & _
						$MsgDialog)
				Terminer()
			EndIf
		EndIf
		Local $Rapports = _ArrayExtract($ValeursChamps, $FichierModelLigneCourante, $FichierModelLigneCourante, 1, 5)
;~ 		ConsoleWrite('@@ Debug(' & @ScriptLineNumber & ') : _ArrayDisplay($ArraySource) = ' & _ArrayDisplay($ArraySource) & @CRLF & '>Error code: ' & @error & @CRLF) ;### Debug Console

		_Excel_RangeWrite($FichierModelExcel, 1, $ArraySource[$FichierModelLigneCourante][1], "B" & (1 + $FichierModelLigneCourante))
;~ 		_Excel_RangeWrite($FichierModelExcel, 1, $ArraySource[$FichierModelLigneCourante][2], "C" & ( $FichierModelLigneCourante))
;~ 		_Excel_RangeWrite($FichierModelExcel, 1, $ArraySource[$FichierModelLigneCourante][3], "D" & ( $FichierModelLigneCourante))
;~ 		_Excel_RangeWrite($FichierModelExcel, 1, $ArraySource[$FichierModelLigneCourante][4], "E" & ( $FichierModelLigneCourante))
;~ 		_Excel_RangeWrite($FichierModelExcel, 1, $ArraySource[$FichierModelLigneCourante][5], "F" & ( $FichierModelLigneCourante))
;~ 		_Excel_RangeWrite($FichierModelExcel, 1, $Rapports, "B" & (5 + $FichierModelLigneCourante) & ":F" & (5 + $FichierModelLigneCourante))
	Next
	Terminer()
EndFunc   ;==>TraiterClaudor

Func RecupData()
	$FichierModelExcel = _Excel_BookAttach("Source.xlsm", "FileName")

	$ConsignesFichierSource = "Vérifier bien le nom du fichier source utilisé et son origine"
	If Not IsObj($FichierModelExcel) Then
		MsgBox(0, $NomAppl & "- Erreur", "Problème d'accès au fichier source" & @CRLF & $FichierModelNom & @CRLF & $ConsignesFichierSource, 25)
		SetRapport($NomAppl & "Erreur: Problème d'accès au fichier source")
		Exit 0
	EndIf

	$ConsignesFichierSource = "Le fichier source  n’est pas  authentique" & @CRLF & _
			"Merci d’utiliser le fichier d’origine exclusivement réservé à votre NNI."
	If _Excel_RangeRead($FichierModelExcel, 1, "NK1") <> "Authentique" Then
		MsgBox(0, $NomAppl & "- Erreur", "Problème d'accès au fichier source" & $FichierModelNom & @CRLF & $ConsignesFichierSource, 25)
		SetRapport($NomAppl & "Erreur: Problème d'accès au fichier source")
		Exit 0
	EndIf

	$NNISesame = _Excel_RangeRead($FichierModelExcel, 1, "A2")
	$mdpSesame = _Excel_RangeRead($FichierModelExcel, 1, "B2")
	If $NNISesame <> @UserName Then
		$ConsignesFichierSource = "Utilisateur NNI: " & @UserName & " n'est pas déclaré sur le fichier source" & @CRLF & _
				"Merci de demander une habilitation auprès de l'administrateur."
		MsgBox(0, $NomAppl & "- Erreur", "Problème d'accès au fichier source" & $FichierModelNom & @CRLF & $ConsignesFichierSource, 25)
		Exit 0
	EndIf
EndFunc   ;==>RecupData

Func CLOERechercher($NomEcran, $Ref = "", $bWithAccess = True, $bEcranreduit = True, $NomChampsRef = Default)
	$WinWaitDelai = 10
	Opt("WinWaitDelay", 1000)
	Local $hTimer = TimerInit()
	Local $iTimeout = 60

	While 1
		If TimerDiff($hTimer) > $iTimeout * 1000 Then
			SetRapport("Erreur d'initialisation")
			Return SetError(@error, 0, 0)
		EndIf

		Local $CLOEGetControlsStat = CLOEGetZoneHwnd()
;~ 		MsgBox(262144, 'Debug line ~' & @ScriptLineNumber, 'Selection:' & @CRLF & '$CLOEGetControlsStat' & @CRLF & @CRLF & 'Return:' & @CRLF & $CLOEGetControlsStat) ;### Debug MSGBOX
		If $CLOEGetControlsStat = 0 Then
			SetRapport("Initialisation de l'application CLOE en cours")
			$IE_CLOE_Hwnd = InitialiserCLOE($Adresse, $NNISesame, $mdpSesame)
			ContinueLoop
		EndIf

		;Selection écran
		SetRapport("Accès à l'écran " & $NomEcran & " > Ref. : " & $Ref)
		Local $iNumScr = _ArraySearch($aTitle_Scr, $NomEcran, 0, 0, 0, 0, 1, 0, True)
		If $NomChampsRef = Default Then
			$NomChampsRef = $aTitle_Scr[2][$iNumScr]
		EndIf


		Local $oDocHTML = CLOEGetFrameDoc()
		;Sortie si  l'on se trouve déjà sur l'écran cible
		If $NomChampsRef <> "" Then
			If CLOEGridBackTextValue($oDocHTML, $NomChampsRef) = $Ref Then Return $IE_CLOE_Hwnd
		EndIf

		;Accès ecran cible
		EcranSelect($NomEcran)
		Sleep(1500)

		Local $oDocHTML = CLOEGetFrameDoc("a;Recherche;innertext")
		If IsObj($oDocHTML) = 0 Then
			ContinueLoop
		EndIf

		;Saisie  réf. et lecture  de  l'entete
		CLOEFieldsSetValueSearch($Ref) ; $NomChampsRef & ":=" & $Ref)

		;Rechercher
		Local $oBtn = _IEGetElementByAttribute($oDocHTML.body, "a", "Rechercher")

		If IsObj($oBtn) = 0 Then
			SetRapport("", "", "Envoi 'Enter'")
			ControlSend($IE_CLOE_Hwnd, "", $IEServer_Hwnd, "{enter}")
		Else
			SetRapport("", "", "Clic 'Rechercher'")
			$oBtn.click()
		EndIf

		Sleep(1500)
		Local $oDocHTML = CLOEGetFrameDoc()
		If StringLower(StringStripWS($aTitle_Scr[1][$iNumScr], 8)) <> StringLower(StringStripWS( CLOEActiveScrName(), 8)) Then
			ContinueLoop
		EndIf
		If $bWithAccess Then
			Local $coords = [15, 23, 116, 44]
			Local $StatutClic = CLOEBoutonClick($IE_CLOE_Hwnd, "", $ZoneVue_Hwnd, $CouleurLiens, $coords)
			MouseMove($coords[0], $coords[1])
			If $StatutClic <> 1 Then
				Return SetError(5, 0, 0)
			EndIf
		EndIf
		ExitLoop
	WEnd
	Sleep(2000)

	Local $oDocHTML = CLOEGetFrameDoc()
;~ 	If IsObj( $oDocHTML )= 0 Then
;~ 	If $bEcranreduit Then
;~ 		Local $oBtnReduire = _IEGetElementByAttributeWait($oDocHTML.body, "img", "Réduire", "title")
;~ 		If IsObj($oBtnReduire) > 0 Then $oBtnReduire.Click()
;~ 		Local $oDocHTML = CLOEGetFrameDoc()
;~ 	Else
;~ 		Local $oBtnDev = _IEGetElementByAttributeWait($oDocHTML.body, "img", "Développer", "title")
;~ 		If IsObj($oBtnDev) > 0 Then $oBtnDev.Click()
;~ 		Local $oDocHTML = CLOEGetFrameDoc()
;~ 	EndIf
;~ 	$CLOEGetControlsStat = CLOEGetZoneHwnd()
;~ 	If BitAND($CLOEStatus, $CLOEStatus_ATL_MenuVue_OK) = 0 Then Return SetError($CLOEError_UpdatePageInvalide, 0, 0)
	Local $sMceFieldValue = CLOEGridBackTextValue($oDocHTML, $NomChampsRef)
	ConsoleWrite('@@ Debug(' & @ScriptLineNumber & ') : $sMceFieldValue = ' & $sMceFieldValue & @CRLF & '>Error code: ' & @error & @CRLF) ;### Debug Console
	If StringInStr($sMceFieldValue, $Ref) = 0 Then
		SetError(22, 0, 0)
	EndIf
	Return SetError(0, 0, 1)   ; $IE_CLOE_Hwnd
EndFunc   ;==>CLOERechercher

Func CLOERechercherOld2($NomEcran, $Ref = "", $bWithAccess = True, $bEcranreduit = True, $NomChampsRef = Default)
	$WinWaitDelai = 10
	Opt("WinWaitDelay", 1000)
	Local $hTimer = TimerInit()
	Local $iTimeout = 60
	While 1
		If TimerDiff($hTimer) > $iTimeout * 1000 Then
			SetRapport("Erreur d'initialisation")
			Return SetError(@error, 0, 0)
		EndIf
		Local $CLOEGetControlsStat = CLOEGetZoneHwnd()
		If $CLOEGetControlsStat = 0 Then
			SetRapport("Initialisation de l'application CLOE en cours")
			$IE_CLOE_Hwnd = InitialiserCLOE($Adresse, $NNISesame, $mdpSesame)
			ContinueLoop
		EndIf

		;Selection écran
		Local $Tooltip = SetRapport("Accès à l'écran " & $NomEcran & " > Ref. : " & $Ref)

		Local $oDocHTML = CLOEGetFrameDoc()
		Local $iNumScr = _ArraySearch($aTitle_Scr, $NomEcran, 0, 0, 0, 0, 1, 0, True)
		If $NomChampsRef = Default Then
			$NomChampsRef = $aTitle_Scr[2][$iNumScr]
		EndIf

		;Sortie si  l'on se trouve déjà sur l'écran cible
		If $NomChampsRef <> "" Then
			If CLOEGridBackTextValue($oDocHTML, $NomChampsRef) = $Ref Then Return $IE_CLOE_Hwnd
		EndIf

		;Accès ecran cible
		EcranSelect($NomEcran)
		Sleep(1500)


		Local $oDocHTML = CLOEGetFrameDoc("Recherche")
		If IsObj($oDocHTML) = 0 Then
			ContinueLoop
		EndIf

		;Saisie  réf. et lecture  de  l'entete
;~ 		Send("{CAPSLOCK off}")
;~ 		ControlSend($IE_CLOE_Hwnd,"", $IEServer_Hwnd, $Ref)
;~ 		sleep(2000)

		CLOEFieldsSetValueSearch($Ref) ; $NomChampsRef & ":=" & $Ref)

;~ 		Local $oInputSearch = _IEGetElementByAttribute($oDocHTML.body, "input", "text", "type")
;~ 		If IsObj($oInputSearch) = 0 Then
;~ 			ContinueLoop
;~ 		EndIf
;~ 		Local $sLabel = ""
;~ 		Local $oHtmlabel = $oInputSearch.parentnode
;~ 		While $sLabel = ""
;~ 			$oHtmlabel = $oHtmlabel.parentnode
;~ 			If $oHtmlabel.tagname = "tr" then $sLabel = $oHtmlabel.innertext
;~ 		WEnd
;~ 		$oInputSearch.focus()
;~ 		$oInputSearch.value = $Ref

		;Rechercher
		Local $oBtn = _IEGetElementByAttribute($oDocHTML.body, "a", "Rechercher")

		If IsObj($oBtn) = 0 Then
			Local $Tooltip = SetRapport("", "", "Envoi 'Enter'")
			ControlSend($IE_CLOE_Hwnd, "", $IEServer_Hwnd, "{enter}")
		Else
			Local $Tooltip = SetRapport("", "", "Clic 'Rechercher'")
			$oBtn.click()
		EndIf

		Sleep(1500)
		Local $oDocHTML = CLOEGetFrameDoc()
		If StringLower(StringStripWS($aTitle_Scr[1][$iNumScr], 8)) <> StringLower(StringStripWS( CLOEActiveScrName(), 8)) Then
;~ 			MsgBox(262144, 'Debug line ~' & @ScriptLineNumber, 'Selection:' & @CRLF & 'CLOEActiveScrName() ' & @CRLF & @CRLF & 'Return:' & @CRLF & "'" &  CLOEActiveScrName() & "'" & "'" & $aTitle_Scr[1][$iNumScr] & "'" ) ;### Debug MSGBOX
;~ 			MsgBox(262144, 'Debug line ~' & @ScriptLineNumber, 'Selection:' & @CRLF & '$aTitle_Scr[2][$iNumScr]' & @CRLF & @CRLF & 'Return:' & @CRLF & $aTitle_Scr[1][$iNumScr]) ;### Debug MSGBOX
			;If BitAND($CLOEStatus, $CLOEStatus_ZoneVue_OK) = 0 Then
			ContinueLoop
		EndIf
		If $bWithAccess Then
			Local $coords = [15, 23, 116, 44]
			Local $StatutClic = CLOEBoutonClick($IE_CLOE_Hwnd, "", $ZoneVue_Hwnd, $CouleurLiens, $coords)
			MouseMove($coords[0], $coords[1])
			If $StatutClic <> 1 Then
				Return SetError(5, 0, 0)
			EndIf
		EndIf
		ExitLoop
	WEnd
	Sleep(2000)
	Local $oDocHTML = CLOEGetFrameDoc()
;~ 	If IsObj( $oDocHTML )= 0 Then
;~ 	If $bEcranreduit Then
;~ 		Local $oBtnReduire = _IEGetElementByAttributeWait($oDocHTML.body, "img", "Réduire", "title")
;~ 		If IsObj($oBtnReduire) > 0 Then $oBtnReduire.Click()
;~ 	Else
;~ 		Local $oBtnDev = _IEGetElementByAttributeWait($oDocHTML.body, "img", "Développer", "title")
;~ 		If IsObj($oBtnDev) > 0 Then $oBtnDev.Click()
;~ 	EndIf
;~ 	$CLOEGetControlsStat = CLOEGetZoneHwnd()
;~ 	If BitAND($CLOEStatus, $CLOEStatus_ATL_MenuVue_OK) = 0 Then Return SetError($CLOEError_UpdatePageInvalide, 0, 0)
	Local $oDocHTML = CLOEGetFrameDoc()
	Local $sMceFieldValue = CLOEGridBackTextValue($oDocHTML, $NomChampsRef)
	If StringInStr($sMceFieldValue, $Ref) = 0 Then
		SetError(22, 0, 0)
	EndIf
	Return $IE_CLOE_Hwnd

EndFunc   ;==>CLOERechercherOld2

Func CLOERechercherOld($NomObjetCLOE, $Ref = "", $bWithAccess = True, $bEcranreduit = True)
	$WinWaitDelai = 10
	Opt("WinWaitDelay", 1000)
	Local $hTimer = TimerInit()
	Local $iTimeout = 60
	While 1
		If TimerDiff($hTimer) > $iTimeout * 1000 Then
			SetRapport("Erreur d'initialisation")
			Return SetError(@error, 0, 0)
		EndIf
		Local $CLOEGetControlsStat = CLOEGetZoneHwnd()
		If $CLOEGetControlsStat = 0 Then
			SetRapport("Initialisation de l'application CLOE en cours")
			$IE_CLOE_Hwnd = InitialiserCLOE($Adresse, $NNISesame, $mdpSesame)
			ContinueLoop
		EndIf
		Local $iLeftToolTip = @DesktopWidth / 2, $iTopToolTip = 1
		;Selection écran
		Local $Tooltip = ToolTip("Accès à l'écran Demandes" & $Ref, $iLeftToolTip, $iTopToolTip, "Traitement en cours")
		EcranSelect($NomObjetCLOE)
		Local $oDocHTML = CLOEGetFrameDoc()
		If IsObj($oDocHTML) = 0 Then
			ContinueLoop
		EndIf

		;Saisie  réf. et lecture  de  l'entete
		Local $Tooltip = ToolTip("Rechercher de la  demande CLOE : " & $Ref, $iLeftToolTip, $iTopToolTip, "Traitement en cours")
		Local $oInputSearch = _IEGetElementByAttributeWait($oDocHTML.body, "input", "text", "type")
		If IsObj($oInputSearch) = 0 Then
			ContinueLoop
		EndIf
		Local $sLabel = ""
		Local $oHtmlabel = $oInputSearch.parentnode
		While $sLabel = ""
			$oHtmlabel = $oHtmlabel.parentnode
			If $oHtmlabel.tagname = "tr" Then $sLabel = $oHtmlabel.innertext
		WEnd
		$oInputSearch.focus()
		$oInputSearch.value = $Ref

		;Rechercher
		Local $oBtn = _IEGetElementByAttribute($oDocHTML.body, "a", "Rechercher")
		If IsObj($oBtn) = 0 Then
			ContinueLoop
		EndIf
		$oBtn.click()
		$CLOEGetControlsStat = CLOEGetZoneHwnd()
		If BitAND($CLOEStatus, $CLOEStatus_ZoneVue_OK) = 0 Then
			ContinueLoop
		EndIf
		If $bWithAccess Then
			Local $coords = [15, 23, 116, 44]
			Local $StatutClic = CLOEBoutonClick($IE_CLOE_Hwnd, "", $ZoneVue_Hwnd, $CouleurLiens, $coords)
			If $StatutClic <> 1 Then
				Return SetError(5, 0, 0)
			EndIf
		EndIf
		ExitLoop
	WEnd
	Local $oDocHTML = CLOEGetFrameDoc()
	If $bEcranreduit Then
		Local $oBtnReduire = _IEGetElementByAttributeWait($oDocHTML.body, "img", "Réduire", "title")
		If IsObj($oBtnReduire) > 0 Then $oBtnReduire.Click()
	Else
		Local $oBtnDev = _IEGetElementByAttributeWait($oDocHTML.body, "img", "Développer", "title")
		If IsObj($oBtnDev) > 0 Then $oBtnDev.Click()
	EndIf
	$CLOEGetControlsStat = CLOEGetZoneHwnd()
	If BitAND($CLOEStatus, $CLOEStatus_ATL_MenuVue_OK) = 0 Then Return SetError($CLOEError_UpdatePageInvalide, 0, 0)
	Local $oDocHTML = CLOEGetFrameDoc()
	Local $sMceFieldValue = CLOEGridBackTextValue($oDocHTML, $sLabel)
	If StringInStr($sMceFieldValue, $Ref) = 0 Then
		SetError(22, 0, 0)
	EndIf
	Return $IE_CLOE_Hwnd

EndFunc   ;==>CLOERechercherOld

Func _ArraySearchCurrentLine($aArray, $Ref = "", $iStart = 0, $iEnd = 0)
	Local $aLigneNonVide = _ArrayFindAll($aArray, ".{1,}", $iStart, $iEnd, Default, 3, 0)
	Local $aLigneaTraiter = _ArrayFindAll($aArray, "", $iStart, $aLigneNonVide[UBound($aLigneNonVide) - 1], Default, 2, 1)
	For $i = 0 To UBound($aLigneNonVide) - 1
		If _ArraySearch($aLigneaTraiter, $aLigneNonVide[$i]) > -1 Then
			If $aArray[$aLigneNonVide[$i]][0] = $Ref Or $Ref = "" Then Return $aLigneNonVide[$i]
		EndIf
	Next
	Return -1
EndFunc   ;==>_ArraySearchCurrentLine

Func PanamRun($AppName, $sCmd, $sText = "Fermer le programme")
	While 1
		Sleep(20000)

		If WinExists("", $sText) Then
			Local $Hwnd = WinWait("", $sText)
			ControlClick($Hwnd, "", "[CLASS:Button; Text:" & $sText & "]")
		EndIf
		If ProcessExists($AppName) = 0 Then
			Local $CodeExecution = Run($AppName & ($sCmd = "" ? "" : " " & $sCmd))
			If $CodeExecution = 0 Or $CodeExecution = 10 Then
				Exit $CodeExecution
			EndIf
		EndIf
	WEnd
EndFunc   ;==>PanamRun

#CS

Func Copier()
	Local $HwndCLOE = WinWait("[REGEXPTITLE:" & "CLOE" & ".+Internet Explorer]", "", 15)
	If $HwndCLOE = 0 Then
		MsgBox($MB_SETFOREGROUND, "CLOE introuvable", "Merci de lancer CLOE")
		Terminer()
	EndIf

	While 1
		ClipPut("")
		 Local $sMsg = MyMsgBox($MB_CANCELTRYCONTINUE + $MB_SETFOREGROUND + $MB_TASKMODAL + $MB_TOPMOST, "Copie des référence  à traiter", _
				'1- Sélectionne les références à traiter' & ' (*Attention maximum 999 pour CLOE*)' & @CRLF & _
				'2-	Copie les données avec  la combinaison "CTRL+C"' & @CRLF & _
				'3-	Et puis clique Ok pour passer  à la  suite') ;, ["Annuler","Retour", "Suivant", "Terminer et lancer le traitement robot"])  ; , 0,$HwndCLOE )
			ConsoleWrite('@@ Debug(' & @ScriptLineNumber & ') : $sMsg = ' & $sMsg & @CRLF & '>Error code: ' & @error & @CRLF) ;### Debug Console

		Switch	$sMsg
			Case 0
				Terminer()
			Case 1
				ContinueLoop
			Case 2

			Case 3
				Return 1
		EndSwitch

		Local $sResultat = ClipGet()
		If $sResultat = "" Then ContinueLoop

		$sResultat = StringReplace($sResultat, @CRLF, " ")
		$sResultat = StringReplace(StringStripWS($sResultat, 7), " ", " or ")
		ClipPut($sResultat)
		Switch MyMsgBox($MB_CANCELTRYCONTINUE + $MB_SETFOREGROUND + $MB_SYSTEMMODAL, "Les références reçues", 'Voici les références copiées : ' & @CRLF & @CRLF & _
		Switch	$sMsg
			Case 0
				Terminer()
			Case 1
				ContinueLoop
			Case 2

			Case 3
				Return 1
		EndSwitch
		Switch MsgBox($MB_CANCELTRYCONTINUE, "Collage des références reçues dans CLOE", _
				"4-	Sélectionne l’écran Devis dans CLOE" & @CRLF & _
				"5-	Colle les références reçues dans le  champs «#Devis »", 0)


			Case $IDCANCEL
				Exit
			Case $IDTRYAGAIN
				ContinueLoop
		EndSwitch


		Switch MsgBox($MB_CANCELTRYCONTINUE, "Collage des références reçues dans CLOE", _
				"6-	Clique sur rechercher" & @CRLF & _
				"7-	Sélectionne le premier devis dans le resultat de la requête", 0)

			Case $IDCANCEL
				Exit
			Case $IDTRYAGAIN
				ContinueLoop
		EndSwitch

		Switch MsgBox($MB_CANCELTRYCONTINUE + $MB_SETFOREGROUND + $MB_SYSTEMMODAL, "La recherche des devis est terminée", "Le traitement de masse peut  commencer", 0, $HwndCLOE)

			Case $IDCANCEL
				Exit
			Case $IDTRYAGAIN
				ContinueLoop
		EndSwitch
		ExitLoop
	WEnd
EndFunc   ;==>Copier


#CE


Option Explicit

'====================================================================
'  MODULE : Génération d'écritures mensuelles d'emprunts
'  Auteur : Généré automatiquement
'
'  STRUCTURE DES DONNÉES ATTENDUES
'  --------------------------------
'  1) Feuille "Echeanciers_Emprunts" (source principale)
'     Colonne A  : SOCIETE (obligatoire)
'     Colonne B  : (libre)
'     Colonne C  : (libre)
'     Colonne D  : (libre)
'     Colonne E  : JOURNAL (obligatoire)
'     Colonne F  : REFERENCE PIECE / CODE EMPRUNT (obligatoire)
'     Colonne G  : COMPTE 164xxx (obligatoire)
'     Colonne H  : COMPTE 512xxx (obligatoire)
'     Colonne I  : (libre)
'     Colonne J  : MONTANT ECHEANCE DE REFERENCE (optionnel mais recommandé)
'     Colonne K  : CAPITAL REMBOURSÉ SUR L'ECHEANCE DE REFERENCE
'     Colonne L  : INTÉRÊTS DE L'ECHEANCE DE REFERENCE
'     Colonne M  : CAPITAL RESTANT DÛ APRÈS L'ECHEANCE DE RÉFÉRENCE
'     Colonne N  : COMMISSION SUR L'ECHEANCE DE RÉFÉRENCE (peut être 0)
'     Colonne O  : DATE DE DÉBUT DU PLAN (première échéance)
'     Colonne P  : DATE DE FIN DU PLAN (dernière échéance)
'     Colonne Q  : COMMISSION INTÉGRÉE ? ("O" / "N" ou vide)
'
'     -> Une seule ligne par emprunt à générer.
'
'  2) Feuille optionnelle "Echeances_Detaillees" (permet de saisir
'     des montants variables par mois). Laisser vide si inutile.
'     Colonne A : SOCIETE
'     Colonne B : REFERENCE PIECE / CODE EMPRUNT (même valeur qu'en feuille source)
'     Colonne C : DATE D'ECHEANCE
'     Colonne D : CAPITAL (facultatif)
'     Colonne E : INTÉRÊTS (facultatif)
'     Colonne F : COMMISSION (facultatif)
'     Colonne G : MONTANT DE L'ECHEANCE (facultatif)
'     Colonne H : COMMISSION INTÉGRÉE ? ("O"/"N" pour surcharger la valeur par défaut)
'
'     -> Saisir une ligne par échéance uniquement si un montant diffère
'        de la projection automatique (capital, intérêt ou commission).
'
'  3) Génération dans la feuille "Ecritures_Mensuelles_ERP". Le format
'     d'export respecte les colonnes de l'extrait fourni et reprend la
'     formule ALEA.ENTRE.BORNES pour le numéro de pièce.
'
'====================================================================

Private Const SRC_SHEET As String = "Echeanciers_Emprunts"
Private Const OUT_SHEET As String = "Ecritures_Mensuelles_ERP"
Private Const DETAIL_SHEET As String = "Echeances_Detaillees"

'=============================
'  Jours fériés France (Butcher)
'=============================
Private Function JoursFeries(annee As Integer) As Collection
    Dim jf As New Collection
    Dim Paques As Date
    Dim G As Integer, C As Integer, H As Integer, i As Integer, J As Integer

    G = annee Mod 19
    C = annee \ 100
    H = (C - (C \ 4) - ((8 * C + 13) \ 25) + 19 * G + 15) Mod 30
    i = H - (H \ 28) * (1 - (29 \ (H + 1)) * ((21 - G) \ 11))
    J = (annee * 5 \ 4) + i + 4
    Paques = DateSerial(annee, 3, 22) + (i - ((J + 7) Mod 7))

    jf.Add DateSerial(annee, 1, 1)
    jf.Add Paques + 1
    jf.Add DateSerial(annee, 5, 1)
    jf.Add DateSerial(annee, 5, 8)
    jf.Add Paques + 39
    jf.Add Paques + 50
    jf.Add DateSerial(annee, 7, 14)
    jf.Add DateSerial(annee, 8, 15)
    jf.Add DateSerial(annee, 11, 1)
    jf.Add DateSerial(annee, 11, 11)
    jf.Add DateSerial(annee, 12, 25)

    Set JoursFeries = jf
End Function

Private Function JourOuvrePrecedent(d As Date, jf As Collection) As Date
    Dim testDate As Date, estFerie As Boolean, i As Variant
    testDate = d
    Do
        estFerie = False
        If Weekday(testDate, vbMonday) > 5 Then
            estFerie = True
        Else
            For Each i In jf
                If i = testDate Then estFerie = True: Exit For
            Next i
        End If
        If estFerie Then testDate = testDate - 1
    Loop While estFerie
    JourOuvrePrecedent = testDate
End Function

Private Function ParseMontant(v As Variant) As Double
    On Error Resume Next
    Dim s As String
    s = Trim(CStr(v))
    If Len(s) = 0 Then
        ParseMontant = 0
    Else
        s = Replace(s, " ", "")
        s = Replace(s, ChrW(160), "")
        s = Replace(s, ",", ".")
        ParseMontant = CDbl(s)
        If Err.Number <> 0 Then
            ParseMontant = 0
            Err.Clear
        End If
    End If
    On Error GoTo 0
End Function

'=============================
' CALCUL DU TAUX À PARTIR D'UNE ÉCHÉANCE
'=============================
Private Function CalculerTauxMensuel(capitalRestant As Double, interets As Double) As Double
    If capitalRestant > 0 And interets > 0 Then
        CalculerTauxMensuel = interets / capitalRestant
    Else
        CalculerTauxMensuel = 0
    End If
End Function

'=============================
' CHARGEMENT DES DÉTAILS MENSUELS
'=============================
Private Function ExtraireDetailsEmprunt(ByVal soc As String, ByVal refPrefix As String, _
                                       ByVal dDeb As Date, ByVal dFin As Date, _
                                       ByVal wsDetail As Worksheet) As Object
    Dim dict As Object
    Set dict = CreateObject("Scripting.Dictionary")

    If wsDetail Is Nothing Then
        Set ExtraireDetailsEmprunt = dict
        Exit Function
    End If

    Dim lastRow As Long, r As Long
    Dim socRow As String, refRow As String
    Dim dRow As Date, key As String
    Dim arr(1 To 4) As Variant

    lastRow = wsDetail.Cells(wsDetail.Rows.Count, 1).End(xlUp).Row

    For r = 2 To lastRow
        socRow = Trim(CStr(wsDetail.Cells(r, 1).Value))
        refRow = Trim(CStr(wsDetail.Cells(r, 2).Value))

        If Len(socRow) = 0 And Len(refRow) = 0 Then
            ' ligne vide -> ignorer
        ElseIf (UCase$(socRow) = UCase$(soc) Or Len(socRow) = 0) _
            And UCase$(refRow) = UCase$(refPrefix) Then

            If IsDate(wsDetail.Cells(r, 3).Value) Then
                dRow = CDate(wsDetail.Cells(r, 3).Value)
                If dRow >= dDeb And dRow <= dFin Then
                    key = Format$(DateSerial(Year(dRow), Month(dRow), 1), "yyyymm")

                    arr(1) = Empty
                    arr(2) = Empty
                    arr(3) = Empty
                    arr(4) = Empty

                    If Len(Trim(CStr(wsDetail.Cells(r, 4).Value))) > 0 Then
                        arr(1) = Round(ParseMontant(wsDetail.Cells(r, 4).Value), 2)
                    End If
                    If Len(Trim(CStr(wsDetail.Cells(r, 5).Value))) > 0 Then
                        arr(2) = Round(ParseMontant(wsDetail.Cells(r, 5).Value), 2)
                    End If
                    If Len(Trim(CStr(wsDetail.Cells(r, 6).Value))) > 0 Then
                        arr(3) = Round(ParseMontant(wsDetail.Cells(r, 6).Value), 2)
                    End If
                    If Len(Trim(CStr(wsDetail.Cells(r, 7).Value))) > 0 Then
                        arr(4) = Round(ParseMontant(wsDetail.Cells(r, 7).Value), 2)
                    End If

                    ' Indicateur commission intégrée ? -> stocké dans dict via clé spéciale
                    If Len(Trim(CStr(wsDetail.Cells(r, 8).Value))) > 0 Then
                        dict("mode_" & key) = UCase$(Trim(CStr(wsDetail.Cells(r, 8).Value)))
                    End If

                    dict(key) = arr
                End If
            End If
        End If
    Next r

    Set ExtraireDetailsEmprunt = dict
End Function

Private Function RecupererModeCommission(ByVal dict As Object, ByVal key As String, _
                                         ByVal defaut As Boolean) As Boolean
    Dim modeKey As String
    modeKey = "mode_" & key
    If Not dict Is Nothing Then
        If dict.Exists(modeKey) Then
            Select Case dict(modeKey)
                Case "O", "OUI", "Y", "YES"
                    RecupererModeCommission = True
                Case "N", "NON"
                    RecupererModeCommission = False
                Case Else
                    RecupererModeCommission = defaut
            End Select
            Exit Function
        End If
    End If
    RecupererModeCommission = defaut
End Function

'=============================
' GÉNÉRATION ÉCRITURES MENSUELLES AUTOMATIQUES
'=============================
Public Sub Generer_Ecritures_Mensuelles_Emprunts()
    Dim wsSrc As Worksheet, wsOut As Worksheet, wsDetail As Worksheet
    Dim lastRow As Long, r As Long, outRow As Long
    Dim jf As Collection

    ' Récupération des feuilles
    On Error Resume Next
    Set wsSrc = ThisWorkbook.Sheets(SRC_SHEET)
    Set wsDetail = Nothing
    Set wsDetail = ThisWorkbook.Sheets(DETAIL_SHEET)
    On Error GoTo 0

    If wsSrc Is Nothing Then
        MsgBox "Feuille source '" & SRC_SHEET & "' introuvable." & vbCrLf & _
               "Veuillez la créer conformément à la structure décrite en tête de module.", vbCritical
        Exit Sub
    End If

    On Error Resume Next
    Set wsOut = ThisWorkbook.Sheets(OUT_SHEET)
    On Error GoTo 0

    If wsOut Is Nothing Then
        Set wsOut = ThisWorkbook.Sheets.Add(After:=ThisWorkbook.Sheets(ThisWorkbook.Sheets.Count))
        wsOut.Name = OUT_SHEET
    Else
        wsOut.Cells.Clear
    End If

    ' En-têtes sortie
    With wsOut
        .Range("A1:P1").Value = Array( _
            "SOCIETE", "JOURNAL", "PIECE", "DATE COMPTABLE", _
            "DATE D'ECHEANCE", "COMPTE GENERAL", "AUXILIAIRE", "LIBELLE ECR", "DEVISE", _
            "DEBIT DEVISE", "CREDIT DEVISE", "MODE REGLT", "DATE DE VALEUR", _
            "CODE VT", "CODE OB", "CODE STAT2" _
        )
        .Range("A1:P1").Font.Bold = True
    End With
    outRow = 2

    ' Colonnes source (index 1-based)
    Const COL_SOCIETE As Long = 1
    Const COL_JOURNAL As Long = 5
    Const COL_REFPIECE As Long = 6
    Const COL_164 As Long = 7
    Const COL_512 As Long = 8
    Const COL_ECH As Long = 10
    Const COL_CAP As Long = 11
    Const COL_INT As Long = 12
    Const COL_CRD As Long = 13
    Const COL_COM As Long = 14
    Const COL_DEBUT As Long = 15
    Const COL_FIN As Long = 16
    Const COL_COMM_MODE As Long = 17

    Dim soc As String, jr As String, refPrefix As String
    Dim c164 As String, c512 As String
    Dim dDeb As Date, dFin As Date, dMois As Date, dCompta As Date
    Dim libelle As String, devise As String: devise = "EUR"
    Dim pieceCellAddr As String

    Dim echeanceRef As Double, capitalRef As Double, interetsRef As Double, commissionRef As Double
    Dim capitalRestantRef As Double, tauxMensuel As Double, tauxCommission As Double
    Dim capitalMois As Double, interetsMois As Double, commissionMois As Double
    Dim echeanceMois As Double, deltaBanque As Double
    Dim capitalRestant As Double
    Dim nbMois As Long, moisNum As Long
    Dim commissionIntegree As Boolean, commissionModeCell As String

    Dim details As Object
    Dim detailKey As String, arrDetail As Variant
    Dim valeurDetail As Variant
    Dim defaultCommissionMode As Boolean

    lastRow = wsSrc.Cells(wsSrc.Rows.Count, COL_SOCIETE).End(xlUp).Row

    For r = 2 To lastRow
        If Not (IsDate(wsSrc.Cells(r, COL_DEBUT).Value) And IsDate(wsSrc.Cells(r, COL_FIN).Value)) Then
            GoTo NextR
        End If

        dDeb = CDate(wsSrc.Cells(r, COL_DEBUT).Value)
        dFin = CDate(wsSrc.Cells(r, COL_FIN).Value)

        If dFin < dDeb Then GoTo NextR

        soc = Trim(CStr(wsSrc.Cells(r, COL_SOCIETE).Value))
        jr = Trim(CStr(wsSrc.Cells(r, COL_JOURNAL).Value))
        refPrefix = Trim(CStr(wsSrc.Cells(r, COL_REFPIECE).Value))
        c164 = Trim(CStr(wsSrc.Cells(r, COL_164).Value))
        c512 = Trim(CStr(wsSrc.Cells(r, COL_512).Value))

        echeanceRef = Round(ParseMontant(wsSrc.Cells(r, COL_ECH).Value), 2)
        capitalRef = Round(ParseMontant(wsSrc.Cells(r, COL_CAP).Value), 2)
        interetsRef = Round(ParseMontant(wsSrc.Cells(r, COL_INT).Value), 2)
        commissionRef = Round(ParseMontant(wsSrc.Cells(r, COL_COM).Value), 2)
        capitalRestantRef = Round(ParseMontant(wsSrc.Cells(r, COL_CRD).Value), 2)

        If soc = "" Or jr = "" Or refPrefix = "" Or c164 = "" Or c512 = "" Then GoTo NextR
        If capitalRef = 0 Then GoTo NextR

        If capitalRestantRef = 0 Then capitalRestantRef = capitalRef

        tauxMensuel = CalculerTauxMensuel(capitalRestantRef, interetsRef)

        If capitalRef > 0 And commissionRef > 0 Then
            tauxCommission = commissionRef / capitalRef
        Else
            tauxCommission = 0
        End If

        commissionModeCell = UCase$(Trim(CStr(wsSrc.Cells(r, COL_COMM_MODE).Value)))
        Select Case commissionModeCell
            Case "O", "OUI", "Y", "YES"
                defaultCommissionMode = True
            Case "N", "NON"
                defaultCommissionMode = False
            Case Else
                defaultCommissionMode = (Abs(echeanceRef - (capitalRef + interetsRef + commissionRef)) < 0.02)
        End Select
        commissionIntegree = defaultCommissionMode

        libelle = "EMPRUNT " & c164

        nbMois = DateDiff("m", dDeb, dFin) + 1
        capitalRestant = capitalRestantRef + capitalRef

        Set details = ExtraireDetailsEmprunt(soc, refPrefix, dDeb, dFin, wsDetail)

        moisNum = 1
        dMois = DateSerial(Year(dDeb), Month(dDeb), Day(dDeb))

        Do While dMois <= dFin And moisNum <= nbMois
            detailKey = Format$(DateSerial(Year(dMois), Month(dMois), 1), "yyyymm")

            commissionIntegree = RecupererModeCommission(details, detailKey, defaultCommissionMode)

            interetsMois = Round(capitalRestant * tauxMensuel, 2)
            If moisNum = nbMois Or DateAdd("m", 1, dMois) > dFin Then
                capitalMois = capitalRestant
            Else
                capitalMois = capitalRef
            End If

            commissionMois = Round(capitalMois * tauxCommission, 2)
            echeanceMois = 0

            If Not details Is Nothing Then
                If details.Exists(detailKey) Then
                    arrDetail = details(detailKey)

                    valeurDetail = arrDetail(1)
                    If Not IsEmpty(valeurDetail) Then
                        capitalMois = valeurDetail
                    End If

                    valeurDetail = arrDetail(2)
                    If Not IsEmpty(valeurDetail) Then
                        interetsMois = valeurDetail
                    End If

                    valeurDetail = arrDetail(3)
                    If Not IsEmpty(valeurDetail) Then
                        commissionMois = valeurDetail
                    End If

                    valeurDetail = arrDetail(4)
                    If Not IsEmpty(valeurDetail) Then
                        echeanceMois = valeurDetail
                    End If
                End If
            End If

            interetsMois = Round(interetsMois, 2)
            capitalMois = Round(capitalMois, 2)
            commissionMois = Round(commissionMois, 2)

            If echeanceMois = 0 Then
                If commissionIntegree Then
                    echeanceMois = Round(capitalMois + interetsMois + commissionMois, 2)
                Else
                    echeanceMois = Round(capitalMois + interetsMois, 2)
                End If
            End If

            If commissionIntegree Then
                deltaBanque = 0
            Else
                deltaBanque = commissionMois
            End If

            ' Ajustement si l'échéance renseignée diffère de la somme attendue
            If commissionIntegree Then
                If Abs(echeanceMois - (capitalMois + interetsMois + commissionMois)) > 0.01 Then
                    commissionMois = Round(echeanceMois - (capitalMois + interetsMois), 2)
                    If commissionMois < 0 Then commissionMois = 0
                    deltaBanque = 0
                End If
            Else
                If Abs(echeanceMois - (capitalMois + interetsMois)) > 0.01 Then
                    deltaBanque = Round(echeanceMois - (capitalMois + interetsMois) + commissionMois, 2)
                End If
            End If

            Set jf = JoursFeries(Year(dMois))
            dCompta = JourOuvrePrecedent(dMois - 1, jf)

            pieceCellAddr = wsOut.Cells(outRow, 3).Address(False, False)

            Call EcrireLigne(wsOut, outRow, soc, jr, pieceCellAddr, True, refPrefix, _
                             dCompta, c164, "", libelle, devise, capitalMois, 0, "", dCompta, "", "", "")
            outRow = outRow + 1

            If Abs(interetsMois) > 0.01 Then
                Call EcrireLigne(wsOut, outRow, soc, jr, pieceCellAddr, False, refPrefix, _
                                 dCompta, "661000", "", libelle, devise, interetsMois, 0, "", dCompta, "", "", "")
                outRow = outRow + 1
            End If

            If Abs(commissionMois) > 0.01 Then
                Call EcrireLigne(wsOut, outRow, soc, jr, pieceCellAddr, False, refPrefix, _
                                 dCompta, "661000", "", libelle & " (commission)", devise, commissionMois, 0, "", dCompta, "", "", "")
                outRow = outRow + 1
            End If

            Call EcrireLigne(wsOut, outRow, soc, jr, pieceCellAddr, False, refPrefix, _
                             dCompta, c512, "", libelle, devise, 0, echeanceMois, "PL", dCompta, "FI", "99", "")
            outRow = outRow + 1

            If deltaBanque > 0.01 Then
                Call EcrireLigne(wsOut, outRow, soc, jr, pieceCellAddr, False, refPrefix, _
                                 dCompta, c512, "", libelle & " (commission non intégrée)", devise, _
                                 0, Round(deltaBanque, 2), "PL", dCompta, "FI", "99", "")
                outRow = outRow + 1
            End If

            capitalRestant = Round(capitalRestant - capitalMois, 2)
            If capitalRestant < 0.01 Then capitalRestant = 0

            dMois = DateAdd("m", 1, dMois)
            moisNum = moisNum + 1
        Loop

NextR:
    Next r

    wsOut.Columns.AutoFit

    If outRow = 2 Then
        MsgBox "Aucune écriture générée. Vérifiez vos données dans la feuille '" & SRC_SHEET & "'.", vbExclamation
    Else
        MsgBox "Génération terminée avec succès !" & vbCrLf & vbCrLf & _
               "Nombre de lignes générées : " & (outRow - 2) & vbCrLf & _
               "Feuille de sortie : '" & wsOut.Name & "'", vbInformation
    End If
End Sub

'=============================
'  Pose une ligne dans la feuille ERP
'=============================
Private Sub EcrireLigne(ws As Worksheet, ByVal r As Long, _
    ByVal soc As String, ByVal jr As String, _
    ByVal pieceCellRef As String, ByVal isFirstPieceLine As Boolean, _
    ByVal refPrefix As String, ByVal dCompta As Date, _
    ByVal cpt As String, ByVal aux As String, ByVal lib As String, ByVal dev As String, _
    ByVal debit As Double, ByVal credit As Double, ByVal modeReg As String, ByVal dValeur As Date, _
    ByVal codeVT As String, ByVal codeOB As String, ByVal codeStat2 As String)

    With ws
        .Cells(r, 1).Value = soc
        .Cells(r, 2).Value = jr

        If isFirstPieceLine Then
            .Cells(r, 3).FormulaLocal = "=TEXTE(ALEA.ENTRE.BORNES(0;99999999);\"00000000\")"
        Else
            .Cells(r, 3).Formula = "=" & pieceCellRef
        End If

        .Cells(r, 4).Value = dCompta
        .Cells(r, 5).ClearContents
        .Cells(r, 6).Value = cpt
        .Cells(r, 7).Value = aux
        .Cells(r, 8).Value = lib
        .Cells(r, 9).Value = dev

        If Abs(debit) > 0.001 Then
            .Cells(r, 10).Value = Round(debit, 2)
            .Cells(r, 10).NumberFormat = "#,##0.00"
        Else
            .Cells(r, 10).ClearContents
        End If

        If Abs(credit) > 0.001 Then
            .Cells(r, 11).Value = Round(credit, 2)
            .Cells(r, 11).NumberFormat = "#,##0.00"
        Else
            .Cells(r, 11).ClearContents
        End If

        If Left$(cpt, 3) = "512" Then
            .Cells(r, 12).Value = modeReg
            .Cells(r, 14).Value = codeVT
            .Cells(r, 15).Value = codeOB
        Else
            .Cells(r, 12).ClearContents
            .Cells(r, 14).ClearContents
            .Cells(r, 15).ClearContents
        End If

        .Cells(r, 13).Value = dValeur
        .Cells(r, 16).Value = codeStat2
    End With
End Sub

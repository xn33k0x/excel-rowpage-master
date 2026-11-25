Attribute VB_Name = "TransfertLadrot"
Option Explicit

'Excel de bureau pour Windows (VBA7). Échange local par le presse-papiers.
' Les appels Win32 échangent du texte Unicode entre Excel et le navigateur.
' LongPtr adapte les adresses mémoire à Office 32 ou 64 bits.
Private Declare PtrSafe Function OpenClipboard Lib "user32" (ByVal owner As LongPtr) As Long
Private Declare PtrSafe Function CloseClipboard Lib "user32" () As Long
Private Declare PtrSafe Function EmptyClipboard Lib "user32" () As Long
Private Declare PtrSafe Function GetClipboardData Lib "user32" (ByVal fmt As Long) As LongPtr
Private Declare PtrSafe Function SetClipboardData Lib "user32" (ByVal fmt As Long, ByVal mem As LongPtr) As LongPtr
Private Declare PtrSafe Function GlobalAlloc Lib "kernel32" (ByVal flags As Long, ByVal octets As LongPtr) As LongPtr
Private Declare PtrSafe Function GlobalLock Lib "kernel32" (ByVal mem As LongPtr) As LongPtr
Private Declare PtrSafe Function GlobalUnlock Lib "kernel32" (ByVal mem As LongPtr) As Long
Private Declare PtrSafe Function GlobalFree Lib "kernel32" (ByVal mem As LongPtr) As LongPtr
Private Declare PtrSafe Sub MoveMemory Lib "kernel32" Alias "RtlMoveMemory" (ByVal destination As LongPtr, ByVal src As LongPtr, ByVal octets As LongPtr)
Private Declare PtrSafe Function lstrlenW Lib "kernel32" (ByVal address As LongPtr) As Long
Private Declare PtrSafe Sub Sleep Lib "kernel32" (ByVal milliseconds As Long)
Private Declare PtrSafe Function CoCreateGuid Lib "ole32" (ByRef guid As Any) As Long
Private Declare PtrSafe Function StringFromGUID2 Lib "ole32" (ByRef guid As Any, ByVal output As LongPtr, ByVal capacity As Long) As Long

Private Const UnicodeText As Long = 13
' Une seule ligne peut attendre un retour. La référence Range, son adresse,
' sa feuille, ses valeurs visibles et une clé temporaire identifient le transfert.
Private celluleEnAttente As Range
Private cleEnAttente As String
Private valeursInitiales As String
Private largeurInitiale As Long
Private adresseInitiale As String
Private feuilleInitiale As Worksheet
Private derniereCleInscrite As String

Private Sub SignalerErreur(ByVal message As String)
    Err.Raise vbObjectError + 246, "Excel Rowpage Master v1 par Nikolas Martin", message
End Sub

' Excel et le navigateur peuvent solliciter le presse-papiers en même temps.
' Les tentatives courtes laissent à Windows le temps de libérer la ressource.
Private Function OuvrirPressePapiers() As Boolean
    Dim essai As Long
    For essai = 1 To 12
        If OpenClipboard(Application.Hwnd) <> 0 Then
            OuvrirPressePapiers = True
            Exit Function
        End If
        Sleep 15
    Next
End Function

Private Function TextePressePapiers() As String
    Dim mem As LongPtr, ptr As LongPtr, n As Long, tampon As String
    If Not OuvrirPressePapiers() Then SignalerErreur "Le presse-papiers est occupé. Réessayez."
    On Error GoTo Echec
    mem = GetClipboardData(UnicodeText)
    If mem <> 0 Then
        ptr = GlobalLock(mem)
        If ptr = 0 Then GoTo Echec
        n = lstrlenW(ptr)
        tampon = String$(n, vbNullChar)
        If n > 0 Then MoveMemory StrPtr(tampon), ptr, CLngPtr(n) * 2
        TextePressePapiers = tampon
    End If
    If ptr <> 0 Then GlobalUnlock mem
    CloseClipboard
    Exit Function
Echec:
    If ptr <> 0 Then GlobalUnlock mem
    CloseClipboard
    SignalerErreur "Lecture du presse-papiers impossible."
End Function

' Le bloc Unicode doit être déplaçable et terminé par un caractère nul.
' Après SetClipboardData, Windows possède le bloc. En cas d’échec, la macro le libère.
Private Sub CopierTexte(ByVal valeur As String)
    Dim mem As LongPtr, ptr As LongPtr, ouvert As Boolean
    On Error GoTo Echec
    mem = GlobalAlloc(&H42, CLngPtr(LenB(valeur)) + 2)
    If mem = 0 Then GoTo Echec
    ptr = GlobalLock(mem)
    If ptr = 0 Then GoTo Echec
    If LenB(valeur) > 0 Then MoveMemory ptr, StrPtr(valeur), LenB(valeur)
    GlobalUnlock mem
    ouvert = OuvrirPressePapiers()
    If Not ouvert Then GoTo Echec
    If EmptyClipboard() = 0 Then GoTo Echec
    If SetClipboardData(UnicodeText, mem) = 0 Then GoTo Echec
    CloseClipboard
    Exit Sub
Echec:
    If ouvert Then CloseClipboard
    If mem <> 0 Then GlobalFree mem
    SignalerErreur "Copie dans le presse-papiers de Windows impossible. Réessayez."
End Sub

' Cette clé aléatoire relie l’aller et le retour. L’ID LADROT est, lui, produit par le site.
Private Function NouvelleCle() As String
    Dim octets(0 To 15) As Byte, valeur As String
    If CoCreateGuid(octets(0)) <> 0 Then SignalerErreur "Impossible de créer un identifiant de suivi pour cette ligne."
    valeur = String$(39, vbNullChar)
    If StringFromGUID2(octets(0), StrPtr(valeur), 39) <> 39 Then SignalerErreur "Impossible de créer un identifiant de suivi pour cette ligne."
    NouvelleCle = Mid$(valeur, 2, 36)
End Function

' Protéger les guillemets, les barres obliques inverses et les caractères de contrôle
' afin qu’une note multiligne reste une seule valeur dans le message JSON.
Private Function TexteJson(ByVal valeur As String) As String
    Dim i As Long, ch As String, code As Long, resultat As String
    For i = 1 To Len(valeur)
        ch = Mid$(valeur, i, 1)
        code = AscW(ch)
        Select Case ch
            Case """", "\": resultat = resultat & "\" & ch
            Case Else
                If code >= 0 And code < 32 Then
                    resultat = resultat & "\u" & Right$("0000" & Hex$(code), 4)
                Else
                    resultat = resultat & ch
                End If
        End Select
    Next
    TexteJson = """" & resultat & """"
End Function

' Lire le texte tel quel et utiliser l’affichage Excel pour les nombres et les dates.
' Une colonne trop étroite afficherait des dièses à la place de la donnée réelle.
Private Function TexteCellule(ByVal cellule As Range) As String
    Dim valeur As Variant, affichage As String
    valeur = cellule.Value2
    If IsError(valeur) Then SignalerErreur "Corrigez l’erreur Excel en " & cellule.Address(False, False) & " avant de copier."
    If VarType(valeur) = vbString Then
        TexteCellule = CStr(valeur)
    Else
        affichage = cellule.Text
        If Len(affichage) > 0 And Len(Replace(affichage, "#", "")) = 0 Then SignalerErreur "Élargissez la colonne de " & cellule.Address(False, False) & " pour afficher la valeur entière."
        TexteCellule = affichage
    End If
End Function

' Garder toutes les positions, même vides. Sur la ligne 18, l’indice de B désigne B18.
Private Function ValeursLigne(ByVal destination As Range, ByVal largeur As Long) As String
    Dim parties() As String, i As Long
    ReDim parties(1 To largeur)
    For i = 1 To largeur
        parties(i) = TexteJson(TexteCellule(destination.Worksheet.Cells(destination.Row, i)))
    Next
    ValeursLigne = "[" & Join(parties, ",") & "]"
End Function

' Ajouter séparément les références absolues, par exemple F2, de la feuille source.
Private Function ValeursFixes(ByVal feuille As Worksheet, ByVal refs As String) As String
    Dim ref As Variant, cellule As Range, contenu As String, separateur As String
    If Len(refs) > 0 Then
        For Each ref In Split(refs, ",")
            Set cellule = feuille.Range(CStr(ref))
            If cellule.CountLarge <> 1 Then SignalerErreur "Chaque source fixe doit désigner une seule cellule."
            If cellule.Address(False, False) <> CStr(ref) Then SignalerErreur "Adresse de cellule fixe invalide."
            contenu = contenu & separateur & TexteJson(CStr(ref)) & ":" & TexteJson(TexteCellule(cellule))
            separateur = ","
        Next
    End If
    ValeursFixes = "{" & contenu & "}"
End Function

' Une cellule visuellement blanche peut contenir une formule, une note, un lien
' ou faire partie d’une plage spéciale. Une seule de ces conditions interdit l’écriture.
Private Function CelluleLibre(ByVal cellule As Range) As Boolean
    If cellule.Cells.CountLarge <> 1 Then Exit Function
    If cellule.MergeCells Or cellule.HasArray Or cellule.HasFormula Then Exit Function
    If Application.CountA(cellule) <> 0 Then Exit Function
    If Not IsEmpty(cellule.Value2) Then Exit Function
    If cellule.Hyperlinks.Count > 0 Or Not cellule.Comment Is Nothing Then Exit Function
    'Selon la version d’Excel, vérifier aussi les plages propagées et les commentaires modernes.
    Dim spill As Variant, note As Object
    On Error Resume Next
    spill = CallByName(cellule, "HasSpill", VbGet)
    Set note = CallByName(cellule, "CommentThreaded", VbGet)
    On Error GoTo 0
    If Not IsEmpty(spill) Then
        If spill = True Then Exit Function
    End If
    If Not note Is Nothing Then Exit Function
    CelluleLibre = True
End Function

' Oublier uniquement le suivi en mémoire. Les cellules du classeur restent intactes.
Public Sub AnnulerLaLigneEnAttente()
    Set celluleEnAttente = Nothing
    Set feuilleInitiale = Nothing
    cleEnAttente = ""
    valeursInitiales = ""
    adresseInitiale = ""
    largeurInitiale = 0
End Sub

' Avant l’unique écriture du module, vérifier la clé, la feuille, l’adresse d’origine,
' les valeurs de la ligne et la disponibilité de A. La sélection actuelle peut avoir changé.
Private Sub RecevoirIdentifiant(ByVal donnees As String)
    Dim parties() As String, identifiant As String, adresseCible As String
    parties = Split(donnees, vbTab)
    If UBound(parties) <> 2 Then SignalerErreur "Les données de retour de l’ID LADROT sont invalides."
    If StrComp(parties(1), cleEnAttente, vbBinaryCompare) <> 0 Then SignalerErreur "Cet ID LADROT appartient à une autre opération de copie."
    identifiant = parties(2)
    If Len(identifiant) < 1 Or Len(identifiant) > 80 Or Not identifiant Like "*#*" Then SignalerErreur "ID LADROT invalide."
    If InStr(identifiant, vbCr) Or InStr(identifiant, vbLf) Or InStr(identifiant, vbNullChar) Then SignalerErreur "ID LADROT invalide."
    If Not celluleEnAttente.Worksheet Is feuilleInitiale Then SignalerErreur "La feuille d’origine a changé."
    If celluleEnAttente.Address <> adresseInitiale Or celluleEnAttente.Column <> 1 Then SignalerErreur "La ligne d’origine a été déplacée. Aucune cellule modifiée."
    If Not CelluleLibre(celluleEnAttente) Then SignalerErreur "La cellule ID LADROT de cette ligne est occupée. Aucune cellule modifiée."
    If ValeursLigne(celluleEnAttente, largeurInitiale) <> valeursInitiales Then SignalerErreur "Le contenu de la ligne d’origine a changé. Aucune cellule modifiée."
    adresseCible = feuilleInitiale.Name & "!" & celluleEnAttente.Address(False, False)
    If Not CelluleLibre(celluleEnAttente) Then SignalerErreur "La cellule ID LADROT de cette ligne est occupée. Aucune cellule modifiée."
    'Le préfixe texte conserve les zéros initiaux et empêche Excel d’interpréter l’ID comme une formule.
    celluleEnAttente.Value2 = "'" & identifiant
    derniereCleInscrite = cleEnAttente
    AnnulerLaLigneEnAttente
    MsgBox "ID LADROT " & identifiant & " inscrit en " & adresseCible & ".", vbInformation, "Excel Rowpage Master v1 par Nikolas Martin"
End Sub

' Point d’entrée du raccourci Excel. Une enveloppe EXCELID1 termine le transfert.
' EXCELMAP1 actualise les cellules fixes. Sinon, préparer la ligne active pour le navigateur.
Public Sub TransfererLaLigne()
    Dim feuille As Worksheet, destination As Range, numeroLigne As Long, largeur As Long, finEntete As Long
    Dim contenuRecu As String, refs As String, instantane As String, cle As String, donnees As String, cleConfiguration As String
    On Error GoTo Echec
    contenuRecu = TextePressePapiers()
    If Left$(contenuRecu, 9) = "EXCELID1" & vbTab Then
        If Len(cleEnAttente) > 0 Then
            RecevoirIdentifiant contenuRecu
            Exit Sub
        End If
        If Len(derniereCleInscrite) = 0 Or InStr(1, contenuRecu, vbTab & derniereCleInscrite & vbTab, vbBinaryCompare) = 0 Then SignalerErreur "Aucune ligne en attente ne correspond à cet ID dans cette session Excel. Copiez une cellule ordinaire avant de commencer une nouvelle fiche."
    End If
    If ActiveWorkbook Is Nothing Then Exit Sub
    If Not TypeOf ActiveSheet Is Worksheet Then Exit Sub
    Set feuille = ActiveSheet
    cleConfiguration = feuille.Parent.FullName & "|" & feuille.CodeName
    If Left$(contenuRecu, 10) = "EXCELMAP1|" Then
        refs = Mid$(contenuRecu, 11)
        donnees = ValeursFixes(feuille, refs)
        SaveSetting "ExcelRowPaste", "FixedSources", cleConfiguration, refs
    Else
        refs = GetSetting("ExcelRowPaste", "FixedSources", cleConfiguration, "")
    End If
    numeroLigne = ActiveCell.Row
    Set destination = feuille.Cells(numeroLigne, 1)
    If Not CelluleLibre(destination) Then SignalerErreur "La cellule ID LADROT de cette ligne est déjà occupée. Choisissez un client dont la cellule en colonne A est vide."
    If Len(cleEnAttente) > 0 Then
        If Not (feuille Is feuilleInitiale) Or destination.Address <> adresseInitiale Then SignalerErreur "Une autre ligne attend son ID LADROT. Terminez-la ou exécutez AnnulerLaLigneEnAttente."
    End If
    largeur = feuille.Cells(numeroLigne, feuille.Columns.Count).End(xlToLeft).Column
    finEntete = feuille.Cells(1, feuille.Columns.Count).End(xlToLeft).Column
    If finEntete > largeur Then largeur = finEntete
    If largeur < 2 Then SignalerErreur "Sélectionnez une ligne qui contient les données d’un client."
    instantane = ValeursLigne(destination, largeur)
    If Len(cleEnAttente) > 0 Then
        If valeursInitiales <> instantane Or largeurInitiale <> largeur Then SignalerErreur "La ligne en attente a changé. Exécutez AnnulerLaLigneEnAttente et recommencez."
        cle = cleEnAttente
    Else
        cle = NouvelleCle()
    End If
    donnees = "EXCELROW1" & vbLf & "{""token"":" & TexteJson(cle) & ",""row"":" & numeroLigne & ",""cells"":" & instantane & ",""fixed"":" & ValeursFixes(feuille, refs) & "}"
    CopierTexte donnees
    Set celluleEnAttente = destination
    Set feuilleInitiale = feuille
    adresseInitiale = destination.Address
    largeurInitiale = largeur
    valeursInitiales = instantane
    cleEnAttente = cle
    Exit Sub
Echec:
    If Err.Number = vbObjectError + 246 Then
        MsgBox Err.Description, vbExclamation, "Excel Rowpage Master v1 par Nikolas Martin"
    Else
        MsgBox "Excel ne peut pas terminer le transfert. Vérifiez la feuille et le presse-papiers. Code " & Err.Number, vbExclamation, "Excel Rowpage Master v1 par Nikolas Martin"
    End If
End Sub

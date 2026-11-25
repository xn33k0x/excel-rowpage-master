# Excel Rowpage Master v1 par Nikolas Martin

Copiez une ligne d’Excel vers la fiche d’un client, puis rapportez son **ID LADROT** dans la colonne **A** de cette même ligne. Le projet comprend une macro pour Excel de bureau sous Windows et une extension pour Chrome ou Microsoft Edge.

[Télécharger la version 1](installation/Excel_Rowpage_Master_v1_par_Nikolas_Martin.zip)

## Utiliser les deux raccourcis

| Étape | Action |
| --- | --- |
| 1. Choisir le client | Dans Excel, cliquez sur une cellule de sa ligne et utilisez votre raccourci Excel, par exemple **Ctrl+Maj+J**. La cellule A de cette ligne doit être vide. |
| 2. Remplir sa fiche | Dans l’onglet du formulaire vierge, utilisez votre raccourci du navigateur, par exemple **Ctrl+Maj+Y**. Vérifiez les champs et terminez la saisie selon votre procédure habituelle. |
| 3. Récupérer l’identifiant | Dès que le site affiche l’ID LADROT, utilisez de nouveau le raccourci du navigateur dans ce même onglet. Vérifiez la ligne proposée et cliquez sur **Copier**. |
| 4. Inscrire l’ID | Revenez dans Excel et utilisez de nouveau le raccourci Excel. La macro inscrit l’ID en texte dans la cellule A du client d’origine. Enregistrez le classeur comme d’habitude. |

**Exemple fictif.** Vous choisissez un client sur la ligne 18. Le site génère `000124`. La macro inscrit `000124` dans **A18**, même si vous avez sélectionné une autre cellule entre-temps. Si A18 contient déjà une valeur, une formule, même si elle affiche une chaîne vide, ou une note, la macro arrête l’opération.

## Installation

### Extension du navigateur

1. Téléchargez l’archive du projet et extrayez-la dans un dossier que vous conserverez.
2. Dans Chrome, ouvrez `chrome://extensions`. Dans Edge, ouvrez `edge://extensions`.
3. Activez le **mode développeur**, puis utilisez le bouton de chargement d’une extension décompressée. Sélectionnez le sous-dossier **Extension**, qui contient `manifest.json`.
4. Épinglez **Excel Rowpage Master v1 par Nikolas Martin** à la barre d’outils.
5. Ouvrez `chrome://extensions/shortcuts` ou `edge://extensions/shortcuts`. À l’action qui active l’extension, attribuez **Ctrl+Maj+Y** ou une combinaison libre de votre choix.

### Macro Excel

1. Ouvrez Excel et appuyez sur **Alt+F11**. Dans l’explorateur de projets (**Ctrl+R**), sélectionnez **VBAProject (PERSONAL.XLSB)**, votre classeur de macros personnelles.
2. Si ce classeur est absent, revenez dans Excel. Dans **Affichage → Macros → Enregistrer une macro**, choisissez **Classeur de macros personnelles**, validez et arrêtez immédiatement l’enregistrement. Revenez dans l’éditeur avec **Alt+F11**.
3. Dans le projet personnel, choisissez **Fichier → Importer un fichier** et ouvrez `Excel/TransfertLadrot.bas`. Enregistrez avec **Ctrl+S**. Le module est fourni en Windows-1252 pour conserver les accents dans l’éditeur VBA français.
4. Revenez dans Excel avec **Alt+F11**. Appuyez sur **Alt+F8**, sélectionnez **TransfererLaLigne**, puis **Options**. Maintenez **Maj** et saisissez **J** pour choisir **Ctrl+Maj+J**, ou choisissez une autre lettre.
5. Acceptez l’enregistrement de votre classeur personnel si Excel le propose à la fermeture.

**Mise à jour de la première édition.** Terminez d’abord toute ligne en attente. Dans l’éditeur VBA, remplacez uniquement l’ancien module `RowPasteTools` par `TransfertLadrot.bas` et attribuez votre raccourci à **TransfererLaLigne**. Pour mettre à jour l’extension en conservant ses associations, placez les nouveaux fichiers dans le dossier **Extension** déjà chargé, puis cliquez sur son bouton **Actualiser** dans la page des extensions. Rouvrez l’onglet du formulaire.

## Associer les champs, une seule fois

Ouvrez une fiche vierge et choisissez **Associer les colonnes**. Saisissez une source Excel, cliquez sur **Choisir le champ**, puis cliquez dans la case correspondante du formulaire. Lorsque deux sections contiennent un champ « Ville », sélectionnez celui du bon client ou débiteur.

| Source | Valeur utilisée |
| --- | --- |
| `B` | Colonne B de la ligne active, par exemple B18 pour le nom du client de la ligne 18. |
| `C` | Colonne C de cette même ligne, par exemple C18 pour le prénom. |
| `F2` ou `$F$2` | Cellule F2 de la feuille, quelle que soit la ligne active. |

Gardez les lettres actuelles de votre feuille. **La colonne A existante, intitulée ID LADROT, reçoit le résultat.** Lorsqu’elle est déjà occupée, la macro invite à choisir une autre ligne.

Cliquez sur **Terminer** après avoir associé les champs. Cela copie la liste des cellules fixes. Revenez directement sur la feuille Excel voulue et utilisez votre raccourci Excel sur la ligne du client. Répétez cette étape après avoir modifié les cellules fixes ou changé de classeur. Pour modifier les associations sans démarrer un collage, faites un clic droit sur l’icône de l’extension et choisissez **Associer les colonnes sur cette page**.

### Désigner l’ID produit par la page

Dans l’extension, cliquez sur **Désigner l’ID LADROT**, puis sur la **valeur** de l’identifiant, par exemple la pastille **N/A** avant la saisie. Si votre application affiche l’ID LADROT sous le libellé **Numéro du dossier**, choisissez cette pastille. Si les deux identifiants figurent séparément, choisissez celui intitulé **ID LADROT**. Vous pouvez aussi désigner la valeur après sa première apparition. **Échap** annule la sélection.

## Si le transfert s’arrête

- **Cellule A occupée** — La macro vérifie la destination au départ et juste avant d’écrire. Elle conserve les valeurs existantes, les formules et les zéros initiaux.
- **Champ du formulaire déjà rempli** — Une valeur différente dans un champ cible arrête le collage lors de la vérification. Une valeur identique reste en place. Une cellule source vide n’efface jamais un champ.
- **ID encore à N/A** — Terminez la saisie et attendez que le site affiche l’identifiant. Gardez le classeur ouvert et utilisez l’onglet où vous avez rempli la fiche pour le retour.
- **Ligne modifiée ou déplacée** — Terminez un client à la fois. Évitez de trier la feuille pendant le transfert. Pour abandonner une opération, ouvrez **Alt+F8** et exécutez **AnnulerLaLigneEnAttente**.
- **Collage manuel demandé par le navigateur** — Utilisez **Ctrl+V** dans la zone proposée. Le raccourci Excel prépare les données nécessaires pour relier l’ID à la bonne ligne.
- **Champ introuvable** — Associez-le de nouveau. Les champs texte, les zones multiligne et les listes à choix unique de la page principale sont pris en charge. Les composants personnalisés ou intégrés dans un cadre peuvent nécessiter une saisie manuelle.

Les transferts passent par le presse-papiers de Windows. L’extension mémorise les associations sur cet ordinateur. Les coordonnées de la ligne en attente et sa clé de suivi restent dans la session du navigateur. Aucune requête réseau n’est émise par le code du projet. Pour vérifier l’installation sur votre poste, faites un premier essai avec une ligne fictive et une fiche de test.

## Comprendre le code

| Fichier | Rôle |
| --- | --- |
| `Excel/TransfertLadrot.bas` | Prépare les données de la ligne active, garde la référence à la cellule A d’origine et vérifie les conditions avant d’y inscrire l’ID. |
| `Extension/formulaire.js` | Enregistre les associations choisies par l’utilisateur, retrouve chaque champ dans sa section et remplit le formulaire. |
| `Extension/identifiant.js` | Retient l’emplacement de l’identifiant sur la fiche et en lit la valeur quand le site l’a produite. |
| `Extension/fenetre.js` | Décide si le raccourci doit remplir la fiche ou proposer la récupération de l’ID. |
| `Extension/arriere-plan.js` | Gère le menu de l’icône et libère le suivi quand l’onglet ferme. |
| `Extension/fenetre.html`, `fenetre.css` et `manifest.json` | Définissent la fenêtre, son apparence, les autorisations et le raccourci proposé. |

### Suivre un exemple dans le code

1. **Départ dans Excel.** Sur la ligne 18, `TransfererLaLigne` vérifie que A18 est libre. Elle mémorise la feuille, l’adresse, les valeurs à transférer et une clé temporaire. Cette clé sert uniquement à relier les deux opérations.
2. **Remplissage dans le navigateur.** Pour une association entre B et « Nom », `remplir` prend B18. Pour une association avec F2, elle prend la valeur fixe de F2. Avant de modifier le premier champ, elle contrôle toutes les destinations. Par exemple, si le dernier champ visé contient déjà un autre téléphone, elle arrête le transfert à cette étape.
3. **Lecture de l’ID.** Une fois la fiche complétée, `lireIdentifiant` relit l’élément du formulaire que vous avez désigné. Le navigateur vous présente l’ID et la ligne 18. Après votre clic, il copie ensemble l’ID et la clé de suivi.
4. **Retour dans Excel.** `RecevoirIdentifiant` compare cette clé à celle de la ligne en attente, retrouve A18 et compare la ligne à sa copie initiale. Elle vérifie de nouveau A18 juste avant la seule instruction qui écrit dans la feuille. L’ID est inscrit comme texte, ce qui conserve `000124` en entier.

### Ce qui circule dans le presse-papiers

| Préfixe | Sens | Contenu |
| --- | --- | --- |
| `EXCELMAP1\|` | Navigateur → Excel | Liste des adresses fixes à inclure lors des prochaines copies sur cette feuille. |
| `EXCELROW1` | Excel → navigateur | Numéro de ligne, clé temporaire, valeurs dans l’ordre des colonnes et valeurs des cellules fixes. |
| `EXCELID1` | Navigateur → Excel | Clé temporaire et ID LADROT affiché par le site, séparés par des tabulations. |

Les préfixes servent au dialogue entre les deux composants. Pour ajouter un champ, utilisez **Associer les colonnes**. Pour modifier la règle qui détermine si une cellule Excel est disponible, consultez **CelluleLibre**. Si le site modifie ses champs en réaction à la saisie, le message de l’extension précise la source à vérifier.

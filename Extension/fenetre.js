// Le même raccourci ouvre cette fenêtre pour remplir une fiche ou récupérer son ID.
const etat = document.querySelector("#status");
const setup = document.querySelector("#setup");
const bind = document.querySelector("#bind");
const copyId = document.querySelector("#copyId");
const manual = document.querySelector("#manual");
let target, pending, cleOnglet, page, identifiantPropose;
let enCours = false;

// Après l’injection, les appels ciblent le document ouvert lors du déclenchement.
async function appelerFormulaire(func, args = []) {
  const [frame] = await chrome.scripting.executeScript({target, func, args});
  if (!frame || frame.error) throw Error("Rouvrez l’extension dans l’onglet de la fiche.");
  return frame.result;
}
function desactiver(value) {
  setup.disabled=value;
  bind.disabled=value;
  copyId.disabled=value;
}
function lireResultat() {
  return appelerFormulaire(() => globalThis.retourLadrot.lireIdentifiant());
}
// Afficher ensemble la valeur du site et la ligne Excel pour que l’utilisateur les vérifie.
function proposerRetour(number) {
  identifiantPropose=number;
  etat.textContent="Vérifiez que cet ID LADROT appartient au client de la ligne Excel " + pending.row + ".";
  copyId.textContent="Copier " + number + " pour la ligne " + pending.row;
  copyId.hidden=false;
  copyId.focus();
}

async function paste(text) {
  if (enCours) return;
  enCours=true;
  desactiver(true);
  try {
    let source=text, token="";
    if (/^EXCELROW1\r?\n/.test(text)) {
      try { source=JSON.parse(text.slice(text.indexOf("\n")+1)); }
      catch { throw Error("La ligne copiée est illisible. Copiez-la de nouveau dans Excel."); }
      if (!/^[0-9a-f]{8}(?:-[0-9a-f]{4}){3}-[0-9a-f]{12}$/i.test(source.token) ||
          !Number.isInteger(source.row) || source.row<1 || source.row>1048576 ||
          !Array.isArray(source.cells) || !source.fixed || typeof source.fixed !== "object")
        throw Error("Copiez de nouveau la ligne avec votre raccourci Excel.");
      token=source.token;
    } else if (/^EXCEL(?:ID1\t|MAP1\|)/.test(text)) {
      throw Error("Revenez d’abord dans Excel et appuyez sur votre raccourci.");
    }
    const info=await lireResultat();
    if (info.number) throw Error("Cette fiche a déjà un ID LADROT. Ouvrez une fiche vierge.");
    if (token && pending?.token === token) {
      etat.textContent="En attente de l’ID LADROT pour la ligne " + pending.row + ". Terminez la saisie. Si l’identifiant est visible, cliquez sur « Désigner l’ID LADROT ».";
      return;
    }
    if (!token && pending) throw Error("Une ligne attend son ID LADROT. Utilisez le raccourci Excel pour démarrer le prochain dossier.");
    const result=await appelerFormulaire(value => globalThis.formulaireExcel.remplir(value),[source]);
    etat.textContent=result.message;
    if (result.ok && token) {
      // Mémoriser le lien avec cet onglet uniquement après un remplissage réussi.
      // La fiche complète ne fait pas partie de la session mémorisée.
      pending={token,row:source.row,origin:page.origin,path:page.pathname};
      await chrome.storage.session.set({[cleOnglet]:pending});
      etat.textContent+=" Dès que l’ID LADROT apparaît, utilisez de nouveau le raccourci du navigateur.";
    } else if (result.ok) {
      etat.textContent+=" Pour récupérer l’ID LADROT, démarrez la saisie avec la macro Excel.";
    }
  } catch(error) {
    etat.textContent=error.message || "Rouvrez l’extension dans l’onglet de la fiche.";
  } finally {
    desactiver(false);enCours=false;
  }
}

manual.addEventListener("paste", event => {
  event.preventDefault();
  void paste(event.clipboardData.getData("text/plain"));
});
async function ouvrirOutil(func) {
  try {await appelerFormulaire(func);window.close();}
  catch {etat.textContent="Rouvrez l’extension sur le formulaire et réessayez.";}
}
setup.addEventListener("click",() => ouvrirOutil(() => globalThis.formulaireExcel.ouvrirAssociations()));
bind.addEventListener("click",() => ouvrirOutil(() => globalThis.retourLadrot.designer()));
copyId.addEventListener("click",async () => {
  try {
    const info=await lireResultat();
    if (!info.number || info.number!==identifiantPropose) throw Error("L’identifiant affiché a changé. Rouvrez l’extension pour le vérifier.");
    // Relire la valeur affichée juste avant la copie. Excel vérifiera la même clé au retour.
    await navigator.clipboard.writeText(["EXCELID1",pending.token,info.number].join("\t"));
    etat.textContent="Revenez dans Excel et utilisez votre raccourci pour inscrire l’ID dans la cellule A" + pending.row + ".";
    copyId.hidden=true;
  } catch {etat.textContent="Impossible de copier cet ID LADROT. Rouvrez l’extension et vérifiez la valeur affichée.";}
});
async function demarrer() {
  const [tab]=await chrome.tabs.query({active:true,currentWindow:true});
  if (!tab?.id || !/^https?:\/\//i.test(tab.url || "")) {
    etat.textContent="Ouvrez d’abord le formulaire dans un onglet du navigateur.";return;
  }
  // Une navigation vers le dossier enregistré peut changer le chemin tout en gardant
  // la même origine et le même onglet. Le lien reste valable pour ce dossier.
  page=new URL(tab.url);
  target={tabId:tab.id};cleOnglet="pending|"+tab.id;
  const [frame]=await chrome.scripting.executeScript({target,files:["identifiant.js","formulaire.js"]});
  if (frame.documentId) target={tabId:tab.id,documentIds:[frame.documentId]};
  pending=(await chrome.storage.session.get(cleOnglet))[cleOnglet];
  if (pending?.origin!==page.origin) pending=null;
  desactiver(false);
  const info=await lireResultat();
  if (pending && info.number) {proposerRetour(info.number);return;}
  if (pending && pending.path!==page.pathname) {
    etat.textContent="En attente de l’ID LADROT pour la ligne " + pending.row + ". Si l’identifiant est visible, cliquez sur « Désigner l’ID LADROT », sélectionnez sa valeur, puis rouvrez l’extension.";return;
  }
  const count=await appelerFormulaire(() => globalThis.formulaireExcel.nombreAssociations());
  if (!count) {
    etat.textContent="Cliquez sur « Associer les colonnes » pour choisir les champs. Avec « Désigner l’ID LADROT », sélectionnez la valeur que le site génère pour le client.";return;
  }
  let text;
  try {text=await navigator.clipboard.readText();}
  catch {
    manual.hidden=false;manual.focus();
    etat.textContent="Appuyez sur Ctrl+V dans la zone ci-dessous pour coller la ligne.";return;
  }
  await paste(text);
}
demarrer().catch(() => {etat.textContent="Onglet inaccessible. Rouvrez l’extension dans le formulaire.";});

(() => {
  if (globalThis.formulaireExcel) return;

  // Ce module agit sur les champs de la page principale après activation de l’extension.
  const selecteurChamp = 'textarea,select,input:not([type]),input:is([type="text" i],[type="email" i],[type="tel" i],[type="url" i],[type="search" i],[type="number" i])';
  const prefix = () => "rowPaste|" + JSON.stringify([location.origin, location.pathname]) + "|";
  const colonnesSuggerees = {B:"Nom", C:"Prénom", E:"Rue / app.", F:"Ville", G:"Code postal", H:"Province", I:"Courriel", J:"Téléphone principal"};
  let panneauAssociations;
  let remplissageEnCours = false;

  function indiceColonne(letter) {
    return [...letter].reduce((n, ch) => n * 26 + ch.charCodeAt(0) - 64, 0) - 1;
  }
  // Chaque formulaire a ses associations, regroupées par origine et chemin de page.
  async function lireAssociations() {
    const key = prefix();
    const data = await chrome.storage.local.get(null);
    return Object.entries(data).filter(([name]) => name.startsWith(key) && !/^A(?:[1-9]\d*)?$/.test(name.slice(key.length)))
      .map(([name, field]) => ({col:name.slice(key.length), field}))
      .sort((a,b) => a.col.localeCompare(b.col, "fr-CA", {numeric:true}));
  }
  function clean(value) {
    return (value || "").trim().replace(/\s+/g, " ");
  }
  function libelleChamp(el) {
    return clean(el.labels?.[0]?.textContent);
  }
  function group(el) {
    const titres = "legend,h1,h2,h3,h4,h5,h6,[role=heading],.card-title,.card-header,.panel-heading";
    for (let parent = el.parentElement; parent && parent !== document.body; parent = parent.parentElement) {
      const names = [...parent.querySelectorAll(titres)].filter(h =>
        !h.contains(el) && (h.compareDocumentPosition(el) & Node.DOCUMENT_POSITION_FOLLOWING));
      if (names.length) return clean(names.at(-1).textContent);
    }
    return "";
  }
  // Mémoriser les caractéristiques du champ et son titre de section. Deux champs
  // « Ville » peuvent avoir la même apparence, mais appartenir à des sections différentes.
  function describe(el) {
    return {
      tag:el.tagName, type:el.type, id:el.id, name:el.getAttribute("name") || "",
      placeholder:el.getAttribute("placeholder") || "", aria:el.getAttribute("aria-label") || "",
      label:libelleChamp(el), group:group(el)
    };
  }
  function available(el) {
    return el.matches(selecteurChamp) && !el.matches(":disabled") && !el.readOnly &&
      !el.multiple && !el.closest("[inert],[hidden]") && el.getClientRects().length > 0 &&
      getComputedStyle(el).visibility === "visible";
  }
  function fields() {
    return [...document.querySelectorAll(selecteurChamp)].filter(available);
  }
  function matches(el, meta) {
    if (el.tagName !== meta.tag || el.type !== meta.type) return false;
    if (meta.id && el.id !== meta.id) return false;
    if (meta.name && el.getAttribute("name") !== meta.name) return false;
    if (meta.placeholder && el.getAttribute("placeholder") !== meta.placeholder) return false;
    if (meta.aria && el.getAttribute("aria-label") !== meta.aria) return false;
    if (meta.label && libelleChamp(el) !== meta.label) return false;
    return !meta.group || group(el) === meta.group;
  }
  // Une destination doit correspondre à un seul champ visible et modifiable.
  function retrouverChamp(meta) {
    const candidats = fields().filter(el => matches(el,meta));
    return candidats.length === 1 ? candidats[0] : null;
  }
  function intitule(meta) {
    return meta.placeholder || meta.aria || meta.label || meta.name || meta.id;
  }

  // Excel sépare les cellules par des tabulations et double les guillemets à l’intérieur d’une cellule.
  function lireLigne(raw) {
    if (typeof raw !== "string" || !raw.trim()) throw Error("Copiez d’abord une ligne de données dans Excel.");
    const text = raw.replace(/\r\n?/g, "\n").replace(/\n$/, "");
    const cells = [];
    let value = "", quoted = false, closed = false;
    for (let i = 0; i < text.length; i++) {
      const c = text[i];
      if (quoted) {
        if (c !== '"') value += c;
        else if (text[i + 1] === '"') { value += '"'; i++; }
        else { quoted = false; closed = true; }
      } else if (c === "\t") {
        cells.push(value); value = ""; closed = false;
      } else if (c === "\n") {
        throw Error("Plusieurs lignes ont été copiées. Sélectionnez une seule ligne à partir de la colonne A.");
      } else if (closed) {
        throw Error("Format de ligne illisible. Copiez de nouveau les cellules dans Excel.");
      } else if (c === '"' && !value) quoted = true;
      else value += c;
    }
    if (quoted) throw Error("Une cellule entre guillemets est incomplète. Copiez de nouveau la ligne.");
    cells.push(value);
    if (cells.length < 2) throw Error("Copiez la ligne Excel à partir de la colonne A.");
    return cells;
  }
  function valeurCible(el, raw, col) {
    if (el instanceof HTMLSelectElement) {
      const key = raw.trim().toLocaleLowerCase();
      const options = [...el.options].filter(o => !o.disabled && !o.closest("optgroup:disabled") &&
        [o.value, o.label].some(v => v.trim().toLocaleLowerCase() === key));
      if (options.length !== 1) throw Error("La source " + col + " ne correspond pas à un choix unique dans la liste.");
      return options[0].value;
    }
    const probe = el.cloneNode(false);
    probe.value = raw;
    if (probe.value !== raw || (el.maxLength >= 0 && raw.length > el.maxLength) || !probe.validity.valid)
      throw Error("Vérifiez le format de la source " + col + " pour ce champ.");
    return raw;
  }
  // Utiliser le setter natif puis émettre les événements attendus par le formulaire.
  // La page peut ainsi appliquer ses règles de validation et ses formats habituels.
  function ecrireValeur(el, value) {
    let proto = HTMLInputElement.prototype;
    if (el instanceof HTMLTextAreaElement) proto = HTMLTextAreaElement.prototype;
    if (el instanceof HTMLSelectElement) proto = HTMLSelectElement.prototype;
    Object.getOwnPropertyDescriptor(proto, "value").set.call(el,value);
    el.dispatchEvent(new Event("input",{bubbles:true, composed:true}));
    el.dispatchEvent(new Event("change",{bubbles:true, composed:true}));
  }
  async function remplir(text) {
    if (remplissageEnCours || panneauAssociations) return {ok:false, message:"Terminez l’association des champs avant de coller."};
    remplissageEnCours = true;
    let champsRemplis = 0;
    try {
      const settings = await lireAssociations();
      if (!settings.length) throw Error("Associez d’abord les colonnes aux champs de ce formulaire.");
      const ligneSuivie = typeof text !== "string";
      const cells = ligneSuivie ? text.cells : lireLigne(text);
      if (/^nom$/i.test(cells[1]?.trim()) && /^pr[eé]nom$/i.test(cells[2]?.trim()))
        throw Error("Cette ligne contient les en-têtes. Copiez une ligne de client.");
      // Examiner toute la ligne avant la première modification pour détecter un conflit tardif.
      if (cells[0] !== "") throw Error("La cellule ID LADROT de la ligne copiée est déjà occupée.");
      const plan = [];
      const used = new Set();
      for (const {col,field} of settings) {
        const fixed = /\d$/.test(col);
        const idx = indiceColonne(col.match(/^[A-Z]+/)[0]);
        let source;
        if (fixed && ligneSuivie && String(text.row) === col.match(/\d+$/)[0]) source = cells[idx];
        else if (fixed) source = ligneSuivie ? text.fixed?.[col] : undefined;
        else source = cells[idx];
        if (typeof source !== "string") throw Error(fixed ? "La cellule fixe " + col + " est absente. Cliquez sur Terminer dans le panneau d’association, puis utilisez le raccourci Excel." : "La sélection se termine avant " + col + ". Copiez de A jusqu’à la dernière colonne associée.");
        const el = retrouverChamp(field);
        if (!el) throw Error("Impossible de retrouver le champ de la source " + col + ". Vérifiez le formulaire ou associez de nouveau ce champ.");
        if (used.has(el)) throw Error("Deux sources visent le même champ. Corrigez les associations.");
        used.add(el);
        const value = source === "" ? "" : valeurCible(el,source.replace(/\r\n?/g,"\n"),col);
        if (el.value && el.value !== value) throw Error("La source " + col + " vise un champ déjà rempli avec une autre valeur. Ouvrez une fiche vierge.");
        plan.push({col,field,value});
      }
      if (!plan.some(item => item.value)) throw Error("Toutes les cellules associées sont vides.");
      // Retrouver la destination avant chaque écriture, car la page peut recréer ses champs.
      for (const {col,field,value} of plan) {
        const el = retrouverChamp(field);
        if (!el || (el.value && el.value !== value)) throw Error("Le formulaire a changé pendant le remplissage de la source " + col + ".");
        if (el.value === value) continue;
        champsRemplis++;
        ecrireValeur(el,value);
      }
      await new Promise(done => setTimeout(done,100));
      for (const {col,field,value} of plan) {
        if (retrouverChamp(field)?.value !== value) throw Error("La page a modifié la valeur de la source " + col + ".");
      }
      return {ok:true, message:champsRemplis ? "Remplissage terminé pour " + champsRemplis + (champsRemplis === 1 ? " champ. " : " champs. ") + "Vérifiez la fiche et enregistrez-la selon votre procédure habituelle." :
        "Les champs associés contiennent déjà les valeurs de cette ligne."};
    } catch (error) {
      return {ok:false, message:error.message + (champsRemplis ? " Vérifiez les champs déjà remplis avant de continuer." : " Aucun champ modifié.")};
    } finally {
      remplissageEnCours = false;
    }
  }

  function node(tag, text, parent) {
    const el = document.createElement(tag);
    if (text) el.textContent = text;
    parent.append(el);
    return el;
  }
  function button(text, parent, run) {
    const el = node("button", text, parent);
    el.type = "button";
    el.addEventListener("click", run);
    return el;
  }
  async function ouvrirAssociations() {
    if (panneauAssociations?.isConnected || remplissageEnCours) return;
    const page = prefix();
    const host = document.createElement("div");
    host.style.cssText = "all:initial!important;position:fixed!important;top:10px!important;right:10px!important;max-width:calc(100vw - 20px)!important;z-index:2147483647!important;";
    // Isoler le panneau des styles de la fiche et de la recherche des champs de saisie.
    const root = host.attachShadow({mode:"closed"});
    document.documentElement.append(host);
    panneauAssociations = host;
    const style = node("style", "", root);
    style.textContent = [
      "*{box-sizing:border-box}main{width:min(345px,calc(100vw - 20px));max-height:85vh;overflow:auto;background:#fbfefc;color:#183b32;border:1px solid #a5bdb3;border-top:4px solid #196752;border-radius:8px;padding:14px;font:14px/1.45 'Segoe UI',sans-serif;box-shadow:0 6px 25px #0003}",
      "header,.tools,li{display:flex;gap:8px;align-items:center}header,li{justify-content:space-between}h2{font-size:17px;margin:0}p{margin:10px 0}input,button{font:inherit}input{width:78px;text-transform:uppercase;padding:4px}button{cursor:pointer;border:1px solid #a5bdb3;border-radius:4px;background:white;color:#17372c;padding:4px 8px}.primary{background:#195c4b;color:white}ul{padding:0;margin:8px 0 0}li{border-top:1px solid #dde8e3;padding:4px 0;overflow-wrap:anywhere}small{display:block;margin:10px 0}[hidden]{display:none!important}"
    ].join("");
    const panel = node("main","",root);
    const head = node("header","",panel);
    node("h2","Associer les colonnes",head);
    const actions = node("div","",head); actions.className="tools";
    button("Déplacer",actions,() => {
      const left = host.style.left === "10px";
      host.style.setProperty("left",left ? "auto" : "10px","important");
      host.style.setProperty("right",left ? "10px" : "auto","important");
    });
    button("Terminer",actions,() => {
      if (enregistrementEnCours) return;
      const refs = current.map(item => item.col).filter(col => /\d$/.test(col));
      // Excel doit connaître les adresses fixes avant la prochaine copie de la ligne.
      const packet = "EXCELMAP1|" + refs.join(",");
      const box = node("textarea","",panel); box.value=packet; box.select();
      let copied=false;
      try {copied=document.execCommand("copy");} catch {}
      box.remove(); close();
      if (!copied) window.prompt("Copiez cette configuration, puis utilisez le raccourci Excel sur la ligne du client.",packet);
    });
    const body = node("div","",panel);
    node("p","Saisissez B pour la colonne B de la ligne active, ou F2 pour la cellule fixe. Sélectionnez ensuite le champ du formulaire.",body);
    const controls = node("div","",body); controls.className = "tools";
    const label = node("label","Source ",controls);
    const sourceExcel = node("input","",label); sourceExcel.maxLength=12; sourceExcel.value="B";
    button("Choisir le champ",controls,choisirSource).className="primary";
    node("small","La colonne A reçoit l’ID LADROT. Associez les autres colonnes aux champs à remplir.",body);
    const list = node("ul","",body);
    const notice = node("p","",panel); notice.setAttribute("role","status");
    let sourceChoisie = "";
    let enregistrementEnCours = false;
    let current = [];
    function close() {
      document.removeEventListener("click",designer,true);
      document.removeEventListener("keydown",gererTouche,true);
      host.remove(); panneauAssociations = null;
    }
    function gererTouche(e) {
      if (e.key === "Escape") { e.stopPropagation(); e.preventDefault(); if (sourceChoisie) annulerChoix(); else close(); }
    }
    function annulerChoix() {
      sourceChoisie = ""; body.hidden = false; notice.textContent = "";
    }
    function choisirSource() {
      if (enregistrementEnCours) return;
      const col = sourceExcel.value.trim().toUpperCase().replace(/\$/g,"");
      const parts = /^([A-Z]{1,3})([1-9]\d{0,6})?$/.exec(col);
      if (!parts || indiceColonne(parts[1]) >= 16384 || Number(parts[2] || 1) > 1048576) {
        notice.textContent = "Saisissez une colonne ou une adresse de cellule valide, par exemple B ou F2."; return;
      }
      if (parts[1] === "A") {
        notice.textContent="La colonne A est réservée au retour de l’ID LADROT."; return;
      }
      sourceChoisie=col; body.hidden=true;
      notice.textContent="Cliquez sur le champ correspondant à " + col + ". Échap pour annuler.";
    }
    async function actualiserListe() {
      current = await lireAssociations();
      list.replaceChildren();
      for (const {col,field} of current) {
        const row = node("li","",list);
        node("span",col + " → " + intitule(field) + (field.group ? " (" + field.group + ")" : ""),row);
        button("Retirer",row,async () => {
          try { await chrome.storage.local.remove(page + col); await actualiserListe(); }
          catch { notice.textContent = "Impossible de retirer cette association. Rouvrez le panneau."; }
        });
      }
      const next = Object.keys(colonnesSuggerees).find(col => !current.some(item => item.col === col));
      if (next) sourceExcel.value = next;
    }
    async function designer(e) {
      if (!sourceChoisie || e.composedPath().includes(host)) return;
      e.preventDefault(); e.stopImmediatePropagation();
      if (enregistrementEnCours) return;
      if (page !== prefix()) {close();return;}
      const el = e.target;
      if (!(el instanceof Element) || !available(el)) {
        notice.textContent="Cliquez dans un champ modifiable ou une liste à choix unique pour " + sourceChoisie + ". Échap pour annuler."; return;
      }
      const meta = describe(el);
      // Le libellé et la section permettent de retrouver le champ même si son identifiant HTML change.
      const sansIdentifiant = {...meta, id:""};
      if (intitule(sansIdentifiant) && retrouverChamp(sansIdentifiant) === el) meta.id = "";
      if (!intitule(meta) || retrouverChamp(meta) !== el) {
        notice.textContent="Impossible de reconnaître ce champ de manière unique. Choisissez un autre champ ou appuyez sur Échap."; return;
      }
      if (current.some(item => item.col !== sourceChoisie && matches(el,item.field))) {
        notice.textContent="Ce champ a déjà une source. Appuyez sur Échap puis retirez son association."; return;
      }
      const col = sourceChoisie;
      enregistrementEnCours=true;
      try {
        await chrome.storage.local.set({[page + col]:meta});
        annulerChoix(); await actualiserListe(); notice.textContent="Association enregistrée pour " + col + " → " + intitule(meta) + ".";
      } catch { annulerChoix(); notice.textContent="Impossible d’enregistrer cette association. Rouvrez le panneau."; }
      finally { enregistrementEnCours=false; }
    }
    document.addEventListener("keydown",gererTouche,true);
    document.addEventListener("click",designer,true);
    try { await actualiserListe(); } catch { close(); throw Error("Les associations sont indisponibles."); }
  }
  globalThis.formulaireExcel = {nombreAssociations:async () => (await lireAssociations()).length, ouvrirAssociations, remplir};
})();

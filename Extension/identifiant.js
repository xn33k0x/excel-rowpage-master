(() => {
  if (globalThis.retourLadrot) return;
  // Décrire l’emplacement de la valeur de sortie et relire sa valeur au moment du retour.
  const cleReglage = () => "rowPasteOutput|" + location.origin;
  const normal = text => (text || "").normalize("NFD").replace(/[\u0300-\u036f]/g,"").toLowerCase().replace(/[\s:]+/g," ").trim();
  const estLibelle = text => ["id ladrot", "numero du dossier", "file number", "dossier number"].includes(normal(text));
  const read = el => ((el.matches("input,textarea") ? el.value : el.textContent) || "").trim();
  function repererLibelles(label) {
    const w = document.createTreeWalker(document.body, NodeFilter.SHOW_TEXT);
    const found = [];
    while (w.nextNode()) {
      const el = w.currentNode.parentElement;
      if (estLibelle(w.currentNode.textContent) && (!label || normal(w.currentNode.textContent) === label) && el && el.getClientRects().length && !found.includes(el)) found.push(el);
    }
    return found;
  }
  // Retrouver le libellé mémorisé, remonter à son ancêtre commun avec la valeur,
  // puis suivre le chemin des éléments. Un chemin devenu ambigu interrompt la lecture.
  function retrouverCible(rule) {
    if (!rule) return null;
    if (rule.kind === "id") {
      const all = document.querySelectorAll("#" + CSS.escape(rule.id));
      return all.length === 1 && all[0].tagName === rule.tag ? all[0] : null;
    }
    const labels = rule.label ? repererLibelles(rule.label) : repererLibelles().filter(el => ![...el.childNodes].some(n => n.nodeType === Node.TEXT_NODE && normal(n.textContent) === "id ladrot"));
    if (labels.length !== 1) return null;
    let el = labels[0];
    for (let i = 0; el && i < rule.up; i++) el = el.parentElement;
    for (const [idx,tag] of rule.down) {
      el = el?.children[idx]; if (!el || el.tagName !== tag) return null;
    }
    return el || null;
  }
  function decrireCible(target) {
    // Conserver le libellé choisi lorsque la fiche affiche plusieurs identifiants.
    const candidats = [];
    for (const ancre of repererLibelles()) {
      for (let parent=ancre, up=0; parent && up<=4 && parent!==document.body; up++,parent=parent.parentElement) {
        if (!parent.contains(target)) continue;
        const down=[];
        for (let el=target; el!==parent; el=el.parentElement)
          down.unshift([[...el.parentElement.children].indexOf(el),el.tagName]);
        const texte=[...ancre.childNodes].find(n => n.nodeType===Node.TEXT_NODE && estLibelle(n.textContent));
        if (texte && target!==ancre && down.length<=5) {
          const regle={kind:"label",label:normal(texte.textContent),up,down};
          if (retrouverCible(regle)===target) candidats.push({regle,distance:up+down.length});
        }
        break;
      }
    }
    candidats.sort((a,b)=>a.distance-b.distance);
    if (candidats.length && candidats[0].distance!==candidats[1]?.distance) return candidats[0].regle;
    if (target.id && document.querySelectorAll("#" + CSS.escape(target.id)).length === 1)
      return {kind:"id",id:target.id,tag:target.tagName};
    throw Error("Sélectionnez la valeur de l’ID LADROT ou sa pastille N/A.");
  }
  // Le texte provisoire N/A ne contient aucun chiffre. Conserver la casse et les zéros
  // de la valeur réellement affichée par l’application.
  function estIdentifiant(text) {
    return text.length <= 80 && /\d/.test(text) && /^[\p{L}\d][\p{L}\d ._/-]*$/u.test(text);
  }
  async function lireIdentifiant() {
    const key=cleReglage();
    const rule=(await chrome.storage.local.get(key))[key];
    const el=retrouverCible(rule);
    if (!el || !el.getClientRects().length) return {configured:!!rule,number:""};
    const value=read(el);
    return {configured:true,number:estIdentifiant(value)?value:""};
  }
  let selectionEnCours=false;
  function designer() {
    if (selectionEnCours) return;
    selectionEnCours=true;
    const prompt=document.createElement("div");
    prompt.style.cssText="position:fixed!important;left:8px!important;bottom:8px!important;z-index:2147483647!important;max-width:380px!important;padding:12px!important;background:#173c32!important;color:white!important;font:14px/1.5 'Segoe UI',sans-serif!important;";
    prompt.textContent="Cliquez sur la valeur de l’ID LADROT généré pour ce client, ou sur sa pastille N/A. Échap pour annuler.";
    document.documentElement.append(prompt);
    function close() {selectionEnCours=false;prompt.remove();document.removeEventListener("click",selectionnerCible,true);document.removeEventListener("keydown",escape,true);}
    function escape(e) {if(e.key==="Escape"){e.preventDefault();e.stopImmediatePropagation();close();}}
    async function selectionnerCible(e) {
      e.preventDefault(); e.stopImmediatePropagation();
      if (!(e.target instanceof Element) || e.target === prompt) return;
      try {
        const val=read(e.target);
        if (!(estIdentifiant(val) || /^(n\s*\/?\s*a|[-–—]|)$/i.test(val))) throw Error("Cliquez directement sur l’identifiant affiché ou sur N/A.");
        const rule=decrireCible(e.target);
        if(retrouverCible(rule)!==e.target) throw Error("Impossible de reconnaître cet identifiant de manière unique.");
        await chrome.storage.local.set({[cleReglage()]:rule});
        close();
      } catch(error){prompt.textContent=error.message+" Échap pour annuler.";}
    }
    document.addEventListener("click",selectionnerCible,true);document.addEventListener("keydown",escape,true);
  }
  globalThis.retourLadrot={lireIdentifiant,designer};
})();

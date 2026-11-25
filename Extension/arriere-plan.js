// Le menu de l’icône ouvre les associations sans déclencher le collage d’une ligne.
chrome.runtime.onInstalled.addListener(() => {
  chrome.contextMenus.create({ id: "map-columns", title: "Associer les colonnes sur cette page", contexts: ["action"] });
});

chrome.contextMenus.onClicked.addListener(async (info, tab) => {
  if (info.menuItemId !== "map-columns" || !tab?.id) return;
  try {
    await chrome.scripting.executeScript({target: {tabId: tab.id}, files: ["identifiant.js","formulaire.js"]});
    await chrome.scripting.executeScript({target: {tabId: tab.id}, func: () => globalThis.formulaireExcel.ouvrirAssociations()});
  } catch {
    await chrome.action.setBadgeText({tabId: tab.id, text: "!"});
    await chrome.action.setTitle({tabId: tab.id, title: "Ouvrez le formulaire dans un onglet ordinaire, puis réessayez."});
  }
});

// À la fermeture de l’onglet, supprimer sa clé de suivi temporaire.
chrome.tabs.onRemoved.addListener(tabId => { void chrome.storage.session.remove("pending|" + tabId); });

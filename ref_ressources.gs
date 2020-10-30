// Auteur : Pôle Qualité 022 (Aubin Tchoï)

// Voici quelques lignes que j'ai utilisées de manière récurrente dans différents scripts .gs

// Transforme un array en object à partir d'un header heads, les clés étant les objets de heads
// Utile pour traiter un tableau de données (combiné avec un map)
// Pas sensible au fait que le header peut présenter plusieurs fois la même valeur (utile pour un GForms)
function ArrayToObj(arr, heads) {
    return heads.reduce((o, k, i) => (o[k] = (r[i] != "") ? arr[i] : o[k] || '', o), {});
}

// Écran de chargement
function display_LoadingScreen(msg) {
    let htmlLoading = HtmlService
    .createHtmlOutput(`<img src="https://raw.githubusercontent.com/aubin-tchoi/Polkali/main/images/Kkooljem.gif" alt="Loading" width="885" height="498">`)
    .setWidth(900)
    .setHeight(500);
    SpreadsheetApp.getUi().showModelessDialog(htmlLoading, msg);
  }
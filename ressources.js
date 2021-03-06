/* Si vous avez des questions à propos de ce script contactez Aubin Tchoï (Directeur Qualité 022) */

// Voici quelques lignes que j'ai utilisées de manière récurrente dans différents scripts .gs
// Rappels importants : toutes les fonctions ne peuvent être appelées à partir de n'importe quel contexte (ex : SpreadsheetApp.getUi)

// Transforme un array en object à partir d'un header heads, les clés étant les objets de heads
// Utile pour traiter un tableau de données (combiné avec un map)
// Pas sensible au fait que le header peut présenter plusieurs fois la même valeur (utile pour un GForms)
function ArrayToObj(arr, heads) {
    return heads.reduce((obj, key, idx) => (obj[key] = (arr[idx] != "") ? arr[idx] : obj[key] || '', obj), {});
}

// Écran de chargement
function displayLoadingScreen(msg) {
    let htmlLoading = HtmlService
    .createHtmlOutput(`<img src="https://raw.githubusercontent.com/aubin-tchoi/Polkali/main/images/loadingScreen.gif" alt="Loading" width="885" height="498">`)
    .setWidth(900)
    .setHeight(500);
    SpreadsheetApp.getUi().showModelessDialog(htmlLoading, msg);
}

// Formatage de phrase pour leur ajouter 1 majuscule au début de chaque mot avec le reste en minuscules
function properCase(phrase) {
    return (phrase.toLowerCase().replace(/([^A-Za-zÀ-ÖØ-öø-ÿ])([A-Za-zÀ-ÖØ-öø-ÿ])(?=[A-Za-zÀ-ÖØ-öø-ÿ]{2})|^([A-Za-zÀ-ÖØ-öø-ÿ])/g, function(_, g1, g2, g3) {
      return (typeof g1 === 'undefined') ? g3.toUpperCase() : g1 + g2.toUpperCase(); }));
}

// Formatage de mot pour leur ajouter 1 espace toutes les n lettres (par exemple : numéros de tel)
function addSpaces(x,n) {
    let regexp = new RegExp(`\\d{${n}}(?=.)`, "g");
    return x.toString().replace(regexp, "$& ");
}

// Couleurs de PEP
let colors = {
    Bordeaux: "#8E3232",
    Jaune: "#FFBE2B",
    Gris: "#404040",
    Châtain: "#A29386"
};

// Retrouve toutes les valeurs uniques dans un array d'objets et les renvoie dans un array
const uniqueValues = (key, data) => data.map((row) => row[key]).filter((val, idx, arr) => arr.indexOf(val) == idx);

// Renvoie true ssi str peut être converti en flottant
function isNumeric(str) {
    if (typeof str != "string") {return false;} // Only processes strings
    return !isNaN(str) && // use type coercion to parse the _entirety_ of the string (`parseFloat` alone does not do this)...
      !isNaN(parseFloat(str)) // ...and ensure strings of whitespace fail
}

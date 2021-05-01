/* Si vous avez des questions à propos de ce script contactez Aubin Tchoï (Directeur Qualité 022) */

/** Main function */
function duplication(URLs) {
  let sourceFolder = DriveApp.getFolderById(URLtoID(URLs.source)),
    targetFolder = DriveApp.getFolderById(URLtoID(URLs.target));
  duplicateFolder(sourceFolder, targetFolder);
}

/** Including a file */
function include(filename) {
  return HtmlService.createHtmlOutputFromFile(filename).getContent();
}

/** Choosing a background image randomly */
function findURL() {
  let urls = ["https://images.unsplash.com/photo-1536679203421-b6d947f6f78d?ixid=MnwxMjA3fDB8MHxwaG90by1wYWdlfHx8fGVufDB8fHx8&ixlib=rb-1.2.1&auto=format&fit=crop&w=1950&q=80",
    "https://images.unsplash.com/photo-1536678089453-b1e655b5fb5a?ixlib=rb-1.2.1&ixid=MnwxMjA3fDB8MHxwaG90by1wYWdlfHx8fGVufDB8fHx8&auto=format&fit=crop&w=1782&q=80",
    "https://images.unsplash.com/photo-1545569369-b0f813083ddb?ixid=MnwxMjA3fDB8MHxwaG90by1wYWdlfHx8fGVufDB8fHx8&ixlib=rb-1.2.1&auto=format&fit=crop&w=1834&q=80",
    "https://images.unsplash.com/photo-1531065520053-29ba2ae5409a?ixid=MnwxMjA3fDB8MHxwaG90by1wYWdlfHx8fGVufDB8fHx8&ixlib=rb-1.2.1&auto=format&fit=crop&w=1834&q=80",
    "https://images.pexels.com/photos/6619945/pexels-photo-6619945.jpeg?auto=compress&cs=tinysrgb&dpr=3&h=750&w=1260",
    "https://images.pexels.com/photos/7693918/pexels-photo-7693918.jpeg?auto=compress&cs=tinysrgb&dpr=3&h=750&w=1260",
    "https://images.unsplash.com/photo-1575856323725-50dca3234ede?ixid=MnwxMjA3fDB8MHxwaG90by1wYWdlfHx8fGVufDB8fHx8&ixlib=rb-1.2.1&auto=format&fit=crop&w=1834&q=80",
    "https://images.unsplash.com/photo-1602280618240-81d57d8d6eeb?ixid=MnwxMjA3fDB8MHxwaG90by1wYWdlfHx8fGVufDB8fHx8&ixlib=rb-1.2.1&auto=format&fit=crop&w=1833&q=80",
    "https://images.squarespace-cdn.com/content/v1/59440628b3db2b2f36a66f10/1600045727721-EG0AGNGTPP5KP8FNV2DZ/ke17ZwdGBToddI8pDm48kJx1dffM5XEzku2PcltQNyF7gQa3H78H3Y0txjaiv_0fDoOvxcdMmMKkDsyUqMSsMWxHk725yiiHCCLfrh8O1z4YTzHvnKhyp6Da-NYroOW3ZGjoBKy3azqku80C789l0hGaawTDWlunVGEFKwsEdnFEPBzLZDmAkZBVqdIKUbMcnUUnBgR-Z4zEehg67J4_kg/92EF321D-A366-4DDE-BD55-F2C44BE38EE8.jpeg?format=2500w",
    "https://images.squarespace-cdn.com/content/v1/59440628b3db2b2f36a66f10/1600045876317-7WY8LOTTJBGCTUE784FH/ke17ZwdGBToddI8pDm48kMXRibDYMhUiookWqwUxEZ97gQa3H78H3Y0txjaiv_0fDoOvxcdMmMKkDsyUqMSsMWxHk725yiiHCCLfrh8O1z4YTzHvnKhyp6Da-NYroOW3ZGjoBKy3azqku80C789l0luUmcNM2NMBIHLdYyXL-Jww_XBra4mrrAHD6FMA3bNKOBm5vyMDUBjVQdcIrt03OQ/E4943A7A-0434-4CE2-B7D2-0838C29CA84D.jpeg?format=2500w",
    "https://images.unsplash.com/photo-1613369307387-48c79d566dea?ixid=MnwxMjA3fDB8MHxwaG90by1wYWdlfHx8fGVufDB8fHx8&ixlib=rb-1.2.1&auto=format&fit=crop&w=1834&q=80",
    "https://images.pexels.com/photos/6620070/pexels-photo-6620070.jpeg?auto=compress&cs=tinysrgb&dpr=3&h=750&w=1260",
    "https://images.pexels.com/photos/35629/bing-cherries-ripe-red-fruit.jpg?auto=compress&cs=tinysrgb&dpr=3&h=750&w=1260",
    "https://images.unsplash.com/photo-1619787942043-99ae128ca5c1?ixid=MnwxMjA3fDB8MHxwaG90by1wYWdlfHx8fGVufDB8fHx8&ixlib=rb-1.2.1&auto=format&fit=crop&w=1875&q=80",
    "https://images.pexels.com/photos/1995730/pexels-photo-1995730.jpeg?auto=compress&cs=tinysrgb&dpr=3&h=750&w=1260",
    "https://images.pexels.com/photos/132419/pexels-photo-132419.jpeg?auto=compress&cs=tinysrgb&dpr=3&h=750&w=1260",
    "https://images.unsplash.com/photo-1619544345442-6508f6c27679?ixid=MnwxMjA3fDB8MHxwaG90by1wYWdlfHx8fGVufDB8fHx8&ixlib=rb-1.2.1&auto=format&fit=crop&w=1950&q=80"
  ];
  return urls[Math.floor(Math.random() * urls.length)];
}

/** GET method */
function doGet(e) {
  return HtmlService.createTemplateFromFile("home").evaluate();
}

/** POST method */
function doPost(e) {
  return HtmlService.createTemplateFromFile("home").evaluate();
}

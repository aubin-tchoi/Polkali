/* Si vous avez des questions à propos de ce script contactez Aubin Tchoï (Directeur Qualité 022) */

/** Main function */
function duplication(IDs) {
  let sourceFolder = DriveApp.getFolderById(IDs.source),
    targetFolder = DriveApp.getFolderById(IDs.target);
  duplicateFolder(sourceFolder, targetFolder);
}

/** Including a file */
function include(filename) {
  return HtmlService.createHtmlOutputFromFile(filename).getContent();
}

/** Choosing a background image randomly */
function findURL() {
  let urls = ["https://images.unsplash.com/photo-1604214590838-e440efa410b1?ixid=MnwxMjA3fDB8MHxwaG90by1wYWdlfHx8fGVufDB8fHx8&ixlib=rb-1.2.1&auto=format&fit=crop&w=1834&q=80",
    "https://images.unsplash.com/photo-1604214932030-7821ab71fe81?ixid=MnwxMjA3fDB8MHxwaG90by1wYWdlfHx8fGVufDB8fHx8&ixlib=rb-1.2.1&auto=format&fit=crop&w=1834&q=80",
    "https://images.unsplash.com/photo-1604214717891-90769c5537ba?ixid=MnwxMjA3fDB8MHxwaG90by1wYWdlfHx8fGVufDB8fHx8&ixlib=rb-1.2.1&auto=format&fit=crop&w=1834&q=80",
    "https://images.unsplash.com/photo-1536679203421-b6d947f6f78d?ixid=MnwxMjA3fDB8MHxwaG90by1wYWdlfHx8fGVufDB8fHx8&ixlib=rb-1.2.1&auto=format&fit=crop&w=1950&q=80",
    "https://images.unsplash.com/photo-1536678089453-b1e655b5fb5a?ixlib=rb-1.2.1&ixid=MnwxMjA3fDB8MHxwaG90by1wYWdlfHx8fGVufDB8fHx8&auto=format&fit=crop&w=1782&q=80",
    "https://images.unsplash.com/photo-1545569369-b0f813083ddb?ixid=MnwxMjA3fDB8MHxwaG90by1wYWdlfHx8fGVufDB8fHx8&ixlib=rb-1.2.1&auto=format&fit=crop&w=1834&q=80",
    "https://images.unsplash.com/photo-1531065520053-29ba2ae5409a?ixid=MnwxMjA3fDB8MHxwaG90by1wYWdlfHx8fGVufDB8fHx8&ixlib=rb-1.2.1&auto=format&fit=crop&w=1834&q=80",
    "https://images.unsplash.com/photo-1619187076652-0bd72f9f2469?ixid=MnwxMjA3fDB8MHxwaG90by1wYWdlfHx8fGVufDB8fHx8&ixlib=rb-1.2.1&auto=format&fit=crop&w=1782&q=80",
    "https://images.pexels.com/photos/6619945/pexels-photo-6619945.jpeg?auto=compress&cs=tinysrgb&dpr=3&h=750&w=1260",
    "https://images.pexels.com/photos/7557420/pexels-photo-7557420.jpeg?auto=compress&cs=tinysrgb&dpr=3&h=750&w=1260",
    "https://images.pexels.com/photos/7693918/pexels-photo-7693918.jpeg?auto=compress&cs=tinysrgb&dpr=3&h=750&w=1260",
    "https://images.unsplash.com/photo-1575856323725-50dca3234ede?ixid=MnwxMjA3fDB8MHxwaG90by1wYWdlfHx8fGVufDB8fHx8&ixlib=rb-1.2.1&auto=format&fit=crop&w=1834&q=80",
    "https://images.unsplash.com/photo-1602280618240-81d57d8d6eeb?ixid=MnwxMjA3fDB8MHxwaG90by1wYWdlfHx8fGVufDB8fHx8&ixlib=rb-1.2.1&auto=format&fit=crop&w=1833&q=80"
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

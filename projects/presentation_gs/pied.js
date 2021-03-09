// Author : Pôle Qualité 022 (Aubin Tchoï)

function onOpen() {
  const ui = DocumentApp.getUi();
  ui.alert("Follow-up présentation gs", "Veuillez vous rendre dans le menu 'Formations' puis sélectionner 'Booking' afin d'obtenir quelques approfondissements si nécessaire.", ui.ButtonSet.OK);
  ui.createMenu('Formations')
  .addItem("Booking", "formations")
  .addToUi();
}

function formations() {
  const ui = DocumentApp.getUi();

  ui.alert("Follow-up présentation gs", "Les questions suivantes vont vous proposer de booker différentes sessions d'aide, à la suite desquelles vous serez redirigé vers un membre émérite de la communauté des programmaticiens afin d'augmenter votre patrimoine intellectuel sur le sujet demandé. \n \n Les six premiers à répondre auront une carotte.", ui.ButtonSet.OK);
  
  // Retrieving the user's choices
  let book = [];
  book.push(["Git", ui.alert("C'est Aubin qui régale !", "Souhaitez-vous apprendre les bases de git ?", ui.ButtonSet.YES_NO) == ui.Button.YES ? "Oui" : "Non"]);
  book.push(["Installation d'outils", ui.alert("C'est Aubin qui régale !", "Souhaitez-vous être accompagé dans l'installation des différents outils présentés ?", ui.ButtonSet.YES_NO) == ui.Button.YES ? "Oui" : "Non"]);
  book.push(["Apps script", ui.alert("C'est Aubin qui régale !", "Souhaitez-vous une mini-formation Apps Script pour bien débuter ?", ui.ButtonSet.YES_NO) == ui.Button.YES ? "Oui" : "Non"]);
  let name = ui.prompt("C'est Aubin qui régale !", "Quel est votre prénom ? (Mettez votre vrai prénom pour que je puisse vous envoyer un mail auto)", ui.ButtonSet.OK).getResponseText() || "";

  try {
  // Thank you message
  let img_url = "https://raw.githubusercontent.com/aubin-tchoi/Polkali/main/images/thumbsUp.png",
      thankyou = HtmlService
  .createHtmlOutput(`<span style='font-size: 12pt;'> <span style="font-family: 'roboto', sans-serif;">Le pôle Qualité a été notifié de votre demande.<br/><br/> &nbsp; La bise.</span></span><p style="text-align:center;"><img src="${img_url}" alt="Thumbs up" width="130" height="131"></p>`)
  .setHeight(230)
  .setWidth(500);
  ui.showModelessDialog(thankyou, "Merci de votre participation !");
  } catch(e) {Logger.log(`Could not show thank you display : ${e}`)}
  
  try {
  // “A mind forever Voyaging through strange seas of Thought, alone.”
  addImage(name);
  } catch(e) {Logger.log(`Could not add image ${e}`)}

  // Writing the mail's content
  let msgHtml = `Voici ce qu'a demandé ${name} : <br/> <br/> <ul>`;
  book.forEach(function(request) {msgHtml += `<li> ${request[0]} : ${request[1]} </li>`;});
  msgHtml += "</ul>";

  let htmlOutput = HtmlService.createHtmlOutput(msgHtml),
    msgPlain = msgHtml.replace(/\<br\/\>/gi, '\n').replace(/(<([^>]+)>)/ig, "");
  
  // Sending the mail
  GmailApp.sendEmail("aubin.tchoi@junior-pep.fr", `Formation GS : ${name}`, msgPlain, {htmlBody: htmlOutput.getContent()});
}

function addImage(username) {
  // Computes the position of the nth image
  function compute_position(n) {
    let left = (n <= 2) ? 60 : 550,
        top = (n <= 2) ? n*100 + 40 : (n - 3)*100 + 60;
    return {left: left, top: top};
  }
  
  // Opening our slide and hardcoding IDs for each image
  const slide = SlidesApp.openById("1ms9VJ4Mr1hPgb1aljm1EyGMlVIh2QRNj9KKmiH67ONI").getSlides()[5],
      img_ids = {thibaut: "1Kc3qqdS-kM75ZCrhEYlrXI2QCQJkmD1L",
                 nathan: "1kHtf7JQdyWmatovptvV6T2MM0uNoipvJ",
                 mai: "1Ps0II3fd5kwjwbSP-ggpCONMgQQSGDlY",
                 melodie: "1C6UCFTMpviGnEV3zFXbyY-Pip-1E9u7O",
                 lois: "1u9Q3-pDBuIxlGpphMScN8s4nRuYmhtcM",
                 charlotte: "1wfX8RU4FmOMbi-eK_9YGv4s4GzjQavuL",
                 arnaud: "1KKcyBL438tIPJS_VdNGbXVi1T922EXJp",
                 antoine: "1FYY6UX_EwGk-zKc9VSh16InVWOrLlqFK",
                 anatole: "1XC3H09q3lLUBu5XHGwK4QOkDUbVKx_9e",
                 aubin: "1Kc3qqdS-kM75ZCrhEYlrXI2QCQJkmD1L"},
      name = username.toString().toLowerCase().replace(/[ïî]/gi, "i").replace(/[êéè]/gi, "e").match(/([a-z]+)/gi)[0];
  
  Logger.log(`Users list : ${Object.keys(img_ids)}`);
  Logger.log(`${name} has logged in ! (${Object.keys(img_ids).includes(name)})`);

  if (Object.keys(img_ids).includes(name)) {
    const img_blob = DriveApp.getFileById(img_ids[name]).getBlob();
    
    // Retrieving how many images we've already put
    let cache = CacheService.getDocumentCache(),
        nb_imgs = cache.get("nb_imgs");
    Logger.log(`Initially : ${nb_imgs} images.`);
    
    // Updating the cache
    if (nb_imgs == null) {cache.put("nb_imgs", 1); nb_imgs = 0;}
    else {cache.remove("nb_imgs"); cache.put("nb_imgs", (parseInt(nb_imgs, 10) + 1).toString());}
    Logger.log(`Now : ${parseInt(nb_imgs, 10) + 1} images.`);
    
    if (nb_imgs < 8) {
      // Computing the position of the new images depeding on how many there already are
      const position = compute_position(nb_imgs),
          size = {width: 70, height: 70};
      
      // Inserting the new image
      slide.insertImage(img_blob, position.left, position.top, size.width, size.height);
    }
  }
}

function reset_cache() {
  let cache = CacheService.getDocumentCache();
  cache.remove("nb_imgs");
}

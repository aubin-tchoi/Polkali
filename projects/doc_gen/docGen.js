function templating() {
    const doc = DocumentApp.getActiveDocument(),
        ui = DocumentApp.getUi();
    let body = doc.getBody(),
        infos = ["date", "lieu", "odg"];
    infos.forEach(function(info) {
        let input = ui.prompt("Génération de document", `Entrez le·s ${info}`, ui.ButtonSet.OK);
        body.replaceText(`{${info}}`, input.getResponseText());
    });
}

// Author : Pôle Qualité 022 (Aubin Tchoï)

function addImage(e) {
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
        name = e.user.toString().match(/([a-z]+)/gi)[0];
    
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
Office.onReady((info) => {
  if (info.host === Office.HostType.Excel && Office.context.platform === Office.PlatformType.OfficeOnline) {
    document.getElementById("sideload-msg").style.display = "none";
    document.getElementById("app-body").style.display = "flex";
    
    Excel.run(async (context) => {
      const sheet = context.workbook.worksheets.getActiveWorksheet();
      
      // Ajouter un gestionnaire d'événement pour le clic sur la feuille
      sheet.onSingleClicked.add(handleClick);
      await context.sync();
    }).catch((error) => {
      console.error(error);
    });
  }
});

async function handleClick(event) {
  try {
    await Excel.run(async (context) => {
      const sheet = context.workbook.worksheets.getActiveWorksheet();
      const range = sheet.getRange(event.address);
      range.load(["hyperlink"]);
      
      await context.sync();
      
      if (range.hyperlink && range.hyperlink.address) {
        window.open(range.hyperlink.address, '_blank');
      }
    });
  } catch (error) {
    console.error(error);
  }
}
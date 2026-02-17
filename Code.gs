/**
 * Creates a custom menu in the Google Sheet UI.
 */
function onOpen() {
  const ui = SpreadsheetApp.getUi();
  ui.createMenu('🚀 Slides Sync')
    .addItem('Sync Colors Now', 'updateSlidesFromSheet')
    .addToUi();
}

/**
 * Optimized sync function handling multiple shapes with the same name.
 */
function updateSlidesFromSheet() {
  const sheetId = 'INSERT_GOOGLE_SHEETS_ID'; //Google Sheets ID
  const slideId = 'INSERT_GOOGLE_SLIDES_ID'; //Google Slides ID
  
  const ss = SpreadsheetApp.openById(sheetId);
  const sheet = ss.getSheetByName('INSERT_TAB_NAME'); //Spreadsheet Tab Name
  const data = sheet.getDataRange().getValues();
  
  const deck = SlidesApp.openById(slideId);
  const slides = deck.getSlides();

  // --- STEP 1: Map Titles to ARRAYS of Shapes ---
  let shapeMap = {}; 
  
  slides.forEach(slide => {
    slide.getPageElements().forEach(element => {
      if (element.getPageElementType() === SlidesApp.PageElementType.SHAPE) {
        let shape = element.asShape();
        let title = shape.getTitle();
        
        if (title) {
          // If this title isn't in our map yet, create an empty list for it
          if (!shapeMap[title]) {
            shapeMap[title] = [];
          }
          // Add this shape to the list for this title
          shapeMap[title].push(shape);
        }
      }
    });
  });

  // --- STEP 2: Process Sheet Data ---
  for (let i = 1; i < data.length; i++) {
    let shapeName = data[i][0];                 // 0 = Column A, 1 = Column B etc
    let val = data[i][7].toString().trim();     // Column H

    // If the name from the sheet exists in our map
    if (shapeMap[shapeName]) {
      
      let hexColor = '#ffffff'; 
      switch(val) {
        case 'Purple': hexColor = '#ff00ff'; break;
        case 'Blue':   hexColor = '#0000ff'; break;
        case 'Orange': hexColor = '#ff9900'; break;
        case 'Green':  hexColor = '#00ff00'; break;
        default:              hexColor = '#ffffff';
      }

      // --- STEP 3: Loop through all shapes sharing that name ---
      shapeMap[shapeName].forEach(shape => {
        shape.getFill().setSolidFill(hexColor);
      });
    }
  }

  SpreadsheetApp.getActiveSpreadsheet().toast("Synced all matching shapes!", "Success", 5);
}

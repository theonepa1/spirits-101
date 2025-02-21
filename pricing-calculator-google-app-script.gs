/**
 * Sets up the Cost Analysis Tool UI on a single sheet.
 */
function setupCostAnalysisUI() {
  var ss = SpreadsheetApp.getActiveSpreadsheet();
  var sheetName = "Cost Analysis Tool";
  var sheet = ss.getSheetByName(sheetName);
  if (!sheet) {
    sheet = ss.insertSheet(sheetName);
  } else {
    sheet.clear();
  }

  // -----------------------------
  // Setup the Input Section (Columns A:C)
  // -----------------------------
  sheet.getRange("A1:C1").merge();
  sheet.getRange("A1").setValue("INPUTS")
      .setFontWeight("bold")
      .setFontSize(14)
      .setHorizontalAlignment("center")
      .setBackground("#d9ead3");

  var inputData = [
    ["Parameter", "Value", "Description"],
    ["Type of Spirit", "whiskey", "Select: whiskey, vodka, gin, rum"],
    ["Total Cases", 500, "Total number of cases (min 500)"],
    ["COGS", 50, "Cost of Goods Sold"],
    ["Shipping", 7, "Shipping cost"],
    ["Warehousing (Importer)", 7.5, "Importer warehousing cost"],
    ["Transportation", 0, "Transportation cost"],
    ["Import Tariffs (%)", 0, "Tariff percentage applied to COGS"],
    ["Misc Costs", 5000, "Total miscellaneous costs"],
    ["State Tax (per liter)", "Georgia", "Select a state"],
    ["Inland Transportation", 3.5, "Distributor inland transportation cost"],
    ["Warehousing (Distributor)", 1.5, "Distributor warehousing cost"],
    ["Distributor Margin (%)", 25, "Distributor margin percentage"],
    ["Retailer Margin (%)", 30, "Retailer margin percentage"],
    ["Retail Shelf Price", 0, "Final price per bottle at retail"]
  ];

  sheet.getRange("A2:C16").setValues(inputData);
  sheet.getRange("A2:C2").setFontWeight("bold").setBackground("#c9daf8");
  sheet.getRange("A2:C16").setBorder(true, true, true, true, true, true);

  var states = [
    "Alabama", "Alaska", "Arizona", "Arkansas", "California", "Colorado", 
    "Connecticut", "Delaware", "Florida", "Georgia", "Hawaii", "Idaho", 
    "Illinois", "Indiana", "Iowa", "Kansas", "Kentucky", "Louisiana", "Maine", 
    "Maryland", "Massachusetts", "Michigan", "Minnesota", "Mississippi", 
    "Missouri", "Montana", "Nebraska", "Nevada", "New Hampshire", "New Jersey", 
    "New Mexico", "New York", "North Carolina", "North Dakota", "Ohio", 
    "Oklahoma", "Oregon", "Pennsylvania", "Rhode Island", "South Carolina", 
    "South Dakota", "Tennessee", "Texas", "Utah", "Vermont", "Virginia", 
    "Washington", "West Virginia", "Wisconsin", "Wyoming"
  ];
  var stateValidation = SpreadsheetApp.newDataValidation()
      .requireValueInList(states, true)
      .build();
  sheet.getRange("B11").setDataValidation(stateValidation);

  // -----------------------------
  // Setup the Output Section (Columns E–F)
  // -----------------------------
  var sections = [
    { title: "Cost per Case", startRow: 2, data: [
      ["Federal Taxes", ""],
      ["Import Duty", ""],
      ["Distributor Costs", ""],
      ["Importer Total Cost", ""]
    ]},
    { title: "Price per Case", startRow: 8, data: [
      ["Importer Selling Price", ""],
      ["Distributor Selling Price", ""],
      ["Retailer Shelf Price", ""]
    ]},
    { title: "Price per Bottle", startRow: 13, data: [
      ["Importer Selling Price", ""],
      ["Distributor Selling Price", ""],
      ["Retailer Shelf Price", ""]
    ]},
    { title: "Profit Analysis", startRow: 18, data: [
      ["Importer Margin (%)", ""],
      ["Importer Profit per Case", ""],
      ["Importer Total Profit", ""]
    ]}
  ];

  sections.forEach(section => {
    sheet.getRange(`E${section.startRow}:F${section.startRow}`).merge();
    sheet.getRange(`E${section.startRow}`).setValue(section.title)
      .setFontWeight("bold").setFontSize(14).setHorizontalAlignment("center")
      .setBackground("#f4cccc");
    sheet.getRange(`E${section.startRow + 1}:F${section.startRow + section.data.length}`).setValues(section.data);
    sheet.getRange(`E${section.startRow}:F${section.startRow + section.data.length}`).setBorder(true, true, true, true, true, true);
  });

  SpreadsheetApp.flush();
}

/**
 * Calculates cost analysis based on inputs.
 */
function calculateCostAnalysisUI() {
  var sheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName("Cost Analysis Tool");
  if (!sheet) return;

  var inputs = {};
  var inputData = sheet.getRange("A3:C16").getValues();
  inputData.forEach(row => inputs[row[0]] = row[1]);

  var cases = Number(inputs["Total Cases"]);
  var cogs = Number(inputs["COGS"]);
  var shipping = Number(inputs["Shipping"]);
  var warehousingImporter = Number(inputs["Warehousing (Importer)"]);
  var transportation = Number(inputs["Transportation"]);
  var importTariffs = Number(inputs["Import Tariffs (%)"]) / 100;
  var miscCosts = Number(inputs["Misc Costs"]);
  var retailerPricePerBottle = Number(inputs["Retail Shelf Price"]);
  var retailerMargin = Number(inputs["Retailer Margin (%)"]) / 100;
  var distributorMargin = Number(inputs["Distributor Margin (%)"]) / 100;
  var inlandTransportation = Number(inputs["Inland Transportation"]);
  var warehousingDistributor = Number(inputs["Warehousing (Distributor)"]);

  var litersPerCase = 9;
  var federalTaxPerCase = litersPerCase * 0.40;
  var retailerPrice = retailerPricePerBottle * 12;
  var importDutyPerCase = cogs * importTariffs;
  var baseCost = cogs + shipping + warehousingImporter + transportation + 
                 (miscCosts / cases) + federalTaxPerCase + importDutyPerCase;
  var distributorCosts = inlandTransportation + warehousingDistributor;
  var distributorPrice = retailerPrice * (1 - retailerMargin);
  var importerPrice = (distributorPrice / (1 + distributorMargin)) - distributorCosts;
  var importerMargin = ((importerPrice - baseCost) / baseCost) * 100;
  var importerProfitPerCase = importerPrice - baseCost;
  var importerTotalProfit = importerProfitPerCase * cases;

  sheet.getRange("F3").setValue("$" + federalTaxPerCase.toFixed(2));
  sheet.getRange("F4").setValue("$" + importDutyPerCase.toFixed(2));
  sheet.getRange("F5").setValue("$" + distributorCosts.toFixed(2));
  sheet.getRange("F6").setValue("$" + baseCost.toFixed(2));
  sheet.getRange("F9").setValue("$" + importerPrice.toFixed(2));
  sheet.getRange("F10").setValue("$" + distributorPrice.toFixed(2));
  sheet.getRange("F11").setValue("$" + retailerPrice.toFixed(2));
  sheet.getRange("F14").setValue("$" + (importerPrice / 12).toFixed(2));
  sheet.getRange("F15").setValue("$" + (distributorPrice / 12).toFixed(2));
  sheet.getRange("F16").setValue("$" + (retailerPrice / 12).toFixed(2));
  sheet.getRange("F19").setValue(importerMargin.toFixed(2) + "%");
  sheet.getRange("F20").setValue("$" + importerProfitPerCase.toFixed(2));
  sheet.getRange("F21").setValue("$" + importerTotalProfit.toFixed(2));

  SpreadsheetApp.flush();
}

/**
 * Auto-recalculate when input values change.
 */
function onEdit(e) {
  var sheet = e.range.getSheet();
  if (sheet.getName() !== "Cost Analysis Tool") return;
  calculateCostAnalysisUI();
}

/**
 * Adds a menu to run setup and calculations.
 */
function onOpen() {
  SpreadsheetApp.getUi().createMenu("Cost Analysis UI")
    .addItem("Setup UI", "setupCostAnalysisUI")
    .addItem("Calculate", "calculateCostAnalysisUI")
    .addToUi();
}

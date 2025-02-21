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
    ["Total Cases", 4000, "Total number of cases (min 500)"],
    ["COGS", 37, "Cost of Goods Sold"],
    ["Shipping", 7, "Shipping cost"],
    ["Warehousing (Importer)", 7.5, "Importer warehousing cost"],
    ["Transportation", 0, "Transportation cost"],
    ["Import Tariffs (%)", 0, "Tariff percentage applied to COGS"],
    ["Misc Costs", 100000, "Total miscellaneous costs"],
    ["State Tax (per liter)", "Georgia", "Select a state"],
    ["Inland Transportation", 3.5, "Distributor inland transportation cost"],
    ["Warehousing (Distributor)", 5, "Distributor warehousing cost"],
    ["Distributor Margin (%)", 25, "Distributor margin percentage"],
    ["Retailer Margin (%)", 30, "Retailer margin percentage"],
    ["Retail Shelf Price", 18, "Final price per bottle at retail"]
  ];

  sheet.getRange("A2:C16").setValues(inputData);
  sheet.getRange("A2:C2").setFontWeight("bold").setBackground("#c9daf8");
  sheet.getRange("A2:C16").setBorder(true, true, true, true, true, true);

  // Set data validation for the state input (all 50 states)
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
  // Rearranged sections: Profit Analysis is now at the top.
  var sections = [
    { title: "Profit Analysis", startRow: 2, data: [
      ["Importer Margin (%)", ""],
      ["Importer Profit per Case", ""],
      ["Importer Total Profit", ""]
    ]},
    { title: "Cost per Case", startRow: 7, data: [
      ["Federal Taxes", ""],
      ["Import Duty", ""],
      ["State Tax", ""],
      ["Distributor Costs", ""],
      ["Importer Total Cost", ""]
    ]},
    { title: "Price per Case", startRow: 14, data: [
      ["Importer Selling Price", ""],
      ["Distributor Selling Price", ""],
      ["Retailer Shelf Price", ""]
    ]},
    { title: "Price per Bottle", startRow: 19, data: [
      ["Importer Selling Price", ""],
      ["Distributor Selling Price", ""],
      ["Retailer Shelf Price", ""]
    ]}
  ];

  sections.forEach(function(section) {
    sheet.getRange("E" + section.startRow + ":F" + section.startRow).merge();
    sheet.getRange("E" + section.startRow)
      .setValue(section.title)
      .setFontWeight("bold")
      .setFontSize(14)
      .setHorizontalAlignment("center")
      .setBackground("#f4cccc");
    sheet.getRange("E" + (section.startRow + 1) + ":F" + (section.startRow + section.data.length))
      .setValues(section.data);
    sheet.getRange("E" + section.startRow + ":F" + (section.startRow + section.data.length))
      .setBorder(true, true, true, true, true, true);
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
  inputData.forEach(function(row) {
    inputs[row[0]] = row[1];
  });

  // Retrieve and convert inputs
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

  // Assume each case contains 12 bottles of 0.75 liters each (total 9 liters per case)
  var litersPerCase = 9;

  // Calculate Federal tax using CBMA tax rate for distilled spirits.
  // CBMA tax rate is $2.7 per proof gallon, where one proof gallon is one gallon of spirit at 50% alcohol.
  // For our 40% v/v spirits, the conversion factor is 40/50 = 0.8.
  var gallonsPerLiter = 0.264172;
  var gallonsPerCase = litersPerCase * gallonsPerLiter;
  var proofGallonsPerCase = gallonsPerCase * (40 / 50);
  var federalTaxPerCase = proofGallonsPerCase * 2.7;

  // Use the provided state tax rate mapping (rates are per liter)
  var stateTaxRates = {
    "Alabama": 5.73, "Alaska": 3.38, "Arizona": 0.79, "Arkansas": 2.12,
    "California": 0.87, "Colorado": 0.60, "Connecticut": 1.57, "Delaware": 1.19,
    "Florida": 1.72, "Georgia": 1.00, "Hawaii": 1.58, "Idaho": 3.21,
    "Illinois": 2.26, "Indiana": 0.71, "Iowa": 3.73, "Kansas": 0.66,
    "Kentucky": 2.44, "Louisiana": 0.80, "Maine": 3.16, "Maryland": 1.44,
    "Massachusetts": 1.07, "Michigan": 3.59, "Minnesota": 2.30, "Mississippi": 2.25,
    "Missouri": 0.53, "Montana": 2.79, "Nebraska": 0.99, "Nevada": 0.95,
    "New Hampshire": 0.00, "New Jersey": 1.45, "New Mexico": 1.60, "New York": 1.70,
    "North Carolina": 4.33, "North Dakota": 1.24, "Ohio": 3.01, "Oklahoma": 1.47,
    "Oregon": 6.04, "Pennsylvania": 1.96, "Rhode Island": 1.43, "South Carolina": 1.43,
    "South Dakota": 1.29, "Tennessee": 1.18, "Texas": 0.63, "Utah": 4.21,
    "Vermont": 2.22, "Virginia": 5.83, "Washington": 9.66, "West Virginia": 2.20,
    "Wisconsin": 0.86, "Wyoming": 0.00
  };
  var stateSelected = inputs["State Tax (per liter)"];
  var stateTaxRate = stateTaxRates[stateSelected] || 0;
  var stateTaxPerCase = litersPerCase * stateTaxRate;

  var importDutyPerCase = cogs * importTariffs;
  var baseCost = cogs + shipping + warehousingImporter + transportation +
                 (miscCosts / cases) + federalTaxPerCase + importDutyPerCase + stateTaxPerCase;

  var distributorCosts = inlandTransportation + warehousingDistributor;
  // Retailer price per case (assuming 12 bottles per case)
  var retailerPrice = retailerPricePerBottle * 12;
  var distributorPrice = retailerPrice * (1 - retailerMargin);
  var importerPrice = (distributorPrice / (1 + distributorMargin)) - distributorCosts;
  var importerMargin = ((importerPrice - baseCost) / baseCost) * 100;
  var importerProfitPerCase = importerPrice - baseCost;
  var importerTotalProfit = importerProfitPerCase * cases;

  // -----------------------------
  // Set Output Values (updated cell addresses based on rearranged sections)
  // -----------------------------
  // Profit Analysis (Rows 2–5)
  sheet.getRange("F3").setValue(importerMargin.toFixed(2) + "%");
  sheet.getRange("F4").setValue("$" + importerProfitPerCase.toFixed(2));
  sheet.getRange("F5").setValue("$" + importerTotalProfit.toFixed(2));

  // Cost per Case (Rows 7–12)
  sheet.getRange("F8").setValue("$" + federalTaxPerCase.toFixed(2));
  sheet.getRange("F9").setValue("$" + importDutyPerCase.toFixed(2));
  sheet.getRange("F10").setValue("$" + stateTaxPerCase.toFixed(2));
  sheet.getRange("F11").setValue("$" + distributorCosts.toFixed(2));
  sheet.getRange("F12").setValue("$" + baseCost.toFixed(2));

  // Price per Case (Rows 14–17)
  sheet.getRange("F15").setValue("$" + importerPrice.toFixed(2));
  sheet.getRange("F16").setValue("$" + distributorPrice.toFixed(2));
  sheet.getRange("F17").setValue("$" + retailerPrice.toFixed(2));

  // Price per Bottle (Rows 19–22)
  sheet.getRange("F20").setValue("$" + (importerPrice / 12).toFixed(2));
  sheet.getRange("F21").setValue("$" + (distributorPrice / 12).toFixed(2));
  sheet.getRange("F22").setValue("$" + (retailerPrice / 12).toFixed(2));

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

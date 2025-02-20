/**
 * Sets up the Cost Analysis Tool UI on a single sheet.
 * Left side (columns A–C) holds the input table.
 * Right side (columns E–F) shows the outputs divided vertically
 * into four groups:
 *   1. Cost per Case
 *   2. Price per Case
 *   3. Price per Bottle
 *   4. Profit Analysis
 */
function setupCostAnalysisUI() {
  var ss = SpreadsheetApp.getActiveSpreadsheet();
  var sheetName = "Cost Analysis Tool";
  var sheet = ss.getSheetByName(sheetName);
  if (!sheet) {
    sheet = ss.insertSheet(sheetName);
  } else {
    sheet.clear(); // Clear previous data/formatting.
  }
  
  // -----------------------------
  // Setup the Input Section (Columns A:C)
  // -----------------------------
  // Merge A1:C1 for the INPUTS header.
  sheet.getRange("A1:C1").merge();
  sheet.getRange("A1").setValue("INPUTS")
      .setFontWeight("bold")
      .setFontSize(14)
      .setHorizontalAlignment("center")
      .setBackground("#d9ead3");
  
  // Input table data.
  // We'll update the rows as follows:
  // Row2: header
  // Row3: Type of Spirit
  // Row4: Total Cases
  // Row5: COGS
  // Row6: Shipping
  // Row7: Warehousing (Importer)
  // Row8: Transportation
  // Row9: Import Tariffs (%)  <-- NEW
  // Row10: Misc Costs
  // Row11: State Tax (per liter)
  // Row12: Inland Transportation
  // Row13: Warehousing (Distributor)
  // Row14: Importer Margin (%)
  // Row15: Distributor Margin (%)
  // Row16: Retailer Margin (%)
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
    ["Importer Margin (%)", 35, "Importer margin percentage"],
    ["Distributor Margin (%)", 25, "Distributor margin percentage"],
    ["Retailer Margin (%)", 30, "Retailer margin percentage"]
  ];
  // Write the input table starting at A2.
  sheet.getRange("A2:C16").setValues(inputData);
  sheet.getRange("A2:C2").setFontWeight("bold").setBackground("#c9daf8");
  sheet.setColumnWidth(1, 180);
  sheet.setColumnWidth(2, 120);
  sheet.setColumnWidth(3, 300);
  sheet.getRange("A2:C16").setBorder(true, true, true, true, true, true);
  
  // --- Data validation for dropdowns ---
  // "Type of Spirit" in B3.
  var spiritValidation = SpreadsheetApp.newDataValidation()
      .requireValueInList(["whiskey", "vodka", "gin", "rum"], true)
      .build();
  sheet.getRange("B3").setDataValidation(spiritValidation);
  
  // "State Tax (per liter)" in B11.
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
  // Setup the Output Section (Columns E–F, arranged vertically)
  // -----------------------------
  // We will arrange the four output groups vertically with spacing.
  // Group 1: Cost per Case (rows 2–6)
  // Group 2: Price per Case (rows 8–11)
  // Group 3: Price per Bottle (rows 13–16)
  // Group 4: Profit Analysis (rows 18–20)
  
  // Group 1 Header
  sheet.getRange("E2:F2").merge();
  sheet.getRange("E2").setValue("Cost per Case")
      .setFontWeight("bold")
      .setFontSize(14)
      .setHorizontalAlignment("center")
      .setBackground("#f4cccc");
  var group1Data = [
    ["Federal Taxes", ""],
    ["Import Duty", ""],
    ["Distributor Costs", ""],
    ["Importer Total Cost", ""]
  ];
  sheet.getRange("E3:F6").setValues(group1Data);
  sheet.getRange("E3:E6").setFontWeight("bold");
  sheet.getRange("E2:F6").setBorder(true, true, true, true, true, true);
  sheet.setRowHeights(2, 5, 25);
  
  // Group 2 Header (row 8)
  sheet.getRange("E8:F8").merge();
  sheet.getRange("E8").setValue("Price per Case")
      .setFontWeight("bold")
      .setFontSize(14)
      .setHorizontalAlignment("center")
      .setBackground("#d9d2e9");
  var group2Data = [
    ["Importer Selling Price", ""],
    ["Distributor Selling Price", ""],
    ["Retailer Shelf Price", ""]
  ];
  sheet.getRange("E9:F11").setValues(group2Data);
  sheet.getRange("E9:E11").setFontWeight("bold");
  sheet.getRange("E8:F11").setBorder(true, true, true, true, true, true);
  sheet.setRowHeights(8, 4, 25);
  
  // Group 3 Header (row 13)
  sheet.getRange("E13:F13").merge();
  sheet.getRange("E13").setValue("Price per Bottle")
      .setFontWeight("bold")
      .setFontSize(14)
      .setHorizontalAlignment("center")
      .setBackground("#d0e0e3");
  var group3Data = [
    ["Importer Selling Price", ""],
    ["Distributor Selling Price", ""],
    ["Retailer Shelf Price", ""]
  ];
  sheet.getRange("E14:F16").setValues(group3Data);
  sheet.getRange("E14:E16").setFontWeight("bold");
  sheet.getRange("E13:F16").setBorder(true, true, true, true, true, true);
  sheet.setRowHeights(13, 4, 25);
  
  // Group 4 Header (row 18)
  sheet.getRange("E18:F18").merge();
  sheet.getRange("E18").setValue("Profit Analysis")
      .setFontWeight("bold")
      .setFontSize(14)
      .setHorizontalAlignment("center")
      .setBackground("#f9cb9c");
  var group4Data = [
    ["Importer Profit per Case", ""],
    ["Importer Total Profit", ""]
  ];
  sheet.getRange("E19:F20").setValues(group4Data);
  sheet.getRange("E19:E20").setFontWeight("bold");
  sheet.getRange("E18:F20").setBorder(true, true, true, true, true, true);
  sheet.setRowHeights(18, 3, 25);
  
  SpreadsheetApp.flush();
  Logger.log("Cost Analysis UI has been set up with grouped outputs.");
}

/**
 * Reads inputs from the UI, performs the cost analysis calculation,
 * and writes the computed outputs (prefixed with "$") into the appropriate groups.
 */
function calculateCostAnalysisUI() {
  var ss = SpreadsheetApp.getActiveSpreadsheet();
  var sheet = ss.getSheetByName("Cost Analysis Tool");
  if (!sheet) {
    SpreadsheetApp.getUi().alert("Please run 'Setup UI' first.");
    return;
  }
  
  // Read input data from A3:C16.
  var inputData = sheet.getRange("A3:C16").getValues();
  var inputs = {};
  for (var i = 0; i < inputData.length; i++) {
    var key = inputData[i][0];
    var value = inputData[i][1];
    inputs[key] = value;
  }
  
  // Map parameters.
  var type = inputs["Type of Spirit"].toString().toLowerCase();
  var cases = Number(inputs["Total Cases"]);
  var cogs = Number(inputs["COGS"]);
  var shipping = Number(inputs["Shipping"]);
  var warehousingImporter = Number(inputs["Warehousing (Importer)"]);
  var transportation = Number(inputs["Transportation"]);
  var importTariffs = Number(inputs["Import Tariffs (%)"]); // New field
  var misc = Number(inputs["Misc Costs"]);
  
  // Look up state tax rate by state name.
  var selectedState = inputs["State Tax (per liter)"].toString();
  var stateTaxMapping = {
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
  var stateTax = Number(stateTaxMapping[selectedState] || 1.00);
  
  var inlandTransportation = Number(inputs["Inland Transportation"]);
  var warehousingDistributor = Number(inputs["Warehousing (Distributor)"]);
  var importerMargin = Number(inputs["Importer Margin (%)"]) / 100;
  var distributorMargin = Number(inputs["Distributor Margin (%)"]) / 100;
  var retailerMargin = Number(inputs["Retailer Margin (%)"]) / 100;
  
  // Constants and calculations.
  var litersPerCase = 9;
  var federalTaxPerCase = litersPerCase * 0.40;
  var CBP_RATES = {whiskey: 2.06, vodka: 1.78, gin: 2.38, rum: 1.59};
  // Updated Import Duty: now calculated as a percentage of COGS.
  var importDutyPerCase = cogs * (importTariffs / 100);
  
  // Base cost per case now includes import duty instead of the fixed CBP duty.
  var baseCost = cogs + shipping + warehousingImporter + transportation +
                 (misc / cases) + federalTaxPerCase + importDutyPerCase;
  var importerPrice = baseCost * (1 + importerMargin);
  var stateTaxPerCase = litersPerCase * stateTax;
  var distributorCosts = stateTaxPerCase + inlandTransportation + warehousingDistributor;
  var distributorPrice = (importerPrice + distributorCosts) * (1 + distributorMargin);
  var retailerPrice = distributorPrice / (1 - retailerMargin);
  var importerPriceBottle = importerPrice / 12;
  var distributorPriceBottle = distributorPrice / 12;
  var retailerPriceBottle = retailerPrice / 12;
  var importerProfitPerCase = importerPrice - baseCost;
  var importerTotalProfit = importerProfitPerCase * cases;
  
  // -----------------------------
  // Write outputs into each group, prefixing with "$".
  // -----------------------------
  
  // Group 1: Cost per Case.
  var group1Outputs = [
    ["$" + federalTaxPerCase.toFixed(2)],
    ["$" + importDutyPerCase.toFixed(2)],  // Updated label: Import Duty.
    ["$" + distributorCosts.toFixed(2)],
    ["$" + baseCost.toFixed(2)]
  ];
  sheet.getRange("F3:F6").setValues(group1Outputs);
  
  // Group 2: Price per Case.
  var group2Outputs = [
    ["$" + importerPrice.toFixed(2)],
    ["$" + distributorPrice.toFixed(2)],
    ["$" + retailerPrice.toFixed(2)]
  ];
  sheet.getRange("F9:F11").setValues(group2Outputs);
  
  // Group 3: Price per Bottle.
  var group3Outputs = [
    ["$" + importerPriceBottle.toFixed(2)],
    ["$" + distributorPriceBottle.toFixed(2)],
    ["$" + retailerPriceBottle.toFixed(2)]
  ];
  sheet.getRange("F14:F16").setValues(group3Outputs);
  
  // Group 4: Profit Analysis.
  var group4Outputs = [
    ["$" + importerProfitPerCase.toFixed(2)],
    ["$" + importerTotalProfit.toFixed(2)]
  ];
  sheet.getRange("F19:F20").setValues(group4Outputs);
  
  SpreadsheetApp.flush();
  Logger.log("Cost analysis calculation complete.");
}

/**
 * Automatically recalculate if changes are made within the input range.
 */
function onEdit(e) {
  var sheet = e.range.getSheet();
  if (sheet.getName() !== "Cost Analysis Tool") return;
  var inputRange = sheet.getRange("A3:C16");
  if (e.range.getRow() >= inputRange.getRow() &&
      e.range.getLastRow() <= inputRange.getLastRow() &&
      e.range.getColumn() >= inputRange.getColumn() &&
      e.range.getLastColumn() <= inputRange.getLastColumn()) {
    calculateCostAnalysisUI();
  }
}

/**
 * Adds a custom menu and triggers calculation on open.
 */
function onOpen() {
  SpreadsheetApp.getUi()
    .createMenu("Cost Analysis UI")
    .addItem("Setup UI", "setupCostAnalysisUI")
    .addItem("Calculate", "calculateCostAnalysisUI")
    .addToUi();
    
  var sheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName("Cost Analysis Tool");
  if (sheet) {
    calculateCostAnalysisUI();
  }
}

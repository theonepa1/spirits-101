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
  // 1) Setup the Input Section (A:C)
  // -----------------------------
  sheet.getRange("A1:C1").merge();
  sheet.getRange("A1")
      .setValue("INPUTS")
      .setFontWeight("bold")
      .setFontSize(14)
      .setHorizontalAlignment("center")
      .setBackground("#d9ead3");

  var inputData = [
    ["Parameter",               "Value",   "Description"],
    ["Type of Spirit",          "whiskey", "Select: whiskey, vodka, gin, rum"],
    ["Total Cases",             4000,      "Total number of cases (min 500)"],
    ["COGS",                    37,        "Cost of Goods Sold"],
    ["Shipping",                7,         "Shipping cost"],
    ["Warehousing (Importer)",  7.5,       "Importer warehousing cost"],
    ["Transportation",          0,         "Transportation cost"],
    ["Import Tariffs (%)",      0,         "Tariff percentage applied to COGS"],
    ["Misc Costs",              100000,    "Total miscellaneous costs"],
    ["State Tax (per liter)",   "Georgia", "Select a state"],
    ["Inland Transportation",   3.5,       "Distributor inland transportation"],
    ["Warehousing (Distributor)", 5,       "Distributor warehousing cost"],
    ["Distributor Markup (%)",  30,        "Distributor markup percentage"],
    ["Retailer Markup (%)",     30,        "Retailer markup percentage"],
    ["Retail Shelf Price",      18,        "Final price per bottle at retail"]
  ];
  sheet.getRange("A2:C16").setValues(inputData);

  // Slight formatting
  sheet.getRange("A2:C2").setFontWeight("bold").setBackground("#c9daf8");
  sheet.getRange("A2:C16").setBorder(true, true, true, true, true, true);

  // -----------------------------
  // 2) Setup the State Tax Table (columns AA–AB)
  //    We'll do a simple VLOOKUP later
  // -----------------------------
  // Put a label in AA1
  sheet.getRange("AA1").setValue("State Tax Table");
  
  // Full list of states + tax rates
  var stateTaxData = [
    ["Alabama", 5.73], ["Alaska", 3.38], ["Arizona", 0.79], ["Arkansas", 2.12],
    ["California", 0.87], ["Colorado", 0.60], ["Connecticut", 1.57], ["Delaware", 1.19],
    ["Florida", 1.72], ["Georgia", 1.00], ["Hawaii", 1.58], ["Idaho", 3.21],
    ["Illinois", 2.26], ["Indiana", 0.71], ["Iowa", 3.73], ["Kansas", 0.66],
    ["Kentucky", 2.44], ["Louisiana", 0.80], ["Maine", 3.16], ["Maryland", 1.44],
    ["Massachusetts", 1.07], ["Michigan", 3.59], ["Minnesota", 2.30], ["Mississippi", 2.25],
    ["Missouri", 0.53], ["Montana", 2.79], ["Nebraska", 0.99], ["Nevada", 0.95],
    ["New Hampshire", 0.00], ["New Jersey", 1.45], ["New Mexico", 1.60], ["New York", 1.70],
    ["North Carolina", 4.33], ["North Dakota", 1.24], ["Ohio", 3.01], ["Oklahoma", 1.47],
    ["Oregon", 6.04], ["Pennsylvania", 1.96], ["Rhode Island", 1.43], ["South Carolina", 1.43],
    ["South Dakota", 1.29], ["Tennessee", 1.18], ["Texas", 0.63], ["Utah", 4.21],
    ["Vermont", 2.22], ["Virginia", 5.83], ["Washington", 9.66], ["West Virginia", 2.20],
    ["Wisconsin", 0.86], ["Wyoming", 0.00]
  ];
  // Write this table starting at AA2
  sheet.getRange(2, 27, stateTaxData.length, 2).setValues(stateTaxData);

  // Provide data validation for state selection (B11) from our table in column AA
  var statesRange = sheet.getRange(2, 27, stateTaxData.length, 1); // AA2:AA51
  var stateValidation = SpreadsheetApp.newDataValidation()
      .requireValueInRange(statesRange, true)
      .build();
  sheet.getRange("B11").setDataValidation(stateValidation);

  // -----------------------------
  // 3) Setup the Output Section (E–F)
  // -----------------------------
  // We'll define four sections in the same structure as before.
  var sections = [
    {
      title: "Profit Analysis",
      startRow: 2,
      data: [
        ["Importer Markup (%)", ""],
        ["Importer Profit per Case", ""],
        ["Importer Total Profit", ""]
      ]
    },
    {
      title: "Cost per Case",
      startRow: 7,
      data: [
        ["Federal Taxes", ""],
        ["Import Duty", ""],
        ["State Tax", ""],
        ["Distributor Costs", ""],
        ["Importer Total Cost", ""]
      ]
    },
    {
      title: "Price per Case",
      startRow: 14,
      data: [
        ["FOB Price", ""],
        ["Landed Cost", ""],
        ["Wholesale Price", ""],
        ["Shelf Price", ""]
      ]
    },
    {
      title: "Price per Bottle",
      startRow: 19,
      data: [
        ["Importer Selling Price", ""],
        ["Distributor Selling Price", ""],
        ["Retailer Shelf Price", ""]
      ]
    }
  ];

  sections.forEach(function(section) {
    var startRow = section.startRow;
    var rows = section.data.length;
    // Title row
    sheet.getRange("E" + startRow + ":F" + startRow).merge();
    sheet.getRange("E" + startRow)
      .setValue(section.title)
      .setFontWeight("bold")
      .setFontSize(14)
      .setHorizontalAlignment("center")
      .setBackground("#f4cccc");

    // Data rows (labels)
    sheet.getRange(startRow + 1, 5, rows, 2).setValues(section.data);

    // Borders
    sheet.getRange(startRow, 5, rows + 1, 2)
      .setBorder(true, true, true, true, true, true);
  });

  // -----------------------------
  // 4) Insert the Spreadsheet Formulas
  //    All references are absolute so they update reliably.
  // -----------------------------

  // (A) Cost per Case items
  //
  // F8 (Federal Taxes):
  //   = 9 liters/case * 0.264172 gal/liter * (40/50) * 2.7
  sheet.getRange("F8").setFormula("=9 * 0.264172 * (40/50) * 2.7");

  // F9 (Import Duty):
  //   = B5 (COGS) * (B9 / 100) (Import Tariffs %)
  sheet.getRange("F9").setFormula("=$B$5 * ($B$9 / 100)");

  // F10 (State Tax):
  //   = 9 liters/case * VLOOKUP(B11, AA2:AB51, 2, FALSE)
  sheet.getRange("F10").setFormula("=9 * VLOOKUP($B$11,$AA$2:$AB$51,2,FALSE)");

  // F11 (Distributor Costs):
  //   = B12 + B13 (Inland Transportation + Warehousing (Distributor))
  sheet.getRange("F11").setFormula("=$B$12 + $B$13");

  // F12 (Importer Total Cost, i.e. baseCost):
  //   = B5 + B6 + B7 + B8 + (B10 / B4) + F8 + F9 + F10
  //   (COGS + Shipping + Warehousing(Importer) + Transportation + (MiscCosts / Cases) + FedTax + Duty + StateTax)
  sheet.getRange("F12").setFormula(
    "=$B$5 + $B$6 + $B$7 + $B$8 + ($B$10/$B$4) + F8 + F9 + F10"
  );

  // (B) Price per Case items
  //
  // F18 (Shelf Price): = B16 (Retail Shelf Price per bottle) * 12
  sheet.getRange("F18").setFormula("=$B$16 * 12");

  // F17 (Wholesale Price):
  //   The formula (distributor → retailer) = shelfPrice / (1 + retailerMarkup)
  //   = F18 / (1 + (B15 / 100))
  sheet.getRange("F17").setFormula("=F18 / (1 + ($B$15 / 100))");

  // F15 (FOB Price):
  //   The script does: importerPrice = (wholesalePrice / (1+ distributorMarkup)) - distributorCosts
  //   So = F17 / (1 + (B14 / 100)) - F11
  sheet.getRange("F15").setFormula("=F17 / (1 + ($B$14 / 100)) - F11");

  // F16 (Landed Cost):
  //   = F12 (baseCost) + F11 (distributorCosts)
  sheet.getRange("F16").setFormula("=F15 + F11");

  // (C) Profit Analysis
  //
  // F3 (Importer Markup (%)):
  //   = ((F15 - F12) / F12) * 100
  sheet.getRange("F3").setFormula("=((F15 - F12)/F12)*100");

  // F4 (Importer Profit per Case):
  //   = (FOB Price - baseCost)
  sheet.getRange("F4").setFormula("=F15 - F12");

  // F5 (Importer Total Profit):
  //   = (profit per case) * (total cases) => F4 * B4
  sheet.getRange("F5").setFormula("=F4 * $B$4");

  // (D) Price per Bottle (Columns E–F, rows 19–22)
  //    We’ll keep the original labeling, but set formulas accordingly:
  //
  // F20 = (FOB Price / 12)
  sheet.getRange("F20").setFormula("=F15 / 12");

  // F21 = (Wholesale Price / 12)
  sheet.getRange("F21").setFormula("=F17 / 12");

  // F22 = (Shelf Price / 12)
  sheet.getRange("F22").setFormula("=F18 / 12");

  // -----------------------------
  // 5) Formatting or finishing touches
  // -----------------------------
  // Optionally, force number formatting on these cells.
  // For example, set them to currency or 2-decimal places:
  var currencyCells = [
    "F4","F5","F8","F9","F10","F11","F12","F15","F16","F17","F18","F20","F21","F22"
  ];
  var percentageCells = ["F3"];
  currencyCells.forEach(function(cell){
    sheet.getRange(cell).setNumberFormat("$#,##0.00");
  });
  percentageCells.forEach(function(cell){
    sheet.getRange(cell).setNumberFormat("0.00\"%\"");
  });

  // Auto-fit columns
  sheet.autoResizeColumn(1);
  sheet.autoResizeColumn(2);
  sheet.autoResizeColumn(3);
  sheet.autoResizeColumn(5);
  sheet.autoResizeColumn(6);

  SpreadsheetApp.flush();
}

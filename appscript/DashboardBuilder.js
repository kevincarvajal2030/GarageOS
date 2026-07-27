

/**
 * ============================================================================
 * GarageOS - Dashboard Builder | DashboardBuilder.gs
 * ============================================================================
 * Native Google Sheets Dashboard layout generator.
 *
 * This module is responsible for generating the Dashboard layout inside
 * the 00_Dashboard worksheet using ONLY native Google Sheets features.
 *
 * No HTML, No CSS, No Sidebar, No Dialog.
 *
 * The Dashboard is fully rebuildable by executing: buildDashboard()
 * ============================================================================
 */

const DashboardBuilder = (() => {

  const DASHBOARD_SHEET = "00_Dashboard";

  // Color palette matching Excel design
  const COLORS = {
    // Backgrounds
    WHITE: "#FFFFFF",
    LIGHT_GRAY: "#F9FAFB",
    BORDER_GRAY: "#E5E7EB",

    // KPI Card backgrounds
    BLUE_LIGHT: "#EFF6FF",
    BLUE_ICON: "#DBEAFE",
    BLUE_TEXT: "#3B82F6",

    GREEN_LIGHT: "#ECFDF5",
    GREEN_ICON: "#D1FAE5",
    GREEN_TEXT: "#22C55E",

    ORANGE_LIGHT: "#FFF7ED",
    ORANGE_ICON: "#FFEDD5",
    ORANGE_TEXT: "#F59E0B",

    PURPLE_LIGHT: "#F5F3FF",
    PURPLE_ICON: "#EDE9FE",
    PURPLE_TEXT: "#8B5CF6",

    EMERALD_LIGHT: "#ECFDF5",
    EMERALD_ICON: "#D1FAE5",
    EMERALD_TEXT: "#16A34A",

    RED_LIGHT: "#FEF2F2",
    RED_ICON: "#FEE2E2",
    RED_TEXT: "#EF4444",

    // Text colors
    TEXT_PRIMARY: "#111827",
    TEXT_SECONDARY: "#6B7280",
    TEXT_DARK: "#1F2937",

    // Status colors
    STATUS_IN_PROGRESS_BG: "#DBEAFE",
    STATUS_IN_PROGRESS_TEXT: "#1E40AF",
    STATUS_WAITING_PARTS_BG: "#FEF3C7",
    STATUS_WAITING_PARTS_TEXT: "#92400E",
    STATUS_WAITING_APPROVAL_BG: "#EDE9FE",
    STATUS_WAITING_APPROVAL_TEXT: "#5B21B6",
    STATUS_PENDING_BG: "#FFEDD5",
    STATUS_PENDING_TEXT: "#9A3412",
    STATUS_ON_HOLD_BG: "#D1FAE5",
    STATUS_ON_HOLD_TEXT: "#065F46",
    STATUS_COMPLETED_BG: "#D1FAE5",
    STATUS_COMPLETED_TEXT: "#065F46",
    STATUS_CANCELLED_BG: "#FEE2E2",
    STATUS_CANCELLED_TEXT: "#991B1B",
    STATUS_PAID_BG: "#D1FAE5",
    STATUS_PAID_TEXT: "#065F46",
    STATUS_REFUNDED_BG: "#FEE2E2",
    STATUS_REFUNDED_TEXT: "#991B1B"
  };

  /**
   * Main entry point - builds the complete dashboard layout.
   */
  function buildDashboard() {
    const sheet = getSheet(DASHBOARD_SHEET);
    if (!sheet) {
      throw new Error("Dashboard sheet '00_Dashboard' not found");
    }

    // Clear existing content
    clearDashboard(sheet);

    // Build each section
    buildDashboardHeader(sheet);
    buildKPICards(sheet);
    buildChartsSection(sheet);
    buildTablesSection(sheet);
    buildVehicleGallery(sheet);
    buildFooter(sheet);

    // Apply final styling
    styleDashboard(sheet);

    showToast("Dashboard rebuilt successfully", "GarageOS");
  }

  /**
   * Clears all content from the dashboard sheet.
   * @param {Sheet} sheet
   */
  function clearDashboard(sheet) {
    const lastRow = sheet.getLastRow();
    const lastCol = sheet.getLastColumn();

    if (lastRow > 0 && lastCol > 0) {
      sheet.getRange(1, 1, Math.max(lastRow, 60), Math.max(lastCol, 26)).clearContent();
      sheet.getRange(1, 1, Math.max(lastRow, 60), Math.max(lastCol, 26)).clearFormatting();
    }

    // Reset column widths and row heights
    sheet.setColumnWidths(1, 26, 100);
    sheet.setRowHeights(1, 60, 21);
  }

  /**
   * Builds the dashboard header section.
   * @param {Sheet} sheet
   */
  function buildDashboardHeader(sheet) {
    // Title
    sheet.getRange("A1").setValue("Dashboard");
    sheet.getRange("A1").setFontWeight("bold");
    sheet.getRange("A1").setFontSize(28);
    sheet.getRange("A1").setFontColor(COLORS.TEXT_PRIMARY);

    // Subtitle
    sheet.getRange("A2").setValue("Overview of your garage operations");
    sheet.getRange("A2").setFontSize(14);
    sheet.getRange("A2").setFontColor(COLORS.TEXT_SECONDARY);

    // Last updated
    sheet.getRange("H1").setValue("Last Updated:");
    sheet.getRange("H1").setFontSize(12);
    sheet.getRange("H1").setFontColor(COLORS.TEXT_SECONDARY);
    sheet.getRange("H1").setHorizontalAlignment("right");

    sheet.getRange("I1").setFormula('=TEXT(NOW(),"MM/dd/yyyy HH:mm")');
    sheet.getRange("I1").setFontSize(12);
    sheet.getRange("I1").setFontColor(COLORS.TEXT_SECONDARY);
  }

  /**
   * Builds the KPI cards section with live data formulas.
   * @param {Sheet} sheet
   */
  function buildKPICards(sheet) {
    const startRow = 4;
    const cardHeight = 7;
    const cardWidth = 3;
    const gap = 1;

    // KPI definitions with data sources - corrected formulas
    const kpis = [
      {
        title: "Total Customers",
        formula: '=COUNTA(FILTER(01_Customers!K:K, 01_Customers!K:K<>""))',
        trendFormula: '=TEXT(COUNTA(FILTER(01_Customers!K:K, 01_Customers!K:K<>""))-COUNTA(FILTER(01_Customers!K:K, ROW(01_Customers!K:K)>ROW(01_Customers!K7)-30, 01_Customers!K:K<>"")), "+#;-#;0") & " vs last month"',
        icon: "👥",
        bg: COLORS.BLUE_LIGHT,
        iconBg: COLORS.BLUE_ICON,
        iconColor: COLORS.BLUE_TEXT
      },
      {
        title: "Total Vehicles",
        formula: '=COUNTA(FILTER(02_Vehicles!K:K, 02_Vehicles!K:K<>""))',
        trendFormula: '=TEXT(COUNTA(FILTER(02_Vehicles!K:K, 02_Vehicles!K:K<>""))-COUNTA(FILTER(02_Vehicles!K:K, ROW(02_Vehicles!K:K)>ROW(02_Vehicles!K7)-30, 02_Vehicles!K:K<>"")), "+#;-#;0") & " vs last month"',
        icon: "🚗",
        bg: COLORS.GREEN_LIGHT,
        iconBg: COLORS.GREEN_ICON,
        iconColor: COLORS.GREEN_TEXT
      },
      {
        title: "Open Work Orders",
        formula: '=COUNTIFS(03_Work_Orders!I:I, "In Progress")+COUNTIFS(03_Work_Orders!I:I, "Waiting Parts")+COUNTIFS(03_Work_Orders!I:I, "Waiting Approval")+COUNTIFS(03_Work_Orders!I:I, "Pending")+COUNTIFS(03_Work_Orders!I:I, "On Hold")',
        trendFormula: '""',
        icon: "🔧",
        bg: COLORS.ORANGE_LIGHT,
        iconBg: COLORS.ORANGE_ICON,
        iconColor: COLORS.ORANGE_TEXT
      },
      {
        title: "Completed Orders",
        formula: '=COUNTIF(03_Work_Orders!I:I, "Completed")',
        trendFormula: '=TEXT(COUNTIF(03_Work_Orders!I:I, "Completed")-COUNTIFS(03_Work_Orders!I:I, "Completed", 03_Work_Orders!R:R, ">"&TODAY()-30), "+#;-#;0") & " this month"',
        icon: "✅",
        bg: COLORS.PURPLE_LIGHT,
        iconBg: COLORS.PURPLE_ICON,
        iconColor: COLORS.PURPLE_TEXT
      },
      {
        title: "Total Revenue",
        formula: '=SUMIF(08_Payments!G:G, "Paid", 08_Payments!F:F)',
        trendFormula: '=TEXT(SUMIF(08_Payments!G:G, "Paid", 08_Payments!F:F)-SUMIFS(08_Payments!F:F, 08_Payments!G:G, "Paid", 08_Payments!H:H, ">"&TODAY()-30), "+$#,##0;-$#,##0;$0") & " vs last month"',
        icon: "💰",
        bg: COLORS.EMERALD_LIGHT,
        iconBg: COLORS.EMERALD_ICON,
        iconColor: COLORS.EMERALD_TEXT,
        numberFormat: "$#,##0"
      },
      {
        title: "Outstanding Balance",
        formula: '=SUMIF(08_Payments!G:G, "Pending", 08_Payments!F:F)+SUMIF(08_Payments!G:G, "Unpaid", 08_Payments!F:F)',
        trendFormula: '""',
        icon: "⏳",
        bg: COLORS.RED_LIGHT,
        iconBg: COLORS.RED_ICON,
        iconColor: COLORS.RED_TEXT,
        numberFormat: "$#,##0"
      }
    ];

    for (let i = 0; i < kpis.length; i++) {
      const kpi = kpis[i];
      const cardStartRow = startRow + Math.floor(i / 6) * (cardHeight + gap);
      const cardStartCol = 1 + (i % 6) * (cardWidth + gap);

      // Create merged cell for card background
      const cardRange = sheet.getRange(cardStartRow, cardStartCol, cardHeight, cardWidth);
      cardRange.setBackground(kpi.bg);
      cardRange.setBorder(true, true, true, true, true, true, COLORS.BORDER_GRAY, SpreadsheetApp.BorderStyle.SOLID);
      cardRange.setBorderStyle(SpreadsheetApp.BorderStyle.ROUNDED);

      // Icon area (merged) - larger icon
      const iconRange = sheet.getRange(cardStartRow + 1, cardStartCol, 2, 1);
      iconRange.merge();
      iconRange.setBackground(kpi.iconBg);
      iconRange.setFontColor(kpi.iconColor);
      iconRange.setHorizontalAlignment("center");
      iconRange.setVerticalAlignment("middle");
      iconRange.setFontSize(24);
      iconRange.setValue(kpi.icon);

      // Value (merged) - with formula, larger font
      const valueRange = sheet.getRange(cardStartRow + 1, cardStartCol + 1, 2, 2);
      valueRange.merge();
      valueRange.setFormula(kpi.formula);
      valueRange.setNumberFormat(kpi.numberFormat || "#,##0");
      valueRange.setFontSize(32);
      valueRange.setFontWeight("bold");
      valueRange.setFontColor(COLORS.TEXT_PRIMARY);
      valueRange.setHorizontalAlignment("right");
      valueRange.setVerticalAlignment("bottom");

      // Title
      const titleRange = sheet.getRange(cardStartRow + 3, cardStartCol, 1, cardWidth);
      titleRange.merge();
      titleRange.setValue(kpi.title);
      titleRange.setFontSize(12);
      titleRange.setFontColor(COLORS.TEXT_SECONDARY);
      titleRange.setHorizontalAlignment("left");
      titleRange.setVerticalAlignment("bottom");
      titleRange.setFontWeight("bold");

      // Trend with percentage
      const trendRange = sheet.getRange(cardStartRow + 4, cardStartCol, 1, cardWidth);
      trendRange.merge();
      trendRange.setFormula(kpi.trendFormula);
      trendRange.setFontSize(11);
      trendRange.setFontColor(COLORS.TEXT_SECONDARY);
      trendRange.setHorizontalAlignment("left");
      trendRange.setVerticalAlignment("top");
    }
  }

  /**
   * Builds the charts section with native Google Sheets charts.
   * Improved layout matching Dashboard-example.png design reference.
   * @param {Sheet} sheet
   */
  function buildChartsSection(sheet) {
    const startRow = 13;

    // Section title with border
    sheet.getRange(startRow, 1).setValue("Analytics");
    sheet.getRange(startRow, 1).setFontWeight("bold");
    sheet.getRange(startRow, 1).setFontSize(16);
    sheet.getRange(startRow, 1).setFontColor(COLORS.TEXT_PRIMARY);

    // Create data ranges for charts (hidden area)
    const dataStartRow = 55;

    // Revenue chart data area
    sheet.getRange(dataStartRow, 1).setValue("Revenue Data");
    sheet.getRange(dataStartRow + 1, 1).setValue("Month");
    sheet.getRange(dataStartRow + 1, 2).setValue("Revenue");

    // Vehicle make chart data area
    sheet.getRange(dataStartRow, 10).setValue("Vehicle Make Data");
    sheet.getRange(dataStartRow + 1, 10).setValue("Make");
    sheet.getRange(dataStartRow + 1, 11).setValue("Count");

    // Work order status chart data area
    sheet.getRange(dataStartRow, 15).setValue("WO Status Data");
    sheet.getRange(dataStartRow + 1, 15).setValue("Status");
    sheet.getRange(dataStartRow + 1, 16).setValue("Count");

    // Chart areas with card styling
    // Revenue chart area - larger, prominent position
    const revenueChartArea = sheet.getRange(startRow + 2, 1, 12, 12);
    revenueChartArea.setBorder(true, true, true, true, true, true, COLORS.BORDER_GRAY, SpreadsheetApp.BorderStyle.SOLID);
    revenueChartArea.setBackground(COLORS.WHITE);

    // Revenue chart title inside card
    sheet.getRange(startRow + 2, 1).setValue("Revenue Trend (6 months)");
    sheet.getRange(startRow + 2, 1).setFontWeight("bold");
    sheet.getRange(startRow + 2, 1).setFontSize(14);
    sheet.getRange(startRow + 2, 1).setFontColor(COLORS.TEXT_PRIMARY);

    // Vehicle make chart area
    const vehicleChartArea = sheet.getRange(startRow + 2, 14, 12, 6);
    vehicleChartArea.setBorder(true, true, true, true, true, true, COLORS.BORDER_GRAY, SpreadsheetApp.BorderStyle.SOLID);
    vehicleChartArea.setBackground(COLORS.WHITE);

    sheet.getRange(startRow + 2, 14).setValue("Vehicles by Make");
    sheet.getRange(startRow + 2, 14).setFontWeight("bold");
    sheet.getRange(startRow + 2, 14).setFontSize(14);
    sheet.getRange(startRow + 2, 14).setFontColor(COLORS.TEXT_PRIMARY);

    // Work order status chart area
    const woChartArea = sheet.getRange(startRow + 2, 21, 12, 6);
    woChartArea.setBorder(true, true, true, true, true, true, COLORS.BORDER_GRAY, SpreadsheetApp.BorderStyle.SOLID);
    woChartArea.setBackground(COLORS.WHITE);

    sheet.getRange(startRow + 2, 21).setValue("Work Order Status");
    sheet.getRange(startRow + 2, 21).setFontWeight("bold");
    sheet.getRange(startRow + 2, 21).setFontSize(14);
    sheet.getRange(startRow + 2, 21).setFontColor(COLORS.TEXT_PRIMARY);

    // Build actual charts using Apps Script ChartBuilder
    buildRevenueChart(sheet, startRow + 3, 1);
    buildVehicleMakeChart(sheet, startRow + 3, 14);
    buildWorkOrderStatusChart(sheet, startRow + 3, 21);
  }

  /**
   * Builds the revenue trend line chart.
   * @param {Sheet} sheet
   * @param {number} anchorRow
   * @param {number} anchorCol
   */
  function buildRevenueChart(sheet, anchorRow, anchorCol) {
    // Remove existing charts in this area
    const charts = sheet.getCharts();
    for (const chart of charts) {
      const pos = chart.getPosition();
      if (pos.getRow() >= anchorRow - 1 && pos.getRow() <= anchorRow + 12 &&
        pos.getColumn() >= anchorCol && pos.getColumn() <= anchorCol + 11) {
        sheet.removeChart(chart);
      }
    }

    // Create a new chart builder
    const chartBuilder = sheet.newChart()
      .setChartType(Charts.ChartType.LINE)
      .addRange(sheet.getRange("A56:B61"))
      .setPosition(anchorRow, anchorCol, 0, 0)
      .setOption('title', '')
      .setOption('legend', { position: 'bottom' })
      .setOption('colors', ['#3B82F6'])
      .setOption('curveType', 'function')
      .setOption('pointSize', 5)
      .setOption('chartArea', { width: '90%', height: '80%' })
      .setOption('hAxis', { textStyle: { fontSize: 10 } })
      .setOption('vAxis', {
        textStyle: { fontSize: 10 },
        format: '$#,##0'
      })
      .setOption('height', 280)
      .setOption('width', 400);

    try {
      sheet.insertChart(chartBuilder.build());
    } catch (e) {
      // Chart creation failed - data may not be ready yet
      console.log("Revenue chart creation deferred: " + e.message);
    }
  }

  /**
   * Builds the vehicle make doughnut chart.
   * @param {Sheet} sheet
   * @param {number} anchorRow
   * @param {number} anchorCol
   */
  function buildVehicleMakeChart(sheet, anchorRow, anchorCol) {
    // Remove existing charts in this area
    const charts = sheet.getCharts();
    for (const chart of charts) {
      const pos = chart.getPosition();
      if (pos.getRow() >= anchorRow - 1 && pos.getRow() <= anchorRow + 12 &&
        pos.getColumn() >= anchorCol && pos.getColumn() <= anchorCol + 5) {
        sheet.removeChart(chart);
      }
    }

    const chartBuilder = sheet.newChart()
      .setChartType(Charts.ChartType.PIE)
      .addRange(sheet.getRange("J56:K61"))
      .setPosition(anchorRow, anchorCol, 0, 0)
      .setOption('title', '')
      .setOption('pieHole', 0.5)
      .setOption('legend', { position: 'right', alignment: 'center' })
      .setOption('colors', ['#3B82F6', '#22C55E', '#F59E0B', '#EF4444', '#8B5CF6', '#F97316'])
      .setOption('chartArea', { width: '100%', height: '90%' })
      .setOption('height', 280)
      .setOption('width', 200);

    try {
      sheet.insertChart(chartBuilder.build());
    } catch (e) {
      console.log("Vehicle make chart creation deferred: " + e.message);
    }
  }

  /**
   * Builds the work order status doughnut chart.
   * @param {Sheet} sheet
   * @param {number} anchorRow
   * @param {number} anchorCol
   */
  function buildWorkOrderStatusChart(sheet, anchorRow, anchorCol) {
    // Remove existing charts in this area
    const charts = sheet.getCharts();
    for (const chart of charts) {
      const pos = chart.getPosition();
      if (pos.getRow() >= anchorRow - 1 && pos.getRow() <= anchorRow + 12 &&
        pos.getColumn() >= anchorCol && pos.getColumn() <= anchorCol + 5) {
        sheet.removeChart(chart);
      }
    }

    const chartBuilder = sheet.newChart()
      .setChartType(Charts.ChartType.PIE)
      .addRange(sheet.getRange("O56:P62"))
      .setPosition(anchorRow, anchorCol, 0, 0)
      .setOption('title', '')
      .setOption('pieHole', 0.5)
      .setOption('legend', { position: 'right', alignment: 'center' })
      .setOption('colors', ['#3B82F6', '#F59E0B', '#8B5CF6', '#F97316', '#22C55E', '#10B981', '#EF4444'])
      .setOption('chartArea', { width: '100%', height: '90%' })
      .setOption('height', 280)
      .setOption('width', 200);

    try {
      sheet.insertChart(chartBuilder.build());
    } catch (e) {
      console.log("Work order status chart creation deferred: " + e.message);
    }

  }

  /**
   * Builds the tables section with native Google Sheets tables.
   * Improved layout matching Dashboard-example.png design reference.
   * @param {Sheet} sheet
   */
  function buildTablesSection(sheet) {
    const startRow = 27;

    // Recent Work Orders table - card style
    sheet.getRange(startRow, 1).setValue("Recent Work Orders");
    sheet.getRange(startRow, 1).setFontWeight("bold");
    sheet.getRange(startRow, 1).setFontSize(14);
    sheet.getRange(startRow, 1).setFontColor(COLORS.TEXT_PRIMARY);


    // Table container with border
    const woContainer = sheet.getRange(startRow, 1, 12, 6);
    woContainer.setBorder(true, true, true, true, true, true, COLORS.BORDER_GRAY, SpreadsheetApp.BorderStyle.SOLID);
    woContainer.setBackground(COLORS.WHITE);

    // Table headers
    const woHeaders = [["WO ID", "Customer", "Vehicle", "Status", "Total", "Date"]];
    sheet.getRange(startRow + 1, 1, 1, 6).setValues(woHeaders);
    sheet.getRange(startRow + 1, 1, 1, 6).setFontWeight("bold");
    sheet.getRange(startRow + 1, 1, 1, 6).setFontSize(12);
    sheet.getRange(startRow + 1, 1, 1, 6).setFontColor(COLORS.TEXT_SECONDARY);
    sheet.getRange(startRow + 1, 1, 1, 6).setBackground(COLORS.LIGHT_GRAY);

    // Data range with alternating row colors and borders
    const woDataRange = sheet.getRange(startRow + 2, 1, 9, 6);
    woDataRange.setBorder(false, false, false, false, true, false, COLORS.BORDER_GRAY, SpreadsheetApp.BorderStyle.SOLID);

    // Recent Payments table - card style
    sheet.getRange(startRow, 8).setValue("Recent Payments");
    sheet.getRange(startRow, 8).setFontWeight("bold");
    sheet.getRange(startRow, 8).setFontSize(14);
    sheet.getRange(startRow, 8).setFontColor(COLORS.TEXT_PRIMARY);

    // Table container with border
    const payContainer = sheet.getRange(startRow, 8, 12, 6);
    payContainer.setBorder(true, true, true, true, true, true, COLORS.BORDER_GRAY, SpreadsheetApp.BorderStyle.SOLID);
    payContainer.setBackground(COLORS.WHITE);

    // Table headers
    const payHeaders = [["Payment ID", "Customer", "Amount", "Method", "Date", "Status"]];
    sheet.getRange(startRow + 1, 8, 1, 6).setValues(payHeaders);
    sheet.getRange(startRow + 1, 8, 1, 6).setFontWeight("bold");
    sheet.getRange(startRow + 1, 8, 1, 6).setFontSize(11);
    sheet.getRange(startRow + 1, 8, 1, 6).setFontColor(COLORS.TEXT_SECONDARY);
    sheet.getRange(startRow + 1, 8, 1, 6).setBackground(COLORS.LIGHT_GRAY);

    // Data range with alternating row colors and borders
    const payDataRange = sheet.getRange(startRow + 2, 8, 9, 6);
    payDataRange.setBorder(false, false, false, false, true, false, COLORS.BORDER_GRAY, SpreadsheetApp.BorderStyle.SOLID);

    // Top Customers table - card style (new addition)
    sheet.getRange(startRow, 15).setValue("Top Customers");
    sheet.getRange(startRow, 15).setFontWeight("bold");
    sheet.getRange(startRow, 15).setFontSize(14);
    sheet.getRange(startRow, 15).setFontColor(COLORS.TEXT_PRIMARY);

    // Table container with border
    const custContainer = sheet.getRange(startRow, 15, 12, 5);
    custContainer.setBorder(true, true, true, true, true, true, COLORS.BORDER_GRAY, SpreadsheetApp.BorderStyle.SOLID);
    custContainer.setBackground(COLORS.WHITE);

    // Table headers
    const custHeaders = [["Customer", "Orders", "Total Spent"]];
    sheet.getRange(startRow + 1, 15, 1, 3).setValues(custHeaders);
    sheet.getRange(startRow + 1, 15, 1, 3).setFontWeight("bold");
    sheet.getRange(startRow + 1, 15, 1, 3).setFontSize(11);
    sheet.getRange(startRow + 1, 15, 1, 3).setFontColor(COLORS.TEXT_SECONDARY);
    sheet.getRange(startRow + 1, 15, 1, 3).setBackground(COLORS.LIGHT_GRAY);

    // Data range
    const custDataRange = sheet.getRange(startRow + 2, 15, 9, 3);
    custDataRange.setBorder(false, false, false, false, true, false, COLORS.BORDER_GRAY, SpreadsheetApp.BorderStyle.SOLID);
  }

  /**
   * Builds the vehicle gallery section.
   * DISABLED - Temporarily removed per requirements.
   * Code preserved for future implementation.
   * @param {Sheet} sheet
   */
  function buildVehicleGallery(sheet) {
    /*
 
     const startRow = 36;
 
     // Section title
     sheet.getRange(startRow, 1).setValue("Vehicle Gallery");
     sheet.getRange(startRow, 1).setFontWeight("bold");
     sheet.getRange(startRow, 1).setFontSize(14);
     sheet.getRange(startRow, 1).setFontColor(COLORS.TEXT_PRIMARY);
 
     // Gallery grid - 6 columns x 2 rows = 12 vehicles
     const cardWidth = 3;
     const cardHeight = 4;
     const gap = 1;
 
     for (let row = 0; row < 2; row++) {
       for (let col = 0; col < 6; col++) {
         const startCellRow = startRow + 2 + row * (cardHeight + gap);
         const startCellCol = 1 + col * (cardWidth + gap);
 
         // Vehicle card container
         const cardRange = sheet.getRange(startCellRow, startCellCol, cardHeight, cardWidth);
         cardRange.setBackground(COLORS.WHITE);
         cardRange.setBorder(true, true, true, true, true, true, COLORS.BORDER_GRAY, SpreadsheetApp.BorderStyle.SOLID);
 
         // Image placeholder
         const imageRange = sheet.getRange(startCellRow, startCellCol, 2, cardWidth);
         imageRange.merge();
         imageRange.setBackground(COLORS.LIGHT_GRAY);
         imageRange.setHorizontalAlignment("center");
         imageRange.setVerticalAlignment("middle");
         imageRange.setValue("[Image]");
         imageRange.setFontColor(COLORS.TEXT_SECONDARY);
         imageRange.setFontSize(10);
 
         // Vehicle info
         const infoRange = sheet.getRange(startCellRow + 2, startCellCol, 2, cardWidth);
         infoRange.merge();
         infoRange.setHorizontalAlignment("left");
         infoRange.setVerticalAlignment("top");
         infoRange.setFontSize(11);
         infoRange.setWrap(true);
       }
     }*/
  }

  /**
   * Builds the footer section.
   * @param {Sheet} sheet
   */
  function buildFooter(sheet) {
    const footerRow = 50;

    sheet.getRange(footerRow, 1).setValue("GarageOS v1.0 | Built with Google Sheets & Apps Script");
    sheet.getRange(footerRow, 1).setFontSize(10);
    sheet.getRange(footerRow, 1).setFontColor(COLORS.TEXT_SECONDARY);
    sheet.getRange(footerRow, 1).setHorizontalAlignment("center");
  }

  /**
   * Applies final styling to the dashboard.
   * @param {Sheet} sheet
   */
  function styleDashboard(sheet) {
    // Set column widths - optimized for dashboard layout
    sheet.setColumnWidth(1, 140);
    sheet.setColumnWidth(2, 140);
    sheet.setColumnWidth(3, 140);
    sheet.setColumnWidth(4, 140);
    sheet.setColumnWidth(5, 140);
    sheet.setColumnWidth(6, 140);
    sheet.setColumnWidth(7, 40); // gap
    sheet.setColumnWidth(8, 140);
    sheet.setColumnWidth(9, 140);
    sheet.setColumnWidth(10, 140);
    sheet.setColumnWidth(11, 140);
    sheet.setColumnWidth(12, 140);
    sheet.setColumnWidth(13, 40); // gap
    sheet.setColumnWidth(14, 140);
    sheet.setColumnWidth(15, 140);
    sheet.setColumnWidth(16, 140);
    sheet.setColumnWidth(17, 140);
    sheet.setColumnWidth(18, 140);
    sheet.setColumnWidth(19, 140);
    sheet.setColumnWidth(20, 140);
    sheet.setColumnWidth(21, 140);
    sheet.setColumnWidth(22, 140);
    sheet.setColumnWidth(23, 140);
    sheet.setColumnWidth(24, 140);
    sheet.setColumnWidth(25, 40); // gap

    // Set row heights - optimized for visual hierarchy
    sheet.setRowHeight(1, 50);  // Title row
    sheet.setRowHeight(2, 30);  // Subtitle row
    sheet.setRowHeight(3, 10);  // Spacer

    // KPI card rows
    for (let row = 4; row <= 10; row++) {
      sheet.setRowHeight(row, 35);
    }

    // Analytics section rows
    for (let row = 13; row <= 26; row++) {
      sheet.setRowHeight(row, 30);
    }

    // Tables section rows
    for (let row = 27; row <= 40; row++) {
      sheet.setRowHeight(row, 28);
    }

    // Hide gridlines for cleaner look
    sheet.setHiddenGridlines(true);

    // Set default font for entire sheet
    sheet.getRange("A1:Z100").setFontFamily("Roboto");
  }

  return {
    buildDashboard,
    buildDashboardLayout: buildDashboard,
    buildDashboardHeader,
    buildKPICards,
    buildChartsSection,
    buildTablesSection,
    buildVehicleGallery,
    buildFooter,
    clearDashboard,
    styleDashboard
  };

})();

/**
 * Global wrapper function for menu access.
 */
function buildDashboard() {
  DashboardBuilder.buildDashboard();
}
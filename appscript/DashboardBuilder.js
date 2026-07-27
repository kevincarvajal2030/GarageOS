

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
    const cardHeight = 6;
    const cardWidth = 3;
    const gap = 1;

    // KPI definitions with data sources
    const kpis = [
      {
        title: "Total Customers",
        formula: '=COUNTA(FILTER(01_Customers!K:K, 01_Customers!K:K<>""))',
        icon: "👥",
        bg: COLORS.BLUE_LIGHT,
        iconBg: COLORS.BLUE_ICON,
        iconColor: COLORS.BLUE_TEXT
      },
      {
        title: "Total Vehicles",
        formula: '=COUNTA(FILTER(02_Vehicles!K:K, 02_Vehicles!K:K<>""))',
        icon: "🚗",
        bg: COLORS.GREEN_LIGHT,
        iconBg: COLORS.GREEN_ICON,
        iconColor: COLORS.GREEN_TEXT
      },
      {
        title: "Open Work Orders",
        formula: '=COUNTIFS(03_Work_Orders!I:I, "In Progress", 03_Work_Orders!I:I, "Waiting Parts", 03_Work_Orders!I:I, "Waiting Approval", 03_Work_Orders!I:I, "Pending", 03_Work_Orders!I:I, "On Hold")',
        icon: "🔧",
        bg: COLORS.ORANGE_LIGHT,
        iconBg: COLORS.ORANGE_ICON,
        iconColor: COLORS.ORANGE_TEXT
      },
      {
        title: "Completed Orders",
        formula: '=COUNTIF(03_Work_Orders!I:I, "Completed")',
        icon: "✅",
        bg: COLORS.PURPLE_LIGHT,
        iconBg: COLORS.PURPLE_ICON,
        iconColor: COLORS.PURPLE_TEXT
      },
      {
        title: "Total Revenue",
        formula: '=SUMIF(08_Payments!G:G, "Paid", 08_Payments!F:F)',
        icon: "💰",
        bg: COLORS.EMERALD_LIGHT,
        iconBg: COLORS.EMERALD_ICON,
        iconColor: COLORS.EMERALD_TEXT
      },
      {
        title: "Outstanding Balance",
        formula: '=SUMIF(08_Payments!G:G, "Pending", 08_Payments!F:F)',
        icon: "⏳",
        bg: COLORS.RED_LIGHT,
        iconBg: COLORS.RED_ICON,
        iconColor: COLORS.RED_TEXT
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

      // Icon area (merged)
      const iconRange = sheet.getRange(cardStartRow, cardStartCol, 2, 1);
      iconRange.merge();
      iconRange.setBackground(kpi.iconBg);
      iconRange.setFontColor(kpi.iconColor);
      iconRange.setHorizontalAlignment("center");
      iconRange.setVerticalAlignment("middle");
      iconRange.setFontSize(20);
      iconRange.setValue(kpi.icon);

      // Value (merged) - with formula
      const valueRange = sheet.getRange(cardStartRow, cardStartCol + 1, 2, 2);
      valueRange.merge();
      valueRange.setFormula(kpi.formula);
      valueRange.setNumberFormat("#,##0");
      valueRange.setFontSize(28);
      valueRange.setFontWeight("bold");
      valueRange.setFontColor(COLORS.TEXT_PRIMARY);
      valueRange.setHorizontalAlignment("left");
      valueRange.setVerticalAlignment("bottom");

      // Title
      const titleRange = sheet.getRange(cardStartRow + 2, cardStartCol, 1, cardWidth);
      titleRange.merge();
      titleRange.setValue(kpi.title);
      titleRange.setFontSize(13);
      titleRange.setFontColor(COLORS.TEXT_SECONDARY);
      titleRange.setHorizontalAlignment("left");
      titleRange.setVerticalAlignment("bottom");

      // Trend placeholder
      const trendRange = sheet.getRange(cardStartRow + 3, cardStartCol, 1, cardWidth);
      trendRange.merge();
      trendRange.setFontSize(12);
      trendRange.setFontWeight("bold");
      trendRange.setHorizontalAlignment("left");
    }
  }

  /**
   * Builds the charts section with native Google Sheets charts.
   * @param {Sheet} sheet
   */
  function buildChartsSection(sheet) {
    const startRow = 12;

    // Section title
    sheet.getRange(startRow, 1).setValue("Analytics");
    sheet.getRange(startRow, 1).setFontWeight("bold");
    sheet.getRange(startRow, 1).setFontSize(16);
    sheet.getRange(startRow, 1).setFontColor(COLORS.TEXT_PRIMARY);

    // Create data ranges for charts (hidden area)
    const dataStartRow = 50;

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

    // Chart placeholders - will be populated with actual charts
    // Revenue chart area
    const revenueChartArea = sheet.getRange(startRow + 2, 1, 10, 8);
    revenueChartArea.merge();
    sheet.getRange(startRow + 2, 1).setValue("Revenue Trend (6 months)");
    sheet.getRange(startRow + 2, 1).setFontWeight("bold");
    sheet.getRange(startRow + 2, 1).setFontSize(14);

    // Vehicle make chart area
    const vehicleChartArea = sheet.getRange(startRow + 2, 10, 10, 4);
    vehicleChartArea.merge();
    sheet.getRange(startRow + 2, 10).setValue("Vehicles by Make");
    sheet.getRange(startRow + 2, 10).setFontWeight("bold");
    sheet.getRange(startRow + 2, 10).setFontSize(14);

    // Work order status chart area
    const woChartArea = sheet.getRange(startRow + 2, 15, 10, 4);
    woChartArea.merge();
    sheet.getRange(startRow + 2, 15).setValue("Work Order Status");
    sheet.getRange(startRow + 2, 15).setFontWeight("bold");
    sheet.getRange(startRow + 2, 15).setFontSize(14);
  }

  /**
   * Builds the tables section with native Google Sheets tables.
   * @param {Sheet} sheet
   */
  function buildTablesSection(sheet) {
    const startRow = 24;

    // Recent Work Orders table
    sheet.getRange(startRow, 1).setValue("Recent Work Orders");
    sheet.getRange(startRow, 1).setFontWeight("bold");
    sheet.getRange(startRow, 1).setFontSize(14);
    sheet.getRange(startRow, 1).setFontColor(COLORS.TEXT_PRIMARY);

    // Table headers
    const woHeaders = [["WO ID", "Customer", "Vehicle", "Status", "Total", "Date"]];
    sheet.getRange(startRow + 1, 1, 1, 6).setValues(woHeaders);
    sheet.getRange(startRow + 1, 1, 1, 6).setFontWeight("bold");
    sheet.getRange(startRow + 1, 1, 1, 6).setFontSize(12);
    sheet.getRange(startRow + 1, 1, 1, 6).setFontColor(COLORS.TEXT_SECONDARY);
    sheet.getRange(startRow + 1, 1, 1, 6).setBackground(COLORS.LIGHT_GRAY);

    // Data range with borders
    const woDataRange = sheet.getRange(startRow + 2, 1, 10, 6);
    woDataRange.setBorder(true, true, true, true, true, false, COLORS.BORDER_GRAY, SpreadsheetApp.BorderStyle.SOLID);

    // Recent Payments table
    sheet.getRange(startRow, 8).setValue("Recent Payments");
    sheet.getRange(startRow, 8).setFontWeight("bold");
    sheet.getRange(startRow, 8).setFontSize(14);
    sheet.getRange(startRow, 8).setFontColor(COLORS.TEXT_PRIMARY);

    // Table headers
    const payHeaders = [["Payment ID", "Customer", "Amount", "Method", "Date", "Status"]];
    sheet.getRange(startRow + 1, 8, 1, 6).setValues(payHeaders);
    sheet.getRange(startRow + 1, 8, 1, 6).setFontWeight("bold");
    sheet.getRange(startRow + 1, 8, 1, 6).setFontSize(12);
    sheet.getRange(startRow + 1, 8, 1, 6).setFontColor(COLORS.TEXT_SECONDARY);
    sheet.getRange(startRow + 1, 8, 1, 6).setBackground(COLORS.LIGHT_GRAY);

    // Data range with borders
    const payDataRange = sheet.getRange(startRow + 2, 8, 10, 6);
    payDataRange.setBorder(true, true, true, true, true, false, COLORS.BORDER_GRAY, SpreadsheetApp.BorderStyle.SOLID);
  }

  /**
   * Builds the vehicle gallery section.
   * @param {Sheet} sheet
   */
  function buildVehicleGallery(sheet) {
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
    }
  }

  /**
   * Builds the footer section.
   * @param {Sheet} sheet
   */
  function buildFooter(sheet) {
    const footerRow = 48;

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
    // Set column widths
    sheet.setColumnWidth(1, 120);
    sheet.setColumnWidth(2, 120);
    sheet.setColumnWidth(3, 120);
    sheet.setColumnWidth(4, 120);
    sheet.setColumnWidth(5, 120);
    sheet.setColumnWidth(6, 120);
    sheet.setColumnWidth(7, 50); // gap
    sheet.setColumnWidth(8, 120);
    sheet.setColumnWidth(9, 120);
    sheet.setColumnWidth(10, 120);
    sheet.setColumnWidth(11, 120);
    sheet.setColumnWidth(12, 120);
    sheet.setColumnWidth(13, 50); // gap
    sheet.setColumnWidth(14, 120);
    sheet.setColumnWidth(15, 120);
    sheet.setColumnWidth(16, 120);
    sheet.setColumnWidth(17, 120);
    sheet.setColumnWidth(18, 120);

    // Set row heights
    sheet.setRowHeight(1, 40);
    sheet.setRowHeight(2, 25);

    // Hide gridlines for cleaner look
    sheet.setHiddenGridlines(true);
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
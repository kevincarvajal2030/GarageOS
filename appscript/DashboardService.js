
/**
 * ============================================================================
 * GarageOS - Dashboard Service | DashboardService.gs
 * ============================================================================
 * Central data collection module for the Dashboard.
 *
 * Uses CacheService with 20-second lifetime.
 * ============================================================================
 */

const DashboardService = (() => {

  const CACHE_KEY = "dashboard_data";
  const CACHE_TTL = 20; // seconds

  /**
   * Returns cached dashboard data if available and not expired.
   * @returns {Object|null}
   */
  function getCachedData() {
    const cache = CacheService.getUserCache();
    const cached = cache.get(CACHE_KEY);

    if (cached) {
      try {
        return JSON.parse(cached);
      } catch (e) {
        cache.remove(CACHE_KEY);
        return null;
      }
    }

    return null;
  }

  /**
   * Stores dashboard data in cache.
   * @param {Object} data
   */
  function storeInCache(data) {
    const cache = CacheService.getUserCache();
    cache.put(CACHE_KEY, JSON.stringify(data), CACHE_TTL);
  }

  /**
   * Clears the dashboard cache.
   */
  function clearCache() {
    const cache = CacheService.getUserCache();
    cache.remove(CACHE_KEY);
  }

  /**
   * Main entry point - returns complete dashboard data in a single call.
   * @returns {Object}
   */
  function getDashboardData() {
    // Check cache first
    const cached = getCachedData();
    if (cached) {
      return cached;
    }

    // Calculate fresh data
    const data = {
      customers: getCustomerStats(),
      vehicles: getVehicleStats(),
      workOrders: getWorkOrderStats(),
      payments: getPaymentStats(),
      revenue: getRevenueStats(),
      charts: {
        revenue: getRevenueChart(),
        vehicleMake: getVehicleMakeChart(),
        workOrderStatus: getWorkOrderStatusChart()
      },
      recentOrders: getRecentWorkOrders(),
      recentPayments: getRecentPayments(),
      topCustomers: getTopCustomers(),
      vehicleGallery: getVehicleGallery()
    };

    // Store in cache
    storeInCache(data);

    return data;
  }

  /**
   * Refreshes the dashboard by clearing cache and returning fresh data.
   * @returns {Object}
   */
  function refreshDashboard() {
    clearCache();
    return getDashboardData();
  }

  /**
   * Returns customer statistics.
   * @returns {Object}
   */
  function getCustomerStats() {
    const sheet = getSheet(SHEETS.CUSTOMERS);
    const lastRow = sheet.getLastRow();

    if (lastRow < TABLE.FIRST_DATA_ROW) {
      return { total: 0, active: 0, blocked: 0, inactive: 0 };
    }

    const config = ModuleConfig.get(SHEETS.CUSTOMERS);
    const values = sheet.getRange(
      TABLE.FIRST_DATA_ROW,
      1,
      lastRow - TABLE.FIRST_DATA_ROW + 1,
      config.fields.Status
    ).getValues();

    let total = 0;
    let active = 0;
    let blocked = 0;
    let inactive = 0;

    for (const row of values) {
      total++;
      const status = String(row[config.fields.Status - 1] || "").trim().toLowerCase();

      if (status === "active") active++;
      else if (status === "blocked") blocked++;
      else if (status === "inactive") inactive++;
    }

    return { total, active, blocked, inactive };
  }

  /**
   * Returns vehicle statistics.
   * @returns {Object}
   */
  function getVehicleStats() {
    const sheet = getSheet(SHEETS.VEHICLES);
    const lastRow = sheet.getLastRow();

    if (lastRow < TABLE.FIRST_DATA_ROW) {
      return { total: 0, active: 0, inRepair: 0, inactive: 0 };
    }

    const config = ModuleConfig.get(SHEETS.VEHICLES);
    const values = sheet.getRange(
      TABLE.FIRST_DATA_ROW,
      1,
      lastRow - TABLE.FIRST_DATA_ROW + 1,
      config.fields.Status
    ).getValues();

    let total = 0;
    let active = 0;
    let inRepair = 0;
    let inactive = 0;

    for (const row of values) {
      total++;
      const status = String(row[config.fields.Status - 1] || "").trim().toLowerCase();

      if (status === "active") active++;
      else if (status === "in repair") inRepair++;
      else if (status === "inactive") inactive++;
    }

    return { total, active, inRepair, inactive };
  }

  /**
   * Returns work order statistics.
   * @returns {Object}
   */
  function getWorkOrderStats() {
    const sheet = getSheet(SHEETS.WORK_ORDERS);
    const lastRow = sheet.getLastRow();

    if (lastRow < TABLE.FIRST_DATA_ROW) {
      return {
        total: 0,
        open: 0,
        completed: 0,
        inProgress: 0,
        waitingParts: 0,
        waitingApproval: 0,
        pending: 0,
        onHold: 0,
        cancelled: 0
      };
    }

    const config = ModuleConfig.get(SHEETS.WORK_ORDERS);
    const values = sheet.getRange(
      TABLE.FIRST_DATA_ROW,
      1,
      lastRow - TABLE.FIRST_DATA_ROW + 1,
      config.fields.Status
    ).getValues();

    let total = 0;
    let open = 0;
    let completed = 0;
    let inProgress = 0;
    let waitingParts = 0;
    let waitingApproval = 0;
    let pending = 0;
    let onHold = 0;
    let cancelled = 0;

    for (const row of values) {
      total++;
      const status = String(row[config.fields.Status - 1] || "").trim();
      const statusLower = status.toLowerCase();

      if (statusLower === "completed") completed++;
      else if (statusLower === "in progress") inProgress++;
      else if (statusLower === "waiting parts") waitingParts++;
      else if (statusLower === "waiting approval") waitingApproval++;
      else if (statusLower === "pending") pending++;
      else if (statusLower === "on hold") onHold++;
      else if (statusLower === "cancelled") cancelled++;
      else open++; // Count any other status as open
    }

    return {
      total,
      open,
      completed,
      inProgress,
      waitingParts,
      waitingApproval,
      pending,
      onHold,
      cancelled
    };
  }

  /**
   * Returns payment statistics.
   * @returns {Object}
   */
  function getPaymentStats() {
    const sheet = getSheet(SHEETS.PAYMENTS);
    const lastRow = sheet.getLastRow();

    if (lastRow < TABLE.FIRST_DATA_ROW) {
      return { total: 0, paid: 0, pending: 0, refunded: 0, totalRevenue: 0, outstandingBalance: 0 };
    }

    const config = ModuleConfig.get(SHEETS.PAYMENTS);
    const values = sheet.getRange(
      TABLE.FIRST_DATA_ROW,
      1,
      lastRow - TABLE.FIRST_DATA_ROW + 1,
      Math.max(config.fields.PaymentStatus, config.fields.TotalCost)
    ).getValues();

    let total = 0;
    let paid = 0;
    let pending = 0;
    let refunded = 0;
    let totalRevenue = 0;
    let outstandingBalance = 0;

    for (const row of values) {
      total++;
      const status = String(row[config.fields.PaymentStatus - 1] || "").trim();
      const amount = parseFloat(row[config.fields.TotalCost - 1]) || 0;

      totalRevenue += amount;

      if (status === "Paid") {
        paid++;
      } else if (status === "Pending") {
        pending++;
        outstandingBalance += amount;
      } else if (status === "Refunded") {
        refunded++;
        totalRevenue -= amount;
      }
    }

    return { total, paid, pending, refunded, totalRevenue, outstandingBalance };
  }

  /**
   * Returns revenue statistics.
   * @returns {Object}
   */
  function getRevenueStats() {
    const payments = getPaymentStats();

    // Calculate monthly revenue trend
    const monthlyTrend = calculateMonthlyRevenueTrend();

    return {
      total: payments.totalRevenue,
      outstanding: payments.outstandingBalance,
      monthlyTrend: monthlyTrend
    };
  }

  /**
   * Calculates monthly revenue trend for the last 6 months.
   * @returns {Array}
   */
  function calculateMonthlyRevenueTrend() {
    const sheet = getSheet(SHEETS.PAYMENTS);
    const lastRow = sheet.getLastRow();

    if (lastRow < TABLE.FIRST_DATA_ROW) {
      return [];
    }

    const config = ModuleConfig.get(SHEETS.PAYMENTS);
    const now = new Date();
    const months = [];

    // Generate last 6 months
    for (let i = 5; i >= 0; i--) {
      const date = new Date(now.getFullYear(), now.getMonth() - i, 1);
      months.push({
        label: Utilities.formatDate(date, Session.getScriptTimeZone(), "MMM yyyy"),
        value: 0,
        month: date.getMonth(),
        year: date.getFullYear()
      });
    }

    const values = sheet.getRange(
      TABLE.FIRST_DATA_ROW,
      1,
      lastRow - TABLE.FIRST_DATA_ROW + 1,
      Math.max(config.fields.PaymentStatus, config.fields.TotalCost, config.fields.PaymentDate)
    ).getValues();

    for (const row of values) {
      const status = String(row[config.fields.PaymentStatus - 1] || "").trim();
      const amount = parseFloat(row[config.fields.TotalCost - 1]) || 0;
      const dateValue = row[config.fields.PaymentDate - 1];

      if (status !== "Paid" || !dateValue) continue;

      const date = new Date(dateValue);

      for (const month of months) {
        if (date.getMonth() === month.month && date.getFullYear() === month.year) {
          month.value += amount;
          break;
        }
      }
    }

    return months.map(m => ({ label: m.label, value: Math.round(m.value * 100) / 100 }));
  }

  /**
   * Returns recent work orders (last 10).
   * @returns {Array}
   */
  function getRecentWorkOrders() {
    const sheet = getSheet(SHEETS.WORK_ORDERS);
    const lastRow = sheet.getLastRow();

    if (lastRow < TABLE.FIRST_DATA_ROW) {
      return [];
    }

    const config = ModuleConfig.get(SHEETS.WORK_ORDERS);
    const maxColumn = Math.max(
      config.fields.WorkOrderID,
      config.fields.CustomerName,
      config.fields.VehicleName,
      config.fields.Status,
      config.fields.TotalCost,
      config.fields.OpenDate
    );

    const values = sheet.getRange(
      TABLE.FIRST_DATA_ROW,
      1,
      lastRow - TABLE.FIRST_DATA_ROW + 1,
      maxColumn
    ).getValues();

    // Get last 10 orders (most recent first)
    const recent = [];
    const startIndex = Math.max(0, values.length - 10);

    for (let i = values.length - 1; i >= startIndex && recent.length < 10; i--) {
      const row = values[i];
      recent.push({
        woId: String(row[config.fields.WorkOrderID - 1] || "").trim(),
        customer: String(row[config.fields.CustomerName - 1] || "").trim(),
        vehicle: String(row[config.fields.VehicleName - 1] || "").trim(),
        status: String(row[config.fields.Status - 1] || "").trim(),
        total: parseFloat(row[config.fields.TotalCost - 1]) || 0,
        openDate: formatDate(row[config.fields.OpenDate - 1])
      });
    }

    return recent;
  }

  /**
   * Returns recent payments (last 10).
   * @returns {Array}
   */
  function getRecentPayments() {
    const sheet = getSheet(SHEETS.PAYMENTS);
    const lastRow = sheet.getLastRow();

    if (lastRow < TABLE.FIRST_DATA_ROW) {
      return [];
    }

    const config = ModuleConfig.get(SHEETS.PAYMENTS);
    const maxColumn = Math.max(
      config.fields.PaymentID,
      config.fields.CustomerName,
      config.fields.TotalCost,
      config.fields.PaymentMethod,
      config.fields.PaymentDate,
      config.fields.PaymentStatus
    );

    const values = sheet.getRange(
      TABLE.FIRST_DATA_ROW,
      1,
      lastRow - TABLE.FIRST_DATA_ROW + 1,
      maxColumn
    ).getValues();

    // Get last 10 payments (most recent first)
    const recent = [];
    const startIndex = Math.max(0, values.length - 10);

    for (let i = values.length - 1; i >= startIndex && recent.length < 10; i--) {
      const row = values[i];
      recent.push({
        paymentId: String(row[config.fields.PaymentID - 1] || "").trim(),
        customer: String(row[config.fields.CustomerName - 1] || "").trim(),
        amount: parseFloat(row[config.fields.TotalCost - 1]) || 0,
        method: String(row[config.fields.PaymentMethod - 1] || "").trim() || "Not specified",
        date: formatDate(row[config.fields.PaymentDate - 1]),
        status: String(row[config.fields.PaymentStatus - 1] || "").trim()
      });
    }

    return recent;
  }

  /**
   * Returns top customers by total spent.
   * @returns {Array}
   */
  function getTopCustomers() {
    const paymentsSheet = getSheet(SHEETS.PAYMENTS);
    const paymentsLastRow = paymentsSheet.getLastRow();

    if (paymentsLastRow < TABLE.FIRST_DATA_ROW) {
      return [];
    }

    const payConfig = ModuleConfig.get(SHEETS.PAYMENTS);
    const values = paymentsSheet.getRange(
      TABLE.FIRST_DATA_ROW,
      1,
      paymentsLastRow - TABLE.FIRST_DATA_ROW + 1,
      Math.max(payConfig.fields.CustomerName, payConfig.fields.TotalCost, payConfig.fields.PaymentStatus)
    ).getValues();

    // Aggregate by customer
    const customerMap = {};

    for (const row of values) {
      const customer = String(row[payConfig.fields.CustomerName - 1] || "").trim();
      const status = String(row[payConfig.fields.PaymentStatus - 1] || "").trim();
      const amount = parseFloat(row[payConfig.fields.TotalCost - 1]) || 0;

      if (!customer) continue;

      if (!customerMap[customer]) {
        customerMap[customer] = { totalSpent: 0, orders: 0 };
      }

      if (status === "Paid") {
        customerMap[customer].totalSpent += amount;
      }
      customerMap[customer].orders++;
    }

    // Convert to array and sort
    const result = Object.entries(customerMap)
      .map(([name, data]) => ({
        customer: name,
        totalSpent: Math.round(data.totalSpent * 100) / 100,
        orders: data.orders
      }))
      .sort((a, b) => b.totalSpent - a.totalSpent)
      .slice(0, 10);

    return result;
  }

  /**
   * Returns vehicle gallery data for the dashboard.
   * Reuses existing VehicleViewerService logic.
   * @returns {Array}
   */
  function getVehicleGallery() {
    const sheet = getSheet(SHEETS.VEHICLES);
    const lastRow = sheet.getLastRow();

    if (lastRow < TABLE.FIRST_DATA_ROW) {
      return [];
    }

    const config = ModuleConfig.get(SHEETS.VEHICLES);
    const maxColumn = Math.max(
      config.fields.VehicleID,
      config.fields.Make,
      config.fields.Model,
      config.fields.Year,
      config.fields.Status,
      config.fields.ImageFileID
    );

    const values = sheet.getRange(
      TABLE.FIRST_DATA_ROW,
      1,
      lastRow - TABLE.FIRST_DATA_ROW + 1,
      maxColumn
    ).getValues();

    const gallery = [];

    for (const row of values) {
      const status = String(row[config.fields.Status - 1] || "").trim().toLowerCase();

      // Only show active vehicles
      if (status !== "active") continue;

      const make = String(row[config.fields.Make - 1] || "").trim();
      const model = String(row[config.fields.Model - 1] || "").trim();
      const year = String(row[config.fields.Year - 1] || "").trim();
      const imageFileId = String(row[config.fields.ImageFileID - 1] || "").trim();

      gallery.push({
        vehicleId: String(row[config.fields.VehicleID - 1] || "").trim(),
        displayName: `${make} ${model} ${year}`.trim(),
        make: make,
        model: model,
        year: year,
        imageFileId: imageFileId,
        imageUrl: imageFileId ? `https://drive.google.com/thumbnail?id=${imageFileId}&sz=w400` : ""
      });

      // Limit to 12 vehicles for gallery
      if (gallery.length >= 12) break;
    }

    return gallery;
  }

  /**
   * Returns data for the revenue line chart.
   * @returns {Object}
   */
  function getRevenueChart() {
    const monthlyData = calculateMonthlyRevenueTrend();

    return {
      labels: monthlyData.map(m => m.label),
      values: monthlyData.map(m => m.value)
    };
  }

  /**
   * Returns data for the vehicle make doughnut chart.
   * @returns {Object}
   */
  function getVehicleMakeChart() {
    const sheet = getSheet(SHEETS.VEHICLES);
    const lastRow = sheet.getLastRow();

    if (lastRow < TABLE.FIRST_DATA_ROW) {
      return { labels: [], values: [] };
    }

    const config = ModuleConfig.get(SHEETS.VEHICLES);
    const values = sheet.getRange(
      TABLE.FIRST_DATA_ROW,
      1,
      lastRow - TABLE.FIRST_DATA_ROW + 1,
      config.fields.Make
    ).getValues();

    const makeCount = {};
    const knownMakes = ["Toyota", "Honda", "Ford", "Chevrolet", "Nissan"];

    for (const row of values) {
      const make = String(row[config.fields.Make - 1] || "").trim();
      if (!make) continue;

      const category = knownMakes.includes(make) ? make : "Others";

      if (!makeCount[category]) {
        makeCount[category] = 0;
      }
      makeCount[category]++;
    }

    // Ensure consistent ordering
    const orderedLabels = [...knownMakes, "Others"];
    const labels = [];
    const values_arr = [];

    for (const label of orderedLabels) {
      if (makeCount[label] > 0) {
        labels.push(label);
        values_arr.push(makeCount[label]);
      }
    }

    return { labels, values: values_arr };
  }

  /**
   * Returns data for the work order status doughnut chart.
   * @returns {Object}
   */
  function getWorkOrderStatusChart() {
    const stats = getWorkOrderStats();

    // Only include statuses that have values
    const allStatuses = [
      { label: "In Progress", value: stats.inProgress, color: "#3B82F6" },
      { label: "Waiting Parts", value: stats.waitingParts, color: "#F59E0B" },
      { label: "Waiting Approval", value: stats.waitingApproval, color: "#8B5CF6" },
      { label: "Pending", value: stats.pending, color: "#F97316" },
      { label: "On Hold", value: stats.onHold, color: "#22C55E" },
      { label: "Completed", value: stats.completed, color: "#10B981" },
      { label: "Cancelled", value: stats.cancelled, color: "#EF4444" }
    ];

    const filtered = allStatuses.filter(s => s.value > 0);

    return {
      labels: filtered.map(s => s.label),
      values: filtered.map(s => s.value),
      colors: filtered.map(s => s.color)
    };
  }

  /**
   * Formats a date value for display.
   * @param {*} dateValue
   * @returns {string}
   */
  function formatDate(dateValue) {
    if (!dateValue) return "";

    try {
      const date = new Date(dateValue);
      return Utilities.formatDate(date, Session.getScriptTimeZone(), "MM/dd/yyyy");
    } catch (e) {
      return String(dateValue);
    }
  }

  return {
    getDashboardData,
    getCustomerStats,
    getVehicleStats,
    getWorkOrderStats,
    getRevenueStats,
    getRecentWorkOrders,
    getRecentPayments,
    getTopCustomers,
    getVehicleGallery,
    getRevenueChart,
    getVehicleMakeChart,
    getWorkOrderStatusChart,
    refreshDashboard
  };

})();

/**
 * Global wrapper functions for google.script.run
 */
function getDashboardData() {
  return DashboardService.getDashboardData();
}

function getCustomerStats() {
  return DashboardService.getCustomerStats();
}

function getVehicleStats() {
  return DashboardService.getVehicleStats();
}

function getWorkOrderStats() {
  return DashboardService.getWorkOrderStats();
}

function getRevenueStats() {
  return DashboardService.getRevenueStats();
}

function getRecentWorkOrders() {
  return DashboardService.getRecentWorkOrders();
}

function getRecentPayments() {
  return DashboardService.getRecentPayments();
}

function getTopCustomers() {
  return DashboardService.getTopCustomers();
}

function getVehicleGallery() {
  return DashboardService.getVehicleGallery();
}

function getRevenueChart() {
  return DashboardService.getRevenueChart();
}

function getVehicleMakeChart() {
  return DashboardService.getVehicleMakeChart();
}

function getWorkOrderStatusChart() {
  return DashboardService.getWorkOrderStatusChart();
}

function refreshDashboard() {
  return DashboardService.refreshDashboard();
}


/**
 * Refreshes dashboard data by updating KPI formulas and chart data ranges.
 * This does NOT rebuild the layout - only refreshes data.
 */
function refreshDashboardData() {
  const sheet = getSheet(SHEETS.DASHBOARD);
  if (!sheet) {
    throw new Error("Dashboard sheet not found");
  }

  // Force recalculation of all formulas by touching a cell
  // This triggers Google Sheets to recalculate all dependent formulas
  const lastUpdated = sheet.getRange("I1");
  lastUpdated.setFormula('=TEXT(NOW(),"MM/dd/yyyy HH:mm")');

  // Trigger chart data refresh by updating hidden data areas
  DashboardBuilder.buildChartsSection(sheet);

  showToast("Dashboard data refreshed", "GarageOS");
}



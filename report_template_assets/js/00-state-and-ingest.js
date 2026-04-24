    const dashboardData = JSON.parse(document.getElementById("dashboard-data").textContent);

    const ui = {
      startDate: document.getElementById("start-date"),
      endDate: document.getElementById("end-date"),
      week: document.getElementById("week-filter"),
      month: document.getElementById("month-filter"),
      dashboardMode: document.getElementById("dashboard-mode"),
      rig: document.getElementById("rig-filter"),
      field: document.getElementById("field-filter"),
      well: document.getElementById("well-filter"),
      category: document.getElementById("category-filter"),
      type: document.getElementById("type-filter"),
      app: document.getElementById("app-filter"),
      validation: document.getElementById("validation-filter"),
      granularity: document.getElementById("granularity-filter"),
      search: document.getElementById("search-filter"),
      reset: document.getElementById("reset-filters"),
      themeToggle: document.getElementById("theme-toggle"),
      themeToggleState: document.getElementById("theme-toggle-state"),
      themeToggleThumb: document.getElementById("theme-toggle-thumb"),
      resultsTitle: document.getElementById("results-title"),
      resultsSubtitle: document.getElementById("results-subtitle"),
      activeFilters: document.getElementById("active-filters"),
      kpiGrid: document.getElementById("kpi-grid"),
      trendChart: document.getElementById("trend-chart"),
      categoryChart: document.getElementById("category-chart"),
      rigChart: document.getElementById("rig-chart"),
      typeChart: document.getElementById("type-chart"),
      appChart: document.getElementById("app-chart"),
      categoryTable: document.getElementById("category-table"),
      fieldTable: document.getElementById("field-table"),
      wellTable: document.getElementById("well-table"),
      interventionTable: document.getElementById("intervention-table"),
      caChart: document.getElementById("ca-chart"),
      presetButtons: Array.from(document.querySelectorAll(".preset-btn")),
      toggles: Array.from(document.querySelectorAll("[data-target]")),
      viewTabs: Array.from(document.querySelectorAll(".view-tab")),
      viewPanels: Array.from(document.querySelectorAll(".view-panel")),
      weeklyReportStartDate: document.getElementById("weekly-report-start-date"),
      weeklyReportEndDate: document.getElementById("weekly-report-end-date"),
      weeklyExportPdf: document.getElementById("weekly-export-pdf"),
      weeklyReportTitle: document.getElementById("weekly-report-title"),
      weeklyReportSubtitle: document.getElementById("weekly-report-subtitle"),
      weeklyReportRange: document.getElementById("weekly-report-range"),
      weeklyBannerCopy: document.getElementById("weekly-banner-copy"),
      weeklyBannerChip1: document.getElementById("weekly-banner-chip-1"),
      weeklyBannerChip2: document.getElementById("weekly-banner-chip-2"),
      weeklyBannerChip3: document.getElementById("weekly-banner-chip-3"),
      weeklyCategoryTable: document.getElementById("weekly-category-table"),
      weeklyCategoryChart: document.getElementById("weekly-category-chart"),
      cumulativeCategoryTable: document.getElementById("cumulative-category-table"),
      cumulativeCategoryChart: document.getElementById("cumulative-category-chart"),
      wiperSummaryTable: document.getElementById("wiper-summary-table"),
      wiperSummaryChart: document.getElementById("wiper-summary-chart"),
      ropSummaryTable: document.getElementById("rop-summary-table"),
      ropSummaryChart: document.getElementById("rop-summary-chart"),
      kpiSummaryTable: document.getElementById("kpi-summary-table"),
      kpiSummaryChart: document.getElementById("kpi-summary-chart"),
      weeklyHighlightMetrics: document.getElementById("weekly-highlight-metrics"),
      actualHighlightsTable: document.getElementById("actual-highlights-table"),
      potentialHighlightsTable: document.getElementById("potential-highlights-table"),
      weeklyStatsMetrics: document.getElementById("weekly-stats-metrics"),
      weeklyStatsTable: document.getElementById("weekly-stats-table"),
      weeklyMonitoredWellsTable: document.getElementById("weekly-monitored-wells-table"),
      flatTimeTitle: document.getElementById("flat-time-title"),
      flatTimeSubtitle: document.getElementById("flat-time-subtitle"),
      flatTimeRig: document.getElementById("flat-time-rig"),
      flatTimeSection: document.getElementById("flat-time-section"),
      flatTimeMetric: document.getElementById("flat-time-metric"),
      flatTimeTopN: document.getElementById("flat-time-top-n"),
      flatTimeMode: document.getElementById("flat-time-mode"),
      flatTimeHeatmapMode: document.getElementById("flat-time-heatmap-mode"),
      flatTimeWell: document.getElementById("flat-time-well"),
      flatTimeUpload: document.getElementById("flat-time-upload"),
      flatTimeRecalculate: document.getElementById("flat-time-recalculate"),
      flatTimeClearUploads: document.getElementById("flat-time-clear-uploads"),
      flatTimeDatasetTags: document.getElementById("flat-time-dataset-tags"),
      flatTimeSummary: document.getElementById("flat-time-summary"),
      flatTimeTableSummary: document.getElementById("flat-time-table-summary"),
      flatTimeAllWellsLayout: document.getElementById("flat-time-all-wells-layout"),
      flatTimeWellRanking: document.getElementById("flat-time-well-ranking"),
      flatTimeParetoChart: document.getElementById("flat-time-pareto-chart"),
      flatTimeAllWellsSections: document.getElementById("flat-time-all-wells-sections"),
      flatTimeSelectedWellSections: document.getElementById("flat-time-selected-well-sections"),
      flatTimeOffsetComparison: document.getElementById("flat-time-offset-comparison"),
      flatTimeSectionHeatmap: document.getElementById("flat-time-section-heatmap"),
      flatTimeSavingsSummary: document.getElementById("flat-time-savings-summary"),
      flatTimeNarrative: document.getElementById("flat-time-narrative"),
      flatTimeWaterfallChart: document.getElementById("flat-time-waterfall-chart"),
      flatTimeSectionBenchmarkChart: document.getElementById("flat-time-section-benchmark-chart"),
      flatTimeRigSummary: document.getElementById("flat-time-rig-summary"),
      flatTimeOpportunityPipeline: document.getElementById("flat-time-opportunity-pipeline"),
      flatTimeGroupLegend: document.getElementById("flat-time-group-legend"),
      flatTimeGroupChart: document.getElementById("flat-time-group-chart"),
      flatTimeActivityChart: document.getElementById("flat-time-activity-chart"),
      flatTimeBenchmarkTable: document.getElementById("flat-time-benchmark-table"),
      flatTimeDrilldownNote: document.getElementById("flat-time-drilldown-note"),
      flatTimeWellDrilldown: document.getElementById("flat-time-well-drilldown"),
      flatTimeActivityDrilldown: document.getElementById("flat-time-activity-drilldown"),
      flatTimeOpportunityTable: document.getElementById("flat-time-opportunity-table"),
      flatTimeGroupTable: document.getElementById("flat-time-group-table"),
      flatTimeLossDrivers: document.getElementById("flat-time-loss-drivers"),
      flatTimeVariabilityChart: document.getElementById("flat-time-variability-chart"),
      flatTimeHeatmapNote: document.getElementById("flat-time-heatmap-note"),
      flatTimeHeatmap: document.getElementById("flat-time-heatmap"),
      flatTimePerfectChart: document.getElementById("flat-time-perfect-chart"),
      flatTimeModeSections: Array.from(document.querySelectorAll("#flat-time-view [data-flat-mode]")),
    };

    const currencyFormatter = new Intl.NumberFormat("en-US", {
      style: "currency",
      currency: "USD",
      maximumFractionDigits: 0,
    });

    const numberFormatter = new Intl.NumberFormat("en-US", {
      maximumFractionDigits: 1,
      minimumFractionDigits: 0,
    });

    const CATEGORY_ORDER = ["Stuck pipe", "Optimization", "Operational Compliance", "Well Control", "Reporting"];
    const THEME_STORAGE_KEY = "weekly-report-theme";
    const WEEKLY_MONITORED_OVERRIDES_KEY = "weekly-monitored-wells-overrides";
    const FLAT_TIME_SERIES_COLORS = ["#1264d6", "#0f766e", "#c06a0a", "#be123c", "#7c3aed", "#0891b2", "#16a34a", "#dc2626"];
    const FLAT_TIME_ACTIVITY_TRANSLATIONS = dashboardData.activityCodeTranslations || {
      loaded: false,
      source: "",
      wellSections: {},
      operations: {},
      activities: {},
      generic: {},
    };
    const FLAT_TIME_RIG_LOOKUP = buildFlatTimeRigLookup(Array.isArray(dashboardData.interventions) ? dashboardData.interventions : []);
    const flatTimeState = {
      baseDatasets: [],
      uploadedDatasets: [],
      focusWell: "",
      focusActivity: "",
      focusActivities: [],
      heatmapMode: "gap",
    };

    function getChartTheme() {
      const isCorona = document.body.classList.contains("theme-corona");
      return isCorona
        ? {
            text: "#f5f5f5",
            muted: "#a1aab8",
            grid: "#2c2e33",
            axis: "#3a3d46",
            line: "#0090e7",
            area: "rgba(0, 144, 231, 0.16)",
            pointLabel: "#f5f5f5",
            valueLabel: "#d5d9e0",
          }
        : {
            text: "#1f2d3d",
            muted: "#607085",
            grid: "#d8e2ef",
            axis: "#9fb3c8",
            line: "#1264d6",
            area: "rgba(18, 100, 214, 0.10)",
            pointLabel: "#34475d",
            valueLabel: "#6b7b8d",
          };
    }

    function applyTheme(theme) {
      const resolvedTheme = theme === "corona" ? "corona" : "classic";
      document.body.classList.toggle("theme-corona", resolvedTheme === "corona");
      ui.themeToggleState.textContent = resolvedTheme === "corona" ? "Corona" : "Classic";
      ui.themeToggleThumb.textContent = resolvedTheme === "corona" ? "Corona" : "Classic";
      localStorage.setItem(THEME_STORAGE_KEY, resolvedTheme);
    }

    function toggleTheme() {
      applyTheme(document.body.classList.contains("theme-corona") ? "classic" : "corona");
      applyFilters();
      renderWeeklyReport();
      renderFlatTime();
    }

    function escapeHtml(value) {
      return String(value ?? "")
        .replaceAll("&", "&amp;")
        .replaceAll("<", "&lt;")
        .replaceAll(">", "&gt;")
        .replaceAll('"', "&quot;")
        .replaceAll("'", "&#39;");
    }

    function formatCurrency(value) {
      return currencyFormatter.format(Number(value || 0));
    }

    function formatNumber(value) {
      return numberFormatter.format(Number(value || 0));
    }

    function getWeeklyMonitoredOverrides() {
      try {
        const raw = localStorage.getItem(WEEKLY_MONITORED_OVERRIDES_KEY);
        const parsed = raw ? JSON.parse(raw) : {};
        return parsed && typeof parsed === "object" ? parsed : {};
      } catch (error) {
        return {};
      }
    }

    function setWeeklyMonitoredOverride(rowKey, fieldName, value) {
      const overrides = getWeeklyMonitoredOverrides();
      if (!overrides[rowKey] || typeof overrides[rowKey] !== "object") {
        overrides[rowKey] = {};
      }
      overrides[rowKey][fieldName] = String(value || "").trim();
      localStorage.setItem(WEEKLY_MONITORED_OVERRIDES_KEY, JSON.stringify(overrides));
    }

    function weeklyMonitoredRowKey(row) {
      return String(row.rig || "") + "||" + String(row.well || "");
    }

    function parseEditableNumber(value, fallback = 0) {
      const text = String(value ?? "").trim();
      if (!text) return Number(fallback || 0);
      const parsed = Number(text.replace(/,/g, ""));
      return Number.isFinite(parsed) ? parsed : Number(fallback || 0);
    }

    function parseEditablePercent(value, fallback = 0) {
      const text = String(value ?? "").trim().replace(/%/g, "");
      if (!text) return Number(fallback || 0);
      const parsed = Number(text);
      return Number.isFinite(parsed) ? parsed : Number(fallback || 0);
    }

    function slugify(value) {
      return String(value || "")
        .toLowerCase()
        .replace(/[^a-z0-9]+/g, "-")
        .replace(/^-+|-+$/g, "");
    }

    function normalizeFlatTimeWellToken(value) {
      return String(value || "").toUpperCase().replace(/[^A-Z0-9]/g, "");
    }

    function baseFlatTimeWellToken(value) {
      return normalizeFlatTimeWellToken(String(value || "").replace(/_\d+$/, ""));
    }

    function buildFlatTimeRigLookup(rows) {
      const lookup = new Map();

      function addRig(key, rigName) {
        if (!key || !rigName) return;
        if (!lookup.has(key)) lookup.set(key, new Set());
        lookup.get(key).add(rigName);
      }

      rows.forEach((row) => {
        const wellName = row.wellName || "";
        const rigName = row.rigName || "";
        addRig(normalizeFlatTimeWellToken(wellName), rigName);
        addRig(baseFlatTimeWellToken(wellName), rigName);
      });

      return lookup;
    }

    function resolveFlatTimeRigLabel(subjectWell) {
      const exact = FLAT_TIME_RIG_LOOKUP.get(normalizeFlatTimeWellToken(subjectWell));
      if (exact && exact.size) return Array.from(exact).sort().join(" / ");

      const base = FLAT_TIME_RIG_LOOKUP.get(baseFlatTimeWellToken(subjectWell));
      if (base && base.size) return Array.from(base).sort().join(" / ");

      return "Rig not mapped";
    }

    function deriveFlatTimeRigLabelFromFileName(fileName) {
      const baseName = String(fileName || "").replace(/\.[^.]+$/, "");
      const prefix = baseName.split("_")[0].trim();
      if (prefix) return prefix;
      return "Rig not mapped";
    }

    function enrichFlatTimeDataset(dataset) {
      if (!dataset) return dataset;
      return {
        ...dataset,
        rigLabel: dataset.rigLabel || resolveFlatTimeRigLabel(dataset.subjectWell),
      };
    }

    function parseCsvLine(line) {
      const values = [];
      let current = "";
      let inQuotes = false;
      for (let index = 0; index < line.length; index += 1) {
        const char = line[index];
        if (char === '"') {
          if (inQuotes && line[index + 1] === '"') {
            current += '"';
            index += 1;
          } else {
            inQuotes = !inQuotes;
          }
        } else if (char === "," && !inQuotes) {
          values.push(current.trim());
          current = "";
        } else {
          current += char;
        }
      }
      values.push(current.trim());
      return values;
    }

    function parseFlatTimeCsvText(fileName, text) {
      const rows = String(text || "")
        .replace(/\r\n/g, "\n")
        .replace(/\r/g, "\n")
        .split("\n")
        .map((line) => parseCsvLine(line));
      const rigLabel = deriveFlatTimeRigLabelFromFileName(fileName);
      const datasetMap = new Map();
      let currentGroupName = "";
      let currentHeader = null;

      function getDataset(subjectWell) {
        if (!datasetMap.has(subjectWell)) {
          datasetMap.set(subjectWell, {
            id: slugify(fileName + "-" + subjectWell),
            fileName,
            subjectWell,
            rigLabel,
            groupsMap: new Map(),
          });
        }
        return datasetMap.get(subjectWell);
      }

      function getGroup(dataset, groupName) {
        if (!dataset.groupsMap.has(groupName)) {
          dataset.groupsMap.set(groupName, {
            groupName,
            activities: [],
            totalSubjectHours: 0,
            totalMeanHours: 0,
            totalMedianHours: 0,
          });
        }
        return dataset.groupsMap.get(groupName);
      }

      rows.forEach((row) => {
        const first = (row[0] || "").trim();
        if (first === "Group Name") {
          currentGroupName = (row[1] || "Unknown").trim();
          currentHeader = null;
          return;
        }

        if (!currentGroupName) return;
        if (!first || first === "Group Type") return;

        if (first === "Activity") {
          const wellColumns = [];
          let meanIndex = -1;
          let medianIndex = -1;

          row.forEach((cell, index) => {
            const label = String(cell || "").trim();
            if (index === 0 || !label) return;
            if (/^mean/i.test(label)) {
              meanIndex = index;
              return;
            }
            if (/^median/i.test(label)) {
              medianIndex = index;
              return;
            }
            wellColumns.push({ index, label });
          });

          currentHeader = { wellColumns, meanIndex, medianIndex };
          return;
        }

        if (!currentHeader || !currentHeader.wellColumns.length) return;

        if (first === "Total") {
          currentHeader.wellColumns.forEach((column) => {
            const dataset = getDataset(column.label);
            const group = getGroup(dataset, currentGroupName);
            group.totalSubjectHours = Number(row[column.index] || group.totalSubjectHours || 0);
          });
          return;
        }

        currentHeader.wellColumns.forEach((column) => {
          const subjectHours = Number(row[column.index] || 0);
          if (!subjectHours) return;
          const dataset = getDataset(column.label);
          const group = getGroup(dataset, currentGroupName);
          group.activities.push({
            activity: first,
            sectionSize: extractFlatTimeSectionSize(first),
            subjectHours,
            meanHours: currentHeader.meanIndex >= 0 ? Number(row[currentHeader.meanIndex] || 0) : 0,
            medianHours: currentHeader.medianIndex >= 0 ? Number(row[currentHeader.medianIndex] || 0) : 0,
          });
        });
      });

      return Array.from(datasetMap.values()).map((dataset) => {
        const groups = Array.from(dataset.groupsMap.values()).filter((group) => group.activities.length || group.totalSubjectHours);
        groups.forEach((group) => {
          if (!group.totalSubjectHours) {
            group.totalSubjectHours = group.activities.reduce((sum, item) => sum + item.subjectHours, 0);
          }
          // Recompute aggregate benchmarks from activity rows because several CSV
          // exports carry inconsistent group total mean/median values.
          group.totalMeanHours = group.activities.reduce((sum, item) => sum + item.meanHours, 0);
          group.totalMedianHours = group.activities.reduce((sum, item) => sum + item.medianHours, 0);
        });

        return {
          id: dataset.id,
          fileName: dataset.fileName,
          subjectWell: dataset.subjectWell,
          rigLabel: dataset.rigLabel,
          groups,
          totalSubjectHours: groups.reduce((sum, group) => sum + group.totalSubjectHours, 0),
          totalMeanHours: groups.reduce((sum, group) => sum + group.totalMeanHours, 0),
          totalMedianHours: groups.reduce((sum, group) => sum + group.totalMedianHours, 0),
        };
      });
    }

    function getFlatTimeDatasets() {
      return [...flatTimeState.baseDatasets, ...flatTimeState.uploadedDatasets].map(enrichFlatTimeDataset);
    }

    function getAvailableFlatTimeRigs(datasets) {
      return Array.from(new Set(datasets.map((dataset) => dataset.rigLabel || "Rig not mapped"))).sort((left, right) => left.localeCompare(right));
    }

    function populateFlatTimeRigOptions(datasets) {
      const current = ui.flatTimeRig.value;
      const options = ['<option value="">All rigs</option>'];
      const rigs = getAvailableFlatTimeRigs(datasets);
      rigs.forEach((rig) => {
        options.push('<option value="' + escapeHtml(rig) + '">' + escapeHtml(rig) + "</option>");
      });
      ui.flatTimeRig.innerHTML = options.join("");
      if (rigs.includes(current)) {
        ui.flatTimeRig.value = current;
      }
    }

    function filterFlatTimeDatasetsByRig(datasets, rigLabel) {
      if (!rigLabel) return datasets;
      return datasets.filter((dataset) => (dataset.rigLabel || "Rig not mapped") === rigLabel);
    }

    function getFlatTimeMetricKey() {
      const metric = ui.flatTimeMetric.value || "subject";
      return metric === "mean" ? "meanHours" : metric === "median" ? "medianHours" : "subjectHours";
    }

    function getFlatTimeTotalKey() {
      const metric = ui.flatTimeMetric.value || "subject";
      return metric === "mean" ? "totalMeanHours" : metric === "median" ? "totalMedianHours" : "totalSubjectHours";
    }

    function getFlatTimeMode() {
      return ui.flatTimeMode.value || "executive";
    }

    function updateFlatTimeModeVisibility() {
      const mode = getFlatTimeMode();
      ui.flatTimeModeSections.forEach((section) => {
      const allowedModes = String(section.dataset.flatMode || "executive engineering").split(/\s+/).filter(Boolean);
        section.hidden = !allowedModes.includes(mode);
      });
    }

    function populateFlatTimeWellOptions(datasets, preferredWell) {
      const current = ui.flatTimeWell.value;
      const options = ['<option value="">Auto-select worst well</option>'];
      const wellNames = datasets.map((dataset) => dataset.subjectWell);
      wellNames.forEach((wellName) => {
        const dataset = datasets.find((item) => item.subjectWell === wellName);
        const label = wellName + (dataset && dataset.rigLabel ? " • " + dataset.rigLabel : "");
        options.push('<option value="' + escapeHtml(wellName) + '">' + escapeHtml(label) + "</option>");
      });
      ui.flatTimeWell.innerHTML = options.join("");
      if (wellNames.includes(current)) {
        ui.flatTimeWell.value = current;
      } else if (preferredWell && wellNames.includes(preferredWell)) {
        ui.flatTimeWell.value = preferredWell;
      } else {
        ui.flatTimeWell.value = "";
      }
    }

    function createFlatTimeUploadId(fileName, subjectWell) {
      return slugify(fileName + "-" + subjectWell + "-" + Date.now() + "-" + Math.random().toString(36).slice(2, 8));
    }

    function extractFlatTimeSectionSize(activityName) {
      const match = String(activityName || "").match(/^(\d+(?:\.\d+)?)(?=-)/);
      return match ? match[1] : "__no_section__";
    }

    function formatFlatTimeSectionSize(sectionSize) {
      return sectionSize === "__no_section__" ? "No section size" : sectionSize + '"';
    }

    function deriveFlatTimeAreaLabel(subjectWell) {
      const token = String(subjectWell || "").trim().split(/[-_]/)[0] || "";
      const match = token.match(/[A-Za-z]+/);
      return match ? match[0].toUpperCase() : "N/A";
    }

    const FLAT_TIME_GROUP_LABELS = {
      DRLG: "Drilling Flat Time",
      WBCO: "WBCO",
      CT: "Hole Condition Trip",
      CSG: "Run & CMT Csg / Liner",
      BOP: "Work on BOP",
      WH: "Well Head Work",
      LOG: "Logging",
      COMP: "Run Completion",
      MAIN: "Main Operations",
      RM: "Reaming / Conditioning",
      OT: "Other",
      SUSP: "Suspension",
      TSMF: "Tubular / Completion Tools",
      TINL: "Tubing / Liner",
      FISH: "Fishing",
      KILL: "Kill Operations",
      MILL: "Milling",
      PA: "Pressure Activities",
      DCOM: "Downhole Completion",
    };

    const FLAT_TIME_GROUP_ORDER = ["DRLG", "WBCO", "CT", "CSG", "BOP", "WH", "LOG", "COMP", "MAIN", "RM", "OT", "SUSP", "TSMF", "TINL", "FISH", "KILL", "MILL", "PA", "DCOM"];

    function flatTimeGroupDisplayLabel(groupName) {
      return FLAT_TIME_GROUP_LABELS[groupName] || groupName || "Unknown";
    }

    function compareFlatTimeGroups(left, right) {
      const leftIndex = FLAT_TIME_GROUP_ORDER.indexOf(left);
      const rightIndex = FLAT_TIME_GROUP_ORDER.indexOf(right);
      if (leftIndex >= 0 && rightIndex >= 0) return leftIndex - rightIndex;
      if (leftIndex >= 0) return -1;
      if (rightIndex >= 0) return 1;
      return String(left).localeCompare(String(right));
    }

    function summarizeFlatTimeActivityCodes(activityLabels) {
      const codes = Array.from(
        new Set(
          (activityLabels || [])
            .map((label) => String(label || "").split("-").filter(Boolean).pop() || "")
            .filter(Boolean)
        )
      ).sort((left, right) => left.localeCompare(right));

      if (!codes.length) return "";

      return codes
        .map((code) => {
          const description =
            (FLAT_TIME_ACTIVITY_TRANSLATIONS.activities || {})[code] ||
            (FLAT_TIME_ACTIVITY_TRANSLATIONS.generic || {})[code] ||
            "";
          return description ? code + " (" + description + ")" : code;
        })
        .join(", ");
    }

    function buildSelectedWellTableSummaryRows(selectedDataset, opportunities, selectedSectionSize) {
      if (!selectedDataset || !opportunities.length) return [];

      const sectionMap = new Map();
      opportunities.forEach((opportunity) => {
        const sectionSize = opportunity.sectionSize || "__no_section__";
        if (selectedSectionSize && sectionSize !== selectedSectionSize) return;
        if (!sectionMap.has(sectionSize)) sectionMap.set(sectionSize, new Map());
        const groupMap = sectionMap.get(sectionSize);
        if (!groupMap.has(opportunity.groupLabel)) {
          groupMap.set(opportunity.groupLabel, {
            groupLabel: opportunity.groupLabel,
            label: flatTimeGroupDisplayLabel(opportunity.groupLabel),
            meanHours: 0,
            actualHours: 0,
            activityLabels: [],
          });
        }
        const entry = groupMap.get(opportunity.groupLabel);
        entry.meanHours += Number(opportunity.meanValue || 0);
        entry.actualHours += Number(opportunity.ranked.find((item) => item.datasetId === selectedDataset.id)?.value || 0);
        entry.activityLabels.push(opportunity.activityLabel);
      });

      return Array.from(sectionMap.entries())
        .map(([sectionSize, groupMap]) => ({
          sectionSize,
          sectionLabel: formatFlatTimeSectionSize(sectionSize),
          groups: Array.from(groupMap.values())
            .filter((item) => item.meanHours > 0 || item.actualHours > 0)
            .sort((left, right) => compareFlatTimeGroups(left.groupLabel, right.groupLabel)),
        }))
        .filter((item) => item.groups.length)
        .sort((left, right) => compareFlatTimeSectionSizes(left.sectionSize, right.sectionSize));
    }

    function renderFlatTimeTableSummary(target, selectedDataset, opportunities, selectedSectionSize) {
      if (!selectedDataset || !opportunities.length) {
        target.innerHTML = '<div class="empty">Choose a well and load comparable offsets to build the table summary.</div>';
        return;
      }

      const sectionRows = buildSelectedWellTableSummaryRows(selectedDataset, opportunities, selectedSectionSize);
      if (!sectionRows.length) {
        target.innerHTML = '<div class="empty">No section summary is available for the selected well in the current scope.</div>';
        return;
      }

      const areaLabel = deriveFlatTimeAreaLabel(selectedDataset.subjectWell);
      const blocks = sectionRows.map((sectionRow, sectionIndex) => {
        const rowSpan = sectionRow.groups.length;
        const rowsHtml = sectionRow.groups
          .map((groupRow, rowIndex) => {
            const metaColumns = rowIndex === 0
              ? (
                  '<td class="table-summary-meta" rowspan="' + rowSpan + '">' + escapeHtml(String(sectionIndex + 1)) + '</td>' +
                  '<td class="table-summary-meta" rowspan="' + rowSpan + '">' + escapeHtml(areaLabel) + '</td>' +
                  '<td class="table-summary-meta" rowspan="' + rowSpan + '">' + escapeHtml(selectedDataset.rigLabel || "Rig not mapped") + '</td>' +
                  '<td class="table-summary-meta" rowspan="' + rowSpan + '">' + escapeHtml(selectedDataset.subjectWell) + '</td>' +
                  '<td class="table-summary-meta" rowspan="' + rowSpan + '">' + escapeHtml(sectionRow.sectionLabel) + '</td>'
                )
              : "";
            const codesSummary = summarizeFlatTimeActivityCodes(groupRow.activityLabels);
            const meanHtml = groupRow.meanHours > 0 ? escapeHtml(formatNumber(groupRow.meanHours)) : '<span class="table-summary-missing">-</span>';
            const actualHtml = groupRow.actualHours > 0 ? escapeHtml(formatNumber(groupRow.actualHours)) : '<span class="table-summary-missing">-</span>';
            return (
              '<tr>' +
              metaColumns +
              '<td class="table-summary-major" title="' + escapeHtml(codesSummary) + '">' +
              '<strong>' + escapeHtml(groupRow.label) + '</strong>' +
              '<span>(' + escapeHtml(codesSummary || "No coded activities") + ')</span>' +
              '</td>' +
              '<td style="text-align:center; font-weight:700;">' + meanHtml + '</td>' +
              '<td style="text-align:center; font-weight:700;">' + actualHtml + '</td>' +
              '</tr>'
            );
          })
          .join("");

        return (
          '<div class="table-summary-block">' +
          '<div class="table-summary-heading"><strong>' + escapeHtml(sectionRow.sectionLabel + " Summary") + '</strong><span>' + escapeHtml(selectedDataset.subjectWell + " • " + (selectedDataset.rigLabel || "Rig not mapped")) + '</span></div>' +
          '<div class="table-wrap table-summary-table"><table><thead><tr>' +
          '<th>Num. #</th><th>Area</th><th>Rig Name</th><th>Well Name</th><th>Phase</th><th>Major Flat time OPS</th><th>Mean Flat Time (hrs)</th><th>Actual (hrs)</th>' +
          '</tr></thead><tbody>' + rowsHtml + '</tbody></table></div>' +
          '</div>'
        );
      }).join("");

      target.innerHTML = '<div class="table-summary-stack">' + blocks + '</div>';
    }

    function compareFlatTimeSectionSizes(left, right) {
      if (left === right) return 0;
      if (left === "__no_section__") return 1;
      if (right === "__no_section__") return -1;
      return Number(left) - Number(right) || left.localeCompare(right);
    }

    function getAvailableFlatTimeSectionSizes(datasets) {
      return Array.from(
        new Set(
          datasets.flatMap((dataset) =>
            dataset.groups.flatMap((group) =>
              group.activities.map((activity) => activity.sectionSize || extractFlatTimeSectionSize(activity.activity))
            )
          )
        )
      ).sort(compareFlatTimeSectionSizes);
    }

    function populateFlatTimeSectionOptions(datasets) {
      const current = ui.flatTimeSection.value;
      const options = ['<option value="">All section sizes</option>'];
      getAvailableFlatTimeSectionSizes(datasets).forEach((sectionSize) => {
        options.push(
          '<option value="' + escapeHtml(sectionSize) + '">' + escapeHtml(formatFlatTimeSectionSize(sectionSize)) + "</option>"
        );
      });
      ui.flatTimeSection.innerHTML = options.join("");
      const available = getAvailableFlatTimeSectionSizes(datasets);
      if (available.includes(current)) {
        ui.flatTimeSection.value = current;
      }
    }

    function filterFlatTimeDatasetsBySection(datasets, sectionSize) {
      if (!sectionSize) return datasets;

      return datasets
        .map((dataset) => {
          const groups = dataset.groups
            .map((group) => {
              const activities = group.activities.filter(
                (activity) => (activity.sectionSize || extractFlatTimeSectionSize(activity.activity)) === sectionSize
              );
              if (!activities.length) return null;
              return {
                groupName: group.groupName,
                activities,
                totalSubjectHours: activities.reduce((sum, activity) => sum + Number(activity.subjectHours || 0), 0),
                totalMeanHours: activities.reduce((sum, activity) => sum + Number(activity.meanHours || 0), 0),
                totalMedianHours: activities.reduce((sum, activity) => sum + Number(activity.medianHours || 0), 0),
              };
            })
            .filter(Boolean);

          if (!groups.length) return null;

          return {
            ...dataset,
            groups,
            totalSubjectHours: groups.reduce((sum, group) => sum + Number(group.totalSubjectHours || 0), 0),
            totalMeanHours: groups.reduce((sum, group) => sum + Number(group.totalMeanHours || 0), 0),
            totalMedianHours: groups.reduce((sum, group) => sum + Number(group.totalMedianHours || 0), 0),
          };
        })
        .filter(Boolean);
    }

    function annotateFlatTimeScopedBenchmarks(datasets) {
      if (!datasets.length) return datasets;

      const activityScopeMap = new Map();

      datasets.forEach((dataset) => {
        dataset.groups.forEach((group) => {
          group.activities.forEach((activity) => {
            const sectionSize = activity.sectionSize || extractFlatTimeSectionSize(activity.activity);
            const scopeKey = [sectionSize, group.groupName, activity.activity].join("||");
            const subjectHours = Number(activity.subjectHours || 0);
            if (subjectHours <= 0) return;
            if (!activityScopeMap.has(scopeKey)) activityScopeMap.set(scopeKey, []);
            activityScopeMap.get(scopeKey).push(subjectHours);
          });
        });
      });

      return datasets.map((dataset) => {
        const groups = dataset.groups.map((group) => {
          const activities = group.activities.map((activity) => {
            const sectionSize = activity.sectionSize || extractFlatTimeSectionSize(activity.activity);
            const scopeKey = [sectionSize, group.groupName, activity.activity].join("||");
            const values = activityScopeMap.get(scopeKey) || [];
            return {
              ...activity,
              meanHours: values.length ? average(values) : 0,
              medianHours: values.length ? percentile(values, 0.5) : 0,
            };
          });

          return {
            ...group,
            activities,
            totalSubjectHours: activities.reduce((sum, item) => sum + Number(item.subjectHours || 0), 0),
            totalMeanHours: activities.reduce((sum, item) => sum + Number(item.meanHours || 0), 0),
            totalMedianHours: activities.reduce((sum, item) => sum + Number(item.medianHours || 0), 0),
          };
        });

        return {
          ...dataset,
          groups,
          totalSubjectHours: groups.reduce((sum, group) => sum + Number(group.totalSubjectHours || 0), 0),
          totalMeanHours: groups.reduce((sum, group) => sum + Number(group.totalMeanHours || 0), 0),
          totalMedianHours: groups.reduce((sum, group) => sum + Number(group.totalMedianHours || 0), 0),
        };
      });
    }

    function buildPerfectFlatTimeSections(datasets) {
      const sectionMap = new Map();

      datasets.forEach((dataset) => {
        dataset.groups.forEach((group) => {
          group.activities.forEach((activity) => {
            const sectionSize = activity.sectionSize || extractFlatTimeSectionSize(activity.activity);
            if (sectionSize === "__no_section__") return;
            if (!sectionMap.has(sectionSize)) sectionMap.set(sectionSize, new Map());
            const activityMap = sectionMap.get(sectionSize);
            const value = Number(activity.subjectHours || 0);
            if (value <= 0) return;
            const current = activityMap.get(activity.activity);
            if (!current || value < current) activityMap.set(activity.activity, value);
          });
        });
      });

      return Array.from(sectionMap.entries())
        .map(([sectionSize, activityMap]) => ({
          sectionSize,
          bestHours: Array.from(activityMap.values()).reduce((sum, value) => sum + value, 0),
        }))
        .filter((item) => item.bestHours > 0)
        .sort((left, right) => Number(right.sectionSize) - Number(left.sectionSize));
    }

    function renderPerfectFlatTimeChart(target, datasets, metricKey) {
      const idealSections = buildPerfectFlatTimeSections(datasets);
      if (!idealSections.length) {
        target.innerHTML = '<div class="empty">No section-sized activities available to draw the perfect flat time curve.</div>';
        return;
      }

      const sectionOrder = Array.from(
        new Set(
          datasets.flatMap((dataset) =>
            dataset.groups.flatMap((group) =>
              group.activities
                .map((activity) => activity.sectionSize || extractFlatTimeSectionSize(activity.activity))
                .filter((sectionSize) => sectionSize && sectionSize !== "__no_section__")
            )
          )
        )
      ).sort((left, right) => Number(right) - Number(left) || left.localeCompare(right));

      const idealMap = new Map(idealSections.map((section) => [section.sectionSize, section.bestHours]));

      function buildSeries(label, color, sectionHours, isIdeal) {
        let cumulativeDays = 0;
        const points = sectionOrder.map((sectionSize, index) => {
          cumulativeDays += Number(sectionHours.get(sectionSize) || 0) / 24;
          return {
            sectionSize,
            cumulativeDays,
            depthIndex: index + 1,
          };
        });
        return { label, color, isIdeal, points };
      }

      const series = [
        buildSeries(
          "Ideal curve",
          "#1264d6",
          new Map(sectionOrder.map((sectionSize) => [sectionSize, Number(idealMap.get(sectionSize) || 0)])),
          true
        ),
        ...datasets.map((dataset, index) => {
          const sectionHours = new Map();
          dataset.groups.forEach((group) => {
            group.activities.forEach((activity) => {
              const sectionSize = activity.sectionSize || extractFlatTimeSectionSize(activity.activity);
              if (!sectionSize || sectionSize === "__no_section__") return;
              const value = Number(activity[metricKey] || 0);
              sectionHours.set(sectionSize, (sectionHours.get(sectionSize) || 0) + value);
            });
          });
          return buildSeries(
            dataset.subjectWell,
            FLAT_TIME_SERIES_COLORS[index % FLAT_TIME_SERIES_COLORS.length],
            sectionHours,
            false
          );
        }),
      ];

      const chartTheme = getChartTheme();
      const width = 960;
      const height = 430;
      const margin = { top: 30, right: 28, bottom: 52, left: 88 };
      const chartWidth = width - margin.left - margin.right;
      const chartHeight = height - margin.top - margin.bottom;
      const maxDays = Math.max(
        ...series.flatMap((line) => line.points.map((point) => point.cumulativeDays)),
        1
      );
      const maxDepth = Math.max(sectionOrder.length, 1);

      const scaledSeries = series.map((line) => ({
        ...line,
        points: line.points.map((point) => ({
          ...point,
          x: margin.left + (point.cumulativeDays / maxDays) * chartWidth,
          y: margin.top + ((point.depthIndex - 1) / Math.max(maxDepth - 1, 1)) * chartHeight,
        })),
      }));

      const xTicks = Array.from({ length: 6 }, (_, index) => {
        const value = (maxDays / 5) * index;
        const x = margin.left + (value / maxDays) * chartWidth;
        return (
          '<g>' +
          '<line x1="' + x.toFixed(2) + '" y1="' + margin.top + '" x2="' + x.toFixed(2) + '" y2="' + (height - margin.bottom) + '" stroke="' + chartTheme.grid + '" stroke-dasharray="4 5"></line>' +
          '<text x="' + x.toFixed(2) + '" y="' + (height - 12) + '" text-anchor="middle" font-size="11" fill="' + chartTheme.valueLabel + '">' + escapeHtml(formatNumber(value)) + "</text>" +
          "</g>"
        );
      }).join("");

      const yTicks = sectionOrder
        .map((sectionSize, index) => {
          const y = margin.top + (index / Math.max(maxDepth - 1, 1)) * chartHeight;
          return (
            '<g>' +
            '<line x1="' + margin.left + '" y1="' + y.toFixed(2) + '" x2="' + (width - margin.right) + '" y2="' + y.toFixed(2) + '" stroke="' + chartTheme.grid + '" stroke-dasharray="4 5"></line>' +
            '<text x="' + (margin.left - 12) + '" y="' + (y + 4).toFixed(2) + '" text-anchor="end" font-size="11" font-weight="700" fill="' + chartTheme.text + '">' + escapeHtml(formatFlatTimeSectionSize(sectionSize)) + "</text>" +
            '</g>'
          );
        })
        .join("");

      const lineSvg = scaledSeries
        .map((line) => {
          const path = line.points.map((point, index) => (index === 0 ? "M" : "L") + point.x.toFixed(2) + " " + point.y.toFixed(2)).join(" ");
          const endPoint = line.points[line.points.length - 1];
          const labelX = Math.min(width - margin.right + 4, endPoint.x + 10);
          const pointsSvg = line.points
            .map((point) => (
              '<circle cx="' + point.x.toFixed(2) + '" cy="' + point.y.toFixed(2) + '" r="' + (line.isIdeal ? "4.5" : "3.5") + '" fill="' + line.color + '" opacity="' + (line.isIdeal ? "1" : "0.85") + '"></circle>'
            ))
            .join("");
          return (
            '<g>' +
            '<path d="' + path + '" fill="none" stroke="' + line.color + '" stroke-width="' + (line.isIdeal ? "4.5" : "2.5") + '" stroke-linecap="round" stroke-linejoin="round" opacity="' + (line.isIdeal ? "1" : "0.9") + '"></path>' +
            pointsSvg +
            '<text x="' + labelX.toFixed(2) + '" y="' + (endPoint.y + (line.isIdeal ? -10 : 10)).toFixed(2) + '" font-size="11" font-weight="700" fill="' + chartTheme.pointLabel + '">' + escapeHtml(line.label + " • " + formatNumber(endPoint.cumulativeDays) + " d") + '</text>' +
            '</g>'
          );
        })
        .join("");

      const legend = series
        .map((line) => (
          '<span class="legend-item" style="margin-right:12px;">' +
          '<span class="legend-dot" style="background:' + line.color + '; width:14px; height:14px;"></span>' +
          escapeHtml(line.label) +
          '</span>'
        ))
        .join("");

      target.innerHTML =
        '<div class="legend" style="margin-bottom:12px; flex-wrap:wrap;">' + legend + '</div>' +
        '<div class="column-chart-wrap">' +
        '<svg class="column-chart-svg" viewBox="0 0 ' + width + " " + height + '" role="img" aria-label="Perfect flat time compared with wells">' +
        xTicks +
        yTicks +
        '<line x1="' + margin.left + '" y1="' + margin.top + '" x2="' + margin.left + '" y2="' + (height - margin.bottom) + '" stroke="' + chartTheme.axis + '"></line>' +
        '<line x1="' + margin.left + '" y1="' + (height - margin.bottom) + '" x2="' + (width - margin.right) + '" y2="' + (height - margin.bottom) + '" stroke="' + chartTheme.axis + '"></line>' +
        lineSvg +
        '<text x="' + (margin.left + chartWidth / 2).toFixed(2) + '" y="' + (height - 6) + '" text-anchor="middle" font-size="12" font-weight="700" fill="' + chartTheme.text + '">Days</text>' +
        '<text x="18" y="' + (margin.top + chartHeight / 2).toFixed(2) + '" text-anchor="middle" font-size="12" font-weight="700" fill="' + chartTheme.text + '" transform="rotate(-90 18 ' + (margin.top + chartHeight / 2).toFixed(2) + ')">Depth / Section Progression</text>' +
        "</svg>" +
        "</div>";
    }

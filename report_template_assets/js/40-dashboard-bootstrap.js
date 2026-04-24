    function renderWeeklyReport() {
      const range = getWeeklyReportDateRange();
      if (!range.start && !range.end) {
        return;
      }

      const bounds = {
        start: range.start,
        end: range.end,
      };

      const periodRows = dashboardData.interventions.filter((row) => {
        if (!row.date) return false;
        if (bounds.start && row.date < bounds.start) return false;
        if (bounds.end && row.date > bounds.end) return false;
        return true;
      });
      const cumulativeRows = dashboardData.interventions.filter((row) => row.date && bounds.end && row.date <= bounds.end);

      const weeklyCategories = buildCategorySummary(periodRows);
      const cumulativeCategories = buildCategorySummary(cumulativeRows);

      ui.weeklyReportTitle.textContent = "Weekly report for " + (bounds.start && bounds.end ? formatDateHuman(bounds.start) + " - " + formatDateHuman(bounds.end) : "selected range");
      ui.weeklyReportSubtitle.textContent = periodRows.length + " interventions, " + uniqueCount(periodRows, "rigName") + " rigs, " + uniqueCount(periodRows, "wellName") + " wells";
      ui.weeklyReportRange.textContent = bounds.start && bounds.end ? formatDateHuman(bounds.start) + " - " + formatDateHuman(bounds.end) : "Range pending";
      ui.weeklyBannerCopy.textContent = "Selected period: " + ui.weeklyReportRange.textContent + ". The blocks below follow the same report storytelling used in the weekly Excel workbook.";
      ui.weeklyBannerChip1.textContent = periodRows.length + " interventions";
      ui.weeklyBannerChip2.textContent = uniqueCount(periodRows, "rigName") + " active rigs";
      ui.weeklyBannerChip3.textContent = formatCurrency(periodRows.reduce((sum, row) => sum + row.costSavingValue + row.potentialAvoidanceValue, 0)) + " total impact";

      renderTable(
        ui.weeklyCategoryTable,
        ["Category", "Number of Interventions", "Rig Action Taken", "Validated", "Validation %"],
        weeklyCategories.map((item) => [
          item.label,
          String(item.interventions),
          String(item.rigAction),
          String(item.validated),
          formatPercent(item.validationRate),
        ])
      );

      renderMultiSeriesChart(ui.weeklyCategoryChart, weeklyCategories, [
        { key: "interventions", label: "Interventions", color: "#1264d6" },
        { key: "rigAction", label: "Rig Action", color: "#0f766e" },
        { key: "validated", label: "Validated", color: "#be123c" },
        { key: "validationRate", label: "Validation %", color: "#c06a0a", format: (value) => formatPercent(value), scale: (value, context) => value * context.primaryMax, isSecondary: true },
      ]);

      renderTable(
        ui.cumulativeCategoryTable,
        ["Category", "# of Interventions", "Rig Action Taken", "Validated", "Validation %"],
        cumulativeCategories.map((item) => [
          item.label,
          String(item.interventions),
          String(item.rigAction),
          String(item.validated),
          formatPercent(item.validationRate),
        ])
      );

      renderMultiSeriesChart(ui.cumulativeCategoryChart, cumulativeCategories, [
        { key: "interventions", label: "Interventions", color: "#1264d6" },
        { key: "rigAction", label: "Rig Action", color: "#0f766e" },
        { key: "validated", label: "Validated", color: "#be123c" },
        { key: "validationRate", label: "Validation %", color: "#c06a0a", format: (value) => formatPercent(value), scale: (value, context) => value * context.primaryMax, isSecondary: true },
      ]);

      const wiperSummary = buildRigSummary(periodRows, isWiperTripRow);
      const ropSummary = buildRigSummary(periodRows, isRopRow);
      const kpiSummary = buildRigSummary(periodRows, isKpiRow);
      const summarySeries = [
        { key: "savedTime", label: "Saved Time", color: "#1264d6", format: (value) => formatNumber(value) },
        { key: "lossTime", label: "Loss Time", color: "#c81e5a", format: (value) => formatNumber(value) },
      ];

      renderSummaryBlock(ui.wiperSummaryTable, ui.wiperSummaryChart, wiperSummary, summarySeries);
      renderSummaryBlock(ui.ropSummaryTable, ui.ropSummaryChart, ropSummary, summarySeries);
      renderSummaryBlock(ui.kpiSummaryTable, ui.kpiSummaryChart, kpiSummary, summarySeries);

      const actualWeek = buildHighlightEntries(periodRows, "actual");
      const actualYtd = buildHighlightEntries(cumulativeRows, "actual");
      const potentialWeek = buildHighlightEntries(periodRows, "potential");
      const potentialYtd = buildHighlightEntries(cumulativeRows, "potential");
      const weeklyStatsRows = buildWeeklyStatsRows(periodRows, cumulativeRows);
      const monitoredThisWeekDays = daysBetweenInclusive(bounds.start, bounds.end);
      const monitoredTotalDays = daysBetweenInclusive(dashboardData.meta.monitoringStartDate, bounds.end);

      renderWeeklyMetrics(actualWeek, actualYtd, potentialWeek, potentialYtd);
      renderHighlightTable(ui.actualHighlightsTable, actualWeek, actualYtd, true);
      renderHighlightTable(ui.potentialHighlightsTable, potentialWeek, potentialYtd, false);
      renderWeeklyStatsMetrics(ui.weeklyStatsMetrics, monitoredThisWeekDays, monitoredTotalDays);
      renderWeeklyStatsTable(ui.weeklyStatsTable, weeklyStatsRows);
      renderWeeklyMonitoredWellsTable(ui.weeklyMonitoredWellsTable, weeklyStatsRows);
    }

    function setActiveView(viewId) {
      ui.viewPanels.forEach((panel) => {
        panel.hidden = panel.id !== viewId;
      });
      ui.viewTabs.forEach((button) => {
        const shouldBeActive =
          button.dataset.view === viewId ||
          (button.dataset.view === "dashboard-view" && viewId === "weekly-report-view");
        button.classList.toggle("is-active", shouldBeActive);
      });
    }

    function setDashboardMode(mode) {
      const resolvedMode = mode === "weekly" ? "weekly" : "interactive";
      if (ui.dashboardMode) {
        ui.dashboardMode.value = resolvedMode;
      }
      if (resolvedMode === "weekly") {
        setActiveView("weekly-report-view");
        renderWeeklyReport();
      } else {
        setActiveView("dashboard-view");
      }
    }

    function exportWeeklyReportPdf() {
      setDashboardMode("weekly");
      renderWeeklyReport();
      window.setTimeout(() => {
        window.print();
      }, 80);
    }

    function getActiveFilters() {
      return {
        startDate: ui.startDate.value,
        endDate: ui.endDate.value,
        week: ui.week.value,
        month: ui.month.value,
        rig: ui.rig.value,
        field: ui.field.value,
        well: ui.well.value,
        category: ui.category.value,
        type: ui.type.value,
        app: ui.app.value,
        validation: ui.validation.value,
        granularity: ui.granularity.value,
        search: ui.search.value.trim().toLowerCase(),
      };
    }

    function rowMatches(row, filters) {
      if (filters.startDate && row.date && row.date < filters.startDate) return false;
      if (filters.endDate && row.date && row.date > filters.endDate) return false;
      if (filters.week && row.week !== filters.week) return false;
      if (filters.month && row.month !== filters.month) return false;
      if (filters.rig && row.rigName !== filters.rig) return false;
      if (filters.field && row.field !== filters.field) return false;
      if (filters.well && row.wellName !== filters.well) return false;
      if (filters.category && row.category !== filters.category) return false;
      if (filters.type && row.type !== filters.type) return false;
      if (filters.app && row.app !== filters.app) return false;
      if (filters.validation === "validated" && !row.isValidated) return false;
      if (filters.validation === "not_validated" && row.isValidated) return false;
      if (filters.search && !row.searchText.includes(filters.search)) return false;
      return true;
    }

    function resetFilters() {
      ui.startDate.value = "";
      ui.endDate.value = "";
      ui.week.value = "";
      ui.month.value = "";
      ui.rig.value = "";
      ui.field.value = "";
      ui.well.value = "";
      ui.category.value = "";
      ui.type.value = "";
      ui.app.value = "";
      ui.validation.value = "";
      ui.granularity.value = "day";
      ui.search.value = "";
      ui.presetButtons.forEach((button) => button.classList.remove("is-active"));
      applyFilters();
    }

    function applyPreset(preset) {
      const defaultRange = getDefaultLastTuesdayRange();
      const referenceEndDate = ui.endDate.value || defaultRange.end;
      if (!referenceEndDate) return;
      const weeklyAllStartDate = dashboardData.meta.monitoringStartDate || dashboardData.meta.minDate || "";
      const end = new Date(referenceEndDate + "T00:00:00");
      let start = null;

      if (preset === "last7") start = new Date(end.getTime() - 6 * 86400000);
      if (preset === "last30") start = new Date(end.getTime() - 29 * 86400000);
      if (preset === "last90") start = new Date(end.getTime() - 89 * 86400000);

      ui.presetButtons.forEach((button) => button.classList.toggle("is-active", button.dataset.preset === preset));
      ui.week.value = "";
      ui.month.value = "";

      if (preset === "all") {
        ui.startDate.value = "";
        ui.endDate.value = "";
        ui.weeklyReportStartDate.value = weeklyAllStartDate;
        ui.weeklyReportEndDate.value = referenceEndDate;
      } else if (start) {
        const startDate = start.toISOString().slice(0, 10);
        ui.startDate.value = startDate;
        ui.endDate.value = referenceEndDate;
        ui.weeklyReportStartDate.value = startDate;
        ui.weeklyReportEndDate.value = referenceEndDate;
      }

      applyFilters();
      renderWeeklyReport();
    }

    function updateSectionVisibility() {
      ui.toggles.forEach((toggle) => {
        const target = document.getElementById(toggle.dataset.target);
        if (target) target.hidden = !toggle.checked;
      });
    }

    function renderKpis(filteredRows) {
      const validated = filteredRows.filter((row) => row.isValidated).length;
      const validationRate = filteredRows.length ? (validated / filteredRows.length) * 100 : 0;
      const costSavingHours = filteredRows.reduce((sum, row) => sum + row.costSavingHours, 0);
      const potentialAvoidanceHours = filteredRows.reduce((sum, row) => sum + row.potentialAvoidanceHours, 0);
      const costSavingValue = filteredRows.reduce((sum, row) => sum + row.costSavingValue, 0);
      const potentialAvoidanceValue = filteredRows.reduce((sum, row) => sum + row.potentialAvoidanceValue, 0);
      const avgSpreadRate = average(filteredRows.map((row) => row.rigSpreadRate));

      const cards = [
        {
          label: "Interventions",
          value: String(filteredRows.length),
          meta: validated + " validated (" + validationRate.toFixed(1) + "%)",
        },
        {
          label: "Coverage",
          value: String(uniqueCount(filteredRows, "rigName")),
          meta: uniqueCount(filteredRows, "field") + " fields and " + uniqueCount(filteredRows, "wellName") + " wells",
        },
        {
          label: "Hours Impact",
          value: formatNumber(costSavingHours + potentialAvoidanceHours),
          meta: formatNumber(costSavingHours) + " saved + " + formatNumber(potentialAvoidanceHours) + " avoided",
        },
        {
          label: "Financial Impact",
          value: formatCurrency(costSavingValue + potentialAvoidanceValue),
          meta: formatCurrency(costSavingValue) + " saved + " + formatCurrency(potentialAvoidanceValue) + " avoided",
        },
        {
          label: "Avg Spread Rate",
          value: formatCurrency(avgSpreadRate),
          meta: "Average across filtered intervention rows",
        },
      ];

      ui.kpiGrid.innerHTML = cards
        .map((card) => {
          return (
            '<article class="card">' +
            '<div class="card-label">' + escapeHtml(card.label) + "</div>" +
            '<div class="card-value">' + escapeHtml(card.value) + "</div>" +
            '<div class="card-meta">' + escapeHtml(card.meta) + "</div>" +
            "</article>"
          );
        })
        .join("");
    }

    function renderFilterChips(filters) {
      const chips = [];
      if (filters.startDate) chips.push("Start: " + filters.startDate);
      if (filters.endDate) chips.push("End: " + filters.endDate);
      if (filters.week) chips.push("Week: " + filters.week);
      if (filters.month) chips.push("Month: " + filters.month);
      if (filters.rig) chips.push("Rig: " + filters.rig);
      if (filters.field) chips.push("Field: " + filters.field);
      if (filters.well) chips.push("Well: " + filters.well);
      if (filters.category) chips.push("Category: " + filters.category);
      if (filters.type) chips.push("Type: " + filters.type);
      if (filters.app) chips.push("App: " + filters.app);
      if (filters.validation === "validated") chips.push("Validated only");
      if (filters.validation === "not_validated") chips.push("Not validated only");
      if (filters.search) chips.push('Search: "' + filters.search + '"');
      ui.activeFilters.innerHTML = chips.length
        ? chips.map((chip) => '<span class="chip">' + escapeHtml(chip) + "</span>").join("")
        : '<span class="chip">No active filters. Showing the full dataset.</span>';
    }

    function renderRankings(filteredRows) {
      renderTable(
        ui.categoryTable,
        ["Category", "Count"],
        buildCounter(filteredRows, "category").slice(0, 10).map((item) => [item.label, String(item.value)])
      );
      renderTable(
        ui.fieldTable,
        ["Field", "Count"],
        buildCounter(filteredRows, "field").slice(0, 10).map((item) => [item.label, String(item.value)])
      );
      renderTable(
        ui.wellTable,
        ["Well", "Count"],
        buildCounter(filteredRows, "wellName").slice(0, 10).map((item) => [item.label, String(item.value)])
      );
    }

    function renderDetails(filteredRows) {
      const rows = filteredRows
        .slice()
        .sort((left, right) => (right.date || "").localeCompare(left.date || ""))
        .slice(0, 150)
        .map((row) => [
          row.index,
          row.date,
          row.week,
          row.rigName,
          row.field,
          row.wellName,
          row.category,
          row.type,
          row.app,
          row.isValidated ? "Yes" : "No",
          row.description.slice(0, 90),
        ]);

      renderTable(
        ui.interventionTable,
        ["#", "Date", "Week", "Rig", "Field", "Well", "Category", "Type", "App", "Validated", "Description"],
        rows
      );
    }

    function renderCostAvoidance(filteredRows) {
      const items = buildValueCounter(
        filteredRows.filter((row) => row.costSavingValue > 0),
        "rigName",
        "costSavingValue"
      );
      renderBarChart(ui.caChart, items, "#c81e5a", (value) => formatCurrency(value), { maxItems: 0 });
    }

    function applyFilters() {
      const filters = getActiveFilters();
      const filteredRows = dashboardData.interventions.filter((row) => rowMatches(row, filters));

      ui.resultsTitle.textContent = filteredRows.length + " interventions in view";
      ui.resultsSubtitle.textContent =
        uniqueCount(filteredRows, "rigName") + " rigs, " +
        uniqueCount(filteredRows, "wellName") + " wells";

      renderFilterChips(filters);
      renderKpis(filteredRows);
      renderTrendChart(ui.trendChart, buildTrend(filteredRows, filters.granularity));
      renderBarChart(ui.categoryChart, buildCounter(filteredRows, "category"), "#1264d6", (value) => String(value));
      renderBarChart(ui.rigChart, buildCounter(filteredRows, "rigName"), "#0f766e", (value) => String(value));
      renderBarChart(ui.typeChart, buildCounter(filteredRows, "type"), "#c06a0a", (value) => String(value));
      renderBarChart(ui.appChart, buildCounter(filteredRows, "app"), "#7c3aed", (value) => String(value));
      renderRankings(filteredRows);
      renderDetails(filteredRows);
      renderCostAvoidance(filteredRows);
      updateSectionVisibility();
    }

    function wireEvents() {
      [
        ui.startDate,
        ui.endDate,
        ui.week,
        ui.month,
        ui.rig,
        ui.field,
        ui.well,
        ui.category,
        ui.type,
        ui.app,
        ui.validation,
        ui.granularity,
      ].forEach((element) => {
        element.addEventListener("change", () => {
          ui.presetButtons.forEach((button) => button.classList.remove("is-active"));
          applyFilters();
        });
      });

      ui.search.addEventListener("input", applyFilters);
      ui.reset.addEventListener("click", resetFilters);
      ui.themeToggle.addEventListener("click", toggleTheme);
      ui.presetButtons.forEach((button) => {
        button.addEventListener("click", () => applyPreset(button.dataset.preset));
      });
      ui.toggles.forEach((toggle) => {
        toggle.addEventListener("change", updateSectionVisibility);
      });

      ui.viewTabs.forEach((button) => {
        button.addEventListener("click", () => {
          if (button.dataset.view === "dashboard-view") {
            setDashboardMode("interactive");
            return;
          }
          setActiveView(button.dataset.view);
        });
      });

      if (ui.dashboardMode) {
        ui.dashboardMode.addEventListener("change", () => {
          setDashboardMode(ui.dashboardMode.value);
        });
      }

      ui.weeklyExportPdf.addEventListener("click", exportWeeklyReportPdf);
      ui.weeklyReportStartDate.addEventListener("change", () => {
        ui.presetButtons.forEach((button) => button.classList.remove("is-active"));
        renderWeeklyReport();
      });
      ui.weeklyReportEndDate.addEventListener("change", () => {
        ui.presetButtons.forEach((button) => button.classList.remove("is-active"));
        renderWeeklyReport();
      });
      ui.flatTimeRig.addEventListener("change", renderFlatTime);
      ui.flatTimeSection.addEventListener("change", renderFlatTime);
      ui.flatTimeMetric.addEventListener("change", renderFlatTime);
      ui.flatTimeTopN.addEventListener("change", renderFlatTime);
      ui.flatTimeMode.addEventListener("change", renderFlatTime);
      if (ui.flatTimeAllWellsLayout) {
        ui.flatTimeAllWellsLayout.addEventListener("change", renderFlatTime);
      }
      ui.flatTimeHeatmapMode.addEventListener("change", () => {
        flatTimeState.heatmapMode = ui.flatTimeHeatmapMode.value || "gap";
        renderFlatTime();
      });
      ui.flatTimeWell.addEventListener("change", () => {
        flatTimeState.focusWell = ui.flatTimeWell.value || "";
        renderFlatTime();
      });
      ui.flatTimeRecalculate.addEventListener("click", renderFlatTime);
      ui.flatTimeClearUploads.addEventListener("click", () => {
        flatTimeState.uploadedDatasets = [];
        ui.flatTimeUpload.value = "";
        ui.flatTimeRig.value = "";
        ui.flatTimeSection.value = "";
        ui.flatTimeWell.value = "";
        ui.flatTimeHeatmapMode.value = "gap";
        flatTimeState.focusWell = "";
        flatTimeState.focusActivity = "";
        flatTimeState.focusActivities = [];
        flatTimeState.heatmapMode = "gap";
        renderFlatTime();
      });
      ui.flatTimeUpload.addEventListener("change", async () => {
        const files = Array.from(ui.flatTimeUpload.files || []);
        if (!files.length) return;
        const parsed = await Promise.all(
          files.map(async (file) => parseFlatTimeCsvText(file.name, await file.text()))
        );
        const prepared = parsed
          .flat()
          .filter((item) => item && item.groups && item.groups.length)
          .map((item) => ({ ...item, id: createFlatTimeUploadId(item.fileName, item.subjectWell) }));
        flatTimeState.uploadedDatasets = [
          ...flatTimeState.uploadedDatasets,
          ...prepared,
        ];
        renderFlatTime();
      });
    }

    function initialize() {
      applyTheme(localStorage.getItem(THEME_STORAGE_KEY) || "classic");
      populateSelect(ui.week, dashboardData.filters.weeks, "All weeks");
      populateSelect(ui.month, dashboardData.filters.months, "All months");
      populateSelect(ui.rig, dashboardData.filters.rigs, "All rigs");
      populateSelect(ui.field, dashboardData.filters.fields, "All fields");
      populateSelect(ui.well, dashboardData.filters.wells, "All wells");
      populateSelect(ui.category, dashboardData.filters.categories, "All categories");
      populateSelect(ui.type, dashboardData.filters.types, "All types");
      populateSelect(ui.app, dashboardData.filters.apps, "All apps");
      if (dashboardData.meta.minDate) ui.startDate.min = dashboardData.meta.minDate;
      if (dashboardData.meta.maxDate) ui.startDate.max = dashboardData.meta.maxDate;
      if (dashboardData.meta.minDate) ui.endDate.min = dashboardData.meta.minDate;
      if (dashboardData.meta.maxDate) ui.endDate.max = dashboardData.meta.maxDate;
      if (dashboardData.meta.minDate) ui.weeklyReportStartDate.min = dashboardData.meta.minDate;
      if (dashboardData.meta.maxDate) ui.weeklyReportStartDate.max = dashboardData.meta.maxDate;
      if (dashboardData.meta.minDate) ui.weeklyReportEndDate.min = dashboardData.meta.minDate;
      if (dashboardData.meta.maxDate) ui.weeklyReportEndDate.max = dashboardData.meta.maxDate;

      const defaultRange = getDefaultLastTuesdayRange();
      ui.startDate.value = defaultRange.start;
      ui.endDate.value = defaultRange.end;
      ui.weeklyReportStartDate.value = defaultRange.start;
      ui.weeklyReportEndDate.value = defaultRange.end;
      if (ui.dashboardMode) ui.dashboardMode.value = "interactive";

      wireEvents();
      setDashboardMode((ui.dashboardMode && ui.dashboardMode.value) || "interactive");
      updateSectionVisibility();
      applyFilters();
      renderWeeklyReport();
      renderFlatTime();
    }

    initialize();

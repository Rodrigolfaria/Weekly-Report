    function textBlob(row) {
      return [
        row.description,
        row.recommendation,
        row.justification,
        row.parameter,
        row.type,
        row.app,
        row.expected,
        row.actual,
        row.rigComment,
        row.rtocComments,
      ]
        .join(" ")
        .toLowerCase();
    }

    function isWiperTripRow(row) {
      return textBlob(row).includes("wiper");
    }

    function isRopRow(row) {
      return row.type.toLowerCase() === "rop" || row.parameter.toLowerCase() === "rop" || /rop/.test(textBlob(row));
    }

    function isKpiRow(row) {
      const parameter = row.parameter.toLowerCase();
      const type = row.type.toLowerCase();
      return (
        type === "kpi" ||
        ["w2w", "s2s", "tripping speed", "tripping speed ", "connection", "kpi", "time"].includes(parameter) ||
        textBlob(row).includes("kpi")
      );
    }

    function buildCategorySummary(rows) {
      return CATEGORY_ORDER.map((category) => {
        const matches = rows.filter((row) => normalizeCategory(row.category) === category);
        const interventions = matches.length;
        const rigAction = matches.filter((row) => isYesLike(row.rigAction)).length;
        const validated = matches.filter((row) => row.isValidated).length;
        return {
          label: category,
          interventions,
          rigAction,
          validated,
          validationRate: interventions ? validated / interventions : 0,
        };
      });
    }

    function buildRigSummary(rows, matcher) {
      const groups = new Map();
      rows.filter(matcher).forEach((row) => {
        const key = row.rigName + "||" + row.wellName;
        if (!groups.has(key)) {
          groups.set(key, {
            label: row.rigName + " / " + row.wellName,
            rig: row.rigName,
            well: row.wellName,
            interventions: 0,
            rigAction: 0,
            savedTime: 0,
            lossTime: 0,
          });
        }
        const entry = groups.get(key);
        entry.interventions += 1;
        entry.rigAction += isYesLike(row.rigAction) ? 1 : 0;
        entry.savedTime += row.costSavingHours;
        entry.lossTime += row.potentialAvoidanceHours;
      });

      const items = Array.from(groups.values()).sort(
        (left, right) =>
          (right.savedTime + right.lossTime + right.interventions) -
            (left.savedTime + left.lossTime + left.interventions) ||
          left.label.localeCompare(right.label)
      );

      const totals = items.reduce(
        (acc, item) => {
          acc.interventions += item.interventions;
          acc.rigAction += item.rigAction;
          acc.savedTime += item.savedTime;
          acc.lossTime += item.lossTime;
          return acc;
        },
        { interventions: 0, rigAction: 0, savedTime: 0, lossTime: 0 }
      );

      return { items, totals };
    }

    function providerLabel(row) {
      return row.engDept || row.optDept || "N/A";
    }

    function buildHighlightEntries(rows, mode) {
      const isActual = mode === "actual";
      return rows
        .filter((row) =>
          isActual
            ? row.costSavingHours > 0 || row.costSavingValue > 0
            : row.potentialAvoidanceHours > 0 || row.potentialAvoidanceValue > 0
        )
        .sort((left, right) =>
          isActual
            ? right.costSavingValue - left.costSavingValue || (right.date || "").localeCompare(left.date || "")
            : right.potentialAvoidanceValue - left.potentialAvoidanceValue || (right.date || "").localeCompare(left.date || "")
        )
        .map((row) => ({
          week: row.week,
          date: row.date,
          rig: row.rigName,
          well: row.wellName,
          provider: providerLabel(row),
          action: isActual ? row.description || row.recommendation : row.recommendation || row.description,
          hours: isActual ? row.costSavingHours : row.potentialAvoidanceHours,
          value: isActual ? row.costSavingValue : row.potentialAvoidanceValue,
        }));
    }

    function renderWeeklyMetrics(actualWeek, actualYtd, potentialWeek, potentialYtd) {
      const cards = [
        {
          label: "Total This Week",
          value: formatNumber(actualWeek.reduce((sum, row) => sum + row.hours, 0)),
          suffix: "hours saved",
          meta: formatCurrency(actualWeek.reduce((sum, row) => sum + row.value, 0)) + " value realized",
        },
        {
          label: "Total YTD",
          value: formatNumber(actualYtd.reduce((sum, row) => sum + row.hours, 0)),
          suffix: "hours saved",
          meta: formatCurrency(actualYtd.reduce((sum, row) => sum + row.value, 0)) + " value realized",
        },
        {
          label: "Potential This Week",
          value: formatNumber(potentialWeek.reduce((sum, row) => sum + row.hours, 0)),
          suffix: "hours potential",
          meta: formatCurrency(potentialWeek.reduce((sum, row) => sum + row.value, 0)) + " value potential",
        },
        {
          label: "Potential YTD",
          value: formatNumber(potentialYtd.reduce((sum, row) => sum + row.hours, 0)),
          suffix: "hours potential",
          meta: formatCurrency(potentialYtd.reduce((sum, row) => sum + row.value, 0)) + " value potential",
        },
      ];

      ui.weeklyHighlightMetrics.innerHTML = cards
        .map(
          (card) =>
            '<div class="metric-pill">' +
            '<div class="label">' + escapeHtml(card.label) + "</div>" +
            '<div class="value"><span class="value-main">' + escapeHtml(card.value) + '</span><span class="value-suffix">' + escapeHtml(card.suffix) + "</span></div>" +
            '<div class="meta">' + escapeHtml(card.meta) + "</div>" +
            "</div>"
        )
        .join("");
    }

    function dominantGroupValue(rows, key, fallback = "N/A") {
      const counts = new Map();
      rows.forEach((row) => {
        const value = String(row?.[key] || "").trim();
        if (!value) return;
        counts.set(value, (counts.get(value) || 0) + 1);
      });
      const top = Array.from(counts.entries())
        .sort((left, right) => right[1] - left[1] || left[0].localeCompare(right[0]))[0];
      return top ? top[0] : fallback;
    }

    function resolveWeeklyMonitoredType(rows, wellName, fieldName) {
      const dominantType = dominantGroupValue(rows, "type", "");
      if (dominantType && dominantType.toLowerCase() !== "n/a") {
        return dominantType;
      }

      const wellText = String(wellName || "").toUpperCase();
      const fieldText = String(fieldName || "").toUpperCase();
      if (wellText.includes("HRDH") || fieldText.includes("HRDH")) {
        return "Gas";
      }

      return "N/A";
    }

    function countAndValidityCell(bucket) {
      return String(bucket.count) + " / " + bucket.validity;
    }

    function renderHighlightTable(target, weekEntries, ytdEntries, isActual) {
      const headers = isActual
        ? ["Rig", "Well", "Provider / Dept", "Operations / Action", "Saved Time (hrs)", "Cost Saving (US$)"]
        : ["Rig", "Well", "Provider / Dept", "Operations / Action", "Potential Saved Time (hrs)", "Potential Cost Saving/Avoidance (US$)"];

      const bodyRows = weekEntries
        .slice(0, 25)
        .map((row) => {
          return (
            "<tr>" +
            "<td>" + escapeHtml(row.rig) + "</td>" +
            "<td>" + escapeHtml(row.well) + "</td>" +
            "<td>" + escapeHtml(row.provider) + "</td>" +
            "<td>" + escapeHtml(row.action) + "</td>" +
            "<td>" + escapeHtml(formatNumber(row.hours)) + "</td>" +
            "<td>" + escapeHtml(formatCurrency(row.value)) + "</td>" +
            "</tr>"
          );
        })
        .join("");

      const totalWeekHours = weekEntries.reduce((sum, row) => sum + row.hours, 0);
      const totalWeekValue = weekEntries.reduce((sum, row) => sum + row.value, 0);
      const totalYtdHours = ytdEntries.reduce((sum, row) => sum + row.hours, 0);
      const totalYtdValue = ytdEntries.reduce((sum, row) => sum + row.value, 0);

      const totalsRows =
        '<tr>' +
        '<td colspan="4" style="text-align:right; font-weight:700;">Total This Week</td>' +
        '<td><strong>' + escapeHtml(formatNumber(totalWeekHours)) + "</strong></td>" +
        '<td><strong>' + escapeHtml(formatCurrency(totalWeekValue)) + "</strong></td>" +
        "</tr>" +
        '<tr>' +
        '<td colspan="4" style="text-align:right; font-weight:700;">Total YTD</td>' +
        '<td><strong>' + escapeHtml(formatNumber(totalYtdHours)) + "</strong></td>" +
        '<td><strong>' + escapeHtml(formatCurrency(totalYtdValue)) + "</strong></td>" +
        "</tr>";

      target.innerHTML =
        '<div class="table-wrap"><table><thead><tr>' +
        headers.map((header) => "<th>" + escapeHtml(header) + "</th>").join("") +
        "</tr></thead><tbody>" +
        bodyRows +
        totalsRows +
        "</tbody></table></div>";
    }

    function renderSummaryBlock(tableTarget, chartTarget, summary, chartSeries) {
      const tableRows = summary.items
        .map((item) => [
          item.rig,
          item.well,
          String(item.interventions),
          String(item.rigAction),
          formatNumber(item.savedTime),
          formatNumber(item.lossTime),
        ])
        .concat([
          [
            "Total",
            "",
            String(summary.totals.interventions),
            String(summary.totals.rigAction),
            formatNumber(summary.totals.savedTime),
            formatNumber(summary.totals.lossTime),
          ],
        ]);

      renderTable(
        tableTarget,
        ["Rig Name", "Well Name", "Number of Interventions", "Rig Action Taken", "Saved Time (hrs)", "Loss Time (hrs)"],
        tableRows
      );

      renderMultiSeriesChart(chartTarget, summary.items, chartSeries);
    }

    function daysBetweenInclusive(start, end) {
      if (!start || !end) return 0;
      const startDate = new Date(start + "T00:00:00");
      const endDate = new Date(end + "T00:00:00");
      const diff = Math.round((endDate - startDate) / 86400000);
      return diff >= 0 ? diff + 1 : 0;
    }

    function buildWeeklyStatsRows(rows, cumulativeRows) {
      const groups = new Map();
      const cumulativeGroups = new Map();

      rows.forEach((row) => {
        const key = row.rigName + "||" + row.wellName;
        if (!groups.has(key)) {
          groups.set(key, { rig: row.rigName, well: row.wellName, rows: [] });
        }
        groups.get(key).rows.push(row);
      });

      cumulativeRows.forEach((row) => {
        const key = row.rigName + "||" + row.wellName;
        if (!cumulativeGroups.has(key)) {
          cumulativeGroups.set(key, []);
        }
        cumulativeGroups.get(key).push(row);
      });

      return Array.from(groups.values())
        .sort((left, right) => left.rig.localeCompare(right.rig) || left.well.localeCompare(right.well))
        .map((group) => {
          const byCategory = (categoryName) => {
            const matches = group.rows.filter((row) => normalizeCategory(row.category) === categoryName);
            const count = matches.length;
            const validated = matches.filter((row) => row.isValidated).length;
            return {
              count,
              validity: count ? formatPercent(validated / count) : "0.0%",
            };
          };

          const totalCount = group.rows.length;
          const totalValidated = group.rows.filter((row) => row.isValidated).length;
          const key = group.rig + "||" + group.well;
          const cumulativeGroupRows = cumulativeGroups.get(key) || [];
          const monitoredThisWeek = new Set(group.rows.map((row) => row.date).filter(Boolean)).size;
          const monitoredSinceStart = new Set(cumulativeGroupRows.map((row) => row.date).filter(Boolean)).size;

          return {
            rig: group.rig,
            well: group.well,
            section: dominantGroupValue(group.rows, "holeSize"),
            field: dominantGroupValue(group.rows, "field"),
            type: resolveWeeklyMonitoredType(group.rows, group.well, dominantGroupValue(group.rows, "field")),
            optimization: byCategory("Optimization"),
            stuckPipe: byCategory("Stuck pipe"),
            wellControl: byCategory("Well Control"),
            operationalCompliance: byCategory("Operational Compliance"),
            reporting: byCategory("Reporting"),
            total: {
              count: totalCount,
              validity: totalCount ? formatPercent(totalValidated / totalCount) : "0.0%",
            },
            monitoredThisWeek,
            monitoredSinceStart,
          };
        });
    }

    function renderWeeklyStatsMetrics(target, thisWeekDays, totalDays) {
      const cards = [
        {
          label: "Days Monitored This Week",
          value: String(thisWeekDays),
          suffix: "days",
          meta: "Selected reporting range",
        },
        {
          label: "Days Monitored Since Start",
          value: String(totalDays),
          suffix: "days",
          meta: "From monitoring start date to selected end date",
        },
      ];

      target.innerHTML = cards
        .map(
          (card) =>
            '<div class="metric-pill">' +
            '<div class="label">' + escapeHtml(card.label) + "</div>" +
            '<div class="value"><span class="value-main">' + escapeHtml(card.value) + '</span><span class="value-suffix">' + escapeHtml(card.suffix) + "</span></div>" +
            '<div class="meta">' + escapeHtml(card.meta) + "</div>" +
            "</div>"
        )
        .join("");
    }

    function renderWeeklyStatsTable(target, rows) {
      if (!rows.length) {
        target.innerHTML = '<div class="empty">No intervention statistics available for the selected period.</div>';
        return;
      }

      const headerHtml =
        "<thead>" +
        '<tr><th colspan="8">Weekly Intervention Statistics and Analysis</th></tr>' +
        '<tr>' +
        '<th>Rig Name</th><th>Well Name</th>' +
        '<th>Optimization<br><span style="font-weight:500;">Interventions / Validity</span></th>' +
        '<th>Stuck Pipe<br><span style="font-weight:500;">Interventions / Validity</span></th>' +
        '<th>Well Control<br><span style="font-weight:500;">Interventions / Validity</span></th>' +
        '<th>Operational Compliance<br><span style="font-weight:500;">Interventions / Validity</span></th>' +
        '<th>Reporting<br><span style="font-weight:500;">Interventions / Validity</span></th>' +
        '<th>Total<br><span style="font-weight:500;">Interventions / Validity</span></th>' +
        "</tr>" +
        "</thead>";

      const totalRow = rows.reduce(
        (acc, row) => {
          acc.optimization.count += row.optimization.count;
          acc.optimization.validityCount += row.optimization.count ? Number(row.optimization.validity.replace("%", "")) * row.optimization.count : 0;
          acc.optimization.den += row.optimization.count;
          acc.stuckPipe.count += row.stuckPipe.count;
          acc.stuckPipe.validityCount += row.stuckPipe.count ? Number(row.stuckPipe.validity.replace("%", "")) * row.stuckPipe.count : 0;
          acc.stuckPipe.den += row.stuckPipe.count;
          acc.wellControl.count += row.wellControl.count;
          acc.wellControl.validityCount += row.wellControl.count ? Number(row.wellControl.validity.replace("%", "")) * row.wellControl.count : 0;
          acc.wellControl.den += row.wellControl.count;
          acc.operationalCompliance.count += row.operationalCompliance.count;
          acc.operationalCompliance.validityCount += row.operationalCompliance.count ? Number(row.operationalCompliance.validity.replace("%", "")) * row.operationalCompliance.count : 0;
          acc.operationalCompliance.den += row.operationalCompliance.count;
          acc.reporting.count += row.reporting.count;
          acc.reporting.validityCount += row.reporting.count ? Number(row.reporting.validity.replace("%", "")) * row.reporting.count : 0;
          acc.reporting.den += row.reporting.count;
          acc.monitoredThisWeek += row.monitoredThisWeek;
          acc.monitoredSinceStart += row.monitoredSinceStart;
          return acc;
        },
        {
          optimization: { count: 0, validityCount: 0, den: 0 },
          stuckPipe: { count: 0, validityCount: 0, den: 0 },
          wellControl: { count: 0, validityCount: 0, den: 0 },
          operationalCompliance: { count: 0, validityCount: 0, den: 0 },
          reporting: { count: 0, validityCount: 0, den: 0 },
          monitoredThisWeek: 0,
          monitoredSinceStart: 0,
        }
      );

      const pct = (bucket) => bucket.den ? (bucket.validityCount / bucket.den).toFixed(1) + "%" : "0.0%";

      const bodyHtml = rows
        .map((row) => {
          return (
            "<tr>" +
            "<td>" + escapeHtml(row.rig) + "</td>" +
            "<td>" + escapeHtml(row.well) + "</td>" +
            "<td>" + escapeHtml(countAndValidityCell(row.optimization)) + "</td>" +
            "<td>" + escapeHtml(countAndValidityCell(row.stuckPipe)) + "</td>" +
            "<td>" + escapeHtml(countAndValidityCell(row.wellControl)) + "</td>" +
            "<td>" + escapeHtml(countAndValidityCell(row.operationalCompliance)) + "</td>" +
            "<td>" + escapeHtml(countAndValidityCell(row.reporting)) + "</td>" +
            "<td><strong>" + escapeHtml(countAndValidityCell(row.total)) + "</strong></td>" +
            "</tr>"
          );
        })
        .join("") +
        (
          "<tr>" +
          "<td><strong>Total</strong></td>" +
          "<td></td>" +
          "<td><strong>" + escapeHtml(String(totalRow.optimization.count) + " / " + pct(totalRow.optimization)) + "</strong></td>" +
          "<td><strong>" + escapeHtml(String(totalRow.stuckPipe.count) + " / " + pct(totalRow.stuckPipe)) + "</strong></td>" +
          "<td><strong>" + escapeHtml(String(totalRow.wellControl.count) + " / " + pct(totalRow.wellControl)) + "</strong></td>" +
          "<td><strong>" + escapeHtml(String(totalRow.operationalCompliance.count) + " / " + pct(totalRow.operationalCompliance)) + "</strong></td>" +
          "<td><strong>" + escapeHtml(String(totalRow.reporting.count) + " / " + pct(totalRow.reporting)) + "</strong></td>" +
          "<td><strong>" + escapeHtml(String(
            totalRow.optimization.count +
            totalRow.stuckPipe.count +
            totalRow.wellControl.count +
            totalRow.operationalCompliance.count +
            totalRow.reporting.count
          )) + "</strong></td>" +
          "</tr>"
        );

      target.innerHTML = '<div class="table-wrap"><table class="stats-table">' + headerHtml + "<tbody>" + bodyHtml + "</tbody></table></div>";
    }

    function renderWeeklyMonitoredWellsTable(target, rows) {
      if (!rows.length) {
        target.innerHTML = '<div class="empty">No monitored wells available for the selected period.</div>';
        return;
      }
      const overrides = getWeeklyMonitoredOverrides();
      const effectiveRows = rows.map((row) => {
        const rowKey = weeklyMonitoredRowKey(row);
        const rowOverrides = overrides[rowKey] || {};
        return {
          rowKey,
          rig: rowOverrides.rig || row.rig || "",
          well: rowOverrides.well || row.well || "",
          section: rowOverrides.section || row.section || "",
          field: rowOverrides.field || row.field || "",
          type: rowOverrides.type || row.type || "",
          monitoredThisWeek: parseEditableNumber(rowOverrides.monitoredThisWeek, row.monitoredThisWeek),
          monitoredSinceStart: parseEditableNumber(rowOverrides.monitoredSinceStart, row.monitoredSinceStart),
          totalInterventions: parseEditableNumber(rowOverrides.totalInterventions, row.total.count),
          validatedPercent: parseEditablePercent(rowOverrides.validatedPercent, Number(String(row.total.validity || "0").replace("%", ""))),
        };
      });

      const totalRow = effectiveRows.reduce(
        (acc, row) => {
          acc.monitoredThisWeek += row.monitoredThisWeek;
          acc.monitoredSinceStart += row.monitoredSinceStart;
          acc.totalInterventions += row.totalInterventions;
          acc.validatedWeighted += row.totalInterventions ? row.validatedPercent * row.totalInterventions : 0;
          acc.validatedDen += row.totalInterventions;
          return acc;
        },
        { monitoredThisWeek: 0, monitoredSinceStart: 0, totalInterventions: 0, validatedWeighted: 0, validatedDen: 0 }
      );

      const bodyHtml = effectiveRows
        .map((row) => {
          return (
            "<tr>" +
            '<td><input type="text" class="table-inline-input" data-weekly-monitored-key="' + escapeHtml(row.rowKey) + '" data-weekly-monitored-field="rig" value="' + escapeHtml(row.rig) + '" placeholder="Rig"></td>' +
            '<td><input type="text" class="table-inline-input" data-weekly-monitored-key="' + escapeHtml(row.rowKey) + '" data-weekly-monitored-field="well" value="' + escapeHtml(row.well) + '" placeholder="Well"></td>' +
            '<td><input type="text" class="table-inline-input" data-weekly-monitored-key="' + escapeHtml(row.rowKey) + '" data-weekly-monitored-field="section" value="' + escapeHtml(row.section) + '" placeholder="Section"></td>' +
            '<td><input type="text" class="table-inline-input" data-weekly-monitored-key="' + escapeHtml(row.rowKey) + '" data-weekly-monitored-field="field" value="' + escapeHtml(row.field) + '" placeholder="Field"></td>' +
            '<td><input type="text" class="table-inline-input" data-weekly-monitored-key="' + escapeHtml(row.rowKey) + '" data-weekly-monitored-field="type" value="' + escapeHtml(row.type) + '" placeholder="Type"></td>' +
            '<td><input type="text" class="table-inline-input" data-weekly-monitored-key="' + escapeHtml(row.rowKey) + '" data-weekly-monitored-field="monitoredThisWeek" value="' + escapeHtml(String(row.monitoredThisWeek)) + '" placeholder="Days"></td>' +
            '<td><input type="text" class="table-inline-input" data-weekly-monitored-key="' + escapeHtml(row.rowKey) + '" data-weekly-monitored-field="monitoredSinceStart" value="' + escapeHtml(String(row.monitoredSinceStart)) + '" placeholder="Days"></td>' +
            '<td><input type="text" class="table-inline-input" data-weekly-monitored-key="' + escapeHtml(row.rowKey) + '" data-weekly-monitored-field="totalInterventions" value="' + escapeHtml(String(row.totalInterventions)) + '" placeholder="Interventions"></td>' +
            '<td><input type="text" class="table-inline-input" data-weekly-monitored-key="' + escapeHtml(row.rowKey) + '" data-weekly-monitored-field="validatedPercent" value="' + escapeHtml(String(row.validatedPercent)) + '" placeholder="%"></td>' +
            "</tr>"
          );
        })
        .join("") +
        (
          "<tr>" +
          "<td><strong>Total</strong></td>" +
          "<td></td><td></td><td></td><td></td>" +
          "<td><strong>" + escapeHtml(String(totalRow.monitoredThisWeek)) + "</strong></td>" +
          "<td><strong>" + escapeHtml(String(totalRow.monitoredSinceStart)) + "</strong></td>" +
          "<td><strong>" + escapeHtml(String(totalRow.totalInterventions)) + "</strong></td>" +
          "<td><strong>" + escapeHtml(totalRow.validatedDen ? formatPercent((totalRow.validatedWeighted / 100) / totalRow.validatedDen) : "0.0%") + "</strong></td>" +
          "</tr>"
        );

      target.innerHTML =
        '<div class="table-wrap"><table class="stats-table"><thead><tr>' +
        '<th>Rig Name</th><th>Well Name</th><th>Section</th><th>Field</th><th>Type</th><th>Days Monitored (This Week)</th><th>Days Monitored (Cumulative)</th><th>Total Number of Interventions</th><th>Validated Interventions %</th>' +
        "</tr></thead><tbody>" +
        bodyHtml +
        "</tbody></table></div>";

      Array.from(target.querySelectorAll("[data-weekly-monitored-key]")).forEach((input) => {
        input.addEventListener("change", () => {
          setWeeklyMonitoredOverride(
            input.dataset.weeklyMonitoredKey || "",
            input.dataset.weeklyMonitoredField || "",
            input.value || ""
          );
          renderWeeklyMonitoredWellsTable(target, rows);
        });
      });
    }

    function average(numbers) {
      const values = numbers.filter((value) => Number.isFinite(value) && value > 0);
      if (!values.length) return 0;
      return values.reduce((sum, value) => sum + value, 0) / values.length;
    }

    function uniqueCount(rows, key) {
      return new Set(rows.map((row) => row[key]).filter(Boolean)).size;
    }

    function buildCounter(rows, key) {
      const counts = new Map();
      rows.forEach((row) => {
        const label = (row[key] || "").trim();
        if (!label) return;
        counts.set(label, (counts.get(label) || 0) + 1);
      });
      return Array.from(counts.entries())
        .map(([label, value]) => ({ label, value }))
        .sort((left, right) => right.value - left.value || left.label.localeCompare(right.label));
    }

    function buildValueCounter(rows, key, valueKey) {
      const counts = new Map();
      rows.forEach((row) => {
        const label = (row[key] || "").trim();
        const value = Number(row[valueKey] || 0);
        if (!label || !Number.isFinite(value)) return;
        counts.set(label, (counts.get(label) || 0) + value);
      });
      return Array.from(counts.entries())
        .map(([label, value]) => ({ label, value }))
        .sort((left, right) => right.value - left.value || left.label.localeCompare(right.label));
    }

    function buildTrend(rows, granularity) {
      const key = granularity === "week" ? "week" : granularity === "month" ? "month" : "date";
      const counts = new Map();
      rows.forEach((row) => {
        const label = row[key];
        if (!label) return;
        counts.set(label, (counts.get(label) || 0) + 1);
      });
      return Array.from(counts.entries())
        .map(([label, value]) => ({ label, value }))
        .sort((left, right) => left.label.localeCompare(right.label))
        .slice(-16);
    }

    function populateSelect(element, options, allLabel) {
      const current = element.value;
      const htmlOptions = ['<option value="">' + escapeHtml(allLabel) + "</option>"];
      options.forEach((option) => {
        htmlOptions.push('<option value="' + escapeHtml(option) + '">' + escapeHtml(option) + "</option>");
      });
      element.innerHTML = htmlOptions.join("");
      if (options.includes(current)) {
        element.value = current;
      }
    }

    function renderBarChart(target, items, color, formatter, options = {}) {
      if (!items.length) {
        target.innerHTML = '<div class="empty">No data available for the selected filters.</div>';
        return;
      }
      const maxItems = Number.isFinite(options.maxItems) ? options.maxItems : 8;
      const trimmed = maxItems > 0 ? items.slice(0, maxItems) : items.slice();
      const maxValue = Math.max(...trimmed.map((item) => item.value), 0);
      target.innerHTML =
        '<div class="bar-list">' +
        trimmed
          .map((item) => {
            const width = maxValue === 0 ? 0 : (item.value / maxValue) * 100;
            return (
              '<div class="bar-row">' +
              '<div class="bar-label" title="' + escapeHtml(item.label) + '">' + escapeHtml(item.label) + "</div>" +
              '<div class="bar-track"><div class="bar-fill" style="width:' + width.toFixed(1) + "%; background:" + color + ';"></div></div>' +
              '<div class="bar-value">' + escapeHtml(formatter(item.value)) + "</div>" +
              "</div>"
            );
          })
          .join("") +
        "</div>";
    }

    function renderTrendChart(target, items) {
      if (!items.length) {
        target.innerHTML = '<div class="empty">No trend data available for the selected filters.</div>';
        return;
      }

      const chartTheme = getChartTheme();

      const width = Math.max(760, items.length * 60);
      const height = 248;
      const paddingX = 42;
      const paddingTop = 20;
      const paddingBottom = 40;
      const chartWidth = width - paddingX * 2;
      const chartHeight = height - paddingTop - paddingBottom;
      const maxValue = Math.max(...items.map((item) => item.value), 1);

      const points = items.map((item, index) => {
        const x = items.length === 1 ? width / 2 : paddingX + (chartWidth * index) / (items.length - 1);
        const y = paddingTop + chartHeight - (item.value / maxValue) * chartHeight;
        return { ...item, x, y };
      });

      const path = points.map((point, index) => (index === 0 ? "M" : "L") + point.x.toFixed(2) + " " + point.y.toFixed(2)).join(" ");
      const areaPath = path + " L " + points[points.length - 1].x.toFixed(2) + " " + (paddingTop + chartHeight) + " L " + points[0].x.toFixed(2) + " " + (paddingTop + chartHeight) + " Z";

      const labels = points
        .map((point) => {
          return '<text x="' + point.x.toFixed(2) + '" y="' + (height - 10) + '" text-anchor="middle" font-size="10" fill="' + chartTheme.muted + '">' + escapeHtml(point.label) + "</text>";
        })
        .join("");

      const dots = points
        .map((point) => {
          return (
            '<g>' +
            '<circle cx="' + point.x.toFixed(2) + '" cy="' + point.y.toFixed(2) + '" r="4" fill="' + chartTheme.line + '"></circle>' +
            '<text x="' + point.x.toFixed(2) + '" y="' + (point.y - 9).toFixed(2) + '" text-anchor="middle" font-size="10" fill="' + chartTheme.pointLabel + '">' + escapeHtml(String(point.value)) + "</text>" +
            "</g>"
          );
        })
        .join("");

      const grid = Array.from({ length: 5 }, (_, index) => {
        const y = paddingTop + (chartHeight * index) / 4;
        return '<line x1="' + paddingX + '" y1="' + y.toFixed(2) + '" x2="' + (width - paddingX) + '" y2="' + y.toFixed(2) + '" stroke="' + chartTheme.grid + '" stroke-dasharray="4 5"></line>';
      }).join("");

      target.innerHTML =
        '<svg class="line-svg" viewBox="0 0 ' + width + " " + height + '" style="width:' + width + 'px" role="img" aria-label="Intervention trend">' +
        grid +
        '<path d="' + areaPath + '" fill="' + chartTheme.area + '"></path>' +
        '<path d="' + path + '" fill="none" stroke="' + chartTheme.line + '" stroke-width="3" stroke-linecap="round" stroke-linejoin="round"></path>' +
        dots +
        labels +
        "</svg>";
    }

    function wrapChartLabel(label, maxLineLength) {
      const words = String(label || "").split(/\s+/).filter(Boolean);
      if (!words.length) return [""];
      const lines = [];
      let current = words[0];
      for (let index = 1; index < words.length; index += 1) {
        const next = words[index];
        if ((current + " " + next).length <= maxLineLength) {
          current += " " + next;
        } else {
          lines.push(current);
          current = next;
        }
      }
      lines.push(current);
      return lines.slice(0, 3);
    }

    function niceMax(value) {
      if (value <= 0) return 1;
      const exponent = Math.floor(Math.log10(value));
      const fraction = value / Math.pow(10, exponent);
      let niceFraction = 1;
      if (fraction <= 1) niceFraction = 1;
      else if (fraction <= 2) niceFraction = 2;
      else if (fraction <= 5) niceFraction = 5;
      else niceFraction = 10;
      return niceFraction * Math.pow(10, exponent);
    }

    function renderMultiSeriesChart(target, items, seriesDefs, options = {}) {
      if (!items.length) {
        target.innerHTML = '<div class="empty">No data available for this block.</div>';
        return;
      }

      const chartTheme = getChartTheme();

      const primaryMax = Math.max(
        ...items.flatMap((item) =>
          seriesDefs
            .filter((series) => !series.isSecondary)
            .map((series) => Number(item[series.key] || 0))
        ),
        0
      );

      const rawMax = Math.max(
        ...items.flatMap((item) => seriesDefs.map((series) => Number(item[series.key] || 0))),
        0
      );

      const context = { primaryMax: primaryMax || rawMax || 1, rawMax: rawMax || 1 };
      const preparedItems = items.map((item) => ({
        ...item,
        series: seriesDefs.map((series) => {
          const rawValue = Number(item[series.key] || 0);
          const scaledValue = series.scale ? Number(series.scale(rawValue, context)) : rawValue;
          return {
            ...series,
            rawValue,
            scaledValue,
            formatted: series.format ? series.format(rawValue) : formatNumber(rawValue),
          };
        }),
      }));

      const maxValue = niceMax(
        Math.max(
          ...preparedItems.flatMap((item) => item.series.map((series) => series.scaledValue)),
          0
        )
      );

      const width = Math.max(options.minWidth || 620, items.length * Math.max(options.groupMinWidth || 150, seriesDefs.length * 44 + 70));
      const height = options.height || 360;
      const margin = { top: 28, right: 18, bottom: 88, left: 52 };
      const chartWidth = width - margin.left - margin.right;
      const chartHeight = height - margin.top - margin.bottom;
      const groupWidth = chartWidth / items.length;
      const clusterWidth = Math.min(groupWidth - 20, seriesDefs.length * 38 + (seriesDefs.length - 1) * 10);
      const barWidth = Math.max(16, Math.min(34, (clusterWidth - (seriesDefs.length - 1) * 10) / seriesDefs.length));
      const clusterOffset = (groupWidth - (barWidth * seriesDefs.length + (seriesDefs.length - 1) * 10)) / 2;
      const tickCount = 4;

      const grid = Array.from({ length: tickCount + 1 }, (_, index) => {
        const value = (maxValue / tickCount) * index;
        const y = margin.top + chartHeight - (value / maxValue) * chartHeight;
        return (
          '<g>' +
          '<line x1="' + margin.left + '" y1="' + y.toFixed(2) + '" x2="' + (width - margin.right) + '" y2="' + y.toFixed(2) + '" stroke="' + chartTheme.grid + '" stroke-dasharray="4 5"></line>' +
          '<text x="' + (margin.left - 10) + '" y="' + (y + 4).toFixed(2) + '" text-anchor="end" font-size="11" fill="' + chartTheme.valueLabel + '">' + escapeHtml(formatNumber(value)) + "</text>" +
          "</g>"
        );
      }).join("");

      const bars = preparedItems
        .map((item, itemIndex) => {
          const groupX = margin.left + itemIndex * groupWidth;
          const labelLines = wrapChartLabel(item.label, 18);
          const labelX = groupX + groupWidth / 2;
          const labelY = height - 38;

          const labelSvg = labelLines
            .map((line, lineIndex) => {
              const dy = lineIndex === 0 ? 0 : 14;
              return '<tspan x="' + labelX.toFixed(2) + '" dy="' + dy + '">' + escapeHtml(line) + "</tspan>";
            })
            .join("");

          const barsSvg = item.series
            .map((series, seriesIndex) => {
              const x = groupX + clusterOffset + seriesIndex * (barWidth + 10);
              const barHeight = maxValue === 0 ? 0 : (series.scaledValue / maxValue) * chartHeight;
              const y = margin.top + chartHeight - barHeight;
              const displayY = Math.max(margin.top + 12, y - 8);
              return (
                '<g>' +
                '<rect x="' + x.toFixed(2) + '" y="' + y.toFixed(2) + '" width="' + barWidth.toFixed(2) + '" height="' + Math.max(2, barHeight).toFixed(2) + '" rx="10" fill="' + series.color + '"></rect>' +
                '<text x="' + (x + barWidth / 2).toFixed(2) + '" y="' + displayY.toFixed(2) + '" text-anchor="middle" font-size="11" font-weight="700" fill="' + chartTheme.pointLabel + '">' + escapeHtml(series.formatted) + "</text>" +
                "</g>"
              );
            })
            .join("");

          return (
            '<g>' +
            barsSvg +
            '<text x="' + labelX.toFixed(2) + '" y="' + labelY + '" text-anchor="middle" font-size="12" font-weight="700" fill="' + chartTheme.text + '">' + labelSvg + "</text>" +
            "</g>"
          );
        })
        .join("");

      target.innerHTML =
        '<div class="column-chart-wrap">' +
        '<svg class="column-chart-svg" viewBox="0 0 ' + width + " " + height + '" role="img" aria-label="Grouped column chart">' +
        grid +
        '<line x1="' + margin.left + '" y1="' + (margin.top + chartHeight).toFixed(2) + '" x2="' + (width - margin.right) + '" y2="' + (margin.top + chartHeight).toFixed(2) + '" stroke="' + chartTheme.axis + '"></line>' +
        bars +
        "</svg>" +
        "</div>";
    }

    function renderTable(target, headers, rows) {
      if (!target) return;
      if (!rows.length) {
        target.innerHTML = '<div class="empty">No rows available for the selected filters.</div>';
        return;
      }

      const headerHtml = headers.map((header) => "<th>" + escapeHtml(header) + "</th>").join("");
      const bodyHtml = rows
        .map((row) => {
          return "<tr>" + row.map((cell) => "<td>" + escapeHtml(cell) + "</td>").join("") + "</tr>";
        })
        .join("");

      target.innerHTML = '<div class="table-wrap"><table><thead><tr>' + headerHtml + "</tr></thead><tbody>" + bodyHtml + "</tbody></table></div>";
    }

    function renderTableHtml(target, headers, rows) {
      if (!target) return;
      if (!rows.length) {
        target.innerHTML = '<div class="empty">No rows available for the selected filters.</div>';
        return;
      }

      const headerHtml = headers.map((header) => "<th>" + escapeHtml(header) + "</th>").join("");
      const bodyHtml = rows
        .map((row) => {
          return "<tr>" + row.map((cell) => "<td>" + cell + "</td>").join("") + "</tr>";
        })
        .join("");

      target.innerHTML = '<div class="table-wrap"><table><thead><tr>' + headerHtml + "</tr></thead><tbody>" + bodyHtml + "</tbody></table></div>";
    }

    function confidenceBadgeHtml(confidence) {
      const tone = String(confidence || "low").trim().toLowerCase();
      return '<span class="confidence-badge ' + escapeHtml(tone) + '">' + escapeHtml(confidence || "Low") + '</span>';
    }

    function flatTimeActionButtonHtml(type, value, label) {
      const title = type === "activity" ? getFlatTimeActivityTooltip(label) : "";
      return '<button type="button" class="table-action" data-flat-focus-' + escapeHtml(type) + '="' + escapeHtml(value) + '"' + (title ? ' title="' + escapeHtml(title) + '"' : "") + '>' + escapeHtml(label) + '</button>';
    }

    function getFlatTimeActivityTooltip(activityLabel) {
      const label = String(activityLabel || "").trim();
      if (!label) return "";

      const translations = FLAT_TIME_ACTIVITY_TRANSLATIONS || {};
      const wellSections = translations.wellSections || {};
      const operations = translations.operations || {};
      const activities = translations.activities || {};
      const generic = translations.generic || {};
      const parts = label.split("-").filter(Boolean);
      const lines = [];

      if (parts.length >= 3) {
        const sectionCode = parts[0];
        const operationCode = parts[1];
        const activityCode = parts[parts.length - 1];

        if (wellSections[sectionCode]) {
          lines.push("Section " + sectionCode + ': ' + wellSections[sectionCode]);
        }
        if (operations[operationCode]) {
          lines.push("Operation " + operationCode + ': ' + operations[operationCode]);
        }
        if (activities[activityCode]) {
          lines.push("Activity " + activityCode + ': ' + activities[activityCode]);
        } else if (generic[activityCode]) {
          lines.push("Activity " + activityCode + ': ' + generic[activityCode]);
        }
      }

      if (!lines.length && generic[label]) {
        lines.push(generic[label]);
      }

      return lines.join("\n");
    }

    function flatTimeActivityLabelHtml(label) {
      const tooltip = getFlatTimeActivityTooltip(label);
      return '<span' + (tooltip ? ' title="' + escapeHtml(tooltip) + '"' : "") + '>' + escapeHtml(label) + '</span>';
    }

    function flatTimeTrendHtml(excessHours) {
      const value = Number(excessHours || 0);
      const isSlow = value > 0.01;
      const tone = isSlow ? "slow" : "fast";
      const arrow = isSlow ? "▲" : "▼";
      const label = isSlow ? "slower" : "faster";
      const days = value / 24;
      return (
        '<span class="trend-indicator ' + tone + '">' +
        '<span>' + escapeHtml(formatHours(value) + " hr (" + formatDays(days) + " d)") + '</span>' +
        '<span class="arrow">' + arrow + '</span>' +
        '<span>' + label + '</span>' +
        '</span>'
      );
    }

    function isYesLike(value) {
      return ["yes", "y", "true", "confirmed"].includes(String(value || "").trim().toLowerCase());
    }

    function normalizeCategory(category) {
      const value = String(category || "").trim().toLowerCase();
      if (value === "stuck pipe") return "Stuck pipe";
      if (value === "optimization") return "Optimization";
      if (value === "operational compliance") return "Operational Compliance";
      if (value === "well control") return "Well Control";
      if (value === "reporting" || value === "reporting ") return "Reporting";
      return category || "Other";
    }

    function formatPercent(value) {
      return (Number(value || 0) * 100).toFixed(1) + "%";
    }

    function formatHoursWithDays(value) {
      const hours = Number(value || 0);
      return formatHours(hours) + " hr / " + formatDays(hours / 24) + " d";
    }

    function formatDateHuman(dateString) {
      if (!dateString) return "";
      return new Date(dateString + "T00:00:00").toLocaleDateString("en-GB", {
        day: "2-digit",
        month: "short",
        year: "numeric",
      });
    }

    function weekBounds(week) {
      const dates = dashboardData.interventions
        .filter((row) => row.week === week && row.date)
        .map((row) => row.date)
        .sort();
      return {
        start: dates[0] || "",
        end: dates[dates.length - 1] || "",
      };
    }

    function formatWeekLabel(week) {
      const bounds = weekBounds(week);
      if (!bounds.start || !bounds.end) return week || "Unknown week";
      return formatDateHuman(bounds.start) + " - " + formatDateHuman(bounds.end);
    }

    function toIsoDate(date) {
      const year = date.getFullYear();
      const month = String(date.getMonth() + 1).padStart(2, "0");
      const day = String(date.getDate()).padStart(2, "0");
      return year + "-" + month + "-" + day;
    }

    function getDefaultLastTuesdayRange(referenceDate) {
      const baseDate = referenceDate ? new Date(referenceDate + "T00:00:00") : new Date();
      if (Number.isNaN(baseDate.getTime())) {
        baseDate.setTime(Date.now());
      }
      baseDate.setHours(0, 0, 0, 0);
      const dayOfWeek = baseDate.getDay();
      const diffToTuesday = (dayOfWeek - 2 + 7) % 7;
      const end = new Date(baseDate);
      end.setDate(baseDate.getDate() - diffToTuesday);
      const start = new Date(end);
      start.setDate(end.getDate() - 6);
      return {
        start: toIsoDate(start),
        end: toIsoDate(end),
      };
    }

    function getWeeklyReportDateRange() {
      let start = ui.weeklyReportStartDate.value;
      let end = ui.weeklyReportEndDate.value;

      if (start && end && start > end) {
        return { start: end, end: start };
      }
      return { start, end };
    }

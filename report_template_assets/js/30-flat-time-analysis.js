    function buildFlatTimeGroupItems(datasets, totalKey) {
      const groupNames = Array.from(
        new Set(datasets.flatMap((dataset) => dataset.groups.map((group) => group.groupName)))
      );

      return groupNames
        .map((groupName) => {
          const item = { label: groupName };
          datasets.forEach((dataset) => {
            const match = dataset.groups.find((group) => group.groupName === groupName);
            item[dataset.id] = match ? Number(match[totalKey] || 0) : 0;
          });
          item.total = datasets.reduce((sum, dataset) => sum + Number(item[dataset.id] || 0), 0);
          return item;
        })
        .sort((left, right) => right.total - left.total || left.label.localeCompare(right.label));
    }

    function buildFlatTimeActivityItems(datasets, metricKey) {
      const activityMap = new Map();
      datasets.forEach((dataset) => {
        dataset.groups.forEach((group) => {
          group.activities.forEach((activity) => {
            const key = activity.activity;
            if (!activityMap.has(key)) {
              activityMap.set(key, {
                label: key,
                groupLabel: group.groupName,
                sectionSize: activity.sectionSize || extractFlatTimeSectionSize(activity.activity),
              });
            }
            activityMap.get(key)[dataset.id] = Number(activity[metricKey] || 0);
          });
        });
      });

      return Array.from(activityMap.values())
        .map((item) => {
          datasets.forEach((dataset) => {
            item[dataset.id] = Number(item[dataset.id] || 0);
          });
          item.total = datasets.reduce((sum, dataset) => sum + Number(item[dataset.id] || 0), 0);
          return item;
        })
        .sort((left, right) => right.total - left.total || left.label.localeCompare(right.label));
    }

    function percentile(numbers, ratio) {
      const values = numbers.filter((value) => Number.isFinite(value)).slice().sort((left, right) => left - right);
      if (!values.length) return 0;
      if (values.length === 1) return values[0];
      const position = (values.length - 1) * ratio;
      const lower = Math.floor(position);
      const upper = Math.ceil(position);
      if (lower === upper) return values[lower];
      const weight = position - lower;
      return values[lower] * (1 - weight) + values[upper] * weight;
    }

    function standardDeviation(numbers) {
      const values = numbers.filter((value) => Number.isFinite(value));
      if (values.length <= 1) return 0;
      const mean = values.reduce((sum, value) => sum + value, 0) / values.length;
      const variance = values.reduce((sum, value) => sum + Math.pow(value - mean, 2), 0) / values.length;
      return Math.sqrt(variance);
    }

    function flatTimeConfidence(sampleSize, cv) {
      if (sampleSize >= 5 && cv <= 0.2) return "High";
      if (sampleSize >= 3 && cv <= 0.35) return "Medium";
      return "Low";
    }

    function flatTimeVariabilityLabel(cv) {
      if (cv <= 0.2) return "Low";
      if (cv <= 0.35) return "Moderate";
      return "High";
    }

    function normalizeFlatTimeFocusActivities(opportunities, fallbackActivity) {
      const validActivities = new Set(opportunities.map((opportunity) => opportunity.activityLabel));
      const currentValues = Array.isArray(flatTimeState.focusActivities) && flatTimeState.focusActivities.length
        ? flatTimeState.focusActivities
        : flatTimeState.focusActivity
          ? [flatTimeState.focusActivity]
          : [];
      const sanitized = Array.from(
        new Set(
          currentValues
            .filter((value) => typeof value === "string" && value.trim())
            .map((value) => value.trim())
            .filter((value) => validActivities.has(value))
        )
      );
      if (!sanitized.length && fallbackActivity && validActivities.has(fallbackActivity)) {
        sanitized.push(fallbackActivity);
      }
      flatTimeState.focusActivities = sanitized;
      flatTimeState.focusActivity = sanitized[0] || "";
      return sanitized;
    }

    function buildSelectedActivityAggregate(opportunities, selectedActivityLabels) {
      const selectedLabels = Array.from(
        new Set((selectedActivityLabels || []).filter((value) => typeof value === "string" && value.trim()))
      );
      const selectedItems = opportunities.filter((opportunity) => selectedLabels.includes(opportunity.activityLabel));
      if (!selectedItems.length) return null;

      const peerMap = new Map();
      selectedItems.forEach((opportunity) => {
        opportunity.ranked.forEach((entry) => {
          if (!peerMap.has(entry.datasetId)) {
            peerMap.set(entry.datasetId, {
              datasetId: entry.datasetId,
              label: entry.label,
              rigLabel: entry.rigLabel || "Rig not mapped",
              value: 0,
            });
          }
          peerMap.get(entry.datasetId).value += Number(entry.value || 0);
        });
      });

      const ranked = Array.from(peerMap.values())
        .filter((entry) => entry.value > 0)
        .sort((left, right) => right.value - left.value || left.label.localeCompare(right.label));
      const values = ranked.map((entry) => entry.value);
      const occurrenceCount = values.length;
      const topEntry = ranked[0] || { label: "N/A", rigLabel: "Rig not mapped", value: 0 };
      const meanValue = occurrenceCount ? values.reduce((sum, value) => sum + value, 0) / occurrenceCount : 0;
      const peerValues = ranked.slice(1).map((entry) => entry.value).filter((value) => value > 0);
      const peerAverage = peerValues.length ? average(peerValues) : 0;
      const fastestTime = occurrenceCount ? Math.min(...values) : 0;
      const p25Value = percentile(values, 0.25);
      const sortedValues = values.slice().sort((left, right) => left - right);
      const medianValue = occurrenceCount
        ? (occurrenceCount % 2
            ? sortedValues[(occurrenceCount - 1) / 2]
            : (sortedValues[occurrenceCount / 2 - 1] + sortedValues[occurrenceCount / 2]) / 2)
        : 0;
      const stdDev = standardDeviation(values);
      const cv = meanValue > 0 ? stdDev / meanValue : 0;
      const idealTime = selectedItems.reduce((sum, item) => sum + Number(item.idealTime || 0), 0);
      const gapToIdeal = Math.max(topEntry.value - idealTime, 0);
      const totalRecoverableHours = ranked.reduce((sum, entry) => sum + Math.max(entry.value - idealTime, 0), 0);
      const confidence = flatTimeConfidence(occurrenceCount, cv);
      const variability = flatTimeVariabilityLabel(cv);
      const sectionLabels = Array.from(new Set(selectedItems.map((item) => formatFlatTimeSectionSize(item.sectionSize))));
      const groupLabels = Array.from(new Set(selectedItems.map((item) => item.groupLabel || "Unknown")));

      return {
        activityLabels: selectedLabels,
        activityLabel: selectedLabels.join(" + "),
        labelCount: selectedLabels.length,
        groupLabel: groupLabels.join(" • "),
        sectionLabel: sectionLabels.join(" • "),
        idealTime,
        idealRule: selectedItems.length === 1 ? selectedItems[0].idealRule : (selectedItems.length + " activity ideals summed"),
        confidence,
        variability,
        occurrenceCount,
        fastestTime,
        meanValue,
        medianValue,
        p25Value,
        stdDev,
        cv,
        topEntry,
        peerAverage,
        gapToIdeal,
        totalRecoverableHours,
        values,
        ranked,
        summaryText:
          occurrenceCount >= 2
            ? selectedItems.length + " activities across " + occurrenceCount + " wells; peers avg " + formatNumber(peerAverage || meanValue || 0) + " hr, top well " + topEntry.label + " ran " + formatNumber(topEntry.value) + " hr"
            : "Only one well observed for the selected activity set",
      };
    }

    function computeFlatTimeOpportunity(item, datasets) {
      const ranked = datasets
        .map((dataset) => ({
          datasetId: dataset.id,
          label: dataset.subjectWell,
          rigLabel: dataset.rigLabel || "Rig not mapped",
          value: Number(item[dataset.id] || 0),
        }))
        .filter((entry) => entry.value > 0)
        .sort((left, right) => right.value - left.value);

      const values = ranked.map((entry) => entry.value);
      const occurrenceCount = values.length;
      const topEntry = ranked[0] || { label: "N/A", rigLabel: "Rig not mapped", value: 0 };
      const meanValue = occurrenceCount ? values.reduce((sum, value) => sum + value, 0) / occurrenceCount : 0;
      const peerValues = ranked.slice(1).map((entry) => entry.value).filter((value) => value > 0);
      const peerAverage = peerValues.length ? average(peerValues) : 0;
      const fastestTime = occurrenceCount ? Math.min(...values) : 0;
      const p25Value = percentile(values, 0.25);
      const sortedValues = values.slice().sort((left, right) => left - right);
      const medianValue = occurrenceCount
        ? (occurrenceCount % 2
            ? sortedValues[(occurrenceCount - 1) / 2]
            : (sortedValues[occurrenceCount / 2 - 1] + sortedValues[occurrenceCount / 2]) / 2)
        : 0;
      const stdDev = standardDeviation(values);
      const cv = meanValue > 0 ? stdDev / meanValue : 0;

      let idealTime = fastestTime;
      let idealRule = "fastest";

      if (occurrenceCount >= 3 && fastestTime > 0) {
        const meanGapRatio = meanValue > 0 ? Math.abs(meanValue - fastestTime) / meanValue : 0;
        const medianGapRatio = medianValue > 0 ? Math.abs(medianValue - fastestTime) / medianValue : 0;
        if (meanGapRatio > 0.35 || medianGapRatio > 0.35) {
          idealTime = Math.min(
            ...[p25Value, medianValue, meanValue].filter((value) => Number.isFinite(value) && value > 0)
          );
          idealRule = "stable benchmark";
        }
      }

      if (!Number.isFinite(idealTime) || idealTime <= 0) {
        idealTime = fastestTime || p25Value || meanValue || medianValue || 0;
      }

      const gapToIdeal = Math.max(topEntry.value - idealTime, 0);
      const gapVsPeerAverage = peerAverage > 0 ? Math.max(topEntry.value - peerAverage, 0) : 0;
      const totalRecoverableHours = ranked.reduce((sum, entry) => sum + Math.max(entry.value - idealTime, 0), 0);
      const confidence = flatTimeConfidence(occurrenceCount, cv);
      const variability = flatTimeVariabilityLabel(cv);

      return {
        sectionSize: item.sectionSize || "__no_section__",
        groupLabel: item.groupLabel || "Unknown",
        activityLabel: item.label,
        totalTime: item.total,
        averagePerWell: occurrenceCount ? item.total / occurrenceCount : 0,
        peerAverage,
        meanValue,
        p25Value,
        medianValue,
        stdDev,
        cv,
        fastestTime,
        idealTime,
        idealRule,
        gapToIdeal,
        gapVsPeerAverage,
        totalRecoverableHours,
        confidence,
        variability,
        topEntry,
        occurrenceCount,
        values,
        ranked,
        summaryText:
          occurrenceCount >= 2
            ? occurrenceCount + " wells; peers avg " + formatNumber(peerAverage || meanValue || 0) + " hr, top well " + topEntry.label + " ran " + formatNumber(topEntry.value) + " hr"
            : "Only one well observed for this activity",
      };
    }

    function buildWellRanking(datasets, opportunities) {
      return datasets
        .map((dataset) => {
          let actualTotal = 0;
          let idealTotal = 0;
          let excessTotal = 0;
          const drivers = [];

          opportunities.forEach((opportunity) => {
            const actual = Number(opportunity.ranked.find((entry) => entry.datasetId === dataset.id)?.value || 0);
            if (!actual) return;
            actualTotal += actual;
            idealTotal += opportunity.idealTime;
            const gap = Math.max(actual - opportunity.idealTime, 0);
            excessTotal += gap;
            if (gap > 0) {
              drivers.push({
                activity: opportunity.activityLabel,
                group: opportunity.groupLabel,
                gap,
              });
            }
          });

          drivers.sort((left, right) => right.gap - left.gap || left.activity.localeCompare(right.activity));
          const topDrivers = drivers.slice(0, 3);
          const otherDriversGap = Math.max(
            drivers.slice(3).reduce((sum, driver) => sum + driver.gap, 0),
            0
          );

          return {
            rigLabel: dataset.rigLabel || "Rig not mapped",
            wellLabel: dataset.subjectWell,
            actualTotal,
            idealTotal,
            excessTotal,
            topDriver: drivers[0] || null,
            topDrivers,
            otherDriversGap,
          };
        })
        .sort((left, right) => right.excessTotal - left.excessTotal || right.actualTotal - left.actualTotal || left.wellLabel.localeCompare(right.wellLabel));
    }

    function buildSectionBenchmarkItems(datasets, metricKey, opportunities) {
      const sectionMap = new Map();
      const opportunityMap = new Map(
        opportunities.map((opportunity) => [opportunity.activityLabel, opportunity])
      );

      datasets.forEach((dataset) => {
        const totalsBySection = new Map();
        dataset.groups.forEach((group) => {
          group.activities.forEach((activity) => {
            const sectionSize = activity.sectionSize || extractFlatTimeSectionSize(activity.activity);
            if (!sectionSize || sectionSize === "__no_section__") return;
            const actual = Number(activity[metricKey] || 0);
            if (actual <= 0) return;
            const opportunity = opportunityMap.get(activity.activity);
            const ideal = Number(opportunity?.idealTime || 0);
            if (!totalsBySection.has(sectionSize)) {
              totalsBySection.set(sectionSize, { actual: 0, ideal: 0 });
            }
            const bucket = totalsBySection.get(sectionSize);
            bucket.actual += actual;
            bucket.ideal += ideal;
          });
        });

        totalsBySection.forEach((totals, sectionSize) => {
          if (!sectionMap.has(sectionSize)) {
            sectionMap.set(sectionSize, { actualValues: [], idealValues: [] });
          }
          const bucket = sectionMap.get(sectionSize);
          bucket.actualValues.push(totals.actual);
          bucket.idealValues.push(totals.ideal);
        });
      });

      return Array.from(sectionMap.entries())
        .map(([sectionSize, bucket]) => {
          const actualAverage = average(bucket.actualValues || []);
          const idealTime = average(bucket.idealValues || []);
          return {
            label: formatFlatTimeSectionSize(sectionSize),
            sectionSize,
            actualAverage,
            idealTime,
            spread: Math.max(actualAverage - idealTime, 0),
          };
        })
        .sort((left, right) => Number(right.sectionSize) - Number(left.sectionSize) || left.label.localeCompare(right.label));
    }

    function buildRigBenchmarkSummary(datasets, opportunities) {
      const rigMap = new Map();

      datasets.forEach((dataset) => {
        const rigLabel = dataset.rigLabel || "Rig not mapped";
        if (!rigMap.has(rigLabel)) {
          rigMap.set(rigLabel, {
            rigLabel,
            datasets: [],
            gapByActivity: new Map(),
          });
        }
        rigMap.get(rigLabel).datasets.push(dataset);
      });

      return Array.from(rigMap.values())
        .map((bucket) => {
          let actualTotal = 0;
          let idealTotal = 0;
          let excessTotal = 0;
          opportunities.forEach((opportunity) => {
            bucket.datasets.forEach((dataset) => {
              const actual = Number(opportunity.ranked.find((entry) => entry.datasetId === dataset.id)?.value || 0);
              if (!actual) return;
              actualTotal += actual;
              idealTotal += opportunity.idealTime;
              const gap = Math.max(actual - opportunity.idealTime, 0);
              excessTotal += gap;
              if (gap > 0) {
                bucket.gapByActivity.set(
                  opportunity.activityLabel,
                  (bucket.gapByActivity.get(opportunity.activityLabel) || 0) + gap
                );
              }
            });
          });

          const mainRepeatingActivity = Array.from(bucket.gapByActivity.entries())
            .sort((left, right) => right[1] - left[1] || left[0].localeCompare(right[0]))[0];

          return {
            rigLabel: bucket.rigLabel,
            wellCount: bucket.datasets.length,
            averageFlatTime: bucket.datasets.length ? actualTotal / bucket.datasets.length : 0,
            averageIdealTime: bucket.datasets.length ? idealTotal / bucket.datasets.length : 0,
            excessTime: excessTotal,
            mainRepeatingActivity: mainRepeatingActivity ? mainRepeatingActivity[0] : "No repeated excess",
          };
        })
        .sort((left, right) => right.excessTime - left.excessTime || right.averageFlatTime - left.averageFlatTime || left.rigLabel.localeCompare(right.rigLabel));
    }

    function buildOpportunityPipeline(opportunities) {
      return opportunities
        .map((opportunity) => {
          const wellsImpacted = opportunity.ranked.filter((entry) => entry.value > opportunity.idealTime).length;
          let priority = "Monitor";
          if (opportunity.totalRecoverableHours >= 40 && opportunity.confidence === "High") priority = "Act now";
          else if (opportunity.totalRecoverableHours >= 20 || opportunity.confidence === "Medium") priority = "Next wave";

          return {
            activityLabel: opportunity.activityLabel,
            groupLabel: opportunity.groupLabel,
            occurrenceCount: opportunity.occurrenceCount,
            wellsImpacted,
            idealTime: opportunity.idealTime,
            totalRecoverableHours: opportunity.totalRecoverableHours,
            priority,
          };
        })
        .sort((left, right) => right.totalRecoverableHours - left.totalRecoverableHours || right.wellsImpacted - left.wellsImpacted || left.activityLabel.localeCompare(right.activityLabel));
    }

    function buildSectionBreakdownMap(datasets, opportunities, metricKey) {
      const opportunityMap = new Map(opportunities.map((opportunity) => [opportunity.activityLabel, opportunity]));
      const breakdownMap = new Map();

      datasets.forEach((dataset) => {
        const sectionMap = new Map();
        dataset.groups.forEach((group) => {
          group.activities.forEach((activity) => {
            const sectionSize = activity.sectionSize || extractFlatTimeSectionSize(activity.activity);
            if (!sectionSize || sectionSize === "__no_section__") return;
            const actual = Number(activity[metricKey] || 0);
            if (actual <= 0) return;

            const opportunity = opportunityMap.get(activity.activity);
            const ideal = Number(opportunity?.idealTime || 0);
            const gap = Math.max(actual - ideal, 0);

            if (!sectionMap.has(sectionSize)) {
              sectionMap.set(sectionSize, {
                sectionSize,
                label: formatFlatTimeSectionSize(sectionSize),
                actualTotal: 0,
                idealTotal: 0,
                gapTotal: 0,
                activities: [],
              });
            }

            const bucket = sectionMap.get(sectionSize);
            bucket.actualTotal += actual;
            bucket.idealTotal += ideal;
            bucket.gapTotal += gap;
            bucket.activities.push({
              activityLabel: activity.activity,
              groupLabel: group.groupName,
              actual,
              idealTime: ideal,
              gap,
              confidence: opportunity?.confidence || "Low",
            });
          });
        });

        const sections = Array.from(sectionMap.values())
          .map((section) => {
            const topActivities = section.activities
              .slice()
              .sort((left, right) => right.actual - left.actual || right.gap - left.gap || left.activityLabel.localeCompare(right.activityLabel))
              .slice(0, 3);
            const topGapActivities = section.activities
              .slice()
              .sort((left, right) => right.gap - left.gap || right.actual - left.actual || left.activityLabel.localeCompare(right.activityLabel))
              .slice(0, 3);
            return {
              ...section,
              topActivities,
              topGapActivities,
            };
          })
          .sort((left, right) => Number(right.sectionSize) - Number(left.sectionSize) || left.label.localeCompare(right.label));

        breakdownMap.set(dataset.id, sections);
      });

      return breakdownMap;
    }

    function sectionRecommendation(activityLabel, groupLabel) {
      const token = String(activityLabel || "").toUpperCase();
      const group = String(groupLabel || "").toUpperCase();
      if (token.includes("LOG")) return "Tighten logging sequence and remove waiting between runs.";
      if (token.includes("WBCO") || token.includes("CIRC")) return "Reduce waiting and optimize circulation / conditioning steps.";
      if (token.includes("CSG") || token.includes("CEM")) return "Improve casing and cement readiness before execution.";
      if (token.includes("BOP")) return "Standardize BOP test and handling sequence with pre-job readiness.";
      if (group.includes("RM") || group.includes("TRIP")) return "Replicate the best tripping sequence and reduce connection / handling delays.";
      return "Replicate the best observed procedure and remove non-productive waiting in this activity.";
    }

    function buildAllWellsSectionSummaryRows(datasets, breakdownMap) {
      return datasets
        .flatMap((dataset) =>
          (breakdownMap.get(dataset.id) || []).map((section) => ({
            rigLabel: dataset.rigLabel || "Rig not mapped",
            wellLabel: dataset.subjectWell,
            sectionSize: section.sectionSize,
            sectionLabel: section.label,
            actualTotal: section.actualTotal,
            idealTotal: section.idealTotal,
            gapTotal: section.gapTotal,
            topActivities: section.topActivities,
          }))
        )
        .sort((left, right) =>
          right.gapTotal - left.gapTotal ||
          right.actualTotal - left.actualTotal ||
          left.rigLabel.localeCompare(right.rigLabel) ||
          left.wellLabel.localeCompare(right.wellLabel)
        );
    }

    function buildWellSectionGroups(datasets, breakdownMap, selectedWell) {
      return datasets
        .map((dataset, index) => {
          const sections = (breakdownMap.get(dataset.id) || []).slice();
          const actualTotal = sections.reduce((sum, section) => sum + Number(section.actualTotal || 0), 0);
          const idealTotal = sections.reduce((sum, section) => sum + Number(section.idealTotal || 0), 0);
          const gapTotal = sections.reduce((sum, section) => sum + Number(section.gapTotal || 0), 0);
          return {
            dataset,
            rigLabel: dataset.rigLabel || "Rig not mapped",
            wellLabel: dataset.subjectWell,
            sections,
            actualTotal,
            idealTotal,
            gapTotal,
            isSelected: dataset.subjectWell === selectedWell,
            orderIndex: index,
          };
        })
        .sort((left, right) =>
          Number(right.isSelected) - Number(left.isSelected) ||
          left.orderIndex - right.orderIndex ||
          left.wellLabel.localeCompare(right.wellLabel)
        );
    }

    function renderSectionDriverChips(activities, valueKey) {
      if (!activities || !activities.length) {
        return '<span class="section-driver-chip is-muted">No recurring activity</span>';
      }
      return activities
        .map((activity) => (
          '<span class="section-driver-chip">' +
          flatTimeActivityLabelHtml(activity.activityLabel) +
          '<span class="section-driver-chip-value">' + escapeHtml(formatNumber(Number(activity[valueKey] || 0)) + " hr") + "</span>" +
          "</span>"
        ))
        .join("");
    }

    function renderGroupedWellSectionSummary(target, wellSectionGroups) {
      if (!wellSectionGroups.length) {
        target.innerHTML = '<div class="empty">No well-section summary available for this section size.</div>';
        return;
      }

      target.innerHTML =
        '<div class="well-section-browser">' +
        wellSectionGroups
          .map((group) => (
            '<details class="well-section-accordion"' + (group.isSelected ? " open" : "") + ">" +
            '<summary>' +
            '<div class="well-section-accordion-title">' +
            '<strong>' + escapeHtml(group.wellLabel) + '</strong>' +
            '<span>' + escapeHtml(group.rigLabel) + '</span>' +
            "</div>" +
            '<div class="well-section-accordion-metrics">' +
            '<span>' + escapeHtml(formatNumber(group.actualTotal)) + ' hr actual</span>' +
            '<span>' + escapeHtml(formatNumber(group.idealTotal)) + ' hr ideal</span>' +
            '<span>' + escapeHtml(formatNumber(group.gapTotal)) + ' hr gap</span>' +
            "</div>" +
            '<span class="check-dropdown-caret">▼</span>' +
            "</summary>" +
            '<div class="well-section-accordion-body">' +
            '<div class="table-wrap">' +
            "<table><thead><tr><th>Section</th><th>Actual (hr)</th><th>Ideal (hr)</th><th>Gap (hr)</th><th>Top 3 Time Consumers</th></tr></thead><tbody>" +
            group.sections
              .map((section) => (
                "<tr>" +
                "<td>" + escapeHtml(section.label) + "</td>" +
                "<td>" + escapeHtml(formatNumber(section.actualTotal)) + "</td>" +
                "<td>" + escapeHtml(formatNumber(section.idealTotal)) + "</td>" +
                "<td>" + escapeHtml(formatNumber(section.gapTotal)) + "</td>" +
                '<td><div class="section-driver-chip-list">' + renderSectionDriverChips(section.topActivities, "actual") + "</div></td>" +
                "</tr>"
              ))
              .join("") +
            "</tbody></table></div>" +
            "</div>" +
            "</details>"
          ))
          .join("") +
        "</div>";
    }

    function renderMatrixGroupedWellSectionSummary(target, wellSectionGroups) {
      if (!wellSectionGroups.length) {
        target.innerHTML = '<div class="empty">No well-section summary available for this section size.</div>';
        return;
      }

      const sectionSizes = Array.from(
        new Set(
          wellSectionGroups.flatMap((group) => group.sections.map((section) => section.sectionSize))
        )
      ).sort((left, right) => Number(right) - Number(left) || String(left).localeCompare(String(right)));

      const maxGap = Math.max(
        ...wellSectionGroups.flatMap((group) => group.sections.map((section) => Number(section.gapTotal || 0))),
        0
      );

      const matrixHeader =
        "<tr><th>Well</th>" +
        sectionSizes.map((sectionSize) => "<th>" + escapeHtml(formatFlatTimeSectionSize(sectionSize)) + "</th>").join("") +
        "</tr>";
      const matrixBody = wellSectionGroups
        .map((group) => {
          const cells = sectionSizes
            .map((sectionSize) => {
              const section = group.sections.find((item) => item.sectionSize === sectionSize);
              if (!section) {
                return '<td><div class="section-matrix-cell is-empty"><strong>-</strong><span>No section</span></div></td>';
              }
              const gap = Number(section.gapTotal || 0);
              const opacity = maxGap > 0 ? Math.max(0.10, gap / maxGap) : 0;
              const bg = gap > 0 ? "rgba(200, 30, 90, " + opacity.toFixed(2) + ")" : "rgba(18, 100, 214, 0.06)";
              const fg = gap > 0.01 ? "#ffffff" : "var(--ink)";
              return (
                '<td>' +
                '<div class="section-matrix-cell" style="background:' + bg + "; color:" + fg + ';">' +
                "<strong>" + escapeHtml(formatNumber(section.actualTotal)) + "</strong>" +
                '<span>actual</span>' +
                "<small>+" + escapeHtml(formatNumber(section.gapTotal)) + " hr gap</small>" +
                "</div>" +
                "</td>"
              );
            })
            .join("");
          return (
            "<tr>" +
            "<td><strong>" + escapeHtml(group.wellLabel) + '</strong><br><span style="color:var(--muted); font-size:12px;">' + escapeHtml(group.rigLabel) + "</span></td>" +
            cells +
            "</tr>"
          );
        })
        .join("");

      target.innerHTML =
        '<div class="stack-gap">' +
        '<div class="table-wrap"><table class="section-matrix-table"><thead>' + matrixHeader + "</thead><tbody>" + matrixBody + "</tbody></table></div>" +
        '<div class="report-note">Matrix shows <strong>actual hours</strong> in each section cell and the section <strong>gap</strong> below it. The grouped detail remains below so you keep the full activity context.</div>' +
        '<div id="flat-time-all-wells-grouped-detail"></div>' +
        "</div>";

      const groupedTarget = target.querySelector("#flat-time-all-wells-grouped-detail");
      renderGroupedWellSectionSummary(groupedTarget, wellSectionGroups);
    }

    function renderAllWellsSectionSummary(target, allWellSectionRows, datasets, breakdownMap, selectedWell) {
      const layout = ui.flatTimeAllWellsLayout ? ui.flatTimeAllWellsLayout.value || "ranked" : "ranked";
      const wellSectionGroups = buildWellSectionGroups(datasets, breakdownMap, selectedWell);

      if (layout === "grouped") {
        renderGroupedWellSectionSummary(target, wellSectionGroups);
        return;
      }

      if (layout === "matrix") {
        renderMatrixGroupedWellSectionSummary(target, wellSectionGroups);
        return;
      }

      renderTableHtml(
        target,
        ["Rig", "Well", "Section", "Actual (hr)", "Ideal (hr)", "Gap (hr)", "Top 3 Time Consumers"],
        allWellSectionRows.map((row) => [
          escapeHtml(row.rigLabel),
          flatTimeActionButtonHtml("well", row.wellLabel, row.wellLabel),
          escapeHtml(row.sectionLabel),
          escapeHtml(formatNumber(row.actualTotal)),
          escapeHtml(formatNumber(row.idealTotal)),
          escapeHtml(formatNumber(row.gapTotal)),
          row.topActivities.map((activity) => flatTimeActivityLabelHtml(activity.activityLabel) + escapeHtml(" (" + formatNumber(activity.actual) + " hr)")).join("<br>"),
        ])
      );
    }

    function buildSelectedWellSectionItems(datasets, selectedDataset, breakdownMap) {
      if (!selectedDataset) return [];
      const selectedSections = breakdownMap.get(selectedDataset.id) || [];
      return selectedSections.map((section) => {
        const peerSections = datasets
          .filter((dataset) => dataset.id !== selectedDataset.id)
          .map((dataset) => {
            const peerSection = (breakdownMap.get(dataset.id) || []).find((item) => item.sectionSize === section.sectionSize);
            if (!peerSection || peerSection.actualTotal <= 0) return null;
            return {
              dataset,
              actualTotal: peerSection.actualTotal,
            };
          })
          .filter(Boolean);

        const peerActuals = peerSections.map((item) => item.actualTotal).filter((value) => value > 0);
        const bestPeer = peerSections
          .slice()
          .sort((left, right) => left.actualTotal - right.actualTotal || left.dataset.subjectWell.localeCompare(right.dataset.subjectWell))[0] || null;

        const offsetAverage = peerActuals.length ? average(peerActuals) : 0;
        const bestOffset = peerActuals.length ? Math.min(...peerActuals) : 0;

        return {
          ...section,
          offsetAverage,
          bestOffset,
          bestOffsetWell: bestPeer ? bestPeer.dataset.subjectWell : "No peer",
          bestOffsetRig: bestPeer ? (bestPeer.dataset.rigLabel || "Rig not mapped") : "No peer",
          peerCount: peerActuals.length,
          gapVsOffsetAverage: offsetAverage > 0 ? section.actualTotal - offsetAverage : 0,
        };
      });
    }

    function buildSectionSavingsSummary(datasets, breakdownMap) {
      const sectionMap = new Map();

      datasets.forEach((dataset) => {
        (breakdownMap.get(dataset.id) || []).forEach((section) => {
          if (!sectionMap.has(section.sectionSize)) {
            sectionMap.set(section.sectionSize, {
              sectionSize: section.sectionSize,
              sectionLabel: section.label,
              recoverableHours: 0,
              actualTotal: 0,
              idealTotal: 0,
              impactedWells: 0,
              activityGaps: new Map(),
              confidenceCounts: { High: 0, Medium: 0, Low: 0 },
            });
          }

          const bucket = sectionMap.get(section.sectionSize);
          bucket.recoverableHours += section.gapTotal;
          bucket.actualTotal += section.actualTotal;
          bucket.idealTotal += section.idealTotal;
          if (section.gapTotal > 0) bucket.impactedWells += 1;
          section.topGapActivities.forEach((activity) => {
            bucket.activityGaps.set(
              activity.activityLabel,
              (bucket.activityGaps.get(activity.activityLabel) || 0) + activity.gap
            );
            bucket.confidenceCounts[activity.confidence] = (bucket.confidenceCounts[activity.confidence] || 0) + 1;
          });
        });
      });

      return Array.from(sectionMap.values())
        .map((bucket) => {
          const topActivity = Array.from(bucket.activityGaps.entries())
            .sort((left, right) => right[1] - left[1] || left[0].localeCompare(right[0]))[0];
          const confidence = bucket.confidenceCounts.High > 0 ? "High" : bucket.confidenceCounts.Medium > 0 ? "Medium" : "Low";
          return {
            sectionSize: bucket.sectionSize,
            sectionLabel: bucket.sectionLabel,
            recoverableHours: bucket.recoverableHours,
            recoverableDays: bucket.recoverableHours / 24,
            actualAverage: datasets.length ? bucket.actualTotal / datasets.length : 0,
            idealAverage: datasets.length ? bucket.idealTotal / datasets.length : 0,
            impactedWells: bucket.impactedWells,
            topActivity: topActivity ? topActivity[0] : "No recurring excess",
            confidence,
            recommendation: sectionRecommendation(topActivity ? topActivity[0] : "", ""),
          };
        })
        .sort((left, right) => right.recoverableHours - left.recoverableHours || Number(right.sectionSize) - Number(left.sectionSize));
    }

    function buildExecutiveNarrativeHtml(selectedDataset, sectionItems, savingsRows, wellRanking) {
      if (!selectedDataset || !sectionItems.length) {
        return '<div class="empty">Select a well and load at least one comparable offset to generate the executive narrative.</div>';
      }

      const worstSection = sectionItems
        .slice()
        .sort((left, right) => right.gapTotal - left.gapTotal || right.actualTotal - left.actualTotal)[0];
      const strongestSavings = savingsRows[0];
      const rankingRow = wellRanking.find((row) => row.wellLabel === selectedDataset.subjectWell);
      const topActivities = (worstSection?.topGapActivities || [])
        .slice(0, 3)
        .map((item) => item.activityLabel + " (" + formatNumber(item.gap) + " hr)")
        .join(", ");

      const paragraphs = [
        '<p><strong>' + escapeHtml(selectedDataset.subjectWell) + '</strong> is currently running with <strong>' + escapeHtml(formatNumber(rankingRow?.excessTotal || 0)) + ' hr</strong> above the recommended flat time benchmark. The section creating the largest burden is <strong>' + escapeHtml(worstSection?.label || "N/A") + '</strong>, where the well is spending <strong>' + escapeHtml(formatNumber(worstSection?.actualTotal || 0)) + ' hr</strong> against an ideal target of <strong>' + escapeHtml(formatNumber(worstSection?.idealTotal || 0)) + ' hr</strong>.</p>',
        '<p>The system selected this section because it combines the largest recoverable gap with repeatable activities seen across the loaded offsets. In this section, the main drivers are <strong>' + escapeHtml(topActivities || "no recurring drivers identified") + '</strong>. This means the opportunity is not just the single slowest event, but a recurring pattern where the selected well is running slower than the best validated performance.</p>',
        strongestSavings
          ? '<p>The most actionable recovery area across the current benchmark set is <strong>' + escapeHtml(strongestSavings.sectionLabel) + '</strong>, with about <strong>' + escapeHtml(formatNumber(strongestSavings.recoverableHours)) + ' hr</strong> (' + escapeHtml(formatNumber(strongestSavings.recoverableDays)) + ' d) available to recover. The recommended action is to attack <strong>' + flatTimeActivityLabelHtml(strongestSavings.topActivity) + '</strong> first because it is the strongest repeated source of excess time and already has a <strong>' + confidenceBadgeHtml(strongestSavings.confidence) + '</strong> benchmark behind it. To recover time, replicate the best observed sequence, remove waiting between dependent steps, and standardize the execution around the loaded offsets that already achieved the lower time.</p>'
          : '<p>No strong recurring section-level savings case is available yet. Load more offsets to increase benchmark confidence and reveal stable section opportunities.</p>',
      ];

      return '<div class="drill-list">' + paragraphs.map((paragraph) => '<div class="drill-item">' + paragraph + '</div>').join('') + '</div>';
    }

    function renderExecutiveSelectedWellSections(target, sectionItems, selectedDataset) {
      if (!selectedDataset || !sectionItems.length) {
        target.innerHTML = '<div class="empty">Select a well and load section-sized activities to compare the selected well by section.</div>';
        return;
      }

      target.innerHTML =
        '<div class="legend" style="margin-bottom:10px;">' +
        '<span class="legend-item"><span class="legend-dot" style="background:#1264d6;"></span>' + escapeHtml(selectedDataset.subjectWell + " actual") + '</span>' +
        '<span class="legend-item"><span class="legend-dot" style="background:#0f766e;"></span>Synthetic ideal</span>' +
        '<span class="legend-item"><span class="legend-dot" style="background:#c06a0a;"></span>Recoverable gap</span>' +
        '</div>' +
        '<p class="report-note" style="margin-bottom:14px;">Synthetic ideal is an activity-by-activity benchmark assembled from the loaded offsets. It is not a single real offset well, so it can be lower than the best real well in a section.</p>' +
        '<div id="flat-time-selected-well-sections-chart"></div><div id="flat-time-selected-well-sections-table" style="margin-top:16px;"></div>';
      const chartTarget = target.querySelector("#flat-time-selected-well-sections-chart");
      const tableTarget = target.querySelector("#flat-time-selected-well-sections-table");

      renderMultiSeriesChart(
        chartTarget,
        sectionItems.map((item) => ({
          label: item.label,
          actual: item.actualTotal,
          ideal: item.idealTotal,
          gap: item.gapTotal,
        })),
        [
          { key: "actual", label: selectedDataset.subjectWell + " actual", color: "#1264d6", format: (value) => formatNumber(value) },
          { key: "ideal", label: "Recommended ideal", color: "#0f766e", format: (value) => formatNumber(value) },
          { key: "gap", label: "Recoverable gap", color: "#c06a0a", format: (value) => formatNumber(value) },
        ],
        { height: 420, minWidth: 760, groupMinWidth: 150 }
      );

      renderTableHtml(
        tableTarget,
        ["Section", "Actual (hr)", "Synthetic Ideal (hr)", "Gap (hr)", "Top 3 Time Consumers"],
        sectionItems.map((item) => [
          escapeHtml(item.label),
          escapeHtml(formatNumber(item.actualTotal)),
          escapeHtml(formatNumber(item.idealTotal)),
          escapeHtml(formatNumber(item.gapTotal)),
          item.topActivities.map((activity) => flatTimeActivityLabelHtml(activity.activityLabel) + escapeHtml(" (" + formatNumber(activity.actual) + " hr)")).join("<br>"),
        ])
      );
    }

    function renderExecutiveOffsetComparison(target, sectionItems, selectedDataset) {
      if (!selectedDataset || !sectionItems.length) {
        target.innerHTML = '<div class="empty">Need one selected well and at least one comparable offset to build the section comparison.</div>';
        return;
      }

      target.innerHTML =
        '<div class="legend" style="margin-bottom:10px;">' +
        '<span class="legend-item"><span class="legend-dot" style="background:#1264d6;"></span>' + escapeHtml(selectedDataset.subjectWell + " selected") + '</span>' +
        '<span class="legend-item"><span class="legend-dot" style="background:#0f766e;"></span>Offset average</span>' +
        '<span class="legend-item"><span class="legend-dot" style="background:#7c3aed;"></span>Best real offset</span>' +
        '<span class="legend-item"><span class="legend-dot" style="background:#c06a0a;"></span>Synthetic ideal</span>' +
        '</div>' +
        '<p class="report-note" style="margin-bottom:14px;">Best real offset is the lowest real well total observed in that section. Synthetic ideal is built from activity-level validated targets and can be lower than the best real offset.</p>' +
        '<div id="flat-time-offset-comparison-chart"></div><div id="flat-time-offset-comparison-table" style="margin-top:16px;"></div>';
      const chartTarget = target.querySelector("#flat-time-offset-comparison-chart");
      const tableTarget = target.querySelector("#flat-time-offset-comparison-table");

      renderMultiSeriesChart(
        chartTarget,
        sectionItems.map((item) => ({
          label: item.label,
          selected: item.actualTotal,
          offsetAverage: item.offsetAverage,
          bestOffset: item.bestOffset,
          ideal: item.idealTotal,
        })),
        [
          { key: "selected", label: selectedDataset.subjectWell, color: "#1264d6", format: (value) => formatNumber(value) },
          { key: "offsetAverage", label: "Offset avg", color: "#0f766e", format: (value) => formatNumber(value) },
          { key: "bestOffset", label: "Best offset", color: "#7c3aed", format: (value) => formatNumber(value) },
          { key: "ideal", label: "Recommended ideal", color: "#c06a0a", format: (value) => formatNumber(value) },
        ],
        { height: 420, minWidth: 820, groupMinWidth: 160 }
      );

      renderTable(
        tableTarget,
        ["Section", selectedDataset.subjectWell + " (hr)", "Offset Avg (hr)", "Best Offset (hr)", "Best Offset Well", "Synthetic Ideal (hr)", "Gap vs Offset Avg (hr)"],
        sectionItems.map((item) => [
          item.label,
          formatNumber(item.actualTotal),
          formatNumber(item.offsetAverage),
          formatNumber(item.bestOffset),
          item.bestOffsetWell === "No peer" ? item.bestOffsetWell : item.bestOffsetWell + " (" + item.bestOffsetRig + ")",
          formatNumber(item.idealTotal),
          formatNumber(item.gapVsOffsetAverage),
        ])
      );
    }

    function renderSectionHeatmap(target, datasets, breakdownMap) {
      if (!datasets.length) {
        target.innerHTML = '<div class="empty">No section heatmap data available.</div>';
        return;
      }

      const sectionSizes = Array.from(
        new Set(
          datasets.flatMap((dataset) => (breakdownMap.get(dataset.id) || []).map((item) => item.sectionSize))
        )
      ).sort((left, right) => Number(right) - Number(left) || String(left).localeCompare(String(right)));

      if (!sectionSizes.length) {
        target.innerHTML = '<div class="empty">No section-sized activities available to build the section heatmap.</div>';
        return;
      }

      const maxGap = Math.max(
        ...datasets.flatMap((dataset) =>
          sectionSizes.map((sectionSize) => {
            const item = (breakdownMap.get(dataset.id) || []).find((entry) => entry.sectionSize === sectionSize);
            return Number(item?.gapTotal || 0);
          })
        ),
        0
      );

      const headerHtml = '<tr><th>Well</th>' + sectionSizes.map((sectionSize) => '<th>' + escapeHtml(formatFlatTimeSectionSize(sectionSize)) + '</th>').join('') + '</tr>';
      const bodyHtml = datasets.map((dataset) => {
        const cells = sectionSizes.map((sectionSize) => {
          const item = (breakdownMap.get(dataset.id) || []).find((entry) => entry.sectionSize === sectionSize);
          const gap = Number(item?.gapTotal || 0);
          const opacity = maxGap > 0 ? Math.max(0.12, gap / maxGap) : 0;
          const bg = gap > 0 ? 'rgba(200, 30, 90, ' + opacity.toFixed(2) + ')' : 'rgba(18, 100, 214, 0.06)';
          const color = gap > 0.01 ? '#ffffff' : 'var(--ink)';
          return '<td style="background:' + bg + '; color:' + color + '; font-weight:700; text-align:center;">' + escapeHtml(formatNumber(gap)) + '</td>';
        }).join('');
        return '<tr><td><strong>' + escapeHtml(dataset.subjectWell) + '</strong><br><span style="color:var(--muted); font-size:12px;">' + escapeHtml(dataset.rigLabel || "") + '</span></td>' + cells + '</tr>';
      }).join('');

      target.innerHTML = '<div class="table-wrap"><table><thead>' + headerHtml + '</thead><tbody>' + bodyHtml + '</tbody></table></div>';
    }

    function renderWaterfallChart(target, dataset, opportunities) {
      if (!dataset || !opportunities.length) {
        target.innerHTML = '<div class="empty">Choose a well to draw the waterfall.</div>';
        return;
      }

      const chartTheme = getChartTheme();
      const contributions = opportunities
        .map((opportunity) => {
          const actual = Number(opportunity.ranked.find((entry) => entry.datasetId === dataset.id)?.value || 0);
          return {
            label: opportunity.activityLabel,
            gap: Math.max(actual - opportunity.idealTime, 0),
          };
        })
        .filter((item) => item.gap > 0)
        .sort((left, right) => right.gap - left.gap || left.label.localeCompare(right.label));

      const topDrivers = contributions.slice(0, 5);
      const otherGap = contributions.slice(5).reduce((sum, item) => sum + item.gap, 0);
      if (otherGap > 0) topDrivers.push({ label: "Other gaps", gap: otherGap });

      const actualTotal = topDrivers.reduce((sum, item) => sum + item.gap, 0) + opportunities.reduce((sum, opportunity) => {
        const actual = Number(opportunity.ranked.find((entry) => entry.datasetId === dataset.id)?.value || 0);
        return sum + Math.min(actual, opportunity.idealTime || 0);
      }, 0);
      const idealTotal = Math.max(actualTotal - topDrivers.reduce((sum, item) => sum + item.gap, 0), 0);

      const steps = [{ label: "Actual total", start: 0, end: actualTotal, type: "total" }];
      let running = actualTotal;
      topDrivers.forEach((driver) => {
        steps.push({ label: driver.label, start: running, end: running - driver.gap, type: "delta", delta: -driver.gap });
        running -= driver.gap;
      });
      steps.push({ label: "Ideal total", start: 0, end: idealTotal, type: "total-ideal" });

      const maxValue = niceMax(Math.max(...steps.map((step) => Math.max(step.start, step.end)), 1));
      const width = Math.max(780, steps.length * 120);
      const height = 360;
      const margin = { top: 24, right: 24, bottom: 86, left: 56 };
      const chartWidth = width - margin.left - margin.right;
      const chartHeight = height - margin.top - margin.bottom;
      const barWidth = Math.min(78, chartWidth / steps.length - 18);

      const grid = Array.from({ length: 5 }, (_, index) => {
        const value = (maxValue / 4) * index;
        const y = margin.top + chartHeight - (value / maxValue) * chartHeight;
        return (
          '<g>' +
          '<line x1="' + margin.left + '" y1="' + y.toFixed(2) + '" x2="' + (width - margin.right) + '" y2="' + y.toFixed(2) + '" stroke="' + chartTheme.grid + '" stroke-dasharray="4 5"></line>' +
          '<text x="' + (margin.left - 10) + '" y="' + (y + 4).toFixed(2) + '" text-anchor="end" font-size="11" fill="' + chartTheme.valueLabel + '">' + escapeHtml(formatNumber(value)) + '</text>' +
          '</g>'
        );
      }).join("");

      const bars = steps.map((step, index) => {
        const x = margin.left + index * (chartWidth / steps.length) + ((chartWidth / steps.length) - barWidth) / 2;
        const topValue = Math.max(step.start, step.end);
        const bottomValue = Math.min(step.start, step.end);
        const y = margin.top + chartHeight - (topValue / maxValue) * chartHeight;
        const yBottom = margin.top + chartHeight - (bottomValue / maxValue) * chartHeight;
        const heightValue = Math.max(yBottom - y, 2);
        const fill = step.type === "delta" ? "#be123c" : step.type === "total-ideal" ? "#0f766e" : "#1264d6";
        const centerX = x + barWidth / 2;
        const labelLines = wrapChartLabel(step.label, 16);
        const labelSvg = labelLines.map((line, lineIndex) => '<tspan x="' + centerX.toFixed(2) + '" dy="' + (lineIndex === 0 ? 0 : 13) + '">' + escapeHtml(line) + '</tspan>').join("");
        const displayValue = step.type === "delta" ? formatNumber(Math.abs(step.delta || 0)) : formatNumber(step.end);
        return (
          '<g>' +
          '<rect x="' + x.toFixed(2) + '" y="' + y.toFixed(2) + '" width="' + barWidth.toFixed(2) + '" height="' + heightValue.toFixed(2) + '" rx="10" fill="' + fill + '"></rect>' +
          '<text x="' + centerX.toFixed(2) + '" y="' + Math.max(margin.top + 12, y - 8).toFixed(2) + '" text-anchor="middle" font-size="11" font-weight="700" fill="' + chartTheme.pointLabel + '">' + escapeHtml(displayValue) + '</text>' +
          '<text x="' + centerX.toFixed(2) + '" y="' + (height - 42) + '" text-anchor="middle" font-size="11" font-weight="700" fill="' + chartTheme.text + '">' + labelSvg + '</text>' +
          '</g>'
        );
      }).join("");

      target.innerHTML =
        '<div class="legend" style="margin-bottom:10px;">' +
        '<span class="legend-item"><span class="legend-dot" style="background:#1264d6;"></span>' + escapeHtml(dataset.subjectWell + " actual") + '</span>' +
        '<span class="legend-item"><span class="legend-dot" style="background:#be123c;"></span>Recoverable gap</span>' +
        '<span class="legend-item"><span class="legend-dot" style="background:#0f766e;"></span>Recommended ideal</span>' +
        '</div>' +
        '<div class="column-chart-wrap">' +
        '<svg class="column-chart-svg" viewBox="0 0 ' + width + ' ' + height + '" role="img" aria-label="Well versus ideal waterfall chart">' +
        grid +
        bars +
        '<line x1="' + margin.left + '" y1="' + (margin.top + chartHeight).toFixed(2) + '" x2="' + (width - margin.right) + '" y2="' + (margin.top + chartHeight).toFixed(2) + '" stroke="' + chartTheme.axis + '"></line>' +
        '</svg></div>';
    }

    function renderVariabilityChart(target, opportunities, topN) {
      const items = opportunities
        .filter((opportunity) => opportunity.occurrenceCount >= 2)
        .slice()
        .sort((left, right) => right.cv - left.cv || right.totalTime - left.totalTime)
        .slice(0, topN);

      if (!items.length) {
        target.innerHTML = '<div class="empty">Need at least two wells per activity to show variability.</div>';
        return;
      }

      const chartTheme = getChartTheme();
      const width = 920;
      const height = Math.max(340, items.length * 48 + 70);
      const margin = { top: 26, right: 28, bottom: 34, left: 190 };
      const chartWidth = width - margin.left - margin.right;
      const chartHeight = height - margin.top - margin.bottom;
      const maxValue = niceMax(Math.max(...items.map((item) => Math.max(...item.values, item.idealTime)), 1));
      const rowHeight = chartHeight / items.length;

      const grid = Array.from({ length: 6 }, (_, index) => {
        const value = (maxValue / 5) * index;
        const x = margin.left + (value / maxValue) * chartWidth;
        return (
          '<g>' +
          '<line x1="' + x.toFixed(2) + '" y1="' + margin.top + '" x2="' + x.toFixed(2) + '" y2="' + (height - margin.bottom) + '" stroke="' + chartTheme.grid + '" stroke-dasharray="4 5"></line>' +
          '<text x="' + x.toFixed(2) + '" y="' + (height - 10) + '" text-anchor="middle" font-size="11" fill="' + chartTheme.valueLabel + '">' + escapeHtml(formatNumber(value)) + '</text>' +
          '</g>'
        );
      }).join("");

      const rows = items.map((item, index) => {
        const y = margin.top + index * rowHeight + rowHeight / 2;
        const minValue = Math.min(...item.values);
        const maxObserved = Math.max(...item.values);
        const q1 = percentile(item.values, 0.25);
        const q3 = percentile(item.values, 0.75);
        const xMin = margin.left + (minValue / maxValue) * chartWidth;
        const xQ1 = margin.left + (q1 / maxValue) * chartWidth;
        const xMedian = margin.left + (item.medianValue / maxValue) * chartWidth;
        const xQ3 = margin.left + (q3 / maxValue) * chartWidth;
        const xMax = margin.left + (maxObserved / maxValue) * chartWidth;
        const xIdeal = margin.left + (item.idealTime / maxValue) * chartWidth;
        return (
          '<g>' +
          '<line x1="' + xMin.toFixed(2) + '" y1="' + y.toFixed(2) + '" x2="' + xMax.toFixed(2) + '" y2="' + y.toFixed(2) + '" stroke="' + chartTheme.axis + '" stroke-width="2"></line>' +
          '<rect x="' + xQ1.toFixed(2) + '" y="' + (y - 10).toFixed(2) + '" width="' + Math.max(xQ3 - xQ1, 2).toFixed(2) + '" height="20" rx="8" fill="rgba(18, 100, 214, 0.25)" stroke="#1264d6"></rect>' +
          '<line x1="' + xMedian.toFixed(2) + '" y1="' + (y - 12).toFixed(2) + '" x2="' + xMedian.toFixed(2) + '" y2="' + (y + 12).toFixed(2) + '" stroke="#1264d6" stroke-width="3"></line>' +
          '<circle cx="' + xIdeal.toFixed(2) + '" cy="' + y.toFixed(2) + '" r="5" fill="#c06a0a"></circle>' +
          '<text x="' + (margin.left - 12) + '" y="' + (y + 4).toFixed(2) + '" text-anchor="end" font-size="11" font-weight="700" fill="' + chartTheme.text + '">' + escapeHtml(item.activityLabel) + '</text>' +
          '</g>'
        );
      }).join("");

      target.innerHTML =
        '<div class="legend" style="margin-bottom:10px;">' +
        '<span class="legend-item"><span class="legend-dot" style="background:#1264d6;"></span>Interquartile range / median</span>' +
        '<span class="legend-item"><span class="legend-dot" style="background:#c06a0a;"></span>Recommended ideal</span>' +
        '</div>' +
        '<div class="column-chart-wrap">' +
        '<svg class="column-chart-svg" viewBox="0 0 ' + width + ' ' + height + '" role="img" aria-label="Variability box plot chart">' +
        grid +
        rows +
        '</svg></div>';
    }

    function wireFlatTimeFocusActions() {
      Array.from(document.querySelectorAll("[data-flat-focus-well]")).forEach((button) => {
        button.addEventListener("click", () => {
          flatTimeState.focusWell = button.dataset.flatFocusWell || "";
          renderFlatTime();
        });
      });

      Array.from(document.querySelectorAll("[data-flat-focus-activity]")).forEach((button) => {
        button.addEventListener("click", () => {
          flatTimeState.focusActivities = button.dataset.flatFocusActivity ? [button.dataset.flatFocusActivity] : [];
          flatTimeState.focusActivity = flatTimeState.focusActivities[0] || "";
          renderFlatTime();
        });
      });
    }

    function renderFlatTimeDrilldown(datasets, opportunities, selectedWell, selectedActivities, sectionOptionDatasets) {
      const selectedDataset = datasets.find((dataset) => dataset.subjectWell === selectedWell) || datasets[0];
      const selectedOpportunity = buildSelectedActivityAggregate(opportunities, selectedActivities);

      if (!selectedDataset || !selectedOpportunity) {
        ui.flatTimeWellDrilldown.innerHTML = '<div class="empty">Select more CSVs to unlock the drill-down.</div>';
        ui.flatTimeActivityDrilldown.innerHTML = '<div class="empty">Select more CSVs to unlock the activity benchmark.</div>';
        return;
      }

      const allWellDrivers = opportunities
        .map((opportunity) => {
          const actual = Number(opportunity.ranked.find((entry) => entry.datasetId === selectedDataset.id)?.value || 0);
          const gap = Math.max(actual - opportunity.idealTime, 0);
          return {
            activityLabel: opportunity.activityLabel,
            groupLabel: opportunity.groupLabel,
            actual,
            idealTime: opportunity.idealTime,
            gap,
            peerAverage: opportunity.peerAverage || opportunity.meanValue || 0,
          };
        })
        .filter((item) => item.actual > 0)
        .sort((left, right) => right.gap - left.gap || right.actual - left.actual || left.activityLabel.localeCompare(right.activityLabel));

      const wellDrivers = allWellDrivers.slice(0, 6);
      const wellOptions = datasets
        .slice()
        .sort((left, right) => left.subjectWell.localeCompare(right.subjectWell))
        .map((dataset) => (
          '<option value="' + escapeHtml(dataset.subjectWell) + '"' +
          (dataset.subjectWell === selectedDataset.subjectWell ? " selected" : "") +
          '>' + escapeHtml(dataset.subjectWell + (dataset.rigLabel ? " • " + dataset.rigLabel : "")) + "</option>"
        ))
        .join("");
      const sectionOptions = ['<option value="">All section sizes</option>']
        .concat(
          getAvailableFlatTimeSectionSizes(sectionOptionDatasets || datasets).map((sectionSize) => (
            '<option value="' + escapeHtml(sectionSize) + '"' +
            (sectionSize === (ui.flatTimeSection.value || "") ? " selected" : "") +
            '>' + escapeHtml(formatFlatTimeSectionSize(sectionSize)) + "</option>"
          ))
        )
        .join("");

      const wellActualTotal = allWellDrivers.reduce((sum, item) => sum + item.actual, 0);
      const wellIdealTotal = allWellDrivers.reduce((sum, item) => sum + item.idealTime, 0);
      const wellExcess = allWellDrivers.reduce((sum, item) => sum + item.gap, 0);

      ui.flatTimeWellDrilldown.innerHTML =
        '<div class="drill-grid" style="margin-bottom:14px;">' +
        '<div class="field">' +
        '<label for="flat-time-drilldown-well-select">Selected Well</label>' +
        '<select id="flat-time-drilldown-well-select">' + wellOptions + '</select>' +
        '</div>' +
        '<div class="field">' +
        '<label for="flat-time-drilldown-section-select">Selected Section</label>' +
        '<select id="flat-time-drilldown-section-select">' + sectionOptions + '</select>' +
        '</div>' +
        '</div>' +
        '<div class="metric-strip" style="margin-bottom:14px;">' +
        '<div class="metric-pill"><div class="label">Selected Well</div><div class="value"><span class="value-main">' + escapeHtml(selectedDataset.subjectWell) + '</span></div><div class="meta">' + escapeHtml(selectedDataset.rigLabel || "Rig not mapped") + '</div></div>' +
        '<div class="metric-pill"><div class="label">Actual Flat Time</div><div class="value"><span class="value-main">' + escapeHtml(formatNumber(wellActualTotal)) + '</span><span class="value-suffix">' + escapeHtml("hr / " + formatNumber(wellActualTotal / 24) + " d") + '</span></div><div class="meta">All activities in the selected filter context</div></div>' +
        '<div class="metric-pill"><div class="label">Ideal Flat Time</div><div class="value"><span class="value-main">' + escapeHtml(formatNumber(wellIdealTotal)) + '</span><span class="value-suffix">' + escapeHtml("hr / " + formatNumber(wellIdealTotal / 24) + " d") + '</span></div><div class="meta">Recommended achievable total for the same activities</div></div>' +
        '<div class="metric-pill"><div class="label">Recoverable Gap</div><div class="value"><span class="value-main">' + escapeHtml(formatNumber(wellExcess)) + '</span><span class="value-suffix">' + escapeHtml("hr / " + formatNumber(wellExcess / 24) + " d") + '</span></div><div class="meta">Time above the recommended ideal</div></div>' +
        '</div>' +
        '<div class="drill-list">' +
        wellDrivers.map((item) =>
          '<div class="drill-item">' +
          '<strong>' + flatTimeActionButtonHtml("activity", item.activityLabel, item.activityLabel) + '</strong>' +
          '<span>' + escapeHtml(item.groupLabel + " • actual " + formatHoursWithDays(item.actual) + " • peers avg " + formatHoursWithDays(item.peerAverage) + " • ideal " + formatHoursWithDays(item.idealTime) + " • gap " + formatHoursWithDays(item.gap)) + '</span>' +
          '</div>'
        ).join('') +
        '</div>';

      const drilldownWellSelect = document.getElementById("flat-time-drilldown-well-select");
      if (drilldownWellSelect) {
        drilldownWellSelect.addEventListener("change", () => {
          flatTimeState.focusWell = drilldownWellSelect.value || "";
          ui.flatTimeWell.value = flatTimeState.focusWell;
          renderFlatTime();
        });
      }

      const drilldownSectionSelect = document.getElementById("flat-time-drilldown-section-select");
      if (drilldownSectionSelect) {
        drilldownSectionSelect.addEventListener("change", () => {
          ui.flatTimeSection.value = drilldownSectionSelect.value || "";
          renderFlatTime();
        });
      }

      const peerRows = selectedOpportunity.ranked
        .slice()
        .sort((left, right) => right.value - left.value || left.label.localeCompare(right.label))
        .map((entry) => ({
          rigLabel: entry.rigLabel || "Rig not mapped",
          label: entry.label,
          value: entry.value,
          gap: Math.max(entry.value - selectedOpportunity.idealTime, 0),
        }));

      const activityOptions = opportunities
        .slice()
        .sort((left, right) => left.activityLabel.localeCompare(right.activityLabel))
        .map((opportunity) => (
          '<label class="check-option">' +
          '<input type="checkbox" class="flat-time-activity-check" value="' + escapeHtml(opportunity.activityLabel) + '"' +
          (selectedOpportunity.activityLabels.includes(opportunity.activityLabel) ? " checked" : "") +
          '>' +
          '<span>' + escapeHtml(opportunity.activityLabel) + "</span>" +
          "</label>"
        ))
        .join("");
      const selectedActivitiesLabel = selectedOpportunity.activityLabels.length
        ? (selectedOpportunity.activityLabels.length <= 2
            ? selectedOpportunity.activityLabels.join(", ")
            : (selectedOpportunity.activityLabels.length + " activities selected"))
        : "Choose activities";

      ui.flatTimeActivityDrilldown.innerHTML =
        '<div class="field" style="margin-bottom:14px; max-width:420px;">' +
        '<label for="flat-time-drilldown-activity-picker">Selected Activities</label>' +
        '<details class="check-dropdown" id="flat-time-drilldown-activity-picker">' +
        '<summary><span class="check-dropdown-label">' + escapeHtml(selectedActivitiesLabel) + '</span><span class="check-dropdown-caret">▼</span></summary>' +
        '<div class="check-dropdown-menu">' + activityOptions + '</div>' +
        '</details>' +
        '<div class="field-help">Tick the activities you want to combine. The benchmark below sums the selected activities.</div>' +
        '</div>' +
        '<div class="metric-strip" style="margin-bottom:14px;">' +
        '<div class="metric-pill"><div class="label">Selected Activities</div><div class="value"><span class="value-main">' + escapeHtml(String(selectedOpportunity.labelCount)) + '</span><span class="value-suffix">' + escapeHtml(selectedOpportunity.labelCount === 1 ? "activity" : "activities") + '</span></div><div class="meta">' + selectedOpportunity.activityLabels.map((label) => flatTimeActivityLabelHtml(label)).join("<br>") + '</div></div>' +
        '<div class="metric-pill"><div class="label">Recommended Ideal</div><div class="value"><span class="value-main">' + escapeHtml(formatNumber(selectedOpportunity.idealTime)) + '</span><span class="value-suffix">' + escapeHtml("hr / " + formatNumber(selectedOpportunity.idealTime / 24) + " d") + '</span></div><div class="meta">' + escapeHtml(selectedOpportunity.idealRule) + '</div></div>' +
        '<div class="metric-pill"><div class="label">Confidence</div><div class="value">' + confidenceBadgeHtml(selectedOpportunity.confidence) + '</div><div class="meta">' + escapeHtml(selectedOpportunity.variability + " variability • sample " + selectedOpportunity.occurrenceCount) + '</div></div>' +
        '<div class="metric-pill"><div class="label">Observed Range</div><div class="value"><span class="value-main">' + escapeHtml(formatNumber(selectedOpportunity.fastestTime)) + " - " + escapeHtml(formatNumber(Math.max(...selectedOpportunity.values, 0))) + '</span><span class="value-suffix">' + escapeHtml("hr / " + formatNumber(selectedOpportunity.fastestTime / 24) + " - " + formatNumber(Math.max(...selectedOpportunity.values, 0) / 24) + " d") + '</span></div><div class="meta">Fastest to slowest observed execution</div></div>' +
        '</div>' +
        '<div class="table-wrap"><table><thead><tr><th>Rig</th><th>Well</th><th>Observed Time (hr)</th><th>Gap vs Ideal (hr)</th></tr></thead><tbody>' +
        peerRows.map((row) =>
          '<tr>' +
          '<td>' + escapeHtml(row.rigLabel) + '</td>' +
          '<td>' + flatTimeActionButtonHtml("well", row.label, row.label) + '</td>' +
          '<td>' + escapeHtml(formatHoursWithDays(row.value)) + '</td>' +
          '<td>' + escapeHtml(formatHoursWithDays(row.gap)) + '</td>' +
          '</tr>'
        ).join('') +
        '</tbody></table></div>';

      Array.from(document.querySelectorAll(".flat-time-activity-check")).forEach((checkbox) => {
        checkbox.addEventListener("change", () => {
          flatTimeState.focusActivities = Array.from(document.querySelectorAll(".flat-time-activity-check:checked"))
            .map((input) => input.value)
            .filter(Boolean);
          if (!flatTimeState.focusActivities.length && checkbox.value) {
            flatTimeState.focusActivities = [checkbox.value];
          }
          flatTimeState.focusActivity = flatTimeState.focusActivities[0] || "";
          renderFlatTime();
        });
      });

      ui.flatTimeDrilldownNote.textContent =
        'Selected well: ' + selectedDataset.subjectWell + ' • selected activities: ' + selectedOpportunity.activityLabels.join(", ") + '. ' +
        'Headline well totals are now calculated from all activities in the selected filter context, while the list below keeps only the top loss drivers. ' +
        'Depth-based drill-down is still limited because the uploaded CSVs do not contain true depth fields such as section top/bottom, measured depth or TD.';
    }

    function renderParetoChart(target, opportunities) {
      if (!opportunities.length) {
        target.innerHTML = '<div class="empty">No recoverable-hour opportunities available.</div>';
        return;
      }

      const items = opportunities
        .slice()
        .sort((left, right) => right.totalRecoverableHours - left.totalRecoverableHours)
        .slice(0, 10);
      const totalRecoverable = items.reduce((sum, item) => sum + item.totalRecoverableHours, 0) || 1;
      const chartTheme = getChartTheme();
      const width = Math.max(820, items.length * 120);
      const height = 360;
      const margin = { top: 24, right: 40, bottom: 86, left: 52 };
      const chartWidth = width - margin.left - margin.right;
      const chartHeight = height - margin.top - margin.bottom;
      const maxBar = niceMax(Math.max(...items.map((item) => item.totalRecoverableHours), 1));
      const groupWidth = chartWidth / items.length;
      let cumulative = 0;
      const points = [];

      const grid = Array.from({ length: 5 }, (_, index) => {
        const value = (maxBar / 4) * index;
        const y = margin.top + chartHeight - (value / maxBar) * chartHeight;
        return (
          '<g>' +
          '<line x1="' + margin.left + '" y1="' + y.toFixed(2) + '" x2="' + (width - margin.right) + '" y2="' + y.toFixed(2) + '" stroke="' + chartTheme.grid + '" stroke-dasharray="4 5"></line>' +
          '<text x="' + (margin.left - 10) + '" y="' + (y + 4).toFixed(2) + '" text-anchor="end" font-size="11" fill="' + chartTheme.valueLabel + '">' + escapeHtml(formatNumber(value)) + "</text>" +
          "</g>"
        );
      }).join("");

      const bars = items.map((item, index) => {
        const x = margin.left + index * groupWidth + groupWidth * 0.18;
        const barWidth = groupWidth * 0.64;
        const barHeight = (item.totalRecoverableHours / maxBar) * chartHeight;
        const y = margin.top + chartHeight - barHeight;
        cumulative += item.totalRecoverableHours;
        const cumulativePct = (cumulative / totalRecoverable) * 100;
        const pointX = x + barWidth / 2;
        const pointY = margin.top + chartHeight - (cumulativePct / 100) * chartHeight;
        points.push({ x: pointX, y: pointY, pct: cumulativePct });
        const labelLines = wrapChartLabel(item.activityLabel, 16);
        const labelSvg = labelLines
          .map((line, lineIndex) => '<tspan x="' + pointX.toFixed(2) + '" dy="' + (lineIndex === 0 ? 0 : 13) + '">' + escapeHtml(line) + "</tspan>")
          .join("");
        return (
          '<g>' +
          '<rect x="' + x.toFixed(2) + '" y="' + y.toFixed(2) + '" width="' + barWidth.toFixed(2) + '" height="' + Math.max(barHeight, 2).toFixed(2) + '" rx="10" fill="#1264d6"></rect>' +
          '<text x="' + pointX.toFixed(2) + '" y="' + Math.max(margin.top + 12, y - 8).toFixed(2) + '" text-anchor="middle" font-size="11" font-weight="700" fill="' + chartTheme.pointLabel + '">' + escapeHtml(formatNumber(item.totalRecoverableHours)) + "</text>" +
          '<text x="' + pointX.toFixed(2) + '" y="' + (height - 42) + '" text-anchor="middle" font-size="11" font-weight="700" fill="' + chartTheme.text + '">' + labelSvg + "</text>" +
          "</g>"
        );
      }).join("");

      const linePath = points.map((point, index) => (index === 0 ? "M" : "L") + point.x.toFixed(2) + " " + point.y.toFixed(2)).join(" ");
      const lineDots = points
        .map((point) => (
          '<g>' +
          '<circle cx="' + point.x.toFixed(2) + '" cy="' + point.y.toFixed(2) + '" r="4" fill="#c06a0a"></circle>' +
          '<text x="' + point.x.toFixed(2) + '" y="' + (point.y - 10).toFixed(2) + '" text-anchor="middle" font-size="10" fill="' + chartTheme.pointLabel + '">' + escapeHtml(formatNumber(point.pct) + "%") + "</text>" +
          '</g>'
        ))
        .join("");

      target.innerHTML =
        '<div class="column-chart-wrap">' +
        '<svg class="column-chart-svg" viewBox="0 0 ' + width + ' ' + height + '" role="img" aria-label="Pareto recoverable hours chart">' +
        grid +
        bars +
        '<path d="' + linePath + '" fill="none" stroke="#c06a0a" stroke-width="3.5" stroke-linecap="round" stroke-linejoin="round"></path>' +
        lineDots +
        '<line x1="' + margin.left + '" y1="' + (margin.top + chartHeight).toFixed(2) + '" x2="' + (width - margin.right) + '" y2="' + (margin.top + chartHeight).toFixed(2) + '" stroke="' + chartTheme.axis + '"></line>' +
        '</svg></div>';
    }

    function renderHeatmap(target, datasets, opportunities, topN, mode) {
      if (!datasets.length || !opportunities.length) {
        target.innerHTML = '<div class="empty">No heatmap data available.</div>';
        return;
      }

      const heatmapMode = mode === "actual" ? "actual" : "gap";
      const rows = datasets.slice();
      const columns = opportunities
        .slice()
        .sort((left, right) => right.totalRecoverableHours - left.totalRecoverableHours)
        .slice(0, Math.max(6, topN));
      const maxValue = Math.max(
        ...rows.flatMap((dataset) =>
          columns.map((opportunity) => {
            const actual = Number(opportunity.ranked.find((entry) => entry.datasetId === dataset.id)?.value || 0);
            return heatmapMode === "actual" ? actual : Math.max(actual - opportunity.idealTime, 0);
          })
        ),
        0
      );

      const headerHtml =
        '<tr><th>Well</th>' +
        columns.map((column) => '<th title="' + escapeHtml(getFlatTimeActivityTooltip(column.activityLabel) || column.activityLabel) + '">' + escapeHtml(column.activityLabel) + '</th>').join("") +
        '</tr>';
      const bodyHtml = rows.map((dataset) => {
        const cells = columns.map((opportunity) => {
          const actual = Number(opportunity.ranked.find((entry) => entry.datasetId === dataset.id)?.value || 0);
          const value = heatmapMode === "actual" ? actual : Math.max(actual - opportunity.idealTime, 0);
          const opacity = maxValue > 0 ? Math.max(0.12, value / maxValue) : 0;
          const bg = value > 0 ? 'rgba(200, 30, 90, ' + opacity.toFixed(2) + ')' : 'rgba(18, 100, 214, 0.06)';
          const color = value > 0.01 ? '#ffffff' : 'var(--ink)';
          const modeLabel = heatmapMode === "actual" ? "Actual hours" : "Gap vs ideal";
          const tooltip = modeLabel + ': ' + formatNumber(value) + ' hr' + ' | Actual: ' + formatNumber(actual) + ' hr | Ideal: ' + formatNumber(opportunity.idealTime) + ' hr';
          return '<td title="' + escapeHtml(tooltip) + '" style="background:' + bg + '; color:' + color + '; font-weight:700; text-align:center;">' + escapeHtml(formatNumber(value)) + '</td>';
        }).join("");
        return '<tr><td><strong>' + escapeHtml(dataset.subjectWell) + '</strong><br><span style="color:var(--muted); font-size:12px;">' + escapeHtml(dataset.rigLabel || '') + '</span></td>' + cells + '</tr>';
      }).join("");

      target.innerHTML = '<div class="table-wrap"><table><thead>' + headerHtml + '</thead><tbody>' + bodyHtml + '</tbody></table></div>';
    }

    function renderFlatTimeDatasetTags(datasets) {
      if (!datasets.length) {
        ui.flatTimeDatasetTags.innerHTML = '<span class="tag tag-muted">No flat time CSV datasets loaded yet.</span>';
        return;
      }

      ui.flatTimeDatasetTags.innerHTML = datasets
        .map((dataset) => {
          const removable = flatTimeState.uploadedDatasets.some((item) => item.id === dataset.id);
          return (
            '<span class="tag">' +
            '<strong>' + escapeHtml(dataset.subjectWell) + "</strong>" +
            '<span class="tag-muted">' + escapeHtml(dataset.rigLabel || "Rig not mapped") + " • " + escapeHtml(dataset.fileName) + "</span>" +
            (removable ? '<button type="button" data-flat-time-remove="' + escapeHtml(dataset.id) + '">Remove</button>' : "") +
            "</span>"
          );
        })
        .join("");

      Array.from(ui.flatTimeDatasetTags.querySelectorAll("[data-flat-time-remove]")).forEach((button) => {
        button.addEventListener("click", () => {
          flatTimeState.uploadedDatasets = flatTimeState.uploadedDatasets.filter((item) => item.id !== button.dataset.flatTimeRemove);
          renderFlatTime();
        });
      });
    }

    function renderFlatTimeSummary(datasets, metricKey, totalKey, activityItems, groupItems) {
      if (!datasets.length) {
        ui.flatTimeSummary.innerHTML = '<div class="empty">Upload flat time CSV files to start the comparison.</div>';
        return;
      }

      const topActivity = activityItems[0];
      const topGroup = groupItems[0];
      const highestDataset = datasets
        .map((dataset) => ({ label: dataset.subjectWell, rigLabel: dataset.rigLabel || "Rig not mapped", value: Number(dataset[totalKey] || 0) }))
        .sort((left, right) => right.value - left.value)[0];
      const opportunities = activityItems.map((item) => computeFlatTimeOpportunity(item, datasets));
      const topSpread = opportunities
        .slice()
        .sort((left, right) => right.gapToIdeal - left.gapToIdeal || right.gapVsPeerAverage - left.gapVsPeerAverage)[0];
      const totalRecoverable = opportunities.reduce((sum, opportunity) => sum + opportunity.totalRecoverableHours, 0);
      const mostReliableIdeal = opportunities
        .filter((opportunity) => opportunity.occurrenceCount >= 3)
        .slice()
        .sort((left, right) => {
          const confidenceScore = { High: 3, Medium: 2, Low: 1 };
          return (
            (confidenceScore[right.confidence] || 0) - (confidenceScore[left.confidence] || 0) ||
            left.cv - right.cv ||
            right.occurrenceCount - left.occurrenceCount
          );
        })[0];
      const overallHours = datasets.reduce((sum, dataset) => sum + Number(dataset[totalKey] || 0), 0);

      const cards = [
        { label: "Benchmarks Compared", value: String(datasets.length), meta: datasets.map((dataset) => dataset.subjectWell).join(", ") },
        { label: "Top Consuming Activity", value: topActivity ? topActivity.label : "N/A", valueHtml: topActivity ? flatTimeActivityLabelHtml(topActivity.label) : escapeHtml("N/A"), meta: topActivity ? formatNumber(topActivity.total) + " hr total" : "No activity data" },
        { label: "Largest Group", value: topGroup ? topGroup.label : "N/A", meta: topGroup ? formatNumber(topGroup.total) + " hr total" : "No group data" },
        { label: "Highest Burden Well", value: highestDataset ? highestDataset.label : "N/A", meta: highestDataset ? (highestDataset.rigLabel + " • " + formatNumber(highestDataset.value) + " hr total") : "No dataset totals" },
        {
          label: "Best Reduction Opportunity",
          value: topSpread ? topSpread.topEntry.rigLabel : "N/A",
          meta: topSpread ? (topSpread.topEntry.label + " • " + topSpread.activityLabel + " • gap " + formatNumber(topSpread.gapToIdeal) + " hr vs ideal") : "Need more than one dataset",
          metaHtml: topSpread ? (escapeHtml(topSpread.topEntry.label + " • ") + flatTimeActivityLabelHtml(topSpread.activityLabel) + escapeHtml(" • gap " + formatNumber(topSpread.gapToIdeal) + " hr vs ideal")) : escapeHtml("Need more than one dataset"),
        },
        {
          label: "Total Recoverable Hours",
          value: formatNumber(totalRecoverable),
          meta: "Sum of time above the recommended ideal across the comparison set",
        },
        {
          label: "Most Reliable Ideal",
          value: mostReliableIdeal ? mostReliableIdeal.activityLabel : "N/A",
          valueHtml: mostReliableIdeal ? flatTimeActivityLabelHtml(mostReliableIdeal.activityLabel) : escapeHtml("N/A"),
          meta: mostReliableIdeal ? (mostReliableIdeal.confidence + " confidence • target " + formatNumber(mostReliableIdeal.idealTime) + " hr") : "Need at least 3 wells for a strong benchmark",
        },
        { label: "Total Compared Time", value: formatNumber(overallHours), meta: metricKey === "subjectHours" ? "Subject well hours" : metricKey === "meanHours" ? "Mean hours" : "Median hours" },
      ];

      ui.flatTimeSummary.innerHTML = cards
        .map((card) => (
          '<div class="metric-pill">' +
          '<div class="label">' + escapeHtml(card.label) + "</div>" +
          '<div class="value"><span class="value-main">' + (card.valueHtml || escapeHtml(card.value)) + '</span></div>' +
          '<div class="meta">' + (card.metaHtml || escapeHtml(card.meta)) + "</div>" +
          "</div>"
        ))
        .join("");
    }

    function renderFlatTimeSeriesLegend(target, seriesDefs) {
      if (!target) return;
      if (!seriesDefs || !seriesDefs.length) {
        target.innerHTML = "";
        return;
      }
      target.innerHTML = seriesDefs
        .map((series) => (
          '<span class="legend-item">' +
          '<span class="legend-dot" style="background:' + escapeHtml(series.color) + ';"></span>' +
          escapeHtml(series.label) +
          '</span>'
        ))
        .join("");
    }

    function renderFlatTime() {
      const allDatasets = getFlatTimeDatasets();
      const metricKey = getFlatTimeMetricKey();
      const totalKey = getFlatTimeTotalKey();
      const topN = Number(ui.flatTimeTopN.value || 10);
      const selectedRig = ui.flatTimeRig.value || "";
      const selectedSectionSize = ui.flatTimeSection.value || "";
      flatTimeState.heatmapMode = ui.flatTimeHeatmapMode.value || flatTimeState.heatmapMode || "gap";

      renderFlatTimeDatasetTags(allDatasets);
      populateFlatTimeRigOptions(allDatasets);
      updateFlatTimeModeVisibility();

      if (!allDatasets.length) {
        ui.flatTimeTitle.textContent = "No flat time datasets loaded";
        ui.flatTimeSubtitle.textContent = "Use the CSV uploader to compare benchmark files.";
        ui.flatTimeSummary.innerHTML = '<div class="empty">Upload flat time CSV files to start the comparison.</div>';
        ui.flatTimeTableSummary.innerHTML = '<div class="empty">Upload flat time CSV files to build the selected well table summary.</div>';
        ui.flatTimeWellRanking.innerHTML = '<div class="empty">Upload flat time CSV files to rank wells by excess time.</div>';
        ui.flatTimeParetoChart.innerHTML = '<div class="empty">Upload flat time CSV files to build a Pareto of recoverable hours.</div>';
        ui.flatTimeAllWellsSections.innerHTML = '<div class="empty">Upload flat time CSV files to summarize all wells by section.</div>';
        ui.flatTimeSelectedWellSections.innerHTML = '<div class="empty">Upload flat time CSV files to compare the selected well by section.</div>';
        ui.flatTimeOffsetComparison.innerHTML = '<div class="empty">Upload flat time CSV files to compare the selected well against loaded offsets.</div>';
        ui.flatTimeSectionHeatmap.innerHTML = '<div class="empty">Upload flat time CSV files to draw the section heat map.</div>';
        ui.flatTimeSavingsSummary.innerHTML = '<div class="empty">Upload flat time CSV files to summarize recoverable time by section.</div>';
        ui.flatTimeNarrative.innerHTML = '<div class="empty">Upload flat time CSV files to generate the executive explanation.</div>';
        ui.flatTimeWaterfallChart.innerHTML = '<div class="empty">Upload flat time CSV files to build the waterfall.</div>';
        ui.flatTimeSectionBenchmarkChart.innerHTML = '<div class="empty">Upload flat time CSV files to compare sections.</div>';
        ui.flatTimeRigSummary.innerHTML = '<div class="empty">Upload flat time CSV files to summarize rigs.</div>';
        ui.flatTimeOpportunityPipeline.innerHTML = '<div class="empty">Upload flat time CSV files to build the opportunity pipeline.</div>';
        if (ui.flatTimeGroupLegend) ui.flatTimeGroupLegend.innerHTML = "";
        ui.flatTimeGroupChart.innerHTML = '<div class="empty">No flat time group data available.</div>';
        ui.flatTimeActivityChart.innerHTML = '<div class="empty">No flat time activity data available.</div>';
        ui.flatTimeBenchmarkTable.innerHTML = '<div class="empty">No activity benchmark table available.</div>';
        ui.flatTimeWellDrilldown.innerHTML = '<div class="empty">Upload flat time CSV files to inspect a well.</div>';
        ui.flatTimeActivityDrilldown.innerHTML = '<div class="empty">Upload flat time CSV files to inspect an activity.</div>';
        ui.flatTimeDrilldownNote.textContent = "Click a well or activity in the tables above to open the benchmark, peer comparison and ideal-time logic.";
        ui.flatTimeOpportunityTable.innerHTML = '<div class="empty">No flat time comparison table available.</div>';
        ui.flatTimeGroupTable.innerHTML = '<div class="empty">No flat time group table available.</div>';
        ui.flatTimeLossDrivers.innerHTML = '<div class="empty">Upload flat time CSV files to list top loss drivers by well.</div>';
        ui.flatTimeVariabilityChart.innerHTML = '<div class="empty">Upload flat time CSV files to show variability.</div>';
        ui.flatTimeHeatmap.innerHTML = '<div class="empty">No heatmap data available.</div>';
        if (ui.flatTimeHeatmapNote) {
          ui.flatTimeHeatmapNote.textContent = "Switch between Actual Hours and Gap vs Ideal to see either observed activity duration or only the hours above the recommended benchmark.";
        }
        ui.flatTimePerfectChart.innerHTML = '<div class="empty">Upload flat time CSV files to draw the perfect flat time curve.</div>';
        populateFlatTimeWellOptions([], "");
        return;
      }

      const rigDatasets = filterFlatTimeDatasetsByRig(allDatasets, selectedRig);
      populateFlatTimeSectionOptions(rigDatasets);
      const datasets = annotateFlatTimeScopedBenchmarks(
        filterFlatTimeDatasetsBySection(rigDatasets, selectedSectionSize)
      );

      if (!datasets.length) {
        ui.flatTimeTitle.textContent = "No data for selected section size";
        ui.flatTimeSubtitle.textContent = "Try another rig or section size, or switch back to all sections.";
        ui.flatTimeSummary.innerHTML = '<div class="empty">No benchmark activities match the selected section size.</div>';
        ui.flatTimeTableSummary.innerHTML = '<div class="empty">No selected well table summary is available for this section size.</div>';
        ui.flatTimeWellRanking.innerHTML = '<div class="empty">No well ranking available for this section size.</div>';
        ui.flatTimeParetoChart.innerHTML = '<div class="empty">No Pareto data available for this section size.</div>';
        ui.flatTimeAllWellsSections.innerHTML = '<div class="empty">No well-section summary available for this section size.</div>';
        ui.flatTimeSelectedWellSections.innerHTML = '<div class="empty">No selected-well section comparison available for this section size.</div>';
        ui.flatTimeOffsetComparison.innerHTML = '<div class="empty">No offset comparison available for this section size.</div>';
        ui.flatTimeSectionHeatmap.innerHTML = '<div class="empty">No section heat map available for this section size.</div>';
        ui.flatTimeSavingsSummary.innerHTML = '<div class="empty">No savings summary available for this section size.</div>';
        ui.flatTimeNarrative.innerHTML = '<div class="empty">No executive explanation available for this section size.</div>';
        ui.flatTimeWaterfallChart.innerHTML = '<div class="empty">No waterfall available for this section size.</div>';
        ui.flatTimeSectionBenchmarkChart.innerHTML = '<div class="empty">No section benchmark available for this section size.</div>';
        ui.flatTimeRigSummary.innerHTML = '<div class="empty">No rig benchmark summary available for this section size.</div>';
        ui.flatTimeOpportunityPipeline.innerHTML = '<div class="empty">No opportunity pipeline available for this section size.</div>';
        if (ui.flatTimeGroupLegend) ui.flatTimeGroupLegend.innerHTML = "";
        ui.flatTimeGroupChart.innerHTML = '<div class="empty">No flat time group data available for this section size.</div>';
        ui.flatTimeActivityChart.innerHTML = '<div class="empty">No flat time activity data available for this section size.</div>';
        ui.flatTimeBenchmarkTable.innerHTML = '<div class="empty">No benchmark table available for this section size.</div>';
        ui.flatTimeWellDrilldown.innerHTML = '<div class="empty">No well drill-down available for this section size.</div>';
        ui.flatTimeActivityDrilldown.innerHTML = '<div class="empty">No activity drill-down available for this section size.</div>';
        ui.flatTimeDrilldownNote.textContent = "Click a well or activity in the tables above to open the benchmark, peer comparison and ideal-time logic.";
        ui.flatTimeOpportunityTable.innerHTML = '<div class="empty">No flat time comparison table available for this section size.</div>';
        ui.flatTimeGroupTable.innerHTML = '<div class="empty">No flat time group table available for this section size.</div>';
        ui.flatTimeLossDrivers.innerHTML = '<div class="empty">No loss driver ranking available for this section size.</div>';
        ui.flatTimeVariabilityChart.innerHTML = '<div class="empty">No variability view available for this section size.</div>';
        ui.flatTimeHeatmap.innerHTML = '<div class="empty">No heatmap data available for this section size.</div>';
        if (ui.flatTimeHeatmapNote) {
          ui.flatTimeHeatmapNote.textContent = "Switch between Actual Hours and Gap vs Ideal to see either observed activity duration or only the hours above the recommended benchmark.";
        }
        ui.flatTimePerfectChart.innerHTML = '<div class="empty">No section-sized activities available for the perfect flat time curve.</div>';
        populateFlatTimeWellOptions([], "");
        return;
      }

      const groupItems = buildFlatTimeGroupItems(datasets, totalKey);
      const activityItems = buildFlatTimeActivityItems(datasets, metricKey);
      const sectionLabel = selectedSectionSize ? formatFlatTimeSectionSize(selectedSectionSize) : "All section sizes";
      if (ui.flatTimeHeatmapNote) {
        ui.flatTimeHeatmapNote.textContent =
          flatTimeState.heatmapMode === "actual"
            ? "You are seeing actual observed hours for each well and activity in the current comparison set. Use this view to compare raw execution time."
            : "You are seeing gap vs ideal hours for each well and activity. A value of 0 means the well met or beat the recommended ideal benchmark for that activity.";
      }

      ui.flatTimeTitle.textContent = datasets.length + " wells compared";
      ui.flatTimeSubtitle.textContent = (selectedRig || "All rigs") + " • " + sectionLabel + " • " + datasets.map((dataset) => dataset.subjectWell).join(" vs ");

      renderFlatTimeSummary(datasets, metricKey, totalKey, activityItems, groupItems);

      const seriesDefs = datasets.map((dataset, index) => ({
        key: dataset.id,
        label: dataset.subjectWell,
        color: FLAT_TIME_SERIES_COLORS[index % FLAT_TIME_SERIES_COLORS.length],
        format: (value) => formatNumber(value),
      }));
      renderFlatTimeSeriesLegend(ui.flatTimeGroupLegend, seriesDefs);

      renderMultiSeriesChart(ui.flatTimeGroupChart, groupItems.slice(0, 8), seriesDefs, {
        height: 430,
        minWidth: 880,
        groupMinWidth: 180,
      });
      renderMultiSeriesChart(ui.flatTimeActivityChart, activityItems.slice(0, topN), seriesDefs, {
        height: 430,
        minWidth: 880,
        groupMinWidth: 180,
      });

      const allOpportunities = activityItems
        .map((item) => computeFlatTimeOpportunity(item, datasets))
        .sort((left, right) => right.gapToIdeal - left.gapToIdeal || right.totalRecoverableHours - left.totalRecoverableHours || right.totalTime - left.totalTime);
      const rankedOpportunities = allOpportunities.slice(0, topN);
      const wellRanking = buildWellRanking(datasets, allOpportunities);
      const worstWell = wellRanking[0] ? wellRanking[0].wellLabel : "";
      const topActivityFocus = rankedOpportunities[0] ? rankedOpportunities[0].activityLabel : "";
      if (!flatTimeState.focusWell || !datasets.some((dataset) => dataset.subjectWell === flatTimeState.focusWell)) {
        flatTimeState.focusWell = worstWell;
      }
      const selectedActivities = normalizeFlatTimeFocusActivities(allOpportunities, topActivityFocus);
      populateFlatTimeWellOptions(datasets, flatTimeState.focusWell || worstWell);
      const selectedWell = ui.flatTimeWell.value || flatTimeState.focusWell || worstWell;
      const selectedDataset = datasets.find((dataset) => dataset.subjectWell === selectedWell) || datasets[0];
      const sectionBreakdownMap = buildSectionBreakdownMap(datasets, allOpportunities, metricKey);
      const selectedWellSectionItems = buildSelectedWellSectionItems(datasets, selectedDataset, sectionBreakdownMap);
      const allWellSectionRows = buildAllWellsSectionSummaryRows(datasets, sectionBreakdownMap);
      const sectionSavingsRows = buildSectionSavingsSummary(datasets, sectionBreakdownMap);
      const sectionBenchmarkItems = buildSectionBenchmarkItems(datasets, metricKey, allOpportunities);
      const rigBenchmarkRows = buildRigBenchmarkSummary(datasets, allOpportunities);
      const opportunityPipeline = buildOpportunityPipeline(allOpportunities).slice(0, Math.max(topN, 8));

      renderFlatTimeTableSummary(ui.flatTimeTableSummary, selectedDataset, allOpportunities, selectedSectionSize);

      renderTableHtml(
        ui.flatTimeWellRanking,
        ["Rig", "Well", "Actual Total (hr)", "Ideal Total (hr)", "Excess Time (hr)", "Top Drivers"],
        wellRanking.map((row) => [
          escapeHtml(row.rigLabel),
          flatTimeActionButtonHtml("well", row.wellLabel, row.wellLabel),
          escapeHtml(formatNumber(row.actualTotal)),
          escapeHtml(formatNumber(row.idealTotal)),
          flatTimeTrendHtml(row.excessTotal),
          row.topDrivers.length || row.otherDriversGap > 0
            ? [
                ...row.topDrivers.map((driver) => flatTimeActivityLabelHtml(driver.activity) + escapeHtml(" (" + driver.group + ", +" + formatNumber(driver.gap) + " hr)")),
                ...(row.otherDriversGap > 0 ? ['<span style="color:var(--muted); font-weight:700;">' + escapeHtml("Other drivers (+" + formatNumber(row.otherDriversGap) + " hr)") + '</span>'] : []),
              ].join("<br>")
            : escapeHtml("No excess detected"),
        ])
      );

      renderParetoChart(ui.flatTimeParetoChart, rankedOpportunities);
      renderAllWellsSectionSummary(
        ui.flatTimeAllWellsSections,
        allWellSectionRows,
        datasets,
        sectionBreakdownMap,
        selectedWell
      );
      renderExecutiveSelectedWellSections(ui.flatTimeSelectedWellSections, selectedWellSectionItems, selectedDataset);
      renderExecutiveOffsetComparison(ui.flatTimeOffsetComparison, selectedWellSectionItems, selectedDataset);
      renderSectionHeatmap(ui.flatTimeSectionHeatmap, datasets, sectionBreakdownMap);
      renderTableHtml(
        ui.flatTimeSavingsSummary,
        ["Section", "Recoverable (hr)", "Recoverable (d)", "Impacted Wells", "Top Activity", "Confidence", "Recommended Action"],
        sectionSavingsRows.map((row) => [
          escapeHtml(row.sectionLabel),
          escapeHtml(formatNumber(row.recoverableHours)),
          escapeHtml(formatNumber(row.recoverableDays)),
          escapeHtml(String(row.impactedWells)),
          flatTimeActivityLabelHtml(row.topActivity),
          confidenceBadgeHtml(row.confidence),
          escapeHtml(row.recommendation),
        ])
      );
      ui.flatTimeNarrative.innerHTML = buildExecutiveNarrativeHtml(selectedDataset, selectedWellSectionItems, sectionSavingsRows, wellRanking);
      renderWaterfallChart(ui.flatTimeWaterfallChart, selectedDataset, allOpportunities);
      renderMultiSeriesChart(
        ui.flatTimeSectionBenchmarkChart,
        sectionBenchmarkItems,
        [
          { key: "actualAverage", label: "Actual avg", color: "#1264d6", format: (value) => formatNumber(value) },
          { key: "idealTime", label: "Recommended ideal", color: "#0f766e", format: (value) => formatNumber(value) },
          { key: "spread", label: "Spread", color: "#c06a0a", format: (value) => formatNumber(value) },
        ],
        {
          height: 420,
          minWidth: 760,
          groupMinWidth: 150,
        }
      );

      renderTable(
        ui.flatTimeRigSummary,
        ["Rig", "# Wells", "Avg Flat Time (hr)", "Ideal Flat Time (hr)", "Excess Time (hr)", "Main Repeating Activity"],
        rigBenchmarkRows.map((row) => [
          row.rigLabel,
          String(row.wellCount),
          formatNumber(row.averageFlatTime),
          formatNumber(row.averageIdealTime),
          formatNumber(row.excessTime),
          row.mainRepeatingActivity,
        ])
      );

      renderTableHtml(
        ui.flatTimeOpportunityPipeline,
        ["Activity", "Group", "Occurrences", "Wells Impacted", "Ideal Time (hr)", "Recoverable Hours", "Priority"],
        opportunityPipeline.map((row) => [
          flatTimeActionButtonHtml("activity", row.activityLabel, row.activityLabel),
          escapeHtml(row.groupLabel),
          escapeHtml(String(row.occurrenceCount)),
          escapeHtml(String(row.wellsImpacted)),
          escapeHtml(formatNumber(row.idealTime)),
          escapeHtml(formatNumber(row.totalRecoverableHours)),
          escapeHtml(row.priority),
        ])
      );

      renderTableHtml(
        ui.flatTimeBenchmarkTable,
        ["Section", "Group", "Activity", "Sample", "Fastest", "P25", "Median", "Mean", "Recommended Ideal", "Variability", "Confidence", "Recoverable Hours", "Highest Well"],
        allOpportunities
          .slice(0, Math.max(topN * 2, 20))
          .map((opportunity) => [
            escapeHtml(formatFlatTimeSectionSize(opportunity.sectionSize)),
            escapeHtml(opportunity.groupLabel),
            flatTimeActionButtonHtml("activity", opportunity.activityLabel, opportunity.activityLabel),
            escapeHtml(String(opportunity.occurrenceCount)),
            escapeHtml(formatNumber(opportunity.fastestTime)),
            escapeHtml(formatNumber(opportunity.p25Value)),
            escapeHtml(formatNumber(opportunity.medianValue)),
            escapeHtml(formatNumber(opportunity.meanValue)),
            escapeHtml(formatNumber(opportunity.idealTime) + " (" + opportunity.idealRule + ")"),
            escapeHtml(opportunity.variability),
            confidenceBadgeHtml(opportunity.confidence),
            escapeHtml(formatNumber(opportunity.totalRecoverableHours)),
            flatTimeActionButtonHtml("well", opportunity.topEntry.label || "", (opportunity.topEntry.rigLabel || "Rig not mapped") + " • " + (opportunity.topEntry.label || "N/A")),
          ])
      );

      renderTableHtml(
        ui.flatTimeOpportunityTable,
        ["Section", "Group", "Activity", "Sample", "Highest Rig", "Highest Well", "Actual Time (hr)", "Peer Avg (hr)", "Ideal Time (hr)", "Gap To Ideal (hr)", "How Gap Was Calculated"],
        rankedOpportunities.map((opportunity) => {
          const peerReference = opportunity.peerAverage || opportunity.meanValue || opportunity.medianValue || 0;
          const explanation =
            opportunity.occurrenceCount >= 2
              ? (
                  (opportunity.occurrenceCount - 1) + " peer wells avg " + formatNumber(peerReference) +
                  " hr; " + opportunity.topEntry.label + " ran " + formatNumber(opportunity.topEntry.value) +
                  " hr; ideal = " + formatNumber(opportunity.idealTime) + " hr; gap = " + formatNumber(opportunity.gapToIdeal) + " hr"
                )
              : "Only one well available, so no peer comparison yet";

          return [
            escapeHtml(formatFlatTimeSectionSize(opportunity.sectionSize)),
            escapeHtml(opportunity.groupLabel),
            flatTimeActionButtonHtml("activity", opportunity.activityLabel, opportunity.activityLabel),
            escapeHtml(String(opportunity.occurrenceCount)),
            escapeHtml(opportunity.topEntry.rigLabel || "Rig not mapped"),
            flatTimeActionButtonHtml("well", opportunity.topEntry.label || "", opportunity.topEntry.label || "N/A"),
            escapeHtml(formatNumber(opportunity.topEntry.value)),
            escapeHtml(formatNumber(peerReference)),
            escapeHtml(formatNumber(opportunity.idealTime)),
            escapeHtml(formatNumber(opportunity.gapToIdeal)),
            escapeHtml(explanation + " (" + opportunity.idealRule + ")"),
          ];
        })
      );

      renderTable(
        ui.flatTimeGroupTable,
        ["Group", ...datasets.map((dataset) => dataset.subjectWell), "Total"],
        groupItems.map((item) => [
          item.label,
          ...datasets.map((dataset) => formatNumber(item[dataset.id] || 0)),
          formatNumber(item.total),
        ])
      );

      renderTable(
        ui.flatTimeLossDrivers,
        ["Rig", "Well", "Top Driver 1", "Top Driver 2", "Top Driver 3", "Excess Time (hr)"],
        wellRanking.map((row) => [
          row.rigLabel,
          row.wellLabel,
          row.topDrivers[0] ? row.topDrivers[0].activity + " (+" + formatNumber(row.topDrivers[0].gap) + " hr)" : "-",
          row.topDrivers[1] ? row.topDrivers[1].activity + " (+" + formatNumber(row.topDrivers[1].gap) + " hr)" : "-",
          row.topDrivers[2] ? row.topDrivers[2].activity + " (+" + formatNumber(row.topDrivers[2].gap) + " hr)" : "-",
          formatNumber(row.excessTotal),
        ])
      );

      renderVariabilityChart(ui.flatTimeVariabilityChart, allOpportunities, topN);
      renderHeatmap(ui.flatTimeHeatmap, datasets, allOpportunities, topN, flatTimeState.heatmapMode);
      renderPerfectFlatTimeChart(ui.flatTimePerfectChart, datasets, metricKey);
      renderFlatTimeDrilldown(datasets, allOpportunities, selectedWell, selectedActivities, rigDatasets);
      wireFlatTimeFocusActions();
    }

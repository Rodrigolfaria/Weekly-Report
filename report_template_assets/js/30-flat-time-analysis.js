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
                sectionSize: resolveFlatTimeSectionToken(activity.activity, activity.sectionSize),
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

    function normalizeFlatTimeFocusActivities(opportunities, fallbackActivity, options = {}) {
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
      if (!sanitized.length && !options.allowEmpty && fallbackActivity && validActivities.has(fallbackActivity)) {
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
            ? selectedItems.length + " activities across " + occurrenceCount + " wells; peers avg " + formatHours(peerAverage || meanValue || 0) + " hr, top well " + topEntry.label + " ran " + formatHours(topEntry.value) + " hr"
            : "Only one well observed for the selected activity set",
      };
    }

    function groupFlatTimeActivitiesForPicker(opportunities, selectedLabels) {
      const selectedSet = new Set(selectedLabels || []);
      const grouped = new Map();

      opportunities
        .slice()
        .sort((left, right) =>
          compareFlatTimeGroups(left.groupLabel, right.groupLabel) ||
          left.activityLabel.localeCompare(right.activityLabel)
        )
        .forEach((opportunity) => {
          const key = opportunity.groupLabel || "UNGROUPED";
          if (!grouped.has(key)) {
            grouped.set(key, {
              groupKey: key,
              groupDisplay: flatTimeGroupDisplayLabel(key),
              activities: [],
            });
          }
          grouped.get(key).activities.push({
            label: opportunity.activityLabel,
            checked: selectedSet.has(opportunity.activityLabel),
          });
        });

      return Array.from(grouped.values());
    }

    function flatTimeDatasetHasSection(dataset, sectionSize) {
      if (!dataset) return false;
      if (!sectionSize) return true;
      return (dataset.groups || []).some((group) =>
        (group.activities || []).some((activity) => {
          const activitySection = resolveFlatTimeSectionToken(activity.activity, activity.sectionSize);
          return activitySection === sectionSize;
        })
      );
    }

    function getFlatTimeSectionSizesForDataset(dataset) {
      if (!dataset) return [];
      return Array.from(
        new Set(
          (dataset.groups || []).flatMap((group) =>
            (group.activities || []).map((activity) => resolveFlatTimeSectionToken(activity.activity, activity.sectionSize))
          )
        )
      ).sort(compareFlatTimeSectionSizes);
    }

    function isFlatTimeOnlyOpportunity(opportunity) {
      if (!opportunity) return false;
      const sectionSize = String(resolveFlatTimeSectionToken(opportunity.activityLabel, opportunity.sectionSize) || "").trim();
      if (!sectionSize || sectionSize === "__no_section__") return false;
      const parts = extractFlatTimeComparisonParts(opportunity.activityLabel, opportunity.groupLabel);
      return !isFlatTimeDrillingAction(parts.action);
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
            ? occurrenceCount + " wells; peers avg " + formatHours(peerAverage || meanValue || 0) + " hr, top well " + topEntry.label + " ran " + formatHours(topEntry.value) + " hr"
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

    function buildSelectedWellParetoItems(selectedDataset, opportunities) {
      if (!selectedDataset) return [];

      return opportunities
        .map((opportunity) => {
          const actual = Number(opportunity.ranked.find((entry) => entry.datasetId === selectedDataset.id)?.value || 0);
          const recoverableHours = Math.max(actual - Number(opportunity.idealTime || 0), 0);
          return {
            ...opportunity,
            selectedWellActualHours: actual,
            selectedWellRecoverableHours: recoverableHours,
          };
        })
        .filter((opportunity) => opportunity.selectedWellActualHours > 0 && opportunity.selectedWellRecoverableHours > 0)
        .sort(
          (left, right) =>
            right.selectedWellRecoverableHours - left.selectedWellRecoverableHours ||
            right.selectedWellActualHours - left.selectedWellActualHours ||
            left.activityLabel.localeCompare(right.activityLabel)
        );
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
            const sectionSize = resolveFlatTimeSectionToken(activity.activity, activity.sectionSize);
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
            const sectionSize = resolveFlatTimeSectionToken(activity.activity, activity.sectionSize);
            if (!sectionSize || sectionSize === "__no_section__") return;
            const actual = Number(activity[metricKey] || 0);
            if (actual <= 0) return;

            const opportunity = opportunityMap.get(activity.activity);
            if (!opportunity) return;
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
          '<span class="section-driver-chip-value">' + escapeHtml(formatHours(Number(activity[valueKey] || 0)) + " hr") + "</span>" +
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
            '<span>' + escapeHtml(formatHours(group.actualTotal)) + ' hr actual</span>' +
            '<span>' + escapeHtml(formatHours(group.idealTotal)) + ' hr ideal</span>' +
            '<span>' + escapeHtml(formatHours(group.gapTotal)) + ' hr gap</span>' +
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
                "<td>" + escapeHtml(formatHours(section.actualTotal)) + "</td>" +
                "<td>" + escapeHtml(formatHours(section.idealTotal)) + "</td>" +
                "<td>" + escapeHtml(formatHours(section.gapTotal)) + "</td>" +
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
                "<strong>" + escapeHtml(formatHours(section.actualTotal)) + "</strong>" +
                '<span>actual</span>' +
                "<small>+" + escapeHtml(formatHours(section.gapTotal)) + " hr gap</small>" +
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
      if (!target) return;
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
          escapeHtml(formatHours(row.actualTotal)),
          escapeHtml(formatHours(row.idealTotal)),
          escapeHtml(formatHours(row.gapTotal)),
          row.topActivities.map((activity) => flatTimeActivityLabelHtml(activity.activityLabel) + escapeHtml(" (" + formatHours(activity.actual) + " hr)")).join("<br>"),
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

    const FLAT_TIME_DRILLING_ACTION_CODES = new Set(["D", "DMR", "DMS"]);

    function extractFlatTimeComparisonParts(activityLabel, fallbackGroupLabel) {
      const parts = String(activityLabel || "").split("-").filter(Boolean);
      return {
        phase: parts[0] || "__unknown__",
        majorOp: parts[1] || fallbackGroupLabel || "Unknown",
        action: parts.length ? parts[parts.length - 1] : "Unknown",
      };
    }

    function isFlatTimeDrillingAction(actionCode) {
      return FLAT_TIME_DRILLING_ACTION_CODES.has(String(actionCode || "").trim().toUpperCase());
    }

    function formatFlatTimeComparisonPhaseLabel(phase) {
      const token = String(phase || "").trim();
      if (!token) return "Unknown phase";
      if (token === "PRE" || token === "EOW") return token;
      if (/^\d+(?:\.\d+)?$/.test(token)) return formatFlatTimeSectionSize(token);
      return token;
    }

    function compareFlatTimePhases(left, right) {
      const leftToken = String(left || "").trim();
      const rightToken = String(right || "").trim();
      if (leftToken === rightToken) return 0;
      if (leftToken === "PRE") return -1;
      if (rightToken === "PRE") return 1;
      if (leftToken === "EOW") return 1;
      if (rightToken === "EOW") return -1;

      const leftIsNumeric = /^\d+(?:\.\d+)?$/.test(leftToken);
      const rightIsNumeric = /^\d+(?:\.\d+)?$/.test(rightToken);
      if (leftIsNumeric && rightIsNumeric) {
        return Number(rightToken) - Number(leftToken) || leftToken.localeCompare(rightToken);
      }
      if (leftIsNumeric) return -1;
      if (rightIsNumeric) return 1;
      return leftToken.localeCompare(rightToken);
    }

    function addFlatTimeComparisonValue(map, key, label, datasetId, value, meta = {}) {
      if (!map.has(key)) {
        map.set(key, {
          key,
          label,
          total: 0,
          ...meta,
        });
      }
      const item = map.get(key);
      item[datasetId] = Number(item[datasetId] || 0) + Number(value || 0);
      item.total += Number(value || 0);
      return item;
    }

    function finalizeFlatTimeComparisonItems(map, datasets, sorter) {
      return Array.from(map.values())
        .map((item) => {
          datasets.forEach((dataset) => {
            item[dataset.id] = Number(item[dataset.id] || 0);
          });
          return item;
        })
        .sort(
          sorter ||
            ((left, right) =>
              Number(right.total || 0) - Number(left.total || 0) ||
              String(left.label || "").localeCompare(String(right.label || "")))
        );
    }

    function buildFlatTimeComparisonData(datasets) {
      const phaseTimesMap = new Map();
      const drillingTimesMap = new Map();
      const phaseDetailsMap = new Map();
      const phasesSeen = new Set();

      datasets.forEach((dataset) => {
        dataset.groups.forEach((group) => {
          group.activities.forEach((activity) => {
            const hours = Number(activity.subjectHours || 0);
            if (hours <= 0) return;

            const parts = extractFlatTimeComparisonParts(activity.activity, group.groupName);
            const phase = parts.phase;
            phasesSeen.add(phase);

            if (!phaseDetailsMap.has(phase)) {
              phaseDetailsMap.set(phase, {
                phase,
                majorOpMap: new Map(),
                actionMap: new Map(),
                activityMap: new Map(),
              });
            }

            const phaseDetails = phaseDetailsMap.get(phase);

            if (isFlatTimeDrillingAction(parts.action)) {
              addFlatTimeComparisonValue(
                drillingTimesMap,
                phase,
                formatFlatTimeComparisonPhaseLabel(phase),
                dataset.id,
                hours,
                { phase }
              );
              return;
            }

            addFlatTimeComparisonValue(
              phaseDetails.activityMap,
              activity.activity,
              activity.activity,
              dataset.id,
              hours,
              {
                phase,
                majorOp: parts.majorOp,
                action: parts.action,
              }
            );

            addFlatTimeComparisonValue(
              phaseTimesMap,
              phase,
              formatFlatTimeComparisonPhaseLabel(phase),
              dataset.id,
              hours,
              { phase }
            );
            addFlatTimeComparisonValue(
              phaseDetails.majorOpMap,
              parts.majorOp,
              parts.majorOp,
              dataset.id,
              hours,
              { phase, majorOp: parts.majorOp }
            );
            addFlatTimeComparisonValue(
              phaseDetails.actionMap,
              parts.action,
              parts.action,
              dataset.id,
              hours,
              { phase, action: parts.action }
            );
          });
        });
      });

      const phaseOrder = Array.from(phasesSeen).sort(compareFlatTimePhases);
      const phaseTimesItems = phaseOrder
        .map((phase) => {
          const item =
            phaseTimesMap.get(phase) ||
            { key: phase, phase, label: formatFlatTimeComparisonPhaseLabel(phase), total: 0 };
          datasets.forEach((dataset) => {
            item[dataset.id] = Number(item[dataset.id] || 0);
          });
          return item;
        })
        .filter((item) => Number(item.total || 0) > 0);
      const drillingTimeItems = phaseOrder
        .map((phase) => {
          const item =
            drillingTimesMap.get(phase) ||
            { key: phase, phase, label: formatFlatTimeComparisonPhaseLabel(phase), total: 0 };
          datasets.forEach((dataset) => {
            item[dataset.id] = Number(item[dataset.id] || 0);
          });
          return item;
        })
        .filter((item) => Number(item.total || 0) > 0);
      const progressionItems = phaseOrder
        .map((phase) => {
          const phaseItem =
            phaseTimesMap.get(phase) ||
            { key: phase, phase, label: formatFlatTimeComparisonPhaseLabel(phase), total: 0 };
          const drillingItem =
            drillingTimesMap.get(phase) ||
            { key: phase, phase, label: formatFlatTimeComparisonPhaseLabel(phase), total: 0 };
          const item = {
            key: phase,
            phase,
            label: formatFlatTimeComparisonPhaseLabel(phase),
            totalFlatHours: Number(phaseItem.total || 0),
            totalDrillingHours: Number(drillingItem.total || 0),
          };
          datasets.forEach((dataset) => {
            item[dataset.id + "__flat"] = Number(phaseItem[dataset.id] || 0);
            item[dataset.id + "__drill"] = Number(drillingItem[dataset.id] || 0);
          });
          return item;
        })
        .filter((item) => Number(item.totalFlatHours || 0) > 0 || Number(item.totalDrillingHours || 0) > 0);

      const phaseDetails = phaseOrder
        .map((phase) => {
          const detail = phaseDetailsMap.get(phase);
          if (!detail) return null;
          const majorOpItems = finalizeFlatTimeComparisonItems(detail.majorOpMap, datasets);
          const actionItems = finalizeFlatTimeComparisonItems(detail.actionMap, datasets).slice(0, 10);
          const topActivities = finalizeFlatTimeComparisonItems(detail.activityMap, datasets).slice(0, 10);
          return {
            phase,
            phaseLabel: formatFlatTimeComparisonPhaseLabel(phase),
            majorOpItems,
            actionItems,
            topActivities,
          };
        })
        .filter(Boolean)
        .filter((detail) => detail.majorOpItems.length || detail.actionItems.length || detail.topActivities.length);

      return {
        phaseOrder,
        phaseTimesItems,
        drillingTimeItems,
        progressionItems,
        phaseDetails,
      };
    }

    function buildComparisonDatasetsOrder(datasets, selectedDataset) {
      return datasets
        .slice()
        .sort((left, right) => {
          if (selectedDataset) {
            if (left.id === selectedDataset.id && right.id !== selectedDataset.id) return -1;
            if (right.id === selectedDataset.id && left.id !== selectedDataset.id) return 1;
          }
          return left.subjectWell.localeCompare(right.subjectWell);
        });
    }

    function buildComparisonSeriesDefs(datasets, selectedDataset) {
      return buildComparisonDatasetsOrder(datasets, selectedDataset).map((dataset, index) => ({
        key: dataset.id,
        label: dataset.id === (selectedDataset && selectedDataset.id) ? dataset.subjectWell + " (selected)" : dataset.subjectWell,
        color: FLAT_TIME_SERIES_COLORS[index % FLAT_TIME_SERIES_COLORS.length],
        format: (value) => formatHours(value),
        isSelected: dataset.id === (selectedDataset && selectedDataset.id),
      }));
    }

    function buildFlatTimeComparisonDeltaHtml(value) {
      const numeric = Number(value || 0);
      const tone = numeric > 0 ? "comparison-delta-positive" : numeric < 0 ? "comparison-delta-negative" : "comparison-delta-neutral";
      const sign = numeric > 0 ? "+" : numeric < 0 ? "-" : "";
      return (
        '<div class="comparison-delta ' + tone + '">' +
        '<strong>' + escapeHtml(sign + formatHours(Math.abs(numeric)) + " hr") + '</strong>' +
        '<span>' + escapeHtml(formatDays(Math.abs(numeric) / 24) + " d") + '</span>' +
        "</div>"
      );
    }

    function buildFlatTimeComparisonValueHtml(hours, options = {}) {
      const numeric = Number(hours || 0);
      const tone = options.primary ? " comparison-table-value-primary" : "";
      return (
        '<div class="comparison-table-value' + tone + '">' +
        '<strong>' + escapeHtml(formatHours(numeric) + " hr") + '</strong>' +
        '<span>' + escapeHtml(formatDays(numeric / 24) + " d") + '</span>' +
        "</div>"
      );
    }

    function buildFlatTimeComparisonOffsetSnapshot(item, datasets, selectedDataset) {
      const selectedValue = Number(item[selectedDataset.id] || 0);
      const offsetValues = datasets
        .filter((dataset) => dataset.id !== selectedDataset.id)
        .map((dataset) => Number(item[dataset.id] || 0))
        .filter((value) => value > 0);
      const offsetAverage = offsetValues.length ? average(offsetValues) : 0;
      const bestOffset = offsetValues.length ? Math.min(...offsetValues) : 0;
      return {
        selectedValue,
        offsetAverage,
        bestOffset,
        differenceVsAverage: selectedValue - offsetAverage,
        differenceVsBest: selectedValue - bestOffset,
        offsetCount: offsetValues.length,
        selectedOnly: selectedValue > 0 && !offsetValues.length,
      };
    }

    function buildFlatTimeComparisonSelectedOnlyBadgeHtml() {
      return '<span class="comparison-flag">Selected well only</span>';
    }

    function buildFlatTimeComparisonDisplayActivities(items, datasets, selectedDataset, limit = 10) {
      const enriched = items.map((item) => ({
        item,
        snapshot: buildFlatTimeComparisonOffsetSnapshot(item, datasets, selectedDataset),
      }));

      const selectedOnly = enriched
        .filter((entry) => entry.snapshot.selectedOnly)
        .sort((left, right) =>
          right.snapshot.selectedValue - left.snapshot.selectedValue ||
          String(left.item.label || "").localeCompare(String(right.item.label || ""))
        );

      const remaining = enriched
        .filter((entry) => !entry.snapshot.selectedOnly)
        .sort((left, right) =>
          Number(right.item.total || 0) - Number(left.item.total || 0) ||
          String(left.item.label || "").localeCompare(String(right.item.label || ""))
        );

      return [...selectedOnly, ...remaining].slice(0, limit);
    }

    function renderFlatTimeComparisonOverview(target, items, datasets, selectedDataset, emptyMessage, options = {}) {
      if (!target) return;
      if (!selectedDataset || !items.length) {
        target.innerHTML = '<div class="empty">' + escapeHtml(emptyMessage) + '</div>';
        return;
      }

      const orderedDatasets = buildComparisonDatasetsOrder(datasets, selectedDataset);
      const seriesDefs = buildComparisonSeriesDefs(orderedDatasets, selectedDataset);

      target.innerHTML =
        '<div class="legend"></div>' +
        '<div class="comparison-chart-body"></div>' +
        '<div class="comparison-table-toolbar"><span class="chip">Selected well: ' + escapeHtml(selectedDataset.subjectWell) + '</span></div>' +
        '<div class="comparison-phase-table"></div>';
      renderFlatTimeSeriesLegend(target.querySelector(".legend"), seriesDefs);
      renderMultiSeriesChart(
        target.querySelector(".comparison-chart-body"),
        items,
        seriesDefs,
        {
          height: 400,
          minWidth: 860,
          groupMinWidth: 150,
        }
      );
      renderTableHtml(
        target.querySelector(".comparison-phase-table"),
        [
          options.labelHeader || "Phase",
          "Selected",
          "Offset Avg",
          "Best Offset",
          "Vs Avg",
          "Vs Best",
        ],
        items.map((item) => {
          const snapshot = buildFlatTimeComparisonOffsetSnapshot(item, orderedDatasets, selectedDataset);
          return [
            escapeHtml(item.label || item.key || ""),
            buildFlatTimeComparisonValueHtml(snapshot.selectedValue, { primary: true }),
            snapshot.offsetCount
              ? buildFlatTimeComparisonValueHtml(snapshot.offsetAverage)
              : '<span class="table-summary-missing">-</span>',
            snapshot.offsetCount
              ? buildFlatTimeComparisonValueHtml(snapshot.bestOffset)
              : '<span class="table-summary-missing">-</span>',
            snapshot.offsetCount ? buildFlatTimeComparisonDeltaHtml(snapshot.differenceVsAverage) : '<span class="table-summary-missing">-</span>',
            snapshot.offsetCount ? buildFlatTimeComparisonDeltaHtml(snapshot.differenceVsBest) : '<span class="table-summary-missing">-</span>',
          ];
        })
      );
    }

    function renderFlatTimeComparisonPhaseDetails(target, phaseDetails, datasets, selectedDataset) {
      if (!target) return;
      if (!selectedDataset || !phaseDetails.length) {
        target.innerHTML = '<div class="empty">No comparison detail is available for the loaded phases.</div>';
        return;
      }

      const orderedDatasets = buildComparisonDatasetsOrder(datasets, selectedDataset);
      const seriesDefs = buildComparisonSeriesDefs(orderedDatasets, selectedDataset);

      target.innerHTML =
        '<div class="comparison-phase-stack">' +
        phaseDetails
          .map((detail, index) => (
            '<div class="comparison-phase-block">' +
            '<div class="card-toolbar">' +
            '<div><h3 style="margin-bottom:6px;">' + escapeHtml(detail.phaseLabel + " Phase") + '</h3><p class="report-note">Major OP, Action, and top activity comparison for this phase. Drilling actions D, DMR, and DMS are excluded from the charts and the table below. Activities flagged as <strong>Selected well only</strong> appear only in the selected well and should be reviewed for possible time loss.</p></div>' +
            '<span class="chip">' + escapeHtml(String(detail.topActivities.length) + " top activities shown") + '</span>' +
            '</div>' +
            '<div class="comparison-phase-grid">' +
            '<div class="comparison-phase-panel"><h4>Major OP</h4><div id="flat-time-comparison-major-legend-' + index + '" class="legend"></div><div id="flat-time-comparison-major-chart-' + index + '"></div></div>' +
            '<div class="comparison-phase-panel"><h4>Actions</h4><div id="flat-time-comparison-action-legend-' + index + '" class="legend"></div><div id="flat-time-comparison-action-chart-' + index + '"></div></div>' +
            '</div>' +
            '<div class="comparison-table-toolbar"><span class="chip">Selected well: ' + escapeHtml(selectedDataset.subjectWell) + '</span></div>' +
            '<div class="comparison-phase-table" id="flat-time-comparison-phase-table-' + index + '"></div>' +
            '</div>'
          ))
          .join("") +
        '</div>';

      phaseDetails.forEach((detail, index) => {
        const majorLegend = target.querySelector("#flat-time-comparison-major-legend-" + index);
        const majorChart = target.querySelector("#flat-time-comparison-major-chart-" + index);
        const actionLegend = target.querySelector("#flat-time-comparison-action-legend-" + index);
        const actionChart = target.querySelector("#flat-time-comparison-action-chart-" + index);
        const tableTarget = target.querySelector("#flat-time-comparison-phase-table-" + index);
        const displayActivities = buildFlatTimeComparisonDisplayActivities(detail.topActivities, orderedDatasets, selectedDataset, 10);

        renderFlatTimeSeriesLegend(majorLegend, seriesDefs);
        renderMultiSeriesChart(
          majorChart,
          detail.majorOpItems,
          seriesDefs,
          {
            height: 360,
            minWidth: 760,
            groupMinWidth: 140,
          }
        );

        renderFlatTimeSeriesLegend(actionLegend, seriesDefs);
        renderMultiSeriesChart(
          actionChart,
          detail.actionItems,
          seriesDefs,
          {
            height: 360,
            minWidth: 760,
            groupMinWidth: 140,
          }
        );

        renderTableHtml(
          tableTarget,
          [
            "Activity",
            "Major OP",
            "Action",
            "Selected",
            "Offset Avg",
            "Best Offset",
            "Vs Avg",
            "Vs Best",
          ],
          displayActivities.map(({ item, snapshot }) => {
            return [
              flatTimeActivityLabelHtml(item.label) + (snapshot.selectedOnly ? "<br>" + buildFlatTimeComparisonSelectedOnlyBadgeHtml() : ""),
              escapeHtml(item.majorOp || "Unknown"),
              escapeHtml(item.action || "Unknown"),
              buildFlatTimeComparisonValueHtml(snapshot.selectedValue, { primary: true }),
              snapshot.offsetCount
                ? buildFlatTimeComparisonValueHtml(snapshot.offsetAverage)
                : '<span class="table-summary-missing">No offset record</span>',
              snapshot.offsetCount
                ? buildFlatTimeComparisonValueHtml(snapshot.bestOffset)
                : '<span class="table-summary-missing">No offset record</span>',
              snapshot.offsetCount ? buildFlatTimeComparisonDeltaHtml(snapshot.differenceVsAverage) : buildFlatTimeComparisonSelectedOnlyBadgeHtml(),
              snapshot.offsetCount ? buildFlatTimeComparisonDeltaHtml(snapshot.differenceVsBest) : buildFlatTimeComparisonSelectedOnlyBadgeHtml(),
            ];
          })
        );
      });
    }

    function renderFlatTimeComparisonProgression(target, progressionItems, datasets, selectedDataset) {
      if (!target) return;
      if (!progressionItems.length || !datasets.length || !selectedDataset) {
        target.innerHTML = '<div class="empty">No phase progression data is available for the loaded wells.</div>';
        return;
      }

      const visibleItems = progressionItems.filter((item) => String(item.phase || "").trim() !== "PRE");
      if (!visibleItems.length) {
        target.innerHTML = '<div class="empty">No section progression is available after excluding the PRE phase.</div>';
        return;
      }

      const phaseOrder = visibleItems.map((item) => item.phase);
      const chartTheme = getChartTheme();
      const width = 1160;
      const height = 520;
      const margin = { top: 34, right: 170, bottom: 78, left: 104 };
      const chartWidth = width - margin.left - margin.right;
      const chartHeight = height - margin.top - margin.bottom;
      const yPadTop = 12;
      const yPadBottom = 46;
      const yUsableHeight = Math.max(chartHeight - yPadTop - yPadBottom, 1);

      const orderedDatasets = buildComparisonDatasetsOrder(datasets, selectedDataset);
      const hiddenDatasetIds = new Set(Array.isArray(flatTimeState.progressionHiddenDatasetIds) ? flatTimeState.progressionHiddenDatasetIds : []);
      if (orderedDatasets.length && orderedDatasets.every((dataset) => hiddenDatasetIds.has(dataset.id))) {
        flatTimeState.progressionHiddenDatasetIds = [];
        hiddenDatasetIds.clear();
      }
      const allSeries = orderedDatasets.map((dataset, index) => {
        let cumulativeDays = 0;
        const points = [];
        visibleItems.forEach((item, phaseIndex) => {
          const currentPhaseIndex = phaseIndex + 1;
          const flatDays = Number(item[dataset.id + "__flat"] || 0) / 24;
          const drillDays = Number(item[dataset.id + "__drill"] || 0) / 24;
          if (!points.length) {
            points.push({
              label: item.label,
              cumulativeDays: 0,
              phaseIndex: currentPhaseIndex,
              segment: "start",
            });
          }
          cumulativeDays += flatDays;
          points.push({
            label: item.label,
            cumulativeDays,
            phaseIndex: currentPhaseIndex,
            segment: "flat",
            flatDays,
            drillDays,
          });
          if (phaseIndex < visibleItems.length - 1) {
            cumulativeDays += drillDays;
            points.push({
              label: item.label,
              cumulativeDays,
              phaseIndex: currentPhaseIndex + 1,
              segment: "drill",
              flatDays,
              drillDays,
            });
          } else if (drillDays > 0) {
            cumulativeDays += drillDays;
            points.push({
              label: item.label,
              cumulativeDays,
              phaseIndex: currentPhaseIndex,
              segment: "drill-final",
              flatDays,
              drillDays,
            });
          }
        });
        return {
          id: dataset.id,
          label: dataset.id === selectedDataset.id ? dataset.subjectWell + " (selected)" : dataset.subjectWell,
          color: FLAT_TIME_SERIES_COLORS[index % FLAT_TIME_SERIES_COLORS.length],
          isSelected: dataset.id === selectedDataset.id,
          isHidden: hiddenDatasetIds.has(dataset.id),
          points,
        };
      });
      const series = allSeries.filter((line) => !line.isHidden);

      const domainMaxDays = Math.max(1, Math.max(
        ...series.flatMap((line) => line.points.map((point) => point.cumulativeDays)),
        1
      ) * 1.12);
      const maxDepth = Math.max(phaseOrder.length, 1);

      const scaledSeries = series.map((line) => ({
        ...line,
        points: line.points.map((point) => ({
          ...point,
          x: margin.left + (point.cumulativeDays / domainMaxDays) * chartWidth,
          y: margin.top + yPadTop + ((point.phaseIndex - 1) / Math.max(maxDepth - 1, 1)) * yUsableHeight,
        })),
      }));

      const phaseBands = phaseOrder
        .map((_phase, index) => {
          if (index % 2 !== 0) return "";
          const y1 = margin.top + yPadTop + (index / Math.max(maxDepth - 1, 1)) * yUsableHeight;
          const y2 = index >= maxDepth - 1
            ? height - margin.bottom
            : margin.top + yPadTop + ((index + 1) / Math.max(maxDepth - 1, 1)) * yUsableHeight;
          return '<rect x="' + margin.left + '" y="' + y1.toFixed(2) + '" width="' + chartWidth + '" height="' + Math.max(y2 - y1, 0).toFixed(2) + '" fill="rgba(21,95,184,0.025)"></rect>';
        })
        .join("");

      const xTicks = Array.from({ length: 6 }, (_, index) => {
        const value = (domainMaxDays / 5) * index;
        const x = margin.left + (value / domainMaxDays) * chartWidth;
        return (
          '<g>' +
          '<line x1="' + x.toFixed(2) + '" y1="' + margin.top + '" x2="' + x.toFixed(2) + '" y2="' + (height - margin.bottom) + '" stroke="' + chartTheme.grid + '" stroke-dasharray="4 5"></line>' +
          '<text x="' + x.toFixed(2) + '" y="' + (height - 12) + '" text-anchor="middle" font-size="11" fill="' + chartTheme.valueLabel + '">' + escapeHtml(formatDays(value)) + "</text>" +
          "</g>"
        );
      }).join("");

      const yTicks = phaseOrder
        .map((phase, index) => {
          const y = margin.top + yPadTop + (index / Math.max(maxDepth - 1, 1)) * yUsableHeight;
          return (
            '<g>' +
            '<line x1="' + margin.left + '" y1="' + y.toFixed(2) + '" x2="' + (width - margin.right) + '" y2="' + y.toFixed(2) + '" stroke="' + chartTheme.grid + '" stroke-dasharray="4 5"></line>' +
            '<text x="' + (margin.left - 12) + '" y="' + (y + 4).toFixed(2) + '" text-anchor="end" font-size="11" font-weight="700" fill="' + chartTheme.text + '">' + escapeHtml(formatFlatTimeComparisonPhaseLabel(phase)) + "</text>" +
            '</g>'
          );
        })
        .join("");

      const lineSvg = scaledSeries
        .map((line, index) => {
          const path = line.points.map((point, index) => (index === 0 ? "M" : "L") + point.x.toFixed(2) + " " + point.y.toFixed(2)).join(" ");
          const endPoint = line.points[line.points.length - 1];
          const labelYOffset = (index - (scaledSeries.length - 1) / 2) * 12;
          const labelY = Math.max(margin.top + 12, Math.min(height - margin.bottom - 8, endPoint.y + labelYOffset));
          const labelAnchorLeft = endPoint.x > width - margin.right - 120;
          const labelX = labelAnchorLeft ? Math.max(margin.left + 8, endPoint.x - 12) : Math.min(width - 10, endPoint.x + 14);
          const labelAnchor = labelAnchorLeft ? "end" : "start";
          return (
            '<g class="trajectory-series' + (line.isSelected ? " is-selected" : "") + '">' +
            '<path d="' + path + '" fill="none" stroke="rgba(255,255,255,0.72)" stroke-width="' + (line.isSelected ? "7.2" : "5.4") + '" stroke-linecap="round" stroke-linejoin="round"></path>' +
            '<path d="' + path + '" fill="none" stroke="' + line.color + '" stroke-width="' + (line.isSelected ? "4.6" : "3.1") + '" stroke-linecap="round" stroke-linejoin="round">' +
            '<title>' + escapeHtml(line.label + " • " + formatDays(endPoint.cumulativeDays) + " d") + '</title>' +
            '</path>' +
            '<circle cx="' + endPoint.x.toFixed(2) + '" cy="' + endPoint.y.toFixed(2) + '" r="' + (line.isSelected ? "4.2" : "3.2") + '" fill="' + line.color + '" stroke="#ffffff" stroke-width="2"></circle>' +
            '<text x="' + labelX.toFixed(2) + '" y="' + labelY.toFixed(2) + '" text-anchor="' + labelAnchor + '" font-size="10.5" font-weight="' + (line.isSelected ? "800" : "680") + '" fill="' + chartTheme.pointLabel + '">' + escapeHtml(line.label + " • " + formatDays(endPoint.cumulativeDays) + " d") + '</text>' +
            '</g>'
          );
        })
        .join("");

      const legend = allSeries
        .map((line) => (
          '<button class="trajectory-legend-item' + (line.isHidden ? " is-muted" : "") + (line.isSelected ? " is-selected" : "") + '" type="button" data-progression-toggle="' + escapeHtml(line.id) + '" title="Click to show or hide this well">' +
          '<span class="legend-dot" style="background:' + line.color + ';"></span>' +
          '<span>' + escapeHtml(line.label) + '</span>' +
          '</button>'
        ))
        .join("");

      target.innerHTML =
        '<div class="trajectory-chart-shell">' +
        '<div class="trajectory-chart-topline">' +
        '<div class="trajectory-legend" aria-label="Toggle wells in chart">' + legend + '</div>' +
        '<span class="trajectory-hint">Click a well to hide/show it</span>' +
        '</div>' +
        '<p class="report-note">PRE is excluded. Horizontal segments are flat time by phase; vertical transitions represent drilling time from actions D, DMR, and DMS.</p>' +
        '<div class="column-chart-wrap trajectory-chart-wrap">' +
        '<svg class="column-chart-svg trajectory-svg" viewBox="0 0 ' + width + " " + height + '" role="img" aria-label="Phase progression by days chart">' +
        phaseBands +
        xTicks +
        yTicks +
        '<line x1="' + margin.left + '" y1="' + margin.top + '" x2="' + margin.left + '" y2="' + (height - margin.bottom) + '" stroke="' + chartTheme.axis + '"></line>' +
        '<line x1="' + margin.left + '" y1="' + (height - margin.bottom) + '" x2="' + (width - margin.right) + '" y2="' + (height - margin.bottom) + '" stroke="' + chartTheme.axis + '"></line>' +
        lineSvg +
        '<text x="' + (margin.left + chartWidth / 2).toFixed(2) + '" y="' + (height - 8) + '" text-anchor="middle" font-size="12" font-weight="700" fill="' + chartTheme.text + '">Days</text>' +
        '<text x="20" y="' + (margin.top + chartHeight / 2).toFixed(2) + '" text-anchor="middle" font-size="12" font-weight="700" fill="' + chartTheme.text + '" transform="rotate(-90 20 ' + (margin.top + chartHeight / 2).toFixed(2) + ')">Depth / Phase Progression</text>' +
        "</svg></div></div>";

      target.querySelectorAll("[data-progression-toggle]").forEach((button) => {
        button.addEventListener("click", () => {
          const datasetId = button.getAttribute("data-progression-toggle") || "";
          if (!datasetId) return;
          const currentHidden = new Set(Array.isArray(flatTimeState.progressionHiddenDatasetIds) ? flatTimeState.progressionHiddenDatasetIds : []);
          if (currentHidden.has(datasetId)) {
            currentHidden.delete(datasetId);
          } else if (orderedDatasets.filter((dataset) => !currentHidden.has(dataset.id)).length > 1) {
            currentHidden.add(datasetId);
          }
          flatTimeState.progressionHiddenDatasetIds = Array.from(currentHidden);
          renderFlatTime();
        });
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
        .map((item) => item.activityLabel + " (" + formatHours(item.gap) + " hr)")
        .join(", ");

      const paragraphs = [
        '<p><strong>' + escapeHtml(selectedDataset.subjectWell) + '</strong> is currently running with <strong>' + escapeHtml(formatHours(rankingRow?.excessTotal || 0)) + ' hr</strong> above the recommended flat time benchmark. The section creating the largest burden is <strong>' + escapeHtml(worstSection?.label || "N/A") + '</strong>, where the well is spending <strong>' + escapeHtml(formatHours(worstSection?.actualTotal || 0)) + ' hr</strong> against an ideal target of <strong>' + escapeHtml(formatHours(worstSection?.idealTotal || 0)) + ' hr</strong>.</p>',
        '<p>The system selected this section because it combines the largest recoverable gap with repeatable activities seen across the loaded offsets. In this section, the main drivers are <strong>' + escapeHtml(topActivities || "no recurring drivers identified") + '</strong>. This means the opportunity is not just the single slowest event, but a recurring pattern where the selected well is running slower than the best validated performance.</p>',
        strongestSavings
          ? '<p>The most actionable recovery area across the current benchmark set is <strong>' + escapeHtml(strongestSavings.sectionLabel) + '</strong>, with about <strong>' + escapeHtml(formatHours(strongestSavings.recoverableHours)) + ' hr</strong> (' + escapeHtml(formatDays(strongestSavings.recoverableDays)) + ' d) available to recover. The recommended action is to attack <strong>' + flatTimeActivityLabelHtml(strongestSavings.topActivity) + '</strong> first because it is the strongest repeated source of excess time and already has a <strong>' + confidenceBadgeHtml(strongestSavings.confidence) + '</strong> benchmark behind it. To recover time, replicate the best observed sequence, remove waiting between dependent steps, and standardize the execution around the loaded offsets that already achieved the lower time.</p>'
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
          { key: "actual", label: selectedDataset.subjectWell + " actual", color: "#1264d6", format: (value) => formatHours(value) },
          { key: "ideal", label: "Recommended ideal", color: "#0f766e", format: (value) => formatHours(value) },
          { key: "gap", label: "Recoverable gap", color: "#c06a0a", format: (value) => formatHours(value) },
        ],
        { height: 420, minWidth: 760, groupMinWidth: 150 }
      );

      renderTableHtml(
        tableTarget,
        ["Section", "Actual (hr)", "Synthetic Ideal (hr)", "Gap (hr)", "Top 3 Time Reduction Opportunities"],
        sectionItems.map((item) => [
          escapeHtml(item.label),
          escapeHtml(formatHours(item.actualTotal)),
          escapeHtml(formatHours(item.idealTotal)),
          escapeHtml(formatHours(item.gapTotal)),
          (item.topGapActivities && item.topGapActivities.length ? item.topGapActivities : item.topActivities)
            .map((activity) => flatTimeActivityLabelHtml(activity.activityLabel) + escapeHtml(" (+" + formatHours(activity.gap || 0) + " hr / " + formatDays((activity.gap || 0) / 24) + " d)"))
            .join("<br>"),
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
          { key: "selected", label: selectedDataset.subjectWell, color: "#1264d6", format: (value) => formatHours(value) },
          { key: "offsetAverage", label: "Offset avg", color: "#0f766e", format: (value) => formatHours(value) },
          { key: "bestOffset", label: "Best offset", color: "#7c3aed", format: (value) => formatHours(value) },
          { key: "ideal", label: "Recommended ideal", color: "#c06a0a", format: (value) => formatHours(value) },
        ],
        { height: 420, minWidth: 820, groupMinWidth: 160 }
      );

      renderTable(
        tableTarget,
        ["Section", selectedDataset.subjectWell + " (hr)", "Offset Avg (hr)", "Best Offset (hr)", "Best Offset Well", "Synthetic Ideal (hr)", "Gap vs Offset Avg (hr)"],
        sectionItems.map((item) => [
          item.label,
          formatHours(item.actualTotal),
          formatHours(item.offsetAverage),
          formatHours(item.bestOffset),
          item.bestOffsetWell === "No peer" ? item.bestOffsetWell : item.bestOffsetWell + " (" + item.bestOffsetRig + ")",
          formatHours(item.idealTotal),
          formatHours(item.gapVsOffsetAverage),
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
          return '<td style="background:' + bg + '; color:' + color + '; font-weight:700; text-align:center;">' + escapeHtml(formatHours(gap)) + '</td>';
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
          '<text x="' + (margin.left - 10) + '" y="' + (y + 4).toFixed(2) + '" text-anchor="end" font-size="11" fill="' + chartTheme.valueLabel + '">' + escapeHtml(formatHours(value)) + '</text>' +
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
        const displayValue = step.type === "delta" ? formatHours(Math.abs(step.delta || 0)) : formatHours(step.end);
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
          '<text x="' + x.toFixed(2) + '" y="' + (height - 10) + '" text-anchor="middle" font-size="11" fill="' + chartTheme.valueLabel + '">' + escapeHtml(formatHours(value)) + '</text>' +
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
      const wellOptionDatasets = sectionOptionDatasets && sectionOptionDatasets.length ? sectionOptionDatasets : datasets;
      const drilldownSelectedDataset = wellOptionDatasets.find((dataset) => dataset.subjectWell === selectedWell) || datasets.find((dataset) => dataset.subjectWell === selectedWell) || wellOptionDatasets[0] || datasets[0];
      const selectedDataset = datasets.find((dataset) => dataset.subjectWell === selectedWell) || datasets[0] || drilldownSelectedDataset;
      const explorerOpportunities = (opportunities || []).filter(isFlatTimeOnlyOpportunity);
      const selectedOpportunity = buildSelectedActivityAggregate(explorerOpportunities, selectedActivities);

      if (!selectedDataset) {
        ui.flatTimeWellDrilldown.innerHTML = '<div class="empty">Select more CSVs to unlock the drill-down.</div>';
        ui.flatTimeActivityDrilldown.innerHTML = '<div class="empty">Select more CSVs to unlock the activity benchmark.</div>';
        return;
      }

      const allWellDrivers = explorerOpportunities
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
      const wellOptions = wellOptionDatasets
        .slice()
        .sort((left, right) => left.subjectWell.localeCompare(right.subjectWell))
        .map((dataset) => (
          '<option value="' + escapeHtml(dataset.subjectWell) + '"' +
          (dataset.subjectWell === drilldownSelectedDataset.subjectWell ? " selected" : "") +
          '>' + escapeHtml(dataset.subjectWell + (dataset.rigLabel ? " • " + dataset.rigLabel : "")) + "</option>"
        ))
        .join("");
      const availableSections = getFlatTimeSectionSizesForDataset(drilldownSelectedDataset);
      const selectedSectionValue = availableSections.includes(ui.flatTimeSection.value || "") ? (ui.flatTimeSection.value || "") : "";
      const sectionOptions = ['<option value="">All section sizes</option>']
        .concat(
          availableSections.map((sectionSize) => (
            '<option value="' + escapeHtml(sectionSize) + '"' +
            (sectionSize === selectedSectionValue ? " selected" : "") +
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
        '<div class="metric-pill"><div class="label">Actual Flat Time</div><div class="value"><span class="value-main">' + escapeHtml(formatHours(wellActualTotal)) + '</span><span class="value-suffix">' + escapeHtml("hr / " + formatDays(wellActualTotal / 24) + " d") + '</span></div><div class="meta">All activities in the selected filter context</div></div>' +
        '<div class="metric-pill"><div class="label">Ideal Flat Time</div><div class="value"><span class="value-main">' + escapeHtml(formatHours(wellIdealTotal)) + '</span><span class="value-suffix">' + escapeHtml("hr / " + formatDays(wellIdealTotal / 24) + " d") + '</span></div><div class="meta">Recommended achievable total for the same activities</div></div>' +
        '<div class="metric-pill"><div class="label">Recoverable Gap</div><div class="value"><span class="value-main">' + escapeHtml(formatHours(wellExcess)) + '</span><span class="value-suffix">' + escapeHtml("hr / " + formatDays(wellExcess / 24) + " d") + '</span></div><div class="meta">Time above the recommended ideal</div></div>' +
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
          const nextWell = drilldownWellSelect.value || "";
          flatTimeState.focusWell = nextWell;
          ui.flatTimeWell.value = nextWell;
          const activeSection = ui.flatTimeSection.value || "";
          const matchingDataset = wellOptionDatasets.find((dataset) => dataset.subjectWell === nextWell);
          if (activeSection && !flatTimeDatasetHasSection(matchingDataset, activeSection)) {
            ui.flatTimeSection.value = "";
          }
          renderFlatTime();
        });
      }

      const drilldownSectionSelect = document.getElementById("flat-time-drilldown-section-select");
      if (drilldownSectionSelect) {
        drilldownSectionSelect.addEventListener("change", () => {
          const nextSection = drilldownSectionSelect.value || "";
          ui.flatTimeSection.value = nextSection;
          renderFlatTime();
        });
      }

      const activityGroups = groupFlatTimeActivitiesForPicker(explorerOpportunities, selectedActivities);
      const activityOptions = activityGroups
        .map((group) => {
          const totalActivities = group.activities.length;
          const selectedCount = group.activities.filter((activity) => activity.checked).length;
          const groupChecked = totalActivities > 0 && selectedCount === totalActivities;
          const groupPartial = selectedCount > 0 && selectedCount < totalActivities;
          const children = group.activities
            .map((activity) => (
              '<label class="check-option check-option-child">' +
              '<input type="checkbox" class="flat-time-activity-check" value="' + escapeHtml(activity.label) + '"' +
              (activity.checked ? " checked" : "") +
              '>' +
              '<span>' + escapeHtml(activity.label) + "</span>" +
              "</label>"
            ))
            .join("");
          return (
            '<div class="check-tree-group">' +
            '<label class="check-option check-option-parent">' +
            '<input type="checkbox" class="flat-time-activity-group-check" data-group="' + escapeHtml(group.groupKey) + '"' +
            (groupChecked ? " checked" : "") +
            (groupPartial ? ' data-indeterminate="true"' : "") +
            '>' +
            '<span><strong>' + escapeHtml(group.groupKey) + '</strong> • ' + escapeHtml(group.groupDisplay) + "</span>" +
            "</label>" +
            '<div class="check-tree-children" data-group="' + escapeHtml(group.groupKey) + '">' + children + "</div>" +
            "</div>"
          );
        })
        .join("");
      const selectedActivitiesLabel = selectedActivities && selectedActivities.length
        ? (((selectedOpportunity && selectedOpportunity.activityLabels.length) || selectedActivities.length) <= 2
            ? selectedActivities.join(", ")
            : (selectedActivities.length + " activities selected"))
        : "Choose activities";

      let activityDrilldownHtml =
        '<div class="field" style="margin-bottom:14px; max-width:420px;">' +
        '<label for="flat-time-drilldown-activity-picker">Selected Activities</label>' +
        '<details class="check-dropdown" id="flat-time-drilldown-activity-picker"' + (flatTimeState.activityPickerOpen ? " open" : "") + '>' +
        '<summary><span class="check-dropdown-label">' + escapeHtml(selectedActivitiesLabel) + '</span><span class="check-dropdown-caret">▼</span></summary>' +
        '<div class="check-dropdown-menu">' + activityOptions + '</div>' +
        '</details>' +
        '<div class="field-help">Tick individual activities or select a whole group such as BOP. The benchmark below sums everything selected.</div>' +
        '</div>';

      if (!selectedOpportunity) {
        ui.flatTimeActivityDrilldown.innerHTML =
          activityDrilldownHtml +
          '<div class="empty">No activity is selected. Mark one or more flat time activities to build the benchmark. PRE and drilling actions D, DMR, and DMS are excluded here.</div>';
      } else {
        const peerRows = selectedOpportunity.ranked
          .slice()
          .sort((left, right) => right.value - left.value || left.label.localeCompare(right.label))
          .map((entry) => ({
            rigLabel: entry.rigLabel || "Rig not mapped",
            label: entry.label,
            value: entry.value,
            gap: Math.max(entry.value - selectedOpportunity.idealTime, 0),
          }));

        activityDrilldownHtml +=
        '<div class="metric-strip" style="margin-bottom:14px;">' +
        '<div class="metric-pill"><div class="label">Selected Activities</div><div class="value"><span class="value-main">' + escapeHtml(String(selectedOpportunity.labelCount)) + '</span><span class="value-suffix">' + escapeHtml(selectedOpportunity.labelCount === 1 ? "activity" : "activities") + '</span></div><div class="meta">' + selectedOpportunity.activityLabels.map((label) => flatTimeActivityLabelHtml(label)).join("<br>") + '</div></div>' +
        '<div class="metric-pill"><div class="label">Recommended Ideal</div><div class="value"><span class="value-main">' + escapeHtml(formatHours(selectedOpportunity.idealTime)) + '</span><span class="value-suffix">' + escapeHtml("hr / " + formatDays(selectedOpportunity.idealTime / 24) + " d") + '</span></div><div class="meta">' + escapeHtml(selectedOpportunity.idealRule) + '</div></div>' +
        '<div class="metric-pill"><div class="label">Confidence</div><div class="value">' + confidenceBadgeHtml(selectedOpportunity.confidence) + '</div><div class="meta">' + escapeHtml(selectedOpportunity.variability + " variability • sample " + selectedOpportunity.occurrenceCount) + '</div></div>' +
        '<div class="metric-pill"><div class="label">Observed Range</div><div class="value"><span class="value-main">' + escapeHtml(formatHours(selectedOpportunity.fastestTime)) + " - " + escapeHtml(formatHours(Math.max(...selectedOpportunity.values, 0))) + '</span><span class="value-suffix">' + escapeHtml("hr / " + formatDays(selectedOpportunity.fastestTime / 24) + " - " + formatDays(Math.max(...selectedOpportunity.values, 0) / 24) + " d") + '</span></div><div class="meta">Fastest to slowest observed execution</div></div>' +
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

        ui.flatTimeActivityDrilldown.innerHTML = activityDrilldownHtml;
      }

      const activityPicker = document.getElementById("flat-time-drilldown-activity-picker");
      if (activityPicker) {
        Array.from(activityPicker.querySelectorAll(".flat-time-activity-group-check[data-indeterminate='true']")).forEach((input) => {
          input.indeterminate = true;
        });
        activityPicker.addEventListener("toggle", () => {
          flatTimeState.activityPickerOpen = activityPicker.open;
        });
      }

      const collectSelectedFlatTimeActivities = () => Array.from(document.querySelectorAll(".flat-time-activity-check:checked"))
        .map((input) => input.value)
        .filter(Boolean);

      Array.from(document.querySelectorAll(".flat-time-activity-group-check")).forEach((groupCheckbox) => {
        groupCheckbox.addEventListener("change", () => {
          const groupKey = groupCheckbox.dataset.group || "";
          const childCheckboxes = Array.from(document.querySelectorAll(".check-tree-children"))
            .filter((container) => container.dataset.group === groupKey)
            .flatMap((container) => Array.from(container.querySelectorAll(".flat-time-activity-check")));
          childCheckboxes.forEach((checkbox) => {
            checkbox.checked = groupCheckbox.checked;
          });

          const selected = collectSelectedFlatTimeActivities();
          flatTimeState.activityPickerOpen = true;
          flatTimeState.focusActivities = selected;
          flatTimeState.focusActivity = selected[0] || "";
          renderFlatTime();
        });
      });

      Array.from(document.querySelectorAll(".flat-time-activity-check")).forEach((checkbox) => {
        checkbox.addEventListener("change", () => {
          const selected = collectSelectedFlatTimeActivities();
          flatTimeState.activityPickerOpen = true;
          flatTimeState.focusActivities = selected;
          flatTimeState.focusActivity = selected[0] || "";
          renderFlatTime();
        });
      });

      ui.flatTimeDrilldownNote.textContent =
        'Selected well: ' + selectedDataset.subjectWell + ' • selected activities: ' + ((selectedOpportunity && selectedOpportunity.activityLabels.length) ? selectedOpportunity.activityLabels.join(", ") : "none") + '. ' +
        'This explorer is using flat time only, excluding PRE plus drilling actions D, DMR, and DMS. ' +
        'Use the selectors in this explorer to move across wells, sections, and activity sets while reviewing the benchmark and ideal-time logic. ' +
        'Depth-based drill-down is still limited because the uploaded CSVs do not contain true depth fields such as section top/bottom, measured depth or TD.';
    }

    function renderParetoChart(target, opportunities, valueKey = "totalRecoverableHours") {
      if (!target) return;
      if (!opportunities.length) {
        target.innerHTML = '<div class="empty">No recoverable-hour opportunities available.</div>';
        return;
      }

      const items = opportunities
        .slice()
        .sort((left, right) => Number(right[valueKey] || 0) - Number(left[valueKey] || 0))
        .filter((item) => Number(item[valueKey] || 0) > 0)
        .slice(0, 10);
      if (!items.length) {
        target.innerHTML = '<div class="empty">No recoverable-hour opportunities available for the selected well.</div>';
        return;
      }
      const totalRecoverable = items.reduce((sum, item) => sum + Number(item[valueKey] || 0), 0) || 1;
      const chartTheme = getChartTheme();
      const width = Math.max(820, items.length * 120);
      const height = 360;
      const margin = { top: 24, right: 40, bottom: 86, left: 52 };
      const chartWidth = width - margin.left - margin.right;
      const chartHeight = height - margin.top - margin.bottom;
      const maxBar = niceMax(Math.max(...items.map((item) => Number(item[valueKey] || 0)), 1));
      const groupWidth = chartWidth / items.length;
      let cumulative = 0;
      const points = [];

      const grid = Array.from({ length: 5 }, (_, index) => {
        const value = (maxBar / 4) * index;
        const y = margin.top + chartHeight - (value / maxBar) * chartHeight;
        return (
          '<g>' +
          '<line x1="' + margin.left + '" y1="' + y.toFixed(2) + '" x2="' + (width - margin.right) + '" y2="' + y.toFixed(2) + '" stroke="' + chartTheme.grid + '" stroke-dasharray="4 5"></line>' +
          '<text x="' + (margin.left - 10) + '" y="' + (y + 4).toFixed(2) + '" text-anchor="end" font-size="11" fill="' + chartTheme.valueLabel + '">' + escapeHtml(formatHours(value)) + "</text>" +
          "</g>"
        );
      }).join("");

      const bars = items.map((item, index) => {
        const x = margin.left + index * groupWidth + groupWidth * 0.18;
        const barWidth = groupWidth * 0.64;
        const itemValue = Number(item[valueKey] || 0);
        const barHeight = (itemValue / maxBar) * chartHeight;
        const y = margin.top + chartHeight - barHeight;
        cumulative += itemValue;
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
          '<text x="' + pointX.toFixed(2) + '" y="' + Math.max(margin.top + 12, y - 8).toFixed(2) + '" text-anchor="middle" font-size="11" font-weight="700" fill="' + chartTheme.pointLabel + '">' + escapeHtml(formatHours(itemValue)) + "</text>" +
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
          const tooltip = modeLabel + ': ' + formatHours(value) + ' hr' + ' | Actual: ' + formatHours(actual) + ' hr | Ideal: ' + formatHours(opportunity.idealTime) + ' hr';
          return '<td title="' + escapeHtml(tooltip) + '" style="background:' + bg + '; color:' + color + '; font-weight:700; text-align:center;">' + escapeHtml(formatHours(value)) + '</td>';
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
      if (!ui.flatTimeSummary) return;
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
        { label: "Top Consuming Activity", value: topActivity ? topActivity.label : "N/A", valueHtml: topActivity ? flatTimeActivityLabelHtml(topActivity.label) : escapeHtml("N/A"), meta: topActivity ? formatHours(topActivity.total) + " hr total" : "No activity data" },
        { label: "Largest Group", value: topGroup ? topGroup.label : "N/A", meta: topGroup ? formatHours(topGroup.total) + " hr total" : "No group data" },
        { label: "Highest Burden Well", value: highestDataset ? highestDataset.label : "N/A", meta: highestDataset ? (highestDataset.rigLabel + " • " + formatHours(highestDataset.value) + " hr total") : "No dataset totals" },
        {
          label: "Best Reduction Opportunity",
          value: topSpread ? topSpread.topEntry.rigLabel : "N/A",
          meta: topSpread ? (topSpread.topEntry.label + " • " + topSpread.activityLabel + " • gap " + formatHours(topSpread.gapToIdeal) + " hr vs ideal") : "Need more than one dataset",
          metaHtml: topSpread ? (escapeHtml(topSpread.topEntry.label + " • ") + flatTimeActivityLabelHtml(topSpread.activityLabel) + escapeHtml(" • gap " + formatHours(topSpread.gapToIdeal) + " hr vs ideal")) : escapeHtml("Need more than one dataset"),
        },
        {
          label: "Total Recoverable Hours",
          value: formatHours(totalRecoverable),
          meta: "Sum of time above the recommended ideal across the comparison set",
        },
        {
          label: "Most Reliable Ideal",
          value: mostReliableIdeal ? mostReliableIdeal.activityLabel : "N/A",
          valueHtml: mostReliableIdeal ? flatTimeActivityLabelHtml(mostReliableIdeal.activityLabel) : escapeHtml("N/A"),
          meta: mostReliableIdeal ? (mostReliableIdeal.confidence + " confidence • target " + formatHours(mostReliableIdeal.idealTime) + " hr") : "Need at least 3 wells for a strong benchmark",
        },
        { label: "Total Compared Time", value: formatHours(overallHours), meta: metricKey === "subjectHours" ? "Subject well hours" : metricKey === "meanHours" ? "Mean hours" : "Median hours" },
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

    function flatTimeAiScopeLabel(scope) {
      if (scope === "selected-well-section") return "Selected well + section";
      if (scope === "selected-activities") return "Selected activities";
      if (scope === "current-scope") return "Current filtered comparison set";
      return "Selected well";
    }

    function setFlatTimeAiOutput(message, isError = false) {
      if (!ui.flatTimeAiOutput) return;
      ui.flatTimeAiOutput.classList.toggle("empty", !message);
      ui.flatTimeAiOutput.classList.toggle("is-error", isError);
      ui.flatTimeAiOutput.innerHTML = message
        ? message
        : "Generate a report to see the report summary here.";
    }

    function setFlatTimeAiBusyState(isBusy) {
      flatTimeState.aiBusy = Boolean(isBusy);
      if (!ui.flatTimeAiGenerate) return;
      ui.flatTimeAiGenerate.disabled = flatTimeState.aiBusy || !flatTimeState.aiContext;
      ui.flatTimeAiGenerate.textContent = flatTimeState.aiBusy ? "Generating..." : "Generate Report";
      if (ui.flatTimeAiExport) {
        ui.flatTimeAiExport.disabled = flatTimeState.aiBusy || !flatTimeState.aiReportText;
      }
    }

    function buildFlatTimeAiContext(
      datasets,
      selectedDataset,
      selectedWell,
      selectedRig,
      selectedSectionSize,
      selectedActivities,
      selectedOpportunity,
      metricKey,
      totalKey,
      wellRanking,
      sectionSavingsRows,
      selectedWellSectionItems,
      rigBenchmarkRows,
      opportunityPipeline,
      allOpportunities
    ) {
      const reportOpportunities = (allOpportunities || []).filter((opportunity) => {
        const activityLabel = String(opportunity?.activityLabel || "");
        const parts = extractFlatTimeComparisonParts(activityLabel, opportunity?.groupLabel);
        if (parts.phase === "PRE") return false;
        if (isFlatTimeDrillingAction(parts.action)) return false;
        return true;
      });
      const reportSelectedLabels = Array.from(
        new Set(
          (selectedActivities || []).filter((label) =>
            reportOpportunities.some((opportunity) => opportunity.activityLabel === label)
          )
        )
      );
      const reportSelectedOpportunity = reportSelectedLabels.length
        ? buildSelectedActivityAggregate(reportOpportunities, reportSelectedLabels)
        : null;
      const reportWellRanking = buildWellRanking(datasets, reportOpportunities);
      const selectedWellRanking = reportWellRanking.find((row) => row.wellLabel === selectedWell) || null;
      const reportBreakdownMap = buildSectionBreakdownMap(datasets, reportOpportunities, metricKey);
      const reportSelectedWellSectionItems = buildSelectedWellSectionItems(datasets, selectedDataset, reportBreakdownMap);
      const reportSectionSavingsRows = buildSectionSavingsSummary(datasets, reportBreakdownMap);
      const comparisonSet = datasets.map((dataset) => {
        const ranking = reportWellRanking.find((row) => row.wellLabel === dataset.subjectWell);
        return {
          rig: dataset.rigLabel || "Rig not mapped",
          well: dataset.subjectWell,
          totalHours: Number(ranking ? ranking.actualTotal : 0),
        };
      });
      const summaryOpportunities = reportOpportunities
        .slice(0, 8)
        .map((item) => ({
          section: formatFlatTimeSectionSize(item.sectionSize),
          group: item.groupLabel,
          activity: item.activityLabel,
          idealHours: Number(item.idealTime || 0),
          recoverableHours: Number(item.totalRecoverableHours || 0),
          confidence: item.confidence,
          highestWell: item.topEntry.label || "N/A",
          highestRig: item.topEntry.rigLabel || "Rig not mapped",
        }));

      return {
        sourceFile: dashboardData.meta.sourceFile || "Unknown source",
        generatedAt: dashboardData.meta.generatedAt || "",
        reportAgent: dashboardData.meta.weeklyReportAgent || "WeeklyReport",
        currentSelection: {
          rig: selectedRig || "All rigs",
          sectionSize: selectedSectionSize ? formatFlatTimeSectionSize(selectedSectionSize) : "All section sizes",
          metric: metricKey === "subjectHours" ? "Subject Well Time" : metricKey === "meanHours" ? "Mean Time" : "Median Time",
          selectedWell: selectedWell || "",
          selectedActivities: reportSelectedLabels,
          analysisMode: ui.flatTimeMode ? ui.flatTimeMode.value : "executive",
        },
        comparisonSet: {
          datasetsLoaded: datasets.length,
          wells: comparisonSet,
        },
        summary: {
          totalComparedHours: comparisonSet.reduce((sum, item) => sum + Number(item.totalHours || 0), 0),
          totalRecoverableHours: reportSectionSavingsRows.reduce((sum, item) => sum + Number(item.recoverableHours || 0), 0),
          highestBurdenWell: selectedWellRanking
            ? {
                rig: selectedWellRanking.rigLabel,
                well: selectedWellRanking.wellLabel,
                actualHours: selectedWellRanking.actualTotal,
                idealHours: selectedWellRanking.idealTotal,
                excessHours: selectedWellRanking.excessTotal,
              }
            : null,
        },
        selectedWell: selectedDataset
          ? {
              rig: selectedDataset.rigLabel || "Rig not mapped",
              well: selectedDataset.subjectWell,
              actualHours: selectedWellRanking ? selectedWellRanking.actualTotal : Number(selectedDataset[totalKey] || 0),
              idealHours: selectedWellRanking ? selectedWellRanking.idealTotal : 0,
              excessHours: selectedWellRanking ? selectedWellRanking.excessTotal : 0,
              sections: reportSelectedWellSectionItems.map((item) => ({
                section: item.label,
                actualHours: item.actualTotal,
                idealHours: item.idealTotal,
                gapHours: item.gapTotal,
                offsetAverageHours: item.offsetAverage,
                bestOffsetHours: item.bestOffset,
                bestOffsetWell: item.bestOffsetWell,
                bestOffsetRig: item.bestOffsetRig,
                topDrivers: (item.topGapActivities && item.topGapActivities.length ? item.topGapActivities : item.topActivities).map((activity) => ({
                  activity: activity.activityLabel,
                  actualHours: activity.actual,
                  idealHours: activity.ideal,
                  gapHours: activity.gap,
                  confidence: activity.confidence,
                })).filter((activity) => Number(activity.gapHours || 0) > 0 || !item.topGapActivities || !item.topGapActivities.length),
              })),
            }
          : null,
        selectedActivities: reportSelectedOpportunity
          ? {
              labels: reportSelectedOpportunity.activityLabels,
              section: reportSelectedOpportunity.sectionLabel,
              group: reportSelectedOpportunity.groupLabel,
              selectedWellHours: Number(reportSelectedOpportunity.ranked.find((entry) => entry.label === selectedWell)?.value || 0),
              idealHours: reportSelectedOpportunity.idealTime,
              gapHours: Math.max(Number(reportSelectedOpportunity.ranked.find((entry) => entry.label === selectedWell)?.value || 0) - Number(reportSelectedOpportunity.idealTime || 0), 0),
              peerAverageHours: reportSelectedOpportunity.peerAverage,
              occurrenceCount: reportSelectedOpportunity.occurrenceCount,
              confidence: reportSelectedOpportunity.confidence,
              highestWell: reportSelectedOpportunity.topEntry.label,
              highestRig: reportSelectedOpportunity.topEntry.rigLabel,
              highestObservedHours: reportSelectedOpportunity.topEntry.value,
            }
          : null,
          topWells: reportWellRanking.slice(0, 5).map((row) => ({
          rig: row.rigLabel,
          well: row.wellLabel,
          actualHours: row.actualTotal,
          idealHours: row.idealTotal,
          excessHours: row.excessTotal,
          topDrivers: row.topDrivers.map((driver) => ({
            activity: driver.activity,
            group: driver.group,
            gapHours: driver.gap,
          })),
          otherDriversGapHours: row.otherDriversGap,
        })),
        savingsBySection: reportSectionSavingsRows.slice(0, 6).map((row) => ({
          section: row.sectionLabel,
          recoverableHours: row.recoverableHours,
          recoverableDays: row.recoverableDays,
          impactedWells: row.impactedWells,
          topActivity: row.topActivity,
          confidence: row.confidence,
          recommendation: row.recommendation,
        })),
        rigBenchmark: rigBenchmarkRows.slice(0, 6).map((row) => ({
          rig: row.rigLabel,
          wells: row.wellCount,
          averageFlatTimeHours: row.averageFlatTime,
          idealFlatTimeHours: row.averageIdealTime,
          excessHours: row.excessTime,
          mainRepeatingActivity: row.mainRepeatingActivity,
        })),
        opportunityPipeline: opportunityPipeline.slice(0, 8).map((row) => ({
          activity: row.activityLabel,
          group: row.groupLabel,
          section: formatFlatTimeSectionSize(row.sectionSize),
          occurrences: row.occurrenceCount,
          wellsImpacted: row.wellsImpacted,
          idealHours: row.idealTime,
          recoverableHours: row.totalRecoverableHours,
          priority: row.priority,
        })),
        topActivities: summaryOpportunities,
      };
    }

    function flatTimeHoursWithDaysNarrative(hours) {
      return formatHours(hours) + " hr (" + formatDays((hours || 0) / 24) + " d)";
    }

    function flatTimeSentenceList(items) {
      const cleaned = (items || []).filter(Boolean);
      if (!cleaned.length) return "";
      if (cleaned.length === 1) return cleaned[0];
      if (cleaned.length === 2) return cleaned[0] + " and " + cleaned[1];
      return cleaned.slice(0, -1).join(", ") + ", and " + cleaned[cleaned.length - 1];
    }

    function buildFlatTimeScopeInsights(aiContext, scope) {
      const selectedWell = aiContext.selectedWell || null;
      const selectedActivities = aiContext.selectedActivities || null;
      const selectedWellSections = selectedWell && Array.isArray(selectedWell.sections) ? selectedWell.sections.slice() : [];
      const selectedSection = selectedWellSections.find((section) => section.section === aiContext.currentSelection.sectionSize) || null;
      const topSelectedSection = selectedWellSections
        .slice()
        .sort((left, right) => Number(right.gapHours || 0) - Number(left.gapHours || 0) || Number(right.actualHours || 0) - Number(left.actualHours || 0))[0] || null;
      const topSelectedActivity = selectedWellSections
        .flatMap((section) => (section.topDrivers || []).map((driver) => ({ ...driver, section: section.section })))
        .sort((left, right) => Number(right.gapHours || 0) - Number(left.gapHours || 0) || Number(right.actualHours || 0) - Number(left.actualHours || 0))[0] || null;
      const topScopeSection = (aiContext.savingsBySection || [])[0] || null;
      const topScopeActivity = (aiContext.topActivities || [])[0] || null;

      if (scope === "selected-activities" && selectedActivities) {
        return {
          recoverableHours: Number(selectedActivities.gapHours || 0),
          recoverableMeta: flatTimeHoursWithDaysNarrative(selectedActivities.gapHours || 0) + " for the selected activity set in the selected well",
          topSectionLabel: selectedActivities.section || "N/A",
          topSectionHours: Number(selectedActivities.gapHours || 0),
          topSectionMeta: selectedActivities.section
            ? "Selected activity set belongs to " + selectedActivities.section + " and carries " + flatTimeHoursWithDaysNarrative(selectedActivities.gapHours || 0) + " of excess time."
            : "No section mapping was found for the selected activities.",
          topActivityLabel: selectedActivities.labels && selectedActivities.labels.length === 1
            ? selectedActivities.labels[0]
            : String((selectedActivities.labels || []).length) + " selected activities",
          topActivityValueHtml: selectedActivities.labels && selectedActivities.labels.length === 1
            ? flatTimeActivityLabelHtml(selectedActivities.labels[0])
            : escapeHtml(String((selectedActivities.labels || []).length) + " selected activities"),
          topActivityHours: Number(selectedActivities.gapHours || 0),
          topActivityMeta: flatTimeHoursWithDaysNarrative(selectedActivities.gapHours || 0) + " above ideal • " + (selectedActivities.group || "Combined job"),
          headlineSection: selectedActivities.section || "",
          headlineActivity: selectedActivities.labels && selectedActivities.labels.length ? selectedActivities.labels[0] : "",
        };
      }

      if (scope === "selected-well-section" && selectedSection) {
        const topSectionDriver = (selectedSection.topDrivers || [])
          .slice()
          .sort((left, right) => Number(right.gapHours || 0) - Number(left.gapHours || 0))[0] || null;
        return {
          recoverableHours: Number(selectedSection.gapHours || 0),
          recoverableMeta: flatTimeHoursWithDaysNarrative(selectedSection.gapHours || 0) + " in the selected section of the selected well",
          topSectionLabel: selectedSection.section,
          topSectionHours: Number(selectedSection.gapHours || 0),
          topSectionMeta: flatTimeHoursWithDaysNarrative(selectedSection.actualHours || 0) + " actual vs " + flatTimeHoursWithDaysNarrative(selectedSection.idealHours || 0) + " ideal",
          topActivityLabel: topSectionDriver ? topSectionDriver.activity : "N/A",
          topActivityValueHtml: topSectionDriver ? flatTimeActivityLabelHtml(topSectionDriver.activity) : escapeHtml("N/A"),
          topActivityHours: Number(topSectionDriver ? topSectionDriver.gapHours || 0 : 0),
          topActivityMeta: topSectionDriver
            ? flatTimeHoursWithDaysNarrative(topSectionDriver.gapHours || 0) + " above ideal • " + selectedSection.section
            : "No excess activity found in this section",
          headlineSection: selectedSection.section,
          headlineActivity: topSectionDriver ? topSectionDriver.activity : "",
        };
      }

      if (scope === "selected-well" && selectedWell) {
        return {
          recoverableHours: Number(selectedWell.excessHours || 0),
          recoverableMeta: flatTimeHoursWithDaysNarrative(selectedWell.excessHours || 0) + " across the selected well",
          topSectionLabel: topSelectedSection ? topSelectedSection.section : "N/A",
          topSectionHours: Number(topSelectedSection ? topSelectedSection.gapHours || 0 : 0),
          topSectionMeta: topSelectedSection
            ? flatTimeHoursWithDaysNarrative(topSelectedSection.gapHours || 0) + " recoverable in the selected well"
            : "No section gap found in the selected well",
          topActivityLabel: topSelectedActivity ? topSelectedActivity.activity : "N/A",
          topActivityValueHtml: topSelectedActivity ? flatTimeActivityLabelHtml(topSelectedActivity.activity) : escapeHtml("N/A"),
          topActivityHours: Number(topSelectedActivity ? topSelectedActivity.gapHours || 0 : 0),
          topActivityMeta: topSelectedActivity
            ? flatTimeHoursWithDaysNarrative(topSelectedActivity.gapHours || 0) + " above ideal • " + topSelectedActivity.section
            : "No recurring activity found in the selected well",
          headlineSection: topSelectedSection ? topSelectedSection.section : "",
          headlineActivity: topSelectedActivity ? topSelectedActivity.activity : "",
        };
      }

      return {
        recoverableHours: Number(aiContext.summary.totalRecoverableHours || 0),
        recoverableMeta: flatTimeHoursWithDaysNarrative(aiContext.summary.totalRecoverableHours || 0) + " across current scope",
        topSectionLabel: topScopeSection ? topScopeSection.section : "N/A",
        topSectionHours: Number(topScopeSection ? topScopeSection.recoverableHours || 0 : 0),
        topSectionMeta: topScopeSection
          ? flatTimeHoursWithDaysNarrative(topScopeSection.recoverableHours || 0) + " across " + topScopeSection.impactedWells + " well(s)"
          : "No section savings identified",
        topActivityLabel: topScopeActivity ? topScopeActivity.activity : "N/A",
        topActivityValueHtml: topScopeActivity ? flatTimeActivityLabelHtml(topScopeActivity.activity) : escapeHtml("N/A"),
        topActivityHours: Number(topScopeActivity ? topScopeActivity.recoverableHours || 0 : 0),
        topActivityMeta: topScopeActivity
          ? flatTimeHoursWithDaysNarrative(topScopeActivity.recoverableHours || 0) + " recoverable • " + topScopeActivity.group
          : "No recurring activity found",
        headlineSection: topScopeSection ? topScopeSection.section : "",
        headlineActivity: topScopeActivity ? topScopeActivity.activity : "",
      };
    }

    function buildFlatTimeStandardReport(aiContext, scope) {
      if (!aiContext) return "";

      const scopeLabel = flatTimeAiScopeLabel(scope);
      const selectedWell = aiContext.selectedWell;
      const selectedActivities = scope === "selected-activities" ? aiContext.selectedActivities : null;
      const scopeInsights = buildFlatTimeScopeInsights(aiContext, scope);
      const topSection = scopeInsights.headlineSection
        ? ((selectedWell && selectedWell.sections || []).find((section) => section.section === scopeInsights.headlineSection) || (aiContext.savingsBySection || []).find((section) => section.section === scopeInsights.headlineSection) || null)
        : null;
      const topActivity = scopeInsights.headlineActivity
        ? ((aiContext.topActivities || []).find((item) => item.activity === scopeInsights.headlineActivity) || null)
        : null;
      const topSectionRecoverableHours = Number(
        scopeInsights.topSectionHours ||
        topSection?.gapHours ||
        topSection?.recoverableHours ||
        0
      );
      const topActivityRecoverableHours = Number(
        scopeInsights.topActivityHours ||
        selectedActivities?.gapHours ||
        topActivity?.recoverableHours ||
        0
      );
      const sectionDrivers = selectedWell && Array.isArray(selectedWell.sections)
        ? selectedWell.sections
            .slice()
            .sort((left, right) => Number(right.gapHours || 0) - Number(left.gapHours || 0))
            .slice(0, 3)
        : [];
      const strongestSections = sectionDrivers.map((section) =>
        section.section + " at " + flatTimeHoursWithDaysNarrative(section.actualHours) +
        " versus " + flatTimeHoursWithDaysNarrative(section.idealHours) +
        ", leaving a gap of " + flatTimeHoursWithDaysNarrative(section.gapHours)
      );
      const topDriverBullets = sectionDrivers.flatMap((section) =>
        (section.topDrivers || []).slice(0, 2).map((driver) =>
          driver.activity + " in the " + section.section + " section contributed " +
          flatTimeHoursWithDaysNarrative(driver.gapHours || 0) +
          " above the recommended benchmark."
        )
      ).slice(0, 5);
      const sectionBullets = sectionDrivers.map((section) => {
        const leadDriver = (section.topDrivers || [])[0];
        return section.section + " offers " + flatTimeHoursWithDaysNarrative(section.gapHours || 0) +
          " of recoverable time in the selected well" +
          (leadDriver ? ", driven mainly by " + leadDriver.activity + "." : ".");
      });
      const lines = [
        "## 1. Executive Summary",
        "This report evaluates **flat time only** and explicitly excludes the **PRE** phase plus drilling actions **D, DMR, and DMS**.",
        selectedWell
          ? "The selected well, **" + selectedWell.well + "** on **" + selectedWell.rig + "**, accumulated **" +
            flatTimeHoursWithDaysNarrative(selectedWell.actualHours) + "** against a recommended ideal of **" +
            flatTimeHoursWithDaysNarrative(selectedWell.idealHours) + "**. This leaves a current excess of **" +
            flatTimeHoursWithDaysNarrative(selectedWell.excessHours) + "** within the selected scope."
          : "The current filtered comparison set contains **" + aiContext.comparisonSet.datasetsLoaded +
            "** loaded well(s) with **" + flatTimeHoursWithDaysNarrative(aiContext.summary.totalComparedHours || 0) +
            "** of compared flat time and **" + flatTimeHoursWithDaysNarrative(aiContext.summary.totalRecoverableHours || 0) +
            "** of recoverable opportunity.",
        scopeInsights.headlineSection
          ? "The current scope highlights **" + scopeInsights.headlineSection + "** as the main section focus, with **" +
            flatTimeHoursWithDaysNarrative(scopeInsights.topSectionHours || 0) + "** of recoverable time tied to the selected job."
          : "No dominant recovery section was identified in the current scope.",
        "",
        "## 2. Selected Job / Scope",
        "- **Scope:** " + scopeLabel,
        "- **Current rig filter:** " + aiContext.currentSelection.rig,
        "- **Current section filter:** " + aiContext.currentSelection.sectionSize,
        "- **Flat time basis:** PRE, D, DMR, and DMS are excluded from all report calculations.",
        "- **Selected well:** " + (aiContext.currentSelection.selectedWell || "Auto-selected from the current scope"),
        selectedActivities && selectedActivities.labels && selectedActivities.labels.length
          ? "- **Selected activities:** " + selectedActivities.labels.join(", ")
          : "- **Selected activities:** No individual activity set was forced.",
        "",
        "## 3. Main Findings",
        selectedActivities && selectedActivities.labels && selectedActivities.labels.length
          ? "The selected activity set combines **" + selectedActivities.labels.length + "** activity label(s). In the selected well, those activities total **" +
            flatTimeHoursWithDaysNarrative(selectedActivities.selectedWellHours || 0) + "** against a combined ideal of **" +
            flatTimeHoursWithDaysNarrative(selectedActivities.idealHours || 0) + "**, leaving **" +
            flatTimeHoursWithDaysNarrative(selectedActivities.gapHours || 0) + "** of excess time."
          : "The section-by-section review points to a concentrated loss profile rather than a uniform performance issue across the full well.",
        ...(strongestSections.length ? strongestSections.map((item) => "- " + item) : ["- No section-specific gap was found in the current scope."]),
        ...(topDriverBullets.length ? topDriverBullets.map((item) => "- " + item) : []),
        "",
        "## 4. Time Recovery Opportunities",
        topSection
          ? "The highest value opportunities come from repeating the best validated section performance while removing the main repeating delays in the selected well."
          : "The current scope does not show a single dominant section, so the opportunity should be managed through the combined activity set.",
        ...(sectionBullets.length ? sectionBullets.map((item) => "- " + item) : ["- No recoverable section time was identified from the current scope."]),
        topActivity
          ? "- The strongest activity-level opportunity is **" + topActivity.activity + "** in **" + topActivity.section +
            "**, where the selected well still carries **" + flatTimeHoursWithDaysNarrative(topActivityRecoverableHours) +
            "** of recoverable time."
          : "",
      ];

      return lines
        .filter((line) => line !== null && line !== undefined)
        .join("\n");
    }

    function renderFlatTimeAiSummary(aiContext) {
      if (!ui.flatTimeAiSummary) return;
      if (!aiContext) {
        ui.flatTimeAiSummary.innerHTML = '<div class="empty">Report summary metrics will appear here after Flat Time context is available.</div>';
        return;
      }

      const scopeInsights = buildFlatTimeScopeInsights(aiContext, ui.flatTimeAiScope ? ui.flatTimeAiScope.value : "selected-well");
      const selectedWell = aiContext.selectedWell;
      const cards = [
        {
          label: "Report Scope",
          value: flatTimeAiScopeLabel(ui.flatTimeAiScope ? ui.flatTimeAiScope.value : "selected-well"),
          meta: aiContext.currentSelection.selectedWell || "Current filtered comparison set",
        },
        {
          label: "Selected Well",
          value: selectedWell ? selectedWell.well : "N/A",
          meta: selectedWell ? (selectedWell.rig + " • " + formatHours(selectedWell.actualHours) + " hr flat time actual") : "No well in scope",
        },
        {
          label: "Recoverable Time",
          value: formatHours(scopeInsights.recoverableHours || 0) + " hr",
          meta: scopeInsights.recoverableMeta,
        },
        {
          label: "Top Section",
          value: scopeInsights.topSectionLabel || "N/A",
          meta: scopeInsights.topSectionMeta,
        },
        {
          label: "Top Activity",
          value: scopeInsights.topActivityLabel || "N/A",
          valueHtml: scopeInsights.topActivityValueHtml || escapeHtml(scopeInsights.topActivityLabel || "N/A"),
          meta: scopeInsights.topActivityMeta,
        },
      ];

      ui.flatTimeAiSummary.innerHTML = cards
        .map((card) => (
          '<div class="metric-pill">' +
          '<div class="label">' + escapeHtml(card.label) + "</div>" +
          '<div class="value"><span class="value-main">' + (card.valueHtml || escapeHtml(card.value)) + '</span></div>' +
          '<div class="meta">' + escapeHtml(card.meta) + "</div>" +
          "</div>"
        ))
        .join("");
    }

    function renderFlatTimeAiChart(aiContext) {
      if (!ui.flatTimeAiChart) return;
      if (!aiContext) {
        ui.flatTimeAiChart.innerHTML = '<div class="empty">Snapshot chart will appear here after Flat Time context is available.</div>';
        return;
      }

      if (aiContext.selectedWell && Array.isArray(aiContext.selectedWell.sections) && aiContext.selectedWell.sections.length) {
        renderMultiSeriesChart(
          ui.flatTimeAiChart,
          aiContext.selectedWell.sections.map((section) => ({
            label: section.section,
            actual: Number(section.actualHours || 0),
            ideal: Number(section.idealHours || 0),
            gap: Number(section.gapHours || 0),
          })),
          [
            { key: "actual", label: "Actual", color: "#1264d6", format: (value) => formatHours(value) },
            { key: "ideal", label: "Ideal", color: "#0f766e", format: (value) => formatHours(value) },
            { key: "gap", label: "Gap", color: "#c06a0a", format: (value) => formatHours(value) },
          ],
          { height: 360, minWidth: 760, groupMinWidth: 150 }
        );
        return;
      }

      if (aiContext.savingsBySection && aiContext.savingsBySection.length) {
        renderBarChart(
          ui.flatTimeAiChart,
          aiContext.savingsBySection.map((row) => ({
            label: row.section,
            value: Number(row.recoverableHours || 0),
          })),
          "#c81e5a",
          (value) => formatHours(value) + " hr",
          { maxItems: 0 }
        );
        return;
      }

      ui.flatTimeAiChart.innerHTML = '<div class="empty">Not enough scoped data to draw the snapshot chart.</div>';
    }

    function renderFlatTimeAiMarkdown(markdown) {
      const source = String(markdown || "").replace(/\\n/g, "\n").trim();
      if (!source) return "";

      const blocks = source
        .split(/\n\s*\n/)
        .map((block) => block.trim())
        .filter(Boolean);

      function inlineFormat(text) {
        let escaped = escapeHtml(text);
        escaped = escaped.replace(/\*\*(.+?)\*\*/g, "<strong>$1</strong>");
        return escaped;
      }

      function buildBodyHtml(lines) {
        const html = [];
        let paragraphLines = [];
        let bulletLines = [];

        function flushParagraph() {
          if (!paragraphLines.length) return;
          html.push('<p>' + paragraphLines.map((line) => inlineFormat(line)).join("<br>") + "</p>");
          paragraphLines = [];
        }

        function flushBullets() {
          if (!bulletLines.length) return;
          html.push("<ul>" + bulletLines.map((line) => "<li>" + inlineFormat(line) + "</li>").join("") + "</ul>");
          bulletLines = [];
        }

        lines.forEach((line) => {
          if (/^(\*\s|-\s)/.test(line)) {
            flushParagraph();
            bulletLines.push(line.replace(/^(\*\s|-\s)/, ""));
            return;
          }
          flushBullets();
          paragraphLines.push(line);
        });

        flushParagraph();
        flushBullets();
        return html.join("");
      }

      return blocks.map((block) => {
        const lines = block.split("\n").map((line) => line.trim()).filter(Boolean);
        if (!lines.length) return "";

        const first = lines[0];
        if (/^(\*\s|-\s)/.test(first) || lines.every((line) => /^(\*\s|-\s)/.test(line))) {
          return '<div class="ai-report-section"><ul>' +
            lines.map((line) => '<li>' + inlineFormat(line.replace(/^(\*\s|-\s)/, "")) + "</li>").join("") +
            "</ul></div>";
        }

        if (/^#{1,6}\s+/.test(first) || /^\*\*.+\*\*$/.test(first) || /^\d+\.\s+/.test(first) || /^[A-Z][A-Za-z\s/&-]+:$/.test(first)) {
          const heading = inlineFormat(first.replace(/^#{1,6}\s+/, "").replace(/^\d+\.\s+/, "").replace(/:$/, ""));
          const bodyLines = lines.slice(1);
          const bodyHtml = bodyLines.length ? buildBodyHtml(bodyLines) : "";
          return '<div class="ai-report-section"><h4>' + heading + "</h4>" + bodyHtml + "</div>";
        }

        return '<div class="ai-report-section">' + buildBodyHtml(lines) + "</div>";
      }).join("");
    }

    function exportFlatTimeAiPdf() {
      if (!flatTimeState.aiReportText || !ui.flatTimeAiOutput || !ui.flatTimeAiChart) {
        setFlatTimeAiOutput("Generate the report first, then export it as PDF.", true);
        return;
      }

      const title = dashboardData.meta.weeklyReportAgent || "WeeklyReport";
      const scopeLabel = flatTimeAiScopeLabel(ui.flatTimeAiScope ? ui.flatTimeAiScope.value : "selected-well");
      const chartHtml = ui.flatTimeAiChart.innerHTML || '<div>No chart available.</div>';
      const summaryHtml = ui.flatTimeAiSummary ? ui.flatTimeAiSummary.innerHTML : "";
      const contextHtml = ui.flatTimeAiContext ? ui.flatTimeAiContext.innerHTML : "";
      const reportHtml = ui.flatTimeAiOutput.innerHTML || "";
      const sourceFile = dashboardData.meta.sourceFile || "Unknown source";
      const generatedAt = dashboardData.meta.generatedAt || "";
      const specificWork = ui.flatTimeAiWork && ui.flatTimeAiWork.value.trim()
        ? ui.flatTimeAiWork.value.trim()
        : "Not provided";

      const printHtml = [
        "<!DOCTYPE html>",
        '<html lang="en"><head><meta charset="utf-8"><meta name="viewport" content="width=device-width, initial-scale=1">',
        "<title>WeeklyReport Report PDF</title>",
        "<style>",
        "body{margin:0;font-family:'Avenir Next','Segoe UI',sans-serif;color:#102033;background:#f4f7fb;}",
        ".page{width:min(1100px,calc(100vw - 28px));margin:0 auto;padding:28px 0 40px;}",
        ".hero{padding:24px 28px;border-radius:28px;background:linear-gradient(135deg,#102033,#1264d6 60%,#0f766e);color:#fff;}",
        ".hero h1{margin:0 0 10px;font-size:34px;letter-spacing:-0.03em;}",
        ".hero p{margin:6px 0;color:rgba(255,255,255,0.88);line-height:1.55;}",
        ".chips{display:flex;flex-wrap:wrap;gap:8px;margin-top:14px;}",
        ".chip{padding:7px 11px;border-radius:999px;background:rgba(255,255,255,0.14);border:1px solid rgba(255,255,255,0.16);color:#fff;font-size:13px;}",
        ".panel{margin-top:18px;background:#fff;border:1px solid #d8e2ef;border-radius:24px;padding:22px;box-shadow:0 16px 32px rgba(15,23,42,0.08);}",
        ".panel h2{margin:0 0 10px;font-size:23px;letter-spacing:-0.02em;}",
        ".panel p.note{margin:0 0 14px;color:#546579;line-height:1.55;}",
        ".metric-strip{display:grid;grid-template-columns:repeat(auto-fit,minmax(190px,1fr));gap:14px;}",
        ".metric-pill{padding:16px;border-radius:18px;background:linear-gradient(180deg,#ffffff,#f8fbff);border:1px solid #d8e2ef;}",
        ".metric-pill .label{font-size:12px;text-transform:uppercase;letter-spacing:0.08em;color:#607085;margin-bottom:8px;}",
        ".metric-pill .value{font-size:24px;font-weight:800;line-height:1.08;color:#102033;margin-bottom:8px;}",
        ".metric-pill .meta{font-size:13px;color:#546579;line-height:1.45;}",
        ".chart-box{padding:14px;border-radius:18px;background:linear-gradient(180deg,#ffffff,#f8fbff);border:1px solid #d8e2ef;overflow:hidden;}",
        ".chart-box svg{width:100%;height:auto;display:block;}",
        ".ai-report-section{padding:14px 16px;border-radius:16px;border:1px solid rgba(18,100,214,0.10);background:rgba(18,100,214,0.03);margin-bottom:12px;}",
        ".ai-report-section:last-child{margin-bottom:0;}",
        ".ai-report-section h4{margin:0 0 10px;font-size:17px;color:#102033;}",
        ".ai-report-section p{margin:0 0 12px;line-height:1.65;color:#102033;}",
        ".ai-report-section ul{margin:0 0 14px;padding-left:22px;}",
        ".ai-report-section li{margin:0 0 8px;line-height:1.6;color:#102033;}",
        ".ai-report-section strong{color:#102033;}",
        "@page{size:A4 portrait;margin:12mm;}",
        "@media print{body{background:#fff}.page{width:100%;padding:0}.panel,.hero,.metric-pill,.chart-box,.ai-report-section{box-shadow:none !important;break-inside:avoid;page-break-inside:avoid;}}",
        "</style></head><body>",
        '<div class="page">',
        '<section class="hero">' +
          '<h1>' + escapeHtml(title) + " Report</h1>" +
          '<p><strong>Scope:</strong> ' + escapeHtml(scopeLabel) + "</p>" +
          '<p><strong>Specific Work:</strong> ' + escapeHtml(specificWork) + "</p>" +
          '<p><strong>Source File:</strong> ' + escapeHtml(sourceFile) + "<br><strong>Generated At:</strong> " + escapeHtml(generatedAt) + "</p>" +
          '<div class="chips">' + contextHtml + "</div>" +
        "</section>",
        '<section class="panel"><h2>Opportunity Summary</h2><p class="note">Key metrics and context used to generate this report.</p><div class="metric-strip">' + summaryHtml + "</div></section>",
        '<section class="panel"><h2>Opportunity Snapshot</h2><p class="note">Visual comparison from the same Flat Time scope used by the report summary.</p><div class="chart-box">' + chartHtml + "</div></section>",
        '<section class="panel"><h2>Written Report</h2><p class="note">English narrative generated from the scoped Flat Time benchmark context.</p>' + reportHtml + "</section>",
        "</div>",
        "<script>window.onload=function(){setTimeout(function(){window.print();},120);};<\\/script>",
        "</body></html>",
      ].join("");

      const existingFrame = document.getElementById("flat-time-ai-print-frame");
      if (existingFrame) {
        existingFrame.remove();
      }

      const printFrame = document.createElement("iframe");
      printFrame.id = "flat-time-ai-print-frame";
      printFrame.setAttribute("aria-hidden", "true");
      printFrame.style.position = "fixed";
      printFrame.style.right = "0";
      printFrame.style.bottom = "0";
      printFrame.style.width = "0";
      printFrame.style.height = "0";
      printFrame.style.border = "0";
      printFrame.style.opacity = "0";
      printFrame.style.pointerEvents = "none";

      const cleanup = () => {
        window.setTimeout(() => {
          if (printFrame.parentNode) {
            printFrame.remove();
          }
        }, 1200);
      };

      printFrame.onload = () => {
        try {
          const frameWindow = printFrame.contentWindow;
          if (!frameWindow) {
            setFlatTimeAiOutput("Could not prepare the PDF preview. Please try again.", true);
            cleanup();
            return;
          }
          frameWindow.focus();
          if ("onafterprint" in frameWindow) {
            frameWindow.onafterprint = cleanup;
          } else {
            cleanup();
          }
          window.setTimeout(() => {
            try {
              frameWindow.print();
            } catch (error) {
              setFlatTimeAiOutput("Could not start the PDF export. Please try again.", true);
              cleanup();
            }
          }, 160);
        } catch (error) {
          setFlatTimeAiOutput("Could not prepare the PDF preview. Please try again.", true);
          cleanup();
        }
      };

      document.body.appendChild(printFrame);
      const frameDocument = printFrame.contentDocument || (printFrame.contentWindow && printFrame.contentWindow.document);
      if (!frameDocument) {
        setFlatTimeAiOutput("Could not prepare the PDF document. Please try again.", true);
        cleanup();
        return;
      }
      frameDocument.open();
      frameDocument.write(printHtml);
      frameDocument.close();
    }

    function renderFlatTimeAiPanel(aiContext) {
      if (!ui.flatTimeAiHeading || !ui.flatTimeAiStatus || !ui.flatTimeAiContext) return;

      ui.flatTimeAiHeading.textContent = "Report Analysis";
      ui.flatTimeAiStatus.textContent = "Standard report";
      if (ui.flatTimeAiNote) {
        ui.flatTimeAiNote.textContent = "";
      }

      if (!aiContext) {
        ui.flatTimeAiContext.innerHTML = '<span class="chip">No Flat Time benchmark context available yet.</span>';
        renderFlatTimeAiSummary(null);
        renderFlatTimeAiChart(null);
        if (!flatTimeState.aiReportText) {
          setFlatTimeAiOutput("Load Flat Time CSV files to generate a report summary for the selected job or benchmark scope.");
        }
        setFlatTimeAiBusyState(false);
        return;
      }

      const chips = [
        "Scope: " + flatTimeAiScopeLabel(ui.flatTimeAiScope ? ui.flatTimeAiScope.value : "selected-well"),
        "Selected well: " + (aiContext.currentSelection.selectedWell || "Auto"),
        "Section: " + aiContext.currentSelection.sectionSize,
        "Benchmarks: " + aiContext.comparisonSet.datasetsLoaded,
      ];
      if ((ui.flatTimeAiScope ? ui.flatTimeAiScope.value : "selected-well") === "selected-activities" && aiContext.selectedActivities && aiContext.selectedActivities.labels.length) {
        chips.push("Activities: " + aiContext.selectedActivities.labels.length);
      }
      if (ui.flatTimeAiWork && ui.flatTimeAiWork.value.trim()) {
        chips.push("Focus: " + ui.flatTimeAiWork.value.trim());
      }
      ui.flatTimeAiContext.innerHTML = chips.map((item) => '<span class="chip">' + escapeHtml(item) + '</span>').join("");
      renderFlatTimeAiSummary(aiContext);
      renderFlatTimeAiChart(aiContext);

      if (flatTimeState.aiReportText) {
        setFlatTimeAiOutput(renderFlatTimeAiMarkdown(flatTimeState.aiReportText), false);
      } else {
        setFlatTimeAiOutput("Click Generate Report to produce a narrative for the current Flat Time scope.");
      }
      setFlatTimeAiBusyState(false);
    }

    async function requestFlatTimeAdvancedReport() {
      if (!flatTimeState.aiContext) {
        setFlatTimeAiOutput("Flat Time context is not ready yet. Load CSV files and try again.", true);
        return;
      }

      flatTimeState.aiReportText = "";
      setFlatTimeAiBusyState(true);
      setFlatTimeAiOutput("Generating the report...", false);

      try {
        const response = await fetch("/ai/flat-time-report", {
          method: "POST",
          headers: { "Content-Type": "application/json" },
          body: JSON.stringify({
            scope: ui.flatTimeAiScope ? ui.flatTimeAiScope.value : "selected-well",
            specificWork: ui.flatTimeAiWork ? ui.flatTimeAiWork.value.trim() : "",
            context: flatTimeState.aiContext,
          }),
        });
        const payload = await response.json();
        if (!response.ok || !payload.ok) {
          throw new Error(payload.error || "WeeklyReport could not generate the report right now.");
        }
        flatTimeState.aiReportText = String(payload.report || "").trim();
        setFlatTimeAiOutput(
          flatTimeState.aiReportText
            ? renderFlatTimeAiMarkdown(flatTimeState.aiReportText)
            : "WeeklyReport returned an empty response.",
          !flatTimeState.aiReportText
        );
      } catch (error) {
        const message = error && error.message ? error.message : "WeeklyReport could not generate the report right now.";
        setFlatTimeAiOutput(message, true);
      } finally {
        setFlatTimeAiBusyState(false);
      }
    }

    async function requestFlatTimeAnalystReport() {
      if (!flatTimeState.aiContext) {
        setFlatTimeAiOutput("Flat Time context is not ready yet. Load CSV files and try again.", true);
        return;
      }

      flatTimeState.aiReportText = buildFlatTimeStandardReport(
        flatTimeState.aiContext,
        ui.flatTimeAiScope ? ui.flatTimeAiScope.value : "selected-well"
      );
      setFlatTimeAiOutput(
        flatTimeState.aiReportText
          ? renderFlatTimeAiMarkdown(flatTimeState.aiReportText)
          : "WeeklyReport could not build the report for the current scope.",
        !flatTimeState.aiReportText
      );
      setFlatTimeAiBusyState(false);
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
        flatTimeState.aiContext = null;
        flatTimeState.aiReportText = "";
        if (ui.flatTimeComparisonPhaseTimes) ui.flatTimeComparisonPhaseTimes.innerHTML = '<div class="empty">Upload flat time CSV files to compare phase times.</div>';
        if (ui.flatTimeComparisonDrillingTime) ui.flatTimeComparisonDrillingTime.innerHTML = '<div class="empty">Upload flat time CSV files to compare drilling-only time.</div>';
        if (ui.flatTimeComparisonPhaseDetails) ui.flatTimeComparisonPhaseDetails.innerHTML = '<div class="empty">Upload flat time CSV files to review phase detail comparisons.</div>';
        if (ui.flatTimeComparisonDepthDays) ui.flatTimeComparisonDepthDays.innerHTML = '<div class="empty">Upload flat time CSV files to draw the phase progression chart.</div>';
        if (ui.flatTimeSummary) ui.flatTimeSummary.innerHTML = '<div class="empty">Upload flat time CSV files to start the comparison.</div>';
        ui.flatTimeTableSummary.innerHTML = '<div class="empty">Upload flat time CSV files to build the selected well table summary.</div>';
        ui.flatTimeWellRanking.innerHTML = '<div class="empty">Upload flat time CSV files to rank wells by excess time.</div>';
        if (ui.flatTimeParetoChart) ui.flatTimeParetoChart.innerHTML = '<div class="empty">Upload flat time CSV files to build a Pareto of recoverable hours.</div>';
        if (ui.flatTimeAllWellsSections) ui.flatTimeAllWellsSections.innerHTML = '<div class="empty">Upload flat time CSV files to summarize all wells by section.</div>';
        ui.flatTimeSelectedWellSections.innerHTML = '<div class="empty">Upload flat time CSV files to compare the selected well by section.</div>';
        ui.flatTimeOffsetComparison.innerHTML = '<div class="empty">Upload flat time CSV files to compare the selected well against loaded offsets.</div>';
        ui.flatTimeSectionHeatmap.innerHTML = '<div class="empty">Upload flat time CSV files to draw the section heat map.</div>';
        if (ui.flatTimeSavingsSummary) ui.flatTimeSavingsSummary.innerHTML = '<div class="empty">Upload flat time CSV files to summarize recoverable time by section.</div>';
        if (ui.flatTimeNarrative) ui.flatTimeNarrative.innerHTML = '<div class="empty">Upload flat time CSV files to generate the executive explanation.</div>';
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
        ui.flatTimeDrilldownNote.textContent = "Use the selectors below to inspect a well, section, or activity benchmark and review the ideal-time logic.";
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
        renderFlatTimeAiPanel(null);
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
        flatTimeState.aiContext = null;
        flatTimeState.aiReportText = "";
        if (ui.flatTimeComparisonPhaseTimes) ui.flatTimeComparisonPhaseTimes.innerHTML = '<div class="empty">No flat time comparison is available for this section size.</div>';
        if (ui.flatTimeComparisonDrillingTime) ui.flatTimeComparisonDrillingTime.innerHTML = '<div class="empty">No drilling-only comparison is available for this section size.</div>';
        if (ui.flatTimeComparisonPhaseDetails) ui.flatTimeComparisonPhaseDetails.innerHTML = '<div class="empty">No phase detail comparison is available for this section size.</div>';
        if (ui.flatTimeComparisonDepthDays) ui.flatTimeComparisonDepthDays.innerHTML = '<div class="empty">No phase progression is available for this section size.</div>';
        if (ui.flatTimeSummary) ui.flatTimeSummary.innerHTML = '<div class="empty">No benchmark activities match the selected section size.</div>';
        ui.flatTimeTableSummary.innerHTML = '<div class="empty">No selected well table summary is available for this section size.</div>';
        ui.flatTimeWellRanking.innerHTML = '<div class="empty">No well ranking available for this section size.</div>';
        if (ui.flatTimeParetoChart) ui.flatTimeParetoChart.innerHTML = '<div class="empty">No Pareto data available for this section size.</div>';
        if (ui.flatTimeAllWellsSections) ui.flatTimeAllWellsSections.innerHTML = '<div class="empty">No well-section summary available for this section size.</div>';
        ui.flatTimeSelectedWellSections.innerHTML = '<div class="empty">No selected-well section comparison available for this section size.</div>';
        ui.flatTimeOffsetComparison.innerHTML = '<div class="empty">No offset comparison available for this section size.</div>';
        ui.flatTimeSectionHeatmap.innerHTML = '<div class="empty">No section heat map available for this section size.</div>';
        if (ui.flatTimeSavingsSummary) ui.flatTimeSavingsSummary.innerHTML = '<div class="empty">No savings summary available for this section size.</div>';
        if (ui.flatTimeNarrative) ui.flatTimeNarrative.innerHTML = '<div class="empty">No executive explanation available for this section size.</div>';
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
        ui.flatTimeDrilldownNote.textContent = "Use the selectors below to inspect a well, section, or activity benchmark and review the ideal-time logic.";
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
        renderFlatTimeAiPanel(null);
        return;
      }

      const groupItems = buildFlatTimeGroupItems(datasets, totalKey);
      const activityItems = buildFlatTimeActivityItems(datasets, metricKey);
      const comparisonData = buildFlatTimeComparisonData(datasets);
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
        format: (value) => formatHours(value),
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
      const selectedActivities = normalizeFlatTimeFocusActivities(allOpportunities, topActivityFocus, { allowEmpty: true });
      populateFlatTimeWellOptions(datasets, flatTimeState.focusWell || worstWell);
      const selectedWell = ui.flatTimeWell.value || flatTimeState.focusWell || worstWell;
      const selectedDataset = datasets.find((dataset) => dataset.subjectWell === selectedWell) || datasets[0];
      renderFlatTimeComparisonOverview(
        ui.flatTimeComparisonPhaseTimes,
        comparisonData.phaseTimesItems,
        datasets,
        selectedDataset,
        "No phase totals are available after excluding drilling actions D, DMR, and DMS.",
        { labelHeader: "Phase" }
      );
      renderFlatTimeComparisonOverview(
        ui.flatTimeComparisonDrillingTime,
        comparisonData.drillingTimeItems,
        datasets,
        selectedDataset,
        "No drilling-action totals are available for D, DMR, or DMS in the current scope.",
        { labelHeader: "Phase" }
      );
      renderFlatTimeComparisonPhaseDetails(
        ui.flatTimeComparisonPhaseDetails,
        comparisonData.phaseDetails,
        datasets,
        selectedDataset
      );
      renderFlatTimeComparisonProgression(
        ui.flatTimeComparisonDepthDays,
        comparisonData.progressionItems,
        datasets,
        selectedDataset
      );
      const selectedWellParetoItems = buildSelectedWellParetoItems(selectedDataset, allOpportunities);
      const sectionBreakdownMap = buildSectionBreakdownMap(datasets, allOpportunities, metricKey);
      const selectedWellSectionItems = buildSelectedWellSectionItems(datasets, selectedDataset, sectionBreakdownMap);
      const allWellSectionRows = buildAllWellsSectionSummaryRows(datasets, sectionBreakdownMap);
      const sectionSavingsRows = buildSectionSavingsSummary(datasets, sectionBreakdownMap);
      const sectionBenchmarkItems = buildSectionBenchmarkItems(datasets, metricKey, allOpportunities);
      const rigBenchmarkRows = buildRigBenchmarkSummary(datasets, allOpportunities);
      const opportunityPipeline = buildOpportunityPipeline(allOpportunities).slice(0, Math.max(topN, 8));
      const selectedActivityAggregate = buildSelectedActivityAggregate(allOpportunities, selectedActivities);
      flatTimeState.aiContext = buildFlatTimeAiContext(
        datasets,
        selectedDataset,
        selectedWell,
        selectedRig,
        selectedSectionSize,
        selectedActivities,
        selectedActivityAggregate,
        metricKey,
        totalKey,
        wellRanking,
        sectionSavingsRows,
        selectedWellSectionItems,
        rigBenchmarkRows,
        opportunityPipeline,
        allOpportunities
      );
      flatTimeState.aiReportText = "";
      renderFlatTimeAiPanel(flatTimeState.aiContext);

      renderFlatTimeTableSummary(ui.flatTimeTableSummary, selectedDataset, allOpportunities, selectedSectionSize);

      renderTableHtml(
        ui.flatTimeWellRanking,
        ["Rig", "Well", "Actual Total (hr / d)", "Ideal Total (hr / d)", "Excess Time (hr / d)", "Top Drivers"],
        wellRanking.map((row) => [
          escapeHtml(row.rigLabel),
          flatTimeActionButtonHtml("well", row.wellLabel, row.wellLabel),
          escapeHtml(formatHours(row.actualTotal) + " hr / " + formatDays(row.actualTotal / 24) + " d"),
          escapeHtml(formatHours(row.idealTotal) + " hr / " + formatDays(row.idealTotal / 24) + " d"),
          flatTimeTrendHtml(row.excessTotal),
          row.topDrivers.length || row.otherDriversGap > 0
            ? [
                ...row.topDrivers.map((driver) => flatTimeActivityLabelHtml(driver.activity) + escapeHtml(" (" + driver.group + ", +" + formatHours(driver.gap) + " hr)")),
                ...(row.otherDriversGap > 0 ? ['<span style="color:var(--muted); font-weight:700;">' + escapeHtml("Other drivers (+" + formatHours(row.otherDriversGap) + " hr)") + '</span>'] : []),
              ].join("<br>")
            : escapeHtml("No excess detected"),
        ])
      );

      renderParetoChart(ui.flatTimeParetoChart, selectedWellParetoItems, "selectedWellRecoverableHours");
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
          escapeHtml(formatHours(row.recoverableHours)),
          escapeHtml(formatDays(row.recoverableDays)),
          escapeHtml(String(row.impactedWells)),
          flatTimeActivityLabelHtml(row.topActivity),
          confidenceBadgeHtml(row.confidence),
          escapeHtml(row.recommendation),
        ])
      );
      if (ui.flatTimeNarrative) {
        ui.flatTimeNarrative.innerHTML = buildExecutiveNarrativeHtml(selectedDataset, selectedWellSectionItems, sectionSavingsRows, wellRanking);
      }
      renderWaterfallChart(ui.flatTimeWaterfallChart, selectedDataset, allOpportunities);
      renderMultiSeriesChart(
        ui.flatTimeSectionBenchmarkChart,
        sectionBenchmarkItems,
        [
          { key: "actualAverage", label: "Actual avg", color: "#1264d6", format: (value) => formatHours(value) },
          { key: "idealTime", label: "Recommended ideal", color: "#0f766e", format: (value) => formatHours(value) },
          { key: "spread", label: "Spread", color: "#c06a0a", format: (value) => formatHours(value) },
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
          formatHours(row.averageFlatTime),
          formatHours(row.averageIdealTime),
          formatHours(row.excessTime),
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
          escapeHtml(formatHours(row.idealTime)),
          escapeHtml(formatHours(row.totalRecoverableHours)),
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
            escapeHtml(formatHours(opportunity.fastestTime)),
            escapeHtml(formatHours(opportunity.p25Value)),
            escapeHtml(formatHours(opportunity.medianValue)),
            escapeHtml(formatHours(opportunity.meanValue)),
            escapeHtml(formatHours(opportunity.idealTime) + " (" + opportunity.idealRule + ")"),
            escapeHtml(opportunity.variability),
            confidenceBadgeHtml(opportunity.confidence),
            escapeHtml(formatHours(opportunity.totalRecoverableHours)),
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
                  (opportunity.occurrenceCount - 1) + " peer wells avg " + formatHours(peerReference) +
                  " hr; " + opportunity.topEntry.label + " ran " + formatHours(opportunity.topEntry.value) +
                  " hr; ideal = " + formatHours(opportunity.idealTime) + " hr; gap = " + formatHours(opportunity.gapToIdeal) + " hr"
                )
              : "Only one well available, so no peer comparison yet";

          return [
            escapeHtml(formatFlatTimeSectionSize(opportunity.sectionSize)),
            escapeHtml(opportunity.groupLabel),
            flatTimeActionButtonHtml("activity", opportunity.activityLabel, opportunity.activityLabel),
            escapeHtml(String(opportunity.occurrenceCount)),
            escapeHtml(opportunity.topEntry.rigLabel || "Rig not mapped"),
            flatTimeActionButtonHtml("well", opportunity.topEntry.label || "", opportunity.topEntry.label || "N/A"),
            escapeHtml(formatHours(opportunity.topEntry.value)),
            escapeHtml(formatHours(peerReference)),
            escapeHtml(formatHours(opportunity.idealTime)),
            escapeHtml(formatHours(opportunity.gapToIdeal)),
            escapeHtml(explanation + " (" + opportunity.idealRule + ")"),
          ];
        })
      );

      renderTable(
        ui.flatTimeGroupTable,
        ["Group", ...datasets.map((dataset) => dataset.subjectWell), "Total"],
        groupItems.map((item) => [
          item.label,
          ...datasets.map((dataset) => formatHours(item[dataset.id] || 0)),
          formatHours(item.total),
        ])
      );

      renderTable(
        ui.flatTimeLossDrivers,
        ["Rig", "Well", "Top Driver 1", "Top Driver 2", "Top Driver 3", "Excess Time (hr)"],
        wellRanking.map((row) => [
          row.rigLabel,
          row.wellLabel,
          row.topDrivers[0] ? row.topDrivers[0].activity + " (+" + formatHours(row.topDrivers[0].gap) + " hr)" : "-",
          row.topDrivers[1] ? row.topDrivers[1].activity + " (+" + formatHours(row.topDrivers[1].gap) + " hr)" : "-",
          row.topDrivers[2] ? row.topDrivers[2].activity + " (+" + formatHours(row.topDrivers[2].gap) + " hr)" : "-",
          formatHours(row.excessTotal),
        ])
      );

      renderVariabilityChart(ui.flatTimeVariabilityChart, allOpportunities, topN);
      renderHeatmap(ui.flatTimeHeatmap, datasets, allOpportunities, topN, flatTimeState.heatmapMode);
      renderPerfectFlatTimeChart(ui.flatTimePerfectChart, datasets, metricKey);
      renderFlatTimeDrilldown(datasets, allOpportunities, selectedWell, selectedActivities, rigDatasets);
      wireFlatTimeFocusActions();
    }

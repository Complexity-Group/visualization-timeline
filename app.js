import { ACTIVE_TAB_KEY, TIMELINE_CONFIGS } from "./config.js";
import { normalizeParadigmKey, parseWorkbookData } from "./workbook.js";

const DEFAULT_GROUP_ID = "default";
const PARADIGM_GROUP_ID = "paradigms";
const ADJUSTMENT_GROUP_ID = "adjustments";
const ONGOING_ENDPOINT_YEAR = 2025;
const TREND_ROW_HEIGHT = 34;
const TREND_TIME_AXIS_HEIGHT = 34;
const SECONDARY_ROW_HEIGHT = 36;
const TIMELINE_CONTAINER_TOP_PADDING = 8;
const PARADIGM_CONTAINER_BOTTOM_PADDING = 8;
const ADJUSTMENT_CONTAINER_BOTTOM_PADDING = 8;
const EXPORT_TREND_ROW_HEIGHT = 52;
const EXPORT_TIME_AXIS_HEIGHT = 54;

let trendTimelineInstance = null;
let paradigmTimelineInstance = null;
let adjustmentTimelineInstance = null;
let activeTabKey = null;
let loadGeneration = 0;
const tabCache = {};
let currentTrendItems = [];
let currentParadigmItems = [];
let currentAdjustmentItems = [];
let trendItemsById = new Map();
let paradigmItemsById = new Map();
let adjustmentItemsById = new Map();
let eventsByTrendSubgroup = new Map();
let eventsByParadigm = new Map();
let eventsByAdjustment = new Map();
let currentTrendOptions = null;
let currentParadigmOptions = null;
let currentAdjustmentOptions = null;
let pendingInitialWindow = null;
let isExportingImage = false;
let activeExportWindow = null;
let paradigmRenderedHeightCorrection = 0;
let paradigmHeightFitGeneration = 0;

const el = {
  widgetRoot: document.getElementById("widgetRoot"),
  timelineTop: document.getElementById("timelineTop"),
  timelineBottom: document.getElementById("timelineBottom"),
  timelineAdjustments: document.getElementById("timelineAdjustments"),
  timelineShell: document.getElementById("timelineShell"),
  paradigmShell: document.getElementById("paradigmShell"),
  adjustmentShell: document.getElementById("adjustmentShell"),
  timelineStage: document.getElementById("timelineStage"),
  timelineTitle: document.getElementById("timelineTitle"),
  timelineChartArea: document.getElementById("timelineChartArea"),
  timelineLabelRail: document.getElementById("timelineLabelRail"),
  trendSectionLabel: document.getElementById("trendSectionLabel"),
  exportLoadingOverlay: document.getElementById("exportLoadingOverlay"),
  tabs: document.getElementById("tabs"),
  status: document.getElementById("status"),
  statusText: document.getElementById("statusText"),
  spinner: document.getElementById("spinner"),
  skeleton: document.getElementById("skeleton"),
  shareBtn: document.getElementById("shareBtn"),
  exportBtn: document.getElementById("exportBtn"),
  zoomInBtn: document.getElementById("zoomInBtn"),
  zoomOutBtn: document.getElementById("zoomOutBtn"),
  resetZoomBtn: document.getElementById("resetZoomBtn"),
  fullscreenBtn: document.getElementById("fullscreenBtn"),
  floatingControls: document.getElementById("floatingControls"),
  shareBtnFs: document.getElementById("shareBtnFs"),
  exportBtnFs: document.getElementById("exportBtnFs"),
  zoomInBtnFs: document.getElementById("zoomInBtnFs"),
  zoomOutBtnFs: document.getElementById("zoomOutBtnFs"),
  resetZoomBtnFs: document.getElementById("resetZoomBtnFs"),
  fullscreenBtnFs: document.getElementById("fullscreenBtnFs"),
  modalOverlay: document.getElementById("modalOverlay"),
  modalTitle: document.getElementById("modalTitle"),
  modalBody: document.getElementById("modalBody"),
  modalClose: document.getElementById("modalClose"),
  shareBtnMobile: document.getElementById("shareBtnMobile"),
  exportBtnMobile: document.getElementById("exportBtnMobile"),
  zoomInBtnMobile: document.getElementById("zoomInBtnMobile"),
  zoomOutBtnMobile: document.getElementById("zoomOutBtnMobile"),
  resetZoomBtnMobile: document.getElementById("resetZoomBtnMobile"),
  fullscreenBtnMobile: document.getElementById("fullscreenBtnMobile")
};

let isFullscreen = false;
let fullscreenTransitionTimer = null;

function setLoadingState(message) {
  el.status.classList.remove("hidden", "error");
  el.spinner.classList.remove("hidden");
  el.statusText.textContent = message;
  el.skeleton.classList.remove("hidden");
  el.timelineTitle.classList.add("hidden");
  el.timelineChartArea.classList.add("hidden");
}

function setReadyState(message) {
  el.status.classList.remove("error");
  el.spinner.classList.add("hidden");
  el.statusText.textContent = message;
  el.skeleton.classList.add("hidden");
  el.timelineTitle.classList.remove("hidden");
  el.timelineChartArea.classList.remove("hidden");
}

function setErrorState(message) {
  el.status.classList.remove("hidden");
  el.status.classList.add("error");
  el.spinner.classList.add("hidden");
  el.statusText.textContent = message;
  el.skeleton.classList.add("hidden");
  el.timelineTitle.classList.add("hidden");
  el.timelineChartArea.classList.add("hidden");
}

async function fetchWorkbook(config) {
  const requestUrl = config.workbookUrl;
  let response;

  try {
    response = await fetch(requestUrl, { cache: "no-store" });
  } catch (error) {
    throw new Error("Network/CORS error while downloading the Google Sheets workbook.");
  }

  if (!response.ok) {
    throw new Error("Failed to fetch the Google Sheets workbook (" + response.status + ")");
  }

  const bytes = await response.arrayBuffer();
  const signature = new Uint8Array(bytes, 0, Math.min(bytes.byteLength, 2));
  if (signature.length < 2 || signature[0] !== 0x50 || signature[1] !== 0x4b) {
    throw new Error("The Google Sheets export did not return a valid XLSX workbook.");
  }

  if (typeof XLSX === "undefined" || typeof XLSX.read !== "function") {
    throw new Error("SheetJS failed to load.");
  }

  try {
    return XLSX.read(new Uint8Array(bytes), {
      type: "array",
      dense: true,
      cellFormula: false,
      cellHTML: false,
      cellStyles: false
    });
  } catch (error) {
    throw new Error("Unable to parse the XLSX workbook: " + (error.message || error));
  }
}

function createYearDate(year) {
  const date = new Date(0);
  date.setHours(0, 0, 0, 0);
  date.setFullYear(year, 0, 1);
  return date;
}

function yearDateInfo(year) {
  const isBce = year < 0;
  return {
    date: createYearDate(year),
    isBce,
    raw: isBce ? Math.abs(year) + " BCE" : String(year)
  };
}

function getRenderedRangeYears(startYear, endYear, timelineStartYear) {
  if (endYear <= timelineStartYear) {
    return {
      startYear: timelineStartYear,
      endYear: timelineStartYear + 10
    };
  }
  return {
    startYear,
    endYear: startYear === endYear ? endYear + 10 : endYear
  };
}

function normalizeSubgroupId(value) {
  return String(value == null ? "" : value).trim().toLowerCase();
}

function hashString(value) {
  const text = String(value == null ? "" : value);
  let hash = 0;
  for (let i = 0; i < text.length; i += 1) {
    hash = ((hash << 5) - hash) + text.charCodeAt(i);
    hash |= 0;
  }
  return hash;
}

function createGroupWithSubgroups(groupId, subgroupIds, subgroupOrderer) {
  const subgroupStack = {};
  for (let i = 0; i < subgroupIds.length; i += 1) {
    subgroupStack[subgroupIds[i]] = false;
  }

  return {
    id: groupId,
    content: "",
    subgroupOrder: subgroupOrderer || compareSubgroupsByDuration,
    subgroupStack
  };
}

function compareSubgroupsByDuration(first, second) {
  const durationDifference = second.subgroupDuration - first.subgroupDuration;
  if (durationDifference !== 0) {
    return durationDifference;
  }

  return String(first.subgroupTitle).localeCompare(String(second.subgroupTitle), undefined, {
    sensitivity: "base"
  });
}

function compareTopTimelineSubgroups(first, second) {
  return compareSubgroupsByStartDate(first, second);
}

function compareSubgroupsByStartDate(first, second) {
  const startDifference = first.subgroupStartYear - second.subgroupStartYear;
  if (startDifference !== 0) {
    return startDifference;
  }

  return String(first.subgroupTitle).localeCompare(String(second.subgroupTitle), undefined, {
    sensitivity: "base"
  });
}

function orderSubgroupsByStartDate(subgroupIds, items) {
  const summaries = {};

  subgroupIds.forEach(function (subgroupId) {
    summaries[subgroupId] = {
      startYear: Infinity,
      title: subgroupId
    };
  });

  items.forEach(function (item) {
    if (item.type !== "range" || !item.subgroup || !summaries[item.subgroup]) {
      return;
    }

    const summary = summaries[item.subgroup];
    summary.startYear = Math.min(summary.startYear, item.startYear);
    summary.title = item.titleText || item.content || item.subgroup;
  });

  items.forEach(function (item) {
    const summary = item.subgroup ? summaries[item.subgroup] : null;
    if (summary) {
      item.subgroupStartYear = summary.startYear;
      item.subgroupTitle = summary.title;
    }
  });

  return Array.from(subgroupIds).sort(function (firstId, secondId) {
    return compareSubgroupsByStartDate(
      {
        subgroupStartYear: summaries[firstId].startYear,
        subgroupTitle: summaries[firstId].title
      },
      {
        subgroupStartYear: summaries[secondId].startYear,
        subgroupTitle: summaries[secondId].title
      }
    );
  });
}

function orderSubgroupsByDuration(subgroupIds, items) {
  const summaries = {};

  subgroupIds.forEach(function (subgroupId) {
    summaries[subgroupId] = {
      earliestStart: Infinity,
      latestEnd: -Infinity,
      duration: 0,
      title: subgroupId
    };
  });

  items.forEach(function (item) {
    if (item.type !== "range" || !item.subgroup || !summaries[item.subgroup]) {
      return;
    }

    const summary = summaries[item.subgroup];
    summary.earliestStart = Math.min(summary.earliestStart, item.start.getTime());
    summary.latestEnd = Math.max(summary.latestEnd, item.end.getTime());
    summary.duration = Math.max(summary.latestEnd - summary.earliestStart, 0);
    summary.title = item.titleText || item.content || item.subgroup;
  });

  items.forEach(function (item) {
    const summary = item.subgroup ? summaries[item.subgroup] : null;
    if (summary) {
      item.subgroupDuration = summary.duration;
      item.subgroupTitle = summary.title;
    }
  });

  return Array.from(subgroupIds).sort(function (firstId, secondId) {
    return compareSubgroupsByDuration(
      {
        subgroupDuration: summaries[firstId].duration,
        subgroupTitle: summaries[firstId].title
      },
      {
        subgroupDuration: summaries[secondId].duration,
        subgroupTitle: summaries[secondId].title
      }
    );
  });
}

function nextTick() {
  return new Promise((resolve) => setTimeout(resolve, 0));
}

function nextAnimationFrame() {
  return new Promise((resolve) => requestAnimationFrame(resolve));
}

function buildTimelineData(workbookData, config) {
  const timelineStart = createYearDate(workbookData.timeline.startYear);
  const timelineEnd = createYearDate(workbookData.timeline.endYear);
  const timeline = {
    title: workbookData.timeline.title,
    trendSubgroups: new Set(),
    trendItems: [],
    paradigmSubgroups: new Set(),
    paradigmItems: [],
    adjustmentSubgroups: new Set(),
    adjustmentItems: []
  };
  const paradigms = new Map();
  const adjustments = new Map();
  let trendCounter = 1;
  let eventCounter = 1;
  let eventRowCount = 0;

  workbookData.tabs.forEach(function (events, sheetName) {
    const subgroup = normalizeSubgroupId(sheetName);
    timeline.trendSubgroups.add(subgroup);

    let minYear = Infinity;
    let maxYear = -Infinity;
    for (let i = 0; i < events.length; i += 1) {
      minYear = Math.min(minYear, events[i].startYear);
      maxYear = Math.max(maxYear, events[i].startYear);
    }

    const rangeStart = yearDateInfo(minYear);
    const rangeEnd = yearDateInfo(maxYear);
    const renderedRangeYears = getRenderedRangeYears(minYear, maxYear, workbookData.timeline.startYear);
    const renderedRangeStart = yearDateInfo(renderedRangeYears.startYear);
    const renderedRangeEnd = yearDateInfo(renderedRangeYears.endYear);
    const startEndpointClass = rangeStart.date.getTime() === timelineStart.getTime() ? "" : " dot-left";
    const endEndpointClass = maxYear < ONGOING_ENDPOINT_YEAR ? " dot-right" : " arrow-right";
    const endpointClasses = startEndpointClass + endEndpointClass;
    timeline.trendItems.push({
      id: "trend-" + trendCounter++,
      group: DEFAULT_GROUP_ID,
      subgroup,
      type: "range",
      className: "trend-range trend-color-" + (Math.abs(hashString(subgroup)) % 8) + endpointClasses,
      start: renderedRangeStart.date,
      end: renderedRangeEnd.date,
      startYear: minYear,
      startRaw: rangeStart.raw,
      endRaw: rangeEnd.raw,
      startIsBce: rangeStart.isBce,
      endIsBce: rangeEnd.isBce,
      startYearLabel: rangeStart.raw,
      endYearLabel: minYear === maxYear || maxYear === workbookData.timeline.endYear
        ? ""
        : rangeEnd.raw,
      titleText: sheetName,
      displayTitleText: sheetName,
      content: sheetName
    });

    for (let i = 0; i < events.length; i += 1) {
      const event = events[i];
      const startInfo = yearDateInfo(event.startYear);
      timeline.trendItems.push({
        id: "event-" + eventCounter++,
        group: DEFAULT_GROUP_ID,
        subgroup,
        type: "point",
        start: startInfo.date,
        startRaw: startInfo.raw,
        startIsBce: startInfo.isBce,
        title: event.event,
        paradigm: event.paradigm,
        adjustment: event.adjustment,
        energySource: sheetName
      });
      eventRowCount += 1;

      const paradigmKey = normalizeParadigmKey(event.paradigm);
      if (paradigmKey) {
        const existing = paradigms.get(paradigmKey);
        if (existing) {
          existing.minYear = Math.min(existing.minYear, event.startYear);
          existing.maxYear = Math.max(existing.maxYear, event.startYear);
        } else {
          paradigms.set(paradigmKey, {
            title: event.paradigm,
            minYear: event.startYear,
            maxYear: event.startYear
          });
        }
      }

      const adjustmentKey = normalizeParadigmKey(event.adjustment);
      if (adjustmentKey) {
        const existingAdjustment = adjustments.get(adjustmentKey);
        if (existingAdjustment) {
          existingAdjustment.minYear = Math.min(existingAdjustment.minYear, event.startYear);
          existingAdjustment.maxYear = Math.max(existingAdjustment.maxYear, event.startYear);
        } else {
          adjustments.set(adjustmentKey, {
            title: event.adjustment,
            minYear: event.startYear,
            maxYear: event.startYear
          });
        }
      }
    }
  });

  let paradigmCounter = 1;
  paradigms.forEach(function (paradigm, paradigmKey) {
    const startInfo = yearDateInfo(paradigm.minYear);
    const endInfo = yearDateInfo(paradigm.maxYear);
    const renderedRangeYears = getRenderedRangeYears(
      paradigm.minYear,
      paradigm.maxYear,
      workbookData.timeline.startYear
    );
    const renderedStartInfo = yearDateInfo(renderedRangeYears.startYear);
    const renderedEndInfo = yearDateInfo(renderedRangeYears.endYear);
    timeline.paradigmSubgroups.add(paradigmKey);
    timeline.paradigmItems.push({
      id: "paradigm-" + paradigmCounter++,
      group: PARADIGM_GROUP_ID,
      subgroup: paradigmKey,
      type: "range",
      className: "paradigm-range"
        + (paradigm.minYear === workbookData.timeline.startYear ? "" : " dot-left")
        + (paradigm.maxYear < ONGOING_ENDPOINT_YEAR ? " dot-right" : " arrow-right"),
      start: renderedStartInfo.date,
      end: renderedEndInfo.date,
      startYear: paradigm.minYear,
      startRaw: startInfo.raw,
      endRaw: endInfo.raw,
      startIsBce: startInfo.isBce,
      endIsBce: endInfo.isBce,
      titleText: paradigm.title,
      displayText: paradigm.title + " (" + startInfo.raw + "-" + endInfo.raw + ")",
      content: escapeHtml(paradigm.title) + " (" + escapeHtml(startInfo.raw) + "-" + escapeHtml(endInfo.raw) + ")"
    });
  });

  let adjustmentCounter = 1;
  adjustments.forEach(function (adjustment, adjustmentKey) {
    const startInfo = yearDateInfo(adjustment.minYear);
    const endInfo = yearDateInfo(adjustment.maxYear);
    const renderedRangeYears = getRenderedRangeYears(
      adjustment.minYear,
      adjustment.maxYear,
      workbookData.timeline.startYear
    );
    const renderedStartInfo = yearDateInfo(renderedRangeYears.startYear);
    const renderedEndInfo = yearDateInfo(renderedRangeYears.endYear);
    timeline.adjustmentSubgroups.add(adjustmentKey);
    timeline.adjustmentItems.push({
      id: "adjustment-" + adjustmentCounter++,
      group: ADJUSTMENT_GROUP_ID,
      subgroup: adjustmentKey,
      type: "range",
      className: "adjustment-range"
        + (adjustment.minYear === workbookData.timeline.startYear ? "" : " dot-left")
        + (adjustment.maxYear < ONGOING_ENDPOINT_YEAR ? " dot-right" : " arrow-right"),
      start: renderedStartInfo.date,
      end: renderedEndInfo.date,
      startYear: adjustment.minYear,
      startRaw: startInfo.raw,
      endRaw: endInfo.raw,
      startIsBce: startInfo.isBce,
      endIsBce: endInfo.isBce,
      titleText: adjustment.title,
      displayText: adjustment.title + " (" + startInfo.raw + "-" + endInfo.raw + ")",
      content: escapeHtml(adjustment.title) + " (" + escapeHtml(startInfo.raw) + "-" + escapeHtml(endInfo.raw) + ")"
    });
  });

  const span = timelineEnd.getTime() - timelineStart.getTime();
  const bufferFactor = window.innerWidth <= 768 ? 0.02 : 0.01;
  const endWithBuffer = new Date(timelineEnd.getTime() + span * bufferFactor);

  const orderedTrendSubgroups = orderSubgroupsByStartDate(
    timeline.trendSubgroups,
    timeline.trendItems
  );
  const orderedParadigmSubgroups = orderSubgroupsByStartDate(
    timeline.paradigmSubgroups,
    timeline.paradigmItems
  );
  const orderedAdjustmentSubgroups = orderSubgroupsByStartDate(
    timeline.adjustmentSubgroups,
    timeline.adjustmentItems
  );

  return {
    title: timeline.title,
    // Items are top-oriented while the time axis remains at the bottom, so
    // chronological subgroup order maps directly to top-to-bottom row order.
    trendGroups: [createGroupWithSubgroups(
      DEFAULT_GROUP_ID,
      orderedTrendSubgroups,
      compareTopTimelineSubgroups
    )],
    paradigmGroups: [createGroupWithSubgroups(
      PARADIGM_GROUP_ID,
      orderedParadigmSubgroups,
      compareSubgroupsByStartDate
    )],
    adjustmentGroups: [createGroupWithSubgroups(
      ADJUSTMENT_GROUP_ID,
      orderedAdjustmentSubgroups,
      compareSubgroupsByStartDate
    )],
    trendItems: timeline.trendItems,
    paradigmItems: timeline.paradigmItems,
    adjustmentItems: timeline.adjustmentItems,
    trendOptions: {
      stack: false,
      stackSubgroups: true,
      start: timelineStart,
      end: endWithBuffer,
      height: Math.max(timeline.trendSubgroups.size * TREND_ROW_HEIGHT, TREND_ROW_HEIGHT) + TREND_TIME_AXIS_HEIGHT
    },
    paradigmOptions: {
      stack: false,
      stackSubgroups: true,
      start: timelineStart,
      end: timelineEnd,
      height: Math.max(timeline.paradigmSubgroups.size * SECONDARY_ROW_HEIGHT, SECONDARY_ROW_HEIGHT)
    },
    adjustmentOptions: {
      stack: false,
      stackSubgroups: true,
      start: timelineStart,
      end: timelineEnd,
      height: Math.max(timeline.adjustmentSubgroups.size * SECONDARY_ROW_HEIGHT, SECONDARY_ROW_HEIGHT)
        + TREND_TIME_AXIS_HEIGHT
    },
    timeAxisStep: workbookData.timeline.increment,
    eventRowCount
  };
}

function buildBounds(baseOptions) {
  const startYear = baseOptions && baseOptions.start instanceof Date
    ? baseOptions.start.getFullYear()
    : null;
  const endYear = baseOptions && baseOptions.end instanceof Date
    ? baseOptions.end.getFullYear()
    : null;

  return {
    min: startYear === null ? undefined : createYearDate(startYear),
    max: endYear === null ? undefined : (function () {
      const date = createYearDate(endYear);
      date.setMonth(11, 31);
      return date;
    })()
  };
}

function getAvailableTimelineHeight() {
  const chartArea = el.timelineChartArea;
  if (chartArea && !chartArea.classList.contains("hidden") && chartArea.clientHeight > 0) {
    return Math.max(chartArea.clientHeight, 2);
  }

  const stage = el.timelineStage;
  const stageHeight = stage && stage.clientHeight > 0
    ? stage.clientHeight
    : Math.max(window.innerHeight - 160, 240);
  const titleHeight = el.timelineTitle && !el.timelineTitle.classList.contains("hidden")
    ? el.timelineTitle.getBoundingClientRect().height
    : 0;

  // The stage is the viewport space left after the chrome header.
  return Math.max(Math.floor(stageHeight - titleHeight), 2);
}

function calculateTimelineHeights(trendOptions) {
  const requiredTrendHeight = Math.max(
    Math.round(trendOptions && trendOptions.height
      ? trendOptions.height
      : TREND_ROW_HEIGHT + TREND_TIME_AXIS_HEIGHT),
    TREND_ROW_HEIGHT
  );
  const requiredParadigmHeight = Math.max(
    Math.round(currentParadigmOptions && currentParadigmOptions.height
      ? currentParadigmOptions.height
      : SECONDARY_ROW_HEIGHT),
    SECONDARY_ROW_HEIGHT
  ) + paradigmRenderedHeightCorrection;
  const requiredAdjustmentHeight = Math.max(
    Math.round(currentAdjustmentOptions && currentAdjustmentOptions.height
      ? currentAdjustmentOptions.height
      : SECONDARY_ROW_HEIGHT),
    SECONDARY_ROW_HEIGHT
  );
  const fixedShellHeight = requiredTrendHeight
    + TIMELINE_CONTAINER_TOP_PADDING
    + requiredParadigmHeight
    + TIMELINE_CONTAINER_TOP_PADDING
    + PARADIGM_CONTAINER_BOTTOM_PADDING
    + TIMELINE_CONTAINER_TOP_PADDING
    + ADJUSTMENT_CONTAINER_BOTTOM_PADDING;
  const remainingAdjustmentHeight = Math.max(
    Math.floor(getAvailableTimelineHeight() - fixedShellHeight),
    0
  );

  return {
    trendHeight: requiredTrendHeight,
    paradigmHeight: requiredParadigmHeight,
    adjustmentHeight: Math.max(requiredAdjustmentHeight, remainingAdjustmentHeight)
  };
}

function getInitialTimelineHeight() {
  return Math.max(Math.floor(getAvailableTimelineHeight() / 3), 1);
}

function applyTimelineShellHeights(heights) {
  const trendShellHeight = heights.trendHeight + TIMELINE_CONTAINER_TOP_PADDING;
  const paradigmShellHeight = heights.paradigmHeight
    + TIMELINE_CONTAINER_TOP_PADDING
    + PARADIGM_CONTAINER_BOTTOM_PADDING;
  const adjustmentShellHeight = heights.adjustmentHeight
    + TIMELINE_CONTAINER_TOP_PADDING
    + ADJUSTMENT_CONTAINER_BOTTOM_PADDING;
  el.timelineShell.style.flex = "0 0 " + trendShellHeight + "px";
  el.timelineShell.style.height = trendShellHeight + "px";
  el.paradigmShell.style.flex = "0 0 " + paradigmShellHeight + "px";
  el.paradigmShell.style.height = paradigmShellHeight + "px";
  el.adjustmentShell.style.flex = "0 0 " + adjustmentShellHeight + "px";
  el.adjustmentShell.style.height = adjustmentShellHeight + "px";
  const chartContentHeight = trendShellHeight + paradigmShellHeight + adjustmentShellHeight;
  el.timelineChartArea.style.setProperty("--chart-content-height", chartContentHeight + "px");
  el.timelineLabelRail.style.setProperty("--trend-label-height", trendShellHeight + "px");
  el.timelineLabelRail.style.setProperty("--paradigm-label-height", paradigmShellHeight + "px");
  el.timelineLabelRail.style.setProperty("--adjustment-label-height", adjustmentShellHeight + "px");
}

function resetTimelineShellHeights() {
  el.timelineShell.style.removeProperty("flex");
  el.timelineShell.style.removeProperty("height");
  el.paradigmShell.style.removeProperty("flex");
  el.paradigmShell.style.removeProperty("height");
  el.adjustmentShell.style.removeProperty("flex");
  el.adjustmentShell.style.removeProperty("height");
  el.timelineChartArea.style.removeProperty("--chart-content-height");
}

function buildTrendItemTemplate(item) {
  if (!item || item.type !== "range") {
    return item ? (item.content || item.title || "") : "";
  }

  const wrapper = document.createElement("span");
  wrapper.className = "trend-content";

  const exportLine = document.createElement("span");
  exportLine.className = "trend-export-line";
  exportLine.setAttribute("aria-hidden", "true");
  wrapper.appendChild(exportLine);

  const startYear = document.createElement("span");
  startYear.className = "trend-start-year";
  startYear.textContent = (item.startYearLabel || "");
  wrapper.appendChild(startYear);

  const title = document.createElement("span");
  title.className = "trend-title-text";
  title.textContent = item.displayTitleText || item.titleText || item.content || "";
  wrapper.appendChild(title);

  if (item.endYearLabel) {
    const endYear = document.createElement("span");
    endYear.className = "trend-end-year";
    endYear.textContent = item.endYearLabel;
    wrapper.appendChild(endYear);
  }

  return wrapper;
}

function buildParadigmItemTemplate(item) {
  if (!item || item.type !== "range") {
    return item ? (item.content || item.title || "") : "";
  }

  const wrapper = document.createElement("span");
  wrapper.className = "paradigm-title-text";
  const exportLine = document.createElement("span");
  exportLine.className = "trend-export-line";
  exportLine.setAttribute("aria-hidden", "true");
  wrapper.appendChild(exportLine);
  wrapper.appendChild(document.createTextNode(item.displayText || item.titleText || item.content || ""));
  wrapper.dataset.exportItemId = String(item.id || "");
  wrapper.dataset.exportStartMs = String(new Date(item.start).getTime());
  wrapper.dataset.exportEndMs = String(new Date(item.end).getTime());
  return wrapper;
}

function buildTopOptions(baseOptions, timeAxisStep) {
  const bounds = buildBounds(baseOptions);
  const initialHeight = getInitialTimelineHeight();
  return {
    ...baseOptions,
    ...bounds,
    height: (initialHeight + TIMELINE_CONTAINER_TOP_PADDING) + "px",
    zoomMin: 1000 * 60 * 60 * 24 * 365,
    orientation: {
      axis: "bottom",
      item: "top"
    },
    verticalScroll: false,
    zoomKey: "ctrlKey",
    margin: {
      item: {
        horizontal: 8,
        vertical: 2
      },
      axis: 14
    },
    timeAxis: {
      scale: "year",
      step: timeAxisStep || 25
    },
    groupHeightMode: "fitItems",
    showCurrentTime: false,
    template: buildTrendItemTemplate,
    onInitialDrawComplete: scheduleInitialTopTimelineScroll,
  };
}

function buildBottomOptions(baseOptions, timeAxisStep) {
  const bounds = buildBounds(baseOptions);
  const initialHeight = getInitialTimelineHeight();
  return {
    ...baseOptions,
    ...bounds,
    height: (initialHeight + TIMELINE_CONTAINER_TOP_PADDING + PARADIGM_CONTAINER_BOTTOM_PADDING) + "px",
    orientation: "top",
    verticalScroll: false,
    zoomable: true,
    zoomKey: "ctrlKey",
    moveable: true,
    selectable: true,
    groupHeightMode: "fitItems",
    margin: {
      item: {
        horizontal: 0,
        vertical: 4
      },
      axis: 0
    },
    timeAxis: {
      scale: "year",
      step: timeAxisStep || 25
    },
    showMajorLabels: false,
    showMinorLabels: true,
    showCurrentTime: false,
    template: buildParadigmItemTemplate
  };
}

let lastAppliedTimelineHeights = null;

function applyTimelineHeights() {
  // Export temporarily expands all timelines to their complete content
  // heights. A ResizeObserver fires during that expansion, so responsive
  // sizing must not overwrite the export-only dimensions.
  if (isExportingImage || !trendTimelineInstance || !paradigmTimelineInstance || !adjustmentTimelineInstance) {
    return false;
  }
  const heights = calculateTimelineHeights(currentTrendOptions);
  const heightKey = heights.trendHeight + ":" + heights.paradigmHeight + ":" + heights.adjustmentHeight;
  if (lastAppliedTimelineHeights === heightKey) {
    return false;
  }
  lastAppliedTimelineHeights = heightKey;
  applyTimelineShellHeights(heights);
  trendTimelineInstance.setOptions({ height: (heights.trendHeight + TIMELINE_CONTAINER_TOP_PADDING) + "px" });
  paradigmTimelineInstance.setOptions({
    height: (heights.paradigmHeight + TIMELINE_CONTAINER_TOP_PADDING + PARADIGM_CONTAINER_BOTTOM_PADDING) + "px"
  });
  adjustmentTimelineInstance.setOptions({
    height: (heights.adjustmentHeight + TIMELINE_CONTAINER_TOP_PADDING + ADJUSTMENT_CONTAINER_BOTTOM_PADDING) + "px"
  });
  return true;
}

function fitParadigmTimelineHeightToRenderedContent() {
  if (isExportingImage || !paradigmTimelineInstance) {
    return false;
  }

  const heightDeficit = measureTimelineHeightDeficit(
    el.timelineBottom,
    ".vis-item.paradigm-range",
    PARADIGM_CONTAINER_BOTTOM_PADDING
  );
  if (heightDeficit <= 0) {
    return false;
  }

  // Row-count estimates can be short by a few pixels per subgroup. Preserve
  // the measured correction for this tab so later resizes and export cleanup
  // cannot restore the clipped estimated height.
  const previousAdjustmentHeight = calculateTimelineHeights(currentTrendOptions).adjustmentHeight;
  paradigmRenderedHeightCorrection += heightDeficit;
  const heights = calculateTimelineHeights(currentTrendOptions);
  lastAppliedTimelineHeights = heights.trendHeight + ":" + heights.paradigmHeight + ":" + heights.adjustmentHeight;
  applyTimelineShellHeights(heights);
  paradigmTimelineInstance.setOptions({
    height: (heights.paradigmHeight + TIMELINE_CONTAINER_TOP_PADDING + PARADIGM_CONTAINER_BOTTOM_PADDING) + "px"
  });
  paradigmTimelineInstance.redraw();
  scrollTimelineToFirstRow(paradigmTimelineInstance, el.timelineBottom);
  if (heights.adjustmentHeight !== previousAdjustmentHeight) {
    adjustmentTimelineInstance.setOptions({
      height: (heights.adjustmentHeight + TIMELINE_CONTAINER_TOP_PADDING + ADJUSTMENT_CONTAINER_BOTTOM_PADDING) + "px"
    });
    adjustmentTimelineInstance.redraw();
    scrollTimelineToFirstRow(adjustmentTimelineInstance, el.timelineAdjustments);
  }
  return true;
}

function scheduleParadigmTimelineContentFit() {
  const generation = ++paradigmHeightFitGeneration;

  function runFitPass(attempt) {
    requestAnimationFrame(function () {
      if (generation !== paradigmHeightFitGeneration || isExportingImage || !paradigmTimelineInstance) {
        return;
      }

      scrollTimelineToFirstRow(paradigmTimelineInstance, el.timelineBottom);
      requestAnimationFrame(function () {
        if (generation !== paradigmHeightFitGeneration || isExportingImage) {
          return;
        }

        const changed = fitParadigmTimelineHeightToRenderedContent();
        if (changed && attempt < 2) {
          runFitPass(attempt + 1);
        }
      });
    });
  }

  runFitPass(0);
}

function syncTimelineWindows(instances) {
  let isSyncing = false;

  function wire(source) {
    source.on("rangechange", (props) => {
      if (isSyncing || isExportingImage) {
        return;
      }

      isSyncing = true;
      instances.forEach(function (target) {
        if (target !== source) {
          target.setWindow(props.start, props.end, { animation: false });
        }
      });
      isSyncing = false;
    });
  }

  instances.forEach(wire);
 }

function refreshVisibleTimelines() {
  if (isExportingImage || !trendTimelineInstance || !paradigmTimelineInstance || !adjustmentTimelineInstance) {
    return;
  }

  // Timelines are created while containers are hidden; redraw once visible so axis/layout paints correctly.
  requestAnimationFrame(() => {
    // First paint all timelines, then apply their full content heights after
    // vis has a valid visible DOM.
    trendTimelineInstance.redraw();
    paradigmTimelineInstance.redraw();
    adjustmentTimelineInstance.redraw();

    requestAnimationFrame(() => {
      applyTimelineHeights();
      trendTimelineInstance.redraw();
      paradigmTimelineInstance.redraw();
      adjustmentTimelineInstance.redraw();

      if (pendingInitialWindow) {
        trendTimelineInstance.setWindow(pendingInitialWindow.start, pendingInitialWindow.end, { animation: false });
        pendingInitialWindow = null;
      } else if (currentTrendOptions) {
        trendTimelineInstance.setWindow(currentTrendOptions.start, currentTrendOptions.end, { animation: false });
      }

      const topWindow = trendTimelineInstance.getWindow();
      paradigmTimelineInstance.setWindow(topWindow.start, topWindow.end, { animation: false });
      adjustmentTimelineInstance.setWindow(topWindow.start, topWindow.end, { animation: false });
      scheduleInitialTopTimelineScroll();
      scheduleParadigmTimelineContentFit();
    });
  });
}

function anchorTrendEndpointYears() {
  const ranges = el.timelineTop.querySelectorAll(".vis-item.trend-range");
  for (let i = 0; i < ranges.length; i += 1) {
    const range = ranges[i];
    const startYears = Array.from(range.querySelectorAll(".trend-start-year"));
    const endYears = Array.from(range.querySelectorAll(".trend-end-year"));
    const startYear = startYears.find(function (node) {
      return node.parentElement === range;
    }) || startYears[0];
    const endYear = endYears.find(function (node) {
      return node.parentElement === range;
    }) || endYears[0];

    // A vis-timeline redraw may render a new template after an endpoint was
    // moved onto the range. Keep exactly one node so a year is never painted
    // twice for the same energy source/use.
    for (let nodeIndex = 0; nodeIndex < startYears.length; nodeIndex += 1) {
      if (startYears[nodeIndex] !== startYear) {
        startYears[nodeIndex].remove();
      }
    }
    for (let nodeIndex = 0; nodeIndex < endYears.length; nodeIndex += 1) {
      if (endYears[nodeIndex] !== endYear) {
        endYears[nodeIndex].remove();
      }
    }

    // vis-timeline nests item content in a shrink-to-fit wrapper. Moving the
    // endpoint labels to the range itself makes 100% mean the full arrow length.
    if (startYear && startYear.parentElement !== range) {
      range.appendChild(startYear);
    }
    if (endYear && endYear.parentElement !== range) {
      range.appendChild(endYear);
    }
  }
}

function unwrapDefaultTimelineRangeContent() {
  const rangeCollections = [
    [el.timelineTop, ".vis-item.trend-range"],
    [el.timelineBottom, ".vis-item.paradigm-range"],
    [el.timelineAdjustments, ".vis-item.adjustment-range"]
  ];

  for (let collectionIndex = 0; collectionIndex < rangeCollections.length; collectionIndex += 1) {
    const root = rangeCollections[collectionIndex][0];
    const selector = rangeCollections[collectionIndex][1];
    const ranges = root.querySelectorAll(selector);

    for (let rangeIndex = 0; rangeIndex < ranges.length; rangeIndex += 1) {
      const range = ranges[rangeIndex];
      const content = range.querySelector(".vis-item-content");

      // vis-timeline puts range labels inside an overflow:hidden helper.
      // Keeping the content directly on the range lets short ranges display
      // their complete labels while preserving the range as the click target.
      if (content && content.parentElement !== range) {
        range.appendChild(content);
      }

      if (!range.querySelector(":scope > .range-click-target")) {
        const clickTarget = document.createElement("span");
        clickTarget.className = "range-click-target";
        clickTarget.setAttribute("aria-hidden", "true");
        range.appendChild(clickTarget);
      }
    }
  }
}

function positionDefaultTimelineRangeTitles() {
  unwrapDefaultTimelineRangeContent();

  const trendPanel = el.timelineTop.querySelector(".vis-panel.vis-center");
  const trendRanges = el.timelineTop.querySelectorAll(".vis-item.trend-range");
  if (trendPanel) {
    const panelBounds = trendPanel.getBoundingClientRect();
    const panelLeft = panelBounds.left + 8;
    const panelRight = panelBounds.right - 8;

    for (let i = 0; i < trendRanges.length; i += 1) {
      const range = trendRanges[i];
      const title = range.querySelector(".trend-title-text");
      if (!title) {
        continue;
      }

      title.style.removeProperty("--trend-title-left");
      const titleBounds = title.getBoundingClientRect();
      let targetLeft = Math.max(titleBounds.left, panelLeft);
      if (targetLeft + titleBounds.width > panelRight) {
        targetLeft = Math.max(panelLeft, panelRight - titleBounds.width);
      }

      // Translate from the title's actual rendered location instead of from
      // the range start. Long-running ranges can begin thousands of years
      // outside the current window, so their range-relative coordinates are
      // not a reliable on-screen anchor.
      title.style.setProperty("--trend-title-left", (targetLeft - titleBounds.left) + "px");
    }
  }

  const secondaryCollections = [
    [el.timelineBottom, ".vis-item.paradigm-range"],
    [el.timelineAdjustments, ".vis-item.adjustment-range"]
  ];
  for (let collectionIndex = 0; collectionIndex < secondaryCollections.length; collectionIndex += 1) {
    const root = secondaryCollections[collectionIndex][0];
    const selector = secondaryCollections[collectionIndex][1];
    const panel = root.querySelector(".vis-panel.vis-center");
    const ranges = root.querySelectorAll(selector);
    if (!panel) {
      continue;
    }
    const panelBounds = panel.getBoundingClientRect();
    const panelLeft = panelBounds.left + 8;
    const panelRight = panelBounds.right - 8;

    for (let i = 0; i < ranges.length; i += 1) {
      const range = ranges[i];
      const title = range.querySelector(".paradigm-title-text");
      if (!title) {
        continue;
      }

      title.style.removeProperty("--paradigm-title-left");
      const titleBounds = title.getBoundingClientRect();
      let targetLeft = Math.max(titleBounds.left, panelLeft);
      if (targetLeft + titleBounds.width > panelRight) {
        targetLeft = Math.max(panelLeft, panelRight - titleBounds.width);
      }

      title.style.setProperty("--paradigm-title-left", (targetLeft - titleBounds.left) + "px");
    }
  }
}

function buildAdjustmentOptions(baseOptions, timeAxisStep) {
  const topOptions = buildTopOptions(baseOptions, timeAxisStep);
  return {
    ...topOptions,
    height: (parseFloat(topOptions.height) + ADJUSTMENT_CONTAINER_BOTTOM_PADDING) + "px",
    orientation: {
      axis: "top",
      item: "top"
    },
    showMajorLabels: false,
    showMinorLabels: true,
    margin: {
      item: {
        horizontal: 0,
        vertical: 4
      },
      axis: 14
    },
    template: buildParadigmItemTemplate
  };
}

let defaultRangeLabelLayoutFrame = null;

function scheduleDefaultTimelineRangeTitleLayout() {
  if (defaultRangeLabelLayoutFrame !== null) {
    cancelAnimationFrame(defaultRangeLabelLayoutFrame);
  }
  defaultRangeLabelLayoutFrame = requestAnimationFrame(function () {
    defaultRangeLabelLayoutFrame = null;
    if (!isExportingImage) {
      positionDefaultTimelineRangeTitles();
    }
  });
}

function scheduleTrendEndpointAnchoring() {
  requestAnimationFrame(function () {
    anchorTrendEndpointYears();
    scheduleDefaultTimelineRangeTitleLayout();
  });
}

function scrollTopTimelineToFirstRow() {
  scrollTimelineToFirstRow(trendTimelineInstance, el.timelineTop);
}

function scrollTimelineToFirstRow(timelineInstance, timelineRoot) {
  if (!timelineInstance || !timelineRoot) {
    return;
  }

  // vis-timeline stores vertical position as an internal negative offset;
  // the center panel itself is not a native scrolling element.
  if (typeof timelineInstance._setScrollTop === "function") {
    timelineInstance._setScrollTop(0);
    timelineInstance.redraw();
    return;
  }

  const sideScrollPanel = timelineRoot.querySelector(".vis-panel.vis-left.vis-vertical-scroll");
  if (sideScrollPanel) {
    sideScrollPanel.scrollTop = 0;
  }
}

function scrollAllTimelinesToFirstRow() {
  scrollTimelineToFirstRow(trendTimelineInstance, el.timelineTop);
  scrollTimelineToFirstRow(paradigmTimelineInstance, el.timelineBottom);
  scrollTimelineToFirstRow(adjustmentTimelineInstance, el.timelineAdjustments);
}

function measureTimelineHeightDeficit(root, rangeSelector, bottomClearance) {
  const itemPanel = root.querySelector(".vis-panel.vis-center");
  const ranges = root.querySelectorAll(rangeSelector);
  if (!itemPanel || !ranges.length) {
    return 0;
  }

  const panelBounds = itemPanel.getBoundingClientRect();
  let lastRangeBottom = panelBounds.top;
  for (let i = 0; i < ranges.length; i += 1) {
    lastRangeBottom = Math.max(lastRangeBottom, ranges[i].getBoundingClientRect().bottom);
  }

  return Math.max(Math.ceil(
    lastRangeBottom - panelBounds.bottom + (bottomClearance || 0)
  ), 0);
}

function measureTimelineExportHeightDeficit(root, rangeSelector) {
  // Keep a small margin between the final arrow and the bottom time axis.
  return measureTimelineHeightDeficit(root, rangeSelector, 4);
}

function measureTrendExportHeightDeficit() {
  return measureTimelineExportHeightDeficit(el.timelineTop, ".vis-item.trend-range");
}

function scheduleInitialTopTimelineScroll() {
  requestAnimationFrame(function () {
    requestAnimationFrame(scrollAllTimelinesToFirstRow);
  });
}

function indexTimelineItems(trendItems, paradigmItems, adjustmentItems) {
  trendItemsById = new Map(trendItems.map((item) => [item.id, item]));
  paradigmItemsById = new Map(paradigmItems.map((item) => [item.id, item]));
  adjustmentItemsById = new Map(adjustmentItems.map((item) => [item.id, item]));
  eventsByTrendSubgroup = new Map();
  eventsByParadigm = new Map();
  eventsByAdjustment = new Map();

  for (const item of trendItems) {
    if (item.type !== "point") {
      continue;
    }

    if (!eventsByTrendSubgroup.has(item.subgroup)) {
      eventsByTrendSubgroup.set(item.subgroup, []);
    }
    eventsByTrendSubgroup.get(item.subgroup).push(item);

    const paradigmKey = normalizeParadigmKey(item.paradigm);
    if (paradigmKey) {
      if (!eventsByParadigm.has(paradigmKey)) {
        eventsByParadigm.set(paradigmKey, []);
      }
      eventsByParadigm.get(paradigmKey).push(item);
    }

    const adjustmentKey = normalizeParadigmKey(item.adjustment);
    if (adjustmentKey) {
      if (!eventsByAdjustment.has(adjustmentKey)) {
        eventsByAdjustment.set(adjustmentKey, []);
      }
      eventsByAdjustment.get(adjustmentKey).push(item);
    }
  }
}

function renderTimeline(data) {
  currentTrendItems = data.trendItems;
  currentParadigmItems = data.paradigmItems;
  currentAdjustmentItems = data.adjustmentItems;
  indexTimelineItems(currentTrendItems, currentParadigmItems, currentAdjustmentItems);
  currentTrendOptions = data.trendOptions;
  currentParadigmOptions = data.paradigmOptions;
  currentAdjustmentOptions = data.adjustmentOptions;
  paradigmRenderedHeightCorrection = 0;
  paradigmHeightFitGeneration += 1;
  lastAppliedTimelineHeights = null;
  el.timelineTitle.textContent = data.title;
  el.trendSectionLabel.textContent = TIMELINE_CONFIGS[activeTabKey].label;
  resetTimelineShellHeights();
  const initialHeight = getInitialTimelineHeight();
  applyTimelineShellHeights({
    trendHeight: initialHeight,
    paradigmHeight: initialHeight,
    adjustmentHeight: initialHeight
  });

  const trendItems = new vis.DataSet(data.trendItems);
  const trendGroups = new vis.DataSet(data.trendGroups);
  const paradigmItems = new vis.DataSet(data.paradigmItems);
  const paradigmGroups = new vis.DataSet(data.paradigmGroups);
  const adjustmentItems = new vis.DataSet(data.adjustmentItems);
  const adjustmentGroups = new vis.DataSet(data.adjustmentGroups);

  if (trendTimelineInstance) {
    trendTimelineInstance.destroy();
  }
  if (paradigmTimelineInstance) {
    paradigmTimelineInstance.destroy();
  }
  if (adjustmentTimelineInstance) {
    adjustmentTimelineInstance.destroy();
  }

  try {
    trendTimelineInstance = new vis.Timeline(el.timelineTop, trendItems, trendGroups, buildTopOptions(data.trendOptions, data.timeAxisStep));
  } catch (error) {
    trendTimelineInstance = new vis.Timeline(el.timelineTop, trendItems, trendGroups, buildTopOptions(data.trendOptions, data.timeAxisStep));
  }

  trendTimelineInstance.setWindow(data.trendOptions.start, data.trendOptions.end, { animation: false });
  scheduleInitialTopTimelineScroll();

  attachTrendClickHandler(trendTimelineInstance);
  trendTimelineInstance.on("changed", scheduleTrendEndpointAnchoring);
  trendTimelineInstance.on("rangechanged", function () {
    updateShareUrl();
    scheduleTrendEndpointAnchoring();
  });
  scheduleTrendEndpointAnchoring();

  paradigmTimelineInstance = new vis.Timeline(
    el.timelineBottom,
    paradigmItems,
    paradigmGroups,
    buildBottomOptions(data.paradigmOptions, data.timeAxisStep)
  );

  adjustmentTimelineInstance = new vis.Timeline(
    el.timelineAdjustments,
    adjustmentItems,
    adjustmentGroups,
    buildAdjustmentOptions(data.adjustmentOptions, data.timeAxisStep)
  );

  syncTimelineWindows([trendTimelineInstance, paradigmTimelineInstance, adjustmentTimelineInstance]);

  paradigmTimelineInstance.on("changed", scheduleDefaultTimelineRangeTitleLayout);
  paradigmTimelineInstance.on("rangechanged", scheduleDefaultTimelineRangeTitleLayout);
  adjustmentTimelineInstance.on("changed", scheduleDefaultTimelineRangeTitleLayout);
  adjustmentTimelineInstance.on("rangechanged", scheduleDefaultTimelineRangeTitleLayout);

  attachParadigmClickHandler(paradigmTimelineInstance);
  attachAdjustmentClickHandler(adjustmentTimelineInstance);

  // Keep the initial viewport aligned after creation.
  paradigmTimelineInstance.setWindow(data.trendOptions.start, data.trendOptions.end, { animation: false });
  adjustmentTimelineInstance.setWindow(data.trendOptions.start, data.trendOptions.end, { animation: false });
  scheduleDefaultTimelineRangeTitleLayout();
}

function formatDate(date) {
  if (!(date instanceof Date) || isNaN(date)) {
    return "";
  }
  const y = date.getFullYear();
  const m = date.getMonth();
  const d = date.getDate();
  const hasMonth = m !== 0 || d !== 1;
  const hasDay = d !== 1;
  if (hasDay) {
    return y + "-" + String(m + 1).padStart(2, "0") + "-" + String(d).padStart(2, "0");
  }
  if (hasMonth) {
    return y + "-" + String(m + 1).padStart(2, "0");
  }
  return String(y);
}

function displayStart(item) {
  if (!item) {
    return "";
  }
  if (item.startIsBce && item.startRaw) {
    return item.startRaw;
  }
  return formatDate(item.start);
}

function parseBceYear(raw) {
  if (!raw) {
    return 0;
  }
  const cleaned = String(raw).replace(/[,\s]/g, "");
  const match = cleaned.match(/(\d+)/);
  return match ? parseInt(match[1], 10) : 0;
}

function eventSortKey(item) {
  if (item.startIsBce) {
    return -parseBceYear(item.startRaw);
  }
  if (item.start instanceof Date && !isNaN(item.start)) {
    return item.start.getFullYear()
      + item.start.getMonth() / 12
      + (item.start.getDate() - 1) / 365;
  }
  return 0;
}

function showEventsModal(title, events, emptyMessage) {
  const sortedEvents = events.slice().sort(function (a, b) {
    return eventSortKey(a) - eventSortKey(b);
  });

  let titleHtml = escapeHtml(title);
  if (sortedEvents.length > 0) {
    const minEvt = sortedEvents[0];
    const maxEvt = sortedEvents[sortedEvents.length - 1];
    const minStr = displayStart(minEvt);
    const maxStr = displayStart(maxEvt);
    const dateSpan = minStr !== maxStr
      ? minStr + " – " + maxStr
      : minStr;
    titleHtml += ' <span class="modal-title-meta"> · ' + events.length + ' event' + (events.length !== 1 ? 's' : '') + ' · ' + escapeHtml(dateSpan) + '</span>';
  }

  el.modalTitle.innerHTML = titleHtml;

  if (sortedEvents.length === 0) {
    el.modalBody.innerHTML = '<div class="event-empty">' + escapeHtml(emptyMessage || "No events recorded.") + '</div>';
  } else {
    let html = '<div class="event-timeline">';
    for (let i = 0; i < sortedEvents.length; i += 1) {
      const evt = sortedEvents[i];
      const paradigmHtml = evt.paradigm
        ? '<span class="event-card-badge" style="' + badgeColors(evt.paradigm) + '">' + escapeHtml(evt.paradigm) + '</span>'
        : '';
      const adjustmentHtml = evt.adjustment
        ? '<span class="event-card-badge" style="' + badgeColors(evt.adjustment) + '">' + escapeHtml(evt.adjustment) + '</span>'
        : '';
      const energySourceHtml = evt.energySource
        ? '<span class="event-card-badge" style="' + badgeColors(evt.energySource) + '">' + escapeHtml(evt.energySource) + '</span>'
        : '';
      html += '<div class="event-node">' +
        '<div class="event-card">' +
        '<div class="event-card-header">' +
        '<span class="event-card-date">' + escapeHtml(displayStart(evt)) + '</span>' +
        paradigmHtml +
        adjustmentHtml +
        energySourceHtml +
        '</div>' +
        '<div class="event-card-body">' + escapeHtml(evt.title || evt.content || "") + '</div>' +
        '</div>' +
        '</div>';
    }
    html += '</div>';
    el.modalBody.innerHTML = html;
  }

  el.modalOverlay.classList.remove("hidden");
  document.body.style.overflow = "hidden";
}

function showTrendModal(trendTitle, subgroup) {
  const events = eventsByTrendSubgroup.get(subgroup) || [];
  showEventsModal(trendTitle, events, "No events recorded for this trend.");
}

function showParadigmModal(paradigmTitle) {
  const paradigmKey = normalizeParadigmKey(paradigmTitle);
  const events = eventsByParadigm.get(paradigmKey) || [];
  showEventsModal(paradigmTitle, events, "No events recorded for this paradigm.");
}

function showAdjustmentModal(adjustmentTitle) {
  const adjustmentKey = normalizeParadigmKey(adjustmentTitle);
  const events = eventsByAdjustment.get(adjustmentKey) || [];
  showEventsModal(adjustmentTitle, events, "No events recorded for this adjustment.");
}

function hideModal() {
  el.modalOverlay.classList.add("hidden");
  document.body.style.overflow = "";
}

function escapeHtml(text) {
  const div = document.createElement("div");
  div.appendChild(document.createTextNode(text));
  return div.innerHTML;
}

const BADGE_PALETTE = [
  { color: "#063a3f", bg: "#dceced", border: "#0b5f66" },
  { color: "#1a4630", bg: "#e3f2ea", border: "#2f6f4e" },
  { color: "#6b3420", bg: "#f6e8e0", border: "#a65d3f" },
  { color: "#6b5114", bg: "#f7efd6", border: "#b0892c" },
  { color: "#2c4068", bg: "#e6edf7", border: "#4f6fa8" },
  { color: "#1f4b57", bg: "#dff1f5", border: "#3d7a8c" },
  { color: "#5a3544", bg: "#f4e6eb", border: "#8b5a6b" },
  { color: "#374151", bg: "#eef0f3", border: "#6b7280" }
];

function badgeColors(paradigm) {
  const idx = Math.abs(hashString(paradigm)) % BADGE_PALETTE.length;
  const p = BADGE_PALETTE[idx];
  return "color:" + p.color + ";background:" + p.bg + ";border:1px solid " + p.border;
}

function attachTrendClickHandler(timelineInstance) {
  timelineInstance.on("click", function (props) {
    if (!props.item) {
      return;
    }

    const clickedItem = trendItemsById.get(props.item);

    if (!clickedItem || clickedItem.type !== "range") {
      return;
    }

    showTrendModal(clickedItem.titleText || clickedItem.content, clickedItem.subgroup);
  });
}

function attachParadigmClickHandler(timelineInstance) {
  timelineInstance.on("click", function (props) {
    if (!props.item) {
      return;
    }

    const clickedItem = paradigmItemsById.get(props.item);

    if (clickedItem) {
      showParadigmModal(clickedItem.titleText || clickedItem.content);
    }
  });
}

function attachAdjustmentClickHandler(timelineInstance) {
  timelineInstance.on("click", function (props) {
    if (!props.item) {
      return;
    }

    const clickedItem = adjustmentItemsById.get(props.item);
    if (clickedItem) {
      showAdjustmentModal(clickedItem.titleText || clickedItem.content);
    }
  });
}

el.modalClose.addEventListener("click", (event) => {
  event.stopPropagation();
  hideModal();
});

el.modalOverlay.addEventListener("click", (event) => {
  if (event.target === el.modalOverlay) {
    hideModal();
  }
});

document.addEventListener("keydown", (event) => {
  if (event.key === "Escape") {
    hideModal();
  }
});

function updateShareUrl() {
  const params = new URLSearchParams();
  params.set("tab", activeTabKey || "");

  if (trendTimelineInstance) {
    const w = trendTimelineInstance.getWindow();
    if (w && w.start && w.end) {
      params.set("start", w.start.toISOString().slice(0, 10));
      params.set("end", w.end.toISOString().slice(0, 10));
    }
  }

  const url = window.location.pathname + "#" + params.toString();
  window.history.replaceState(null, "", url);
}

function getInitialTabKey() {
  const hash = window.location.hash.replace(/^#/, "");
  if (hash) {
    const params = new URLSearchParams(hash);
    const tab = params.get("tab");
    const startStr = params.get("start");
    const endStr = params.get("end");
    const start = startStr ? new Date(startStr) : null;
    const end = endStr ? new Date(endStr) : null;
    if (start instanceof Date && !isNaN(start) && end instanceof Date && !isNaN(end) && start <= end) {
      pendingInitialWindow = { start, end };
    }
    if (tab && TIMELINE_CONFIGS[tab]) {
      return tab;
    }
  }
  return ACTIVE_TAB_KEY;
}

function forEachButton(buttons, fn) {
  for (let i = 0; i < buttons.length; i += 1) {
    const btn = buttons[i];
    if (btn) {
      fn(btn, i);
    }
  }
}

function getShareButtons() {
  return [el.shareBtn, el.shareBtnFs, el.shareBtnMobile].filter(Boolean);
}

function getExportButtons() {
  return [el.exportBtn, el.exportBtnFs, el.exportBtnMobile].filter(Boolean);
}

function setButtonText(btn, text) {
  if (!btn) {
    return;
  }
  if (btn.classList.contains("mobile-icon-btn")) {
    return;
  }
  btn.textContent = text;
}

async function shareUrl() {
  updateShareUrl();
  const url = window.location.href;
  const buttons = getShareButtons();
  const originalLabels = buttons.map(function (btn) { return btn.textContent; });
  forEachButton(buttons, function (btn) {
    btn.disabled = true;
  });

  let copied = false;
  try {
    if (navigator.clipboard && window.isSecureContext) {
      await navigator.clipboard.writeText(url);
      copied = true;
    }
  } catch (error) {
    console.warn("Clipboard write failed, falling back:", error);
  }

  if (!copied) {
    const ta = document.createElement("textarea");
    ta.value = url;
    ta.setAttribute("readonly", "");
    ta.style.position = "fixed";
    ta.style.left = "-9999px";
    document.body.appendChild(ta);
    ta.select();
    try {
      copied = document.execCommand("copy");
    } catch (e) {
      copied = false;
    }
    document.body.removeChild(ta);
  }

  forEachButton(buttons, function (btn) {
    setButtonText(btn, copied ? "Link copied" : "Copy failed");
    if (copied) {
      btn.classList.add("success");
    } else {
      btn.classList.remove("success");
    }
  });
  setTimeout(() => {
    forEachButton(buttons, function (btn, idx) {
      setButtonText(btn, originalLabels[idx]);
      btn.classList.remove("success");
      btn.disabled = false;
    });
  }, 2000);
}

function removeInlineStyleProperties(element, propertyNames) {
  for (let i = 0; i < propertyNames.length; i += 1) {
    element.style.removeProperty(propertyNames[i]);
  }
  if (!element.getAttribute("style")) {
    element.removeAttribute("style");
  }
}

function clearExportLayoutStyles() {
  removeInlineStyleProperties(el.timelineStage, [
    "flex", "width", "max-width", "height", "min-height", "padding-bottom", "overflow"
  ]);
  removeInlineStyleProperties(el.timelineChartArea, ["flex", "width", "height"]);
  removeInlineStyleProperties(el.timelineShell, ["flex", "width", "height"]);
  removeInlineStyleProperties(el.paradigmShell, ["flex", "width", "height"]);
  removeInlineStyleProperties(el.adjustmentShell, ["flex", "width", "height"]);
  removeInlineStyleProperties(el.timelineTop, ["width"]);
  removeInlineStyleProperties(el.timelineBottom, ["width"]);
  removeInlineStyleProperties(el.timelineAdjustments, ["width"]);
  removeInlineStyleProperties(el.exportLoadingOverlay, [
    "position", "top", "left", "width", "height"
  ]);
}

function setExportWidths(widths) {
  const widthTargets = {
    stage: el.timelineStage,
    chartArea: el.timelineChartArea,
    timelineShell: el.timelineShell,
    paradigmShell: el.paradigmShell,
    adjustmentShell: el.adjustmentShell,
    timelineTop: el.timelineTop,
    timelineBottom: el.timelineBottom,
    timelineAdjustments: el.timelineAdjustments
  };

  for (const [key, element] of Object.entries(widthTargets)) {
    element.style.width = widths[key] + "px";
  }
}

function addExportSecondaryTimelineGridLines(clonedDocument) {
  const energyTimeline = clonedDocument.getElementById("timelineTop");
  if (!energyTimeline) {
    return;
  }

  const energyPanel = energyTimeline.querySelector(".vis-panel.vis-center");
  if (!energyPanel) {
    return;
  }

  const energyPanelBounds = energyPanel.getBoundingClientRect();
  if (energyPanelBounds.width <= 0) {
    return;
  }

  const linePositions = [];
  const seenPositions = new Set();
  const energyGridLines = energyTimeline.querySelectorAll(".vis-grid.vis-vertical");
  for (let lineIndex = 0; lineIndex < energyGridLines.length; lineIndex += 1) {
    const lineBounds = energyGridLines[lineIndex].getBoundingClientRect();
    const ratio = (lineBounds.left - energyPanelBounds.left) / energyPanelBounds.width;
    if (ratio < -0.001 || ratio > 1.001) {
      continue;
    }

    const normalizedRatio = Math.min(Math.max(ratio, 0), 1);
    const positionKey = normalizedRatio.toFixed(6);
    if (!seenPositions.has(positionKey)) {
      seenPositions.add(positionKey);
      linePositions.push(normalizedRatio);
    }
  }

  if (!linePositions.length) {
    return;
  }

  ["timelineBottom", "timelineAdjustments"].forEach(function (timelineId) {
    const timeline = clonedDocument.getElementById(timelineId);
    const panel = timeline && timeline.querySelector(".vis-panel.vis-center");
    if (!panel) {
      return;
    }

    const gridLayer = clonedDocument.createElement("div");
    gridLayer.className = "export-secondary-time-grid";
    gridLayer.setAttribute("aria-hidden", "true");
    for (let positionIndex = 0; positionIndex < linePositions.length; positionIndex += 1) {
      const gridLine = clonedDocument.createElement("span");
      gridLine.className = "export-secondary-time-grid-line";
      gridLine.style.left = (linePositions[positionIndex] * 100) + "%";
      gridLayer.appendChild(gridLine);
    }
    panel.appendChild(gridLayer);
  });
}

function neutralizeExportBackgroundImages(clonedDocument) {
  if (!clonedDocument.head) {
    return;
  }

  // html2canvas rasterizes every CSS gradient into an intermediate canvas
  // and calls createPattern with it. vis-timeline can produce zero-width
  // range/background boxes, making that intermediate canvas invalid. The
  // meaningful timeline gradients have export-safe CSS/DOM replacements,
  // so disable background images throughout the detached export document.
  const exportSafetyStyle = clonedDocument.createElement("style");
  exportSafetyStyle.textContent = [
    "html, body, #timelineStage, #timelineStage * { background-image: none !important; }",
    "#timelineStage *::before, #timelineStage *::after { background-image: none !important; }"
  ].join("\n");
  clonedDocument.head.appendChild(exportSafetyStyle);
  addExportSecondaryTimelineGridLines(clonedDocument);

  const clonedTrendTimeline = clonedDocument.getElementById("timelineTop");
  if (clonedTrendTimeline) {
    const trendRanges = clonedTrendTimeline.querySelectorAll(".vis-item.trend-range");
    for (let i = 0; i < trendRanges.length; i += 1) {
      const range = trendRanges[i];
      const exportLine = range.querySelector(".trend-export-line");
      if (!exportLine) {
        continue;
      }

      // The export line starts inside the label template. Once vis-timeline
      // reparents that content, width:100% may resolve against the item set
      // instead of the dated range. Anchor it directly to the native range
      // and stretch between that range's own left and right edges.
      range.appendChild(exportLine);
      exportLine.style.setProperty("position", "absolute", "important");
      exportLine.style.setProperty("left", "0", "important");
      exportLine.style.setProperty("right", "0", "important");
      exportLine.style.setProperty("width", "auto", "important");
      exportLine.style.setProperty("transform", "none", "important");

      if (range.classList.contains("arrow-right")) {
        const exportArrowhead = clonedDocument.createElement("span");
        exportArrowhead.className = "range-export-arrowhead trend-export-arrowhead";
        exportArrowhead.setAttribute("aria-hidden", "true");
        range.appendChild(exportArrowhead);
      }
    }
  }

  const secondaryTimelineDescriptors = [
    { id: "timelineBottom", selector: ".vis-item.paradigm-range", options: currentParadigmOptions },
    { id: "timelineAdjustments", selector: ".vis-item.adjustment-range", options: currentAdjustmentOptions }
  ];
  for (let timelineIndex = 0; timelineIndex < secondaryTimelineDescriptors.length; timelineIndex += 1) {
  const descriptor = secondaryTimelineDescriptors[timelineIndex];
  const clonedTimeline = clonedDocument.getElementById(descriptor.id);
  if (!clonedTimeline) {
    continue;
  }

  const itemPanel = clonedTimeline.querySelector(".vis-panel.vis-center") || clonedTimeline;
  const panelBounds = itemPanel.getBoundingClientRect();
  const exportWindowStartMs = new Date(
    activeExportWindow ? activeExportWindow.start : descriptor.options.start
  ).getTime();
  const exportWindowEndMs = new Date(
    activeExportWindow ? activeExportWindow.end : descriptor.options.end
  ).getTime();
  const exportWindowSpanMs = Math.max(exportWindowEndMs - exportWindowStartMs, 1);
  const paradigmRangeCandidates = Array.from(clonedTimeline.querySelectorAll(descriptor.selector));
  const paradigmRanges = [];
  const seenParadigmItemIds = new Set();
  for (let candidateIndex = 0; candidateIndex < paradigmRangeCandidates.length; candidateIndex += 1) {
    const candidate = paradigmRangeCandidates[candidateIndex];
    const identityNode = candidate.querySelector(".paradigm-title-text[data-export-item-id]");
    const itemId = identityNode ? identityNode.dataset.exportItemId : "";
    if (itemId && seenParadigmItemIds.has(itemId)) {
      candidate.remove();
      continue;
    }
    if (itemId) {
      seenParadigmItemIds.add(itemId);
    }
    paradigmRanges.push(candidate);
  }

  for (let i = 0; i < paradigmRanges.length; i += 1) {
    const range = paradigmRanges[i];
    const sourceTitle = range.querySelector(".paradigm-title-text");
    if (!sourceTitle) {
      continue;
    }
    const itemStartMs = Number(sourceTitle.dataset.exportStartMs);
    const itemEndMs = Number(sourceTitle.dataset.exportEndMs);

    // Do not reuse vis-timeline's cloned content node. Its computed paint
    // layers can survive DOM cleanup and appear as a translucent rectangle.
    // A fresh text-only label leaves the native outer range as the sole box.
    const content = clonedDocument.createElement("span");
    content.className = "paradigm-export-label";
    content.textContent = sourceTitle.textContent;
    const exportLine = clonedDocument.createElement("span");
    exportLine.className = "trend-export-line";
    exportLine.setAttribute("aria-hidden", "true");
    range.replaceChildren(exportLine, content);
    if (range.classList.contains("dot-left")) {
      const startDot = clonedDocument.createElement("span");
      startDot.className = "paradigm-export-start-dot";
      startDot.setAttribute("aria-hidden", "true");
      range.appendChild(startDot);
    }
    if (range.classList.contains("dot-right")) {
      const endDot = clonedDocument.createElement("span");
      endDot.className = "paradigm-export-end-dot";
      endDot.setAttribute("aria-hidden", "true");
      range.appendChild(endDot);
    }
    if (range.classList.contains("arrow-right")) {
      const arrowhead = clonedDocument.createElement("span");
      arrowhead.className = "range-export-arrowhead paradigm-export-arrowhead";
      arrowhead.setAttribute("aria-hidden", "true");
      range.appendChild(arrowhead);
    }
    content.style.setProperty("position", "absolute", "important");
    content.style.setProperty("display", "block", "important");
    content.style.setProperty("top", "auto", "important");
    content.style.setProperty("bottom", "18px", "important");
    content.style.setProperty("left", "20px", "important");
    content.style.setProperty("width", "max-content", "important");
    content.style.setProperty("min-width", "0", "important");
    content.style.setProperty("max-width", "none", "important");
    content.style.setProperty("min-height", "0", "important");
    content.style.setProperty("padding", "0", "important");
    content.style.setProperty("white-space", "nowrap", "important");
    content.style.setProperty("background", "transparent", "important");
    content.style.setProperty("border", "0", "important");
    content.style.setProperty("box-shadow", "none", "important");
    content.style.setProperty("transform", "none", "important");
    content.style.setProperty("z-index", "2", "important");

    let rangeBounds = range.getBoundingClientRect();
    const cloneView = clonedDocument.defaultView;
    const computedRangeStyle = cloneView ? cloneView.getComputedStyle(range) : range.style;
    const currentLeft = parseFloat(computedRangeStyle.left) || 0;

    if (Number.isFinite(itemStartMs) && Number.isFinite(itemEndMs) && itemEndMs >= exportWindowStartMs) {
      // Recompute the cloned native range from its real dates. This handles
      // ranges such as Extractivism, whose start is before the export window
      // and whose end sits exactly on the right boundary.
      const visibleStartMs = Math.max(itemStartMs, exportWindowStartMs);
      const visibleEndMs = Math.min(Math.max(itemEndMs, visibleStartMs), exportWindowEndMs);
      const innerPanelLeft = panelBounds.left + 1;
      const innerPanelRight = panelBounds.right - 1;
      const innerPanelWidth = Math.max(innerPanelRight - innerPanelLeft, 1);
      let desiredLeft = innerPanelLeft + ((visibleStartMs - exportWindowStartMs) / exportWindowSpanMs) * innerPanelWidth;
      const desiredRight = innerPanelLeft + ((visibleEndMs - exportWindowStartMs) / exportWindowSpanMs) * innerPanelWidth;
      const desiredWidth = Math.max(desiredRight - desiredLeft, 4);
      desiredLeft = Math.min(desiredLeft, innerPanelRight - desiredWidth);

      range.style.setProperty("left", (currentLeft + desiredLeft - rangeBounds.left) + "px", "important");
      range.style.setProperty("width", desiredWidth + "px", "important");
      rangeBounds = range.getBoundingClientRect();
    } else {
      const clippedLeft = Math.max(panelBounds.left + 1 - rangeBounds.left, 0);
      const clippedRight = Math.max(rangeBounds.right - (panelBounds.right - 1), 0);
      if (clippedLeft > 0 || clippedRight > 0) {
        // Fallback for older items without export metadata.
        const visibleWidth = Math.max(rangeBounds.width - clippedLeft - clippedRight, 1);
        range.style.setProperty("left", (currentLeft + clippedLeft) + "px", "important");
        range.style.setProperty("width", visibleWidth + "px", "important");
        rangeBounds = range.getBoundingClientRect();
      }
    }

    const contentBounds = content.getBoundingClientRect();
    const horizontalPadding = 10;
    const panelLeft = panelBounds.left + horizontalPadding;
    const panelRight = panelBounds.right - horizontalPadding;
    let labelLeft = Math.max(rangeBounds.left, panelLeft);
    if (labelLeft + contentBounds.width > panelRight) {
      labelLeft = Math.max(panelLeft, panelRight - contentBounds.width);
    }

    const labelTop = Math.max(0, panelBounds.top - rangeBounds.top);
    content.style.setProperty("left", (labelLeft - rangeBounds.left) + "px", "important");
    content.style.setProperty("top", labelTop + "px", "important");
  }
  }
}

async function exportTimelineAsJpeg() {
  const buttons = getExportButtons();
  const originalLabels = buttons.map(function (btn) { return btn.textContent; });
  let exportSucceeded = false;
  const originalChartScrollTop = el.timelineChartArea.scrollTop;
  forEachButton(buttons, function (btn) {
    btn.disabled = true;
    setButtonText(btn, "Exporting…");
  });
  el.exportLoadingOverlay.classList.remove("hidden");
  el.timelineStage.setAttribute("aria-busy", "true");
  try {
    await nextAnimationFrame();

    if (!trendTimelineInstance || !paradigmTimelineInstance || !adjustmentTimelineInstance ||
        !currentTrendOptions || !currentParadigmOptions || !currentAdjustmentOptions) {
      throw new Error("Timeline is not ready to export.");
    }

    const originalStageRect = el.timelineStage.getBoundingClientRect();
    const isMobileExport = window.innerWidth <= 768;
    const targetExportWidth = 1920;
    const exportRailWidth = 72;
    const exportTimelineWidth = targetExportWidth - exportRailWidth;
    const exportWidths = {
      stage: targetExportWidth,
      chartArea: targetExportWidth,
      timelineShell: exportTimelineWidth,
      paradigmShell: exportTimelineWidth,
      adjustmentShell: exportTimelineWidth,
      timelineTop: exportTimelineWidth,
      timelineBottom: exportTimelineWidth,
      timelineAdjustments: exportTimelineWidth
    };
    el.exportLoadingOverlay.style.position = "fixed";
    el.exportLoadingOverlay.style.top = originalStageRect.top + "px";
    el.exportLoadingOverlay.style.left = originalStageRect.left + "px";
    el.exportLoadingOverlay.style.width = originalStageRect.width + "px";
    el.exportLoadingOverlay.style.height = originalStageRect.height + "px";
    isExportingImage = true;
    el.timelineStage.classList.add("is-exporting-image");
    el.timelineStage.classList.toggle("is-mobile-export", isMobileExport);

    const trendRangeCount = currentTrendItems.filter(function (item) {
      return item.type === "range";
    }).length;
    // The export axis is 52px high with a 54px rail band. Include that full
    // band so the final top-timeline row cannot be laid out behind the axis.
    let fullTrendHeight = Math.max(
      trendRangeCount * EXPORT_TREND_ROW_HEIGHT + EXPORT_TIME_AXIS_HEIGHT,
      78
    );
    const paradigmRangeCount = currentParadigmItems.filter(function (item) {
      return item.type === "range";
    }).length;
    let fullParadigmHeight = Math.max(paradigmRangeCount * 52, 76);
    const adjustmentRangeCount = currentAdjustmentItems.filter(function (item) {
      return item.type === "range";
    }).length;
    let fullAdjustmentHeight = Math.max(
      adjustmentRangeCount * EXPORT_TREND_ROW_HEIGHT + EXPORT_TIME_AXIS_HEIGHT,
      78
    );
    let fullChartHeight = fullTrendHeight + fullParadigmHeight + fullAdjustmentHeight
      + TIMELINE_CONTAINER_TOP_PADDING * 3
      + PARADIGM_CONTAINER_BOTTOM_PADDING
      + ADJUSTMENT_CONTAINER_BOTTOM_PADDING;
    const titleHeight = Math.ceil(el.timelineTitle.getBoundingClientRect().height);

    el.timelineLabelRail.style.setProperty("--paradigm-label-height", (fullParadigmHeight + TIMELINE_CONTAINER_TOP_PADDING + PARADIGM_CONTAINER_BOTTOM_PADDING) + "px");
    el.timelineLabelRail.style.setProperty("--trend-label-height", (fullTrendHeight + TIMELINE_CONTAINER_TOP_PADDING) + "px");
    el.timelineLabelRail.style.setProperty("--adjustment-label-height", (fullAdjustmentHeight + TIMELINE_CONTAINER_TOP_PADDING + ADJUSTMENT_CONTAINER_BOTTOM_PADDING) + "px");
    el.timelineChartArea.style.setProperty("--chart-content-height", fullChartHeight + "px");

    el.timelineStage.style.flex = "0 0 auto";
    setExportWidths(exportWidths);
    el.timelineStage.style.maxWidth = "none";
    el.timelineStage.style.height = (titleHeight + fullChartHeight) + "px";
    el.timelineStage.style.minHeight = "0";
    el.timelineStage.style.paddingBottom = "0";
    el.timelineStage.style.overflow = "visible";
    el.timelineChartArea.scrollTop = 0;
    el.timelineChartArea.style.flex = "0 0 auto";
    el.timelineChartArea.style.height = fullChartHeight + "px";
    el.timelineShell.style.flex = "0 0 auto";
    el.timelineShell.style.height = (fullTrendHeight + TIMELINE_CONTAINER_TOP_PADDING) + "px";
    el.paradigmShell.style.flex = "0 0 auto";
    el.paradigmShell.style.height = (fullParadigmHeight + TIMELINE_CONTAINER_TOP_PADDING + PARADIGM_CONTAINER_BOTTOM_PADDING) + "px";
    el.adjustmentShell.style.flex = "0 0 auto";
    el.adjustmentShell.style.height = (fullAdjustmentHeight + TIMELINE_CONTAINER_TOP_PADDING + ADJUSTMENT_CONTAINER_BOTTOM_PADDING) + "px";

    const exportStart = new Date(currentParadigmOptions.start);
    const exactExportEnd = new Date(currentParadigmOptions.end);
    const exportEnd = new Date(exactExportEnd);
    exportEnd.setFullYear(exportEnd.getFullYear() + 10);
    activeExportWindow = { start: exportStart, end: exportEnd };

    trendTimelineInstance.setOptions({
      height: (fullTrendHeight + TIMELINE_CONTAINER_TOP_PADDING) + "px",
      verticalScroll: false,
      groupHeightMode: "fitItems",
      min: exportStart,
      max: exportEnd
    });
    paradigmTimelineInstance.setOptions({
      height: (fullParadigmHeight + TIMELINE_CONTAINER_TOP_PADDING + PARADIGM_CONTAINER_BOTTOM_PADDING) + "px",
      verticalScroll: false,
      groupHeightMode: "fitItems",
      min: exportStart,
      max: exportEnd
    });
    adjustmentTimelineInstance.setOptions({
      height: (fullAdjustmentHeight + TIMELINE_CONTAINER_TOP_PADDING + ADJUSTMENT_CONTAINER_BOTTOM_PADDING) + "px",
      verticalScroll: false,
      groupHeightMode: "fitItems",
      min: exportStart,
      max: exportEnd
    });

    // Height changes can make vis-timeline recalculate its root width.
    // Reapply the on-screen widths before the export redraw.
    setExportWidths(exportWidths);

    trendTimelineInstance.redraw();
    paradigmTimelineInstance.redraw();
    adjustmentTimelineInstance.redraw();
    trendTimelineInstance.setWindow(exportStart, exportEnd, { animation: false });
    paradigmTimelineInstance.setWindow(exportStart, exportEnd, { animation: false });
    adjustmentTimelineInstance.setWindow(exportStart, exportEnd, { animation: false });

    await nextAnimationFrame();
    await nextAnimationFrame();

    // vis-timeline can use a slightly larger subgroup pitch than the nominal
    // export row height. Measure the actual rendered rows and expand until the
    // final arrow clears the time-axis panel instead of relying on estimation.
    for (let correctionIndex = 0; correctionIndex < 3; correctionIndex += 1) {
      const trendHeightDeficit = measureTrendExportHeightDeficit();
      const paradigmHeightDeficit = measureTimelineExportHeightDeficit(el.timelineBottom, ".vis-item.paradigm-range");
      const adjustmentHeightDeficit = measureTimelineExportHeightDeficit(el.timelineAdjustments, ".vis-item.adjustment-range");
      if (trendHeightDeficit <= 0 && paradigmHeightDeficit <= 0 && adjustmentHeightDeficit <= 0) {
        break;
      }

      fullTrendHeight += trendHeightDeficit;
      fullParadigmHeight += paradigmHeightDeficit;
      fullAdjustmentHeight += adjustmentHeightDeficit;
      fullChartHeight += trendHeightDeficit + paradigmHeightDeficit + adjustmentHeightDeficit;
      el.timelineStage.style.height = (titleHeight + fullChartHeight) + "px";
      el.timelineChartArea.style.height = fullChartHeight + "px";
      el.timelineChartArea.style.setProperty("--chart-content-height", fullChartHeight + "px");
      el.timelineShell.style.height = (fullTrendHeight + TIMELINE_CONTAINER_TOP_PADDING) + "px";
      el.paradigmShell.style.height = (fullParadigmHeight + TIMELINE_CONTAINER_TOP_PADDING + PARADIGM_CONTAINER_BOTTOM_PADDING) + "px";
      el.adjustmentShell.style.height = (fullAdjustmentHeight + TIMELINE_CONTAINER_TOP_PADDING + ADJUSTMENT_CONTAINER_BOTTOM_PADDING) + "px";
      el.timelineLabelRail.style.setProperty("--trend-label-height", (fullTrendHeight + TIMELINE_CONTAINER_TOP_PADDING) + "px");
      el.timelineLabelRail.style.setProperty("--paradigm-label-height", (fullParadigmHeight + TIMELINE_CONTAINER_TOP_PADDING + PARADIGM_CONTAINER_BOTTOM_PADDING) + "px");
      el.timelineLabelRail.style.setProperty("--adjustment-label-height", (fullAdjustmentHeight + TIMELINE_CONTAINER_TOP_PADDING + ADJUSTMENT_CONTAINER_BOTTOM_PADDING) + "px");
      trendTimelineInstance.setOptions({ height: (fullTrendHeight + TIMELINE_CONTAINER_TOP_PADDING) + "px" });
      paradigmTimelineInstance.setOptions({ height: (fullParadigmHeight + TIMELINE_CONTAINER_TOP_PADDING + PARADIGM_CONTAINER_BOTTOM_PADDING) + "px" });
      adjustmentTimelineInstance.setOptions({ height: (fullAdjustmentHeight + TIMELINE_CONTAINER_TOP_PADDING + ADJUSTMENT_CONTAINER_BOTTOM_PADDING) + "px" });
      trendTimelineInstance.redraw();
      paradigmTimelineInstance.redraw();
      adjustmentTimelineInstance.redraw();
      trendTimelineInstance.setWindow(exportStart, exportEnd, { animation: false });
      paradigmTimelineInstance.setWindow(exportStart, exportEnd, { animation: false });
      adjustmentTimelineInstance.setWindow(exportStart, exportEnd, { animation: false });
      await nextAnimationFrame();
      await nextAnimationFrame();
    }

    // Reset any scroll offset retained from the compact interactive viewport
    // after the expanded export layout settles.
    scrollTopTimelineToFirstRow();
    await nextAnimationFrame();
    await nextAnimationFrame();
    anchorTrendEndpointYears();

    const options = {
      backgroundColor: getComputedStyle(document.body).backgroundColor || "#ffffff",
      scale: 1,
      width: targetExportWidth,
      windowWidth: targetExportWidth,
      useCORS: true,
      logging: false,
      onclone: neutralizeExportBackgroundImages
    };
    const timelineCanvas = await html2canvas(el.timelineStage, options);
    let exportCanvas = timelineCanvas;
    if (timelineCanvas.width !== targetExportWidth) {
      exportCanvas = document.createElement("canvas");
      exportCanvas.width = targetExportWidth;
      exportCanvas.height = Math.round(
        timelineCanvas.height * (targetExportWidth / timelineCanvas.width)
      );
      const exportContext = exportCanvas.getContext("2d");
      exportContext.drawImage(
        timelineCanvas,
        0,
        0,
        timelineCanvas.width,
        timelineCanvas.height,
        0,
        0,
        exportCanvas.width,
        exportCanvas.height
      );
    }
    const jpegUrl = exportCanvas.toDataURL("image/jpeg", 0.92);
    const a = document.createElement("a");
    const stamp = new Date().toISOString().replace(/[:T]/g, "-").slice(0, 19);
    a.href = jpegUrl;
    a.download = "timeline-" + (activeTabKey || "view") + "-" + stamp + ".jpg";
    document.body.appendChild(a);
    a.click();
    document.body.removeChild(a);
    exportSucceeded = true;
  } catch (error) {
    console.error("Export failed:", error);
  } finally {
    try {
      clearExportLayoutStyles();
      isExportingImage = false;
      activeExportWindow = null;
      el.timelineStage.classList.remove("is-exporting-image");
      el.timelineStage.classList.remove("is-mobile-export");

      if (trendTimelineInstance && paradigmTimelineInstance && adjustmentTimelineInstance &&
          currentTrendOptions && currentParadigmOptions && currentAdjustmentOptions) {
        const trendBounds = buildBounds(currentTrendOptions);
        const paradigmBounds = buildBounds(currentParadigmOptions);
        const adjustmentBounds = buildBounds(currentAdjustmentOptions);
        lastAppliedTimelineHeights = null;
        const timelineHeights = calculateTimelineHeights(currentTrendOptions);
        applyTimelineShellHeights(timelineHeights);
        trendTimelineInstance.setOptions({
          ...trendBounds,
          height: (timelineHeights.trendHeight + TIMELINE_CONTAINER_TOP_PADDING) + "px",
          verticalScroll: false,
          groupHeightMode: "fitItems"
        });
        paradigmTimelineInstance.setOptions({
          ...paradigmBounds,
          height: (timelineHeights.paradigmHeight + TIMELINE_CONTAINER_TOP_PADDING + PARADIGM_CONTAINER_BOTTOM_PADDING) + "px",
          verticalScroll: false,
          groupHeightMode: "fitItems"
        });
        adjustmentTimelineInstance.setOptions({
          ...adjustmentBounds,
          height: (timelineHeights.adjustmentHeight + TIMELINE_CONTAINER_TOP_PADDING + ADJUSTMENT_CONTAINER_BOTTOM_PADDING) + "px",
          verticalScroll: false,
          groupHeightMode: "fitItems"
        });
        trendTimelineInstance.redraw();
        paradigmTimelineInstance.redraw();
        adjustmentTimelineInstance.redraw();
        await nextAnimationFrame();
        await nextAnimationFrame();
        resetTimelineZoom(false);
        await nextAnimationFrame();
        anchorTrendEndpointYears();
      }
      el.timelineChartArea.scrollTop = originalChartScrollTop;
    } finally {
      el.exportLoadingOverlay.classList.add("hidden");
      el.timelineStage.removeAttribute("aria-busy");
    }
  }
  forEachButton(buttons, function (btn) {
    setButtonText(btn, exportSucceeded ? "Saved" : "Export failed");
    btn.classList.toggle("success", exportSucceeded);
  });
  setTimeout(() => {
    forEachButton(buttons, function (btn, idx) {
      setButtonText(btn, originalLabels[idx]);
      btn.classList.remove("success");
      btn.disabled = false;
    });
  }, 2000);
}

function activateTab(tabKey) {
  activeTabKey = tabKey;

  const tabButtons = el.tabs.querySelectorAll(".tab");
  for (let i = 0; i < tabButtons.length; i += 1) {
    const t = tabButtons[i];
    const isActive = t.getAttribute("data-tab") === tabKey;
    t.classList.toggle("active", isActive);
    t.setAttribute("aria-selected", String(isActive));
  }

  updateShareUrl();
  loadAndRenderTimeline();
}

function configIsValid(config) {
  return config && config.workbookUrl;
}

async function loadAndRenderTimeline(forceRefresh) {
  const gen = ++loadGeneration;

  const config = TIMELINE_CONFIGS[activeTabKey];
  if (!configIsValid(config)) {
    setErrorState("No data source configured for " + (config ? config.label : activeTabKey));
    return;
  }

  const t0 = performance.now();
  let workbookData;
  let fromCache = false;

  if (forceRefresh) {
    delete tabCache[activeTabKey];
  }

  if (tabCache[activeTabKey]) {
    workbookData = tabCache[activeTabKey].workbookData;
    fromCache = true;
  } else {
    setLoadingState("Loading Google Sheets workbook...");

    try {
      const workbook = await fetchWorkbook(config);
      if (gen !== loadGeneration) {
        return;
      }
      setLoadingState("Parsing workbook sheets...");
      await nextTick();
      workbookData = parseWorkbookData(workbook);
      tabCache[activeTabKey] = { workbookData };
    } catch (error) {
      if (gen !== loadGeneration) {
        return;
      }
      const msg = error && error.message ? error.message : "Unknown error";
      setErrorState("Unable to load timeline: " + msg);
      console.error(error);
      return;
    }
  }

  try {
    if (gen !== loadGeneration) {
      return;
    }
    await nextTick();
    setLoadingState("Preparing timeline data...");

    const data = buildTimelineData(workbookData, config);
    const elapsed = ((performance.now() - t0) / 1000).toFixed(1);

    if (data.trendItems.length + data.paradigmItems.length + data.adjustmentItems.length === 0) {
      setErrorState(config.label + " has no data yet.");
      return;
    }

    if (gen !== loadGeneration) {
      return;
    }
    await nextTick();
    setLoadingState("Rendering timeline...");

    renderTimeline(data);

    const source = fromCache ? " (cached)" : "";
    // const message = "Timeline loaded in " + elapsed + "s" + source + ". Event rows: " + data.eventRowCount +
    //     ". Items: " + (data.trendItems.length + data.paradigmItems.length) + ".";
    setReadyState();
    refreshVisibleTimelines();
  } catch (error) {
    const msg = error && error.message ? error.message : "Unknown error";
    setErrorState("Unable to prepare timeline: " + msg);
    console.error(error);
  } finally {
    if (gen === loadGeneration) {
    }
  }
}

function zoomTimeline(amount) {
  if (!trendTimelineInstance) {
    return;
  }
  if (amount < 0) {
    trendTimelineInstance.zoomOut(Math.abs(amount), { animation: true });
  } else {
    trendTimelineInstance.zoomIn(amount, { animation: true });
  }
}

function resetTimelineZoom(animate) {
  if (!trendTimelineInstance) {
    return;
  }
  trendTimelineInstance.fit({ animation: animate !== false });
}

function updateFullscreenControl(button) {
  if (!button) {
    return;
  }
  button.setAttribute("aria-pressed", String(isFullscreen));
  button.title = isFullscreen ? "Exit full screen" : "Enter full screen";

  const enterIcon = button.querySelector(".fs-icon-enter");
  const exitIcon = button.querySelector(".fs-icon-exit");
  if (enterIcon) {
    enterIcon.classList.toggle("hidden", isFullscreen);
  }
  if (exitIcon) {
    exitIcon.classList.toggle("hidden", !isFullscreen);
  }
}

function updateFullscreenButtonState() {
  [el.fullscreenBtn, el.fullscreenBtnFs, el.fullscreenBtnMobile]
    .forEach(updateFullscreenControl);

  if (el.floatingControls) {
    el.floatingControls.setAttribute("aria-hidden", String(!isFullscreen));
  }
}

function setFullscreenMode(nextFullscreen) {
  if (!el.widgetRoot || isFullscreen === nextFullscreen) {
    return;
  }

  isFullscreen = nextFullscreen;
  el.widgetRoot.classList.toggle("is-fullscreen", isFullscreen);
  el.widgetRoot.classList.remove("fs-enter", "fs-exit");
  el.widgetRoot.classList.add(isFullscreen ? "fs-enter" : "fs-exit");
  updateFullscreenButtonState();
  lastAppliedTimelineHeights = null;

  if (fullscreenTransitionTimer !== null) {
    clearTimeout(fullscreenTransitionTimer);
  }

  requestAnimationFrame(function () {
    scheduleTimelineResize();
    requestAnimationFrame(function () {
      applyTimelineHeights();
      if (trendTimelineInstance) {
        trendTimelineInstance.redraw();
      }
      if (paradigmTimelineInstance) {
        paradigmTimelineInstance.redraw();
      }
      if (adjustmentTimelineInstance) {
        adjustmentTimelineInstance.redraw();
      }
    });
  });

  fullscreenTransitionTimer = setTimeout(function () {
    el.widgetRoot.classList.remove("fs-enter", "fs-exit");
    fullscreenTransitionTimer = null;
    lastAppliedTimelineHeights = null;
    scheduleTimelineResize();
  }, 360);
}

function toggleFullscreenMode() {
  setFullscreenMode(!isFullscreen);
}

function addClickHandler(element, handler) {
  if (element) {
    element.addEventListener("click", handler);
  }
}

function bindControlGroup(controls) {
  addClickHandler(controls.share, shareUrl);
  addClickHandler(controls.export, exportTimelineAsJpeg);
  addClickHandler(controls.zoomIn, function () { zoomTimeline(0.2); });
  addClickHandler(controls.zoomOut, function () { zoomTimeline(-0.2); });
  addClickHandler(controls.resetZoom, resetTimelineZoom);
  addClickHandler(controls.fullscreen, toggleFullscreenMode);
}

el.tabs.addEventListener("click", function (event) {
  const tab = event.target.closest(".tab");
  if (!tab) {
    return;
  }
  const tabKey = tab.getAttribute("data-tab");
  if (tabKey && tabKey !== activeTabKey) {
    activateTab(tabKey);
  }
});

bindControlGroup({
  share: el.shareBtn,
  export: el.exportBtn,
  zoomIn: el.zoomInBtn,
  zoomOut: el.zoomOutBtn,
  resetZoom: el.resetZoomBtn,
  fullscreen: el.fullscreenBtn
});
bindControlGroup({
  share: el.shareBtnFs,
  export: el.exportBtnFs,
  zoomIn: el.zoomInBtnFs,
  zoomOut: el.zoomOutBtnFs,
  resetZoom: el.resetZoomBtnFs,
  fullscreen: el.fullscreenBtnFs
});
bindControlGroup({
  share: el.shareBtnMobile,
  export: el.exportBtnMobile,
  zoomIn: el.zoomInBtnMobile,
  zoomOut: el.zoomOutBtnMobile,
  resetZoom: el.resetZoomBtnMobile,
  fullscreen: el.fullscreenBtnMobile
});

document.addEventListener("keydown", function (event) {
  if (event.key === "Escape") {
    if (isFullscreen && el.modalOverlay.classList.contains("hidden")) {
      setFullscreenMode(false);
    }
  }
});

updateFullscreenButtonState();

let resizeFrame = null;
function scheduleTimelineResize() {
  if (resizeFrame !== null) {
    cancelAnimationFrame(resizeFrame);
  }
  resizeFrame = requestAnimationFrame(function () {
    resizeFrame = null;
    if (isExportingImage || !trendTimelineInstance || !paradigmTimelineInstance || !adjustmentTimelineInstance) {
      return;
    }
    const changed = applyTimelineHeights();
    if (changed) {
      trendTimelineInstance.redraw();
      paradigmTimelineInstance.redraw();
      adjustmentTimelineInstance.redraw();
    }
    scheduleDefaultTimelineRangeTitleLayout();
    scheduleParadigmTimelineContentFit();
  });
}

window.addEventListener("resize", scheduleTimelineResize);

if (typeof ResizeObserver !== "undefined" && el.timelineShell) {
  const shellObserver = new ResizeObserver(function () {
    scheduleTimelineResize();
  });
  shellObserver.observe(el.timelineShell);
  if (el.timelineStage) {
    shellObserver.observe(el.timelineStage);
  }
}

activateTab(getInitialTabKey());

import { ACTIVE_TAB_KEY, TIMELINE_CONFIGS } from "./config.js";
import { normalizeParadigmKey, parseWorkbookData } from "./workbook.js";

const DEFAULT_GROUP_ID = "default";
const PARADIGM_GROUP_ID = "paradigms";
const TREND_ROW_HEIGHT = 34;
const TREND_TIME_AXIS_HEIGHT = 34;

let trendTimelineInstance = null;
let paradigmTimelineInstance = null;
let activeTabKey = null;
let loadGeneration = 0;
const tabCache = {};
let currentTrendItems = [];
let currentParadigmItems = [];
let trendItemsById = new Map();
let paradigmItemsById = new Map();
let eventsByTrendSubgroup = new Map();
let eventsByParadigm = new Map();
let currentTrendOptions = null;
let currentParadigmOptions = null;
let pendingInitialWindow = null;
let isExportingImage = false;
let activeExportWindow = null;

const el = {
  widgetRoot: document.getElementById("widgetRoot"),
  timelineTop: document.getElementById("timelineTop"),
  timelineBottom: document.getElementById("timelineBottom"),
  timelineShell: document.getElementById("timelineShell"),
  paradigmShell: document.getElementById("paradigmShell"),
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

function yearDateInfo(year, timelineStart) {
  const isBce = year < 0;
  return {
    date: isBce ? new Date(timelineStart) : createYearDate(year),
    isBce,
    raw: isBce ? Math.abs(year) + " BCE" : String(year)
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
  return compareSubgroupsByDuration(second, first);
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
    paradigmItems: []
  };
  const paradigms = new Map();
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

    const rangeStart = yearDateInfo(minYear, timelineStart);
    const rangeEnd = yearDateInfo(maxYear, timelineStart);
    const endpointClasses = rangeStart.date.getTime() === timelineStart.getTime()
      ? " arrow-right"
      : " arrow-right dot-left";
    timeline.trendItems.push({
      id: "trend-" + trendCounter++,
      group: DEFAULT_GROUP_ID,
      subgroup,
      type: "range",
      className: "trend-range trend-color-" + (Math.abs(hashString(subgroup)) % 8) + endpointClasses,
      start: rangeStart.date,
      end: rangeEnd.date,
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
      const startInfo = yearDateInfo(event.startYear, timelineStart);
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
        energySource: sheetName
      });
      eventRowCount += 1;

      const paradigmKey = normalizeParadigmKey(event.paradigm);
      if (!paradigmKey) {
        continue;
      }
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
  });

  let paradigmCounter = 1;
  paradigms.forEach(function (paradigm, paradigmKey) {
    const startInfo = yearDateInfo(paradigm.minYear, timelineStart);
    const endInfo = yearDateInfo(paradigm.maxYear, timelineStart);
    timeline.paradigmSubgroups.add(paradigmKey);
    timeline.paradigmItems.push({
      id: "paradigm-" + paradigmCounter++,
      group: PARADIGM_GROUP_ID,
      subgroup: paradigmKey,
      type: "range",
      className: "paradigm-range",
      start: startInfo.date,
      end: endInfo.date,
      startRaw: startInfo.raw,
      endRaw: endInfo.raw,
      startIsBce: startInfo.isBce,
      endIsBce: endInfo.isBce,
      titleText: paradigm.title,
      displayText: paradigm.title + " (" + startInfo.raw + "-" + endInfo.raw + ")",
      content: escapeHtml(paradigm.title) + " (" + escapeHtml(startInfo.raw) + "-" + escapeHtml(endInfo.raw) + ")"
    });
  });

  const span = timelineEnd.getTime() - timelineStart.getTime();
  const bufferFactor = window.innerWidth <= 768 ? 0.02 : 0.01;
  const endWithBuffer = new Date(timelineEnd.getTime() + span * bufferFactor);

  const orderedTrendSubgroups = orderSubgroupsByDuration(
    timeline.trendSubgroups,
    timeline.trendItems
  );
  const orderedParadigmSubgroups = orderSubgroupsByDuration(
    timeline.paradigmSubgroups,
    timeline.paradigmItems
  );

  return {
    title: timeline.title,
    // The top timeline is bottom-oriented, so reverse its subgroup comparator
    // to make the desired order read correctly from top to bottom on screen.
    trendGroups: [createGroupWithSubgroups(
      DEFAULT_GROUP_ID,
      orderedTrendSubgroups,
      compareTopTimelineSubgroups
    )],
    paradigmGroups: [createGroupWithSubgroups(PARADIGM_GROUP_ID, orderedParadigmSubgroups)],
    trendItems: timeline.trendItems,
    paradigmItems: timeline.paradigmItems,
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
      end: timelineEnd
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
  const availableHeight = getAvailableTimelineHeight();
  const requiredTrendHeight = Math.max(
    Math.round(trendOptions && trendOptions.height
      ? trendOptions.height
      : TREND_ROW_HEIGHT + TREND_TIME_AXIS_HEIGHT),
    TREND_ROW_HEIGHT
  );

  if (requiredTrendHeight * 1.5 > availableHeight) {
    const equalHeight = Math.floor(availableHeight / 2);
    return {
      trendHeight: equalHeight,
      paradigmHeight: equalHeight
    };
  }

  return {
    trendHeight: requiredTrendHeight,
    paradigmHeight: availableHeight - requiredTrendHeight
  };
}

function getInitialTimelineHeight() {
  return Math.max(Math.floor(getAvailableTimelineHeight() / 2), 1);
}

function applyTimelineShellHeights(heights) {
  el.timelineShell.style.flex = "0 0 " + heights.trendHeight + "px";
  el.timelineShell.style.height = heights.trendHeight + "px";
  el.paradigmShell.style.flex = "0 0 " + heights.paradigmHeight + "px";
  el.paradigmShell.style.height = heights.paradigmHeight + "px";
}

function resetTimelineShellHeights() {
  el.timelineShell.style.removeProperty("flex");
  el.timelineShell.style.removeProperty("height");
  el.paradigmShell.style.removeProperty("flex");
  el.paradigmShell.style.removeProperty("height");
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
  wrapper.textContent = item.displayText || item.titleText || item.content || "";
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
    height: initialHeight + "px",
    zoomMin: 1000 * 60 * 60 * 24 * 365,
    orientation: "bottom",
    verticalScroll: true,
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
    groupHeightMode: "fixed",
    showCurrentTime: false,
    template: buildTrendItemTemplate,
    onInitialDrawComplete: scheduleInitialTopTimelineScroll,
  };
}

function buildBottomOptions(baseOptions) {
  const bounds = buildBounds(baseOptions);
  const initialHeight = getInitialTimelineHeight();
  return {
    ...baseOptions,
    ...bounds,
    height: initialHeight + "px",
    orientation: "top",
    verticalScroll: true,
    zoomable: true,
    zoomKey: "ctrlKey",
    moveable: true,
    selectable: true,
    groupHeightMode: "fixed",
    margin: {
      item: {
        horizontal: 0,
        vertical: 4
      },
      axis: 0
    },
    showCurrentTime: false,
    template: buildParadigmItemTemplate,
    // timeAxis: {
    //   scale: "year",
    //   step: 1
    // }
  };
}

let lastAppliedTimelineHeights = null;

function applyTimelineHeights() {
  // Export temporarily expands both timelines to their complete content
  // heights. A ResizeObserver fires during that expansion, so responsive
  // sizing must not overwrite the export-only dimensions.
  if (isExportingImage || !trendTimelineInstance || !paradigmTimelineInstance) {
    return false;
  }
  const heights = calculateTimelineHeights(currentTrendOptions);
  const heightKey = heights.trendHeight + ":" + heights.paradigmHeight;
  if (lastAppliedTimelineHeights === heightKey) {
    return false;
  }
  lastAppliedTimelineHeights = heightKey;
  applyTimelineShellHeights(heights);
  trendTimelineInstance.setOptions({ height: heights.trendHeight + "px" });
  paradigmTimelineInstance.setOptions({ height: heights.paradigmHeight + "px" });
  el.timelineLabelRail.style.setProperty("--paradigm-label-height", heights.paradigmHeight + "px");
  return true;
}

function syncTimelineWindowsBidirectional(first, second) {
  let isSyncing = false;

  function wire(source, target) {
    source.on("rangechange", (props) => {
      if (isSyncing || isExportingImage) {
        return;
      }

      isSyncing = true;
      target.setWindow(props.start, props.end, { animation: false });
      isSyncing = false;
    });
  }

  wire(first, second);
  wire(second, first);
 }

function refreshVisibleTimelines() {
  if (isExportingImage || !trendTimelineInstance || !paradigmTimelineInstance) {
    return;
  }

  // Timelines are created while containers are hidden; redraw once visible so axis/layout paints correctly.
  requestAnimationFrame(() => {
    // First paint both timelines using the equal heights they were initialized
    // with. Adaptive sizing is applied only after vis has a valid visible DOM.
    trendTimelineInstance.redraw();
    paradigmTimelineInstance.redraw();

    requestAnimationFrame(() => {
      applyTimelineHeights();
      trendTimelineInstance.redraw();
      paradigmTimelineInstance.redraw();

      if (pendingInitialWindow) {
        trendTimelineInstance.setWindow(pendingInitialWindow.start, pendingInitialWindow.end, { animation: false });
        pendingInitialWindow = null;
      } else if (currentTrendOptions) {
        trendTimelineInstance.setWindow(currentTrendOptions.start, currentTrendOptions.end, { animation: false });
      }

      const topWindow = trendTimelineInstance.getWindow();
      paradigmTimelineInstance.setWindow(topWindow.start, topWindow.end, { animation: false });
      scheduleInitialTopTimelineScroll();
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
    [el.timelineBottom, ".vis-item.paradigm-range"]
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

  const paradigmPanel = el.timelineBottom.querySelector(".vis-panel.vis-center");
  const paradigmRanges = el.timelineBottom.querySelectorAll(".vis-item.paradigm-range");
  if (paradigmPanel) {
    const panelBounds = paradigmPanel.getBoundingClientRect();
    const panelLeft = panelBounds.left + 8;
    const panelRight = panelBounds.right - 8;

    for (let i = 0; i < paradigmRanges.length; i += 1) {
      const range = paradigmRanges[i];
      const content = range.querySelector(".vis-item-content");
      if (!content) {
        continue;
      }

      content.style.removeProperty("--paradigm-title-left");
      const contentBounds = content.getBoundingClientRect();
      let targetLeft = Math.max(contentBounds.left, panelLeft);
      if (targetLeft + contentBounds.width > panelRight) {
        targetLeft = Math.max(panelLeft, panelRight - contentBounds.width);
      }

      content.style.setProperty("--paradigm-title-left", (targetLeft - contentBounds.left) + "px");
    }
  }
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
  if (!trendTimelineInstance) {
    return;
  }

  // vis-timeline stores vertical position as an internal negative offset;
  // the center panel itself is not a native scrolling element.
  if (typeof trendTimelineInstance._setScrollTop === "function") {
    trendTimelineInstance._setScrollTop(0);
    trendTimelineInstance.redraw();
    return;
  }

  const sideScrollPanel = el.timelineTop.querySelector(".vis-panel.vis-left.vis-vertical-scroll");
  if (sideScrollPanel) {
    sideScrollPanel.scrollTop = 0;
  }
}

function scheduleInitialTopTimelineScroll() {
  requestAnimationFrame(function () {
    requestAnimationFrame(scrollTopTimelineToFirstRow);
  });
}

function indexTimelineItems(trendItems, paradigmItems) {
  trendItemsById = new Map(trendItems.map((item) => [item.id, item]));
  paradigmItemsById = new Map(paradigmItems.map((item) => [item.id, item]));
  eventsByTrendSubgroup = new Map();
  eventsByParadigm = new Map();

  for (const item of trendItems) {
    if (item.type !== "point") {
      continue;
    }

    if (!eventsByTrendSubgroup.has(item.subgroup)) {
      eventsByTrendSubgroup.set(item.subgroup, []);
    }
    eventsByTrendSubgroup.get(item.subgroup).push(item);

    const paradigmKey = normalizeParadigmKey(item.paradigm);
    if (!paradigmKey) {
      continue;
    }
    if (!eventsByParadigm.has(paradigmKey)) {
      eventsByParadigm.set(paradigmKey, []);
    }
    eventsByParadigm.get(paradigmKey).push(item);
  }
}

function renderTimeline(data) {
  currentTrendItems = data.trendItems;
  currentParadigmItems = data.paradigmItems;
  indexTimelineItems(currentTrendItems, currentParadigmItems);
  currentTrendOptions = data.trendOptions;
  currentParadigmOptions = data.paradigmOptions;
  lastAppliedTimelineHeights = null;
  el.timelineTitle.textContent = data.title;
  el.trendSectionLabel.textContent = TIMELINE_CONFIGS[activeTabKey].label;
  resetTimelineShellHeights();
  const initialHeight = getInitialTimelineHeight();
  el.timelineLabelRail.style.setProperty(
    "--paradigm-label-height",
    initialHeight + "px"
  );

  const trendItems = new vis.DataSet(data.trendItems);
  const trendGroups = new vis.DataSet(data.trendGroups);
  const paradigmItems = new vis.DataSet(data.paradigmItems);
  const paradigmGroups = new vis.DataSet(data.paradigmGroups);

  if (trendTimelineInstance) {
    trendTimelineInstance.destroy();
  }
  if (paradigmTimelineInstance) {
    paradigmTimelineInstance.destroy();
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
    buildBottomOptions(data.paradigmOptions)
  );

  syncTimelineWindowsBidirectional(trendTimelineInstance, paradigmTimelineInstance);

  paradigmTimelineInstance.on("changed", scheduleDefaultTimelineRangeTitleLayout);
  paradigmTimelineInstance.on("rangechanged", scheduleDefaultTimelineRangeTitleLayout);

  attachParadigmClickHandler(paradigmTimelineInstance);

  // Keep the initial viewport aligned after creation.
  paradigmTimelineInstance.setWindow(data.trendOptions.start, data.trendOptions.end, { animation: false });
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
      const energySourceHtml = evt.energySource
        ? '<span class="event-card-badge" style="' + badgeColors(evt.energySource) + '">' + escapeHtml(evt.energySource) + '</span>'
        : '';
      html += '<div class="event-node">' +
        '<div class="event-card">' +
        '<div class="event-card-header">' +
        '<span class="event-card-date">' + escapeHtml(displayStart(evt)) + '</span>' +
        paradigmHtml +
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
  removeInlineStyleProperties(el.timelineTop, ["width"]);
  removeInlineStyleProperties(el.timelineBottom, ["width"]);
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
    timelineTop: el.timelineTop,
    timelineBottom: el.timelineBottom
  };

  for (const [key, element] of Object.entries(widthTargets)) {
    element.style.width = widths[key] + "px";
  }
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
    }
  }

  const clonedTimeline = clonedDocument.getElementById("timelineBottom");
  if (!clonedTimeline) {
    return;
  }

  const itemPanel = clonedTimeline.querySelector(".vis-panel.vis-center") || clonedTimeline;
  const panelBounds = itemPanel.getBoundingClientRect();
  const exportWindowStartMs = new Date(
    activeExportWindow ? activeExportWindow.start : currentParadigmOptions.start
  ).getTime();
  const exportWindowEndMs = new Date(
    activeExportWindow ? activeExportWindow.end : currentParadigmOptions.end
  ).getTime();
  const exportWindowSpanMs = Math.max(exportWindowEndMs - exportWindowStartMs, 1);
  const paradigmRangeCandidates = Array.from(clonedTimeline.querySelectorAll(".vis-item.paradigm-range"));
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
    range.replaceChildren(content);
    content.style.setProperty("position", "absolute", "important");
    content.style.setProperty("display", "flex", "important");
    content.style.setProperty("align-items", "center", "important");
    content.style.setProperty("top", "0", "important");
    content.style.setProperty("left", "0", "important");
    content.style.setProperty("width", "max-content", "important");
    content.style.setProperty("min-width", "0", "important");
    content.style.setProperty("max-width", "none", "important");
    content.style.setProperty("min-height", "48px", "important");
    content.style.setProperty("padding", "6px 10px", "important");
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

async function exportTimelineAsJpeg() {
  const buttons = getExportButtons();
  const originalLabels = buttons.map(function (btn) { return btn.textContent; });
  let exportSucceeded = false;
  forEachButton(buttons, function (btn) {
    btn.disabled = true;
    setButtonText(btn, "Exporting…");
  });
  el.exportLoadingOverlay.classList.remove("hidden");
  el.timelineStage.setAttribute("aria-busy", "true");
  try {
    await nextAnimationFrame();

    if (!trendTimelineInstance || !paradigmTimelineInstance || !currentTrendOptions || !currentParadigmOptions) {
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
      timelineTop: exportTimelineWidth,
      timelineBottom: exportTimelineWidth
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
    const fullTrendHeight = Math.max(trendRangeCount * 52 + 80, 78);
    const paradigmRangeCount = currentParadigmItems.filter(function (item) {
      return item.type === "range";
    }).length;
    const fullParadigmHeight = Math.max(paradigmRangeCount * 52 + 50, 76);
    const fullChartHeight = fullTrendHeight + fullParadigmHeight;
    const titleHeight = Math.ceil(el.timelineTitle.getBoundingClientRect().height);

    el.timelineLabelRail.style.setProperty("--paradigm-label-height", fullParadigmHeight + "px");

    el.timelineStage.style.flex = "0 0 auto";
    setExportWidths(exportWidths);
    el.timelineStage.style.maxWidth = "none";
    el.timelineStage.style.height = (titleHeight + fullChartHeight) + "px";
    el.timelineStage.style.minHeight = "0";
    el.timelineStage.style.paddingBottom = "0";
    el.timelineStage.style.overflow = "visible";
    el.timelineChartArea.style.flex = "0 0 auto";
    el.timelineChartArea.style.height = fullChartHeight + "px";
    el.timelineShell.style.flex = "0 0 auto";
    el.timelineShell.style.height = fullTrendHeight + "px";
    el.paradigmShell.style.flex = "0 0 auto";
    el.paradigmShell.style.height = fullParadigmHeight + "px";

    const exportStart = new Date(currentParadigmOptions.start);
    const exactExportEnd = new Date(currentParadigmOptions.end);
    const exportEnd = new Date(exactExportEnd);
    exportEnd.setFullYear(exportEnd.getFullYear() + 10);
    activeExportWindow = { start: exportStart, end: exportEnd };

    trendTimelineInstance.setOptions({
      height: fullTrendHeight + "px",
      verticalScroll: false,
      groupHeightMode: "fitItems",
      min: exportStart,
      max: exportEnd
    });
    paradigmTimelineInstance.setOptions({
      height: fullParadigmHeight + "px",
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
    trendTimelineInstance.setWindow(exportStart, exportEnd, { animation: false });
    paradigmTimelineInstance.setWindow(exportStart, exportEnd, { animation: false });

    await nextAnimationFrame();
    await nextAnimationFrame();
    // Expanding a bottom-oriented timeline can preserve its former negative
    // vertical offset, which pushes the first rows down and leaves blank space
    // above them in the exported image. Reset after the expanded layout has
    // settled, then allow the corrective redraw to finish before capture.
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

      if (trendTimelineInstance && paradigmTimelineInstance && currentTrendOptions && currentParadigmOptions) {
        const trendBounds = buildBounds(currentTrendOptions);
        const paradigmBounds = buildBounds(currentParadigmOptions);
        lastAppliedTimelineHeights = null;
        const timelineHeights = calculateTimelineHeights(currentTrendOptions);
        applyTimelineShellHeights(timelineHeights);
        el.timelineLabelRail.style.setProperty(
          "--paradigm-label-height",
          timelineHeights.paradigmHeight + "px"
        );
        trendTimelineInstance.setOptions({
          ...trendBounds,
          height: timelineHeights.trendHeight + "px",
          verticalScroll: true,
          groupHeightMode: "fixed"
        });
        paradigmTimelineInstance.setOptions({
          ...paradigmBounds,
          height: timelineHeights.paradigmHeight + "px",
          verticalScroll: true,
          groupHeightMode: "fixed"
        });
        trendTimelineInstance.redraw();
        paradigmTimelineInstance.redraw();
        await nextAnimationFrame();
        await nextAnimationFrame();
        resetTimelineZoom(false);
        await nextAnimationFrame();
        anchorTrendEndpointYears();
      }
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

    if (data.trendItems.length + data.paradigmItems.length === 0) {
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
    if (isExportingImage || !trendTimelineInstance || !paradigmTimelineInstance) {
      return;
    }
    const changed = applyTimelineHeights();
    if (changed) {
      trendTimelineInstance.redraw();
      paradigmTimelineInstance.redraw();
    }
    scheduleDefaultTimelineRangeTitleLayout();
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

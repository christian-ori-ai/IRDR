const SETTINGS_DB_NAME = "irdr-mobile";
const SETTINGS_STORE_NAME = "settings";
const RESULTS_DIRECTORY_KEY = "results-directory-handle";

const state = {
  weeks: [],
  activeWeekId: "",
  activeFacility: "",
  currentIndex: 0,
  supportsFolderSave: false,
  resultsDirectoryHandle: null,
  resultsDirectoryName: "",
};

const elements = {
  setupScreen: document.getElementById("setup-screen"),
  countScreen: document.getElementById("count-screen"),
  weekSelect: document.getElementById("week-select"),
  facilityOptions: document.getElementById("facility-options"),
  facilityOptionTemplate: document.getElementById("facility-option-template"),
  startCount: document.getElementById("start-count"),
  summaryLocations: document.getElementById("summary-locations"),
  summaryCompleted: document.getElementById("summary-completed"),
  summaryVariance: document.getElementById("summary-variance"),
  summaryNote: document.getElementById("summary-note"),
  resultsFolderStatus: document.getElementById("results-folder-status"),
  chooseResultsFolder: document.getElementById("choose-results-folder"),
  clearResultsFolder: document.getElementById("clear-results-folder"),
  activeWeek: document.getElementById("active-week"),
  activeFacility: document.getElementById("active-facility"),
  progressLabel: document.getElementById("progress-label"),
  progressStats: document.getElementById("progress-stats"),
  progressFill: document.getElementById("progress-fill"),
  cardBin: document.getElementById("card-bin"),
  cardCostCenter: document.getElementById("card-cost-center"),
  cardOrder: document.getElementById("card-order"),
  cardDescription: document.getElementById("card-description"),
  varianceCount: document.getElementById("variance-count"),
  statusText: document.getElementById("status-text"),
  notesInput: document.getElementById("notes-input"),
  minusButton: document.getElementById("minus-button"),
  plusButton: document.getElementById("plus-button"),
  previousLocation: document.getElementById("previous-location"),
  nextLocation: document.getElementById("next-location"),
  backToSetup: document.getElementById("back-to-setup"),
  exportResults: document.getElementById("export-results"),
  resetSession: document.getElementById("reset-session"),
  sessionNote: document.getElementById("session-note"),
};

document.addEventListener("DOMContentLoaded", init);

async function init() {
  state.supportsFolderSave = supportsDirectorySave();
  registerServiceWorker();
  bindEvents();
  await hydrateResultsDirectory();

  try {
    const response = await fetch("data/samples.json", { cache: "no-store" });
    const payload = await response.json();
    state.weeks = payload.weeks ?? [];
    populateWeekSelect();
  } catch (error) {
    elements.weekSelect.innerHTML = '<option value="">Unable to load samples</option>';
    elements.summaryNote.textContent = "The sample file could not be loaded. Check data/samples.json.";
    console.error(error);
  }
}

function bindEvents() {
  elements.weekSelect.addEventListener("change", () => {
    state.activeWeekId = elements.weekSelect.value;
    state.activeFacility = "";
    renderFacilityOptions();
    updateSelectionSummary();
  });

  elements.startCount.addEventListener("click", startCountSession);
  elements.backToSetup.addEventListener("click", () => setActiveScreen("setup"));
  elements.minusButton.addEventListener("click", () => adjustVariance(-1));
  elements.plusButton.addEventListener("click", () => adjustVariance(1));
  elements.notesInput.addEventListener("input", () => saveCurrentEntry(true));
  elements.previousLocation.addEventListener("click", () => {
    void moveLocation(-1);
  });
  elements.nextLocation.addEventListener("click", () => {
    void moveLocation(1);
  });
  elements.exportResults.addEventListener("click", () => {
    void handleManualExport();
  });
  elements.resetSession.addEventListener("click", resetSession);
  elements.chooseResultsFolder.addEventListener("click", () => {
    void chooseResultsFolder();
  });
  elements.clearResultsFolder.addEventListener("click", () => {
    void clearResultsFolder();
  });
}

async function hydrateResultsDirectory() {
  if (!state.supportsFolderSave) {
    updateResultsFolderUI();
    return;
  }

  try {
    const storedHandle = await settingsGet(RESULTS_DIRECTORY_KEY);
    if (storedHandle) {
      state.resultsDirectoryHandle = storedHandle;
      state.resultsDirectoryName = storedHandle.name ?? "";
    }
  } catch (error) {
    console.warn("Unable to restore results directory handle", error);
  }

  updateResultsFolderUI();
}

function populateWeekSelect() {
  const options = ['<option value="">Select a weekly sample</option>'];
  state.weeks.forEach((week) => {
    options.push(`<option value="${week.id}">${week.label}</option>`);
  });
  elements.weekSelect.innerHTML = options.join("");
  renderFacilityOptions();
  updateSelectionSummary();
}

function renderFacilityOptions() {
  elements.facilityOptions.innerHTML = "";
  const week = getActiveWeek();
  if (!week) {
    elements.startCount.disabled = true;
    return;
  }

  Object.entries(week.facilities ?? {}).forEach(([facilityName, facility]) => {
    const node = elements.facilityOptionTemplate.content.firstElementChild.cloneNode(true);
    const progress = getSessionProgress(week.id, facilityName, facility.locations ?? []);

    node.dataset.facility = facilityName;
    node.querySelector(".facility-card__name").textContent = facilityName;
    node.querySelector(".facility-card__meta").textContent =
      `${facility.sampleSize} locations • ${progress.completed} complete`;
    node.classList.toggle("is-selected", facilityName === state.activeFacility);
    node.addEventListener("click", () => {
      state.activeFacility = facilityName;
      renderFacilityOptions();
      updateSelectionSummary();
    });

    elements.facilityOptions.appendChild(node);
  });

  elements.startCount.disabled = !(state.activeWeekId && state.activeFacility);
}

function updateSelectionSummary() {
  const week = getActiveWeek();
  const facility = getActiveFacility();

  if (!week || !facility) {
    elements.summaryLocations.textContent = "--";
    elements.summaryCompleted.textContent = "--";
    elements.summaryVariance.textContent = "--";
    elements.summaryNote.textContent = week
      ? "Pick a facility to see the current progress."
      : "Choose a week and facility to see the sample details.";
    elements.startCount.disabled = true;
    return;
  }

  const progress = getSessionProgress(week.id, state.activeFacility, facility.locations);
  elements.summaryLocations.textContent = String(facility.sampleSize);
  elements.summaryCompleted.textContent = String(progress.completed);
  elements.summaryVariance.textContent = String(progress.totalVariance);
  elements.summaryNote.textContent =
    `${week.company} • ${facility.population} eligible locations in this facility`;
  elements.startCount.disabled = false;
}

function startCountSession() {
  const week = getActiveWeek();
  const facility = getActiveFacility();
  if (!week || !facility) {
    return;
  }

  const session = loadSession(week.id, state.activeFacility);
  state.currentIndex = clamp(session.currentIndex ?? 0, 0, Math.max((facility.locations?.length ?? 1) - 1, 0));
  setActiveScreen("count");
  renderCountScreen();
}

function setActiveScreen(screenName) {
  const isSetup = screenName === "setup";
  elements.setupScreen.classList.toggle("screen--active", isSetup);
  elements.countScreen.classList.toggle("screen--active", !isSetup);
}

function renderCountScreen() {
  const week = getActiveWeek();
  const facility = getActiveFacility();
  if (!week || !facility || !(facility.locations?.length)) {
    setActiveScreen("setup");
    return;
  }

  const location = facility.locations[state.currentIndex];
  const session = loadSession(week.id, state.activeFacility);
  const entry = session.entries[String(location.sampleOrder)] ?? { varianceCount: 0, notes: "" };
  const progress = getSessionProgress(week.id, state.activeFacility, facility.locations);
  const percentage = ((state.currentIndex + 1) / facility.locations.length) * 100;

  elements.activeWeek.textContent = week.label;
  elements.activeFacility.textContent = state.activeFacility;
  elements.progressLabel.textContent = `Location ${state.currentIndex + 1} of ${facility.locations.length}`;
  elements.progressStats.textContent = `${progress.completed} complete • ${progress.defectLocations} with variance`;
  elements.progressFill.style.width = `${percentage}%`;
  elements.cardBin.textContent = String(location.binCode ?? "");
  elements.cardCostCenter.textContent = String(location.costCenter);
  elements.cardOrder.textContent = String(location.sampleOrder);
  elements.cardDescription.textContent = location.binDescription || "No description";
  elements.varianceCount.textContent = String(entry.varianceCount ?? 0);
  elements.notesInput.value = entry.notes ?? "";
  elements.statusText.textContent = getStatusText(entry.varianceCount ?? 0);
  elements.minusButton.disabled = (entry.varianceCount ?? 0) <= 0;
  elements.previousLocation.disabled = state.currentIndex === 0;
  elements.nextLocation.textContent =
    state.currentIndex === facility.locations.length - 1 ? "Finish Count" : "Next Location";
}

function adjustVariance(delta) {
  const week = getActiveWeek();
  const facility = getActiveFacility();
  if (!week || !facility) {
    return;
  }

  const location = facility.locations[state.currentIndex];
  const session = loadSession(week.id, state.activeFacility);
  const entryKey = String(location.sampleOrder);
  const existing = session.entries[entryKey] ?? { varianceCount: 0, notes: "" };
  const nextValue = Math.max(0, (existing.varianceCount ?? 0) + delta);

  session.entries[entryKey] = {
    ...existing,
    varianceCount: nextValue,
    notes: elements.notesInput.value,
    reviewed: true,
    updatedAt: new Date().toISOString(),
  };
  session.currentIndex = state.currentIndex;
  saveSession(week.id, state.activeFacility, session);
  renderCountScreen();
  updateSelectionSummary();
  renderFacilityOptions();
}

function saveCurrentEntry(markReviewed = false) {
  const week = getActiveWeek();
  const facility = getActiveFacility();
  if (!week || !facility) {
    return;
  }

  const location = facility.locations[state.currentIndex];
  const session = loadSession(week.id, state.activeFacility);
  const entryKey = String(location.sampleOrder);
  const existing = session.entries[entryKey] ?? { varianceCount: 0 };

  session.entries[entryKey] = {
    ...existing,
    notes: elements.notesInput.value,
    varianceCount: existing.varianceCount ?? 0,
    reviewed: markReviewed || Boolean(existing.reviewed),
    updatedAt: new Date().toISOString(),
  };
  session.currentIndex = state.currentIndex;
  saveSession(week.id, state.activeFacility, session);
  updateSelectionSummary();
  renderFacilityOptions();
}

async function moveLocation(direction) {
  saveCurrentEntry(true);

  const week = getActiveWeek();
  const facility = getActiveFacility();
  if (!week || !facility) {
    return;
  }

  if (direction > 0 && state.currentIndex === facility.locations.length - 1) {
    const outcome = await exportResults({
      preferFolder: true,
      interactive: Boolean(state.resultsDirectoryHandle),
    });
    setActiveScreen("setup");
    updateSelectionSummary();
    renderFacilityOptions();
    window.alert(buildExportMessage(outcome, true));
    return;
  }

  state.currentIndex = clamp(state.currentIndex + direction, 0, facility.locations.length - 1);
  const session = loadSession(week.id, state.activeFacility);
  session.currentIndex = state.currentIndex;
  saveSession(week.id, state.activeFacility, session);
  renderCountScreen();
}

async function handleManualExport() {
  const outcome = await exportResults({ preferFolder: true, interactive: true });
  if (outcome) {
    window.alert(buildExportMessage(outcome, false));
  }
}

async function exportResults(options = {}) {
  const week = getActiveWeek();
  const facility = getActiveFacility();
  if (!week || !facility) {
    return null;
  }

  const exportData = buildResultsExport(week, facility);
  const shouldPreferFolder = options.preferFolder !== false;

  if (shouldPreferFolder) {
    const saveOutcome = await saveResultsToDirectory(exportData, options.interactive === true);
    if (saveOutcome) {
      return saveOutcome;
    }
  }

  downloadCsv(exportData.filename, exportData.csvContent);
  return {
    method: "download",
    filename: exportData.filename,
    reason: state.supportsFolderSave ? "download-fallback" : "unsupported",
  };
}

function buildResultsExport(week, facility) {
  const session = loadSession(week.id, state.activeFacility);
  const progress = getSessionProgress(week.id, state.activeFacility, facility.locations);
  const exportedAt = new Date().toISOString();
  const filename = buildResultsFileName(week.id, state.activeFacility, exportedAt);
  const lines = [
    [
      "company",
      "week",
      "facility",
      "source_file",
      "exported_at",
      "facility_population",
      "sample_size",
      "completed_locations",
      "defect_locations",
      "total_variance_cases",
      "sample_order",
      "random_draw_order",
      "cost_center",
      "bin_code",
      "bin_description",
      "reviewed",
      "variance_count",
      "notes",
    ],
  ];

  facility.locations.forEach((location) => {
    const entry = session.entries[String(location.sampleOrder)] ?? { varianceCount: 0, notes: "" };
    lines.push([
      week.company ?? "",
      week.id,
      state.activeFacility,
      week.sourceFile ?? "",
      exportedAt,
      facility.population ?? "",
      facility.sampleSize ?? "",
      progress.completed,
      progress.defectLocations,
      progress.totalVariance,
      location.sampleOrder,
      location.randomDrawOrder ?? "",
      location.costCenter,
      location.binCode,
      location.binDescription ?? "",
      entry.reviewed ?? false,
      entry.varianceCount ?? 0,
      entry.notes ?? "",
    ]);
  });

  return {
    filename,
    csvContent: lines.map((row) => row.map(csvEscape).join(",")).join("\n"),
  };
}

async function saveResultsToDirectory(exportData, interactive) {
  const directoryHandle = await getResultsDirectoryHandle({ interactive });
  if (!directoryHandle) {
    return null;
  }

  try {
    const fileHandle = await directoryHandle.getFileHandle(exportData.filename, { create: true });
    const writable = await fileHandle.createWritable();
    await writable.write(exportData.csvContent);
    await writable.close();

    return {
      method: "folder",
      filename: exportData.filename,
      folderName: directoryHandle.name ?? "Results",
      reason: "saved-to-folder",
    };
  } catch (error) {
    console.warn("Unable to save results directly to the selected directory", error);
    updateResultsFolderUI(
      "The selected folder could not be written to just now. The app will fall back to a normal download until access is restored.",
    );
    return null;
  }
}

async function getResultsDirectoryHandle({ interactive = false } = {}) {
  if (!state.supportsFolderSave) {
    return null;
  }

  let directoryHandle = state.resultsDirectoryHandle;

  if (!directoryHandle) {
    if (!interactive) {
      return null;
    }

    try {
      directoryHandle = await window.showDirectoryPicker({ id: "irdr-results", mode: "readwrite" });
    } catch (error) {
      if (error?.name !== "AbortError") {
        console.warn("Results folder selection failed", error);
      }
      return null;
    }

    state.resultsDirectoryHandle = directoryHandle;
    state.resultsDirectoryName = directoryHandle.name ?? "";

    try {
      await settingsSet(RESULTS_DIRECTORY_KEY, directoryHandle);
    } catch (error) {
      console.warn("Unable to persist results directory handle", error);
    }
  }

  const permissionGranted = await verifyDirectoryPermission(directoryHandle, interactive);
  if (!permissionGranted) {
    updateResultsFolderUI(
      interactive
        ? "The device did not grant write access to that folder. Results will download instead until access is allowed."
        : "",
    );
    return null;
  }

  state.resultsDirectoryHandle = directoryHandle;
  state.resultsDirectoryName = directoryHandle.name ?? "";
  updateResultsFolderUI();
  return directoryHandle;
}

async function chooseResultsFolder() {
  if (!state.supportsFolderSave) {
    updateResultsFolderUI();
    return;
  }

  const directoryHandle = await getResultsDirectoryHandle({ interactive: true });
  if (!directoryHandle) {
    updateResultsFolderUI();
    return;
  }

  updateResultsFolderUI(
    `Direct save is ready for folder "${directoryHandle.name}". Finish Count will try to write the results CSV there.`,
  );
}

async function clearResultsFolder() {
  state.resultsDirectoryHandle = null;
  state.resultsDirectoryName = "";

  try {
    await settingsDelete(RESULTS_DIRECTORY_KEY);
  } catch (error) {
    console.warn("Unable to clear saved results directory handle", error);
  }

  updateResultsFolderUI();
}

async function verifyDirectoryPermission(directoryHandle, askForPermission) {
  if (!directoryHandle) {
    return false;
  }

  const options = { mode: "readwrite" };

  try {
    if (typeof directoryHandle.queryPermission === "function") {
      const query = await directoryHandle.queryPermission(options);
      if (query === "granted") {
        return true;
      }
    }

    if (!askForPermission) {
      return false;
    }

    if (typeof directoryHandle.requestPermission === "function") {
      const request = await directoryHandle.requestPermission(options);
      return request === "granted";
    }
  } catch (error) {
    console.warn("Directory permission check failed", error);
  }

  return false;
}

function updateResultsFolderUI(overrideMessage = "") {
  if (!state.supportsFolderSave) {
    elements.resultsFolderStatus.textContent =
      "This browser does not expose direct folder save here. Results will download to the device instead.";
    elements.chooseResultsFolder.disabled = true;
    elements.clearResultsFolder.disabled = true;
    elements.sessionNote.innerHTML =
      "Tapping <strong>Finish Count</strong> on the last location will download the results CSV to the device.";
    return;
  }

  if (overrideMessage) {
    elements.resultsFolderStatus.textContent = overrideMessage;
  } else if (state.resultsDirectoryHandle) {
    elements.resultsFolderStatus.textContent =
      `Direct save is configured for folder "${state.resultsDirectoryName || "Results"}".`;
  } else {
    elements.resultsFolderStatus.textContent =
      "No results folder is selected yet. Choose the device's local IRDR/Results folder to save files there directly.";
  }

  elements.chooseResultsFolder.disabled = false;
  elements.clearResultsFolder.disabled = !state.resultsDirectoryHandle;
  elements.sessionNote.innerHTML = state.resultsDirectoryHandle
    ? "Tapping <strong>Finish Count</strong> on the last location will try to save straight into the selected results folder. If that fails, the CSV will download to the device."
    : "Tapping <strong>Finish Count</strong> on the last location will download the results CSV unless a results folder has been selected.";
}

function resetSession() {
  const week = getActiveWeek();
  const facility = getActiveFacility();
  if (!week || !facility) {
    return;
  }

  const confirmed = window.confirm(`Reset saved progress for ${state.activeFacility} on ${week.label}?`);
  if (!confirmed) {
    return;
  }

  localStorage.removeItem(getSessionKey(week.id, state.activeFacility));
  state.currentIndex = 0;
  renderCountScreen();
  updateSelectionSummary();
  renderFacilityOptions();
}

function getActiveWeek() {
  return state.weeks.find((week) => week.id === state.activeWeekId) ?? null;
}

function getActiveFacility() {
  const week = getActiveWeek();
  if (!week || !state.activeFacility) {
    return null;
  }
  return week.facilities?.[state.activeFacility] ?? null;
}

function getSessionProgress(weekId, facilityName, locations) {
  const session = loadSession(weekId, facilityName);
  let completed = 0;
  let totalVariance = 0;
  let defectLocations = 0;

  locations.forEach((location) => {
    const entry = session.entries[String(location.sampleOrder)];
    if (!entry) {
      return;
    }

    const hasNotes = Boolean((entry.notes ?? "").trim());
    const varianceCount = Number(entry.varianceCount ?? 0);
    if (Boolean(entry.reviewed) || hasNotes || varianceCount > 0) {
      completed += 1;
    }
    if (varianceCount > 0) {
      totalVariance += varianceCount;
      defectLocations += 1;
    }
  });

  return { completed, totalVariance, defectLocations };
}

function loadSession(weekId, facilityName) {
  const raw = localStorage.getItem(getSessionKey(weekId, facilityName));
  if (!raw) {
    return { currentIndex: 0, entries: {} };
  }

  try {
    const parsed = JSON.parse(raw);
    return {
      currentIndex: Number(parsed.currentIndex ?? 0),
      entries: parsed.entries ?? {},
    };
  } catch (error) {
    console.warn("Unable to parse session", error);
    return { currentIndex: 0, entries: {} };
  }
}

function saveSession(weekId, facilityName, session) {
  localStorage.setItem(getSessionKey(weekId, facilityName), JSON.stringify(session));
}

function getSessionKey(weekId, facilityName) {
  return `irdr-progress::${weekId}::${facilityName}`;
}

function getStatusText(varianceCount) {
  if (varianceCount > 0) {
    return `${varianceCount} cases marked incorrect or missing for this location.`;
  }
  return "No variance recorded for this location.";
}

function buildExportMessage(outcome, isComplete) {
  const prefix = isComplete ? "Count complete. " : "";
  if (!outcome) {
    return `${prefix}The results could not be exported.`;
  }

  if (outcome.method === "folder") {
    return `${prefix}Results were saved to ${outcome.folderName}/${outcome.filename}.`;
  }

  if (state.supportsFolderSave && !state.resultsDirectoryHandle) {
    return `${prefix}Results were downloaded as ${outcome.filename}. Choose the local IRDR/Results folder on the setup screen if you want direct save next time.`;
  }

  return `${prefix}Results were downloaded as ${outcome.filename}.`;
}

function downloadCsv(filename, csvContent) {
  const blob = new Blob([csvContent], { type: "text/csv;charset=utf-8" });
  const url = URL.createObjectURL(blob);
  const anchor = document.createElement("a");
  anchor.href = url;
  anchor.download = filename;
  anchor.click();
  URL.revokeObjectURL(url);
}

function csvEscape(value) {
  const text = String(value ?? "");
  return `"${text.replace(/"/g, '""')}"`;
}

function buildResultsFileName(weekId, facilityName, exportedAt) {
  const safeFacility = String(facilityName ?? "facility")
    .trim()
    .toLowerCase()
    .replace(/[^a-z0-9]+/g, "-")
    .replace(/^-|-$/g, "");
  const stamp = String(exportedAt ?? "")
    .replace(/[:]/g, "")
    .replace(/\.\d+Z$/, "Z");
  return `irdr-results-${weekId}-${safeFacility}-${stamp}.csv`;
}

function clamp(value, min, max) {
  return Math.min(Math.max(value, min), max);
}

function supportsDirectorySave() {
  return window.isSecureContext && "showDirectoryPicker" in window && "indexedDB" in window;
}

function openSettingsDb() {
  return new Promise((resolve, reject) => {
    const request = indexedDB.open(SETTINGS_DB_NAME, 1);

    request.addEventListener("upgradeneeded", () => {
      request.result.createObjectStore(SETTINGS_STORE_NAME);
    });
    request.addEventListener("success", () => resolve(request.result));
    request.addEventListener("error", () => reject(request.error));
  });
}

async function settingsGet(key) {
  const db = await openSettingsDb();
  return new Promise((resolve, reject) => {
    const transaction = db.transaction(SETTINGS_STORE_NAME, "readonly");
    const store = transaction.objectStore(SETTINGS_STORE_NAME);
    const request = store.get(key);
    request.addEventListener("success", () => resolve(request.result));
    request.addEventListener("error", () => reject(request.error));
  });
}

async function settingsSet(key, value) {
  const db = await openSettingsDb();
  return new Promise((resolve, reject) => {
    const transaction = db.transaction(SETTINGS_STORE_NAME, "readwrite");
    const store = transaction.objectStore(SETTINGS_STORE_NAME);
    store.put(value, key);
    transaction.addEventListener("complete", () => resolve());
    transaction.addEventListener("error", () => reject(transaction.error));
    transaction.addEventListener("abort", () => reject(transaction.error));
  });
}

async function settingsDelete(key) {
  const db = await openSettingsDb();
  return new Promise((resolve, reject) => {
    const transaction = db.transaction(SETTINGS_STORE_NAME, "readwrite");
    const store = transaction.objectStore(SETTINGS_STORE_NAME);
    store.delete(key);
    transaction.addEventListener("complete", () => resolve());
    transaction.addEventListener("error", () => reject(transaction.error));
    transaction.addEventListener("abort", () => reject(transaction.error));
  });
}

function registerServiceWorker() {
  if (!("serviceWorker" in navigator)) {
    return;
  }

  window.addEventListener("load", () => {
    navigator.serviceWorker.register("sw.js").catch((error) => {
      console.warn("Service worker registration failed", error);
    });
  });
}

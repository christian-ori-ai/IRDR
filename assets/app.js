import { oneDriveConfig } from "./onedrive-config.js?v=20260417f";

const SETTINGS_DB_NAME = "irdr-mobile";
const SETTINGS_STORE_NAME = "settings";
const RESULTS_DIRECTORY_KEY = "results-directory-handle";
const COUNTER_NAME_KEY = "irdr-counter-name";
const DEVICE_ID_KEY = "irdr-device-id";
const CLAIM_LEASE_MS = 15 * 60 * 1000;
const HEARTBEAT_INTERVAL_MS = 60 * 1000;
const RUNTIME_DIR_NAME = ".irdr-runtime";
const CLAIMS_DIR_NAME = "claims";
const SESSIONS_DIR_NAME = "sessions";
const STATUS_DIR_NAME = "status";

const state = {
  weeks: [],
  activeWeekId: "",
  activeFacility: "",
  currentIndex: 0,
  supportsFolderSave: false,
  supportsFileShare: false,
  resultsDirectoryHandle: null,
  resultsDirectoryName: "",
  resultsDirectoryNeedsReconnect: false,
  oneDriveConfigured: false,
  oneDriveReady: false,
  oneDriveBusy: false,
  oneDriveClient: null,
  oneDriveAccount: null,
  oneDriveUploadFolderId: "",
  oneDriveUploadFolderWebUrl: "",
  counterName: "",
  deviceId: "",
  sharedStatus: null,
  activeClaimKey: "",
  activeClaimContext: null,
  heartbeatId: null,
};

const elements = {
  setupScreen: document.getElementById("setup-screen"),
  countScreen: document.getElementById("count-screen"),
  weekSelect: document.getElementById("week-select"),
  facilityOptions: document.getElementById("facility-options"),
  facilityOptionTemplate: document.getElementById("facility-option-template"),
  startCount: document.getElementById("start-count"),
  counterName: document.getElementById("counter-name"),
  summaryLocations: document.getElementById("summary-locations"),
  summaryCompleted: document.getElementById("summary-completed"),
  summaryVariance: document.getElementById("summary-variance"),
  summaryNote: document.getElementById("summary-note"),
  sharedStatus: document.getElementById("shared-status"),
  resultsFolderStatus: document.getElementById("results-folder-status"),
  chooseResultsFolder: document.getElementById("choose-results-folder"),
  clearResultsFolder: document.getElementById("clear-results-folder"),
  oneDriveStatus: document.getElementById("onedrive-status"),
  oneDriveFootnote: document.getElementById("onedrive-footnote"),
  connectOneDrive: document.getElementById("connect-onedrive"),
  disconnectOneDrive: document.getElementById("disconnect-onedrive"),
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
  uploadOneDrive: document.getElementById("upload-onedrive"),
  shareResults: document.getElementById("share-results"),
  resetSession: document.getElementById("reset-session"),
  sessionNote: document.getElementById("session-note"),
};

document.addEventListener("DOMContentLoaded", init);

async function init() {
  state.supportsFolderSave = supportsDirectorySave();
  state.supportsFileShare = supportsFileShare();
  hydrateOperatorIdentity();
  registerServiceWorker();
  await initializeOneDrive();
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
    void refreshSelectionSummary();
  });

  elements.counterName.addEventListener("input", () => {
    state.counterName = elements.counterName.value.trim();
    localStorage.setItem(COUNTER_NAME_KEY, elements.counterName.value);
    void refreshSelectionSummary();
  });

  elements.startCount.addEventListener("click", () => {
    void startCountSession();
  });
  elements.backToSetup.addEventListener("click", () => {
    void pauseCountSession({ returnToSetup: true });
  });
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
  elements.shareResults.addEventListener("click", () => {
    void handleManualShare();
  });
  elements.resetSession.addEventListener("click", () => {
    void resetSession();
  });
  elements.chooseResultsFolder.addEventListener("click", () => {
    void chooseResultsFolder();
  });
  elements.clearResultsFolder.addEventListener("click", () => {
    void clearResultsFolder();
  });
  elements.connectOneDrive.addEventListener("click", () => {
    void connectOneDrive();
  });
  elements.disconnectOneDrive.addEventListener("click", () => {
    void disconnectOneDrive();
  });
  elements.uploadOneDrive.addEventListener("click", () => {
    void handleManualOneDriveUpload();
  });
  document.addEventListener("visibilitychange", () => {
    if (document.hidden) {
      saveCountSnapshot();
      return;
    }

    if (!isCountScreenActive()) {
      void refreshSelectionSummary();
    }
  });
}

function hydrateOperatorIdentity() {
  const storedName = localStorage.getItem(COUNTER_NAME_KEY) ?? "";
  const storedDeviceId = localStorage.getItem(DEVICE_ID_KEY);
  state.counterName = storedName.trim();
  state.deviceId = storedDeviceId || generateDeviceId();
  elements.counterName.value = storedName;
  if (!storedDeviceId) {
    localStorage.setItem(DEVICE_ID_KEY, state.deviceId);
  }
}

async function initializeOneDrive() {
  state.oneDriveConfigured = hasConfiguredOneDrive();
  renderOneDriveUI();

  if (!state.oneDriveConfigured) {
    return;
  }

  if (!window.msal?.PublicClientApplication) {
    elements.oneDriveStatus.textContent =
      "Microsoft sign-in could not be loaded on this page, so OneDrive upload is unavailable right now.";
    elements.oneDriveFootnote.textContent =
      "The local auth library is missing or blocked. The app can still save, share, and download results locally.";
    return;
  }

  try {
    state.oneDriveClient = new window.msal.PublicClientApplication({
      auth: {
        clientId: oneDriveConfig.clientId,
        authority: oneDriveConfig.authority,
        redirectUri: oneDriveConfig.redirectUri,
        postLogoutRedirectUri: oneDriveConfig.postLogoutRedirectUri ?? oneDriveConfig.redirectUri,
      },
      cache: {
        cacheLocation: "localStorage",
        storeAuthStateInCookie: false,
      },
    });

    if (typeof state.oneDriveClient.initialize === "function") {
      await state.oneDriveClient.initialize();
    }

    const redirectResult = await state.oneDriveClient.handleRedirectPromise();
    if (redirectResult?.account) {
      state.oneDriveClient.setActiveAccount(redirectResult.account);
    }

    const activeAccount = state.oneDriveClient.getActiveAccount()
      || state.oneDriveClient.getAllAccounts?.()[0]
      || null;

    if (activeAccount) {
      state.oneDriveClient.setActiveAccount(activeAccount);
    }

    state.oneDriveAccount = activeAccount;
    state.oneDriveUploadFolderId = "";
    state.oneDriveUploadFolderWebUrl = "";
    state.oneDriveReady = true;
  } catch (error) {
    console.warn("Unable to initialize Microsoft sign-in", error);
    elements.oneDriveStatus.textContent =
      "Microsoft sign-in could not initialize, so OneDrive upload is disabled for now.";
    elements.oneDriveFootnote.textContent =
      "The app will still work locally. Recheck assets/onedrive-config.js and the Entra app redirect URI if you want direct OneDrive upload.";
  }

  renderOneDriveUI();
}

function hasConfiguredOneDrive() {
  return Boolean(
    oneDriveConfig?.enabled
    && oneDriveConfig.clientId
    && !String(oneDriveConfig.clientId).includes("YOUR-"),
  );
}

function renderOneDriveUI() {
  const isConnected = Boolean(state.oneDriveAccount);
  const canConnect = state.oneDriveConfigured && Boolean(state.oneDriveClient || window.msal?.PublicClientApplication);

  if (!state.oneDriveConfigured) {
    elements.oneDriveStatus.textContent =
      "OneDrive upload is off until assets/onedrive-config.js has a real Entra app client ID and enabled is set to true.";
    elements.oneDriveFootnote.textContent =
      `When you are ready, register the GitHub Pages URL as a SPA redirect URI and point uploadPath at the OneDrive folder you want, such as ${getOneDriveUploadPathDisplay()}.`;
  } else if (!state.oneDriveReady) {
    elements.oneDriveStatus.textContent = "Preparing Microsoft sign-in for direct OneDrive upload...";
    elements.oneDriveFootnote.textContent =
      `Once connected, the app can upload the results CSV into OneDrive/${getOneDriveUploadPathDisplay()}.`;
  } else if (isConnected) {
    elements.oneDriveStatus.textContent = `Connected to OneDrive as ${state.oneDriveAccount.username || "your Microsoft account"}.`;
    elements.oneDriveFootnote.textContent =
      `Finish & Upload will send the CSV to OneDrive/${getOneDriveUploadPathDisplay()} while still keeping the local export fallback.`;
  } else {
    elements.oneDriveStatus.textContent =
      "OneDrive upload is configured, but this device is not connected yet.";
    elements.oneDriveFootnote.textContent =
      `Connect once on this device and the app can upload straight into OneDrive/${getOneDriveUploadPathDisplay()} without relying on folder picking.`;
  }

  elements.connectOneDrive.disabled = !canConnect || state.oneDriveBusy || isConnected;
  elements.disconnectOneDrive.disabled = !isConnected || state.oneDriveBusy;
  elements.uploadOneDrive.hidden = !isConnected;
}

async function connectOneDrive() {
  if (!state.oneDriveConfigured || !state.oneDriveClient) {
    renderOneDriveUI();
    return;
  }

  state.oneDriveBusy = true;
  renderOneDriveUI();

  try {
    await state.oneDriveClient.loginRedirect({
      scopes: getOneDriveScopes(),
      prompt: "select_account",
    });
  } catch (error) {
    state.oneDriveBusy = false;
    console.warn("Unable to start Microsoft sign-in", error);
    elements.oneDriveStatus.textContent =
      "The app could not start Microsoft sign-in just now. You can still count locally and try connecting again later.";
    renderOneDriveUI();
  }
}

async function disconnectOneDrive() {
  if (!state.oneDriveClient || !state.oneDriveAccount) {
    renderOneDriveUI();
    return;
  }

  state.oneDriveBusy = true;
  renderOneDriveUI();

  try {
    await state.oneDriveClient.logoutRedirect({
      account: state.oneDriveAccount,
      postLogoutRedirectUri: oneDriveConfig.postLogoutRedirectUri ?? oneDriveConfig.redirectUri,
    });
  } catch (error) {
    state.oneDriveBusy = false;
    console.warn("Unable to sign out of Microsoft", error);
    elements.oneDriveStatus.textContent =
      "The app could not sign out of Microsoft just now. You can still continue counting locally.";
    renderOneDriveUI();
  }
}

function getOneDriveScopes() {
  return Array.from(new Set(oneDriveConfig.graphScopes ?? ["Files.ReadWrite"]));
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

  if (state.resultsDirectoryHandle) {
    const verifiedHandle = await getResultsDirectoryHandle({ interactive: false });
    if (verifiedHandle || state.resultsDirectoryNeedsReconnect) {
      return;
    }
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
  void refreshSelectionSummary();
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
    const countKey = buildCountKey(week.id, facilityName);
    const progress = state.sharedStatus?.key === countKey
      ? calculateSessionProgress(
        choosePreferredSession(loadSession(week.id, facilityName), state.sharedStatus?.session),
        facility.locations ?? [],
      )
      : getSessionProgress(week.id, facilityName, facility.locations ?? []);

    node.dataset.facility = facilityName;
    node.querySelector(".facility-card__name").textContent = facilityName;
    node.querySelector(".facility-card__meta").textContent =
      `${facility.sampleSize} locations • ${progress.completed} complete`;
    node.classList.toggle("is-selected", facilityName === state.activeFacility);
    node.addEventListener("click", () => {
      state.activeFacility = facilityName;
      renderFacilityOptions();
      void refreshSelectionSummary();
    });

    elements.facilityOptions.appendChild(node);
  });

  renderStartButtonState();
}

async function refreshSelectionSummary() {
  const week = getActiveWeek();
  const facility = getActiveFacility();

  if (!week || !facility) {
    state.sharedStatus = null;
    elements.summaryLocations.textContent = "--";
    elements.summaryCompleted.textContent = "--";
    elements.summaryVariance.textContent = "--";
    elements.sharedStatus.textContent = "";
    elements.summaryNote.textContent = week
      ? "Pick a facility to see the current progress."
      : "Choose a week and facility to see the sample details.";
    renderStartButtonState();
    return;
  }

  await refreshSharedStatus(week.id, state.activeFacility);
  const progress = calculateSessionProgress(
    loadSession(week.id, state.activeFacility),
    facility.locations,
  );
  elements.summaryLocations.textContent = String(facility.sampleSize);
  elements.summaryCompleted.textContent = String(progress.completed);
  elements.summaryVariance.textContent = String(progress.totalVariance);
  elements.summaryNote.textContent =
    `${week.company} • ${facility.population} eligible locations in this facility`;
  elements.sharedStatus.textContent = state.resultsDirectoryNeedsReconnect
    ? "Saved Results folder access expired on this device. Choose Results Folder again if you want direct save."
    : "Progress is saved on this device/browser until you export or reset it.";
  renderStartButtonState();
}

async function startCountSession() {
  const week = getActiveWeek();
  const facility = getActiveFacility();
  if (!week || !facility) {
    return;
  }

  const session = loadSession(week.id, state.activeFacility);
  const sessionToUse = normalizeSession(session);
  saveSession(week.id, state.activeFacility, sessionToUse);
  state.currentIndex = clamp(
    sessionToUse.currentIndex ?? 0,
    0,
    Math.max((facility.locations?.length ?? 1) - 1, 0),
  );
  setActiveScreen("count");
  renderCountScreen();
}

function renderStartButtonState() {
  const hasBasics = Boolean(state.activeWeekId && state.activeFacility);
  let label = "Start Count";
  let disabled = !hasBasics;

  if (hasLocalProgress()) {
    label = "Resume Count";
  }

  elements.startCount.textContent = label;
  elements.startCount.disabled = disabled;
}

function renderSharedStatus() {
  elements.sharedStatus.textContent = describeSharedStatus(state.sharedStatus);
}

function setActiveScreen(screenName) {
  const isSetup = screenName === "setup";
  elements.setupScreen.classList.toggle("screen--active", isSetup);
  elements.countScreen.classList.toggle("screen--active", !isSetup);
}

function isCountScreenActive() {
  return elements.countScreen.classList.contains("screen--active");
}

async function pauseCountSession({ returnToSetup = false } = {}) {
  saveCountSnapshot();

  if (returnToSetup) {
    setActiveScreen("setup");
    await refreshSelectionSummary();
    renderFacilityOptions();
  }
}

function saveCountSnapshot() {
  const week = getActiveWeek();
  const facility = getActiveFacility();

  if (!week || !facility || !isCountScreenActive()) {
    return;
  }

  saveCurrentEntry(false);
  const session = loadSession(week.id, state.activeFacility);
  session.currentIndex = state.currentIndex;
  saveSession(week.id, state.activeFacility, session);
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
  elements.uploadOneDrive.hidden = !Boolean(state.oneDriveAccount);
  elements.shareResults.hidden = !state.supportsFileShare;
  elements.nextLocation.textContent =
    state.currentIndex === facility.locations.length - 1
      ? (shouldAutoUploadToOneDrive()
        ? "Finish & Upload"
        : (state.supportsFileShare ? "Finish & Share" : "Finish Count"))
      : "Next Location";
  elements.sessionNote.innerHTML = buildSessionNoteMessage();
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
    updatedBy: state.counterName,
    updatedByDeviceId: state.deviceId,
    updatedAt: new Date().toISOString(),
  };
  session.currentIndex = state.currentIndex;
  saveSession(week.id, state.activeFacility, session);
  renderCountScreen();
  void refreshSelectionSummary();
  renderFacilityOptions();
  void syncSharedSession();
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
    updatedBy: state.counterName,
    updatedByDeviceId: state.deviceId,
    updatedAt: new Date().toISOString(),
  };
  session.currentIndex = state.currentIndex;
  saveSession(week.id, state.activeFacility, session);
  void refreshSelectionSummary();
  renderFacilityOptions();
  void syncSharedSession();
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
      share: shouldAutoUploadToOneDrive() ? false : state.supportsFileShare,
      uploadToOneDrive: shouldAutoUploadToOneDrive(),
    });
    setActiveScreen("setup");
    await refreshSelectionSummary();
    renderFacilityOptions();
    window.alert(buildExportMessage(outcome, true));
    return;
  }

  state.currentIndex = clamp(state.currentIndex + direction, 0, facility.locations.length - 1);
  const session = loadSession(week.id, state.activeFacility);
  session.currentIndex = state.currentIndex;
  saveSession(week.id, state.activeFacility, session);
  void syncSharedSession();
  renderCountScreen();
}

async function handleManualExport() {
  const outcome = await exportResults({ preferFolder: true, interactive: true });
  if (outcome) {
    window.alert(buildExportMessage(outcome, false));
  }
}

async function handleManualShare() {
  const outcome = await exportResults({
    preferFolder: true,
    interactive: false,
    share: true,
  });
  if (outcome) {
    window.alert(buildExportMessage(outcome, false));
  }
}

async function handleManualOneDriveUpload() {
  if (!state.oneDriveAccount) {
    window.alert("Connect OneDrive on the setup screen first, then try the upload again.");
    return;
  }

  const outcome = await exportResults({
    preferFolder: true,
    interactive: false,
    uploadToOneDrive: true,
  });
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
  const shareOutcome = options.share ? await shareExportData(exportData) : { status: "not-requested" };
  let localOutcome = null;

  if (shouldPreferFolder) {
    const saveOutcome = await saveResultsToDirectory(exportData, options.interactive === true);
    if (saveOutcome) {
      localOutcome = saveOutcome;
    }
  }

  if (!localOutcome) {
    downloadCsv(exportData.filename, exportData.csvContent);
    localOutcome = {
      method: "download",
      filename: exportData.filename,
      reason: state.supportsFolderSave ? "download-fallback" : "unsupported",
    };
  }

  const oneDriveOutcome = options.uploadToOneDrive
    ? await uploadResultsToOneDrive(exportData)
    : { status: "not-requested" };

  return {
    ...localOutcome,
    shareStatus: shareOutcome.status,
    oneDriveStatus: oneDriveOutcome.status,
    oneDriveLocation: oneDriveOutcome.location ?? "",
    oneDriveWebUrl: oneDriveOutcome.webUrl ?? "",
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
      "exported_by_counter",
      "exported_by_device",
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
      "last_updated_by",
      "last_updated_at",
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
      state.counterName,
      state.deviceId,
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
      entry.updatedBy ?? "",
      entry.updatedAt ?? "",
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
    if (state.resultsDirectoryHandle && !interactive) {
      await forgetResultsDirectory({
        message:
          "Saved results folder access expired on this device. Choose Results Folder again if you want direct save.",
        needsReconnect: true,
      });
    } else {
      updateResultsFolderUI(
        interactive
          ? "The device did not grant write access to that folder. Results will download instead until access is allowed."
          : "",
      );
    }
    return null;
  }

  state.resultsDirectoryHandle = directoryHandle;
  state.resultsDirectoryName = directoryHandle.name ?? "";
  state.resultsDirectoryNeedsReconnect = false;
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
  await refreshSelectionSummary();
}

async function clearResultsFolder() {
  await forgetResultsDirectory();
  await refreshSelectionSummary();
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

async function refreshSharedStatus(weekId, facilityName) {
  if (!weekId || !facilityName) {
    state.sharedStatus = null;
    return null;
  }

  state.sharedStatus = {
    type: state.resultsDirectoryNeedsReconnect
      ? "folder-access-needed"
      : (state.supportsFolderSave ? "local-progress" : "local-only"),
  };
  return state.sharedStatus;
}

function describeSharedStatus(sharedStatus) {
  if (!sharedStatus) {
    return "";
  }

  if (sharedStatus.type === "folder-not-selected") {
    return "Direct save is inactive until a Results folder is selected.";
  }

  if (sharedStatus.type === "folder-access-needed") {
    return "Saved Results folder access expired on this device. Choose Results Folder again if you want direct save.";
  }

  if (sharedStatus.type === "local-progress") {
    return "Progress is saved on this device/browser until you export or reset it.";
  }

  if (sharedStatus.type === "local-only") {
    return "This browser is using local-only progress for this count.";
  }

  if (sharedStatus.type === "claimed-by-me") {
    return "This count is currently claimed by you. Resume when you're ready.";
  }

  if (sharedStatus.type === "claimed-by-other") {
    return `In progress by ${sharedStatus.claim?.counterName || "another counter"} • last seen ${formatTimestamp(sharedStatus.claim?.lastSeenAt)}.`;
  }

  if (sharedStatus.type === "completed") {
    return `Completed by ${sharedStatus.status?.counterName || "a counter"} at ${formatTimestamp(sharedStatus.status?.completedAt)}. Starting again will reopen the count.`;
  }

  if (sharedStatus.type === "available-resume-me") {
    return `Your shared progress is saved from ${formatTimestamp(sharedStatus.session?.lastUpdatedAt)}. Starting will resume it.`;
  }

  if (sharedStatus.type === "available-resume") {
    return `Shared progress exists from ${sharedStatus.session?.counterName || "another counter"} • last updated ${formatTimestamp(sharedStatus.session?.lastUpdatedAt)}. Starting will continue that work.`;
  }

  return "Shared count is available to claim.";
}

async function ensureCountClaim(weekId, facilityName) {
  if (!state.resultsDirectoryHandle) {
    return { ok: true, mode: "local-only" };
  }

  const key = buildCountKey(weekId, facilityName);
  if (state.activeClaimKey && state.activeClaimKey !== key) {
    await releaseActiveClaim("paused");
  }

  const currentStatus = state.sharedStatus?.key === key
    ? state.sharedStatus
    : await refreshSharedStatus(weekId, facilityName);

  if (currentStatus?.type === "claimed-by-other") {
    const shouldTakeOver = window.confirm(
      `${currentStatus.claim?.counterName || "Another counter"} is already working this count. Taking over will move the shared lock to you. Continue?`,
    );
    if (!shouldTakeOver) {
      return { ok: false, reason: "claimed-by-other" };
    }
  }

  if (currentStatus?.type === "completed") {
    const shouldReopen = window.confirm(
      `This count was completed by ${currentStatus.status?.counterName || "another counter"}. Reopen it and continue from the saved progress?`,
    );
    if (!shouldReopen) {
      return { ok: false, reason: "completed" };
    }
  }

  const now = Date.now();
  const existingClaim = currentStatus?.claim;
  const claim = {
    key,
    weekId,
    facilityName,
    counterName: state.counterName,
    deviceId: state.deviceId,
    sessionId: existingClaim?.sessionId || generateDeviceId(),
    claimedAt: existingClaim?.claimedAt || new Date(now).toISOString(),
    lastSeenAt: new Date(now).toISOString(),
    expiresAt: new Date(now + CLAIM_LEASE_MS).toISOString(),
  };

  await writeRuntimeJson(CLAIMS_DIR_NAME, `${key}.json`, claim);
  await writeRuntimeJson(STATUS_DIR_NAME, `${key}.json`, {
    state: "in_progress",
    key,
    weekId,
    facilityName,
    counterName: state.counterName,
    deviceId: state.deviceId,
    lastSeenAt: claim.lastSeenAt,
  });

  if (!state.resultsDirectoryHandle) {
    return { ok: true, mode: "local-only" };
  }

  const confirmedClaim = await readRuntimeJson(CLAIMS_DIR_NAME, `${key}.json`);
  if (!confirmedClaim || confirmedClaim.sessionId !== claim.sessionId) {
    if (!state.resultsDirectoryHandle) {
      return { ok: true, mode: "local-only" };
    }
    await refreshSharedStatus(weekId, facilityName);
    return { ok: false, reason: "claim-conflict" };
  }

  state.activeClaimKey = key;
  state.activeClaimContext = { key, weekId, facilityName };
  await refreshSharedStatus(weekId, facilityName);
  return { ok: true, claim };
}

function startHeartbeat(weekId, facilityName) {
  stopHeartbeat();
  state.activeClaimContext = {
    key: buildCountKey(weekId, facilityName),
    weekId,
    facilityName,
  };
  void pulseHeartbeat();
  state.heartbeatId = window.setInterval(() => {
    void pulseHeartbeat();
  }, HEARTBEAT_INTERVAL_MS);
}

function stopHeartbeat() {
  if (state.heartbeatId) {
    window.clearInterval(state.heartbeatId);
    state.heartbeatId = null;
  }
}

async function pulseHeartbeat() {
  if (!state.activeClaimContext || !state.resultsDirectoryHandle || document.hidden) {
    return;
  }

  const { key, weekId, facilityName } = state.activeClaimContext;
  const existingClaim = await readRuntimeJson(CLAIMS_DIR_NAME, `${key}.json`);
  if (existingClaim && !isOwnedByCurrentCounter(existingClaim)) {
    await handleClaimLost(weekId, facilityName, existingClaim);
    return;
  }

  const now = Date.now();
  const updatedClaim = {
    key,
    weekId,
    facilityName,
    counterName: state.counterName,
    deviceId: state.deviceId,
    sessionId: existingClaim?.sessionId || generateDeviceId(),
    claimedAt: existingClaim?.claimedAt || new Date(now).toISOString(),
    lastSeenAt: new Date(now).toISOString(),
    expiresAt: new Date(now + CLAIM_LEASE_MS).toISOString(),
  };

  await writeRuntimeJson(CLAIMS_DIR_NAME, `${key}.json`, updatedClaim);
  await writeRuntimeJson(STATUS_DIR_NAME, `${key}.json`, {
    state: "in_progress",
    key,
    weekId,
    facilityName,
    counterName: state.counterName,
    deviceId: state.deviceId,
    lastSeenAt: updatedClaim.lastSeenAt,
  });
}

async function loadSharedSession(weekId, facilityName) {
  if (!state.resultsDirectoryHandle) {
    return null;
  }

  const key = buildCountKey(weekId, facilityName);
  return readRuntimeJson(SESSIONS_DIR_NAME, `${key}.json`);
}

function choosePreferredSession(localSession, sharedSession) {
  if (!sharedSession) {
    return normalizeSession(localSession);
  }

  const localStamp = getSessionTimestamp(localSession);
  const sharedStamp = getSessionTimestamp(sharedSession);
  return sharedStamp > localStamp ? normalizeSession(sharedSession) : normalizeSession(localSession);
}

async function syncSharedSession() {
  if (!state.activeClaimContext || !state.resultsDirectoryHandle) {
    return;
  }

  const { key, weekId, facilityName } = state.activeClaimContext;
  const activeClaim = await readRuntimeJson(CLAIMS_DIR_NAME, `${key}.json`);
  if (activeClaim && !isClaimExpired(activeClaim) && !isOwnedByCurrentCounter(activeClaim)) {
    await handleClaimLost(weekId, facilityName, activeClaim);
    return;
  }

  const localSession = loadSession(weekId, facilityName);
  const payload = {
    key,
    weekId,
    facilityName,
    counterName: state.counterName,
    deviceId: state.deviceId,
    currentIndex: localSession.currentIndex ?? 0,
    entries: localSession.entries ?? {},
    lastUpdatedAt: new Date().toISOString(),
  };

  await writeRuntimeJson(SESSIONS_DIR_NAME, `${key}.json`, payload);
}

async function handleClaimLost(weekId, facilityName, claim) {
  stopHeartbeat();
  state.activeClaimKey = "";
  state.activeClaimContext = null;
  await refreshSharedStatus(weekId, facilityName);
  renderSharedStatus();
  renderStartButtonState();
  renderFacilityOptions();

  if (!isCountScreenActive()) {
    return;
  }

  setActiveScreen("setup");
  await refreshSelectionSummary();
  window.alert(
    `This count was taken over by ${claim?.counterName || "another counter"}. Your last saved progress is still on this device, but the shared lock has moved to them.`,
  );
}

async function markCountCompleted(week, facility, resultFileName) {
  if (state.activeClaimContext) {
    await syncSharedSession();
  }

  if (state.resultsDirectoryHandle) {
    const key = buildCountKey(week.id, state.activeFacility);
    const progress = getSessionProgress(week.id, state.activeFacility, facility.locations);
    await writeRuntimeJson(STATUS_DIR_NAME, `${key}.json`, {
      state: "completed",
      key,
      weekId: week.id,
      facilityName: state.activeFacility,
      counterName: state.counterName,
      deviceId: state.deviceId,
      completedAt: new Date().toISOString(),
      resultFileName,
      completedLocations: progress.completed,
      defectLocations: progress.defectLocations,
      totalVariance: progress.totalVariance,
    });
    await deleteRuntimeFile(CLAIMS_DIR_NAME, `${key}.json`);
  }

  stopHeartbeat();
  state.activeClaimKey = "";
  state.activeClaimContext = null;
  await refreshSharedStatus(week.id, state.activeFacility);
}

async function releaseActiveClaim(nextState = "paused") {
  if (!state.activeClaimContext || !state.resultsDirectoryHandle) {
    stopHeartbeat();
    state.activeClaimKey = "";
    state.activeClaimContext = null;
    return;
  }

  const { key, weekId, facilityName } = state.activeClaimContext;
  const claim = await readRuntimeJson(CLAIMS_DIR_NAME, `${key}.json`);
  if (claim && isOwnedByCurrentCounter(claim)) {
    await deleteRuntimeFile(CLAIMS_DIR_NAME, `${key}.json`);
    await writeRuntimeJson(STATUS_DIR_NAME, `${key}.json`, {
      state: nextState,
      key,
      weekId,
      facilityName,
      counterName: state.counterName,
      deviceId: state.deviceId,
      lastSeenAt: new Date().toISOString(),
    });
  }

  stopHeartbeat();
  state.activeClaimKey = "";
  state.activeClaimContext = null;
}

function hasLocalProgress() {
  const week = getActiveWeek();
  const facility = getActiveFacility();
  if (!week || !facility) {
    return false;
  }
  const session = loadSession(week.id, state.activeFacility);
  return Boolean(Object.keys(session.entries ?? {}).length);
}

function normalizeSession(session) {
  return {
    currentIndex: Number(session?.currentIndex ?? 0),
    entries: session?.entries ?? {},
  };
}

function getSessionTimestamp(session) {
  if (!session) {
    return 0;
  }
  const directStamp = Date.parse(session.lastUpdatedAt ?? "");
  if (Number.isFinite(directStamp)) {
    return directStamp;
  }
  return Math.max(
    0,
    ...Object.values(session.entries ?? {}).map((entry) => Date.parse(entry?.updatedAt ?? "") || 0),
  );
}

function buildCountKey(weekId, facilityName) {
  return `${weekId}--${slugify(facilityName)}`;
}

function slugify(value) {
  return String(value ?? "")
    .trim()
    .toLowerCase()
    .replace(/[^a-z0-9]+/g, "-")
    .replace(/^-|-$/g, "");
}

function isOwnedByCurrentCounter(record) {
  if (!record) {
    return false;
  }

  if (record.deviceId && state.deviceId) {
    return record.deviceId === state.deviceId;
  }

  return Boolean(record.counterName && state.counterName && record.counterName === state.counterName);
}

function isClaimExpired(claim) {
  return Date.parse(claim?.expiresAt ?? "") <= Date.now();
}

function formatTimestamp(value) {
  const stamp = Date.parse(value ?? "");
  if (!Number.isFinite(stamp)) {
    return "recently";
  }
  return new Date(stamp).toLocaleString([], {
    month: "numeric",
    day: "numeric",
    hour: "numeric",
    minute: "2-digit",
  });
}

async function getRuntimeDirectory(directoryName, create = false) {
  const base = state.resultsDirectoryHandle;
  if (!base) {
    return null;
  }

  try {
    const runtime = await base.getDirectoryHandle(RUNTIME_DIR_NAME, { create });
    return runtime.getDirectoryHandle(directoryName, { create });
  } catch (error) {
    if (error?.name === "NotFoundError") {
      return null;
    }
    if (isResultsDirectoryAccessError(error)) {
      await handleResultsDirectoryAccessLost(error);
      return null;
    }
    throw error;
  }
}

async function readRuntimeJson(directoryName, fileName) {
  const directory = await getRuntimeDirectory(directoryName, false);
  if (!directory) {
    return null;
  }

  try {
    const fileHandle = await directory.getFileHandle(fileName);
    const file = await fileHandle.getFile();
    return JSON.parse(await file.text());
  } catch (error) {
    if (error?.name === "NotFoundError") {
      return null;
    }
    if (isResultsDirectoryAccessError(error)) {
      await handleResultsDirectoryAccessLost(error);
      return null;
    }
    console.warn(`Unable to read shared runtime file ${fileName}`, error);
    return null;
  }
}

async function writeRuntimeJson(directoryName, fileName, data) {
  const directory = await getRuntimeDirectory(directoryName, true);
  if (!directory) {
    return false;
  }

  try {
    const fileHandle = await directory.getFileHandle(fileName, { create: true });
    const writable = await fileHandle.createWritable();
    await writable.write(JSON.stringify(data, null, 2));
    await writable.close();
    return true;
  } catch (error) {
    if (isResultsDirectoryAccessError(error)) {
      await handleResultsDirectoryAccessLost(error);
      return false;
    }
    throw error;
  }
}

async function deleteRuntimeFile(directoryName, fileName) {
  const directory = await getRuntimeDirectory(directoryName, false);
  if (!directory) {
    return;
  }

  try {
    await directory.removeEntry(fileName);
  } catch (error) {
    if (isResultsDirectoryAccessError(error)) {
      await handleResultsDirectoryAccessLost(error);
      return;
    }
    if (error?.name !== "NotFoundError") {
      console.warn(`Unable to delete shared runtime file ${fileName}`, error);
    }
  }
}

async function forgetResultsDirectory({ message = "", needsReconnect = false } = {}) {
  state.resultsDirectoryHandle = null;
  state.resultsDirectoryName = "";
  state.resultsDirectoryNeedsReconnect = needsReconnect;

  try {
    await settingsDelete(RESULTS_DIRECTORY_KEY);
  } catch (error) {
    console.warn("Unable to clear saved results directory handle", error);
  }

  updateResultsFolderUI(message);
}

async function handleResultsDirectoryAccessLost(
  error,
  message = "Saved results folder access is no longer available on this device. Choose Results Folder again if you want direct save.",
) {
  console.warn("Results directory access is no longer available", error);
  await forgetResultsDirectory({ message, needsReconnect: true });
}

function isResultsDirectoryAccessError(error) {
  return ["NotAllowedError", "SecurityError", "InvalidStateError"].includes(error?.name);
}

function updateResultsFolderUI(overrideMessage = "") {
  if (!state.supportsFolderSave) {
    elements.resultsFolderStatus.textContent =
      "This browser does not expose direct folder save here. Results will download to the device instead.";
    elements.chooseResultsFolder.disabled = true;
    elements.clearResultsFolder.disabled = true;
    elements.sessionNote.innerHTML = buildSessionNoteMessage();
    return;
  }

  if (overrideMessage) {
    elements.resultsFolderStatus.textContent = overrideMessage;
  } else if (state.resultsDirectoryHandle) {
    elements.resultsFolderStatus.textContent =
      `Direct save is configured for folder "${state.resultsDirectoryName || "Results"}".`;
  } else {
    elements.resultsFolderStatus.textContent =
      "No results folder is selected yet. Choose the device's local IRDR/Results folder if you want the CSV saved there directly.";
  }

  elements.chooseResultsFolder.disabled = false;
  elements.clearResultsFolder.disabled = !state.resultsDirectoryHandle;
  elements.sessionNote.innerHTML = buildSessionNoteMessage();
}

async function resetSession() {
  const week = getActiveWeek();
  const facility = getActiveFacility();
  if (!week || !facility) {
    return;
  }

  const confirmed = window.confirm(
    `Reset saved progress for ${state.activeFacility} on ${week.label} on this device?`,
  );
  if (!confirmed) {
    return;
  }

  localStorage.removeItem(getSessionKey(week.id, state.activeFacility));
  state.currentIndex = 0;
  setActiveScreen("setup");
  await refreshSelectionSummary();
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
  return calculateSessionProgress(loadSession(weekId, facilityName), locations);
}

function calculateSessionProgress(session, locations) {
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

  const shareMessage = buildShareMessage(outcome.shareStatus);
  const oneDriveMessage = buildOneDriveMessage(outcome);

  if (outcome.method === "folder") {
    return `${prefix}Results were saved to ${outcome.folderName}/${outcome.filename}.${shareMessage}${oneDriveMessage}`;
  }

  if (state.supportsFolderSave && !state.resultsDirectoryHandle) {
    return `${prefix}Results were downloaded as ${outcome.filename}. Choose the local IRDR/Results folder on the setup screen if you want direct save next time.${shareMessage}${oneDriveMessage}`;
  }

  return `${prefix}Results were downloaded as ${outcome.filename}.${shareMessage}${oneDriveMessage}`;
}

function buildShareMessage(shareStatus) {
  if (shareStatus === "shared") {
    return " The file was also shared from the device.";
  }

  if (shareStatus === "cancelled") {
    return " Sharing was canceled, but the export was still saved.";
  }

  if (shareStatus === "failed") {
    return " The device could not open the share sheet, but the export was still saved.";
  }

  return "";
}

function buildOneDriveMessage(outcome) {
  if (outcome.oneDriveStatus === "uploaded") {
    return outcome.oneDriveLocation
      ? ` The file was also uploaded to OneDrive at ${outcome.oneDriveLocation}.`
      : " The file was also uploaded to OneDrive.";
  }

  if (outcome.oneDriveStatus === "auth-required") {
    return " OneDrive upload needs you to reconnect your Microsoft account on the setup screen.";
  }

  if (outcome.oneDriveStatus === "failed") {
    return " OneDrive upload did not finish, but the local export still succeeded.";
  }

  return "";
}

function buildSessionNoteMessage() {
  if (shouldAutoUploadToOneDrive()) {
    return "Progress stays on this device/browser, and <strong>Finish &amp; Upload</strong> will also send the CSV to OneDrive while keeping the local export fallback.";
  }

  if (state.resultsDirectoryHandle) {
    return state.supportsFileShare
      ? "Progress stays on this device/browser, and <strong>Finish &amp; Share</strong> will open the device share sheet before saving straight into the selected results folder."
      : "Progress stays on this device/browser, and <strong>Finish Count</strong> will try to save straight into the selected results folder.";
  }

  return state.supportsFileShare
    ? "Progress stays on this device/browser, and <strong>Finish &amp; Share</strong> will open the device share sheet before downloading the CSV unless a results folder has been selected."
    : "Tapping <strong>Finish Count</strong> on the last location will download the results CSV unless a results folder has been selected.";
}

function shouldAutoUploadToOneDrive() {
  return Boolean(state.oneDriveAccount && oneDriveConfig.autoUploadOnFinish !== false);
}

async function uploadResultsToOneDrive(exportData) {
  if (!state.oneDriveConfigured) {
    return { status: "not-configured" };
  }

  if (!state.oneDriveClient || !state.oneDriveAccount) {
    return { status: "not-connected" };
  }

  try {
    const accessToken = await acquireOneDriveAccessToken();
    if (!accessToken) {
      return { status: "auth-required" };
    }

    const folder = await ensureOneDriveResultsFolder(accessToken);
    if (!folder?.id) {
      return { status: "failed" };
    }

    const response = await fetch(
      `https://graph.microsoft.com/v1.0/me/drive/items/${folder.id}:/${encodeGraphPathSegment(exportData.filename)}:/content`,
      {
        method: "PUT",
        headers: {
          Authorization: `Bearer ${accessToken}`,
          "Content-Type": "text/csv;charset=utf-8",
        },
        body: exportData.csvContent,
      },
    );

    if (!response.ok) {
      const errorText = await response.text();
      console.warn("OneDrive upload failed", response.status, errorText);
      return { status: "failed" };
    }

    const uploadedItem = await response.json();
    return {
      status: "uploaded",
      location: `${getOneDriveUploadPathDisplay()}/${exportData.filename}`,
      webUrl: uploadedItem.webUrl ?? folder.webUrl ?? "",
    };
  } catch (error) {
    if (isInteractionRequiredError(error)) {
      return { status: "auth-required" };
    }
    console.warn("Unable to upload results to OneDrive", error);
    return { status: "failed" };
  }
}

async function acquireOneDriveAccessToken() {
  if (!state.oneDriveClient || !state.oneDriveAccount) {
    return "";
  }

  try {
    const tokenResponse = await state.oneDriveClient.acquireTokenSilent({
      account: state.oneDriveAccount,
      scopes: getOneDriveScopes(),
    });
    return tokenResponse.accessToken;
  } catch (error) {
    if (isInteractionRequiredError(error)) {
      console.warn("OneDrive token needs user interaction again", error);
      return "";
    }
    throw error;
  }
}

async function ensureOneDriveResultsFolder(accessToken) {
  if (state.oneDriveUploadFolderId) {
    return {
      id: state.oneDriveUploadFolderId,
      webUrl: state.oneDriveUploadFolderWebUrl,
    };
  }

  let currentFolder = await graphJson("/me/drive/root?$select=id,webUrl", accessToken);
  for (const segment of getOneDriveUploadPathSegments()) {
    const children = await graphJson(
      `/me/drive/items/${currentFolder.id}/children?$select=id,name,webUrl,folder`,
      accessToken,
    );

    let nextFolder = (children?.value ?? []).find((child) => child.name === segment && child.folder);
    if (!nextFolder) {
      nextFolder = await graphJson(`/me/drive/items/${currentFolder.id}/children`, accessToken, {
        method: "POST",
        headers: {
          "Content-Type": "application/json",
        },
        body: JSON.stringify({
          name: segment,
          folder: {},
          "@microsoft.graph.conflictBehavior": "fail",
        }),
      });
    }

    currentFolder = nextFolder;
  }

  state.oneDriveUploadFolderId = currentFolder.id ?? "";
  state.oneDriveUploadFolderWebUrl = currentFolder.webUrl ?? "";
  return currentFolder;
}

async function graphJson(path, accessToken, options = {}) {
  const response = await fetch(`https://graph.microsoft.com/v1.0${path}`, {
    method: options.method ?? "GET",
    headers: {
      Authorization: `Bearer ${accessToken}`,
      ...(options.headers ?? {}),
    },
    body: options.body,
  });

  if (!response.ok) {
    const errorText = await response.text();
    throw new Error(`Graph request failed (${response.status}): ${errorText}`);
  }

  return response.status === 204 ? null : response.json();
}

function encodeGraphPathSegment(value) {
  return encodeURIComponent(String(value ?? "")).replace(/%2F/g, "/");
}

function getOneDriveUploadPathSegments() {
  return String(oneDriveConfig.uploadPath ?? "IRDR/Results")
    .split("/")
    .map((segment) => segment.trim())
    .filter(Boolean);
}

function getOneDriveUploadPathDisplay() {
  return getOneDriveUploadPathSegments().join("/");
}

function isInteractionRequiredError(error) {
  return String(error?.name ?? "").includes("InteractionRequired")
    || String(error?.errorCode ?? "").includes("interaction_required");
}

async function shareExportData(exportData) {
  if (!state.supportsFileShare) {
    return { status: "unsupported" };
  }

  const payload = buildSharePayload(exportData);
  if (!payload) {
    return { status: "unsupported" };
  }

  try {
    await navigator.share(payload);
    return { status: "shared" };
  } catch (error) {
    if (error?.name === "AbortError") {
      return { status: "cancelled" };
    }
    console.warn("Unable to open the device share sheet", error);
    return { status: "failed" };
  }
}

function buildSharePayload(exportData) {
  if (!state.supportsFileShare || typeof File !== "function") {
    return null;
  }

  const shareFile = new File([exportData.csvContent], exportData.filename, {
    type: "text/csv;charset=utf-8",
  });
  const payload = {
    files: [shareFile],
    title: `IRDR Results • ${state.activeFacility}`,
    text: `IRDR results for ${state.activeFacility} (${getActiveWeek()?.label ?? ""})`,
  };

  try {
    return navigator.canShare?.(payload) ? payload : null;
  } catch (error) {
    console.warn("The device cannot share the generated CSV payload", error);
    return null;
  }
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

function generateDeviceId() {
  if (window.crypto?.randomUUID) {
    return window.crypto.randomUUID();
  }

  return `device-${Date.now().toString(36)}-${Math.random().toString(36).slice(2, 10)}`;
}

function supportsDirectorySave() {
  return window.isSecureContext && "showDirectoryPicker" in window && "indexedDB" in window;
}

function supportsFileShare() {
  return window.isSecureContext
    && typeof navigator.share === "function"
    && typeof navigator.canShare === "function"
    && typeof File === "function";
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

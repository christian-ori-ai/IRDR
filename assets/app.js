const state = {
  weeks: [],
  activeWeekId: "",
  activeFacility: "",
  currentIndex: 0,
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
};

document.addEventListener("DOMContentLoaded", init);

async function init() {
  registerServiceWorker();
  bindEvents();

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
  elements.previousLocation.addEventListener("click", () => moveLocation(-1));
  elements.nextLocation.addEventListener("click", () => moveLocation(1));
  elements.exportResults.addEventListener("click", exportResults);
  elements.resetSession.addEventListener("click", resetSession);
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

function moveLocation(direction) {
  saveCurrentEntry(true);

  const facility = getActiveFacility();
  if (!facility) {
    return;
  }

  state.currentIndex = clamp(state.currentIndex + direction, 0, facility.locations.length - 1);
  const week = getActiveWeek();
  const session = loadSession(week.id, state.activeFacility);
  session.currentIndex = state.currentIndex;
  saveSession(week.id, state.activeFacility, session);
  renderCountScreen();
}

function exportResults() {
  const week = getActiveWeek();
  const facility = getActiveFacility();
  if (!week || !facility) {
    return;
  }

  const session = loadSession(week.id, state.activeFacility);
  const lines = [
    ["week", "facility", "sample_order", "cost_center", "bin_code", "bin_description", "reviewed", "variance_count", "notes"],
  ];

  facility.locations.forEach((location) => {
    const entry = session.entries[String(location.sampleOrder)] ?? { varianceCount: 0, notes: "" };
    lines.push([
      week.id,
      state.activeFacility,
      location.sampleOrder,
      location.costCenter,
      location.binCode,
      location.binDescription ?? "",
      entry.reviewed ?? false,
      entry.varianceCount ?? 0,
      entry.notes ?? "",
    ]);
  });

  const csvContent = lines.map((row) => row.map(csvEscape).join(",")).join("\n");
  const blob = new Blob([csvContent], { type: "text/csv;charset=utf-8" });
  const url = URL.createObjectURL(blob);
  const anchor = document.createElement("a");
  anchor.href = url;
  anchor.download = `irdr-${week.id}-${state.activeFacility.toLowerCase()}.csv`;
  anchor.click();
  URL.revokeObjectURL(url);
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

function csvEscape(value) {
  const text = String(value ?? "");
  return `"${text.replace(/"/g, '""')}"`;
}

function clamp(value, min, max) {
  return Math.min(Math.max(value, min), max);
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

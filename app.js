(function () {
  "use strict";

  const STORAGE_KEY = "activity-log-state-v1";
  const AUTO_SAVE_INTERVAL_OPTIONS = [15, 30, 60, 120, 180, 360];
  const AUTO_SAVE_DEFAULT_INTERVAL = 60;
  const AUTO_SAVE_FILE_NAME = "activity-log-autosave.json";
  const AUTO_SAVE_DB_NAME = "activity-log-autosave";
  const AUTO_SAVE_DB_VERSION = 1;
  const AUTO_SAVE_DB_STORE = "handles";
  const AUTO_SAVE_HANDLE_KEY = "primary";
  const SNAP_OPTIONS = [5, 10, 15, 30, 60];
  const MINUTES_PER_DAY = 24 * 60;
  const ROWS_PER_PAGE = 15;
  const DRAG_START_THRESHOLD = 4;
  const HISTORY_LIMIT = 50;
  const HOLIDAY_API_PORT = 8765;
  const WEEKDAY_LABELS = ["日", "月", "火", "水", "木", "金", "土"];
  const ACTIVITY_TYPES = [
    { type: "食事", className: "type-meal", color: "#f0a34a", defaultDuration: 30 },
    { type: "外出", className: "type-outing", color: "#5aa1e3", defaultDuration: 60 },
    { type: "運動", className: "type-exercise", color: "#72be57", defaultDuration: 60 },
    { type: "作業", className: "type-work", color: "#9c77d9", defaultDuration: 60 },
    { type: "風呂", className: "type-bath", color: "#e57d75", defaultDuration: 15 },
    { type: "睡眠", className: "type-sleep", color: "#7286e6", defaultDuration: 360 },
  ];

  const ACTIVITY_META = Object.fromEntries(
    ACTIVITY_TYPES.map((item) => [item.type, item])
  );

  const state = {
    settings: buildDefaultSettings(),
    days: [],
    activities: [],
    selectedDayId: null,
    selectedActivityId: null,
    selectedPaletteType: ACTIVITY_TYPES[0].type,
    interaction: null,
    dialogActivityId: null,
    holidays: {},
    holidaySyncInfo: null,
    holidaySyncInFlight: false,
    historyPast: [],
    historyFuture: [],
    historyLimit: HISTORY_LIMIT,
    settingsDialogOriginSnapshot: null,
    settingsDialogCommitted: false,
    autoSaveHandle: null,
    autoSaveHandleName: "",
    autoSavePermission: "prompt",
    autoSaveStatusMessage: "",
    autoSaveLastSavedAt: "",
    autoSaveTimerId: null,
    autoSaveInFlight: false,
    autoSaveDirty: false,
  };

  const dom = {
    entryDate: document.getElementById("entry-date"),
    addDayButton: document.getElementById("add-day-button"),
    deleteDayButton: document.getElementById("delete-day-button"),
    snapSelect: document.getElementById("snap-select"),
    settingsButton: document.getElementById("settings-button"),
    saveButton: document.getElementById("save-button"),
    loadButton: document.getElementById("load-button"),
    loadInput: document.getElementById("load-input"),
    printButton: document.getElementById("print-button"),
    palette: document.getElementById("activity-palette"),
    activityEditorForm: document.getElementById("activity-editor-form"),
    activityEditorSummary: document.getElementById("activity-editor-summary"),
    activityEditorType: document.getElementById("activity-editor-type"),
    activityEditorDate: document.getElementById("activity-editor-date"),
    activityEditorTime: document.getElementById("activity-editor-time"),
    activityEditorDuration: document.getElementById("activity-editor-duration"),
    activityEditorApply: document.getElementById("activity-editor-apply"),
    selectedDaySummaryDate: document.getElementById("selected-day-summary-date"),
    selectedDaySummaryList: document.getElementById("selected-day-summary-list"),
    pages: document.getElementById("pages-container"),
    emptyState: document.getElementById("empty-state"),
    dialog: document.getElementById("activity-dialog"),
    dialogForm: document.getElementById("activity-form"),
    dialogTypeLabel: document.getElementById("dialog-type-label"),
    dialogDate: document.getElementById("dialog-date"),
    dialogTime: document.getElementById("dialog-time"),
    dialogDuration: document.getElementById("dialog-duration"),
    dialogCancelButton: document.getElementById("dialog-cancel-button"),
    settingsDialog: document.getElementById("settings-dialog"),
    settingsForm: document.getElementById("settings-form"),
    settingsMinFontSize: document.getElementById("settings-min-font-size"),
    settingsMinFontSizeValue: document.getElementById("settings-min-font-size-value"),
    settingsCornerRadius: document.getElementById("settings-corner-radius"),
    settingsCornerRadiusValue: document.getElementById("settings-corner-radius-value"),
    settingsSaturdayColor: document.getElementById("settings-saturday-color"),
    settingsSundayColor: document.getElementById("settings-sunday-color"),
    settingsHolidayColor: document.getElementById("settings-holiday-color"),
    settingsAutoSaveEnabled: document.getElementById("settings-autosave-enabled"),
    settingsAutoSaveInterval: document.getElementById("settings-autosave-interval"),
    settingsAutoSaveStatus: document.getElementById("settings-autosave-status"),
    autoSavePickButton: document.getElementById("autosave-pick-button"),
    autoSaveSaveNowButton: document.getElementById("autosave-save-now-button"),
    settingsHistoryStatus: document.getElementById("settings-history-status"),
    holidaySyncButton: document.getElementById("holiday-sync-button"),
    holidaySyncStatus: document.getElementById("holiday-sync-status"),
    settingsCancelButton: document.getElementById("settings-cancel-button"),
  };

  init().catch((error) => {
    console.error(error);
  });

  async function init() {
    bindEvents();
    const storageSnapshot = loadStateFromStorage();
    dom.snapSelect.value = String(state.settings.snapMinutes);
    applyVisualSettings();
    syncEntryDate();
    render();
    await initializeAutoSave(storageSnapshot);
  }

  function bindEvents() {
    dom.addDayButton.addEventListener("click", handleAddDay);
    dom.deleteDayButton.addEventListener("click", handleDeleteDay);
    dom.snapSelect.addEventListener("change", handleSnapChange);
    dom.settingsButton.addEventListener("click", handleOpenSettingsDialog);
    dom.saveButton.addEventListener("click", handleExportJson);
    dom.loadButton.addEventListener("click", () => dom.loadInput.click());
    dom.loadInput.addEventListener("change", handleImportJson);
    dom.printButton.addEventListener("click", () => window.print());

    dom.palette.addEventListener("click", handlePaletteClick);
    dom.palette.addEventListener("pointerdown", handlePalettePointerDown);

    dom.pages.addEventListener("click", handlePagesClick);
    dom.pages.addEventListener("dblclick", handlePageDoubleClick);
    dom.pages.addEventListener("pointerdown", handlePagesPointerDown);
    dom.pages.addEventListener("input", handlePagesInput);
    dom.pages.addEventListener("focusout", handlePagesFocusOut);
    dom.pages.addEventListener("focusin", handlePagesFocusIn);
    dom.pages.addEventListener("keydown", handlePagesKeyDown);

    document.addEventListener("pointermove", handlePointerMove);
    document.addEventListener("pointerup", handlePointerUp);
    document.addEventListener("keydown", handleDocumentKeyDown);

    dom.activityEditorForm.addEventListener("submit", handleActivityEditorSubmit);
    dom.dialogForm.addEventListener("submit", handleDialogSubmit);
    dom.dialogCancelButton.addEventListener("click", () => dom.dialog.close());
    dom.settingsForm.addEventListener("submit", handleSettingsSubmit);
    dom.settingsForm.addEventListener("input", handleSettingsInput);
    dom.settingsAutoSaveEnabled.addEventListener("change", handleSettingsInput);
    dom.settingsAutoSaveInterval.addEventListener("change", handleSettingsInput);
    dom.autoSavePickButton.addEventListener("click", handleChooseAutoSaveFile);
    dom.autoSaveSaveNowButton.addEventListener("click", handleAutoSaveNowClick);
    dom.holidaySyncButton.addEventListener("click", handleHolidaySyncClick);
    dom.settingsCancelButton.addEventListener("click", cancelSettingsDialog);
    dom.settingsDialog.addEventListener("cancel", handleSettingsDialogCancel);
    dom.settingsDialog.addEventListener("close", handleSettingsDialogClose);
    document.addEventListener("visibilitychange", handleDocumentVisibilityChange);
  }

  function buildDefaultSettings() {
    return {
      snapMinutes: 15,
      autoSaveEnabled: true,
      autoSaveIntervalMinutes: AUTO_SAVE_DEFAULT_INTERVAL,
      rowsPerPage: ROWS_PER_PAGE,
      printOrientation: "landscape",
      activityCornerRadius: 9,
      activityMinFontSize: 0.32,
      saturdayDateColor: "#e8f0ff",
      sundayDateColor: "#fde4e4",
      holidayDateColor: "#ffe6c8",
    };
  }

  function handleAddDay() {
    const dateValue = dom.entryDate.value || getSuggestedEntryDate();
    const historyPushed = pushHistorySnapshot();
    const result = addDay(dateValue, { select: true });
    if (!result.ok) {
      if (historyPushed) {
        discardLastHistorySnapshot();
      }
      window.alert(result.message);
      return;
    }
    syncEntryDate();
    saveAndRender();
  }

  function handleDeleteDay() {
    const day = getSelectedDay();
    if (!day) {
      window.alert("削除する日付エントリを選択してください。");
      return;
    }
    const accepted = window.confirm(`${day.date} を削除します。よろしいですか？`);
    if (!accepted) {
      return;
    }
    pushHistorySnapshot();
    removeDayById(day.id);
    syncEntryDate();
    saveAndRender();
  }

  function handleSnapChange() {
    const nextValue = Number(dom.snapSelect.value);
    if (!SNAP_OPTIONS.includes(nextValue) || nextValue === state.settings.snapMinutes) {
      return;
    }
    pushHistorySnapshot();
    state.settings.snapMinutes = nextValue;
    state.activities = normalizeActivities(state.activities);
    saveAndRender();
  }

  function handleExportJson() {
    const snapshot = exportState();
    const blob = new Blob([JSON.stringify(snapshot, null, 2)], {
      type: "application/json",
    });
    const url = URL.createObjectURL(blob);
    const anchor = document.createElement("a");
    anchor.href = url;
    anchor.download = `activity-log-${buildTimestampSlug(new Date())}.json`;
    document.body.appendChild(anchor);
    anchor.click();
    anchor.remove();
    URL.revokeObjectURL(url);
  }

  function handleImportJson(event) {
    const [file] = event.target.files || [];
    if (!file) {
      return;
    }
    const reader = new FileReader();
    reader.onload = () => {
      let historyPushed = false;
      try {
        const parsed = JSON.parse(String(reader.result || ""));
        historyPushed = pushHistorySnapshot();
        importState(parsed);
        syncEntryDate();
        saveAndRender();
        syncAutoSaveSchedule();
      } catch (error) {
        if (historyPushed) {
          discardLastHistorySnapshot();
        }
        console.error(error);
        window.alert("JSONの読み込みに失敗しました。形式を確認してください。");
      } finally {
        dom.loadInput.value = "";
      }
    };
    reader.readAsText(file, "utf-8");
  }

  function handlePaletteClick(event) {
    const button = event.target.closest(".palette-item");
    if (!button) {
      return;
    }
    const { type } = button.dataset;
    if (!ACTIVITY_META[type]) {
      return;
    }
    state.selectedPaletteType = type;
    renderPalette();
  }

  function handlePalettePointerDown(event) {
    if (event.button !== 0) {
      return;
    }
    const button = event.target.closest(".palette-item");
    if (!button) {
      return;
    }
    const { type } = button.dataset;
    if (!ACTIVITY_META[type]) {
      return;
    }
    state.selectedPaletteType = type;
    event.preventDefault();
    state.interaction = {
      mode: "create",
      type,
      hoverDayId: null,
      hoverRect: null,
      hoverMinute: null,
      releaseClientX: event.clientX,
      releaseClientY: event.clientY,
      previewActivity: null,
      valid: false,
    };
    renderPalette();
  }

  function handlePagesClick(event) {
    if (event.target.closest(".note-input")) {
      return;
    }
    const activityBlock = event.target.closest(".activity-block");
    if (activityBlock) {
      selectActivity(activityBlock.dataset.activityId, activityBlock.dataset.dayId, {
        renderNow: false,
        focusSurface: true,
      });
      return;
    }

    const dateButton = event.target.closest(".date-button");
    if (dateButton) {
      state.selectedActivityId = null;
      selectDay(dateButton.dataset.dayId);
      return;
    }

    const row = event.target.closest(".day-row");
    if (row) {
      state.selectedActivityId = null;
      selectDay(row.dataset.dayId);
    }
  }

  function handlePageDoubleClick(event) {
    const activityBlock = event.target.closest(".activity-block");
    if (!activityBlock) {
      return;
    }
    const activity = getActivityById(activityBlock.dataset.activityId);
    if (!activity) {
      return;
    }
    selectActivity(activity.id, activityBlock.dataset.dayId, {
      renderNow: false,
      focusSurface: true,
    });
    openActivityDialog(activity);
  }

  function handlePagesPointerDown(event) {
    if (event.button !== 0) {
      return;
    }
    const handle = event.target.closest(".resize-handle");
    const activityBlock = event.target.closest(".activity-block");

    if (handle && activityBlock) {
      const activity = getActivityById(activityBlock.dataset.activityId);
      const hostTimeline = activityBlock.closest(".timeline");
      const timelineRect = hostTimeline ? hostTimeline.getBoundingClientRect() : null;
      if (!activity || !timelineRect) {
        return;
      }
      selectActivity(activity.id, activityBlock.dataset.dayId, {
        renderNow: false,
        focusSurface: true,
      });
      event.preventDefault();
      state.interaction = {
        mode: "resize",
        edge: handle.dataset.edge,
        activityId: activity.id,
        rect: timelineRect,
        anchorClientX: event.clientX,
        originalStart: activity.start,
        originalDuration: activity.durationMinutes,
        activityType: activity.type,
        previewActivity: { ...activity },
        valid: true,
      };
      render();
      return;
    }

    if (activityBlock) {
      const activity = getActivityById(activityBlock.dataset.activityId);
      const hostTimeline = activityBlock.closest(".timeline");
      const timelineRect = hostTimeline ? hostTimeline.getBoundingClientRect() : null;
      if (!activity || !timelineRect) {
        return;
      }
      selectActivity(activity.id, activityBlock.dataset.dayId, {
        renderNow: false,
        focusSurface: true,
      });
      event.preventDefault();
      state.interaction = {
        mode: "move-pending",
        activityId: activity.id,
        rect: timelineRect,
        anchorClientX: event.clientX,
        anchorClientY: event.clientY,
        originalStart: activity.start,
        originalDuration: activity.durationMinutes,
        activityType: activity.type,
        previewActivity: { ...activity },
        valid: true,
      };
      return;
    }
  }

  function handlePagesInput(event) {
    const input = event.target.closest(".note-input");
    if (!input) {
      return;
    }
    const day = getDayById(input.dataset.dayId);
    if (!day) {
      return;
    }
    day.note = input.value;
    persistState();
    syncPrintNote(input.dataset.dayId, input.value);
  }

  function handlePagesFocusOut(event) {
    const input = event.target.closest(".note-input");
    if (!input) {
      return;
    }
    commitNoteInput(input);
  }

  function handlePagesFocusIn(event) {
    const input = event.target.closest(".note-input");
    if (!input) {
      return;
    }
    input.dataset.originalValue = input.value;
    state.selectedActivityId = null;
    selectDay(input.dataset.dayId, { renderNow: false });
    renderRowsSelectionOnly();
    renderActivitySelectionOnly();
    renderSelectedActivityEditor();
  }

  function handlePagesKeyDown(event) {
    const input = event.target.closest(".note-input");
    if (!input) {
      return;
    }
    if (event.key === "Enter") {
      event.preventDefault();
      input.blur();
    }
  }

  function handlePointerMove(event) {
    if (!state.interaction) {
      return;
    }
    if (state.interaction.mode === "move-pending") {
      const movedX = event.clientX - state.interaction.anchorClientX;
      const movedY = event.clientY - state.interaction.anchorClientY;
      if (Math.hypot(movedX, movedY) < DRAG_START_THRESHOLD) {
        return;
      }
      state.interaction.mode = "move";
    }
    if (state.interaction.mode === "create") {
      updateCreateInteraction(event.clientX, event.clientY);
      render();
      return;
    }
    if (state.interaction.mode === "move") {
      updateMoveInteraction(event.clientX);
      render();
      return;
    }
    if (state.interaction.mode === "resize") {
      updateResizeInteraction(event.clientX);
      render();
    }
  }

  function handlePointerUp() {
    if (!state.interaction) {
      return;
    }

    const interaction = state.interaction;
    state.interaction = null;

    if (interaction.mode === "move-pending") {
      syncSelectionUi();
      return;
    }

    if (interaction.mode === "create") {
      if (!interaction.previewActivity) {
        render();
        return;
      }
      const candidate = buildCreateCandidateFromPoint(
        interaction,
        interaction.releaseClientX,
        interaction.releaseClientY
      );
      if (!candidate) {
        render();
        return;
      }
      const result = upsertActivity(candidate);
      if (!result.ok) {
        window.alert(result.message);
      }
      render();
      return;
    }

    if (interaction.mode === "move" || interaction.mode === "resize") {
      const result = upsertActivity(interaction.previewActivity, {
        ignoreId: interaction.activityId,
      });
      if (!result.ok) {
        window.alert(result.message);
      }
      render();
    }
  }

  function handleDocumentKeyDown(event) {
    if (event.key !== "Delete" && event.key !== "Del") {
      if (event.ctrlKey || event.metaKey) {
        const lowerKey = String(event.key).toLowerCase();
        if (lowerKey === "z") {
          event.preventDefault();
          commitPendingInputBeforeHistory();
          if (!dom.settingsDialog.open && !dom.dialog.open) {
            performUndo();
          }
          return;
        }
        if (lowerKey === "y") {
          event.preventDefault();
          commitPendingInputBeforeHistory();
          if (!dom.settingsDialog.open && !dom.dialog.open) {
            performRedo();
          }
          return;
        }
      }
      return;
    }
    const activeElement = document.activeElement;
    const activeTag = activeElement ? activeElement.tagName : "";
    if (
      activeTag === "INPUT" ||
      activeTag === "TEXTAREA" ||
      activeTag === "SELECT" ||
      activeTag === "BUTTON" ||
      dom.dialog.open
    ) {
      return;
    }
    const activity = getSelectedActivity();
    if (!activity) {
      return;
    }
    pushHistorySnapshot();
    state.activities = state.activities.filter((item) => item.id !== activity.id);
    state.selectedActivityId = null;
    saveAndRender();
  }

  function handleActivityEditorSubmit(event) {
    event.preventDefault();
    const activity = getSelectedActivity();
    if (!activity) {
      window.alert("編集する行動ログをクリックで選択してください。");
      return;
    }

    const startDate = dom.activityEditorDate.value;
    const startTime = dom.activityEditorTime.value;
    const durationValue = Number(dom.activityEditorDuration.value);

    if (!startDate || !startTime || !Number.isFinite(durationValue)) {
      window.alert("開始日、開始時刻、長さを入力してください。");
      return;
    }

    const candidate = {
      id: activity.id,
      type: activity.type,
      start: snapDateTimeString(`${startDate}T${startTime}`),
      durationMinutes: snapDuration(durationValue),
    };

    const result = upsertActivity(candidate, { ignoreId: activity.id });
    if (!result.ok) {
      window.alert(result.message);
    }
  }

  function handleOpenSettingsDialog() {
    state.settingsDialogOriginSnapshot = createHistorySnapshot();
    state.settingsDialogCommitted = false;
    syncSettingsDialogFields();
    renderSettingsHistoryStatus();
    dom.settingsDialog.showModal();
  }

  function handleSettingsInput() {
    state.settings.activityMinFontSize = normalizeFontSizeSetting(dom.settingsMinFontSize.value);
    state.settings.activityCornerRadius = normalizeCornerRadiusSetting(dom.settingsCornerRadius.value);
    state.settings.saturdayDateColor = normalizeColorSetting(dom.settingsSaturdayColor.value, "#e8f0ff");
    state.settings.sundayDateColor = normalizeColorSetting(dom.settingsSundayColor.value, "#fde4e4");
    state.settings.holidayDateColor = normalizeColorSetting(dom.settingsHolidayColor.value, "#ffe6c8");
    state.settings.autoSaveEnabled = dom.settingsAutoSaveEnabled.value === "enabled";
    state.settings.autoSaveIntervalMinutes = normalizeAutoSaveIntervalSetting(
      dom.settingsAutoSaveInterval.value
    );
    state.autoSaveStatusMessage = "";
    syncSettingsDialogFields();
    syncAutoSaveSchedule();
    applyVisualSettings();
    render();
  }

  function handleSettingsSubmit(event) {
    event.preventDefault();
    if (
      state.settingsDialogOriginSnapshot &&
      !areSnapshotsEqual(state.settingsDialogOriginSnapshot, createHistorySnapshot())
    ) {
      pushHistorySnapshot(state.settingsDialogOriginSnapshot);
    }
    state.settingsDialogCommitted = true;
    persistState();
    syncAutoSaveSchedule();
    render();
    dom.settingsDialog.close();
  }

  function handleSettingsDialogCancel(event) {
    event.preventDefault();
    cancelSettingsDialog();
  }

  function handleSettingsDialogClose() {
    if (!state.settingsDialogCommitted && state.settingsDialogOriginSnapshot) {
      applyHistorySnapshot(state.settingsDialogOriginSnapshot);
      persistState();
      syncAutoSaveSchedule();
      render();
    }
    state.settingsDialogOriginSnapshot = null;
    state.settingsDialogCommitted = false;
  }

  function cancelSettingsDialog() {
    state.settingsDialogCommitted = false;
    dom.settingsDialog.close();
  }

  async function handleHolidaySyncClick() {
    if (state.holidaySyncInFlight) {
      return;
    }

    const years = getHolidaySyncYears();
    state.holidaySyncInFlight = true;
    state.holidaySyncInfo = {
      ...(state.holidaySyncInfo || {}),
      status: "syncing",
      message: `${years.join(", ")}年の祝日を同期中です。`,
    };
    renderHolidaySyncStatus();

    try {
      const payload = await fetchHolidayData(years);
      const historySnapshot = createHistorySnapshot();
      const nextHolidays = mergeHolidayEntries(state.holidays, payload.holidays, years, payload);
      if (!areHolidayMapsEqual(state.holidays, nextHolidays)) {
        pushHistorySnapshot(historySnapshot);
      }
      state.holidays = nextHolidays;
      state.holidaySyncInfo = {
        status: "success",
        source: payload.source,
        fetchedAt: payload.fetchedAt,
        syncedYears: years,
        count: payload.holidays.length,
        message: `${years.join(", ")}年の祝日を同期しました。`,
      };
      saveAndRender();
    } catch (error) {
      console.error(error);
      state.holidaySyncInfo = {
        ...(state.holidaySyncInfo || {}),
        status: "error",
        syncedYears: years,
        message:
          "祝日同期に失敗しました。`python server.py` を起動してから、もう一度試してください。",
      };
      renderHolidaySyncStatus();
      window.alert("祝日同期に失敗しました。ローカルサーバーが起動しているか確認してください。");
    } finally {
      state.holidaySyncInFlight = false;
      renderHolidaySyncStatus();
    }
  }

  async function handleChooseAutoSaveFile() {
    if (!isAutoSaveSupported()) {
      window.alert("このブラウザでは単一ファイルの自動保存に対応していません。");
      return;
    }

    try {
      const handle = await window.showSaveFilePicker({
        suggestedName: state.autoSaveHandleName || AUTO_SAVE_FILE_NAME,
        types: [
          {
            description: "JSON files",
            accept: {
              "application/json": [".json"],
            },
          },
        ],
      });

      state.autoSaveHandle = handle;
      state.autoSaveHandleName = typeof handle.name === "string" ? handle.name : AUTO_SAVE_FILE_NAME;
      state.autoSaveStatusMessage = "";
      state.settings.autoSaveEnabled = true;
      state.settings.autoSaveIntervalMinutes = normalizeAutoSaveIntervalSetting(
        state.settings.autoSaveIntervalMinutes
      );

      try {
        await saveAutoSaveHandleToDb(handle);
      } catch (error) {
        console.error(error);
        state.autoSaveStatusMessage =
          "保存先の記憶に失敗しました。今回の起動中は自動保存できます。";
      }

      const saved = await saveSnapshotToAutoSaveFile({
        reason: "manual",
        requestPermission: true,
        force: true,
      });
      persistState({ markAutoSaveDirty: !saved });
      syncAutoSaveSchedule();
      render();
    } catch (error) {
      if (error && error.name === "AbortError") {
        return;
      }
      console.error(error);
      state.autoSaveStatusMessage = "自動保存先の設定に失敗しました。";
      renderAutoSaveStatus();
      window.alert("自動保存先の設定に失敗しました。");
    }
  }

  async function handleAutoSaveNowClick() {
    const saved = await saveSnapshotToAutoSaveFile({
      reason: "manual",
      requestPermission: true,
      force: true,
    });
    if (saved) {
      render();
    }
  }

  function handleDocumentVisibilityChange() {
    if (document.visibilityState !== "hidden") {
      return;
    }
    void saveSnapshotToAutoSaveFile({
      reason: "hidden",
      requestPermission: false,
      force: false,
    });
  }

  function handleDialogSubmit(event) {
    event.preventDefault();
    const activity = getActivityById(state.dialogActivityId);
    if (!activity) {
      dom.dialog.close();
      return;
    }

    const startDate = dom.dialogDate.value;
    const startTime = dom.dialogTime.value;
    const durationValue = Number(dom.dialogDuration.value);

    if (!startDate || !startTime || !Number.isFinite(durationValue)) {
      window.alert("開始日、開始時刻、長さを入力してください。");
      return;
    }

    const candidate = {
      id: activity.id,
      type: activity.type,
      start: snapDateTimeString(`${startDate}T${startTime}`),
      durationMinutes: snapDuration(durationValue),
    };

    const result = upsertActivity(candidate, { ignoreId: activity.id });
    if (!result.ok) {
      window.alert(result.message);
      return;
    }
    dom.dialog.close();
  }

  function render() {
    applyVisualSettings();
    renderPalette();
    renderPages();
    updateControls();
  }

  function renderPalette() {
    dom.palette.innerHTML = ACTIVITY_TYPES.map((item) => {
      const isSelected = item.type === state.selectedPaletteType;
      return `
        <button
          type="button"
          class="palette-item ${isSelected ? "is-selected" : ""}"
          data-type="${escapeAttr(item.type)}"
          title="${escapeAttr(item.type)}"
        >
          <span class="palette-chip" style="background:${item.color}"></span>
          <span>${escapeHtml(item.type)}</span>
        </button>
      `;
    }).join("");
  }

  function renderPages() {
    const days = getSortedDays();
    dom.emptyState.hidden = days.length > 0;
    if (!days.length) {
      dom.pages.innerHTML = "";
      return;
    }

    const snapCount = MINUTES_PER_DAY / state.settings.snapMinutes;
    const pages = chunk(days, state.settings.rowsPerPage);
    const renderableActivities = getRenderableActivities();

    dom.pages.innerHTML = pages
      .map((pageDays, pageIndex) => {
        const yearLabel = buildYearLabel(pageDays);
        const filledDays = [...pageDays];
        while (filledDays.length < state.settings.rowsPerPage) {
          filledDays.push(null);
        }
        return `
          <section class="page" data-page-index="${pageIndex}">
            <div class="page-header">
              <div class="page-years">${escapeHtml(yearLabel)}</div>
              <div class="time-ruler">
                <div class="time-ruler-label">月/日</div>
                <div class="ruler-track">
                  ${renderHourLabels()}
                </div>
                <div class="time-ruler-label">備考</div>
              </div>
            </div>
            <div class="page-rows">
              ${filledDays
                .map((day) => {
                  if (!day) {
                    return renderPlaceholderRow(snapCount);
                  }
                  const segments = getSegmentsForDay(day, renderableActivities);
                  return renderDayRow(day, segments, snapCount);
                })
                .join("")}
            </div>
          </section>
        `;
      })
      .join("");
  }

  function renderDayRow(day, segments, snapCount) {
    const isSelected = day.id === state.selectedDayId;
    const dateButtonState = getDateButtonState(day.date);
    return `
      <div class="day-row ${isSelected ? "is-selected" : ""}" data-day-id="${escapeAttr(day.id)}">
        <button
          type="button"
          class="date-button ${escapeAttr(dateButtonState.className)}"
          data-day-id="${escapeAttr(day.id)}"
          title="${escapeAttr(dateButtonState.title)}"
        >
          ${escapeHtml(formatMonthDay(day.date))}
        </button>
        <div
          class="timeline ${isSelected ? "is-selected" : ""}"
          data-day-id="${escapeAttr(day.id)}"
          style="--snap-count:${snapCount};"
        >
          ${segments.join("")}
        </div>
        <div class="note-cell">
          <input
            class="note-input"
            data-day-id="${escapeAttr(day.id)}"
            type="text"
            value="${escapeAttr(day.note || "")}"
            placeholder="備考"
          />
          <div class="note-print" data-note-print-for="${escapeAttr(day.id)}">${escapeHtml(
      day.note || ""
    )}</div>
        </div>
      </div>
    `;
  }

  function renderPlaceholderRow(snapCount) {
    return `
      <div class="day-row-placeholder">
        <button type="button" class="date-button" tabindex="-1" disabled>--/--</button>
        <div class="timeline" style="--snap-count:${snapCount};"></div>
        <div class="note-cell">
          <input class="note-input" type="text" tabindex="-1" disabled placeholder="備考" />
        </div>
      </div>
    `;
  }

  function renderHourLabels() {
    return Array.from({ length: 25 }, (_, hour) => {
      const pct = (hour / 24) * 100;
      let edgeClass = "";
      if (hour === 0) {
        edgeClass = "is-edge-left";
      } else if (hour === 24) {
        edgeClass = "is-edge-right";
      }
      return `<span class="ruler-hour ${edgeClass}" style="left:${pct}%">${hour}</span>`;
    }).join("");
  }

  function renderRowsSelectionOnly() {
    document.querySelectorAll(".day-row").forEach((row) => {
      row.classList.toggle("is-selected", row.dataset.dayId === state.selectedDayId);
      const timeline = row.querySelector(".timeline");
      if (timeline) {
        timeline.classList.toggle("is-selected", row.dataset.dayId === state.selectedDayId);
      }
    });
  }

  function updateControls() {
    dom.deleteDayButton.disabled = !state.selectedDayId;
    dom.snapSelect.value = String(state.settings.snapMinutes);
    renderSelectedActivityEditor();
    renderSelectedDaySummary();
    syncSettingsDialogFields();
    renderSettingsHistoryStatus();
    renderAutoSaveStatus();
    renderHolidaySyncStatus();
  }

  function renderActivitySelectionOnly() {
    document.querySelectorAll(".activity-block").forEach((block) => {
      block.classList.toggle("is-selected", block.dataset.activityId === state.selectedActivityId);
    });
  }

  function renderSelectedActivityEditor() {
    const activity = getSelectedActivity();
    const disabled = !activity;

    dom.activityEditorType.value = activity ? activity.type : "";
    dom.activityEditorDate.value = activity ? activity.start.slice(0, 10) : "";
    dom.activityEditorTime.value = activity ? activity.start.slice(11, 16) : "";
    dom.activityEditorDuration.value = activity ? String(activity.durationMinutes) : "";
    dom.activityEditorSummary.textContent = activity
      ? `開始時刻 ${activity.start.slice(11, 16)} / 長さ ${activity.durationMinutes} 分`
      : "行動ログをクリックすると、開始日・開始時刻・長さをここで文字入力して反映できます。";

    [
      dom.activityEditorType,
      dom.activityEditorDate,
      dom.activityEditorTime,
      dom.activityEditorDuration,
      dom.activityEditorApply,
    ].forEach((element) => {
      element.disabled = disabled;
    });
  }

  function renderSelectedDaySummary() {
    if (!dom.selectedDaySummaryDate || !dom.selectedDaySummaryList) {
      return;
    }
    const day = getSelectedDay();
    if (!day) {
      dom.selectedDaySummaryDate.textContent =
        "日付エントリを選択すると、ここに集計を表示します。";
      dom.selectedDaySummaryList.innerHTML =
        '<div class="day-summary-empty">まだ選択されている日付がありません。</div>';
      return;
    }

    const summaryItems = summarizeActivitiesForDay(day);
    dom.selectedDaySummaryDate.textContent = `${day.date} の行動時間`;
    if (!summaryItems.length) {
      dom.selectedDaySummaryList.innerHTML =
        '<div class="day-summary-empty">この日はまだ行動ログがありません。</div>';
      return;
    }

    const totalMinutes = summaryItems.reduce((sum, item) => sum + item.minutes, 0);
    dom.selectedDaySummaryList.innerHTML = `
      <div class="day-summary-item">
        <span class="day-summary-label">合計</span>
        <span class="day-summary-duration">${escapeHtml(formatDurationMinutes(totalMinutes))}</span>
      </div>
      ${summaryItems
        .map(
          (item) => `
            <div class="day-summary-item">
              <span class="day-summary-label">${escapeHtml(item.type)}</span>
              <span class="day-summary-duration">${escapeHtml(formatDurationMinutes(item.minutes))}</span>
            </div>
          `
        )
        .join("")}
    `;
  }

  function syncSettingsDialogFields() {
    if (!dom.settingsMinFontSize) {
      return;
    }
    dom.settingsMinFontSize.value = String(state.settings.activityMinFontSize);
    dom.settingsCornerRadius.value = String(state.settings.activityCornerRadius);
    dom.settingsSaturdayColor.value = normalizeColorSetting(
      state.settings.saturdayDateColor,
      "#e8f0ff"
    );
    dom.settingsSundayColor.value = normalizeColorSetting(
      state.settings.sundayDateColor,
      "#fde4e4"
    );
    dom.settingsHolidayColor.value = normalizeColorSetting(
      state.settings.holidayDateColor,
      "#ffe6c8"
    );
    dom.settingsAutoSaveEnabled.value = state.settings.autoSaveEnabled ? "enabled" : "disabled";
    dom.settingsAutoSaveInterval.value = String(
      normalizeAutoSaveIntervalSetting(state.settings.autoSaveIntervalMinutes)
    );
    dom.settingsMinFontSizeValue.textContent = `${state.settings.activityMinFontSize.toFixed(2)}rem`;
    dom.settingsCornerRadiusValue.textContent = `${state.settings.activityCornerRadius}px`;
  }

  function renderSettingsHistoryStatus() {
    if (!dom.settingsHistoryStatus) {
      return;
    }
    dom.settingsHistoryStatus.textContent = `戻せる: ${state.historyPast.length} / やり直し: ${state.historyFuture.length}`;
  }

  function renderHolidaySyncStatus() {
    if (!dom.holidaySyncStatus || !dom.holidaySyncButton) {
      return;
    }
    if (state.holidaySyncInFlight) {
      dom.holidaySyncStatus.textContent =
        state.holidaySyncInfo?.message || "祝日を同期中です。";
    } else if (state.holidaySyncInfo?.message) {
      dom.holidaySyncStatus.textContent = state.holidaySyncInfo.message;
    } else {
      dom.holidaySyncStatus.textContent = "未同期です。";
    }
    dom.holidaySyncButton.disabled = state.holidaySyncInFlight;
  }

  function renderAutoSaveStatus() {
    if (
      !dom.settingsAutoSaveStatus ||
      !dom.settingsAutoSaveEnabled ||
      !dom.settingsAutoSaveInterval ||
      !dom.autoSavePickButton ||
      !dom.autoSaveSaveNowButton
    ) {
      return;
    }

    const supported = isAutoSaveSupported();
    dom.settingsAutoSaveEnabled.disabled = !supported;
    dom.settingsAutoSaveInterval.disabled = !supported || !state.settings.autoSaveEnabled;
    dom.autoSavePickButton.disabled = !supported || state.autoSaveInFlight;
    dom.autoSaveSaveNowButton.disabled =
      !supported || !state.autoSaveHandle || state.autoSaveInFlight;

    if (state.autoSaveInFlight) {
      dom.settingsAutoSaveStatus.textContent = "自動保存ファイルへ保存中です。";
      return;
    }

    if (state.autoSaveStatusMessage) {
      dom.settingsAutoSaveStatus.textContent = state.autoSaveStatusMessage;
      return;
    }

    dom.settingsAutoSaveStatus.textContent = buildAutoSaveStatusMessage();
  }

  function applyVisualSettings() {
    document.documentElement.style.setProperty(
      "--activity-corner-radius",
      `${state.settings.activityCornerRadius}px`
    );
    document.documentElement.style.setProperty(
      "--entry-corner-radius",
      `${state.settings.activityCornerRadius}px`
    );
    document.documentElement.style.setProperty(
      "--saturday-date-color",
      normalizeColorSetting(state.settings.saturdayDateColor, "#e8f0ff")
    );
    document.documentElement.style.setProperty(
      "--sunday-date-color",
      normalizeColorSetting(state.settings.sundayDateColor, "#fde4e4")
    );
    document.documentElement.style.setProperty(
      "--holiday-date-color",
      normalizeColorSetting(state.settings.holidayDateColor, "#ffe6c8")
    );
  }

  function selectDay(dayId, options = {}) {
    state.selectedDayId = dayId;
    syncEntryDate();
    if (options.renderNow !== false) {
      render();
    }
  }

  function selectActivity(activityId, dayId, options = {}) {
    state.selectedActivityId = activityId;
    if (dayId) {
      state.selectedDayId = dayId;
    }
    syncEntryDate();
    if (options.renderNow === false) {
      syncSelectionUi({ focusSurface: options.focusSurface !== false });
      return;
    }
    render();
    if (options.focusSurface !== false) {
      focusInteractionSurface();
    }
  }

  function syncSelectionUi(options = {}) {
    renderRowsSelectionOnly();
    renderActivitySelectionOnly();
    updateControls();
    if (options.focusSurface) {
      focusInteractionSurface();
    }
  }

  function focusInteractionSurface() {
    if (!dom.pages) {
      return;
    }
    try {
      dom.pages.focus({ preventScroll: true });
    } catch (error) {
      dom.pages.focus();
    }
  }

  function addDay(dateValue, options = {}) {
    if (!isValidDateString(dateValue)) {
      return { ok: false, message: "日付が不正です。" };
    }
    if (findDayByDate(dateValue)) {
      return { ok: false, message: "同じ日付のエントリは作成できません。" };
    }
    const day = {
      id: createId("day"),
      date: dateValue,
      note: "",
    };
    state.days.push(day);
    sortStateDays();
    if (options.select) {
      state.selectedDayId = day.id;
      state.selectedActivityId = null;
    }
    return { ok: true, day };
  }

  function removeDayById(dayId) {
    const day = getDayById(dayId);
    if (!day) {
      return;
    }
    const sortedDays = getSortedDays();
    const removedIndex = sortedDays.findIndex((item) => item.id === dayId);
    trimActivitiesByDay(day.date);
    state.days = state.days.filter((item) => item.id !== dayId);
    state.selectedActivityId = null;
    const nextDays = getSortedDays();
    if (!nextDays.length) {
      state.selectedDayId = null;
      return;
    }
    const fallbackIndex = clamp(removedIndex, 0, nextDays.length - 1);
    state.selectedDayId = nextDays[fallbackIndex].id;
  }

  function trimActivitiesByDay(dateValue) {
    const dayStart = parseDateString(dateValue);
    const dayEnd = addDays(dayStart, 1);
    const nextActivities = [];

    state.activities.forEach((activity) => {
      const start = parseDateTimeString(activity.start);
      const end = addMinutes(start, activity.durationMinutes);
      if (end <= dayStart || start >= dayEnd) {
        nextActivities.push(activity);
        return;
      }

      if (start < dayStart) {
        nextActivities.push({
          ...activity,
          durationMinutes: differenceInMinutes(dayStart, start),
        });
      }

      if (end > dayEnd) {
        nextActivities.push({
          ...activity,
          id: createId("activity"),
          start: formatDateTimeString(dayEnd),
          durationMinutes: differenceInMinutes(end, dayEnd),
        });
      }
    });

    state.activities = normalizeActivities(nextActivities);
  }

  function upsertActivity(candidate, options = {}) {
    const ignoreId = options.ignoreId || null;
    const recordHistory = options.recordHistory !== false;
    const normalized = normalizeActivity(candidate);
    if (!normalized.ok) {
      return normalized;
    }
    if (hasActivityOverlap(normalized.activity, ignoreId)) {
      return { ok: false, message: "行動ログの時間帯が重なっています。" };
    }

    if (recordHistory) {
      pushHistorySnapshot();
    }
    ensureDaysForActivity(normalized.activity);

    if (ignoreId) {
      state.activities = state.activities.map((item) =>
        item.id === ignoreId ? normalized.activity : item
      );
    } else {
      state.activities.push(normalized.activity);
    }

    state.activities = normalizeActivities(state.activities);
    state.selectedActivityId = normalized.activity.id;
    saveAndRender();
    return { ok: true, activity: normalized.activity };
  }

  function normalizeActivity(candidate) {
    if (!candidate || !ACTIVITY_META[candidate.type]) {
      return { ok: false, message: "行動種別が不正です。" };
    }
    if (!candidate.start || !isValidDateTimeString(candidate.start)) {
      return { ok: false, message: "開始日時が不正です。" };
    }
    const duration = snapDuration(candidate.durationMinutes);
    if (!Number.isFinite(duration) || duration <= 0) {
      return { ok: false, message: "長さが不正です。" };
    }
    return {
      ok: true,
      activity: {
        id: candidate.id || createId("activity"),
        type: candidate.type,
        start: snapDateTimeString(candidate.start),
        durationMinutes: duration,
      },
    };
  }

  function normalizeActivities(items) {
    return [...items].sort((left, right) => left.start.localeCompare(right.start));
  }

  function hasActivityOverlap(candidate, ignoreId) {
    const candidateStart = parseDateTimeString(candidate.start);
    const candidateEnd = addMinutes(candidateStart, candidate.durationMinutes);

    return state.activities.some((activity) => {
      if (activity.id === ignoreId) {
        return false;
      }
      const start = parseDateTimeString(activity.start);
      const end = addMinutes(start, activity.durationMinutes);
      return candidateStart < end && start < candidateEnd;
    });
  }

  function ensureDaysForActivity(activity) {
    const start = parseDateTimeString(activity.start);
    const end = addMinutes(start, activity.durationMinutes);
    let cursor = startOfDay(start);
    const lastCovered = startOfDay(addMinutes(end, -1));

    while (cursor <= lastCovered) {
      const dateValue = formatDateString(cursor);
      if (!findDayByDate(dateValue)) {
        addDay(dateValue);
      }
      cursor = addDays(cursor, 1);
    }
    sortStateDays();
  }

  function updateCreateInteraction(clientX, clientY) {
    if (!state.interaction || state.interaction.mode !== "create") {
      return;
    }
    updatePaletteCreateInteraction(clientX, clientY);
  }

  function updatePaletteCreateInteraction(clientX, clientY) {
    const interaction = state.interaction;
    const timeline = getTimelineFromPoint(clientX, clientY);

    interaction.hoverDayId = timeline ? timeline.dataset.dayId : null;
    interaction.hoverRect = timeline ? timeline.getBoundingClientRect() : null;
    interaction.releaseClientX = clientX;
    interaction.releaseClientY = clientY;

    if (!timeline || !interaction.hoverDayId || !interaction.hoverRect) {
      interaction.hoverMinute = null;
      interaction.previewActivity = null;
      interaction.valid = false;
      return;
    }

    interaction.hoverMinute = getSnappedMinuteFromClientX(
      interaction.hoverRect,
      clientX,
      { allowOverflow: false, clampInside: true }
    );

    const previewActivity = buildCreatePreviewForDay(
      interaction.type,
      interaction.hoverDayId,
      interaction.hoverMinute,
      interaction.hoverMinute
    );

    interaction.previewActivity = previewActivity;
    interaction.valid = !!previewActivity && !hasActivityOverlap(previewActivity, null);
  }

  function buildCreateCandidateFromPoint(interaction, clientX, clientY) {
    const timeline = getTimelineFromPoint(clientX, clientY);
    if (!timeline) {
      return null;
    }

    const hoverDayId = timeline.dataset.dayId;
    const hoverRect = timeline.getBoundingClientRect();
    const hoverMinute = getSnappedMinuteFromClientX(hoverRect, clientX, {
      allowOverflow: false,
      clampInside: true,
    });

    return buildCreatePreviewForDay(interaction.type, hoverDayId, hoverMinute, hoverMinute);
  }

  function buildCreatePreviewForDay(type, dayId, anchorMinute, currentMinute) {
    const day = getDayById(dayId);
    if (!day || !ACTIVITY_META[type]) {
      return null;
    }

    const defaultDuration = ACTIVITY_META[type].defaultDuration;
    let startMinute;
    let endMinute;
    if (Math.abs(currentMinute - anchorMinute) < state.settings.snapMinutes) {
      startMinute = anchorMinute;
      endMinute = anchorMinute + defaultDuration;
    } else {
      startMinute = Math.min(anchorMinute, currentMinute);
      endMinute = Math.max(anchorMinute, currentMinute);
    }

    const dayStart = parseDateString(day.date);
    const start = addMinutes(dayStart, startMinute);
    return {
      id: createId("preview"),
      type,
      start: formatDateTimeString(start),
      durationMinutes: Math.max(state.settings.snapMinutes, endMinute - startMinute),
    };
  }

  function updateMoveInteraction(clientX) {
    const interaction = state.interaction;
    if (!interaction || interaction.mode !== "move") {
      return;
    }
    const deltaMinutes = snapToGrid(
      ((clientX - interaction.anchorClientX) / interaction.rect.width) * MINUTES_PER_DAY,
      state.settings.snapMinutes
    );
    const nextStart = addMinutes(parseDateTimeString(interaction.originalStart), deltaMinutes);
    const previewActivity = {
      id: interaction.activityId,
      type: interaction.activityType,
      start: formatDateTimeString(nextStart),
      durationMinutes: interaction.originalDuration,
    };
    interaction.previewActivity = previewActivity;
    interaction.valid = !hasActivityOverlap(previewActivity, interaction.activityId);
  }

  function updateResizeInteraction(clientX) {
    const interaction = state.interaction;
    if (!interaction || interaction.mode !== "resize") {
      return;
    }
    const deltaMinutes = snapToGrid(
      ((clientX - interaction.anchorClientX) / interaction.rect.width) * MINUTES_PER_DAY,
      state.settings.snapMinutes
    );
    const originalStart = parseDateTimeString(interaction.originalStart);
    const originalEnd = addMinutes(originalStart, interaction.originalDuration);
    let nextStart = originalStart;
    let nextEnd = originalEnd;

    if (interaction.edge === "left") {
      nextStart = addMinutes(originalStart, deltaMinutes);
      if (nextStart >= addMinutes(originalEnd, -state.settings.snapMinutes)) {
        nextStart = addMinutes(originalEnd, -state.settings.snapMinutes);
      }
    } else {
      nextEnd = addMinutes(originalEnd, deltaMinutes);
      if (nextEnd <= addMinutes(originalStart, state.settings.snapMinutes)) {
        nextEnd = addMinutes(originalStart, state.settings.snapMinutes);
      }
    }

    const previewActivity = {
      id: interaction.activityId,
      type: interaction.activityType,
      start: formatDateTimeString(nextStart),
      durationMinutes: differenceInMinutes(nextEnd, nextStart),
    };

    interaction.previewActivity = previewActivity;
    interaction.valid = !hasActivityOverlap(previewActivity, interaction.activityId);
  }

  function getRenderableActivities() {
    const activities = state.activities.map((activity) => ({
      ...activity,
      preview: false,
      invalid: false,
    }));
    const preview = state.interaction ? state.interaction.previewActivity : null;
    if (!preview) {
      return activities;
    }
    return activities
      .filter((activity) => activity.id !== preview.id)
      .concat({
        ...preview,
        preview: true,
        invalid: !state.interaction.valid,
      });
  }

  function getSegmentsForDay(day, activities) {
    const dayStart = parseDateString(day.date);
    const dayEnd = addDays(dayStart, 1);

    return activities
      .map((activity) => {
        const start = parseDateTimeString(activity.start);
        const end = addMinutes(start, activity.durationMinutes);
        if (end <= dayStart || start >= dayEnd) {
          return null;
        }
        const segmentStart = start > dayStart ? start : dayStart;
        const segmentEnd = end < dayEnd ? end : dayEnd;
        const visibleMinutes = differenceInMinutes(segmentEnd, segmentStart);
        const left = (differenceInMinutes(segmentStart, dayStart) / MINUTES_PER_DAY) * 100;
        const width = (visibleMinutes / MINUTES_PER_DAY) * 100;
        const isSelected = activity.id === state.selectedActivityId;
        const meta = ACTIVITY_META[activity.type];
        const blockStyle = getActivityBlockStyle(visibleMinutes);

        return `
          <div
            class="activity-block ${meta.className} ${isSelected ? "is-selected" : ""} ${
          activity.preview ? "is-preview" : ""
        } ${activity.invalid ? "is-invalid" : ""}"
            data-activity-id="${escapeAttr(activity.id)}"
            data-day-id="${escapeAttr(day.id)}"
            style="left:${left}%;width:calc(${width}% - 2px);${blockStyle}"
            title="${escapeAttr(activity.type)}"
          >
            ${segmentStart.getTime() === start.getTime()
              ? '<span class="resize-handle left" data-edge="left"></span>'
              : ""}
            <span class="activity-label">${escapeHtml(activity.type)}</span>
            ${segmentEnd.getTime() === end.getTime()
              ? '<span class="resize-handle right" data-edge="right"></span>'
              : ""}
          </div>
        `;
      })
      .filter(Boolean);
  }

  function getActivityBlockStyle(visibleMinutes) {
    const minFontSize = normalizeFontSizeSetting(state.settings.activityMinFontSize);
    if (visibleMinutes <= 15) {
      return `--block-font-size:${minFontSize}rem;--block-padding:1px;--handle-width:4px;`;
    }
    if (visibleMinutes <= 30) {
      return `--block-font-size:${Math.max(minFontSize + 0.06, 0.46).toFixed(
        2
      )}rem;--block-padding:2px;--handle-width:5px;`;
    }
    if (visibleMinutes <= 60) {
      return `--block-font-size:${Math.max(minFontSize + 0.14, 0.56).toFixed(
        2
      )}rem;--block-padding:4px;--handle-width:6px;`;
    }
    return "--block-font-size:0.8rem;--block-padding:8px;--handle-width:10px;";
  }

  function openActivityDialog(activity) {
    state.dialogActivityId = activity.id;
    dom.dialogTypeLabel.textContent = activity.type;
    dom.dialogDate.value = activity.start.slice(0, 10);
    dom.dialogTime.value = activity.start.slice(11, 16);
    dom.dialogDuration.value = String(activity.durationMinutes);
    dom.dialog.showModal();
  }

  function importState(snapshot) {
    if (!snapshot || typeof snapshot !== "object") {
      throw new Error("Invalid JSON");
    }

    const settings = snapshot.settings || {};
    const nextSettings = {
      ...buildDefaultSettings(),
      ...settings,
    };
    const snapMinutes = SNAP_OPTIONS.includes(Number(nextSettings.snapMinutes))
      ? Number(nextSettings.snapMinutes)
      : 15;
    const days = Array.isArray(snapshot.days) ? snapshot.days : [];
    const activities = Array.isArray(snapshot.activities) ? snapshot.activities : [];

    const seenDates = new Set();
    const nextDays = days.map((day) => {
      if (!day || !isValidDateString(day.date)) {
        throw new Error("Invalid day");
      }
      if (seenDates.has(day.date)) {
        throw new Error("Duplicate date");
      }
      seenDates.add(day.date);
      return {
        id: day.id || createId("day"),
        date: day.date,
        note: typeof day.note === "string" ? day.note : "",
      };
    });

    const nextActivities = activities.map((activity) => {
      const normalized = normalizeActivity({
        id: activity.id || createId("activity"),
        type: activity.type,
        start: activity.start,
        durationMinutes: activity.durationMinutes,
      });
      if (!normalized.ok) {
        throw new Error("Invalid activity");
      }
      return normalized.activity;
    });

    const sortedActivities = normalizeActivities(nextActivities);
    assertNoOverlap(sortedActivities);

    state.settings = {
      ...buildDefaultSettings(),
      ...state.settings,
      ...nextSettings,
      snapMinutes,
      autoSaveEnabled: normalizeAutoSaveEnabledSetting(nextSettings.autoSaveEnabled),
      autoSaveIntervalMinutes: normalizeAutoSaveIntervalSetting(
        nextSettings.autoSaveIntervalMinutes
      ),
      activityCornerRadius: normalizeCornerRadiusSetting(nextSettings.activityCornerRadius),
      activityMinFontSize: normalizeFontSizeSetting(nextSettings.activityMinFontSize),
      saturdayDateColor: normalizeColorSetting(nextSettings.saturdayDateColor, "#e8f0ff"),
      sundayDateColor: normalizeColorSetting(nextSettings.sundayDateColor, "#fde4e4"),
      holidayDateColor: normalizeColorSetting(nextSettings.holidayDateColor, "#ffe6c8"),
    };
    state.holidays = normalizeHolidayMap(snapshot.holidays || {});
    state.holidaySyncInfo = normalizeHolidaySyncInfo(snapshot.holidaySyncInfo || null);
    state.days = nextDays.sort((left, right) => left.date.localeCompare(right.date));
    state.activities = sortedActivities;
    sortedActivities.forEach((activity) => ensureDaysForActivity(activity));
    state.selectedDayId = state.days[0] ? state.days[0].id : null;
    state.selectedActivityId = null;
    state.selectedPaletteType = ACTIVITY_TYPES[0].type;
    state.holidaySyncInFlight = false;
  }

  function exportState() {
    return {
      version: 1,
      exportedAt: formatDateTimeString(new Date()),
      settings: {
        snapMinutes: state.settings.snapMinutes,
        autoSaveEnabled: state.settings.autoSaveEnabled,
        autoSaveIntervalMinutes: state.settings.autoSaveIntervalMinutes,
        rowsPerPage: state.settings.rowsPerPage,
        printOrientation: state.settings.printOrientation,
        activityCornerRadius: state.settings.activityCornerRadius,
        activityMinFontSize: state.settings.activityMinFontSize,
        saturdayDateColor: state.settings.saturdayDateColor,
        sundayDateColor: state.settings.sundayDateColor,
        holidayDateColor: state.settings.holidayDateColor,
      },
      holidays: state.holidays,
      holidaySyncInfo: state.holidaySyncInfo,
      days: getSortedDays().map((day) => ({
        id: day.id,
        date: day.date,
        note: day.note,
      })),
      activities: normalizeActivities(state.activities).map((activity) => ({
        id: activity.id,
        type: activity.type,
        start: activity.start,
        durationMinutes: activity.durationMinutes,
      })),
    };
  }

  function loadStateFromStorage() {
    try {
      const raw = window.localStorage.getItem(STORAGE_KEY);
      if (!raw) {
        return null;
      }
      const parsed = JSON.parse(raw);
      importState(parsed);
      state.historyPast = [];
      state.historyFuture = [];
      return parsed;
    } catch (error) {
      console.error(error);
      window.localStorage.removeItem(STORAGE_KEY);
      return null;
    }
  }

  function persistState(options = {}) {
    const markAutoSaveDirty = options.markAutoSaveDirty !== false;
    try {
      window.localStorage.setItem(STORAGE_KEY, JSON.stringify(exportState()));
      if (markAutoSaveDirty) {
        state.autoSaveDirty = true;
      }
    } catch (error) {
      console.error(error);
    }
  }

  function saveAndRender(options = {}) {
    persistState(options);
    render();
  }

  async function initializeAutoSave(storageSnapshot) {
    state.autoSaveStatusMessage = "";
    if (!isAutoSaveSupported()) {
      renderAutoSaveStatus();
      return;
    }

    try {
      const handle = await loadAutoSaveHandleFromDb();
      if (!handle) {
        renderAutoSaveStatus();
        return;
      }

      state.autoSaveHandle = handle;
      state.autoSaveHandleName = typeof handle.name === "string" ? handle.name : AUTO_SAVE_FILE_NAME;

      const readPermission = await ensureAutoSavePermission(handle, {
        mode: "read",
        request: false,
      });
      let autoSaveSnapshot = null;
      if (readPermission === "granted") {
        autoSaveSnapshot = await readSnapshotFromAutoSaveHandle(handle);
      }

      const storageKey = getSnapshotSortKey(storageSnapshot);
      const autoSaveKey = getSnapshotSortKey(autoSaveSnapshot);

      if (autoSaveKey && (!storageKey || autoSaveKey > storageKey)) {
        importState(autoSaveSnapshot);
        state.historyPast = [];
        state.historyFuture = [];
        syncEntryDate();
        persistState({ markAutoSaveDirty: false });
        state.autoSaveDirty = false;
        state.autoSaveLastSavedAt = autoSaveKey;
        state.autoSaveStatusMessage = `${state.autoSaveHandleName} から新しいバックアップを復元しました。`;
        render();
      } else {
        state.autoSaveLastSavedAt = autoSaveKey || "";
        state.autoSaveDirty = Boolean(storageKey && (!autoSaveKey || storageKey > autoSaveKey));
      }

      const writePermission = await ensureAutoSavePermission(handle, {
        mode: "readwrite",
        request: false,
      });
      if (writePermission !== "granted" && state.autoSaveHandle) {
        state.autoSaveStatusMessage = `${state.autoSaveHandleName} へ保存するには再承認が必要です。`;
      } else if (
        state.autoSaveStatusMessage !==
        `${state.autoSaveHandleName} から新しいバックアップを復元しました。`
      ) {
        state.autoSaveStatusMessage = "";
      }
    } catch (error) {
      console.error(error);
      state.autoSaveStatusMessage = "自動保存ファイルの読み込みに失敗しました。";
    }

    syncAutoSaveSchedule();
    renderAutoSaveStatus();
  }

  async function saveSnapshotToAutoSaveFile(options = {}) {
    const {
      reason = "timer",
      requestPermission = false,
      force = false,
    } = options;

    if (!isAutoSaveSupported()) {
      return false;
    }
    if (!state.autoSaveHandle) {
      renderAutoSaveStatus();
      return false;
    }
    if (!force && !state.settings.autoSaveEnabled) {
      return false;
    }
    if (!force && !state.autoSaveDirty) {
      return true;
    }
    if (state.autoSaveInFlight) {
      return false;
    }

    state.autoSaveInFlight = true;
    renderAutoSaveStatus();

    try {
      const permission = await ensureAutoSavePermission(state.autoSaveHandle, {
        mode: "readwrite",
        request: requestPermission,
      });
      if (permission !== "granted") {
        state.autoSaveStatusMessage = `${state.autoSaveHandleName || AUTO_SAVE_FILE_NAME} へ保存するには再承認が必要です。`;
        syncAutoSaveSchedule();
        return false;
      }

      const snapshot = exportState();
      const writable = await state.autoSaveHandle.createWritable();
      await writable.write(JSON.stringify(snapshot, null, 2));
      await writable.close();

      state.autoSaveLastSavedAt = snapshot.exportedAt;
      state.autoSaveDirty = false;
      state.autoSaveStatusMessage =
        reason === "manual"
          ? `${state.autoSaveHandleName || AUTO_SAVE_FILE_NAME} に保存しました。`
          : "";
      syncAutoSaveSchedule();
      return true;
    } catch (error) {
      console.error(error);
      state.autoSaveStatusMessage = "自動保存ファイルへの書き込みに失敗しました。";
      return false;
    } finally {
      state.autoSaveInFlight = false;
      renderAutoSaveStatus();
    }
  }

  function syncAutoSaveSchedule() {
    if (state.autoSaveTimerId) {
      window.clearInterval(state.autoSaveTimerId);
      state.autoSaveTimerId = null;
    }

    if (
      !state.settings.autoSaveEnabled ||
      !state.autoSaveHandle ||
      state.autoSavePermission !== "granted"
    ) {
      return;
    }

    const intervalMinutes = normalizeAutoSaveIntervalSetting(state.settings.autoSaveIntervalMinutes);
    state.autoSaveTimerId = window.setInterval(() => {
      void saveSnapshotToAutoSaveFile({
        reason: "timer",
        requestPermission: false,
        force: false,
      });
    }, intervalMinutes * 60 * 1000);
  }

  function buildAutoSaveStatusMessage() {
    if (!isAutoSaveSupported()) {
      return "このブラウザでは単一ファイルの自動保存に対応していません。";
    }
    if (!state.autoSaveHandle) {
      return "自動保存先が未設定です。設定からファイルを選択してください。";
    }
    if (state.autoSavePermission !== "granted") {
      return `${state.autoSaveHandleName || AUTO_SAVE_FILE_NAME} へ保存するには再承認が必要です。`;
    }

    const interval = normalizeAutoSaveIntervalSetting(state.settings.autoSaveIntervalMinutes);
    const modeLabel = state.settings.autoSaveEnabled
      ? `${interval}分ごとに自動保存します。`
      : "自動保存はオフです。";

    if (!state.autoSaveLastSavedAt) {
      return `${state.autoSaveHandleName || AUTO_SAVE_FILE_NAME} / ${modeLabel}`;
    }

    return `${
      state.autoSaveHandleName || AUTO_SAVE_FILE_NAME
    } / 最終保存 ${formatAutoSaveLabel(state.autoSaveLastSavedAt)} / ${modeLabel}`;
  }

  function isAutoSaveSupported() {
    return (
      typeof window.showSaveFilePicker === "function" &&
      typeof window.indexedDB !== "undefined"
    );
  }

  async function ensureAutoSavePermission(handle, options = {}) {
    const { mode = "readwrite", request = false } = options;
    if (!handle) {
      state.autoSavePermission = "denied";
      return "denied";
    }
    if (typeof handle.queryPermission !== "function") {
      state.autoSavePermission = "granted";
      return "granted";
    }

    const descriptor = { mode };
    let permission = await handle.queryPermission(descriptor);
    if (permission !== "granted" && request && typeof handle.requestPermission === "function") {
      permission = await handle.requestPermission(descriptor);
    }
    state.autoSavePermission = permission;
    return permission;
  }

  async function readSnapshotFromAutoSaveHandle(handle) {
    const file = await handle.getFile();
    if (!file || !file.size) {
      return null;
    }
    const text = await file.text();
    if (!text.trim()) {
      return null;
    }
    return JSON.parse(text);
  }

  async function loadAutoSaveHandleFromDb() {
    const db = await openAutoSaveDb();
    try {
      return await new Promise((resolve, reject) => {
        const transaction = db.transaction(AUTO_SAVE_DB_STORE, "readonly");
        const store = transaction.objectStore(AUTO_SAVE_DB_STORE);
        const request = store.get(AUTO_SAVE_HANDLE_KEY);
        request.onsuccess = () => resolve(request.result || null);
        request.onerror = () => reject(request.error || new Error("Failed to load handle"));
      });
    } finally {
      db.close();
    }
  }

  async function saveAutoSaveHandleToDb(handle) {
    const db = await openAutoSaveDb();
    try {
      await new Promise((resolve, reject) => {
        const transaction = db.transaction(AUTO_SAVE_DB_STORE, "readwrite");
        const store = transaction.objectStore(AUTO_SAVE_DB_STORE);
        const request = store.put(handle, AUTO_SAVE_HANDLE_KEY);
        request.onsuccess = () => resolve();
        request.onerror = () => reject(request.error || new Error("Failed to save handle"));
      });
    } finally {
      db.close();
    }
  }

  function openAutoSaveDb() {
    return new Promise((resolve, reject) => {
      const request = window.indexedDB.open(AUTO_SAVE_DB_NAME, AUTO_SAVE_DB_VERSION);
      request.onupgradeneeded = () => {
        const db = request.result;
        if (!db.objectStoreNames.contains(AUTO_SAVE_DB_STORE)) {
          db.createObjectStore(AUTO_SAVE_DB_STORE);
        }
      };
      request.onsuccess = () => resolve(request.result);
      request.onerror = () => reject(request.error || new Error("Failed to open IndexedDB"));
    });
  }

  function getSnapshotSortKey(snapshot) {
    if (!snapshot || typeof snapshot.exportedAt !== "string") {
      return "";
    }
    return snapshot.exportedAt;
  }

  function formatAutoSaveLabel(value) {
    return String(value || "").replace("T", " ");
  }

  function createHistorySnapshot() {
    return JSON.parse(
      JSON.stringify({
        settings: state.settings,
        holidays: state.holidays,
        holidaySyncInfo: state.holidaySyncInfo,
        days: state.days,
        activities: state.activities,
        selectedDayId: state.selectedDayId,
        selectedActivityId: state.selectedActivityId,
        selectedPaletteType: state.selectedPaletteType,
      })
    );
  }

  function applyHistorySnapshot(snapshot) {
    state.settings = {
      ...buildDefaultSettings(),
      ...snapshot.settings,
      autoSaveEnabled: normalizeAutoSaveEnabledSetting(snapshot.settings.autoSaveEnabled),
      autoSaveIntervalMinutes: normalizeAutoSaveIntervalSetting(
        snapshot.settings.autoSaveIntervalMinutes
      ),
      activityCornerRadius: normalizeCornerRadiusSetting(snapshot.settings.activityCornerRadius),
      activityMinFontSize: normalizeFontSizeSetting(snapshot.settings.activityMinFontSize),
      saturdayDateColor: normalizeColorSetting(snapshot.settings.saturdayDateColor, "#e8f0ff"),
      sundayDateColor: normalizeColorSetting(snapshot.settings.sundayDateColor, "#fde4e4"),
      holidayDateColor: normalizeColorSetting(snapshot.settings.holidayDateColor, "#ffe6c8"),
    };
    state.holidays = normalizeHolidayMap(snapshot.holidays || {});
    state.holidaySyncInfo = normalizeHolidaySyncInfo(snapshot.holidaySyncInfo || null);
    state.days = (snapshot.days || []).map((day) => ({ ...day }));
    state.activities = (snapshot.activities || []).map((activity) => ({ ...activity }));
    state.selectedDayId = snapshot.selectedDayId || null;
    state.selectedActivityId = snapshot.selectedActivityId || null;
    state.selectedPaletteType = snapshot.selectedPaletteType || ACTIVITY_TYPES[0].type;
    state.interaction = null;
    state.dialogActivityId = null;
    state.holidaySyncInFlight = false;
  }

  function pushHistorySnapshot(snapshot = createHistorySnapshot()) {
    const previous = state.historyPast[state.historyPast.length - 1];
    if (previous && areSnapshotsEqual(previous, snapshot)) {
      return false;
    }
    state.historyPast.push(snapshot);
    if (state.historyPast.length > state.historyLimit) {
      state.historyPast.shift();
    }
    state.historyFuture = [];
    return true;
  }

  function discardLastHistorySnapshot() {
    state.historyPast.pop();
  }

  function performUndo() {
    if (!state.historyPast.length) {
      return;
    }
    const current = createHistorySnapshot();
    const previous = state.historyPast.pop();
    state.historyFuture.push(current);
    applyHistorySnapshot(previous);
    persistState();
    syncAutoSaveSchedule();
    render();
  }

  function performRedo() {
    if (!state.historyFuture.length) {
      return;
    }
    const current = createHistorySnapshot();
    const next = state.historyFuture.pop();
    state.historyPast.push(current);
    if (state.historyPast.length > state.historyLimit) {
      state.historyPast.shift();
    }
    applyHistorySnapshot(next);
    persistState();
    syncAutoSaveSchedule();
    render();
  }

  function areSnapshotsEqual(left, right) {
    return JSON.stringify(left) === JSON.stringify(right);
  }

  function commitPendingInputBeforeHistory() {
    const activeElement = document.activeElement;
    if (!activeElement) {
      return;
    }
    if (activeElement.matches && activeElement.matches(".note-input")) {
      commitNoteInput(activeElement);
    }
  }

  function commitNoteInput(input) {
    const day = getDayById(input.dataset.dayId);
    if (!day) {
      return;
    }
    const originalValue = input.dataset.originalValue ?? day.note ?? "";
    if (originalValue !== day.note) {
      const snapshot = createHistorySnapshot();
      const snapshotDay = snapshot.days.find((item) => item.id === day.id);
      if (snapshotDay) {
        snapshotDay.note = originalValue;
      }
      pushHistorySnapshot(snapshot);
      persistState();
      renderSettingsHistoryStatus();
    } else {
      persistState();
    }
    input.dataset.originalValue = day.note;
  }

  function syncEntryDate() {
    dom.entryDate.value = getSuggestedEntryDate();
  }

  function getSuggestedEntryDate() {
    const selectedDay = getSelectedDay();
    if (!selectedDay) {
      return formatDateString(new Date());
    }
    return formatDateString(addDays(parseDateString(selectedDay.date), 1));
  }

  function syncPrintNote(dayId, value) {
    const node = document.querySelector(`[data-note-print-for="${escapeSelector(dayId)}"]`);
    if (node) {
      node.textContent = value;
    }
  }

  function getSelectedDay() {
    return getDayById(state.selectedDayId);
  }

  function getSelectedActivity() {
    return getActivityById(state.selectedActivityId);
  }

  function getDayById(dayId) {
    return state.days.find((day) => day.id === dayId) || null;
  }

  function getActivityById(activityId) {
    return state.activities.find((activity) => activity.id === activityId) || null;
  }

  function findDayByDate(dateValue) {
    return state.days.find((day) => day.date === dateValue) || null;
  }

  function getSortedDays() {
    return [...state.days].sort((left, right) => left.date.localeCompare(right.date));
  }

  function sortStateDays() {
    state.days.sort((left, right) => left.date.localeCompare(right.date));
  }

  function buildYearLabel(days) {
    const years = [...new Set(days.map((day) => day.date.slice(0, 4)))];
    return years.join(", ");
  }

  function getDateButtonState(dateValue) {
    const holiday = state.holidays[dateValue] || null;
    if (holiday) {
      return {
        kind: "holiday",
        className: "is-holiday",
        title: holiday.name ? `${formatMonthDay(dateValue)} ${holiday.name}` : formatMonthDay(dateValue),
      };
    }
    const weekday = parseDateString(dateValue).getDay();
    if (weekday === 0) {
      return { kind: "sunday", className: "is-sunday", title: formatMonthDay(dateValue) };
    }
    if (weekday === 6) {
      return { kind: "saturday", className: "is-saturday", title: formatMonthDay(dateValue) };
    }
    return { kind: "normal", className: "", title: formatMonthDay(dateValue) };
  }

  function getHolidaySyncYears() {
    const years = [...new Set(state.days.map((day) => Number(day.date.slice(0, 4))))].filter((year) =>
      Number.isFinite(year)
    );
    if (years.length) {
      return years.sort((left, right) => left - right);
    }
    return [new Date().getFullYear()];
  }

  async function fetchHolidayData(years) {
    const endpoint = getHolidayApiEndpoint(years);
    const response = await fetch(endpoint, {
      headers: {
        Accept: "application/json",
      },
    });
    if (!response.ok) {
      throw new Error(`Holiday sync failed: ${response.status}`);
    }
    return response.json();
  }

  function getHolidayApiEndpoint(years) {
    const query = encodeURIComponent(years.join(","));
    if (window.location.protocol === "http:" || window.location.protocol === "https:") {
      return new URL(`/api/holidays?years=${query}`, window.location.origin).toString();
    }
    return `http://127.0.0.1:${HOLIDAY_API_PORT}/api/holidays?years=${query}`;
  }

  function mergeHolidayEntries(currentMap, holidays, years, payload) {
    const targetYears = new Set(years.map((year) => String(year)));
    const nextMap = {};
    Object.entries(currentMap || {}).forEach(([date, info]) => {
      if (!targetYears.has(String(date).slice(0, 4))) {
        nextMap[date] = { ...info };
      }
    });
    (holidays || []).forEach((holiday) => {
      nextMap[holiday.date] = {
        name: holiday.name,
        source: payload.source || "cao",
        syncedAt: payload.fetchedAt || formatDateTimeString(new Date()),
      };
    });
    return nextMap;
  }

  function areHolidayMapsEqual(left, right) {
    return JSON.stringify(left || {}) === JSON.stringify(right || {});
  }

  function normalizeHolidayMap(input) {
    const next = {};
    Object.entries(input || {}).forEach(([date, value]) => {
      if (!isValidDateString(date) || !value || typeof value !== "object") {
        return;
      }
      next[date] = {
        name: typeof value.name === "string" ? value.name : "",
        source: typeof value.source === "string" ? value.source : "cao",
        syncedAt: typeof value.syncedAt === "string" ? value.syncedAt : "",
      };
    });
    return next;
  }

  function normalizeHolidaySyncInfo(value) {
    if (!value || typeof value !== "object") {
      return null;
    }
    return {
      status: typeof value.status === "string" ? value.status : "idle",
      source: typeof value.source === "string" ? value.source : "cao",
      fetchedAt: typeof value.fetchedAt === "string" ? value.fetchedAt : "",
      syncedYears: Array.isArray(value.syncedYears) ? value.syncedYears.map(Number).filter(Number.isFinite) : [],
      count: Number.isFinite(Number(value.count)) ? Number(value.count) : 0,
      message: typeof value.message === "string" ? value.message : "",
    };
  }

  function normalizeAutoSaveEnabledSetting(value) {
    if (value === undefined || value === null) {
      return true;
    }
    return value !== false;
  }

  function normalizeAutoSaveIntervalSetting(value) {
    const nextValue = Number(value);
    return AUTO_SAVE_INTERVAL_OPTIONS.includes(nextValue)
      ? nextValue
      : AUTO_SAVE_DEFAULT_INTERVAL;
  }

  function summarizeActivitiesForDay(day) {
    const dayStart = parseDateString(day.date);
    const dayEnd = addDays(dayStart, 1);
    const totals = new Map(ACTIVITY_TYPES.map((item) => [item.type, 0]));

    state.activities.forEach((activity) => {
      const start = parseDateTimeString(activity.start);
      const end = addMinutes(start, activity.durationMinutes);
      const overlapStart = start > dayStart ? start : dayStart;
      const overlapEnd = end < dayEnd ? end : dayEnd;
      if (overlapEnd <= overlapStart) {
        return;
      }
      totals.set(
        activity.type,
        (totals.get(activity.type) || 0) + differenceInMinutes(overlapEnd, overlapStart)
      );
    });

    return ACTIVITY_TYPES.map((item) => ({
      type: item.type,
      minutes: totals.get(item.type) || 0,
    })).filter((item) => item.minutes > 0);
  }

  function formatDurationMinutes(totalMinutes) {
    const minutes = Math.max(0, Math.round(totalMinutes));
    const hoursPart = Math.floor(minutes / 60);
    const minutesPart = minutes % 60;
    if (hoursPart && minutesPart) {
      return `${hoursPart}時間${minutesPart}分`;
    }
    if (hoursPart) {
      return `${hoursPart}時間`;
    }
    return `${minutesPart}分`;
  }

  function buildTimestampSlug(date) {
    const parts = [
      date.getFullYear(),
      pad(date.getMonth() + 1),
      pad(date.getDate()),
      "-",
      pad(date.getHours()),
      pad(date.getMinutes()),
      pad(date.getSeconds()),
    ];
    return parts.join("");
  }

  function getTimelineFromPoint(clientX, clientY) {
    const element = document.elementFromPoint(clientX, clientY);
    return element ? element.closest(".timeline[data-day-id]") : null;
  }

  function getSnappedMinuteFromClientX(rect, clientX, options) {
    const allowOverflow = options.allowOverflow;
    const clampInside = options.clampInside;
    let minutes = ((clientX - rect.left) / rect.width) * MINUTES_PER_DAY;
    if (!allowOverflow) {
      minutes = clamp(minutes, 0, MINUTES_PER_DAY);
    }
    let snapped = snapToGrid(minutes, state.settings.snapMinutes);
    if (clampInside) {
      snapped = clamp(snapped, 0, MINUTES_PER_DAY - state.settings.snapMinutes);
    }
    return snapped;
  }

  function snapToGrid(value, step) {
    return Math.round(value / step) * step;
  }

  function normalizeCornerRadiusSetting(value) {
    const numeric = Number(value);
    if (!Number.isFinite(numeric)) {
      return buildDefaultSettings().activityCornerRadius;
    }
    return clamp(Math.round(numeric), 0, 16);
  }

  function normalizeColorSetting(value, fallback) {
    const candidate = String(value || "").trim();
    if (/^#[0-9a-fA-F]{6}$/.test(candidate)) {
      return candidate.toLowerCase();
    }
    return fallback;
  }

  function normalizeFontSizeSetting(value) {
    const numeric = Number(value);
    if (!Number.isFinite(numeric)) {
      return buildDefaultSettings().activityMinFontSize;
    }
    return clamp(Math.round(numeric * 100) / 100, 0.3, 0.7);
  }

  function snapDuration(value) {
    const numeric = Number(value);
    const snapped = snapToGrid(
      Number.isFinite(numeric) ? numeric : state.settings.snapMinutes,
      state.settings.snapMinutes
    );
    return Math.max(state.settings.snapMinutes, snapped);
  }

  function snapDateTimeString(value) {
    const date = parseDateTimeString(value);
    const startOfCurrentDay = startOfDay(date);
    const minutes = differenceInMinutes(date, startOfCurrentDay);
    const snapped = snapToGrid(minutes, state.settings.snapMinutes);
    return formatDateTimeString(addMinutes(startOfCurrentDay, snapped));
  }

  function parseDateString(value) {
    const [year, month, day] = value.split("-").map(Number);
    return new Date(year, month - 1, day, 0, 0, 0, 0);
  }

  function parseDateTimeString(value) {
    const [datePart, timePart] = value.split("T");
    const [year, month, day] = datePart.split("-").map(Number);
    const [hour, minute] = timePart.split(":").map(Number);
    return new Date(year, month - 1, day, hour, minute, 0, 0);
  }

  function formatDateString(date) {
    return `${date.getFullYear()}-${pad(date.getMonth() + 1)}-${pad(date.getDate())}`;
  }

  function formatDateTimeString(date) {
    return `${formatDateString(date)}T${pad(date.getHours())}:${pad(date.getMinutes())}`;
  }

  function formatMonthDay(dateValue) {
    const [, month, day] = dateValue.split("-");
    const weekday = WEEKDAY_LABELS[parseDateString(dateValue).getDay()];
    return `${Number(month)}/${Number(day)}(${weekday})`;
  }

  function addDays(date, value) {
    const next = new Date(date.getTime());
    next.setDate(next.getDate() + value);
    return next;
  }

  function addMinutes(date, value) {
    return new Date(date.getTime() + value * 60 * 1000);
  }

  function startOfDay(date) {
    return new Date(date.getFullYear(), date.getMonth(), date.getDate(), 0, 0, 0, 0);
  }

  function differenceInMinutes(left, right) {
    return Math.round((left.getTime() - right.getTime()) / (60 * 1000));
  }

  function assertNoOverlap(activities) {
    for (let index = 0; index < activities.length; index += 1) {
      const current = activities[index];
      const currentStart = parseDateTimeString(current.start);
      const currentEnd = addMinutes(currentStart, current.durationMinutes);
      for (let inner = index + 1; inner < activities.length; inner += 1) {
        const next = activities[inner];
        const nextStart = parseDateTimeString(next.start);
        const nextEnd = addMinutes(nextStart, next.durationMinutes);
        if (currentStart < nextEnd && nextStart < currentEnd) {
          throw new Error("Activity overlap");
        }
      }
    }
  }

  function isValidDateString(value) {
    return /^\d{4}-\d{2}-\d{2}$/.test(value);
  }

  function isValidDateTimeString(value) {
    return /^\d{4}-\d{2}-\d{2}T\d{2}:\d{2}$/.test(value);
  }

  function chunk(items, size) {
    const result = [];
    for (let index = 0; index < items.length; index += size) {
      result.push(items.slice(index, index + size));
    }
    return result;
  }

  function createId(prefix) {
    return `${prefix}-${Math.random().toString(36).slice(2, 10)}-${Date.now().toString(36)}`;
  }

  function escapeHtml(value) {
    return String(value)
      .replaceAll("&", "&amp;")
      .replaceAll("<", "&lt;")
      .replaceAll(">", "&gt;")
      .replaceAll('"', "&quot;")
      .replaceAll("'", "&#39;");
  }

  function escapeAttr(value) {
    return escapeHtml(value);
  }

  function escapeSelector(value) {
    return String(value).replaceAll('"', '\\"');
  }

  function clamp(value, min, max) {
    return Math.min(Math.max(value, min), max);
  }

  function pad(value) {
    return String(value).padStart(2, "0");
  }
})();

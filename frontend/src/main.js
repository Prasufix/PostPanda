import "./styles.css";

const providerPresets = {
  outlook: { host: "smtp-mail.outlook.com", port: "587", readonly: true },
  gmail: { host: "smtp.gmail.com", port: "587", readonly: true },
  custom: { host: "", port: "587", readonly: false },
};

const backendOrigin = `${window.location.protocol}//${window.location.hostname}:8000`;

const state = {
  sessionId: "",
  totalRows: 0,
  index: 0,
  columns: [],
  busy: false,
  oauthClientId: "",
  oauthStatus: {},
};

const el = {
  excelFile: document.querySelector("#excelFile"),
  excelPickButton: document.querySelector("#excelPickButton"),
  excelSelectedName: document.querySelector("#excelSelectedName"),
  fileMeta: document.querySelector("#fileMeta"),
  emailCol: document.querySelector("#emailCol"),
  addVariable: document.querySelector("#addVariable"),
  variableRows: document.querySelector("#variableRows"),
  subject: document.querySelector("#subject"),
  template: document.querySelector("#template"),
  fmtBold: document.querySelector("#fmtBold"),
  fmtLink: document.querySelector("#fmtLink"),
  fmtVariable: document.querySelector("#fmtVariable"),
  fmtParagraph: document.querySelector("#fmtParagraph"),
  fmtBullet: document.querySelector("#fmtBullet"),
  variablePicker: document.querySelector("#variablePicker"),
  variablePickerInput: document.querySelector("#variablePickerInput"),
  variablePickerSuggestions: document.querySelector("#variablePickerSuggestions"),
  variablePickerInsert: document.querySelector("#variablePickerInsert"),
  variablePickerClose: document.querySelector("#variablePickerClose"),
  previewMeta: document.querySelector("#previewMeta"),
  previewHtml: document.querySelector("#previewHtml"),
  prevBtn: document.querySelector("#prevBtn"),
  nextBtn: document.querySelector("#nextBtn"),
  refreshPreview: document.querySelector("#refreshPreview"),
  authMode: document.querySelector("#authMode"),
  provider: document.querySelector("#provider"),
  smtpFields: document.querySelector("#smtpFields"),
  oauthFields: document.querySelector("#oauthFields"),
  mailAppFields: document.querySelector("#mailAppFields"),
  oauthConnect: document.querySelector("#oauthConnect"),
  oauthDisconnect: document.querySelector("#oauthDisconnect"),
  oauthInfo: document.querySelector("#oauthInfo"),
  smtpHost: document.querySelector("#smtpHost"),
  smtpPort: document.querySelector("#smtpPort"),
  smtpSender: document.querySelector("#smtpSender"),
  smtpPassword: document.querySelector("#smtpPassword"),
  sendTest: document.querySelector("#sendTest"),
  sendAll: document.querySelector("#sendAll"),
  status: document.querySelector("#status"),
};

init();

async function init() {
  state.oauthClientId = ensureClientId();
  bindEvents();
  updateSelectedFileName("");
  applyProviderPreset();
  hydrateDefaultColumns([]);
  updateAuthModeView();
  await refreshOAuthStatus();
}

function bindEvents() {
  el.excelPickButton.addEventListener("click", onPickExcelFile);
  el.excelFile.addEventListener("change", onUploadFile);

  el.authMode.addEventListener("change", () => {
    updateAuthModeView();
    setBusy(state.busy);
  });

  el.provider.addEventListener("change", () => {
    applyProviderPreset();
    updateAuthModeView();
    refreshOAuthInfo();
  });

  el.emailCol.addEventListener("change", () => refreshPreviewSafe(false));
  el.addVariable.addEventListener("click", onAddVariableRow);
  el.template.addEventListener("input", () => refreshPreviewSafe(false));

  el.fmtBold.addEventListener("click", onInsertBold);
  el.fmtLink.addEventListener("click", onInsertLink);
  el.fmtVariable.addEventListener("click", onOpenVariablePicker);
  el.fmtParagraph.addEventListener("click", onInsertParagraph);
  el.fmtBullet.addEventListener("click", onInsertBullet);
  el.variablePickerClose.addEventListener("click", closeVariablePicker);
  el.variablePickerInsert.addEventListener("click", onInsertVariableFromPicker);
  el.variablePickerInput.addEventListener("keydown", (event) => {
    if (event.key === "Enter") {
      event.preventDefault();
      onInsertVariableFromPicker();
      return;
    }
    if (event.key === "Escape") {
      event.preventDefault();
      closeVariablePicker();
    }
  });
  el.variablePickerInput.addEventListener("input", () => {
    renderVariableSuggestions(el.variablePickerInput.value);
  });

  document.addEventListener("mousedown", (event) => {
    if (el.variablePicker.classList.contains("hidden")) return;
    const target = event.target;
    const clickedPicker = el.variablePicker.contains(target);
    const clickedTrigger = el.fmtVariable.contains(target);
    if (!clickedPicker && !clickedTrigger) {
      closeVariablePicker();
    }
  });

  el.prevBtn.addEventListener("click", async () => {
    if (!hasData()) return;
    state.index = (state.index - 1 + state.totalRows) % state.totalRows;
    await refreshPreviewSafe(true);
  });

  el.nextBtn.addEventListener("click", async () => {
    if (!hasData()) return;
    state.index = (state.index + 1) % state.totalRows;
    await refreshPreviewSafe(true);
  });

  el.refreshPreview.addEventListener("click", async () => {
    await refreshPreviewSafe(true);
  });

  el.oauthConnect.addEventListener("click", onOAuthConnect);
  el.oauthDisconnect.addEventListener("click", onOAuthDisconnect);

  window.addEventListener("message", async (event) => {
    if (!event.data || event.data.type !== "postpanda-oauth") return;

    const isTrustedOrigin = event.origin === backendOrigin || event.origin.endsWith(":8000");
    if (!isTrustedOrigin) return;

    if (event.data.status === "ok") {
      setStatus(`OAuth connected (${event.data.provider}): ${event.data.email || "ok"}`, "ok");
    } else {
      setStatus(event.data.message || "OAuth failed.", "error");
    }

    await refreshOAuthStatus();
  });

  el.sendTest.addEventListener("click", onSendTest);
  el.sendAll.addEventListener("click", onSendAll);
}

function ensureClientId() {
  const key = "postpanda_client_id";
  const existing = localStorage.getItem(key);
  if (existing) return existing;

  let generated;
  if (window.crypto && typeof window.crypto.randomUUID === "function") {
    generated = window.crypto.randomUUID().replace(/-/g, "");
  } else {
    generated = `client_${Math.random().toString(36).slice(2)}${Date.now().toString(36)}`;
  }

  localStorage.setItem(key, generated);
  return generated;
}

function isOAuthMode() {
  return el.authMode.value === "oauth";
}

function isMailAppMode() {
  return el.authMode.value === "mailapp";
}

function isPasswordMode() {
  return el.authMode.value === "password";
}

function selectedOAuthProvider() {
  if (el.provider.value === "outlook") return "microsoft";
  if (el.provider.value === "gmail") return "google";
  return "";
}

function applyProviderPreset() {
  const preset = providerPresets[el.provider.value] || providerPresets.custom;
  el.smtpHost.value = preset.host;
  el.smtpPort.value = preset.port;
  el.smtpHost.readOnly = preset.readonly;
  el.smtpPort.readOnly = preset.readonly;
}

function updateAuthModeView() {
  if (isOAuthMode() && el.provider.value === "custom") {
    el.provider.value = "outlook";
    applyProviderPreset();
  }

  el.smtpFields.classList.toggle("hidden", !isPasswordMode());
  el.oauthFields.classList.toggle("hidden", !isOAuthMode());
  el.mailAppFields.classList.toggle("hidden", !isMailAppMode());

  if (isMailAppMode()) {
    el.sendTest.classList.add("hidden");
    el.sendAll.textContent = "Open All in Mail App";
  } else {
    el.sendTest.classList.remove("hidden");
    el.sendTest.textContent = "Send Test Email";
    el.sendAll.textContent = "Send All Emails";
  }

  refreshOAuthInfo();
}

async function refreshOAuthStatus() {
  try {
    const data = await callApi(`/api/oauth/status?clientId=${encodeURIComponent(state.oauthClientId)}`, {
      method: "GET",
    });
    state.oauthStatus = data.providers || {};
  } catch (_error) {
    state.oauthStatus = {};
  }
  refreshOAuthInfo();
}

function refreshOAuthInfo() {
  if (!isOAuthMode()) {
    el.oauthConnect.disabled = true;
    el.oauthDisconnect.disabled = true;
    return;
  }

  const provider = selectedOAuthProvider();

  if (!provider) {
    el.oauthInfo.textContent = "OAuth is only available for Outlook or Gmail.";
    el.oauthConnect.disabled = true;
    el.oauthDisconnect.disabled = true;
    return;
  }

  const providerInfo = state.oauthStatus[provider] || null;
  const configured = Boolean(providerInfo && providerInfo.configured);
  const connected = Boolean(providerInfo && providerInfo.connected);
  const label = providerInfo?.label || provider;

  if (!configured) {
    el.oauthInfo.textContent = `${label}: OAuth is not configured (missing client ID/secret).`;
  } else if (connected) {
    el.oauthInfo.textContent = `${label} connected as ${providerInfo.email}.`;
  } else {
    el.oauthInfo.textContent = `${label} not connected yet.`;
  }

  el.oauthConnect.disabled = state.busy || !configured;
  el.oauthDisconnect.disabled = state.busy || !connected;
}

async function onOAuthConnect() {
  const provider = selectedOAuthProvider();
  if (!provider) {
    setStatus("OAuth is only available for Outlook or Gmail.", "error");
    return;
  }

  const providerInfo = state.oauthStatus[provider] || null;
  if (!providerInfo || !providerInfo.configured) {
    setStatus("OAuth is not configured for this provider yet.", "error");
    return;
  }

  const url = new URL(`/api/oauth/login/${provider}`, backendOrigin);
  url.searchParams.set("clientId", state.oauthClientId);
  url.searchParams.set("frontendOrigin", window.location.origin);

  const popup = window.open(url.toString(), "postpanda_oauth", "width=520,height=720");
  if (!popup) {
    setStatus("Popup was blocked. Please allow popups for this page.", "error");
    return;
  }

  setStatus("OAuth login opened. Please complete it in the popup.", "info");
}

async function onOAuthDisconnect() {
  const provider = selectedOAuthProvider();
  if (!provider) {
    setStatus("OAuth is only available for Outlook or Gmail.", "error");
    return;
  }

  setBusy(true);
  try {
    await callApi("/api/oauth/logout", {
      method: "POST",
      body: JSON.stringify({
        clientId: state.oauthClientId,
        provider,
      }),
    });

    setStatus("OAuth disconnected.", "ok");
    await refreshOAuthStatus();
  } catch (error) {
    setStatus(error.message, "error");
  } finally {
    setBusy(false);
  }
}

async function onUploadFile(event) {
  const file = event.target.files?.[0];
  updateSelectedFileName(file?.name || "");
  if (!file) return;

  setBusy(true);
  setStatus("Uploading Excel file...", "info");

  try {
    const formData = new FormData();
    formData.append("file", file);

    const data = await callApi("/api/upload", {
      method: "POST",
      body: formData,
    });

    state.sessionId = data.sessionId;
    state.totalRows = data.totalRows;
    state.index = 0;

    updateSelectedFileName(data.filename || file.name);
    el.fileMeta.textContent = `${data.filename} - ${data.totalRows} recipients`;
    hydrateDefaultColumns(data.columns);

    setStatus("File uploaded. Generating preview...", "ok");
    await refreshPreviewSafe(true);
  } catch (error) {
    setStatus(error.message, "error");
  } finally {
    setBusy(false);
  }
}

function onPickExcelFile() {
  if (state.busy) return;
  el.excelFile.click();
}

function updateSelectedFileName(name) {
  if (!el.excelSelectedName) return;
  const value = (name || "").trim();
  el.excelSelectedName.textContent = value || "No file selected";
  el.excelSelectedName.title = value;
}

function hydrateDefaultColumns(columns) {
  state.columns = Array.isArray(columns) ? [...columns] : [];
  setSelectOptions(el.emailCol, state.columns, pickColumn(state.columns, ["mail", "email", "e-mail"], 2));
  clearVariableRows();
  el.addVariable.disabled = !state.columns.length || state.busy;
}

function setSelectOptions(select, values, defaultValue = "", allowEmpty = false, emptyLabel = "-") {
  select.innerHTML = "";
  if (!values.length) {
    const option = new Option(emptyLabel, "", true, true);
    select.add(option);
    select.disabled = true;
    return;
  }

  if (allowEmpty) {
    select.add(new Option(emptyLabel, ""));
  }

  values.forEach((value) => {
    const option = new Option(value, value);
    select.add(option);
  });

  const selected = allowEmpty
    ? (defaultValue && values.includes(defaultValue) ? defaultValue : "")
    : (values.includes(defaultValue) ? defaultValue : values[0]);
  select.value = selected;
  select.disabled = false;
}

function pickColumn(columns, candidates, fallbackIndex = 0) {
  if (!columns.length) return "";
  const normalize = (value) => value.trim().toLowerCase().replace(/[^a-z0-9]+/g, "");
  const normalized = columns.map((col) => ({ raw: col, norm: normalize(col) }));

  for (const candidate of candidates) {
    const candidateNorm = normalize(candidate);
    const exact = normalized.find((entry) => entry.norm === candidateNorm);
    if (exact) return exact.raw;
  }

  for (const candidate of candidates) {
    if (candidate === "name") continue;
    const candidateNorm = normalize(candidate);
    const fuzzy = normalized.find((entry) => entry.norm.includes(candidateNorm));
    if (fuzzy) return fuzzy.raw;
  }

  return columns[Math.min(fallbackIndex, columns.length - 1)];
}

function clearVariableRows() {
  el.variableRows.innerHTML = "";
  renderVariableSuggestions();
}

function sanitizeVariableName(value) {
  return (value || "").trim().replace(/^\{+|\}+$/g, "").trim();
}

function getAvailableVariableNames() {
  const inputs = el.variableRows.querySelectorAll(".variable-row input");
  const seen = new Set();
  const names = [];

  inputs.forEach((input) => {
    const name = sanitizeVariableName(input.value);
    if (!name) return;
    const key = name.toLowerCase();
    if (seen.has(key)) return;
    seen.add(key);
    names.push(name);
  });

  return names;
}

function renderVariableSuggestions(query = "") {
  if (!el.variablePickerSuggestions) return;
  el.variablePickerSuggestions.innerHTML = "";

  const names = getAvailableVariableNames();
  if (!names.length) {
    const empty = document.createElement("span");
    empty.className = "hint-inline";
    empty.textContent = "No mapped variables yet.";
    el.variablePickerSuggestions.append(empty);
    return;
  }

  const needle = sanitizeVariableName(query).toLowerCase();
  const filteredNames = needle
    ? names.filter((name) => name.toLowerCase().includes(needle))
    : names;

  if (!filteredNames.length) {
    const empty = document.createElement("span");
    empty.className = "hint-inline";
    empty.textContent = "No variable matches your input.";
    el.variablePickerSuggestions.append(empty);
    return;
  }

  filteredNames.forEach((name) => {
    const button = document.createElement("button");
    button.type = "button";
    button.textContent = `{{${name}}}`;
    button.addEventListener("click", () => {
      el.variablePickerInput.value = name;
      el.variablePickerInput.focus();
      renderVariableSuggestions(name);
    });
    el.variablePickerSuggestions.append(button);
  });
}

function findBestVariableMatch(input, names) {
  if (!input || !names.length) return "";
  const target = input.toLowerCase();

  const exact = names.find((name) => name.toLowerCase() === target);
  if (exact) return exact;

  const startsWith = names.filter((name) => name.toLowerCase().startsWith(target));
  if (startsWith.length === 1) return startsWith[0];

  const contains = names.filter((name) => name.toLowerCase().includes(target));
  if (contains.length === 1) return contains[0];

  return "";
}

function onAddVariableRow() {
  addVariableRow("", "");
  refreshPreviewSafe(false);
}

function addVariableRow(variableName, defaultColumn) {
  const row = document.createElement("div");
  row.className = "variable-row";

  const variableLabel = document.createElement("label");
  const variableTitle = document.createElement("span");
  variableTitle.textContent = "Variable";
  const variableInput = document.createElement("input");
  variableInput.type = "text";
  variableInput.placeholder = "e.g. Company";
  variableInput.value = variableName || "";
  variableLabel.append(variableTitle, variableInput);

  const columnLabel = document.createElement("label");
  const columnTitle = document.createElement("span");
  columnTitle.textContent = "Excel Column";
  const columnSelect = document.createElement("select");
  setSelectOptions(columnSelect, state.columns, defaultColumn || "", true, "Select column");
  columnLabel.append(columnTitle, columnSelect);

  const removeButton = document.createElement("button");
  removeButton.type = "button";
  removeButton.textContent = "Remove";
  removeButton.addEventListener("click", () => {
    row.remove();
    renderVariableSuggestions();
    refreshPreviewSafe(false);
  });

  variableInput.addEventListener("input", () => {
    renderVariableSuggestions();
    refreshPreviewSafe(false);
  });
  columnSelect.addEventListener("change", () => refreshPreviewSafe(false));

  row.append(variableLabel, columnLabel, removeButton);
  el.variableRows.append(row);
  renderVariableSuggestions();
}

function collectVariableMap() {
  const mapping = {};
  const rows = el.variableRows.querySelectorAll(".variable-row");

  rows.forEach((row) => {
    const variableInput = row.querySelector("input");
    const columnSelect = row.querySelector("select");
    const name = sanitizeVariableName(variableInput?.value || "");
    const column = columnSelect?.value.trim() || "";
    if (!name || !column) return;
    mapping[name] = column;
  });

  return mapping;
}

function hasData() {
  return Boolean(state.sessionId && state.totalRows > 0);
}

function onInsertBold() {
  wrapSelection(el.template, "**", "**", "bold");
}

function onInsertLink() {
  const current = getSelectedText(el.template) || "Link text";
  const url = window.prompt("Enter link URL (https://...)", "https://");
  if (!url) return;
  const snippet = `[${current}](${url.trim()})`;
  replaceSelection(el.template, snippet);
  el.template.dispatchEvent(new Event("input"));
}

function onOpenVariablePicker() {
  const selectedText = sanitizeVariableName(getSelectedText(el.template));
  el.variablePickerInput.value = selectedText || "";

  el.variablePicker.classList.remove("hidden");
  renderVariableSuggestions(el.variablePickerInput.value);
  el.variablePickerInput.focus();
  el.variablePickerInput.select();
}

function closeVariablePicker() {
  el.variablePicker.classList.add("hidden");
}

function onInsertVariableFromPicker() {
  const typedVariable = sanitizeVariableName(el.variablePickerInput.value);
  const availableNames = getAvailableVariableNames();
  const matchedVariable = findBestVariableMatch(typedVariable, availableNames);
  const variable = matchedVariable || typedVariable || (availableNames.length === 1 ? availableNames[0] : "");
  if (!variable) {
    setStatus("Please enter a variable name.", "error");
    return;
  }

  replaceSelection(el.template, `{{${variable}}}`);
  el.template.dispatchEvent(new Event("input"));
  closeVariablePicker();
  refreshPreviewSafe(false);
}

function onInsertParagraph() {
  replaceSelection(el.template, "\n\n");
  el.template.dispatchEvent(new Event("input"));
}

function onInsertBullet() {
  const selected = getSelectedText(el.template);
  if (!selected) {
    replaceSelection(el.template, "- ");
  } else {
    const lines = selected.split("\n").map((line) => {
      if (!line.trim()) return line;
      return line.trimStart().startsWith("- ") ? line : `- ${line}`;
    });
    replaceSelection(el.template, lines.join("\n"));
  }
  el.template.dispatchEvent(new Event("input"));
}

function getSelectedText(textarea) {
  const start = textarea.selectionStart ?? 0;
  const end = textarea.selectionEnd ?? 0;
  return textarea.value.slice(start, end);
}

function wrapSelection(textarea, prefix, suffix, fallbackText) {
  const selected = getSelectedText(textarea) || fallbackText;
  replaceSelection(textarea, `${prefix}${selected}${suffix}`);
  textarea.dispatchEvent(new Event("input"));
}

function replaceSelection(textarea, insertion) {
  const start = textarea.selectionStart ?? 0;
  const end = textarea.selectionEnd ?? 0;
  const before = textarea.value.slice(0, start);
  const after = textarea.value.slice(end);
  textarea.value = `${before}${insertion}${after}`;
  const cursor = start + insertion.length;
  textarea.focus();
  textarea.selectionStart = cursor;
  textarea.selectionEnd = cursor;
}

function previewPayload() {
  return {
    sessionId: state.sessionId,
    index: state.index,
    template: el.template.value,
    mapping: {
      emailCol: el.emailCol.value,
      variableMap: collectVariableMap(),
    },
  };
}

function sendPayload() {
  return {
    ...previewPayload(),
    authMode: el.authMode.value,
    subject: el.subject.value,
    smtp: {
      host: el.smtpHost.value.trim(),
      port: Number(el.smtpPort.value),
      sender: el.smtpSender.value.trim(),
      password: el.smtpPassword.value,
    },
    oauth: {
      provider: selectedOAuthProvider(),
      clientId: state.oauthClientId,
    },
    mailApp: {
      provider: el.provider.value,
    },
  };
}

function isOAuthConnectedForCurrentProvider() {
  const provider = selectedOAuthProvider();
  if (!provider) return false;
  const info = state.oauthStatus[provider];
  return Boolean(info && info.connected);
}

async function refreshPreviewSafe(showErrors) {
  if (!hasData()) {
    el.previewMeta.textContent = "Recipient: -";
    el.previewHtml.textContent = "No rendered preview available yet.";
    return;
  }

  try {
    const data = await callApi("/api/preview", {
      method: "POST",
      body: JSON.stringify(previewPayload()),
    });

    state.index = data.index;
    el.previewMeta.textContent = `Recipient ${data.index + 1}/${data.totalRows} - ${data.recipient || "without email"}`;
    el.previewHtml.innerHTML = data.previewHtml || "";
  } catch (error) {
    if (showErrors) {
      setStatus(error.message, "error");
    }
  }
}

async function onSendTest() {
  if (!hasData()) {
    setStatus("Please upload an Excel file first.", "error");
    return;
  }

  if (isOAuthMode() && !isOAuthConnectedForCurrentProvider()) {
    setStatus("Please connect OAuth for the selected provider first.", "error");
    return;
  }

  setBusy(true);
  setStatus(isMailAppMode() ? "Opening message in mail app..." : "Sending test email...", "info");

  try {
    const data = await callApi("/api/send-test", {
      method: "POST",
      body: JSON.stringify(sendPayload()),
    });

    setStatus(data.message || "Action completed successfully.", "ok");
  } catch (error) {
    setStatus(error.message, "error");
  } finally {
    setBusy(false);
  }
}

async function onSendAll() {
  if (!hasData()) {
    setStatus("Please upload an Excel file first.", "error");
    return;
  }

  if (isOAuthMode() && !isOAuthConnectedForCurrentProvider()) {
    setStatus("Please connect OAuth for the selected provider first.", "error");
    return;
  }

  const question = isMailAppMode()
    ? `Open ${state.totalRows} messages in the mail app?`
    : `Send ${state.totalRows} emails?`;
  const shouldProceed = window.confirm(question);
  if (!shouldProceed) return;

  setBusy(true);
  setStatus(isMailAppMode() ? "Opening messages in mail app..." : "Sending bulk emails...", "info");

  try {
    const data = await callApi("/api/send-all", {
      method: "POST",
      body: JSON.stringify(sendPayload()),
    });

    if (data.mode === "mailapp") {
      setStatus(`Opened: ${data.drafted} drafted, ${data.skipped} skipped, ${data.total} total.`, "ok");
    } else {
      setStatus(`Done: ${data.sent} sent, ${data.skipped} skipped, ${data.total} total.`, "ok");
    }
  } catch (error) {
    setStatus(error.message, "error");
  } finally {
    setBusy(false);
  }
}

async function callApi(path, init) {
  const headers = new Headers(init?.headers || {});
  if (!(init?.body instanceof FormData)) {
    headers.set("Content-Type", "application/json");
  }

  const response = await fetch(path, {
    ...init,
    headers,
  });

  const json = await response.json().catch(() => ({}));
  if (!response.ok) {
    throw new Error(json.error || "Unknown error");
  }

  return json;
}

function setBusy(value) {
  state.busy = value;
  if (value) {
    closeVariablePicker();
  }

  const controls = [
    el.excelFile,
    el.excelPickButton,
    el.emailCol,
    el.addVariable,
    el.subject,
    el.template,
    el.fmtBold,
    el.fmtLink,
    el.fmtVariable,
    el.fmtParagraph,
    el.fmtBullet,
    el.variablePickerInput,
    el.variablePickerInsert,
    el.variablePickerClose,
    el.prevBtn,
    el.nextBtn,
    el.refreshPreview,
    el.authMode,
    el.provider,
    el.oauthConnect,
    el.oauthDisconnect,
    el.smtpHost,
    el.smtpPort,
    el.smtpSender,
    el.smtpPassword,
    el.sendTest,
    el.sendAll,
  ];

  controls.forEach((control) => {
    if (!control) return;

    if (control === el.smtpHost || control === el.smtpPort) {
      const preset = providerPresets[el.provider.value] || providerPresets.custom;
      const disabled = value || !isPasswordMode() || preset.readonly;
      control.disabled = disabled;
      return;
    }

    if (control === el.smtpSender || control === el.smtpPassword) {
      control.disabled = value || !isPasswordMode();
      return;
    }

    if (control === el.oauthConnect || control === el.oauthDisconnect) {
      refreshOAuthInfo();
      return;
    }

    control.disabled = value;
  });

  const variableControls = el.variableRows.querySelectorAll("input, select, button");
  variableControls.forEach((control) => {
    control.disabled = value || !state.columns.length;
  });
  el.addVariable.disabled = value || !state.columns.length;

  const suggestionButtons = el.variablePickerSuggestions.querySelectorAll("button");
  suggestionButtons.forEach((button) => {
    button.disabled = value;
  });

  refreshOAuthInfo();
}

function setStatus(message, tone) {
  el.status.textContent = message;
  el.status.dataset.tone = tone;
}

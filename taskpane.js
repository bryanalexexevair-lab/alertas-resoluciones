/* global Office, Excel */
(function () {
  "use strict";

  const STORAGE_KEY = "resolucion-alertas-enabled-v1";
  const MONTHS = ["ENERO", "FEBRERO", "MARZO", "ABRIL", "MAYO", "JUNIO", "JULIO", "AGOSTO", "SEPTIEMBRE", "OCTUBRE", "NOVIEMBRE", "DICIEMBRE"];
  const registeredSheets = new Set();
  let enabled = false;
  let activeSheetId = null;
  let activationHandlerReady = false;

  const toggleButton = () => document.getElementById("toggleButton");
  const setSheetStatus = (message, state) => {
    const el = document.getElementById("sheetStatus");
    el.textContent = message;
    el.className = "sheet-status" + (state ? " " + state : "");
  };

  function setEnabled(value) {
    enabled = value;
    try { window.localStorage.setItem(STORAGE_KEY, value ? "1" : "0"); } catch (_) { /* Storage can be restricted. */ }
    const button = toggleButton();
    button.setAttribute("aria-pressed", value ? "true" : "false");
    document.getElementById("toggleTitle").textContent = value ? "Alertas activadas" : "Alertas desactivadas";
    document.getElementById("toggleHint").textContent = value ? "Vigilando la hoja mensual mientras este panel permanezca abierto." : "Actívalas para vigilar la hoja mensual.";
    document.getElementById("toggleLabel").textContent = value ? "Desactivar" : "Activar";
    button.classList.toggle("on", value);
    if (!value) {
      activeSheetId = null;
      setSheetStatus("Alertas pausadas.");
    }
  }

  function normalized(value) {
    return String(value == null ? "" : value).trim().toUpperCase().normalize("NFD").replace(/[\u0300-\u036f]/g, "");
  }

  function normalizedId(value) {
    return normalized(value).replace(/[^0-9A-Z]/g, "");
  }

  function isMonthlyName(name) {
    const match = normalized(name).match(/^([A-Z]+)\s+(\d{4})$/);
    return !!match && MONTHS.indexOf(match[1]) >= 0;
  }

  function isExpectedHeaders(values) {
    const nameHeader = normalized(values[0]);
    return (nameHeader === "NOMBRE" || nameHeader === "NOMBRES") && normalized(values[1]) === "CEDULA" && normalized(values[2]) === "ESTATUS" && normalized(values[3]).startsWith("MOTIVO");
  }

  async function registerActivationHandler() {
    if (activationHandlerReady) return;
    await Excel.run(async (context) => {
      context.workbook.worksheets.onActivated.add(onWorksheetActivated);
      await context.sync();
    });
    activationHandlerReady = true;
  }

  async function onWorksheetActivated() {
    if (enabled) await bindActiveMonthlySheet();
  }

  async function bindActiveMonthlySheet() {
    if (!enabled) return;
    try {
      await Excel.run(async (context) => {
        const sheet = context.workbook.worksheets.getActiveWorksheet();
        sheet.load("id,name");
        await context.sync();
        if (!isMonthlyName(sheet.name)) {
          activeSheetId = null;
          setSheetStatus("Abre la hoja del mes, por ejemplo NOVIEMBRE 2026.");
          return;
        }
        const headerRanges = ["B1", "E1", "F1", "G1"].map((address) => sheet.getRange(address));
        headerRanges.forEach((range) => range.load("text"));
        await context.sync();
        const headerValues = headerRanges.map((range) => range.text[0][0]);
        if (!isExpectedHeaders(headerValues)) {
          activeSheetId = null;
          setSheetStatus("La hoja tiene nombre mensual, pero sus encabezados no coinciden con B, E, F y G.", "error");
          return;
        }
        activeSheetId = sheet.id;
        if (!registeredSheets.has(sheet.id)) {
          sheet.onChanged.add(onWorksheetChanged);
          await context.sync();
          registeredSheets.add(sheet.id);
        }
        setSheetStatus("Lista en " + sheet.name + ". Cambia una resolución en F para revisar antecedentes.", "ready");
      });
    } catch (error) {
      setSheetStatus("No pude conectarme a la hoja activa: " + safeError(error), "error");
    }
  }

  function parseStatusRows(address) {
    const cellAddress = String(address || "").split("!").pop().replace(/'/g, "").replace(/\$/g, "");
    const match = cellAddress.match(/^F(\d+)(?::F(\d+))?$/i);
    if (!match) return [];
    const first = Number(match[1]);
    const last = Number(match[2] || match[1]);
    if (first < 2 || last < first || last - first > 500) return [];
    const rows = [];
    for (let row = first; row <= last; row += 1) rows.push(row);
    return rows;
  }

  async function onWorksheetChanged(event) {
    if (!enabled || event.source !== "Local" || event.worksheetId !== activeSheetId) return;
    const rows = parseStatusRows(event.address);
    if (!rows.length) return;
    for (const rowNumber of rows) {
      try {
        const alertData = await findConflicts(event.worksheetId, rowNumber);
        if (alertData) await openAlert(alertData);
      } catch (error) {
        setSheetStatus("No pude revisar la resolución: " + safeError(error), "error");
      }
    }
  }

  async function findConflicts(sheetId, rowNumber) {
    return Excel.run(async (context) => {
      const sheet = context.workbook.worksheets.getItem(sheetId);
      const currentRanges = ["B", "E", "F", "G"].map((column) => sheet.getRange(column + rowNumber));
      currentRanges.forEach((range) => range.load("text"));
      const used = sheet.getUsedRange(true);
      used.load("rowIndex,rowCount");
      await context.sync();

      const currentCells = currentRanges.map((range) => range.text[0][0] || "");
      const currentName = currentCells[0];
      const currentId = currentCells[1];
      const currentStatus = normalizeResolution(currentCells[2]);
      const currentMotive = currentCells[3];
      const idKey = normalizedId(currentId);
      const currentUserKey = normalized(currentName);
      if (!idKey || !currentStatus || currentStatus === "OTHER") return null;

      const lastRow = used.rowIndex + used.rowCount;
      if (lastRow < 2 || rowNumber <= 2) return null;
      const endPreviousRow = Math.min(lastRow, rowNumber - 1);
      if (endPreviousRow < 2) return null;

      const names = sheet.getRange("B2:B" + endPreviousRow);
      const ids = sheet.getRange("E2:E" + endPreviousRow);
      const statuses = sheet.getRange("F2:F" + endPreviousRow);
      const motives = sheet.getRange("G2:G" + endPreviousRow);
      names.load("text"); ids.load("text"); statuses.load("text"); motives.load("text");
      await context.sync();

      const conflicts = [];
      for (let i = 0; i < ids.text.length; i += 1) {
        if (normalizedId(ids.text[i][0]) !== idKey) continue;
        const priorName = (names.text[i][0] || "").trim();
        if (currentUserKey && normalized(priorName) === currentUserKey) continue;
        const priorStatus = normalizeResolution(statuses.text[i][0]);
        if (!priorStatus || priorStatus === "OTHER" || priorStatus === currentStatus) continue;
        conflicts.push({
          row: i + 2,
          name: priorName || "Usuario no indicado",
          status: priorStatus,
          motive: (motives.text[i][0] || "").trim() || "Sin motivo registrado"
        });
      }
      if (!conflicts.length) return null;
      return {
        sheet: activeSheetName(),
        row: rowNumber,
        id: currentId,
        name: currentName.trim() || "",
        status: currentStatus,
        motive: currentMotive.trim(),
        conflicts: conflicts
      };
    });
  }

  function activeSheetName() {
    const status = document.getElementById("sheetStatus").textContent;
    const match = status.match(/Lista en (.+?)\./);
    return match ? match[1] : "hoja mensual";
  }

  function normalizeResolution(value) {
    const status = normalized(value).replace(/\s+/g, " ");
    if (status === "APLICA") return "APLICA";
    if (status === "NO APLICA") return "NO APLICA";
    return status ? "OTHER" : "";
  }

  function openAlert(data) {
    return new Promise((resolve) => {
      const url = new URL("alert.html", window.location.href).href;
      let dialog = null;
      Office.context.ui.displayDialogAsync(url, { height: 60, width: 48, displayInIframe: true }, (result) => {
        if (result.status !== Office.AsyncResultStatus.Succeeded) {
          showInlineAlert(data);
          resolve();
          return;
        }
        dialog = result.value;
        dialog.addEventHandler(Office.EventType.DialogMessageReceived, (message) => {
          if (message.message === "READY") {
            dialog.messageChild(JSON.stringify(data), { targetOrigin: window.location.origin });
          } else if (message.message === "ACCEPT") {
            dialog.close();
            resolve();
          }
        });
        dialog.addEventHandler(Office.EventType.DialogEventReceived, () => resolve());
      });
    });
  }

  function showInlineAlert(data) {
    const box = document.getElementById("inlineAlert");
    const other = data.conflicts.map((item) => item.status + " · " + item.name + " · " + item.motive).join(" | ");
    box.replaceChildren();
    const message = document.createElement("p");
    message.textContent = "Revisa la cédula " + data.id + ": tu resolución es " + data.status + ", pero hay antecedentes distintos. " + other;
    const accept = document.createElement("button");
    accept.type = "button";
    accept.className = "inline-accept";
    accept.textContent = "Aceptar";
    accept.addEventListener("click", () => { box.hidden = true; });
    box.append(message, accept);
    box.hidden = false;
  }

  function safeError(error) {
    return error && error.message ? error.message : "error desconocido";
  }

  Office.onReady(async (info) => {
    if (info.host !== Office.HostType.Excel) {
      setSheetStatus("Este complemento funciona dentro de Excel Online.", "error");
      return;
    }
    try { enabled = window.localStorage.getItem(STORAGE_KEY) === "1"; } catch (_) { enabled = false; }
    setEnabled(enabled);
    document.getElementById("toggleButton").addEventListener("click", async () => {
      if (!enabled) {
        setEnabled(true);
        await registerActivationHandler();
        await bindActiveMonthlySheet();
      } else {
        setEnabled(false);
      }
    });
    if (enabled) {
      await registerActivationHandler();
      await bindActiveMonthlySheet();
    }
  });
})();


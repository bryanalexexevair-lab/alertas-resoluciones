/* global Office */
(function () {
  "use strict";

  function addConflict(item) {
    const card = document.createElement("article");
    card.className = "conflict";
    const status = document.createElement("div");
    status.className = "conflict-status";
    status.textContent = "RESOLUCIÓN ANTERIOR: " + item.status;
    const name = document.createElement("div");
    name.className = "conflict-name";
    name.textContent = item.name;
    const motive = document.createElement("div");
    motive.className = "conflict-motive";
    motive.textContent = item.motive;
    const row = document.createElement("div");
    row.className = "conflict-row";
    row.textContent = "Fila " + item.row;
    card.append(status, name, motive, row);
    document.getElementById("conflictList").appendChild(card);
  }

  Office.onReady(() => {
    const accept = document.getElementById("acceptButton");
    Office.context.ui.addHandlerAsync(Office.EventType.DialogParentMessageReceived, (event) => {
      if (event.message === "READY") return;
      try {
        const data = JSON.parse(event.message);
        const userLine = data.name ? data.name + ", registraste " + data.status + ". " : "Registraste " + data.status + ". ";
        document.getElementById("summary").textContent = userLine + "La cédula " + data.id + " tiene " + data.conflicts.length + " antecedente(s) de otros usuarios con una resolución distinta.";
        data.conflicts.forEach(addConflict);
        accept.disabled = false;
      } catch (_) {
        document.getElementById("summary").textContent = "No se pudieron mostrar los antecedentes. Cierra la ventana y revisa la cédula en la hoja.";
        accept.disabled = false;
      }
    });
    Office.context.ui.messageParent("READY");
    accept.addEventListener("click", () => Office.context.ui.messageParent("ACCEPT"));
  });
})();


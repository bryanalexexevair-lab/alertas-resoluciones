/* global Office */
(function () {
  "use strict";

  Office.onReady(() => {
    const accept = document.getElementById("acceptButton");
    Office.context.ui.addHandlerAsync(Office.EventType.DialogParentMessageReceived, (event) => {
      if (event.message === "READY") return;
      try {
        const data = JSON.parse(event.message);
        document.getElementById("message").textContent = data.message;
        document.getElementById("detail").textContent = data.detail || "";
      } catch (_) {
        document.getElementById("message").textContent = "Esta cédula ya aparece en los antecedentes de la hoja.";
        document.getElementById("detail").textContent = "Revisa la celda seleccionada antes de continuar.";
      }
      accept.disabled = false;
    });
    Office.context.ui.messageParent("READY");
    accept.addEventListener("click", () => Office.context.ui.messageParent("ACCEPT"));
  });
})();

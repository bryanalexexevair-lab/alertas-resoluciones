/* global Office */
(function () {
  "use strict";

  Office.onReady(() => {
    Office.context.ui.addHandlerAsync(Office.EventType.DialogParentMessageReceived, (event) => {
      try {
        const data = JSON.parse(event.message);
        document.getElementById("message").textContent = data.message || "Esta cédula tiene antecedentes en la hoja.";
        const detail = [];
        if (data.id) detail.push("Cédula " + data.id);
        if (data.detail) detail.push(data.detail);
        document.getElementById("detail").textContent = detail.join(" · ");
      } catch (_) {
        document.getElementById("message").textContent = "Esta cédula tiene antecedentes en la hoja.";
        document.getElementById("detail").textContent = "";
      }
    });
    Office.context.ui.messageParent("READY");
  });
})();

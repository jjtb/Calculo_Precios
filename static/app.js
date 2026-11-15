document.addEventListener("htmx:beforeRequest", (event) => {
  const form = event.detail.elt;
  const button = form && form.querySelector('button[type="submit"]');
  if (button) {
    button.dataset.originalText = button.textContent;
    button.disabled = true;
    button.textContent = "Validando y calculando…";
  }
});

document.addEventListener("htmx:afterRequest", (event) => {
  const form = event.detail.elt;
  const button = form && form.querySelector('button[type="submit"]');
  if (button) {
    button.disabled = false;
    button.textContent = button.dataset.originalText || "Validar insumos y calcular";
  }
  const result = document.getElementById("calculation-result");
  if (result && result.firstElementChild) {
    const success = result.querySelector(".calculation-success");
    if (success) {
      const version = success.dataset.inputSetVersion;
      const options = document.getElementById("calculation-options");
      const status = document.querySelector(".input-set-status");
      if (options) options.open = false;
      if (status && version) {
        status.textContent = `Set de insumos vigente: v${version}.`;
      }
    }
    result.scrollIntoView({ behavior: "smooth", block: "start" });
  }
});

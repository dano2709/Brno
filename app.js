(() => {
  "use strict";
  const selections = { payments: null, base: null, target: null };
  const fields = [
    { key: "payments", input: "paymentsFile", meta: "paymentsMeta", drop: "paymentsDrop", label: "Screenshot plateb" },
    { key: "base", input: "baseFile", meta: "baseMeta", drop: "baseDrop", label: "Základní dokument" },
    { key: "target", input: "targetFile", meta: "targetMeta", label: "Cílový Excel" }
  ];

  const formatSize = (bytes) => {
    if (bytes < 1024) return bytes + " B";
    if (bytes < 1024 * 1024) return (bytes / 1024).toLocaleString("cs-CZ", { maximumFractionDigits: 1 }) + " kB";
    return (bytes / (1024 * 1024)).toLocaleString("cs-CZ", { maximumFractionDigits: 2 }) + " MB";
  };

  function updateField(config, file) {
    selections[config.key] = file || null;
    const meta = document.getElementById(config.meta);
    if (file) {
      meta.textContent = file.name + " · " + formatSize(file.size) + " · načteno pouze lokálně";
      meta.classList.add("selected");
    } else {
      meta.textContent = config.key === "target" ? "Cílová šablona nebyla vybrána" : "Zatím nebyl vybrán soubor";
      meta.classList.remove("selected");
    }
    updateSummary();
  }

  fields.forEach((config) => {
    const input = document.getElementById(config.input);
    input.addEventListener("change", () => updateField(config, input.files && input.files[0]));

    if (config.drop) {
      const drop = document.getElementById(config.drop);
      ["dragenter", "dragover"].forEach(type => drop.addEventListener(type, event => {
        event.preventDefault();
        drop.classList.add("dragging");
      }));
      ["dragleave", "drop"].forEach(type => drop.addEventListener(type, event => {
        event.preventDefault();
        drop.classList.remove("dragging");
      }));
      drop.addEventListener("drop", event => {
        const file = event.dataTransfer && event.dataTransfer.files && event.dataTransfer.files[0];
        if (!file) return;
        const transfer = new DataTransfer();
        transfer.items.add(file);
        input.files = transfer.files;
        updateField(config, file);
      });
    }
  });

  function updateSummary() {
    const ready = Boolean(selections.payments && selections.base);
    document.getElementById("selectedSummary").hidden = !ready;
  }

  document.getElementById("reviewButton").addEventListener("click", () => {
    const list = document.getElementById("reviewList");
    list.replaceChildren();
    fields.forEach(config => {
      const file = selections[config.key];
      if (!file) return;
      const item = document.createElement("div");
      item.className = "review-item";
      const title = document.createElement("strong");
      title.textContent = config.label;
      const details = document.createElement("span");
      details.textContent = file.name + " · " + formatSize(file.size) + " · " + (file.type || "typ souboru nerozpoznán");
      item.append(title, details);
      list.append(item);
    });
    document.getElementById("reviewPanel").hidden = false;
    document.getElementById("reviewPanel").scrollIntoView({ behavior: "smooth", block: "start" });
  });

  document.getElementById("backButton").addEventListener("click", () => {
    document.getElementById("reviewPanel").hidden = true;
    document.getElementById("selectedSummary").scrollIntoView({ behavior: "smooth", block: "center" });
  });
})();
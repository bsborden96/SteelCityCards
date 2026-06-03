const sampleInventory = [
  { SKU: "SC-PSA-1001", Item: "Rookie Heat Slab", Game: "Football", Category: "Slab", Condition: "PSA 10", Qty: 1, Price: 740 },
  { SKU: "SC-WAX-2204", Item: "Friday Night Wax Box", Game: "Baseball", Category: "Sealed", Condition: "Factory sealed", Qty: 8, Price: 189 },
  { SKU: "SC-TCG-3319", Item: "Commander Stack Lot", Game: "Magic", Category: "TCG Singles", Condition: "NM-MT", Qty: 12, Price: 96 },
  { SKU: "SC-PKM-0442", Item: "Charizard Chase Binder", Game: "Pokemon", Category: "TCG Binder", Condition: "Near mint", Qty: 2, Price: 420 },
  { SKU: "SC-MER-7611", Item: "Steel City Black Hoodie", Game: "Merch", Category: "Merch", Condition: "New", Qty: 22, Price: 58 },
  { SKU: "SC-YGO-7602", Item: "Blue-Eyes Display Lot", Game: "Yu-Gi-Oh", Category: "TCG Lot", Condition: "Light play", Qty: 3, Price: 144 }
];

let inventoryRows = [];

const currency = new Intl.NumberFormat("en-US", {
  style: "currency",
  currency: "USD",
  maximumFractionDigits: 0
});

function estimateOffer() {
  const type = document.querySelector("#collectionType").value;
  const count = Number(document.querySelector("#cardCount").value || 0);
  const condition = Number(document.querySelector("#condition").value || 1);
  const deal = document.querySelector("#dealType").value;

  const typeMultipliers = {
    raw: 14,
    slab: 68,
    sealed: 42,
    mixed: 26
  };
  const dealMultipliers = {
    cash: 0.68,
    credit: 0.82,
    consign: 0.9
  };
  const estimated = count * typeMultipliers[type] * (condition / 10) * dealMultipliers[deal];

  document.querySelector("#conditionValue").textContent = condition;
  document.querySelector("#offerValue").textContent = currency.format(Math.max(25, estimated));
}

function normalizeRow(row) {
  const pick = (...keys) => keys.map((key) => row[key]).find((value) => value !== undefined && value !== "");
  return {
    SKU: pick("SKU", "Sku", "sku", "Item ID", "ID") || "UNLISTED",
    Item: pick("Item", "Name", "Product", "Title", "Description") || "Untitled item",
    Game: pick("Game", "Sport", "TCG", "Brand", "Franchise") || "General",
    Category: pick("Category", "Type", "Department") || "Card",
    Condition: pick("Condition", "Grade", "Quality") || "Review",
    Qty: pick("Qty", "Quantity", "Stock") || 1,
    Price: pick("Price", "Retail", "Amount", "Value") || 0
  };
}

function renderInventory(rows = inventoryRows) {
  const body = document.querySelector("#inventoryBody");
  const search = document.querySelector("#inventorySearch").value.trim().toLowerCase();
  const visibleRows = rows.filter((row) => Object.values(row).join(" ").toLowerCase().includes(search));

  body.innerHTML = visibleRows.map((row) => `
    <tr>
      <td>${escapeHtml(row.SKU)}</td>
      <td>${escapeHtml(row.Item)}</td>
      <td>${escapeHtml(row.Game)}</td>
      <td>${escapeHtml(row.Category)}</td>
      <td>${escapeHtml(row.Condition)}</td>
      <td>${escapeHtml(String(row.Qty))}</td>
      <td>${formatPrice(row.Price)}</td>
    </tr>
  `).join("");

  document.querySelector("#inventoryCount").textContent = `${visibleRows.length} item${visibleRows.length === 1 ? "" : "s"}`;
}

function formatPrice(value) {
  const number = Number(String(value).replace(/[$,]/g, ""));
  return Number.isFinite(number) ? currency.format(number) : escapeHtml(String(value));
}

function escapeHtml(value) {
  return String(value)
    .replaceAll("&", "&amp;")
    .replaceAll("<", "&lt;")
    .replaceAll(">", "&gt;")
    .replaceAll('"', "&quot;")
    .replaceAll("'", "&#039;");
}

function parseCsv(text) {
  const lines = text.trim().split(/\r?\n/);
  const headers = lines.shift().split(",").map((header) => header.trim());
  return lines.map((line) => {
    const values = line.split(",").map((value) => value.trim());
    return headers.reduce((row, header, index) => {
      row[header] = values[index] || "";
      return row;
    }, {});
  });
}

function handleInventoryFile(file) {
  const reader = new FileReader();
  const isCsv = file.name.toLowerCase().endsWith(".csv");

  reader.onload = (event) => {
    let rows = [];
    if (isCsv || !window.XLSX) {
      if (!isCsv && !window.XLSX) {
        document.querySelector("#databaseNote").textContent = "Excel parsing needs the connected XLSX library. Export the stock sheet as CSV or enable the library on the live site.";
        return;
      }
      const text = isCsv ? event.target.result : "";
      rows = text ? parseCsv(text) : [];
    } else {
      const workbook = XLSX.read(event.target.result, { type: "array" });
      const sheet = workbook.Sheets[workbook.SheetNames[0]];
      rows = XLSX.utils.sheet_to_json(sheet);
    }

    inventoryRows = rows.map(normalizeRow);
    renderInventory();
    document.querySelector("#databaseNote").textContent = `${inventoryRows.length} stock rows loaded for staff review. Connect the endpoint to sync them live.`;
  };

  if (isCsv) {
    reader.readAsText(file);
  } else {
    reader.readAsArrayBuffer(file);
  }
}

function showToast(message) {
  const toast = document.querySelector("#cartToast");
  toast.textContent = message;
  toast.classList.add("show");
  window.clearTimeout(showToast.timer);
  showToast.timer = window.setTimeout(() => toast.classList.remove("show"), 2400);
}

document.querySelectorAll("#collectionType, #cardCount, #condition, #dealType").forEach((field) => {
  field.addEventListener("input", estimateOffer);
});

document.querySelector("#sellForm").addEventListener("submit", (event) => {
  event.preventDefault();
  document.querySelector("#sellNote").textContent = "Seller review queued. Connect this form to email, CRM, or checkout intake when the site goes live.";
});

document.querySelector("#newsletterForm").addEventListener("submit", (event) => {
  event.preventDefault();
  const email = document.querySelector("#newsletterEmail").value;
  document.querySelector("#newsletterNote").textContent = `${email} is on the early-drop list.`;
  event.currentTarget.reset();
});

document.querySelector("#staffToggle").addEventListener("click", () => {
  const panel = document.querySelector("#staffPanel");
  panel.hidden = !panel.hidden;
  document.querySelector("#staffToggle").textContent = panel.hidden ? "Staff Stock Portal" : "Hide Staff Portal";
});

document.querySelector("#saveConnection").addEventListener("click", () => {
  const endpoint = document.querySelector("#databaseEndpoint").value.trim();
  const staffKey = document.querySelector("#staffKey").value.trim();
  const note = document.querySelector("#databaseNote");

  if (!endpoint || !staffKey) {
    note.textContent = "Add a database endpoint and staff key before saving the connection.";
    return;
  }

  localStorage.setItem("steelCityInventoryEndpoint", endpoint);
  note.textContent = "Connection saved in this browser. A backend API can now receive stock uploads from this portal.";
});

document.querySelectorAll(".filter").forEach((button) => {
  button.addEventListener("click", () => {
    document.querySelectorAll(".filter").forEach((filter) => filter.classList.remove("is-active"));
    button.classList.add("is-active");

    const type = button.dataset.filter;
    document.querySelectorAll(".product-card").forEach((card) => {
      card.hidden = type !== "all" && card.dataset.type !== type;
    });
  });
});

document.querySelectorAll("[data-cart]").forEach((button) => {
  button.addEventListener("click", () => showToast(`${button.dataset.cart} added to inquiry cart.`));
});

document.querySelector("#inventoryUpload").addEventListener("change", (event) => {
  const [file] = event.target.files;
  if (file) {
    handleInventoryFile(file);
  }
});

document.querySelector("#inventorySearch").addEventListener("input", () => renderInventory());

document.querySelector("#loadSample").addEventListener("click", () => {

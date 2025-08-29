/* global Office, Word */
const TILE_TEMPLATES = [
  { id:"TextTile",      type:"TEXT",      label:"Pole tekstowe",
    sample:"{{ FIELD_NAME }}" },
  { id:"ConditionTile", type:"CONDITION", label:"Blok warunkowy",
    sample:"{{ START_is_visible }}\n<treść>\n{{ END_is_visible }}" },
  { id:"ProductTable",  type:"TABLE",     label:"Tabela produktów (znacznik)",
    sample:"{{ INSERT_PRODUCT_TABLE }}" }
];


Office.onReady(() => {
  renderTiles();
  bindUI();
  refreshPlaceholders();
});

function bindUI(){
  const plus = document.getElementById("btnInsertPlus");
  if (plus) plus.addEventListener("click", insertSelectedOrFirstTile);
  const createBtn = document.getElementById("btnCreateTile");
  if (createBtn) createBtn.addEventListener("click", createTileFromForm);

  window.addEventListener("keydown", (e) => {
    if (e.key === "+" || (e.key === "=" && e.shiftKey)) {
      e.preventDefault();
      insertSelectedOrFirstTile();
    }
  });
}

let selectedTileId = TILE_TEMPLATES[0].id;

function renderTiles(){
  const root = document.getElementById("tiles");
  root.innerHTML = "";
  TILE_TEMPLATES.forEach(tile => {
    const row = document.createElement("div");
    row.className = "item tile " + (tile.type ? `type-${tile.type}` : "");
  const root = document.getElementById("tiles");
  root.innerHTML = "";
  TILE_TEMPLATES.forEach(tile => {
    const row = document.createElement("div");
    row.className = "item";
    const left = document.createElement("div");
    left.style.display="flex"; left.style.alignItems="center"; left.style.gap="8px";

    const radio = document.createElement("input");
    radio.type="radio"; radio.name="tile"; radio.checked = tile.id === selectedTileId;
    radio.addEventListener("change", () => selectedTileId = tile.id);

    const label = document.createElement("div");
    label.innerHTML = `<strong>${tile.label}</strong><br/><code>${escapeHtml(tile.sample)}</code>`;

    left.appendChild(radio); left.appendChild(label);

    const btn = document.createElement("button");
    btn.textContent = "Wstaw";
    btn.addEventListener("click", () => insertTile(tile.sample));

    row.appendChild(left); row.appendChild(btn);
    root.appendChild(row);
  });
}

function insertSelectedOrFirstTile(){
  const tile = TILE_TEMPLATES.find(t => t.id === selectedTileId) || TILE_TEMPLATES[0];
  insertTile(tile.sample);
}

async function insertTile(text){
  await Word.run(async (context) => {
    const range = context.document.getSelection();
    range.insertText(text, Word.InsertLocation.replace);
    await context.sync();
  });
}

async function refreshPlaceholders(){
  const container = document.getElementById("phList");
  container.textContent = "(wczytywanie…)";
  try {
    const list = await extractPlaceholdersFromActiveDocument();
    if (!list.length){ container.textContent = "Brak placeholderów w dokumencie."; return; }
    container.innerHTML = "";
    list.sort().forEach(name => {
      const row = document.createElement("div");
      row.className = "item";
      row.innerHTML = `<div><code>{{ ${escapeHtml(name)} }}</code></div>`;
      const btn = document.createElement("button");
      btn.textContent = "Wstaw";
      btn.addEventListener("click", () => insertTile(`{{ ${name} }}`));
      row.appendChild(btn);
      container.appendChild(row);
    });
  } catch(e){
    container.textContent = `Błąd: ${e.message || e}`;
  }
}

async function extractPlaceholdersFromActiveDocument(){
  return Word.run(async (context) => {
    const ooxml = context.document.body.getOoxml();
    await context.sync();
    const xml = ooxml.value || "";
    const set = new Set();
    const re = /\{\{\s*([^}]+?)\s*\}\}/g;
    let m;
    while ((m = re.exec(xml)) !== null){
      const raw = m[1].trim();
      set.add(raw);
    }
    return Array.from(set);
  });
}

function escapeHtml(s){
  return s.replaceAll("&","&amp;").replaceAll("<","&lt;").replaceAll(">","&gt;");
}

function createTileFromForm(){
  const name = (document.getElementById("tileName").value || "").trim();
  const type = document.getElementById("tileType").value;
  if (!name && type === "TEXT"){ alert("Podaj nazwę pola."); return; }

  let sample = "";
  switch(type){
    case "TEXT":      sample = `{{ ${name} }}`; break;
    case "CONDITION": const cond = name || "is_visible";
                      sample = `{{ START_${cond} }}\n<treść>\n{{ END_${cond} }}`; break;
    case "TABLE":     sample = "{{ INSERT_PRODUCT_TABLE }}"; break;
    case "IMAGE":     sample = `{{ IMAGE_${name || "logo"} }}`; break;
  }

  const id = `Custom_${Date.now()}`;
  TILE_TEMPLATES.unshift({
    id, type,
    label: name ? `${name} (${type})` : type,
    sample
  });
  selectedTileId = id;
  renderTiles();
}


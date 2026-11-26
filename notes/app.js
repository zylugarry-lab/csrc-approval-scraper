const KEY = "tool-hub-notes-v1";
const authorEl = document.querySelector("#author");
const targetEl = document.querySelector("#target");
const bodyEl = document.querySelector("#body");
const listEl = document.querySelector("#list");
const statusEl = document.querySelector("#status");
const addBtn = document.querySelector("#add");

let notes = [];
let editingId = null;

function now() {
  return new Date().toISOString();
}

function formatTime(iso) {
  if (!iso) return "";
  const d = new Date(iso);
  if (Number.isNaN(d.getTime())) return iso;
  const pad = (n) => String(n).padStart(2, "0");
  return `${d.getFullYear()}-${pad(d.getMonth() + 1)}-${pad(d.getDate())} ${pad(d.getHours())}:${pad(d.getMinutes())}`;
}

function load() {
  try {
    const raw = JSON.parse(localStorage.getItem(KEY) || "{}");
    notes = Array.isArray(raw.notes) ? raw.notes : [];
    if (raw.author) authorEl.value = raw.author;
  } catch {
    notes = [];
  }
}

function persist() {
  localStorage.setItem(KEY, JSON.stringify({
    author: authorEl.value.trim(),
    notes,
    savedAt: now(),
  }));
}

function setStatus(text) {
  statusEl.textContent = text;
}

function render() {
  listEl.innerHTML = "";
  if (!notes.length) {
    const empty = document.createElement("p");
    empty.className = "empty";
    empty.textContent = "还没有意见。写完点添加。";
    listEl.append(empty);
    return;
  }
  notes.slice().reverse().forEach((note) => {
    const li = document.createElement("li");
    const meta = document.createElement("div");
    meta.className = "meta";
    const who = note.author || "未署名";
    const about = note.target ? ` · ${note.target}` : "";
    const edited = note.updatedAt && note.updatedAt !== note.createdAt ? ` · 改于 ${formatTime(note.updatedAt)}` : "";
    meta.textContent = `${who}${about} · ${formatTime(note.createdAt)}${edited}`;
    const body = document.createElement("div");
    body.className = "body";
    body.textContent = note.body;
    const actions = document.createElement("div");
    actions.className = "actions";
    const edit = document.createElement("button");
    edit.className = "ghost";
    edit.type = "button";
    edit.textContent = "改这条";
    edit.addEventListener("click", () => startEdit(note.id));
    const remove = document.createElement("button");
    remove.className = "danger";
    remove.type = "button";
    remove.textContent = "删除";
    remove.addEventListener("click", () => removeNote(note.id));
    actions.append(edit, remove);
    li.append(meta, body, actions);
    listEl.append(li);
  });
}

function startEdit(id) {
  const note = notes.find((item) => item.id === id);
  if (!note) return;
  editingId = id;
  authorEl.value = note.author || "";
  targetEl.value = note.target || "";
  bodyEl.value = note.body || "";
  addBtn.textContent = "保存这条";
  bodyEl.focus();
  setStatus("正在改已有意见。保存这条后会写回列表。");
}

function clearForm() {
  editingId = null;
  targetEl.value = "";
  bodyEl.value = "";
  addBtn.textContent = "添加";
}

function upsert() {
  const body = bodyEl.value.trim();
  if (!body) {
    setStatus("先写意见再添加。");
    bodyEl.focus();
    return;
  }
  const author = authorEl.value.trim();
  const target = targetEl.value.trim();
  if (editingId) {
    const note = notes.find((item) => item.id === editingId);
    if (note) {
      note.author = author;
      note.target = target;
      note.body = body;
      note.updatedAt = now();
    }
    setStatus("这条已保存。");
  } else {
    notes.push({
      id: crypto.randomUUID ? crypto.randomUUID() : String(Date.now()) + Math.random(),
      author,
      target,
      body,
      createdAt: now(),
      updatedAt: now(),
    });
    setStatus("已添加，并保存在这台电脑。");
  }
  persist();
  clearForm();
  render();
}

function removeNote(id) {
  notes = notes.filter((item) => item.id !== id);
  if (editingId === id) clearForm();
  persist();
  render();
  setStatus("已删除。");
}

function download(name, text, type) {
  const blob = new Blob([text], { type });
  const url = URL.createObjectURL(blob);
  const a = document.createElement("a");
  a.href = url;
  a.download = name;
  document.body.append(a);
  a.click();
  a.remove();
  setTimeout(() => URL.revokeObjectURL(url), 1500);
}

function stamp() {
  const d = new Date();
  const pad = (n) => String(n).padStart(2, "0");
  return `${d.getFullYear()}${pad(d.getMonth() + 1)}${pad(d.getDate())}`;
}

function saveBackup() {
  persist();
  const payload = {
    kind: "tool-hub-notes",
    savedAt: now(),
    notes,
  };
  download(`修改意见-${stamp()}.json`, JSON.stringify(payload, null, 2), "application/json");
  setStatus("已下载备份。发给同事后，对方点导入备份即可合并。");
}

function saveText() {
  persist();
  if (!notes.length) {
    setStatus("还没有意见可下载。");
    return;
  }
  const lines = ["修改意见", ""];
  notes.forEach((note, i) => {
    lines.push(`${i + 1}. ${note.author || "未署名"}${note.target ? " · " + note.target : ""} · ${formatTime(note.createdAt)}`);
    lines.push(note.body);
    lines.push("");
  });
  download(`修改意见-${stamp()}.txt`, lines.join("\n"), "text/plain;charset=utf-8");
  setStatus("已下载文本。");
}

function importBackup(file) {
  const reader = new FileReader();
  reader.onload = () => {
    try {
      const data = JSON.parse(String(reader.result || "{}"));
      const incoming = Array.isArray(data.notes) ? data.notes : Array.isArray(data) ? data : [];
      if (!incoming.length) {
        setStatus("这个文件里没有意见。");
        return;
      }
      const map = new Map(notes.map((item) => [item.id, item]));
      let added = 0;
      incoming.forEach((item) => {
        if (!item || !item.body) return;
        const id = item.id || (crypto.randomUUID ? crypto.randomUUID() : String(Date.now()) + Math.random());
        const prev = map.get(id);
        if (!prev) {
          map.set(id, { ...item, id });
          added += 1;
          return;
        }
        if (String(item.updatedAt || "") > String(prev.updatedAt || "")) map.set(id, { ...item, id });
      });
      notes = [...map.values()].sort((a, b) => String(a.createdAt).localeCompare(String(b.createdAt)));
      persist();
      render();
      setStatus(added ? `已导入，新增 ${added} 条。` : "已导入。没有新的条目，较新的修改已覆盖。");
    } catch {
      setStatus("这个文件不是能导入的备份。");
    }
  };
  reader.readAsText(file, "utf-8");
}

addBtn.addEventListener("click", upsert);
document.querySelector("#save").addEventListener("click", saveBackup);
document.querySelector("#text").addEventListener("click", saveText);
document.querySelector("#import").addEventListener("change", (event) => {
  const file = event.target.files && event.target.files[0];
  if (file) importBackup(file);
  event.target.value = "";
});
bodyEl.addEventListener("keydown", (event) => {
  if ((event.ctrlKey || event.metaKey) && event.key === "Enter") {
    event.preventDefault();
    upsert();
  }
});
authorEl.addEventListener("change", persist);

load();
render();
setStatus(notes.length ? `本机已有 ${notes.length} 条。可继续添加，或保存备份。` : "写好后点添加。需要给同事时再保存备份。");

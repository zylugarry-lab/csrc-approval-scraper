const KIT_FILES = [
  "__init__.py",
  "engine.py",
  "formatter.py",
  "formula.py",
  "classifier.py",
  "config.py",
  "xmlutil.py",
  "sample.py",
  "__main__.py",
  "config.yaml",
];

const statusEl = document.querySelector("#status");
const listEl = document.querySelector("#list");
const runBtn = document.querySelector("#run");
const zipBtn = document.querySelector("#zip");
const inputEl = document.querySelector("#files");
const dropEl = document.querySelector("#drop");

let pyodide = null;
let loading = null;
const queue = [];
const done = [];

function setStatus(text) {
  statusEl.textContent = text;
}

function render() {
  listEl.innerHTML = "";
  queue.forEach((item, index) => {
    const li = document.createElement("li");
    const name = document.createElement("span");
    name.textContent = item.file.name;
    const state = document.createElement("span");
    state.textContent = item.state;
    state.className = "state";
    li.append(name, state);
    if (item.blob) {
      const a = document.createElement("a");
      a.href = URL.createObjectURL(item.blob);
      a.download = item.outName;
      a.textContent = "下载";
      li.append(a);
    }
    if (item.error) {
      const err = document.createElement("div");
      err.className = "err";
      err.textContent = item.error;
      li.append(err);
    }
    listEl.append(li);
    item.index = index;
  });
  const ready = queue.some((item) => item.state === "完成");
  zipBtn.hidden = queue.filter((item) => item.blob).length < 2;
  runBtn.disabled = !queue.length || queue.some((item) => item.state === "刷新中");
}

function addFiles(fileList) {
  for (const file of fileList) {
    const lower = file.name.toLowerCase();
    if (file.name.startsWith("~$")) continue;
    if (!lower.endsWith(".docx")) {
      queue.push({
        file,
        state: "跳过",
        error: "只接受 .docx。旧版 .doc 请先在 Word 或 WPS 里另存为 .docx。",
      });
      continue;
    }
    if (queue.some((item) => item.file.name === file.name && item.file.size === file.size)) continue;
    queue.push({ file, state: "等待" });
  }
  render();
}

inputEl.addEventListener("change", () => {
  addFiles(inputEl.files);
  inputEl.value = "";
});

dropEl.addEventListener("dragover", (event) => {
  event.preventDefault();
  dropEl.classList.add("over");
});
dropEl.addEventListener("dragleave", () => dropEl.classList.remove("over"));
dropEl.addEventListener("drop", (event) => {
  event.preventDefault();
  dropEl.classList.remove("over");
  addFiles(event.dataTransfer.files);
});

async function ensureEngine() {
  if (pyodide) return pyodide;
  if (loading) return loading;
  loading = (async () => {
    setStatus("正在加载排版程序，第一次大约需要十几秒。");
    const engine = await loadPyodide({
      indexURL: "https://cdn.jsdelivr.net/pyodide/v0.26.4/full/",
    });
    await engine.loadPackage(["micropip", "lxml", "pyyaml", "typing-extensions"]);
    const wheel = new Uint8Array(await (await fetch("vendor/python_docx-1.1.2-py3-none-any.whl")).arrayBuffer());
    engine.FS.writeFile("/python_docx-1.1.2-py3-none-any.whl", wheel);
    await engine.runPythonAsync(`
import micropip
await micropip.install("emfs:/python_docx-1.1.2-py3-none-any.whl", deps=False)
`);
    engine.FS.mkdirTree("/format_kit");
    for (const name of KIT_FILES) {
      const response = await fetch("kit/" + name);
      if (!response.ok) throw new Error("缺少排版文件 " + name);
      const text = await response.text();
      const target = name === "config.yaml" ? "/config.yaml" : "/format_kit/" + name;
      engine.FS.writeFile(target, text);
    }
    engine.FS.mkdirTree("/work/in");
    engine.FS.mkdirTree("/work/out");
    engine.runPython(`
import sys
sys.path.insert(0, "/")
`);
    setStatus("排版程序已就绪。选择 Word 后点开始刷新。");
    pyodide = engine;
    return engine;
  })();
  try {
    return await loading;
  } catch (error) {
    loading = null;
    throw error;
  }
}

function downloadBlob(blob, name) {
  const url = URL.createObjectURL(blob);
  const a = document.createElement("a");
  a.href = url;
  a.download = name;
  document.body.append(a);
  a.click();
  a.remove();
  setTimeout(() => URL.revokeObjectURL(url), 2000);
}

function outName(name) {
  return name.replace(/\.docx$/i, "") + "_已刷新.docx";
}

async function refreshOne(engine, item) {
  const buf = new Uint8Array(await item.file.arrayBuffer());
  engine.FS.writeFile("/work/in/input.docx", buf);
  const result = engine.runPython(`
from pathlib import Path
try:
    from format_kit.engine import refresh_document
    report = refresh_document(
        "/work/in/input.docx",
        "/work/out/out.docx",
        config_path=Path("/config.yaml"),
    )
    "OK\\n" + report.as_text()
except Exception as exc:
    "ERR\\n" + str(exc)
`);
  if (String(result).startsWith("ERR")) {
    throw new Error(String(result).slice(4).trim());
  }
  const bytes = engine.FS.readFile("/work/out/out.docx");
  return new Blob([bytes], {
    type: "application/vnd.openxmlformats-officedocument.wordprocessingml.document",
  });
}

runBtn.addEventListener("click", async () => {
  runBtn.disabled = true;
  try {
    const engine = await ensureEngine();
    for (const item of queue) {
      if (item.state === "完成" || item.state === "跳过") continue;
      item.state = "刷新中";
      item.error = "";
      render();
      setStatus("正在刷新：" + item.file.name);
      try {
        item.blob = await refreshOne(engine, item);
        item.outName = outName(item.file.name);
        item.state = "完成";
        done.push(item);
      } catch (error) {
        item.state = "失败";
        item.error = error && error.message ? error.message : String(error);
      }
      render();
    }
    const ok = queue.filter((item) => item.blob);
    if (ok.length === 1) downloadBlob(ok[0].blob, ok[0].outName);
    if (ok.length > 1) await downloadZip(ok);
    const failed = queue.filter((item) => item.state === "失败").length;
    setStatus(
      failed
        ? "完成 " + ok.length + " 份，失败 " + failed + " 份。"
        : "已刷新 " + ok.length + " 份，文件已开始下载。"
    );
  } catch (error) {
    setStatus(error && error.message ? error.message : String(error));
  }
  render();
});

async function downloadZip(items) {
  const zip = new JSZip();
  const used = new Map();
  for (const item of items) {
    let name = item.outName;
    if (used.has(name)) {
      used.set(name, used.get(name) + 1);
      name = name.replace(/\.docx$/i, "-" + used.get(item.outName) + ".docx");
    } else {
      used.set(name, 1);
    }
    zip.file(name, await item.blob.arrayBuffer());
  }
  const blob = await zip.generateAsync({ type: "blob" });
  downloadBlob(blob, "格式刷新.zip");
}

zipBtn.addEventListener("click", async () => {
  const ok = queue.filter((item) => item.blob);
  if (ok.length) await downloadZip(ok);
});

ensureEngine().catch((error) => {
  setStatus(error && error.message ? error.message : String(error));
});

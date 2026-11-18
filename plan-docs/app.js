const KIT_FILES = [
  "extract_plan.py",
  "generate_说明书.py",
  "generate_风险揭示书.py",
  "contract_sections.py",
  "risk_sections.py",
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

function setStatus(text) {
  statusEl.textContent = text;
}

function stemOf(name) {
  return name.replace(/\.docx$/i, "").replace(/管理合同.*$/, "").replace(/合同$/, "");
}

function render() {
  listEl.innerHTML = "";
  queue.forEach((item) => {
    const li = document.createElement("li");
    const name = document.createElement("div");
    name.className = "name";
    name.textContent = item.file.name;
    li.append(name);
    if (item.sms && item.risk) {
      const links = document.createElement("div");
      links.className = "links";
      const a = document.createElement("a");
      a.href = URL.createObjectURL(item.sms);
      a.download = item.smsName;
      a.textContent = "下载说明书";
      const b = document.createElement("a");
      b.href = URL.createObjectURL(item.risk);
      b.download = item.riskName;
      b.textContent = "下载风险揭示书";
      links.append(a, b);
      li.append(links);
    }
    if (item.report) {
      const report = document.createElement("div");
      report.className = "report";
      report.textContent = item.report;
      li.append(report);
    }
    if (item.error) {
      const err = document.createElement("div");
      err.className = "report";
      err.textContent = item.error;
      li.append(err);
    }
    listEl.append(li);
  });
  const blobs = queue.filter((item) => item.sms && item.risk);
  zipBtn.hidden = blobs.length < 1;
  runBtn.disabled = !queue.length || queue.some((item) => item.state === "生成中");
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
    setStatus("正在加载生成程序，第一次大约需要十几秒。");
    const engine = await loadPyodide({
      indexURL: "https://cdn.jsdelivr.net/pyodide/v0.26.4/full/",
    });
    await engine.loadPackage(["micropip", "lxml", "pyyaml", "typing-extensions"]);
    const wheel = new Uint8Array(
      await (await fetch("vendor/python_docx-1.1.2-py3-none-any.whl")).arrayBuffer()
    );
    engine.FS.writeFile("/python_docx-1.1.2-py3-none-any.whl", wheel);
    await engine.runPythonAsync(`
import micropip
await micropip.install("emfs:/python_docx-1.1.2-py3-none-any.whl", deps=False)
`);
    engine.FS.mkdirTree("/kit");
    engine.FS.mkdirTree("/template");
    engine.FS.mkdirTree("/work");
    for (const name of KIT_FILES) {
      const response = await fetch("kit/" + name);
      if (!response.ok) throw new Error("缺少程序文件 " + name);
      engine.FS.writeFile("/kit/" + name, await response.text());
    }
    const smsTpl = new Uint8Array(await (await fetch("template/sms-template.docx")).arrayBuffer());
    const riskTpl = new Uint8Array(await (await fetch("template/risk-template.docx")).arrayBuffer());
    engine.FS.writeFile("/template/sms-template.docx", smsTpl);
    engine.FS.writeFile("/template/risk-template.docx", riskTpl);
    engine.runPython(`
import sys
sys.path.insert(0, "/kit")
`);
    setStatus("生成程序已就绪。选择合同后点生成并下载。");
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

async function generateOne(engine, item) {
  const buf = new Uint8Array(await item.file.arrayBuffer());
  engine.FS.writeFile("/work/in.docx", buf);
  const report = engine.runPython(`
from extract_plan import write_pair
try:
    text = write_pair(
        "/work/in.docx",
        "/template/sms-template.docx",
        "/template/risk-template.docx",
        "/work/sms.docx",
        "/work/risk.docx",
        year=2026,
    )
    "OK\\n" + text
except Exception as exc:
    "ERR\\n" + str(exc)
`);
  if (String(report).startsWith("ERR")) {
    throw new Error(String(report).slice(4).trim());
  }
  const smsBytes = engine.FS.readFile("/work/sms.docx");
  const riskBytes = engine.FS.readFile("/work/risk.docx");
  const type = "application/vnd.openxmlformats-officedocument.wordprocessingml.document";
  const stem = stemOf(item.file.name);
  item.smsName = stem + "说明书-自动生成.docx";
  item.riskName = stem + "风险揭示书-自动生成.docx";
  item.sms = new Blob([smsBytes], { type });
  item.risk = new Blob([riskBytes], { type });
  item.report = String(report).replace(/^OK\n/, "");
}

async function zipAll() {
  const zip = new JSZip();
  queue.forEach((item) => {
    if (!item.sms || !item.risk) return;
    zip.file(item.smsName, item.sms);
    zip.file(item.riskName, item.risk);
  });
  const blob = await zip.generateAsync({ type: "blob" });
  downloadBlob(blob, "说明书与风险揭示书.zip");
}

runBtn.addEventListener("click", async () => {
  runBtn.disabled = true;
  try {
    const engine = await ensureEngine();
    for (const item of queue) {
      if (item.state === "完成" || item.state === "跳过") continue;
      item.state = "生成中";
      item.error = "";
      render();
      setStatus("正在生成：" + item.file.name);
      try {
        await generateOne(engine, item);
        item.state = "完成";
      } catch (error) {
        item.state = "失败";
        item.error = error && error.message ? error.message : String(error);
      }
      render();
    }
    const done = queue.filter((item) => item.sms && item.risk);
    if (done.length === 1) {
      downloadBlob(done[0].sms, done[0].smsName);
      downloadBlob(done[0].risk, done[0].riskName);
    } else if (done.length > 1) {
      await zipAll();
    }
    setStatus(done.length ? "已生成。请对照合同看一遍费率、业绩报酬和风险正文。" : "没有生成文件。");
  } catch (error) {
    setStatus(error && error.message ? error.message : String(error));
  }
  render();
});

zipBtn.addEventListener("click", () => {
  zipAll().catch((error) => setStatus(error && error.message ? error.message : String(error)));
});

ensureEngine().catch((error) => {
  setStatus(error && error.message ? error.message : String(error));
});

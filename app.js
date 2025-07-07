<script src="https://cdnjs.cloudflare.com/ajax/libs/jszip/3.7.1/jszip.min.js"></script>

const editor = document.getElementById("editor");
const highlighter = document.getElementById("highlighting");
const themeSelector = document.getElementById("themeSelector");
const projectSelector = document.getElementById("projectSelector");
const moduleSelector = document.getElementById("moduleSelector");

const KEYWORDS = [
  "Sub", "End Sub", "Function", "End Function", "Dim", "Set", "If", "Then", "Else", "End If",
  "For", "Each", "In", "Next", "Do", "Loop", "While", "Wend", "With", "End With",
  "Select Case", "Case", "End Select", "Const", "Private", "Public", "ByVal", "ByRef", "Exit Sub",
  "Integer", "Long", "String", "Boolean", "Variant", "Object", "Nothing", "True", "False",
  "Call", "GoTo", "On Error", "Resume", "Redim", "Preserve", "Option Explicit"
];

let currentProject = "Progetto1";
let currentModule = "Modulo1";
let projects = JSON.parse(localStorage.getItem("vbaber-projects") || "{}");

function getModules(projectName) {
  return projects[projectName] || {};
}

function saveModule(project, module, content) {
  if (!projects[project]) projects[project] = {};
  projects[project][module] = content;
  localStorage.setItem("vbaber-projects", JSON.stringify(projects));
}

function loadModule(project, module) {
  editor.innerText = (projects[project] && projects[project][module]) || "";
  currentModule = module;
  highlight();
}

function updateProjectSelector() {
  projectSelector.innerHTML = "";
  Object.keys(projects).forEach(name => {
    const opt = document.createElement("option");
    opt.value = name;
    opt.textContent = name;
    if (name === currentProject) opt.selected = true;
    projectSelector.appendChild(opt);
  });
}

function updateModuleSelector() {
  moduleSelector.innerHTML = "";
  const modules = getModules(currentProject);
  Object.keys(modules).forEach(name => {
    const opt = document.createElement("option");
    opt.value = name;
    opt.textContent = name;
    if (name === currentModule) opt.selected = true;
    moduleSelector.appendChild(opt);
  });
}

function exportProject() {
  const content = projects[currentProject][currentModule];
  const blob = new Blob([content], {type: "text/plain"});
  const link = document.createElement("a");
  link.href = URL.createObjectURL(blob);
  link.download = `${currentModule}.bas`;
  link.click();
}

function addModule() {
  const name = prompt("Nome nuovo modulo:");
  if (name) {
    if (!projects[currentProject]) projects[currentProject] = {};
    projects[currentProject][name] = "";
    localStorage.setItem("vbaber-projects", JSON.stringify(projects));
    updateModuleSelector();
    loadModule(currentProject, name);
  }
}

editor.addEventListener("input", () => {
  saveModule(currentProject, currentModule, editor.innerText);
  highlight();
});

moduleSelector.addEventListener("change", () => {
  saveModule(currentProject, currentModule, editor.innerText);
  loadModule(currentProject, moduleSelector.value);
});

projectSelector.addEventListener("change", () => {
  saveModule(currentProject, currentModule, editor.innerText);
  loadModule(projectSelector.value, Object.keys(getModules(projectSelector.value))[0] || "Modulo1");
  currentProject = projectSelector.value;
  updateModuleSelector();
});

themeSelector.addEventListener("change", () => {
  document.body.className = themeSelector.value;
  localStorage.setItem("vbaber-theme", themeSelector.value);
});

window.addEventListener("DOMContentLoaded", () => {
  const savedTheme = localStorage.getItem("vbaber-theme") || "light";
  document.body.className = savedTheme;
  themeSelector.value = savedTheme;
  if (!projects["Progetto1"]) projects["Progetto1"] = { "Modulo1": "" };
  updateProjectSelector();
  updateModuleSelector();
  loadModule(currentProject, currentModule);
});

function updateProcedureList() {
  const code = editor.innerText;
  const lines = code.split("\n");
  const procedures = [];
  lines.forEach((line, idx) => {
    const match = line.match(/\b(Sub|Function)\s+(\w+)/i);
    if (match) {
      procedures.push({ name: match[2], line: idx });
    }
  });
  const procSel = document.getElementById("procedureSelector");
  procSel.innerHTML = '<option value="">Vai a Sub/Function</option>';
  procedures.forEach(proc => {
    const opt = document.createElement("option");
    opt.value = proc.line;
    opt.textContent = proc.name;
    procSel.appendChild(opt);
  });
}

function goToProcedure() {
  const line = parseInt(document.getElementById("procedureSelector").value);
  if (!isNaN(line)) {
    const range = document.createRange();
    const sel = window.getSelection();
    const lines = editor.childNodes;
    let charCount = 0, currentLine = 0;
    for (let node of editor.childNodes) {
      if (node.nodeType === Node.TEXT_NODE) {
        const nodeLines = node.textContent.split("\n");
        for (let i = 0; i < nodeLines.length; i++) {
          if (currentLine === line) {
            range.setStart(node, charCount);
            range.collapse(true);
            sel.removeAllRanges();
            sel.addRange(range);
            editor.focus();
            return;
          }
          currentLine++;
          charCount += nodeLines[i].length + 1;
        }
      }
    }
  }
}

editor.addEventListener("input", () => {
  saveModule(currentProject, currentModule, editor.innerText);
  highlight();
  updateProcedureList();
});

function importModule() {
  const file = document.getElementById('fileInput').files[0];
  if (!file) return;
  const reader = new FileReader();
  reader.onload = () => {
    const name = file.name.replace(/\.(bas|cls)$/i, "") || "Importato";
    if (!projects[currentProject]) projects[currentProject] = {};
    projects[currentProject][name] = reader.result;
    localStorage.setItem("vbaber-projects", JSON.stringify(projects));
    updateModuleSelector();
    loadModule(currentProject, name);
  };
  reader.readAsText(file);
}

function exportAll() {
  const allModules = projects[currentProject];
  const zip = new JSZip();
  Object.entries(allModules).forEach(([name, content]) => {
    zip.file(name + ".bas", content);
  });
  zip.generateAsync({type: "blob"}).then(function(blob) {
    const link = document.createElement("a");
    link.href = URL.createObjectURL(blob);
    link.download = currentProject + ".zip";
    link.click();
  });
}

function checkBlockCoherence() {
  const code = editor.innerText;
  const lines = code.split('\n');
  const stack = [];
  const errors = [];

  const pairs = {
    "If": "End If",
    "For": "Next",
    "While": "Wend",
    "Do": "Loop",
    "Select Case": "End Select",
    "With": "End With"
  };

  const openers = Object.keys(pairs);
  const closers = Object.values(pairs);

  lines.forEach((line, idx) => {
    const trimmed = line.trim();
    for (let open of openers) {
      if (trimmed.match(new RegExp(`^${open}\b`, "i"))) {
        stack.push({ type: open, line: idx + 1 });
      }
    }
    for (let close of closers) {
      if (trimmed.match(new RegExp(`^${close}\b`, "i"))) {
        const expectedOpen = openers[closers.indexOf(close)];
        const last = stack.pop();
        if (!last || last.type !== expectedOpen) {
          errors.push(`Errore a riga ${idx + 1}: '${close}' senza '${expectedOpen}' corrispondente.`);
        }
      }
    }
  });

  stack.forEach(unclosed => {
    errors.push(`Blocco '${unclosed.type}' aperto a riga ${unclosed.line} non chiuso.`);
  });

  showBlockErrors(errors);
}

function showBlockErrors(errors) {
  let msg = document.getElementById("blockErrors");
  if (!msg) {
    msg = document.createElement("div");
    msg.id = "blockErrors";
    msg.style = "padding:0.5em; font-family:monospace; font-size:0.9em;";
    document.body.insertBefore(msg, document.getElementById("app"));
  }
  if (errors.length > 0) {
    msg.innerHTML = "<strong>⚠️ Errori di struttura:</strong><br>" + errors.join("<br>");
    msg.style.color = "red";
  } else {
    msg.innerHTML = "<strong>✅ Nessun errore di struttura nei blocchi</strong>";
    msg.style.color = "green";
  }
}

editor.addEventListener("input", () => {
  saveModule(currentProject, currentModule, editor.innerText);
  highlight();
  updateProcedureList();
  checkBlockCoherence();
});

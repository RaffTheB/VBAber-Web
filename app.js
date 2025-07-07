
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

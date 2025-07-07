
const editor = document.getElementById("editor");
const highlighter = document.getElementById("highlighting");
const themeSelector = document.getElementById("themeSelector");
const projectSelector = document.getElementById("projectSelector");

const KEYWORDS = [
  "Sub", "End Sub", "Function", "End Function", "Dim", "Set", "If", "Then", "Else", "End If",
  "For", "Each", "In", "Next", "Do", "Loop", "While", "Wend", "With", "End With",
  "Select Case", "Case", "End Select", "Const", "Private", "Public", "ByVal", "ByRef", "Exit Sub",
  "Integer", "Long", "String", "Boolean", "Variant", "Object", "Nothing", "True", "False",
  "Call", "GoTo", "On Error", "Resume", "Redim", "Preserve", "Option Explicit"
];

let currentProject = "Progetto1";
let projects = JSON.parse(localStorage.getItem("vbaber-projects") || "{}");

function saveProject(name, content) {
  projects[name] = content;
  localStorage.setItem("vbaber-projects", JSON.stringify(projects));
}

function loadProject(name) {
  editor.innerText = projects[name] || "";
  currentProject = name;
  highlight();
}

function exportProject() {
  const blob = new Blob([editor.innerText], {type: "text/plain"});
  const link = document.createElement("a");
  link.href = URL.createObjectURL(blob);
  link.download = currentProject + ".bas";
  link.click();
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

editor.addEventListener("input", () => {
  saveProject(currentProject, editor.innerText);
  highlight();
});

function escapeHtml(str) {
  return str.replace(/&/g, "&amp;")
            .replace(/</g, "&lt;")
            .replace(/>/g, "&gt;");
}

function highlight() {
  let code = escapeHtml(editor.innerText);
  code = code.replace(/"(.*?)"/g, '<span class="token string">"$1"</span>');
  code = code.replace(/'(.*)/g, '<span class="token comment">'$1</span>');
  code = code.replace(/\b(\d+(\.\d+)?)\b/g, '<span class="token number">$1</span>');
  KEYWORDS.forEach(kw => {
    const pattern = "\b" + kw.replace(/ /g, "\s+") + "\b";
    const regex = new RegExp(pattern, "gi");
    code = code.replace(regex, `<span class="token keyword">${kw}</span>`);
  });
  highlighter.innerHTML = code + "<br/>";
}

themeSelector.addEventListener("change", () => {
  document.body.className = themeSelector.value;
  localStorage.setItem("vbaber-theme", themeSelector.value);
});

projectSelector.addEventListener("change", () => {
  saveProject(currentProject, editor.innerText);
  loadProject(projectSelector.value);
});

window.addEventListener("DOMContentLoaded", () => {
  const savedTheme = localStorage.getItem("vbaber-theme") || "light";
  themeSelector.value = savedTheme;
  document.body.className = savedTheme;
  if (!projects["Progetto1"]) projects["Progetto1"] = "";
  updateProjectSelector();
  loadProject(currentProject);
});

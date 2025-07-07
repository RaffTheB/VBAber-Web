
const editor = document.getElementById("editor");
const themeSelector = document.getElementById("themeSelector");
const projectSelector = document.getElementById("projectSelector");

const KEYWORDS = [
  "Sub", "End Sub", "Function", "End Function", "Dim", "Set", "If", "Then", "Else", "End If",
  "For", "Each", "In", "Next", "Do", "Loop", "While", "Wend", "With", "End With",
  "Select Case", "Case", "End Select", "Const", "Private", "Public", "ByVal", "ByRef", "Exit Sub"
];

let currentProject = "Progetto1";
let projects = JSON.parse(localStorage.getItem("vbaber-projects") || "{}");

function saveProject(name, content) {
  projects[name] = content;
  localStorage.setItem("vbaber-projects", JSON.stringify(projects));
}

function loadProject(name) {
  editor.value = projects[name] || "";
  currentProject = name;
  updateSuggestions();
}

function exportProject() {
  const blob = new Blob([editor.value], {type: "text/plain"});
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

editor.addEventListener("input", (e) => {
  saveProject(currentProject, editor.value);
  updateSuggestions();
});

editor.addEventListener("keydown", (e) => {
  const val = editor.value;
  const pos = editor.selectionStart;
  if (e.key === '"') {
    e.preventDefault();
    editor.value = val.slice(0, pos) + '""' + val.slice(pos);
    editor.selectionStart = editor.selectionEnd = pos + 1;
  } else if (e.key === "Enter") {
    const lineStart = val.lastIndexOf("\n", pos - 1) + 1;
    const currentLine = val.slice(lineStart, pos);
    if (/\bIf .* Then$/.test(currentLine)) {
      const indent = currentLine.match(/^\s*/)[0];
      const addition = `\n${indent}    \n${indent}End If`;
      editor.value = val.slice(0, pos) + addition + val.slice(pos);
      editor.selectionStart = editor.selectionEnd = pos + indent.length + 5;
      e.preventDefault();
    }
  } else if (e.key === "Tab" && suggestionBox.style.display !== "none") {
    e.preventDefault();
    applySuggestion();
  }
});

themeSelector.addEventListener("change", () => {
  document.body.className = themeSelector.value;
  localStorage.setItem("vbaber-theme", themeSelector.value);
});

projectSelector.addEventListener("change", () => {
  saveProject(currentProject, editor.value);
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

// Suggestion dropdown
const suggestionBox = document.createElement("div");
suggestionBox.id = "suggestionBox";
document.body.appendChild(suggestionBox);

function updateSuggestions() {
  const pos = editor.selectionStart;
  const text = editor.value.slice(0, pos);
  const lastWord = text.split(/\s+/).pop();
  if (lastWord.length < 2) {
    suggestionBox.style.display = "none";
    return;
  }
  const matches = KEYWORDS.filter(k => k.toLowerCase().startsWith(lastWord.toLowerCase()));
  if (matches.length === 0) {
    suggestionBox.style.display = "none";
    return;
  }
  suggestionBox.innerHTML = "";
  matches.forEach(match => {
    const div = document.createElement("div");
    div.className = "suggestion";
    div.textContent = match;
    div.onclick = () => applySuggestion(match);
    suggestionBox.appendChild(div);
  });
  const rect = editor.getBoundingClientRect();
  suggestionBox.style.display = "block";
  suggestionBox.style.left = rect.left + 10 + "px";
  suggestionBox.style.top = rect.top + editor.offsetHeight - 100 + "px";
}

function applySuggestion(word = null) {
  const pos = editor.selectionStart;
  const text = editor.value;
  const parts = text.slice(0, pos).split(/\s+/);
  const last = parts.pop();
  const rest = parts.join(" ");
  const insert = word || suggestionBox.firstChild.textContent;
  const newText = text.slice(0, pos - last.length) + insert + text.slice(pos);
  editor.value = newText;
  editor.selectionStart = editor.selectionEnd = pos - last.length + insert.length;
  saveProject(currentProject, editor.value);
  suggestionBox.style.display = "none";
}

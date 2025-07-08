
function switchTheme() {
  const theme = document.getElementById("themeSelector").value;
  document.body.className = theme;
}

function highlight() {
  const text = editor.innerText;
  const html = text
    .replace(/&/g, "&amp;")
    .replace(/</g, "&lt;")
    .replace(/>/g, "&gt;")
    .replace(/"[^"]*"/g, match => `<span class="string">${match}</span>`)
    .replace(/'.*/g, match => `<span class="comment">${match}</span>`)
    .replace(/\b(If|Then|Else|End|Sub|Function|Select|Case|For|Next|While|Wend|Do|Loop|With|Const|Dim|Set|As|New|Enum|Public|Private)\b/gi, match => `<span class="keyword">${match}</span>`);
  document.getElementById("highlighting-content").innerHTML = html;
}

editor.addEventListener("input", () => {
  highlight();
  showAutocomplete();
});

editor.addEventListener("keydown", function(e) {
  if (e.key === "Tab") {
    e.preventDefault();
    const sel = window.getSelection();
    const range = sel.getRangeAt(0);
    const tabNode = document.createTextNode("  ");
    range.insertNode(tabNode);
    range.setStartAfter(tabNode);
    range.setEndAfter(tabNode);
    sel.removeAllRanges();
    sel.addRange(range);
  }
});

const autocompleteWords = ["Dim", "If", "Then", "Else", "End", "Sub", "Function", "For", "Next", "Do", "Loop", "While", "Wend"];

function showAutocomplete() {
  const text = editor.innerText;
  const lastWord = text.split(/\s+/).pop().toLowerCase();
  const matches = autocompleteWords.filter(word => word.toLowerCase().startsWith(lastWord));

  const list = document.getElementById("autocomplete-list");
  list.innerHTML = "";
  if (matches.length === 0 || lastWord.length === 0) {
    list.style.display = "none";
    return;
  }

  matches.forEach(match => {
    const li = document.createElement("li");
    li.textContent = match;
    li.onclick = () => insertAutocomplete(match);
    list.appendChild(li);
  });

  list.style.display = "block";
}

function insertAutocomplete(word) {
  const text = editor.innerText;
  const words = text.split(/\s+/);
  words.pop();
  words.push(word);
  editor.innerText = words.join(" ") + " ";
  highlight();
  document.getElementById("autocomplete-list").style.display = "none";
}

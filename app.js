
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

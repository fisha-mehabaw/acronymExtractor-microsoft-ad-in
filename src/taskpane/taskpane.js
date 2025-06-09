Office.onReady(() => {
  document.getElementById("run").onclick = run;
});

function extractAcronymDefinitions(text) {
  const pairs = [];

  // Pattern 1: Definition (ACRONYM)
  const pattern1 = /\b([A-Z][a-z]+(?: [A-Z][a-z]+)*)\s+\((\b[A-Z]{2,}\b)\)/g;
  let match;
  while ((match = pattern1.exec(text)) !== null) {
    pairs.push({ acronym: match[2], definition: match[1] });
  }

  // Pattern 2: ACRONYM (Definition)
  const pattern2 = /\b([A-Z]{2,})\s+\(([^)]+)\)/g;
  while ((match = pattern2.exec(text)) !== null) {
    if (/^[A-Z][a-z]+(?: [A-Z][a-z]+)*$/.test(match[2])) {
      pairs.push({ acronym: match[1], definition: match[2] });
    }
  }

  // Pattern 3: ACRONYM – Definition
  const pattern3 = /\b([A-Z]{2,})\s*[-–—]\s*([A-Z][a-z]+(?: [A-Z][a-z]+)*)/g;
  while ((match = pattern3.exec(text)) !== null) {
    pairs.push({ acronym: match[1], definition: match[2] });
  }

  // Remove duplicates
  const unique = {};
  pairs.forEach(({ acronym, definition }) => {
    if (!unique[acronym]) {
      unique[acronym] = definition;
    }
  });

  return Object.entries(unique).map(([acronym, definition]) => ({ acronym, definition }));
}

async function run() {
  await Word.run(async (context) => {
    const body = context.document.body;
    body.load("text");
    await context.sync();

    const results = extractAcronymDefinitions(body.text);

    const output = document.getElementById("output");

    output.innerHTML = `
      <div style="display: flex; justify-content: space-between; align-items: center;">
        <h3 style="margin: 0;">Acronyms Detected</h3>
        <button id="copyButton" title="Copy to clipboard" style="border: none; background: none; cursor: pointer;">
          🗐 Copy
        </button>
      </div>
      <div id="acronymList">
        ${
          results.length
            ? `<ul>${results.map(r => `<li><b>${r.acronym}</b> – ${r.definition}</li>`).join("")}</ul>`
            : "No acronyms found."
        }
      </div>
    `;

    // Setup copy functionality
    setTimeout(() => {
  const copyButton = document.getElementById("copyButton");
  const acronymList = document.getElementById("acronymList");

  if (copyButton && acronymList) {
    copyButton.addEventListener("click", async () => {
      try {
        const text = acronymList.innerText || acronymList.textContent;

        if (navigator.clipboard && window.isSecureContext) {
          await navigator.clipboard.writeText(text);
        } else {
          const textarea = document.createElement("textarea");
          textarea.value = text;
          document.body.appendChild(textarea);
          textarea.select();
          document.execCommand("copy");
          document.body.removeChild(textarea);
        }

        const status = document.getElementById("status");
        if (status) {
          status.textContent = "✅ Acronym list copied!";
          setTimeout(() => (status.textContent = ""), 3000);
        }

      } catch (err) {
        console.error("Copy failed:", err);
        const status = document.getElementById("status");
        if (status) {
          status.textContent = "⚠️ Copy failed.";
          setTimeout(() => (status.textContent = ""), 3000);
        }
      }
    });
  }
}, 0);

  });
}

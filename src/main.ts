import "./style.css";

Office.onReady(() => {
  renderApp();
});

function renderApp() {
  const app = document.getElementById("app")!;

  app.innerHTML = `
    <div class="container">
      <header class="header">
        <div class="logo">GCA</div>
        <h1 class="title">BOL Summary</h1>
      </header>

      <main class="content">
        <div id="status" class="status idle">
          Attach a BOL PDF to see the summary.
        </div>

        <div id="summary" class="summary hidden">
          <div class="field">
            <span class="label">Pickup From</span>
            <span id="pickup" class="value">—</span>
          </div>
          <div class="field">
            <span class="label">Deliver To</span>
            <span id="deliver" class="value">—</span>
          </div>
          <div class="field">
            <span class="label">Dimensions</span>
            <span id="dimensions" class="value">—</span>
          </div>
          <div class="field">
            <span class="label">Weight</span>
            <span id="weight" class="value">—</span>
          </div>
        </div>
      </main>

      <footer class="footer">
        <button id="btn-read" class="btn">Read BOL</button>
      </footer>
    </div>
  `;

  document.getElementById("btn-read")!.addEventListener("click", readBOL);
}

// ─── Paso 2: Leer adjuntos ───────────────────────────────────────────────────

function readBOL() {
  setStatus("loading", "Reading attachments...");

  const item = Office.context.mailbox.item;

  if (!item) {
    setStatus("error", "No email is open.");
    return;
  }

  const attachments = item.attachments;

  if (!attachments || attachments.length === 0) {
    setStatus("idle", "No attachments found in this email.");
    return;
  }

  // Busca el primer PDF adjunto
  const bolAttachment = attachments.find(
    (att) =>
      att.attachmentType === Office.MailboxEnums.AttachmentType.File &&
      att.name.toLowerCase().endsWith(".pdf")
  );

  if (!bolAttachment) {
    setStatus("idle", "No PDF attachment found.");
    return;
  }

  setStatus("loading", `Found: ${bolAttachment.name}. Reading...`);

  // Descarga el contenido del adjunto como base64
  item.getAttachmentContentAsync(
    bolAttachment.id,
    (result) => {
      if (result.status === Office.AsyncResultStatus.Failed) {
        setStatus("error", `Could not read attachment: ${result.error.message}`);
        return;
      }

      const base64 = result.value.content;
      setStatus("ok", `PDF loaded (${bolAttachment.name}). Parsing...`);

      parsePDF(base64);
    }
  );
}

// Placeholder — lo implementamos en el Paso 3
function parsePDF(base64: string) {
  console.log("parsePDF called with base64 length:", base64.length);
  setStatus("ok", "PDF received. Parser coming in next step.");
}

// ─── Helpers ─────────────────────────────────────────────────────────────────

function setStatus(type: "idle" | "loading" | "error" | "ok", message: string) {
  const el = document.getElementById("status")!;
  el.className = `status ${type}`;
  el.textContent = message;
}

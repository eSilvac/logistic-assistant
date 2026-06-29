(function(){let e=document.createElement(`link`).relList;if(e&&e.supports&&e.supports(`modulepreload`))return;for(let e of document.querySelectorAll(`link[rel="modulepreload"]`))n(e);new MutationObserver(e=>{for(let t of e)if(t.type===`childList`)for(let e of t.addedNodes)e.tagName===`LINK`&&e.rel===`modulepreload`&&n(e)}).observe(document,{childList:!0,subtree:!0});function t(e){let t={};return e.integrity&&(t.integrity=e.integrity),e.referrerPolicy&&(t.referrerPolicy=e.referrerPolicy),e.crossOrigin===`use-credentials`?t.credentials=`include`:e.crossOrigin===`anonymous`?t.credentials=`omit`:t.credentials=`same-origin`,t}function n(e){if(e.ep)return;e.ep=!0;let n=t(e);fetch(e.href,n)}})(),Office.onReady(()=>{e()});function e(){let e=document.getElementById(`app`);e.innerHTML=`
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
  `,document.getElementById(`btn-read`).addEventListener(`click`,t)}function t(){r(`loading`,`Reading attachments...`);let e=Office.context.mailbox.item;if(!e){r(`error`,`No email is open.`);return}let t=e.attachments;if(!t||t.length===0){r(`idle`,`No attachments found in this email.`);return}let i=t.find(e=>e.attachmentType===Office.MailboxEnums.AttachmentType.File&&e.name.toLowerCase().endsWith(`.pdf`));if(!i){r(`idle`,`No PDF attachment found.`);return}r(`loading`,`Found: ${i.name}. Reading...`),e.getAttachmentContentAsync(i.id,e=>{if(e.status===Office.AsyncResultStatus.Failed){r(`error`,`Could not read attachment: ${e.error.message}`);return}let t=e.value.content;r(`ok`,`PDF loaded (${i.name}). Parsing...`),n(t)})}function n(e){console.log(`parsePDF called with base64 length:`,e.length),r(`ok`,`PDF received. Parser coming in next step.`)}function r(e,t){let n=document.getElementById(`status`);n.className=`status ${e}`,n.textContent=t}
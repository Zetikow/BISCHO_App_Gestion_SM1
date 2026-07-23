// ===================================================================
// SUPPORT — messages envoyés à la gestion du club (feuille "Support"),
// avec historique et réponse visible depuis l'appli.
// ===================================================================

let supportHistoryState = { loading: false, loaded: false, history: [] };
let adminSupportState = { loading: false, loaded: false, requests: [] };

// Reformate "dd/MM/yyyy HH:mm" (venant du serveur) en quelque chose de plus lisible, cohérent
// avec le reste de l'appli (ex: "8 juillet à 14h32").
function formatSupportDate(dateStr) {
  const m = String(dateStr || "").match(/(\d{2})\/(\d{2})\/(\d{4}) (\d{2}):(\d{2})/);
  if (!m) return dateStr || "";
  const d = new Date(+m[3], +m[2] - 1, +m[1], +m[4], +m[5]);
  const dateLabel = d.toLocaleDateString("fr-FR", { day: "numeric", month: "long" });
  return `${dateLabel} à ${m[4]}h${m[5]}`;
}

function renderSupportPage() {
  // Vue Admin : uniquement les demandes reçues, regroupées par personne — pas d'utilité pour
  // l'Admin d'avoir aussi son propre formulaire ou son propre historique sur cette page.
  if (hasRole("Admin")) {
    let html = `<button class="back-link" data-goto-page="home">← Retour à l'accueil</button>
    <div class="page-title">Support — Demandes reçues</div>
    <div class="page-sub">Une carte par personne ayant fait une demande.</div>`;

    if (adminSupportState.loading && !adminSupportState.loaded) {
      html += `<div class="card"><div class="muted">Chargement…</div></div>`;
      return html;
    }
    if (adminSupportState.requests.length === 0) {
      html += `<div class="card"><div class="muted">Aucune demande reçue pour le moment.</div></div>`;
      return html;
    }

    // Regroupe les demandes par personne (une carte par personne, messages du plus ancien au plus récent).
    const byPerson = {};
    const order = [];
    adminSupportState.requests.forEach(r => {
      if (!byPerson[r.nom]) { byPerson[r.nom] = []; order.push(r.nom); }
      byPerson[r.nom].push(r);
    });
    order.forEach(nom => {
      const msgs = byPerson[nom].slice().reverse(); // chronologique (l'état est déjà du plus récent au plus ancien)
      const lastUnanswered = [...msgs].reverse().find(m => !m.reponse);
      const editingId = window.__replyEditingId;

      html += `<div class="card">
        <div style="font-size:14px; font-weight:800; color:#fff; margin-bottom:8px;">${escapeHtml(nom)}</div>
        ${msgs.map(r => `
          <div style="margin-bottom:10px; padding-bottom:10px; border-bottom:1px solid rgba(255,255,255,0.06);">
            <div class="muted" style="font-size:10px; margin-bottom:4px;">${escapeHtml(formatSupportDate(r.date))}</div>
            <div style="font-size:12.5px; color:#e8e8ee; line-height:1.5;">${escapeHtml(r.message)}</div>
            ${(r.reponse && editingId !== r.id) ? `
              <div style="margin-top:8px; padding-top:8px; border-top:1px solid rgba(255,255,255,0.08);">
                <div style="font-size:9px; font-weight:800; text-transform:uppercase; color:#33d17a; margin-bottom:4px;">Ta réponse</div>
                <div style="font-size:12px; color:#e4e8f2; line-height:1.5;">${escapeHtml(r.reponse)}</div>
                <div class="expand-toggle" data-edit-reply="${escapeHtml(r.id)}" style="margin-top:5px;">Modifier la réponse</div>
              </div>
            ` : (r.id === (lastUnanswered ? lastUnanswered.id : (editingId || "")) ? `
              <textarea data-reply-input="${escapeHtml(r.id)}" rows="3" placeholder="Ta réponse..." style="margin-top:8px;">${escapeHtml(r.reponse || "")}</textarea>
              <button class="btn" style="margin-top:8px;" data-reply-submit="${escapeHtml(r.id)}">Envoyer la réponse</button>
            ` : (r.reponse ? "" : `<div class="muted" style="margin-top:6px; font-size:10.5px; font-style:italic;">En attente de réponse</div>`))}
          </div>
        `).join("")}
      </div>`;
    });
    return html;
  }

  // Vue standard (non-Admin) : formulaire + historique personnel.
  let html = `<button class="back-link" data-goto-page="home">← Retour à l'accueil</button>
  <div class="page-title">Support / une question ?</div>
  <div class="page-sub">Ton message part directement à l'équipe de gestion du club.</div>
  <div class="card">
    <label class="field-label">Ton message</label>
    <textarea id="support-message" rows="6" placeholder="Décris ta question ou ton souci..."></textarea>
    <button class="btn" id="support-submit" style="margin-top:10px;" ${window.__supportSending ? "disabled" : ""}>${window.__supportSending ? "Envoi en cours..." : "Envoyer"}</button>
    ${window.__supportSent ? `<div class="muted" style="margin-top:8px; color:#33d17a; font-weight:700;">✓ Message envoyé, merci !</div>` : ""}
  </div>`;

  html += `<div class="section-h">Mes demandes précédentes</div>`;
  if (supportHistoryState.loading && !supportHistoryState.loaded) {
    html += `<div class="card"><div class="muted">Chargement…</div></div>`;
  } else if (supportHistoryState.history.length === 0) {
    html += `<div class="card"><div class="muted">Aucune demande envoyée pour le moment.</div></div>`;
  } else {
    supportHistoryState.history.forEach(h => {
      html += `<div class="card">
        <div class="muted" style="font-size:10.5px; margin-bottom:6px;">${escapeHtml(formatSupportDate(h.date))}</div>
        <div style="font-size:12.5px; color:#e8e8ee; line-height:1.5;">${escapeHtml(h.message)}</div>
        ${h.reponse ? `<div style="margin-top:10px; padding-top:10px; border-top:1px solid rgba(255,255,255,0.08);">
          <div style="font-size:9.5px; font-weight:800; text-transform:uppercase; color:#33d17a; margin-bottom:5px;">Réponse du club</div>
          <div style="font-size:12.5px; color:#e4e8f2; line-height:1.5;">${escapeHtml(h.reponse)}</div>
        </div>` : `<div class="muted" style="margin-top:8px; font-size:10.5px; font-style:italic;">En attente de réponse</div>`}
      </div>`;
    });
  }
  return html;
}

// ===================== ACTIONS API =====================

async function fetchSupportHistory() {
  const cached = supportHistoryState.loaded;
  supportHistoryState.loading = true;
  if (!cached) render();
  try {
    const params = new URLSearchParams({ action: "getMySupportHistory", authNom: session.nom, authCode: session.code });
    const res = await fetch(`${GOOGLE_SCRIPT_URL}?${params.toString()}`);
    const data = await res.json();
    if (data.ok) supportHistoryState = { loading: false, loaded: true, history: data.history || [] };
    else supportHistoryState.loading = false;
  } catch (err) { supportHistoryState.loading = false; }
  render();
}

async function sendSupportMessageApi(message) {
  window.__supportSending = true;
  render();
  try {
    const params = new URLSearchParams({ action: "sendSupportMessage", message, authNom: session.nom, authCode: session.code });
    await fetch(`${GOOGLE_SCRIPT_URL}?${params.toString()}`);
    window.__supportSending = false;
    window.__supportSent = true;
    const msgField = document.getElementById("support-message");
    if (msgField) msgField.value = ""; // réinitialise le champ après envoi
    fetchSupportHistory(); // recharge la liste pour y voir la nouvelle demande
  } catch (err) {
    window.__supportSending = false;
    isOnline = false;
    render();
  }
}

async function fetchAdminSupportRequests() {
  const cached = adminSupportState.loaded;
  adminSupportState.loading = true;
  if (!cached) render();
  try {
    const params = new URLSearchParams({ action: "getAllSupportRequests", authNom: session.nom, authCode: session.code });
    const res = await fetch(`${GOOGLE_SCRIPT_URL}?${params.toString()}`);
    const data = await res.json();
    if (data.ok) adminSupportState = { loading: false, loaded: true, requests: data.requests || [] };
    else adminSupportState.loading = false;
  } catch (err) { adminSupportState.loading = false; }
  render();
}

async function replySupportMessageApi(id, reponse) {
  try {
    const params = new URLSearchParams({ action: "replySupportMessage", id, reponse, authNom: session.nom, authCode: session.code });
    await fetch(`${GOOGLE_SCRIPT_URL}?${params.toString()}`);
    window.__replyEditingId = null;
    fetchAdminSupportRequests();
  } catch (err) { isOnline = false; render(); }
}

function attachSupportEvents() {
  const supportSubmit = document.getElementById("support-submit");
  if (supportSubmit) supportSubmit.onclick = () => {
    const msg = (document.getElementById("support-message").value || "").trim();
    if (!msg) { alert("Merci d'écrire un message avant d'envoyer."); return; }
    vibrate();
    sendSupportMessageApi(msg);
  };

  document.querySelectorAll("[data-edit-reply]").forEach(el => {
    el.onclick = () => { vibrate(); window.__replyEditingId = el.dataset.editReply; render(); };
  });

  document.querySelectorAll("[data-reply-submit]").forEach(el => {
    el.onclick = () => {
      const id = el.dataset.replySubmit;
      const textarea = document.querySelector(`[data-reply-input="${id}"]`);
      const reponse = (textarea ? textarea.value : "").trim();
      if (!reponse) { alert("Écris une réponse avant d'envoyer."); return; }
      vibrate();
      replySupportMessageApi(id, reponse);
    };
  });
}

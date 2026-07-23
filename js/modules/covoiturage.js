// ===================================================================
// COVOITURAGE — pour les matchs à l'extérieur (feuille "Covoiturage").
// ===================================================================

function covoitEntryFor(eventId, nom) {
  return covoiturage.find(r => r[0] === eventId && r[1] === nom) || null;
}

async function setCovoiturageApi(nom, eventId, jeConduit, places, besoinPlace) {
  // Mise à jour optimiste locale
  const existing = covoitEntryFor(eventId, nom);
  if (existing) {
    existing[2] = jeConduit; existing[3] = places; existing[4] = besoinPlace;
  } else {
    covoiturage.push([eventId, nom, jeConduit, places, besoinPlace]);
  }
  render();
  try {
    const params = new URLSearchParams({ action: "setCovoiturage", nom, eventId, jeConduit, places, besoinPlace: besoinPlace, authNom: session.nom, authCode: session.code });
    await fetch(`${GOOGLE_SCRIPT_URL}?${params.toString()}`);
  } catch (err) { isOnline = false; render(); }
}

function renderCovoiturageePage() {
  const teams = myCarpoolTeams();
  if (teams.length === 0) {
    return `<div class="page-title">Covoiturage</div><div class="card"><div class="muted">Aucune équipe concernée pour ce compte.</div></div>`;
  }
  const activeTeam = (window.__covoitTeamView && teams.includes(window.__covoitTeamView)) ? window.__covoitTeamView : teams[0];

  const now = new Date();
  const matches = evenements.filter(ev => typeClass(ev[3]) === "match" && eventEquipe(ev) === activeTeam && !isHomeMatch(ev[5]) && eventDateObj(ev) >= now)
    .sort((a, b) => eventDateObj(a) - eventDateObj(b));

  let html = `<div class="page-title">Covoiturage</div><div class="page-sub">Matchs à l'extérieur — équipe ${escapeHtml(activeTeam)}</div>`;
  html += renderTeamSwitcher(teams, activeTeam, "covoit-team");

  if (matches.length === 0) {
    html += `<div class="card"><div class="muted">Aucun match à l'extérieur à venir pour cette équipe.</div></div>`;
    return html;
  }

  const identities = myCarpoolIdentitiesForTeam(activeTeam);

  matches.forEach(ev => {
    const [id, date, heure, , titre, lieu] = ev;
    const d = eventDateObj(ev);
    const dateLabel = d.toLocaleDateString("fr-FR", { day: "2-digit", month: "short" }).replace(".", "").toUpperCase();
    const entries = covoiturage.filter(r => r[0] === id);
    const drivers = entries.filter(r => r[2] === "Oui");
    const needers = entries.filter(r => r[4] === "Oui");
    const totalPlaces = drivers.reduce((s, r) => s + (parseInt(r[3], 10) || 0), 0);

    html += `<div class="cp-match-card">
      <div class="cp-match-head">
        <div><div class="cp-match-title">${escapeHtml(titre || "Match")}</div><div class="cp-match-sub">${dateLabel} · ${formatHeure(ev) || ""} · ${escapeHtml(lieu || "")}</div></div>
        <div style="display:flex; gap:6px;">
          <div class="cp-summary-badge"><div class="num" style="color:#33d17a;">${totalPlaces}</div><div class="lbl">Places</div></div>
          <div class="cp-summary-badge"><div class="num" style="color:#ffb43c;">${needers.length}</div><div class="lbl">Demandes</div></div>
        </div>
      </div>
      <div class="cp-cols">
        <div class="cp-col">
          <div class="cp-col-h driver">🚗 Conducteurs</div>
          ${drivers.length === 0 ? `<div class="cp-empty">Personne pour l'instant</div>` : drivers.map(r => `<div class="cp-row"><span>${escapeHtml(r[1])}</span><span class="places">${escapeHtml(r[3] || "?")} pl.</span></div>`).join("")}
        </div>
        <div class="cp-col">
          <div class="cp-col-h need">🙋 Cherchent une place</div>
          ${needers.length === 0 ? `<div class="cp-empty">Personne pour l'instant</div>` : needers.map(r => `<div class="cp-row"><span>${escapeHtml(r[1])}</span></div>`).join("")}
        </div>
      </div>`;

    if (identities.length === 0) {
      html += `<div class="muted" style="font-size:9.5px; margin-top:10px; text-align:center;">Seul ton parent peut modifier cette page pour toi.</div>`;
    } else {
      identities.forEach(idt => {
        const entry = covoitEntryFor(id, idt.nom);
        const jeConduit = entry ? entry[2] : "";
        const places = entry ? entry[3] : "3";
        const besoinPlace = entry ? entry[4] : "";
        html += `<div class="cp-edit-box">
          <div class="cp-edit-label">${idt.isChild ? `Pour ${escapeHtml(idt.nom)} <span class="cp-for-child">ton enfant</span>` : "Toi"}</div>
          <div class="cp-toggle-row">
            <button type="button" class="cp-toggle-btn ${jeConduit === "Oui" ? "active-yes" : ""}" data-cp-conduit="${escapeHtml(id)}|||${escapeHtml(idt.nom)}">Je conduis</button>
            <button type="button" class="cp-toggle-btn ${besoinPlace === "Oui" ? "active-need" : ""}" data-cp-besoin="${escapeHtml(id)}|||${escapeHtml(idt.nom)}">J'ai besoin d'une place</button>
          </div>
          ${jeConduit === "Oui" ? `<div class="cp-edit-label">Nombre de places disponibles</div>
          <select data-cp-places="${escapeHtml(id)}|||${escapeHtml(idt.nom)}">
            ${[1,2,3,4,5].map(n => `<option value="${n}" ${String(places) === String(n) ? "selected" : ""}>${n}</option>`).join("")}
          </select>` : ""}
        </div>`;
      });
    }

    html += `</div>`;
  });

  return html;
}

// Historique du covoiturage (matchs passés) pour une personne donnée — utilisé notamment sur
// le Profil des parents, pour voir ce qui a été renseigné pour leur enfant au fil de la saison.
function renderCovoiturageHistoryCard(nom) {
  const now = new Date();
  const entries = covoiturage.filter(r => r[1] === nom).map(r => {
    const ev = evenements.find(e => e[0] === r[0]);
    return ev ? { ev, jeConduit: r[2], besoinPlace: r[4] } : null;
  }).filter(Boolean).filter(x => eventDateObj(x.ev) < now).sort((a, b) => eventDateObj(b.ev) - eventDateObj(a.ev));

  let html = `<div class="card"><div class="section-h" style="margin-top:0;">Historique covoiturage</div>`;
  if (entries.length === 0) {
    html += `<div class="muted">Aucun historique pour le moment.</div>`;
  } else {
    entries.slice(0, 8).forEach(({ ev, jeConduit, besoinPlace }) => {
      const d = eventDateObj(ev);
      const dateLabel = d.toLocaleDateString("fr-FR", { day: "2-digit", month: "short" }).replace(".", "").toUpperCase();
      const statut = jeConduit === "Oui" ? "🚗 A conduit" : (besoinPlace === "Oui" ? "🙋 A eu besoin d'une place" : "—");
      html += `<div class="paiement-row"><div>${dateLabel} — ${escapeHtml(ev[4] || "Match")}</div><div class="muted" style="font-size:11px;">${statut}</div></div>`;
    });
  }
  html += `</div>`;
  return html;
}

function attachCovoiturageEvents() {
  document.querySelectorAll("[data-covoit-team]").forEach(el => {
    el.onclick = () => { vibrate(); window.__covoitTeamView = el.dataset.covoitTeam; render(); };
  });

  document.querySelectorAll("[data-cp-conduit]").forEach(el => {
    el.onclick = () => {
      vibrate();
      const [eventId, nom] = el.dataset.cpConduit.split("|||");
      const entry = covoitEntryFor(eventId, nom);
      const newVal = (entry && entry[2] === "Oui") ? "" : "Oui";
      const places = (entry && entry[3]) || "3";
      const besoin = newVal === "Oui" ? "" : (entry ? entry[4] : "");
      setCovoiturageApi(nom, eventId, newVal, places, besoin);
    };
  });

  document.querySelectorAll("[data-cp-besoin]").forEach(el => {
    el.onclick = () => {
      vibrate();
      const [eventId, nom] = el.dataset.cpBesoin.split("|||");
      const entry = covoitEntryFor(eventId, nom);
      const newVal = (entry && entry[4] === "Oui") ? "" : "Oui";
      const jeConduit = newVal === "Oui" ? "" : (entry ? entry[2] : "");
      const places = (entry && entry[3]) || "";
      setCovoiturageApi(nom, eventId, jeConduit, places, newVal);
    };
  });

  document.querySelectorAll("[data-cp-places]").forEach(el => {
    el.onchange = () => {
      const [eventId, nom] = el.dataset.cpPlaces.split("|||");
      const entry = covoitEntryFor(eventId, nom);
      setCovoiturageApi(nom, eventId, "Oui", el.value, entry ? entry[4] : "");
    };
  });
}

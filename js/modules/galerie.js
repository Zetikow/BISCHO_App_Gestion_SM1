// ===================================================================
// GALERIE PHOTOS — photos de match hébergées sur Google Drive
// (voir apps-script/Photos.gs).
// ===================================================================

let photoViewState = { loading: false, loaded: false, error: null, photos: [], folderName: "" };
const photoViewCache = {};

async function fetchPhotosForMatch(equipe, date, titre) {
  const cacheKey = `${equipe}__${date}__${titre}`;
  const cached = photoViewCache[cacheKey];
  if (cached) {
    photoViewState = { loading: false, loaded: true, error: null, photos: cached.photos, folderName: cached.folderName };
    render();
  } else {
    photoViewState = { loading: true, loaded: false, error: null, photos: [], folderName: "" };
    render();
  }
  try {
    const params = new URLSearchParams({ action: "photosListForMatch", equipe, date, titre, authNom: session.nom, authCode: session.code });
    const res = await fetch(`${GOOGLE_SCRIPT_URL}?${params.toString()}`);
    const data = await res.json();
    if (data.ok) {
      photoViewCache[cacheKey] = { photos: data.photos, folderName: data.folderName };
      photoViewState = { loading: false, loaded: true, error: null, photos: data.photos, folderName: data.folderName };
    } else {
      photoViewState = { loading: false, loaded: true, error: "Impossible de charger les photos.", photos: [], folderName: "" };
    }
  } catch (err) {
    if (!cached) photoViewState = { loading: false, loaded: true, error: "Connexion impossible.", photos: [], folderName: "" };
  }
  render();
}

function goToGallery(equipe) {
  window.__galleryEquipe = equipe;
  currentPage = "gallery";
  render();
}

function goToPhotoView(equipe, date, titre) {
  window.__photoViewMatch = { equipe, date, titre };
  currentPage = "photoview";
  fetchPhotosForMatch(equipe, date, titre);
}

function renderGalleryPage() {
  const allowedTeams = myAllowedEquipes();
  const requested = window.__galleryEquipe || allowedTeams[0] || "SM1";
  const equipe = allowedTeams.includes(requested) ? requested : (allowedTeams[0] || "SM1");
  const matches = evenements.filter(ev => typeClass(ev[3]) === "match" && (eventEquipe(ev) === equipe || eventEquipe(ev) === "Toutes"))
    .sort((a, b) => eventDateObj(b) - eventDateObj(a));

  let html = `<button class="back-link" data-goto-page="home">← Retour à l'accueil</button>`;
  html += `<div class="page-title">Galerie photos</div><div class="page-sub">Équipe ${escapeHtml(equipe)}</div>`;
  html += renderTeamSwitcher(allowedTeams, equipe, "gallery-team");

  if (matches.length === 0) {
    html += `<div class="card"><div class="muted">Aucun match enregistré pour cette équipe pour le moment.</div></div>`;
    return html;
  }

  html += `<div class="section-h">Matchs de l'équipe</div>`;
  html += `<div class="matchcard-carousel-wrap">
    <button class="carousel-arrow left" data-carousel-scroll="left" aria-label="Précédent">‹</button>
    <div class="matchcard-carousel">`;
  matches.forEach(ev => {
    const [, date, , , titre, lieu] = ev;
    const d = eventDateObj(ev);
    const dateLabel = d.toLocaleDateString("fr-FR", { day: "2-digit", month: "short" }).replace(".", "").toUpperCase();
    const m = formatMatchDisplay(titre, lieu);
    html += `<div class="mc2-card" data-open-match-photos="${escapeHtml(date)}|||${escapeHtml(titre)}|||${escapeHtml(equipe)}">
      <div class="mc2-datebar">${dateLabel}</div>
      <div style="padding:16px 12px; text-align:center;">
        <div style="font-size:13px; font-weight:700; color:#fff; line-height:1.4;">${escapeHtml(m.label)}</div>
      </div>
      <div class="mc2-time">📷 Voir les photos</div>
    </div>`;
  });
  html += `</div>
    <button class="carousel-arrow right" data-carousel-scroll="right" aria-label="Suivant">›</button>
  </div>`;
  return html;
}

function renderPhotoViewPage() {
  const match = window.__photoViewMatch || {};
  const m = formatMatchDisplay(match.titre, "");
  let html = `<button class="back-link" data-goto-page="gallery">← Retour à la galerie</button>`;
  html += `<div class="page-title">${escapeHtml(m.label || match.titre || "Match")}</div><div class="page-sub">${escapeHtml(photoViewState.folderName || "")}</div>`;

  if (photoViewState.loading && !photoViewState.loaded) {
    html += `<div class="card"><div class="muted">Chargement des photos…</div></div>`;
    return html;
  }
  if (photoViewState.error) {
    html += `<div class="card"><div class="muted">${escapeHtml(photoViewState.error)}</div></div>`;
    return html;
  }
  if (photoViewState.photos.length === 0) {
    html += `<div class="card"><div class="muted">Aucune photo pour ce match pour le moment. Dépose-les directement dans le dossier Google Drive correspondant.</div></div>`;
    return html;
  }
  html += `<div class="photo-grid">`;
  photoViewState.photos.forEach(p => {
    html += `<a href="${escapeHtml(p.viewUrl)}" target="_blank" rel="noopener" class="photo-grid-item"><img src="${escapeHtml(p.imageUrl)}" alt="${escapeHtml(p.name)}" loading="lazy"/></a>`;
  });
  html += `</div>`;
  return html;
}

function attachGalerieEvents() {
  document.querySelectorAll("[data-goto-gallery]").forEach(el => {
    el.onclick = () => { vibrate(); goToGallery(el.dataset.gotoGallery); };
  });

  document.querySelectorAll("[data-open-match-photos]").forEach(el => {
    el.onclick = () => {
      vibrate();
      const [date, titre, equipe] = el.dataset.openMatchPhotos.split("|||");
      goToPhotoView(equipe, date, titre);
    };
  });

  document.querySelectorAll("[data-gallery-team]").forEach(el => {
    el.onclick = () => { vibrate(); window.__galleryEquipe = el.dataset.galleryTeam; render(); };
  });
}

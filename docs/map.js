/* ═══════════════════════════════════════════════════════════════
   All content comes from places.xlsx. Nothing in this file needs
   editing except the Street View key on the next line.

   Paste a Google Maps Embed API key here and every place gets a
   panorama built from its coordinates. Leave it empty and cards
   get a button that opens Street View in a new tab instead — or
   fill in the "embed" column per row to place your own.
   ═══════════════════════════════════════════════════════════════ */

const GOOGLE_KEY = "";

/* ─────────────────────────────────────────────────────────────── */

const story = document.getElementById("story");
const oops  = document.getElementById("oops");

let PLACES  = [];
let markers = [];
let current = -1;

/* ---------- map ---------- */

const map = L.map("map", { zoomControl: true }).setView([52.1, 5.1], 7);

L.tileLayer("https://{s}.basemaps.cartocdn.com/light_all/{z}/{x}/{y}{r}.png", {
  maxZoom: 19,
  attribution: "&copy; OpenStreetMap &middot; &copy; CARTO"
}).addTo(map);

/* ---------- text helpers ----------
   A spreadsheet cell holds plain text, so paragraph breaks and *italics*
   are turned into HTML here. Everything is escaped first, so a stray < or
   & in a story can't break the page. */

const escapeHtml = s =>
  String(s).replace(/[&<>]/g, c => ({ "&": "&amp;", "<": "&lt;", ">": "&gt;" }[c]));

const italics = s => s.replace(/\*([^*\n]+)\*/g, "<em>$1</em>");

const toParagraphs = s =>
  String(s || "")
    .split(/\r?\n/)
    .map(t => t.trim())
    .filter(Boolean)
    .map(t => `<p>${italics(escapeHtml(t))}</p>`)
    .join("");

const splitList = s =>
  String(s || "").split(",").map(t => t.trim()).filter(Boolean);

/* ---------- load the spreadsheet ----------
   Cells arrive as numbers or strings depending on what you typed, so
   everything gets coerced to a string before it's used. */

const cell = v => (v === undefined || v === null ? "" : String(v).trim());

/* Excel stores a typed 215 as the number 215.0, which Street View won't
   take. This drops the pointless decimal without touching real ones. */
const num = v => cell(v).replace(/\.0+$/, "");

async function load() {
  let rows;
  try {
    const res = await fetch("places.xlsx");
    if (!res.ok) throw new Error(res.status);
    const book  = XLSX.read(await res.arrayBuffer(), { type: "array" });
    const sheet = book.Sheets["places"] || book.Sheets[book.SheetNames[0]];
    rows = XLSX.utils.sheet_to_json(sheet, { defval: "" });
  } catch (err) {
    console.error(err);
    return fail(
      "Couldn't read places.xlsx. Run <code>quarto preview</code> rather than opening " +
      "the HTML file directly — the browser blocks local file reads otherwise."
    );
  }

  PLACES = rows.map((row, i) => {
    const lat = parseFloat(cell(row.lat));
    const lng = parseFloat(cell(row.lng));
    if (!isFinite(lat) || !isFinite(lng)) {
      console.warn(`Row ${i + 2} ("${cell(row.name) || "unnamed"}") has no usable lat/lng — skipped.`);
      return null;
    }
    return {
      name:    cell(row.name) || "Untitled",
      place:   cell(row.place),
      coords:  [lat, lng],
      text:    toParagraphs(cell(row.story)),
      photos:  splitList(cell(row.photos)),
      heading: num(row.heading),
      pitch:   num(row.pitch),
      embed:   cell(row.embed),
      pano:    cell(row.pano).toLowerCase() !== "false"
    };
  }).filter(Boolean);

  if (!PLACES.length) {
    return fail("No usable rows on the <strong>places</strong> sheet. Every row needs a lat and a lng.");
  }
  build();
}

load();

function fail(msg) {
  oops.innerHTML = `<div class="oops-box"><strong>Nothing to show</strong><p>${msg}</p></div>`;
  oops.hidden = false;
}

/* ---------- build ---------- */

function build() {
  map.fitBounds(L.latLngBounds(PLACES.map(p => p.coords)), { padding: [70, 70] });
  if (PLACES.length === 1) map.setView(PLACES[0].coords, 15);

  markers = PLACES.map((p, i) => {
    const m = L.marker(p.coords, {
      icon: L.divIcon({
        className: "",
        html:
          `<div class="pin-wrap" style="--i:${i}">` +
            `<div class="pin"><span class="pin-num">${i + 1}</span></div>` +
          `</div>`,
        iconSize:   [32, 42],
        iconAnchor: [16, 40]
      }),
      title: p.name,
      riseOnHover: true
    }).addTo(map);

    m.on("click", () => (current === i && !story.hidden ? closeCard() : openCard(i)));
    return m;
  });
}

const pinOf = i => markers[i].getElement().querySelector(".pin");

/* ---------- street view ---------- */

function panoEmbedSrc(p) {
  if (!p.pano) return "";
  if (p.embed) return p.embed;
  if (!GOOGLE_KEY) return "";
  const q = new URLSearchParams({
    key: GOOGLE_KEY,
    location: `${p.coords[0]},${p.coords[1]}`,
    fov: "90"
  });
  if (p.heading) q.set("heading", p.heading);
  if (p.pitch)   q.set("pitch",   p.pitch);
  return `https://www.google.com/maps/embed/v1/streetview?${q}`;
}

const panoLinkHref = p =>
  `https://www.google.com/maps/@?api=1&map_action=pano&viewpoint=${p.coords[0]},${p.coords[1]}`;

/* ---------- media frame ----------
   Photos and the panorama share one rectangle. Photos first, Street View
   last. The iframe is only created when its slide is opened. */

function buildMedia(p) {
  const frames = p.photos.map(f => ({ kind: "photo", src: `photos/${f}` }));
  const panoSrc = panoEmbedSrc(p);
  if (panoSrc) frames.push({ kind: "pano", src: panoSrc });
  if (!frames.length) return "";

  const slides = frames.map((f, n) =>
    f.kind === "photo"
      ? `<div class="media-frame${n === 0 ? " is-shown" : ""}">
           <img src="${f.src}" alt="" loading="lazy" onerror="this.closest('.media-frame').dataset.broken='1'">
         </div>`
      : `<div class="media-frame${n === 0 ? " is-shown" : ""}" data-pano-src="${f.src}"></div>`
  ).join("");

  const dots = frames.length > 1
    ? '<div class="media-dots">' + frames.map((f, n) =>
        `<button type="button" class="dot${f.kind === "pano" ? " is-pano" : ""}${n === 0 ? " is-on" : ""}"
                 data-n="${n}" aria-label="${f.kind === "pano" ? "Street View" : "Photo " + (n + 1)}">${
          f.kind === "pano" ? "360" : ""
        }</button>`).join("") + "</div>"
    : "";

  return `<div class="card-media">${slides}${dots}</div>`;
}

function wireMedia() {
  const media = story.querySelector(".card-media");
  if (!media) return;

  const frames = [...media.querySelectorAll(".media-frame")];
  const dots   = [...media.querySelectorAll(".dot")];

  const show = n => {
    frames.forEach((f, j) => f.classList.toggle("is-shown", j === n));
    dots.forEach((d, j) => d.classList.toggle("is-on", j === n));

    const f = frames[n];
    if (f.dataset.panoSrc && !f.firstElementChild) {
      const frame = document.createElement("iframe");
      frame.src = f.dataset.panoSrc;
      frame.loading = "lazy";
      frame.allowFullscreen = true;
      frame.referrerPolicy = "no-referrer-when-downgrade";
      frame.title = "Street View";
      f.appendChild(frame);
    }
  };

  dots.forEach(d => (d.onclick = () => show(Number(d.dataset.n))));
  show(0);
}

/* ---------- card ---------- */

function openCard(i) {
  const p = PLACES[i];
  current = i;

  markers.forEach((_, j) => pinOf(j).classList.toggle("is-active", j === i));
  pinOf(i).classList.add("is-read");

  const showLink = p.pano && !panoEmbedSrc(p);

  story.innerHTML =
    '<div class="card-bar"></div>' +
    buildMedia(p) +
    '<div class="card-body">' +
      (p.place ? `<p class="card-place">${escapeHtml(p.place)}</p>` : "") +
      `<h2 class="card-title"><span>${escapeHtml(p.name)}</span></h2>` +
      `<div class="card-text">${p.text}</div>` +
      (showLink
        ? `<a class="card-sv" href="${panoLinkHref(p)}" target="_blank" rel="noopener">Párate de nuevo en este lugar &rarr;</a>`
        : "") +
      '<div class="card-nav">' +
        '<button class="navbtn" data-go="-1" type="button">&larr;</button>' +
        `<span class="card-count">${i + 1} / ${PLACES.length}</span>` +
        '<button class="navbtn" data-go="1" type="button">&rarr;</button>' +
        '<button class="card-close" type="button">Close</button>' +
      "</div>" +
    "</div>";

  story.hidden = false;
  wireMedia();

  story.classList.remove("is-lit");
  void story.offsetWidth;
  story.classList.add("is-lit");

  story.querySelectorAll(".navbtn").forEach(b => {
    const target = i + Number(b.dataset.go);
    b.disabled = target < 0 || target >= PLACES.length;
    b.onclick = () => openCard(target);
  });
  story.querySelector(".card-close").onclick = closeCard;

  map.panTo(p.coords, { animate: true, duration: 0.6 });
}

function closeCard() {
  story.hidden = true;
  story.classList.remove("is-lit");
  markers.forEach((_, j) => pinOf(j).classList.remove("is-active"));
  current = -1;
}

document.addEventListener("keydown", e => {
  if (story.hidden) return;
  if (e.key === "Escape") closeCard();
  if (e.key === "ArrowRight" && current < PLACES.length - 1) openCard(current + 1);
  if (e.key === "ArrowLeft"  && current > 0)                 openCard(current - 1);
});
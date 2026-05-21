console.log("✅ app.js chargé");

/* =========================================================
   INIT
========================================================= */
document.addEventListener("DOMContentLoaded", () => {
  fileExcel.addEventListener("change", chargerExcel);
  runBtn.addEventListener("click", lancerOptimisation);
});

/* =========================================================
   CONSTANTES MÉTIER
========================================================= */

const TOLERANCE_HAUTEUR = 20; // cm autorisés en dépassement
const HAUTEURS = {
  "ETAGERE": 52,
  "1R": 83,
  "2R": 105,
  "3R": 118,
  "HAUT": 250,
  "VIDE": 0
};
const ZONES_BLOQUEES = {
  "M": [14, 15, 16],
  "N": [14, 15, 16]
};
function positionsParNiveau(travee) {
  return Number(travee) === 9 ? 3 : 4;
}
function getNiveau(index) {
  let result = "";
  let i = index;

  while (i >= 0) {
    result = String.fromCharCode((i % 26) + 65) + result;
    i = Math.floor(i / 26) - 1;
  }

  return result;
}
function blocDeTravee(travee) {
  const t = Number(travee);
  if (t >= 3 && t <= 9) return "BAS";
  if (t >= 10 && t <= 16) return "HAUT";
  return null;
}

/* =========================================================
   DONNÉES GLOBALES
========================================================= */

let emplacementsAvant = [];
let emplacementsApres = [];
let implantationAvant = null;
let implantationApres = null;


/* =========================================================
   IMPORT EXCEL
========================================================= */

function chargerExcel(e) {
  const file = e.target.files[0];
  if (!file) return;

  const reader = new FileReader();
  reader.onload = ev => {
    const wb = XLSX.read(ev.target.result, { type: "binary" });
    const rows = XLSX.utils.sheet_to_json(wb.Sheets["Emplacements"]);

    emplacementsAvant = rows
  .map(r => {

    const typeBrut = String(r["Type"] || "").trim().toUpperCase();

    return {
      allee: r["Allée"],
      travee: Number(r["Travée"]),
      niveau: String(r["Niveau"]),
      position: Number(r["Position"]),
      type: HAUTEURS[typeBrut] ? typeBrut : "VIDE"
    };

  })
  .filter(e => {
    return !(
      ZONES_BLOQUEES[e.allee] &&
      ZONES_BLOQUEES[e.allee].includes(e.travee)
    );
  });

console.log("Excel brut :", rows.length);
console.log("Après import :", emplacementsAvant.length);

    implantationAvant = compter(emplacementsAvant);
    afficherAvant();
  };

  reader.readAsBinaryString(file);
}

/* =========================================================
   CALCUL HAUTEUR D’UNE TRAVÉE
========================================================= */

function hauteurTravee(emps) {
  const niveauxVus = new Set();
  let h = 0;

  emps.forEach(e => {
    if (!niveauxVus.has(e.niveau)) {
      niveauxVus.add(e.niveau);
      h += HAUTEURS[e.type];
    }
  });

  return h;
}

/* =========================================================
   OPTIMISATION
========================================================= */

function lancerOptimisation() {

  const allowExtraLevels = document.getElementById("allowExtraLevels").checked;
  const HAUTEUR_MAX = Number(document.getElementById("hMax").value);

  if (!emplacementsAvant.length) {
    alert("Importer un Excel d'abord");
    return;
  }

  emplacementsApres = [];

  const structure = buildStructure(emplacementsAvant);
  const { objectifs, cibles } = computeTargets(structure);

  const ctx = {
    objectifs,
    cibles,
    HAUTEUR_MAX,
    allowExtraLevels,
    compteurs: { "1R":0,"2R":0,"3R":0,"HAUT":0 },
    compteursParAllee: {},
    pileAvant: buildPileAvant(emplacementsAvant)
  };

  const piles = buildPiles(structure, ctx);
  emplacementsApres = applyPiles(structure, piles);

  implantationApres = compter(emplacementsApres);

  const analyse = analyserLisses(emplacementsAvant, emplacementsApres);

  afficherApres();
  afficherComparaison();
  afficherParAllee(emplacementsApres);
  afficherHauteursParAllee(emplacementsApres);
  afficherPlanParAlleeDetaille(emplacementsApres);
  afficherBI(analyse);

  console.log("Mouvements de lisses :", calculerMouvementsLisses(emplacementsAvant, emplacementsApres));
}
function buildPileAvant(emps) {

  const map = {};

  emps.forEach(e => {
    const key = `${e.allee}_${e.travee}`;
    if (!map[key]) map[key] = {};

    if (!map[key][e.niveau]) {
      map[key][e.niveau] = e.type;
    }
  });

  const result = {};

  Object.entries(map).forEach(([key, niveaux]) => {
    const tri = Object.keys(niveaux)
      .sort((a,b) => a.localeCompare(b, undefined, {numeric:true}));

    result[key] = tri.map(n => niveaux[n]);
  });

  return result;
}
function buildStructure(emps) {

  const allees = [...new Set(emps.map(e => e.allee))];

  const structure = {};

  allees.forEach(a => {
    structure[a] = {
      BAS: [3,4,5,6,7,8,9],
      HAUT: [10,11,12,13,14,15,16]
    };
  });

  return structure;
}
function computeTargets(structure) {

  const objectifs = {
    "1R": Number(p1R.value) / 100,
    "2R": Number(p2R.value) / 100,
    "3R": Number(p3R.value) / 100,
    "HAUT": Number(pH.value) / 100
  };

  let total = 0;

  Object.values(structure).forEach(blocs => {
    Object.values(blocs).forEach(trs => {
      trs.forEach(tr => {
        total += positionsParNiveau(tr) * 10;
      });
    });
  });

  const cibles = {};
  Object.keys(objectifs).forEach(t => {
    cibles[t] = Math.round(total * objectifs[t]);
  });

  return { objectifs, cibles };
}
function choisirType(ctx, hauteur, niveauIdx, allee, tr) {

  const { HAUTEUR_MAX } = ctx;

  const candidats = Object.keys(HAUTEURS)
    .filter(t => hauteur + HAUTEURS[t] <= HAUTEUR_MAX + TOLERANCE_HAUTEUR);

  if (!candidats.length) return null;

  const pileAvant = ctx.pileAvant[`${allee}_${tr}`] || [];
  const typeAvant = pileAvant[niveauIdx];

  let best = null;
  let bestScore = -9999;

  candidats.forEach(t => {

  if (t === "ETAGERE") return;
  if (ctx.compteurs[t] >= ctx.cibles[t] * 1.15) return;

  let score = 0;

  // =========================
  // 1. BESOIN (%)
  // =========================
  const ratio = ctx.compteurs[t] / (ctx.cibles[t] || 1);
  score += (1 - ratio) * 6;

  if (ratio < 0.5) score += 2;
  if (ratio > 1) score -= 2;

  // =========================
  // 2. GLOBAL
  // =========================
  const total =
    ctx.compteurs["1R"] +
    ctx.compteurs["2R"] +
    ctx.compteurs["3R"] +
    ctx.compteurs["HAUT"];

  const part = ctx.compteurs[t] / (total || 1);
  score -= Math.abs(part - ctx.objectifs[t]) * 1;

// =========================
// ✅ ÉQUILIBRAGE PAR ALLÉE CORRIGÉ
// =========================

const ratioAllee =
  ctx.compteursParAllee[allee][t] / (ctx.ciblesParAllee[t] || 1);

// ✅ prioriser ceux en retard
score += (1 - ratioAllee) * 4;

// ✅ pénalité si trop avancé
if (ratioAllee > 1) score -= 3;


  // =========================
  // 4. VARIÉTÉ
  // =========================
  if (typeAvant) {
    if (t === typeAvant) score += 0.1;
    else score += 0.2;
  }

  // =========================
  // ✅ 5. ANTI ZEBRA
  // =========================
  if (typeAvant && t !== typeAvant) {
    score -= 0.15;
  }

  // =========================
  // 6. MÉTIER
  // =========================
  if (t === "3R") score += 0.2;
  if (t === "2R") score += 0.1;
  if (t === "1R") score -= 0.05;

  if (score > bestScore) {
    bestScore = score;
    best = t;
  }

});


  return best;
}
function buildPiles(structure, ctx) {

  const piles = {};

const NB_ALLEES = Object.keys(structure).length;

ctx.ciblesParAllee = {};

Object.keys(ctx.cibles).forEach(t => {
  ctx.ciblesParAllee[t] = ctx.cibles[t] / NB_ALLEES;
});

  Object.entries(structure).forEach(([allee, blocs]) => {

    ctx.compteursParAllee[allee] = { "1R":0,"2R":0,"3R":0,"HAUT":0 };

    Object.entries(blocs).forEach(([bloc, travees]) => {

      const key = `${allee}_${bloc}`;

      let pile = [];
      let hauteur = 0;

      while (true) {

        if (!ctx.allowExtraLevels && pile.length >= 10) break;

        const type = choisirType(ctx, hauteur, pile.length, allee, travees[0]);

        if (!type) break;

        pile.push(type);
        hauteur += HAUTEURS[type];

        const nb = positionsParNiveau(travees[0]);

        ctx.compteurs[type] += nb;
        ctx.compteursParAllee[allee][type] += nb;

        if (hauteur >= ctx.HAUTEUR_MAX) break;
      }

      piles[key] = pile;

    });

  });

  return piles;
}
function applyPiles(structure, piles) {

  const result = [];

  Object.entries(structure).forEach(([allee, blocs]) => {

    Object.entries(blocs).forEach(([bloc, travees]) => {

      const pile = piles[`${allee}_${bloc}`];

      travees.forEach(tr => {

  const zoneBloquee =
    ZONES_BLOQUEES[allee] &&
    ZONES_BLOQUEES[allee].includes(tr);

  if (zoneBloquee) {

    emplacementsAvant
      .filter(e => e.allee === allee && e.travee === tr)
      .forEach(e => result.push({ ...e }));

    return;
  }

        pile.forEach((type, idx) => {

          const niveau = getNiveau(idx); 
          const posMax = positionsParNiveau(tr);
          for (let p = 1; p <= posMax; p++) {
            result.push({ allee, travee: tr, niveau, position: p, type });
          }

        });

      });

    });

  });

  return result;
}

// ✅ Détecter incohérence dans l'existant
function estNiveauIncoherent(allee, tr, niveau) {

  const types = emplacementsAvant
    .filter(e =>
      e.allee === allee &&
      e.travee === tr &&
      e.niveau === niveau
    )
    .map(e => e.type);

  return new Set(types).size > 1;
}
function tauxHomogeneBloc(allee, tr) {

  const bloc = blocDeTravee(tr);

  const niveaux = [...new Set(
    emplacementsAvant
      .filter(e =>
        e.allee === allee &&
        blocDeTravee(e.travee) === bloc
      )
      .map(e => e.niveau)
  )];

  let nbHomogenes = 0;

  niveaux.forEach(niveau => {

    const types = emplacementsAvant
      .filter(e =>
        e.allee === allee &&
        blocDeTravee(e.travee) === bloc &&
        e.niveau === niveau
      )
      .map(e => e.type);

    if (new Set(types).size === 1) {
      nbHomogenes++;
    }

  });

  return nbHomogenes / niveaux.length; // 0 → 1
}


/* =========================================================
   COMPTAGE / AFFICHAGE
========================================================= */

function compter(list) {
  const r = { total: 0, repartition: {} };
  list.forEach(e => {
    r.total++;
    r.repartition[e.type] = (r.repartition[e.type] || 0) + 1;
  });
  return r;
}
function calculerMouvementsLisses(avant, apres) {

  const mapAvant = {};
  const mapApres = {};

  // 🔹 index AVANT
  avant.forEach(e => {
    const key = `${e.allee}_${e.travee}_${e.niveau}`;
    if (!mapAvant[key]) mapAvant[key] = e.type;
  });

  // 🔹 index APRES
  apres.forEach(e => {
    const key = `${e.allee}_${e.travee}_${e.niveau}`;
    if (!mapApres[key]) mapApres[key] = e.type;
  });

  let mouvements = 0;

  const travees = new Set(
    [...avant, ...apres].map(e => `${e.allee}_${e.travee}`)
  );

  travees.forEach(trKey => {

    let niveaux = [];

    Object.keys(mapApres).forEach(k => {
      if (k.startsWith(trKey)) {
        niveaux.push(k.split("_")[2]);
      }
    });

    // ✅ tri correct des niveaux (A → Z → AA)
    niveaux.sort((a, b) => a.localeCompare(b, undefined, { numeric: true }));

    for (let i = 0; i < niveaux.length - 1; i++) {

      const n1 = niveaux[i];
      const n2 = niveaux[i + 1];

      const key1 = `${trKey}_${n1}`;
      const key2 = `${trKey}_${n2}`;

      const avant1 = mapAvant[key1];
      const avant2 = mapAvant[key2];

      const apres1 = mapApres[key1];
      const apres2 = mapApres[key2];

      if (avant1 !== apres1 || avant2 !== apres2) {
        mouvements++;
      }

    }

  });

  return mouvements;
}
function analyserLisses(avant, apres) {

  const mapAvant = {};
  const mapApres = {};

  // index AVANT
  avant.forEach(e => {
    const key = `${e.allee}_${e.travee}_${e.niveau}`;
    if (!mapAvant[key]) mapAvant[key] = e.type;
  });

  // index APRES
  apres.forEach(e => {
    const key = `${e.allee}_${e.travee}_${e.niveau}`;
    if (!mapApres[key]) mapApres[key] = e.type;
  });

  const resultat = {};

  const travees = new Set(
    [...avant, ...apres].map(e => `${e.allee}_${e.travee}`)
  );

  let totalMove = 0;
  let totalNew = 0;

  travees.forEach(trKey => {

    let niveaux = [];

    Object.keys(mapApres).forEach(k => {
      if (k.startsWith(trKey)) {
        niveaux.push(k.split("_")[2]);
      }
    });

    niveaux.sort((a, b) => a.localeCompare(b, undefined, { numeric: true }));

    let move = 0;
    let added = 0;

    for (let i = 0; i < niveaux.length - 1; i++) {

      const n1 = niveaux[i];
      const n2 = niveaux[i + 1];

      const key1 = `${trKey}_${n1}`;
      const key2 = `${trKey}_${n2}`;

      const avant1 = mapAvant[key1];
      const avant2 = mapAvant[key2];

      const apres1 = mapApres[key1];
      const apres2 = mapApres[key2];

      // 🆕 nouvelle lisse (niveau inexistant avant)
      if (!avant1 || !avant2) {
        added++;
        totalNew++;
      }

      // 🔧 lisse modifiée
      else if (avant1 !== apres1 || avant2 !== apres2) {
        move++;
        totalMove++;
      }

    }

    const [allee, travee] = trKey.split("_");

    if (!resultat[allee]) {
      resultat[allee] = {
        move: 0,
        added: 0
      };
    }

    resultat[allee].move += move;
    resultat[allee].added += added;

  });

  return {
    totalMove,
    totalNew,
    detail: resultat
  };
}
function afficherAvant() {
  avantResult.innerHTML =
    `<b>Total :</b> ${implantationAvant.total}<br>` +
    Object.entries(implantationAvant.repartition)
      .map(([t,n]) => `${t} : ${n}`)
      .join("<br>");
}

function afficherApres() {

const mode = "BALANCE";

  const mouvements = calculerMouvementsLisses(
    emplacementsAvant,
    emplacementsApres
  );

  apresResult.innerHTML =
    `<b>Mode :</b> ${mode}<br>` +
    `<b>Total :</b> ${implantationApres.total}<br>` +
    `<b>Mouvements de lisses :</b> ${mouvements}<br><br>` +
    Object.entries(implantationApres.repartition)
      .map(([t,n]) => `${t} : ${n} (${((n/implantationApres.total)*100).toFixed(1)}%)`)
      .join("<br>");
}
function afficherComparaison() {
  const tbody = compareTable.querySelector("tbody");
  tbody.innerHTML = "";

  Object.keys(HAUTEURS).forEach(t => {
    tbody.innerHTML += `
      <tr>
        <td>GLOBAL</td>
        <td>${t}</td>
        <td>${implantationAvant.repartition[t] || 0}</td>
        <td>${implantationApres.repartition[t] || 0}</td>
        <td>${(implantationApres.repartition[t] || 0) -
              (implantationAvant.repartition[t] || 0)}</td>
      </tr>`;
  });
}

function afficherParAllee(emps) {
  const tbody = document.querySelector("#alleeTable tbody");
  tbody.innerHTML = "";

  const data = {};
  emps.forEach(e => {
    if (!data[e.allee])
      data[e.allee] = { "1R":0,"2R":0,"3R":0,"HAUT":0,total:0 };
    data[e.allee][e.type]++;
    data[e.allee].total++;
  });

  Object.entries(data).forEach(([allee,v]) => {
    tbody.innerHTML += `
      <tr>
        <td>${allee}</td>
        <td>${v["1R"]}</td>
        <td>${v["2R"]}</td>
        <td>${v["3R"]}</td>
        <td>${v["HAUT"]}</td>
        <td>${v.total}</td>
        <td>${(v.total * 0.96).toFixed(1)}</td>
      </tr>`;
  });
}
function afficherHauteursParAllee(emplacements) {

  const HAUTEUR_MAX = Number(document.getElementById("hMax").value); // ✅ AJOUT

  const tbody = document.querySelector("#hauteurTable tbody");
  if (!tbody) return;

  tbody.innerHTML = "";

  const niveauxVus = new Set();
  const hauteursParTravee = {};

  // 1️⃣ Calcul hauteur par travée
  emplacements.forEach(e => {
    const keyNiveau = `${e.allee}_${e.travee}_${e.niveau}`;
    if (niveauxVus.has(keyNiveau)) return;
    niveauxVus.add(keyNiveau);

    const keyTrav = `${e.allee}_${e.travee}`;
    if (!hauteursParTravee[keyTrav]) {
      hauteursParTravee[keyTrav] = 0;
    }
    hauteursParTravee[keyTrav] += HAUTEURS[e.type];
  });

  // 2️⃣ Hauteur MAX par allée (indicateur valide)
  const hauteursParAllee = {};

  Object.entries(hauteursParTravee).forEach(([key, h]) => {
    const [allee] = key.split("_");
    if (!hauteursParAllee[allee] || h > hauteursParAllee[allee]) {
      hauteursParAllee[allee] = h;
    }
  });

  // 3️⃣ Affichage
  Object.entries(hauteursParAllee).forEach(([allee, h]) => {

  const HAUTEUR_MAX = Number(document.getElementById("hMax").value); // ✅ important
  const pct = ((h / HAUTEUR_MAX) * 100).toFixed(1);                  // ✅ création
  const marge = HAUTEUR_MAX - h;

  tbody.innerHTML += `
    <tr>
      <td>${allee}</td>
      <td>${h}</td>
      <td>${pct} %</td>
      <td>${marge}</td>
    </tr>
  `;
});
}
function hauteurMaxParAllee(emplacements, allee) {
  const niveauxVus = new Set();
  const hauteurParTravee = {};

  emplacements.forEach(e => {
    if (e.allee !== allee) return;

    const keyN = `${e.travee}_${e.niveau}`;
    if (niveauxVus.has(keyN)) return;
    niveauxVus.add(keyN);

    if (!hauteurParTravee[e.travee]) {
      hauteurParTravee[e.travee] = 0;
    }
    hauteurParTravee[e.travee] += HAUTEURS[e.type];
  });

  return Math.max(...Object.values(hauteurParTravee));
}
function afficherPlanParAlleeDetaille(emplacements) {

  const container = document.getElementById("planContainer");
  container.innerHTML = "";

  const data = {};

  // 🔹 Structuration
  emplacements.forEach(e => {
    if (!data[e.allee]) data[e.allee] = {};
    if (!data[e.allee][e.travee]) data[e.allee][e.travee] = {};

    if (!data[e.allee][e.travee][e.niveau]) {
  data[e.allee][e.travee][e.niveau] = [];
}

data[e.allee][e.travee][e.niveau].push(e.type);
  });

  // 🔹 Parcours des allées
  Object.entries(data).forEach(([allee, travees]) => {

    const div = document.createElement("div");
    div.innerHTML = `<h3>Allée ${allee}</h3>`;

    const table = document.createElement("table");
    table.border = "1";
    table.style.marginBottom = "25px";
    table.style.borderCollapse = "collapse";

    const thead = document.createElement("thead");
    const tbody = document.createElement("tbody");

    // 🔹 récupérer tous les niveaux possibles
    const niveauxSet = new Set();
    Object.values(travees).forEach(niv => {
      Object.keys(niv).forEach(n => niveauxSet.add(n));
    });

    const niveaux = Array.from(niveauxSet).sort();

    // 🔹 HEADER
    const headerRow = document.createElement("tr");

    headerRow.innerHTML =
      `<th>Travée</th>` +
      niveaux.map(n => `<th>${n}</th>`).join("") +
      `<th>1R</th><th>2R</th><th>3R</th><th>HAUT</th><th>Hauteur</th>`;

    thead.appendChild(headerRow);

    // 🔹 Lignes
    Object.entries(travees)
      .sort((a, b) => a[0] - b[0])
      .forEach(([travee, niveauxData]) => {

        const tr = document.createElement("tr");

        // stats
        const stats = { "1R":0,"2R":0,"3R":0,"HAUT":0 };
        let hauteur = 0;

        let rowHTML = `<td>${travee}</td>`;

        niveaux.forEach(n => {

  const types = niveauxData[n] || [];

  // ✅ comptage réel des positions
  let display = "VIDE";
  let dominantType = null;

  if (types.length > 0) {

    const counts = {};
    types.forEach(t => {
      counts[t] = (counts[t] || 0) + 1;
    });

    // affichage type : 3R:8 HAUT:8
    display = Object.entries(counts)
  .map(([t,n]) => `${t}:${n}`)
  .join("\n");

    // ✅ choisir un type dominant (pour stats + couleur)
    dominantType = Object.keys(counts)
      .reduce((a,b) => counts[a] > counts[b] ? a : b);

    stats[dominantType] += types.length;
    hauteur += HAUTEURS[dominantType];
  }

  let color = "";
  if (dominantType === "1R") color = "#A8D5A2";
  if (dominantType === "2R") color = "#FFD966";
  if (dominantType === "3R") color = "#F4B183";
  if (dominantType === "HAUT") color = "#9BC2E6";
  if (dominantType === "ETAGERE") color = "#00B050";
  if (types.length === 0) color = "#EEEEEE";

  rowHTML += `<td style="background:${color}">${display}</td>`;
});

        rowHTML += `
          <td>${stats["1R"]}</td>
          <td>${stats["2R"]}</td>
          <td>${stats["3R"]}</td>
          <td>${stats["HAUT"]}</td>
          <td>${hauteur}</td>
        `;

        tr.innerHTML = rowHTML;
        tbody.appendChild(tr);
      });

    table.appendChild(thead);
    table.appendChild(tbody);
    div.appendChild(table);

    container.appendChild(div);
  });
}
function afficherBI(analyse) {

  let html = `<h3>Analyse des lisses</h3>`;

  html += `
    <b>Lisses déplacées :</b> ${analyse.totalMove}<br>
    <b>Lisses ajoutées :</b> ${analyse.totalNew}<br><br>
  `;

  html += `
    <table border="1" style="border-collapse:collapse">
      <tr>
        <th>Allée</th>
        <th>Déplacements</th>
        <th>Ajouts</th>
      </tr>
  `;

  Object.entries(analyse.detail).forEach(([allee, d]) => {
    html += `
      <tr>
        <td>${allee}</td>
        <td>${d.move}</td>
        <td>${d.added}</td>
      </tr>
    `;
  });

  html += "</table>";

  document.getElementById("apresResult").innerHTML += "<br>" + html;
}
function estBlocIncoherent(allee, tr) {

  const bloc = blocDeTravee(tr);

  let incoherent = false;

  const niveaux = [...new Set(
    emplacementsAvant
      .filter(e =>
        e.allee === allee &&
        blocDeTravee(e.travee) === bloc
      )
      .map(e => e.niveau)
  )];

  niveaux.forEach(niveau => {

    const types = emplacementsAvant
      .filter(e =>
        e.allee === allee &&
        blocDeTravee(e.travee) === bloc &&
        e.niveau === niveau
      )
      .map(e => e.type);

    if (new Set(types).size > 1) {
      incoherent = true;
    }

  });

  return incoherent;
}

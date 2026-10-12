/* ===============================
  VARIABLES GLOBALES
  =============================== */
let driversList = []; // Liste complète des conducteurs (depuis Excel)
let teams = [];       // Liste des écuries avec leurs couleurs et historique (depuis Excel)
let races = [];       // Calendrier des courses (depuis Excel)
let sprintResults = [];
let seasonTeamStandings = [];
let currentYear = 2026; // Saison actuellement affichée
let availableYears = [];

const FUTURE_LINEUPS = {
  2027: [
    { teamName: 'McLaren', drivers: ['Lando Norris', 'Oscar Piastri'] },
    { teamName: 'Red Bull Racing', drivers: ['Max Verstappen', 'Isack Hadjar'] },
    { teamName: 'Aston Martin', drivers: ['Fernando Alonso', 'Lance Stroll'] },
    { teamName: 'Haas F1 Team', drivers: [null, null] },
    { teamName: 'Williams', drivers: ['Alexander Albon', 'Carlos Sainz'] },
    { teamName: 'Cadillac', drivers: ['Valtteri Bottas', 'Sergio Perez'] },
    { teamName: 'Ferrari', drivers: ['Charles Leclerc', 'Lewis Hamilton'] },
    { teamName: 'Mercedes', drivers: ['Kimi Antonelli', 'George Russell'] },
    { teamName: 'Alpine', drivers: ['Pierre Gasly', 'Franco Colapinto'] },
    { teamName: 'Racing Bulls Formula One Team', drivers: [null, null] },
    { teamName: 'Audi', drivers: ['Nico Hulkenberg', 'Gabriel Bortoleto'] }
  ]
};

// Barème des points standard F1 pour les 10 premières positions
const BAREME_POINTS_F1 = [25, 18, 15, 12, 10, 8, 6, 4, 2, 1];

// État de l'interface
let dataLoaded = false;
let raceFilter = 'all';      // all | done | todo
let rankMode = 'drivers';    // drivers | teams
let driverFilter = 'all';    // all | active | past
let driverQuery = '';
let countdownTimer = null;

const $ = id => document.getElementById(id);

function flagCountryCode(flag) {
  const indicators = Array.from(flag, character => character.codePointAt(0))
    .filter(codePoint => codePoint >= 0x1f1e6 && codePoint <= 0x1f1ff);

  if (indicators.length === 2) {
    return indicators
      .map(codePoint => String.fromCharCode(codePoint - 0x1f1e6 + 65))
      .join('')
      .toLowerCase();
  }

  const countryCode = flag.trim();
  return /^[a-z]{2}$/i.test(countryCode) ? countryCode.toLowerCase() : '';
}

/* Codes pays ISO déduits des NOMS (plus fiable que les emojis, mal décodés par la lib Excel) */
const COUNTRY_CODES = {
  'allemagne': 'de', 'arabie saoudite': 'sa', 'australie': 'au', 'autriche': 'at', 'azerbaidjan': 'az',
  'bahrein': 'bh', 'belgique': 'be', 'bresil': 'br', 'canada': 'ca', 'chine': 'cn', 'coree du sud': 'kr',
  'emirats arabes unis': 'ae', 'espagne': 'es', 'etats-unis': 'us', 'france': 'fr', 'hongrie': 'hu',
  'inde': 'in', 'italie': 'it', 'japon': 'jp', 'malaisie': 'my', 'mexique': 'mx', 'monaco': 'mc',
  'pays-bas': 'nl', 'portugal': 'pt', 'qatar': 'qa', 'royaume-uni': 'gb', 'russie': 'ru',
  'singapour': 'sg', 'turquie': 'tr'
};
const NATIONALITY_CODES = {
  'allemand': 'de', 'americain': 'us', 'argentin': 'ar', 'australien': 'au', 'autrichien': 'at', 'belge': 'be',
  'britannique': 'gb', 'united kingdom': 'gb', 'bresilien': 'br', 'canadien': 'ca', 'chinois': 'cn',
  'colombien': 'co', 'danois': 'dk', 'espagnol': 'es', 'finlandais': 'fi', 'francais': 'fr', 'hongrois': 'hu',
  'indien': 'in', 'indonesien': 'id', 'irlandais': 'ie', 'italien': 'it', 'japonais': 'jp', 'malaisien': 'my',
  'mexicain': 'mx', 'monegasque': 'mc', 'neerlandais': 'nl', 'neo-zelandais': 'nz', 'polonais': 'pl',
  'portugais': 'pt', 'russe': 'ru', 'suisse': 'ch', 'suedois': 'se', 'tcheque': 'cz', 'thailandais': 'th',
  'venezuelien': 've'
};
function lookupCode(table, name) {
  const norm = String(name || '').normalize('NFD').replace(/[\u0300-\u036f]/g, '').trim().toLowerCase();
  if (!norm) return '';
  if (table[norm]) return table[norm];
  // Texte abîmé ("Bahre?n") : on traite chaque "?" comme un caractère joker
  if (norm.includes('?')) {
    const re = new RegExp('^' + norm.replace(/[.*+()\[\]\\^$|{}]/g, '\\$&').replace(/\?/g, '.') + '$');
    const key = Object.keys(table).find(k => re.test(k));
    if (key) return table[key];
  }
  return '';
}
const countryCodeFromName = name => lookupCode(COUNTRY_CODES, name);
const nationalityCode = nat => lookupCode(NATIONALITY_CODES, nat);

/* ===============================
  CHARGEMENT DES DONNÉES
  =============================== */
async function loadData() {
  const statusElement = $('loadingStatus');
  const retryButton = $('retryDataLoadBtn');
  const loaderView = $('data-loader-view');

  retryButton.classList.add('hidden');
  statusElement.classList.add('visually-hidden');
  statusElement.textContent = 'Chargement de la saison…';
  loaderView.setAttribute('aria-busy', 'true');

  try {
    const response = await fetch('datas.xlsx');
    if (!response.ok) {
      throw new Error(`Impossible de charger datas.xlsx (HTTP ${response.status}).`);
    }

    const data = new Uint8Array(await response.arrayBuffer());
    const workbook = XLSX.read(data, { type: 'array', cellDates: true, dateNF: 'yyyy-mm-dd' });
    processData(workbook);

    loaderView.setAttribute('aria-busy', 'false');
    loaderView.classList.add('hidden');
    initApp();
  } catch (error) {
    statusElement.textContent = `Erreur lors du chargement des données. Vérifiez que datas.xlsx est publié avec le site et contient les onglets requis : ${error.message}`;
    statusElement.classList.remove('visually-hidden');
    loaderView.setAttribute('aria-busy', 'false');
    retryButton.classList.remove('hidden');
    console.error(error);
  }
}

function refreshSeasonSelector() {
  const years = Array.from(new Set([
    ...races.map(r => raceYear(r)),
    ...teams.flatMap(t => t.history.map(h => h.year)),
    ...driversList.map(d => d.latestYear),
    ...Object.keys(FUTURE_LINEUPS).map(Number)
  ].filter(y => Number.isInteger(y) && y > 0))).sort((a, b) => b - a);

  availableYears = years.length ? years : [2026];
  if (!availableYears.includes(currentYear)) currentYear = availableYears[0];

  const seasonSelect = $('seasonSelect');
  if (!seasonSelect) return;

  seasonSelect.innerHTML = availableYears.map(y => `<option value="${y}">Saison ${y}</option>`).join('');
  seasonSelect.value = String(currentYear);
  const seasonOptions = $('seasonOptions');
  if (seasonOptions) {
    seasonOptions.innerHTML = availableYears
      .map(y => `<button type="button" class="season-option" role="option" data-season-year="${y}">Saison ${y}</button>`)
      .join('');
  }
  updateSeasonSelectButton();
}

function updateSeasonSelectButton() {
  const button = $('seasonSelectButton');
  const options = $('seasonOptions');
  if (!button || !options) return;

  button.textContent = `Saison ${currentYear}`;
  options.querySelectorAll('.season-option').forEach(option => {
    const selected = Number(option.dataset.seasonYear) === currentYear;
    option.setAttribute('aria-selected', String(selected));
    option.tabIndex = selected ? 0 : -1;
  });
}

function setSeasonMenuOpen(open, focusSelected = false) {
  const wrapper = document.querySelector('.season-select-wrap');
  const button = $('seasonSelectButton');
  const options = $('seasonOptions');
  if (!wrapper || !button || !options) return;

  wrapper.classList.toggle('is-open', open);
  options.classList.toggle('hidden', !open);
  button.setAttribute('aria-expanded', String(open));
  if (open && focusSelected) {
    const selected = options.querySelector('[aria-selected="true"]');
    selected?.focus();
    selected?.scrollIntoView({ block: 'nearest' });
  }
}

function processData(workbook) {
  // 1. Conducteurs
  const driversSheet = workbook.Sheets['Conducteurs'];
  if (!driversSheet) throw new Error("Onglet 'Conducteurs' manquant.");
  driversList = XLSX.utils.sheet_to_json(driversSheet);

  // 2. Écuries
  const teamsBaseSheet = workbook.Sheets['Écuries_Base'];
  const teamsHistorySheet = workbook.Sheets['Historique_Écuries'];
  if (!teamsBaseSheet || !teamsHistorySheet) throw new Error("Onglets 'Écuries_Base' ou 'Historique_Écuries' manquants.");

  const teamsBase = XLSX.utils.sheet_to_json(teamsBaseSheet);
  const teamsHistory = XLSX.utils.sheet_to_json(teamsHistorySheet);

  teams = teamsBase.map(baseTeam => {
    const teamHistory = teamsHistory
      .filter(h => h.name === baseTeam.name)
      .map(h => ({
        year: h.year,
        drivers: [h.driver1, h.driver2, h.driver3].filter(d => d),
        engine: h.moteur || '',
        pos: Number(h.pos_constructeurs) || null,
        points: h.points ?? null,
        wins: h.victoires ?? null,
        races: h.courses ?? null
      }))
      .sort((a, b) => b.year - a.year);

    return {
      ...baseTeam,
      color: String(baseTeam.color),
      color2: String(baseTeam.color2),
      history: teamHistory
    };
  });

  // Infos historiques des conducteurs
  const activeDriversInCurrentYear = teams.flatMap(t => t.history.find(h => h.year === currentYear)?.drivers || []);

  driversList = driversList
    .map(driver => {
      let latestYear = 0;
      teams.forEach(team => {
        team.history.forEach(h => {
          if (h.drivers.includes(driver.name) && h.year > latestYear) latestYear = h.year;
        });
      });

      const isDriverActive = activeDriversInCurrentYear.includes(driver.name);
      const yearToUse = latestYear === 0 ? (isDriverActive ? currentYear : 0) : latestYear;

      return { ...driver, latestYear: yearToUse, isActive: isDriverActive, points: 0, wins: 0, podiums: 0 };
    })
    .filter(d => d.latestYear > 0);

  // 3. Courses
  const racesSheet = workbook.Sheets['Courses'];
  if (!racesSheet) throw new Error("Onglet 'Courses' manquant.");

  const rawRacesData = XLSX.utils.sheet_to_json(racesSheet);

  races = rawRacesData.map(r => {
    const dateData = r.Date;
    const flagCode = countryCodeFromName(r.Pays) || (r.Flag ? flagCountryCode(String(r.Flag).trim()) : '');

    let dateStr;
    if (dateData instanceof Date) {
      dateStr = dateData.toISOString().split('T')[0];
    } else if (typeof dateData === 'string' || typeof dateData === 'number') {
      dateStr = String(dateData).trim();
    } else {
      dateStr = '2026-12-31';
    }

    const classification = [];
    let positionCounter = 1;

    for (let i = 1; i <= 24; i++) {
      const driverData = r[`P${i}`];
      if (driverData) {
        const driverDataStr = String(driverData).trim();
        if (driverDataStr) {
          const statusMatch = driverDataStr.match(/\((DNF|DSQ|DNS)\)$/i);
          const status = statusMatch ? statusMatch[1].toUpperCase() : 'Finished';
          const cleanName = driverDataStr.replace(/\s+\((DNF|DSQ|DNS)\)$/i, '').trim();

          if (cleanName) {
            classification.push({ pos: positionCounter++, name: cleanName, status: status });
          }
        }
      }
    }

    return {
      id: r.ID,
      name: r.Name,
      location: r.Location,
      country: r.Pays || '',
      circuit: r.Circuit || r.Location || '',
      date: dateStr,
      flag: flagCode,
      laps: Number(r.Tours) || null,
      distance: Number(r.Distance_km) || null,
      sprint: /^(oui|yes|true|1)$/i.test(String(r.Sprint || '').trim()),
      classification: classification
    };
  }).sort((a, b) => new Date(a.date) - new Date(b.date));

  const detailedResultsSheet = workbook.Sheets['Résultats'];
  if (detailedResultsSheet) {
    const detailsByRace = new Map();
    XLSX.utils.sheet_to_json(detailedResultsSheet).forEach(row => {
      const id = Number(row.ID);
      if (!id || !row.Pilote) return;
      if (!detailsByRace.has(id)) detailsByRace.set(id, new Map());
      detailsByRace.get(id).set(String(row.Pilote), row);
    });

    races.forEach(race => {
      const raceDetails = detailsByRace.get(Number(race.id));
      if (!raceDetails) return;
      race.classification.forEach(result => {
        const detail = raceDetails.get(result.name);
        if (!detail) return;
        result.number = Number(detail.Numero) || null;
        result.laps = Number(detail.Tours) || null;
        result.time = detail.Temps ? String(detail.Temps) : '';
        result.gap = detail.Ecart ? String(detail.Ecart) : '';
        result.displayPosition = String(detail.Position || '') === 'NC' ? 'NC' : result.pos;
      });
    });
  }

  const sprintsSheet = workbook.Sheets['Sprints'];
  sprintResults = sprintsSheet ? XLSX.utils.sheet_to_json(sprintsSheet).map(result => {
    const id = Number(result.ID);
    const race = races.find(item => Number(item.id) === id);
    return {
      id,
      year: Number(result.Annee) || Math.floor(id / 100),
      name: result.Pilote,
      points: Number(result.Points) || 0,
      date: result.Date || race?.date
    };
  }) : [];

  const teamStandingsSheet = workbook.Sheets['Classement_Écuries'];
  seasonTeamStandings = teamStandingsSheet
    ? XLSX.utils.sheet_to_json(teamStandingsSheet).map(row => ({
      year: Number(row.Annee),
      teamName: row.Ecurie,
      points: Number(row.Points) || 0,
      wins: Number(row.Victoires) || 0
    }))
    : [];

  refreshSeasonSelector();
}

/* ===============================
  CALCUL DES CLASSEMENTS
  =============================== */
function calculateDriverRanking() {
  driversList.forEach(d => { d.points = 0; d.wins = 0; d.podiums = 0; });

  const season = seasonRaces();
  const seasonDrivers = new Set(
    teams.flatMap(team => getTeamLineup(team, currentYear))
  );
  season.forEach(race => {
    race.classification.forEach(result => seasonDrivers.add(result.name));
  });

  season
    .filter(isRaceCompleted)
    .forEach(race => {
      race.classification.forEach((result, index) => {
        if (result.status !== 'Finished') return;
        const driver = driversList.find(d => d.name === result.name);
        if (!driver) return;

        if (index < 10) driver.points += BAREME_POINTS_F1[index];
        if (index === 0) driver.wins += 1;
        if (index < 3) driver.podiums += 1;
      });
    });

  sprintResults
    .filter(result => result.year === currentYear && isPast(result.date))
    .forEach(result => {
      const driver = driversList.find(d => d.name === result.name);
      if (driver) driver.points += result.points;
    });

  return [...driversList]
    .filter(d => seasonDrivers.has(d.name))
    .sort((a, b) => b.points - a.points || b.wins - a.wins);
}

// Classement des constructeurs : somme des points des pilotes de chaque écurie
function calculateTeamRanking() {
  const savedStandings = seasonTeamStandings
    .filter(row => row.year === currentYear)
    .map(row => {
      const team = teams.find(item => item.name === row.teamName);
      return team && hasTeamLineup(team, currentYear) ? {
        team,
        name: team.name,
        drivers: getTeamLineup(team, currentYear),
        points: row.points,
        wins: row.wins
      } : null;
    })
    .filter(Boolean);

  if (savedStandings.length) return savedStandings;

  const ranking = calculateDriverRanking();

  const rows = teams
    .filter(t => hasTeamLineup(t, currentYear))
    .map(t => ({
      team: t,
      name: t.name,
      drivers: getTeamLineup(t, currentYear),
      points: 0,
      wins: 0
    }));

  ranking.forEach(d => {
    const team = getTeamOfDriver(d.name, currentYear);
    const row = team && rows.find(r => r.team === team);
    if (row) { row.points += d.points; row.wins += d.wins; }
  });

  return rows.sort((a, b) => b.points - a.points || b.wins - a.wins);
}

/* ===============================
  UTILITAIRES
  =============================== */
const DRIVER_IMAGE_YEARS = {
  aitken: [2022],
  albon: [2022, 2024, 2025, 2026],
  alonso: [2003, 2004, 2017, 2021, 2022, 2023, 2024, 2025, 2026],
  antonelli: [2025, 2026],
  bearman: [2025],
  bortoleto: [2025, 2026],
  bottas: [2017, 2022, 2023, 2024, 2026],
  buemi: [2022],
  carlos: [2016, 2026],
  colapinto: [2025, 2026],
  devries: [2022],
  doohan: [2024, 2025],
  drugovich: [2023],
  ericsson: [2017],
  fittipaldi: [2022, 2023],
  gasly: [2022, 2023, 2024, 2025, 2026],
  giovinazzi: [2022, 2023],
  grosjean: [2016, 2020],
  hadjar: [2025, 2026],
  hamilton: [2017, 2022, 2023, 2024, 2025, 2026],
  hartley: [2018],
  hulkenberg: [2017, 2022, 2024, 2025, 2026],
  kubica: [2022],
  kvyat: [2020],
  kyvat: [2017],
  latifi: [2020, 2022],
  lawson: [2023, 2025],
  leclerc: [2022, 2023, 2024, 2025, 2026],
  lindblad: [2026],
  m_schumacher: [2022, 2023],
  magnussen: [2017, 2022, 2024],
  massa: [2017],
  mazepin: [2021],
  norris: [2022, 2024, 2025, 2026],
  ocon: [2017, 2021, 2022, 2024, 2025, 2026],
  palmer: [2017],
  palou: [2023],
  perez: [2017, 2022, 2023, 2024, 2026],
  piastri: [2022, 2024, 2025, 2026],
  raikkonen: [2016, 2021],
  ricciardo: [2017, 2022, 2023, 2024],
  rosberg: [2016],
  russell: [2020, 2022, 2024, 2025, 2026],
  sainz: [2022, 2023, 2024, 2025],
  sargeant: [2023, 2024],
  schumacher: [2000],
  stroll: [2017, 2022, 2025, 2026],
  trulli: [2003],
  tsunoda: [2022, 2023, 2025],
  vandoorne: [2017, 2022],
  verstappen: [2017, 2022, 2023, 2024, 2025, 2026],
  vettel: [2015, 2019, 2020, 2022],
  wehrlein: [2017],
  zhou: [2022, 2023, 2024]
};

function closestAvailableImageYear(years, requestedYear) {
  return years.reduce((closest, year) => {
    const distance = Math.abs(year - requestedYear);
    const closestDistance = Math.abs(closest - requestedYear);
    return distance < closestDistance || (distance === closestDistance && year > closest)
      ? year
      : closest;
  });
}

function getDriverImg(driverName, year) {
  const imageName = getDriverImageGeneric(driverName);
  const availableImageYears = DRIVER_IMAGE_YEARS[imageName];
  if (!availableImageYears) return 'assets/drivers/unkown.png';

  const imageYear = closestAvailableImageYear(availableImageYears, Number(year));
  return `assets/drivers/${imageYear}_full_${imageName}.png`;
}

function getDriverHeadImg(driverName, year) {
  return getDriverImg(driverName, year);
}

function getDriverImageGeneric(driverName) {
  const d = driversList.find(dr => dr.name === driverName);
  const aliases = { arvid_lindblad: 'lindblad' };
  return aliases[d?.generic] || d?.generic || 'unknown';
}

function getTeamLineup(team, year) {
  const futureLineup = FUTURE_LINEUPS[year]?.find(lineup => lineup.teamName === team.name);
  return futureLineup ? futureLineup.drivers.filter(Boolean) : team.history.find(item => item.year === year)?.drivers || [];
}

function hasTeamLineup(team, year) {
  return Boolean(FUTURE_LINEUPS[year]?.some(lineup => lineup.teamName === team.name) ||
    team.history.some(item => item.year === year));
}

function getLineupEntries(year) {
  const futureLineup = FUTURE_LINEUPS[year];
  if (futureLineup) {
    return futureLineup
      .map(entry => ({
        team: teams.find(team => team.name === entry.teamName),
        drivers: entry.drivers
      }))
      .filter(entry => entry.team);
  }

  return teams
    .filter(team => hasTeamLineup(team, year))
    .map(team => ({ team, drivers: getMainDrivers(getTeamLineup(team, year), year) }))
    .sort((a, b) => a.team.name.localeCompare(b.team.name));
}

// Ne garde que les 2 pilotes titulaires : ceux qui ont disputé le plus de courses
// dans la saison (écarte remplaçants et 3es pilotes). À égalité / sans données : l'ordre du fichier.
function getMainDrivers(drivers, year) {
  if (drivers.length <= 2) return drivers;
  const seasonOfYear = races.filter(r => raceYear(r) === year);
  const starts = name => seasonOfYear.filter(r => r.classification.some(c => c.name === name)).length;
  const keep = drivers
    .map((name, order) => ({ name, order, count: starts(name) }))
    .sort((a, b) => b.count - a.count || a.order - b.order)
    .slice(0, 2)
    .sort((a, b) => a.order - b.order);
  return keep.map(d => d.name);
}

function setCarBackground(element, team, requestedYear) {
  const carPath = getCarPath(team, requestedYear);
  if (carPath) element.style.backgroundImage = `url("${carPath}")`;
}

const CAR_IMAGE_NAME_ALIASES = {
  sauber_f1: 'sauber',
  caterham: 'catheram'
};

const CAR_IMAGE_YEARS = {
  alpine: [2025, 2026],
  arrows: [2002],
  astonmartin: [2025, 2026],
  audi: [2026],
  bar: [2005],
  benetton: [1994],
  bmwsauber: [2008],
  brawn: [2009],
  cadillac: [2026],
  catheram: [2012, 2014],
  ferrari: [2023, 2024, 2025, 2026],
  forceindia: [2018],
  haas: [2017, 2025, 2026],
  honda: [2008],
  hrt: [2006],
  jaguar: [2004],
  jordan: [2000],
  lotusf1: [2012],
  lotusracing: [2010],
  manor: [2016],
  marussia: [2013],
  mclaren: [2017, 2025, 2026],
  mercedes: [2014, 2020, 2025, 2026],
  midland: [2006],
  minardi: [2000],
  prost: [1999],
  rbd: [2025, 2026],
  redbull: [2017, 2025, 2026],
  renault: [2003, 2004, 2005, 2006, 2008, 2009, 2017],
  sauber: [2018, 2025],
  spyker: [2007],
  superaguri: [2008],
  tororosso: [2015],
  toyota: [2009],
  virgin: [2011],
  williams: [2017, 2025, 2026]
};

function getCarPath(team, year) {
  const imageName = CAR_IMAGE_NAME_ALIASES[team.genericName] || team.genericName;
  const availableImageYears = CAR_IMAGE_YEARS[imageName];
  if (!availableImageYears) return null;

  const imageYear = closestAvailableImageYear(availableImageYears, Number(year));
  return `assets/cars/${imageYear}${imageName}.png`;
}
function getTeamLogoPath(team) { return `assets/logos/${team.genericName}_logo.png`; }

const getCircuitImagePath = (simpleName) => {
  const standardizedName = String(simpleName || '').toLowerCase().replace(/[\s-]/g, '_').replace('é', 'e').replace('ï', 'i');
  const circuitImages = { barcelona: 'montmelo', mexico_city: 'mexico_city', sao_paulo: 'bresil', 'são_paulo': 'bresil' };
  return `assets/circuits/${circuitImages[standardizedName] || standardizedName}.png`;
};

function isPast(dateStr) { return new Date(dateStr + 'T23:59:59') < new Date(); }
const isRaceCompleted = race => race.classification.length > 0 || isPast(race.date);

const raceYear = r => (r.id ? Math.floor(r.id / 100) : null);
const seasonRaces = () => races.filter(r => raceYear(r) === currentYear);
const lastName = n => String(n).split(' ').pop();
const firstName = n => String(n).split(' ').slice(0, -1).join(' ');

function getTeamOfDriver(name, year) {
  return teams.find(team => getTeamLineup(team, year).includes(name)) || null;
}
const teamColors = t => (t ? [t.color, t.color2 || t.color] : ['#4a4d55', '#23252a']);
const FALLBACK_LOGO = 'assets/logos/unknown_logo.png';
const hideOnError = `onerror="this.style.visibility='hidden'"`;

function parseDate(s) {
  const d = new Date(/^\d{4}-\d{2}-\d{2}$/.test(s) ? s + 'T12:00:00' : s);
  return isNaN(d) ? null : d;
}
function fmtDate(s, opts = { day: 'numeric', month: 'short' }) {
  const d = parseDate(s);
  return d ? d.toLocaleDateString('fr-FR', opts).replace(/\./g, '') : s;
}

/* ===============================
  COMPOSANTS : LIGNES DE CLASSEMENT (style "The Grid")
  =============================== */
function driverRankRow(d, i, maxPts) {
  const team = getTeamOfDriver(d.name, currentYear);
  const [c1, c2] = teamColors(team);
  const pct = maxPts ? Math.max(3, Math.round(d.points / maxPts * 100)) : 0;
  const cls = i < 3 ? ` is-p${i + 1}` : '';
  const gap = i === 0 ? 'Leader' : `−${maxPts - d.points} pts`;

  return `
    <div class="rank-row${cls}" style="--i:${i};--c1:${c1};--c2:${c2};--pct:${pct}%">
      <div class="rk-pos"><span>${i + 1}</span></div>
      <div class="rk-band">
        <img class="rk-logo" src="${team ? getTeamLogoPath(team) : FALLBACK_LOGO}" alt="${team ? team.name : ''}" loading="lazy" ${hideOnError}>
        <img class="rk-head" src="${getDriverHeadImg(d.name, currentYear)}" alt="${d.name}" loading="lazy" ${hideOnError}>
        <div class="rk-name">
          <span class="rk-last">${lastName(d.name)}</span>
        </div>
        <div class="rk-meta">
          <b>${gap}</b>
          <small title="Victoires · Podiums">${d.wins} V · ${d.podiums} P</small>
        </div>
        <i class="rk-bar"></i>
      </div>
      <div class="rk-pts"><b>${d.points}</b><small>PTS</small></div>
    </div>`;
}

function homePodiumCard(d, i) {
  const team = getTeamOfDriver(d.name, currentYear);
  const [c1, c2] = teamColors(team);

  return `
    <article class="home-podium-card is-p${i + 1}" style="--i:${i};--c1:${c1};--c2:${c2}" aria-label="${i + 1}. ${d.name}, ${d.points} points">
      <span class="home-podium-card__position">${i + 1}</span>
      <img class="home-podium-card__driver" src="${getDriverImg(d.name, currentYear)}" alt="${d.name}" loading="lazy" ${hideOnError}>
      <div class="home-podium-card__points">${d.points}<small>PTS</small></div>
      <div class="home-podium-card__info">
        <div class="home-podium-card__name">${lastName(d.name)}</div>
        <small class="home-podium-card__detail">${team ? team.name : ''}</small>
      </div>
    </article>`;
}

function teamRankRow(t, i, maxPts) {
  const [c1, c2] = teamColors(t.team);
  const pct = maxPts ? Math.max(3, Math.round(t.points / maxPts * 100)) : 0;
  const cls = i < 3 ? ` is-p${i + 1}` : '';
  const carPath = getCarPath(t.team, currentYear);

  return `
    <div class="rank-row is-team${cls}" style="--i:${i};--c1:${c1};--c2:${c2};--pct:${pct}%">
      <div class="rk-pos"><span>${i + 1}</span></div>
      <div class="rk-band">
        <img class="rk-logo lg" src="${getTeamLogoPath(t.team)}" alt="${t.name}" loading="lazy" ${hideOnError}>
        <div class="rk-name">
          <span class="rk-first">${t.drivers.map(lastName).join(' · ')}</span>
          <span class="rk-last">${t.name}</span>
        </div>
        <img class="rk-car" ${carPath ? `src="${carPath}"` : 'hidden'} alt="" loading="lazy" ${hideOnError}>
        <i class="rk-bar"></i>
      </div>
      <div class="rk-pts"><b>${t.points}</b><small>PTS</small></div>
    </div>`;
}

/* ===============================
  COMPOSANT : TUILE DE COURSE
  =============================== */
function raceTile(r, round, nextId, i = 0, calendarMode = false) {
  const past = isRaceCompleted(r);
  const state = past ? 'past' : (r.id === nextId ? 'next' : 'future');
  const podium = past ? r.classification.slice(0, 3) : [];
  const year = raceYear(r) || currentYear;

  let c1 = '#e10600', c2 = '#8a0500';
  if (podium[0]) [c1, c2] = teamColors(getTeamOfDriver(podium[0].name, year));
  else if (calendarMode) [c1, c2] = ['#5a5d66', '#2c2e34'];   // course pas encore disputée : pan gris neutre

  let badge = '';
  if (state === 'next') badge = '<span class="rt-badge is-next">Prochaine</span>';
  else if (state === 'future') badge = '<span class="rt-badge">À venir</span>';
  else if (!podium.length) badge = '<span class="rt-badge">Résultats à venir</span>';

  const podiumHtml = !calendarMode && podium.length ? `
    <div class="rt-podium">
      ${podium.map(p => {
        const [pc] = teamColors(getTeamOfDriver(p.name, year));
        return `<span class="pod pod-${p.pos}" style="--c:${pc}"><i>${p.pos}</i><b>${lastName(p.name)}</b></span>`;
      }).join('')}
    </div>` : '';

  return `
    <article class="race-tile is-${state}${calendarMode ? ` is-calendar${i % 4 === 1 || i % 4 === 2 ? ' is-light' : ''}` : ''}" data-race="${r.id}" style="--i:${i};--c1:${c1};--c2:${c2}">
      <div class="rt-round"><small>RD</small><b>${round}</b></div>
      <div class="rt-body">
        <div class="rt-head">
          <div class="rt-meta"><span class="rt-date">${fmtDate(r.date)}</span>${calendarMode ? '' : badge}</div>
          <h4 class="rt-name">${r.name}</h4>
          <p class="rt-loc">${calendarMode ? `<img class="rt-flag" src="https://flagcdn.com/w40/${r.flag}.png" alt="" loading="lazy" ${hideOnError}>` : ''}${r.location}</p>
        </div>
        ${podiumHtml}
      </div>
      <img class="rt-circuit" src="${getCircuitImagePath(r.location)}" alt="Circuit ${r.location}" loading="lazy" ${hideOnError}>
    </article>`;
}

/* ===============================
  RENDER ACCUEIL
  =============================== */
function renderCourses() {
  $('bannerYear').textContent = currentYear;
  renderHero();

  const list = seasonRaces();
  const upcoming = list.filter(r => !isRaceCompleted(r));
  const next = upcoming[0];

  let shown = next ? upcoming.slice(1, 6) : [];
  let title = 'Ensuite au calendrier';
  if (!shown.length) {
    shown = list.filter(isRaceCompleted).slice(-5).reverse();
    title = 'Derniers Grands Prix';
  }
  $('nextRacesTitle').textContent = title;

  $('nextRacesList').innerHTML =
    shown.map((r, i) => raceTile(r, list.indexOf(r) + 1, next ? next.id : null, i)).join('') +
    `<button class="more-tile" data-view="season" data-season-section="calendar">
       <b>Voir plus de GP</b>
       <span class="muted" style="color:rgba(255,255,255,.85)">Calendrier complet</span>
     </button>`;

  const ranking = calculateDriverRanking();
  $('homeStandings').innerHTML = ranking.slice(0, 5).map(homePodiumCard).join('');
}

function renderHero() {
  clearInterval(countdownTimer);
  const el = $('nextHero');
  const list = seasonRaces();
  const next = list.find(r => !isRaceCompleted(r));

  if (next) {
    const round = list.indexOf(next) + 1;
    el.innerHTML = `
      <div class="hero-card">
        <div class="hero-text">
          <span class="kicker">Prochain Grand Prix · Round ${round}</span>
          <h3 class="hero-title">${next.name}</h3>
          <p class="hero-loc"><img class="flag" src="https://flagcdn.com/w80/${next.flag}.png" alt="" loading="lazy" ${hideOnError}>${next.location} · ${fmtDate(next.date, { weekday: 'long', day: 'numeric', month: 'long' })}</p>
          <div class="countdown">
            ${['jours', 'heures', 'min', 'sec'].map((l, i) => `<div class="cd"><b id="cd${i}">--</b><span>${l}</span></div>`).join('')}
          </div>
          <button class="btn-skew" data-race="${next.id}">Voir détails</button>
        </div>
        <div class="hero-art">
          <img class="hero-circuit" src="${getCircuitImagePath(next.location)}" alt="Circuit ${next.location}" ${hideOnError}>
        </div>
      </div>`;
    startCountdown(next);
    return;
  }

  // Saison terminée : on met le champion à l'honneur
  const ranking = calculateDriverRanking();
  const champ = ranking[0];
  if (!list.length || !champ) { el.innerHTML = ''; return; }

  const team = getTeamOfDriver(champ.name, currentYear);
  const [c1, c2] = teamColors(team);
  el.innerHTML = `
    <div class="hero-card hero-champion" style="--c1:${c1};--c2:${c2}">
      <div class="hero-text">
        <span class="kicker">Saison ${currentYear} terminée</span>
        <div class="hero-first">${firstName(champ.name)}</div>
        <div class="hero-big-name">${lastName(champ.name)}</div>
        <div class="hero-first">${team ? team.name : ''}</div>
        <div class="hero-stats">
          <div><b>${champ.points}</b><span>Points</span></div>
          <div><b>${champ.wins}</b><span>Victoires</span></div>
          <div><b>${champ.podiums}</b><span>Podiums</span></div>
        </div>
        <button class="btn-skew" data-view="season" data-season-section="ranking">Voir le classement</button>
      </div>
      <div class="hero-art">
        <img class="hero-driver" src="${getDriverImg(champ.name, currentYear)}" alt="${champ.name}" ${hideOnError}>
      </div>
    </div>`;
}

function startCountdown(race) {
  const target = new Date(race.date + 'T14:00:00');
  const tick = () => {
    let diff = Math.max(0, target - new Date());
    const vals = [
      Math.floor(diff / 86400000),
      Math.floor(diff / 3600000) % 24,
      Math.floor(diff / 60000) % 60,
      Math.floor(diff / 1000) % 60
    ];
    vals.forEach((v, i) => {
      const el = $(`cd${i}`);
      if (el) el.textContent = String(v).padStart(2, '0');
    });
  };
  tick();
  countdownTimer = setInterval(tick, 1000);
}

/* ===============================
  RENDER SAISON
  =============================== */
function renderSeason() {
  $('seasonYearLabel').textContent = currentYear;
  renderSeasonProgress();
  renderSeasonRaces();
  renderSeasonRanking();
}

function renderSeasonProgress() {
  const list = seasonRaces();
  const done = list.filter(isRaceCompleted).length;
  $('spDone').textContent = done;
  $('spTotal').textContent = list.length;
  $('spFill').style.width = list.length ? `${Math.round(done / list.length * 100)}%` : '0%';
}

function renderSeasonRaces() {
  const list = seasonRaces();
  const next = list.find(r => !isRaceCompleted(r));

  const items = list
    .map((r, idx) => ({ r, round: idx + 1 }))
    .filter(({ r }) => raceFilter === 'all' || (raceFilter === 'done' ? isRaceCompleted(r) : !isRaceCompleted(r)));

  $('seasonRacesList').innerHTML = items.length
    ? items.map(({ r, round }, i) => raceTile(r, round, next ? next.id : null, i, true)).join('')
    : '<p class="empty">Aucune course dans cette catégorie.</p>';
}

function renderSeasonRanking() {
  const list = $('rankingList');
  const sub = $('rankSub');

  // Sous-titre "ROUND 17 / SINGAPORE"
  const season = seasonRaces();
  const finished = season.filter(isRaceCompleted);
  const last = finished[finished.length - 1];
  sub.innerHTML = last ? `Round ${season.indexOf(last) + 1} <span>/ ${String(last.location).toUpperCase()}</span>` : '';

  if (rankMode === 'teams') {
    $('rankVertical').textContent = 'Constructeurs';
    const rows = calculateTeamRanking();
    const max = rows.length ? rows[0].points : 0;
    list.innerHTML = rows.map((t, i) => teamRankRow(t, i, max)).join('');
  } else {
    $('rankVertical').textContent = 'Classement';
    const rows = calculateDriverRanking();
    const max = rows.length ? rows[0].points : 0;
    list.innerHTML = rows.map((d, i) => driverRankRow(d, i, max)).join('');
  }
}

/* ===============================
  RENDER CONDUCTEURS
  =============================== */
// Saison à afficher pour un pilote : l'année choisie s'il y a roulé, sinon sa dernière saison avant celle-ci
function driverSeasonYear(d) {
  if (getTeamOfDriver(d.name, currentYear)) return currentYear;
  let best = 0;
  teams.forEach(t => t.history.forEach(h => {
    if (h.year <= currentYear && h.year > best && h.drivers.includes(d.name)) best = h.year;
  }));
  return best || d.latestYear;
}

function renderDrivers() {
  const grid = $('driversGrid');
  calculateDriverRanking(); // met à jour les points

  const q = driverQuery.trim().toLowerCase();

  const list = driversList
    .filter(d => d.latestYear > 0)
    .filter(d => {
      const isActive = Boolean(getTeamOfDriver(d.name, currentYear));
      return driverFilter === 'all' || (driverFilter === 'active' ? isActive : !isActive);
    })
    .filter(d => !q || d.name.toLowerCase().includes(q))
    .map(d => ({ ...d, _y: driverSeasonYear(d) }))
    .sort((a, b) => {
      // 1) pilotes de la saison choisie (au calendrier ou ayant marqué des points), classés par points
      const inSeason = d => Boolean(getTeamOfDriver(d.name, currentYear)) || d.points > 0;
      const sa = inSeason(a), sb = inSeason(b);
      if (sa !== sb) return sa ? -1 : 1;
      if (sa) return b.points - a.points || b.wins - a.wins || a.name.localeCompare(b.name);
      // 2) les autres (anciens pilotes) : du plus récent au plus ancien
      return b._y - a._y || a.name.localeCompare(b.name);
    });

  if (!list.length) {
    grid.innerHTML = '<p class="empty">Aucun pilote ne correspond.</p>';
    return;
  }

  // Numéros de course par saison (issus de l'onglet Résultats) : le dernier de l'année l'emporte
  const numbers = new Map();
  races.forEach(r => (r.classification || []).forEach(c => {
    if (c.number) numbers.set(`${raceYear(r)}|${c.name}`, c.number);
  }));

  grid.innerHTML = list.map((d, i) => {
    const isActive = Boolean(getTeamOfDriver(d.name, currentYear));
    const year = d._y;                                   // saison réellement affichée
    const team = getTeamOfDriver(d.name, year);
    const [c1, c2] = teamColors(team);
    const badge = isActive ? `${d.points} pts` : year;
    const num = numbers.get(`${year}|${d.name}`) || (isActive ? Number(d.numero) || null : null);

    return `
      <div class="driver-tile${isActive ? '' : ' is-inactive'}" style="--i:${Math.min(i, 15)};--c1:${c1};--c2:${c2}">
        <img class="dt-img" src="${getDriverImg(d.name, year)}" alt="${d.name}" loading="lazy" ${hideOnError}>
        <span class="dt-badge">${badge}</span>
        ${num ? `<span class="dt-number">${num}</span>` : ''}
        <div class="dt-info">
          <span class="dt-first">${nationalityCode(d.nationalite) ? `<img class="dt-flag" src="https://flagcdn.com/w40/${nationalityCode(d.nationalite)}.png" alt="${d.nationalite || ''}" title="${d.nationalite || ''}" loading="lazy" ${hideOnError}>` : ''}${firstName(d.name)}</span>
          <span class="dt-last">${lastName(d.name)}</span>
          <div class="dt-team">
            ${team ? `<img src="${getTeamLogoPath(team)}" alt="" loading="lazy" ${hideOnError}>` : ''}
            <span>${team ? team.name : 'Historique'} · ${year}</span>
          </div>
        </div>
      </div>`;
  }).join('');
}

// Photos en pied (très hautes) : on les zoome sur le haut pour n'afficher que le buste
function markFullBody(img) {
  img.parentNode.classList.toggle('is-fullbody', img.naturalHeight / img.naturalWidth > 1.8);
}

function lineupTeamHtml({ team, drivers }, i, pos) {
  const [c1, c2] = teamColors(team);
  const driversHtml = drivers.map(driverName => `
    <figure class="lineup-driver${driverName ? '' : ' is-tbc'}">
      ${driverName
        ? `<div class="lineup-driver__photo"><img src="${getDriverImg(driverName, currentYear)}" alt="${driverName}" loading="lazy" onload="markFullBody(this)" ${hideOnError}></div>`
        : '<div class="lineup-driver__photo"><img class="lineup-driver__silhouette" src="assets/drivers/unkown.png" alt="" loading="lazy"></div>'}
      <figcaption>${driverName ? lastName(driverName) : 'TBC'}</figcaption>
    </figure>`).join('');
  return `<article class="lineup-team" style="--i:${i};--x:${pos.x.toFixed(2)};--y:${pos.y.toFixed(2)};--c1:${c1};--c2:${c2}">
    <div class="lineup-team__brand">
      <img src="${getTeamLogoPath(team)}" alt="${team.name}" loading="lazy" ${hideOnError}>
      <span>${team.name}</span>
    </div>
    <div class="lineup-team__drivers">${driversHtml}</div>
  </article>`;
}

// Dimensions (en cqw, 1 = 1 % de la largeur de la zone) — doivent rester cohérentes avec le CSS
const LINEUP_RING = { cardW: 16.5, cardH: 10.6, height: 50 };

// Répartit n points à égale distance le long d'une ellipse (sens horaire, 4 cartes centrées sur le haut)
function lineupRingPositions(n) {
  const { cardW, cardH, height } = LINEUP_RING;
  const a = 50 - cardW / 2 - 0.6;
  const b = height / 2 - cardH / 2 - 0.6;
  const cx = 50, cy = height / 2;

  // longueur d'arc cumulée (t = 0 en haut, sens horaire)
  const steps = 2000, cum = [0];
  let px = 0, py = -b;
  for (let i = 1; i <= steps; i++) {
    const t = (i / steps) * 2 * Math.PI;
    const x = a * Math.sin(t), y = -b * Math.cos(t);
    cum.push(cum[i - 1] + Math.hypot(x - px, y - py));
    px = x; py = y;
  }
  const total = cum[steps];
  const at = len => {
    len = ((len % total) + total) % total;
    let lo = 0, hi = steps;
    while (lo < hi) { const mid = (lo + hi) >> 1; if (cum[mid] < len) lo = mid + 1; else hi = mid; }
    const t = (lo / steps) * 2 * Math.PI;
    return { x: cx + a * Math.sin(t), y: cy - b * Math.cos(t) };
  };
  const gap = total / n;
  const start = -gap * (n % 2 ? 1.5 : 1.0);   // première carte en haut à gauche
  return Array.from({ length: n }, (_, k) => at(start + k * gap));
}

function renderLineup() {
  const grid = $('lineupGrid');

  const entries = getLineupEntries(currentYear);

  if (!entries.length) {
    grid.innerHTML = '<p class="empty lineup-empty">Les compositions seront affichées dès qu’elles seront disponibles.</p>';
    return;
  }

  // Les écuries forment un anneau aplati autour du titre
  const positions = lineupRingPositions(entries.length);
  grid.innerHTML = `
    <div class="lineup-center" aria-hidden="true">
      <b class="lineup-center__year">${currentYear}</b>
      <span class="lineup-center__label">Driver line-up</span>
      <span class="lineup-center__arrows">▷▷▷</span>
    </div>
    ${entries.map((entry, i) => lineupTeamHtml(entry, i, positions[i])).join('')}`;
}

/* ===============================
  FRISE CHRONOLOGIQUE DES ÉCURIES
  =============================== */
// Filiations entre écuries (le tableur ne contient pas les rachats / renommages).
// Chaque ligne = UNE piste de la frise : les écuries y sont listées dans l'ordre.
// Les écuries absentes de cette liste sont ajoutées automatiquement.
const TEAM_LINEAGES = [
  ['Ferrari'],
  ['McLaren'],
  ['Williams'],
  ['BAR Honda', 'Honda Racing F1', 'Brawn GP', 'Mercedes'],
  ['Jaguar Racing', 'Red Bull Racing'],
  ['Benetton', 'Renault F1 Team', 'Lotus F1 Team', 'Alpine'],
  ['Jordan Grand Prix', 'Midland F1 Racing', 'Spyker F1 Team', 'Force India', 'Aston Martin'],
  ['Minardi', 'Toro Rosso', 'Racing Bulls Formula One Team'],
  ['Sauber', 'BMW Sauber', 'Kick Sauber', 'Audi'],
  // Filiations plus courtes : rangées ensemble dans les pistes restantes
  ['Virgin Racing', 'Marussia F1 Team', 'Manor Racing'],
  ['Lotus Racing', 'Caterham F1 Team']
];
const TIMELINE_MAIN_LANES = 9; // les 9 premières filiations ont chacune leur propre piste

// Découpe les saisons d'une écurie en périodes continues (ex. Renault 2002-2011 puis 2016-2019)
function teamRuns(team) {
  const years = [...new Set(team.history.map(h => h.year))].sort((a, b) => a - b);
  const runs = [];
  years.forEach(y => {
    const last = runs[runs.length - 1];
    if (last && y === last.end + 1) last.end = y;
    else runs.push({ start: y, end: y });
  });
  return runs.map(r => ({ team, start: r.start, end: r.end }));
}

function buildTimelineLanes() {
  const byName = new Map(teams.map(t => [t.name, t]));
  const used = new Set();
  const chains = [];

  TEAM_LINEAGES.forEach(names => {
    const members = names.map(n => byName.get(n)).filter(t => t && t.history.length);
    members.forEach(t => used.add(t.name));
    if (members.length) chains.push(members);
  });
  // Écuries absentes de TEAM_LINEAGES : une filiation d'une seule écurie
  teams.forEach(t => { if (!used.has(t.name) && t.history.length) chains.push([t]); });

  const toItem = members => {
    const segments = members.flatMap(teamRuns).sort((a, b) => a.start - b.start);
    return { segments, start: segments[0].start, end: segments[segments.length - 1].end };
  };

  const items = chains.map(toItem);
  const lanes = items.slice(0, TIMELINE_MAIN_LANES).map(it => ({ segments: it.segments, end: it.end }));
  const firstPackedLane = lanes.length;

  // Le reste est rangé "façon tétris" : même piste, avec au moins 1 saison vide entre deux filiations
  items.slice(TIMELINE_MAIN_LANES).sort((a, b) => a.start - b.start).forEach(it => {
    const lane = lanes.slice(firstPackedLane).find(l => it.start >= l.end + 2);
    if (lane) { lane.segments.push(...it.segments); lane.end = it.end; }
    else lanes.push({ segments: [...it.segments], end: it.end });
  });
  return lanes;
}

// Texte clair ou foncé selon la luminosité de la couleur de fond
function readableTextColor(hex) {
  const m = /^#?([0-9a-f]{6})$/i.exec(String(hex || '').trim());
  if (!m) return '#fff';
  const n = parseInt(m[1], 16);
  const lum = (0.299 * (n >> 16) + 0.587 * ((n >> 8) & 255) + 0.114 * (n & 255)) / 255;
  return lum > 0.62 ? '#111' : '#fff';
}

function renderTeamTimeline(container) {
  const lanes = buildTimelineLanes();
  const allSegs = lanes.flatMap(l => l.segments);
  if (!allSegs.length) return;

  const minYear = Math.min(...allSegs.map(s => s.start));
  const maxYear = Math.max(...allSegs.map(s => s.end));
  const total = maxYear - minYear + 1;
  const pct = y => ((y - minYear) / total) * 100;
  const mid = y => pct(y) + 50 / total; // milieu de la saison y

  // Graduations tous les 5 ans (+ première et dernière saison, sans chevauchement)
  const ticks = [];
  for (let y = minYear; y <= maxYear; y++) {
    if (y === minYear || y === maxYear || (y % 5 === 0 && y - minYear >= 2 && maxYear - y >= 2)) ticks.push(y);
  }
  const cyTick = Math.min(Math.max(currentYear, minYear), maxYear);
  const axisHtml = ticks.filter(y => y !== cyTick).map(y => `<span class="tl-tick" style="left:${mid(y)}%">${y}</span>`).join('');
  const gridHtml = ticks.map(y => `<i class="tl-gridline" style="left:${mid(y)}%"></i>`).join('');

  const lanesHtml = lanes.map(lane => {
    const segs = lane.segments.map(s => {
      const t = s.team;
      const years = s.end - s.start + 1;
      const size = years >= 3 ? 'is-wide' : years >= 2 ? 'is-mid' : 'is-thin';
      const current = hasTeamLineup(t, currentYear) && currentYear >= s.start && currentYear <= s.end;
      return `<div class="tl-seg ${size}${current ? ' is-current' : ''}"
          data-team="${t.name.replace(/"/g, '&quot;')}" data-start="${s.start}" data-end="${s.end}"
          style="left:${pct(s.start)}%;width:${(years / total) * 100}%;--c1:${t.color};--c2:${t.color2 || t.color};--fg:${readableTextColor(t.color)}">
          <img class="tl-logo" src="${getTeamLogoPath(t)}" alt="" loading="lazy" onerror="this.parentElement.classList.add('no-logo');this.remove()">
          <span class="tl-abbr">${t.name.split(' ')[0].slice(0, 3)}</span>
          <span class="tl-name">${t.name}</span>
        </div>`;
    }).join('');
    return `<div class="tl-lane">${segs}</div>`;
  }).join('');

  const cy = Math.min(Math.max(currentYear, minYear), maxYear);
  container.innerHTML = `
    <div class="section-head teams-timeline-head">
      <h3>Frise chronologique</h3><span class="section-line"></span>
      <span class="muted tl-hint">${minYear} – ${maxYear} · survolez une écurie</span>
    </div>
    <div class="tl-scroll">
      <div class="tl-board">
        <div class="tl-axis">${axisHtml}</div>
        <div class="tl-body">
          <div class="tl-grid">${gridHtml}</div>
          <div class="tl-now" style="left:${mid(cy)}%"><b>${cy}</b></div>
          ${lanesHtml}
        </div>
      </div>
    </div>`;

  bindTimelineTooltip(container, maxYear);
}

/* Infobulle de la frise (une seule pour toute la page) */
let tlTip = null;
function bindTimelineTooltip(container, maxYear) {
  if (!tlTip) {
    tlTip = document.createElement('div');
    tlTip.className = 'tl-tip';
    document.body.appendChild(tlTip);
  }
  const hide = () => tlTip.classList.remove('show');

  const show = (seg, x, y) => {
    const t = teams.find(tt => tt.name === seg.dataset.team);
    if (!t) return;
    const start = Number(seg.dataset.start), end = Number(seg.dataset.end);
    const ongoing = end >= maxYear;
    const n = end - start + 1;

    // Prédécesseur / successeur = bloc voisin, collé à celui-ci, sur la même piste
    const sibs = [...seg.parentElement.querySelectorAll('.tl-seg')];
    const idx = sibs.indexOf(seg);
    const prev = sibs[idx - 1], next = sibs[idx + 1];
    const prevName = prev && Number(prev.dataset.end) + 1 === start && prev.dataset.team !== t.name ? prev.dataset.team : null;
    const nextName = next && Number(next.dataset.start) === end + 1 && next.dataset.team !== t.name ? next.dataset.team : null;

    tlTip.style.setProperty('--c1', t.color);
    tlTip.style.setProperty('--c2', t.color2 || t.color);
    tlTip.innerHTML = `
      <div class="tl-tip__bar"></div>
      <div class="tl-tip__title">${t.drapeau ? t.drapeau + ' ' : ''}${t.name}</div>
      ${t.nom_complet && t.nom_complet !== t.name ? `<div class="tl-tip__sub">${t.nom_complet}</div>` : ''}
      <div class="tl-tip__row"><span>Arrivée</span><b>${start}</b></div>
      <div class="tl-tip__row"><span>Départ</span><b>${ongoing ? 'Toujours en F1' : end}</b></div>
      <div class="tl-tip__row"><span>Durée</span><b>${n} saison${n > 1 ? 's' : ''}</b></div>
      ${prevName ? `<div class="tl-tip__link">↳ Succède à <b>${prevName}</b></div>` : ''}
      ${nextName ? `<div class="tl-tip__link">↳ Devient <b>${nextName}</b></div>` : ''}`;

    tlTip.classList.add('show');
    const w = tlTip.offsetWidth, h = tlTip.offsetHeight;
    tlTip.style.left = Math.min(Math.max(8, x + 16), window.innerWidth - w - 8) + 'px';
    tlTip.style.top = Math.max(8, y + 18 + h > window.innerHeight ? y - h - 14 : y + 18) + 'px';
  };

  container.onmousemove = e => {
    const seg = e.target.closest('.tl-seg');
    seg ? show(seg, e.clientX, e.clientY) : hide();
  };
  container.onmouseleave = hide;
  container.onclick = e => {
    const seg = e.target.closest('.tl-seg');
    if (!seg) return;
    // Écran tactile : 1er tap = infobulle, 2e tap = détail de l'écurie
    if (window.matchMedia('(hover: none)').matches && !seg.classList.contains('tl-armed')) {
      container.querySelectorAll('.tl-armed').forEach(s => s.classList.remove('tl-armed'));
      seg.classList.add('tl-armed');
      const r = seg.getBoundingClientRect();
      show(seg, r.left + r.width / 2 - 100, r.bottom - 10);
      return;
    }
    hide();
    openTeamDetail(seg.dataset.team);
  };
}
window.addEventListener('scroll', () => tlTip && tlTip.classList.remove('show'), { passive: true });

/* ===============================
  RENDER ÉCURIES & DÉTAIL
  =============================== */
function renderTeams() {
  const grid = $('teamsGrid');

  const processedTeams = teams.map(t => {
    const isActive = hasTeamLineup(t, currentYear);
    const latestYear = t.history.length > 0 ? t.history[0].year : 0;
    return { ...t, isActive, latestYear: isActive ? currentYear : latestYear };
  }).filter(t => t.latestYear > 0 || t.isActive);

  processedTeams.sort((a, b) => {
    if (a.isActive && !b.isActive) return -1;
    if (!a.isActive && b.isActive) return 1;
    return a.name.localeCompare(b.name);
  });

  grid.innerHTML = '';

  grid.insertAdjacentHTML('beforeend', `
    <nav class="jump-bar teams-jump" aria-label="Aller à une section">
      <span class="jump-bar__label">Aller à</span>
      <button class="chip" data-jump="teams-current">Écuries actuelles</button>
      <button class="chip" data-jump="teams-timeline">Frise chronologique</button>
      <button class="chip" data-jump="teams-past">Écuries historiques</button>
    </nav>`);

  const activeTeams = processedTeams.filter(t => t.isActive);
  const pastTeams = processedTeams.filter(t => !t.isActive);

  // --- Écuries actuelles : bandes verticales ---
  if (activeTeams.length) {
    const strip = document.createElement('div');
    strip.className = 'teams-strip';
    strip.id = 'teams-current';
    strip.style.setProperty('--n', activeTeams.length);
    activeTeams.forEach((t, i) => {
      const el = document.createElement('div');
      el.className = 'team-strip';
      el.style.setProperty('--team-color', t.color);
      el.style.setProperty('--team-color2', t.color2 || t.color);
      el.style.setProperty('--i', i);
      el.title = t.name;
      el.onclick = () => openTeamDetail(t.name);
      el.innerHTML = `
        <img class="ts-logo" src="${getTeamLogoPath(t)}" alt="" loading="lazy" ${hideOnError}>
        <div class="ts-car"></div>
        <span class="ts-name">${t.name}</span>`;
      strip.appendChild(el);
      setCarBackground(el.querySelector('.ts-car'), t, currentYear);
    });
    grid.appendChild(strip);
  }

  // --- Frise chronologique (entre actuelles et historiques) ---
  const timelineBlock = document.createElement('div');
  timelineBlock.className = 'teams-timeline';
  timelineBlock.id = 'teams-timeline';
  grid.appendChild(timelineBlock);
  renderTeamTimeline(timelineBlock);

  // --- Écuries historiques : cartes ---
  if (pastTeams.length) {
    const head = document.createElement('div');
    head.className = 'section-head teams-past-head';
    head.id = 'teams-past';
    head.innerHTML = '<h3>Écuries historiques</h3><span class="section-line"></span>';
    grid.appendChild(head);
  }
  pastTeams.forEach((t, i) => {
    const card = document.createElement('div');
    card.className = 'team-card inactive-team';
    card.style.setProperty('--team-color', t.color);
    card.style.setProperty('--team-color2', t.color2);
    card.style.setProperty('--i', Math.min(i, 12));
    card.onclick = () => openTeamDetail(t.name);

    card.innerHTML = `
      <div class="team-car-preview"></div>
      <div class="team-card-info">
        <img class="tc-logo" src="${getTeamLogoPath(t)}" alt="" loading="lazy" ${hideOnError}>
        <div>
          <h4>${t.name}</h4>
          <div class="muted">${t.latestYear} · Historique</div>
        </div>
      </div>`;
    grid.appendChild(card);
    setCarBackground(card.querySelector('.team-car-preview'), t, t.latestYear);
  });
}

/* Aller à une écurie : crée une entrée d'historique (le bouton retour revient à la grille) */
function openTeamDetail(teamName) {
  navigate({ view: 'teams', team: teamName });
}

let teamSpy = null;   // IntersectionObserver de la barre d'années

function showTeamDetail(teamName) {
  const team = teams.find(t => t.name === teamName);
  if (!team) return;

  $('teamsGrid').classList.add('hidden');
  $('teamDetail').classList.remove('hidden');
  if (teamSpy) { teamSpy.disconnect(); teamSpy = null; }

  const esc = v => String(v ?? '').replace(/[&<>"']/g, c => ({ '&': '&amp;', '<': '&lt;', '>': '&gt;', '"': '&quot;', "'": '&#39;' })[c]);
  const c1 = team.color, c2 = team.color2 || team.color;
  const hist = team.history;                                   // déjà trié du plus récent au plus ancien
  const dataMax = Math.max(...teams.flatMap(t => t.history.map(h => h.year)));
  const first = team.premiere_saison || hist[hist.length - 1]?.year || '–';
  const lastY = team.derniere_saison || hist[0]?.year;
  const ongoing = hist.some(h => h.year >= dataMax);
  const flagCode = countryCodeFromName(team.pays) || flagCountryCode(String(team.drapeau || ''));
  const dash = v => (v === null || v === undefined || v === '' ? '–' : v);

  const stats = [
    ['Saisons', hist.length],
    ['Première', first],
    ['Dernière', ongoing ? 'En cours' : dash(lastY)],
    ['Victoires', dash(team.victoires)],
    ['Titres', dash(team.titres_constructeurs)]
  ].map(([label, value]) => `<div><b>${esc(value)}</b><span>${label}</span></div>`).join('');

  const hero = `
    <header class="td-hero" style="--c1:${c1};--c2:${c2}">
      <div class="td-hero__text">
        <span class="kicker">Écurie${team.pays ? ` · ${esc(team.pays)}` : ''}</span>
        <h3 class="td-title">${esc(team.name)}</h3>
        <p class="td-full">
          ${flagCode ? `<img class="td-flag" src="https://flagcdn.com/w40/${flagCode}.png" alt="" onerror="this.remove()">` : ''}
          ${esc(team.nom_complet && team.nom_complet !== team.name ? team.nom_complet : (team.pays || ''))}
        </p>
        <div class="td-stats">${stats}</div>
      </div>
      <div class="td-hero__art">
        <img class="td-hero__logo" src="${getTeamLogoPath(team)}" alt="" ${hideOnError}>
        <div class="td-hero__car"></div>
      </div>
    </header>`;

  const rail = hist.length > 1 ? `
    <nav class="jump-bar td-years" aria-label="Aller à une saison">
      <span class="jump-bar__label">Saison</span>
      ${hist.map(h => `<button class="chip" data-jump="td-${h.year}">${h.year}</button>`).join('')}
    </nav>` : '';

  const fg = readableTextColor(c1);
  const seasons = hist.map((h, i) => {
    const podiumCls = h.pos && h.pos <= 3 ? ` is-p${h.pos}` : '';
    const facts = [
      ['Points', h.points],
      ['Victoires', h.wins],
      ['Courses', h.races],
      ['Moteur', h.engine]
    ].filter(([, v]) => v !== null && v !== undefined && v !== '')
      .map(([label, v]) => `<div class="td-fact"><span>${label}</span><b>${esc(v)}</b></div>`).join('');

    const drivers = h.drivers.map(name => `
      <figure class="td-driver" title="${esc(name)}">
        <img src="${getDriverImg(name, h.year)}" alt="${esc(name)}" loading="lazy" ${hideOnError}>
        <figcaption>${esc(lastName(name))}</figcaption>
      </figure>`).join('');

    return `
      <article class="td-season" id="td-${h.year}" style="--c1:${c1};--c2:${c2};--fg:${fg};--i:${Math.min(i, 10)}">
        <div class="td-year">
          <b>${h.year}</b>
          ${h.pos ? `<span class="td-pos${podiumCls}">P${h.pos}</span><small>Constructeurs${h.pos === 1 ? ' · Champion' : ''}</small>` : ''}
        </div>
        <div class="td-main">
          <div class="td-car" data-car="${h.year}"></div>
          ${facts ? `<div class="td-facts">${facts}</div>` : ''}
          <div class="td-drivers">${drivers}</div>
        </div>
      </article>`;
  }).join('');

  $('teamDetailBody').innerHTML = `
    ${hero}
    <div class="section-head td-seasons-head"><h3>Saison par saison</h3><span class="section-line"></span></div>
    ${rail}
    <div class="td-seasons">${seasons || '<p class="empty">Aucune saison enregistrée.</p>'}</div>`;

  setCarBackground($('teamDetailBody').querySelector('.td-hero__car'), team, hist[0]?.year || currentYear);
  $('teamDetailBody').querySelectorAll('.td-car').forEach(el => setCarBackground(el, team, Number(el.dataset.car)));

  // Mise en surbrillance de la saison visible dans la barre
  const railEl = $('teamDetailBody').querySelector('.td-years');
  if (railEl && 'IntersectionObserver' in window) {
    teamSpy = new IntersectionObserver(entries => {
      // En remontant tout en haut (première saison repassée sous la zone de lecture) : plus aucune surbrillance
      entries.filter(e => !e.isIntersecting && e.target === $('teamDetailBody').querySelector('.td-season') && e.boundingClientRect.top > 0)
        .forEach(() => railEl.querySelectorAll('.chip').forEach(c => c.classList.remove('active')));
      entries.filter(e => e.isIntersecting).forEach(e => {
        railEl.querySelectorAll('.chip').forEach(c => c.classList.toggle('active', c.dataset.jump === e.target.id));
        const on = railEl.querySelector('.chip.active');
        if (on) railEl.scrollLeft = on.offsetLeft - railEl.clientWidth / 2 + on.offsetWidth / 2;
      });
    }, { rootMargin: '-30% 0px -60% 0px' });
    $('teamDetailBody').querySelectorAll('.td-season').forEach(el => teamSpy.observe(el));
  }
}

$('backToTeams').onclick = () => {
  if (route.team && history.state?.pushed) history.back();           // revient à la grille, au même endroit
  else navigate({ view: 'teams' }, { replace: true });
};

/* ===============================
  MODALE COURSE (DÉTAIL)
  =============================== */
const modal = $('modal');

function openRaceModal(raceId) {
  const race = races.find(r => r.id === raceId);
  if (!race) return;

  const year = raceYear(race) || currentYear;
  const round = races.filter(item => raceYear(item) === year).indexOf(race) + 1;
  const raceLabel = race.location || race.name;
  const escapedRaceLabel = String(raceLabel).replace(/[&<>"']/g, character => ({
    '&': '&amp;', '<': '&lt;', '>': '&gt;', '"': '&quot;', "'": '&#39;'
  })[character]);
  const escapeHtml = value => String(value ?? '').replace(/[&<>"']/g, character => ({
    '&': '&amp;', '<': '&lt;', '>': '&gt;', '"': '&quot;', "'": '&#39;'
  })[character]);
  const podium = race.classification.slice(0, 3);
  const podiumOrder = [podium[1], podium[0], podium[2]].filter(Boolean);

  $('modalTitle').textContent = `${race.name} — ${fmtDate(race.date, { day: 'numeric', month: 'long', year: 'numeric' })}`;
  modal.classList.add('modal--race-results');

  let html = `
    <div class="race-results">
      <div class="race-results__banner">
        <span class="race-results__stripe"></span>
        <img src="assets/f1.png" alt="Formule 1" class="race-results__logo">
        <span class="race-results__stripe"></span>
      </div>
      <div class="race-results__heading">
        <span>${round > 0 ? `ROUND ${round}` : 'RACE'}</span><i>/</i><b>${escapedRaceLabel}</b>
      </div>`;

  if (race.classification.length === 0) {
    html += `<div class="race-results__upcoming">
      <div class="race-results__upcoming-circuit">
        <img src="${getCircuitImagePath(race.location)}" alt="" ${hideOnError}>
        <span>${escapeHtml(race.circuit || race.location)}</span>
      </div>
      <div class="race-results__upcoming-info">
        <span class="race-results__upcoming-status">Week-end à venir</span>
        <h4>${escapeHtml(fmtDate(race.date, { weekday: 'long', day: 'numeric', month: 'long' }))}<span>${year}</span></h4>
        <p>${escapeHtml(race.location)}${race.country ? ` · ${escapeHtml(race.country)}` : ''}</p>
        <div class="race-results__upcoming-stats">
          ${race.laps ? `<span><b>${race.laps}</b><small>tours</small></span>` : ''}
          ${race.distance ? `<span><b>${escapeHtml(Number(race.distance).toLocaleString('fr-FR'))}</b><small>km</small></span>` : ''}
          <span><b>${race.sprint ? 'Sprint' : 'Classique'}</b><small>format</small></span>
        </div>
        <p class="race-results__upcoming-note">Le classement et les résultats détaillés seront disponibles après la course.</p>
      </div>
    </div>`;
  } else {
    html += `<section class="race-podium" aria-label="Podium">${podiumOrder.map(p => {
      const team = getTeamOfDriver(p.name, year);
      const [c1, c2] = teamColors(team);
      const pts = (p.pos <= 10 && p.status === 'Finished') ? BAREME_POINTS_F1[p.pos - 1] : 0;
      const detail = p.laps ? `${p.laps} tours · ${p.time || p.gap || p.status}` : '';
      return `<article class="race-podium-card is-p${p.pos}" style="--c1:${c1};--c2:${c2}">
        <span class="race-podium-card__position">${p.pos}</span>
        <img class="race-podium-card__driver" src="${getDriverImg(p.name, year)}" alt="${p.name}" ${hideOnError}>
        <div class="race-podium-card__info">
          <div class="race-podium-card__name">${lastName(p.name)}</div>
          ${detail ? `<small class="race-podium-card__detail">${detail}</small>` : ''}
          <div class="race-podium-card__points">${pts}<small>PTS</small></div>
        </div>
      </article>`;
    }).join('')}</section>`;

    const remainingResults = race.classification.filter(result => result.pos > 3);
    if (remainingResults.length) {
      html += `<section class="race-results__table" aria-label="Classement de la course">${remainingResults.map((p, i) => {
        const team = getTeamOfDriver(p.name, year);
        const pts = (p.pos <= 10 && p.status === 'Finished') ? BAREME_POINTS_F1[p.pos - 1] : 0;
        const status = p.status !== 'Finished' ? `<small class="race-results__status">${p.status}</small>` : '';
        const raceDetail = p.laps ? `${p.laps} tours · ${p.time || p.gap || p.status}` : '';
        return `<div class="race-result-row" style="--i:${i}">
          <span class="race-result-row__position">${p.displayPosition || p.pos}</span>
          <span class="race-result-row__driver">
            <span>${p.number ? `<small class="race-results__number">#${p.number}</small>` : ''}${p.name}${status}</span>
            ${raceDetail ? `<small class="race-results__detail">${raceDetail}</small>` : ''}
          </span>
          <span class="race-result-row__team">
            <img src="${team ? getTeamLogoPath(team) : FALLBACK_LOGO}" alt="" ${hideOnError}>
            <span>${team ? team.name : 'Écurie inconnue'}</span>
          </span>
          <span class="race-result-row__points"><b>${pts || '–'}</b><small>${pts === 1 ? 'PT' : 'PTS'}</small></span>
        </div>`;
      }).join('')}</section>`;
    }
  }

  html += '</div>';
  $('modalContent').innerHTML = html;
  modal.classList.add('visible');
  modal.setAttribute('aria-hidden', 'false');
}

/* Fermeture demandée par l'utilisateur (croix, Échap, clic sur le fond) : on dépile l'historique */
function requestCloseModal() {
  if (!modal.classList.contains('visible')) return;
  if (route.race && history.state?.pushed) history.back();
  else if (route.race) navigate({ ...route, race: null }, { replace: true });
  else closeModal();
}

function closeModal() {
  modal.classList.remove('visible');
  modal.classList.remove('modal--race-results');
  modal.setAttribute('aria-hidden', 'true');
}

/* ===============================
  NAVIGATION (délégation d'événements)
  =============================== */
const VIEWS = ['courses', 'season', 'drivers', 'teams', 'lineup'];

/* ===============================
  DRAPEAU À DAMIER (bannière d'accueil)
=============================== */
function buildCheckeredFlag() {
  const host = $('checkerFlag');
  if (!host) return;
  const w = host.clientWidth;
  if (!w) return;                                   // vue masquée : on reconstruira à l'affichage

  const rows = 4;
  const size = w < 600 ? 16 : 26;                   // taille d'une case
  const cols = Math.ceil(w / size);
  const fade = Math.max(4, Math.min(16, Math.floor(cols / 4)));   // zone de dispersion aux extrémités

  // petit générateur pseudo-aléatoire : rendu identique à chaque redimensionnement
  let seed = 1991;
  const rnd = () => { seed |= 0; seed = (seed + 0x6D2B79F5) | 0; let t = Math.imul(seed ^ (seed >>> 15), 1 | seed);
    t = (t + Math.imul(t ^ (t >>> 7), 61 | t)) ^ t; return ((t ^ (t >>> 14)) >>> 0) / 4294967296; };

  host.style.height = `${rows * size}px`;
  const frag = document.createDocumentFragment();

  for (let r = 0; r < rows; r++) {
    for (let c = 0; c < cols; c++) {
      if ((r + c) % 2) continue;                    // damier : une case sur deux
      const edge = Math.min(c, cols - 1 - c);
      let scale = 1, jy = 0, jx = 0;
      if (edge < fade) {
        const p = edge / fade;                      // 0 = bord, 1 = début du damier plein
        if (rnd() > p * 1.15 + 0.06) continue;      // des cases disparaissent vers les bords
        scale = 0.22 + 0.78 * p;
        jy = (rnd() - 0.5) * size * 0.7 * (1 - p);
        jx = (rnd() - 0.5) * size * 0.5 * (1 - p);
      }
      const d = size * 0.84 * scale;
      const el = document.createElement('i');
      el.style.cssText = `width:${d}px;height:${d}px;left:${c * size + (size - d) / 2 + jx}px;top:${r * size + (size - d) / 2 + jy}px;` +
        `--d:${Math.round(Math.abs(c - cols / 2) / (cols / 2) * 650 + rnd() * 260)}ms`;
      frag.appendChild(el);
    }
  }
  host.replaceChildren(frag);
}

let flagResizeTimer;
window.addEventListener('resize', () => {
  clearTimeout(flagResizeTimer);
  flagResizeTimer = setTimeout(buildCheckeredFlag, 150);
});

/* ===============================
  ROUTEUR : historique du navigateur + liens d'ancrage ("jump links")
  Chaque point clé (page, écurie, Grand Prix, ancre) = une entrée d'historique,
  avec la position de défilement mémorisée pour y revenir via le bouton Retour.
  Format : #/teams?saison=2026&ecurie=Alpine&course=202601&ancre=td-2025
  =============================== */
history.scrollRestoration = 'manual';

let route = { view: 'courses', team: null, race: null, at: null, year: null, section: 'calendar' };
let shown = {};              // page actuellement affichée (évite de tout re-rendre pour une simple modale)
let scrollSaveTimer = null;

function routeToHash(r) {
  const p = new URLSearchParams();
  if (r.year) p.set('saison', r.year);
  if (r.team) p.set('ecurie', r.team);
  if (r.race) p.set('course', r.race);
  if (r.at) p.set('ancre', r.at);
  if (r.view === 'season' && r.section === 'ranking') p.set('section', 'ranking');
  const q = p.toString();
  return `#/${r.view}${q ? '?' + q : ''}`;
}

function parseHash() {
  const m = location.hash.match(/^#\/([a-z]+)(?:\?(.*))?$/);
  if (!m || !VIEWS.includes(m[1])) return { view: 'courses', team: null, race: null, at: null, year: null };
  const p = new URLSearchParams(m[2] || '');
  return {
    view: m[1],
    team: m[1] === 'teams' ? (p.get('ecurie') || null) : null,
    race: Number(p.get('course')) || null,
    at: p.get('ancre') || null,
    year: Number(p.get('saison')) || null,
    section: p.get('section') === 'ranking' ? 'ranking' : 'calendar'
  };
}

function saveScroll() {
  if (!dataLoaded) return;
  history.replaceState({ ...(history.state || {}), scrollY: window.scrollY }, '');
}
window.addEventListener('scroll', () => {
  clearTimeout(scrollSaveTimer);
  scrollSaveTimer = setTimeout(saveScroll, 120);
}, { passive: true });

function restoreScroll(y) {
  window.scrollTo({ top: y, behavior: 'auto' });
  if (y) {   // les images / animations peuvent encore modifier la hauteur de page
    requestAnimationFrame(() => window.scrollTo({ top: y, behavior: 'auto' }));
    setTimeout(() => window.scrollTo({ top: y, behavior: 'auto' }), 300);
  }
}

function jumpTo(id, instant) {
  requestAnimationFrame(() => {
    const el = document.getElementById(id);
    if (el) el.scrollIntoView({ behavior: instant ? 'auto' : 'smooth', block: 'start' });
  });
}

function navigate(next, { replace = false } = {}) {
  const target = {
    view: next.view,
    team: next.view === 'teams' ? (next.team || null) : null,
    race: next.race || null,
    at: next.at || null,
    year: next.year || currentYear,
    section: next.view === 'season' ? (next.section || seasonSection) : 'calendar'
  };
  const hash = routeToHash(target);
  clearTimeout(scrollSaveTimer);

  if (hash === location.hash) {          // même destination : on rafraîchit la page sans empiler d'entrée
    replace = true;
    shown = {};
  }
  if (!replace) saveScroll();            // mémorise où l'on était avant de partir

  const state = { route: target, scrollY: replace ? window.scrollY : 0, pushed: replace ? Boolean(history.state?.pushed) : true };
  history[replace ? 'replaceState' : 'pushState'](state, '', hash);
  applyRoute(target);
}

function applyRoute(r, { pop = false, scrollY = 0 } = {}) {
  if (!dataLoaded) return;

  if (r.year && r.year !== currentYear && availableYears.includes(r.year)) {
    currentYear = r.year;
    shown = {};
  }
  if ($('seasonSelect')) $('seasonSelect').value = String(currentYear);
  updateSeasonSelectButton();

  const pageChanged = shown.view !== r.view ||
    shown.team !== (r.team || null) ||
    shown.year !== currentYear;
  if (pageChanged) {
    showView(r.view, r.team);
    shown = { view: r.view, team: r.team || null, year: currentYear };
  }
  if (r.view === 'season') setSeasonSection(r.section);
  route = r;

  if (r.race) openRaceModal(r.race); else closeModal();

  if (pop && scrollY) restoreScroll(scrollY);
  else if (r.at) jumpTo(r.at, pageChanged || pop);
  else if (pageChanged || pop) restoreScroll(0);
}

window.addEventListener('popstate', e => {
  if (!dataLoaded) return;
  applyRoute(e.state?.route || parseHash(), { pop: true, scrollY: e.state?.scrollY || 0 });
});

let seasonSection = 'calendar';

function setSeasonSection(section) {
  seasonSection = section === 'ranking' ? 'ranking' : 'calendar';
  $('seasonSectionTitle').textContent = seasonSection === 'ranking' ? 'Classement' : 'Calendrier';
  $('seasonCalendar').classList.toggle('hidden', seasonSection !== 'calendar');
  $('seasonRanking').classList.toggle('hidden', seasonSection !== 'ranking');
  document.querySelectorAll('header .nav-btn[data-view="season"]').forEach(button => {
    button.classList.toggle('active', button.dataset.seasonSection === seasonSection);
  });
}

function showView(v, team = null) {
  if (!dataLoaded || !VIEWS.includes(v)) return;

  VIEWS.forEach(name => $(`view-${name}`).classList.toggle('hidden', name !== v));
  document.querySelectorAll('header .nav-btn[data-view]').forEach(b => {
    b.classList.toggle('active', b.dataset.view === v &&
      (v !== 'season' || b.dataset.seasonSection === seasonSection));
  });

  // Par défaut : grille des écuries (le détail est affiché plus bas si une écurie est demandée)
  $('teamDetail').classList.add('hidden');
  $('teamsGrid').classList.remove('hidden');

  if (v === 'courses') { renderCourses(); buildCheckeredFlag(); }
  if (v === 'season') renderSeason();
  if (v === 'drivers') renderDrivers();
  if (v === 'teams') renderTeams();
  if (v === 'lineup') renderLineup();
  if (v === 'teams' && team) showTeamDetail(team);
}

function setActiveChip(container, attr, value) {
  container.querySelectorAll('.chip').forEach(c => c.classList.toggle('active', c.dataset[attr] === value));
}

document.addEventListener('click', e => {
  const seasonWrapper = document.querySelector('.season-select-wrap');
  if (seasonWrapper && !seasonWrapper.contains(e.target)) setSeasonMenuOpen(false);

  const seasonOption = e.target.closest('[data-season-year]');
  if (seasonOption) {
    const seasonSelect = $('seasonSelect');
    seasonSelect.value = seasonOption.dataset.seasonYear;
    seasonSelect.dispatchEvent(new Event('change', { bubbles: true }));
    setSeasonMenuOpen(false);
    $('seasonSelectButton').focus();
    return;
  }

  const raceEl = e.target.closest('[data-race]');
  if (raceEl) { navigate({ ...route, race: Number(raceEl.dataset.race), at: null }); return; }

  const viewEl = e.target.closest('[data-view]');
  if (viewEl) {
    navigate({
      view: viewEl.dataset.view,
      section: viewEl.dataset.view === 'season' ? viewEl.dataset.seasonSection : undefined
    });
    return;
  }

  const jumpEl = e.target.closest('[data-jump]');
  if (jumpEl) { navigate({ ...route, race: null, at: jumpEl.dataset.jump }); return; }

  const filterEl = e.target.closest('[data-filter]');
  if (filterEl) {
    raceFilter = filterEl.dataset.filter;
    setActiveChip($('raceFilters'), 'filter', raceFilter);
    renderSeasonRaces();
    return;
  }

  const rankEl = e.target.closest('[data-rank]');
  if (rankEl) {
    rankMode = rankEl.dataset.rank;
    setActiveChip($('rankTabs'), 'rank', rankMode);
    renderSeasonRanking();
    return;
  }

  const dEl = e.target.closest('[data-dfilter]');
  if (dEl) {
    driverFilter = dEl.dataset.dfilter;
    setActiveChip($('driverFilters'), 'dfilter', driverFilter);
    renderDrivers();
    return;
  }

  if (e.target === modal) requestCloseModal();
});

document.addEventListener('keydown', e => {
  const seasonButton = $('seasonSelectButton');
  const seasonOptions = $('seasonOptions');
  if (e.key === 'Escape' && seasonButton?.getAttribute('aria-expanded') === 'true') {
    e.preventDefault();
    setSeasonMenuOpen(false);
    seasonButton.focus();
    return;
  }
  if (e.target === seasonButton && ['ArrowDown', 'ArrowUp'].includes(e.key)) {
    e.preventDefault();
    setSeasonMenuOpen(true, true);
    return;
  }
  if (e.target.matches?.('.season-option') && ['ArrowDown', 'ArrowUp', 'Home', 'End'].includes(e.key)) {
    e.preventDefault();
    const options = Array.from(seasonOptions.querySelectorAll('.season-option'));
    const currentIndex = options.indexOf(e.target);
    const nextIndex = e.key === 'Home' ? 0 : e.key === 'End' ? options.length - 1 :
      (currentIndex + (e.key === 'ArrowDown' ? 1 : -1) + options.length) % options.length;
    options[nextIndex]?.focus();
    return;
  }
  if (e.key === 'Escape') requestCloseModal();
});
$('driverSearch').addEventListener('input', e => { driverQuery = e.target.value; renderDrivers(); });
$('seasonSelectButton')?.addEventListener('click', () => {
  setSeasonMenuOpen($('seasonSelectButton').getAttribute('aria-expanded') !== 'true', true);
});
$('seasonSelect')?.addEventListener('change', e => {
  const year = Number(e.target.value || currentYear);
  currentYear = year;
  updateSeasonSelectButton();
  if (!dataLoaded) { currentYear = year; return; }
  // Changer de saison ne crée pas d'entrée d'historique : on remplace l'entrée courante
  navigate({ view: route.view, team: route.team, year, section: route.section }, { replace: true });
});
window.matchMedia('(max-width: 620px)').addEventListener('change', () => {
  if (dataLoaded) renderLineup();
});

/* ===============================
  INITIALISATION
  =============================== */
function initApp() {
  dataLoaded = true;
  const initial = parseHash();   // permet de retomber sur la bonne page après un rechargement
  if (initial.year && availableYears.includes(initial.year)) currentYear = initial.year;
  const r = { ...initial, year: currentYear };
  history.replaceState({ route: r, scrollY: 0, pushed: false }, '', routeToHash(r));
  shown = {};
  applyRoute(r);
}

window.addEventListener('DOMContentLoaded', () => {
  const closeButton = document.querySelector('#modal .close-btn');
  if (closeButton) closeButton.onclick = requestCloseModal;
  closeModal();

  // Keep the application hidden until its published workbook has loaded.
  VIEWS.forEach(name => $(`view-${name}`).classList.add('hidden'));
  $('data-loader-view').classList.remove('hidden');
  $('retryDataLoadBtn').addEventListener('click', loadData);
  loadData();
});

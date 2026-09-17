// ============================================================
// TEAM DAYS WORKED - Google Apps Script v1
// ============================================================
// Configuration via the "Teams" tab in your Sheet:
//
//   | Team   | Name      | Email                |
//   |--------|-----------|----------------------|
//   | Team A | Pier      | pier@acme.be       |
//   | Team A | Pol       | pol@acme.be        |
//
// Shows the previous calendar month and the current calendar month, split
// into weeks, with the number of days actually worked per week (derived
// from the ABSENCE of an "Out of Office" Google Calendar event that day).
// A month rarely starts or ends on a Monday/Friday, so the first/last week
// of a month is clipped to the days that actually fall inside that month.
// ============================================================

const WERKUREN_PER_DAG = 8;
const SHEET_TEAMS = "Teams";
const SHEET_DAGEN = "Days Worked";

// Acme brand colors
const GRIJS           = "#3C3C3B";
const BLAUW_LICHT     = "#d6e8f5";

const TEAM_KLEUREN = {
  "Team A": "#3E7CBF",
  "Team B": "#27B4AF",
  "Team C": "#2e6da4",
  "Team D": "#1e9490",
};

function teamKleur(teamNaam) {
  return TEAM_KLEUREN[teamNaam] || "#5f6368";
}

// ============================================================
// MONTH CONFIG — current calendar month + next calendar month
// ============================================================
function bepaalMaanden(vandaag) {
  const vorigeStart = new Date(vandaag.getFullYear(), vandaag.getMonth() - 1, 1);
  const huidigeStart = new Date(vandaag.getFullYear(), vandaag.getMonth(), 1);
  const volgendeStart = new Date(vandaag.getFullYear(), vandaag.getMonth() + 1, 1);
  return [
    { naam: formatMaandNaam(vorigeStart), start: vorigeStart, einde: huidigeStart },
    { naam: formatMaandNaam(huidigeStart), start: huidigeStart, einde: volgendeStart },
  ];
}

// Build the calendar weeks a month spans, with actual working days per week.
// Convention: [maand.start, maand.einde) — first day in, last day out. A week
// straddling the month boundary only counts the days that fall inside it.
function buildMaandWeken(maand) {
  const weken = [];
  const cursor = getMaandagVanWeek(maand.start);
  while (cursor < maand.einde) {
    const weekStart = new Date(cursor);
    const weekEinde = new Date(cursor); weekEinde.setDate(cursor.getDate() + 4);
    let werkdagen = 0;
    const dc = new Date(weekStart);
    while (dc <= weekEinde) {
      const dag = dc.getDay();
      if (dag !== 0 && dag !== 6 && dc >= maand.start && dc < maand.einde) werkdagen++;
      dc.setDate(dc.getDate() + 1);
    }
    if (werkdagen > 0) {
      weken.push({ week: { start: weekStart, einde: weekEinde }, werkdagen });
    }
    cursor.setDate(cursor.getDate() + 7);
  }
  return weken;
}

// Count absent hours from cache, converted to days, only for days within [periodeStart, periodeEinde)
function berekenAfwezigInWeek(dagCache, weekStart, weekEinde, periodeStart, periodeEinde) {
  let totaalUren = 0;
  const cursor = new Date(weekStart);
  while (cursor <= weekEinde) {
    if (cursor >= periodeStart && cursor < periodeEinde) {
      totaalUren += dagCache[cursor.toDateString()] || 0;
    }
    cursor.setDate(cursor.getDate() + 1);
  }
  return Math.round((totaalUren / WERKUREN_PER_DAG) * 4) / 4;
}

// ============================================================
// MAIN FUNCTION
// refreshDagenGewerkt() shows the loading dialog; the dialog calls doRefresh()
// ============================================================
function refreshDagenGewerkt() {
  // Touch the Calendar service here, directly from the menu click, so that on
  // first use Apps Script can show the real Google authorization prompt.
  // google.script.run calls made from inside the dialog below cannot trigger
  // that prompt — they just fail with a generic "permission needed" error.
  Calendar.CalendarList.list({ maxResults: 1 });

  const html = HtmlService.createHtmlOutputFromFile('Loading')
    .setWidth(480).setHeight(480);
  SpreadsheetApp.getUi().showModalDialog(html, ' ');
}

function doRefresh(vanSidebar) {
  const ss = SpreadsheetApp.getActiveSpreadsheet();

  const teamsSheet = ss.getSheetByName(SHEET_TEAMS);
  if (!teamsSheet) { SpreadsheetApp.getUi().alert('❌ Tab "Teams" not found.'); return; }

  const data = teamsSheet.getDataRange().getValues();
  const teams = {};
  data.forEach(rij => {
    const [team, naam, email] = rij.map(v => v.toString().trim());
    if (email && email.includes("@")) {
      if (!teams[team]) teams[team] = [];
      teams[team].push({ naam, email });
    }
  });

  const teamNamen = Object.keys(teams);
  if (teamNamen.length === 0) { SpreadsheetApp.getUi().alert('❌ No teams found.'); return; }

  const vandaag = new Date();
  const maandGroepen = bepaalMaanden(vandaag).map(maand => ({ maand, weken: buildMaandWeken(maand) }));

  // Column layout: Name | [Total | weeks...] per month — the total column
  // precedes its month's weeks, and doubles as the divider between months.
  let kolIdx = 2;
  maandGroepen.forEach(groep => {
    groep.totalKolIdx = kolIdx; kolIdx++;
    groep.startKolIdx = kolIdx; kolIdx += groep.weken.length;
  });
  const aantalKolommen = kolIdx - 1;

  // ONE API call per person for the full period
  const alleEmails = [...new Set(teamNamen.flatMap(t => teams[t].map(l => l.email)))];
  const periodeStart = maandGroepen[0].weken[0].week.start;
  const periodeEinde = new Date(maandGroepen[maandGroepen.length - 1].weken.slice(-1)[0].week.einde);
  periodeEinde.setHours(23, 59, 59, 999);

  const afwezigCache = {};
  const kalenderFouten = {};
  alleEmails.forEach(email => {
    const { dagData, fout } = haalAfwezighedenOp(email, periodeStart, periodeEinde);
    afwezigCache[email] = dagData;
    if (fout) kalenderFouten[email] = fout;
  });

  let sheet = ss.getSheetByName(SHEET_DAGEN);
  if (!sheet) sheet = ss.insertSheet(SHEET_DAGEN);
  sheet.getImages().forEach(img => img.remove());
  sheet.clearContents();
  sheet.clearFormats();

  sheet.setColumnWidth(1, 200);
  maandGroepen.forEach(groep => {
    sheet.setColumnWidth(groep.totalKolIdx, 90);
    for (let k = groep.startKolIdx; k < groep.startKolIdx + groep.weken.length; k++) sheet.setColumnWidth(k, 120);
  });

  // Main title
  sheet.getRange(1, 1, 1, aantalKolommen).merge();
  sheet.getRange(1, 1)
    .setValue("📅  Team Days Worked — updated on " + formatDatum(vandaag))
    .setFontFamily("Montserrat").setFontSize(13).setFontWeight("bold")
    .setBackground(GRIJS).setFontColor("#ffffff")
    .setHorizontalAlignment("center");
  sheet.setRowHeight(1, 36);

  let rij = 2;

  teamNamen.forEach(teamNaam => {
    const leden = teams[teamNaam];
    const teamKleurWaarde = teamKleur(teamNaam);
    const teamHeaderRij = rij;

    // Team header — one colored bar spanning the full width
    sheet.getRange(rij, 1, 1, aantalKolommen).merge()
      .setValue(teamNaam)
      .setFontFamily("Montserrat").setFontSize(11).setFontWeight("bold")
      .setBackground(teamKleurWaarde).setFontColor("#ffffff")
      .setHorizontalAlignment("left");
    sheet.setRowHeight(rij, 28);
    rij++;

    // Month header row — "Total" above each total column, month name merged over its weeks
    sheet.getRange(rij, 1).setValue("").setBackground(BLAUW_LICHT);
    maandGroepen.forEach(groep => {
      sheet.getRange(rij, groep.totalKolIdx)
        .setValue("Total")
        .setFontFamily("Montserrat").setFontWeight("bold").setFontSize(10)
        .setBackground(BLAUW_LICHT).setFontColor(GRIJS).setHorizontalAlignment("center");

      const kolBreedte = groep.weken.length;
      const cel = sheet.getRange(rij, groep.startKolIdx, 1, kolBreedte);
      // Break apart first: on a range that's already fully merged from an
      // earlier run, merge() is a silent no-op, leaving a stale merge in place.
      cel.breakApart();
      if (kolBreedte > 1) cel.merge();
      // setRichTextValue instead of setValue: a bare month name like
      // "September" or "November" is also a valid Dutch month name, and
      // Range.setValue() unconditionally auto-parses date-looking strings
      // and silently converts them to a date — redisplayed in the
      // spreadsheet's locale (lowercase Dutch month names) — regardless of
      // the cell's number format. Rich text bypasses that parser entirely.
      cel.setRichTextValue(SpreadsheetApp.newRichTextValue().setText(groep.maand.naam).build())
        .setFontFamily("Montserrat").setFontWeight("bold").setFontSize(10)
        .setBackground(BLAUW_LICHT).setFontColor(GRIJS).setHorizontalAlignment("center");
    });
    sheet.setRowHeight(rij, 22);
    rij++;

    // Week headers — date range + working days count
    const headers = new Array(aantalKolommen).fill("");
    headers[0] = "Team member";
    maandGroepen.forEach(groep => {
      groep.weken.forEach((wk, i) => {
        headers[groep.startKolIdx - 1 + i] = `${formatDatumKort(wk.week.start)} – ${formatDatumKort(wk.week.einde)}\n(${wk.werkdagen})`;
      });
    });
    sheet.getRange(rij, 1, 1, aantalKolommen).setValues([headers])
      .setFontFamily("Montserrat").setFontWeight("normal")
      .setBackground(BLAUW_LICHT).setFontColor(GRIJS)
      .setHorizontalAlignment("center").setWrap(true);
    sheet.getRange(rij, 1).setFontWeight("bold");
    sheet.setRowHeight(rij, 44);
    rij++;

    // Team members
    leden.forEach(lid => {
      const dagCache = afwezigCache[lid.email] || {};
      const heeftFout = !!kalenderFouten[lid.email];
      const rowData = new Array(aantalKolommen).fill("");
      const kleurCellen = [];

      maandGroepen.forEach(groep => {
        let gewerktTotaal = 0;
        let werkdagenTotaal = 0;
        groep.weken.forEach((wk, i) => {
          const kolIdx = groep.startKolIdx + i;
          werkdagenTotaal += wk.werkdagen;
          if (heeftFout) {
            rowData[kolIdx - 1] = "⚠️ no access";
            return;
          }
          const afwezig = berekenAfwezigInWeek(dagCache, wk.week.start, wk.week.einde, groep.maand.start, groep.maand.einde);
          const gewerkt = Math.max(0, Math.round((wk.werkdagen - afwezig) * 4) / 4);
          gewerktTotaal += gewerkt;
          const label = Number.isInteger(gewerkt) ? gewerkt : gewerkt.toFixed(2).replace('.', ',');
          rowData[kolIdx - 1] = `${label} / ${wk.werkdagen}`;
          kleurCellen.push({ kolIdx, ratio: wk.werkdagen > 0 ? gewerkt / wk.werkdagen : 1 });
        });

        if (heeftFout) {
          rowData[groep.totalKolIdx - 1] = "⚠️ no access";
        } else {
          const totLabel = Number.isInteger(gewerktTotaal) ? gewerktTotaal : gewerktTotaal.toFixed(2).replace('.', ',');
          rowData[groep.totalKolIdx - 1] = `${totLabel} / ${werkdagenTotaal}`;
          kleurCellen.push({ kolIdx: groep.totalKolIdx, ratio: werkdagenTotaal > 0 ? gewerktTotaal / werkdagenTotaal : 1 });
        }
      });

      sheet.getRange(rij, 1, 1, aantalKolommen).setValues([rowData]).setFontFamily("Montserrat");
      // Name as a link straight to that person's Google Calendar, for a quick cross-check
      const calUrl = "https://calendar.google.com/calendar/u/0/r?cid=" + encodeURIComponent(lid.email);
      sheet.getRange(rij, 1).setFormula(`=HYPERLINK("${calUrl}";"${lid.naam.replace(/"/g, '""')}")`);
      sheet.getRange(rij, 2, 1, aantalKolommen - 1).setHorizontalAlignment("center");

      if (heeftFout) {
        sheet.getRange(rij, 1, 1, aantalKolommen).setBackground("#fff3e0").setFontColor("#e65100");
      } else {
        kleurCellen.forEach(({ kolIdx, ratio }) => {
          const cel = sheet.getRange(rij, kolIdx);
          if (ratio >= 1) cel.setBackground("#e6f4ea");
          else if (ratio >= 0.5) cel.setBackground("#fef9c3");
          else cel.setBackground("#fce8e6");
        });
      }
      rij++;
    });

    // Vertical divider between month blocks — right border on the last week
    // column of every month group except the last (the next group's Total
    // column already reads as a visual break, this reinforces it).
    const rijTeamEinde = rij - 1;
    maandGroepen.slice(0, -1).forEach(groep => {
      const laatsteWeekKol = groep.startKolIdx + groep.weken.length - 1;
      sheet.getRange(teamHeaderRij, laatsteWeekKol, rijTeamEinde - teamHeaderRij + 1, 1)
        .setBorder(null, null, null, true, null, null, "#888888", SpreadsheetApp.BorderStyle.SOLID_MEDIUM);
    });

    // Empty row
    sheet.setRowHeight(rij, 16);
    rij++;
  });

  // Errors at the bottom
  const foutenLijst = Object.entries(kalenderFouten);
  if (foutenLijst.length > 0) {
    rij++;
    sheet.getRange(rij, 1, 1, aantalKolommen).merge();
    sheet.getRange(rij, 1)
      .setValue("⚠️ Calendar access errors — check permissions for the accounts below")
      .setFontWeight("bold").setBackground("#fff3e0").setFontColor("#e65100");
    rij++;
    foutenLijst.forEach(([email, fout]) => {
      sheet.getRange(rij, 1, 1, aantalKolommen).merge();
      sheet.getRange(rij, 1).setValue(`${email}: ${fout}`)
        .setBackground("#fff8f0").setFontColor("#bf360c");
      rij++;
    });
    Logger.log(`Calendar access errors: ${foutenLijst.map(([e]) => e).join(', ')}`);
  }

  if (!vanSidebar) ss.toast('✅ Days worked updated!', '🗓️ Team Days Worked', 5);
}

// ============================================================
// ONE API call per person — returns data + any error
// ============================================================
function haalAfwezighedenOp(email, periodeStart, periodeEinde) {
  const dagData = {};
  try {
    const response = Calendar.Events.list(email, {
      timeMin: periodeStart.toISOString(),
      timeMax: periodeEinde.toISOString(),
      eventTypes: ["outOfOffice"],
      singleEvents: true,
    });
    (response.items || []).forEach(event => {
      const isAllDay = !!event.start.date && !event.start.dateTime;
      // All-day events carry bare "YYYY-MM-DD" strings, which JS Date always
      // parses as UTC midnight. In any timezone ahead of UTC (e.g. Brussels)
      // that lands just after local midnight, so the day-walking loop below
      // would count one extra calendar day per event. Parse those as local
      // dates directly to avoid the shift; timed events already carry a
      // proper offset in their dateTime string, so they're safe as-is.
      const evStart = event.start.dateTime ? new Date(event.start.dateTime) : parseLokaleDatum(event.start.date);
      const evEinde = event.end.dateTime ? new Date(event.end.dateTime) : parseLokaleDatum(event.end.date);
      const cursor = new Date(evStart); cursor.setHours(0, 0, 0, 0);
      while (cursor < evEinde) {
        const dag = cursor.getDay();
        if (dag !== 0 && dag !== 6) {
          const key = cursor.toDateString();
          if (isAllDay) {
            dagData[key] = WERKUREN_PER_DAG;
          } else {
            const dagStart = new Date(cursor); dagStart.setHours(9, 0, 0, 0);
            const dagEinde = new Date(cursor); dagEinde.setHours(17, 0, 0, 0);
            const overlapMs = Math.min(evEinde, dagEinde) - Math.max(evStart, dagStart);
            if (overlapMs > 0) {
              dagData[key] = Math.min(WERKUREN_PER_DAG, (dagData[key] || 0) + overlapMs / (1000 * 60 * 60));
            }
          }
        }
        cursor.setDate(cursor.getDate() + 1);
      }
    });
    return { dagData, fout: null };
  } catch (e) {
    Logger.log("Error for " + email + ": " + e.message);
    return { dagData, fout: e.message };
  }
}

// ============================================================
// MENU
// ============================================================
function onOpen() {
  SpreadsheetApp.getUi()
    .createMenu("📅 Days Worked")
    .addItem("🔄 Refresh days worked", "refreshDagenGewerkt")
    .addItem("⚙️ Install auto-refresh trigger", "installeerTrigger")
    .addToUi();
}

// Run once via the menu to install the installable on-open trigger.
// An installable trigger CAN call the Calendar API; simple onOpen() cannot.
function installeerTrigger() {
  ScriptApp.getProjectTriggers()
    .filter(t => t.getHandlerFunction() === 'doRefresh')
    .forEach(t => ScriptApp.deleteTrigger(t));

  ScriptApp.newTrigger('refreshDagenGewerkt')
    .forSpreadsheet(SpreadsheetApp.getActive())
    .onOpen()
    .create();

  SpreadsheetApp.getUi().alert('✅ Auto-refresh installed — days worked will load automatically on open.');
}

// ============================================================
// HELPERS
// ============================================================
function getMaandagVanWeek(datum) {
  const d = new Date(datum);
  const diff = d.getDay() === 0 ? -6 : 1 - d.getDay();
  d.setDate(d.getDate() + diff);
  d.setHours(0, 0, 0, 0);
  return d;
}

function formatDatum(d) {
  return d.toLocaleDateString("en-GB", { weekday: "long", year: "numeric", month: "long", day: "numeric" });
}

function formatDatumKort(d) {
  return d.toLocaleDateString("en-GB", { day: "numeric", month: "short" });
}

// Fixed array instead of toLocaleDateString(..., { month: "long" }): Apps
// Script's V8 sandbox has been observed returning lowercase month names for
// some months (e.g. "september" instead of "September") even with the
// "en-GB" locale explicitly set, unlike plain Node.js. A hardcoded list
// sidesteps that runtime's locale/ICU data entirely.
const MAAND_NAMEN = ["January", "February", "March", "April", "May", "June",
  "July", "August", "September", "October", "November", "December"];

function formatMaandNaam(d) {
  return `${MAAND_NAMEN[d.getMonth()]} ${d.getFullYear()}`;
}

function parseLokaleDatum(dateStr) {
  const [j, m, d] = dateStr.split('-').map(Number);
  return new Date(j, m - 1, d);
}

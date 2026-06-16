/****************************************************************
 * INVLOEDSREGISTER — Apps Script backend (Code.gs)
 * Hoort bij invloedsregister.html (Huurdersvereniging Brederode).
 *
 * Doet twee dingen die een statische pagina niet kan:
 *  1) De Google Sheet als opslag achter de pagina (lezen/schrijven).
 *  2) Eén keer per dag een ALERT-MAIL sturen voor meldingen die de
 *     deadline naderen of overschreden zijn.
 *
 * --- INSTALLEREN (eenmalig) ---
 *  1. Open de Google Sheet "Register".
 *  2. Menu: Extensies > Apps Script. Plak deze code, sla op.
 *  3. Pas hieronder SHEET_NAAM en STANDAARD_ONTVANGER aan.
 *  4. Draai eenmaal de functie  setup()  (kies 'setup' bovenin, klik Uitvoeren,
 *     geef toestemming). Dit zet de kopregel én de dagelijkse trigger klaar.
 *  5. Menu: Implementeren > Nieuwe implementatie > type 'Web-app'.
 *       Uitvoeren als: ikzelf.  Toegang: iedereen.
 *     Kopieer de /exec-URL en zet die in invloedsregister.html bij
 *     CONFIG.BACKEND_URL.
 ****************************************************************/

const SHEET_NAAM        = "Register";
const STANDAARD_ONTVANGER = "secretaris@huurdersverenigingbrederode.nl"; // krijgt altijd het overzicht
const ALERT_DAGEN       = 21;   // meldt wat ≤ dit aantal dagen tot einddatum staat
const ALERT_FRACTIE     = 0.8;  // …of waarvan de termijn ≥ 80% verstreken is
const APP_TITEL         = "Invloedsregister — Huurdersvereniging Brederode";

const KOLOMMEN = ["id","datum","type","onderwerp","omschrijving","melder",
                  "invoerder","afspraak","termijn","einddatum","status",
                  "betrokkenen","toelichting"];

/* ---------- web-app endpoints ---------- */
function doGet(e){
  const actie = (e && e.parameter && e.parameter.action) || "list";
  if (actie === "list") return json({rows: leesRijen()});
  return json({error:"onbekende actie"});
}
function doPost(e){
  const body = JSON.parse(e.postData.contents || "{}");
  if (body.action === "save")   { bewaarRij(body.row); return json({ok:true}); }
  if (body.action === "delete") { verwijderRij(body.id); return json({ok:true}); }
  return json({error:"onbekende actie"});
}
function json(obj){
  return ContentService.createTextOutput(JSON.stringify(obj))
         .setMimeType(ContentService.MimeType.JSON);
}

/* ---------- sheet-laag ---------- */
function sheet(){ return SpreadsheetApp.getActive().getSheetByName(SHEET_NAAM); }

function leesRijen(){
  const sh = sheet(); const data = sh.getDataRange().getValues();
  if (data.length < 2) return [];
  const head = data[0].map(String);
  return data.slice(1).filter(r => r.join("")!=="").map(r=>{
    const o = {}; head.forEach((h,i)=>{
      let v = r[i];
      if (v instanceof Date) v = isoTekst(v);            // datum -> 'yyyy-mm-dd'
      o[h] = v==null ? "" : String(v);
    });
    return o;
  });
}
function isoTekst(d){ const p=n=>String(n).padStart(2,"0");
  return d.getFullYear()+"-"+p(d.getMonth()+1)+"-"+p(d.getDate()); }
function bewaarRij(row){
  const sh = sheet(); const data = sh.getDataRange().getValues();
  const ids = data.map(r=> String(r[0]));
  const idx = ids.indexOf(String(row.id));
  const rij = KOLOMMEN.map(k=> row[k]!=null ? row[k] : "");
  if (idx > 0) sh.getRange(idx+1,1,1,KOLOMMEN.length).setValues([rij]);
  else         sh.appendRow(rij);
}
function verwijderRij(id){
  const sh = sheet(); const data = sh.getDataRange().getValues();
  for (let i=data.length-1;i>=1;i--){ if (String(data[i][0])===String(id)) sh.deleteRow(i+1); }
}

/* ---------- dagelijkse ALERT-MAIL ---------- */
function sendDeadlineAlerts(){
  const rijen = leesRijen();
  const vandaag = new Date(); vandaag.setHours(0,0,0,0);
  const naderend = [];

  rijen.forEach(r=>{
    if (!r.einddatum) return;
    if (r.status === "Afgerond" || r.status === "Vervallen") return;
    const eind  = parseISO(r.einddatum);
    const start = parseISO(r.datum) || eind;
    if (!eind) return;
    const dagen = Math.round((eind - vandaag)/86400000);
    const totaal = Math.max(1, Math.round((eind - start)/86400000));
    const frac = Math.min(1, Math.max(0, (vandaag - start)/86400000 / totaal));
    if (dagen < 0 || dagen <= ALERT_DAGEN || frac >= ALERT_FRACTIE){
      naderend.push({r, dagen, eind});
    }
  });

  if (!naderend.length) return;
  naderend.sort((a,b)=> a.dagen - b.dagen);

  // ontvangers: standaard-adres + alle 'betrokkenen' van naderende meldingen
  const ontvangers = new Set([STANDAARD_ONTVANGER]);
  naderend.forEach(x=> String(x.r.betrokkenen||"").split(/[,;]/)
    .map(s=>s.trim()).filter(s=>s.includes("@")).forEach(a=>ontvangers.add(a)));

  const regels = naderend.map(x=>{
    const t = x.dagen < 0 ? `${Math.abs(x.dagen)} dagen VERLOPEN`
            : x.dagen === 0 ? "vandaag de deadline"
            : `nog ${x.dagen} dagen`;
    return `• ${x.r.onderwerp} — ${t} (einddatum ${nl(x.eind)}, status: ${x.r.status||"—"})`;
  }).join("\n");

  const body =
`Beste lezer,

De volgende meldingen uit het ${APP_TITEL} naderen of overschrijden hun afgesproken einddatum:

${regels}

Open het register om de afspraken en het vervolg te bekijken.

— Automatisch bericht van het Invloedsregister`;

  MailApp.sendEmail({
    to: Array.from(ontvangers).join(","),
    subject: `Invloedsregister — ${naderend.length} melding(en) naderen de deadline`,
    body: body
  });
}

/* ---------- setup (eenmalig draaien) ---------- */
function setup(){
  const ss = SpreadsheetApp.getActive();
  let sh = ss.getSheetByName(SHEET_NAAM);
  if (!sh) sh = ss.insertSheet(SHEET_NAAM);
  if (sh.getLastRow() === 0) sh.appendRow(KOLOMMEN);          // kopregel
  // datum-kolommen (datum, einddatum) als platte tekst -> Sheets maakt er geen datumwaarde van
  const cDatum = KOLOMMEN.indexOf("datum")+1, cEind = KOLOMMEN.indexOf("einddatum")+1;
  sh.getRange(1, cDatum, sh.getMaxRows(), 1).setNumberFormat("@");
  sh.getRange(1, cEind,  sh.getMaxRows(), 1).setNumberFormat("@");
  // dagelijkse trigger (08:00) — voorkom dubbele
  ScriptApp.getProjectTriggers().forEach(t=>{
    if (t.getHandlerFunction()==="sendDeadlineAlerts") ScriptApp.deleteTrigger(t);
  });
  ScriptApp.newTrigger("sendDeadlineAlerts").timeBased().everyDays(1).atHour(8).create();
  SpreadsheetApp.getUi && SpreadsheetApp.getActive().toast("Setup klaar: kopregel + dagelijkse alert (08:00).");
}

/* ---------- datum-helpers (ISO yyyy-mm-dd) ---------- */
function parseISO(s){ if(!s) return null; const p=String(s).split("-"); 
  return p.length===3 ? new Date(+p[0], +p[1]-1, +p[2]) : null; }
function nl(d){ const p=n=>String(n).padStart(2,"0");
  return p(d.getDate())+"-"+p(d.getMonth()+1)+"-"+d.getFullYear(); }

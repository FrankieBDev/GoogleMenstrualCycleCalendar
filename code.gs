function onOpen() {
  SpreadsheetApp.getUi()
    .createMenu('Cycle Tools')
    .addItem('Predict from Settings only', 'predictFromSettingsOnly')
    .addSeparator()
    .addItem('Delete all created events', 'deleteAllCreatedEvents')
    .addItem('Install monthly trigger', 'installMonthlyTrigger')
    .addToUi();
}

// ===== Main actions =====

function predictFromSettingsOnly() {
  const cfg = readConfig_();
  const anchorStr = cfg.anchorStartStr;
  if (!anchorStr) throw new Error('Config.AnchorStartDate is empty. Set this to your last real Day 1.');
  const anchor = new Date(anchorStr);
  if (isNaN(anchor.getTime())) throw new Error('AnchorStartDate is not a valid date (use YYYY-MM-DD).');

  const cal = getOrCreateCalendar_(cfg.calName);
  deleteCycleEvents_(cal, cfg.titlePrefix, { fromDaysAgo: 0, toDaysAhead: 540, wipeAll: false });
  createPredictions_(cal, anchor, cfg);
  SpreadsheetApp.getActive().toast('Future predictions created.');
}

function deleteAllCreatedEvents() {
  const cfg = readConfig_();
  const cal = getOrCreateCalendar_(cfg.calName);
  deleteCycleEvents_(cal, cfg.titlePrefix, { fromDaysAgo: 3650, toDaysAhead: 1825, wipeAll: true });
  SpreadsheetApp.getActive().toast('All created events removed.');
}

function installMonthlyTrigger() {
  ScriptApp.newTrigger('predictFromSettingsOnly')
    .timeBased()
    .onMonthDay(1)
    .atHour(6)
    .create();
  SpreadsheetApp.getActive().toast('Monthly trigger installed for the 1st at 06:00.');
}

// ===== Core logic =====

function createPredictions_(cal, anchorStart, cfg) {
  const start = stripTime_(anchorStart);
  const endLimit = monthsFrom_(start, cfg.monthsAhead);
  let cycleStart = new Date(start);

  while (cycleStart < endLimit) {
    // Pre-bleed week (continuous 7-day event)
    const preWeekStart = addDays_(cycleStart, -7);
    const preWeekEnd = addDays_(cycleStart, 0); // ends day before Day 1
    cal.createAllDayEvent(`${cfg.titlePrefix} ⚡️💜 Pre-bleed week`, preWeekStart, preWeekEnd, {
      description: '💜 Be extra kind to yourself this week 💜',
      visibility: CalendarApp.Visibility.PRIVATE
    });

    // PMS Day
    const pmsDayStart = addDays_(cycleStart, -1);
    const pmsDayEnd = addDays_(cycleStart, 0);
    cal.createAllDayEvent(`${cfg.titlePrefix} 🌧😡 PMS Day`, pmsDayStart, pmsDayEnd, {
      description: 'Reminder: symptoms may peak today — rest & self-care 💜',
      visibility: CalendarApp.Visibility.PRIVATE
    });

    // Period days
    for (let d = 0; d < cfg.periodLen; d++) {
      const dayStart = addDays_(cycleStart, d);
      const dayEnd = addDays_(dayStart, 1);
      cal.createAllDayEvent(`${cfg.titlePrefix} 🩸 Period Day ${d + 1}`, dayStart, dayEnd, {
        description: 'Predicted',
        visibility: CalendarApp.Visibility.PRIVATE
      });
    }

    // Ovulation estimate
    const ovu = addDays_(cycleStart, Math.max(1, Math.round(cfg.avgCycle - cfg.luteal)));
    const ovuEnd = addDays_(ovu, 1);
    cal.createAllDayEvent(`${cfg.titlePrefix} 🥚 Ovulation`, ovu, ovuEnd, {
      description: 'Predicted fertile peak',
      visibility: CalendarApp.Visibility.PRIVATE
    });

    // Fertile window
    const fertileStart = addDays_(ovu, -5);
    cal.createAllDayEvent(`${cfg.titlePrefix} 🌸 Fertile window`, fertileStart, ovuEnd, {
      description: 'Predicted fertile window',
      visibility: CalendarApp.Visibility.PRIVATE
    });

    cycleStart = addDays_(cycleStart, cfg.avgCycle);
  }
}

// ===== Delete helper (strong prefix match) =====
function deleteCycleEvents_(cal, prefix, opts) {
  const fromDaysAgo = opts.fromDaysAgo || 3650;   // 10 years back
  const toDaysAhead = opts.toDaysAhead || 1825;   // 5 years ahead

  const today = stripTime_(new Date());
  const start = addDays_(today, -fromDaysAgo);
  const end   = addDays_(today,  toDaysAhead);

  const evs = cal.getEvents(start, end);
  let deleted = 0;
  evs.forEach(ev => {
    const title = ev.getTitle() || '';
    if (title.startsWith(prefix)) {
      try { ev.deleteEvent(); deleted++; } catch(e){}
    }
  });
  SpreadsheetApp.getActive().toast(`Deleted ${deleted} events with prefix ${prefix}`);
}

// ===== Config + utils =====
function readConfig_() {
  const s = SpreadsheetApp.getActive().getSheetByName('Config');
  if (!s) throw new Error('Missing "Config" sheet');
  const lastRow = s.getLastRow();
  if (lastRow < 2) throw new Error('Config sheet has no rows under the header.');

  const pairs = s.getRange(2, 1, lastRow - 1, 2).getValues();
  const map = {};
  pairs.forEach(([key, val]) => {
    const k = String(key || '').trim();
    if (k) map[k] = val;
  });

  return {
    calName: String(map['CalendarName'] || 'Cycle'),
    avgCycle: Number(map['AverageCycleDays']),
    periodLen: Number(map['AveragePeriodDays']),
    luteal: Number(map['LutealDays'] || 14),
    monthsAhead: Number(map['PredictMonthsAhead'] || 6),
    anchorStartStr: String(map['AnchorStartDate'] || ''),
    titlePrefix: String(map['TitlePrefix'] || '[Cycle]')
  };
}

function getOrCreateCalendar_(name) {
  const cals = CalendarApp.getCalendarsByName(name);
  return cals.length ? cals[0] : CalendarApp.createCalendar(name);
}

// Date helpers
function stripTime_(d) { return new Date(d.getFullYear(), d.getMonth(), d.getDate()); }
function addDays_(d, n) { const x = new Date(d); x.setDate(x.getDate() + n); return x; }
function monthsFrom_(d, m) { const x = new Date(d); x.setMonth(x.getMonth() + m); return x; }

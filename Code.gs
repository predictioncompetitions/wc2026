const SPREADSHEET_ID = '1vEeV9Lp6myZrlsVd0UQoa7Mry3vFHt7-y8fpwEPPuaU';
const SETTINGS_SHEET = 'SETTINGS';
const GROUP_TEAMS_SHEET = 'GROUP_TEAMS';
const ENTRIES_RAW_SHEET = 'ENTRIES_RAW';
const FIXTURES_SHEET = 'FIXTURES';

/*
  Create a Google Drive folder for saved entry PDFs.
  Put that folder ID below.
*/
const PDF_FOLDER_ID = '1WjxPG9HyeZLj9pGne5PlqWZMaE8p9Ogj';

/*
  Optional image URLs for the PDF template.
  Leave blank if not using yet.
*/
const PDF_HEADER_IMAGE_URL = '';
const PDF_LOGO_IMAGE_URL = '';

/*
  Revolut payment link shown in the confirmation email.
  Replace with the real link once available. Leave blank to hide the button entirely.
*/
const REVOLUT_PAYMENT_LINK = 'https://revolut.me/YOUR-LINK-HERE';
const REVOLUT_ENTRY_FEE_LABEL = '€10'; // Shown on the button and in the email copy. Set to '' to hide.

const ROUND_OF32_MATCH_ORDER = [73, 74, 75, 76, 77, 78, 79, 80, 81, 82, 83, 84, 85, 86, 87, 88];
const ROUND_OF16_MATCH_ORDER = [89, 90, 91, 92, 93, 94, 95, 96];
const QUARTER_FINAL_MATCH_ORDER = [97, 98, 99, 100];
const SEMI_FINAL_MATCH_ORDER = [101, 102];
const FINAL_MATCH_NO = 104;

function doGet() {
  return HtmlService.createHtmlOutputFromFile('Index')
    .setTitle('World Cup 2026 Prediction')
    .setXFrameOptionsMode(HtmlService.XFrameOptionsMode.ALLOWALL);
}

/* =========================================================
   INITIAL LOAD
========================================================= */

function getInitialData() {
  try {
    const settings = getSettingsMap_();
    const groups = getGroupTeams_();

    let entryStatus = String(settings.ENTRY_STATUS || 'OPEN').trim().toUpperCase();

    const deadlineRaw = settings.ENTRY_DEADLINE;
    let deadlineIso = '';
    if (deadlineRaw) {
      const d = new Date(deadlineRaw);
      if (!isNaN(d.getTime())) {
        deadlineIso = d.toISOString();
        if (new Date() > d) {
          entryStatus = 'CLOSED';
        }
      }
    }

    return {
      ok: true,
      competitionName: settings.COMPETITION_NAME || 'World Cup 2026 Predictor',
      entryStatus,
      entryDeadline: deadlineIso,
      maxEntriesPerEmail: Number(settings.MAX_ENTRIES_PER_EMAIL || 1),
      formVersion: settings.FORM_VERSION || 'V1',
      groups
    };
  } catch (error) {
    return {
      ok: false,
      message: error && error.message ? error.message : 'Failed to load initial data.'
    };
  }
}

/* =========================================================
   ENTRY SUBMIT
========================================================= */

function submitPhase1Entry(payload) {
  try {
    validateEntryWindow_();

    if (!payload || typeof payload !== 'object') {
      throw new Error('No entry data was received.');
    }

    const cleaned = normalizePayload_(payload);
    cleaned.FULL_NAME = getUniqueDisplayName_(cleaned.FULL_NAME);
    cleaned.DISPLAY_NAME = cleaned.FULL_NAME;

    validatePayload_(cleaned);
    validateKnockoutPayload_(payload.KNOCKOUT_DATA);
    checkDuplicateEmail_(cleaned.EMAIL);

    const timestamp = new Date();

    const headers = getSheetHeaders_(ENTRIES_RAW_SHEET);
    const rowObject = {};

    headers.forEach(header => {
      rowObject[header] = '';
    });

    rowObject.ENTRY_ID = '';
    rowObject.TIMESTAMP = timestamp;
    rowObject.FULL_NAME = cleaned.FULL_NAME;
    rowObject.EMAIL = cleaned.EMAIL;
    rowObject.DISPLAY_NAME = cleaned.DISPLAY_NAME || cleaned.FULL_NAME;
    rowObject.ENTRY_STATUS = 'LOCKED';
    rowObject.FORM_VERSION = cleaned.FORM_VERSION || 'V1';
    rowObject.VALIDATION_STATUS = 'VALID';

    Object.keys(cleaned.GROUPS).forEach(group => {
      rowObject[`${group}_1ST`] = cleaned.GROUPS[group]['1ST'];
      rowObject[`${group}_2ND`] = cleaned.GROUPS[group]['2ND'];
      rowObject[`${group}_3RD`] = cleaned.GROUPS[group]['3RD'];
    });

    cleaned.THIRD_RANK.forEach((team, index) => {
      rowObject[`THIRD_RANK_${index + 1}`] = team;
    });

    applyKnockoutToRowObject_(rowObject, payload.KNOCKOUT_DATA);

    if ('PDF_SENT' in rowObject) rowObject.PDF_SENT = 'NO';
    if ('PDF_SENT_AT' in rowObject) rowObject.PDF_SENT_AT = '';
    if ('PDF_FILE_URL' in rowObject) rowObject.PDF_FILE_URL = '';
    if ('NOTES' in rowObject) rowObject.NOTES = '';

    const rowValues = headers.map(header => rowObject[header] !== undefined ? rowObject[header] : '');

    const sheet = getSheet_(ENTRIES_RAW_SHEET);
    sheet.appendRow(rowValues);

    const rowNumber = sheet.getLastRow();
    const entryIdFormatted = generateFormattedEntryId_(cleaned.FULL_NAME, rowNumber);

    const headerMap = getHeaderIndexMap_(ENTRIES_RAW_SHEET);

    if (headerMap.ENTRY_ID) {
      sheet.getRange(rowNumber, headerMap.ENTRY_ID).setValue(entryIdFormatted);
    }

    try {
      const entryRowObject = getEntryRowObjectByRowNumber_(rowNumber);
      entryRowObject.ENTRY_ID = entryIdFormatted;

      const pdfData = buildPdfTemplateData_(entryRowObject);
      const pdfBlob = createEntryPdfBlob_(pdfData);
      const savedFile = savePdfToDrive_(pdfBlob);
      sendEntryConfirmationEmail_(pdfData, pdfBlob);

      updatePdfStatus_(rowNumber, 'YES', savedFile.getUrl(), '');
    } catch (pdfError) {
      updatePdfStatus_(
        rowNumber,
        'ERROR',
        '',
        pdfError && pdfError.message ? pdfError.message : 'PDF/email step failed.'
      );
    }

    return {
      ok: true,
      message: `Entry submitted successfully. We will email you a copy of your entry and your Entry ID is ${entryIdFormatted}.`,
      entryId: entryIdFormatted
    };
  } catch (error) {
    return {
      ok: false,
      message: error && error.message ? error.message : 'Submission failed.'
    };
  }
}

/* =========================================================
   KNOCKOUT BRACKET GENERATION
========================================================= */

function generateKnockoutBracket(payload) {
  try {
    if (!payload || typeof payload !== 'object') {
      throw new Error('No entry data was received.');
    }

    const cleaned = normalizePayload_(payload);
    validatePayload_(cleaned);

    const bracket = buildKnockoutBracket_(cleaned);

    return {
      ok: true,
      bracket
    };
  } catch (error) {
    return {
      ok: false,
      message: error && error.message ? error.message : 'Failed to generate knockout bracket.'
    };
  }
}

function buildKnockoutBracket_(cleaned) {
  const standings = buildStandingsFromCleaned_(cleaned);
  const top8Thirds = getTop8ThirdPlacedTeams_(cleaned, standings);
  const allocation = allocateThirdPlaceTeams_(top8Thirds);

  const roundOf32 = buildRoundOf32Fixtures_(standings, allocation.byMatch);
  const roundOf16 = buildFutureRoundFixtures_('Round of 16', ROUND_OF16_MATCH_ORDER);
  const quarterFinals = buildFutureRoundFixtures_('Quarter-final', QUARTER_FINAL_MATCH_ORDER);
  const semiFinals = buildFutureRoundFixtures_('Semi-final', SEMI_FINAL_MATCH_ORDER);
  const final = buildFinalFixture_();
  const progression = getKnockoutProgression_();

  return {
    standings,
    top8Thirds,
    thirdPlaceAssignments: allocation.byMatch,
    roundOf32,
    roundOf16,
    quarterFinals,
    semiFinals,
    final,
    progression
  };
}

function buildStandingsFromCleaned_(cleaned) {
  const result = {
    winners: {},
    runnersUp: {},
    thirds: {}
  };

  Object.keys(cleaned.GROUPS).sort().forEach(group => {
    result.winners[group] = cleaned.GROUPS[group]['1ST'];
    result.runnersUp[group] = cleaned.GROUPS[group]['2ND'];
    result.thirds[group] = cleaned.GROUPS[group]['3RD'];
  });

  return result;
}

function getTop8ThirdPlacedTeams_(cleaned, standings) {
  const thirdTeamToGroup = {};

  Object.keys(standings.thirds).forEach(group => {
    thirdTeamToGroup[standings.thirds[group]] = group;
  });

  const top8 = cleaned.THIRD_RANK.slice(0, 8).map((team, index) => {
    const group = thirdTeamToGroup[team];
    if (!group) {
      throw new Error(`Could not determine the group for ranked 3rd-place team: ${team}`);
    }
    return {
      team,
      group,
      rank: index + 1
    };
  });

  if (top8.length !== 8) {
    throw new Error('Exactly 8 qualifying 3rd-place teams are required.');
  }

  return top8;
}

function allocateThirdPlaceTeams_(top8Thirds) {
  const flexibleMatches = getFlexibleThirdPlaceMatchRules_();

  if (!Array.isArray(top8Thirds) || top8Thirds.length !== 8) {
    throw new Error('Exactly 8 qualifying 3rd-place teams are required.');
  }

  const pairedThirds = top8Thirds.map((teamObj, index) => ({
    ...teamObj,
    pairIndex: Math.floor(index / 2) + 1
  }));

  const orderedMatches = flexibleMatches
    .slice()
    .sort((a, b) => {
      const aEligible = pairedThirds.filter(t => t.group !== a.winnerGroup).length;
      const bEligible = pairedThirds.filter(t => t.group !== b.winnerGroup).length;
      return aEligible - bEligible;
    });

  const assignment = backtrackAllocateThirds_(
    orderedMatches,
    pairedThirds,
    {},
    { top: {}, bottom: {} }
  );

  if (!assignment) {
    throw new Error('Could not find a valid 3rd-place allocation for this ranked top 8.');
  }

  return {
    byMatch: assignment,
    orderedMatches: flexibleMatches.map(m => m.matchNo)
  };
}

function backtrackAllocateThirds_(remainingMatches, remainingTeams, currentAssignment, sidePairUsage) {
  if (remainingMatches.length === 0) {
    return currentAssignment;
  }

  const currentMatch = remainingMatches[0];
  const nextMatches = remainingMatches.slice(1);

  const eligibleTeams = remainingTeams.filter(teamObj => {
    if (teamObj.group === currentMatch.winnerGroup) return false;
    if (sidePairUsage[currentMatch.side] && sidePairUsage[currentMatch.side][teamObj.pairIndex]) {
      return false;
    }
    return true;
  });

  if (eligibleTeams.length === 0) {
    return null;
  }

  for (let i = 0; i < eligibleTeams.length; i++) {
    const candidate = eligibleTeams[i];

    const nextAssignment = Object.assign({}, currentAssignment, {
      [currentMatch.matchNo]: candidate
    });

    const nextRemainingTeams = remainingTeams.filter(t => t.team !== candidate.team);

    const nextSidePairUsage = {
      top: Object.assign({}, sidePairUsage.top),
      bottom: Object.assign({}, sidePairUsage.bottom)
    };
    nextSidePairUsage[currentMatch.side][candidate.pairIndex] = true;

    if (!canFillRemainingSlots_(nextMatches, nextRemainingTeams, nextSidePairUsage)) {
      continue;
    }

    const solved = backtrackAllocateThirds_(
      nextMatches,
      nextRemainingTeams,
      nextAssignment,
      nextSidePairUsage
    );

    if (solved) return solved;
  }

  return null;
}

function canFillRemainingSlots_(remainingMatches, remainingTeams, sidePairUsage) {
  if (remainingMatches.length === 0) return true;

  for (let i = 0; i < remainingMatches.length; i++) {
    const match = remainingMatches[i];

    const hasEligible = remainingTeams.some(teamObj => {
      if (teamObj.group === match.winnerGroup) return false;
      if (sidePairUsage[match.side] && sidePairUsage[match.side][teamObj.pairIndex]) return false;
      return true;
    });

    if (!hasEligible) return false;
  }

  for (let i = 0; i < remainingTeams.length; i++) {
    const teamObj = remainingTeams[i];

    const canFitSomewhere = remainingMatches.some(match => {
      if (teamObj.group === match.winnerGroup) return false;
      if (sidePairUsage[match.side] && sidePairUsage[match.side][teamObj.pairIndex]) return false;
      return true;
    });

    if (!canFitSomewhere) return false;
  }

  return true;
}

function buildRoundOf32Fixtures_(standings, allocation) {
  const fixed = getFixedRoundOf32Rules_();
  const flexible = getFlexibleThirdPlaceMatchRules_();
  const fixturesMeta = getFixturesMap_();
  const fixtures = [];

  fixed.forEach(rule => {
    const meta = getFixtureMeta_(fixturesMeta, rule.matchNo, rule.date, 'Round of 32');
    fixtures.push({
      matchNo: rule.matchNo,
      date: meta.date,
      kickOffGMT: meta.kickOffGMT,
      round: meta.round,
      team1: resolveSlot_(rule.team1, standings),
      team2: resolveSlot_(rule.team2, standings)
    });
  });

  flexible.forEach(rule => {
    const assignedThird = allocation[rule.matchNo];
    if (!assignedThird) {
      throw new Error(`No 3rd-place team was assigned to Match ${rule.matchNo}.`);
    }

    const meta = getFixtureMeta_(fixturesMeta, rule.matchNo, rule.date, 'Round of 32');
    fixtures.push({
      matchNo: rule.matchNo,
      date: meta.date,
      kickOffGMT: meta.kickOffGMT,
      round: meta.round,
      team1: resolveSlot_(rule.team1, standings),
      team2: assignedThird.team,
      thirdPlaceGroup: assignedThird.group,
      thirdPlaceRank: assignedThird.rank
    });
  });

  fixtures.sort((a, b) => a.matchNo - b.matchNo);
  return fixtures;
}

function buildFutureRoundFixtures_(fallbackRoundName, orderedMatchNos) {
  const fixturesMeta = getFixturesMap_();

  return orderedMatchNos.map(matchNo => {
    const meta = getFixtureMeta_(fixturesMeta, matchNo, '', fallbackRoundName);
    return {
      matchNo: matchNo,
      date: meta.date,
      kickOffGMT: meta.kickOffGMT,
      round: meta.round
    };
  });
}

function buildFinalFixture_() {
  const fixturesMeta = getFixturesMap_();
  const meta = getFixtureMeta_(fixturesMeta, FINAL_MATCH_NO, '19 July', 'Final');

  return {
    matchNo: FINAL_MATCH_NO,
    date: meta.date,
    kickOffGMT: meta.kickOffGMT,
    round: meta.round
  };
}

function resolveSlot_(slot, standings) {
  if (slot.type === 'WINNER') return standings.winners[slot.group];
  if (slot.type === 'RUNNER_UP') return standings.runnersUp[slot.group];
  throw new Error(`Unknown slot type: ${slot.type}`);
}

/* =========================================================
   SAVE KNOCKOUT DATA TO ENTRIES_RAW
========================================================= */

function validateKnockoutPayload_(knockoutData) {
  if (!knockoutData || typeof knockoutData !== 'object') {
    throw new Error('Knockout data is missing.');
  }

  if (!Array.isArray(knockoutData.roundOf32) || knockoutData.roundOf32.length !== 16) {
    throw new Error('Round of 32 knockout data is incomplete.');
  }

  if (!Array.isArray(knockoutData.roundOf16) || knockoutData.roundOf16.length !== 8) {
    throw new Error('Round of 16 knockout data is incomplete.');
  }

  if (!Array.isArray(knockoutData.quarterFinals) || knockoutData.quarterFinals.length !== 4) {
    throw new Error('Quarter-final knockout data is incomplete.');
  }

  if (!Array.isArray(knockoutData.semiFinals) || knockoutData.semiFinals.length !== 2) {
    throw new Error('Semi-final knockout data is incomplete.');
  }

  if (!knockoutData.final || typeof knockoutData.final !== 'object') {
    throw new Error('Final knockout data is missing.');
  }

  validateRoundScores_(knockoutData.roundOf32, 'Round of 32');
  validateRoundScores_(knockoutData.roundOf16, 'Round of 16');
  validateRoundScores_(knockoutData.quarterFinals, 'Quarter-finals');
  validateRoundScores_(knockoutData.semiFinals, 'Semi-finals');
  validateRoundScores_([knockoutData.final], 'Final');

  if (!knockoutData.predictedChampion) {
    throw new Error('Predicted champion is missing.');
  }
}

function validateRoundScores_(matches, roundName) {
  matches.forEach(match => {
    if (!match.team1 || !match.team2) {
      throw new Error(`${roundName} contains an incomplete fixture.`);
    }

    if (match.score1 === '' || match.score2 === '' || match.score1 === undefined || match.score2 === undefined) {
      throw new Error(`${roundName} contains an incomplete score.`);
    }

    const n1 = Number(match.score1);
    const n2 = Number(match.score2);

    if (!Number.isInteger(n1) || !Number.isInteger(n2) || n1 < 0 || n2 < 0) {
      throw new Error(`${roundName} contains an invalid score.`);
    }

    if (n1 === n2) {
      throw new Error(`${roundName} cannot contain a draw.`);
    }

    if (!match.winner) {
      throw new Error(`${roundName} contains a match without a winner.`);
    }
  });
}

function applyKnockoutToRowObject_(rowObject, knockoutData) {
  if (!knockoutData) return;

  const top8Thirds = Array.isArray(knockoutData.top8Thirds) ? knockoutData.top8Thirds : [];
  top8Thirds.forEach((item, index) => {
    rowObject[`THIRD_Q${index + 1}`] = item.team || '';
  });

  const flexibleOrder = getFlexibleThirdPlaceMatchRules_().map(rule => rule.matchNo);
  flexibleOrder.forEach((matchNo, index) => {
    const assigned = knockoutData.thirdPlaceAssignments && knockoutData.thirdPlaceAssignments[matchNo];
    rowObject[`TS${index + 1}_TEAM`] = assigned && assigned.team ? assigned.team : '';
  });

  applyRoundMatchesToRowObject_(rowObject, knockoutData.roundOf32 || [], ROUND_OF32_MATCH_ORDER, 'R32');
  applyRoundMatchesToRowObject_(rowObject, knockoutData.roundOf16 || [], ROUND_OF16_MATCH_ORDER, 'R16');
  applyRoundMatchesToRowObject_(rowObject, knockoutData.quarterFinals || [], QUARTER_FINAL_MATCH_ORDER, 'QF');
  applyRoundMatchesToRowObject_(rowObject, knockoutData.semiFinals || [], SEMI_FINAL_MATCH_ORDER, 'SF');
  applyFinalToRowObject_(rowObject, knockoutData.final || null);

  rowObject.PREDICTED_CHAMPION = knockoutData.predictedChampion || '';
}

function applyRoundMatchesToRowObject_(rowObject, matches, orderedMatchNos, prefix) {
  orderedMatchNos.forEach((matchNo, index) => {
    const entry = findMatchByMatchNo_(matches, matchNo);
    const slot = index + 1;

    rowObject[`${prefix}_${slot}_HOME`] = entry ? entry.team1 || '' : '';
    rowObject[`${prefix}_${slot}_AWAY`] = entry ? entry.team2 || '' : '';
    rowObject[`${prefix}_${slot}_HOME_SCORE`] = entry && entry.score1 !== undefined ? entry.score1 : '';
    rowObject[`${prefix}_${slot}_AWAY_SCORE`] = entry && entry.score2 !== undefined ? entry.score2 : '';
    rowObject[`${prefix}_${slot}_WINNER`] = entry ? entry.winner || '' : '';
  });
}

function applyFinalToRowObject_(rowObject, finalMatch) {
  rowObject.FINAL_HOME = finalMatch ? finalMatch.team1 || '' : '';
  rowObject.FINAL_AWAY = finalMatch ? finalMatch.team2 || '' : '';
  rowObject.FINAL_HOME_SCORE = finalMatch && finalMatch.score1 !== undefined ? finalMatch.score1 : '';
  rowObject.FINAL_AWAY_SCORE = finalMatch && finalMatch.score2 !== undefined ? finalMatch.score2 : '';
  rowObject.FINAL_WINNER = finalMatch ? finalMatch.winner || '' : '';
}

function findMatchByMatchNo_(matches, matchNo) {
  return (matches || []).find(match => Number(match.matchNo) === Number(matchNo)) || null;
}

/* =========================================================
   KNOCKOUT CONFIG
========================================================= */

function getFixedRoundOf32Rules_() {
  return [
    { matchNo: 73, date: '28 June', team1: { type: 'RUNNER_UP', group: 'A' }, team2: { type: 'RUNNER_UP', group: 'B' } },
    { matchNo: 75, date: '29 June', team1: { type: 'WINNER', group: 'F' }, team2: { type: 'RUNNER_UP', group: 'C' } },
    { matchNo: 76, date: '29 June', team1: { type: 'WINNER', group: 'C' }, team2: { type: 'RUNNER_UP', group: 'F' } },
    { matchNo: 78, date: '30 June', team1: { type: 'RUNNER_UP', group: 'E' }, team2: { type: 'RUNNER_UP', group: 'I' } },
    { matchNo: 83, date: '2 July',  team1: { type: 'RUNNER_UP', group: 'K' }, team2: { type: 'RUNNER_UP', group: 'L' } },
    { matchNo: 84, date: '2 July',  team1: { type: 'WINNER', group: 'H' }, team2: { type: 'RUNNER_UP', group: 'J' } },
    { matchNo: 86, date: '3 July',  team1: { type: 'WINNER', group: 'J' }, team2: { type: 'RUNNER_UP', group: 'H' } },
    { matchNo: 88, date: '3 July',  team1: { type: 'RUNNER_UP', group: 'D' }, team2: { type: 'RUNNER_UP', group: 'G' } }
  ];
}

function getFlexibleThirdPlaceMatchRules_() {
  return [
    { matchNo: 74, date: '29 June', team1: { type: 'WINNER', group: 'E' }, winnerGroup: 'E', side: 'top' },
    { matchNo: 77, date: '30 June', team1: { type: 'WINNER', group: 'I' }, winnerGroup: 'I', side: 'top' },
    { matchNo: 79, date: '30 June', team1: { type: 'WINNER', group: 'A' }, winnerGroup: 'A', side: 'top' },
    { matchNo: 80, date: '1 July',  team1: { type: 'WINNER', group: 'L' }, winnerGroup: 'L', side: 'top' },

    { matchNo: 81, date: '1 July',  team1: { type: 'WINNER', group: 'D' }, winnerGroup: 'D', side: 'bottom' },
    { matchNo: 82, date: '1 July',  team1: { type: 'WINNER', group: 'G' }, winnerGroup: 'G', side: 'bottom' },
    { matchNo: 85, date: '2 July',  team1: { type: 'WINNER', group: 'B' }, winnerGroup: 'B', side: 'bottom' },
    { matchNo: 87, date: '3 July',  team1: { type: 'WINNER', group: 'K' }, winnerGroup: 'K', side: 'bottom' }
  ];
}

function getKnockoutProgression_() {
  const fixturesMeta = getFixturesMap_();
  const base = {
    89:  { round: 'Round of 16',   team1: 'Winner Match 74',  team2: 'Winner Match 77' },
    90:  { round: 'Round of 16',   team1: 'Winner Match 73',  team2: 'Winner Match 75' },
    91:  { round: 'Round of 16',   team1: 'Winner Match 76',  team2: 'Winner Match 78' },
    92:  { round: 'Round of 16',   team1: 'Winner Match 79',  team2: 'Winner Match 80' },
    93:  { round: 'Round of 16',   team1: 'Winner Match 83',  team2: 'Winner Match 84' },
    94:  { round: 'Round of 16',   team1: 'Winner Match 81',  team2: 'Winner Match 82' },
    95:  { round: 'Round of 16',   team1: 'Winner Match 86',  team2: 'Winner Match 88' },
    96:  { round: 'Round of 16',   team1: 'Winner Match 85',  team2: 'Winner Match 87' },
    97:  { round: 'Quarter-final', team1: 'Winner Match 89',  team2: 'Winner Match 90' },
    98:  { round: 'Quarter-final', team1: 'Winner Match 93',  team2: 'Winner Match 94' },
    99:  { round: 'Quarter-final', team1: 'Winner Match 91',  team2: 'Winner Match 92' },
    100: { round: 'Quarter-final', team1: 'Winner Match 95',  team2: 'Winner Match 96' },
    101: { round: 'Semi-final',    team1: 'Winner Match 97',  team2: 'Winner Match 98' },
    102: { round: 'Semi-final',    team1: 'Winner Match 99',  team2: 'Winner Match 100' },
    103: { round: 'Bronze Final',  team1: 'Runner-up Match 101', team2: 'Runner-up Match 102' },
    104: { round: 'Final',         team1: 'Winner Match 101', team2: 'Winner Match 102' }
  };

  Object.keys(base).forEach(matchNo => {
    const meta = getFixtureMeta_(fixturesMeta, Number(matchNo), '', base[matchNo].round);
    base[matchNo].date = meta.date;
    base[matchNo].kickOffGMT = meta.kickOffGMT;
    base[matchNo].round = meta.round;
  });

  return base;
}

/* =========================================================
   SCORING SYSTEM
========================================================= */

function getScoringSystemData_() {
  return {
    title: 'World Cup 2026 Prediction Competition – Scoring Summary',

    groupStage: [
      '- 4 pts – Correct team qualify in correct position',
      '- 1 pt – Correct team qualify but in wrong position',
      '- 2 pts – Correct group winner',
      '- 5 pts – Perfect group (Bonus)'
    ],

    teamProgression: [
      'Points are cumulative (stack as teams progress):',
      '- Round of 32 – 3 pts',
      '- Round of 16 – 3 pts',
      '- Quarter-finals – 7 pts',
      '- Semi-finals – 10 pts',
      '- Final – 15 pts',
      '- Winner – 20 pts'
    ],

    knockoutScores: [
      'Per match:',
      '- 3 pts – Correct winner',
      '- 2 pts – Correct goals (Team 1)',
      '- 2 pts – Correct goals (Team 2)',
      '- 1 pt – Correct goal difference',
      '- 2 pts – Exact score (Bonus)',
      '',
      'Max per match – 10 pts'
    ],

    importantRules: [
      '- No draws allowed in knockout predictions',
      '- Penalty wins = treated as 1-goal win',
      '- Scores apply to fixture position (not teams)',
      '- Third-place match not included'
    ],

    tiebreakers: [
      '1. Total points',
      '2. Progression points',
      '3. Exact score predictions',
      '4. Group stage points'
    ]
  };
}

/* =========================================================
   PDF / DRIVE / EMAIL
========================================================= */

function getEntryRowObjectByRowNumber_(rowNumber) {
  const sheet = getSheet_(ENTRIES_RAW_SHEET);
  const headers = getSheetHeaders_(ENTRIES_RAW_SHEET);
  const rowValues = sheet.getRange(rowNumber, 1, 1, headers.length).getValues()[0];

  const obj = {};
  headers.forEach((header, index) => {
    obj[header] = rowValues[index];
  });

  return obj;
}

function buildPdfTemplateData_(row) {
  const settings = getSettingsMap_();
  const groups = getGroupTeams_();

  const groupRows = Object.keys(groups).sort().map(group => ({
    group: group,
    first: String(row[`${group}_1ST`] || ''),
    second: String(row[`${group}_2ND`] || ''),
    third: String(row[`${group}_3RD`] || '')
  }));

  const thirdRank = [];
  for (let i = 1; i <= 12; i++) {
    thirdRank.push(String(row[`THIRD_RANK_${i}`] || ''));
  }

  const thirdQualifiers = [];
  for (let i = 1; i <= 8; i++) {
    thirdQualifiers.push(String(row[`THIRD_Q${i}`] || ''));
  }

  const topSeedAssignments = [];
  const flexibleMatchNos = getFlexibleThirdPlaceMatchRules_().map(rule => rule.matchNo);
  for (let i = 1; i <= 8; i++) {
    topSeedAssignments.push({
      slot: i,
      matchNo: flexibleMatchNos[i - 1],
      team: String(row[`TS${i}_TEAM`] || '')
    });
  }

  return {
    competitionName: settings.COMPETITION_NAME || 'World Cup 2026 Predictor',
    headerImageUrl: PDF_HEADER_IMAGE_URL || '',
    logoImageUrl: PDF_LOGO_IMAGE_URL || '',
    entryId: String(row.ENTRY_ID || ''),
    fullName: String(row.FULL_NAME || ''),
    displayName: String(row.DISPLAY_NAME || row.FULL_NAME || ''),
    email: String(row.EMAIL || ''),
    timestamp: formatPdfDateTime_(row.TIMESTAMP),
    predictedChampion: String(row.PREDICTED_CHAMPION || ''),
    groups: groupRows,
    thirdRank: thirdRank,
    thirdQualifiers: thirdQualifiers,
    topSeedAssignments: topSeedAssignments,
    roundOf32: buildRoundMatchesForPdf_(row, 'R32', ROUND_OF32_MATCH_ORDER),
    roundOf16: buildRoundMatchesForPdf_(row, 'R16', ROUND_OF16_MATCH_ORDER),
    quarterFinals: buildRoundMatchesForPdf_(row, 'QF', QUARTER_FINAL_MATCH_ORDER),
    semiFinals: buildRoundMatchesForPdf_(row, 'SF', SEMI_FINAL_MATCH_ORDER),
    final: [{
      matchNo: FINAL_MATCH_NO,
      home: String(row.FINAL_HOME || ''),
      away: String(row.FINAL_AWAY || ''),
      homeScore: row.FINAL_HOME_SCORE,
      awayScore: row.FINAL_AWAY_SCORE,
      winner: String(row.FINAL_WINNER || '')
    }],
    scoringSystem: getScoringSystemData_()
  };
}

function buildRoundMatchesForPdf_(row, prefix, orderedMatchNos) {
  return orderedMatchNos.map((matchNo, index) => {
    const slot = index + 1;
    return {
      matchNo: matchNo,
      home: String(row[`${prefix}_${slot}_HOME`] || ''),
      away: String(row[`${prefix}_${slot}_AWAY`] || ''),
      homeScore: row[`${prefix}_${slot}_HOME_SCORE`],
      awayScore: row[`${prefix}_${slot}_AWAY_SCORE`],
      winner: String(row[`${prefix}_${slot}_WINNER`] || '')
    };
  });
}

function formatPdfDateTime_(value) {
  if (!value) return '';
  const date = new Date(value);
  if (isNaN(date.getTime())) return String(value);

  return Utilities.formatDate(
    date,
    Session.getScriptTimeZone(),
    'dd MMMM yyyy, HH:mm'
  );
}

function createEntryPdfBlob_(pdfData) {
  const template = HtmlService.createTemplateFromFile('EntryPdfTemplate');
  template.data = pdfData;

  const html = template.evaluate().getContent();
  const blob = Utilities.newBlob(html, 'text/html', 'entry.html').getAs('application/pdf');

  const safeName = sanitizeFileName_(pdfData.fullName || 'Entry');
  const fileName = `${pdfData.entryId} - ${safeName}.pdf`;

  return blob.setName(fileName);
}

function savePdfToDrive_(pdfBlob) {
  if (!PDF_FOLDER_ID || String(PDF_FOLDER_ID).trim() === '') {
    throw new Error('ERR - PDF folder ID has not been configured in Code.gs.');
  }

  try {
    const folder = DriveApp.getFolderById(String(PDF_FOLDER_ID).trim());
    return folder.createFile(pdfBlob);
  } catch (err) {
    throw new Error(
      'ERR - Unable to access the PDF folder. Check the folder ID is correct and the script account has permission. Details: ' +
      (err && err.message ? err.message : String(err))
    );
  }
}

function sendEntryConfirmationEmail_(pdfData, pdfBlob) {
  if (!pdfData.email) return;

  const settings = getSettingsMap_();
  const senderName = settings.EMAIL_SENDER_NAME || pdfData.competitionName;
  const subject = `Your World Cup 2026 Predictor Entry (${pdfData.entryId})`;

  const recipientName = escapeHtml_(pdfData.displayName || pdfData.fullName || 'Entrant');
  const competitionName = escapeHtml_(pdfData.competitionName || 'World Cup 2026 Predictor');
  const entryId = escapeHtml_(pdfData.entryId || '');
  const email = escapeHtml_(pdfData.email || '');
  const champion = escapeHtml_(pdfData.predictedChampion || 'Not selected');

  // Revolut button/section — only shown if the link is configured
  const revolutLinkSet = REVOLUT_PAYMENT_LINK && String(REVOLUT_PAYMENT_LINK).trim() !== '';
  const feeLabel = REVOLUT_ENTRY_FEE_LABEL ? ` ${escapeHtml_(REVOLUT_ENTRY_FEE_LABEL)}` : '';
  const feeLabelPlain = REVOLUT_ENTRY_FEE_LABEL ? ` (${REVOLUT_ENTRY_FEE_LABEL})` : '';

  const revolutBlockHtml = revolutLinkSet ? `
    <table role="presentation" width="100%" cellpadding="0" cellspacing="0" border="0" style="margin:22px 0 6px 0;border-collapse:separate;border-spacing:0;background:#f0fdf4;border:1px solid #bbf7d0;border-radius:12px;">
      <tr>
        <td style="padding:18px 20px 16px 20px;">
          <div style="font-size:12px;font-weight:700;letter-spacing:1px;text-transform:uppercase;color:#15803d;margin:0 0 6px 0;">
            Final step — Pay your entry fee${feeLabel}
          </div>
          <p style="margin:0 0 14px 0;font-size:14px;line-height:1.5;color:#166534;">
            Please complete your${feeLabel ? ' ' + escapeHtml_(REVOLUT_ENTRY_FEE_LABEL) : ''} entry fee payment via Revolut to finalise your entry. Your predictions are locked in regardless, but the fee is needed before kickoff.
          </p>
          <table role="presentation" cellpadding="0" cellspacing="0" border="0">
            <tr>
              <td style="border-radius:10px;background:#0075eb;">
                <a href="${escapeHtml_(REVOLUT_PAYMENT_LINK)}" target="_blank" style="display:inline-block;padding:13px 24px;font-size:15px;font-weight:700;color:#ffffff;text-decoration:none;border-radius:10px;font-family:Arial,Helvetica,sans-serif;">
                  Pay${feeLabel} via Revolut →
                </a>
              </td>
            </tr>
          </table>
          <div style="margin-top:10px;font-size:12px;line-height:1.5;color:#166534;">
            Or copy this link: <span style="color:#0075eb;word-break:break-all;">${escapeHtml_(REVOLUT_PAYMENT_LINK)}</span>
          </div>
        </td>
      </tr>
    </table>
  ` : '';

  const revolutBlockPlain = revolutLinkSet
    ? `\nFinal step — Pay your entry fee${feeLabelPlain}:\n${REVOLUT_PAYMENT_LINK}\n`
    : '';

  const plainBody =
    `Hi ${pdfData.displayName || pdfData.fullName},\n\n` +
    `Thank you for submitting your ${pdfData.competitionName} entry.\n\n` +
    `Your entry has been successfully received.\n\n` +
    `Entry ID: ${pdfData.entryId}\n` +
    `Predicted Champion: ${pdfData.predictedChampion || 'Not selected'}\n\n` +
    `A PDF copy of your submitted picks is attached to this email.\n` +
    revolutBlockPlain +
    `\nRegards,\n` +
    `${pdfData.competitionName}`;

  const htmlBody = `
    <div style="margin:0;padding:0;background:#f3f4f6;font-family:Arial,Helvetica,sans-serif;color:#111827;">
      <table role="presentation" width="100%" cellpadding="0" cellspacing="0" border="0" style="background:#f3f4f6;margin:0;padding:0;width:100%;">
        <tr>
          <td align="center" style="padding:32px 16px;">
            <table role="presentation" width="100%" cellpadding="0" cellspacing="0" border="0" style="max-width:620px;background:#ffffff;border:1px solid #e5e7eb;border-radius:16px;overflow:hidden;">

              <tr>
                <td style="padding:28px 32px 18px 32px;background:#14532d;">
                  <div style="font-size:12px;line-height:1.2;font-weight:700;letter-spacing:1.2px;text-transform:uppercase;color:rgba(255,255,255,0.72);margin:0 0 8px 0;">
                    Entry Confirmation
                  </div>
                  <div style="font-size:32px;line-height:1.05;font-weight:800;color:#ffffff;margin:0;">
                    Mick's World Cup 2026 Predictor
                  </div>
                  <div style="font-size:14px;line-height:1.5;color:rgba(255,255,255,0.82);margin:10px 0 0 0;">
                    Your entry has been successfully received and locked in.
                  </div>
                </td>
              </tr>

              <tr>
                <td style="padding:28px 32px 10px 32px;">
                  <p style="margin:0 0 16px 0;font-size:15px;line-height:1.6;color:#111827;">
                    Hi ${recipientName},
                  </p>

                  <p style="margin:0 0 16px 0;font-size:15px;line-height:1.6;color:#374151;">
                    Thank you for submitting your entry to the <strong style="color:#111827;">${competitionName}</strong>.
                    Your picks have been recorded successfully.
                  </p>

                  <table role="presentation" width="100%" cellpadding="0" cellspacing="0" border="0" style="margin:20px 0 22px 0;border-collapse:separate;border-spacing:0;background:#f9fafb;border:1px solid #e5e7eb;border-radius:12px;">
                    <tr>
                      <td colspan="2" style="padding:14px 16px 10px 16px;font-size:12px;font-weight:700;letter-spacing:1px;text-transform:uppercase;color:#6b7280;">
                        Entry Details
                      </td>
                    </tr>
                    <tr>
                      <td style="padding:0 16px 10px 16px;font-size:14px;line-height:1.5;color:#6b7280;width:160px;">
                        Entry ID
                      </td>
                      <td style="padding:0 16px 10px 16px;font-size:14px;line-height:1.5;color:#111827;font-weight:700;">
                        ${entryId}
                      </td>
                    </tr>
                    <tr>
                      <td style="padding:0 16px 10px 16px;font-size:14px;line-height:1.5;color:#6b7280;">
                        Email
                      </td>
                      <td style="padding:0 16px 10px 16px;font-size:14px;line-height:1.5;color:#111827;">
                        ${email}
                      </td>
                    </tr>
                    <tr>
                      <td style="padding:0 16px 16px 16px;font-size:14px;line-height:1.5;color:#6b7280;">
                        Predicted Champion
                      </td>
                      <td style="padding:0 16px 16px 16px;font-size:14px;line-height:1.5;color:#111827;font-weight:600;">
                        ${champion}
                      </td>
                    </tr>
                  </table>

                  <p style="margin:0 0 14px 0;font-size:15px;line-height:1.6;color:#374151;">
                    A PDF copy of your submitted picks is attached to this email for your records.
                  </p>

                  ${revolutBlockHtml}

                  <p style="margin:18px 0 0 0;font-size:15px;line-height:1.6;color:#374151;">
                    Please keep this email safe in case you need to refer back to your entry later.
                  </p>
                </td>
              </tr>

              <tr>
                <td style="padding:20px 32px 28px 32px;">
                  <div style="border-top:1px solid #e5e7eb;padding-top:18px;font-size:14px;line-height:1.7;color:#4b5563;">
                    Regards,<br>
                    <strong style="color:#111827;">${competitionName}</strong>
                  </div>
                </td>
              </tr>

            </table>

            <div style="max-width:620px;margin:12px auto 0 auto;padding:0 8px;font-size:11px;line-height:1.5;color:#9ca3af;text-align:center;">
              Official entry confirmation
            </div>
          </td>
        </tr>
      </table>
    </div>
  `;

  MailApp.sendEmail({
    to: pdfData.email,
    subject: subject,
    body: plainBody,
    htmlBody: htmlBody,
    name: senderName,
    attachments: [pdfBlob]
  });
}

function updatePdfStatus_(rowNumber, status, fileUrl, note) {
  const sheet = getSheet_(ENTRIES_RAW_SHEET);
  const headers = getSheetHeaders_(ENTRIES_RAW_SHEET);

  const headerIndex = {};
  headers.forEach((header, index) => {
    headerIndex[header] = index + 1;
  });

  if (headerIndex.PDF_SENT) {
    sheet.getRange(rowNumber, headerIndex.PDF_SENT).setValue(status);
  }

  if (headerIndex.PDF_SENT_AT) {
    sheet.getRange(rowNumber, headerIndex.PDF_SENT_AT).setValue(new Date());
  }

  if (headerIndex.PDF_FILE_URL) {
    sheet.getRange(rowNumber, headerIndex.PDF_FILE_URL).setValue(fileUrl || '');
  }

  if (headerIndex.NOTES) {
    sheet.getRange(rowNumber, headerIndex.NOTES).setValue(note || '');
  }
}

function sanitizeFileName_(value) {
  return String(value || '')
    .replace(/[\\\/:*?"<>|#%\[\]{}]/g, '')
    .replace(/\s+/g, ' ')
    .trim();
}

function escapeHtml_(value) {
  return String(value || '')
    .replace(/&/g, '&amp;')
    .replace(/</g, '&lt;')
    .replace(/>/g, '&gt;')
    .replace(/"/g, '&quot;')
    .replace(/'/g, '&#39;');
}

/* =========================================================
   FIXTURES HELPERS
========================================================= */

function getFixturesMap_() {
  const sheet = getSheet_(FIXTURES_SHEET);
  const values = sheet.getDataRange().getValues();

  if (values.length < 2) {
    return {};
  }

  const headers = values[0].map(h => String(h || '').trim());
  const matchNoCol = findHeaderIndex_(headers, ['MATCH NUMBER', 'MATCH_NO', 'MATCH NO']);
  const dateCol = findHeaderIndex_(headers, ['DATE (2026)', 'DATE']);
  const roundCol = findHeaderIndex_(headers, ['ROUND']);
  const kickOffCol = findHeaderIndex_(headers, ['KICK OFF GMT', 'KICKOFF GMT', 'KICK OFF', 'KICKOFF']);

  if (matchNoCol === -1) {
    throw new Error(`${FIXTURES_SHEET} must contain a Match Number column.`);
  }

  const map = {};

  for (let i = 1; i < values.length; i++) {
    const row = values[i];
    const matchNo = Number(row[matchNoCol]);
    if (!matchNo) continue;

    map[matchNo] = {
      matchNo: matchNo,
      date: dateCol > -1 ? formatFixtureDateCell_(row[dateCol]) : '',
      round: roundCol > -1 ? String(row[roundCol] || '').trim() : '',
      kickOffGMT: kickOffCol > -1 ? formatFixtureTimeCell_(row[kickOffCol]) : ''
    };
  }

  return map;
}

function getFixtureMeta_(fixturesMap, matchNo, fallbackDate, fallbackRound) {
  const meta = fixturesMap && fixturesMap[matchNo] ? fixturesMap[matchNo] : null;
  return {
    matchNo: matchNo,
    date: meta && meta.date ? meta.date : (fallbackDate || ''),
    round: meta && meta.round ? meta.round : (fallbackRound || ''),
    kickOffGMT: meta && meta.kickOffGMT ? meta.kickOffGMT : ''
  };
}

function findHeaderIndex_(headers, candidates) {
  const normalisedHeaders = headers.map(normalizeHeader_);
  for (let i = 0; i < candidates.length; i++) {
    const idx = normalisedHeaders.indexOf(normalizeHeader_(candidates[i]));
    if (idx > -1) return idx;
  }
  return -1;
}

function normalizeHeader_(value) {
  return String(value || '')
    .trim()
    .toUpperCase()
    .replace(/\s+/g, ' ');
}

function formatFixtureDateCell_(value) {
  if (value === null || value === undefined || value === '') return '';

  if (Object.prototype.toString.call(value) === '[object Date]' && !isNaN(value.getTime())) {
    return Utilities.formatDate(value, Session.getScriptTimeZone(), 'd MMMM');
  }

  return String(value).trim();
}

function formatFixtureTimeCell_(value) {
  if (value === null || value === undefined || value === '') return '';

  if (Object.prototype.toString.call(value) === '[object Date]' && !isNaN(value.getTime())) {
    return Utilities.formatDate(value, Session.getScriptTimeZone(), 'H:mm');
  }

  return String(value).trim();
}

// Kept for any callers outside getFixturesMap_ — delegates to date formatter by default
function formatFixtureCell_(value) {
  return formatFixtureDateCell_(value);
}

/* =========================================================
   HELPERS
========================================================= */

function getSettingsMap_() {
  const sheet = getSheet_(SETTINGS_SHEET);
  const values = sheet.getDataRange().getValues();
  const map = {};

  values.forEach(row => {
    const key = String(row[0] || '').trim();
    const value = row[1];
    if (key) map[key] = value;
  });

  return map;
}

function getGroupTeams_() {
  const sheet = getSheet_(GROUP_TEAMS_SHEET);
  const values = sheet.getDataRange().getValues();

  if (values.length < 2) {
    throw new Error(`${GROUP_TEAMS_SHEET} has no team data.`);
  }

  const headers = values[0].map(h => String(h || '').trim());
  const groupCol = headers.indexOf('GROUP');
  const teamCol = headers.indexOf('TEAM');

  if (groupCol === -1 || teamCol === -1) {
    throw new Error(`${GROUP_TEAMS_SHEET} must contain GROUP and TEAM headers.`);
  }

  const groups = {};

  for (let i = 1; i < values.length; i++) {
    const group = String(values[i][groupCol] || '').trim().toUpperCase();
    const team = String(values[i][teamCol] || '').trim();

    if (!group || !team) continue;

    if (!groups[group]) groups[group] = [];
    groups[group].push(team);
  }

  const groupLetters = Object.keys(groups).sort();

  if (!groupLetters.length) {
    throw new Error(`${GROUP_TEAMS_SHEET} contains no valid group/team rows.`);
  }

  groupLetters.forEach(group => {
    if (groups[group].length < 3) {
      throw new Error(`Group ${group} must contain at least 3 teams in ${GROUP_TEAMS_SHEET}.`);
    }
  });

  return groups;
}

function validateEntryWindow_() {
  const settings = getSettingsMap_();
  const entryStatus = String(settings.ENTRY_STATUS || 'OPEN').trim().toUpperCase();

  if (entryStatus !== 'OPEN') {
    throw new Error('Entries are currently closed.');
  }

  const deadlineRaw = settings.ENTRY_DEADLINE;
  if (deadlineRaw) {
    const deadline = new Date(deadlineRaw);
    if (!isNaN(deadline.getTime()) && new Date() > deadline) {
      throw new Error('The entry deadline has passed.');
    }
  }
}

function checkDuplicateEmail_(email) {
  const settings = getSettingsMap_();
  const maxEntries = Number(settings.MAX_ENTRIES_PER_EMAIL || 1);
  if (!maxEntries || maxEntries < 1) return;

  const headers = getSheetHeaders_(ENTRIES_RAW_SHEET);
  Logger.log('ENTRIES_RAW headers: ' + JSON.stringify(headers));

  const normalisedHeaders = headers.map(h => normalizeHeader_(h));
  const emailCol = normalisedHeaders.indexOf('EMAIL');

  if (emailCol === -1) {
    throw new Error(
      'ENTRIES_RAW must contain an EMAIL column. Headers found: ' + headers.join(' | ')
    );
  }

  const sheet = getSheet_(ENTRIES_RAW_SHEET);
  const lastRow = sheet.getLastRow();
  if (lastRow < 2) return;

  const emailValues = sheet.getRange(2, emailCol + 1, lastRow - 1, 1).getValues();

  let count = 0;
  for (let i = 0; i < emailValues.length; i++) {
    const existingEmail = String(emailValues[i][0] || '').trim().toLowerCase();
    if (existingEmail && existingEmail === String(email).trim().toLowerCase()) {
      count++;
    }
  }

  if (count >= maxEntries) {
    throw new Error('An entry has already been submitted for this email address.');
  }
}

function normalizePayload_(payload) {
  const groups = getGroupTeams_();
  const groupLetters = Object.keys(groups).sort();

  const cleaned = {
    FULL_NAME: toTitleCaseName_(payload.FULL_NAME || ''),
    EMAIL: String(payload.EMAIL || '').trim().toLowerCase(),
    DISPLAY_NAME: toTitleCaseName_(payload.DISPLAY_NAME || payload.FULL_NAME || ''),
    FORM_VERSION: String(payload.FORM_VERSION || 'V1').trim(),
    GROUPS: {},
    THIRD_RANK: []
  };

  groupLetters.forEach(group => {
    cleaned.GROUPS[group] = {
      '1ST': String(payload[`${group}_1ST`] || '').trim(),
      '2ND': String(payload[`${group}_2ND`] || '').trim(),
      '3RD': String(payload[`${group}_3RD`] || '').trim()
    };
  });

  for (let i = 1; i <= groupLetters.length; i++) {
    cleaned.THIRD_RANK.push(String(payload[`THIRD_RANK_${i}`] || '').trim());
  }

  return cleaned;
}

function validatePayload_(cleaned) {
  if (!cleaned.FULL_NAME) throw new Error('Full Name is required.');
  if (!cleaned.EMAIL) throw new Error('Email Address is required.');

  const emailPattern = /^[^\s@]+@[^\s@]+\.[^\s@]+$/;
  if (!emailPattern.test(cleaned.EMAIL)) {
    throw new Error('Please enter a valid email address.');
  }

  const groups = getGroupTeams_();
  const groupLetters = Object.keys(groups).sort();
  const selectedThirds = [];

  groupLetters.forEach(group => {
    const allowedTeams = groups[group];
    const first = cleaned.GROUPS[group]['1ST'];
    const second = cleaned.GROUPS[group]['2ND'];
    const third = cleaned.GROUPS[group]['3RD'];

    if (!first || !second || !third) {
      throw new Error(`Please complete all three positions for Group ${group}.`);
    }

    const picks = [first, second, third];

    if (new Set(picks).size !== 3) {
      throw new Error(`Each team can only be used once in Group ${group}.`);
    }

    picks.forEach(team => {
      if (!allowedTeams.includes(team)) {
        throw new Error(`Invalid selection detected in Group ${group}.`);
      }
    });

    selectedThirds.push(third);
  });

  if (cleaned.THIRD_RANK.length !== groupLetters.length) {
    throw new Error('The 3rd-place ranking is incomplete.');
  }

  if (cleaned.THIRD_RANK.some(team => !team)) {
    throw new Error('Please complete all 3rd-place ranking positions.');
  }

  if (new Set(cleaned.THIRD_RANK).size !== cleaned.THIRD_RANK.length) {
    throw new Error('Each 3rd-place team can only be used once in the ranking.');
  }

  const selectedThirdsSorted = [...selectedThirds].sort();
  const rankedThirdsSorted = [...cleaned.THIRD_RANK].sort();

  if (JSON.stringify(selectedThirdsSorted) !== JSON.stringify(rankedThirdsSorted)) {
    throw new Error('Your 3rd-place ranking must use exactly the teams you selected to finish 3rd.');
  }
}

function toTitleCaseName_(value) {
  return String(value || '')
    .trim()
    .toLowerCase()
    .split(/\s+/)
    .filter(Boolean)
    .map(word => word.charAt(0).toUpperCase() + word.slice(1))
    .join(' ');
}

function getInitials_(fullName) {
  const parts = String(fullName || '')
    .trim()
    .split(/\s+/)
    .filter(Boolean)
    .filter(part => !/^\d+$/.test(part));

  if (!parts.length) return 'XX';

  const initials = parts
    .slice(0, 2)
    .map(part => part.charAt(0).toUpperCase())
    .join('');

  return initials || 'XX';
}

function generateFormattedEntryId_(fullName, rowNumber) {
  const initials = getInitials_(fullName);
  const sequence = String(rowNumber - 1).padStart(4, '0');
  return `WC26-${initials}${sequence}`;
}

function getUniqueDisplayName_(fullName) {
  const baseName = String(fullName || '').trim();
  if (!baseName) return '';

  const sheet = getSheet_(ENTRIES_RAW_SHEET);
  const values = sheet.getDataRange().getValues();
  if (values.length < 2) return baseName;

  const headers = values[0].map(h => String(h || '').trim());
  const fullNameCol = headers.indexOf('FULL_NAME');

  if (fullNameCol === -1) {
    return baseName;
  }

  let highestSuffix = 1;
  let exactFound = false;
  const escapedBase = escapeRegex_(baseName);
  const suffixRegex = new RegExp(`^${escapedBase}\\s+(\\d+)$`, 'i');

  for (let i = 1; i < values.length; i++) {
    const existingName = String(values[i][fullNameCol] || '').trim();
    if (!existingName) continue;

    if (existingName.toLowerCase() === baseName.toLowerCase()) {
      exactFound = true;
      highestSuffix = Math.max(highestSuffix, 1);
      continue;
    }

    const match = existingName.match(suffixRegex);
    if (match) {
      highestSuffix = Math.max(highestSuffix, Number(match[1]) || 1);
    }
  }

  if (!exactFound && highestSuffix === 1) {
    return baseName;
  }

  return `${baseName} ${highestSuffix + 1}`;
}

function escapeRegex_(value) {
  return String(value || '').replace(/[.*+?^${}()|[\]\\]/g, '\\$&');
}

function getHeaderIndexMap_(sheetName) {
  const headers = getSheetHeaders_(sheetName);
  const map = {};
  headers.forEach((header, index) => {
    map[header] = index + 1;
  });
  return map;
}

function getSheetHeaders_(sheetName) {
  const sheet = getSheet_(sheetName);
  const lastColumn = sheet.getLastColumn();

  if (lastColumn === 0) {
    throw new Error(`${sheetName} has no headers.`);
  }

  const headers = sheet.getRange(1, 1, 1, lastColumn).getValues()[0]
    .map(h => String(h || '').trim());

  if (headers.some(h => !h)) {
    throw new Error(`${sheetName} contains one or more blank headers in row 1.`);
  }

  return headers;
}

function getSheet_(sheetName) {
  const ss = SpreadsheetApp.openById(SPREADSHEET_ID);
  const sheet = ss.getSheetByName(sheetName);

  if (!sheet) {
    throw new Error(`Sheet not found: ${sheetName}`);
  }

  return sheet;
}

/* =========================================================
   TEST HELPERS
========================================================= */

function testGroupTeams() {
  Logger.log(JSON.stringify(getGroupTeams_()));
}

function testInitialData() {
  Logger.log(JSON.stringify(getInitialData()));
}

function testGenerateKnockoutFromLastEntry() {
  const sheet = getSheet_(ENTRIES_RAW_SHEET);
  const values = sheet.getDataRange().getValues();
  if (values.length < 2) {
    throw new Error('ENTRIES_RAW has no entries to test.');
  }

  const headers = values[0].map(h => String(h || '').trim());
  const lastRow = values[values.length - 1];
  const payload = {};

  headers.forEach((header, index) => {
    payload[header] = lastRow[index];
  });

  Logger.log(JSON.stringify(generateKnockoutBracket(payload)));
}

function testCreatePdfFromLastEntry() {
  const sheet = getSheet_(ENTRIES_RAW_SHEET);
  const lastRow = sheet.getLastRow();
  if (lastRow < 2) throw new Error('No entries found.');

  const row = getEntryRowObjectByRowNumber_(lastRow);
  const pdfData = buildPdfTemplateData_(row);
  const pdfBlob = createEntryPdfBlob_(pdfData);
  const file = savePdfToDrive_(pdfBlob);

  Logger.log(file.getUrl());
}

function testEmailLastEntryPdf() {
  const sheet = getSheet_(ENTRIES_RAW_SHEET);
  const lastRow = sheet.getLastRow();
  if (lastRow < 2) throw new Error('No entries found.');

  const row = getEntryRowObjectByRowNumber_(lastRow);
  const pdfData = buildPdfTemplateData_(row);
  const pdfBlob = createEntryPdfBlob_(pdfData);
  const file = savePdfToDrive_(pdfBlob);
  sendEntryConfirmationEmail_(pdfData, pdfBlob);

  Logger.log('PDF saved: ' + file.getUrl());
  Logger.log('Email sent to: ' + pdfData.email);
}

function testPdfFolderAccess() {
  const folder = DriveApp.getFolderById(String(PDF_FOLDER_ID).trim());
  Logger.log('Folder name: ' + folder.getName());
  Logger.log('Folder URL: ' + folder.getUrl());
}

function testEntriesRawHeaders() {
  const headers = getSheetHeaders_(ENTRIES_RAW_SHEET);
  Logger.log('Headers: ' + JSON.stringify(headers));
  Logger.log('Normalised: ' + JSON.stringify(headers.map(h => normalizeHeader_(h))));
}
/* =========================================================
   PUBLIC WEBSITE API
========================================================= */

function doGet(e) {
  const action = e && e.parameter && e.parameter.action
    ? e.parameter.action : '';

  if (action === 'getStats')        return jsonResponse(getPublicStats_());
  if (action === 'getLeaderboard')  return jsonResponse(getPublicLeaderboard_());
  if (action === 'getEntrantNames') return jsonResponse(getEntrantNames_());
  if (action === 'lookupEntry') {
    const name = e.parameter.name || e.parameter.q || '';
    return jsonResponse(lookupPublicEntry_(String(name).trim()));
  }

  // Default — serve the entry form
  return HtmlService.createHtmlOutputFromFile('Index')
    .setTitle('World Cup 2026 Prediction')
    .setXFrameOptionsMode(HtmlService.XFrameOptionsMode.ALLOWALL);
}

function jsonResponse(data) {
  return ContentService
    .createTextOutput(JSON.stringify(data))
    .setMimeType(ContentService.MimeType.JSON);
}

function getEntrantNames_() {
  try {
    const sheet = getSheet_(ENTRIES_RAW_SHEET);
    const values = sheet.getDataRange().getValues();
    if (values.length < 2) return { ok: true, names: [] };

    const headers = values[0].map(h => String(h || '').trim());
    const dispCol   = headers.indexOf('DISPLAY_NAME');
    const nameCol   = headers.indexOf('FULL_NAME');
    const statusCol = headers.indexOf('ENTRY_STATUS');

    const names = [];
    for (let i = 1; i < values.length; i++) {
      const row = values[i];
      if (statusCol > -1 &&
          String(row[statusCol] || '').trim().toUpperCase() !== 'LOCKED') continue;
      const name = String(row[dispCol > -1 ? dispCol : nameCol] || '').trim();
      if (name) names.push(name);
    }
    return { ok: true, names };
  } catch(err) {
    return { ok: false, message: err.message };
  }
}

function getPublicStats_() {
  try {
    const sheet = getSheet_(ENTRIES_RAW_SHEET);
    const values = sheet.getDataRange().getValues();
    if (values.length < 2) {
      return { ok: true, entryCount: 0, uniqueChampions: 0,
               championPicks: [], groupWinnerPicks: [],
               thirdQualifierPicks: [], finalistPicks: [] };
    }

    const headers = values[0].map(h => String(h || '').trim());
    const col = name => headers.indexOf(name);
    const groups = getGroupTeams_();
    const groupLetters = Object.keys(groups).sort();

    const champCounts = {}, groupWinnerCounts = {},
          thirdQCounts = {}, finalistCounts = {};
    let entryCount = 0;

    for (let i = 1; i < values.length; i++) {
      const row = values[i];
      if (String(row[col('ENTRY_STATUS')] || '').trim().toUpperCase() !== 'LOCKED') continue;
      entryCount++;

      const champ = String(row[col('PREDICTED_CHAMPION')] || '').trim();
      if (champ) champCounts[champ] = (champCounts[champ] || 0) + 1;

      groupLetters.forEach(g => {
        const c = col(`${g}_1ST`);
        if (c > -1) {
          const team = String(row[c] || '').trim();
          if (team) groupWinnerCounts[team] = (groupWinnerCounts[team] || 0) + 1;
        }
      });

      for (let q = 1; q <= 8; q++) {
        const c = col(`THIRD_Q${q}`);
        if (c > -1) {
          const team = String(row[c] || '').trim();
          if (team) thirdQCounts[team] = (thirdQCounts[team] || 0) + 1;
        }
      }

      ['SF_1_WINNER','SF_2_WINNER'].forEach(f => {
        const c = col(f);
        if (c > -1) {
          const team = String(row[c] || '').trim();
          if (team) finalistCounts[team] = (finalistCounts[team] || 0) + 1;
        }
      });
    }

    const sorted = obj => Object.entries(obj)
      .sort((a, b) => b[1] - a[1])
      .map(([team, count]) => [team, count]);

    return {
      ok: true,
      entryCount,
      uniqueChampions:     Object.keys(champCounts).length,
      championPicks:       sorted(champCounts),
      groupWinnerPicks:    sorted(groupWinnerCounts),
      thirdQualifierPicks: sorted(thirdQCounts),
      finalistPicks:       sorted(finalistCounts)
    };
  } catch(err) {
    return { ok: false, message: err.message };
  }
}

function getPublicLeaderboard_() {
  try {
    const sheet = getSheet_(ENTRIES_RAW_SHEET);
    const values = sheet.getDataRange().getValues();
    if (values.length < 2) {
      return { ok: true, hasScores: false, entryCount: 0, entries: [] };
    }

    const headers = values[0].map(h => String(h || '').trim());
    const col = name => headers.indexOf(name);

    const entries = [];
    for (let i = 1; i < values.length; i++) {
      const row = values[i];
      if (String(row[col('ENTRY_STATUS')] || '').trim().toUpperCase() !== 'LOCKED') continue;
      const totalPoints = Number(row[col('TOTAL_POINTS')] || 0);
      entries.push({
        displayName:      String(row[col('DISPLAY_NAME')] || row[col('FULL_NAME')] || '').trim(),
        entryId:          String(row[col('ENTRY_ID')] || '').trim(),
        predictedChampion:String(row[col('PREDICTED_CHAMPION')] || '').trim(),
        totalScore:       totalPoints,
        groupScore:       Number(row[col('GROUP_POINTS')] || 0) || null,
        progressionScore: Number(row[col('PROGRESSION_POINTS')] || 0) || null,
        knockoutScore:    Number(row[col('KNOCKOUT_POINTS')] || 0) || null
      });
    }

    const hasScores = entries.some(e => e.totalScore > 0);
    entries.sort((a, b) => b.totalScore - a.totalScore);

    const updatedAt = Utilities.formatDate(
      new Date(), Session.getScriptTimeZone(), 'dd MMM yyyy, HH:mm');

    return { ok: true, hasScores, entryCount: entries.length,
             entries, lastUpdated: updatedAt };
  } catch(err) {
    return { ok: false, message: err.message };
  }
}

function lookupPublicEntry_(name) {
  try {
    if (!name) return { ok: false, message: 'No name provided.' };

    const sheet = getSheet_(ENTRIES_RAW_SHEET);
    const values = sheet.getDataRange().getValues();
    if (values.length < 2) return { ok: false, message: 'No entries found.' };

    const headers = values[0].map(h => String(h || '').trim());
    const col = n => headers.indexOf(n);

    const rowIndex = values.findIndex((row, i) => {
      if (i === 0) return false;
      const dispName = String(row[col('DISPLAY_NAME')] || row[col('FULL_NAME')] || '').trim();
      return dispName.toLowerCase() === name.toLowerCase();
    });

    if (rowIndex === -1) return { ok: false, message: 'Entry not found.' };

    const row = values[rowIndex];
    const groups = getGroupTeams_();
    const groupLetters = Object.keys(groups).sort();

    const groupData = groupLetters.map(g => ({
      group:  g,
      first:  String(row[col(`${g}_1ST`)] || ''),
      second: String(row[col(`${g}_2ND`)] || ''),
      third:  String(row[col(`${g}_3RD`)] || '')
    }));

    const buildRound = (prefix, matchNos) => matchNos.map((matchNo, idx) => {
      const s = idx + 1;
      return {
        matchNo,
        home:      String(row[col(`${prefix}_${s}_HOME`)]       || ''),
        away:      String(row[col(`${prefix}_${s}_AWAY`)]       || ''),
        homeScore: row[col(`${prefix}_${s}_HOME_SCORE`)] !== undefined
                   ? row[col(`${prefix}_${s}_HOME_SCORE`)] : '',
        awayScore: row[col(`${prefix}_${s}_AWAY_SCORE`)] !== undefined
                   ? row[col(`${prefix}_${s}_AWAY_SCORE`)] : '',
        winner:    String(row[col(`${prefix}_${s}_WINNER`)]     || ''),
        pts:       Number(row[col(`${prefix}_${s}_PTS`)]        || 0) || null
      };
    });

    return {
      ok: true,
      entry: {
        fullName:         String(row[col('FULL_NAME')]          || ''),
        displayName:      String(row[col('DISPLAY_NAME')]       || ''),
        entryId:          String(row[col('ENTRY_ID')]           || ''),
        predictedChampion:String(row[col('PREDICTED_CHAMPION')] || ''),
        totalScore:       Number(row[col('TOTAL_POINTS')]       || 0) || null,
        groups:           groupData,
        roundOf32:        buildRound('R32', ROUND_OF32_MATCH_ORDER),
        roundOf16:        buildRound('R16', ROUND_OF16_MATCH_ORDER),
        quarterFinals:    buildRound('QF',  QUARTER_FINAL_MATCH_ORDER),
        semiFinals:       buildRound('SF',  SEMI_FINAL_MATCH_ORDER),
        final:            buildRound('FINAL', [FINAL_MATCH_NO])
      }
    };
  } catch(err) {
    return { ok: false, message: err.message };
  }
}

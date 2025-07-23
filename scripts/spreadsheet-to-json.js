const XLSX = require('xlsx');
const { z } = require('zod');
const chalk = require("chalk");

// Zod schema for the spreadsheet JSON structure
const SpreadsheetSchema = z.object({
  client_name: z.string(),
  created_date: z.string(),
  reports: z.array(
    z.object({
      report_name: z.string(),
      questions: z.array(
        z.object({
          question_number: z.number().nullable(),
          question_text: z.string(),
          sub_question_text: z.string().optional(),
          responses: z.array(
            z.object({
              respondent: z.string(),
              score: z.number().optional(),
              comment: z.string().optional(),
              skip_reason: z.string().optional(),
              response: z.string().optional(),
            })
          ),
        })
      ),
    })
  ),
  respondees: z.array(
    z.object({
      name: z.string(),
      position: z.string().optional(),
    })
  ),
});

function normalizeHeader(header) {
  return String(header)
    .toLowerCase()
    .replace(/[^a-z0-9]/g, '');
}

const RAW_HEADER_MAP = {
  'Client Name': 'client_name',
  Created: 'created_date',
  'Question Text': 'question_text',
  'Question #': 'question_number',
  'Question Number': 'question_number',
  'Q Number': 'question_number',
  'Q No': 'question_number',
  'Q#': 'question_number',
  // "Question": "question_text", // Removed to avoid ambiguity
  'Sub-Question': 'sub_question_text',
  'Sub Question': 'sub_question_text',
  category: 'category',
  Respondent: 'respondent',
  Position: 'position',
  Response: 'score',
  Comment: 'comment',
  'Skip Reason': 'skip_reason',
  // Add more mappings as needed
};
const HEADER_MAP = Object.fromEntries(Object.entries(RAW_HEADER_MAP).map(([k, v]) => [normalizeHeader(k), v]));

function mapHeader(header) {
  const norm = normalizeHeader(header);
  if (HEADER_MAP[norm]) return HEADER_MAP[norm];
  for (const key in HEADER_MAP) {
    if (norm.includes(key)) return HEADER_MAP[key];
  }
  return null;
}

function formatDateToYMD(dateStr) {
  if (!dateStr) return '';
  const d = new Date(dateStr);
  if (Number.isNaN(d.getTime())) return dateStr;
  return d.toISOString().slice(0, 10);
}

function getUnmatchedHeaders(fileContent) {
  const workbook = XLSX.read(fileContent, { type: 'buffer' });
  const sheetNames = workbook.SheetNames;
  const unmatched = new Set();
  for (let i = 1; i < sheetNames.length; i++) {
    const sheet = workbook.Sheets[sheetNames[i]];
    const rows = XLSX.utils.sheet_to_json(sheet, { header: 1 });
    if (rows.length < 1) continue;
    const headers = rows[0];
    for (const h of headers) {
      if (!mapHeader(h)) {
        unmatched.add(h);
      }
    }
  }
  return Array.from(unmatched);
}

function parseSpreadsheetBuffer(fileContent) {
  const workbook = XLSX.read(fileContent, { type: 'buffer' });
  const sheetNames = workbook.SheetNames;
  if (sheetNames.length < 2) {
    throw new Error('Excel file must contain at least a details sheet and one data sheet');
  }

  // --- 1. Extract main details from first sheet ---
  const detailsSheet = workbook.Sheets[sheetNames[0]];
  const detailsRows = XLSX.utils.sheet_to_json(detailsSheet, { header: 1 });
  let client_name = '';
  let created_date = '';
  for (const row of detailsRows) {
    if (row.length < 2) continue;
    const fieldName = row[0];
    const value = row[1];
    if (fieldName === 'Client Name') client_name = value;
    if (fieldName === 'Created') created_date = formatDateToYMD(value);
  }
  if (!client_name || !created_date) {
    throw new Error('Could not extract client_name or created_date from details sheet');
  }

  // --- 2. Process each data sheet and build reports structure ---
  const reportMap = new Map();
  const respondeeMap = new Map();
  for (let i = 1; i < sheetNames.length; i++) {
    const originalSheetName = sheetNames[i];
    const report_name = originalSheetName
      .replace(/([a-z])([A-Z0-9])/g, '$1 $2')
      .replace(/([A-Z])([A-Z][a-z])/g, '$1 $2')
      .trim();
    const sheet = workbook.Sheets[originalSheetName];
    const rows = XLSX.utils.sheet_to_json(sheet, { header: 1 });
    if (rows.length < 2) continue;
    const headers = rows[0].map((h) => mapHeader(h) || normalizeHeader(h));
    let lastQuestionNumber = null;
    let lastQuestionText = '';
    let lastSubQuestionText = '';
    const questionMap = new Map();
    for (let r = 1; r < rows.length; r++) {
      const row = rows[r];
      let question_number = lastQuestionNumber;
      let question_text = lastQuestionText;
      let sub_question_text = lastSubQuestionText;
      let respondent = '';
      let position = '';
      let score = undefined;
      let comment = '';
      let skip_reason = '';
      let response = '';
      let newQuestionText = false;
      let explicitQuestionNumber = false;
      let foundQuestionNumber = null;
      for (let c = 0; c < headers.length; c++) {
        const field = headers[c];
        if (!field) continue;
        const value = row[c];
        if (field === 'question_number' && value !== undefined && value !== null && value !== '') {
          const parsed = Number.parseInt(value);
          if (!Number.isNaN(parsed)) {
            foundQuestionNumber = parsed;
            explicitQuestionNumber = true;
          }
        }
        if (field === 'question_text' && value) {
          const cleaned = String(value)
            .replace(/\[[a-zA-Z0-9_]+\]/g, '')
            .trim();
          if (cleaned !== lastQuestionText) {
            question_text = cleaned;
            lastQuestionText = cleaned;
            // If sub_question_text is present in this row, set it, else clear it
            const subIdx = headers.findIndex((h) => h === 'sub_question_text');
            if (subIdx !== -1 && row[subIdx]) {
              sub_question_text = row[subIdx];
              lastSubQuestionText = row[subIdx];
            } else {
              sub_question_text = '';
              lastSubQuestionText = '';
            }
            newQuestionText = true;
          }
        }
        if (field === 'sub_question_text' && value) {
          sub_question_text = value;
          lastSubQuestionText = value;
        }
        if (field === 'respondent') respondent = value || '';
        if (field === 'position') position = value || '';
        if (field === 'score') {
          if (value !== undefined && value !== null && value !== '') {
            if (Number.isNaN(Number(value))) {
              response = value;
            } else {
              score = Number.parseInt(value) || 0;
              response = '';
            }
          }
        }
        if (field === 'comment') comment = value || '';
        if (field === 'skip_reason') skip_reason = value || '';
        if (field === 'response' && response === '') response = value || '';
      }
      // Set question_number logic after parsing all columns
      if (newQuestionText) {
        if (explicitQuestionNumber) {
          question_number = foundQuestionNumber;
          lastQuestionNumber = foundQuestionNumber;
        } else {
          question_number = null;
          lastQuestionNumber = null;
        }
      } else if (explicitQuestionNumber) {
        question_number = foundQuestionNumber;
        lastQuestionNumber = foundQuestionNumber;
      } else {
        question_number = lastQuestionNumber;
      }
      if (respondent && String(respondent).trim() !== '') {
        if (!respondeeMap.has(respondent)) {
          respondeeMap.set(respondent, { name: respondent, position });
        }
        const qKey = `${question_number}|${question_text}|${sub_question_text}`;
        const questionObj = {
          question_number: question_number ?? null,
          question_text: question_text || '',
          responses: [],
        };
        if (sub_question_text) {
          questionObj.sub_question_text = sub_question_text;
        }
        questionMap.set(qKey, questionObj);
        // Build response object, omitting empty fields
        const responseObj = { respondent };
        if (score !== undefined && score !== null && score !== '') responseObj.score = score;
        if (comment) responseObj.comment = comment;
        if (skip_reason) responseObj.skip_reason = skip_reason;
        if (response) responseObj.response = response;
        questionMap.get(qKey).responses.push(responseObj);
      }
    }
    reportMap.set(report_name, {
      report_name,
      questions: Array.from(questionMap.values()),
    });
  }
  const reports = Array.from(reportMap.values());
  const respondees = Array.from(respondeeMap.values());
  const spreadsheetJson = { client_name, created_date, reports, respondees };
  const unmatchedHeaders = getUnmatchedHeaders(fileContent);
  const result = { ...spreadsheetJson };
  if (unmatchedHeaders.length > 0) {
    result.unmatchedHeaders = unmatchedHeaders;
  }
  SpreadsheetSchema.parse(spreadsheetJson);
  return result;
}

/**
 * Compile a summary of companies, each with a list of respondents and a list of report years.
 * @param {Array} spreadsheets - Array of parsed spreadsheet JSON objects.
 * @returns {Array} - Array of companies with respondents and report years.
 */
function compileCompaniesSummary(spreadsheets) {
  const companies = {};
  for (const sheet of spreadsheets) {
    const { client_name, created_date, respondees, reports } = sheet;
    if (!client_name) continue;
    if (!companies[client_name]) {
      companies[client_name] = { name: client_name, respondents: new Set(), years: new Set() };
    }
    if (Array.isArray(respondees)) {
      for (const r of respondees) {
        if (r && r.name) companies[client_name].respondents.add(r.name);
      }
    }
    // Extract year from created_date (YYYY or YYYY-MM-DD)
    const year = created_date ? String(created_date).slice(0, 4) : null;
    if (year) companies[client_name].years.add(year);
  }
  // Convert sets to arrays
  return Object.values(companies).map((c) => ({
    name: c.name,
    respondents: Array.from(c.respondents),
    years: Array.from(c.years).sort(),
  }));
}

/**
 * Compile a summary of questions, each with the years it was used, subquestions, and their years.
 * @param {Array} spreadsheets - Array of parsed spreadsheet JSON objects.
 * @returns {Array} - Array of questions with years and subquestions.
 */
function compileQuestionsSummary(spreadsheets) {
  const questions = {};
  for (const sheet of spreadsheets) {
    const { created_date, reports } = sheet;
    const year = created_date ? String(created_date).slice(0, 4) : null;
    if (!Array.isArray(reports)) continue;
    for (const report of reports) {
      if (!Array.isArray(report.questions)) continue;
      for (const q of report.questions) {
        const qText = q.question_text || '';
        if (!qText) continue;
        if (!questions[qText]) {
          questions[qText] = { question: qText, years: new Set(), subquestions: {} };
        }
        if (year) questions[qText].years.add(year);
        const subText = q.sub_question_text;
        if (subText) {
          if (!questions[qText].subquestions[subText]) {
            questions[qText].subquestions[subText] = new Set();
          }
          if (year) questions[qText].subquestions[subText].add(year);
        }
      }
    }
  }
  // Convert sets to arrays
  return Object.values(questions).map((q) => ({
    question: q.question,
    years: Array.from(q.years).sort(),
    subquestions: Object.entries(q.subquestions).map(([sub, years]) => ({
      subquestion: sub,
      years: Array.from(years).sort(),
    })),
  }));
}

module.exports = {
  SpreadsheetSchema,
  parseSpreadsheetBuffer,
  normalizeHeader,
  mapHeader,
  formatDateToYMD,
  getUnmatchedHeaders,
  compileCompaniesSummary,
  compileQuestionsSummary,
};

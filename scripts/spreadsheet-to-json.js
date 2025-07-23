const ExcelJS = require('exceljs');
const { z } = require('zod');
const chalk = require('chalk');

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
            }),
          ),
        }),
      ),
    }),
  ),
  respondees: z.array(
    z.object({
      name: z.string(),
      position: z.string().optional(),
    }),
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
  'Sub-Question': 'sub_question_text',
  'Sub Question': 'sub_question_text',
  category: 'category',
  Respondent: 'respondent',
  Position: 'position',
  Response: 'score',
  Comment: 'comment',
  'Skip Reason': 'skip_reason',
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

async function getUnmatchedHeaders(fileContent) {
  const workbook = new ExcelJS.Workbook();
  await workbook.xlsx.load(fileContent);
  const unmatched = new Set();
  workbook.worksheets.forEach((sheet, index) => {
    if (index === 0) return; // Skip details sheet
    const headers = sheet.getRow(1).values;
    headers.forEach((h) => {
      if (h && !mapHeader(h)) {
        unmatched.add(h);
      }
    });
  });
  return Array.from(unmatched);
}

async function parseSpreadsheetBuffer(fileContent) {
  const workbook = new ExcelJS.Workbook();
  await workbook.xlsx.load(fileContent);

  if (workbook.worksheets.length < 2) {
    throw new Error('Excel file must contain at least a details sheet and one data sheet');
  }

  // --- 1. Extract main details from first sheet ---
  const detailsSheet = workbook.worksheets[0];
  let client_name = '';
  let created_date = '';
  detailsSheet.eachRow((row) => {
    const fieldName = row.getCell(1).value;
    const value = row.getCell(2).value;
    if (fieldName === 'Client Name') client_name = value;
    if (fieldName === 'Created') created_date = formatDateToYMD(value);
  });

  if (!client_name || !created_date) {
    throw new Error('Could not extract client_name or created_date from details sheet');
  }

  // --- 2. Process each data sheet and build reports structure ---
  const reportMap = new Map();
  const respondeeMap = new Map();

  for (let i = 1; i < workbook.worksheets.length; i++) {
    const sheet = workbook.worksheets[i];
    const report_name = sheet.name
      .replace(/([a-z])([A-Z0-9])/g, '$1 $2')
      .replace(/([A-Z])([A-Z][a-z])/g, '$1 $2')
      .trim();

    const headerRow = sheet.getRow(1).values;
    const headers = headerRow.map((h) => (h ? mapHeader(h) || normalizeHeader(h) : null));

    let lastQuestionNumber = null;
    let lastQuestionText = '';
    let lastSubQuestionText = '';
    const questionMap = new Map();

    sheet.eachRow((row, rowNumber) => {
      if (rowNumber === 1) return; // Skip header row

      const rowValues = row.values;
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

      headers.forEach((field, c) => {
        if (!field) return;
        const value = rowValues[c];
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
            const subIdx = headers.findIndex((h) => h === 'sub_question_text');
            if (subIdx !== -1 && rowValues[subIdx]) {
              sub_question_text = rowValues[subIdx];
              lastSubQuestionText = rowValues[subIdx];
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
      });

      if (newQuestionText) {
        question_number = explicitQuestionNumber ? foundQuestionNumber : null;
        lastQuestionNumber = question_number;
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
        if (!questionMap.has(qKey)) {
          const questionObj = {
            question_number: question_number ?? null,
            question_text: question_text || '',
            responses: [],
          };
          if (sub_question_text) {
            questionObj.sub_question_text = sub_question_text;
          }
          questionMap.set(qKey, questionObj);
        }

        const responseObj = { respondent };
        if (score !== undefined && score !== null && score !== '') responseObj.score = score;
        if (comment) responseObj.comment = comment;
        if (skip_reason) responseObj.skip_reason = skip_reason;
        if (response) responseObj.response = response;
        questionMap.get(qKey).responses.push(responseObj);
      }
    });

    reportMap.set(report_name, {
      report_name,
      questions: Array.from(questionMap.values()),
    });
  }

  const reports = Array.from(reportMap.values());
  const respondees = Array.from(respondeeMap.values());
  const spreadsheetJson = { client_name, created_date, reports, respondees };

  const unmatchedHeaders = await getUnmatchedHeaders(fileContent);
  const result = { ...spreadsheetJson };
  if (unmatchedHeaders.length > 0) {
    result.unmatchedHeaders = unmatchedHeaders;
  }

  SpreadsheetSchema.parse(spreadsheetJson);
  return result;
}

function compileCompaniesSummary(spreadsheets) {
  const companies = {};
  for (const sheet of spreadsheets) {
    const { client_name, created_date, respondees } = sheet;
    if (!client_name) continue;
    if (!companies[client_name]) {
      companies[client_name] = { name: client_name, respondents: new Set(), years: new Set() };
    }
    if (Array.isArray(respondees)) {
      for (const r of respondees) {
        if (r && r.name) companies[client_name].respondents.add(r.name);
      }
    }
    const year = created_date ? String(created_date).slice(0, 4) : null;
    if (year) companies[client_name].years.add(year);
  }
  return Object.values(companies).map((c) => ({
    name: c.name,
    respondents: Array.from(c.respondents),
    years: Array.from(c.years).sort(),
  }));
}

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

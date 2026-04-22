import XLSX from 'xlsx';
import ExcelJS from 'exceljs';
import Candidate from '../models/Candidate.js';

const HEADER_ALIASES = {
  fullName: ['Họ và tên', 'fullName'],
  schoolName: ['Trường cấp 3 đang học', 'Trường', 'schoolName', 'school'],
  examNumber: ['Số báo danh', 'examNumber'],
  examSubjects: ['Môn thi', 'examSubjects'],
  examRoom: ['Phòng Thi', 'examRoom'],
  math: ['Toán', ' Toán ', 'math'],
  physics: ['Vật Lí', 'Lý', 'physics'],
  chemistry: ['Hóa Học', 'Hóa Học ', 'Hóa', 'chemistry']
};

const cleanText = (value) => String(value ?? '').replace(/\u00a0/g, ' ').trim();
const normalizeKey = (value) => cleanText(value).toLowerCase();

const readValue = (row, aliases) => {
  for (const alias of aliases) {
    if (row[alias] !== undefined && row[alias] !== null && row[alias] !== '') {
      return row[alias];
    }
  }

  const normalizedEntries = Object.entries(row).map(([key, value]) => [normalizeKey(key), value]);

  for (const alias of aliases) {
    const matchedEntry = normalizedEntries.find(([key, value]) => key === normalizeKey(alias) && value !== undefined && value !== null && value !== '');

    if (matchedEntry) {
      return matchedEntry[1];
    }
  }

  return '';
};

const normalizeScore = (value) => {
  const text = cleanText(value);

  if (!text) {
    return null;
  }

  const normalized = text.replace(/\s+/g, '').replace(',', '.');
  const parsed = Number(normalized);

  if (!Number.isFinite(parsed)) {
    return null;
  }

  return Math.round(parsed * 100) / 100;
};

export const parseCandidatesFromExcelBuffer = (buffer) => {
  const workbook = XLSX.read(buffer, {
    type: 'buffer',
    cellNF: true,
    cellText: true,
    cellFormula: true
  });
  const firstSheet = workbook.SheetNames[0];

  if (!firstSheet) {
    throw new Error('File Excel không có sheet dữ liệu.');
  }

  const rawRows = XLSX.utils.sheet_to_json(workbook.Sheets[firstSheet], {
    defval: '',
    raw: false
  });

  if (!rawRows.length) {
    throw new Error('File Excel không có dữ liệu.');
  }

  return rawRows
    .map((row, index) => {
      const fullName = cleanText(readValue(row, HEADER_ALIASES.fullName));
      const schoolName = cleanText(readValue(row, HEADER_ALIASES.schoolName));
      const examNumber = cleanText(readValue(row, HEADER_ALIASES.examNumber));
      const rawExamSubjects = cleanText(readValue(row, HEADER_ALIASES.examSubjects));
      const examSubjects = rawExamSubjects || 'Toán, Lý, Hóa';
      const examRoom = cleanText(readValue(row, HEADER_ALIASES.examRoom));
      const scores = {
        math: normalizeScore(readValue(row, HEADER_ALIASES.math)),
        physics: normalizeScore(readValue(row, HEADER_ALIASES.physics)),
        chemistry: normalizeScore(readValue(row, HEADER_ALIASES.chemistry))
      };

      const hasAnyData =
        fullName ||
        schoolName ||
        examNumber ||
        rawExamSubjects ||
        examRoom ||
        Object.values(scores).some((value) => value !== null);

      if (!hasAnyData) {
        return null;
      }

      if (!examNumber || !fullName) {
        throw new Error(`Dòng ${index + 2} thiếu Số báo danh hoặc Họ và tên.`);
      }

      return {
        examNumber,
        fullName,
        schoolName,
        examSubjects,
        examRoom,
        scores
      };
    })
    .filter(Boolean);
};

export const buildCandidatesWorkbookBuffer = async () => {
  const candidates = await Candidate.find().sort({ examNumber: 1 }).lean();
  const workbook = new ExcelJS.Workbook();
  const worksheet = workbook.addWorksheet('Danh sách điểm');

  worksheet.columns = [
    { header: 'STT', key: 'index', width: 8 },
    { header: 'Họ và tên', key: 'fullName', width: 28 },
    { header: 'Trường cấp 3 đang học', key: 'schoolName', width: 32 },
    { header: 'Số báo danh', key: 'examNumber', width: 18 },
    { header: 'Môn thi', key: 'examSubjects', width: 18 },
    { header: 'Phòng Thi', key: 'examRoom', width: 28 },
    { header: 'Toán', key: 'math', width: 12 },
    { header: 'Vật Lí', key: 'physics', width: 12 },
    { header: 'Hóa Học', key: 'chemistry', width: 12 }
  ];

  candidates.forEach((candidate, index) => {
    worksheet.addRow({
      index: index + 1,
      fullName: candidate.fullName,
      schoolName: candidate.schoolName || '',
      examNumber: candidate.examNumber,
      examSubjects: candidate.examSubjects || 'Toán, Lý, Hóa',
      examRoom: candidate.examRoom || '',
      math: candidate.scores?.math ?? '',
      physics: candidate.scores?.physics ?? '',
      chemistry: candidate.scores?.chemistry ?? ''
    });
  });

  worksheet.getRow(1).font = { bold: true };
  worksheet.views = [{ state: 'frozen', ySplit: 1 }];
  ['G', 'H', 'I'].forEach((column) => {
    worksheet.getColumn(column).numFmt = '0.##';
  });

  return workbook.xlsx.writeBuffer();
};

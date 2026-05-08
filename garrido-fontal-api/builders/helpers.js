const { Paragraph, TextRun, Table, TableRow, TableCell, AlignmentType, BorderStyle, WidthType, ShadingType, VerticalAlign, ImageRun, Footer } = require('docx');
const path = require('path');
const fs = require('fs');

// ── Date formatter ────────────────────────────────────────────────
// Converts "2026-04-13" → "13/04/2026"
function formatDate(dateStr) {
  if (!dateStr) return '—';
  const parts = dateStr.split('-');
  if (parts.length === 3) return `${parts[2]}/${parts[1]}/${parts[0]}`;
  return dateStr;
}
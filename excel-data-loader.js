(function () {
  'use strict';

  const CONFIG = Object.assign({
    excelFile: 'ICT-RFA-RFI-lists-New.xlsx',
    dataBranch: 'dashboard-data',
    dataUrl: '',
    refreshMs: 15000
  }, window.LTIA_DASHBOARD_CONFIG || {});

  const SHEET_SPECS = {
    'RFA': {
      type: 'RFA',
      columns: {
        system: 2,
        keyword: 3,
        ref: 4,
        deadline: 5,
        subject: 6,
        subject_vi: 7,
        jcjv_status: 8,
        processing: 9,
        delay: 10,
        withdraw: 11,
        links: [12, 13, 14, 15, 16],
        employer_status: 17,
        vietur_soft: 18,
        vietur_email: 19,
        vietur_hard: 20,
        jcjv_response: 21,
        jcjv_hard: 22,
        signed_by: 23,
        employer_email: 24,
        employer_hard: 25,
        resub_deadline: 26,
        actual_resub: 27,
        resub_assessment: 28,
        resub_delay: 29,
        notes: 30
      }
    },
    'RFA-SDS': {
      type: 'RFA-SDS',
      keyword: 'SDS',
      columns: {
        system: 2,
        ref: 3,
        deadline: 4,
        subject: 5,
        subject_vi: 6,
        jcjv_status: 7,
        processing: 8,
        delay: 9,
        withdraw: 10,
        links: [11, 12, 13, 14, 15],
        employer_status: 16,
        vietur_soft: 17,
        vietur_email: 18,
        vietur_hard: 19,
        jcjv_response: 20,
        jcjv_hard: 21,
        signed_by: 22,
        employer_email: 23,
        employer_hard: 24,
        resub_deadline: 25,
        actual_resub: 26,
        resub_assessment: 27,
        resub_delay: 28,
        notes: 29
      }
    },
    'RFI': {
      type: 'RFI',
      keyword: 'RFI',
      columns: {
        system: 2,
        ref: 3,
        deadline: 4,
        subject: 5,
        subject_vi: 6,
        jcjv_status: 7,
        processing: 8,
        delay: 9,
        withdraw: 10,
        links: [11, 12, 13, 14, 15],
        employer_status: 16,
        vietur_soft: 17,
        vietur_email: 18,
        vietur_hard: 19,
        jcjv_response: 20,
        jcjv_hard: 21,
        signed_by: 22,
        employer_email: 23,
        employer_hard: 24,
        notes: 25
      }
    }
  };

  const EXCEL_EPOCH_UTC = Date.UTC(1899, 11, 30);

  function encodeCell(row, col) {
    return XLSX.utils.encode_cell({
      r: row,
      c: col
    });
  }

  function cell(ws, row, col) {
    return ws[encodeCell(row, col)] || null;
  }

  function value(ws, row, col) {
    const c = cell(ws, row, col);
    return c && c.v !== undefined && c.v !== null ? c.v : '';
  }

  function text(v) {
    if (v === null || v === undefined) return '';
    if (v instanceof Date) return v.toISOString();
    return String(v).trim();
  }

  function number(v) {
    if (typeof v === 'number' && Number.isFinite(v)) return v;
    if (v instanceof Date) {
      return (Date.UTC(v.getUTCFullYear(), v.getUTCMonth(), v.getUTCDate()) - EXCEL_EPOCH_UTC) / 86400000;
    }
    if (typeof v === 'string' && v.trim() !== '') {
      const n = Number(v.trim());
      if (Number.isFinite(n)) return n;
      const m = v.trim().match(/^(\d{1,2})[\\/-](\d{1,2})[\\/-](\d{4})$/);
      if (m) {
        const d = Date.UTC(Number(m[3]), Number(m[2]) - 1, Number(m[1]));
        return Math.round((d - EXCEL_EPOCH_UTC) / 86400000);
      }
    }
    return null;
  }

  function raw(v) {
    const n = number(v);
    if (n !== null && (typeof v === 'number' || v instanceof Date || /^\s*\d+(?:\.\d+)?\s*$/.test(String(v)))) {
      return String(Number.isInteger(n) ? n : n);
    }
    return text(v);
  }

  function dateText(v) {
    const n = number(v);
    if (n === null || !Number.isFinite(n)) return '';
    const d = new Date(EXCEL_EPOCH_UTC + Math.round(n) * 86400000);
    const dd = String(d.getUTCDate()).padStart(2, '0');
    const mm = String(d.getUTCMonth() + 1).padStart(2, '0');
    const yyyy = d.getUTCFullYear();
    return `${dd}/${mm}/${yyyy}`;
  }

  function todaySerial() {
    const now = new Date();
    const midnightUtc = Date.UTC(now.getFullYear(), now.getMonth(), now.getDate());
    return Math.floor((midnightUtc - EXCEL_EPOCH_UTC) / 86400000);
  }

  function shortHash(buffer) {
    const bytes = new Uint8Array(buffer);
    let hash = 2166136261;
    for (let i = 0; i < bytes.length; i++) {
      hash ^= bytes[i];
      hash = Math.imul(hash, 16777619);
    }
    return (hash >>> 0).toString(16).padStart(8, '0');
  }

  function getLink(ws, row, col) {
    const c = cell(ws, row, col);
    if (!c) return '';
    if (c.l) {
      if (c.l.Target) return String(c.l.Target).trim();
      if (c.l.Location) return String(c.l.Location).trim();
      if (c.l.href) return String(c.l.href).trim();
    }
    return '';
  }

  function baseOf(ref) {
    return text(ref).replace(/[A-Z]+$/, '');
  }

  function currentHolder(record) {
    const w = String(record.withdraw || '').trim().toUpperCase();
    if (w === 'X' || w === 'O') return 'Withdrawn / Changed';

    const emp = String(record.employer_status || '').trim().toUpperCase();
    if (record.type === 'RFI') {
      if (emp === 'RBE') return 'Closed';
    } else {
      if (emp === 'A') return 'Closed';
      if ((emp === 'B' || emp === 'C') && record.resub_deadline_raw) return 'Vietur - Resubmission';
    }

    if (!record.jcjv_response_raw && !record.jcjv_hard_raw) return 'JCJV';
    return 'Employer';
  }

  function buildRecord(ws, row, spec) {
    const c = spec.columns;
    const ref = text(value(ws, row, c.ref));
    const tt = value(ws, row, 0);
    const ttText = String(tt || '').trim();
    const ttIsFormula = ttText.startsWith('=');

    if (
      !ref ||
      (!ttIsFormula &&
      (ttText === '' || !Number.isFinite(Number(tt))))
    ) return null;

    const record = {
      type: spec.type,
      system: text(value(ws, row, c.system)),
      keyword: spec.keyword !== undefined ? spec.keyword : text(value(ws, row, c.keyword)),
      ref,
      deadline_raw: raw(value(ws, row, c.deadline)),
      subject: text(value(ws, row, c.subject)),
      subject_vi: text(value(ws, row, c.subject_vi)),
      jcjv_status_display: text(value(ws, row, c.jcjv_status)),
      processing: raw(value(ws, row, c.processing)),
      delay: raw(value(ws, row, c.delay)),
      withdraw: text(value(ws, row, c.withdraw)),
      vietur_soft_raw: raw(value(ws, row, c.vietur_soft)),
      vietur_email_raw: raw(value(ws, row, c.vietur_email)),
      vietur_hard_raw: raw(value(ws, row, c.vietur_hard)),
      jcjv_response_raw: raw(value(ws, row, c.jcjv_response)),
      jcjv_hard_raw: raw(value(ws, row, c.jcjv_hard)),
      signed_by: text(value(ws, row, c.signed_by)),
      employer_email_raw: raw(value(ws, row, c.employer_email)),
      employer_hard_raw: raw(value(ws, row, c.employer_hard)),
      employer_status: text(value(ws, row, c.employer_status)),
      resub_deadline_raw: c.resub_deadline === undefined ? '' : raw(value(ws, row, c.resub_deadline)),
      actual_resub_raw: c.actual_resub === undefined ? '' : raw(value(ws, row, c.actual_resub)),
      resub_assessment: c.resub_assessment === undefined ? '' : text(value(ws, row, c.resub_assessment)),
      resub_delay: c.resub_delay === undefined ? '' : raw(value(ws, row, c.resub_delay)),
      notes: text(value(ws, row, c.notes)),
      links: {
        vietur_official: getLink(ws, row, c.links[0]),
        vietur_internal: getLink(ws, row, c.links[1]),
        jcjv_official: getLink(ws, row, c.links[2]),
        jcjv_internal: getLink(ws, row, c.links[3]),
        employer_official: getLink(ws, row, c.links[4])
      }
    };

    record.base = baseOf(record.ref);
    record.latest = '0';
    record.need_submission = null;
    record.deadline_serial = number(value(ws, row, c.deadline));
    record.deadline = dateText(value(ws, row, c.deadline));
    record.vietur_soft = dateText(value(ws, row, c.vietur_soft));
    record.vietur_soft_serial = number(value(ws, row, c.vietur_soft));
    record.vietur_email = dateText(value(ws, row, c.vietur_email));
    record.vietur_email_serial = number(value(ws, row, c.vietur_email));
    record.vietur_hard = dateText(value(ws, row, c.vietur_hard));
    record.vietur_hard_serial = number(value(ws, row, c.vietur_hard));
    record.jcjv_response = dateText(value(ws, row, c.jcjv_response));
    record.jcjv_response_serial = number(value(ws, row, c.jcjv_response));
    record.jcjv_hard = dateText(value(ws, row, c.jcjv_hard));
    record.jcjv_hard_serial = number(value(ws, row, c.jcjv_hard));
    record.employer_email = dateText(value(ws, row, c.employer_email));
    record.employer_email_serial = number(value(ws, row, c.employer_email));
    record.employer_hard = dateText(value(ws, row, c.employer_hard));
    record.employer_hard_serial = number(value(ws, row, c.employer_hard));
    record.resub_deadline = dateText(value(ws, row, c.resub_deadline));
    record.resub_deadline_serial = number(value(ws, row, c.resub_deadline));
    record.actual_resub = dateText(value(ws, row, c.actual_resub));
    record.actual_resub_serial = number(value(ws, row, c.actual_resub));

    const hard = record.vietur_hard_serial;
    const rawDeadline = record.deadline_serial;
    const deadlineCalcSerial = rawDeadline !== null ? rawDeadline : (hard !== null ? hard + 21 : null);
    record.deadline_calc = deadlineCalcSerial === null ? '' : dateText(deadlineCalcSerial);

    const jcEnd = record.jcjv_hard_serial !== null ? record.jcjv_hard_serial : record.jcjv_response_serial;
    record.jcjv_days_calc = hard === null ? 0 : Math.max(0, Math.round((jcEnd !== null ? jcEnd : todaySerial()) - hard));

    if (deadlineCalcSerial === null) {
      record.response_assessment = '';
      record.response_delay_calc = 0;
    } else if (record.jcjv_hard_serial !== null) {
      record.response_delay_calc = Math.max(0, Math.round(record.jcjv_hard_serial - deadlineCalcSerial));
      record.response_assessment = record.jcjv_hard_serial > deadlineCalcSerial ? 'Late response' : 'Timely response';
    } else if (record.jcjv_response_serial !== null) {
      record.response_assessment = 'Pending hard copy';
      record.response_delay_calc = 0;
    } else {
      record.response_delay_calc = Math.max(0, Math.round(todaySerial() - deadlineCalcSerial));
      record.response_assessment = todaySerial() > deadlineCalcSerial ? 'Overdue / Pending' : 'Pending JCJV';
    }

    record.current_holder = currentHolder(record);
    if (record.current_holder === 'Vietur - Resubmission') record.need_submission = 'YES';
    record.is_latest_calc = false;
    return record;
  }

  function parseWorkbook(buffer) {
    const workbook = XLSX.read(buffer, {
      type: 'array',
      cellDates: false,
      cellHTML: false,
      cellStyles: true
    });
    const records = [];

    Object.keys(SHEET_SPECS).forEach(sheetName => {
      const spec = SHEET_SPECS[sheetName];
      const ws = workbook.Sheets[sheetName];
      if (!ws || !ws['!ref']) return;
      const range = XLSX.utils.decode_range(ws['!ref']);
      for (let r = Math.max(range.s.r, 6); r <= range.e.r; r++) {
        const record = buildRecord(ws, r, spec);
        if (record) records.push(record);
      }
    });

    const groups = new Map();
    records.forEach(record => {
      if (!groups.has(record.base)) groups.set(record.base, []);
      groups.get(record.base).push(record);
    });

    groups.forEach(group => {
      const active = group.filter(
        r => String(r.withdraw || '').trim().toUpperCase() !== 'X' &&
        String(r.withdraw || '').trim().toUpperCase() !== 'O'
      );
      const latest = active.length ? active[active.length - 1] : null;
      if (latest) latest.latest = '1';
      group.forEach(r => {
        r.is_latest_calc = r === latest;
        r.related = group.map(item => ({
          ref: item.ref,
          status: String(item.jcjv_status_display || '').trim() || '—'
        }));
      });
    });

    const sourceDateCell = workbook.Sheets.RFA && workbook.Sheets.RFA['J1'] ? workbook.Sheets.RFA['J1'].v : null;
    const sourceDate = dateText(sourceDateCell);

    return {
      version: shortHash(buffer),
      source_modified: sourceDate,
      published_at: new Date().toLocaleString('en-GB', {
        hour12: false
      }),
      records,
      error: ''
    };
  }

  function githubRawUrl(branch) {
    const host = String(location.hostname || '');
    if (!/\.github\.io$/i.test(host)) return '';
    const owner = host.split('.')[0];
    const pathParts = location.pathname.split('/').filter(Boolean);
    const repo = pathParts.length ? pathParts[0] : `${owner}.github.io`;
    return `https://raw.githubusercontent.com/${encodeURIComponent(owner)}/${encodeURIComponent(repo)}/${encodeURIComponent(branch)}/${CONFIG.excelFile.split('/').map(encodeURIComponent).join('/')}`;
  }

  function candidates() {
    const urls = [];
    if (CONFIG.dataUrl) urls.push(CONFIG.dataUrl);
    const primary = githubRawUrl(CONFIG.dataBranch);
    if (primary) urls.push(primary);

    // Local preview only. GitHub Pages must read from dashboard-data.
    if (location.protocol === 'file:' || /^(localhost|127\.0\.0\.1)$/i.test(location.hostname)) {
      urls.push(`./${CONFIG.excelFile.split('/').map(encodeURIComponent).join('/')}`);
    }
    return [...new Set(urls)];
  }

  function bust(url) {
    const join = url.includes('?') ? '&' : '?';
    return `${url}${join}ltia_excel_cache=${Date.now()}`;
  }

  async function fetchWorkbook() {
    const errors = [];
    for (const url of candidates()) {
      try {
        const response = await fetch(bust(url), {
          cache: 'no-store'
        });
        if (!response.ok) throw new Error(`HTTP ${response.status}`);
        const buffer = await response.arrayBuffer();
        if (!buffer.byteLength) throw new Error('Empty workbook');
        return {
          payload: parseWorkbook(buffer),
          url
        };
      } catch (err) {
        errors.push(`${url}: ${err.message}`);
      }
    }
    throw new Error(`Cannot load Excel source. ${errors.join(' | ')}`);
  }

  window.LTIAExcelLoader = {
    load: fetchWorkbook,
    config: CONFIG
  };
})();
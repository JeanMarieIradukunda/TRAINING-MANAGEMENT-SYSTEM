/**
 * Training Management System — generator.js
 * ------------------------------------------------------------------------
 * Small shared helpers used by the two public, no-login generator pages:
 *   - scheme_of_work_user.html
 *   - lesson_plan_user.html
 *
 * Both pages receive the *entire* live curriculum (sectors, trades,
 * levels, trade/level links, modules, learning outcomes, indicative
 * contents, trainers, logos) as one JSON payload rendered into a
 * <script type="application/json" id="tms-data"> tag by the view. These
 * helpers turn that payload into cascading <select> options and provide
 * a couple of formatting/print utilities so the page-specific scripts
 * stay focused on layout, not plumbing.
 * ------------------------------------------------------------------------
 */

(function () {
  'use strict';

  function readData(elementId) {
    var el = document.getElementById(elementId || 'tms-data');
    if (!el) return null;
    try {
      return JSON.parse(el.textContent);
    } catch (e) {
      console.error('TMSGen: could not parse curriculum data payload', e);
      return null;
    }
  }

  function escapeHtml(value) {
    if (value === null || value === undefined) return '';
    return String(value)
      .replace(/&/g, '&amp;')
      .replace(/</g, '&lt;')
      .replace(/>/g, '&gt;')
      .replace(/"/g, '&quot;')
      .replace(/'/g, '&#39;');
  }

  // Turns "Learning outcome 1: Plan VoIP system installation" into
  // "Plan VoIP system installation" for places that want the short form.
  function shortOutcomeText(text) {
    if (!text) return '';
    var parts = String(text).split(':');
    return parts.length > 1 ? parts.slice(1).join(':').trim() : String(text).trim();
  }

  /**
   * Replaces a <select>'s options with a placeholder plus one <option>
   * per item.
   *   populateSelect(selectEl, items, {
   *     valueKey: 'id', labelFn: function(item) { return item.name; },
   *     placeholder: 'Select trade', disabled: false
   *   });
   */
  function populateSelect(selectEl, items, opts) {
    if (!selectEl) return;
    opts = opts || {};
    var valueKey = opts.valueKey || 'id';
    var labelFn = opts.labelFn || function (item) { return item.label || item.name || ''; };
    var placeholder = opts.placeholder || 'Select';

    selectEl.innerHTML = '';
    var placeholderOpt = document.createElement('option');
    placeholderOpt.value = '';
    placeholderOpt.textContent = placeholder;
    selectEl.appendChild(placeholderOpt);

    (items || []).forEach(function (item) {
      var opt = document.createElement('option');
      opt.value = item[valueKey];
      opt.textContent = labelFn(item);
      selectEl.appendChild(opt);
    });

    selectEl.disabled = !items || items.length === 0;
  }

  function todayLong() {
    var d = new Date();
    return d.toLocaleDateString(undefined, { year: 'numeric', month: 'long', day: 'numeric' });
  }

  // Today's date as "YYYY-MM-DD", suitable for pre-filling an
  // <input type="date">. Built from local date parts (not toISOString)
  // so the value matches the user's own calendar day, not UTC's.
  function todayISO() {
    var d = new Date();
    var mm = String(d.getMonth() + 1).padStart(2, '0');
    var dd = String(d.getDate()).padStart(2, '0');
    return d.getFullYear() + '-' + mm + '-' + dd;
  }

  function formatDateLong(isoString) {
    if (!isoString) return '\u2014';
    var d = new Date(isoString + 'T00:00:00');
    if (isNaN(d.getTime())) return isoString;
    return d.toLocaleDateString(undefined, { year: 'numeric', month: 'long', day: 'numeric' });
  }

  function printDocument() {
    window.print();
  }

  // Reads the csrftoken cookie Django sets (see ensure_csrf_cookie on the
  // generator page views) so export/AI fetch() calls can send it back as
  // the X-CSRFToken header.
  function getCookie(name) {
    var match = document.cookie.match('(^|;)\\s*' + name + '\\s*=\\s*([^;]+)');
    return match ? decodeURIComponent(match.pop()) : '';
  }

  /**
   * Walks a rendered generator document (the <article class="gen-doc">
   * element) and turns it into a plain JSON structure the server can
   * rebuild as a Word/Excel file - the SAME data already on screen,
   * captured straight out of the DOM so it always matches (including any
   * edits the trainer has made) without duplicating the generation logic
   * on the backend.
   *
   *   { title, subtitle, meta: [[label, value], ...],
   *     tables: [ [ {section, cells:[{text,colspan,header}, ...]}, ... ], ... ],
   *     signoff: [{label, name}, ...], footer }
   *
   * Rows/cells carrying the "no-print" class (screen-only pagination
   * controls etc.) are skipped, matching what actually ends up on a
   * printed copy. Rows hidden via inline display:none (paginated week
   * rows) are still included, same as the @media print override does.
   */
  function collectExportPayload(genDocEl) {
    var result = { title: '', subtitle: '', meta: [], tables: [], signoff: [], footer: '' };
    if (!genDocEl) return result;

    var titleEl = genDocEl.querySelector('.gen-doc-title h3');
    if (titleEl) result.title = titleEl.textContent.trim();

    var subtitleEl = genDocEl.querySelector('.gen-doc-title p');
    if (subtitleEl) result.subtitle = subtitleEl.textContent.trim();

    var metaEl = genDocEl.querySelector('.gen-meta-grid');
    if (metaEl) {
      var dts = metaEl.querySelectorAll('dt');
      var dds = metaEl.querySelectorAll('dd');
      for (var i = 0; i < dts.length; i++) {
        result.meta.push([dts[i].textContent.trim(), dds[i] ? dds[i].textContent.trim() : '']);
      }
    }

    genDocEl.querySelectorAll('.gen-table-wrap table, table.gen-sp-table').forEach(function (table) {
      var rows = [];
      table.querySelectorAll('tr').forEach(function (tr) {
        if (tr.classList.contains('no-print')) return;
        var isSection = /gen-term-separator|gen-sp-section-row|gen-sp-subhead-row/.test(tr.className || '');
        var cells = [];
        tr.querySelectorAll('th, td').forEach(function (cell) {
          if (cell.classList.contains('no-print')) return;
          cells.push({
            text: (cell.innerText || cell.textContent || '').trim(),
            colspan: parseInt(cell.getAttribute('colspan'), 10) || 1,
            header: cell.tagName === 'TH',
          });
        });
        if (cells.length) rows.push({ section: isSection, cells: cells });
      });
      if (rows.length) result.tables.push(rows);
    });

    genDocEl.querySelectorAll('.gen-signoff-grid .gen-signoff-line').forEach(function (line) {
      var divs = line.querySelectorAll(':scope > div');
      result.signoff.push({
        label: divs[0] ? divs[0].textContent.trim() : '',
        name: divs[1] ? divs[1].textContent.trim() : '',
      });
    });

    var footerEl = genDocEl.querySelector('.gen-doc-footer');
    if (footerEl) result.footer = footerEl.textContent.trim();

    return result;
  }

  /**
   * POSTs a collectExportPayload() payload to a server export endpoint
   * and downloads the binary (.docx/.xlsx) response it streams back.
   * Returns a Promise so callers can toggle a loading state and surface
   * errors (the server returns a JSON {error: "..."} body on failure).
   */
  function exportDocument(url, payload, filename) {
    return fetch(url, {
      method: 'POST',
      headers: {
        'Content-Type': 'application/json',
        'X-CSRFToken': getCookie('csrftoken'),
      },
      body: JSON.stringify(payload),
    }).then(function (resp) {
      if (!resp.ok) {
        return resp.json().catch(function () { return {}; }).then(function (data) {
          throw new Error(data.error || ('Export failed (HTTP ' + resp.status + ')'));
        });
      }
      return resp.blob();
    }).then(function (blob) {
      var objectUrl = URL.createObjectURL(blob);
      var link = document.createElement('a');
      link.href = objectUrl;
      link.download = filename;
      document.body.appendChild(link);
      link.click();
      document.body.removeChild(link);
      URL.revokeObjectURL(objectUrl);
    });
  }

  // Browsers use document.title as the default filename when the person
  // chooses "Save as PDF" from the print dialog, so we sanitise arbitrary
  // curriculum text into something safe to use as a file name: strip
  // characters that are illegal (or awkward) in file names on Windows/
  // macOS/Linux, collapse whitespace, and cap the length.
  function filenameSafe(text, fallback) {
    var clean = String(text || '')
      .replace(/[\\/:*?"<>|]/g, '-')
      .replace(/\s+/g, ' ')
      .trim();
    if (clean.length > 120) clean = clean.slice(0, 120).trim();
    return clean || fallback || 'Document';
  }

  window.TMSGen = {
    readData: readData,
    escapeHtml: escapeHtml,
    shortOutcomeText: shortOutcomeText,
    populateSelect: populateSelect,
    todayLong: todayLong,
    todayISO: todayISO,
    formatDateLong: formatDateLong,
    filenameSafe: filenameSafe,
    print: printDocument,
    getCookie: getCookie,
    collectExportPayload: collectExportPayload,
    exportDocument: exportDocument,
  };
})();
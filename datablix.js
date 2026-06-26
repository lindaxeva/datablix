/*
 * datablix.js
 *
 * Browser port of the datablix analysis logic. It mirrors the Python tool:
 * the same column type detection, freshness report, and categorical glossary.
 * All work happens in memory, so a CSV opened in the web page never leaves
 * the visitor's computer.
 *
 * Works in the browser (attaches to window.datablix) and in Node (exports),
 * so the same file can be unit tested from the command line.
 */
(function (root) {
  "use strict";

  var TYPE_CONFIDENCE = 0.8;

  var MONTHS = {
    jan: 1, january: 1, feb: 2, february: 2, mar: 3, march: 3,
    apr: 4, april: 4, may: 5, jun: 6, june: 6, jul: 7, july: 7,
    aug: 8, august: 8, sep: 9, sept: 9, september: 9, oct: 10, october: 10,
    nov: 11, november: 11, dec: 12, december: 12
  };

  function daysInMonth(y, m) {
    return new Date(Date.UTC(y, m, 0)).getUTCDate();
  }

  function validYmd(y, m, d) {
    if (m < 1 || m > 12) return false;
    if (d < 1 || d > daysInMonth(y, m)) return false;
    return { y: y, m: m, d: d };
  }

  // Try to read a value as a calendar date. Returns {y, m, d} or null.
  // Format order matches the Python tool, including trying day/month before
  // month/day for slash and dash dates.
  function parseDate(value) {
    var text = String(value).trim();
    if (!text) return null;
    var m;

    m = text.match(/^(\d{4})-(\d{2})-(\d{2})$/);
    if (m) return validYmd(+m[1], +m[2], +m[3]);

    m = text.match(/^(\d{4})\/(\d{2})\/(\d{2})$/);
    if (m) return validYmd(+m[1], +m[2], +m[3]);

    m = text.match(/^(\d{1,2})-(\d{1,2})-(\d{4})$/);
    if (m) return validYmd(+m[3], +m[2], +m[1]) || validYmd(+m[3], +m[1], +m[2]);

    m = text.match(/^(\d{1,2})\/(\d{1,2})\/(\d{4})$/);
    if (m) return validYmd(+m[3], +m[2], +m[1]) || validYmd(+m[3], +m[1], +m[2]);

    m = text.match(/^(\d{4})-(\d{2})-(\d{2})[ T]\d{2}:\d{2}:\d{2}$/);
    if (m) return validYmd(+m[1], +m[2], +m[3]);

    // 12 Jan 2026  /  12 January 2026
    m = text.match(/^(\d{1,2})\s+([A-Za-z]+)\s+(\d{4})$/);
    if (m && MONTHS[m[2].toLowerCase()]) {
      return validYmd(+m[3], MONTHS[m[2].toLowerCase()], +m[1]);
    }

    // Jan 12, 2026  /  January 12, 2026
    m = text.match(/^([A-Za-z]+)\s+(\d{1,2}),\s*(\d{4})$/);
    if (m && MONTHS[m[1].toLowerCase()]) {
      return validYmd(+m[3], MONTHS[m[1].toLowerCase()], +m[2]);
    }

    m = text.match(/^(\d{4})(\d{2})(\d{2})$/);
    if (m) return validYmd(+m[1], +m[2], +m[3]);

    return null;
  }

  function isoString(ymd) {
    function pad(n) { return (n < 10 ? "0" : "") + n; }
    return ymd.y + "-" + pad(ymd.m) + "-" + pad(ymd.d);
  }

  function daysBetween(fromYmd, toYmd) {
    var a = Date.UTC(fromYmd.y, fromYmd.m - 1, fromYmd.d);
    var b = Date.UTC(toYmd.y, toYmd.m - 1, toYmd.d);
    return Math.round((b - a) / 86400000);
  }

  function parseNumber(value) {
    var text = String(value).trim();
    if (!text) return null;
    var cleaned = text.replace(/,/g, "").replace(/\$/g, "").replace(/%/g, "").trim();
    if (cleaned === "" || cleaned === "-" || cleaned === ".") return null;
    var n = Number(cleaned);
    return isNaN(n) ? null : n;
  }

  function normalizeLabel(value) {
    return String(value).trim().toLowerCase().replace(/\s+/g, " ");
  }

  // Minimal RFC-4180-ish CSV parser: handles quoted fields, escaped quotes,
  // commas and newlines inside quotes, and both LF and CRLF line endings.
  function parseCsv(text) {
    var rows = [];
    var row = [];
    var field = "";
    var inQuotes = false;
    var i = 0;
    var c;

    while (i < text.length) {
      c = text[i];
      if (inQuotes) {
        if (c === '"') {
          if (text[i + 1] === '"') { field += '"'; i += 2; continue; }
          inQuotes = false; i++; continue;
        }
        field += c; i++; continue;
      }
      if (c === '"') { inQuotes = true; i++; continue; }
      if (c === ",") { row.push(field); field = ""; i++; continue; }
      if (c === "\r") { i++; continue; }
      if (c === "\n") { row.push(field); rows.push(row); row = []; field = ""; i++; continue; }
      field += c; i++;
    }
    // Flush the final field/row if the file did not end with a newline.
    if (field !== "" || row.length > 0) { row.push(field); rows.push(row); }

    if (!rows.length) return { headers: [], rows: [] };

    var headers = rows[0].map(function (h) { return h.trim(); });
    var data = [];
    for (var r = 1; r < rows.length; r++) {
      // Skip fully blank trailing lines.
      if (rows[r].length === 1 && rows[r][0].trim() === "") continue;
      var record = {};
      for (var col = 0; col < headers.length; col++) {
        var raw = col < rows[r].length ? rows[r][col] : "";
        record[headers[col]] = String(raw).trim();
      }
      data.push(record);
    }
    return { headers: headers, rows: data };
  }

  function classifyColumns(headers, rows, maxCategories) {
    var types = {};
    headers.forEach(function (header) {
      var values = rows.map(function (r) { return r[header]; });
      var nonEmpty = values.filter(function (v) { return v !== ""; });

      if (!nonEmpty.length) { types[header] = "empty"; return; }

      var dateHits = nonEmpty.filter(function (v) { return parseDate(v) !== null; }).length;
      if (dateHits / nonEmpty.length >= TYPE_CONFIDENCE) { types[header] = "date"; return; }

      var numberHits = nonEmpty.filter(function (v) { return parseNumber(v) !== null; }).length;
      var isNumeric = numberHits / nonEmpty.length >= TYPE_CONFIDENCE;

      var distinct = new Set(nonEmpty).size;
      var looksCategorical = distinct <= maxCategories && distinct < nonEmpty.length;

      if (isNumeric && !looksCategorical) types[header] = "numeric";
      else if (looksCategorical) types[header] = "categorical";
      else types[header] = "text";
    });
    return types;
  }

  function round1(n) { return Math.round(n * 10) / 10; }

  function buildFreshness(headers, rows, types, today, freshDays, staleDays) {
    var dateColumns = headers.filter(function (h) { return types[h] === "date"; });
    var summary = [];

    dateColumns.forEach(function (header) {
      var parsed = rows.map(function (r) { return parseDate(r[header]); })
                       .filter(function (d) { return d !== null; });
      if (!parsed.length) return;

      var newest = parsed[0], oldest = parsed[0];
      parsed.forEach(function (d) {
        if (daysBetween(d, newest) < 0) newest = d;
        if (daysBetween(oldest, d) < 0) oldest = d;
      });

      var daysSince = daysBetween(newest, today);
      var status = daysSince <= freshDays ? "Fresh"
                 : daysSince <= staleDays ? "Aging" : "Stale";

      var staleRows = parsed.filter(function (d) {
        return daysBetween(d, today) > staleDays;
      }).length;

      summary.push({
        column: header,
        newest: isoString(newest),
        oldest: isoString(oldest),
        days_since_newest: daysSince,
        status: status,
        rows_with_dates: parsed.length,
        stale_rows: staleRows,
        stale_share_pct: round1(100 * staleRows / parsed.length)
      });
    });

    var overall = null;
    if (summary.length) {
      var freshest = summary.reduce(function (best, c) {
        return c.days_since_newest < best.days_since_newest ? c : best;
      });
      overall = {
        reference_date: isoString(today),
        most_recent_value: freshest.newest,
        days_since: freshest.days_since_newest,
        status: freshest.status,
        driven_by_column: freshest.column
      };
    }

    return {
      reference_date: isoString(today),
      fresh_threshold_days: freshDays,
      stale_threshold_days: staleDays,
      overall: overall,
      date_columns: summary
    };
  }

  function buildGlossary(headers, rows, types, maxCategories, topN) {
    var glossary = [];
    headers.forEach(function (header) {
      if (types[header] !== "categorical") return;

      var values = rows.map(function (r) { return r[header]; });
      var nonEmpty = values.filter(function (v) { return v !== ""; });
      var missing = values.length - nonEmpty.length;

      var counts = {};
      nonEmpty.forEach(function (v) { counts[v] = (counts[v] || 0) + 1; });

      // Sort by count descending, keeping first-seen order on ties (stable).
      var order = {};
      nonEmpty.forEach(function (v, i) { if (!(v in order)) order[v] = i; });
      var ranked = Object.keys(counts).sort(function (a, b) {
        return counts[b] - counts[a] || order[a] - order[b];
      });
      var total = nonEmpty.length;

      var entries = ranked.slice(0, topN).map(function (v) {
        return { value: v, count: counts[v], share_pct: total ? round1(100 * counts[v] / total) : 0 };
      });

      var groups = {};
      Object.keys(counts).forEach(function (v) {
        var key = normalizeLabel(v);
        (groups[key] = groups[key] || []).push(v);
      });
      var inconsistencies = Object.keys(groups)
        .map(function (k) { return groups[k].slice().sort(); })
        .filter(function (g) { return g.length > 1; });

      glossary.push({
        column: header,
        distinct_values: Object.keys(counts).length,
        missing_cells: missing,
        missing_share_pct: values.length ? round1(100 * missing / values.length) : 0,
        top_values: entries,
        more_values_not_shown: Math.max(0, ranked.length - topN),
        inconsistent_value_groups: inconsistencies
      });
    });
    return glossary;
  }

  function analyze(text, options) {
    options = options || {};
    var maxCategories = options.maxCategories || 20;
    var topN = options.top || 10;
    var freshDays = options.freshDays || 30;
    var staleDays = options.staleDays || 90;
    var today = options.today || todayYmd();

    var parsed = parseCsv(text);
    var types = classifyColumns(parsed.headers, parsed.rows, maxCategories);

    return {
      meta: {
        rows: parsed.rows.length,
        columns: parsed.headers.length,
        column_types: types
      },
      freshness: buildFreshness(parsed.headers, parsed.rows, types, today, freshDays, staleDays),
      glossary: buildGlossary(parsed.headers, parsed.rows, types, maxCategories, topN)
    };
  }

  function todayYmd() {
    var d = new Date();
    return { y: d.getFullYear(), m: d.getMonth() + 1, d: d.getDate() };
  }

  function ymdFromIso(text) {
    var m = String(text).match(/^(\d{4})-(\d{2})-(\d{2})$/);
    return m ? { y: +m[1], m: +m[2], d: +m[3] } : null;
  }

  var api = {
    parseDate: parseDate,
    parseNumber: parseNumber,
    normalizeLabel: normalizeLabel,
    parseCsv: parseCsv,
    classifyColumns: classifyColumns,
    buildFreshness: buildFreshness,
    buildGlossary: buildGlossary,
    analyze: analyze,
    todayYmd: todayYmd,
    ymdFromIso: ymdFromIso,
    isoString: isoString,
    daysBetween: daysBetween
  };

  if (typeof module !== "undefined" && module.exports) module.exports = api;
  else root.datablix = api;

})(typeof self !== "undefined" ? self : this);

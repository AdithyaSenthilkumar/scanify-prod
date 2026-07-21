/* ============================================================================
 * Scanify — Shared report PDF letterhead template
 * ----------------------------------------------------------------------------
 * Produces a clean, company-letterhead PDF that matches the printed Stedman
 * ranking/target sheets: STEDMAN logo (top-right), centred company name and
 * Chennai address, a bold-underlined report title, one bordered data table,
 * and the "With our Best Wishes" closing. Used by the Ranking Reports and the
 * Sales Target (year-wise) report so every report shares one look.
 *
 * Requires jsPDF (umd) + jspdf-autotable to be loaded first. Uses only the
 * globals they expose — no other dependencies.
 * ==========================================================================*/
(function () {
    "use strict";

    var LOGO_URL = "/assets/scanify/images/stedman_logo.png";
    var LOGO_DATA = null;          // base64 data URL, filled by preload
    var LOGO_RATIO = 337 / 184;    // native aspect ratio of the source PNG

    // Preload the logo once so addImage() has synchronous data at export time.
    (function preloadLogo() {
        try {
            var img = new Image();
            img.crossOrigin = "anonymous";
            img.onload = function () {
                try {
                    var c = document.createElement("canvas");
                    c.width = img.naturalWidth || img.width;
                    c.height = img.naturalHeight || img.height;
                    if (c.width && c.height) LOGO_RATIO = c.width / c.height;
                    c.getContext("2d").drawImage(img, 0, 0);
                    LOGO_DATA = c.toDataURL("image/png");
                } catch (e) { /* tainted canvas — skip logo silently */ }
            };
            img.src = LOGO_URL;
        } catch (e) { /* no-op */ }
    })();

    // jsPDF's default Helvetica lacks ₹, en/em dashes and smart quotes — map them
    // to safe ASCII so the PDF never shows tofu/garbled glyphs.
    function ascii(s) {
        if (s == null) return "";
        return String(s)
            .replace(/₹/g, "Rs.")
            .replace(/[–—]/g, "-")
            .replace(/[‘’]/g, "'")
            .replace(/[“”]/g, '"')
            .replace(/•/g, "-")
            .replace(/ /g, " ");
    }

    var COMPANY_NAME = "STEDMAN PHARMACEUTICALS PVT LTD., CHENNAI";
    var COMPANY_ADDR = "14A, Nehru Nagar, 3rd Cross Street, Kottivakkam, Chennai - 600 041";
    var HEADER_RESERVE = 33;       // vertical space (mm) the letterhead occupies

    // Named row styles a caller can attach to <tr> classes (see extractTable).
    var ROW_STYLES = {
        dark:       { fillColor: [30, 41, 59],  textColor: [255, 255, 255], fontStyle: "bold" },
        group:      { fillColor: [226, 232, 240], fontStyle: "bold" },
        subtotal:   { fillColor: [255, 251, 235], fontStyle: "bold" },
        total:      { fillColor: [219, 234, 254], fontStyle: "bold" },
        grandtotal: { fillColor: [30, 41, 59],  textColor: [255, 255, 255], fontStyle: "bold" }
    };

    function drawHeader(doc, title, subtitle) {
        var pageW = doc.internal.pageSize.getWidth();
        var margin = 10;

        if (LOGO_DATA) {
            try {
                var lh = 13, lw = lh * LOGO_RATIO;
                doc.addImage(LOGO_DATA, "PNG", pageW - margin - lw, 5.5, lw, lh);
            } catch (e) { /* skip */ }
        }

        doc.setFont("helvetica", "bold");
        doc.setFontSize(13);
        doc.setTextColor(15, 23, 42);
        doc.text(COMPANY_NAME, pageW / 2, 11, { align: "center" });

        doc.setFont("helvetica", "normal");
        doc.setFontSize(8.5);
        doc.setTextColor(90);
        doc.text(COMPANY_ADDR, pageW / 2, 16, { align: "center" });

        var t = ascii((title || "").toUpperCase());
        doc.setFont("helvetica", "bold");
        doc.setFontSize(11);
        doc.setTextColor(0);
        doc.text(t, pageW / 2, 24, { align: "center" });
        var tw = doc.getTextWidth(t);
        doc.setLineWidth(0.4);
        doc.setDrawColor(0);
        doc.line(pageW / 2 - tw / 2, 25.4, pageW / 2 + tw / 2, 25.4);

        if (subtitle) {
            doc.setFont("helvetica", "normal");
            doc.setFontSize(8.5);
            doc.setTextColor(90);
            doc.text(ascii(subtitle), pageW / 2, 29.6, { align: "center" });
        }
        doc.setTextColor(0);
        doc.setDrawColor(0);
    }

    function drawClosing(doc, y) {
        var margin = 10;
        doc.setFont("helvetica", "normal");
        doc.setFontSize(9);
        doc.setTextColor(0);
        doc.text("With our Best Wishes,", margin, y);
        doc.text("for Stedman Pharmaceuticals Pvt Ltd.,", margin, y + 5);
    }

    function drawFooters(doc, division) {
        var pages = doc.internal.getNumberOfPages();
        var pageW = doc.internal.pageSize.getWidth();
        var pageH = doc.internal.pageSize.getHeight();
        var margin = 10;
        var stamp = "Stedman Pharmaceuticals" + (division ? "  —  " + division + " Division" : "") +
                    "  |  Generated " + new Date().toLocaleDateString("en-IN");
        for (var i = 1; i <= pages; i++) {
            doc.setPage(i);
            doc.setFont("helvetica", "normal");
            doc.setFontSize(7.5);
            doc.setTextColor(140);
            doc.text(ascii(stamp), margin, pageH - 6);
            doc.text("Page " + i + " of " + pages, pageW - margin, pageH - 6, { align: "right" });
        }
        doc.setTextColor(0);
    }

    // Read a rendered DOM <table> into autoTable head/body, preserving colspans,
    // right-alignment, and per-row style tags. `rowClassStyles` maps a <tr> CSS
    // class to a ROW_STYLES key. Rows tagged `.no-pdf` are skipped.
    function extractTable(tableEl, rowClassStyles) {
        rowClassStyles = rowClassStyles || {};
        var head = [], body = [], meta = [];

        tableEl.querySelectorAll("thead tr").forEach(function (tr) {
            var row = [];
            tr.querySelectorAll("th").forEach(function (th) {
                var cell = { content: ascii(th.textContent.trim()) };
                var cs = parseInt(th.getAttribute("colspan") || "1", 10);
                var rs = parseInt(th.getAttribute("rowspan") || "1", 10);
                if (cs > 1) cell.colSpan = cs;
                if (rs > 1) cell.rowSpan = rs;
                row.push(cell);
            });
            head.push(row);
        });

        tableEl.querySelectorAll("tbody tr").forEach(function (tr) {
            if (tr.classList.contains("no-pdf")) return;
            var styleName = null;
            Object.keys(rowClassStyles).some(function (cls) {
                if (tr.classList.contains(cls)) { styleName = rowClassStyles[cls]; return true; }
                return false;
            });
            var row = [];
            tr.querySelectorAll("td, th").forEach(function (td) {
                var cell = { content: ascii(td.textContent.trim()) };
                var cs = parseInt(td.getAttribute("colspan") || "1", 10);
                if (cs > 1) cell.colSpan = cs;
                if ((td.className || "").indexOf("text-right") >= 0) cell.styles = { halign: "right" };
                else if ((td.className || "").indexOf("text-center") >= 0) cell.styles = { halign: "center" };
                row.push(cell);
            });
            body.push(row);
            meta.push(styleName);
        });

        return { head: head, body: body, meta: meta };
    }

    /**
     * Export a bordered, letterhead PDF from a DOM table (or explicit head/body).
     * opts:
     *   tableEl        DOM <table> to read (preferred), OR
     *   head, body     explicit autoTable arrays
     *   title          report title (rendered UPPERCASE, bold, underlined)
     *   subtitle       optional line under the title (period / region / …)
     *   filename       download name (".pdf" appended if missing)
     *   orientation    "landscape" (default) | "portrait"
     *   division       shown in the footer
     *   rowClassStyles { "<tr class>": "<ROW_STYLES key>" }
     *   columnStyles   passed straight to autoTable
     *   fontSize       body font size (default 7.5)
     */
    function exportTable(opts) {
        opts = opts || {};
        var jsPDF = window.jspdf && window.jspdf.jsPDF;
        if (!jsPDF) { alert("PDF library not loaded. Please refresh and try again."); return; }

        var orientation = opts.orientation || "landscape";
        var doc = new jsPDF({ orientation: orientation, unit: "mm", format: "a4", compress: true });

        var head, body, meta;
        if (opts.tableEl) {
            var ex = extractTable(opts.tableEl, opts.rowClassStyles);
            head = ex.head; body = ex.body; meta = ex.meta;
        } else {
            head = opts.head || []; body = opts.body || []; meta = [];
        }

        var title = opts.title || "Report";
        var subtitle = opts.subtitle || "";
        var pageH = doc.internal.pageSize.getHeight();

        doc.autoTable({
            head: head,
            body: body,
            startY: HEADER_RESERVE,
            theme: "grid",
            styles: {
                fontSize: opts.fontSize || 7.5, cellPadding: 1.6,
                lineColor: [150, 150, 150], lineWidth: 0.1,
                textColor: [0, 0, 0], valign: "middle", overflow: "linebreak"
            },
            headStyles: {
                fillColor: [30, 41, 59], textColor: [255, 255, 255], fontStyle: "bold",
                halign: "center", lineColor: [150, 150, 150], lineWidth: 0.1
            },
            columnStyles: opts.columnStyles || {},
            margin: { left: 10, right: 10, top: HEADER_RESERVE, bottom: 16 },
            didParseCell: function (data) {
                if (data.section === "body") {
                    var m = meta[data.row.index];
                    if (m && ROW_STYLES[m]) {
                        var s = ROW_STYLES[m];
                        for (var k in s) { if (s.hasOwnProperty(k)) data.cell.styles[k] = s[k]; }
                    }
                }
            },
            didDrawPage: function () {
                drawHeader(doc, title, subtitle);
            }
        });

        // Closing note on the last page (new page if there's no room).
        var cy = (doc.lastAutoTable ? doc.lastAutoTable.finalY : HEADER_RESERVE) + 10;
        if (cy > pageH - 18) {
            doc.addPage();
            drawHeader(doc, title, subtitle);
            cy = HEADER_RESERVE + 8;
        }
        drawClosing(doc, cy);

        drawFooters(doc, opts.division || "");

        var filename = opts.filename || (title.replace(/[^\w]+/g, "_"));
        if (!/\.pdf$/i.test(filename)) filename += ".pdf";
        doc.save(filename);
    }

    window.ScanifyReportPDF = {
        exportTable: exportTable,
        ascii: ascii,
        ROW_STYLES: ROW_STYLES
    };
})();

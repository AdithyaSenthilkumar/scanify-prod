/*
 * Sales Target Entry — Compact Table + Edit Modal approach
 * Each HQ row stores month data in a JS array, shown as a compact summary row.
 * Editing happens in a modal with quarterly-grouped month inputs.
 */
let targetRows = []; // { id, hqName, hqId, regionId, months: {apr,...,mar} }
let rowIdCounter = 0;
let editingDocName = null;
let editingRowIdx = null;
let isApproved = false;

const MONTHS = ["apr","may","jun","jul","aug","sep","oct","nov","dec","jan","feb","mar"];

$(document).ready(function () {
    setDefaultDates();

    const urlParams = new URLSearchParams(window.location.search);
    const targetName = urlParams.get("name");
    if (targetName) {
        loadExistingTarget(targetName);
    }

    $("#add-row-btn").on("click", function () { openEditModal(-1); });
    $("#save-btn").on("click", saveTargets);
    $("#reset-btn").on("click", resetScreen);
    $("#process-bulk-import").on("click", processBulkImport);
    $("#approve-btn").on("click", approveTarget);
    $("#modal-save-row").on("click", saveModalRow);

    // Modal month input live recalc
    $(document).on("input", ".modal-month-input", recalcModalTotals);

    // HQ search in modal
    $(document).on("input focus click", "#modal-hq-search", debounce(handleModalHQSearch, 250));
    $(document).on("click", ".modal-hq-result", handleModalHQClick);
    $(document).on("click", function (e) {
        if (!$(e.target).closest(".hq-cell-wrap, .modal-link-dropdown").length) {
            $(".modal-link-dropdown").remove();
        }
    });

    renderTable();
});

function setDefaultDates() {
    const today = new Date();
    const year = today.getMonth() + 1 >= 4 ? today.getFullYear() : today.getFullYear() - 1;
    const nextShortYear = String((year + 1) % 100).padStart(2, "0");
    $("#financial-year").val(`${year}-${nextShortYear}`);
    $("#start-date").val(`${year}-04-01`);
    $("#end-date").val(`${year + 1}-03-31`);
    $("#target-status").val("Draft");
    setStatusBadge("Draft", 0);
}

// ─── Table Rendering ─────────────────────────────────────

function renderTable() {
    const tbody = document.getElementById("targets-tbody");
    const empty = document.getElementById("emptyTargetRows");
    if (!targetRows.length) {
        tbody.innerHTML = "";
        if (empty) empty.style.display = "";
        recalcSummary();
        return;
    }
    if (empty) empty.style.display = "none";

    let html = "";
    targetRows.forEach(function (row, idx) {
        const m = row.months;
        const q1 = toNum(m.apr) + toNum(m.may) + toNum(m.jun);
        const q2 = toNum(m.jul) + toNum(m.aug) + toNum(m.sep);
        const q3 = toNum(m.oct) + toNum(m.nov) + toNum(m.dec);
        const q4 = toNum(m.jan) + toNum(m.feb) + toNum(m.mar);
        const total = q1 + q2 + q3 + q4;

        html += '<tr>'
            + '<td class="text-center">' + (idx + 1) + '</td>'
            + '<td><span class="hq-name-display">' + escapeHtml(row.hqName || "—") + '</span></td>'
            + '<td class="text-right target-val">' + formatNum(q1) + '</td>'
            + '<td class="text-right target-val">' + formatNum(q2) + '</td>'
            + '<td class="text-right target-val">' + formatNum(q3) + '</td>'
            + '<td class="text-right target-val">' + formatNum(q4) + '</td>'
            + '<td class="text-right target-total">' + formatNum(total) + '</td>'
            + '<td class="text-center">'
            + (isApproved
                ? '<button class="btn-row-edit" onclick="openEditModal(' + idx + ')"><i class="fa fa-eye"></i> View</button>'
                : '<button class="btn-row-edit mr-1" onclick="openEditModal(' + idx + ')"><i class="fa fa-edit"></i> Edit</button>'
                  + '<button class="btn-row-del" onclick="removeTargetRow(' + idx + ')"><i class="fa fa-trash"></i></button>')
            + '</td>'
            + '</tr>';
    });
    tbody.innerHTML = html;
    recalcSummary();
}

function removeTargetRow(idx) {
    if (targetRows.length <= 1) {
        showAlert("At least one row is required", "warning");
        return;
    }
    targetRows.splice(idx, 1);
    renderTable();
}

function recalcSummary() {
    let grandTotal = 0;
    let hqCount = 0;
    targetRows.forEach(function (row) {
        const m = row.months;
        let total = 0;
        MONTHS.forEach(function (k) { total += toNum(m[k]); });
        grandTotal += total;
        if (row.hqId) hqCount++;
    });
    $("#total-hqs").text(hqCount);
    $("#grand-total").text(formatNum(grandTotal));
}

// ─── Edit Modal ──────────────────────────────────────────

function openEditModal(idx) {
    editingRowIdx = idx;
    const isNew = idx < 0;
    const row = isNew ? { hqName: "", hqId: "", regionId: "", months: {} } : targetRows[idx];

    $("#edit-modal-title").text(isApproved ? "View HQ Target" : (isNew ? "Add HQ Target" : "Edit HQ Target"));
    $("#modal-hq-search").val(row.hqName || "").prop("readonly", isApproved);
    $("#modal-hq-id").val(row.hqId || "");
    $("#modal-hq-region-id").val(row.regionId || "");

    MONTHS.forEach(function (m) {
        var $inp = $(".modal-month-input[data-month='" + m + "']");
        $inp.val(toNum(row.months[m] || 0));
        if (isApproved) $inp.prop("readonly", true).css("background", "#f8f9fa");
        else $inp.prop("readonly", false).css("background", "");
    });
    recalcModalTotals();

    if (isApproved) {
        $("#modal-save-row").hide();
    } else {
        $("#modal-save-row").show().text(isNew ? " Add Row" : " Save Row")
            .html('<i class="fa fa-check"></i> ' + (isNew ? "Add Row" : "Save Row"));
    }

    $("#edit-hq-modal").modal("show");
}

function recalcModalTotals() {
    var vals = {};
    $(".modal-month-input").each(function () {
        vals[$(this).data("month")] = toNum($(this).val());
    });
    var q1 = (vals.apr || 0) + (vals.may || 0) + (vals.jun || 0);
    var q2 = (vals.jul || 0) + (vals.aug || 0) + (vals.sep || 0);
    var q3 = (vals.oct || 0) + (vals.nov || 0) + (vals.dec || 0);
    var q4 = (vals.jan || 0) + (vals.feb || 0) + (vals.mar || 0);
    $("#modal-q1").text(formatNum(q1));
    $("#modal-q2").text(formatNum(q2));
    $("#modal-q3").text(formatNum(q3));
    $("#modal-q4").text(formatNum(q4));
    $("#modal-yearly-total").text(formatNum(q1 + q2 + q3 + q4));
}

function saveModalRow() {
    var hqName = ($("#modal-hq-search").val() || "").trim();
    var hqId = ($("#modal-hq-id").val() || "").trim();
    var regionId = ($("#modal-hq-region-id").val() || "").trim();

    if (!hqId) {
        showAlert("Please search and select an HQ", "warning");
        return;
    }

    // Check for duplicate HQ
    for (var i = 0; i < targetRows.length; i++) {
        if (i !== editingRowIdx && targetRows[i].hqId === hqId) {
            showAlert("Duplicate HQ: " + hqName, "warning");
            return;
        }
    }

    var months = {};
    $(".modal-month-input").each(function () {
        months[$(this).data("month")] = toNum($(this).val());
    });

    if (editingRowIdx < 0) {
        // Adding new
        rowIdCounter++;
        targetRows.push({ id: rowIdCounter, hqName: hqName, hqId: hqId, regionId: regionId, months: months });
    } else {
        // Updating existing
        targetRows[editingRowIdx].hqName = hqName;
        targetRows[editingRowIdx].hqId = hqId;
        targetRows[editingRowIdx].regionId = regionId;
        targetRows[editingRowIdx].months = months;
    }

    $("#edit-hq-modal").modal("hide");
    renderTable();
}

// ─── HQ Search in Modal ─────────────────────────────────

async function handleModalHQSearch(e) {
    var search = ($("#modal-hq-search").val() || "").trim();
    var division = $("#division").val();

    if (e.type === "input") {
        $("#modal-hq-id").val("");
        $("#modal-hq-region-id").val("");
    }
    $(".modal-link-dropdown").remove();

    try {
        var rows = await callApi("scanify.api.search_hq_targets", { search: search, division: division });
        if (!rows.length) return;

        var $menu = $('<div class="modal-link-dropdown"></div>');
        rows.forEach(function (row) {
            var meta = [row.team_name, row.region_name, row.zone].filter(Boolean).join(" | ");
            $menu.append(
                '<div class="modal-hq-result" data-hq=\'' + escapeHtml(JSON.stringify(row)) + '\'>'
                + '<div><strong>' + escapeHtml(row.hq_name || row.name) + '</strong></div>'
                + '<small class="text-muted">' + escapeHtml(meta) + '</small>'
                + '</div>'
            );
        });

        var $input = $("#modal-hq-search");
        var offset = $input.position();
        $menu.css({
            position: "absolute",
            top: (offset.top + $input.outerHeight() + 2) + "px",
            left: offset.left + "px",
            width: $input.outerWidth() + "px",
            zIndex: 99999,
            background: "#fff",
            border: "1px solid #ced4da",
            borderRadius: "4px",
            boxShadow: "0 4px 12px rgba(0,0,0,0.1)",
            maxHeight: "250px",
            overflowY: "auto"
        });
        $input.closest(".hq-cell-wrap").append($menu);
    } catch (e) {
        showAlert("Unable to search HQ", "danger");
    }
}

function handleModalHQClick() {
    var row = JSON.parse(unescapeHtml($(this).attr("data-hq")));
    $("#modal-hq-search").val(row.hq_name || row.name);
    $("#modal-hq-id").val(row.name || "");
    $("#modal-hq-region-id").val(row.region || "");
    $(".modal-link-dropdown").remove();
}

// ─── Save / Load / Approve ───────────────────────────────

async function saveTargets() {
    var financial_year = ($("#financial-year").val() || "").trim();
    var start_date = ($("#start-date").val() || "").trim();
    var end_date = ($("#end-date").val() || "").trim();

    if (!financial_year || !start_date || !end_date) {
        showAlert("Financial year, start date and end date are mandatory", "warning");
        return;
    }

    var rows = [];
    var seenHQs = {};
    var seenRegions = {};
    var hasError = false;

    targetRows.forEach(function (row) {
        if (!row.hqId) return;
        if (seenHQs[row.hqId]) {
            showAlert("Duplicate HQ: " + row.hqName, "warning");
            hasError = true;
            return;
        }
        seenHQs[row.hqId] = true;
        if (row.regionId) seenRegions[row.regionId] = true;

        var r = { hq: row.hqId };
        MONTHS.forEach(function (m) { r[m] = toNum(row.months[m]); });
        rows.push(r);
    });

    if (hasError) return;
    if (!rows.length) { showAlert("Add at least one HQ row", "warning"); return; }

    var regionKeys = Object.keys(seenRegions);
    if (regionKeys.length > 1) {
        showAlert("All selected HQs must belong to the same region", "warning");
        return;
    }

    showLoadingOverlay("Saving sales target...");
    try {
        var result;
        if (editingDocName) {
            result = await callApi("scanify.api.update_hq_yearly_target_from_portal", {
                name: editingDocName, financial_year: financial_year,
                start_date: start_date, end_date: end_date, status: "Draft", hq_targets: rows
            });
        } else {
            result = await callApi("scanify.api.create_hq_yearly_target_from_portal", {
                financial_year: financial_year, start_date: start_date,
                end_date: end_date, status: "Draft", hq_targets: rows
            });
        }
        hideLoadingOverlay();
        if (result.success) {
            showAlert("Saved as Draft: " + (result.name || editingDocName), "success");
            setTimeout(function () { window.location.href = "/portal/sales-targets-list"; }, 900);
        } else {
            showAlert(result.message || "Failed to save", "danger");
        }
    } catch (e) {
        hideLoadingOverlay();
        showAlert(e.message || "Failed to save", "danger");
    }
}

function resetScreen() {
    targetRows = [];
    editingDocName = null;
    isApproved = false;
    rowIdCounter = 0;
    setDefaultDates();
    renderTable();
}

async function loadExistingTarget(name) {
    showLoadingOverlay("Loading sales target...");
    try {
        var result = await callApi("scanify.api.get_hq_yearly_target_details", { name: name });
        hideLoadingOverlay();
        if (result.success && result.doc) {
            editingDocName = result.doc.name;
            $("#financial-year").val(result.doc.financial_year);
            $("#start-date").val(result.doc.start_date);
            $("#end-date").val(result.doc.end_date);
            $("#target-status").val(result.doc.status || "Draft");

            isApproved = result.doc.docstatus === 1;
            setStatusBadge(isApproved ? "Approved" : (result.doc.status || "Draft"), result.doc.docstatus);

            if (isApproved) {
                $("#save-btn, #add-row-btn, #bulk-import-btn, #approve-btn, #reset-btn").addClass("d-none");
                $("#financial-year, #start-date, #end-date").prop("readonly", true).css("background", "#f8f9fa");
            } else {
                $("#save-btn").removeClass("d-none");
                $("#approve-btn").removeClass("d-none");
            }

            targetRows = [];
            if (result.doc.items && result.doc.items.length) {
                result.doc.items.forEach(function (item) {
                    rowIdCounter++;
                    var months = {};
                    MONTHS.forEach(function (m) { months[m] = item[m] || 0; });
                    targetRows.push({
                        id: rowIdCounter,
                        hqName: item.hq_name,
                        hqId: item.hq,
                        regionId: item.region || "",
                        months: months
                    });
                });
            }
            renderTable();
        } else {
            showAlert("Failed to load target details", "danger");
        }
    } catch (e) {
        hideLoadingOverlay();
        showAlert(e.message || "Failed to load target details", "danger");
    }
}

async function approveTarget() {
    if (!editingDocName) return;
    if (!confirm("Are you sure you want to approve this Sales Target? It cannot be edited afterwards.")) return;

    showLoadingOverlay("Approving...");
    try {
        var result = await callApi("scanify.api.submit_hq_yearly_target_from_portal", { name: editingDocName });
        hideLoadingOverlay();
        if (result.success) {
            showAlert("Sales Target Approved successfully!", "success");
            isApproved = true;
            setStatusBadge("Approved", 1);
            $("#save-btn, #approve-btn, #add-row-btn, #bulk-import-btn, #reset-btn").addClass("d-none");
            $("#financial-year, #start-date, #end-date").prop("readonly", true).css("background", "#f8f9fa");
            renderTable();
        } else {
            showAlert(result.message || "Approval failed", "danger");
        }
    } catch (e) {
        hideLoadingOverlay();
        showAlert(e.message || "Approval failed", "danger");
    }
}

function setStatusBadge(status, docstatus) {
    var $badge = $("#target-status-badge");
    $badge.removeClass("badge-warning badge-success badge-secondary badge-danger");
    if (docstatus === 1 || status === "Approved") {
        $badge.addClass("badge-success").text("Approved");
    } else if (status === "Draft" || !status) {
        $badge.addClass("badge-warning").text("Draft");
    } else {
        $badge.addClass("badge-secondary").text(status);
    }
}

// ─── Bulk Import ─────────────────────────────────────────

function downloadTargetTemplate() {
    var headers = ["HQ Name","Apr","May","Jun","Jul","Aug","Sep","Oct","Nov","Dec","Jan","Feb","Mar"];
    var sample = ["Mumbai HQ","10.00","10.00","10.00","10.00","10.00","10.00","10.00","10.00","10.00","10.00","10.00","10.00"];
    var csvContent = headers.join(",") + "\n" + sample.join(",");
    var blob = new Blob([csvContent], { type: "text/csv;charset=utf-8;" });
    var url = URL.createObjectURL(blob);
    var link = document.createElement("a");
    link.setAttribute("href", url);
    link.setAttribute("download", "HQ_Target_Template.csv");
    document.body.appendChild(link);
    link.click();
    document.body.removeChild(link);
    URL.revokeObjectURL(url);
    showAlert("Template downloaded. Fill HQ Names exactly as in HQ Master (figures in Lakhs).", "info");
}

async function processBulkImport() {
    var fileInput = document.getElementById("target-import-file");
    var file = fileInput && fileInput.files[0];
    if (!file) { showAlert("Please select an Excel or CSV file to import", "warning"); return; }

    $("#import-target-progress").show();
    $("#import-target-results").hide();

    var formData = new FormData();
    formData.append("file", file);
    formData.append("division", $("#division").val() || "");

    try {
        var response = await fetch("/api/method/scanify.api.resolve_hq_target_rows_from_file", {
            method: "POST",
            headers: { "X-Frappe-CSRF-Token": window.csrf_token },
            body: formData,
        });
        var data = await response.json();
        var result = data.message;
        $("#import-target-progress").hide();

        if (!result || !result.success) {
            showAlert(result?.message || "Failed to process file", "danger");
            return;
        }

        var rows = result.rows || [];
        var errors = result.errors || [];
        if (rows.length === 0) {
            var errMsg = "No valid HQ rows found in the file.";
            if (errors.length) errMsg += " " + errors.slice(0, 3).join("; ");
            showAlert(errMsg, "warning");
            return;
        }

        targetRows = [];
        rows.forEach(function (item) {
            rowIdCounter++;
            var months = {};
            MONTHS.forEach(function (m) { months[m] = item[m] || 0; });
            targetRows.push({ id: rowIdCounter, hqName: item.hq_name, hqId: item.hq, regionId: "", months: months });
        });

        renderTable();
        $("#bulk-import-modal").modal("hide");
        fileInput.value = "";

        var msg = "Imported " + rows.length + " HQ row(s) successfully.";
        if (errors.length) msg += " " + errors.length + " row(s) skipped: " + errors.slice(0, 2).join("; ");
        showAlert(msg, "success");
    } catch (e) {
        $("#import-target-progress").hide();
        showAlert("Failed to process the import file: " + (e.message || ""), "danger");
    }
}

// ─── Utilities ───────────────────────────────────────────

function toNum(val) { var n = parseFloat(val); return Number.isNaN(n) ? 0 : n; }
function formatNum(num) { return toNum(num).toFixed(2); }

function debounce(func, wait) {
    var timeout;
    return function () {
        var ctx = this, args = arguments;
        clearTimeout(timeout);
        timeout = setTimeout(function () { func.apply(ctx, args); }, wait);
    };
}

function callApi(method, args) {
    return new Promise(function (resolve, reject) {
        $.ajax({
            url: "/api/method/" + method,
            type: "POST",
            contentType: "application/json",
            headers: { "X-Frappe-CSRF-Token": window.csrf_token || (window.frappe && frappe.csrf_token) || "" },
            data: JSON.stringify(args || {}),
            success: function (r) { resolve(r.message || {}); },
            error: function (xhr) {
                reject(new Error(xhr.responseJSON?._server_messages || xhr.responseJSON?.message || "API error"));
            }
        });
    });
}

function escapeHtml(s) {
    return String(s || "").replace(/[&<>"']/g, function (m) {
        return { "&": "&amp;", "<": "&lt;", ">": "&gt;", '"': "&quot;", "'": "&#039;" }[m];
    });
}
function unescapeHtml(s) {
    var txt = document.createElement("textarea");
    txt.innerHTML = s || "";
    return txt.value;
}

function showLoadingOverlay(message) {
    $("body").append(
        '<div id="loading-overlay" style="position:fixed;top:0;left:0;width:100%;height:100%;background:rgba(0,0,0,.6);z-index:99999;display:flex;align-items:center;justify-content:center;">'
        + '<div style="background:#fff;padding:18px 28px;border-radius:8px;text-align:center;">'
        + '<div class="spinner-border text-primary mb-2" role="status"></div>'
        + '<div>' + escapeHtml(message) + '</div></div></div>'
    );
}
function hideLoadingOverlay() { $("#loading-overlay").remove(); }

function showAlert(message, type) {
    var $alert = $(
        '<div class="alert alert-' + type + ' alert-dismissible fade show" role="alert"'
        + ' style="position:fixed;top:70px;right:20px;z-index:9999;min-width:320px;box-shadow:0 4px 6px rgba(0,0,0,.1);">'
        + '<strong>' + escapeHtml(message) + '</strong>'
        + '<button type="button" class="close" data-dismiss="alert">&times;</button></div>'
    );
    $("body").append($alert);
    setTimeout(function () { $alert.fadeOut(function () { $(this).remove(); }); }, 3500);
}

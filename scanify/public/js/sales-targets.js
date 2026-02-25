let rowCounter = 0;

$(document).ready(function () {
    setDefaultDates();
    addRow();

    $("#add-row-btn").on("click", addRow);
    $("#save-btn").on("click", saveTargets);
    $("#reset-btn").on("click", resetScreen);

    $(document).on("input focus click", ".hq-search", debounce(handleHQSearchInput, 250));
    $(document).on("click", ".hq-result", handleHQResultClick);
    $(document).on("input", ".month-input", handleMonthInput);
    $(document).on("click", ".remove-row-btn", handleRemoveRow);

    $(document).on("click", function (e) {
        if (!$(e.target).closest(".hq-cell-wrap, .link-dropdown").length) {
            $(".link-dropdown").remove();
        }
    });
});

function setDefaultDates() {
    const today = new Date();
    const year = today.getMonth() + 1 >= 4 ? today.getFullYear() : today.getFullYear() - 1;
    const nextShortYear = String((year + 1) % 100).padStart(2, "0");

    $("#financial-year").val(`${year}-${nextShortYear}`);
    $("#start-date").val(`${year}-04-01`);
    $("#end-date").val(`${year + 1}-03-31`);
    $("#target-status").val("Draft");
}

function addRow() {
    rowCounter += 1;
    const rowId = `target-row-${rowCounter}`;
    const months = ["apr", "may", "jun", "jul", "aug", "sep", "oct", "nov", "dec", "jan", "feb", "mar"];

    let monthCells = "";
    months.forEach((m, idx) => {
        monthCells += `<td><input type="number" min="0" step="0.01" class="form-control form-control-sm month-input" data-month="${m}" value="0"></td>`;
        if ([2, 5, 8, 11].includes(idx)) {
            const qNum = (idx + 1) / 3;
            monthCells += `<td><input type="text" class="form-control form-control-sm calc-input q${qNum}-total" value="0.00" readonly></td>`;
        }
    });

    const rowHtml = `
        <tr id="${rowId}">
            <td class="col-hq">
                <div class="hq-cell-wrap">
                    <input type="text" class="form-control form-control-sm hq-search" placeholder="Search HQ..." autocomplete="off">
                    <input type="hidden" class="hq-id">
                    <input type="hidden" class="hq-region-id">
                </div>
            </td>
            ${monthCells}
            <td><input type="text" class="form-control form-control-sm calc-input yearly-total" value="0.00" readonly></td>
            <td class="text-center">
                <button class="btn btn-sm btn-danger remove-row-btn" type="button" title="Remove row">
                    <i class="fa fa-trash"></i>
                </button>
            </td>
        </tr>
    `;

    $("#targets-tbody").append(rowHtml);
    recalcAll();
}

async function handleHQSearchInput(e) {
    const $input = $(this);
    const search = ($input.val() || "").trim();
    const $wrap = $input.closest(".hq-cell-wrap");
    const division = $("#division").val();

    if (e.type === 'input') {
        $wrap.find(".hq-id").val("");
        $wrap.find(".hq-region-id").val("");
        clearHierarchyFields($wrap.closest("tr"));
    }
    syncDetectedRegion();

    $(".link-dropdown").remove();

    try {
        const rows = await callApi("scanify.api.search_hq_targets", {
            search,
            division
        });

        if (!rows.length) {
            return;
        }

        const rowId = $wrap.closest("tr").attr("id") || "";
        const $menu = $('<div class="link-dropdown"></div>').attr("data-target-row", rowId);
        rows.forEach((row) => {
            const meta = [row.team_name, row.region_name, row.zone].filter(Boolean).join(" | ");
            const item = $(`
                <div class="hq-result" data-hq='${escapeHtml(JSON.stringify(row))}'>
                    <div><strong>${escapeHtml(row.hq_name || row.name)}</strong></div>
                    <small class="text-muted">${escapeHtml(meta)}</small>
                </div>
            `);
            $menu.append(item);
        });

        const offset = $input.offset();
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

        $("body").append($menu);
    } catch (e) {
        showAlert("Unable to search HQ", "danger");
    }
}

function handleHQResultClick() {
    const row = JSON.parse(unescapeHtml($(this).attr("data-hq")));
    const targetRowId = $(this).closest(".link-dropdown").attr("data-target-row") || "";
    const $row = targetRowId ? $("#" + targetRowId) : $();
    const $wrap = $row.find(".hq-cell-wrap").first();
    const $tr = $wrap.closest("tr");

    if (!$wrap.length) {
        $(".link-dropdown").remove();
        return;
    }

    $wrap.find(".hq-search").val(row.hq_name || row.name);
    $wrap.find(".hq-id").val(row.name || "");
    $wrap.find(".hq-region-id").val(row.region || "");

    $(".link-dropdown").remove();
    syncDetectedRegion();
}

function handleMonthInput() {
    recalcRow($(this).closest("tr"));
    recalcAll();
}

function handleRemoveRow() {
    if ($("#targets-tbody tr").length === 1) {
        showAlert("At least one row is required", "warning");
        return;
    }
    $(this).closest("tr").remove();
    recalcAll();
    syncDetectedRegion();
}

function recalcRow($tr) {
    const q1 = sumMonths($tr, ["apr", "may", "jun"]);
    const q2 = sumMonths($tr, ["jul", "aug", "sep"]);
    const q3 = sumMonths($tr, ["oct", "nov", "dec"]);
    const q4 = sumMonths($tr, ["jan", "feb", "mar"]);
    const total = q1 + q2 + q3 + q4;

    $tr.find(".q1-total").val(formatNum(q1));
    $tr.find(".q2-total").val(formatNum(q2));
    $tr.find(".q3-total").val(formatNum(q3));
    $tr.find(".q4-total").val(formatNum(q4));
    $tr.find(".yearly-total").val(formatNum(total));
}

function recalcAll() {
    let grandTotal = 0;
    let hqCount = 0;

    $("#targets-tbody tr").each(function () {
        recalcRow($(this));
        grandTotal += toNum($(this).find(".yearly-total").val());
        if ($(this).find(".hq-id").val()) {
            hqCount += 1;
        }
    });

    $("#total-hqs").text(hqCount);
    $("#grand-total").text(formatNum(grandTotal));
}

function sumMonths($tr, monthKeys) {
    return monthKeys.reduce((acc, key) => acc + toNum($tr.find(`.month-input[data-month="${key}"]`).val()), 0);
}

function clearHierarchyFields($tr) {
    $tr.find(".team-name").val("");
    $tr.find(".region-name").val("");
    $tr.find(".zone-name").val("");
}

function syncDetectedRegion() {
    const regions = new Set();
    let display = "";

    $("#targets-tbody tr").each(function () {
        const regionId = ($(this).find(".hq-region-id").val() || "").trim();
        const regionName = ($(this).find(".region-name").val() || "").trim();
        if (regionId) {
            regions.add(regionId);
            display = regionName || display;
        }
    });

    if (regions.size === 0) {
        $("#detected-region").val("");
        return;
    }
    if (regions.size > 1) {
        $("#detected-region").val("Multiple regions selected");
        return;
    }
    $("#detected-region").val(display);
}

async function saveTargets() {
    const financial_year = ($("#financial-year").val() || "").trim();
    const start_date = ($("#start-date").val() || "").trim();
    const end_date = ($("#end-date").val() || "").trim();
    const status = $("#target-status").val();

    if (!financial_year || !start_date || !end_date) {
        showAlert("Financial year, start date and end date are mandatory", "warning");
        return;
    }

    const rows = [];
    const seenHQs = new Set();
    const seenRegions = new Set();
    let hasValidationError = false;

    $("#targets-tbody tr").each(function () {
        const $tr = $(this);
        const hq = ($tr.find(".hq-id").val() || "").trim();
        const region = ($tr.find(".hq-region-id").val() || "").trim();

        if (!hq) {
            return;
        }

        if (seenHQs.has(hq)) {
            showAlert(`Duplicate HQ selected: ${$tr.find(".hq-search").val() || hq}`, "warning");
            hasValidationError = true;
            return false;
        }
        seenHQs.add(hq);

        if (region) {
            seenRegions.add(region);
        }

        rows.push({
            hq,
            apr: toNum($tr.find('.month-input[data-month="apr"]').val()),
            may: toNum($tr.find('.month-input[data-month="may"]').val()),
            jun: toNum($tr.find('.month-input[data-month="jun"]').val()),
            jul: toNum($tr.find('.month-input[data-month="jul"]').val()),
            aug: toNum($tr.find('.month-input[data-month="aug"]').val()),
            sep: toNum($tr.find('.month-input[data-month="sep"]').val()),
            oct: toNum($tr.find('.month-input[data-month="oct"]').val()),
            nov: toNum($tr.find('.month-input[data-month="nov"]').val()),
            dec: toNum($tr.find('.month-input[data-month="dec"]').val()),
            jan: toNum($tr.find('.month-input[data-month="jan"]').val()),
            feb: toNum($tr.find('.month-input[data-month="feb"]').val()),
            mar: toNum($tr.find('.month-input[data-month="mar"]').val())
        });
    });

    if (hasValidationError) {
        return;
    }

    if (!rows.length) {
        showAlert("Add at least one HQ row", "warning");
        return;
    }
    if (seenRegions.size > 1) {
        showAlert("All selected HQs must belong to the same region", "warning");
        return;
    }

    showLoadingOverlay("Saving sales target...");
    try {
        const result = await callApi("scanify.api.create_hq_yearly_target_from_portal", {
            financial_year,
            start_date,
            end_date,
            status,
            hq_targets: rows
        });

        hideLoadingOverlay();
        if (result.success) {
            showAlert(`Saved successfully: ${result.name}`, "success");
            setTimeout(() => {
                window.location.href = `/app/hq-yearly-target/${result.name}`;
            }, 900);
        } else {
            showAlert(result.message || "Failed to save", "danger");
        }
    } catch (e) {
        hideLoadingOverlay();
        showAlert(e.message || "Failed to save", "danger");
    }
}

function resetScreen() {
    $("#targets-tbody").empty();
    addRow();
    setDefaultDates();
    $("#detected-region").val("");
    recalcAll();
}

function toNum(val) {
    const num = parseFloat(val);
    return Number.isNaN(num) ? 0 : num;
}

function formatNum(num) {
    return toNum(num).toFixed(2);
}

function debounce(func, wait) {
    let timeout;
    return function (...args) {
        const ctx = this;
        clearTimeout(timeout);
        timeout = setTimeout(() => func.apply(ctx, args), wait);
    };
}

function callApi(method, args) {
    return new Promise((resolve, reject) => {
        $.ajax({
            url: `/api/method/${method}`,
            type: "POST",
            contentType: "application/json",
            headers: {
                "X-Frappe-CSRF-Token": window.csrf_token || (window.frappe && frappe.csrf_token) || ""
            },
            data: JSON.stringify(args || {}),
            success: function (r) {
                resolve(r.message || {});
            },
            error: function (xhr) {
                reject(new Error(xhr.responseJSON?._server_messages || xhr.responseJSON?.message || "API error"));
            }
        });
    });
}

function escapeHtml(s) {
    return String(s || "").replace(/[&<>"']/g, function (m) {
        return {
            "&": "&amp;",
            "<": "&lt;",
            ">": "&gt;",
            '"': "&quot;",
            "'": "&#039;"
        }[m];
    });
}

function unescapeHtml(s) {
    const txt = document.createElement("textarea");
    txt.innerHTML = s || "";
    return txt.value;
}

function showLoadingOverlay(message) {
    const overlay = `
        <div id="loading-overlay" style="position:fixed;top:0;left:0;width:100%;height:100%;background:rgba(0,0,0,.6);z-index:99999;display:flex;align-items:center;justify-content:center;">
            <div style="background:#fff;padding:18px 28px;border-radius:8px;text-align:center;">
                <div class="spinner-border text-primary mb-2" role="status"></div>
                <div>${escapeHtml(message)}</div>
            </div>
        </div>
    `;
    $("body").append(overlay);
}

function hideLoadingOverlay() {
    $("#loading-overlay").remove();
}

function showAlert(message, type) {
    const html = `
        <div class="alert alert-${type} alert-dismissible fade show" role="alert"
            style="position:fixed;top:70px;right:20px;z-index:9999;min-width:320px;box-shadow:0 4px 6px rgba(0,0,0,.1);">
            <strong>${escapeHtml(message)}</strong>
            <button type="button" class="close" data-dismiss="alert">&times;</button>
        </div>
    `;
    const $alert = $(html);
    $("body").append($alert);
    setTimeout(() => {
        $alert.fadeOut(function () {
            $(this).remove();
        });
    }, 3500);
}

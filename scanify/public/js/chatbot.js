/**
 * Scanify Agentic Chatbot — Client-side JS
 * Text2SQL pipeline with Chart.js visualization support
 */
window.ScanifyChatbot = (function () {
    'use strict';

    let config = {};
    let conversationHistory = [];
    let chartInstances = {};
    let chartCounter = 0;
    let chartDataStore = {};
    let isProcessing = false;

    function init(opts) {
        config = opts || {};
        bindEvents();
        autoResizeInput();
    }

    // ── Event Binding ──────────────────────────────────
    function bindEvents() {
        const input = document.getElementById('chatInput');
        const sendBtn = document.getElementById('btnSend');
        const clearBtn = document.getElementById('btnClearChat');
        const exportBtn = document.getElementById('btnExportChat');

        if (sendBtn) sendBtn.addEventListener('click', sendMessage);
        if (clearBtn) clearBtn.addEventListener('click', clearChat);
        if (exportBtn) exportBtn.addEventListener('click', exportChat);

        if (input) {
            input.addEventListener('keydown', function (e) {
                if (e.key === 'Enter' && !e.shiftKey) {
                    e.preventDefault();
                    sendMessage();
                }
            });
            input.addEventListener('input', function () {
                autoResizeInput();
                updateCharCount();
            });
        }
    }

    function autoResizeInput() {
        const input = document.getElementById('chatInput');
        if (!input) return;
        input.style.height = 'auto';
        input.style.height = Math.min(input.scrollHeight, 120) + 'px';
    }

    function updateCharCount() {
        const input = document.getElementById('chatInput');
        const counter = document.getElementById('charCount');
        if (input && counter) {
            counter.textContent = input.value.length + ' / 2000';
        }
    }

    // ── Send Message ───────────────────────────────────
    function sendMessage() {
        const input = document.getElementById('chatInput');
        if (!input) return;

        const message = input.value.trim();
        if (!message || isProcessing) return;

        isProcessing = true;
        input.value = '';
        autoResizeInput();
        updateCharCount();

        // Add user message to UI
        appendMessage('user', message);

        // Add to conversation history
        conversationHistory.push({ role: 'user', content: message });

        // Show typing indicator
        const typingId = showTypingIndicator();

        // Call API
        $.ajax({
            url: '/api/method/scanify.api.chatbot_query',
            type: 'POST',
            contentType: 'application/json',
            headers: { 'X-Frappe-CSRF-Token': config.csrfToken || frappe.csrf_token },
            data: JSON.stringify({
                message: message,
                conversation_history: JSON.stringify(conversationHistory)
            }),
            success: function (r) {
                removeTypingIndicator(typingId);
                const data = r.message || r;
                handleResponse(data);
                isProcessing = false;
            },
            error: function (xhr) {
                removeTypingIndicator(typingId);
                let errMsg = 'Something went wrong. Please try again.';
                try {
                    const errData = JSON.parse(xhr.responseText);
                    if (errData._server_messages) {
                        const msgs = JSON.parse(errData._server_messages);
                        errMsg = typeof msgs[0] === 'string' ? JSON.parse(msgs[0]).message : msgs[0].message;
                    }
                } catch (e) { /* use default */ }
                appendMessage('assistant', renderErrorBubble(errMsg));
                isProcessing = false;
            }
        });
    }

    // ── Response Handler ───────────────────────────────
    function handleResponse(data) {
        if (!data.success && data.error) {
            appendMessage('assistant', renderErrorBubble(data.error));
            conversationHistory.push({ role: 'assistant', content: data.error });
            return;
        }

        const type = data.type || 'text';
        let html = '';
        let historyText = '';

        switch (type) {
            case 'text':
                html = renderTextResponse(data);
                historyText = data.message || '';
                break;
            case 'error':
                html = renderErrorBubble(data.message || 'Unknown error');
                historyText = data.message || '';
                break;
            case 'metric':
                html = renderMetricResponse(data);
                historyText = data.title + ': ' + JSON.stringify(data.data);
                break;
            case 'data':
                html = renderTableResponse(data);
                historyText = data.title + ' (' + data.total_rows + ' rows)';
                break;
            case 'chart':
                html = renderChartResponse(data);
                historyText = data.title + ' (chart with ' + data.total_rows + ' data points)';
                break;
            default:
                html = renderTextResponse(data);
                historyText = data.message || JSON.stringify(data);
        }

        appendMessage('assistant', html);
        conversationHistory.push({ role: 'assistant', content: historyText });
    }

    // ── Message Rendering ──────────────────────────────
    function appendMessage(role, contentHtml) {
        const container = document.getElementById('chatMessages');
        if (!container) return;

        const wrapper = document.createElement('div');
        wrapper.className = 'chat-message ' + role + '-message';

        const avatarHtml = role === 'user'
            ? '<div class="message-avatar user-avatar"><i class="fa fa-user"></i></div>'
            : '<div class="message-avatar assistant-avatar"><i class="fa fa-robot"></i></div>';

        const time = new Date().toLocaleTimeString([], { hour: '2-digit', minute: '2-digit' });
        const bubbleClass = role === 'user' ? 'user-bubble' : 'assistant-bubble';

        wrapper.innerHTML =
            avatarHtml +
            '<div class="message-content">' +
            '<div class="message-bubble ' + bubbleClass + '">' + contentHtml + '</div>' +
            '<div class="message-time">' + time + '</div>' +
            '</div>';

        container.appendChild(wrapper);
        scrollToBottom();

        // Initialize charts after DOM layout settles
        setTimeout(function () {
            // Ensure chart/table messages have full width
            var chatMsg = wrapper.closest('.chat-message') || wrapper;
            if (wrapper.querySelector('.chart-response') || wrapper.querySelector('.table-response') || wrapper.querySelector('.metric-response')) {
                chatMsg.style.width = '90%';
                chatMsg.style.maxWidth = '90%';
            }
            initPendingCharts(wrapper);
        }, 200);
    }

    function scrollToBottom() {
        const container = document.getElementById('chatContainer');
        if (container) {
            container.scrollTop = container.scrollHeight;
        }
    }

    // ── Response Renderers ─────────────────────────────

    function renderTextResponse(data) {
        const text = data.message || '';
        let html = '';
        if (typeof marked !== 'undefined' && typeof marked.parse === 'function') {
            html = marked.parse(text);
        } else {
            html = '<p>' + escapeHtml(text).replace(/\n/g, '<br>') + '</p>';
        }
        return '<div class="text-response">' + html + '</div>';
    }

    function renderErrorBubble(msg) {
        return '<div class="error-response">' +
            '<i class="fa fa-exclamation-triangle"></i> ' +
            '<span>' + escapeHtml(msg) + '</span>' +
            '</div>';
    }

    function renderMetricResponse(data) {
        const rows = data.data || [];
        const row = rows[0] || {};
        const keys = Object.keys(row);

        let metricHtml = '<div class="metric-response">';
        metricHtml += '<div class="metric-header">' + escapeHtml(data.title || 'Result') + '</div>';
        if (data.description) {
            metricHtml += '<div class="metric-desc">' + escapeHtml(data.description) + '</div>';
        }

        if (!rows.length || !keys.length) {
            metricHtml += '<div class="empty-state"><i class="fa fa-calculator"></i><span>No result returned for this query</span></div>';
            metricHtml += renderSqlLink(data.sql);
            metricHtml += '</div>';
            return metricHtml;
        }

        metricHtml += '<div class="metric-values">';
        keys.forEach(function (key) {
            const val = row[key];
            const displayVal = typeof val === 'number' ? formatNumber(val) : val;
            metricHtml += '<div class="metric-item">' +
                '<div class="metric-value">' + displayVal + '</div>' +
                '<div class="metric-label">' + formatColumnName(key) + '</div>' +
                '</div>';
        });
        metricHtml += '</div>';
        metricHtml += renderSqlLink(data.sql);
        metricHtml += '</div>';
        return metricHtml;
    }

    function renderTableResponse(data) {
        const rows = data.data || [];

        // Empty state
        if (!rows.length) {
            let html = '<div class="table-response">';
            html += '<div class="table-header-bar">';
            html += '<div class="table-title"><i class="fa fa-table mr-1"></i> ' + escapeHtml(data.title || 'Results') + '</div>';
            html += '</div>';
            if (data.description) {
                html += '<div class="table-desc">' + escapeHtml(data.description) + '</div>';
            }
            html += '<div class="empty-state"><i class="fa fa-inbox"></i><span>No data found for this query</span></div>';
            html += renderSqlLink(data.sql);
            html += '</div>';
            return html;
        }

        // Always use actual data keys for lookups
        const dataKeys = Object.keys(rows[0]);
        // Use AI columns as display headers if count matches, else format data keys
        const aiColumns = data.columns || [];
        const displayHeaders = (aiColumns.length === dataKeys.length) ? aiColumns : dataKeys.map(formatColumnName);

        let html = '<div class="table-response">';
        html += '<div class="table-header-bar">';
        html += '<div class="table-title"><i class="fa fa-table mr-1"></i> ' + escapeHtml(data.title || 'Results') + '</div>';
        html += '<div class="table-meta">' + data.total_rows + ' rows</div>';
        html += '</div>';
        if (data.description) {
            html += '<div class="table-desc">' + escapeHtml(data.description) + '</div>';
        }

        html += '<div class="table-scroll-wrapper">';
        html += '<table class="chat-data-table">';
        html += '<thead><tr>';
        displayHeaders.forEach(function (col) {
            html += '<th>' + escapeHtml(col) + '</th>';
        });
        html += '</tr></thead>';
        html += '<tbody>';

        const displayRows = rows.slice(0, 100);
        displayRows.forEach(function (row, idx) {
            html += '<tr>';
            dataKeys.forEach(function (key) {
                const val = row[key];
                const displayVal = typeof val === 'number' ? formatNumber(val) : (val || '');
                html += '<td>' + escapeHtml(String(displayVal)) + '</td>';
            });
            html += '</tr>';
        });
        html += '</tbody></table>';
        html += '</div>';

        if (rows.length > 100) {
            html += '<div class="table-truncated">Showing 100 of ' + rows.length + ' rows</div>';
        }

        html += '<div class="response-actions">';
        html += renderSqlLink(data.sql);
        html += '<button class="action-link" onclick="ScanifyChatbot.exportTableCsv(this)"><i class="fa fa-download"></i> Export CSV</button>';
        html += '</div>';
        html += '</div>';
        return html;
    }

    function renderChartResponse(data) {
        const rows = data.data || [];
        const dataKeys = rows.length ? Object.keys(rows[0]) : [];
        const aiColumns = data.columns || [];
        const displayHeaders = (aiColumns.length === dataKeys.length) ? aiColumns : dataKeys.map(formatColumnName);
        const chartId = 'chat-chart-' + (++chartCounter);
        const chartType = data.chart_type || 'bar';

        // Empty state
        if (!rows.length) {
            let html = '<div class="chart-response">';
            html += '<div class="chart-title-bar">';
            html += '<div class="chart-title"><i class="fa fa-chart-bar mr-1"></i> ' + escapeHtml(data.title || 'Chart') + '</div>';
            html += '<div class="chart-type-badge">' + chartType + '</div>';
            html += '</div>';
            if (data.description) {
                html += '<div class="chart-desc">' + escapeHtml(data.description) + '</div>';
            }
            html += '<div class="empty-state"><i class="fa fa-chart-area"></i><span>No data available to generate this chart</span></div>';
            html += renderSqlLink(data.sql);
            html += '</div>';
            return html;
        }

        // Resolve label/value columns to actual data keys
        const labelCol = resolveColumnKey(data.label_column, dataKeys, aiColumns) || dataKeys[0] || '';
        const rawValueCols = data.value_columns || [];
        const valueCols = rawValueCols.length
            ? rawValueCols.map(function(c) { return resolveColumnKey(c, dataKeys, aiColumns) || c; })
            : dataKeys.slice(1);

        let html = '<div class="chart-response">';
        html += '<div class="chart-title-bar">';
        html += '<div class="chart-title"><i class="fa fa-chart-bar mr-1"></i> ' + escapeHtml(data.title || 'Chart') + '</div>';
        html += '<div class="chart-type-badge">' + chartType + '</div>';
        html += '</div>';
        if (data.description) {
            html += '<div class="chart-desc">' + escapeHtml(data.description) + '</div>';
        }

        html += '<div class="chart-canvas-wrapper">';
        html += '<canvas id="' + chartId + '"></canvas>';
        html += '</div>';

        // Store chart data in JS map (avoids HTML attribute escaping issues)
        chartDataStore[chartId] = {
            chartType: chartType,
            labelCol: labelCol,
            valueCols: valueCols,
            rows: rows
        };

        // Also show the data table below chart
        html += '<details class="chart-data-details">';
        html += '<summary><i class="fa fa-table mr-1"></i> View Data Table (' + rows.length + ' rows)</summary>';
        html += '<div class="table-scroll-wrapper" style="margin-top:8px;">';
        html += '<table class="chat-data-table compact">';
        html += '<thead><tr>';
        displayHeaders.forEach(function (col) {
            html += '<th>' + escapeHtml(col) + '</th>';
        });
        html += '</tr></thead><tbody>';
        rows.slice(0, 50).forEach(function (row) {
            html += '<tr>';
            dataKeys.forEach(function (key) {
                const val = row[key];
                const dv = typeof val === 'number' ? formatNumber(val) : (val || '');
                html += '<td>' + escapeHtml(String(dv)) + '</td>';
            });
            html += '</tr>';
        });
        html += '</tbody></table></div>';
        if (rows.length > 50) {
            html += '<div class="table-truncated">Showing 50 of ' + rows.length + ' rows</div>';
        }
        html += '</details>';

        html += '<div class="response-actions">';
        html += renderSqlLink(data.sql);
        html += '<button class="action-link" onclick="ScanifyChatbot.exportTableCsv(this)"><i class="fa fa-download"></i> Export CSV</button>';
        html += '</div>';
        html += '</div>';
        return html;
    }

    function renderSqlLink(sql) {
        if (!sql) return '';
        return '<button class="action-link" onclick="ScanifyChatbot.showSql(\'' +
            escapeHtml(btoa(unescape(encodeURIComponent(sql)))) +
            '\')"><i class="fa fa-code"></i> View SQL</button>';
    }

    // ── Chart Initialization ───────────────────────────
    function initPendingCharts(container) {
        const canvases = container.querySelectorAll('canvas[id^="chat-chart-"]');
        canvases.forEach(function (canvas) {
            const chartId = canvas.id;
            const stored = chartDataStore[chartId];
            if (!stored) return;

            const chartType = stored.chartType;
            const rows = stored.rows;
            if (!rows || !rows.length) return;

            // Resolve column names against actual row keys
            const rowKeys = Object.keys(rows[0]);
            var labelCol = resolveColumnKey(stored.labelCol, rowKeys, []);
            var valueCols = stored.valueCols.map(function (c) {
                return resolveColumnKey(c, rowKeys, []);
            });

            // Final fallback: if still unresolved, use first key for label, rest for values
            if (rowKeys.indexOf(labelCol) === -1) labelCol = rowKeys[0];
            valueCols = valueCols.filter(function (c) { return rowKeys.indexOf(c) !== -1; });
            if (!valueCols.length) {
                valueCols = rowKeys.filter(function (k) { return k !== labelCol; });
            }

            const labels = rows.map(function (r) { return r[labelCol] || ''; });

            // Color palette
            const colors = [
                '#4f46e5', '#10b981', '#f59e0b', '#ef4444', '#3b82f6',
                '#8b5cf6', '#ec4899', '#14b8a6', '#f97316', '#06b6d4',
                '#84cc16', '#e11d48', '#7c3aed', '#0ea5e9', '#d946ef'
            ];
            const bgColors = colors.map(function (c) { return c + 'CC'; });

            const isPieType = (chartType === 'pie' || chartType === 'doughnut');
            const realChartType = chartType === 'horizontalBar' ? 'bar' : chartType;

            const datasets = valueCols.map(function (col, idx) {
                const data = rows.map(function (r) {
                    const v = r[col];
                    return typeof v === 'number' ? v : parseFloat(v) || 0;
                });

                if (isPieType) {
                    return {
                        label: formatColumnName(col),
                        data: data,
                        backgroundColor: data.map(function (_, i) { return bgColors[i % bgColors.length]; }),
                        borderColor: data.map(function (_, i) { return colors[i % colors.length]; }),
                        borderWidth: 2
                    };
                }
                return {
                    label: formatColumnName(col),
                    data: data,
                    backgroundColor: bgColors[idx % bgColors.length],
                    borderColor: colors[idx % colors.length],
                    borderWidth: 2,
                    borderRadius: realChartType === 'bar' ? 4 : 0,
                    tension: 0.3,
                    fill: realChartType === 'line' ? 'origin' : undefined,
                    pointRadius: realChartType === 'line' ? 4 : undefined,
                    pointHoverRadius: realChartType === 'line' ? 6 : undefined
                };
            });

            const isHorizontal = chartType === 'horizontalBar';

            const chartConfig = {
                type: realChartType,
                data: { labels: labels, datasets: datasets },
                options: {
                    responsive: true,
                    maintainAspectRatio: false,
                    indexAxis: isHorizontal ? 'y' : 'x',
                    plugins: {
                        legend: {
                            display: valueCols.length > 1 || isPieType,
                            position: isPieType ? 'right' : 'top',
                            labels: {
                                font: { family: 'Inter', size: 12 },
                                padding: 16,
                                usePointStyle: true,
                                pointStyleWidth: 12
                            }
                        },
                        tooltip: {
                            backgroundColor: '#1e293b',
                            titleFont: { family: 'Inter', size: 13, weight: '600' },
                            bodyFont: { family: 'Inter', size: 12 },
                            padding: 12,
                            cornerRadius: 8,
                            displayColors: true,
                            callbacks: {
                                label: function (ctx) {
                                    var value;
                                    if (isHorizontal) {
                                        value = ctx.parsed.x;
                                    } else if (isPieType) {
                                        value = ctx.raw;
                                    } else {
                                        value = ctx.parsed.y;
                                    }
                                    return ctx.dataset.label + ': ' + formatNumber(value);
                                }
                            }
                        }
                    },
                    scales: isPieType ? {} : (isHorizontal ? {
                        x: {
                            grid: { color: '#f1f5f9' },
                            ticks: {
                                font: { family: 'Inter', size: 11 },
                                callback: function (val) { return formatNumber(val); }
                            },
                            beginAtZero: true
                        },
                        y: {
                            grid: { display: false },
                            ticks: {
                                font: { family: 'Inter', size: 11 },
                                autoSkip: false
                            }
                        }
                    } : {
                        x: {
                            grid: { display: false },
                            ticks: {
                                font: { family: 'Inter', size: 11 },
                                maxRotation: 45,
                                autoSkip: true,
                                maxTicksLimit: 20
                            }
                        },
                        y: {
                            grid: { color: '#f1f5f9' },
                            ticks: {
                                font: { family: 'Inter', size: 11 },
                                callback: function (val) { return formatNumber(val); }
                            },
                            beginAtZero: true
                        }
                    }),
                    animation: { duration: 600, easing: 'easeOutQuart' }
                }
            };

            try {
                if (chartInstances[canvas.id]) {
                    chartInstances[canvas.id].destroy();
                }
                chartInstances[canvas.id] = new Chart(canvas, chartConfig);
            } catch (e) {
                console.error('Chart render error:', e);
            }

            // Clean up stored data after rendering
            delete chartDataStore[chartId];
        });
    }

    // ── Typing Indicator ───────────────────────────────
    function showTypingIndicator() {
        const container = document.getElementById('chatMessages');
        const id = 'typing-' + Date.now();
        const div = document.createElement('div');
        div.className = 'chat-message assistant-message';
        div.id = id;
        div.innerHTML =
            '<div class="message-avatar assistant-avatar"><i class="fa fa-robot"></i></div>' +
            '<div class="message-content">' +
            '<div class="message-bubble assistant-bubble">' +
            '<div class="typing-indicator">' +
            '<div class="typing-dot"></div>' +
            '<div class="typing-dot"></div>' +
            '<div class="typing-dot"></div>' +
            '<span class="typing-text">Analyzing your query...</span>' +
            '</div></div></div>';
        container.appendChild(div);
        scrollToBottom();
        return id;
    }

    function removeTypingIndicator(id) {
        const el = document.getElementById(id);
        if (el) el.remove();
    }

    // ── Utility Functions ──────────────────────────────
    function escapeHtml(str) {
        if (!str) return '';
        const div = document.createElement('div');
        div.textContent = str;
        return div.innerHTML;
    }

    function formatNumber(val) {
        if (val === null || val === undefined || val === '') return '';
        const num = Number(val);
        if (isNaN(num)) return val;
        if (Number.isInteger(num) && Math.abs(num) < 1e15) {
            return num.toLocaleString('en-IN');
        }
        if (Math.abs(num) >= 100) {
            return num.toLocaleString('en-IN', { maximumFractionDigits: 2 });
        }
        return num.toLocaleString('en-IN', { maximumFractionDigits: 4 });
    }

    function formatColumnName(col) {
        if (!col) return '';
        return col
            .replace(/_/g, ' ')
            .replace(/\b\w/g, function (c) { return c.toUpperCase(); });
    }

    /**
     * Resolve an AI-provided column name (could be display name like "Product Name"
     * or actual key like "product_name") to the real data key from the row.
     */
    function resolveColumnKey(colName, dataKeys, aiColumns) {
        if (!colName) return '';
        // Direct match — already a data key
        if (dataKeys.indexOf(colName) !== -1) return colName;
        // Match by position in AI columns array
        var idx = aiColumns.indexOf(colName);
        if (idx !== -1 && idx < dataKeys.length) return dataKeys[idx];
        // Fuzzy: normalize both sides (lowercase, strip spaces/underscores) and compare
        var norm = colName.toLowerCase().replace(/[\s_]/g, '');
        for (var i = 0; i < dataKeys.length; i++) {
            if (dataKeys[i].toLowerCase().replace(/[\s_]/g, '') === norm) return dataKeys[i];
        }
        return colName; // fallback — return as-is
    }

    // ── Public Actions ─────────────────────────────────
    function showSql(encodedSql) {
        try {
            const sql = decodeURIComponent(escape(atob(encodedSql)));
            document.getElementById('sqlDisplay').textContent = sql;
            $('#sqlModal').modal('show');
        } catch (e) {
            console.error('SQL decode error:', e);
        }
    }

    function clearChat() {
        if (!confirm('Clear all conversation history?')) return;
        conversationHistory = [];
        const container = document.getElementById('chatMessages');
        // Keep only the first (welcome) message
        while (container.children.length > 1) {
            container.removeChild(container.lastChild);
        }
        // Destroy all chart instances
        Object.keys(chartInstances).forEach(function (key) {
            if (chartInstances[key]) chartInstances[key].destroy();
        });
        chartInstances = {};
        chartDataStore = {};
        chartCounter = 0;
    }

    function exportChat() {
        if (!conversationHistory.length) {
            alert('No conversation to export.');
            return;
        }
        let text = 'Scanify AI Chatbot — Conversation Export\n';
        text += 'Division: ' + (config.division || '') + '\n';
        text += 'Date: ' + new Date().toLocaleString() + '\n';
        text += '='.repeat(60) + '\n\n';

        conversationHistory.forEach(function (entry) {
            const role = entry.role === 'user' ? 'You' : 'AI';
            text += '[' + role + ']\n' + entry.content + '\n\n';
        });

        const blob = new Blob([text], { type: 'text/plain' });
        const a = document.createElement('a');
        a.href = URL.createObjectURL(blob);
        a.download = 'scanify-chat-' + new Date().toISOString().slice(0, 10) + '.txt';
        a.click();
        URL.revokeObjectURL(a.href);
    }

    function exportTableCsv(btn) {
        const tableWrapper = btn.closest('.table-response, .chart-response');
        if (!tableWrapper) return;
        const table = tableWrapper.querySelector('.chat-data-table');
        if (!table) return;

        let csv = '';
        const rows = table.querySelectorAll('tr');
        rows.forEach(function (row) {
            const cells = row.querySelectorAll('th, td');
            const rowData = [];
            cells.forEach(function (cell) {
                let val = cell.textContent.replace(/"/g, '""');
                rowData.push('"' + val + '"');
            });
            csv += rowData.join(',') + '\n';
        });

        const blob = new Blob([csv], { type: 'text/csv' });
        const a = document.createElement('a');
        a.href = URL.createObjectURL(blob);
        a.download = 'scanify-data-' + new Date().toISOString().slice(0, 10) + '.csv';
        a.click();
        URL.revokeObjectURL(a.href);
    }

    // ── Suggestion Click Handler ───────────────────────
    window.useSuggestion = function (el) {
        const small = el.querySelector('small');
        if (small) {
            const input = document.getElementById('chatInput');
            if (input) {
                input.value = small.textContent;
                input.focus();
                autoResizeInput();
                updateCharCount();
            }
        }
    };

    window.copySql = function () {
        const sqlText = document.getElementById('sqlDisplay').textContent;
        if (navigator.clipboard) {
            navigator.clipboard.writeText(sqlText);
        } else {
            const ta = document.createElement('textarea');
            ta.value = sqlText;
            document.body.appendChild(ta);
            ta.select();
            document.execCommand('copy');
            document.body.removeChild(ta);
        }
    };

    // Public API
    return {
        init: init,
        showSql: showSql,
        exportTableCsv: exportTableCsv
    };

})();

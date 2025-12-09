function setupDragAndDrop(type) {
    const dropZone = document.getElementById(`drop-zone-${type}`);
    const fileInput = document.getElementById(`file-${type}`);

    dropZone.addEventListener('click', () => fileInput.click());

    dropZone.addEventListener('dragover', (e) => {
        e.preventDefault();
        dropZone.classList.add('dragover');
    });

    dropZone.addEventListener('dragleave', () => {
        dropZone.classList.remove('dragover');
    });

    dropZone.addEventListener('drop', (e) => {
        e.preventDefault();
        dropZone.classList.remove('dragover');
        if (e.dataTransfer.files.length) {
            handleFiles(e.dataTransfer.files, type);
        }
    });

    fileInput.addEventListener('change', () => {
        if (fileInput.files.length) {
            handleFiles(fileInput.files, type);
        }
    });
}

function handleFiles(files, type) {
    const file = files[0];
    uploadFile(file, type);
}

function uploadFile(file, type) {
    const formData = new FormData();
    formData.append('file', file);
    formData.append('type', type);

    addLog(`正在上传文件: ${file.name}...`);

    fetch('/upload', {
        method: 'POST',
        body: formData
    })
        .then(response => response.json())
        .then(data => {
            if (data.error) {
                addLog(`❌ 上传失败: ${data.error}`, 'error');
            } else {
                addLog(`✅ 上传成功!`, 'system');
            }
        })
        .catch(error => {
            addLog(`❌ 上传错误: ${error}`, 'error');
        });
}

function runTask(type) {
    const statusSpan = document.getElementById(`status-${type}`);
    statusSpan.textContent = "运行中...";
    statusSpan.classList.add('running');

    fetch('/api/run', {
        method: 'POST',
        headers: {
            'Content-Type': 'application/json'
        },
        body: JSON.stringify({ type: type })
    })
        .then(response => response.json())
        .then(data => {
            if (data.error) {
                addLog(`❌ 启动失败: ${data.error}`, 'error');
                resetStatus(type);
            } else {
                addLog(`🚀 ${data.message}`, 'system');
                // Auto reset status after some time or wait for specific log message (simplified here)
                setTimeout(() => resetStatus(type), 5000);
            }
        })
        .catch(error => {
            addLog(`❌ 请求错误: ${error}`, 'error');
            resetStatus(type);
        });
}

function resetStatus(type) {
    const statusSpan = document.getElementById(`status-${type}`);
    statusSpan.textContent = "就绪";
    statusSpan.classList.remove('running');
}

function addLog(message, type = '') {
    const terminal = document.getElementById('terminal');
    const div = document.createElement('div');
    div.className = `log-line ${type}`;
    div.textContent = `[${new Date().toLocaleTimeString()}] ${message}`;
    terminal.appendChild(div);
    terminal.scrollTop = terminal.scrollHeight;
}

function clearLogs() {
    document.getElementById('terminal').innerHTML = '<div class="log-line system">>>> 日志已清除</div>';
}

// SSE for Log Streaming
function setupLogStream() {
    const eventSource = new EventSource('/api/logs');
    eventSource.onmessage = function (event) {
        const data = JSON.parse(event.data);
        if (data.message) {
            // Check for completion messages to reset status if needed
            if (data.message.includes("化验汇总任务完成")) resetStatus('hy');
            if (data.message.includes("称重汇总任务完成")) resetStatus('cz');

            addLog(data.message);
        }
    };
    eventSource.onerror = function () {
        // addLog("⚠️ 日志连接断开，尝试重连...", 'error');
        eventSource.close();
        setTimeout(setupLogStream, 2000);
    };
}

// Preview Results with Tabs
function previewResults(type) {
    const container = document.getElementById(`preview-${type}`);
    if (!container) return;

    container.innerHTML = '<div class="loading">正在加载数据预览...</div>';
    container.style.display = 'block';

    fetch(`/api/preview/${type}`)
        .then(response => {
            if (!response.ok) {
                return response.json().then(err => { throw new Error(err.error || 'Network response was not ok'); });
            }
            return response.json();
        })
        .then(data => {
            container.innerHTML = '';
            if (Object.keys(data).length === 0) {
                container.innerHTML = '<div class="no-data">暂无数据 preview available.</div>';
                return;
            }

            const tabHeader = document.createElement('div');
            tabHeader.className = 'tab-header';

            const tabContentContainer = document.createElement('div');
            tabContentContainer.className = 'tab-content-container';

            let first = true;
            const sheets = Object.entries(data);

            sheets.forEach(([sheetName, tableHtml], index) => {
                // Tab Button
                const btn = document.createElement('button');
                btn.className = `tab-btn ${first ? 'active' : ''}`;
                btn.textContent = sheetName;
                btn.dataset.target = `sheet-${type}-${index}`;

                // Tab Content
                const content = document.createElement('div');
                content.id = `sheet-${type}-${index}`;
                content.className = `sheet-content ${first ? 'active' : ''}`;
                // Simplified wrapper without 'table-responsive' since DataTables handles scrolling
                content.innerHTML = `<div>${tableHtml}</div>`;

                // Click Event
                btn.addEventListener('click', () => {
                    // Deactivate all in this container
                    tabHeader.querySelectorAll('.tab-btn').forEach(b => b.classList.remove('active'));
                    tabContentContainer.querySelectorAll('.sheet-content').forEach(c => c.classList.remove('active'));

                    // Activate clicked
                    btn.classList.add('active');
                    content.classList.add('active');

                    // Adjust DataTables columns on tab switch
                    const table = $(content).find('table').DataTable();
                    table.columns.adjust().draw();
                });

                tabHeader.appendChild(btn);
                tabContentContainer.appendChild(content);
                first = false;
            });

            container.appendChild(tabHeader);
            container.appendChild(tabContentContainer);

            // Initialize DataTables
            $(container).find('table').each(function () {
                const table = $(this).DataTable({
                    colReorder: true,
                    paging: true,
                    pageLength: 20,
                    lengthMenu: [[20, 50, 100, -1], [20, 50, 100, "全部"]],
                    scrollX: true,
                    scrollY: '550px',
                    scrollCollapse: true,
                    autoWidth: false, // Disable auto width to let CSS control, helps with alignment
                    language: {
                        "sProcessing": "处理中...",
                        "sLengthMenu": "显示 _MENU_ 项结果",
                        "sZeroRecords": "没有匹配结果",
                        "sInfo": "显示第 _START_ 至 _END_ 项结果，共 _TOTAL_ 项",
                        "sInfoEmpty": "显示第 0 至 0 项结果，共 0 项",
                        "sInfoFiltered": "(由 _MAX_ 项结果过滤)",
                        "sSearch": "搜索:",
                        "sEmptyTable": "表中数据为空",
                        "oPaginate": {
                            "sFirst": "首页",
                            "sPrevious": "上页",
                            "sNext": "下页",
                            "sLast": "末页"
                        }
                    }
                });

                // Fix for alignment issues: Adjust columns after a short delay to ensure valid widths
                setTimeout(() => {
                    table.columns.adjust().draw();
                }, 200);
            });
        })
        .catch(error => {
            container.innerHTML = `<div class="error-msg">加载失败: ${error.message} <br> 请确保已运行任务生成了报表。</div>`;
        });
}

// Main Tab Switching
function switchMainTab(type) {
    // Buttons
    document.querySelectorAll('.main-tab-btn').forEach(btn => {
        btn.classList.remove('active');
        if (btn.getAttribute('onclick').includes(type)) {
            btn.classList.add('active');
        }
    });

    // Content
    document.querySelectorAll('.main-tab-content').forEach(content => {
        content.classList.remove('active');
    });
    document.getElementById(`main-tab-${type}`).classList.add('active');
}

// Initialize
document.addEventListener('DOMContentLoaded', () => {
    setupDragAndDrop('hy');
    setupDragAndDrop('cz');
    setupLogStream();

    // Auto preview if data exists? Maybe later.
});

const API_BASE = '/api';

// UI Elements
const tabs = document.querySelectorAll('.tab-pane');
const navBtns = document.querySelectorAll('.nav-btn');
const toastEl = document.getElementById('toast');
const toastMsg = document.getElementById('toast-msg');
const loadingOverlay = document.getElementById('loading-overlay');
const loadingText = document.getElementById('loading-text');
const repoInfoText = document.getElementById('repo-info-text');
const logText = document.getElementById('log-text');
const repoPathInput = document.getElementById('repo-path');

// Tab Navigation
navBtns.forEach(btn => {
    btn.addEventListener('click', () => {
        // Remove active class from all
        navBtns.forEach(b => b.classList.remove('active'));
        tabs.forEach(t => t.classList.remove('active'));

        // Add active class to clicked
        btn.classList.add('active');
        const tabId = `tab-${btn.dataset.tab}`;
        document.getElementById(tabId).classList.add('active');

        if (btn.dataset.tab === 'logs') {
            fetchLogs();
        }
    });
});

// Toast Notification
function showToast(message, isError = false) {
    toastMsg.textContent = message;
    if (isError) {
        toastEl.classList.add('error');
    } else {
        toastEl.classList.remove('error');
    }

    toastEl.classList.remove('hidden');
    toastEl.classList.add('show');

    setTimeout(() => {
        toastEl.classList.remove('show');
        setTimeout(() => toastEl.classList.add('hidden'), 400);
    }, 4000);
}

// Loading Overlay
function showLoading(text = 'Processing...') {
    loadingText.textContent = text;
    loadingOverlay.classList.remove('hidden');
}

function hideLoading() {
    loadingOverlay.classList.add('hidden');
}

// API Helper
async function apiCall(endpoint, method = 'POST', data = null) {
    try {
        const options = {
            method,
            headers: {
                'Content-Type': 'application/json'
            }
        };
        if (data) {
            options.body = JSON.stringify(data);
        }

        const res = await fetch(`${API_BASE}${endpoint}`, options);
        if (!res.ok) {
            throw new Error(`HTTP ${res.status}`);
        }
        return await res.json();
    } catch (err) {
        console.error(err);
        return { success: false, message: err.message };
    }
}

// Initial Load
async function initApp() {
    const res = await apiCall('/path', 'GET');
    if (res && res.path) {
        repoPathInput.value = res.path;
    }

    // Start polling logs every 3 seconds
    setInterval(fetchLogs, 3000);
}

// Fetch Logs
async function fetchLogs() {
    const res = await apiCall('/logs', 'GET');
    if (res && res.logs) {
        logText.textContent = res.logs.join('\n');
        // auto scroll to bottom
        const viewer = document.getElementById('log-viewer');
        if (viewer.scrollTop + viewer.clientHeight >= viewer.scrollHeight - 50) {
            viewer.scrollTop = viewer.scrollHeight;
        }
    }
}

// Repository Actions
document.getElementById('btn-set-path').addEventListener('click', async () => {
    const path = repoPathInput.value.trim();
    if (!path) return;
    showLoading('Setting Repository Path...');
    const res = await apiCall('/path', 'POST', { path });
    hideLoading();
    if (res.success) {
        showToast(`Repository path set to ${res.path}`);
    } else {
        showToast(res.message || 'Failed to set path', true);
    }
});

document.getElementById('btn-init').addEventListener('click', async () => {
    showLoading('Initializing Repository...');
    const res = await apiCall('/init');
    hideLoading();
    if (res.success) showToast('Repository Initialized');
    else showToast(res.message, true);
});

document.getElementById('btn-sync').addEventListener('click', async () => {
    const refUrl = document.getElementById('ref-url').value.trim();
    if (!refUrl) {
        showToast('Reference URL is required', true);
        return;
    }
    showLoading('Synchronizing Repository... this may take a while.');
    const res = await apiCall('/sync', 'POST', { refUrl });
    hideLoading();
    if (res.success) showToast('Sync Complete');
    else showToast(res.message, true);
});

document.getElementById('btn-cleanup').addEventListener('click', async () => {
    showLoading('Cleaning Up Repository...');
    const res = await apiCall('/cleanup');
    hideLoading();
    if (res.success) showToast('Cleanup Complete');
    else showToast(res.message, true);
});

document.getElementById('btn-info').addEventListener('click', async () => {
    showLoading('Fetching Information...');
    const res = await apiCall('/info', 'GET');
    hideLoading();
    if (res.success) {
        repoInfoText.textContent = res.info || 'No repository info found.';
        if (res.settings) {
            document.getElementById('s-missing').value = res.settings.OnRemoteFileNotFound || 'Fail';
            document.getElementById('s-cache').value = res.settings.OfflineCacheMode || 'Disable';
            document.getElementById('s-report').value = res.settings.RepositoryReport || 'CSV';
        }
    } else {
        showToast(res.message, true);
    }
});

document.getElementById('btn-refresh-logs').addEventListener('click', () => {
    fetchLogs();
});

// Settings Form
document.getElementById('settings-form').addEventListener('submit', async (e) => {
    e.preventDefault();
    const data = {
        missing: document.getElementById('s-missing').value,
        cache: document.getElementById('s-cache').value,
        report: document.getElementById('s-report').value
    };
    showLoading('Applying Settings...');
    const res = await apiCall('/settings', 'POST', data);
    hideLoading();
    if (res.success) showToast('Settings Applied');
    else showToast(res.message, true);
});

// Filters Form
function getCheckedValues(containerId) {
    const inputs = document.querySelectorAll(`#${containerId} input:checked`);
    return Array.from(inputs).map(inp => inp.value);
}

document.getElementById('filter-form').addEventListener('submit', async (e) => {
    e.preventDefault();
    const data = {
        Platform: document.getElementById('f-platform').value.trim(),
        Os: document.getElementById('f-os').value,
        OsVer: document.getElementById('f-osver').value,
        PreferLtsc: document.getElementById('f-ltsc').checked,
        Category: getCheckedValues('f-category'),
        ReleaseType: getCheckedValues('f-release'),
        Characteristic: getCheckedValues('f-char')
    };

    showLoading('Adding Filter...');
    const res = await apiCall('/filter', 'POST', data);
    hideLoading();
    if (res.success) {
        showToast('Filter Added Successfully');
        // Reset form
        document.getElementById('f-platform').value = '';
        document.querySelectorAll('#filter-form input[type="checkbox"]').forEach(c => c.checked = false);
    } else {
        showToast(res.message, true);
    }
});

// Fleet Management
let fleetState = [];

function renderFleetTable() {
    const tbody = document.getElementById('fleet-table-body');
    const emptyState = document.getElementById('fleet-empty-state');
    const bulkActions = document.getElementById('fleet-bulk-actions');

    tbody.innerHTML = '';

    if (fleetState.length === 0) {
        tbody.parentElement.classList.add('hidden');
        emptyState.classList.remove('hidden');
        bulkActions.classList.add('hidden');
        return;
    }

    tbody.parentElement.classList.remove('hidden');
    emptyState.classList.add('hidden');
    bulkActions.classList.remove('hidden');

    fleetState.forEach((endpoint, index) => {
        const tr = document.createElement('tr');

        let statusBadge = `<span class="badge ${endpoint.status === 'Online' ? 'badge-success' : 'badge-danger'}">${endpoint.status || 'Unknown'}</span>`;
        if (endpoint.status === 'Scanning') {
            statusBadge = `<span class="badge badge-warning"><i class="fa-solid fa-spinner fa-spin"></i> Scanning</span>`;
        }

        let updatesHtml = '<span class="text-muted">N/A</span>';
        if (endpoint.applicable && endpoint.applicable.length > 0) {
            updatesHtml = `
                <div style="display: flex; align-items: center; gap: 0.5rem;">
                    <span class="badge badge-warning">${endpoint.applicable.length} Available</span>
                    <button class="btn btn-accent btn-sm btn-deploy-single" data-index="${index}" title="Push Updates"><i class="fa-solid fa-bolt"></i></button>
                </div>
            `;
        } else if (endpoint.applicable && endpoint.applicable.length === 0) {
            updatesHtml = `<span class="badge badge-success">Up to date</span>`;
        }

        tr.innerHTML = `
            <td><input type="checkbox" class="row-checkbox endpoint-select" data-index="${index}"></td>
            <td><strong>${endpoint.hostname}</strong></td>
            <td>${statusBadge}</td>
            <td>
                ${endpoint.system ? `<div class="sys-item"><i class="fa-solid fa-laptop"></i> ${endpoint.system.model}</div>
                   <div class="sys-item small text-muted">ID: ${endpoint.system.platform} | SN: ${endpoint.system.serial}</div>` : '<span class="text-muted">N/A</span>'}
            </td>
            <td>${endpoint.system ? `<i class="fa-brands fa-windows"></i> ${endpoint.system.os}` : '<span class="text-muted">N/A</span>'}</td>
            <td>${updatesHtml}</td>
            <td>
                <div class="table-actions">
                    <button class="btn btn-icon btn-scan" data-index="${index}" title="Scan Endpoint"><i class="fa-solid fa-radar"></i></button>
                    <button class="btn btn-icon btn-remove text-danger" data-index="${index}" title="Remove"><i class="fa-solid fa-trash"></i></button>
                </div>
            </td>
        `;
        tbody.appendChild(tr);
    });

    // Add event listeners to new buttons
    document.querySelectorAll('.btn-scan').forEach(btn => {
        btn.addEventListener('click', (e) => scanEndpoint(e.currentTarget.dataset.index));
    });

    document.querySelectorAll('.btn-remove').forEach(btn => {
        btn.addEventListener('click', (e) => {
            fleetState.splice(e.currentTarget.dataset.index, 1);
            renderFleetTable();
        });
    });

    document.querySelectorAll('.btn-deploy-single').forEach(btn => {
        btn.addEventListener('click', async (e) => {
            const index = e.currentTarget.dataset.index;
            const endpoint = fleetState[index];
            if (!endpoint.applicable || endpoint.applicable.length === 0) return;

            showLoading(`Pushing ${endpoint.applicable.length} updates to ${endpoint.hostname}...`);
            const res = await apiCall('/deploy', 'POST', { targets: endpoint.hostname, packages: endpoint.applicable });
            hideLoading();

            if (res.success) {
                showToast('Deployment Command Sent. Check Logs for status.');
            } else {
                showToast(res.message, true);
            }
        });
    });

    updateBulkSelection();
}

document.getElementById('btn-add-endpoint').addEventListener('click', () => {
    const hostname = prompt('Enter the hostname or IP address of the target PC:');
    if (hostname && hostname.trim()) {
        fleetState.push({
            hostname: hostname.trim(),
            status: 'Pending',
            system: null
        });
        renderFleetTable();
        scanEndpoint(fleetState.length - 1);
    }
});

async function scanEndpoint(index) {
    const endpoint = fleetState[index];
    endpoint.status = 'Scanning';
    renderFleetTable();

    const res = await apiCall('/fleet/scan', 'POST', { hostname: endpoint.hostname });

    if (res.success && res.status) {
        endpoint.status = res.status;
        if (res.system) endpoint.system = res.system;
        if (res.applicable) endpoint.applicable = res.applicable;
    } else {
        endpoint.status = 'Offline';
    }
    renderFleetTable();
}

// Select All logic
document.getElementById('fleet-select-all').addEventListener('change', (e) => {
    const checkboxes = document.querySelectorAll('.endpoint-select');
    checkboxes.forEach(cb => cb.checked = e.target.checked);
});

function updateBulkSelection() {
    const checkAll = document.getElementById('fleet-select-all');
    if (checkAll) checkAll.checked = false;
}

// Bulk Deploy logic
document.getElementById('btn-bulk-deploy').addEventListener('click', async () => {
    const selectedIndexes = Array.from(document.querySelectorAll('.endpoint-select:checked')).map(cb => parseInt(cb.dataset.index));

    if (selectedIndexes.length === 0) {
        showToast('Please select at least one endpoint.', true);
        return;
    }

    let totalDeployments = 0;

    for (const idx of selectedIndexes) {
        const endpoint = fleetState[idx];
        if (endpoint.applicable && endpoint.applicable.length > 0) {
            showLoading(`Pushing ${endpoint.applicable.length} updates to ${endpoint.hostname}...`);
            const res = await apiCall('/deploy', 'POST', { targets: endpoint.hostname, packages: endpoint.applicable });
            if (res.success) {
                totalDeployments++;
            } else {
                showToast(`Failed on ${endpoint.hostname}: ${res.message}`, true);
            }
        }
    }

    hideLoading();

    if (totalDeployments > 0) {
        showToast(`Deployment Commands Sent to ${totalDeployments} endpoints. Check Logs for status.`);
    } else {
        showToast('No applicable updates found for selected endpoints.', true);
    }
});

// OS dropdown logic
document.getElementById('f-os').addEventListener('change', (e) => {
    const val = e.target.value;
    const osverGroup = document.getElementById('group-osver');
    if (val === '*') {
        osverGroup.style.opacity = '0.5';
        document.getElementById('f-osver').disabled = true;
    } else {
        osverGroup.style.opacity = '1';
        document.getElementById('f-osver').disabled = false;
    }
});

// Run Init
window.addEventListener('DOMContentLoaded', initApp);

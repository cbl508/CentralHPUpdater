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
        showToast('Information Updated');
    } else {
        showToast(res.message, true);
    }
});

document.getElementById('btn-refresh-logs').addEventListener('click', () => {
    fetchLogs();
    showToast('Logs refreshed');
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

// Deploy Form
document.getElementById('deploy-form').addEventListener('submit', async (e) => {
    e.preventDefault();
    const data = {
        targets: document.getElementById('d-targets').value,
        packages: document.getElementById('d-packages').value.split(',').map(s => s.trim()).filter(Boolean)
    };

    showLoading('Executing Deployment...');
    const res = await apiCall('/deploy', 'POST', data);
    hideLoading();
    if (res.success) {
        showToast('Deployment Command Sent. Check Logs for status.');
    } else {
        showToast(res.message, true);
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

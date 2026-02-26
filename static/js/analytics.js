// Analytics Dashboard JavaScript
// Global variables
let allClaims = [];
let filteredClaims = [];
let statusChart = null;
let trendChart = null;
let modelChart = null;
let branchChart = null;
let largeBranchChart = null;

// Initialize dashboard
document.addEventListener('DOMContentLoaded', function () {
    loadAnalyticsData();
    attachFilterListeners();
});

// Load data from backend
async function loadAnalyticsData() {
    try {
        console.log('Loading analytics data...');
        const response = await fetch('/api/analytics-data');
        const data = await response.json();

        if (data.success) {
            allClaims = data.claims;
            filteredClaims = [...allClaims];

            console.log('✅ Data loaded successfully');
            console.log('Total claims:', allClaims.length);

            // Debug: Show first claim data structure
            if (allClaims.length > 0) {
                console.log('Sample claim data:', allClaims[0]);
                console.log('Sample mobile number:', allClaims[0].mobile_number, 'Type:', typeof allClaims[0].mobile_number);
            }

            updateAllComponents();
        } else {
            console.error('❌ Failed to load analytics data:', data);
        }
    } catch (error) {
        console.error('❌ Error loading analytics:', error);
    }
}

// Update all dashboard components
function updateAllComponents() {
    updateKPIs();
    updateStatusChart();
    updateTrendChart();
    updateModelChart();
    updateBranchChart();
    updateReplacementFunnel();
    updateRecentClaims();
    updatePendingActions();
    updateClaimsTable();
}

// 🔹 KPI CALCULATIONS
function updateKPIs() {
    const claims = filteredClaims;

    // Total Claims
    document.getElementById('kpi-total').textContent = claims.length;

    // Open Claims
    const openClaims = claims.filter(c => {
        const s = (c.status || '').toLowerCase();

        // Exclude definitively closed statuses
        if (s === 'settled' || s === 'closed' || s === 'repair completed') return false;

        // For Replacement Approved, it's only "closed" if the workflow is complete
        if (s.includes('replacement') && s.includes('approved')) {
            return !c.complete;
        }

        // All other statuses (Registered, Follow Up, etc.) are Open
        return true;
    });
    document.getElementById('kpi-open').textContent = openClaims.length;

    // Replacement Pending (any replacement stage incomplete)
    const replacementPending = claims.filter(c => {
        // Strict check: Only count if status is 'Replacement Approved' or 'Replacement approved'
        const hasReplacementStatus =
            c.status === 'Replacement Approved' ||
            c.status === 'Replacement approved';

        if (!hasReplacementStatus) return false;

        return !c.replacement_confirmation ||
            !c.replacement_osg_approval ||
            !c.replacement_mail_store ||
            !c.replacement_invoice_gen ||
            !c.replacement_invoice_sent ||
            !c.replacement_settled_accounts;
    });
    document.getElementById('kpi-replacement').textContent = replacementPending.length;

    // Repair Completed
    const repairCompleted = claims.filter(c => c.status === 'Repair Completed');
    const kpiRepairEl = document.getElementById('kpi-repair-completed');
    if (kpiRepairEl) kpiRepairEl.textContent = repairCompleted.length;

    // Replacement Approved Completed
    const settledClaims = claims.filter(c => {
        const s = (c.status || '').toLowerCase();
        const hasReplacementStatus = s.includes('replacement') && s.includes('approved');
        return c.complete && hasReplacementStatus;
    });
    document.getElementById('kpi-settled').textContent = settledClaims.length;

    // Closed as Completed (New KPI)
    const closedClaims = claims.filter(c => (c.status || '').toLowerCase() === 'closed');
    const kpiClosedEl = document.getElementById('kpi-closed');
    if (kpiClosedEl) kpiClosedEl.textContent = closedClaims.length;


    // Rejected Claims
    const rejectedClaims = claims.filter(c => {
        const s = (c.status || '').toLowerCase();
        return s === 'rejected';
    });
    const kpiRejectedEl = document.getElementById('kpi-rejected');
    if (kpiRejectedEl) kpiRejectedEl.textContent = rejectedClaims.length;

    // No Issue / On Call Resolution
    const noIssueClaims = claims.filter(c => {
        const s = (c.status || '').toLowerCase();
        return s === 'no issue' || s === 'on call resolution' || s === 'no issue/on call resolution' || s === 'no issue / on call resolution';
    });
    const kpiNoIssueEl = document.getElementById('kpi-no-issue');
    if (kpiNoIssueEl) kpiNoIssueEl.textContent = noIssueClaims.length;

    // Today's New Claims
    const today = new Date().toISOString().split('T')[0];
    const todayClaims = claims.filter(c =>
        c.submitted_date && c.submitted_date.startsWith(today)
    );
    document.getElementById('kpi-today').textContent = todayClaims.length;

    // This Month's Claims
    const now = new Date();
    const monthStart = new Date(now.getFullYear(), now.getMonth(), 1).toISOString().split('T')[0];
    const monthClaims = claims.filter(c =>
        c.submitted_date && c.submitted_date >= monthStart
    );
    document.getElementById('kpi-month').textContent = monthClaims.length;
}

// 🔹 STATUS DISTRIBUTION CHART
function updateStatusChart() {
    const ctx = document.getElementById('statusChart');

    // Count claims by status
    const statusCounts = {};
    filteredClaims.forEach(claim => {
        let status = claim.status || 'Unknown';
        // Normalize status
        if (status.includes('Replacement')) status = 'Replacement';
        statusCounts[status] = (statusCounts[status] || 0) + 1;
    });

    const labels = Object.keys(statusCounts);
    const data = Object.values(statusCounts);

    // Color scheme
    const colors = [
        '#3b82f6', '#10b981', '#f59e0b', '#8b5cf6',
        '#ec4899', '#6366f1', '#14b8a6', '#64748b'
    ];

    // Destroy existing chart
    if (statusChart) {
        statusChart.destroy();
    }

    // Create new chart
    statusChart = new Chart(ctx, {
        type: 'doughnut',
        data: {
            labels: labels,
            datasets: [{
                data: data,
                backgroundColor: colors.slice(0, labels.length),
                borderWidth: 0,
                hoverOffset: 4
            }]
        },
        options: {
            responsive: true,
            maintainAspectRatio: false,
            plugins: {
                legend: {
                    position: 'bottom',
                    labels: {
                        padding: 15,
                        usePointStyle: true,
                        font: { size: 11 }
                    }
                },
                title: { display: false }
            },
            layout: { padding: 10 }
        }
    });
}

// 🔹 TREND ANALYTICS (Line Chart)
function updateTrendChart() {
    const ctx = document.getElementById('trendChart');
    if (!ctx) return;

    // Get last 15 days
    const dates = [];
    const counts = [];
    const today = new Date();

    for (let i = 14; i >= 0; i--) {
        const d = new Date(today);
        d.setDate(today.getDate() - i);
        dates.push(d.toISOString().split('T')[0]);
        counts.push(0);
    }

    // Count claims per date
    filteredClaims.forEach(c => {
        if (c.submitted_date) {
            const idx = dates.indexOf(c.submitted_date);
            if (idx !== -1) counts[idx]++;
        }
    });

    // Destroy existing
    if (trendChart) trendChart.destroy();

    trendChart = new Chart(ctx, {
        type: 'line',
        data: {
            labels: dates.map(d => {
                const parts = d.split('-');
                return `${parts[2]}/${parts[1]}`;
            }),
            datasets: [{
                label: 'Claims Submitted',
                data: counts,
                borderColor: '#3b82f6',
                backgroundColor: 'rgba(59, 130, 246, 0.1)',
                borderWidth: 2,
                fill: true,
                tension: 0.4,
                pointRadius: 3
            }]
        },
        options: {
            responsive: true,
            maintainAspectRatio: false,
            plugins: {
                legend: { display: false }
            },
            scales: {
                y: {
                    beginAtZero: true,
                    ticks: { precision: 0 }
                },
                x: {
                    grid: { display: false }
                }
            }
        }
    });
}

// 🔹 TOP MODELS CHART (Bar Chart)
function updateModelChart() {
    const ctx = document.getElementById('modelChart');
    if (!ctx) return;

    // Count by Model
    const modelCounts = {};
    filteredClaims.forEach(c => {
        const m = c.model || c.product || 'Unknown';
        modelCounts[m] = (modelCounts[m] || 0) + 1;
    });

    // Sort and Top 5
    const sorted = Object.entries(modelCounts)
        .sort((a, b) => b[1] - a[1])
        .slice(0, 5);

    // Destroy existing
    if (modelChart) modelChart.destroy();

    modelChart = new Chart(ctx, {
        type: 'bar',
        data: {
            labels: sorted.map(i => i[0].length > 15 ? i[0].substr(0, 15) + '...' : i[0]),
            datasets: [{
                label: 'Claims',
                data: sorted.map(i => i[1]),
                backgroundColor: [
                    '#3b82f6', '#2563eb', '#1d4ed8', '#1e40af', '#1e3a8a'
                ],
                borderRadius: 4
            }]
        },
        options: {
            indexAxis: 'y',
            responsive: true,
            maintainAspectRatio: false,
            plugins: {
                legend: { display: false }
            },
            scales: {
                x: {
                    beginAtZero: true,
                    grid: { display: false },
                    ticks: { precision: 0 }
                },
                y: {
                    grid: { display: false }
                }
            }
        }
    });
}

// 🔹 BRANCH ANALYSIS CHART
function updateBranchChart() {
    const ctx = document.getElementById('branchChart');
    if (!ctx) return;

    // Count by Branch
    const branchCounts = {};
    filteredClaims.forEach(c => {
        const b = c.branch || 'Unknown';
        branchCounts[b] = (branchCounts[b] || 0) + 1;
    });

    // Sort Descending (Top 7)
    const sorted = Object.entries(branchCounts)
        .sort((a, b) => b[1] - a[1])
        .slice(0, 7);

    // Destroy existing
    if (branchChart) branchChart.destroy();

    branchChart = new Chart(ctx, {
        type: 'bar',
        data: {
            labels: sorted.map(i => i[0]),
            datasets: [{
                label: 'Claims',
                data: sorted.map(i => i[1]),
                backgroundColor: '#8b5cf6',
                borderRadius: 4,
                maxBarThickness: 30
            }]
        },
        options: {
            indexAxis: 'y',
            responsive: true,
            maintainAspectRatio: false,
            plugins: {
                legend: { display: false }
            },
            scales: {
                x: {
                    beginAtZero: true,
                    ticks: { precision: 0 }
                },
                y: {
                    grid: { display: false }
                }
            }
        }
    });
}

// 🔹 REPLACEMENT WORKFLOW FUNNEL
function updateReplacementFunnel() {
    const container = document.getElementById('replacementFunnel');

    // Count each stage
    const stages = [
        { key: 'replacement_confirmation', label: 'Customer Confirmation', icon: 'ri-user-follow-line' },
        { key: 'replacement_osg_approval', label: 'OSG Approval', icon: 'ri-checkbox-circle-line' },
        { key: 'replacement_mail_store', label: 'Mail to Store', icon: 'ri-mail-send-line' },
        { key: 'replacement_invoice_gen', label: 'Invoice Generated', icon: 'ri-file-list-2-line' },
        { key: 'replacement_invoice_sent', label: 'Invoice Sent to OSG', icon: 'ri-send-plane-line' },
        { key: 'replacement_settled_accounts', label: 'Settled with Accounts', icon: 'ri-money-dollar-circle-line' }
    ];

    let html = '<div class="funnel-stages">';

    // ONLY include claims with "Replacement approved" status
    const funnelClaims = filteredClaims.filter(c =>
        c.status === 'Replacement approved' || c.status === 'Replacement Approved'
    );

    stages.forEach((stage, index) => {
        const completed = funnelClaims.filter(c => c[stage.key] === true).length;
        const pending = funnelClaims.filter(c => c[stage.key] === false).length;
        const total = completed + pending;
        const percentage = total > 0 ? Math.round((completed / total) * 100) : 0;

        html += `
            <div class="funnel-stage">
                <div class="stage-header">
                    <i class="${stage.icon}"></i>
                    <span class="stage-label">${stage.label}</span>
                    <span class="stage-count">${completed}/${total}</span>
                </div>
                <div class="stage-bar">
                    <div class="stage-progress" style="width: ${percentage}%"></div>
                </div>
                <div class="stage-stats">
                    <span class="stat-completed">✓ ${completed} Completed</span>
                    <span class="stat-pending">⧗ ${pending} Pending</span>
                </div>
            </div>
        `;
    });

    html += '</div>';
    container.innerHTML = html;
}

// 🔹 RECENT CLAIMS LIST
function updateRecentClaims() {
    const container = document.getElementById('recentClaimsList');

    // Get last 10 claims sorted by date
    const recentClaims = [...filteredClaims]
        .sort((a, b) => new Date(b.submitted_date) - new Date(a.submitted_date))
        .slice(0, 10);

    if (recentClaims.length === 0) {
        container.innerHTML = '<div class="empty-state">No claims found</div>';
        return;
    }

    let html = '';
    recentClaims.forEach(claim => {
        const statusClass = (claim.status || '').toLowerCase().replace(/\s+/g, '-');
        html += `
            <div class="claim-item">
                <div class="claim-id-badge">${claim.claim_id}</div>
                <div class="claim-info">
                    <div class="claim-customer">
                        <strong>${claim.customer_name}</strong>
                        <span class="claim-mobile">${claim.mobile_number}</span>
                    </div>
                    <div class="claim-product">${claim.product || claim.model}</div>
                </div>
                <div class="claim-status">
                    <span class="status-pill ${statusClass}">${claim.status}</span>
                </div>
            </div>
        `;
    });

    container.innerHTML = html;
}

// 🔹 PENDING ACTIONS
function updatePendingActions() {
    // 1. Pending Follow-Up
    const pendingFollowUp = filteredClaims.filter(c =>
        c.status === 'Registered' || c.status === 'Follow Up'
    ).sort((a, b) => new Date(a.submitted_date) - new Date(b.submitted_date));

    document.getElementById('pendingFollowUpCount').textContent = pendingFollowUp.length;

    const followUpContainer = document.getElementById('pendingFollowUpList');
    if (pendingFollowUp.length === 0) {
        followUpContainer.innerHTML = '<div class="empty-state-sm">No pending follow-ups</div>';
    } else {
        let html = '';

        pendingFollowUp.forEach(claim => {
            const daysOld = calculateDaysOld(claim.submitted_date);
            const isOverdue = daysOld > 3;
            html += `
                <div class="action-item ${isOverdue ? 'overdue' : ''}" onclick="openAnalyticsModal('${claim.claim_id}')" style="cursor: pointer;">
                    <div class="action-info">
                        <strong>${claim.claim_id}</strong> - ${claim.customer_name}
                        <span class="action-date">${daysOld} days old</span>
                    </div>
                    ${isOverdue ? '<i class="ri-error-warning-line overdue-icon"></i>' : ''}
                </div>
            `;
        });
        followUpContainer.innerHTML = html;
    }

    // 2. Pending Replacement
    const pendingReplacement = filteredClaims.filter(c => {
        // Strict check: Only count if status is 'Replacement Approved' or 'Replacement approved'
        const s = (c.status || '').toLowerCase();
        const hasReplacementStatus = s.includes('replacement') && s.includes('approved');

        if (!hasReplacementStatus) return false;

        // Use the complete flag from backend which now includes mail_sent_to_store
        return !c.complete;
    });

    document.getElementById('pendingReplacementCount').textContent = pendingReplacement.length;

    const replacementContainer = document.getElementById('pendingReplacementList');
    if (pendingReplacement.length === 0) {
        replacementContainer.innerHTML = '<div class="empty-state-sm">No pending replacements</div>';
    } else {
        let html = '';

        pendingReplacement.forEach(claim => {
            const pendingStage = getPendingReplacementStage(claim);
            html += `
                <div class="action-item" onclick="openAnalyticsModal('${claim.claim_id}')" style="cursor: pointer;">
                    <div class="action-info">
                        <strong>${claim.claim_id}</strong> - ${claim.customer_name}
                        <span class="action-stage">Pending: ${pendingStage}</span>
                    </div>
                </div>
            `;
        });
        replacementContainer.innerHTML = html;
    }

    // Update total badge
    document.getElementById('pendingActionsBadge').textContent =
        pendingFollowUp.length + pendingReplacement.length;
}

// 🔹 FULL CLAIMS TABLE
function updateClaimsTable() {
    const tbody = document.getElementById('claimsTableBody');

    if (filteredClaims.length === 0) {
        tbody.innerHTML = '<tr><td colspan="10" class="empty-state">No claims match your filters</td></tr>';
        document.getElementById('showingCount').textContent = '0';
        document.getElementById('totalCount').textContent = allClaims.length;
        return;
    }

    let html = '';
    filteredClaims.forEach(claim => {
        const statusClass = (claim.status || '').toLowerCase().replace(/\s+/g, '-');
        const replacementProgress = calculateReplacementProgress(claim);

        html += `
            <tr class="data-row clickable-row" onclick="openAnalyticsModal('${claim.claim_id}')">
                <td><span class="id-badge">${claim.claim_id}</span></td>
                <td>${formatDate(claim.submitted_date)}</td>
                <td>${claim.customer_name}</td>
                <td>${claim.mobile_number}</td>
                <td>${claim.branch || '-'}</td>
                <td>${claim.product || claim.model}</td>
                <td class="issue-cell" title="${claim.issue}">${truncate(claim.issue, 40)}</td>
                <td><span class="status-pill ${statusClass}">${claim.status}</span></td>
                <td>
                    <div class="progress-bar-container">
                        <div class="progress-bar" style="width: ${replacementProgress}%"></div>
                        <span class="progress-text">${replacementProgress}%</span>
                    </div>
                </td>
                <td>
                    ${claim.complete
                ? '<i class="ri-checkbox-circle-fill complete-icon" style="color:var(--success)"></i>'
                : '<i class="ri-checkbox-blank-circle-line incomplete-icon" style="color:#d1d5db"></i>'}
                </td>
            </tr>
        `;
    });

    tbody.innerHTML = html;
    document.getElementById('showingCount').textContent = filteredClaims.length;
    document.getElementById('totalCount').textContent = allClaims.length;
}

// 🔹 FILTER LOGIC
function attachFilterListeners() {
    // Listen for Enter key on input fields
    const inputFields = ['filterClaimId', 'filterMobile'];

    inputFields.forEach(id => {
        const element = document.getElementById(id);
        if (element) {
            element.addEventListener('keypress', function (event) {
                if (event.key === 'Enter') {
                    applyFilters();
                }
            });
        }
    });

    // Note: No auto-apply on change - user must click "Apply Filters" button
}

function applyFilters() {
    console.log('=== APPLYING FILTERS ===');
    console.log('Total claims before filter:', allClaims.length);

    filteredClaims = allClaims.filter(claim => {
        // Date Range
        const dateStart = document.getElementById('filterDateFrom').value;
        const dateEnd = document.getElementById('filterDateTo').value;

        if (dateStart && claim.submitted_date) {
            if (claim.submitted_date < dateStart) {
                console.log(`Filtered out ${claim.claim_id}: Date too early`);
                return false;
            }
        }
        if (dateEnd && claim.submitted_date) {
            if (claim.submitted_date > dateEnd) {
                console.log(`Filtered out ${claim.claim_id}: Date too late`);
                return false;
            }
        }

        // Status
        const statusValue = document.getElementById('filterStatus').value;
        if (statusValue) {
            if (statusValue === 'Replacement Completed') {
                const s = (claim.status || '').toLowerCase();
                const isReplacement = s.includes('replacement');
                if (!(isReplacement && claim.complete)) return false;
            } else if (claim.status !== statusValue) {
                console.log(`Filtered out ${claim.claim_id}: Status mismatch (${claim.status} !== ${statusValue})`);
                return false;
            }
        }

        // Replacement Stage (ONLY applies to claims with status "Replacement approved")
        const replStage = document.getElementById('filterReplacementStage').value;
        if (replStage) {
            // First check: Must have "Replacement approved" status
            const hasReplacementStatus =
                claim.status === 'Replacement approved' ||
                claim.status === 'Replacement Approved';

            if (!hasReplacementStatus) {
                console.log(`Filtered out ${claim.claim_id}: Not a replacement claim (Status: ${claim.status})`);
                return false; // Exclude claims that aren't in replacement workflow
            }

            // Second check: Must have selected stage pending (not complete)
            if (claim[replStage] === true) {
                console.log(`Filtered out ${claim.claim_id}: Replacement stage already complete`);
                return false; // Hide if already done
            }
        }

        // Complete
        const completeFilter = document.getElementById('filterComplete').value;
        if (completeFilter === 'true' && !claim.complete) {
            console.log(`Filtered out ${claim.claim_id}: Not complete`);
            return false;
        }
        if (completeFilter === 'false' && claim.complete) {
            console.log(`Filtered out ${claim.claim_id}: Already complete`);
            return false;
        }

        // Claim ID - with null/undefined checks
        const claimIdSearch = document.getElementById('filterClaimId').value.trim().toLowerCase();
        if (claimIdSearch) {
            const claimId = String(claim.claim_id || '').toLowerCase();
            if (!claimId.includes(claimIdSearch)) {
                console.log(`Filtered out ${claim.claim_id}: Claim ID doesn't match search`);
                return false;
            }
        }

        // Mobile - with null/undefined checks and string conversion
        const mobileSearch = document.getElementById('filterMobile').value.trim();
        if (mobileSearch) {
            const claimMobile = String(claim.mobile_number || '').trim();
            console.log(`Checking mobile: "${claimMobile}" includes "${mobileSearch}"?`);
            if (!claimMobile.includes(mobileSearch)) {
                console.log(`Filtered out ${claim.claim_id}: Mobile number doesn't match (${claimMobile} doesn't include ${mobileSearch})`);
                return false;
            }
        }

        return true;
    });

    console.log('Total claims after filter:', filteredClaims.length);
    console.log('Filtered claims:', filteredClaims.map(c => c.claim_id));
    console.log('======================');

    updateAllComponents();

    // Auto-open modal if searching by Claim ID and we found exactly one match
    const searchId = document.getElementById('filterClaimId').value.trim();
    if (searchId && filteredClaims.length === 1) {
        // Optional: Check if the found claim actually contains the search term
        // (It should, because of the filter logic, but good double check)
        openAnalyticsModal(filteredClaims[0].claim_id);
    }
}

function clearAllFilters() {
    document.getElementById('filterDateFrom').value = '';
    document.getElementById('filterDateTo').value = '';
    document.getElementById('filterStatus').selectedIndex = 0;
    document.getElementById('filterReplacementStage').selectedIndex = 0;
    document.getElementById('filterComplete').selectedIndex = 0;
    document.getElementById('filterClaimId').value = '';
    document.getElementById('filterMobile').value = '';

    // Clear active card highlight
    clearActiveCard();

    filteredClaims = [...allClaims];
    updateAllComponents();
}

// 🔹 KPI CARD FILTER
function filterByCard(cardType) {
    // Highlight active card
    clearActiveCard();
    const cardEl = document.getElementById('card-' + cardType);
    if (cardEl) cardEl.classList.add('card-active');

    // Show label in table header
    const labelEl = document.getElementById('activeCardLabel');
    const labelMap = {
        'total': '— All Claims',
        'open': '— Open Claims',
        'replacement': '— Replacement Pending',
        'repair-completed': '— Repair Completed',
        'settled': '— Replacement Completed',
        'closed': '— Closed',
        'today': "— Today's New",
        'month': '— This Month',
        'rejected': '— Rejected',
        'no-issue': '— No Issue / On Call Resolution'
    };
    if (labelEl) {
        labelEl.textContent = labelMap[cardType] || '';
        labelEl.style.display = 'inline';
    }

    const today = new Date().toISOString().split('T')[0];
    const now = new Date();
    const monthStart = new Date(now.getFullYear(), now.getMonth(), 1).toISOString().split('T')[0];

    switch (cardType) {
        case 'total':
            filteredClaims = [...allClaims];
            break;
        case 'open':
            filteredClaims = allClaims.filter(c => {
                const s = (c.status || '').toLowerCase();
                if (s === 'settled' || s === 'closed' || s === 'repair completed') return false;
                if (s.includes('replacement') && s.includes('approved')) return !c.complete;
                return true;
            });
            break;
        case 'replacement':
            filteredClaims = allClaims.filter(c => {
                const hasReplacementStatus =
                    c.status === 'Replacement Approved' ||
                    c.status === 'Replacement approved';
                if (!hasReplacementStatus) return false;
                return !c.replacement_confirmation ||
                    !c.replacement_osg_approval ||
                    !c.replacement_mail_store ||
                    !c.replacement_invoice_gen ||
                    !c.replacement_invoice_sent ||
                    !c.replacement_settled_accounts;
            });
            break;
        case 'repair-completed':
            filteredClaims = allClaims.filter(c => c.status === 'Repair Completed');
            break;
        case 'settled':
            filteredClaims = allClaims.filter(c => {
                const s = (c.status || '').toLowerCase();
                return c.complete && s.includes('replacement') && s.includes('approved');
            });
            break;
        case 'closed':
            filteredClaims = allClaims.filter(c => (c.status || '').toLowerCase() === 'closed');
            break;
        case 'today':
            filteredClaims = allClaims.filter(c =>
                c.submitted_date && c.submitted_date.startsWith(today)
            );
            break;
        case 'month':
            filteredClaims = allClaims.filter(c =>
                c.submitted_date && c.submitted_date >= monthStart
            );
            break;
        case 'rejected':
            filteredClaims = allClaims.filter(c => (c.status || '').toLowerCase() === 'rejected');
            break;
        case 'no-issue':
            filteredClaims = allClaims.filter(c => {
                const s = (c.status || '').toLowerCase();
                return s === 'no issue' || s === 'on call resolution' || s === 'no issue/on call resolution' || s === 'no issue / on call resolution';
            });
            break;
        default:
            filteredClaims = [...allClaims];
    }

    // Update the claims table, KPIs, and counts
    updateClaimsTable();
    updateKPIs();

    // Scroll to All Claims table
    const tableSection = document.getElementById('claimsAnalyticsTable');
    if (tableSection) {
        tableSection.closest('.table-section').scrollIntoView({ behavior: 'smooth', block: 'start' });
    }
}

function clearActiveCard() {
    document.querySelectorAll('.kpi-card.card-active').forEach(el => el.classList.remove('card-active'));
    const labelEl = document.getElementById('activeCardLabel');
    if (labelEl) labelEl.style.display = 'none';
}

// 🔹 UTILITY FUNCTIONS
function calculateReplacementProgress(claim) {
    const stages = [
        claim.replacement_confirmation,
        claim.replacement_osg_approval,
        claim.replacement_mail_store,
        claim.replacement_invoice_gen,
        claim.replacement_invoice_sent,
        claim.replacement_settled_accounts
    ];

    const completed = stages.filter(s => s === true).length;
    return Math.round((completed / stages.length) * 100);
}

function getPendingReplacementStage(claim) {
    if (!claim.replacement_confirmation) return 'Customer Confirmation';
    if (!claim.replacement_osg_approval) return 'OSG Approval';
    if (!claim.replacement_mail_store) return 'Mail to Store';
    if (!claim.replacement_invoice_gen) return 'Invoice Generated';
    if (!claim.replacement_invoice_sent) return 'Invoice Sent';
    if (!claim.replacement_settled_accounts) return 'Settled with Accounts';
    return 'Unknown';
}

function calculateDaysOld(dateStr) {
    if (!dateStr) return 0;
    const submitted = new Date(dateStr);
    const now = new Date();
    const diff = now - submitted;
    return Math.floor(diff / (1000 * 60 * 60 * 24));
}

function formatDate(dateStr) {
    if (!dateStr) return '-';
    const date = new Date(dateStr);
    return date.toLocaleDateString('en-IN', {
        year: 'numeric',
        month: 'short',
        day: 'numeric'
    });
}

function truncate(str, length) {
    if (!str) return '-';
    return str.length > length ? str.substring(0, length) + '...' : str;
}

function debounce(func, wait) {
    let timeout;
    return function executedFunction(...args) {
        const later = () => {
            clearTimeout(timeout);
            func(...args);
        };
        clearTimeout(timeout);
        timeout = setTimeout(later, wait);
    };
}

function refreshDashboard() {
    loadAnalyticsData();
}

async function exportTableData() {
    // Build list of claim IDs to export
    const claimIds = filteredClaims.map(c => c.claim_id);

    try {
        const btn = document.querySelector('.btn-export');
        if (btn) { btn.disabled = true; btn.innerHTML = '<i class="ri-loader-4-line"></i> Exporting...'; }

        const response = await fetch('/api/export-claims-excel', {
            method: 'POST',
            headers: { 'Content-Type': 'application/json' },
            body: JSON.stringify({ claim_ids: claimIds })
        });

        if (!response.ok) throw new Error('Export failed');

        const blob = await response.blob();
        const url = URL.createObjectURL(blob);
        const link = document.createElement('a');
        link.href = url;
        link.download = 'OSG_Claims_Export_' + new Date().toISOString().slice(0, 10) + '.xlsx';
        document.body.appendChild(link);
        link.click();
        document.body.removeChild(link);
        URL.revokeObjectURL(url);
    } catch (err) {
        console.error('Export error:', err);
        alert('Export failed. Please try again.');
    } finally {
        const btn = document.querySelector('.btn-export');
        if (btn) { btn.disabled = false; btn.innerHTML = '<i class="ri-download-line"></i> Export'; }
    }
}

// Navigation helper - NOW OPENS LOCAL MODAL
function openAnalyticsModal(claimId) {
    const claim = allClaims.find(c => c.claim_id === claimId);
    if (!claim) return;

    // ... (rest of local modal logic) ...
    // Since I can't easily see the end of this function without scrolling more, I will insert the NEW functions 
    // for the branch modal clearly at the end of the file or after filteredClaims usage.
    // Actually, I should just append them at the end. I will use a different chunk for that.


    // Populate Modal Fields
    document.getElementById('m_claim_id').textContent = claim.claim_id;
    document.getElementById('m_status').textContent = claim.status;
    document.getElementById('m_customer').textContent = claim.customer_name;
    document.getElementById('m_mobile').textContent = claim.mobile_number;
    document.getElementById('m_product').textContent = claim.product || claim.model;
    document.getElementById('m_date').textContent = formatDate(claim.submitted_date);
    document.getElementById('m_issue').textContent = claim.issue || 'No issue description';
    document.getElementById('m_osid').textContent = claim.osid || '--';
    document.getElementById('m_invoice').textContent = claim.invoice_number || '--';
    document.getElementById('m_sr_no').textContent = claim.sr_no || '--';

    // Populate Workflow Steps
    const stepsContainer = document.getElementById('m_workflow_steps');
    let stepsHtml = '';

    const steps = [
        { key: 'replacement_confirmation', label: 'Customer Confirmation' },
        { key: 'replacement_osg_approval', label: 'OSG Approval' },
        { key: 'replacement_mail_store', label: 'Mail to Store' },
        { key: 'replacement_invoice_gen', label: 'Invoice Generated' },
        { key: 'replacement_invoice_sent', label: 'Invoice Sent to OSG' },
        { key: 'replacement_settled_accounts', label: 'Settled with Accounts' }
    ];

    if (claim.status === 'Replacement approved' || claim.status === 'Replacement Approved') {
        steps.forEach(step => {
            const isDone = claim[step.key] === true;
            const icon = isDone ? '<i class="ri-checkbox-circle-fill" style="color: green"></i>' : '<i class="ri-checkbox-blank-circle-line" style="color: #ccc"></i>';
            stepsHtml += `
                <div class="step-item">
                    <div class="step-icon">${icon}</div>
                    <div class="step-label">${step.label}</div>
                    <div class="step-status">${isDone ? 'Completed' : 'Pending'}</div>
                </div>
            `;
        });
    } else if (claim.status === 'Repair Completed') {
        stepsHtml = `
            <div class="step-item">
                <div class="step-icon"><i class="ri-tools-line" style="color: var(--primary)"></i></div>
                <div class="step-label">Repair Completed</div>
                <div class="step-status">Done</div>
            </div>
         `;
    } else if (claim.status === 'Follow Up') {
        stepsHtml = `
            <div class="step-item" style="flex-direction: column; align-items: flex-start; gap: 0.5rem;">
                <div style="display: flex; align-items: center; gap: 0.5rem; color: var(--primary); font-weight: 600;">
                    <i class="ri-history-line"></i> Follow Up History
                </div>
                <div style="width: 100%; max-height: 200px; overflow-y: auto; background: rgba(0,0,0,0.03); padding: 0.8rem; border-radius: 8px; white-space: pre-wrap; font-size: 0.9rem; border: 1px solid rgba(0,0,0,0.05);">
                    ${claim.follow_up_notes || 'No history available'}
                </div>
            </div>
        `;
    } else {
        stepsHtml = `<div style="color: #666; font-style: italic;">No active workflow for status: ${claim.status}</div>`;
    }

    // Add completion status (Removed per user request)

    stepsContainer.innerHTML = stepsHtml;

    // Show Modal
    document.getElementById('analytics-modal').classList.remove('hidden');
}

// 🔹 BRANCH MODAL LOGIC
function openBranchModal() {
    const modal = document.getElementById('branchModal');
    if (!modal) return;
    modal.classList.remove('hidden');
    modal.classList.add('flex'); // Ensure flex display

    // Render Large Chart
    setTimeout(() => {
        const ctx = document.getElementById('branchChartLarge');
        if (!ctx) return;

        // Recalculate counts for ALL branches
        const branchCounts = {};
        filteredClaims.forEach(c => {
            const b = c.branch || 'Unknown';
            branchCounts[b] = (branchCounts[b] || 0) + 1;
        });

        const sorted = Object.entries(branchCounts)
            .sort((a, b) => b[1] - a[1]); // All branches

        if (largeBranchChart) largeBranchChart.destroy();

        largeBranchChart = new Chart(ctx, {
            type: 'bar',
            data: {
                labels: sorted.map(i => i[0]),
                datasets: [{
                    label: 'Claims',
                    data: sorted.map(i => i[1]),
                    backgroundColor: '#8b5cf6',
                    borderRadius: 4,
                    maxBarThickness: 40
                }]
            },
            options: {
                indexAxis: 'y', // Horizontal Layout
                responsive: true,
                maintainAspectRatio: false,
                plugins: {
                    legend: { display: false },
                    title: {
                        display: true,
                        text: 'Branch Performance (All Branches)',
                        font: { size: 16 }
                    }
                },
                scales: {
                    x: {
                        beginAtZero: true,
                        ticks: { precision: 0 }
                    },
                    y: {
                        grid: { display: false },
                        ticks: { autoSkip: false } // Show all labels in large view
                    }
                }
            }
        });
    }, 100);
}

function closeBranchModal() {
    const modal = document.getElementById('branchModal');
    if (modal) {
        modal.classList.add('hidden');
        modal.classList.remove('flex');
    }
}

function closeAnalyticsModal() {
    document.getElementById('analytics-modal').classList.add('hidden');
}

// Restore table click to open LOCAL modal
// Modification to updateClaimsTable is needed above to use openAnalyticsModal instead of redirect logic.

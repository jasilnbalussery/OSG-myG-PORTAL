/**
 * OSG myG Portal - Optimized Frontend v2.0
 * Clean, Fast, Reliable JavaScript
 * 
 * Features:
 * - State management
 * - Debounced API calls
 * - Error handling
 * - Loading states
 * - Cache management
 */

// =============================================================================
// STATE MANAGEMENT
// =============================================================================

const AppState = {
    claims: [],
    filteredClaims: [],
    currentClaim: null,
    loading: false,
    error: null,
    lastFetch: 0,
    CACHE_DURATION: 30000, // 30 seconds

    setClaims(claims) {
        this.claims = claims;
        this.filteredClaims = claims;
        this.lastFetch = Date.now();
    },

    setLoading(state) {
        this.loading = state;
        this.updateUI();
    },

    setError(error) {
        this.error = error;
        this.showError(error);
    },

    isCacheValid() {
        return (Date.now() - this.lastFetch) < this.CACHE_DURATION;
    },

    updateUI() {
        const loader = document.getElementById('loader');
        if (loader) {
            loader.style.display = this.loading ? 'flex' : 'none';
        }
    },

    showError(message) {
        console.error(message);
        // Could add toast notification here
    }
};

// =============================================================================
// API CLIENT
// =============================================================================

const API = {
    baseURL: window.location.origin,

    async fetchClaims(forceRefresh = false) {
        if (!forceRefresh && AppState.isCacheValid()) {
            console.log('[API] Using cached claims');
            return AppState.claims;
        }

        AppState.setLoading(true);

        try {
            const controller = new AbortController();
            const timeoutId = setTimeout(() => controller.abort(), 30000);
            const response = await fetch(`${this.baseURL}/api/claims`, { signal: controller.signal });
            clearTimeout(timeoutId);
            const data = await response.json();

            if (data.success) {
                AppState.setClaims(data.claims);
                console.log(`[API] Fetched ${data.claims.length} claims`);
                return data.claims;
            } else {
                throw new Error(data.error || 'Failed to fetch claims');
            }
        } catch (error) {
            AppState.setError(`Failed to load claims: ${error.message}`);
            return [];
        } finally {
            AppState.setLoading(false);
        }
    },

    async fetchClaimDetail(claimId) {
        try {
            const response = await fetch(`${this.baseURL}/api/claim/${claimId}`);
            const data = await response.json();

            if (data.success) {
                return data.claim;
            } else {
                throw new Error(data.error || 'Claim not found');
            }
        } catch (error) {
            AppState.setError(`Failed to load claim: ${error.message}`);
            return null;
        }
    },

    async updateClaim(claimId, updates) {
        try {
            const response = await fetch(`${this.baseURL}/api/claim/${claimId}/update`, {
                method: 'POST',
                headers: { 'Content-Type': 'application/json' },
                body: JSON.stringify(updates)
            });

            const data = await response.json();

            if (data.success) {
                console.log('[API] Claim updated successfully');
                // Invalidate cache to fetch fresh data
                AppState.lastFetch = 0;
                return true;
            } else {
                throw new Error(data.error || 'Update failed');
            }
        } catch (error) {
            AppState.setError(`Failed to update claim: ${error.message}`);
            return false;
        }
    },

    async submitClaim(claimData) {
        try {
            const response = await fetch(`${this.baseURL}/api/claim/submit`, {
                method: 'POST',
                headers: { 'Content-Type': 'application/json' },
                body: JSON.stringify(claimData)
            });

            const data = await response.json();

            if (data.success) {
                console.log(`[API] Claim submitted: ${data.claim_id}`);
                return data;
            } else {
                throw new Error(data.error || 'Submission failed');
            }
        } catch (error) {
            AppState.setError(`Failed to submit claim: ${error.message}`);
            return null;
        }
    },

    async clearCache() {
        try {
            await fetch(`${this.baseURL}/api/cache/clear`, { method: 'POST' });
            AppState.lastFetch = 0;
            console.log('[API] Cache cleared');
        } catch (error) {
            console.error('[API] Cache clear failed:', error);
        }
    }
};

// =============================================================================
// MODAL MANAGEMENT
// =============================================================================

const Modal = {
    overlay: null,
    content: null,
    currentClaimId: null,

    init() {
        this.overlay = document.getElementById('claimModalOverlay');
        this.content = document.getElementById('claimDetailContent');

        if (this.overlay) {
            this.overlay.addEventListener('click', (e) => {
                if (e.target === this.overlay) {
                    this.close();
                }
            });
        }
    },

    async open(claimId) {
        this.currentClaimId = claimId;

        // Show modal
        if (this.overlay) {
            this.overlay.classList.remove('hidden');
        }

        // Load claim data
        const claim = await API.fetchClaimDetail(claimId);
        if (claim) {
            this.populate(claim);
        }
    },

    populate(claim) {
        // Populate modal fields with claim data
        this.setValue('modalClaimId', claim.claim_id);
        this.setValue('modalCustomerName', claim.customer_name);
        this.setValue('modalMobile', claim.mobile_number);
        this.setValue('modalProduct', claim.product);
        this.setValue('modalIssue', claim.issue);
        this.setValue('updateStatus', claim.status);
        this.setValue('followUpDate', claim.follow_up_date);
        this.setValue('followUpNotes', claim.follow_up_notes);
        this.setValue('assignedStaff', claim.assigned_staff);

        // Set checkboxes
        this.setChecked('chk_repl_confirmation', claim.replacement_confirmation);
        this.setChecked('chk_repl_osg_approval', claim.replacement_osg_approval);
        this.setChecked('chk_repl_mail_store', claim.replacement_mail_store);
        this.setChecked('chk_repl_invoice_gen', claim.replacement_invoice_gen);
        this.setChecked('chk_repl_invoice_sent', claim.replacement_invoice_sent);
        this.setChecked('chk_repl_settled_accounts', claim.replacement_settled_accounts);
        this.setChecked('chk_complete_repl', claim.complete);

        if (claim.tat !== null && claim.tat !== undefined) {
            this.setValue('disp_tat', claim.tat);
        }

        // Initialize Enhanced Workflow UI
        if (typeof window.initializeWorkflow === 'function') {
            window.initializeWorkflow(claim);
        }
    },

    setValue(id, value) {
        const element = document.getElementById(id);
        if (element) {
            if (element.tagName === 'INPUT' || element.tagName === 'SELECT' || element.tagName === 'TEXTAREA') {
                element.value = value || '';
            } else {
                element.textContent = value || '';
            }
        }
    },

    setChecked(id, checked) {
        const element = document.getElementById(id);
        if (element && element.type === 'checkbox') {
            element.checked = !!checked;
        }
    },

    getChecked(id) {
        const element = document.getElementById(id);
        return element && element.type === 'checkbox' ? element.checked : false;
    },

    async save() {
        if (!this.currentClaimId) return;

        // Collect updates
        const updates = {
            status: document.getElementById('updateStatus')?.value,
            follow_up_date: document.getElementById('followUpDate')?.value,
            follow_up_notes: document.getElementById('followUpNotes')?.value,
            assigned_staff: document.getElementById('assignedStaff')?.value,
            replacement_confirmation: this.getChecked('chk_repl_confirmation'),
            replacement_osg_approval: this.getChecked('chk_repl_osg_approval'),
            replacement_mail_store: this.getChecked('chk_repl_mail_store'),
            replacement_invoice_gen: this.getChecked('chk_repl_invoice_gen'),
            replacement_invoice_sent: this.getChecked('chk_repl_invoice_sent'),
            replacement_settled_accounts: this.getChecked('chk_repl_settled_accounts'),
            complete: this.getChecked('chk_complete_repl')
        };

        // Save via API
        const success = await API.updateClaim(this.currentClaimId, updates);

        if (success) {
            alert('Changes saved successfully!');
            this.close();
            Dashboard.refresh();
        } else {
            alert('Failed to save changes. Please try again.');
        }
    },

    close() {
        if (this.overlay) {
            this.overlay.classList.add('hidden');
        }
        this.currentClaimId = null;
    }
};

// =============================================================================
// DASHBOARD
// =============================================================================

const Dashboard = {
    tableBody: null,
    searchInput: null,

    init() {
        this.tableBody = document.getElementById('claimsTableBody');
        this.searchInput = document.getElementById('searchInput');

        // Setup event listeners
        if (this.searchInput) {
            this.searchInput.addEventListener('input', debounce(() => this.filter(), 300));
        }

        // Load initial data
        this.load();
    },

    async load() {
        const claims = await API.fetchClaims();
        this.render(claims);
        this.updateStats(claims);
    },

    async refresh() {
        const claims = await API.fetchClaims(true);
        this.render(claims);
        this.updateStats(claims);
    },

    filter() {
        const searchTerm = this.searchInput?.value.toLowerCase() || '';

        const filtered = AppState.claims.filter(claim => {
            const searchableText = `
                ${claim.claim_id}
                ${claim.customer_name}
                ${claim.mobile_number}
                ${claim.product}
                ${claim.status}
            `.toLowerCase();

            return searchableText.includes(searchTerm);
        });

        this.render(filtered);
    },

    render(claims) {
        if (!this.tableBody) return;

        if (claims.length === 0) {
            this.tableBody.innerHTML = '<tr><td colspan="6" style="text-align: center; padding: 2rem;">No claims found</td></tr>';
            return;
        }

        this.tableBody.innerHTML = claims.map(claim => `
            <tr onclick="Modal.open('${claim.claim_id}')" style="cursor: pointer;">
                <td><span class="id-badge">${claim.claim_id}</span></td>
                <td>${this.formatDate(claim.submitted_date)}</td>
                <td>${claim.customer_name}</td>
                <td>${claim.mobile_number}</td>
                <td>${claim.product}</td>
                <td><span class="status-pill ${this.getStatusClass(claim.status)}">${claim.status}</span></td>
            </tr>
        `).join('');
    },

    updateStats(claims) {
        // Total claims
        this.setStatValue('total-claims', claims.length);

        // Open claims
        const openClaims = claims.filter(c => !c.complete && c.status !== 'Closed').length;
        this.setStatValue('open-claims', openClaims);

        // This month
        const thisMonth = claims.filter(c => {
            if (!c.submitted_date) return false;
            const claimDate = new Date(c.submitted_date);
            const now = new Date();
            return claimDate.getFullYear() === now.getFullYear() &&
                claimDate.getMonth() === now.getMonth();
        }).length;
        this.setStatValue('this-month', thisMonth);

        // Settled
        const settled = claims.filter(c => c.complete).length;
        this.setStatValue('settled', settled);
    },

    setStatValue(id, value) {
        const element = document.getElementById(id);
        if (element) {
            element.textContent = value;
        }
    },

    formatDate(dateStr) {
        if (!dateStr) return '';
        try {
            const date = new Date(dateStr);
            return date.toLocaleDateString('en-IN', {
                year: 'numeric',
                month: 'short',
                day: 'numeric'
            });
        } catch {
            return dateStr;
        }
    },

    getStatusClass(status) {
        const statusMap = {
            'registered': 'registered',
            'follow up': 'follow-up',
            'replacement approved': 'completed',
            'settled': 'completed',
            'closed': 'closed'
        };
        return statusMap[status?.toLowerCase()] || 'registered';
    }
};

// =============================================================================
// UTILITIES
// =============================================================================

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

// =============================================================================
// INITIALIZATION
// =============================================================================

document.addEventListener('DOMContentLoaded', () => {
    console.log('[App] OSG myG Portal v2.0 initialized');

    // Initialize components
    Modal.init();
    Dashboard.init();

    // Setup save button
    const saveBtn = document.getElementById('saveClaimChanges');
    if (saveBtn) {
        saveBtn.addEventListener('click', () => Modal.save());
    }

    // Setup close button
    const closeBtn = document.getElementById('closeModal');
    if (closeBtn) {
        closeBtn.addEventListener('click', () => Modal.close());
    }

    // Setup refresh button
    const refreshBtn = document.getElementById('refreshBtn');
    if (refreshBtn) {
        refreshBtn.addEventListener('click', () => Dashboard.refresh());
    }
});

// Export for debugging
window.AppState = AppState;
window.API = API;
window.Modal = Modal;
window.Dashboard = Dashboard;

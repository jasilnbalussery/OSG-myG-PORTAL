// --- Global ---
function closeModal() {
    document.getElementById('modal-container').classList.add('hidden');
}

// Format Date like "21 Dec 2025"
function formatDate(dateStr) {
    if (!dateStr) return "--";
    const d = new Date(dateStr);
    if (isNaN(d.getTime())) return dateStr;
    return d.toLocaleDateString('en-GB', { day: 'numeric', month: 'short', year: 'numeric' });
}

function renderWorkflowList(data) {
    const list = document.getElementById('workflowStatusList');
    if (!list) return;
    list.innerHTML = '';

    // Status normalization
    const status = (data.status || '').toLowerCase().trim();

    const steps = [];

    if (status.includes('replacement')) {
        steps.push({ label: "Customer Confirmation", key: "replacement_confirmation" });
        steps.push({ label: "Onsitego Approval", key: "replacement_osg_approval" });
        steps.push({ label: "Mail Sent to Store", key: "replacement_mail_store" });
        steps.push({ label: "Invoice Generated", key: "replacement_invoice_gen" });
        steps.push({ label: "Invoice Sent", key: "replacement_invoice_sent" });
        steps.push({ label: "Settlement Mail to Accounts", key: "settlement_mail_accounts" });
        steps.push({ label: "Settled with Accounts", key: "replacement_settled_accounts" });

    } else if (status === 'follow up') {
        // Follow Up specific step
        // Consider it "Completed" if there are any notes in history, otherwise "Pending"
        const hasNotes = data.follow_up_notes && data.follow_up_notes.length > 0;
        steps.push({ label: "Customer Follow Up", customDone: hasNotes });

    } else if (status === 'closed' || status === 'rejected' || status.includes('no issue')) {
        // Resolved status step
        steps.push({ label: "Claim Resolved", customDone: true });

    } else if (status === 'submitted') {
        steps.push({ label: "Claim Submitted", customDone: true });

    } else if (status === 'registered') {
        steps.push({ label: "Claim Registered", customDone: true });

    } else {
        // Standard / Repair steps logic (Registered, Submitted, Repair Completed)
        steps.push({ label: "Repair Feedback", key: "repair_feedback_completed" });

        // If status is specifically Repair Completed, maybe add a completion step?
        // sticking to requested "Repair Feedback" for now as the core repair metric.
    }

    steps.forEach(step => {
        // Use customDone if present, otherwise check the data key
        let isDone = false;
        if (step.customDone !== undefined) {
            isDone = step.customDone;
        } else {
            isDone = data[step.key] === true;
        }

        const statusText = isDone ? "Completed" : "Pending";
        const iconClass = isDone ? "ri-check-line" : "ri-time-line";
        const stateClass = isDone ? "completed" : "pending";

        const html = `
            <div class="workflow-item">
                <div class="workflow-info">
                    <div class="workflow-icon ${stateClass}">
                        <i class="${iconClass}"></i>
                    </div>
                    <span class="workflow-text">${step.label}</span>
                </div>
                <span class="workflow-state ${stateClass}">${statusText}</span>
            </div>
        `;
        list.innerHTML += html;
    });
}

function toggleFollowUpHistory() {
    const el = document.getElementById('followUpHistory');
    if (el) el.classList.toggle('hidden');
}

function getISTDate() {
    const d = new Date();
    // distinct handling for time if needed, but input type="date" needs YYYY-MM-DD
    // IST is UTC+5:30
    const utc = d.getTime() + (d.getTimezoneOffset() * 60000);
    const titleDate = new Date(utc + (3600000 * 5.5));
    return titleDate.toISOString().split('T')[0];
}

// --- Submit Claim Page ---
// Auto-lookup when mobile hits 10 digits
(function setupMobileAutoLookup() {
    let debounceTimer = null;
    document.addEventListener('DOMContentLoaded', function () {
        const mobileInput = document.getElementById('mobileInput');
        if (!mobileInput) return;

        // Fire search on input once we have 10 digits
        mobileInput.addEventListener('input', function () {
            clearTimeout(debounceTimer);
            const val = mobileInput.value.replace(/\D/g, ''); // digits only
            mobileInput.value = val; // sanitise non-numeric characters
            if (val.length === 10) {
                debounceTimer = setTimeout(() => searchCustomer(), 600);
            }
        });

        // Also trigger on Enter key
        mobileInput.addEventListener('keydown', function (e) {
            if (e.key === 'Enter') {
                e.preventDefault();
                clearTimeout(debounceTimer);
                searchCustomer();
            }
        });
    });
})();

// --- Utility ---
async function fetchWithTimeout(resource, options = {}) {
    const { timeout = 30000 } = options; // 30s max wait for Google Sheets
    const controller = new AbortController();
    const id = setTimeout(() => controller.abort(), timeout);
    try {
        const response = await fetch(resource, {
            ...options,
            signal: controller.signal
        });
        clearTimeout(id);
        return response;
    } catch (error) {
        clearTimeout(id);
        throw error;
    }
}

// --- Submit Claim Page ---
// Global selection state
window.selectedProducts = [];

async function searchCustomer() {
    const mobileInput = document.getElementById('mobileInput');
    const searchBtn = document.getElementById('searchBtn');
    const mobile = mobileInput.value;
    const msgBox = document.getElementById('search-msg');
    const formSection = document.getElementById('claim-form-section');

    if (!mobile || mobile.length !== 10) {
        msgBox.textContent = "Please enter a valid 10-digit number.";
        msgBox.className = "msg-box warning";
        msgBox.classList.remove('hidden');
        return;
    }

    // Show loading & Disable inputs
    msgBox.textContent = "Searching Database...";
    msgBox.className = "msg-box info";
    msgBox.classList.remove('hidden');

    // UI Feedback
    mobileInput.disabled = true;
    searchBtn.disabled = true;
    const originalBtnText = searchBtn.innerHTML;
    searchBtn.innerHTML = '<i class="ri-loader-4-line spin"></i> Searching...';

    // Reset selection
    window.selectedProducts = [];
    renderClaimForms();

    try {
        const response = await fetchWithTimeout('/lookup-customer', {
            method: 'POST',
            headers: { 'Content-Type': 'application/json' },
            body: JSON.stringify({ mobile: mobile }),
            timeout: 15000 // 15s timeout for fast lookup
        });

        const data = await response.json();

        if (data.loading) {
            msgBox.textContent = "⏳ " + (data.message || "Server is updating data. Please try again in 10 seconds.");
            msgBox.className = "msg-box warning";
            msgBox.classList.remove('hidden');
            mobileInput.disabled = false;
            searchBtn.disabled = false;
            searchBtn.innerHTML = originalBtnText;
            return;
        }

        msgBox.classList.add('hidden');
        formSection.classList.remove('hidden');

        if (data.success) {
            // Found existing customer
            document.getElementById('disp-name').textContent = data.customer_name;
            document.getElementById('disp-mobile').textContent = mobile;
            document.getElementById('hidden_name').value = data.customer_name;
            document.getElementById('hidden_mobile').value = mobile;

            // Populate Products
            const productList = document.getElementById('product-list');
            productList.innerHTML = "";

            if (data.products && data.products.length > 0) {
                data.products.forEach((prod, index) => {
                    const el = document.createElement('div');
                    el.className = 'product-card';
                    el.dataset.index = index; // Store index for reference
                    el.innerHTML = `
                            <h4>${prod.model}</h4>
                            <p style="font-size:0.8rem; color:#777">Serial: ${prod.serial}</p>
                            <p style="font-size:0.8rem; color:#777">OSID: ${prod.osid}</p>
                            <p style="font-size:0.8rem; color:#777">Inv: ${prod.invoice}</p>
                        `;
                    el.onclick = () => toggleProductSelection(prod, el);
                    productList.appendChild(el);
                });
            } else {
                productList.innerHTML = '<p style="color:#777; font-style:italic;">No registered products found. Register a new one below.</p>';
            }

        } else {
            // New Customer
            msgBox.textContent = "New Customer! Please enter details.";
            msgBox.className = "msg-box info";
            msgBox.classList.remove('hidden');

            document.getElementById('disp-name').innerHTML = `<input type="text" id="manual_name" placeholder="Enter Name" class="manual-input" onchange="document.getElementById('hidden_name').value = this.value">`;
            document.getElementById('disp-mobile').textContent = mobile;
            document.getElementById('hidden_name').value = "";
            document.getElementById('hidden_mobile').value = mobile;
            const manualProd = { model: '', serial: '', invoice: '', osid: 'MANUAL_ENTRY', isManual: true };
            window.selectedProducts = [manualProd];

            document.getElementById('product-list').innerHTML = `
                    <div class="manual-product-form" style="padding:1rem; background:rgba(67, 97, 238, 0.05); border-radius:12px; border:1px solid rgba(67, 97, 238, 0.2);">
                        <h4 style="margin-bottom:0.8rem; color:var(--primary);">New Product Details</h4>
                        <div class="input-group" style="display:flex; gap:0.5rem; margin-bottom:0.5rem">
                             <input type="text" class="manual-input" style="border:1px solid #ddd; padding:0.5rem; border-radius:8px;" placeholder="Model Name" onchange="updateManualProduct(this, 'model')">
                             <input type="text" class="manual-input" style="border:1px solid #ddd; padding:0.5rem; border-radius:8px;" placeholder="Serial No" onchange="updateManualProduct(this, 'serial')">
                        </div>
                        <div class="input-group">
                             <input type="text" class="manual-input" style="border:1px solid #ddd; padding:0.5rem; border-radius:8px; width:100%;" placeholder="Invoice No" onchange="updateManualProduct(this, 'invoice')">
                        </div>
                    </div>
                `;

            window.updateManualProduct = function (el, field) {
                window.selectedProducts[0][field] = el.value;
                renderClaimForms();
            };

            renderClaimForms();
        }
    } catch (error) {
        msgBox.textContent = "Connection Error. Please try again.";
        msgBox.className = "msg-box error";
        msgBox.classList.remove('hidden');
        console.error("Lookup error:", error);
    } finally {
        mobileInput.disabled = false;
        searchBtn.disabled = false;
        searchBtn.innerHTML = originalBtnText;
    }
}
// Toggle Selection
function toggleProductSelection(product, element) {
    const index = window.selectedProducts.findIndex(p => p.osid === product.osid && p.serial === product.serial);

    if (index === -1) {
        // Add
        window.selectedProducts.push(product);
        element.classList.add('selected');
    } else {
        // Remove
        window.selectedProducts.splice(index, 1);
        element.classList.remove('selected');
    }
    renderClaimForms();
}

// Render Dynamic Forms
function renderClaimForms() {
    const container = document.getElementById('dynamic-claim-forms');

    if (window.selectedProducts.length === 0) {
        container.innerHTML = `
            <div class="empty-selection-msg" style="text-align: center; color: var(--text-tertiary); padding: 2rem; border: 2px dashed #e5e7eb; border-radius: 12px;">
                <i class="ri-shopping-cart-line" style="font-size: 2rem; margin-bottom: 0.5rem; display: block;"></i>
                Please select products above to add claim details
            </div>`;
        return;
    }

    // Preserve existing input values if re-rendering? 
    // For simplicity, we might lose values if toggling. 
    // To fix, we should store values in the product object or a separate map.
    // Basic implementation: Just re-render. User should verify.

    container.innerHTML = "";

    window.selectedProducts.forEach((prod, index) => {
        const div = document.createElement('div');
        div.className = 'claim-card-dynamic';
        div.style.cssText = "background: #fff; border: 1px solid #e5e7eb; border-radius: 12px; padding: 1.5rem; margin-bottom: 1.5rem; box-shadow: 0 4px 6px -1px rgba(0,0,0,0.05);";

        div.innerHTML = `
            <h4 style="color: var(--primary); margin-bottom: 1rem; border-bottom: 1px solid #eee; padding-bottom: 0.5rem;">
                <i class="ri-smartphone-line"></i> ${prod.model || 'New Product'}
            </h4>
            
            <div class="grid-2">
                 <div class="form-group">
                    <label style="color: var(--text-tertiary); font-size: 0.8rem; font-weight:600;">Issue Description</label>
                    <textarea id="issue_${index}" rows="2" placeholder="Describe issue for this product..." required
                        style="width: 100%; padding: 0.8rem; border-radius: 8px; border: 1px solid #ddd; background: #f9fafb; outline: none;"></textarea>
                 </div>
                 <div class="form-group">
                    <label style="color: var(--text-tertiary); font-size: 0.8rem; font-weight:600;">Upload Invoice/Docs</label>
                    <input type="file" id="files_${index}" multiple
                        style="width: 100%; padding: 0.6rem; border: 1px dashed #ccc; border-radius: 8px; background: #fbfbfb;">
                    <p style="font-size:0.75rem; color:#999; margin-top:4px;">If invoice is same as another, you can upload once here.</p>
                 </div>
            </div>
        `;
        container.appendChild(div);
    });
}

// Handle Form Submit
const subForm = document.getElementById('submissionForm');
if (subForm) {
    subForm.onsubmit = async (e) => {
        e.preventDefault();

        if (window.selectedProducts.length === 0) {
            alert("Please select at least one product.");
            return;
        }

        const formData = new FormData(subForm);
        const loader = document.getElementById('loader');
        loader.classList.remove('hidden');

        // Build Claims Data
        const claimsList = window.selectedProducts.map((prod, index) => {
            const issueEl = document.getElementById(`issue_${index}`);
            return {
                ...prod,
                issue: issueEl ? issueEl.value : "No Issue Description",
                file_key: `files_${index}`
            };
        });

        formData.append('claims_data', JSON.stringify(claimsList));

        // Append Files
        window.selectedProducts.forEach((_, index) => {
            const fileInput = document.getElementById(`files_${index}`);
            if (fileInput && fileInput.files.length > 0) {
                for (let i = 0; i < fileInput.files.length; i++) {
                    formData.append(`files_${index}`, fileInput.files[i]);
                }
            }
        });

        try {
            const res = await fetchWithTimeout('/submit-claim', {
                method: 'POST',
                body: formData
            });
            const result = await res.json();

            loader.classList.add('hidden');

            if (result.success) {
                alert(result.message);
                window.location.reload();
            } else {
                alert("Error: " + result.message);
            }
        } catch (err) {
            loader.classList.add('hidden');
            console.error(err);
            if (err.name === 'AbortError') {
                alert("Submission timed out. Server busy.");
            } else {
                alert("Submission failed. Check network.");
            }
        }
    }
}


// --- Dashboard Page ---
function filterTable() {
    const input = document.getElementById("searchInput");
    const filter = input.value.toUpperCase();
    const table = document.querySelector(".premium-table");
    const tr = table.getElementsByTagName("tr");

    for (let i = 1; i < tr.length; i++) {
        let txtValue = tr[i].textContent || tr[i].innerText;
        if (txtValue.toUpperCase().indexOf(filter) > -1) {
            tr[i].style.display = "";
        } else {
            tr[i].style.display = "none";
        }
    }
}

function filterClaims(type) {
    const table = document.querySelector(".premium-table");
    const tr = table.getElementsByTagName("tr");

    // Clear Search Input when KPI clicked
    const searchInput = document.getElementById("searchInput");
    if (searchInput) searchInput.value = "";

    for (let i = 1; i < tr.length; i++) {
        const row = tr[i];
        const isComplete = row.getAttribute('data-completed') === 'true';
        const rowStatus = (row.getAttribute('data-status') || '').toLowerCase().trim();

        if (type === 'all') {
            row.style.display = "";
        } else if (type === 'pending') {
            if (!isComplete) row.style.display = "";
            else row.style.display = "none";
        } else if (type === 'completed') {
            if (isComplete) row.style.display = "";
            else row.style.display = "none";
        } else if (type === 'cancelled') {
            if (rowStatus === 'cancelled') row.style.display = "";
            else row.style.display = "none";
        }
    }
}

// Modal Logic
let currentClaimId = null;

async function openClaimModal(id) {
    currentClaimId = id;
    document.body.style.cursor = 'wait';

    try {
        const res = await fetch(`/claim/${id}`);
        if (!res.ok) throw new Error("Fetch failed");
        const data = await res.json();

        const modal = document.getElementById('modal-container');
        const body = document.getElementById('modal-body');
        // Load READ-ONLY Template
        const tmpl = document.getElementById('claimDetailTemplate').content.cloneNode(true);

        body.innerHTML = '';
        body.appendChild(tmpl);

        // Populate Read-Only Fields
        const setText = (id, val) => { const e = document.getElementById(id); if (e) e.textContent = val || '--'; };

        setText('disp_claimId', data.id); // Ensure property names match API response
        setText('disp_status', data.status);
        setText('disp_name', data.customer_name);
        setText('disp_mobile', data.mobile_no);
        setText('disp_product', data.product_name || data.model);
        setText('disp_date', formatDate(data.date));
        setText('disp_osid', data.osid);
        setText('disp_invoice', data.invoice_no);
        setText('disp_srno', data.serial_no);

        const issueBox = document.getElementById('disp_issue');
        if (issueBox) issueBox.textContent = data.issue || "No issue description provided.";

        // Render Workflow List (Read Only)
        renderWorkflowList(data);

        // Follow Up History (View Only)
        const histEl = document.getElementById('followUpHistory');
        if (histEl) histEl.value = data.follow_up_notes || '';

        modal.classList.remove('hidden');

    } catch (e) {
        console.error(e);
        alert("Failed to load claim details.");
    } finally {
        document.body.style.cursor = 'default';
    }
}

// Editor Modal for Action Button
async function openClaimEditModal(id) {
    currentClaimId = id;
    document.body.style.cursor = 'wait';

    try {
        const res = await fetch(`/claim/${id}`);
        if (!res.ok) throw new Error("Fetch failed");
        const data = await res.json();

        const modal = document.getElementById('modal-container');
        const body = document.getElementById('modal-body');
        // Load EDITOR Template
        const tmpl = document.getElementById('claimEditTemplate').content.cloneNode(true);

        body.innerHTML = '';
        body.appendChild(tmpl);

        // Populate Inputs
        const setVal = (id, val) => { const e = document.getElementById(id); if (e) e.value = val || ''; };

        setVal('detailName', data.customer_name);
        setVal('detailMobile', data.mobile_no);
        setVal('detailAddress', data.address);
        setVal('detailProduct', data.product_name);
        setVal('detailModel', data.model);
        setVal('detailSerial', data.serial_no); // Using serial_no from backend
        setVal('detailInvoice', data.invoice_no);
        setVal('claimOsid', data.osid);
        setVal('claimSrNo', data.sr_no); // Load the actual SR No (separate from product serial number)
        setVal('submittedDate', formatDate(data.date));
        setVal('detailIssue', data.issue);

        // Status Dropdown
        setVal('updateStatus', data.status);

        // Follow Up
        setVal('followUpHistory', data.follow_up_notes || '');
        setVal('followUpDate', data.next_follow_up_date);
        setVal('assignedStaff', data.assigned_staff);
        setVal('settledDate', data.claim_settled_date);

        // Initial UI State
        const statusDropdown = document.getElementById('updateStatus');
        if (statusDropdown) {
            statusDropdown.addEventListener('change', checkStatusUI);
        }

        checkStatusUI();

        // Initialize Workflow Checkboxes
        if (window.initializeWorkflow) {
            window.initializeWorkflow(data);
        }

        modal.classList.remove('hidden');

    } catch (e) {
        console.error(e);
        alert("Failed to load editor.");
    } finally {
        document.body.style.cursor = 'default';
    }
}

function checkStatusUI() {
    const statusEl = document.getElementById('updateStatus');
    if (!statusEl) return;
    const status = statusEl.value;

    // Elements
    const tabsContainer = document.querySelector('#modal-body .tabs');
    const inputSrNo = document.getElementById('claimSrNo');
    const divSrNo = inputSrNo?.closest('.field-group');
    const followUpSection = document.querySelector('.follow-up-section');
    const wfRepair = document.getElementById('wf-section-repair');
    const wfReplacement = document.getElementById('workflow-replacement-section');

    // Default: Reset standard sections
    if (tabsContainer) tabsContainer.classList.remove('hidden'); // Default show tabs
    if (divSrNo) divSrNo.classList.add('hidden');
    if (followUpSection) followUpSection.classList.remove('hidden');
    if (wfRepair) wfRepair.classList.add('hidden');
    if (wfReplacement) wfReplacement.classList.add('hidden');

    // Manage Info Fields Visibility
    const allInfoFields = document.querySelectorAll('#tab-info .details-grid .field-group');
    const showAllInfo = () => {
        allInfoFields.forEach(el => el.classList.remove('hidden'));
        if (divSrNo) divSrNo.classList.add('hidden'); // Re-hide SR by default
    };

    // Status Logic
    if (status === 'Submitted') {
        // Submitted: No Tabs, Full View, SR Input Hidden
        if (tabsContainer) tabsContainer.classList.add('hidden');
        switchTab('info');
        showAllInfo();
    }
    else if (status === 'Registered') {
        // Registered: No Tabs, Full Info + SR No
        if (tabsContainer) tabsContainer.classList.add('hidden');
        switchTab('info');
        showAllInfo();
        if (divSrNo) divSrNo.classList.remove('hidden');
    }
    else if (status === 'Follow Up') {
        // Follow Up: No Tabs, Full Info + Follow Up Section
        if (tabsContainer) tabsContainer.classList.add('hidden');
        switchTab('info');
        showAllInfo();

        const fDate = document.getElementById('followUpDate');
        if (fDate) {
            fDate.removeAttribute('readonly');
            if (!fDate.value) fDate.value = getISTDate();
        }
    }
    else if (status === 'Repair Completed') {
        // Repair: Show Tabs, Workflow Tab
        switchTab('workflow');
        if (wfRepair) wfRepair.classList.remove('hidden');
    }
    else if (status === 'Replacement approved' || status === 'Replacement Approved') {
        // Replacement: Show Tabs, Workflow Tab
        switchTab('workflow');
        if (wfReplacement) wfReplacement.classList.remove('hidden');
    }
    else if (status === 'Closed' || status === 'Rejected' || status.includes('No Issue')) {
        // Resolved: Show Tabs
        switchTab('info');
        showAllInfo();
    }
}

// Rewrite switchTab correctly
window.switchTab = function (tabName) {
    const modalBody = document.getElementById('modal-body'); // Scope to modal
    const buttons = modalBody.querySelectorAll('.tab-btn');
    const contents = modalBody.querySelectorAll('.tab-content');

    buttons.forEach(btn => {
        if (btn.textContent.toLowerCase().includes(tabName.replace('tab-', '')))
            btn.classList.add('active'); // Dirty heuristic, but simpler is:
        else btn.classList.remove('active');
    });

    contents.forEach(c => c.classList.remove('active'));
    const target = modalBody.querySelector(`#tab-${tabName}`);
    if (target) target.classList.add('active');

    // Highlight button
    // We can assume the order Info(0), Workflow(1), Notes(2)
    buttons.forEach(b => b.classList.remove('active'));
    if (tabName === 'info') buttons[0].classList.add('active');
    if (tabName === 'workflow') buttons[1].classList.add('active');
    if (tabName === 'notes') buttons[2].classList.add('active');
}

async function saveClaimChanges() {
    // History handling
    const oldHistory = document.getElementById('followUpHistory').value;
    const newNote = document.getElementById('newFollowUpNote').value.trim();
    let combinedNotes = oldHistory;
    let newRemarks = ""; // Default if no note

    if (newNote) {
        // Append new note with timestamp
        const timestamp = new Date().toLocaleString('en-IN');
        combinedNotes = (oldHistory ? oldHistory + "\n" : "") + `[${timestamp}] ${newNote}`;
        // Remarks is just the latest note
        newRemarks = newNote;
    }

    const payload = {
        status: document.getElementById('updateStatus').value,
        claim_settled_date: document.getElementById('settledDate').value || null,

        follow_up_date: document.getElementById('followUpDate').value || null,

        // Use the combined history for "Follow Up - Notes" column
        follow_up_notes: combinedNotes,
        assigned_staff: (document.getElementById('assignedStaff')?.value || '').trim(),
        sr_no: (document.getElementById('claimSrNo')?.value || '').trim(),
    };

    if (newNote) {
        payload.remarks = newRemarks;
    }

    // Determine values from duplicated checkboxes
    const fb1 = document.getElementById('chk_repair_feedback').checked;
    const isFeedbackDone = fb1;
    
    // Extract feedback rating if feedback is done
    const feedbackRatingEl = document.getElementById('feedback_rating_value');
    const feedbackRating = (isFeedbackDone && feedbackRatingEl) ? parseInt(feedbackRatingEl.value) || 0 : 0;

    const cmp1 = document.getElementById('chk_complete_repair') ? document.getElementById('chk_complete_repair').checked : false;
    const cmp2 = document.getElementById('chk_complete_repl') ? document.getElementById('chk_complete_repl').checked : false;
    let isComplete = cmp1 || cmp2;

    // Auto-mark as complete if status is Closed, Rejected or No Issue
    if (payload.status === 'Closed' || payload.status === 'Rejected' || payload.status.includes('No Issue')) {
        isComplete = true;
    }

    // Replacement Workflow Checkboxes (Columns O-T)
    const getReplCheckbox = (id) => {
        const el = document.getElementById(id);
        return el ? el.checked : false;
    };

    // Other fields

    payload.repair_feedback_completed = isFeedbackDone;
    payload.feedback_rating = feedbackRating;

    // Replacement workflow fields (Columns O-T)
    payload.replacement_confirmation = getReplCheckbox('chk_repl_confirmation');
    payload.replacement_osg_approval = getReplCheckbox('chk_repl_osg_approval');
    payload.replacement_mail_store = getReplCheckbox('chk_repl_mail_store');
    payload.replacement_invoice_gen = getReplCheckbox('chk_repl_invoice_gen');
    payload.replacement_invoice_sent = getReplCheckbox('chk_repl_invoice_sent');
    payload.replacement_settlement_mail = getReplCheckbox('chk_repl_settlement_mail');
    payload.replacement_settled_accounts = getReplCheckbox('chk_repl_settled_accounts');

    // Complete flag (Column U)
    payload.complete = isComplete;

    try {
        const res = await fetch(`/update-claim/${currentClaimId}`, {
            method: 'POST',
            headers: { 'Content-Type': 'application/json' },
            body: JSON.stringify(payload)
        });
        const d = await res.json();
        if (d.success) {
            closeModal();
            alert("Changes saved successfully!");
            window.location.reload();
        } else {
            alert("Failed to update: " + (d.message || "Unknown error"));
        }
    } catch (e) {
        alert("Error updating");
    }
}

// Date Display
const d = new Date();
const dateStr = d.toLocaleDateString('en-IN', { weekday: 'long', year: 'numeric', month: 'long', day: 'numeric' });
const dateEl = document.getElementById('currentDate');
if (dateEl) dateEl.textContent = dateStr;

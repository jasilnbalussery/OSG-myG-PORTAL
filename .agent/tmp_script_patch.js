function checkStatusUI() {
    const status = document.getElementById('updateStatus').value;
    const isSubmitted = status === 'Submitted';
    const isRegistered = status === 'Registered';
    const isFollowUp = status === 'Follow Up';
    const isRepairCompleted = status === 'Repair Completed';

    // Tabs Buttons / Container
    const tabsContainer = document.querySelector('#modal-body .tabs');
    const btnWorkflow = document.getElementById('btn-tab-workflow');
    const btnNotes = document.getElementById('btn-tab-notes');

    // Sections
    const TabWorkflow = document.getElementById('tab-workflow');
    const TabNotes = document.getElementById('tab-notes');

    // Workflow Sub-Sections
    const wfSectionRepair = document.getElementById('wf-section-repair');
    const wfSectionRepl = document.getElementById('wf-section-replacement');

    // Info Fields
    const GrpSettled = document.getElementById('grp-settled');
    const GrpStaff = document.getElementById('grp-staff');
    const GrpSubmitted = document.getElementById('grp-submitted');

    // Controls
    const btnSave = document.getElementById('btn-save-changes');
    const subDateInput = document.getElementById('submittedDate');

    // --- Modes ---

    if (isSubmitted || isRegistered) {
        // Mode 1: Info Only (Date only)
        if (tabsContainer) tabsContainer.classList.add('hidden');

        switchTab('info');

        // Hide other tabs content explicitly
        if (TabWorkflow) TabWorkflow.classList.add('hidden');
        if (TabNotes) TabNotes.classList.add('hidden');

        // Layout Info Fields
        if (GrpSettled) GrpSettled.classList.add('hidden');
        if (GrpStaff) GrpStaff.classList.add('hidden');
        if (GrpSubmitted) GrpSubmitted.classList.remove('hidden');

        // Controls
        if (isRegistered) {
            // Registered: Set Date to Today, Readonly, Show Save
            if (btnSave) btnSave.classList.remove('hidden');
            if (subDateInput) {
                subDateInput.value = getISTDate(); // Default to Today
                subDateInput.setAttribute('readonly', true);
            }
        } else {
            // Submitted: Readonly Date (Historical), Hide Save
            if (btnSave) btnSave.classList.add('hidden');
            if (subDateInput) subDateInput.setAttribute('readonly', true);
        }

    } else if (isFollowUp) {
        // Mode 2: Follow Up Only
        if (tabsContainer) tabsContainer.classList.add('hidden');

        switchTab('notes'); // Force Follow Up Tab

        // Ensure Visibility
        if (TabNotes) TabNotes.classList.remove('hidden');
        if (TabWorkflow) TabWorkflow.classList.add('hidden');

        // Controls
        if (btnSave) btnSave.classList.remove('hidden');

        // Follow Up Date: Today, Readonly
        const followUpDateInput = document.getElementById('followUpDate');
        if (followUpDateInput) {
            followUpDateInput.value = getISTDate();
            followUpDateInput.setAttribute('readonly', true);
        }

        // Safe reset
        if (subDateInput) subDateInput.setAttribute('readonly', true);

    } else if (isRepairCompleted) {
        // Mode 3: Repair Completed (Workflow -> Repair Only)
        if (tabsContainer) tabsContainer.classList.add('hidden');

        switchTab('workflow'); // Force Workflow Tab

        // Ensure Visibility of Workflow Tab Only
        if (TabWorkflow) TabWorkflow.classList.remove('hidden');
        if (TabNotes) TabNotes.classList.add('hidden');

        // Toggle Subsections
        if (wfSectionRepair) wfSectionRepair.classList.remove('hidden');
        if (wfSectionRepl) wfSectionRepl.classList.add('hidden');

        // Controls
        if (btnSave) btnSave.classList.remove('hidden');
        if (subDateInput) subDateInput.setAttribute('readonly', true);

    } else if (status === 'Replacement Approved' || status === 'Replacement approved') {
        // Mode 4: Replacement Approved (Workflow -> Replacement Only)
        if (tabsContainer) tabsContainer.classList.remove('hidden');

        // Ensure content not forced hidden
        if (TabWorkflow) TabWorkflow.classList.remove('hidden');
        if (TabNotes) TabNotes.classList.remove('hidden');

        // Show all Buttons
        if (btnWorkflow) btnWorkflow.classList.remove('hidden');
        if (btnNotes) btnNotes.classList.remove('hidden');

        // Toggle Subsections: HIDE Repair, SHOW Replacement
        if (wfSectionRepair) wfSectionRepair.classList.add('hidden');
        if (wfSectionRepl) wfSectionRepl.classList.remove('hidden');

        // Restore Info Fields
        if (GrpSettled) GrpSettled.classList.remove('hidden');
        if (GrpStaff) GrpStaff.classList.remove('hidden');
        if (GrpSubmitted) GrpSubmitted.classList.remove('hidden');

        // Controls
        if (btnSave) btnSave.classList.remove('hidden');
        if (subDateInput) subDateInput.setAttribute('readonly', true);

    } else {
        // Mode 5: Full Workflow / Other
        if (tabsContainer) tabsContainer.classList.remove('hidden');

        // Show all Buttons
        if (btnWorkflow) btnWorkflow.classList.remove('hidden');
        if (btnNotes) btnNotes.classList.remove('hidden');

        // Ensure content not forced hidden
        if (TabWorkflow) TabWorkflow.classList.remove('hidden');
        if (TabNotes) TabNotes.classList.remove('hidden');

        // Show all Workflow Subsections
        if (wfSectionRepair) wfSectionRepair.classList.remove('hidden');
        if (wfSectionRepl) wfSectionRepl.classList.remove('hidden');

        // Restore Info Fields
        if (GrpSettled) GrpSettled.classList.remove('hidden');
        if (GrpStaff) GrpStaff.classList.remove('hidden');
        if (GrpSubmitted) GrpSubmitted.classList.remove('hidden');

        // Controls
        if (btnSave) btnSave.classList.remove('hidden');
        if (subDateInput) subDateInput.setAttribute('readonly', true);
    }
}

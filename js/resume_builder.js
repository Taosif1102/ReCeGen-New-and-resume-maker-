document.addEventListener('DOMContentLoaded', () => {
    // Start Mappings
    // Mapping of input IDs to preview IDs
    const mappings = {
        'input-name': ['preview-name', 'preview-pd-name', 'preview-sig-name'],
        'input-co': 'preview-co',
        'input-addr1': 'preview-addr1',
        'input-addr2': 'preview-addr2',
        'input-addr3': 'preview-addr3',
        'input-mobile': 'preview-mobile',
        'input-email': 'preview-email',
        'input-objective': 'preview-objective',

        // Diploma
        'input-diploma-inst': 'preview-diploma-inst',
        'input-diploma-board': 'preview-diploma-board',
        'input-diploma-subj': 'preview-diploma-subj',
        'input-diploma-year': 'preview-diploma-year',

        // SSC
        'input-ssc-school': 'preview-ssc-school',
        'input-ssc-board': 'preview-ssc-board',
        'input-ssc-group': 'preview-ssc-group',
        'input-ssc-year': 'preview-ssc-year',
        'input-ssc-gpa': 'preview-ssc-gpa',

        // Computer
        'input-comp-apps': 'preview-comp-apps',
        'input-comp-other': 'preview-comp-other',
        'input-comp-prog': 'preview-comp-prog',

        // Language
        'input-lang-bengali': 'preview-lang-bengali',
        'input-lang-english': 'preview-lang-english',

        // Personal Details
        'input-pd-father': 'preview-pd-father',
        'input-pd-mother': 'preview-pd-mother',
        'input-pd-dob': 'preview-pd-dob',
        'input-pd-gender': 'preview-pd-gender',
        'input-pd-marital': 'preview-pd-marital',
        'input-pd-religion': 'preview-pd-religion',
        'input-pd-nationality': 'preview-pd-nationality',
        'input-pd-blood': 'preview-pd-blood',
        'input-pd-height': 'preview-pd-height',
        'input-pd-nid': 'preview-pd-nid',
        'input-pd-perm-addr': 'preview-pd-perm-addr',

        // Declaration & Date
        'input-declaration': 'preview-declaration',
        'input-date': 'preview-date'
    };
    // End Mappings

    // Start Event Listeners
    // Add event listeners for standard fields
    for (const [inputId, previewIds] of Object.entries(mappings)) {
        const input = document.getElementById(inputId);
        if (!input) continue;

        input.addEventListener('input', (e) => {
            const value = e.target.value;

            if (Array.isArray(previewIds)) {
                previewIds.forEach(id => {
                    const el = document.getElementById(id);
                    if (el) el.textContent = value;
                });
            } else {
                const el = document.getElementById(previewIds);
                if (el) el.textContent = value;
            }
        });
    }
    // End Event Listeners

    // Start Key Qualities
    // Special handling for Key Qualities (List)
    const qualitiesInput = document.getElementById('input-qualities');
    const qualitiesPreview = document.getElementById('preview-qualities');

    if (qualitiesInput && qualitiesPreview) {
        qualitiesInput.addEventListener('input', (e) => {
            const lines = e.target.value.split('\n');
            qualitiesPreview.innerHTML = ''; // Clear existing list

            lines.forEach(line => {
                if (line.trim() !== '') {
                    const li = document.createElement('li');
                    li.textContent = line;
                    qualitiesPreview.appendChild(li);
                }
            });
        });
    }
    // End Key Qualities

    // Start Image Upload Handling
    // Image Upload Handling
    const photoInput = document.getElementById('input-photo');
    const photoPreview = document.getElementById('preview-photo');
    const photoPlaceholder = document.getElementById('photo-placeholder');

    if (photoInput && photoPreview && photoPlaceholder) {
        photoInput.addEventListener('change', (e) => {
            const file = e.target.files[0];
            if (file) {
                const reader = new FileReader();
                reader.onload = (e) => {
                    photoPreview.src = e.target.result;
                    photoPreview.style.display = 'block';
                    photoPlaceholder.style.display = 'none';
                };
                reader.readAsDataURL(file);
            } else {
                photoPreview.src = '';
                photoPreview.style.display = 'none';
                photoPlaceholder.style.display = 'block';
            }
        });
    }
    // End Image Upload Handling
    
    // Start Auto Save & Clear Logic
    const form = document.getElementById('resume-form');
    const STORAGE_KEY = 'resume_builder_data';

    function saveData() {
        if (localStorage.getItem('autoSaveEnabled') !== 'true') return;

        const data = {};
        const inputs = form.querySelectorAll('input, textarea');
        inputs.forEach(input => {
            if (input.type === 'file') return; // Skip files
            data[input.id] = input.value;
        });
        localStorage.setItem(STORAGE_KEY, JSON.stringify(data));
    }

    function loadData() {
        if (localStorage.getItem('autoSaveEnabled') !== 'true') return;

        const data = JSON.parse(localStorage.getItem(STORAGE_KEY));
        if (!data) return;

        Object.keys(data).forEach(id => {
            const input = document.getElementById(id);
            if (input) {
                input.value = data[id];
                // Trigger input event to update preview
                input.dispatchEvent(new Event('input'));
            }
        });
    }

    function clearData() {
        localStorage.removeItem(STORAGE_KEY);
        form.reset();
        // Reset preview by triggering input events on empty fields
        const inputs = form.querySelectorAll('input, textarea');
        inputs.forEach(input => {
            if (input.type !== 'file') {
                input.value = ''; // Ensure value is cleared
                input.dispatchEvent(new Event('input'));
            }
        });
        // Restore default placeholders if needed, or just let them be empty
        // For specific fields like lists, we might need extra handling, but dispatching input handles most.
        document.getElementById('preview-photo').style.display = 'none';
        document.getElementById('photo-placeholder').style.display = 'block';
    }

    // Attach Auto Save to all inputs
    if (form) {
        form.addEventListener('input', saveData);
    }

    // Listen for global events from theme.js
    document.addEventListener('autosave-trigger', saveData);
    document.addEventListener('clear-data-trigger', clearData);

    // Load data on startup
    loadData();
    // End Auto Save & Clear Logic
});

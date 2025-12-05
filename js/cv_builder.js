document.addEventListener('DOMContentLoaded', () => {
    // Start Input Mapping
    // Mapping of input IDs to preview IDs
    const mappings = {
        'input-name': 'preview-name',
        'input-title': 'preview-title',
        'input-phone': 'preview-phone',
        'input-email': 'preview-email',
        'input-linkedin': 'preview-linkedin',
        'input-location': 'preview-location',
        'input-summary': 'preview-summary'
    };
    // End Input Mapping

    // Start Event Listeners
    // Attach event listeners for simple text fields
    for (const [inputId, previewId] of Object.entries(mappings)) {
        const input = document.getElementById(inputId);
        const preview = document.getElementById(previewId);
        if (input && preview) {
            input.addEventListener('input', (e) => {
                preview.textContent = e.target.value;
            });
        }
    }
    // End Event Listeners

    // Start Photo Handling
    // Photo handling
    const photoInput = document.getElementById('input-photo');
    const photoPreview = document.getElementById('preview-photo');
    const photoPlaceholder = document.getElementById('photo-placeholder');

    if (photoInput) {
        photoInput.addEventListener('change', function(e) {
            const file = e.target.files[0];
            if (file) {
                const reader = new FileReader();
                reader.onload = function(e) {
                    photoPreview.src = e.target.result;
                    photoPreview.style.display = 'block';
                    if (photoPlaceholder) photoPlaceholder.style.display = 'none';
                }
                reader.readAsDataURL(file);
            }
        });
    }
    // End Photo Handling

    // Start Dynamic Lists
    // Dynamic Lists (Skills, Languages, Experience, Education)
    // For simplicity in this version, we'll assume textareas with line-separated values for lists
    // and specific fields for Experience/Education, but let's try to make it a bit more structured for Exp/Edu if possible.
    // Actually, to match the image exactly, let's use specific inputs for the first item and maybe a way to add more?
    // For now, let's stick to updating the provided structure.

    // Skills
    const skillsInput = document.getElementById('input-skills');
    const skillsPreview = document.getElementById('preview-skills');
    if (skillsInput && skillsPreview) {
        skillsInput.addEventListener('input', (e) => {
            const skills = e.target.value.split('\n').filter(s => s.trim());
            skillsPreview.innerHTML = skills.map(s => `<li class="skill-item mb-2 text-sm">${s}</li>`).join('');
        });
    }

    // Languages
    const langInput = document.getElementById('input-languages');
    const langPreview = document.getElementById('preview-languages');
    if (langInput && langPreview) {
        langInput.addEventListener('input', (e) => {
            const langs = e.target.value.split('\n').filter(s => s.trim());
            langPreview.innerHTML = langs.map(s => `<li class="lang-item mb-2 text-sm">${s}</li>`).join('');
        });
    }
    // End Dynamic Lists

    // Start Experience Update
    // Experience (Handling one entry for now as a demo, or multiple if we use a container)
    // Let's implement a simple update for the first experience block
    const updateExperience = () => {
        const title = document.getElementById('input-job-title').value;
        const company = document.getElementById('input-job-company').value;
        const date = document.getElementById('input-job-date').value;
        const desc = document.getElementById('input-job-desc').value;

        const previewContainer = document.getElementById('preview-experience');
        // For this demo, we'll just replace the content. In a full app, we'd append.
        previewContainer.innerHTML = `
            <div class="job-item mb-6">
                <div class="job-title font-bold text-base text-gray-800">${title}</div>
                <div class="job-company italic text-gray-500 text-sm mb-1">${company}</div>
                <div class="job-date text-xs text-gray-400 mb-2">${date}</div>
                <div class="job-desc text-sm leading-relaxed text-gray-600">${desc.replace(/\n/g, '<br>')}</div>
            </div>
        `;
    };

    ['input-job-title', 'input-job-company', 'input-job-date', 'input-job-desc'].forEach(id => {
        const el = document.getElementById(id);
        if (el) el.addEventListener('input', updateExperience);
    });
    // End Experience Update

    // Start Education Update
    // Education
    const updateEducation = () => {
        const degree = document.getElementById('input-edu-degree').value;
        const school = document.getElementById('input-edu-school').value;
        const date = document.getElementById('input-edu-date').value;
        const desc = document.getElementById('input-edu-desc').value;

        const previewContainer = document.getElementById('preview-education');
        previewContainer.innerHTML = `
            <div class="edu-item mb-6">
                <div class="edu-degree font-bold text-base text-gray-800">${degree}</div>
                <div class="edu-school italic text-gray-500 text-sm mb-1">${school}</div>
                <div class="edu-date text-xs text-gray-400 mb-2">${date}</div>
                <div class="job-desc text-sm leading-relaxed text-gray-600">${desc.replace(/\n/g, '<br>')}</div>
            </div>
        `;
    };

    ['input-edu-degree', 'input-edu-school', 'input-edu-date', 'input-edu-desc'].forEach(id => {
        const el = document.getElementById(id);
        if (el) el.addEventListener('input', updateEducation);
    });
    // End Education Update
    
    // Start Auto Save & Clear Logic
    const form = document.getElementById('cv-form');
    const STORAGE_KEY = 'cv_builder_data';

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
        document.getElementById('preview-photo').style.display = 'none';
        if (document.getElementById('photo-placeholder')) {
            document.getElementById('photo-placeholder').style.display = 'block';
        }
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

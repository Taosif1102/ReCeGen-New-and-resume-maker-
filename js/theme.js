document.addEventListener('DOMContentLoaded', () => {
    // Start Theme Menu UI
    // Create Theme Menu UI
    const themeContainer = document.createElement('div');
    themeContainer.className = 'theme-container';
    
    const themes = [
        { id: 'light', color: '#fffdf5', title: 'Light (Warm)' },
        { id: 'dark-blue', color: '#0f172a', title: 'Dark Blue' },
        { id: 'red', color: '#fff5f5', title: 'Soft Red' },
        { id: 'orange', color: '#fffaf0', title: 'Orange' },
        { id: 'green', color: '#f0fff4', title: 'Green' },
        { id: 'purple', color: '#faf5ff', title: 'Purple' }
    ];

    const optionsContainer = document.createElement('div');
    optionsContainer.className = 'theme-options';

    themes.forEach(t => {
        const btn = document.createElement('div');
        btn.className = `theme-btn btn-${t.id}`;
        btn.title = t.title;
        btn.dataset.theme = t.id;
        btn.style.backgroundColor = t.color; // Ensure preview color is correct
        
        btn.addEventListener('click', () => applyTheme(t.id));
        optionsContainer.appendChild(btn);
    });

    const toggleBtn = document.createElement('div');
    toggleBtn.className = 'theme-toggle-main';
    toggleBtn.innerHTML = '<i class="fas fa-palette"></i>';
    
    toggleBtn.addEventListener('click', () => {
        themeContainer.classList.toggle('active');
    });

    themeContainer.appendChild(optionsContainer);
    
    // Only add floating toggle button on index.html
    const isIndexPage = window.location.pathname.endsWith('index.html') || window.location.pathname.endsWith('/');
    const isBuilderPage = window.location.pathname.includes('builder.html');
    
    if (isIndexPage) {
        themeContainer.appendChild(toggleBtn);
    }

    if (isBuilderPage) {
        const divider = document.createElement('div');
        divider.style.cssText = 'height: 1px; background: rgba(0,0,0,0.1); margin: 10px 0;';
        optionsContainer.appendChild(divider);

        // Auto Save Toggle
        const autoSaveContainer = document.createElement('div');
        autoSaveContainer.style.cssText = 'display: flex; justify-content: space-between; items-center; padding: 10px; color: var(--text-color); font-size: 0.9rem;';
        
        const autoSaveLabel = document.createElement('span');
        autoSaveLabel.textContent = 'Auto Save';
        
        const autoSaveSwitch = document.createElement('label');
        autoSaveSwitch.className = 'switch';
        autoSaveSwitch.style.cssText = 'position: relative; display: inline-block; width: 40px; height: 20px;';
        
        const autoSaveInput = document.createElement('input');
        autoSaveInput.type = 'checkbox';
        autoSaveInput.id = 'auto-save-toggle';
        autoSaveInput.checked = localStorage.getItem('autoSaveEnabled') === 'true';
        
        const slider = document.createElement('span');
        slider.style.cssText = 'position: absolute; cursor: pointer; top: 0; left: 0; right: 0; bottom: 0; background-color: #ccc; transition: .4s; border-radius: 20px;';
        
        // Add slider pseudo-element styles dynamically
        const style = document.createElement('style');
        style.textContent = `
            #auto-save-toggle:checked + span { background-color: var(--primary-color); }
            #auto-save-toggle:checked + span:before { transform: translateX(20px); }
            #auto-save-toggle + span:before { position: absolute; content: ""; height: 16px; width: 16px; left: 2px; bottom: 2px; background-color: white; transition: .4s; border-radius: 50%; }
        `;
        document.head.appendChild(style);

        autoSaveSwitch.appendChild(autoSaveInput);
        autoSaveSwitch.appendChild(slider);
        
        autoSaveContainer.appendChild(autoSaveLabel);
        autoSaveContainer.appendChild(autoSaveSwitch);
        optionsContainer.appendChild(autoSaveContainer);

        // Clear Data Button
        const clearBtn = document.createElement('button');
        clearBtn.textContent = 'Clear Data';
        clearBtn.style.cssText = 'width: 100%; padding: 8px; margin-top: 5px; background: #fee2e2; color: #dc2626; border: none; border-radius: 6px; cursor: pointer; font-size: 0.9rem; font-weight: 500; transition: background 0.2s;';
        clearBtn.onmouseover = () => clearBtn.style.background = '#fecaca';
        clearBtn.onmouseout = () => clearBtn.style.background = '#fee2e2';
        clearBtn.id = 'clear-data-btn';
        
        optionsContainer.appendChild(clearBtn);

        // Event Listeners for new controls
        autoSaveInput.addEventListener('change', (e) => {
            localStorage.setItem('autoSaveEnabled', e.target.checked);
            if (e.target.checked) {
                // Trigger immediate save if enabled
                document.dispatchEvent(new CustomEvent('autosave-trigger'));
            }
        });

        clearBtn.addEventListener('click', () => {
            if (confirm('Are you sure you want to clear all data? This cannot be undone.')) {
                document.dispatchEvent(new CustomEvent('clear-data-trigger'));
            }
        });
    }

    document.body.appendChild(themeContainer);

    // Expose global function to toggle settings
    window.toggleThemeSettings = () => {
        themeContainer.classList.toggle('active');
    };
    // End Theme Menu UI

    // Start Event Listeners
    // Close menu when clicking outside
    document.addEventListener('click', (e) => {
        if (!themeContainer.contains(e.target)) {
            themeContainer.classList.remove('active');
        }
    });
    // End Event Listeners

    // Start Theme Application Logic
    // Theme Application Logic
    function applyTheme(themeName) {
        // Remove all theme classes
        document.documentElement.classList.remove('theme-dark-blue', 'theme-red', 'theme-orange', 'theme-green', 'theme-purple');
        
        // Add new theme class (if not light, which is default)
        if (themeName !== 'light') {
            document.documentElement.classList.add(`theme-${themeName}`);
        }
        
        localStorage.setItem('theme', themeName);
    }
    // End Theme Application Logic

    // Start Initialize Theme
    // Initialize Theme
    let savedTheme = localStorage.getItem('theme') || 'light';
    
    // Migration: Replace old 'dark' theme with 'dark-blue'
    if (savedTheme === 'dark') {
        savedTheme = 'dark-blue';
        localStorage.setItem('theme', savedTheme);
    }
    
    applyTheme(savedTheme);
    // End Initialize Theme
});

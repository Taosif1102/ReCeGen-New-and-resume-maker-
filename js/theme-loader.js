(function() {
    let savedTheme = localStorage.getItem('theme') || 'light';
    if (savedTheme === 'dark') {
        savedTheme = 'dark-blue';
        localStorage.setItem('theme', savedTheme);
    }
    if (savedTheme !== 'light') {
        document.documentElement.classList.add(`theme-${savedTheme}`);
    }
})();

// Header layout preview switcher: demo aid for choosing between the three
// header arrangements. Delete this file and its <script> tag once a layout
// is chosen and hardcoded.
(function () {
    var STORAGE_KEY = 'portal_header_layout';
    var MODES = [
        { id: 'left', label: 'Seal left' },
        { id: 'right', label: 'Seal right' },
        { id: 'center', label: 'Centered' }
    ];

    function apply(mode) {
        if (mode === 'left') {
            delete document.body.dataset.header;
        } else {
            document.body.dataset.header = mode;
        }
        var buttons = document.querySelectorAll('.header-layout-switcher button');
        buttons.forEach(function (btn) {
            var active = btn.dataset.mode === mode;
            btn.classList.toggle('active', active);
            btn.setAttribute('aria-pressed', active ? 'true' : 'false');
        });
    }

    function init() {
        var bar = document.createElement('div');
        bar.className = 'header-layout-switcher';
        bar.setAttribute('role', 'group');
        bar.setAttribute('aria-label', 'Header layout preview');
        MODES.forEach(function (mode) {
            var btn = document.createElement('button');
            btn.type = 'button';
            btn.textContent = mode.label;
            btn.dataset.mode = mode.id;
            btn.addEventListener('click', function () {
                try { localStorage.setItem(STORAGE_KEY, mode.id); } catch (err) {}
                apply(mode.id);
            });
            bar.appendChild(btn);
        });
        document.body.appendChild(bar);

        var saved = 'left';
        try { saved = localStorage.getItem(STORAGE_KEY) || 'left'; } catch (err) {}
        apply(saved);
    }

    if (document.readyState === 'loading') {
        document.addEventListener('DOMContentLoaded', init);
    } else {
        init();
    }
})();

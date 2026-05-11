// Disable right-click menu
window.addEventListener('contextmenu', function (e) {
    e.preventDefault();
}, false);

// Disable zoom (Ctrl + / Ctrl - / Ctrl + wheel)
window.addEventListener('wheel', function (e) {
    if (e.ctrlKey) e.preventDefault();
}, { passive: false });

window.addEventListener('keydown', function (e) {
    // Block zoom keys
    if (e.ctrlKey && (e.key === '+' || e.key === '-' || e.key === '=')) {
        e.preventDefault();
    }

    // Block reload keys
    if (e.key === 'F5' || (e.ctrlKey && e.key === 'r')) {
        e.preventDefault();
    }

    // Block closing tab/window shortcut
    if (e.ctrlKey && e.key === 'w') {
        e.preventDefault();
    }

    // Block dev tools
    if (e.key === 'F12' || (e.ctrlKey && e.shiftKey && e.key === 'I')) {
        e.preventDefault();
    }

    // Block Ctrl+L (address bar)
    if (e.ctrlKey && e.key === 'l') {
        e.preventDefault();
    }

    // Block Backspace navigation
    if (e.key === 'Backspace' && !e.target.matches("input, textarea")) {
        e.preventDefault();
    }
}, false);

// Disable dragging of page elements/images/links
window.addEventListener("dragstart", function (e) {
    e.preventDefault();
}, false);

// Trigger exit prompt on ESC key
window.addEventListener('keydown', function(e) {
    if (e.key === 'Escape') {
        // Ask Python (PyWebView) to run the exit PIN prompt
        window.pywebview.api.request_exit();
    }
});

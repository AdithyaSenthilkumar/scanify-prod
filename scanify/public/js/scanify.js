// scanify/public/js/scanify.js

frappe.provide('scanify');

// Set home page on boot
frappe.ready(function() {
    // Override default home page
    if (frappe.boot) {
        frappe.boot.home_page = 'scanify';
    }
});

// Redirect to Scanify workspace after login
$(document).on('startup', function() {
    // If user just logged in and on default desk page, redirect
    if (frappe.get_route()[0] === 'workspace' && 
        frappe.get_route()[1] === 'home') {
        frappe.set_route('workspace', 'scanify');
    }
    
    // If on root /app, redirect to scanify
    if (frappe.get_route().length === 0 || 
        (frappe.get_route().length === 1 && frappe.get_route()[0] === 'workspace')) {
        frappe.set_route('workspace', 'scanify');
    }
});

// Override get_home_page function
frappe.get_home_page = function() {
    return 'scanify';
};

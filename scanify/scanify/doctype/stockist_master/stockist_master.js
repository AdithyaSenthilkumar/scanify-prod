frappe.ui.form.on('Stockist Master', {
	refresh: function(frm) {
		if (!frm.is_new() && frm.doc.status === 'Active') {
			frm.add_custom_button(__('View Statements'), function() {
				frappe.set_route('List', 'Stockist Statement', {
					'stockist_code': frm.doc.name
				});
			});
			
			frm.add_custom_button(__('Create New Statement'), function() {
				frappe.new_doc('Stockist Statement', {
					'stockist_code': frm.doc.name
				});
			});
		}
	},
	
	hq: function(frm) {
		if (frm.doc.hq) {
			frappe.db.get_value('HQ Master', frm.doc.hq, ['team', 'region', 'zone'], (r) => {
				frm.set_value('team', r.team);
				frm.set_value('region', r.region);
				frm.set_value('zone', r.zone);
			});
		}
	}
});

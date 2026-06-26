frappe.ui.form.on('Product Master', {
	refresh: function(frm) {
		set_excluded_region_query(frm);
		if (!frm.is_new()) {
			frm.add_custom_button(__('View Usage in Statements'), function() {
				frappe.set_route('query-report', 'Product Wise Sales', {
					'product_code': frm.doc.name
				});
			});
		}
	},

	division: function(frm) {
		// Region picker is scoped to the product's own division.
		set_excluded_region_query(frm);
	},

	// Live-refresh the comma-separated code preview whenever a region is removed.
	excluded_regions_remove: function(frm) {
		refresh_excluded_region_codes(frm);
	},

	mrp: function(frm) {
		if (frm.doc.mrp && !frm.doc.pts) {
			let suggested_pts = frm.doc.mrp * 0.69;
			frappe.msgprint(__('Suggested PTS (69% of MRP): ₹') + suggested_pts.toFixed(2));
		}
	}
});

frappe.ui.form.on('Product Excluded Region', {
	// Fires when a region is picked in a child row — keep the code preview in sync.
	region: function(frm) {
		refresh_excluded_region_codes(frm);
	}
});

function set_excluded_region_query(frm) {
	frm.set_query('region', 'excluded_regions', function() {
		const filters = { status: 'Active' };
		if (frm.doc.division) {
			filters.division = frm.doc.division;
		}
		return { filters: filters };
	});
}

function refresh_excluded_region_codes(frm) {
	const seen = [];
	(frm.doc.excluded_regions || []).forEach(function(row) {
		const code = (row.region || '').trim();
		if (code && seen.indexOf(code) === -1) {
			seen.push(code);
		}
	});
	frm.set_value('excluded_region_codes', seen.join(', '));
}

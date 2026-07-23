import frappe

FIELDS = ("proof_attachment_1", "proof_attachment_2", "proof_attachment_3", "proof_attachment_4")


def execute():
    """Link Scheme Request proof files to their scheme.

    Proof documents are uploaded BEFORE the scheme exists (upload_scheme_attachment) and
    only re-pointed to it after insert. Where that link is missing — legacy rows, an
    interrupted submit, or a scheme created through a path that never re-pointed —
    Frappe's File.has_permission falls through to `return False`, so /private/files/<name>
    returns 403 for everyone except the uploader and Administrator. A portal Admin with
    System Manager still could not download the attachment.

    Downloads now go through scanify.api.download_scheme_attachment (portal-authorised),
    so this patch is a data cleanup: it makes Desk show the files as attachments and
    unblocks any direct /private/files link that is still in circulation.

    Add-only and idempotent: only fills File rows that have no attached_to_name, and
    never re-points a file already attached elsewhere.
    """
    cols = ", ".join("`%s`" % f for f in FIELDS)
    where = " OR ".join("(`%s` IS NOT NULL AND `%s` != '')" % (f, f) for f in FIELDS)
    schemes = frappe.db.sql(
        "SELECT name, %s FROM `tabScheme Request` WHERE %s" % (cols, where), as_dict=True)

    linked = missing = 0
    for sr in schemes:
        for field in FIELDS:
            url = (sr.get(field) or "").strip()
            if not url:
                continue
            # Only claim files not attached to anything yet. "is / not set" maps to
            # IFNULL(field,'')='' so it matches BOTH NULL and '' — an IN [None, '']
            # filter becomes SQL IN (NULL, '') and would never match NULL rows.
            names = frappe.db.get_all(
                "File",
                filters={"file_url": url, "attached_to_name": ["is", "not set"]},
                pluck="name",
            )
            if not names:
                missing += 1
                continue
            for fname in names:
                frappe.db.set_value("File", fname, {
                    "attached_to_doctype": "Scheme Request",
                    "attached_to_name": sr["name"],
                    "attached_to_field": field,
                }, update_modified=False)
                linked += 1

    frappe.db.commit()
    print("link_scheme_proof_files: linked %d file(s); %d attachment url(s) had no "
          "unattached File row (already linked or file record absent)" % (linked, missing))

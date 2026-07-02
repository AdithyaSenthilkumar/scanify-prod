import frappe


def execute():
    """Drop any global UNIQUE index on Product Master.product_code.

    Product Master's autoname changed from `field:product_code` to a series
    (`format:PRD-{####}`) so the same Product Code can be reused across DIFFERENT
    divisions (uniqueness is enforced per-division in the Product Master
    controller). Existing documents keep their old names (name == product_code);
    only new products get series names. MariaDB does not always drop a
    pre-existing unique index during a normal schema sync, so remove it
    explicitly and idempotently.
    """
    table = "tabProduct Master"

    try:
        indexes = frappe.db.sql(
            f"SHOW INDEX FROM `{table}` WHERE Column_name = 'product_code'",
            as_dict=True,
        )
    except Exception as e:
        print(f"Could not inspect indexes on {table}: {e}")
        return

    dropped = set()
    for idx in indexes:
        key = idx.get("Key_name")
        non_unique = str(idx.get("Non_unique"))
        if key and key != "PRIMARY" and key not in dropped and non_unique == "0":
            try:
                frappe.db.sql(f"ALTER TABLE `{table}` DROP INDEX `{key}`")
                dropped.add(key)
                print(f"✓ Dropped unique index `{key}` on {table}.product_code")
            except Exception as e:
                print(f"✗ Could not drop index `{key}`: {e}")

    if not dropped:
        print("No unique index on Product Master.product_code to drop (already removed).")

    frappe.db.commit()

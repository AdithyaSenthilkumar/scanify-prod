import frappe


def execute():
    """Drop the global UNIQUE index on Stockist Master.stockist_code.

    The `unique` flag was removed from the field so the same editable Stockist Code
    can be reused across DIFFERENT divisions (uniqueness is now enforced per-division
    in the Stockist Master controller). MariaDB does not always drop the pre-existing
    unique index during a normal schema sync, so remove it explicitly and idempotently.
    """
    table = "tabStockist Master"

    try:
        indexes = frappe.db.sql(
            f"SHOW INDEX FROM `{table}` WHERE Column_name = 'stockist_code'",
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
                print(f"✓ Dropped unique index `{key}` on {table}.stockist_code")
            except Exception as e:
                print(f"✗ Could not drop index `{key}`: {e}")

    if not dropped:
        print("No unique index on Stockist Master.stockist_code to drop (already removed).")

    frappe.db.commit()

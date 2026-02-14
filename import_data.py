"""Import existing data from Onna Business .xlsx into the SQLite database."""
import os
import sys
from datetime import datetime, date
import openpyxl
from app import app, db, Item


def import_excel(filepath):
    wb = openpyxl.load_workbook(filepath, data_only=True)
    ws = wb['Log']

    items = []
    for row_idx, row in enumerate(ws.iter_rows(min_row=2, max_col=12, values_only=True), start=2):
        date_bought, date_sold, description, cost, listing_price, sold_for, \
            _pred_profit, _actual_profit, _pred_margin, _actual_margin, _days, status = row[:12]

        # Skip completely empty rows
        if not description:
            continue

        # Skip supply/expense rows (like "Furniture Restorer")
        # Still import them but mark appropriately

        # Parse dates
        db_bought = _to_date(date_bought)
        db_sold = _to_date(date_sold)

        # Parse numbers
        db_cost = _to_float(cost) or 0
        db_listing = _to_float(listing_price)
        db_sold_for = _to_float(sold_for)

        # Determine status
        db_status = 'Listed'
        if status and isinstance(status, str):
            if status.strip().lower() == 'sold':
                db_status = 'Sold'
            elif status.strip().lower() == 'listed':
                db_status = 'Listed'
        elif db_sold_for and db_sold:
            # If no explicit status but has sold_for and date_sold, mark as Sold
            db_status = 'Sold'

        item = Item(
            date_bought=db_bought,
            date_sold=db_sold,
            description=str(description).strip(),
            cost=db_cost,
            listing_price=db_listing,
            sold_for=db_sold_for,
            status=db_status,
        )
        items.append(item)
        print(f"  Row {row_idx}: {description} | Cost: ${db_cost} | Status: {db_status}")

    return items


def _to_date(val):
    if val is None:
        return None
    if isinstance(val, datetime):
        return val.date()
    if isinstance(val, date):
        return val
    return None


def _to_float(val):
    if val is None:
        return None
    try:
        f = float(val)
        return f if f >= 0 else None
    except (ValueError, TypeError):
        return None


if __name__ == '__main__':
    excel_path = os.path.join(os.path.dirname(__file__), 'Onna Business .xlsx')

    if not os.path.exists(excel_path):
        print(f"ERROR: Could not find {excel_path}")
        sys.exit(1)

    with app.app_context():
        # Clear existing data
        existing = Item.query.count()
        if existing > 0:
            print(f"Clearing {existing} existing items...")
            Item.query.delete()
            db.session.commit()

        print(f"\nImporting from: {excel_path}")
        print("-" * 60)
        items = import_excel(excel_path)

        db.session.add_all(items)
        db.session.commit()

        print("-" * 60)
        print(f"\nImported {len(items)} items successfully!")

        # Quick summary
        sold = sum(1 for i in items if i.status == 'Sold')
        listed = sum(1 for i in items if i.status == 'Listed')
        print(f"  Sold: {sold}")
        print(f"  Listed: {listed}")
        total_profit = sum(i.actual_profit or 0 for i in items if i.status == 'Sold')
        print(f"  Total profit: ${total_profit:.0f}")

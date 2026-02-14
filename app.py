import io
import os
from datetime import datetime, date
from flask import Flask, render_template, request, jsonify, redirect, url_for, send_file
from flask_sqlalchemy import SQLAlchemy
from reportlab.lib import colors
from reportlab.lib.pagesizes import letter
from reportlab.lib.units import inch
from reportlab.platypus import SimpleDocTemplate, Table, TableStyle, Paragraph, Spacer
from reportlab.lib.styles import getSampleStyleSheet, ParagraphStyle

basedir = os.path.abspath(os.path.dirname(__file__))

app = Flask(__name__)
app.config['SQLALCHEMY_DATABASE_URI'] = 'sqlite:///' + os.path.join(basedir, 'instance', 'onna_business.db')
app.config['SQLALCHEMY_TRACK_MODIFICATIONS'] = False
app.config['SECRET_KEY'] = os.environ.get('SECRET_KEY', 'onna-flips-dev-key')

os.makedirs(os.path.join(basedir, 'instance'), exist_ok=True)

db = SQLAlchemy(app)


class Item(db.Model):
    id = db.Column(db.Integer, primary_key=True)
    date_bought = db.Column(db.Date, nullable=True)
    date_sold = db.Column(db.Date, nullable=True)
    description = db.Column(db.String(200), nullable=False)
    cost = db.Column(db.Float, default=0)
    listing_price = db.Column(db.Float, nullable=True)
    sold_for = db.Column(db.Float, nullable=True)
    status = db.Column(db.String(20), default='Listed')  # Listed, Sold
    notes = db.Column(db.Text, nullable=True)

    @property
    def actual_profit(self):
        if self.sold_for is not None and self.cost is not None:
            return self.sold_for - self.cost
        return None

    @property
    def predicted_profit(self):
        if self.status != 'Sold' and self.listing_price and self.cost is not None:
            return self.listing_price - self.cost
        return None

    @property
    def actual_margin(self):
        if self.sold_for and self.sold_for > 0 and self.cost is not None:
            return round((self.sold_for - self.cost) / self.sold_for, 4)
        return None

    @property
    def days_to_sell(self):
        if self.date_bought and self.date_sold:
            return (self.date_sold - self.date_bought).days
        return None

    @property
    def profit_per_day(self):
        days = self.days_to_sell
        profit = self.actual_profit
        if days and days > 0 and profit is not None:
            return round(profit / days, 2)
        return None

    def to_dict(self):
        return {
            'id': self.id,
            'date_bought': self.date_bought.isoformat() if self.date_bought else None,
            'date_sold': self.date_sold.isoformat() if self.date_sold else None,
            'description': self.description,
            'cost': self.cost,
            'listing_price': self.listing_price,
            'sold_for': self.sold_for,
            'status': self.status,
            'notes': self.notes,
            'actual_profit': self.actual_profit,
            'predicted_profit': self.predicted_profit,
            'actual_margin': self.actual_margin,
            'days_to_sell': self.days_to_sell,
            'profit_per_day': self.profit_per_day,
        }


# --- Page Routes ---

@app.route('/')
def dashboard():
    return render_template('dashboard.html')


@app.route('/inventory')
def inventory():
    return render_template('inventory.html')


@app.route('/add')
def add_item_page():
    return render_template('add_item.html')


@app.route('/edit/<int:item_id>')
def edit_item_page(item_id):
    return render_template('edit_item.html', item_id=item_id)


@app.route('/tax-export')
def tax_export_page():
    return render_template('tax_export.html')


# --- API Routes ---

@app.route('/api/items', methods=['GET'])
def get_items():
    status_filter = request.args.get('status')
    query = Item.query
    if status_filter:
        query = query.filter_by(status=status_filter)
    items = query.order_by(Item.date_bought.desc().nullslast()).all()
    return jsonify([item.to_dict() for item in items])


@app.route('/api/items', methods=['POST'])
def create_item():
    data = request.json
    item = Item(
        date_bought=_parse_date(data.get('date_bought')),
        date_sold=_parse_date(data.get('date_sold')),
        description=data['description'],
        cost=float(data.get('cost', 0) or 0),
        listing_price=_parse_float(data.get('listing_price')),
        sold_for=_parse_float(data.get('sold_for')),
        status=data.get('status', 'Listed'),
        notes=data.get('notes'),
    )
    db.session.add(item)
    db.session.commit()
    return jsonify(item.to_dict()), 201


@app.route('/api/items/<int:item_id>', methods=['GET'])
def get_item(item_id):
    item = Item.query.get_or_404(item_id)
    return jsonify(item.to_dict())


@app.route('/api/items/<int:item_id>', methods=['PUT'])
def update_item(item_id):
    item = Item.query.get_or_404(item_id)
    data = request.json
    item.date_bought = _parse_date(data.get('date_bought'))
    item.date_sold = _parse_date(data.get('date_sold'))
    item.description = data['description']
    item.cost = float(data.get('cost', 0) or 0)
    item.listing_price = _parse_float(data.get('listing_price'))
    item.sold_for = _parse_float(data.get('sold_for'))
    item.status = data.get('status', item.status)
    item.notes = data.get('notes')
    db.session.commit()
    return jsonify(item.to_dict())


@app.route('/api/items/<int:item_id>', methods=['DELETE'])
def delete_item(item_id):
    item = Item.query.get_or_404(item_id)
    db.session.delete(item)
    db.session.commit()
    return jsonify({'message': 'Item deleted'})


@app.route('/api/stats', methods=['GET'])
def get_stats():
    all_items = Item.query.all()
    sold_items = [i for i in all_items if i.status == 'Sold']
    listed_items = [i for i in all_items if i.status == 'Listed']

    total_spent = sum(i.cost or 0 for i in all_items)
    total_revenue = sum(i.sold_for or 0 for i in sold_items)
    total_profit = sum(i.actual_profit or 0 for i in sold_items)
    current_invested = sum(i.cost or 0 for i in listed_items)
    predicted_profit = sum(i.predicted_profit or 0 for i in listed_items)

    avg_margin = 0
    margins = [i.actual_margin for i in sold_items if i.actual_margin is not None]
    if margins:
        avg_margin = round(sum(margins) / len(margins), 4)

    avg_days = 0
    days_list = [i.days_to_sell for i in sold_items if i.days_to_sell is not None and i.days_to_sell > 0]
    if days_list:
        avg_days = round(sum(days_list) / len(days_list), 1)

    # Monthly profit breakdown
    monthly = {}
    for item in sold_items:
        if item.date_sold and item.actual_profit is not None:
            key = item.date_sold.strftime('%Y-%m')
            monthly[key] = monthly.get(key, 0) + item.actual_profit
    monthly_sorted = sorted(monthly.items())

    # Top items by profit
    top_items = sorted(
        [i for i in sold_items if i.actual_profit is not None],
        key=lambda x: x.actual_profit,
        reverse=True
    )[:10]

    # Top items by profit/day
    top_by_efficiency = sorted(
        [i for i in sold_items if i.profit_per_day is not None and i.profit_per_day > 0],
        key=lambda x: x.profit_per_day,
        reverse=True
    )[:10]

    return jsonify({
        'total_items': len(all_items),
        'sold_count': len(sold_items),
        'listed_count': len(listed_items),
        'total_spent': round(total_spent, 2),
        'total_revenue': round(total_revenue, 2),
        'total_profit': round(total_profit, 2),
        'current_invested': round(current_invested, 2),
        'predicted_profit': round(predicted_profit, 2),
        'avg_margin': avg_margin,
        'avg_days_to_sell': avg_days,
        'monthly_profit': [{'month': m, 'profit': round(p, 2)} for m, p in monthly_sorted],
        'top_items': [{'description': i.description, 'profit': i.actual_profit} for i in top_items],
        'top_by_efficiency': [
            {'description': i.description, 'profit_per_day': i.profit_per_day, 'days': i.days_to_sell}
            for i in top_by_efficiency
        ],
    })


@app.route('/api/tax-export', methods=['GET'])
def generate_tax_pdf():
    start = _parse_date(request.args.get('start_date'))
    end = _parse_date(request.args.get('end_date'))
    include_listed = request.args.get('include_listed') == '1'

    # Query sold items in date range
    query = Item.query.filter_by(status='Sold')
    if start:
        query = query.filter(Item.date_sold >= start)
    if end:
        query = query.filter(Item.date_sold <= end)
    sold_items = query.order_by(Item.date_sold).all()

    listed_items = []
    if include_listed:
        listed_items = Item.query.filter_by(status='Listed').all()

    # Build PDF
    buf = io.BytesIO()
    doc = SimpleDocTemplate(buf, pagesize=letter, topMargin=0.5*inch, bottomMargin=0.5*inch)
    styles = getSampleStyleSheet()
    elements = []

    # Title
    title_style = ParagraphStyle('Title', parent=styles['Title'], fontSize=20, spaceAfter=6)
    elements.append(Paragraph("Onna's Flips - Tax Report", title_style))

    date_range = ""
    if start and end:
        date_range = f"{start.strftime('%m/%d/%Y')} - {end.strftime('%m/%d/%Y')}"
    elif start:
        date_range = f"From {start.strftime('%m/%d/%Y')}"
    elif end:
        date_range = f"Through {end.strftime('%m/%d/%Y')}"
    else:
        date_range = "All dates"
    elements.append(Paragraph(f"Period: {date_range}", styles['Normal']))
    elements.append(Paragraph(f"Generated: {datetime.now().strftime('%m/%d/%Y')}", styles['Normal']))
    elements.append(Spacer(1, 0.3*inch))

    # Summary box
    total_cost = sum(i.cost or 0 for i in sold_items)
    total_revenue = sum(i.sold_for or 0 for i in sold_items)
    total_profit = sum(i.actual_profit or 0 for i in sold_items)

    predicted_profit = sum(i.predicted_profit or 0 for i in listed_items)

    summary_data = [
        ['SUMMARY', '', ''],
        ['Items Sold', 'Total Revenue', 'Total Profit'],
        [str(len(sold_items)), f'${total_revenue:,.2f}', f'${total_profit:,.2f}'],
        ['Cost of Goods', 'Avg Margin', 'Predicted Profit (Listed)'],
    ]

    margins = [i.actual_margin for i in sold_items if i.actual_margin is not None]
    avg_margin = f'{sum(margins)/len(margins)*100:.0f}%' if margins else 'N/A'
    summary_data.append([f'${total_cost:,.2f}', avg_margin, f'${predicted_profit:,.2f}'])

    summary_table = Table(summary_data, colWidths=[2.2*inch, 2.2*inch, 2.2*inch])
    summary_table.setStyle(TableStyle([
        ('SPAN', (0, 0), (2, 0)),
        ('BACKGROUND', (0, 0), (2, 0), colors.HexColor('#1a1a2e')),
        ('TEXTCOLOR', (0, 0), (2, 0), colors.white),
        ('FONTNAME', (0, 0), (2, 0), 'Helvetica-Bold'),
        ('FONTSIZE', (0, 0), (2, 0), 14),
        ('ALIGN', (0, 0), (-1, -1), 'CENTER'),
        ('BACKGROUND', (0, 1), (-1, 1), colors.HexColor('#e8f5e9')),
        ('BACKGROUND', (0, 3), (-1, 3), colors.HexColor('#e8f5e9')),
        ('FONTNAME', (0, 1), (-1, 1), 'Helvetica-Bold'),
        ('FONTNAME', (0, 3), (-1, 3), 'Helvetica-Bold'),
        ('FONTSIZE', (0, 1), (-1, -1), 10),
        ('FONTSIZE', (0, 2), (-1, 2), 14),
        ('FONTSIZE', (0, 4), (-1, 4), 14),
        ('BOX', (0, 0), (-1, -1), 1, colors.HexColor('#1a1a2e')),
        ('GRID', (0, 1), (-1, -1), 0.5, colors.lightgrey),
        ('TOPPADDING', (0, 0), (-1, -1), 8),
        ('BOTTOMPADDING', (0, 0), (-1, -1), 8),
    ]))
    elements.append(summary_table)
    elements.append(Spacer(1, 0.3*inch))

    # Sold items table
    if sold_items:
        elements.append(Paragraph("Sold Items", styles['Heading2']))
        elements.append(Spacer(1, 0.1*inch))

        table_data = [['Date Sold', 'Date Bought', 'Item', 'Cost', 'Sold For', 'Profit']]
        for item in sold_items:
            table_data.append([
                item.date_sold.strftime('%m/%d/%y') if item.date_sold else '-',
                item.date_bought.strftime('%m/%d/%y') if item.date_bought else '-',
                item.description[:35],
                f'${item.cost:,.2f}' if item.cost else '$0',
                f'${item.sold_for:,.2f}' if item.sold_for else '-',
                f'${item.actual_profit:,.2f}' if item.actual_profit is not None else '-',
            ])

        # Totals row
        table_data.append(['', '', 'TOTALS', f'${total_cost:,.2f}', f'${total_revenue:,.2f}', f'${total_profit:,.2f}'])

        items_table = Table(table_data, colWidths=[0.8*inch, 0.8*inch, 2.3*inch, 0.85*inch, 0.85*inch, 0.85*inch])
        items_table.setStyle(TableStyle([
            ('BACKGROUND', (0, 0), (-1, 0), colors.HexColor('#1a1a2e')),
            ('TEXTCOLOR', (0, 0), (-1, 0), colors.white),
            ('FONTNAME', (0, 0), (-1, 0), 'Helvetica-Bold'),
            ('FONTSIZE', (0, 0), (-1, 0), 9),
            ('FONTSIZE', (0, 1), (-1, -1), 8),
            ('ALIGN', (3, 0), (-1, -1), 'RIGHT'),
            ('GRID', (0, 0), (-1, -1), 0.5, colors.lightgrey),
            ('ROWBACKGROUNDS', (0, 1), (-1, -2), [colors.white, colors.HexColor('#f8f9fa')]),
            ('BACKGROUND', (0, -1), (-1, -1), colors.HexColor('#e8f5e9')),
            ('FONTNAME', (0, -1), (-1, -1), 'Helvetica-Bold'),
            ('TOPPADDING', (0, 0), (-1, -1), 4),
            ('BOTTOMPADDING', (0, 0), (-1, -1), 4),
        ]))
        elements.append(items_table)

    # Monthly Breakout Analysis
    if sold_items:
        elements.append(Spacer(1, 0.3*inch))
        elements.append(Paragraph("Monthly Breakout Analysis", styles['Heading2']))
        elements.append(Spacer(1, 0.1*inch))

        monthly = {}
        for item in sold_items:
            if item.date_sold:
                key = item.date_sold.strftime('%Y-%m')
                if key not in monthly:
                    monthly[key] = {'items': 0, 'cost': 0, 'revenue': 0, 'profit': 0}
                monthly[key]['items'] += 1
                monthly[key]['cost'] += item.cost or 0
                monthly[key]['revenue'] += item.sold_for or 0
                monthly[key]['profit'] += item.actual_profit or 0

        month_names = {
            '01': 'Jan', '02': 'Feb', '03': 'Mar', '04': 'Apr', '05': 'May', '06': 'Jun',
            '07': 'Jul', '08': 'Aug', '09': 'Sep', '10': 'Oct', '11': 'Nov', '12': 'Dec',
        }

        month_data = [['Month', 'Items Sold', 'Cost', 'Revenue', 'Profit', 'Margin']]
        for key in sorted(monthly.keys()):
            m = monthly[key]
            y, mo = key.split('-')
            margin = f'{m["profit"]/m["revenue"]*100:.0f}%' if m['revenue'] else 'N/A'
            month_data.append([
                f'{month_names[mo]} {y}',
                str(m['items']),
                f'${m["cost"]:,.2f}',
                f'${m["revenue"]:,.2f}',
                f'${m["profit"]:,.2f}',
                margin,
            ])
        # Totals
        month_data.append([
            'TOTAL', str(len(sold_items)),
            f'${total_cost:,.2f}', f'${total_revenue:,.2f}', f'${total_profit:,.2f}',
            f'{total_profit/total_revenue*100:.0f}%' if total_revenue else 'N/A',
        ])

        month_table = Table(month_data, colWidths=[1.1*inch, 0.8*inch, 1*inch, 1*inch, 1*inch, 0.8*inch])
        month_table.setStyle(TableStyle([
            ('BACKGROUND', (0, 0), (-1, 0), colors.HexColor('#1a1a2e')),
            ('TEXTCOLOR', (0, 0), (-1, 0), colors.white),
            ('FONTNAME', (0, 0), (-1, 0), 'Helvetica-Bold'),
            ('FONTSIZE', (0, 0), (-1, 0), 9),
            ('FONTSIZE', (0, 1), (-1, -1), 9),
            ('ALIGN', (1, 0), (-1, -1), 'RIGHT'),
            ('GRID', (0, 0), (-1, -1), 0.5, colors.lightgrey),
            ('ROWBACKGROUNDS', (0, 1), (-1, -2), [colors.white, colors.HexColor('#f8f9fa')]),
            ('BACKGROUND', (0, -1), (-1, -1), colors.HexColor('#e8f5e9')),
            ('FONTNAME', (0, -1), (-1, -1), 'Helvetica-Bold'),
            ('TOPPADDING', (0, 0), (-1, -1), 5),
            ('BOTTOMPADDING', (0, 0), (-1, -1), 5),
        ]))
        elements.append(month_table)

    # Listed items (unsold inventory)
    if listed_items:
        elements.append(Spacer(1, 0.3*inch))
        elements.append(Paragraph("Unsold Inventory", styles['Heading2']))
        elements.append(Spacer(1, 0.1*inch))

        inv_data = [['Date Bought', 'Item', 'Cost', 'Listing Price', 'Est. Profit']]
        inv_total_cost = 0
        for item in listed_items:
            inv_total_cost += item.cost or 0
            inv_data.append([
                item.date_bought.strftime('%m/%d/%y') if item.date_bought else '-',
                item.description[:40],
                f'${item.cost:,.2f}' if item.cost else '$0',
                f'${item.listing_price:,.2f}' if item.listing_price else '-',
                f'${item.predicted_profit:,.2f}' if item.predicted_profit else '-',
            ])
        inv_data.append(['', 'TOTAL INVESTED', f'${inv_total_cost:,.2f}', '', ''])

        inv_table = Table(inv_data, colWidths=[0.9*inch, 2.5*inch, 0.9*inch, 1*inch, 0.9*inch])
        inv_table.setStyle(TableStyle([
            ('BACKGROUND', (0, 0), (-1, 0), colors.HexColor('#6c757d')),
            ('TEXTCOLOR', (0, 0), (-1, 0), colors.white),
            ('FONTNAME', (0, 0), (-1, 0), 'Helvetica-Bold'),
            ('FONTSIZE', (0, 0), (-1, 0), 9),
            ('FONTSIZE', (0, 1), (-1, -1), 8),
            ('ALIGN', (2, 0), (-1, -1), 'RIGHT'),
            ('GRID', (0, 0), (-1, -1), 0.5, colors.lightgrey),
            ('ROWBACKGROUNDS', (0, 1), (-1, -2), [colors.white, colors.HexColor('#f8f9fa')]),
            ('BACKGROUND', (0, -1), (-1, -1), colors.HexColor('#fff3cd')),
            ('FONTNAME', (0, -1), (-1, -1), 'Helvetica-Bold'),
            ('TOPPADDING', (0, 0), (-1, -1), 4),
            ('BOTTOMPADDING', (0, 0), (-1, -1), 4),
        ]))
        elements.append(inv_table)

    # === WILD ANALYSIS ===
    if sold_items:
        elements.append(Spacer(1, 0.3*inch))
        wild_title = ParagraphStyle('WildTitle', parent=styles['Heading1'], fontSize=16,
                                     textColor=colors.HexColor('#1a1a2e'))
        elements.append(Paragraph("Wild Analysis", wild_title))
        elements.append(Spacer(1, 0.1*inch))

        # --- 1. Best & Worst Flips ---
        elements.append(Paragraph("Best & Worst Flips", styles['Heading3']))
        sorted_by_profit = sorted([i for i in sold_items if i.actual_profit is not None],
                                   key=lambda x: x.actual_profit, reverse=True)
        bw_data = [['Rank', 'Item', 'Cost', 'Sold For', 'Profit', 'Margin']]
        for i, item in enumerate(sorted_by_profit[:5], 1):
            margin = f'{item.actual_margin*100:.0f}%' if item.actual_margin else 'N/A'
            bw_data.append([f'#{i}', item.description[:30], f'${item.cost:,.0f}',
                           f'${item.sold_for:,.0f}', f'${item.actual_profit:,.0f}', margin])
        if len(sorted_by_profit) > 5:
            bw_data.append(['', '', '', '', '', ''])
            for i, item in enumerate(sorted_by_profit[-3:], 1):
                margin = f'{item.actual_margin*100:.0f}%' if item.actual_margin else 'N/A'
                bw_data.append([f'Worst #{i}', item.description[:30], f'${item.cost:,.0f}',
                               f'${item.sold_for:,.0f}' if item.sold_for else '-',
                               f'${item.actual_profit:,.0f}' if item.actual_profit else '-', margin])

        bw_table = Table(bw_data, colWidths=[0.7*inch, 2*inch, 0.7*inch, 0.8*inch, 0.7*inch, 0.7*inch])
        bw_table.setStyle(TableStyle([
            ('BACKGROUND', (0, 0), (-1, 0), colors.HexColor('#28a745')),
            ('TEXTCOLOR', (0, 0), (-1, 0), colors.white),
            ('FONTNAME', (0, 0), (-1, 0), 'Helvetica-Bold'),
            ('FONTSIZE', (0, 0), (-1, -1), 8),
            ('ALIGN', (2, 0), (-1, -1), 'RIGHT'),
            ('GRID', (0, 0), (-1, -1), 0.5, colors.lightgrey),
            ('TOPPADDING', (0, 0), (-1, -1), 3),
            ('BOTTOMPADDING', (0, 0), (-1, -1), 3),
        ]))
        elements.append(bw_table)
        elements.append(Spacer(1, 0.2*inch))

        # --- 2. Speed Demons (Fastest Flips) ---
        elements.append(Paragraph("Speed Demons - Fastest Flips", styles['Heading3']))
        fast_flips = sorted([i for i in sold_items if i.days_to_sell is not None and i.days_to_sell >= 0],
                           key=lambda x: x.days_to_sell)
        speed_data = [['Item', 'Days', 'Profit', '$/Day']]
        for item in fast_flips[:7]:
            ppd = f'${item.profit_per_day:,.0f}' if item.profit_per_day else 'N/A'
            speed_data.append([item.description[:35], str(item.days_to_sell),
                              f'${item.actual_profit:,.0f}' if item.actual_profit else '-', ppd])
        speed_table = Table(speed_data, colWidths=[2.5*inch, 0.7*inch, 0.8*inch, 0.8*inch])
        speed_table.setStyle(TableStyle([
            ('BACKGROUND', (0, 0), (-1, 0), colors.HexColor('#17a2b8')),
            ('TEXTCOLOR', (0, 0), (-1, 0), colors.white),
            ('FONTNAME', (0, 0), (-1, 0), 'Helvetica-Bold'),
            ('FONTSIZE', (0, 0), (-1, -1), 8),
            ('ALIGN', (1, 0), (-1, -1), 'RIGHT'),
            ('GRID', (0, 0), (-1, -1), 0.5, colors.lightgrey),
            ('TOPPADDING', (0, 0), (-1, -1), 3),
            ('BOTTOMPADDING', (0, 0), (-1, -1), 3),
        ]))
        elements.append(speed_table)
        elements.append(Spacer(1, 0.2*inch))

        # --- 3. ROI Champions (Best Return on Investment) ---
        elements.append(Paragraph("ROI Champions - Best Return on Investment", styles['Heading3']))
        roi_items = sorted([i for i in sold_items if i.cost and i.cost > 0 and i.actual_profit is not None],
                          key=lambda x: x.actual_profit / x.cost, reverse=True)
        roi_data = [['Item', 'Cost', 'Sold For', 'ROI %']]
        for item in roi_items[:7]:
            roi = (item.actual_profit / item.cost) * 100
            roi_data.append([item.description[:35], f'${item.cost:,.0f}',
                           f'${item.sold_for:,.0f}', f'{roi:,.0f}%'])
        roi_table = Table(roi_data, colWidths=[2.5*inch, 0.8*inch, 0.8*inch, 0.8*inch])
        roi_table.setStyle(TableStyle([
            ('BACKGROUND', (0, 0), (-1, 0), colors.HexColor('#6f42c1')),
            ('TEXTCOLOR', (0, 0), (-1, 0), colors.white),
            ('FONTNAME', (0, 0), (-1, 0), 'Helvetica-Bold'),
            ('FONTSIZE', (0, 0), (-1, -1), 8),
            ('ALIGN', (1, 0), (-1, -1), 'RIGHT'),
            ('GRID', (0, 0), (-1, -1), 0.5, colors.lightgrey),
            ('TOPPADDING', (0, 0), (-1, -1), 3),
            ('BOTTOMPADDING', (0, 0), (-1, -1), 3),
        ]))
        elements.append(roi_table)
        elements.append(Spacer(1, 0.2*inch))

        # --- 4. Price Bracket Analysis ---
        elements.append(Paragraph("Price Bracket Analysis - Where the Money Is", styles['Heading3']))
        brackets = {'$0-25': (0, 25), '$26-50': (26, 50), '$51-100': (51, 100),
                     '$101-200': (101, 200), '$201-300': (201, 300), '$300+': (301, 99999)}
        bracket_data = [['Sale Price Range', 'Items', 'Avg Profit', 'Avg Days', 'Total Profit']]
        for label, (lo, hi) in brackets.items():
            b_items = [i for i in sold_items if i.sold_for and lo <= i.sold_for <= hi]
            if b_items:
                avg_p = sum(i.actual_profit or 0 for i in b_items) / len(b_items)
                d_list = [i.days_to_sell for i in b_items if i.days_to_sell and i.days_to_sell > 0]
                avg_d = f'{sum(d_list)/len(d_list):.0f}' if d_list else 'N/A'
                tot_p = sum(i.actual_profit or 0 for i in b_items)
                bracket_data.append([label, str(len(b_items)), f'${avg_p:,.0f}', avg_d, f'${tot_p:,.0f}'])
        bracket_table = Table(bracket_data, colWidths=[1.2*inch, 0.7*inch, 1*inch, 0.8*inch, 1*inch])
        bracket_table.setStyle(TableStyle([
            ('BACKGROUND', (0, 0), (-1, 0), colors.HexColor('#fd7e14')),
            ('TEXTCOLOR', (0, 0), (-1, 0), colors.white),
            ('FONTNAME', (0, 0), (-1, 0), 'Helvetica-Bold'),
            ('FONTSIZE', (0, 0), (-1, -1), 8),
            ('ALIGN', (1, 0), (-1, -1), 'RIGHT'),
            ('GRID', (0, 0), (-1, -1), 0.5, colors.lightgrey),
            ('TOPPADDING', (0, 0), (-1, -1), 3),
            ('BOTTOMPADDING', (0, 0), (-1, -1), 3),
        ]))
        elements.append(bracket_table)
        elements.append(Spacer(1, 0.2*inch))

        # --- 5. Day of Week Analysis ---
        elements.append(Paragraph("Best Day to Sell", styles['Heading3']))
        day_names = ['Monday', 'Tuesday', 'Wednesday', 'Thursday', 'Friday', 'Saturday', 'Sunday']
        day_stats = {d: {'count': 0, 'profit': 0} for d in day_names}
        for item in sold_items:
            if item.date_sold:
                day = day_names[item.date_sold.weekday()]
                day_stats[day]['count'] += 1
                day_stats[day]['profit'] += item.actual_profit or 0
        dow_data = [['Day', 'Sales', 'Total Profit', 'Avg Profit']]
        for day in day_names:
            s = day_stats[day]
            if s['count'] > 0:
                dow_data.append([day, str(s['count']), f'${s["profit"]:,.0f}',
                               f'${s["profit"]/s["count"]:,.0f}'])
        dow_table = Table(dow_data, colWidths=[1.2*inch, 0.8*inch, 1*inch, 1*inch])
        dow_table.setStyle(TableStyle([
            ('BACKGROUND', (0, 0), (-1, 0), colors.HexColor('#e83e8c')),
            ('TEXTCOLOR', (0, 0), (-1, 0), colors.white),
            ('FONTNAME', (0, 0), (-1, 0), 'Helvetica-Bold'),
            ('FONTSIZE', (0, 0), (-1, -1), 8),
            ('ALIGN', (1, 0), (-1, -1), 'RIGHT'),
            ('GRID', (0, 0), (-1, -1), 0.5, colors.lightgrey),
            ('TOPPADDING', (0, 0), (-1, -1), 3),
            ('BOTTOMPADDING', (0, 0), (-1, -1), 3),
        ]))
        elements.append(dow_table)
        elements.append(Spacer(1, 0.2*inch))

        # --- 6. Fun Facts / Business Scorecard ---
        elements.append(Paragraph("Business Scorecard", styles['Heading3']))

        total_days_in_biz = 0
        all_bought = [i.date_bought for i in sold_items if i.date_bought]
        all_sold_dates = [i.date_sold for i in sold_items if i.date_sold]
        if all_bought and all_sold_dates:
            total_days_in_biz = (max(all_sold_dates) - min(all_bought)).days or 1

        profit_per_week = total_profit / max(total_days_in_biz / 7, 1)
        avg_flip_cost = total_cost / len(sold_items) if sold_items else 0
        biggest_single = max(sold_items, key=lambda x: x.actual_profit or 0)
        fastest = min([i for i in sold_items if i.days_to_sell and i.days_to_sell > 0],
                     key=lambda x: x.days_to_sell, default=None)

        facts = [
            ['Metric', 'Value'],
            ['Days in Business', str(total_days_in_biz)],
            ['Profit per Week', f'${profit_per_week:,.0f}'],
            ['Avg Cost per Flip', f'${avg_flip_cost:,.0f}'],
            ['Avg Sale Price', f'${total_revenue/len(sold_items):,.0f}' if sold_items else 'N/A'],
            ['Biggest Win', f'{biggest_single.description} (${biggest_single.actual_profit:,.0f})' if biggest_single else 'N/A'],
            ['Fastest Flip', f'{fastest.description} ({fastest.days_to_sell}d, ${fastest.actual_profit:,.0f})' if fastest else 'N/A'],
            ['Items Flipped per Week', f'{len(sold_items)/max(total_days_in_biz/7, 1):.1f}'],
            ['Money Multiplier', f'{total_revenue/total_cost:.1f}x' if total_cost else 'N/A'],
            ['Profit if Reinvested at Avg Margin',
             f'${total_profit * (1 + sum(margins)/len(margins)):,.0f}' if margins else 'N/A'],
        ]
        facts_table = Table(facts, colWidths=[2.5*inch, 3.5*inch])
        facts_table.setStyle(TableStyle([
            ('BACKGROUND', (0, 0), (-1, 0), colors.HexColor('#1a1a2e')),
            ('TEXTCOLOR', (0, 0), (-1, 0), colors.white),
            ('FONTNAME', (0, 0), (-1, 0), 'Helvetica-Bold'),
            ('FONTNAME', (0, 1), (0, -1), 'Helvetica-Bold'),
            ('FONTSIZE', (0, 0), (-1, -1), 9),
            ('GRID', (0, 0), (-1, -1), 0.5, colors.lightgrey),
            ('ROWBACKGROUNDS', (0, 1), (-1, -1), [colors.white, colors.HexColor('#f8f9fa')]),
            ('TOPPADDING', (0, 0), (-1, -1), 4),
            ('BOTTOMPADDING', (0, 0), (-1, -1), 4),
        ]))
        elements.append(facts_table)

    doc.build(elements)
    buf.seek(0)

    filename = f"OnnaFlips_TaxReport_{datetime.now().strftime('%Y%m%d')}.pdf"
    return send_file(buf, mimetype='application/pdf', as_attachment=True, download_name=filename)


def _parse_date(val):
    if not val:
        return None
    if isinstance(val, (date, datetime)):
        return val if isinstance(val, date) else val.date()
    try:
        return datetime.strptime(val, '%Y-%m-%d').date()
    except (ValueError, TypeError):
        return None


def _parse_float(val):
    if val is None or val == '':
        return None
    try:
        return float(val)
    except (ValueError, TypeError):
        return None


with app.app_context():
    db.create_all()

if __name__ == '__main__':
    app.run(debug=True, port=5000)

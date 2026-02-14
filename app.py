import os
from datetime import datetime, date
from flask import Flask, render_template, request, jsonify, redirect, url_for
from flask_sqlalchemy import SQLAlchemy

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

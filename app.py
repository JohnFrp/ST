import os
import io
from flask import Flask, render_template, request, jsonify, redirect, url_for, flash
import pandas as pd
from datetime import datetime
from urllib.parse import quote_plus
from dotenv import load_dotenv

# Load environment variables
load_dotenv()

# Database configuration
password = "Johnh4k3r"
encoded_password = quote_plus(password)
DATABASE_URI = f'postgresql://postgres.dtcbyjnvggptyerrbxwp:{encoded_password}@aws-1-ap-southeast-1.pooler.supabase.com:6543/postgres'

def create_app():
    app = Flask(__name__)
    app.config['SECRET_KEY'] = os.environ.get('SECRET_KEY', 'pharmacy-pos-secret-key')
    
    # Configure the database URI
    app.config["SQLALCHEMY_DATABASE_URI"] = DATABASE_URI
    app.config["SQLALCHEMY_TRACK_MODIFICATIONS"] = False

    # Initialize extensions within app context
    from flask_sqlalchemy import SQLAlchemy
    
    db = SQLAlchemy()

    # Define models
    class Stock(db.Model):
        __tablename__ = 'stocks'
        
        id = db.Column(db.Integer, primary_key=True)
        name = db.Column(db.String(255), nullable=False)
        generic_name = db.Column(db.String(255))        
        category = db.Column(db.String(255))
        buy_price = db.Column(db.Float, nullable=False)
        sell_price = db.Column(db.Float, nullable=False)
        stock_quantity = db.Column(db.Integer, nullable=False)
        expiry_date = db.Column(db.Date)
        created_at = db.Column(db.DateTime, default=datetime.utcnow)
        
        def __repr__(self):
            return f'<Stock {self.name}>'
        
        def to_dict(self):
            return {
                'id': self.id,
                'name': self.name,
                'generic_name': self.generic_name,                
                'category': self.category,
                'buy_price': self.buy_price,
                'sell_price': self.sell_price,
                'stock_quantity': self.stock_quantity,
                'expiry_date': self.expiry_date.strftime('%Y-%m-%d') if self.expiry_date else None,
                'profit_per_unit': round(self.sell_price - self.buy_price, 2),
                'total_value': round(self.sell_price * self.stock_quantity, 2),
                'created_at': self.created_at.strftime('%Y-%m-%d %H:%M:%S') if self.created_at else None
            }

    # Initialize the app with extensions
    db.init_app(app)

    # Create tables before first request
    #@app.before_first_request
    def create_tables():
        with app.app_context():
            try:
                db.create_all()
                print("✅ Database tables created successfully!")
            except Exception as e:
                print(f"❌ Error creating tables: {e}")

    @app.route('/')
    def index():
        return redirect(url_for('show_stock'))

    @app.route('/upload', methods=['GET', 'POST'])
    def upload_file():
        if request.method == 'POST':
            if 'file' not in request.files:
                flash('No file selected')
                return redirect(request.url)
            
            file = request.files['file']
            if file.filename == '':
                flash('No file selected')
                return redirect(request.url)
            
            if file and allowed_file(file.filename):
                try:
                    # Read Excel file directly from memory
                    excel_data = file.read()
                    
                    # Use pandas to read the Excel file from bytes
                    df = pd.read_excel(io.BytesIO(excel_data))
                    
                    # Validate required columns
                    required_columns = ['name', 'buy_price', 'sell_price', 'stock_quantity']
                    missing_columns = [col for col in required_columns if col not in df.columns]
                    
                    if missing_columns:
                        flash(f'Missing required columns: {", ".join(missing_columns)}')
                        return redirect(request.url)
                    
                    # Process data
                    stocks_added = 0
                    stocks_updated = 0
                    for _, row in df.iterrows():
                        # Handle expiry date conversion
                        expiry_date = None
                        if 'expiry_date' in df.columns and pd.notna(row['expiry_date']):
                            if isinstance(row['expiry_date'], str):
                                try:
                                    expiry_date = datetime.strptime(row['expiry_date'], '%Y-%m-%d').date()
                                except ValueError:
                                    # Try different date formats
                                    try:
                                        expiry_date = datetime.strptime(row['expiry_date'], '%d/%m/%Y').date()
                                    except ValueError:
                                        # Try parsing as datetime object
                                        try:
                                            expiry_date = row['expiry_date'].date()
                                        except:
                                            expiry_date = None
                            else:
                                try:
                                    expiry_date = row['expiry_date'].date()
                                except AttributeError:
                                    # Handle pandas Timestamp
                                    try:
                                        expiry_date = row['expiry_date'].to_pydatetime().date()
                                    except:
                                        expiry_date = None
                        
                        # Check if stock already exists
                        existing_stock = Stock.query.filter_by(name=row['name']).first()
                        if existing_stock:
                            # Update existing stock
                            existing_stock.generic_name = row.get('generic_name', existing_stock.generic_name)                            
                            existing_stock.category = row.get('category', existing_stock.category)
                            existing_stock.buy_price = float(row['buy_price'])
                            existing_stock.sell_price = float(row['sell_price'])
                            existing_stock.stock_quantity = int(row['stock_quantity'])
                            existing_stock.expiry_date = expiry_date
                            stocks_updated += 1
                        else:
                            # Create new stock
                            stock = Stock(
                                name=row['name'],
                                generic_name=row.get('generic_name', ''),                                
                                category=row.get('category', ''),
                                buy_price=float(row['buy_price']),
                                sell_price=float(row['sell_price']),
                                stock_quantity=int(row['stock_quantity']),
                                expiry_date=expiry_date
                            )
                            db.session.add(stock)
                            stocks_added += 1
                    
                    db.session.commit()
                    
                    if stocks_updated > 0 and stocks_added > 0:
                        flash(f'Successfully imported {stocks_added} new stock items and updated {stocks_updated} existing items')
                    elif stocks_added > 0:
                        flash(f'Successfully imported {stocks_added} new stock items')
                    elif stocks_updated > 0:
                        flash(f'Successfully updated {stocks_updated} existing stock items')
                    else:
                        flash('No changes made to the database')
                    
                except Exception as e:
                    db.session.rollback()
                    flash(f'Error importing file: {str(e)}')
                    print(f"Error details: {e}")
                
                return redirect(url_for('show_stock'))
            else:
                flash('Please upload a valid Excel file (.xlsx or .xls)')
        
        return render_template('upload.html')

    @app.route('/stock')
    def show_stock():
        try:
            # Get filter parameters
            category_filter = request.args.get('category', '')
            search_query = request.args.get('search', '')
            
            # Base query
            query = Stock.query
            
            # Apply filters
            if category_filter:
                query = query.filter(Stock.category == category_filter)
            
            if search_query:
                query = query.filter(
                    (Stock.name.ilike(f'%{search_query}%')) |
                    (Stock.generic_name.ilike(f'%{search_query}%'))                   
                )
            
            stocks = query.order_by(Stock.name).all()
            
            # Get unique categories for filter dropdown
            categories = db.session.query(Stock.category).distinct().all()
            categories = [cat[0] for cat in categories if cat[0]]
            
            # Calculate summary statistics
            total_items = sum(stock.stock_quantity for stock in stocks)
            total_value = sum(stock.sell_price * stock.stock_quantity for stock in stocks)
            total_profit_potential = sum((stock.sell_price - stock.buy_price) * stock.stock_quantity for stock in stocks)
            
            return render_template('stock.html', 
                                 stocks=stocks,
                                 categories=categories,
                                 total_items=total_items,
                                 total_value=round(total_value, 2),
                                 total_profit_potential=round(total_profit_potential, 2),
                                 now=datetime.now())
        except Exception as e:
            flash(f'Error loading stock data: {str(e)}')
            return render_template('stock.html', 
                                 stocks=[],
                                 categories=[],
                                 total_items=0,
                                 total_value=0,
                                 total_profit_potential=0,
                                 now=datetime.now())

    @app.route('/api/stock')
    def api_stock():
        try:
            stocks = Stock.query.all()
            return jsonify([stock.to_dict() for stock in stocks])
        except Exception as e:
            return jsonify({'error': str(e)}), 500

    @app.route('/api/stock/<int:stock_id>')
    def api_stock_detail(stock_id):
        try:
            stock = Stock.query.get_or_404(stock_id)
            return jsonify(stock.to_dict())
        except Exception as e:
            return jsonify({'error': str(e)}), 500

    @app.route('/api/categories')
    def api_categories():
        try:
            categories = db.session.query(Stock.category).distinct().all()
            categories = [cat[0] for cat in categories if cat[0]]
            return jsonify(categories)
        except Exception as e:
            return jsonify({'error': str(e)}), 500

    @app.route('/clear-stock', methods=['POST'])
    def clear_stock():
        try:
            # Delete all stock records
            num_deleted = Stock.query.delete()
            db.session.commit()
            flash(f'Successfully cleared {num_deleted} stock items from database')
        except Exception as e:
            db.session.rollback()
            flash(f'Error clearing stock: {str(e)}')
        
        return redirect(url_for('show_stock'))

    @app.route('/init-db')
    def init_db():
        """Manual route to initialize database"""
        try:
            with app.app_context():
                db.create_all()
            flash('Database initialized successfully!')
        except Exception as e:
            flash(f'Error initializing database: {str(e)}')
        return redirect(url_for('show_stock'))

    def allowed_file(filename):
        return '.' in filename and \
               filename.rsplit('.', 1)[1].lower() in {'xlsx', 'xls'}

    return app

# Create app instance
app = create_app()

# Initialize database when app starts
with app.app_context():
    try:
        from flask_sqlalchemy import SQLAlchemy
        db = SQLAlchemy(app)
        db.create_all()
        print("✅ Database tables created successfully!")
    except Exception as e:
        print(f"❌ Error creating tables: {e}")

if __name__ == '__main__':
    app.run(debug=True)
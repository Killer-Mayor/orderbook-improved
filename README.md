# Orderbook Management System (Improved)

 A robust Flask-based order management system with Google Sheets integration, featuring improved architecture, comprehensive error handling, and production-ready code.

## Features

- **Order Management**: Add, view, and track orders with automatic GST calculation
- **Dispatch Tracking**: Record and monitor product dispatches with order-level granularity
- **Pivot Analytics**: Dynamic pivot tables for product/party analysis with multi-select filtering
- **Google Sheets Integration**: Seamless synchronization with Google Sheets
- **RESTful API**: JSON endpoints for programmatic access
- **Docker Support**: Containerized deployment ready

## Architecture Improvements

This version includes significant improvements over the original:

### Code Organization
- **Modular structure**: Separated routes, services, and models
- **Configuration management**: Centralized config with environment variable validation
- **Service layer**: Business logic isolated from routing
- **Error handling**: Comprehensive exception handling with proper logging

### Security Enhancements
- **Secrets management**: Proper .gitignore for credentials
- **Environment validation**: Required variables checked at startup
- **Input sanitization**: Form data validation

### Performance
- **Efficient data access**: Optimized Google Sheets queries
- **Error recovery**: Graceful degradation on failures
- **Logging**: Structured logging for debugging

## Quick Start

### Prerequisites

- Python 3.8+
- Google Cloud service account with Sheets API access
- Google Sheet with appropriate structure

### Local Development

1. **Clone the repository**
```bash
git clone https://github.com/Killer-Mayor/orderbook-improved.git
cd orderbook-improved
```

2. **Create virtual environment**
```bash
python -m venv venv
source venv/bin/activate  # On Windows: venv\Scripts\activate
```

3. **Install dependencies**
```bash
pip install -r requirements.txt
```

4. **Configure environment variables**

Create a `.env` file:
```env
SHEET_ID=your_google_sheet_id_here
FLASK_SECRET=your_random_secret_key_here
FLASK_ENV=development
LOG_LEVEL=INFO
```

5. **Add service account credentials**

Place your `service_account.json` file in the project root or set `SERVICE_ACCOUNT_JSON` environment variable with the JSON content.

6. **Run the application**
```bash
python run.py
```

Open http://127.0.0.1:5000 in your browser.

## Google Sheets Structure

The application expects the following worksheets:

- **orders**: Main order data (auto-created if missing)
  - Headers: Order Number, Date, Company, Product, Brand, Quantity, Price
- **products**: Product list (Column A: product names)
- **companies**: Company list (Column A: company names)
- **brands**: Brand list (Column A: brand names)
- **dispatch**: Dispatch records (auto-created)
  - Headers: Date, Company, Product, Quantity, Order Number

## API Endpoints

### Orders
- `GET /` - Main dashboard with order form
- `POST /submit` - Add new orders
- `GET /orders` - View recent orders
- `GET /api/products` - Get product list (JSON)
- `GET /api/companies` - Get company list (JSON)

### Dispatch
- `GET /dispatch` - Dispatch management interface
- `POST /dispatch` - Record dispatch (JSON)
- `GET /api/orders_by_product?product=<name>` - Orders by product
- `GET /api/orders_by_party?company=<name>` - Orders by company
- `GET /api/pivot_data?product_filter=<filter>&party_filter=<filter>` - Pivot data

### Health
- `GET /_health` - Health check endpoint

## Deployment

### Docker

```bash
docker build -t orderbook .
docker run -p 5000:5000 \
  -e SHEET_ID=your_sheet_id \
  -e FLASK_SECRET=your_secret \
  -e SERVICE_ACCOUNT_JSON="$(cat service_account.json)" \
  orderbook
```

### Vercel (Serverless)

1. Install Vercel CLI: `npm i -g vercel`
2. Set environment variables in Vercel dashboard
3. Deploy: `vercel --prod`

### Render

1. Connect repository to Render
2. Choose "Web Service" with Python environment
3. Set start command: `./start.sh`
4. Configure environment variables

## Environment Variables

| Variable | Required | Description |
|----------|----------|-------------|
| `SHEET_ID` | Yes | Google Sheets document ID |
| `FLASK_SECRET` | Yes | Secret key for Flask sessions |
| `SERVICE_ACCOUNT_JSON` | Yes* | Service account JSON (content or file path) |
| `FLASK_ENV` | No | Environment (development/production) |
| `LOG_LEVEL` | No | Logging level (DEBUG/INFO/WARNING/ERROR) |
| `PORT` | No | Server port (default: 5000) |

*Either provide `SERVICE_ACCOUNT_JSON` environment variable or place `service_account.json` file in project root.

## Development

### Project Structure

```
orderbook-improved/
├── app/
│   ├── __init__.py          # Flask app factory
│   ├── routes/              # Route blueprints
│   │   ├── __init__.py
│   │   ├── orders.py        # Order routes
│   │   ├── dispatch.py      # Dispatch routes
│   │   └── api.py           # API routes
│   ├── services/            # Business logic
│   │   ├── __init__.py
│   │   └── sheets_service.py # Google Sheets service
│   ├── models/              # Data models
│   │   └── __init__.py
│   └── utils/               # Utilities
│       ├── __init__.py
│       └── validators.py    # Input validation
├── config/
│   ├── __init__.py          # Configuration classes
│   └── example.env          # Example environment file
├── templates/               # HTML templates
├── static/                  # Static files (CSS, JS)
├── tests/                   # Test suite
├── run.py                   # Application entry point
├── requirements.txt         # Python dependencies
├── Dockerfile              # Docker configuration
├── start.sh                # Startup script
└── README.md
```

### Running Tests

```bash
pytest tests/
```

## Contributing

1. Fork the repository
2. Create a feature branch
3. Make your changes
4. Add tests
5. Submit a pull request

## License

MIT License - See LICENSE file for details

## Support

For issues and questions, please open a GitHub issue.

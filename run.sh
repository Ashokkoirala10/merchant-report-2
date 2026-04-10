#!/bin/bash
# Quick start script for NRB Merchant Report Generator
echo "================================"
echo " NRB Merchant Report Generator"
echo "================================"
echo ""

# Install deps if needed
pip install -r requirements.txt -q

# Run migrations
python manage.py migrate --run-syncdb -q

echo "✅ Ready! Starting server..."
echo "   Open: http://127.0.0.1:8000"
echo ""
python manage.py runserver

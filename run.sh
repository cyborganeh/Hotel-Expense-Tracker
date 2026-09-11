#!/bin/bash
# run.sh - Launch the Hotel Spending Tracker & Separator App
cd "$(dirname "$0")"

echo "========================================================="
echo "   Hotel Santika Depok - Expense Separator & Tracker     "
echo "========================================================="
echo "Starting Streamlit web server on http://localhost:8501..."
echo ""

uv run streamlit run app.py --server.port 8501 --server.headless true

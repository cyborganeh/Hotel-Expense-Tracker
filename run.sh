#!/bin/bash
# run.sh - Launch the Hotel Spending Tracker & Separator App
cd "$(dirname "$0")"

PORT=$(python - <<'PY'
import socket
for port in range(8501, 8511):
    s = socket.socket()
    s.setsockopt(socket.SOL_SOCKET, socket.SO_REUSEADDR, 1)
    try:
        s.bind(("127.0.0.1", port))
        print(port)
        break
    except OSError:
        pass
    finally:
        s.close()
else:
    raise SystemExit(1)
PY
)

echo "========================================================="
echo "   Hotel - Expense Separator & Tracker     "
echo "========================================================="
echo "Starting Streamlit web server on http://localhost:${PORT}..."
echo ""

if [ -f .venv/bin/activate ]; then
  . .venv/bin/activate
fi

if command -v streamlit >/dev/null 2>&1; then
  exec streamlit run app.py --server.port "$PORT" --server.address 0.0.0.0 --server.headless true
fi

if command -v uv >/dev/null 2>&1; then
  exec uv run streamlit run app.py --server.port "$PORT" --server.address 0.0.0.0 --server.headless true
fi

python -m streamlit run app.py --server.port "$PORT" --server.address 0.0.0.0 --server.headless true

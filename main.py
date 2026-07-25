import os
import sys

# Add project root and backend directory to sys.path automatically
root_dir = os.path.dirname(os.path.abspath(__file__))
backend_dir = os.path.join(root_dir, "backend")

if root_dir not in sys.path:
    sys.path.insert(0, root_dir)
if os.path.exists(backend_dir) and backend_dir not in sys.path:
    sys.path.insert(0, backend_dir)

# Import app from backend.app or app with fallback
try:
    from backend.app import app
except Exception as e:
    print(f"Importing backend.app failed ({e}), trying direct app import...")
    try:
        from app import app
    except Exception as e2:
        print(f"Direct app import failed: {e2}")
        raise e2

if __name__ == "__main__":
    import uvicorn
    port = int(os.environ.get("PORT", 10000))
    print(f"Starting server on port {port}...")
    uvicorn.run(app, host="0.0.0.0", port=port)

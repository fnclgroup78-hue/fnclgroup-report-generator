import sys
import os
import traceback

print("=== STARTING FNCL SERVER DIAGNOSTIC ===")
print("Current Working Directory:", os.getcwd())

# Ensure current directory and backend directory are in sys.path
cwd = os.getcwd()
backend_dir = os.path.join(cwd, "backend")

if cwd not in sys.path:
    sys.path.insert(0, cwd)
if os.path.exists(backend_dir) and backend_dir not in sys.path:
    sys.path.insert(0, backend_dir)

try:
    print("Attempting to import app module...")
    import app
    fastapi_app = getattr(app, "app", None)
    if not fastapi_app:
        import app.app as fastapi_app
    print("✅ App module imported successfully! Starting Uvicorn server...")
    import uvicorn
    port = int(os.environ.get("PORT", 10000))
    uvicorn.run(fastapi_app, host="0.0.0.0", port=port)
except Exception as e:
    print("\n" + "="*50)
    print("❌ EXCEPTION OCCURRED DURING APP IMPORT:")
    print("ERROR MESSAGE:", str(e))
    print("="*50)
    traceback.print_exc()
    print("="*50 + "\n")
    sys.exit(1)

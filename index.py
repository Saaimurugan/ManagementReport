"""
Vercel entry point - imports and runs the Flask app
"""
from app import app

# Vercel expects the app to be available as 'app' variable
if __name__ == "__main__":
    app.run()
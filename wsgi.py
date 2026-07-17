"""
WSGI entry point for deployment platforms like Vercel, Heroku, etc.
"""
from app import app

if __name__ == "__main__":
    app.run()
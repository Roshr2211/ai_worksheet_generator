from fastapi import FastAPI
import subprocess

app = FastAPI()

@app.get("/")
def read_root():
    """Runs the Streamlit app as a subprocess."""
    subprocess.Popen(["streamlit", "run", "main.py", "--server.port=8501", "--server.headless=true"])
    return {"message": "Streamlit app is running"}

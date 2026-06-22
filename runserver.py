import sys
import uvicorn

port = 8000
for arg in sys.argv:
    if arg.startswith("--port"):
        port = int(sys.argv[sys.argv.index(arg) + 1])

if __name__ == "__main__":
    uvicorn.run("main:app", host="127.0.0.1", port=port)
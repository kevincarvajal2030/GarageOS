from flask import Flask, jsonify
import subprocess

app = Flask(__name__)


@app.get("/")
def HOME():
    return jsonify({
        "PROJECT": "GarageOS Bridge",
        "STATUS": "ONLINE"
    })


@app.post("/github/push")
def GITHUB_PUSH():

    COMMANDS = [
        [r"C:\Users\kevin\AppData\Roaming\npm\clasp.cmd", "pull"],
        [r"C:\Program Files\Git\cmd\git.exe", "-C", "..", "add", "."],
        [r"C:\Program Files\Git\cmd\git.exe", "-C", "..", "commit", "-m", "Auto Sync"],
        [r"C:\Program Files\Git\cmd\git.exe", "-C", "..", "push"]
    ]

    OUTPUT = []

    for COMMAND in COMMANDS:

        print("EXECUTING:", COMMAND)

        RESULT = subprocess.run(
            COMMAND,
            capture_output=True,
            text=True
        )

        OUTPUT.append({
            "COMMAND": " ".join(COMMAND),
            "RETURN_CODE": RESULT.returncode,
            "STDOUT": RESULT.stdout,
            "STDERR": RESULT.stderr
        })

    return jsonify(OUTPUT)


if __name__ == "__main__":
    app.run(debug=True)
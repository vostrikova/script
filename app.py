from flask import Flask

import subprocess


app = Flask(__name__)


@app.route("/", methods=["POST"])
def run_script():
    subprocess.Popen("python project_calculation.py", shell=True)
    return "OK"


if __name__ == "__main__":
    app.run(host="0.0.0.0", port=7878, debug=False)

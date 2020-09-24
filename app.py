from flask import Flask

import subprocess


app = Flask(__name__)


@app.route("/", methods=["POST"])
def run_script():
    with subprocess.Popen("python project_calculation.py", shell=True, stdout=subprocess.PIPE, stderr=subprocess.PIPE) as proc:
        stdout, stderr = proc.communicate()
    return "OK"


if __name__ == "__main__":
    app.run(host="0.0.0.0", port=7878, debug=False)

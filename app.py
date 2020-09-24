from flask import Flask

app = Flask(__name__)


@app.route("/", methods=["POST"])
def run_script():
    # RUN SCRIPT
    pass



if __name__ == "__main__":

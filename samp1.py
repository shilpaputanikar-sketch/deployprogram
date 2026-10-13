from flask import Flask, jsonify

app = Flask(__name__)

@app.route("/")
def home():
    return jsonify({
        "message": "Hello! Python backend is working."
    })

@app.route("/about")
def about():
    return jsonify({
        "name": "My Application",
        "backend": "Python Flask"
    })

if __name__ == "__main__":
    app.run(debug=True)

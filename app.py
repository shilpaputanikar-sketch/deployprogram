from flask import Flask, request, jsonify

app = Flask(__name__)

users = [
    {"id": 1, "name": "Shilpa"},
    {"id": 2, "name": "Yashaswini"},
]

tasks = []
next_task_id = 1


@app.route("/", methods=["GET"])
def home():
    return jsonify({"status": "ok", "message": "Flask API is running"})


@app.route("/users", methods=["GET"])
def get_users():
    return jsonify(users)


@app.route("/tasks", methods=["GET"])
def get_tasks():
    return jsonify(tasks)


@app.route("/tasks", methods=["POST"])
def add_task():
    global next_task_id
    data = request.get_json(silent=True) or {}
    if "title" not in data:
        return jsonify({"error": "title is required"}), 400
    task = {
        "id": next_task_id,
        "title": data["title"],
        "done": bool(data.get("done", False)),
    }
    tasks.append(task)
    next_task_id += 1
    return jsonify(task), 201


@app.route("/tasks/<int:task_id>", methods=["DELETE"])
def delete_task(task_id):
    global tasks
    for i, t in enumerate(tasks):
        if t["id"] == task_id:
            tasks.pop(i)
            return jsonify({"deleted": task_id})
    return jsonify({"error": "not found"}), 404


if __name__ == "__main__":
    app.run(host="0.0.0.0", port=5000, debug=True)

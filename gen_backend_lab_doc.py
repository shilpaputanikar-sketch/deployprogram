from docx import Document
from docx.shared import Pt, RGBColor, Inches
from docx.enum.text import WD_ALIGN_PARAGRAPH

doc = Document()

style = doc.styles['Normal']
style.font.name = 'Calibri'
style.font.size = Pt(11)

def add_code(code, doc):
    for line in code.strip().split('\n'):
        p = doc.add_paragraph()
        run = p.add_run(line)
        run.font.name = 'Consolas'
        run.font.size = Pt(9)
        run.font.color.rgb = RGBColor(0, 0, 128)
        p.paragraph_format.space_before = Pt(0)
        p.paragraph_format.space_after = Pt(0)
        p.paragraph_format.line_spacing = Pt(12)

def add_explanation(lines, doc):
    for label, code, explanation in lines:
        p = doc.add_paragraph()
        run = p.add_run(label + ' ')
        run.bold = True
        run.font.size = Pt(11)
        
        p2 = doc.add_paragraph()
        run2 = p2.add_run(code)
        run2.font.name = 'Consolas'
        run2.font.size = Pt(9)
        run2.font.color.rgb = RGBColor(0, 0, 128)
        p2.paragraph_format.space_before = Pt(0)
        p2.paragraph_format.space_after = Pt(0)
        
        doc.add_paragraph(explanation)

def add_program(doc, title, code, explanations):
    doc.add_heading(title, level=1)
    doc.add_heading('Complete Code', level=2)
    add_code(code, doc)
    if explanations:
        doc.add_paragraph()
        doc.add_heading('Line by Line Explanation', level=2)
        add_explanation(explanations, doc)
    doc.add_page_break()

# ============================================================
title = doc.add_heading('Backend Development Lab - 24BCA52P', level=0)
title.alignment = WD_ALIGN_PARAGRAPH.CENTER
doc.add_paragraph()

# ============================================================
# PROGRAM 1
# ============================================================
prog1_code = """from flask import Flask, jsonify

app = Flask(__name__)

# Static list of 5 users
users = [
    {"id": 1, "name": "Alice", "email": "alice@example.com"},
    {"id": 2, "name": "Bob", "email": "bob@example.com"},
    {"id": 3, "name": "Charlie", "email": "charlie@example.com"},
    {"id": 4, "name": "Diana", "email": "diana@example.com"},
    {"id": 5, "name": "Eve", "email": "eve@example.com"}
]

@app.route('/users', methods=['GET'])
def get_users():
    return jsonify(users)

@app.route('/users/<int:user_id>', methods=['GET'])
def get_user(user_id):
    user = next((u for u in users if u['id'] == user_id), None)
    if user:
        return jsonify(user)
    return jsonify({"error": "User not found"}), 404

if __name__ == '__main__':
    app.run(debug=True)"""

prog1_exp = [
    ("Line 1:", "from flask import Flask, jsonify",
     "Import Flask class and jsonify function. jsonify converts Python dictionaries/lists into JSON format for API responses."),
    ("Line 3:", "app = Flask(__name__)", "Create Flask application instance."),
    ("Lines 5-11:", "users = [...]",
     "Create a static list of 5 user dictionaries. Each user has id, name, and email fields. This acts as our in-memory database."),
    ("Line 13:", "@app.route('/users', methods=['GET'])",
     "Route decorator for /users endpoint. Only accepts GET requests. When someone visits /users, the function below runs."),
    ("Line 14:", "def get_users():",
     "Define function that handles GET /users requests."),
    ("Line 15:", "return jsonify(users)", "Convert the users list to JSON and return it as the HTTP response. Flask automatically sets Content-Type to application/json."),
    ("Line 17:", "@app.route('/users/<int:user_id>', methods=['GET'])",
     "Dynamic route that captures user_id from URL. The <int:> converter ensures only integers are accepted. Example: /users/3 will pass user_id=3."),
    ("Line 18:", "def get_user(user_id):",
     "Define function that receives the captured user_id as parameter."),
    ("Line 19:", "user = next((u for u in users if u['id'] == user_id), None)",
     "Use a generator expression to find the user with matching id. next() returns the first match or None if no user found."),
    ("Line 20:", "if user:", "Check if a user was found (not None)."),
    ("Line 21:", "return jsonify(user)", "If found, return the user dictionary as JSON with status code 200 (default)."),
    ("Line 22:", "return jsonify({\"error\": \"User not found\"}), 404",
     "If not found, return an error message as JSON with HTTP status code 404 (Not Found). The second value after comma is the status code."),
    ("Line 24:", "if __name__ == '__main__':",
     "Check if this script is run directly (not imported)."),
    ("Line 25:", "app.run(debug=True)", "Start the Flask development server with debug mode enabled."),
]
add_program(doc, "Program 1: Build User List API using Flask", prog1_code, prog1_exp)

# ============================================================
# PROGRAM 2
# ============================================================
prog2_code = """from flask import Flask, request, jsonify
import re

app = Flask(__name__)

def is_valid_email(email):
    pattern = r'^[a-zA-Z0-9._%+-]+@[a-zA-Z0-9.-]+\\.[a-zA-Z]{2,}$'
    return re.match(pattern, email) is not None

@app.route('/register', methods=['POST'])
def register():
    data = request.get_json()
    name = data.get('name', '')
    email = data.get('email', '')

    if not name or not email:
        return jsonify({
            "success": False,
            "message": "Name and email are required"
        }), 400

    if not is_valid_email(email):
        return jsonify({
            "success": False,
            "message": "Invalid email format"
        }), 400

    return jsonify({
        "success": True,
        "message": f"User {name} registered successfully with email {email}"
    })

if __name__ == '__main__':
    app.run(debug=True)"""

prog2_exp = [
    ("Line 1:", "from flask import Flask, request, jsonify",
     "Import Flask along with request (to read incoming data) and jsonify (to send JSON responses)."),
    ("Line 2:", "import re", "Import Python's regular expression module for validating email format."),
    ("Line 4:", "app = Flask(__name__)", "Create Flask application."),
    ("Lines 6-7:", "def is_valid_email(email): ...",
     "Define a helper function that checks if an email matches a regex pattern. The pattern ensures: characters before @, domain name, and at least 2-letter TLD. Returns True if valid, False otherwise."),
    ("Line 9:", "@app.route('/register', methods=['POST'])",
     "Define /register endpoint that only accepts POST requests (since we are submitting data)."),
    ("Line 10:", "def register():", "Define the register function."),
    ("Line 11:", "data = request.get_json()",
     "Parse the incoming request body as JSON. The client must send Content-Type: application/json."),
    ("Lines 12-13:", "name = data.get('name', '')  email = data.get('email', '')",
     "Extract 'name' and 'email' fields from the JSON data. get() with default '' prevents KeyError if field is missing."),
    ("Lines 15-19:", "if not name or not email: ...",
     "If either name or email is empty/missing, return a 400 Bad Request with success=False and an error message."),
    ("Lines 21-24:", "if not is_valid_email(email): ...",
     "Call the email validation function. If email is invalid, return 400 with error message."),
    ("Lines 26-29:", "return jsonify({...})",
     "If all validation passes, return success=True with a registration confirmation message."),
    ("Lines 31-32:", "if __name__... app.run...", "Start the Flask server."),
]
add_program(doc, "Program 2: Create Student Registration API", prog2_code, prog2_exp)

# ============================================================
# PROGRAM 3
# ============================================================
prog3_code = """from flask import Flask, request, jsonify
import sqlite3

app = Flask(__name__)

def init_db():
    conn = sqlite3.connect('students.db')
    cursor = conn.cursor()
    cursor.execute('''
        CREATE TABLE IF NOT EXISTS students (
            id INTEGER PRIMARY KEY AUTOINCREMENT,
            name TEXT NOT NULL,
            email TEXT NOT NULL,
            age INTEGER
        )
    ''')
    conn.commit()
    conn.close()

@app.route('/students', methods=['POST'])
def insert_student():
    data = request.get_json()
    conn = sqlite3.connect('students.db')
    cursor = conn.cursor()
    cursor.execute(
        'INSERT INTO students (name, email, age) VALUES (?, ?, ?)',
        (data['name'], data['email'], data.get('age', 0))
    )
    conn.commit()
    student_id = cursor.lastrowid
    conn.close()
    return jsonify({"message": "Student created", "id": student_id}), 201

@app.route('/students', methods=['GET'])
def get_students():
    conn = sqlite3.connect('students.db')
    cursor = conn.cursor()
    cursor.execute('SELECT * FROM students')
    rows = cursor.fetchall()
    conn.close()
    students = [
        {"id": r[0], "name": r[1], "email": r[2], "age": r[3]}
        for r in rows
    ]
    return jsonify(students)

if __name__ == '__main__':
    init_db()
    app.run(debug=True)"""

prog3_exp = [
    ("Line 1:", "from flask import Flask, request, jsonify", "Import Flask modules."),
    ("Line 2:", "import sqlite3", "Import SQLite3 — Python's built-in database module (no installation needed)."),
    ("Line 4:", "app = Flask(__name__)", "Create Flask app."),
    ("Lines 6-14:", "def init_db(): ...",
     "Function to create the database table if it doesn't exist. sqlite3.connect() creates or opens the database file. cursor.execute() runs SQL. CREATE TABLE IF NOT EXISTS ensures we don't get an error if table already exists. AUTOINCREMENT automatically assigns unique IDs."),
    ("Lines 16-26:", "@app.route('/students', methods=['POST'])",
     "POST endpoint to insert a student. Gets JSON data, connects to DB, uses parameterized query (?, ?) to prevent SQL injection, inserts the student, commits, retrieves the auto-generated ID, and returns 201 Created."),
    ("Lines 28-38:", "@app.route('/students', methods=['GET'])",
     "GET endpoint to fetch all students. Connects to DB, runs SELECT *, fetches all rows, converts each row tuple to a dictionary, returns JSON array."),
    ("Lines 40-42:", "if __name__... init_db(); app.run()",
     "Initialize the database table before starting the server."),
]
add_program(doc, "Program 3: Implement CRUD API with Database", prog3_code, prog3_exp)

# ============================================================
# PROGRAM 4
# ============================================================
prog4_code = """from flask import Flask, request, jsonify
import sqlite3

app = Flask(__name__)

def init_db():
    conn = sqlite3.connect('students.db')
    cursor = conn.cursor()
    cursor.execute('''
        CREATE TABLE IF NOT EXISTS students (
            id INTEGER PRIMARY KEY AUTOINCREMENT,
            name TEXT NOT NULL,
            email TEXT NOT NULL,
            age INTEGER
        )
    ''')
    conn.commit()
    conn.close()

@app.route('/students/<int:student_id>', methods=['PUT'])
def update_student(student_id):
    data = request.get_json()
    conn = sqlite3.connect('students.db')
    cursor = conn.cursor()
    cursor.execute(
        'UPDATE students SET name = ?, email = ?, age = ? WHERE id = ?',
        (data['name'], data['email'], data.get('age', 0), student_id)
    )
    conn.commit()
    if cursor.rowcount == 0:
        conn.close()
        return jsonify({"error": "Student not found"}), 404
    conn.close()
    return jsonify({"message": "Student updated"})

@app.route('/students/<int:student_id>', methods=['DELETE'])
def delete_student(student_id):
    conn = sqlite3.connect('students.db')
    cursor = conn.cursor()
    cursor.execute('DELETE FROM students WHERE id = ?', (student_id,))
    conn.commit()
    if cursor.rowcount == 0:
        conn.close()
        return jsonify({"error": "Student not found"}), 404
    conn.close()
    return jsonify({"message": "Student deleted"})

if __name__ == '__main__':
    init_db()
    app.run(debug=True)"""

prog4_exp = [
    ("Lines 1-5:", "Imports and Flask setup", "Same as Program 3 — imports Flask, sqlite3, creates app, defines init_db()."),
    ("Lines 20-31:", "@app.route('/students/<int:student_id>', methods=['PUT'])",
     "PUT endpoint to update a student by ID. Captures student_id from URL. Gets JSON data with updated fields. Runs UPDATE SQL with parameterized query. If cursor.rowcount is 0, no student was found with that ID — returns 404. Otherwise returns success."),
    ("Lines 33-41:", "@app.route('/students/<int:student_id>', methods=['DELETE'])",
     "DELETE endpoint to delete a student by ID. Captures student_id from URL. Runs DELETE SQL. Checks rowcount — if 0, student not found (404). Otherwise returns success message."),
]
add_program(doc, "Program 4: Extend CRUD with Update and Delete", prog4_code, prog4_exp)

# ============================================================
# PROGRAM 5
# ============================================================
prog5_code = """const express = require('express');
const app = express();
app.use(express.json());

let products = [
    { id: 1, name: 'Laptop', price: 75000 },
    { id: 2, name: 'Mouse', price: 500 },
    { id: 3, name: 'Keyboard', price: 1200 }
];
let nextId = 4;

app.get('/products', (req, res) => {
    res.json(products);
});

app.post('/products', (req, res) => {
    const product = {
        id: nextId++,
        name: req.body.name,
        price: req.body.price
    };
    products.push(product);
    res.status(201).json(product);
});

app.listen(3000, () => {
    console.log('Server running on port 3000');
});"""

prog5_exp = [
    ("Line 1:", "const express = require('express');",
     "Import the Express.js library in Node.js. require() is Node's way of importing modules."),
    ("Line 2:", "const app = express();", "Create an Express application instance (similar to Flask(__name__))."),
    ("Line 3:", "app.use(express.json());",
     "Add middleware that automatically parses incoming JSON request bodies. Without this, req.body would be undefined."),
    ("Lines 5-9:", "let products = [...];",
     "Create an in-memory array of product objects. Each product has id, name, and price. This acts as our temporary database."),
    ("Line 10:", "let nextId = 4;", "Counter to auto-generate unique IDs for new products (since 3 products already exist with IDs 1-3)."),
    ("Lines 12-14:", "app.get('/products', (req, res) => {...});",
     "Define GET /products endpoint. req = request object, res = response object. res.json(products) sends the products array as JSON."),
    ("Lines 16-22:", "app.post('/products', (req, res) => {...});",
     "Define POST /products endpoint. Creates a new product object with auto-generated ID, name from req.body.name, and price from req.body.price. Pushes it to the array. Returns 201 (Created) with the new product as JSON."),
    ("Lines 24-26:", "app.listen(3000, () => {...});",
     "Start the server on port 3000. The callback runs once the server starts. To run: npm install express, then node filename.js."),
]
add_program(doc, "Program 5: Build Product API using Node.js (Express)", prog5_code, prog5_exp)

# ============================================================
# PROGRAM 6
# ============================================================
prog6_code = """const express = require('express');
const app = express();
app.use(express.json());

// Middleware to log request method and URL
app.use((req, res, next) => {
    console.log(`[${new Date().toISOString()}] ${req.method} ${req.url}`);
    next();
});

let products = [
    { id: 1, name: 'Laptop', price: 75000 }
];
let nextId = 2;

app.get('/products', (req, res) => {
    res.json(products);
});

app.post('/products', (req, res) => {
    const { name, price } = req.body;

    // Validation
    if (!name || name.trim() === '') {
        return res.status(400).json({
            error: 'Product name is required and cannot be empty'
        });
    }
    if (!price || price <= 0) {
        return res.status(400).json({
            error: 'Product price must be greater than 0'
        });
    }

    const product = { id: nextId++, name, price };
    products.push(product);
    res.status(201).json(product);
});

app.listen(3000, () => console.log('Server on port 3000'));"""

prog6_exp = [
    ("Lines 1-4:", "Imports and setup", "Express setup with JSON body parser."),
    ("Lines 6-9:", "app.use((req, res, next) => {...});",
     "Custom middleware function. app.use() runs on EVERY request. It logs the timestamp, HTTP method (GET/POST), and URL path to the console. next() is called to pass control to the next middleware or route handler. Without next(), the request would hang forever."),
    ("Lines 14-16:", "GET /products", "Returns product list."),
    ("Lines 18-34:", "POST /products with validation",
     "Before adding a product, validate: name must exist and not be empty (name.trim() removes whitespace). If invalid, return 400 with error message. Price must exist and be > 0. If invalid, return 400. Only if both validations pass, the product is created and added."),
]
add_program(doc, "Program 6: Add Middleware and Validation", prog6_code, prog6_exp)

# ============================================================
# PROGRAM 7
# ============================================================
prog7_code = """from flask import Flask, request, jsonify

app = Flask(__name__)

tasks = []
next_id = 1

@app.route('/tasks', methods=['POST'])
def add_task():
    data = request.get_json()
    task = {
        'id': next_id,
        'title': data.get('title', ''),
        'status': data.get('status', 'pending')
    }
    tasks.append(task)
    next_id += 1
    return jsonify(task), 201

@app.route('/tasks', methods=['GET'])
def view_tasks():
    status = request.args.get('status')
    if status:
        filtered = [t for t in tasks if t['status'] == status]
        return jsonify(filtered)
    return jsonify(tasks)

if __name__ == '__main__':
    app.run(debug=True)"""

prog7_exp = [
    ("Lines 1-6:", "Setup", "Flask imports, create app, empty tasks list and ID counter."),
    ("Lines 8-16:", "POST /tasks — add task",
     "Receive JSON with title and optional status (defaults to 'pending'). Create a task dict with auto-generated id. Append to list. Return the created task with 201 status."),
    ("Lines 18-24:", "GET /tasks — view tasks",
     "Fetch all tasks. Optional query parameter ?status=pending filters tasks by status. request.args.get('status') reads query parameters from URL. Example: GET /tasks?status=completed returns only completed tasks."),
]
add_program(doc, "Program 7: Build Task Manager API", prog7_code, prog7_exp)

# ============================================================
# PROGRAM 8
# ============================================================
prog8_code = """from flask import Flask, request, jsonify

app = Flask(__name__)

# Hardcoded credentials
VALID_USERNAME = "admin"
VALID_PASSWORD = "password123"

@app.route('/login', methods=['POST'])
def login():
    data = request.get_json()
    username = data.get('username', '')
    password = data.get('password', '')

    if username == VALID_USERNAME and password == VALID_PASSWORD:
        return jsonify({
            "success": True,
            "message": "Login successful"
        })
    else:
        return jsonify({
            "success": False,
            "message": "Invalid username or password"
        }), 401

if __name__ == '__main__':
    app.run(debug=True)"""

prog8_exp = [
    ("Lines 1-5:", "Setup", "Flask imports, create app."),
    ("Lines 7-8:", "VALID_USERNAME = ... VALID_PASSWORD = ...",
     "Hardcoded credentials stored as constants. In a real app, passwords would be hashed and stored in a database."),
    ("Lines 10-22:", "@app.route('/login', methods=['POST'])",
     "Login endpoint that accepts POST with JSON body containing username and password. Compares with hardcoded values. If match: return success=True with 200. If no match: return success=False with 401 (Unauthorized) status code."),
]
add_program(doc, "Program 8: Build Login API", prog8_code, prog8_exp)

# ============================================================
# PROGRAM 9
# ============================================================
prog9_code = """from flask import Flask, request, jsonify
import requests

app = Flask(__name__)

# Get free API key from https://openweathermap.org/api
API_KEY = 'your_api_key_here'

@app.route('/weather', methods=['GET'])
def get_weather():
    city = request.args.get('city', 'London')
    url = f'https://api.openweathermap.org/data/2.5/weather?q={city}&appid={API_KEY}&units=metric'

    response = requests.get(url)

    if response.status_code != 200:
        return jsonify({
            "error": "City not found or API error"
        }), 404

    data = response.json()
    formatted = {
        "city": data['name'],
        "temperature": data['main']['temp'],
        "feels_like": data['main']['feels_like'],
        "humidity": data['main']['humidity'],
        "description": data['weather'][0]['description'],
        "wind_speed": data['wind']['speed']
    }
    return jsonify(formatted)

if __name__ == '__main__':
    app.run(debug=True)"""

prog9_exp = [
    ("Line 1:", "from flask import Flask, request, jsonify", "Flask imports."),
    ("Line 2:", "import requests", "Import the requests library (pip install requests) to make HTTP calls to external APIs."),
    ("Line 4:", "app = Flask(__name__)", "Create Flask app."),
    ("Line 7:", "API_KEY = 'your_api_key_here'",
     "API key for OpenWeatherMap. User must sign up at openweathermap.org and replace with their own key."),
    ("Lines 9-24:", "@app.route('/weather', methods=['GET'])",
     "GET /weather endpoint. Reads ?city= parameter from URL (defaults to 'London'). Constructs the OpenWeatherMap API URL with city and API key. Uses requests.get() to call the external API. If status is not 200 (city not found), returns error. Otherwise extracts relevant fields from the API response (temperature, humidity, description, wind speed) and returns formatted JSON."),
]
add_program(doc, "Program 9: Build Weather Data API Integration", prog9_code, prog9_exp)

# ============================================================
# PROGRAM 10
# ============================================================
prog10_code = """# n8n Workflow: Webhook to Store Data
#
# Steps to create in n8n (https://n8n.io):
#
# 1. Add Webhook node (trigger)
#    - Set "HTTP Method" to POST
#    - Set "Path" to /webhook
#    - Copy the webhook URL (e.g., https://your-n8n.com/webhook/abc123)
#
# 2. Add Set node (to process data)
#    - Keep all incoming fields
#
# 3. Add HTTP Request node (optional, to call your API)
#    - Method: POST
#    - URL: http://localhost:5000/users
#    - Send the webhook data as JSON body
#
# 4. Add Wait node (to log response)
#
# Python Flask endpoint to RECEIVE data from n8n:"""

prog10_code2 = """from flask import Flask, request, jsonify

app = Flask(__name__)

@app.route('/n8n-webhook', methods=['POST'])
def n8n_webhook():
    data = request.get_json()
    print("Received from n8n:", data)
    # Store or process the data
    return jsonify({"status": "received", "data": data})

if __name__ == '__main__':
    app.run(debug=True, port=5000)"""

prog10_exp = [
    ("n8n Setup:", "Go to n8n dashboard, create new workflow.",
     "n8n is a workflow automation tool. You create visual workflows by connecting nodes. Each node performs an action (webhook, API call, data transformation)."),
    ("Webhook Node:", "Acts as trigger",
     "Creates a URL that external services can POST data to. When data is sent to this URL, the workflow starts."),
    ("Set Node:", "Process or transform data",
     "Modify or enrich the incoming data before sending it further."),
    ("HTTP Request Node:", "Calls your Flask API",
     "Sends the processed data to your Flask endpoint (localhost:5000/n8n-webhook). Your Flask API receives and processes it."),
]
add_program(doc, "Program 10: Create Workflow using n8n", prog10_code + prog10_code2, prog10_exp)

# ============================================================
# PROGRAM 11
# ============================================================
prog11_code = """from flask import Flask, request, jsonify
import openai  # pip install openai

app = Flask(__name__)

# Set your OpenAI API key
openai.api_key = 'your-openai-api-key'

@app.route('/chat', methods=['POST'])
def chat():
    data = request.get_json()
    user_message = data.get('message', '')

    if not user_message:
        return jsonify({"error": "Message is required"}), 400

    response = openai.ChatCompletion.create(
        model="gpt-3.5-turbo",
        messages=[
            {"role": "user", "content": user_message}
        ]
    )

    ai_reply = response['choices'][0]['message']['content']

    return jsonify({
        "user_message": user_message,
        "ai_response": ai_reply
    })

if __name__ == '__main__':
    app.run(debug=True)"""

prog11_exp = [
    ("Line 1:", "Flask imports", "Standard Flask imports."),
    ("Line 2:", "import openai", "Import OpenAI Python library (pip install openai) to access ChatGPT API."),
    ("Line 4:", "app = Flask(__name__)", "Create Flask app."),
    ("Line 7:", "openai.api_key = 'your-openai-api-key'",
     "Set your OpenAI API key. Get it from platform.openai.com/api-keys. Never hardcode this in production — use environment variables."),
    ("Lines 9-22:", "@app.route('/chat', methods=['POST'])",
     "POST /chat endpoint. Expects JSON with 'message' field containing user input. Validates that message is not empty. Calls OpenAI's ChatCompletion API with gpt-3.5-turbo model. Sends the user message and gets AI response. Returns both the user message and AI reply as JSON."),
]
add_program(doc, "Program 11: Build Simple Chat API", prog11_code, prog11_exp)

# ============================================================
# PROGRAM 12
# ============================================================
prog12_code = """from flask import Flask, request, jsonify
import re

app = Flask(__name__)

def is_valid_email(email):
    pattern = r'^[a-zA-Z0-9._%+-]+@[a-zA-Z0-9.-]+\\.[a-zA-Z]{2,}$'
    return re.match(pattern, email) is not None

@app.route('/register', methods=['POST'])
def register():
    data = request.get_json()
    name = data.get('name', '')
    email = data.get('email', '')

    if not name or not email:
        return jsonify({
            "success": False,
            "status_code": 400,
            "message": "Validation Error",
            "errors": {
                "name": "Name is required" if not name else None,
                "email": "Email is required" if not email else None
            }
        }), 400

    if not is_valid_email(email):
        return jsonify({
            "success": False,
            "status_code": 400,
            "message": "Validation Error",
            "errors": {
                "email": "Invalid email format. Please provide a valid email address"
            }
        }), 400

    return jsonify({
        "success": True,
        "status_code": 201,
        "message": f"User {name} registered successfully",
        "data": {
            "name": name,
            "email": email
        }
    }), 201

if __name__ == '__main__':
    app.run(debug=True)"""

prog12_exp = [
    ("What changed:", "Improved response format",
     "This program takes the registration API from Program 2 and improves its response structure. Instead of simple success/message, we now include: status_code in the body (good for debugging), structured errors object showing which field failed and why, data field in success response containing the registered info, proper HTTP status codes (400 for errors, 201 for created)."),
    ("Purpose:", "",
     "A well-structured API response makes it easier for frontend developers to handle errors. Instead of parsing a text message, they can check the 'errors' object and display field-specific error messages."),
]
add_program(doc, "Program 12: Improve API using AI", prog12_code, prog12_exp)

# ============================================================
# PROGRAM 13
# ============================================================
prog13_code = """# Git Commands for Version Control

# Step 1: Initialize Git repository
git init

# Step 2: Check status of files
git status

# Step 3: Add files to staging area
git add .
# Or add specific file: git add app.py

# Step 4: Commit files with message
git commit -m "Initial commit: Backend API"

# Step 5: Create a new repository on GitHub
# Go to https://github.com -> New Repository
# Copy the remote URL

# Step 6: Add remote origin
git remote add origin https://github.com/yourusername/your-repo.git

# Step 7: Push code to GitHub
git push -u origin main

# Additional useful commands:
git log --oneline          # View commit history
git branch -M main        # Rename branch to main
git status                # See modified files
git diff                  # See changes in files
git add <filename>        # Stage specific file
git commit -m "message"   # Commit with message
git push                  # Push commits to remote
git pull                  # Pull latest from remote"""

prog13_exp = [
    ("git init:", "Creates a new .git hidden folder in your project.",
     "This initializes a Git repository. Now Git will track all changes in this folder. Run this once per project."),
    ("git status:", "Shows which files are modified/new/deleted.",
     "Red = untracked/modified. Green = staged (ready to commit). Always run this before committing to see what changed."),
    ("git add .:", "Adds all files to the staging area.",
     "Staging area is like a 'preview' of what will be committed. Use git add filename to stage specific files."),
    ("git commit -m 'message':", "Saves the staged changes permanently.",
     "Each commit creates a checkpoint you can revert to if something breaks. Messages should describe what changed."),
    ("git remote add origin <url>:", "Links your local repo to GitHub.",
     "This connects your local Git repository to a remote repository on GitHub so you can push/pull code."),
    ("git push -u origin main:", "Uploads your code to GitHub.",
     "The -u flag sets upstream (remembers where to push next time). Subsequent pushes only need git push."),
]
add_program(doc, "Program 13: Version Control using Git", prog13_code, prog13_exp)

# ============================================================
# PROGRAM 14
# ============================================================
prog14_code = """# Deploy Flask API on Render (Free)

# Step 1: Prepare your project
# Create a requirements.txt file:
pip freeze > requirements.txt

# Step 2: Create a Procfile (no extension, just 'Procfile')
# Add this line to Procfile:
web: gunicorn app:app

# Step 3: Push code to GitHub (see Program 13)

# Step 4: Deploy on Render
# 1. Go to https://render.com and sign up
# 2. Click "New +" > "Web Service"
# 3. Connect your GitHub repository
# 4. Configure:
#    - Name: your-api-name
#    - Environment: Python
#    - Build Command: pip install -r requirements.txt
#    - Start Command: gunicorn app:app
#    - Choose Free plan
# 5. Click "Create Web Service"
# 6. Wait 2-3 minutes for deployment

# Step 5: Test your deployed API
# Render provides URL like: https://your-api.onrender.com
# Test with:
curl https://your-api.onrender.com/users

# Deploy Flask API on PythonAnywhere (Alternative)
# 1. Go to https://www.pythonanywhere.com
# 2. Create account (Free tier available)
# 3. Go to Dashboard > Web > Add a new web app
# 4. Choose Flask and Python version
# 5. Upload code via Git or file upload
# 6. Reload web app
# 7. URL: https://yourusername.pythonanywhere.com"""

prog14_exp = [
    ("Step 1:", "Prepare requirements.txt",
     "Freeze all Python dependencies into requirements.txt so the deployment platform knows what to install."),
    ("Step 2:", "Create Procfile",
     "A Procfile tells Render how to run your app. gunicorn app:app means: use Gunicorn server, import 'app' from app.py. Gunicorn is a production-ready web server (unlike Flask's built-in debug server)."),
    ("Step 3:", "Push to GitHub",
     "Deployment platforms pull code from GitHub. All files must be committed and pushed."),
    ("Step 4:", "Deploy on Render",
     "Render automatically detects it's a Python app. It runs pip install, then gunicorn. Free tier sleeps after inactivity (wakes up on first request, takes ~30 seconds)."),
    ("Step 5:", "Test endpoints",
     "After deployment, all your endpoints work at the Render URL. Example: https://your-api.onrender.com/users should return your user list."),
]
add_program(doc, "Program 14: Deploy Backend API", prog14_code, prog14_exp)

# ============================================================
doc.add_heading('Appendix: How to Run Flask Apps', level=1)
steps = [
    "1. Install Python from python.org",
    "2. Install Flask: pip install flask",
    "3. Install additional packages as needed: pip install requests openai sqlite3",
    "4. Save the code as app.py (or any .py file)",
    "5. Run: python app.py",
    "6. Open browser and go to http://127.0.0.1:5000",
    "7. To test POST endpoints, use Postman or curl:",
]
for s in steps:
    doc.add_paragraph(s)
p = doc.add_paragraph()
run = p.add_run("curl -X POST http://127.0.0.1:5000/register -H \"Content-Type: application/json\" -d '{\"name\":\"Test\",\"email\":\"test@example.com\"}'")
run.font.name = 'Consolas'
run.font.size = Pt(9)

doc.add_paragraph()
doc.add_heading('How to Run Node.js Apps', level=1)
steps2 = [
    "1. Install Node.js from nodejs.org",
    "2. Create a new directory and navigate to it",
    "3. Run: npm init -y (creates package.json)",
    "4. Install Express: npm install express",
    "5. Save the code as server.js",
    "6. Run: node server.js",
    "7. Server starts on http://localhost:3000",
]
for s in steps2:
    doc.add_paragraph(s)

# Save
filepath = 'D:\\allen\\Backend_Development_Lab_Programs.docx'
doc.save(filepath)
print(f'Saved to {filepath}')

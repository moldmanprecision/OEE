from flask import Flask, send_from_directory, jsonify, request, redirect, session
import subprocess
import os

app = Flask(__name__)
app.secret_key = "secret123"   # change later

BASE_DIR = os.path.dirname(os.path.abspath(__file__))


# ================= LOGIN PAGE =================
@app.route("/login", methods=["GET", "POST"])
def login():
    if request.method == "POST":
        username = request.form.get("username")
        password = request.form.get("password")

        # 👉 यहाँ अपना user/password set करो
        if username == "admin" and password == "1234":
            session["user"] = username
            return redirect("/")
        else:
            return "❌ Invalid credentials"

    return '''
        <h2>Login</h2>
        <form method="post">
            <input name="username" placeholder="Username"><br><br>
            <input name="password" type="password" placeholder="Password"><br><br>
            <button type="submit">Login</button>
        </form>
    '''


# ================= LOGOUT =================
@app.route("/logout")
def logout():
    session.pop("user", None)
    return redirect("/login")


# ================= AUTH CHECK =================
def check_auth():
    return "user" in session


# ================= DASHBOARD =================
@app.route("/")
def home():
    if not check_auth():
        return redirect("/login")
    return send_from_directory(BASE_DIR, "oee_dashboard.html")


# ================= JSON =================
@app.route("/output/oee_data.json")
def json_data():
    if not check_auth():
        return jsonify({"error": "Unauthorized"})
    return send_from_directory(os.path.join(BASE_DIR, "output"), "oee_data.json")


# ================= RUN ETL =================
@app.route("/run-etl")
def run_etl():
    result = subprocess.run(
        ["python3", "oee_etl.py"],
        cwd=BASE_DIR,
        capture_output=True,
        text=True
    )
    return jsonify({
        "stdout": result.stdout,
        "stderr": result.stderr
    })


# ================= UPLOAD =================
@app.route("/upload-json", methods=["POST"])
def upload_json():
    file = request.files.get("file")

    if not file:
        return jsonify({"error": "No file uploaded"})

    save_path = os.path.join(BASE_DIR, "output", "oee_data.json")
    os.makedirs(os.path.dirname(save_path), exist_ok=True)

    file.save(save_path)

    return jsonify({"status": "uploaded successfully"})


# ================= SERVER =================
if __name__ == "__main__":
    port = int(os.environ.get("PORT", 10000))
    app.run(host="0.0.0.0", port=port)
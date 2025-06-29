from flask import Flask, render_template, jsonify, request, redirect, send_file
import json
import os
import urllib.parse
from generate_pdf import generate_report_for_project  # Import your clean PDF generator!
from rq import Queue
from rq.job import Job
import redis
import uuid
import platform

app = Flask(__name__)

CONFIG_FILE = 'app_config.json'

def load_config():
    if os.path.exists(CONFIG_FILE):
        with open(CONFIG_FILE, 'r') as f:
            return json.load(f)
    return {"projects": {}, "current_obs": {}}

def save_config(config):
    with open(CONFIG_FILE, 'w') as f:
        json.dump(config, f, indent=2)

app_config = load_config()

def get_projects():
    return app_config.get("projects", {})

def get_current_obs(project):
    return app_config.get("current_obs", {}).get(project, 0)

def set_current_obs(project, obs_number):
    if "current_obs" not in app_config:
        app_config["current_obs"] = {}
    app_config["current_obs"][project] = obs_number
    save_config(app_config)

def get_next_obs(project):
    current = get_current_obs(project)
    next_number = current + 1
    set_current_obs(project, next_number)
    return next_number

def get_prefilled_form_url(project, obs_number):
    projects = get_projects()
    if project not in projects:
        return None
    google_form_url = app_config.get("form_url")
    obs_field_id = app_config.get("obs_field_id")
    project_field_id = app_config.get("project_field_id")
    params = {
        'usp': 'pp_url',
        obs_field_id: str(obs_number),
        project_field_id: project
    }
    return f"{google_form_url}?{urllib.parse.urlencode(params)}"

@app.route('/')
def index():
    projects = get_projects()
    default_project = next(iter(projects)) if projects else ""
    return render_template('index.html', projects=projects, default_project=default_project)

@app.route('/get_projects')
def get_projects_route():
    return jsonify({"projects": list(get_projects().keys())})

@app.route('/get_current_obs')
def get_current_obs_route():
    project = request.args.get('project')
    if not project or project not in get_projects():
        return jsonify({"error": "Invalid or missing project"}), 400
    current_obs = get_current_obs(project)
    form_url = get_prefilled_form_url(project, current_obs)
    return jsonify({"obs_number": current_obs, "form_url": form_url})

@app.route('/get_next_obs')
def get_next_obs_route():
    project = request.args.get('project')
    if not project or project not in get_projects():
        return jsonify({"error": "Invalid or missing project"}), 400
    next_obs = get_next_obs(project)
    form_url = get_prefilled_form_url(project, next_obs)
    return jsonify({"obs_number": next_obs, "form_url": form_url})

@app.route('/reset_obs', methods=['POST'])
def reset_obs():
    data = request.get_json()
    project = data.get('project')
    new_number = data.get('new_number')
    if not project or project not in get_projects():
        return jsonify({"error": "Invalid or missing project"}), 400
    if not isinstance(new_number, int) or new_number < 1:
        return jsonify({"error": "Invalid new_number"}), 400
    set_current_obs(project, new_number - 1)
    return jsonify({"success": True})

@app.route('/open_observation_form')
def open_observation_form():
    project = request.args.get('project')
    if not project or project not in get_projects():
        return "Invalid or missing project", 400
    current_obs = get_current_obs(project)
    url = get_prefilled_form_url(project, current_obs)
    if not url:
        return "Form URL not found", 404
    return redirect(url)

redis_conn = redis.Redis()
task_queue = Queue(connection=redis_conn)

@app.route('/generate_report', methods=['POST'])
def generate_report():
    project = request.json.get('project')
    if not project or project not in get_projects():
        return jsonify({"error": "Invalid or missing project"}), 400
    try:
        job_id = str(uuid.uuid4())  # Generate a unique job ID
        job = task_queue.enqueue_call(
            func='generate_pdf.generate_report_for_project',
            args=(project,),
            job_id=job_id,
            timeout=600  # 10 minutes
        )
        return jsonify({"job_id": job_id}), 202
    except Exception as e:
        print(f"Error generating report: {e}")
        return jsonify({"error": str(e)}), 500

@app.route('/obs_submitted', methods=['POST'])
def obs_submitted():
    data = request.get_json()
    project = data.get('project')
    obs_number = data.get('obs_number')
    if not project or project not in get_projects():
        return jsonify({"error": "Invalid or missing project"}), 400
    current = get_current_obs(project)
    if obs_number == current:
        set_current_obs(project, current + 1)
    return jsonify({"success": True})

@app.route('/get_report_count')
def get_report_count():
    project = request.args.get('project')
    if not project or project not in get_projects():
        return jsonify({"error": "Invalid or missing project"}), 400
    # Load your sheet and count matching records
    from generate_pdf import get_report_record_count
    count = get_report_record_count(project)
    return jsonify({"count": count})

@app.route('/report_status/<job_id>')
def report_status(job_id):
    try:
        job = Job.fetch(job_id, connection=redis_conn)
    except Exception:
        return jsonify({"status": "not_found"}), 404

    if job.is_finished:
        return jsonify({"status": "finished", "download_url": f"/download_report/{job_id}"})
    elif job.is_failed:
        return jsonify({"status": "failed"})
    else:
        return jsonify({"status": "in_progress"})

@app.route('/download_report/<job_id>')
def download_report(job_id):
    try:
        job = Job.fetch(job_id, connection=redis_conn)
        pdf_path = job.result
        return send_file(pdf_path, as_attachment=True)
    except Exception:
        return "Report not found or not ready.", 404

if __name__ == '__main__':
    port = int(os.environ.get('PORT', 5000))
    app.run(host='0.0.0.0', port=port)






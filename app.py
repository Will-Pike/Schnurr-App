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
    return {"projects": {}}

def save_config(config):
    with open(CONFIG_FILE, 'w') as f:
        json.dump(config, f, indent=2)

app_config = load_config()

def get_projects():
    return app_config.get("projects", {})

def get_prefilled_form_url(project, obs_number, user=None, floor=None, room=None):
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
    
    # Add user, floor, and room field IDs if available
    if user and app_config.get("user_field_id"):
        params[app_config.get("user_field_id")] = user
    if floor and app_config.get("floor_field_id"):
        params[app_config.get("floor_field_id")] = floor
    if room and app_config.get("room_field_id"):
        params[app_config.get("room_field_id")] = room
    
    return f"{google_form_url}?{urllib.parse.urlencode(params)}"

@app.route('/')
def index():
    projects = get_projects()
    default_project = next(iter(projects)) if projects else ""
    return render_template('index.html', projects=projects, default_project=default_project)

@app.route('/get_projects')
def get_projects_route():
    return jsonify({"projects": list(get_projects().keys())})

@app.route('/get_last_row_data')
def get_last_row_data_route():
    try:
        from generate_pdf import get_last_row_data
        data = get_last_row_data()
        if data:
            return jsonify(data)
        else:
            return jsonify({"error": "No data found"}), 404
    except Exception as e:
        print(f"Error in get_last_row_data_route: {e}")
        return jsonify({"error": "Failed to load data"}), 500

@app.route('/get_current_obs')
def get_current_obs_route():
    project = request.args.get('project')
    if not project or project not in get_projects():
        return jsonify({"error": "Invalid or missing project"}), 400
    
    try:
        # Get the highest OBS from spreadsheet and add 1 for the next number
        from generate_pdf import get_highest_obs_for_project, get_last_row_data
        highest_obs = get_highest_obs_for_project(project)
        next_obs = highest_obs + 1
        
        # Get last row data for auto-filling
        last_row_data = get_last_row_data()
        
        if last_row_data:
            form_url = get_prefilled_form_url(
                project, 
                next_obs, 
                user=last_row_data.get("user"),
                floor=last_row_data.get("floor"),
                room=last_row_data.get("room")
            )
        else:
            form_url = get_prefilled_form_url(project, next_obs)
        
        response = {"obs_number": next_obs, "form_url": form_url}
        if last_row_data:
            response["last_row_data"] = last_row_data
        
        return jsonify(response)
    except Exception as e:
        print(f"Error in get_current_obs_route: {e}")
        return jsonify({"error": "Failed to load OBS data"}), 500

@app.route('/get_next_obs')
def get_next_obs_route():
    project = request.args.get('project')
    if not project or project not in get_projects():
        return jsonify({"error": "Invalid or missing project"}), 400
    
    try:
        # Get the highest OBS from spreadsheet and add 1 for the next number
        from generate_pdf import get_highest_obs_for_project, get_last_row_data
        highest_obs = get_highest_obs_for_project(project)
        next_obs = highest_obs + 1
        
        # Get last row data for auto-filling
        last_row_data = get_last_row_data()
        
        if last_row_data:
            form_url = get_prefilled_form_url(
                project, 
                next_obs, 
                user=last_row_data.get("user"),
                floor=last_row_data.get("floor"),
                room=last_row_data.get("room")
            )
        else:
            form_url = get_prefilled_form_url(project, next_obs)
        
        return jsonify({"obs_number": next_obs, "form_url": form_url})
    except Exception as e:
        print(f"Error in get_next_obs_route: {e}")
        return jsonify({"error": "Failed to load OBS data"}), 500

@app.route('/reset_obs', methods=['POST'])
def reset_obs():
    data = request.get_json()
    project = data.get('project')
    new_number = data.get('new_number')
    if not project or project not in get_projects():
        return jsonify({"error": "Invalid or missing project"}), 400
    if not isinstance(new_number, int) or new_number < 1:
        return jsonify({"error": "Invalid new_number"}), 400
    # This route is now mostly obsolete but keeping for compatibility
    return jsonify({"success": True})

@app.route('/open_observation_form')
def open_observation_form():
    project = request.args.get('project')
    if not project or project not in get_projects():
        return "Invalid or missing project", 400
    try:
        from generate_pdf import get_highest_obs_for_project
        highest_obs = get_highest_obs_for_project(project)
        next_obs = highest_obs + 1
        url = get_prefilled_form_url(project, next_obs)
        if not url:
            return "Form URL not found", 404
        return redirect(url)
    except Exception as e:
        print(f"Error in open_observation_form: {e}")
        return "Error loading form", 500

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
    # This route is now mostly obsolete but keeping for compatibility
    return jsonify({"success": True})

@app.route('/get_report_count')
def get_report_count():
    project = request.args.get('project')
    if not project or project not in get_projects():
        return jsonify({"error": "Invalid or missing project"}), 400
    try:
        from generate_pdf import get_report_record_count
        count = get_report_record_count(project)
        return jsonify({"count": count})
    except Exception as e:
        print(f"Error getting report count: {e}")
        return jsonify({"error": "Failed to get report count"}), 500

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






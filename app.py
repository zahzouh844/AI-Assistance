import os
from flask import Flask, request, render_template, send_from_directory, jsonify
from modules.summarizer import (
    init_glpi_session,
    get_glpi_entities,
    get_glpi_tickets_by_entity,
    get_users_by_entity,
    get_sla_alerts,
    format_alerts_java_style,
    create_pptx_from_glpi
)
from modules.ai_assistant import simple_ai_response, generate_response
from flask import Flask, render_template, request
from modules.ml_model import predict_ticket_priority
APP_TOKEN = "nOUdovQeRT0NTkaGlcNY6tWOPIpAf5t1L6xzJ4H0"
API_URL = "http://localhost/glpi/apirest.php"
import requests
# Initialisation de l'application
app = Flask(__name__)

UPLOAD_FOLDER = 'uploads'
STATIC_FOLDER = 'static/files'
os.makedirs(UPLOAD_FOLDER, exist_ok=True)
os.makedirs(STATIC_FOLDER, exist_ok=True)
app.config['UPLOAD_FOLDER'] = UPLOAD_FOLDER
app.config['STATIC_FOLDER'] = STATIC_FOLDER

# Page d'accueil avec génération de PowerPoint par entité GLPI
@app.route("/", methods=["GET", "POST"])
def index():
    session_token = init_glpi_session("glpi", "glpi")
    entities = get_glpi_entities(session_token) if session_token else []
    error = ""

    if request.method == "POST":
        entity_id = request.form.get("entity_id")
        entity_name = request.form.get("entity_name")

        if not entity_id or not entity_name:
            error = "Veuillez sélectionner une entité."
        else:
            output_path = os.path.join(STATIC_FOLDER, "glpi_filtered_tickets.pptx")

            data = get_glpi_tickets_by_entity(entity_id, session_token)
            users_list = get_users_by_entity(entity_id, session_token)

            if "error" in data:
                error = data["error"]
            else:
                try:
                    create_pptx_from_glpi(
                        data,
                        output_path,
                        entity_name=entity_name,
                        users_list=users_list
                    )
                    return send_from_directory(STATIC_FOLDER, "glpi_filtered_tickets.pptx", as_attachment=True)
                except Exception as e:
                    error = f"Erreur lors de la génération du PowerPoint : {e}"

    return render_template("index.html", error=error, entities=entities)


# Chatbot IA avec mode simple ou avancé (Mistral)
@app.route("/chat", methods=["POST"])
def chat():
    user_message = request.form.get("message")
    mode = request.form.get("mode", "simple")  # "simple" ou "mistral"
    session_token = init_glpi_session("glpi", "glpi")

    if not user_message:
        return jsonify({"response": "Aucun message reçu."}), 400

    if mode == "mistral":
        bot_response = generate_response(user_message, session_token)
    else:
        bot_response = simple_ai_response(user_message, session_token)

    return jsonify({"response": bot_response})


# Vérification des alertes SLA
@app.route("/check_sla", methods=["POST"])
def check_sla():
    session_token = init_glpi_session("glpi", "glpi")
    entities_response = get_glpi_entities(session_token)

    all_alerts = []

    for entity in entities_response:
        entity_id = entity.get("id")
        tickets_response = get_glpi_tickets_by_entity(entity_id, session_token)

        for ticket in tickets_response.get("tickets", []):
            if ticket.get("slas_id_ttr") != 0 or ticket.get("slas_id_tto") != 0:
                all_alerts.append({
                    "ticket_id": ticket["id"],
                    "sla_alert": "SLA associé à ce ticket"
                })

    if not all_alerts:
        return jsonify({"response": "Aucune alerte SLA trouvée pour aucune entité."}), 200

    return jsonify({"response": all_alerts})
import re

def split_camel_case(text):
    return re.sub(r'(?<!^)(?=[A-Z])', ' ', text)

@app.route("/predict", methods=["POST"])
def predict_priority():
    subject = request.form.get("subject")
    body = request.form.get("body") or "vide"

    # Ajout de la séparation camelCase
    subject = split_camel_case(subject)

    type_ = "1"
    queue = "1"
    priority = "2"
    language = "1"

    try:
        prediction_result = predict_ticket_priority(subject, body, type_, queue, priority, language)
    except Exception as e:
        prediction_result = f"Erreur: {str(e)}"

    return render_template("index.html", prediction=prediction_result)


    return render_template("index.html", prediction=prediction_result)
from flask import jsonify

@app.route("/api/tickets/<int:entity_id>")
def api_tickets(entity_id):
    session_token = init_glpi_session("glpi", "glpi")
    data = get_glpi_tickets_by_entity(entity_id, session_token)
    if "error" in data:
        return jsonify([])  # ou {"error": data["error"]}
    # Format pour ne retourner que id et subject (ou name)
    tickets = data.get("tickets", [])
    simplified = []
    for t in tickets:
        simplified.append({
            "id": t.get("id"),
            "subject": t.get("name") or t.get("subject") or "Sans titre"
        })
    return jsonify(simplified)


@app.route("/api/predict_priority/<int:ticket_id>")
def api_predict_priority(ticket_id):
    session_token = init_glpi_session("glpi", "glpi")
    # Récupérer le ticket détaillé
    detail_url = f"{API_URL}/Ticket/{ticket_id}"
    headers = {
        "App-Token": APP_TOKEN,
        "Session-Token": session_token
    }
    resp = requests.get(detail_url, headers=headers)
    if resp.status_code != 200:
        return jsonify({"error": "Ticket non trouvé"}), 404
    ticket = resp.json()
    # Extraire champs nécessaires au modèle
    subject = ticket.get("name") or ticket.get("subject") or ""
    body = ticket.get("content") or ""
    type_ = str(ticket.get("type", "1"))
    queue = str(ticket.get("queue", "1"))
    priority = str(ticket.get("priority", "2"))
    language = "1"  # si tu n’as pas l’info, tu peux fixer

    try:
        pred = predict_ticket_priority(subject, body, type_, queue, priority, language)
        return jsonify({"priority": pred})
    except Exception as e:
        return jsonify({"error": str(e)})




# Lancement de l'application
if __name__ == "__main__":
    app.run(debug=True)

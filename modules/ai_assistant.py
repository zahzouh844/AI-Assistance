import random
import ollama  # ➔ Utilisation de Ollama pour IA libre
from modules.summarizer import get_glpi_entities, get_glpi_tickets_by_entity
from modules.summarizer import get_ticket_status, get_ticket_priority

chat_history = []

# Dictionnaire des entités GLPI
entity_ids = {
    "client 1": "1",
    "client 2": "2",
    "client 3": "3",
    "entité racine": "0"
}

# Caches en mémoire pour ne pas recharger à chaque fois
cached_entities = []
cached_tickets = {}

GLPI_BASE_URL = "http://127.0.0.1/glpi/front/ticket.form.php?id="  # URL à adapter selon ton installation

def ask_ollama(prompt):
    """Interroger Ollama"""
    response = ollama.chat(model='mistral', messages=[{'role': 'user', 'content': prompt}])
    return response['message']['content']

def load_glpi_data(session_token):
    """Charger et cacher les données GLPI"""
    global cached_entities, cached_tickets

    cached_entities = get_glpi_entities(session_token)
    for entity_name, entity_id in entity_ids.items():
        tickets_data = get_glpi_tickets_by_entity(entity_id)
        if "tickets" in tickets_data:
            cached_tickets[entity_name] = tickets_data["tickets"]
        else:
            cached_tickets[entity_name] = []

def simple_ai_response(user_message, session_token):
    chat_history.append({"role": "user", "message": user_message})
    user_message = user_message.lower().strip()

    if "entités" in user_message:
        entities = get_glpi_entities(session_token)
        if entities:
            names = "<br>".join(f"• {e['name']}" for e in entities)
            response = f"<h3> Entités disponibles :</h3>{names}"
        else:
            response = "❌ Aucune entité disponible ou erreur d'authentification."

    elif "tickets" in user_message:
        # Recherche d'une entité dans le message
        for entity_name in entity_ids.keys():
            if entity_name.lower() in user_message:
                entity_id = entity_ids[entity_name]
                tickets_data = get_glpi_tickets_by_entity(entity_id, session_token)

                if "tickets" in tickets_data:
                    tickets = tickets_data["tickets"]
                    if tickets:
                        response_lines = [f"<h3> Tickets pour <b>{entity_name}</b> :</h3>"]
                        for idx, t in enumerate(tickets, start=1):
                            ticket_id = t.get("id")
                            name = t.get("name", "Sans titre")
                            status = get_ticket_status(t.get("status"))
                            priority = get_ticket_priority(t.get("priority"))
                            # Générer l'URL dynamique pour chaque ticket
                            url = f"http://localhost/glpi/front/ticket.form.php?id={ticket_id}"

                            line = f"""
                            <div style="margin: 8px 0; padding: 10px; background-color: #eef5f3; border-radius: 6px;">
                                <strong>{idx}. {name}</strong><br>
                                Statut : {status} | Urgence : {priority}<br>
                                <a href="{url}" target="_blank">
                                    <button style="margin-top: 5px;"> Voir le ticket</button>
                                </a>
                            </div>
                            """
                            response_lines.append(line)
                        response = "".join(response_lines)
                    else:
                        response = f"❌ Aucun ticket trouvé pour <b>{entity_name}</b>."
                else:
                    response = f"❌ Erreur lors de la récupération des tickets pour <b>{entity_name}</b>."
                break
        else:
            response = "❗ Merci de préciser une entité (ex : « tickets de client 1 »)."

    elif any(word in user_message for word in ["bonjour", "salut", "hello", "salam", "cc"]):
        response = random.choice([" Bonjour ! Comment puis-je vous aider ?", "Salut ! Prêt à bosser sur les tickets ? 😎"])

    else:
        response = generate_response(user_message, session_token)

    chat_history.append({"role": "bot", "message": response})
    return response



def generate_response(message, session_token):
    message = message.lower()

    for entity_name, entity_id in entity_ids.items():
        if entity_name in message:
            tickets_data = get_glpi_tickets_by_entity(entity_id, session_token)  # Passer session_token ici
            if "tickets" in tickets_data:
                tickets = tickets_data["tickets"]
                if tickets:
                    prompt = build_prompt_from_tickets(tickets, entity_name)
                    response = ask_ollama(prompt)

                else:
                    response = f"Aucun ticket trouvé pour l'entité {entity_name}."
            else:
                response = f"Erreur en récupérant les tickets pour l'entité {entity_name}."
            break
    else:
        # Si aucune entité détectée ➔ On laisse Mistral répondre
        response = ask_ollama(message)

    return response
def build_prompt_from_tickets(tickets, entity_name):
    if not tickets:
        return f"Aucun ticket disponible pour l'entité {entity_name}."

    prompt_lines = [f"Voici une liste de tickets GLPI pour l'entité « {entity_name} » :\n"]
    
    for idx, ticket in enumerate(tickets, 1):
        title = ticket.get("name", "Sans titre")
        description = ticket.get("content", ticket.get("description", "Pas de description fournie."))

        # Tronquer la description si elle est trop longue
        if len(description) > 300:
            description = description[:300] + "..."

        prompt_lines.append(f"{idx}. [Nom : {title}] - Description : {description}")

    prompt_lines.append("\nPeux-tu me faire un résumé **numéroté ligne par ligne** des problèmes les plus fréquents rencontrés par les utilisateurs dans ces tickets ? Réponds en format clair :\n1. ...\n2. ...\n3. ...")

    return "\n".join(prompt_lines)


#!/usr/bin/env python3
# -*- coding: utf-8 -*-
"""
Script d'automatisation d'envoi d'emails
Permet d'envoyer des emails automatiquement avec personnalisation et planification
"""

import smtplib
import ssl
from email.mime.text import MIMEText
from email.mime.multipart import MIMEMultipart
from email.mime.base import MIMEBase
from email import encoders
import csv
import json
import os
from datetime import datetime, timedelta
import time
import logging
import schedule
import threading
from pathlib import Path

# Configuration du logging
logging.basicConfig(
    level=logging.INFO,
    format='%(asctime)s - %(levelname)s - %(message)s',
    handlers=[
        logging.FileHandler('email_automation.log'),
        logging.StreamHandler()
    ]
)

class EmailAutomation:
    def __init__(self, config_file='email_config.json'):
        """
        Initialise le gestionnaire d'emails automatiques
        
        Args:
            config_file (str): Chemin vers le fichier de configuration
        """
        self.config = self.load_config(config_file)
        self.smtp_server = None
        self.scheduler_running = False
        self.sent_tracker_file = 'sent_emails_tracker.json'
        self.sent_emails = self.load_sent_tracker()
        
    def load_config(self, config_file):
        """Charge la configuration depuis un fichier JSON"""
        try:
            with open(config_file, 'r', encoding='utf-8') as f:
                return json.load(f)
        except FileNotFoundError:
            # Crée un fichier de configuration par défaut
            default_config = {
                "smtp_server": "smtp.gmail.com",
                "smtp_port": 587,
                "sender_email": "legofferwanvariant616@gmail.com",
                "sender_password": "CENSORED",
                "sender_name": "Erwan Le Goff",
                "delay_between_emails": 2,
                "max_emails_per_batch": 10,
                "schedule_enabled": False,
                "schedule_interval_hours": 12
            }
            with open(config_file, 'w', encoding='utf-8') as f:
                json.dump(default_config, f, indent=4, ensure_ascii=False)
            logging.info(f"Fichier de configuration créé : {config_file}")
            return default_config
    
    def load_sent_tracker(self):
        """Charge le tracker des emails envoyés"""
        try:
            with open(self.sent_tracker_file, 'r', encoding='utf-8') as f:
                return json.load(f)
        except FileNotFoundError:
            return {}
    
    def save_sent_tracker(self):
        """Sauvegarde le tracker des emails envoyés"""
        try:
            with open(self.sent_tracker_file, 'w', encoding='utf-8') as f:
                json.dump(self.sent_emails, f, indent=4, ensure_ascii=False)
        except Exception as e:
            logging.error(f"Erreur lors de la sauvegarde du tracker : {e}")
    
    def mark_email_as_sent(self, email, timestamp=None):
        """Marque un email comme envoyé"""
        if timestamp is None:
            timestamp = datetime.now().isoformat()
        self.sent_emails[email] = timestamp
        self.save_sent_tracker()
    
    def is_email_sent_recently(self, email, hours=24):
        """Vérifie si un email a été envoyé récemment"""
        if email not in self.sent_emails:
            return False
        
        try:
            sent_time = datetime.fromisoformat(self.sent_emails[email])
            return (datetime.now() - sent_time).total_seconds() < hours * 3600
        except:
            return False
    
    def connect_smtp(self):
        """Établit la connexion SMTP"""
        try:
            context = ssl.create_default_context()
            self.smtp_server = smtplib.SMTP(self.config['smtp_server'], self.config['smtp_port'])
            self.smtp_server.starttls(context=context)
            self.smtp_server.login(self.config['sender_email'], self.config['sender_password'])
            logging.info("Connexion SMTP établie avec succès")
            return True
        except Exception as e:
            logging.error(f"Erreur de connexion SMTP : {e}")
            return False
    
    def disconnect_smtp(self):
        """Ferme la connexion SMTP"""
        if self.smtp_server:
            self.smtp_server.quit()
            logging.info("Connexion SMTP fermée")
    
    def create_message(self, recipient_email, subject, body_text, body_html=None, attachments=None):
        """
        Crée un message email
        
        Args:
            recipient_email (str): Email du destinataire
            subject (str): Sujet de l'email
            body_text (str): Corps du message en texte brut
            body_html (str, optional): Corps du message en HTML
            attachments (list, optional): Liste des fichiers à joindre
        
        Returns:
            MIMEMultipart: Message email prêt à envoyer
        """
        message = MIMEMultipart("alternative")
        message["Subject"] = subject
        message["From"] = f"{self.config['sender_name']} <{self.config['sender_email']}>"
        message["To"] = recipient_email
        
        # Ajoute le corps du message
        text_part = MIMEText(body_text, "plain", "utf-8")
        message.attach(text_part)
        
        if body_html:
            html_part = MIMEText(body_html, "html", "utf-8")
            message.attach(html_part)
        
        # Ajoute les pièces jointes
        if attachments:
            for file_path in attachments:
                self.add_attachment(message, file_path)
        
        return message
    
    def add_attachment(self, message, file_path):
        """Ajoute une pièce jointe au message"""
        try:
            with open(file_path, "rb") as attachment:
                part = MIMEBase('application', 'octet-stream')
                part.set_payload(attachment.read())
            
            encoders.encode_base64(part)
            part.add_header(
                'Content-Disposition',
                f'attachment; filename= {os.path.basename(file_path)}'
            )
            message.attach(part)
            logging.info(f"Pièce jointe ajoutée : {file_path}")
        except Exception as e:
            logging.error(f"Erreur lors de l'ajout de la pièce jointe : {e}")
    
    def personalize_message(self, template, data):
        """
        Personnalise un message avec les données fournies
        
        Args:
            template (str): Template du message avec placeholders {nom}, {email}, etc.
            data (dict): Données pour remplacer les placeholders
        
        Returns:
            str: Message personnalisé
        """
        try:
            return template.format(**data)
        except KeyError as e:
            logging.error(f"Clé manquante dans les données : {e}")
            return template
    
    def send_single_email(self, recipient_email, subject, body_text, body_html=None, attachments=None):
        """
        Envoie un email unique
        
        Args:
            recipient_email (str): Email du destinataire
            subject (str): Sujet de l'email
            body_text (str): Corps du message en texte
            body_html (str, optional): Corps du message en HTML
            attachments (list, optional): Liste des pièces jointes
        
        Returns:
            bool: True si envoyé avec succès, False sinon
        """
        try:
            message = self.create_message(recipient_email, subject, body_text, body_html, attachments)
            self.smtp_server.send_message(message)
            logging.info(f"Email envoyé avec succès à {recipient_email}")
            return True
        except Exception as e:
            logging.error(f"Erreur lors de l'envoi à {recipient_email} : {e}")
            return False
    
    def send_bulk_emails(self, recipients_file, subject_template, body_template, html_template=None):
        """
        Envoie des emails en masse à partir d'un fichier CSV
        
        Args:
            recipients_file (str): Fichier CSV avec les destinataires
            subject_template (str): Template du sujet avec placeholders
            body_template (str): Template du corps avec placeholders
            html_template (str, optional): Template HTML avec placeholders
        
        Returns:
            dict: Statistiques d'envoi
        """
        stats = {"sent": 0, "failed": 0, "total": 0}
        
        try:
            with open(recipients_file, 'r', encoding='utf-8') as file:
                reader = csv.DictReader(file)
                
                for row in reader:
                    stats["total"] += 1
                    
                    # Personnalise le message
                    subject = self.personalize_message(subject_template, row)
                    body_text = self.personalize_message(body_template, row)
                    body_html = self.personalize_message(html_template, row) if html_template else None
                    
                    # Envoie l'email
                    if self.send_single_email(row['email'], subject, body_text, body_html):
                        stats["sent"] += 1
                    else:
                        stats["failed"] += 1
                    
                    # Délai entre les emails pour éviter le spam
                    time.sleep(self.config.get('delay_between_emails', 2))
                    
        except Exception as e:
            logging.error(f"Erreur lors de l'envoi en masse : {e}")
        
        return stats
    
    def create_sample_csv(self, filename='recipients.csv'):
        """Crée un fichier CSV d'exemple avec les bonnes colonnes et les données fournies"""
        sample_data = [
            ['nom', 'email', 'entreprise', 'secteur', 'ville'],
            ['Nicolas Rossi', 'niko@tchokos.net', 'Tchokos', 'Concepteur de sites Web', 'Epinay-Sur-Orge'],
            ['Servive Contact de Creano', 'contact@creano.paris', 'Creano', 'Concepteur de sites Web', 'Juvisy-sur-Orge'],
            ['Servive Contact', 'box@netobox.com','Netobox','Concepteur de sites Web', 'Bretigny-sur-Orge'],
            ['Service Contact de Livinweb', 'contact@livinweb.fr', 'Livinweb', 'Concepteur de sites Web', 'Paris 20'],
            ['Service Contact de Wecom Digital', 'bonjour@wecom.digital', 'Wecom Digital', 'Création et maintenance de sites webs', 'Paris'],
            ['Service Contact de Skyreka', 'contact@skyreka.com', 'Skyreka', 'Concepteur de sites web', 'Paris'],
            ['Service Contact', 'job@glucoz.fr', 'Glucoz', 'Concepteur de sites web', 'Paris'],
            ['Service Contact', 'contact@cubedesigners.com', 'cubedesigners', 'Concepteur de sites web à Paris', 'Paris']
        ]
        
        with open(filename, 'w', newline='', encoding='utf-8') as file:
            writer = csv.writer(file)
            writer.writerows(sample_data)
        
        logging.info(f"Fichier CSV d'exemple créé : {filename}")
    
    def send_cv_application(self, recipient_email, recipient_name, company, position, cv_path, cover_letter=None):
        """
        Envoie une candidature avec CV en PDF
        
        Args:
            recipient_email (str): Email du recruteur
            recipient_name (str): Nom du recruteur
            company (str): Nom de l'entreprise
            position (str): Poste visé
            cv_path (str): Chemin vers le fichier PDF du CV
            cover_letter (str, optional): Lettre de motivation personnalisée
        
        Returns:
            bool: True si envoyé avec succès, False sinon
        """
        # Vérification de l'existence du fichier CV
        if not os.path.exists(cv_path):
            logging.error(f"Fichier CV introuvable : {cv_path}")
            return False
        
        # Sujet de l'email
        subject = f"Candidature spontanée - Alternance Développeur Full Stack - Erwan Le Goff"
        
        # Corps du message personnalisé d'Erwan
        if not cover_letter:
            cover_letter = f"""Bonjour,

Je me permets de vous adresser ma candidature spontanée pour une alternance en développement web Full Stack, à partir de septembre 2025, au sein de {company}.

Actuellement en inscrit en formation Bac+3 Développeur Full Stack chez Cloud Campus, je suis passionné par la création d'applications web modernes, performantes et accessibles. J'ai déjà eu l'opportunité de développer des projets concrets lors de mon stage chez DonkeyCode et à travers mes projets personnels disponibles sur mon portfolio.

Mes compétences :
• Front-end : HTML, CSS, JavaScript, Angular, Bootstrap
• Back-end : PHP, Symfony, MySQL, familiarisation avec Laravel
• Outils : Visual Studio Code, Git, Figma, Wordpress
• Bonnes pratiques UX/UI, intégration responsive, requêtes API, manipulation de données JSON.

Je suis motivé, sérieux, et je souhaite évoluer dans un environnement où je pourrais contribuer activement à des projets tout en continuant à progresser au contact d'une équipe expérimentée.

Vous pouvez consulter mes réalisations sur mon portfolio :
🔗 https://erwan-le-goff.onrender.com

Je reste à votre disposition pour échanger et vous présenter plus en détail mes compétences et mes projets.

Je vous remercie pour votre attention et vous prie d'agréer mes salutations les plus sincères.

Erwan Le Goff
📞 07 67 33 51 17
📧 elegoff296@gmail.com"""
        
        # Version HTML de la lettre
        html_cover_letter = f"""
        <html>
        <body style="font-family: Arial, sans-serif; line-height: 1.6; color: #333; max-width: 600px;">
            <p>Bonjour,</p>
            
            <p>Je me permets de vous adresser ma candidature spontanée pour une <strong>alternance en développement web Full Stack</strong>, à partir de <strong>septembre 2025</strong>, au sein de <strong>{company}</strong>.</p>
            
            <p>Actuellement en formation <strong>Bac+3 Développeur Full Stack</strong> chez Cloud Campus, je suis passionné par la création d'applications web modernes, performantes et accessibles. J'ai déjà eu l'opportunité de développer des projets concrets lors de mon stage chez DonkeyCode et à travers mes projets personnels disponibles sur mon portfolio.</p>
            
            <div style="background-color: #f8f9fa; padding: 15px; border-left: 4px solid #007bff; margin: 20px 0;">
                <h3 style="color: #007bff; margin-top: 0;">Mes compétences :</h3>
                <ul style="margin-bottom: 0;">
                    <li><strong>Front-end :</strong> HTML, CSS, JavaScript, Angular, Bootstrap</li>
                    <li><strong>Back-end :</strong> PHP, Symfony, MySQL, familiarisation avec Laravel</li>
                    <li><strong>Outils :</strong> Visual Studio Code, Git, Figma, Wordpress</li>
                    <li><strong>Pratiques :</strong> UX/UI, intégration responsive, requêtes API, manipulation de données JSON</li>
                </ul>
            </div>
            
            <p>Je suis motivé, sérieux, et je souhaite évoluer dans un environnement où je pourrais contribuer activement à des projets tout en continuant à progresser au contact d'une équipe expérimentée.</p>
            
            <p>Vous pouvez consulter mes réalisations sur mon portfolio :<br>
            🔗 <a href="https://erwan-le-goff.onrender.com" style="color: #007bff;">https://erwan-le-goff.onrender.com</a></p>
            
            <p>Je reste à votre disposition pour échanger et vous présenter plus en détail mes compétences et mes projets.</p>
            
            <p>Je vous remercie pour votre attention et vous prie d'agréer mes salutations les plus sincères.</p>
            
            <div style="margin-top: 30px; padding: 15px; background-color: #f8f9fa; border-radius: 5px;">
                <strong>Erwan Le Goff</strong><br>
                📞 07 67 33 51 17<br>
                📧 elegoff296@gmail.com
            </div>
        </body>
        </html>
        """
        
        # Envoi de l'email avec le CV en pièce jointe
        return self.send_single_email(
            recipient_email=recipient_email,
            subject=subject,
            body_text=cover_letter,
            body_html=html_cover_letter,
            attachments=[cv_path]
        )
    
    def send_bulk_cv_applications(self, recipients_file, cv_path, cover_letter_template=None, check_duplicates=True):
        """
        Envoie des candidatures en masse avec CV
        
        Args:
            recipients_file (str): Fichier CSV avec les destinataires
            cv_path (str): Chemin vers le fichier PDF du CV
            cover_letter_template (str, optional): Template de lettre de motivation
            check_duplicates (bool): Vérifier les doublons avant envoi
        
        Returns:
            dict: Statistiques d'envoi
        """
        if not os.path.exists(cv_path):
            logging.error(f"Fichier CV introuvable : {cv_path}")
            return {"sent": 0, "failed": 0, "total": 0, "skipped": 0, "error": "CV file not found"}
        
        stats = {"sent": 0, "failed": 0, "total": 0, "skipped": 0}
        max_emails = self.config.get('max_emails_per_batch', 10)
        
        try:
            with open(recipients_file, 'r', encoding='utf-8') as file:
                reader = csv.DictReader(file)
                
                for row in reader:
                    stats["total"] += 1
                    
                    # Extraction des données
                    recipient_email = row.get('email', '')
                    recipient_name = row.get('nom', 'Madame/Monsieur')
                    company = row.get('entreprise', 'votre entreprise')
                    position = row.get('poste', 'le poste proposé')
                    
                    # Vérification des doublons
                    if check_duplicates and self.is_email_sent_recently(recipient_email):
                        logging.info(f"Email déjà envoyé récemment à {recipient_email}, ignoré")
                        stats["skipped"] += 1
                        continue
                    
                    # Limite du nombre d'emails par lot
                    if stats["sent"] >= max_emails:
                        logging.info(f"Limite de {max_emails} emails atteinte pour ce lot")
                        break
                    
                    # Personnalisation de la lettre de motivation si template fourni
                    if cover_letter_template:
                        cover_letter = self.personalize_message(cover_letter_template, row)
                    else:
                        cover_letter = None
                    
                    # Envoi de la candidature
                    if self.send_cv_application(recipient_email, recipient_name, company, position, cv_path, cover_letter):
                        stats["sent"] += 1
                        self.mark_email_as_sent(recipient_email)
                        logging.info(f"Candidature envoyée à {company} ({recipient_email})")
                    else:
                        stats["failed"] += 1
                        logging.error(f"Échec envoi à {company} ({recipient_email})")
                    
                    # Délai entre les envois
                    time.sleep(self.config.get('delay_between_emails', 3))
                    
        except Exception as e:
            logging.error(f"Erreur lors de l'envoi en masse : {e}")
            stats["error"] = str(e)
        
        return stats
    
    def scheduled_cv_send(self, recipients_file, cv_path):
        """
        Fonction appelée par le scheduler pour envoyer des candidatures
        
        Args:
            recipients_file (str): Fichier CSV avec les destinataires
            cv_path (str): Chemin vers le fichier PDF du CV
        """
        logging.info("=== Démarrage de l'envoi programmé de candidatures ===")
        
        if not os.path.exists(recipients_file):
            logging.error(f"Fichier recipients introuvable : {recipients_file}")
            return
        
        if not self.validate_pdf(cv_path):
            logging.error(f"Fichier CV invalide : {cv_path}")
            return
        
        # Connexion SMTP
        if not self.connect_smtp():
            logging.error("Impossible de se connecter au serveur SMTP")
            return
        
        try:
            # Envoi des candidatures
            stats = self.send_bulk_cv_applications(recipients_file, cv_path)
            
            # Log des statistiques
            logging.info(f"Envoi programmé terminé - Stats: {stats}")
            
            # Notification si tous les emails ont été envoyés
            if stats["sent"] == 0 and stats["skipped"] > 0:
                logging.info("Tous les emails ont déjà été envoyés récemment")
            
        except Exception as e:
            logging.error(f"Erreur pendant l'envoi programmé : {e}")
        
        finally:
            self.disconnect_smtp()
    
    def start_scheduler(self, recipients_file, cv_path, interval_hours=12):
        """
        Démarre le scheduler pour l'envoi automatique
        
        Args:
            recipients_file (str): Fichier CSV avec les destinataires
            cv_path (str): Chemin vers le fichier PDF du CV
            interval_hours (int): Intervalle en heures entre les envois
        """
        if self.scheduler_running:
            logging.warning("Le scheduler est déjà en cours d'exécution")
            return
        
        self.scheduler_running = True
        
        # Configuration du scheduler
        schedule.every(interval_hours).hours.do(
            self.scheduled_cv_send, 
            recipients_file, 
            cv_path
        )
        
        logging.info(f"Scheduler démarré - Envoi toutes les {interval_hours} heures")
        logging.info(f"Prochaine exécution: {schedule.next_run()}")
        
        # Boucle d'exécution du scheduler
        def run_scheduler():
            while self.scheduler_running:
                try:
                    schedule.run_pending()
                    time.sleep(60)  # Vérification toutes les minutes
                except Exception as e:
                    logging.error(f"Erreur dans le scheduler : {e}")
                    time.sleep(60)
        
        # Démarrage du scheduler dans un thread séparé
        scheduler_thread = threading.Thread(target=run_scheduler, daemon=True)
        scheduler_thread.start()
        
        return scheduler_thread
    
    def stop_scheduler(self):
        """Arrête le scheduler"""
        self.scheduler_running = False
        schedule.clear()
        logging.info("Scheduler arrêté")
    
    def get_scheduler_status(self):
        """Retourne le statut du scheduler"""
        return {
            "running": self.scheduler_running,
            "next_run": str(schedule.next_run()) if schedule.jobs else "Aucune tâche programmée",
            "jobs_count": len(schedule.jobs)
        }
    
    def validate_pdf(self, pdf_path):
        """
        Valide qu'un fichier est bien un PDF
        
        Args:
            pdf_path (str): Chemin vers le fichier
            
        Returns:
            bool: True si c'est un PDF valide
        """
        try:
            if not os.path.exists(pdf_path):
                return False
            
            # Vérification de l'extension
            if not pdf_path.lower().endswith('.pdf'):
                return False
            
            # Vérification du header PDF
            with open(pdf_path, 'rb') as file:
                header = file.read(4)
                return header == b'%PDF'
                
        except Exception as e:
            logging.error(f"Erreur lors de la validation PDF : {e}")
            return False

def main():
    """Fonction principale avec menu interactif"""
    email_bot = EmailAutomation()
    
    while True:
        print("\n=== Script d'Automatisation d'Emails ===")
        print("1. Envoyer un email unique")
        print("2. Envoyer des emails en masse")
        print("3. Envoyer une candidature avec CV (unique)")
        print("4. Envoyer des candidatures avec CV (en masse)")
        print("5. Créer un fichier CSV d'exemple")
        print("6. Tester la connexion SMTP")
        print("7. 🕒 Démarrer l'envoi programmé (toutes les 12h)")
        print("8. 🛑 Arrêter l'envoi programmé")
        print("9. 📊 Statut du scheduler")
        print("10. 📋 Voir les emails envoyés")
        print("0. Quitter")
        
        choice = input("Choisissez une option : ")
        
        if choice == "1":
            # Email unique
            recipient = input("Email du destinataire : ")
            subject = input("Sujet : ")
            body = input("Message : ")
            
            # Option pièce jointe
            attachment = input("Chemin vers une pièce jointe (optionnel) : ")
            attachments = [attachment] if attachment and os.path.exists(attachment) else None
            
            if email_bot.connect_smtp():
                email_bot.send_single_email(recipient, subject, body, attachments=attachments)
                email_bot.disconnect_smtp()
        
        elif choice == "2":
            # Emails en masse
            csv_file = input("Fichier CSV des destinataires (recipients.csv) : ") or "recipients.csv"
            subject_template = input("Template du sujet (ex: Bonjour {nom}) : ")
            body_template = input("Template du message (ex: Bonjour {nom} de {entreprise}) : ")
            
            if email_bot.connect_smtp():
                stats = email_bot.send_bulk_emails(csv_file, subject_template, body_template)
                print(f"Statistiques : {stats['sent']} envoyés, {stats['failed']} échoués sur {stats['total']} total")
                email_bot.disconnect_smtp()
        
        elif choice == "3":
            # Candidature unique avec CV
            print("\n=== Envoi de candidature avec CV ===")
            cv_path = input("Chemin vers votre CV (PDF) : ")
            
            if not email_bot.validate_pdf(cv_path):
                print("❌ Fichier CV invalide ou introuvable. Assurez-vous que c'est un fichier PDF.")
                continue
            
            recipient_email = input("Email du recruteur : ")
            recipient_name = input("Nom du recruteur : ")
            company = input("Nom de l'entreprise : ")
            position = input("Poste visé : ")
            
            # Lettre de motivation personnalisée (optionnel)
            custom_letter = input("Lettre de motivation personnalisée (Entrée pour utiliser le template par défaut) : ")
            
            if email_bot.connect_smtp():
                success = email_bot.send_cv_application(
                    recipient_email, recipient_name, company, position, cv_path, 
                    custom_letter if custom_letter else None
                )
                if success:
                    print("✅ Candidature envoyée avec succès !")
                    email_bot.mark_email_as_sent(recipient_email)
                else:
                    print("❌ Échec de l'envoi de la candidature.")
                email_bot.disconnect_smtp()
        
        elif choice == "4":
            # Candidatures en masse avec CV
            print("\n=== Envoi de candidatures en masse avec CV ===")
            cv_path = input("Chemin vers votre CV (PDF) : ")

            if not email_bot.validate_pdf(cv_path):
                print("❌ Fichier CV invalide ou introuvable. Assurez-vous que c'est un fichier PDF.")
                continue

            csv_file = input("Fichier CSV des destinataires (recipients.csv) : ") or "recipients.csv"

            if not os.path.exists(csv_file):
                print(f"❌ Fichier {csv_file} introuvable.")
                continue

            if email_bot.connect_smtp():
                stats = email_bot.send_bulk_cv_applications(csv_file, cv_path)
                print(f"Statistiques : {stats['sent']} envoyés, {stats['failed']} échoués, {stats['skipped']} ignorés sur {stats['total']} total")
                email_bot.disconnect_smtp()

        elif choice == "5":
            email_bot.create_sample_csv()
            print("✅ Fichier CSV d'exemple créé !")

        elif choice == "6":
            if email_bot.connect_smtp():
                print("✅ Connexion SMTP réussie !")
                email_bot.disconnect_smtp()
            else:
                print("❌ Échec de la connexion SMTP.")

        elif choice == "7":
            # Démarrer l'envoi programmé
            print("\n=== Configuration de l'envoi programmé ===")
            
            cv_path = input("Chemin vers votre CV (PDF) : ")
            if not email_bot.validate_pdf(cv_path):
                print("❌ Fichier CV invalide ou introuvable.")
                continue
            
            csv_file = input("Fichier CSV des destinataires (recipients.csv) : ") or "recipients.csv"
            if not os.path.exists(csv_file):
                print(f"❌ Fichier {csv_file} introuvable.")
                continue
            
            # Configuration de l'intervalle
            interval_input = input("Intervalle entre les envois en heures (12 par défaut) : ")
            try:
                interval_hours = int(interval_input) if interval_input else 12
            except ValueError:
                interval_hours = 12
            
            # Confirmation
            print(f"\n📋 Configuration:")
            print(f"   CV: {cv_path}")
            print(f"   Destinataires: {csv_file}")
            print(f"   Intervalle: {interval_hours} heures")
            
            confirm = input("\nDémarrer l'envoi programmé ? (o/N) : ").lower()
            if confirm in ['o', 'oui', 'y', 'yes']:
                email_bot.start_scheduler(csv_file, cv_path, interval_hours)
                print("✅ Envoi programmé démarré !")
                print("Le script continuera à fonctionner en arrière-plan.")
                print("Appuyez sur Entrée pour revenir au menu...")
                input()
            else:
                print("❌ Envoi programmé annulé.")

        elif choice == "8":
            # Arrêter l'envoi programmé
            status = email_bot.get_scheduler_status()
            if status["running"]:
                email_bot.stop_scheduler()
                print("✅ Envoi programmé arrêté.")
            else:
                print("ℹ️  Aucun envoi programmé en cours.")

        elif choice == "9":
            # Statut du scheduler
            status = email_bot.get_scheduler_status()
            print(f"\n📊 Statut du scheduler:")
            print(f"   État: {'🟢 Actif' if status['running'] else '🔴 Arrêté'}")
            print(f"   Prochaine exécution: {status['next_run']}")
            print(f"   Tâches programmées: {status['jobs_count']}")
            
            if status["running"]:
                print(f"   Emails envoyés: {len(email_bot.sent_emails)}")
                print("\nAppuyez sur Entrée pour continuer...")
                input()

        elif choice == "10":
            # Voir les emails envoyés
            print(f"\n📋 Emails envoyés ({len(email_bot.sent_emails)} total):")
            if email_bot.sent_emails:
                for email, timestamp in list(email_bot.sent_emails.items())[-10:]:  # Derniers 10
                    try:
                        sent_time = datetime.fromisoformat(timestamp)
                        formatted_time = sent_time.strftime("%d/%m/%Y %H:%M")
                        print(f"   📧 {email} - {formatted_time}")
                    except:
                        print(f"   📧 {email} - {timestamp}")
                
                if len(email_bot.sent_emails) > 10:
                    print(f"   ... et {len(email_bot.sent_emails) - 10} autres")
            else:
                print("   Aucun email envoyé pour le moment.")
            
            print("\nAppuyez sur Entrée pour continuer...")
            input()

        elif choice == "0":
            # Quitter
            if email_bot.scheduler_running:
                print("⚠️  Un envoi programmé est en cours.")
                stop_scheduler = input("Voulez-vous l'arrêter avant de quitter ? (o/N) : ").lower()
                if stop_scheduler in ['o', 'oui', 'y', 'yes']:
                    email_bot.stop_scheduler()
                    print("✅ Scheduler arrêté.")
                else:
                    print("⚠️  Le scheduler continuera à fonctionner en arrière-plan.")
            
            print("Fermeture du programme. Au revoir !")
            break

        else:
            print("❌ Option invalide. Veuillez réessayer.")


if __name__ == "__main__":
    try:
        main()
    except KeyboardInterrupt:
        print("\n\n🛑 Interruption par l'utilisateur.")
        print("Fermeture du programme...")
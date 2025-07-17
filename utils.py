import os
import re
import logging
import configparser
from datetime import datetime
from typing import Dict
import pandas as pd
import smtplib
import openai
import json
from email.mime.text import MIMEText
from email.mime.multipart import MIMEMultipart
from faster_whisper import WhisperModel
import sys
from pydub import AudioSegment
import noisereduce as nr
import numpy as np
import tempfile
import scipy.io.wavfile as wav

# Ensure UTF-8 encoding for stdout
sys.stdout.reconfigure(encoding='utf-8')

# Configure logging
logging.basicConfig(
    level=logging.INFO,
    format='%(asctime)s - %(name)s - %(levelname)s - %(message)s',
    handlers=[
        logging.FileHandler('appointment_system.log', encoding='utf-8'),
        logging.StreamHandler(sys.stdout)
    ]
)
logger = logging.getLogger("utils")


class ConfigManager:
    """Handles configuration file operations"""
    def __init__(self, config_path='config.ini'):
        self.config = configparser.ConfigParser()
        self.config.read(config_path)

    def get_path(self, key):
        if 'PATHS' not in self.config:
            raise ValueError("Missing [PATHS] section in config.ini")
        return self.config['PATHS'].get(key)

    def get_email_config(self, key):
        return self.config['EMAIL'].get(key)

    def get_whisper_config(self, key):
        return self.config['WHISPER'].get(key)

    def get_openai_config(self, key):
        return self.config['OPENAI'].get(key)



class AudioProcessor:
    """Handles audio transcription using faster-whisper, with preprocessing"""
    def __init__(self, model_path: str):
        # Use GPU and float32 to avoid warning
        self.model = WhisperModel(model_path, device="cuda", compute_type="float32")

    def preprocess_audio(self, audio_path: str) -> str:
        """Clean and normalize audio, return path to processed WAV"""
        try:
            audio = AudioSegment.from_file(audio_path)

            # Normalize volume
            audio = audio.set_channels(1).set_frame_rate(16000).normalize()

            # Save as temporary WAV file
            with tempfile.NamedTemporaryFile(delete=False, suffix='.wav') as tmp_wav:
                audio.export(tmp_wav.name, format='wav')
                tmp_wav_path = tmp_wav.name

            # Noise reduce
            rate, data = wav.read(tmp_wav_path)
            reduced_noise = nr.reduce_noise(y=data.astype(np.float32), sr=rate)

            # Save noise-reduced version
            with tempfile.NamedTemporaryFile(delete=False, suffix='.wav') as clean_file:
                wav.write(clean_file.name, rate, reduced_noise.astype(np.int16))
                return clean_file.name

        except Exception as e:
            logger.error(f"Error during audio preprocessing: {e}")
            raise

    def transcribe_audio(self, audio_path: str) -> str:
        """Preprocess + transcribe with dynamic language switching (no language lock)"""
        try:
            logger.info(f"Preprocessing and transcribing audio: {audio_path}")
            cleaned_audio = self.preprocess_audio(audio_path)

            # Do NOT set a fixed language; let Whisper auto-detect and switch between English/Arabic
            segments, _ = self.model.transcribe(cleaned_audio, beam_size=5, language=None)
            transcript = " ".join([segment.text for segment in segments])
            logger.info("Transcription complete.")
            return transcript
        except Exception as e:
            logger.error(f"Error during transcription: {e}")
            raise


class DataExtractor:
    """Extracts appointment info using OpenAI LLM"""
    def __init__(self, config):
        # Set the API key for the openai module
        openai.api_key = config.get_openai_config('api_key')

    def extract_info(self, transcript: str) -> Dict[str, str]:
        prompt = f"""
You are a helpful assistant. Extract appointment info from this text.
Return only valid JSON with the following keys:
name, email, phone, plate, model, date, time_of_call, time_of_appointment, services_provided.
If missing, return null for that key.

Transcript: \"\"\"{transcript}\"\"\"
"""
        try:
            response = openai.ChatCompletion.create(
                model="gpt-4o",
                messages=[{"role": "user", "content": prompt}],
                temperature=0,
            )
            content = response.choices[0].message['content'] if isinstance(response.choices[0].message, dict) else response.choices[0].message.content
            logger.info(f"LLM Response: {content}")

            # Remove markdown code block if present
            match = re.search(r"```(?:json)?\s*(\{.*?\})\s*```", content, re.DOTALL)
            if match:
                content = match.group(1).strip()

            return json.loads(content)
        except Exception as e:
            logger.error(f"Error extracting data from LLM: {e}")
            raise


class ExcelManager:
    def __init__(self, file_path):
        self.file_path = file_path
        self.headers = ['name', 'email', 'phone', 'plate', 'model', 'date',
                        'time_of_call', 'time_of_appointment', 'services_provided']
        self._ensure_excel_file_exists()

    def _ensure_excel_file_exists(self):
        if not os.path.exists(self.file_path):
            logger.info(f"Excel file not found. Creating new one at {self.file_path}")
            try:
                df = pd.DataFrame(columns=self.headers)
                df.to_excel(self.file_path, index=False, engine='openpyxl')
                logger.info("Created new Excel file.")
            except Exception as e:
                logger.error(f"Failed to create Excel file: {e}")

    def add_appointment(self, appointment_data):
        try:
            df = pd.read_excel(self.file_path, engine='openpyxl')
            new_entry = {col: appointment_data.get(col, "") for col in self.headers}
            df = pd.concat([df, pd.DataFrame([new_entry])], ignore_index=True)
            df.to_excel(self.file_path, index=False, engine='openpyxl')
            logger.info("Appointment saved to Excel.")
        except Exception as e:
            logger.error(f"Error saving to Excel: {e}")
            raise


class EmailSender:
    def __init__(self, config: ConfigManager):
        self.smtp_server = config.get_email_config('smtp_server')
        self.smtp_port = int(config.get_email_config('smtp_port'))
        self.email_from = config.get_email_config('email_from')
        self.email_password = config.get_email_config('email_password')
        self.use_tls = config.get_email_config('use_tls').lower() == 'yes'

    def send_confirmation(self, recipient: str, details: Dict[str, str]):
        try:
            msg = MIMEMultipart()
            msg['From'] = self.email_from
            msg['To'] = recipient
            msg['Subject'] = 'Your Car Dealership Appointment Confirmation'

            # Format services
            services_list = details.get('services_provided')
            if isinstance(services_list, list):
                services_html = "<ul>" + "".join(f"<li>{service}</li>" for service in services_list) + "</ul>"
            else:
                services_html = f"<p>{services_list or 'Not specified'}</p>"

            html = f"""
            <html>
                <body>
                    <h2>Appointment Confirmation</h2>
                    <p>Dear {details.get('name', 'Customer')},</p>
                    <p>This confirms your appointment:</p>
                    <ul>
                        <li><strong>Name:</strong> {details.get('name')}</li>
                        <li><strong>Email:</strong> {details.get('email')}</li>
                        <li><strong>Phone:</strong> {details.get('phone') or 'Not available'}</li>
                        <li><strong>Date:</strong> {details.get('date')}</li>
                        <li><strong>Time:</strong> {details.get('time_of_appointment') or 'Not specified'}</li>
                        <li><strong>Vehicle:</strong> {details.get('model')}</li>
                        <li><strong>Plate Number:</strong> {details.get('plate')}</li>
                        <li><strong>Requested Services:</strong> {services_html}</li>
                    </ul>
                    <p>Please contact us if you need to reschedule.</p>
                </body>
            </html>
            """

            msg.attach(MIMEText(html, 'html'))

            with smtplib.SMTP(self.smtp_server, self.smtp_port) as server:
                if self.use_tls:
                    server.starttls()
                server.login(self.email_from, self.email_password)
                server.send_message(msg)

            logger.info(f"Confirmation email sent to {recipient}")
        except Exception as e:
            logger.error(f"Failed to send email: {e}")
            raise

model = WhisperModel("large-v2", device="cuda", compute_type="float32")
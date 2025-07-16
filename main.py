import os
import time
from watchdog.observers import Observer
from watchdog.events import FileSystemEventHandler
from utils import ConfigManager, AudioProcessor, DataExtractor, ExcelManager, EmailSender
import sys
sys.stdout.reconfigure(encoding='utf-8')

class NewFileHandler(FileSystemEventHandler):
    """Handles new file creation events"""
    def __init__(self, config):
        self.config = config
        self.audio_processor = AudioProcessor(config.get_whisper_config('model_path'))
        self.data_extractor = DataExtractor(config)  # Pass config to use OpenAI key
        self.excel_manager = ExcelManager(config.get_path('appointments_file'))
        self.email_sender = EmailSender(config)
        
    def on_created(self, event):
        if not event.is_directory and event.src_path.lower().endswith(('.wav', '.mp3', '.mp4')):
            self.process_audio_file(event.src_path)
    
    def process_audio_file(self, file_path):
        """Process a new audio file through the entire pipeline"""
        try:
            print(f"\nProcessing file: {file_path}")

            # Step 1: Transcribe audio
            transcript = self.audio_processor.transcribe_audio(file_path)
            print("✅ Transcription complete")

            # Save transcript
            transcript_file = os.path.join(
                self.config.get_path('transcripts_dir'),
                f"{os.path.splitext(os.path.basename(file_path))[0]}.txt"
            )
            with open(transcript_file, 'w', encoding='utf-8') as f:
                f.write(transcript)
            
            # Step 2: Extract appointment data
            appointment_data = self.data_extractor.extract_info(transcript)

            if not appointment_data:
                raise ValueError("No appointment data could be extracted from transcript")
            
            print(f"✅ Extracted appointment data: {appointment_data}")
            
            # Step 3: Save to Excel
            self.excel_manager.add_appointment(appointment_data)
            print("✅ Saved to Excel")

            ## Step 4: Send confirmation email if email was found
            # if appointment_data.get('email'):
            #     self.email_sender.send_confirmation(
            #         appointment_data['email'],
            #         appointment_data
            #     )
            #     print(f"✅ Sent confirmation email to {appointment_data['email']}")
            # Step 4: Send confirmation email
            recipient_email = appointment_data.get('email') or 'cnayak70@gmail.com'
            self.email_sender.send_confirmation(recipient_email, appointment_data)
            print(f"✅ Sent confirmation email to {recipient_email}")



            
            # Move processed file
            processed_path = os.path.join(
                self.config.get_path('processed_dir'),
                os.path.basename(file_path)
            )
            os.rename(file_path, processed_path)
            print(f"✅ Moved to processed folder: {processed_path}")
            
        except Exception as e:
            print(f"❌ Error processing file {file_path}: {str(e)}")

def main():
    """Main function to start the file watcher"""
    config_path = "c:/Users/Administrator/Desktop/Chirag/car_dealer/config.ini"
    config = ConfigManager(config_path)

    print(f"Loading config from: {config_path}")
    print(f"Loaded sections: {config.config.sections()}")

    # Create directories if they don't exist
    for dir_key in ['audio_input_dir', 'processed_dir', 'transcripts_dir']:
        os.makedirs(config.get_path(dir_key), exist_ok=True)
    
    # Initialize and start file watcher
    event_handler = NewFileHandler(config)
    observer = Observer()
    observer.schedule(
        event_handler,
        path=config.get_path('audio_input_dir'),
        recursive=False
    )
    observer.start()
    
    print("✅ Appointment automation system started. Watching for new audio files...")
    print(f"📂 Input directory: {config.get_path('audio_input_dir')}")
    
    try:
        while True:
            time.sleep(1)
    except KeyboardInterrupt:
        observer.stop()
    observer.join()

if __name__ == "__main__":
    main()

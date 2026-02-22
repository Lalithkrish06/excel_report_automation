"""
Configuration Management - Simplified
"""
import os

class Config:
    # Folders
    INPUT_FOLDER = 'input_files'
    OUTPUT_FOLDER = 'output'
    
    # Email (Optional - can leave empty)
    SENDER_EMAIL = os.getenv('SENDER_EMAIL', '')
    SENDER_PASSWORD = os.getenv('SENDER_PASSWORD', '')
    SMTP_SERVER = 'smtp.gmail.com'
    SMTP_PORT = 587
    RECIPIENT_EMAIL = ''
    
    # Company
    COMPANY_NAME = 'Business Corp'
    
    @classmethod
    def validate(cls):
        errors = []
        
        # Create folders if not exist
        if not os.path.exists(cls.OUTPUT_FOLDER):
            os.makedirs(cls.OUTPUT_FOLDER)
        
        # Check input folder
        if not os.path.exists(cls.INPUT_FOLDER):
            errors.append(f"❌ Input folder '{cls.INPUT_FOLDER}' not found!")
        
        return errors

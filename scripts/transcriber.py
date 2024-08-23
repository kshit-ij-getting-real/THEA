import speech_recognition as sr
import pandas as pd

def transcribe_audio(file_path):
    recognizer = sr.Recognizer()
    with sr.AudioFile(file_path) as source:
        audio = recognizer.record(source)
    text = recognizer.recognize_google(audio)
    return text

def store_transcription(transcription, excel_path):
    # Load the existing Excel file
    excel_data = pd.ExcelFile(excel_path)
    df = pd.read_excel(excel_data, sheet_name='Responses')
    
    # Add the transcription to the DataFrame
    new_row = {'Transcription': transcription}
    df = df.append(new_row, ignore_index=True)
    
    # Write the updated DataFrame back to Excel
    with pd.ExcelWriter(excel_path, engine='openpyxl', mode='a', if_sheet_exists='replace') as writer:
        df.to_excel(writer, sheet_name='Responses', index=False)

def main(file_path, excel_path):
    transcription = transcribe_audio(file_path)
    store_transcription(transcription, excel_path)
    print("Transcription completed and stored in Excel.")

# Example usage
audio_file = "audio_files/sample_audio.wav"  # Replace with your audio file path
excel_file = "Customer Interaction Template.xlsx"  # Replace with your Excel file path
main(audio_file, excel_file)
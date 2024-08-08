import os
import csv
import whisper
from openpyxl import load_workbook

ROOT_DIR="C:\\Users\\dastarr\\Microsoft\\Mastering the Marketplace - Documents\\on-demand"
CSV_META_DATA_FILE_PATH = "C:\\Users\\dastarr\\Microsoft\\Mastering the Marketplace - Documents\\on-demand\\video-meta-data.csv"
TRANSCRIPT_OUTPUT_DIR = "C:\\Users\\dastarr\\Microsoft\\Mastering the Marketplace - Documents\\on-demand\\transcripts"

whisper_model = None

def find_file(file_name, directory_to_search):
    for root, dirs, files in os.walk(directory_to_search):
        if file_name in files:
            return os.path.join(root, file_name)
    return None

def get_files_to_process():

    print("get_files_to_process")

    files_to_process = []
    
    with open(CSV_META_DATA_FILE_PATH, mode='r') as csvfile:
        csvreader = csv.DictReader(csvfile)
        for row in csvreader:
            file_info = {
                "title": row["Title"],
                "file_name": row["Filename"],
                "url": row["URL"]
            }
            files_to_process.append(file_info)
    
    return files_to_process

def process_files(files_to_process):
    print("process_files")
    
    for file in files_to_process:
        file_name = file.get('file_name')
        title = file.get('title')
        url = file.get('url')

        full_path = find_file(file_name, ROOT_DIR)

        print ("====================================")
        print("Processing: " + file_name)

        if full_path is None:
            print("File not found: " + file_name)
            continue

        transcript = transcribe_audio(full_path)
        transcript = add_meta_data(title, url, transcript)
        transcript = clean_up_transcription(transcript)

        transcript_file_name = file_name + '.txt'
        save_file(transcript_file_name, transcript)

def transcribe_audio(path):
    print("transcribe_audio")
    
    result = whisper_model.transcribe(path)
    text = result["text"]
    
    return text

def clean_up_transcription(transcription):
    print("Cleaning transcription")

    transcription = transcription.replace("aka.ms slash mastering the marketplace", "https://aka.ms/masteringthemarketplace")
    transcription = transcription.replace("SAS", "SaaS")
    transcription = transcription.replace("Star ", "Starr ")
    transcription = transcription.replace("Star.", "Starr.")
    transcription = transcription.replace("Stark", "Starr")
    
    return transcription

def add_meta_data(video_title, url, transcription):
    
    # add the title to the top of the transcription
    transcription = f"Title: {video_title}\n\n{transcription}\n\nCitation Links:\n{url}"

    return transcription

def save_file(file_name, transcript):
    print("Saving file: " + file_name)
    
    # get the destination file info
    txt_file_path = os.path.join(TRANSCRIPT_OUTPUT_DIR, file_name)

    # delete any existing file
    if os.path.isfile(txt_file_path):
        os.remove(txt_file_path)

    # Save the transcription to a text file
    with open(txt_file_path, 'w') as txt_file:
        txt_file.write(transcript)

    return txt_file_path

if __name__ == "__main__":
    
    whisper_model = whisper.load_model("base")

    files_to_process = get_files_to_process()

    process_files(files_to_process)
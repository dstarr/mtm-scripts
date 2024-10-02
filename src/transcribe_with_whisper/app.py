import os
import csv
import whisper
from dotenv import load_dotenv
# from openpyxl import load_workbook

load_dotenv()

ROOT_DIR = os.getenv("ROOT_DIR")
CSV_META_DATA_FILE_PATH = os.getenv("CSV_META_DATA_FILE_PATH")
TRANSCRIPT_OUTPUT_DIR = os.getenv("TRANSCRIPT_OUTPUT_DIR")

print("===========================================================")

if ROOT_DIR is None:
    raise ValueError("Environment variable ROOT_DIR is not set")

if CSV_META_DATA_FILE_PATH is None:
    raise ValueError("Environment variable CSV_META_DATA_FILE_PATH is not set")

if TRANSCRIPT_OUTPUT_DIR is None:
    raise ValueError("Environment variable TRANSCRIPT_OUTPUT_DIR is not set")

print(f"ROOT_DIR {ROOT_DIR}")
print(f"CSV_META_DATA_FILE_PATH {CSV_META_DATA_FILE_PATH}")
print(f"TRANSCRIPT_OUTPUT_DIR {TRANSCRIPT_OUTPUT_DIR}")
print("===========================================================")

whisper_model = None

def find_file(file_name, directory_to_search):
    for root, dirs, files in os.walk(directory_to_search):
        if file_name in files:
            return os.path.join(root, file_name)
    return None

def get_files_to_process():

    files_to_process = []
    
    with open(CSV_META_DATA_FILE_PATH, mode='r') as csvfile:
        csvreader = csv.DictReader(csvfile)
        for row in csvreader:

            # skip files that are not marked for transcription
            transcribe_this_file = bool(row["Transcribe"])
            if not transcribe_this_file:
                continue

            file_info = {
                "learning_path": row["Learning Path"],
                "title": row["Title"],
                "file_name": row["Filename"],
                "url": row["URL"]
            }

            # add this file to the list of files to process
            files_to_process.append(file_info)
    
    return files_to_process

def process_files(files_to_process):
    
    for file in files_to_process:
        learning_path = file['learning_path']
        file_name = file['file_name']
        title = file['title']
        url = file['url']

        full_path = find_file(file_name, ROOT_DIR)

        print ("====================================")

        if full_path is None: # can't find the file
            print("File not found: " + learning_path + " : " + file_name)
            continue

        print("Processing: " + title + "\n" + full_path)

        transcript = transcribe_audio(full_path)
        transcript = add_meta_data(title, url, transcript)
        transcript = clean_up_transcription(transcript)

        transcript_file_name = file_name.removesuffix(".mp4")
        transcript_file_name += '.txt'

        save_file(transcript_file_name, learning_path, transcript)

def transcribe_audio(path):

    result = whisper_model.transcribe(path)
    text = result["text"]
    
    return text

def clean_up_transcription(transcription):

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

def save_file(file_name, learning_path, transcript):
    
    # create the output directory if it doesn't exist
    if not os.path.exists(os.path.join(TRANSCRIPT_OUTPUT_DIR, learning_path)):
        os.makedirs(os.path.join(TRANSCRIPT_OUTPUT_DIR, learning_path))

    # get the destination file path
    txt_file_path = os.path.join(TRANSCRIPT_OUTPUT_DIR, learning_path, file_name)

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
    if len(files_to_process) == 0:
        print("No files to process.")
        exit()

    process_files(files_to_process)
import os
import whisper
from openpyxl import load_workbook

DIR_TO_PARSE="C:\\Users\\dastarr\\Microsoft\\Mastering the Marketplace - Documents\\on-demand\\mastering-saas-accelerator\\video"
EXCEL_META_DATA_FILE_PATH = "C:\\Users\\dastarr\\Microsoft\\Mastering the Marketplace - Documents\\on-demand\\video-meta-data.xlsx"
TRANSCRIPT_OUTPUT_DIR = "C:\\Users\\dastarr\\Microsoft\\Mastering the Marketplace - Documents\\on-demand\\transcripts"

whisper_model = None

def transcribe_audio(path):
    print("transcribe_audio")
    
    result = whisper_model.transcribe(path)
    text = result["text"]
    
    return text

def process_videos(input_path):
    items = os.listdir(input_path)
    files = [item for item in items 
             if os.path.isfile(os.path.join(input_path, item))
             and item.endswith('.mp4')]

    for file_name in files:
        full_path = os.path.join(input_path, file_name)
        process_a_video(file_name, full_path)

def process_a_video(file_name, full_path):
    
    print ("====================================")
    print ("FILE NAME:\t" + file_name)

    # meta data from the spreadsheet
    _, title, url = get_metadata_from_spreadsheet(file_name)

    if title is None:
        print("NO METADATA FOUND FOR: " + file_name)
        return

    print("TITLE:\t\t" + title)
    
    # # transcribe the audio, add meta data, and clean the output
    transcription = transcribe_audio(full_path)
    transcription = add_meta_data(title, url, transcription)
    transcription = clean_up_transcription(transcription)
    
    transcript_file_name = file_name + '.txt'
    
    save_file(transcript_file_name, transcription)

def clean_up_transcription(transcription):
    print("Cleaning transcription")

    transcription = transcription.replace("aka.ms slash mastering the marketplace", "https://aka.ms/masteringthemarketplace")
    transcription = transcription.replace("SAS", "SaaS")
    transcription = transcription.replace("Star ", "Starr ")
    transcription = transcription.replace("Star.", "Starr.")
    transcription = transcription.replace("Stark", "Starr")
    
    return transcription

# # uses a specially formatted excel file to add meta data to the transcription
def get_metadata_from_spreadsheet(file_name):

    learning_path = None
    title = None
    url = None

    # load the excel file
    wb = load_workbook(filename=EXCEL_META_DATA_FILE_PATH)
    ws = wb.active

    for row in ws.iter_rows(min_col=1, max_col=6, values_only=True):
        if row[2] == file_name:
            learning_path = row[0]
            title = row[1]
            url = row[3]
            break
    
    wb.close()
            
    return learning_path, title, url

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
    process_videos(DIR_TO_PARSE)
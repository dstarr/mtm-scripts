import os
import whisper
from openpyxl import load_workbook
from azure.storage.blob import BlobServiceClient

DIR_TO_PARSE = "C:\\Users\\dastarr\\Microsoft\\Mastering the Marketplace - Documents\\on-demand\\mastering-saas-offers\\video"

BLOB_STORE_CONNECTION_STRING = "DefaultEndpointsProtocol=https;AccountName=st2yegwmcouglsa;AccountKey=fdLmUY5HiL4AkWHcZGB7AWj69/SMwSDjGS1iIcf+URrklahp9YZV5KEUZFmv2gbcia1JudMm8/Qi+AStC7x7bA==;EndpointSuffix=core.windows.net"
BLOB_STORE_CONTAINER_NAME = "mtm-video-scripts"
EXCEL_META_DATA_FILE_PATH = "C:\\Users\\dastarr\\Microsoft\Mastering the Marketplace - Documents\\on-demand\\video-meta-data.xlsx"
TRANSCRIPT_OUTPUT_DIR = "C:\\Users\\dastarr\\Microsoft\\Mastering the Marketplace - Documents\\on-demand\\transcripts"

def transcribe_audio(path):
    
    print("Transcribing")
    
    model = whisper.load_model("base")
    result = model.transcribe(path)
    text = result["text"]
    
    return text

def process_videos(input_path):
    items = os.listdir(input_path)
    files = [item for item in items 
             if os.path.isfile(os.path.join(input_path, item))
             and item.endswith('.mp4')]

    for file_name in files:
        # get the base file name without extension
        file_name_root = file_name.rsplit('.', 1)[0]
        full_path = os.path.join(input_path, file_name)

        process_a_video(file_name_root, full_path)

def process_a_video(file_name_root, full_path):                
    print ("====================================")
    print ("PROCESSING: " + file_name_root)

    # meta data from the spreadsheet
    playlist_name, title, url, file_name_prefix = get_metadata_from_spreadsheet(file_name_root)

    if title is None:
        print("NO METADATA FOUND FOR: " + file_name_root)
        return

    print("TITLE: " + title)
    
    # transcribe the audio, add meta data, and clean the output
    transcription = transcribe_audio(full_path)
    transcription = add_meta_data(content_category=playlist_name, video_title=title, youtube_video_link=url, transcription=transcription)
    transcription = clean_up_transcription(transcription)
    
    txt_file_name = file_name_prefix + file_name_root + '.txt'
    
    path_to_txt_file = save_file(txt_file_name, transcription)
    upload_to_blob(txt_file_name, path_to_txt_file)

def clean_up_transcription(transcription):
    print("Cleaning transcription")

    transcription = transcription.replace("aka.ms slash mastering the marketplace", "https://aka.ms/masteringthemarketplace")
    transcription = transcription.replace("SAS", "SaaS")
    transcription = transcription.replace("Star ", "Starr ")
    transcription = transcription.replace("Star.", "Starr.")
    
    return transcription

# uses a specially formatted excel file to add meta data to the transcription
def get_metadata_from_spreadsheet(file_name):
    print("Getting metadata from spreadsheet")

    title = None
    youtube_video_link = None
    playlist_name = None
    file_name_prefix = None

    # load the excel file
    wb = load_workbook(filename=EXCEL_META_DATA_FILE_PATH)
    ws = wb.active

    for row in ws.iter_rows(min_col=1, max_col=6, values_only=True):
        # match on the filename root in col 4
        if row[4] == file_name:
            playlist_name = row[0]
            title = row[2]
            youtube_video_link = row[3]
            file_name_prefix = row[5]

            break
    
    wb.close()
            
    return playlist_name, title, youtube_video_link, file_name_prefix

def add_meta_data(content_category, video_title, youtube_video_link, transcription):
    print("Adding metadata to transcription")
    
    # add the title to the top of the transcription
    transcription = "content: " + transcription
    transcription = "category: " + content_category + "\n\n" + transcription
    transcription = "url: " + youtube_video_link + "\n\n" + transcription
    transcription = "title: " + video_title + "\n\n" + transcription

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

def upload_to_blob(file_name, file_path):

    print("Uploading blob: " + file_name)

    # get the blob client
    blob_service_client = BlobServiceClient.from_connection_string(BLOB_STORE_CONNECTION_STRING)
    container_client = blob_service_client.get_container_client(BLOB_STORE_CONTAINER_NAME)
    blob_client = container_client.get_blob_client(file_name)

    # delete any existing blob
    if blob_client.exists():
        blob_client.delete_blob()

    # upload the file
    with open(file_path, "rb") as f:
        blob_client.upload_blob(f)

def DEBUG_upload_files_to_blob():
    for dirpath, dirnames, filenames in os.walk(TRANSCRIPT_OUTPUT_DIR):
        for file_name in filenames:
            if file_name.endswith('.txt'):
                # get the base file name without extension
                file_name_root = file_name.rsplit('.', 1)[0]
                full_path = os.path.join(dirpath, file_name)

                upload_to_blob(file_name_root, full_path)

if __name__ == "__main__":
    process_videos(DIR_TO_PARSE)
    # DEBUG_upload_files_to_blob()

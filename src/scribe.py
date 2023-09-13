import os
import whisper
from openpyxl import load_workbook
from azure.storage.blob import BlobServiceClient


DIR_TO_PARSE = "C:\\Users\\dastarr\\Microsoft\\Mastering the Marketplace - Documents\\on-demand\\mastering-ma-offers\\video"
OUTPUT_DIR = "C:\\Users\\dastarr\\Microsoft\\Mastering the Marketplace - Documents\\on-demand\\transcripts"
EXCEL_META_DATA_FILE_PATH = "C:\\Users\\dastarr\\Microsoft\Mastering the Marketplace - Documents\\on-demand\\video-meta-data.xlsx"
BLOB_STORE_CONNECTION_STRING = "DefaultEndpointsProtocol=https;AccountName=st2yegwmcouglsa;AccountKey=fdLmUY5HiL4AkWHcZGB7AWj69/SMwSDjGS1iIcf+URrklahp9YZV5KEUZFmv2gbcia1JudMm8/Qi+AStC7x7bA==;EndpointSuffix=core.windows.net"
BLOB_STORE_CONTAINER_NAME = "tmp-video-transcripts"

def transcribe_audio(path):
    
    print(f"Transcribing {path}...")
    
    model = whisper.load_model("base")
    result = model.transcribe(path)
    text = result["text"]
    
    return text

def process_videos(input_path, output_dir):
    for dirpath, dirnames, filenames in os.walk(input_path):
        for file_name in filenames:
            if file_name.endswith('.mp4'):
                # get the base file name without extension
                file_name_root = file_name.rsplit('.', 1)[0]
                full_path = os.path.join(dirpath, file_name)

                transcription = transcribe_audio(full_path)
                transcription = add_meta_data(file_name_root, transcription)
                transcription = clean_transcription(transcription)
                
                path_to_txt_file = save_file(file_name_root, transcription)
                upload_to_blob(file_name_root, path_to_txt_file)

def clean_transcription(transcription):
    print("Cleaning transcription")

    transcription = transcription.replace("SAS", "SaaS")
    transcription = transcription.replace("Star ", "Starr ")
    transcription = transcription.replace("Star.", "Starr.")
    
    return transcription

# uses a specially formatted excel file to add meta data to the transcription
def add_meta_data(file_name, transcription):
    print("Adding metadata to " + file_name)
    
    # load the excel file
    wb = load_workbook(filename=EXCEL_META_DATA_FILE_PATH)
    ws = wb.active

    for row in ws.iter_rows(min_col=1, max_col=5, values_only=True):
        # match on the filename root in col 5
        if row[4] == file_name:
            title = row[2]
            youtube_video_link = row[3]

            # add the title to the top of the transcription
            transcription = "Content: " + transcription
            transcription = "URL: " + youtube_video_link + "\n\n" + transcription
            transcription = "Title: " + title + "\n\n" + transcription

            break
    
    wb.close()
            
    return transcription

def save_file(file_name, transcript):
    print("Saving file: " + file_name)
    
    # get the destination file info
    txt_file_name = file_name + '.txt'
    txt_file_path = os.path.join(OUTPUT_DIR, txt_file_name)

    # delete any existing file
    if os.path.isfile(txt_file_path):
        os.remove(txt_file_path)

    # Save the transcription to a text file
    with open(txt_file_path, 'w') as txt_file:
        txt_file.write(transcript)

    return txt_file_path

def upload_to_blob(file_name_root, file_path):
    print("Uploading blob: " + file_name_root)

    # get the blob client
    blob_service_client = BlobServiceClient.from_connection_string(BLOB_STORE_CONNECTION_STRING)
    container_client = blob_service_client.get_container_client(BLOB_STORE_CONTAINER_NAME)
    blob_client = container_client.get_blob_client(file_name_root)

    # delete any existing blob
    if blob_client.exists():
        blob_client.delete_blob()

    # upload the file
    with open(file_path, "rb") as f:
        blob_client.upload_blob(f)

def DEBUG_upload_files_to_blob():
    for dirpath, dirnames, filenames in os.walk(OUTPUT_DIR):
        for file_name in filenames:
            if file_name.endswith('.txt'):
                # get the base file name without extension
                file_name_root = file_name.rsplit('.', 1)[0]
                full_path = os.path.join(dirpath, file_name)

                upload_to_blob(file_name_root, full_path)

if __name__ == "__main__":
    process_videos(DIR_TO_PARSE, OUTPUT_DIR)
    # DEBUG_upload_files_to_blob()

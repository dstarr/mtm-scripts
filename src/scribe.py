import os
import whisper

DIR_TO_PARSE = "../raw/"

def transcribe_audio(path):
    
    print(f"Transcribing {path}...")
    
    model = whisper.load_model("base")
    result = model.transcribe(path)
    text = result["text"]
    
    return text

def process_directory(path):
    for dirpath, dirnames, filenames in os.walk(path):
        for file_name in filenames:
            if file_name.endswith('.mp4'):
                full_path = os.path.join(dirpath, file_name)
                transcription = transcribe_audio(full_path)

                # get the destination file info
                txt_file_name = file_name.rsplit('.', 1)[0] + '.txt'
                txt_file_path = os.path.join(dirpath, txt_file_name)

                delete_file_if_exists(txt_file_path)
                
                # Save the transcription to a text file
                with open(txt_file_path, 'w') as txt_file:
                    txt_file.write(transcription)
                    
def delete_file_if_exists(file_path):
    if os.path.isfile(file_path):
        os.remove(file_path)
        
if __name__ == "__main__":
    process_directory(DIR_TO_PARSE)



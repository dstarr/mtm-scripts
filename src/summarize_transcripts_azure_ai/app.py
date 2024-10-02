import os
from dotenv import load_dotenv
from azure.ai.textanalytics import TextAnalyticsClient, AbstractiveSummaryAction
from azure.core.credentials import AzureKeyCredential

# Example method for summarizing text
def create_summary(client, content):
    
    poller = client.begin_analyze_actions(
        documents=[content],
        actions=[
            AbstractiveSummaryAction(max_sentence_count=3)
        ],
        language="en"
    )

    document_results = poller.result()
    
    for result in document_results:
        extract_summary_result = result[0]  # first document, first result
        
        if extract_summary_result.is_error:
            print("...Is an error with code '{}' and message '{}'".format(
                extract_summary_result.code, extract_summary_result.message
            ))
            return None
        else:
            return extract_summary_result.summaries[0].text

def find_txt_files(directory):
    
    txt_files = []
    for root, _, files in os.walk(directory):
        for file in files:
            if file.endswith('.txt'):
                txt_files.append(os.path.join(root, file))
                
    return txt_files

def main():
    load_dotenv()

    # read the environment variables
    key = os.environ.get('LANGUAGE_KEY')
    endpoint = os.environ.get('LANGUAGE_ENDPOINT')
    transcription_root_dir = os.environ.get('TRANSCRIPTION_ROOT_DIR')

    # create the client
    credential = AzureKeyCredential(key)
    client = TextAnalyticsClient(endpoint=endpoint, credential=credential)

    # get the paths to all the transcription files
    txt_files_paths = find_txt_files(transcription_root_dir)
    
    for file_path in txt_files_paths:
        print(file_path)
        
        with open(file_path, 'r') as file:
            transcription_text = file.read()
            summary = create_summary(client, transcription_text)
            print(summary)
        
        print("==============================")
    
    print("DONE!")
    
if __name__ == "__main__":
    main()


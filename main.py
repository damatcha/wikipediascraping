import os
import time
import pandas as pd
import wikipedia
from flask import Flask
import threading
import webbrowser

REQUEST_DELAY = 30  
app = Flask(__name__)

excel_file_path = None  

@app.route('/')
def home():
    threading.Thread(target=process_excel_file).start()
    return 'Processing started! You can close this window.'

def search_wikipedia_pages(name):
    urls = []
    try:
        search_results = wikipedia.search(name, results=50)
        print(f"Search results for '{name}': {search_results}")

        for result in search_results:
            page = wikipedia.page(result)
            if name in page.title:  
                urls.append(page.url)
                print(f"Found page: {page.title} - URL: {page.url}")

    except wikipedia.exceptions.DisambiguationError as e:
        print(f"DisambiguationError for name '{name}': {e.options}")
    except wikipedia.exceptions.PageError as e:
        print(f"PageError for name '{name}': {e}")
    except Exception as e:
        print(f"Error searching for name '{name}': {e}")

    return urls

def process_excel(file_path, region):
    try:
        df = pd.read_excel(file_path, engine='openpyxl')
        print("Columns in the Excel file:", df.columns)
    except Exception as e:
        print(f"Error reading Excel file: {e}")
        return

    name_column = 'NAMES'  
    if name_column not in df.columns:
        print(f"Error: '{name_column}' column not found in the Excel file.")
        return

    results = []

    all_names = df[name_column].dropna().tolist() 
    wikipedia.set_lang(region)  

    for i in range(0, len(all_names), 10):
        names_batch = all_names[i:i + 10]  
        print(f"Processing batch {i // 10 + 1}: Names {i + 1} to {i + len(names_batch)}")

        for name in names_batch:
            print(f"Searching for '{name}' on Wikipedia...")
            urls = search_wikipedia_pages(name)
            if urls:
                for url in urls:
                    results.append({'NAME': name, 'EXACT WEB SOURCE': url})
            else:
                results.append({'NAME': name, 'EXACT WEB SOURCE': ''})

        time.sleep(REQUEST_DELAY)  

    result_df = pd.DataFrame(results)
    output_path = 'wikipedia_results.xlsx'
    result_df.to_excel(output_path, index=False)
    print(f"Results saved to {output_path}")

def process_excel_file():
    global excel_file_path
    region = 'en'  

    if os.path.isfile(excel_file_path):
        print(f"Processing file: {excel_file_path}")
        process_excel(excel_file_path, region)
    else:
        print("Excel file not found. Exiting.")

def main():
    global excel_file_path
    excel_file_path = input("Enter the path to the Excel file: ")

    def run_server():
        print("Starting Flask server on http://127.0.0.1:XXXX")
        app.run(host='127.0.0.1', port=XXXX)

 
    server_thread = threading.Thread(target=run_server)
    server_thread.start()

    webbrowser.open(f'http://127.0.0.1:XXXX/')

    input("Press Enter once the processing is done or to exit...")

    if os.path.isfile(excel_file_path):
        process_excel_file()
    else:
        print("Excel file not found. Exiting.")

if __name__ == '__main__':
    main()

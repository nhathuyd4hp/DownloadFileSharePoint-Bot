import os
import logging
import pandas as pd
from bot import WebAccess
from bot import SharePoint
from concurrent.futures import ThreadPoolExecutor


logging.basicConfig(
    level=logging.INFO,
    format='%(asctime)s - %(name)s - %(levelname)s - %(message)s',
    encoding='utf-8',
    datefmt="%Y-%m-%d %H:%M:%S",
    handlers=[
        logging.FileHandler('bot.log', mode='a', encoding='utf-8'),
        logging.StreamHandler(),
    ],
)
# Config
HEADLESS = False
TIMEOUT = 5
MAX_THREAD = 5
# Download Path
WA_DOWNLOAD_PATH = os.path.join(os.path.dirname(os.path.abspath(__file__)),"WebAccess")
SP_DOWNLOAD_PATH = os.path.join(os.path.dirname(os.path.abspath(__file__)),"SharePoint")

def download(csv_file):
    path = os.path.join(WA_DOWNLOAD_PATH, csv_file)
    df = pd.read_csv(path,encoding="CP932")
    
    result = SharePoint(
        username="vietnamrpa@nskkogyo.onmicrosoft.com",
        password="Robot159753",
        timeout=TIMEOUT,
        headless=HEADLESS,
        logger_name="SharePoint",
        download_directory= os.path.join(SP_DOWNLOAD_PATH,csv_file),
    ).download_file(
        file_pattern="割付図・エクセル/.*.pln$",
        site_url = df['資料リンク'][:2].to_list(),
    )
    print(f"Result: {result}")

    
if __name__ == "__main__":
    WA = WebAccess(
        username="2909",
        password="159753",
        headless= HEADLESS,
        timeout=TIMEOUT,
        logger_name="WebAccess",
        download_directory= WA_DOWNLOAD_PATH,
    )
    for ビルダー名 in ["014400","026600","001705","004100"]:
        WA.get_information(
            ビルダー名 = ビルダー名,
            図面 = ["作図済","送付済","CBUP済","CB送付済"],
            確定納品日 = ["2025/03/26","2025/04/26"],
            FIELDS=["案件番号","得意先名","物件名","資料リンク"],
            OUTPUT_FILE=os.path.join(WA_DOWNLOAD_PATH,f'{ビルダー名}.csv'),
        )
    del WA
          
    csv_files = [file for file in os.listdir(WA_DOWNLOAD_PATH) if file.endswith(".csv")]
    for i in range(0, len(csv_files), MAX_THREAD):
        batch = csv_files[i:i+MAX_THREAD]
        with ThreadPoolExecutor(max_workers=MAX_THREAD) as executor:
            executor.map(download, batch)
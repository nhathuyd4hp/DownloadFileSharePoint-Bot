import os
import logging
from bot import WebAccess
from bot import SharePoint
import pandas as pd

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

HEADLESS = False
TIMEOUT = 5

WA_DOWNLOAD_PATH = os.path.join(os.path.dirname(os.path.abspath(__file__)),"WebAccess")
SP_DOWNLOAD_PATH = os.path.join(os.path.dirname(os.path.abspath(__file__)),"SharePoint")

if __name__ == "__main__":
    # WA = WebAccess(
    #     username="2909",
    #     password="159753",
    #     headless= HEADLESS,
    #     timeout=TIMEOUT,
    #     logger_name="WebAccess",
    #     download_directory= WA_DOWNLOAD_PATH,
    # )
    # for ビルダー名 in ["014400","026600","001705","004100"]:
    #     WA.get_information(
    #         ビルダー名 = ビルダー名,
    #         図面 = ["作図済","送付済","CBUP済","CB送付済"],
    #         確定納品日 = ["2025/03/26","2025/04/26"],
    #         FIELDS=["案件番号","得意先名","物件名","資料リンク"],
    #         OUTPUT_FILE=f"{WA_DOWNLOAD_PATH}\{ビルダー名}.csv",
    #     )
    # del WA
    
    result = SharePoint(
        username="vietnamrpa@nskkogyo.onmicrosoft.com",
        password="Robot159753",
        timeout=TIMEOUT,
        headless=HEADLESS,
        logger_name="SharePoint",
        download_directory= SP_DOWNLOAD_PATH,
    ).download_file(
        file_pattern="割付図・エクセル/.*.pln$",
        site_url = [
            'https://nskkogyo.sharepoint.com/:f:/s/2019/EtnwHQtR9G9Do6l_rFcpENABIl_UJZrJQJtBPzjJoNtYmw',
            'https://nskkogyo.sharepoint.com/sites/2019/_layouts/15/viewer.aspx?sourcedoc={16398997-0f91-4b25-85bb-adcd95ba1454}',
            'https://nskkogyo.sharepoint.com/sites/2019/_layouts/15/viewer.aspx?sourcedoc={a0391c66-e0c7-4d1c-8415-955dd356f1d4}',
            'https://nskkogyo.sharepoint.com/sites/2019/_layouts/15/viewer.aspx?sourcedoc={2a586df0-b2e8-44a0-b17f-0804b66c9fe1}',
            'https://nskkogyo.sharepoint.com/sites/2019/_layouts/15/viewer.aspx?sourcedoc={70c21d0b-af39-4464-a333-72cee443ef5b}',
            'https://nskkogyo.sharepoint.com/sites/2019/_layouts/15/viewer.aspx?sourcedoc={86024541-7ed5-4e0e-8381-37d172d4f871}',
            'https://nskkogyo.sharepoint.com/:f:/s/2019/EufjhvDxQcFHs-Lx3wOzY3UBsIZ5dG8lePEkVTgGkN3T0Q',
            'https://nskkogyo.sharepoint.com/sites/2019/_layouts/15/viewer.aspx?sourcedoc={d412c823-5626-4bea-bc65-a090c21ee4cc}',
            'https://nskkogyo.sharepoint.com/sites/2019/_layouts/15/viewer.aspx?sourcedoc={3323e71e-dcce-4121-b4fe-804a1f01f056}',
            'https://nskkogyo.sharepoint.com/sites/2019/_layouts/15/viewer.aspx?sourcedoc={907ca56b-354e-4e32-982a-c5ba1c25243f}'
        ]
    )
    print(result)
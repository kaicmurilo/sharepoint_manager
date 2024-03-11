import os, logging
from sharepoint import SharePoint
from domain import SharepointSettings

logging.basicConfig(level=logging.INFO)

if __name__ == "__main__":
    config = SharepointSettings(
        site_url=os.environ.get("SP_SITE_URL"),
        site_path=os.environ.get("SP_SITE_PATH"),
        tenant_id=os.environ.get("SP_TENANT_ID"),
        client_id=os.environ.get("SP_CLIENT_ID"),
        client_secret=os.environ.get("SP_CLIENT_SECRET"),
    )

    sp = SharePoint(config)
    file_path = os.environ.get("SP_FILE_PATH")
    folder_path = os.environ.get("SP_FOLDER_PATH")
    local_save_path = os.environ.get("SP_LOCAL_SAVE_PATH")

    try:
        sp.upload_file_to_sharepoint(file_path, folder_path)
        sp.download_file_from_sharepoint(file_path, local_save_path)
    except Exception as e:
        logging.error(e)

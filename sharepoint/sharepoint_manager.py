import requests
import os
import logging

# Configure o logging
logging.basicConfig(level=logging.INFO)


class SharePoint:
    def __init__(self, config):
        """
        Initializes the class with the provided configuration.

        Args:
            config: The configuration object containing the URL, path, tenant ID, client secret, and client ID.

        Returns:
            None
        """
        self.site_url = config.url
        self.site_path = config.path
        self.tenant_id = config.tenant_id
        self.client_secret = config.client_secret
        self.client_id = config.client_id
        self.access_token = self.get_access_token()
        self.site_id = self.get_site_id()
        self.graph_url_principal = "https://graph.microsoft.com"
        self.graph_url_default = f"{self.graph_url_principal}/.default"
        self.graph_utils = f"{self.graph_url_principal}/v1.0/sites/"
        self.graph_url = f"{self.graph}{self.site_id}/drive/root:/"
        self.root_url = f"{self.graph}{self.site_id}/drive/root/children"
        self.api_url = f"{self.graph}{self.site_url}:{self.site_path}"
        self.token_url = (
            f"https://login.microsoftonline.com/{self.tenant_id}/oauth2/v2.0/token"
        )

    def get_access_token(self):
        """
        Method to retrieve the access token using client credentials and Microsoft Graph API.
        """

        token_data = {
            "grant_type": "client_credentials",
            "scope": self.graph_url_default,
            "client_id": self.client_id,
            "client_secret": self.client_secret,
        }
        token_response = requests.post(self.token_url, data=token_data)
        access_token = token_response.json().get("access_token")
        return access_token

    def get_site_id(self):
        """
        Retrieves the site ID using the access token and site URL.

        Parameters:
            self (obj): The instance of the class.

        Returns:
            str: The ID of the site, or None if the request fails.
        """
        headers = {
            "Authorization": f"Bearer {self.access_token}",
            "Content-Type": "application/json",
        }

        response = requests.get(self.api_url, headers=headers)
        if response.status_code == 200:
            return response.json().get("id", None)
        else:
            response.raise_for_status()

    def list_folders_in_root(self):
        """
        Lista todas as pastas na raiz do site do SharePoint.

        Args:
        - access_token: O token de acesso para autenticação.
        """

        headers = {"Authorization": f"Bearer {self.access_token}"}
        response = requests.get(self.root_url, headers=headers)

        if response.status_code == 200:
            items = response.json().get("value", [])
            print("Pastas na raiz:")
            for item in items:
                if "folder" in item:
                    print(f"Nome: {item['name']}")
            return items
        else:
            print(f"Erro ao listar pastas: {response.status_code}")

    def list_directory_contents(self, folder_path):
        headers = {"Authorization": f"Bearer {self.access_token}"}
        list_folder_url = f"{self.graph_url}{folder_path}:/children"
        response = requests.get(list_folder_url, headers=headers)
        if response.status_code == 200:
            items = response.json().get("value", [])
            for item in items:
                print(
                    f"Name: { item['name']} | Type: {'Folder' if 'folder' in item else 'File'}"
                )
        else:
            print(f"Erro ao acessar a pasta: {response.status_code}")

        return items

    def upload_file_to_sharepoint(self, file_path, folder_path):
        file_size = os.path.getsize(file_path)
        with open(file_path, "rb") as file:
            chunk_size = 4 * 1024 * 1024  # Read 4MB at a time
            while True:
                data = file.read(chunk_size)
                if not data:
                    break

                # Create a new file in SharePoint
                file_url = f"{self.graph_url}{folder_path}/{os.path.basename(file_path)}:/content"
                headers = {
                    "Authorization": f"Bearer {self.access_token}",
                    "Content-Length": str(len(data)),
                    "Content-Range": f"bytes 0-{file_size - 1}/{file_size}",
                }

                # Send the file chunk to SharePoint
                response = requests.put(file_url, headers=headers, data=data)

                if response.status_code in [200, 201]:
                    logging.info(
                        f"Arquivo '{os.path.basename(file_path)}' copiado com sucesso para '{folder_path}'."
                    )
                else:
                    response.raise_for_status()

    def download_file_from_sharepoint(self, file_path, local_save_path):

        download_url = f"{self.graph_url}{file_path}:/content"
        headers = {"Authorization": f"Bearer {self.access_token}"}
        response = requests.get(download_url, headers=headers, stream=True)

        if response.status_code == 200:
            with open(local_save_path, "wb") as file:
                for chunk in response.iter_content(chunk_size=1024):
                    if chunk:  # filter out keep-alive new chunks
                        file.write(chunk)
            print(f"Arquivo baixado com sucesso: {local_save_path}")
        else:
            print(f"Erro ao baixar o arquivo: {response.status_code}")

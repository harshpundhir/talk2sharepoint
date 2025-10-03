'''My libraries'''
from params import sharepoint_email, sharepoint_url_site, sharepoint_site_name, sharepoint_doc_library
from setpwd import setpwd
setpwd()

'''Python libraries'''
import os
sharepoint_password = os.environ.get('sharepoint_password')
import office365
from urllib import response
from office365.sharepoint.client_context import ClientContext
from office365.runtime.auth.user_credential import UserCredential
from office365.sharepoint.files.file import File
import datetime



USERNAME = sharepoint_email
PASSWORD = sharepoint_password
SHAREPOINT_SITE = sharepoint_url_site
SHAREPOINT_SITE_NAME = sharepoint_site_name
SHAREPOINT_DOC = sharepoint_doc_library

class SharePoint:
    def _auth(self):
        conn = ClientContext(SHAREPOINT_SITE).with_credentials(
            UserCredential(
                USERNAME,
                PASSWORD
            )
        )
        return conn

    def _get_files_list(self, folder_name):
        conn = self._auth()
        target_folder_url = f'{SHAREPOINT_DOC}/{folder_name}'
        root_folder = conn.web.get_folder_by_server_relative_url(target_folder_url)
        root_folder.expand(["Files", "Folders"]).get().execute_query()
        return root_folder.files
    
    def get_folder_list(self, folder_name):
        conn = self._auth()
        target_folder_url = f'{SHAREPOINT_DOC}/{folder_name}'
        root_folder = conn.web.get_folder_by_server_relative_url(target_folder_url)
        root_folder.expand(["Folders"]).get().execute_query()
        return root_folder.folders

    def download_file(self, file_name, folder_name):
        conn = self._auth()
        file_url = f'/sites/{SHAREPOINT_SITE_NAME}/{SHAREPOINT_DOC}/{folder_name}/{file_name}'
        file = File.open_binary(conn, file_url)
        return file.content
    
sh = SharePoint()
print(sh.get_folder_list('General'))
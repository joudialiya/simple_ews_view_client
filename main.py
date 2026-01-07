import requests
from requests_ntlm import HttpNtlmAuth
from elements import get_folder, get_folder_by_id, find_folder_by_id, find_item, get_item, get_attachment
import xml.etree.ElementTree as ET
import typing
from pprint import pprint
import html
import os
import base64

ENV_PATH = ".env"

def parse_env():
    if not os.path.exists(ENV_PATH):
        return
    with open(ENV_PATH, "r") as f:
        vars = f.readlines()
        for line in vars:
            line = line.strip()
            if not line or line[0] == "#":
                continue
            key, value = line.split("=")
            key = key.strip()
            value = value.strip()
            os.environ[key] = value

parse_env()

NS_T = "http://schemas.microsoft.com/exchange/services/2006/types"
NS_M = "http://schemas.microsoft.com/exchange/services/2006/types"

namespaces = {
  "t": NS_T,
  "m": NS_M
}

HASH = os.environ.get("HASH", None)
BASE_URL = os.environ.get("BASE_URL", None)
USERNAME = os.environ.get("USERNAME", None)

TMP = os.environ.get("TMP", None)

if not os.path.isdir(TMP):
    os.mkdir(TMP)

def get_path(name: str):
    return os.path.join(TMP, name.replace("/", ""))

def save(name:str, content: bytes):
    path = get_path(name)

    if os.path.exists(path):
        print("File exists", name)

    with open(path, "+wb") as f:
        f.write(content)

class View():
  def exec_command(self, cli: "CMD", command:str, args: str):
    print(f"** Unknow command {command}")

class FolderView(View):
    def __init__(self, name: str = None, id: str = None, change_id: str = None):
        super().__init__()
        self.response_items = None
        self.response_folders = None
        self.response_folder = None
        self.name = name
        self.id = id
        self.change_id = change_id
    
    def do_refresh(self, cli: "CMD", command, args):
        if self.name:
            response = cli.post(get_folder(self.name))
        else:
            response = cli.post(get_folder_by_id(self.id, self.change_id))

        doc = ET.fromstring(response)
        self.response_folder = doc
        save(f"{self.name}_GetFolder.xml", response)

        self.name = doc.find(".//t:DisplayName", namespaces).text
        self.id = doc.find(".//t:FolderId", namespaces).attrib.get("Id")
        self.change_id =  doc.find(".//t:FolderId", namespaces).attrib.get("ChangeKey")

        print("Name:", self.name)
        print("Id:", self.id)
        print("ChangeKey:", self.change_id)

        print("[+] \"folders\" to view subfolders.")
        print("[+] \"items\" to view messages.")
        print("[+] \"back\" to go back.")


    def do_items(self, cli: "CMD", command, args):
        print(self.__class__, "refresh", self.name)
        response = cli.post(find_item(self.id, self.change_id))
        save(f"{self.name}_FindItem.xml", response)
        doc = ET.fromstring(response.decode("utf-8"))
        self.response_items = doc
        ids = [item.attrib.get("Id") for item in doc.findall(".//t:ItemId", namespaces)]
        key = [item.attrib.get("Changekey") for item in doc.findall(".//t:ItemId", namespaces=namespaces)]
        subjects = doc.findall(".//t:Subject", namespaces)
        subjects = [e.text for e in subjects]
        names = [e.text for e in doc.findall(".//t:From/t:Mailbox/t:Name", namespaces)]
        emails = [e.text for e in doc.findall(".//t:From/t:Mailbox/t:EmailAddress", namespaces)]
        for i, s in enumerate(subjects):
            print(f"[{i}] {s} [From] {names[i]} ({emails[i]})")

    def do_folders(self, cli: "CMD", command, args):
        print(self.__class__, "refresh", self.name)
        response = cli.post(find_folder_by_id(self.id, self.change_id))
        save(f"{self.name}_FindFolder.xml", response)
        doc = ET.fromstring(response)
        self.response_folders = doc
        folders = doc.findall(".//t:Folder", namespaces)
        if not folders:
            print("No sub folders ...")
        for i in range(0, len(folders)):
            name = folders[i].find(".//t:DisplayName", namespaces).text
            print(f"[{i}]", name)

    def do_view(self, cli: "CMD", command, args):
        print(self.__class__, "view item")
        i = int(args, 10)
        ids = [(
            e.attrib.get("Id"),
            e.attrib.get("ChangeKey")
            ) for e in self.response_items.findall(".//t:ItemId", namespaces)]
        if i < 0 or i >= len(ids):
            print("Index out of range")
            return
        cli.view = ItemView(*ids[i])

    def do_enter(self, cli: "CMD", command, args):
        print(self.__class__, "enter")
        i = int(args, 10)
        ids = [(
            e.attrib.get("Id"),
            e.attrib.get("ChangeKey")
            ) for e in self.response_folders.findall(".//t:Folder/t:FolderId", namespaces)]
        if i < 0 or i >= len(ids):
            print("Index out of range")
            return
        cli.view = FolderView(id=ids[i][0], change_id=ids[i][1])

    def do_back(self, cli: "CMD", command, args):
        print(self.__class__, "back")
        if self.response_folder is None:
            print("Please refresh ...")
            return
        id = self.response_folder.find(".//t:ParentFolderId", namespaces).attrib.get("Id")
        change_key = self.response_folder.find(".//t:ParentFolderId", namespaces).attrib.get("ChangeKey")
        if not id:
            print("No parrent ...")
        cli.view = FolderView(id=id, change_id=change_key)

    def exec_command(self, cli: "CMD", command: str, args: str):
        if command == "items":
            self.do_items(cli, command, args)
        elif command == "folders":
            self.do_folders(cli, command, args)
        elif command == "view":
            self.do_view(cli, command, args)
        elif command == "enter":
            self.do_enter(cli, command, args)
        elif command == "back":
            self.do_back(cli, command, args)
        elif command == "refresh":
            self.do_refresh(cli, command, args)
        else:
            super().exec_command(cli, command, args)

class ItemView(View):
    def __init__(self, id: str, change_key: str):
        super().__init__()
        self.id = id
        self.change_key = change_key
        self.response = None

    def do_refresh(self, cli: "CMD", command: str, args: str):
        response = cli.post(get_item(self.id, self.change_key))
        # save resp
        save(f"{self.id}.xml", response)
        self.response = ET.fromstring(response)
        self.print_msg()

    def print_msg(self):
        print("[+] Enter \"body\" to preview thw body of the message")
        ids = self.response.findall(".//t:FileAttachment/t:AttachmentId", namespaces)
        ids = [e.attrib.get("Id") for e in ids]
        names = self.response.findall(".//t:FileAttachment/t:Name", namespaces)
        names = [e.text for e in names]
        print("[+] Enter \"attch + [num]\" to preview the attachment")
        for i in range(0, len(ids)):
            print(f"\t[{i}]", names[i])

    def do_body (self, cli, command, args):
        print(self.__class__, "body")
        body = html.unescape(self.response.find(".//t:Message/t:Body", namespaces).text)
        body = body.encode()
        save(f"{self.id}.html", body)
        os.startfile(get_path(f"{self.id}.html"))
        print(f"Openning msg body ...")

    def do_attch(self, cli: "CMD", command: str, args: str):
        print(self.__class__, "attch")
        i = int(args, 10)
        ids = self.response.findall(".//t:FileAttachment/t:AttachmentId", namespaces)
        ids = [e.attrib.get("Id") for e in ids]

        names = self.response.findall(".//t:FileAttachment/t:Name", namespaces)
        names = [e.text for e in names]

        if i < 0 or i >= len(names):
            print("Index out of range")
        
        if not os.path.exists(os.path.join(TMP, names[i])):
            print("Retrieving from server ...")
            resp = cli.post(get_attachment(ids[i]))
            doc = ET.fromstring(resp.decode("utf-8"))
            save(f"{ids[i]}.xml", resp)

            name = doc.find(".//t:FileAttachment/t:Name", namespaces).text
            content = doc.find(".//t:FileAttachment/t:Content", namespaces).text
            content = base64.b64decode(content)
            save(name, content)
        print(f"Openning {names[i]} ...")
        os.startfile(get_path(names[i]))

    def exec_command(self, cli: "CMD", command: str, args: str):
        if command == "body":
            self.do_body(cli, command, args)
        elif command == "attch":
            self.do_attch(cli, command, args)
        elif command == "back":
            print("back ...")
            cli.view = FolderView("inbox")
        elif command == "refresh":
            self.do_refresh(cli, command, args)
        elif command == "id":
            print(self.id)
            print(self.change_key)
        else:
            super().exec_command(cli, command, args)

class CMD():
    def __init__(self, username, password):
        self.username = username
        self.password = password
        self.is_running = True
        self.view:View = FolderView("inbox")
        self.connect()

    def connect(self):
        self.session = requests.Session()
        self.session.auth = HttpNtlmAuth(self.username, self.password)

    def post(self, envolope: str) -> bytes:
        headers = {
            "Content-Type": "text/xml; charset=utf-8",
            "Accept": "text/xml",
        }
        response = self.session.post(
            url=BASE_URL,
            data=envolope,
            headers=headers
        )
        if not response.ok:
            raise Exception("POST error")
        return response.content
    def loop(self):
        while self.is_running:
            prompt = input("> ")
            prompt = prompt.strip()
            if prompt == "":
                continue

            parts = prompt.split(" ")
            command = parts[0]
            args = ""
            if len(parts) > 1:
                args = " ".join(parts[1:])
            self.exec_command(command, args)
    def exec_command(self, command: str, args: str):
        if command == "bye":
            print("bye ...")
            self.is_running = False
        elif command == "info":
            print("[username]", self.username) 
            print("[password]", self.password)
        else:
            self.view.exec_command(self, command, args)

if __name__ == "__main__":
    cli = CMD(USERNAME, HASH)
    cli.loop()


def get_folder(name: str):
  return f"""<?xml version="1.0" encoding="utf-8"?>
<soap:Envelope
  xmlns:soap="http://schemas.xmlsoap.org/soap/envelope/"
  xmlns:t="http://schemas.microsoft.com/exchange/services/2006/types">
  <soap:Body>
    <GetFolder
      xmlns="http://schemas.microsoft.com/exchange/services/2006/messages"
      xmlns:t="http://schemas.microsoft.com/exchange/services/2006/types">
      <FolderShape>
        <t:BaseShape>Default</t:BaseShape>
        <t:AdditionalProperties>
          <t:FieldURI FieldURI="folder:ParentFolderId"/>
        </t:AdditionalProperties>
      </FolderShape>
      <FolderIds>
        <t:DistinguishedFolderId Id="{name}"/>
      </FolderIds>
    </GetFolder>
  </soap:Body>
</soap:Envelope>
"""

def get_folder_by_id(id: str, change_key: str):
  return f"""<?xml version="1.0" encoding="utf-8"?>
<soap:Envelope
  xmlns:soap="http://schemas.xmlsoap.org/soap/envelope/"
  xmlns:t="http://schemas.microsoft.com/exchange/services/2006/types">
  <soap:Body>
    <GetFolder
      xmlns="http://schemas.microsoft.com/exchange/services/2006/messages"
      xmlns:t="http://schemas.microsoft.com/exchange/services/2006/types">
      <FolderShape>
        <t:BaseShape>Default</t:BaseShape>
        <t:AdditionalProperties>
          <t:FieldURI FieldURI="folder:ParentFolderId"/>
        </t:AdditionalProperties>
      </FolderShape>
      <FolderIds>
        <t:FolderId Id="{id}" ChnangeKey="{change_key}"/>
      </FolderIds>
    </GetFolder>
  </soap:Body>
</soap:Envelope>
"""

def find_folder_by_id(id: str, change_key: str):
  return f"""<?xml version="1.0" encoding="utf-8"?>
<soap:Envelope
  xmlns:soap="http://schemas.xmlsoap.org/soap/envelope/"
  xmlns:t="http://schemas.microsoft.com/exchange/services/2006/types">
  <soap:Body>
    <FindFolder Traversal="Shallow" xmlns="http://schemas.microsoft.com/exchange/services/2006/messages">
      <FolderShape>
        <t:BaseShape>Default</t:BaseShape>
        <t:AdditionalProperties>
          <t:FieldURI FieldURI="folder:ParentFolderId"/>
        </t:AdditionalProperties>
      </FolderShape>
      <ParentFolderIds>
        <t:FolderId Id="{id}" ChnangeKey="{change_key}"/>
      </ParentFolderIds>
    </FindFolder>
  </soap:Body>
</soap:Envelope>"""


def find_item(id: str, change_key: str):
  return  f"""<?xml version="1.0" encoding="utf-8"?>
<soap:Envelope
  xmlns:soap="http://schemas.xmlsoap.org/soap/envelope/"
  xmlns:t="http://schemas.microsoft.com/exchange/services/2006/types">
  <soap:Body>
    <FindItem
      xmlns="http://schemas.microsoft.com/exchange/services/2006/messages"
      xmlns:t="http://schemas.microsoft.com/exchange/services/2006/types"
      Traversal="Shallow">
      <ItemShape>
        <t:BaseShape>Default</t:BaseShape>
      </ItemShape>
      <ParentFolderIds>
        <t:FolderId Id="{id}" ChnangeKey="{change_key}"/>
      </ParentFolderIds>
    </FindItem>
  </soap:Body>
</soap:Envelope>"""

def get_item(item_id, change_key):
  return f"""<?xml version="1.0" encoding="utf-8"?>
<soap:Envelope
  xmlns:xsi="http://www.w3.org/2001/XMLSchema-instance"
  xmlns:xsd="http://www.w3.org/2001/XMLSchema"
  xmlns:soap="http://schemas.xmlsoap.org/soap/envelope/"
  xmlns:t="http://schemas.microsoft.com/exchange/services/2006/types">
  <soap:Body>
    <GetItem
      xmlns="http://schemas.microsoft.com/exchange/services/2006/messages"
      xmlns:t="http://schemas.microsoft.com/exchange/services/2006/types">
      <ItemShape>
        <t:BaseShape>Default</t:BaseShape>
        <t:IncludeMimeContent>true</t:IncludeMimeContent>
        <AdditionalProperties>
          <FieldURI FieldURI="item:Attachments"/>
        </AdditionalProperties>
      </ItemShape>
      <ItemIds>
        <t:ItemId Id="{item_id}" ChangeKey="{change_key}" />
      </ItemIds>
    </GetItem>
  </soap:Body>
</soap:Envelope>"""

def get_attachment(id: str):
  return f"""<?xml version="1.0" encoding="utf-8"?>
<soap:Envelope
  xmlns:xsi="http://www.w3.org/2001/XMLSchema-instance"
  xmlns:xsd="http://www.w3.org/2001/XMLSchema"
  xmlns:soap="http://schemas.xmlsoap.org/soap/envelope/"
  xmlns:t="http://schemas.microsoft.com/exchange/services/2006/types">
  <soap:Body>
    <GetAttachment
      xmlns="http://schemas.microsoft.com/exchange/services/2006/messages"
      xmlns:t="http://schemas.microsoft.com/exchange/services/2006/types">
      <AttachmentShape/>
      <AttachmentIds>
        <t:AttachmentId Id="{id}"/>
      </AttachmentIds>
    </GetAttachment>
  </soap:Body>
</soap:Envelope>"""
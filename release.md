emailbrowser-api SGCustoms 
[20/01/2023] v1.6 

## NEW FEATURE
-

## IMPROVED
- Search Sorting by SentDate will sorted the dates and time in based on sorting value
- Remove logic which extract email attachments into a folder in Content Server

## FIXED
- Fix Search Attachment name with Extraction Attachment when the email is not extracted
- Fix Html charset of html header to UTF-8 in Email HTMLviewer API

## NOTE
- The 2 default search option in front end will get it from 1st and 2nd value from config 'FolderUnderEnterprise'.
- need to add nodeID of 'ArchiveFolderID' in config


-------------------------------------------------------------------------------------------

emailbrowser-api SGCustoms 
[20/01/2023] v0.5

## NEW FEATURE
- Add configuration for folder to show on 1st level under enterprise
- Add API to Find folder by Name and filter by defined folder in enterprise 
- Add Default Folder for Search. it get value from same configuration on folder under enterprise
- Default search will include search in Archive / Migrated folder. the folder is defined in config 'ArchiveFolderID'.

## IMPROVED
- Improved API for get folder by name with db query
- Improved API ListSubFolder with CSAccess.ListNodesByPage. it will return folder only and exclude documents.
- Improved API getListFolderByName with permssion check and folder path in single db query

## FIXED
- Fixed fuzzy_search error handle when the result is 0 and return response null. previously it return error 500.

## NOTE
- The 2 default search option in front end will get it from 1st and 2nd value from config 'FolderUnderEnterprise'.
- need to add nodeID of 'ArchiveFolderID' in config



-------------------------------------------------------------------------------------------

emailbrowser-api SGCustoms 
[05/01/2023] v0.4

## NEW FEATURE
- Add email viewer in HTML response
- Add multiple Search Location parameter. 
- Add Get folder info from Folder Path

## IMPROVED
- Reduced loading on search with Backend Pagination in search / advance search
- update SXCommon Lib dll from SXCommmon SGCustom branch

## FIXED
- fixed get conversation api
- Check embeded attachment for .msg attachments

## NOTE
- Included Ale API for OTCS PATCH pat222008003
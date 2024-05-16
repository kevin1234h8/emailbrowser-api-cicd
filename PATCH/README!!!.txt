This is patch for content server
- Copy pat222008003.txt to OTHOME/patch
- open OTHOME/config/opentext.ini file, write the required setting bellow

[SGCustom_ReadMSG]
hostname=  (PUT HOSTNAME HERE, EXAMPLE: 192.168.1.12, or customapp, or localhost)
urlRead=/api/v1/attachments/listAttahcments
urlDownload=/api/v1/attachments/downloadAttachment
urlReadEmbedded=/api/v1/attachments/getEmbeddedAttachment
timeout=90
port= (SET PORT HERE, EXAMPLE: 80, 8080, 25637)
--drop table EmailBrowser_ExtractedEmails
CREATE TABLE EmailBrowser_ExtractedEmails(
	 CSID bigint NOT NULL
	,CSName nvarchar(248) NOT NULL
	,CSVersionNum bigint NULL
	,CSModifyDate datetime NOT NULL
	,CSParentID bigint NOT NULL
	,CSCachingFolderNodeID bigint NOT NULL
	,IsExtracted bit NOT NULL DEFAULT 0
	,IsExtractingNow bit NOT NULL DEFAULT 0
	,LastExtractedDate datetime NULL
	CONSTRAINT PK_EmailBrowser_ExtractedEmails_CSID PRIMARY KEY CLUSTERED 
	(
		CSID
	)
) 
select * from EmailBrowser_ExtractedEmails


--drop table EmailBrowser_ExtractedEmailAttachments

CREATE TABLE EmailBrowser_ExtractedEmailAttachments(
	 CSEmailID bigint NOT NULL
	,FileHash binary(16) NULL
	,CSID bigint NOT NULL
	,FileName nvarchar(248) NOT NULL
	,FileType nvarchar(248) NOT NULL
	,FileSize bigint NOT NULL
	,Deleted bit NOT NULL DEFAULT 0
) 
CREATE CLUSTERED INDEX CI_EmailBrowser_ExtractedEmailAttachments ON EmailBrowser_ExtractedEmailAttachments(CSEmailID, CSID)

select * from EmailBrowser_ExtractedEmailAttachments

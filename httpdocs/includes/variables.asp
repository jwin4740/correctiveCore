<%

'## all global variables

Const hdAdminPath = "/hdadmin/"
Const hdPathToDriverTemplate = "includes/hddrivertemplate.asp"
Const jwVideoPath = "/videos/"

'## to be added to the <body> tag for an online event if needed
hdAddBodyOnLoad = ""

'## Set file I/O constants.
Const ForReading   = 1
Const ForWriting   = 2
Const ForAppending = 8

Const hdMasterUsersEmailAddr = "seh@halsteaddesign.net,blabone@halsteaddesign.net"

'## category constants
Const hdNEWScat = 1
Const hdBLOGcat = 2
Const hdPROJECTScat = 3
Const hdGALLERYcat = 4
Const hdBANNERScat = 5
Const hdCALENDARcat = 6
Const hdTESTIMONYcat = 7
Const hdCONTACTScat = 8

'## Site Settings
Const setMgtPages = "setMgtPages"
Const setMgtBlog = "setMgtBlog"
Const setMgtCalendar = "setMgtCalendar"
Const setMgtContact = "setMgtContact"
Const setMgtGallery = "setMgtGallery"
Const setMgtNews = "setMgtNews"
Const setMgtProjects = "setMgtProjects"
Const setMgtBanners = "setMgtBanners"
Const setMgtMailer = "setMgtMailer"
Const setMgtTestimony = "setMgtTestimony"

'## SQL Paging vars
hdRsDefaultPageSize = 15
QueryStringEnd = ""
SortQueryStringEnd = ""
LocalQueryStringEnd = ""
hsSQLWHERE = ""
hsSQLORDERBY = ""
hdRSRecordTypesTitle = "Displaying Records"
hdRSPagingLineBreakCount = 15

'## ASPupload vars
Const hdUploadTempDir = "/assets/temp/"
Const hdPDFDir = "/assets/pdf/"
Const hdBannerDir = "/assets/bnr/"
Const hdGalleryDir = "/assets/image/"

'## Banner variables
Const hdBannerW = 350       '## width
Const hdBannerH = 110       '## height
Const hdDefaultBannerImage = "banner.jpg"

'## Gallery variables
Const hdGalleryW = 600       '## width
Const hdGalleryH = 338       '## height
Const hdGalleryThumbW = 300
Const hdGalleryThumbH = 169

%>

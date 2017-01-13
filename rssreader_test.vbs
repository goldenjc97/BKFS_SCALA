' ---------------------------------------------------------------------------
'
' RSS Reader for ScalaScript
'
' ---------------------------------------------------------------------------
'
' Copyright 2004-2009, Scala, Inc.
'
' Permission is granted to use this script or derived works in
' Scala InfoChannel applications and installations.
'
' ---------------------------------------------------------------------------
'
' This script reads RSS feeds and presents the data to a matching ScalaScript.
' It supports the common RSS formats (0.91, 1.0, and 2.0).  
' The most important RSS fields are supported.  Most of the
' rest would be straightforward to add.
'
' ---------------------------------------------------------------------------
'
' Modifications by Marc Rifkin
'
' Additions to the HTML character replace list
' The XML file is saved locally so it doesn't have to be re-downloaded 
' A new XML file is only retrieved if the online version is newer than the last downloaded one
' HTTP header tags added to force proxy servers to not cache the file
' If the feed has <media:content> or <enclosure> tags within items, the files are downloaded for each corresponding item (used by Photo RSS feeds and some news feeds)
' If the feed has <image><url> tags within channels, the file are downloaded for each corresponding item (used by some feeds for a channel logo)
' Function ReplaceRegEx is a handy search/replace function using Regular Expression
' Function CleanFileName to generate Windows-safe filenames from a given string
' Function FindFromList searches a string for a list of substrings and return the first match if found 
' Function ChopText cuts off a text string to the nearest word 
' Function InsertReturns inserts carriage returns every so many characters
' Function GetFileDate to find out the last modified date of a file
' Function ParseRSSDate converts RFC822-formatted date to local format
' Functions RSBinaryToString and DownloadFile to allow file retrieval via HTTP (for feeds with links to images)
' Function GetItemImage downloads an item's image, using DownloadFile
' Function CreateFolder creates a folder, delete if already exists
' Function GetParseError retrieves the error details from an XML object after loading
' Function GetScalaVersion checks what version Scala is running
' Variable RSSFeedName is the filename to save the local XML file and folder containing it (if empty, one will be created) 
' Varible ChannelImageStatus = 1 if a channel image was found, 0 if not
' Variable ChannelImage which contains the name of the channel's image (if available)
' Variable ItemImage which contains the name of the items's image (if available)
' Variable ItemImageStatus = 1 if item images are found, 0 if not
' Variable ItemImageUpdate is an integer incremented to signal Scala script when an image is ready
' Variable HideItemImages is a boolean indicating whether item images should be downloaded or not
' Variable HideChannelImage is a boolean indicating whether the channel image should be downloaded or not
' Variable UseItemDescriptionImage forces the use of the img found in the <description> of the item rather than <media:content> or <enclosure>
' Variable HideItemLinkText removes any text in <a> </a> blocks 
' Variable HideItemDates is a boolean indicating whether to show item dates or not 
' Variable DownloadFrequency determines how often to check for updates (in minutes), -1 = disable downloads
' Variable ProxyServer and ProxyPort lets you specify a proxy server and port number (if a proxy server is required)
' Variables evType, evLocation, evStartDate, evEndDate, evOrganizer for RSS calendar feeds
' Variable Style_title is used to set font, color, etc. for titles
' Variable Style_description is used to set font, color, etc. for descriptions                
' Variable Style_date is used to set font, color, etc. for dates                
' Variable Style_gap is used to add a gap space to each line   
' Variables MaxLengthTitle and MaxLengthDescription set the max length allowed for titles and descriptions            
'
' ---------------------------------------------------------------------------
'
' Control variables
'
' Inputs:
' - RSSFeedURL: URL of the RSS feed
' - RSSFeedName: Name to save the local XML file and folder containing it (if empty, one will be created) 
' - LoopForever: Boolean indicating if the VBScript should loop through the items forever (no longer shared).
' - MaxItems: Integer, set to non-zero to limit to that many items from the feed (0 = show all)
' - HideChannelTitle: Boolean, set to true to hide the Channel title 
' - HideChannelDescription: Boolean, set to true to hide the Channel description 
' - HideItemTitles: Boolean, set to true to hide the ItemTitle 
' - HideItemDescriptions: Boolean, set to true to hide the description 
' - HideItemDates: Boolean, set to true to hide the date 
' - MaxLengthTitle: String, max length allowed for titles 
' - MaxLengthDescription: String, max length allowed for descriptions 
' - HideItemImages: Boolean indicating whether item images should be downloaded or not
' - HideChannelImage: Boolean indicating whether the channel image should be downloaded or not
' - UseItemDescriptionImage: Boolean forces the use of the img found in the <description> of the item rather than <media:content> or <enclosure>
' - HideItemLinkText: Boolean removes any text in <a> </a> blocks 
' - DownloadFrequency: Integer determines how often to check for updates (minutes)
' - ProxyServer: String, host or IP of proxy server (if one is required)
' - ProxyPort: String, port number for the proxy server
' - Style_title: String, set font, color, etc. for titles
' - Style_description: String, set font, color, etc. for items                
' - Style_date: String, set font, color, etc. for dates                
' - Style_gap: String, set font, color, etc. for the space between items
'
' Outputs:
' - ChannelErrorNum: Returns non-zero if there was an error loading the channel.
'
' Channel information is returned in the following variables shared:
' - ChannelTitle: The RSS channel/title element
' - ChannelLink: The RSS channel/link element
' - ChannelDescription: The RSS channel/description element
' - ChannelLanguage: The RSS channel/language element
' - ChannelCopyright: The RSS channel/copyright element
' - ChannelDate: The RSS channel/pubDate element
' - ChannelImage: The name of the channel's image (if available)
' - ChannelImageStatus: Whether the channel has an image downloaded (0 or 1)
'
' For each item, information is returned in the following shared variables:
' - ItemTitle: The RSS item/title element
' - ItemLink: The RSS item/link element
' - ItemDescription: The RSS item/description element
' - ItemDate: The RSS item/pubDate element, or if not found, the item/dc:date element
' - ItemImage: Name of image (if available)
' - ItemImageStatus: Whether the item has an image downloaded (0 or 1)
' - evType: Calendar event type (if available)
' - evLocation: Calendar event location (if available
' - evStartdate: Calendar event start date (if available)
' - evEnddate: Calendar event end date (if available)
' - evOrganizer: Calendar event organizer (if available)
'
' Input/Outputs:
' - ChannelRequest: ScalaScript should set to 1 to request the channel,
'       and the VBScript sets it to zero when ready.
' - ItemRequest: ScalaScript should set to 1 request the next item,
'       and the VBScript sets it to zero when ready.  (Use as the crawl
'       "Cue Variable")
' - ItemImageUpdate: Integer, this variable is incremented every time a new item
'       image is ready.

' ---------------------------------------------------------------------------
'
' Funtions and Subroutines
' 
' Search for an XML node and return its text
Function GetElementText(xmlObject, tagname)
    Dim xmlElement
    Set xmlElement = xmlObject.selectSingleNode("*[local-name()='" & tagname & "']")
    If (Not xmlElement is Nothing) Then
        GetElementText = xmlElement.text
    Else
        GetElementText = ""
    End If
End Function

' Call GetElementText then remove HTML and white space
Function GetCleanedElementText(xmlObject, tagname)
    text = GetElementText(xmlObject, tagname)
    text = CleanHTMLMarkup(text)
    text = CleanupWhitespace(text)
    GetCleanedElementText = text
End Function

' Clean up HTML markup included in some RSS feeds
Function CleanHTMLMarkup(text)
    ' Replace any hexadecimal entities
    regEx.Pattern = "&#x[0-9a-f]+;"
    While regEx.Test(text)
        Set Matches = regEx.Execute(text)
        text = Left(text, Matches(0).FirstIndex) & Chrw("&H" + Mid(text, Matches(0).FirstIndex + 4, Matches(0).Length-4)) & Mid(text, Matches(0).FirstIndex + Matches(0).Length + 1)
    Wend
    ' Replace any decimal entities
    regEx.Pattern = "&#[0-9]+;"
    While regEx.Test(text)
        Set Matches = regEx.Execute(text)
        text = Left(text, Matches(0).FirstIndex) & Chrw(Mid(text, Matches(0).FirstIndex + 3, Matches(0).Length-3)) & Mid(text, Matches(0).FirstIndex + Matches(0).Length + 1)
    Wend
    ' Replace named HTML entities
    text = Replace(text, "&Aacute;",chr(193))
    text = Replace(text, "&aacute;",chr(225))
    text = Replace(text, "&Acirc;",chr(194))
    text = Replace(text, "&acirc;",chr(226))
    text = Replace(text, "&acute;",chr(180))
    text = Replace(text, "&AElig;",chr(198))
    text = Replace(text, "&aelig;",chr(230))
    text = Replace(text, "&Agrave;",chr(192))
    text = Replace(text, "&agrave;",chr(224))
    text = Replace(text, "&apos;", "'")
    text = Replace(text, "&Aring;",chr(197))
    text = Replace(text, "&aring;",chr(229))
    text = Replace(text, "&Atilde;",chr(195))
    text = Replace(text, "&atilde;",chr(227))
    text = Replace(text, "&Auml;",chr(196))
    text = Replace(text, "&auml;",chr(228))
    text = Replace(text, "&brvbar;",chr(166))
    text = Replace(text, "&Ccedil;",chr(199))
    text = Replace(text, "&ccedil;",chr(231))
    text = Replace(text, "&cedil;",chr(184))
    text = Replace(text, "&cent;", "¢")
    text = Replace(text, "&cent;",chr(162))
    text = Replace(text, "&copy;", "©")
    text = Replace(text, "&copy;",chr(169))
    text = Replace(text, "&curren;", "¤")
    text = Replace(text, "&curren;",chr(164))
    text = Replace(text, "&deg;", "°")
    text = Replace(text, "&deg;",chr(176))
    text = Replace(text, "&divide;",chr(247))
    text = Replace(text, "&eacute;", "é")
    text = Replace(text, "&Eacute;",chr(201))
    text = Replace(text, "&eacute;",chr(233))
    text = Replace(text, "&Ecirc;",chr(202))
    text = Replace(text, "&ecirc;",chr(234))
    text = Replace(text, "&Egrave;",chr(200))
    text = Replace(text, "&egrave;",chr(232))
    text = Replace(text, "&ETH;",chr(208))
    text = Replace(text, "&eth;",chr(240))
    text = Replace(text, "&Euml;",chr(203))
    text = Replace(text, "&euml;",chr(235))
    text = Replace(text, "&euro;", "€")
    text = Replace(text, "&frac12;",chr(189))
    text = Replace(text, "&frac14;",chr(188))
    text = Replace(text, "&frac34;",chr(190))
    text = Replace(text, "&gt;", ">")
    text = Replace(text, "&iacute;", "í")
    text = Replace(text, "&Iacute;",chr(205))
    text = Replace(text, "&iacute;",chr(237))
    text = Replace(text, "&Icirc;",chr(206))
    text = Replace(text, "&icirc;",chr(238))
    text = Replace(text, "&iexcl;",chr(161))
    text = Replace(text, "&Igrave;",chr(204))
    text = Replace(text, "&igrave;",chr(236))
    text = Replace(text, "&iquest;",chr(191))
    text = Replace(text, "&Iuml;",chr(207))
    text = Replace(text, "&iuml;",chr(239))
    text = Replace(text, "&laquo;", "«")
    text = Replace(text, "&laquo;",chr(171))
    text = Replace(text, "&ldquo;", "“")
    text = Replace(text, "&lsquo;", "‘")
    text = Replace(text, "&lt;", "<")
    text = Replace(text, "&macr;",chr(175))
    text = Replace(text, "&mdash;", "-")
    text = Replace(text, "&micro;",chr(181))
    text = Replace(text, "&middot;", "·")
    text = Replace(text, "&middot;",chr(183))
    text = Replace(text, "&nbsp;", " ")
    text = Replace(text, "&nbsp;",chr(160))
    text = Replace(text, "&ndash;", "-")
    text = Replace(text, "&not;",chr(172))
    text = Replace(text, "&Ntilde;",chr(209))
    text = Replace(text, "&ntilde;",chr(241))
    text = Replace(text, "&Oacute;",chr(211))
    text = Replace(text, "&oacute;",chr(243))
    text = Replace(text, "&Ocirc;",chr(212))
    text = Replace(text, "&ocirc;",chr(244))
    text = Replace(text, "&Ograve;",chr(210))
    text = Replace(text, "&ograve;",chr(242))
    text = Replace(text, "&ordf;",chr(170))
    text = Replace(text, "&ordm;",chr(186))
    text = Replace(text, "&Oslash;",chr(216))
    text = Replace(text, "&oslash;",chr(248))
    text = Replace(text, "&Otilde;",chr(213))
    text = Replace(text, "&otilde;",chr(245))
    text = Replace(text, "&Ouml;",chr(214))
    text = Replace(text, "&ouml;",chr(246))
    text = Replace(text, "&para;",chr(182))
    text = Replace(text, "&plusmn;",chr(177))
    text = Replace(text, "&pound;", "£")
    text = Replace(text, "&pound;",chr(163))
    text = Replace(text, "&quot;", """")
    text = Replace(text, "&raquo;", "»")
    text = Replace(text, "&raquo;",chr(187))
    text = Replace(text, "&rdquo;", "”")
    text = Replace(text, "&reg;", "®")
    text = Replace(text, "&reg;",chr(174))
    text = Replace(text, "&rsquo;", "’")
    text = Replace(text, "&sect;",chr(167))
    text = Replace(text, "&shy;",chr(173))
    text = Replace(text, "&sup1;",chr(185))
    text = Replace(text, "&sup2;",chr(178))
    text = Replace(text, "&sup3;",chr(179))
    text = Replace(text, "&szlig;",chr(223))
    text = Replace(text, "&THORN;",chr(222))
    text = Replace(text, "&thorn;",chr(254))
    text = Replace(text, "&times;",chr(215))
    text = Replace(text, "&Uacute;",chr(218))
    text = Replace(text, "&uacute;",chr(250))
    text = Replace(text, "&Ucirc;",chr(219))
    text = Replace(text, "&ucirc;",chr(251))
    text = Replace(text, "&Ugrave;",chr(217))
    text = Replace(text, "&ugrave;",chr(249))
    text = Replace(text, "&uml;",chr(168))
    text = Replace(text, "&Uuml;",chr(220))
    text = Replace(text, "&uuml;",chr(252))
    text = Replace(text, "&Yacute;",chr(221))
    text = Replace(text, "&yacute;",chr(253))
    text = Replace(text, "&yen;", "¥")
    text = Replace(text, "&yen;",chr(165))
    text = Replace(text, "&yuml;",chr(255))
    ' Ampersand conversion comes last!
    text = Replace(text, "&amp;", "&")
    ' A few other characters that don't seem to translate well
    text = Replace(text, chrw(146), "’")
    text = Replace(text, chrw(151), "–")
    ' Convert <br> and <p> to newline
    text = ReplaceRegEx(text, vbCrLf, "<(br|p)\b[^>]*>")
    ' Remove <a> </a> blocks
    If HideItemLinkText = "True" Then
        text = ReplaceRegEx(text, "", "<a.*</a>")
    End If
   ' Remove any other embedded markup
    text = ReplaceRegEx(text, "", "<[^>]+>")
    CleanHTMLMarkup = text
 End Function

' Collapse whitespace/newlines
Function CleanupWhitespace(text)
    ' Tabs to spaces
    text = Replace(text, vbTab, " ")
    ' Non-splitting spaces to regular spaces
    text = Replace(text, Chr(160), " ")
    ' Temporarily convert all newlines make any kind of newline into vbCr
    text = Replace(text, vbLf, vbCr)
    ' Replace a mix of spaces and newlines, with a single newline
    text = ReplaceRegEx(text, vbCr, vbCr & "[ " & vbCr & "]+")
    ' Collapse multiple spaces
    text = ReplaceRegEx(text, " ", " +")
    ' Convert newlines back to vbCrLf
    text = Replace(text, vbCr, " ")
    ' Remove dash + space at beginning
    text = ReplaceRegEx(text, "", "^- ")
    ' Trim leading and trailing spaces
    text = Trim(text)
    CleanupWhitespace = text
End Function

' Replace patrn with str2 in str1 using regex
Function ReplaceRegEx(str1, str2, patrn)
  regEx.Pattern = patrn 
  ReplaceRegEx = regEx.Replace(str1, str2) 
End Function

' Create a Windows friendly filename from a given string
Function CleanFileName(str)
    ' This is used for each feed's xml filename and sub-folder
    Dim CleanFileNameTemp
    CleanFileNameTemp = str
    ' Remove ugly protocol and file type stuff
    CleanFileNameTemp = ReplaceRegEx(CleanFileNameTemp, "", "(http://|https://|\.html|\.xml|\.rss|\.php|\.asp|\.jsp)")
    ' Remove characters not valid for Windows filenames
    CleanFileNameTemp = ReplaceRegEx(CleanFileNameTemp, "", "[\\\/:*?. \x22 <>|=]")
    CleanFileName = Left(CleanFileNameTemp,40)
End Function

' Search a string for a list of substrings and return the first match if found
Function FindFromList(list, searchtext)
    ' For example, find the file type of an item image. The URL may have extra text at the end,
    ' so you have to search for the extention, you can't just assume it's the last three chars.
    RegEx.Pattern = "(" & Join(list,"|") & ")"
    Set Matches = RegEx.Execute(searchtext)
    If Matches.count > 0 Then
        FindFromList = Matches(0)
    Else
        FindFromList = ""
    End If
End Function

' Cut off a text string to the nearest word 
Function ChopText(text, choplimit)
    If InstrRev(text, " ", choplimit) >0 Then
        ChopText = Left(text, InstrRev(text, " ", choplimit)-1) & "..." 
    ElseIf Len(text) > choplimit Then
        ChopText = Left(text, choplimit) & "..."
    Else
        ChopText = text
    End If
End Function

' Inserts carriage returns every so many characters
Function InsertReturns(text, charcount)
    Dim insert
    Dim breaktext
    insert = charcount
    Do While insert < Len(text)
        breaktext = InStrRev(text, " ", insert)
        If insert - breaktext > charcount Then
            text = Left(text, insert) & VbCrLf &  Mid(text, insert + 1)
        Else
            insert = breaktext
            text = Left(text, insert) & VbCrLf &  Mid(text, insert + 1)
        End If
        insert = insert + charcount + 2
    Loop
    InsertReturns = text
End Function

' Return the last modified date of a file
Function GetFileDate(datefile)
    Dim f
    tempGetFileDate = 0
        If fso.FileExists(datefile) Then 
            Set f = fso.GetFile(datefile)
            tempGetFileDate = f.DateLastModified
        End If
    GetFileDate = tempGetFileDate
End Function

' Convert RFC822-formatted date to local format
Function ParseRSSDate(rssdate)
    ' Note there are at least two common date formats even though RFC822 is
    ' supposed to be the standard. And there is no standard way of expressing
    ' the time zone, so you have to check for the most common formats.
    Dim sDate, tempDate, tempYear, tempMonth, tempMonthNum, tempDay, tempHour, tempMinute, tempSecond, tempZone, tempZoneOffet
    ' Break up each part of the date and time
    ' Example date: "Fri, 13 Jun 2008 16:33:50 GMT"
    RegEx.Pattern = "^([A-Za-z]{3}),\s([\d]{1,2})\s([A-Za-z]*)\s([\d]{4})\s([\d]{2}):([\d]{2}):([\d]{2})\s(.*)"
    Set Matches = RegEx.Execute(rssdate)
    If Matches.count > 0 Then
        With Matches(0)
            tempDay = .SubMatches(1)
            tempMonthName = .SubMatches(2)
            ' Month should be three letter abbreviation but some feeds have full spelling
            ' So check for both
            For i = 1 to 12
                ' Try array of names set in script
                If Left(tempMonthName,3) = Left(MonthNames(i-1),3) Then
                    tempMonthNum = i
                ' Try list of names from Windows MonthName() function (abbreviated)
                ElseIf tempMonthName = MonthName(i, True) Then tempMonthNum = i
                ' Try list of names from Windows MonthName() (full)
                ElseIf tempMonthName = MonthName(i, False) Then tempMonthNum = i
                End If
            Next
            tempYear = .SubMatches(3)
            tempHour = .SubMatches(4)
            tempMinute = .SubMatches(5)
            tempSecond = .SubMatches(6)
            tempZone = .SubMatches(7)
        End With
    Else
        'Check if alternate date format was used
        ' 2008-06-13T16:33:50.45+01:00 or  2008-06-13T16:33:50.45Z
        RegEx.Pattern = "^([\d]{4})-([\d]{2})-([\d]{2})T([\d]{2}):([\d]{2}):([\d]{2})\.*([\d]{0,2})(.*)"
        Set Matches = RegEx.Execute(rssdate)
        If Matches.count > 0 Then
            With Matches(0)
                tempYear = .SubMatches(0)
                tempMonthNum = .SubMatches(1)
                tempDay = .SubMatches(2)
                tempHour = .SubMatches(3)
                tempMinute = .SubMatches(4)
                tempSecond = .SubMatches(5)
                tempZone = .SubMatches(7)
            End With
        End If
    End If
    ' Assemble into VBScript friendly date and time values
    tempDate = DateSerial(tempYear, tempMonthNum, tempDay)
    tempTime = TimeSerial(tempHour, tempMinute, tempSecond)

    ' Calculate time zone offset based on zone specified
    ' Feeds should use GMT but not everyone follows the rules
    ' This list catches common exceptions
    Select Case tempZone
        Case "EST" tempZoneOffset = -5
        Case "EDT" tempZoneOffset = -4
        Case "CST" tempZoneOffset = -6
        Case "CDT" tempZoneOffset = -5
        Case "MST" tempZoneOffset = -7
        Case "MDT" tempZoneOffset = -6
        Case "PST" tempZoneOffset = -8
        Case "PDT" tempZoneOffset = -7
        Case "Z" tempZoneOffset = 0
        Case "A" tempZoneOffset = -1
        Case "M" tempZoneOffset = -12
        Case "N" tempZoneOffset = 1
        Case "Y" tempZoneOffset = 12
        Case "GMT"  tempZoneOffset = 0
    End Select
    
    ' Somtimes the zone offset is explicitly specified in +/-HHMM or +/-HH:MM format
    RegEx.Pattern = "(\+*\-*[\d]{2}):*([\d]{2})"
    Set Matches = RegEx.Execute(tempZone)
    If Matches.count > 0 Then tempZoneOffset = CInt(Matches(0).Submatches(0)) + 60*CInt(Matches(0).Submatches(1))        

    ' Calculate the local date and time using the offset from the feed and the local offset from Windows
    ' This provides the date and time in the local time zone of the viewer
    tempDate = dateadd("n", (0-ActiveTimeBias)-(tempZoneOffset*60), tempDate & " " & tempTime) 
    tempTime = dateadd("n", (0-ActiveTimeBias)-(tempZoneOffset*60), tempTime) 

    ' Convert to the time and date format specified in Windows Regional and Language Options
    tempDate = FormatDateTime(tempDate,1)
    tempTime = FormatDateTime(tempTime,3)
    ParseRSSDate = tempDate ' & " " & tempTime
End Function

' RSBinaryToString converts binary data (used by DownloadFile)
Function RSBinaryToString(Binary)
    ' This is a workaround since VBScript can't really handle binary
    ' data. An ADODB database object is used to store it.
    Dim RS
    Set RS = CreateObject("ADODB.Recordset")
    Dim LBinary
    Const adLongVarChar = 201
    LBinary = LenB(Binary)
    If LBinary>0 Then
        RS.Fields.Append "mBinary", adLongVarChar, LBinary
        RS.Open
        RS.AddNew
        RS("mBinary").AppendChunk Binary 
        RS.Update
        RSBinaryToString = RS("mBinary")
    Else
        RSBinaryToString = ""
    End If
    Set RS = Nothing
End Function

' Download a file
Function DownloadFile(FileURL, FileLocal)
    On Error Resume Next
    Dim DownloadFileTemp, binaryFile, textstream
    DownloadFileTemp = ""
    ' First look in the Content Folder
    ' This would allow a substitute feed to be sent from Content Manager or copied to the LocalIntegratedContent folder
    ' instead of downloading from the URL.
    DownloadFileTemp = oScalaFileLock.LockScalaFile("Content:\" & CleanFileName(RSSFeedName) & "\" & FileLocal)
    If DownloadFileTemp <> "" Then
        DownloadFile = DownloadFileTemp ' Found the file in Content folder
        oScalaFileLock.UnlockScalaFile()
    Else
        If Err.Number <> 0 Then Err.Clear
        ' Then look in the temp folder and decide if it needs to be updated
        ' Update only if Download Frequency is > 0 (-1=disabled) and either the file doesn't exist or is older than the currently downloaded xml file
        DownloadFileTemp = TempFolder & "\" & CleanFileName(RSSFeedName) & "\" & FileLocal
        If (iDownloadFrequency >= 0) and (net_online = True) and ((fso.FileExists(DownloadFileTemp) = False) or (DateDiff("n", GetFileDate(DownloadFileTemp), GetFileDate(LocalFile)) >= 0)) Then
            WinHttpReq.Open "GET", FileURL, False, authUser, authPassword
            If ProxyServer <> "" Then WinHttpReq.setProxy HTTPREQUEST_PROXYSETTING_PROXY, ProxyServer & ":" & ProxyPort
            ' These headers force proxy servers to not cache
            WinHttpReq.SetRequestHeader "Cache-Control", "no-cache" 
            WinHttpReq.SetRequestHeader "Cache-Control", "must-revalidate" 
            WinHttpReq.SetRequestHeader "If-Modified-Since", "Sat, 1 Jan 2000 00:00:00 GMT"  
            WinHttpReq.SetRequestHeader "Pragma", "no-cache"
            WinHttpReq.Send
            ' Save file if request was successful
            If Err.Number = 0 Then
                If WinHttpReq.Status = 200 Then
                    binaryFile = WinHttpReq.ResponseBody
                    ' Convert the image data to a string form that can be saved by the FileSystemObject
                    stringFile = RSBinaryToString(binaryFile)
                    ' Write the image data to the temp folder
                    Set textstream = fso.OpenTextFile(DownloadFileTemp, 2, True)
                    textstream.Write(stringFile)
                    textstream.Close
                    DownloadFile = DownloadFileTemp 
                    net_online = True
                Else 
                    ' Download failed due to HTTP error
                    DownloadFile = "" 
                    net_online = False ' Future downloads will will be disabled until the next loop 
                    oScala.Log("RSSREADER ERROR: A file download failed. Downloading will be disabled until the next loop. Feed: " & RSSFeedURL)
                    oScala.Log("RSSREADER ERROR: " & CStr(WinHttpReq.Status) & " " & WinHttpReq.StatusText)
                End If
            Else 
                ' Download failed due to Windows error
                DownloadFile = "" 
                ' Future downloads will will be disabled until the next loop
                net_online = False 
                oScala.Log("RSSREADER ERROR: A file download failed. Downloading will be disabled until the next loop.  Feed: " & RSSFeedURL)
                oScala.Log("RSSREADER ERROR: " & CStr(Err.Number) & " " & Err.Description)
            End If
        ' File still exists and is newer than the current xml file
        Elseif (fso.FileExists(DownloadFileTemp) = True) and (DateDiff("n", GetFileDate(DownloadFileTemp), GetFileDate(LocalFile)) <= 0) Then
            DownloadFile = DownloadFileTemp 
        End If
    End If
End Function

' Download an item's image
Function GetItemImage(ItemObject, ImageNum)
    ' This function tries to find the item's image
    ' Different feeds use different tags for this
    ' Or you can choose to get the first <img> in the description instead
    On Error Resume Next
    Dim ImageURL
    GetItemImage = ""
    ImageURL = ""
    ' Different feeds use different tags for images
    ' Look for a subnode with a url attribute
    Set xmlsubItem = ItemObject.selectSingleNode("*[@url != '']") 
    If (Not xmlsubItem is Nothing) Then
        If ((xmlsubItem.nodeName = "media:content") or (xmlsubItem.nodeName = "enclosure")) and ((xmlsubItem.getAttribute("type") = "image/jpg") or (xmlsubItem.getAttribute("type") = "image/jpeg") or (xmlsubItem.getAttribute("type") = "image/png") or (xmlsubItem.getAttribute("type") = "image/gif")) Then ImageURL = xmlsubItem.getAttribute("url")
    End If
    ' Sometimes the <description> has a higher quality img, so you can use this with the UseItemDescriptionImage variable
    If  UseItemDescriptionImage = "True" Then 
        ' Try src= with double quotes
        RegEx.Pattern = "src\s*=\s*""([^""\?]*)"
        Set Matches = RegEx.Execute(ItemObject.selectSingleNode("*[local-name()='description']").xml)
        If Matches.count > 0 Then 
            ImageURL = matches(0).submatches(0)
        Else
            ' Try src= with single quotes
            RegEx.Pattern = "src\s*=\s*'([^'\?]*)"
            Set Matches = RegEx.Execute(ItemObject.selectSingleNode("*[local-name()='description']").xml)
            If Matches.count > 0 Then ImageURL = matches(0).submatches(0)
        End If
    End If
    ' If we have a result, then download it
    If ImageURL <> "" Then
        ItemImageTemp = DownloadFile(ImageURL, CleanFileName(RSSFeedName) & "_itemimage_" & ImageNum & FindFromList(FileTypes, ImageURL))
        If ItemImageTemp <> "" Then 
            ' Download succeeded
            GetItemImage = ItemImageTemp
            ItemImageStatus = 1
        Else 
            ' Download failed
            GetItemImage = ""  
            ItemImageStatus = 0
        End If
    Else 
        ' No item image found 
        GetItemImage = "" 
        ItemImageStatus = 0
    End If
End Function

' Create a folder, delete if already exists
Sub CreateFolder(folderpath)
    ' This is used to create (and clear) the sub-folder for the feed
    ' It is only called when the xml file is downloaded or updated
    ' That way old item images get deleted as well
    Dim f
    If fso.FolderExists(folderpath) Then
        fso.DeleteFolder(folderpath)
    End If
    If Not fso.FolderExists(folderpath) Then
        Set f = fso.CreateFolder(folderpath)
    End If
End Sub

' Retrieve the error details from an XML object after loading
Function GetParseError(xmlObj) 
    ' Report the error number and description back to the Scala script
    Dim xPE
    ' Obtain the ParseError object
    Set xPE = xmlObj.parseError
    If xPE.errorCode <> 0 Then
        With xPE
            GetParseError = "Error #: " & .errorCode & ": " & xPE.reason & _
                         " Line #: " & .Line & _
                         " Line Position: " & .linepos &  _
                         " Position In File: " & .filepos 
        End With
    End If
End Function

' Check what version Scala is running
Function GetScalaVersion
    On Error Resume Next
    version_d = WshShell.RegRead("HKEY_LOCAL_MACHINE\SOFTWARE\Scala\InfoChannel Designer 5\ProductVersion")
    version_p = WshShell.RegRead("HKEY_LOCAL_MACHINE\SOFTWARE\Scala\InfoChannel Player 5\ProductVersion")
    If version_d > version_p Then
        GetScalaVersion = version_d
    Else
        GetScalaVersion = version_p
    End If
    If Err.Number <> 0 Then Err.Clear
End Function

' ---------------------------------------------------------------------------
'
' Main Program
'
On Error Resume Next

' Scala Player object
Dim oScala
Set oScala = CreateObject("ScalaPlayer.ScalaPlayer.1")

' Scala FileLock object
Dim oScalaFileLock
Set oScalaFileLock = CreateObject("ScalaFileLock.ScalaFileLock.1")

' XMLHTTP object for downloading the RSS XML file
Err.Clear
Dim downloadXML
Set downloadXML = CreateObject("Msxml2.XMLHTTP.4.0") 
' XMLHTTP 5.0 is required for functions used in this script
If Err.Number = 429 Then oScala.Log("RSSREADER ERROR: XMLHTTP 4.0 not installed.")

' MSXML object to save the XML to a local file
Err.Clear
Dim saveXML
Set saveXML = CreateObject("Msxml2.DOMDocument.4.0") 
saveXML.async = False
' MSXML 4.0 is required for functions used in this script
If Err.Number = 429 Then oScala.Log("RSSREADER ERROR: MSXML 4.0 not installed.")

' MSXML object to parse the XML file
Dim parseXML
Set parseXML = CreateObject("Msxml2.DOMDocument.4.0") 
parseXML.async = False

' WinHTTP object for downloading other files (eg: images)
Err.Clear
Dim WinHttpReq
Set WinHttpReq = CreateObject("WinHttp.WinHttpRequest.5.1")
' WinHttpRequest 5.1 is required for functions used in this script
If Err.Number = 429 Then oScala.Log("RSSREADER ERROR: WinHttp 5.1 not installed.")

' WinHTTP proxy constant values
HTTPREQUEST_PROXYSETTING_DEFAULT = 0
HTTPREQUEST_PROXYSETTING_PRECONFIG = 0
HTTPREQUEST_PROXYSETTING_DIRECT = 1
HTTPREQUEST_PROXYSETTING_PROXY = 2

' FileSystemObject for accessing local files 
Dim fso
Set fso = CreateObject("Scripting.FileSystemObject")
Dim TempFolder, LocalFile, LocalFileTemp
' Path for temp files
TempFolder = fso.GetSpecialFolder(2)

' Variables used in the script
Dim docElement, channelElement
Dim ItemNum, itemElement, itemElements, xmlsubItem
Dim httpFile_local
Dim authUser, authPassword
Dim net_online, first_loop
first_loop = True

' Regular Expression variables used throughout the script
Dim regEx, Match, Matches
Set regEx = New RegExp
regEx.IgnoreCase = True
regEx.Global = True

' Timeout used in case crawl gets stuck
Dim TimeoutStart, TimeoutEnd
TimeoutEnd = 120 ' seconds

' Supported file types used to properly name downloaded image files
Dim FileTypes
FileTypes = Array(".jpg",".jpeg",".gif",".png",".bmp")

' Array of month names used when parsing item dates
Dim MonthNames
MonthNames = Array("January", "February", "March", "April", "May", "June", "July", "August", "September", "October", "November", "December")

' Windows Shell object
Dim WshShell
Set WshShell = CreateObject("WScript.Shell")

' Local time zone offset for this PC
Dim  ActiveTimeBias
ActiveTimeBias = WshShell.RegRead("HKEY_LOCAL_MACHINE\System\CurrentControlSet\Control\TimeZoneInformation\ActiveTimeBias")

' Delimiter for crawl text segments
If GetScalaVersion() > "503.0.0" Then
    crawlsegment = ""
Else
    crawlsegment = vbCrLf
End If

' The options MaxItems, DownloadFrequency, MaxLengthTitle and MaxLengthDescription 
' are strings in case they are used as template data fields
' They are converted to integer equivalents below
Dim iMaxItems, iDownloadFrequency, iMaxLengthTitle, iMaxLengthDescription, iMaxLengthAuthor

' Max number of items to display (0 = display all)
if iMaxItems < 0 Then iMaxItems = 0
iMaxItems = CInt(MaxItems)

' Download frequency is the number of minutes old file must be to be downloaded again 
' (0 = use default value, -1 = downloading disabled)
' Disabling downloads is useful for situations where the xml file is copied to 
' the player locally or sent via Content Manager
' If downloading is disabled, the script will look for the xml file in Content:
' folder if it is not found in the normal location
iDownloadFrequency = CInt(DownloadFrequency)
If iDownloadFrequency = 0 Then iDownloadFrequency = 60 ' default

' Max length for titles and descriptions (0 = use default values, -1 = display all)
iMaxLengthTitle = CInt(MaxLengthTitle)
If iMaxLengthTitle <= 0 Then iMaxLengthTitle = 80 ' default
iMaxLengthDescription = CInt(MaxLengthDescription)
If iMaxLengthDescription <> 0 Then iMaxLengthDescription = 500 ' 500

iMaxLengthAuthor = CInt(MaxLengthAuthor)
If iMaxLengthAuthor <= 0 Then iMaxLengthAuthor = 160 ' default

' Check if HTTP authentication has been specified in URL
' RSSFeedURL should be user@password:url
RegEx.Pattern = "^([^@]*)@([^:]*):(.*)"
Set Matches = RegEx.Execute(RSSFeedURL)
If Matches.count > 0 Then
    With Matches(0)
        authUser = .SubMatches(0)
        authPassword = .SubMatches(1)
        RSSFeedURL = .SubMatches(2)
    End With
End If

Do
    ' If a local filename is not specified, use the URL 
    If RSSFeedName = "" Then RSSFeedName = RSSFeedURL
    ' Create the local filename, make sure it is a valid Windows name
    httpFile_local = CleanFileName(RSSFeedName) & ".xml"
    LocalFile = ""
    
    ' This flag is set to False if an image download files while displaying the feed.
    ' It is reset to true when the feed loops or is restarted.
    net_online = True 
    
    ' First look in the Content folder
    ' The feed should be in a sub-folder with the same name as the file
    LocalFileTemp = oScalaFileLock.LockScalaFile("Content:\" & CleanFileName(RSSFeedName) & "\" & httpFile_local) 
    If LocalFileTemp <> "" Then
        oScalaFileLock.UnlockScalaFile()
        LocalFile = LocalFileTemp
    Else
        ' Then look in the temp folder and decide if it needs to be updated
        If Err.Number <> 0 Then Err.Clear
        LocalFileTemp = TempFolder & "\" & CleanFileName(RSSFeedName) & "\" & httpFile_local

        ' Download the xml file if Download Frequency > 0 and either the file doesn't exist yet or is older than iDownloadFrequency (minutes)
        If (iDownloadFrequency >= 0) and ((fso.FileExists(LocalFileTemp)= False) or (DateDiff("n", GetFileDate(LocalFileTemp), Now) >= iDownloadFrequency)) Then
            downloadXML.Open "GET", RSSFeedURL, False, authUser, authPassword
            If ProxyServer <> "" Then downloadXML.setProxy HTTPREQUEST_PROXYSETTING_PROXY, ProxyServer & ":" & ProxyPort
            ' These headers force proxy servers to not cache
            downloadXML.SetRequestHeader "Cache-Control", "no-cache" 
            downloadXML.SetRequestHeader "Cache-Control", "must-revalidate" 
            downloadXML.SetRequestHeader "If-Modified-Since", "Sat, 1 Jan 2000 00:00:00 GMT"  
            downloadXML.SetRequestHeader "Pragma", "no-cache"
            ' Attempt to download, check status
            Err.Clear
            downloadXML.Send
            'Save to a local file in the temp folder if the download was successful
            If Err.Number = 0 Then
                If downloadXML.Status = 200 Then
                    saveXML.Load(downloadXML.ResponseBody)
                    If GetParseError(saveXML) = "" Then
                        ' Create sub-folder for feed content, delete if already exists since we're downloading updated feed
                        Call CreateFolder(TempFolder & "\" & CleanFileName(RSSFeedName))
                        saveXML.Save(LocalFileTemp)
                    Else
                        ' Report an error if the RSS XML didn't parse correctly
                        oScala.Log("RSSREADER ERROR: Error parsing XML feed: " & RSSFeedURL)
                        oScala.Log("RSSREADER ERROR: " & GetParseError(saveXML))
                    End If
                Else
                    ' Report an error if download was unsuccessful due to server or internet
                    oScala.Log("RSSREADER ERROR: XML file download failed: " & RSSFeedURL)
                    oScala.Log("RSSREADER ERROR: " & CStr(downloadXML.Status) & " " & downloadXML.StatusText)
                End If
            Else
                ' Report an error if download was unsuccessful due to local network
                oScala.Log("RSSREADER ERROR: XML file download failed: " & RSSFeedURL)
                oScala.Log("RSSREADER ERROR: " & CStr(Err.Number) & " " & Err.Description)
            End If
        End If
        ' Use file if it exists in the temp folder
        If fso.FileExists(LocalFileTemp) = True Then LocalFile = LocalFileTemp
    End If
    
    ' Load the local XML file
    If parseXML.Load(LocalFile) Then
    ' Start at Document Element
        Set docElement = parseXML.documentElement
        ' Did we find anything?  No childNodes could mean file not found, for example
        If (parseXML.childNodes.Length > 0) Then
            ' On the first loop, wait for the Scala script to be ready for the channel
            TimeoutStart = Now
            If (first_loop) Then
                While (ChannelRequest = 0 and DateDiff("s",TimeoutStart,Now) < TimeoutEnd)
                    oScala.Sleep(200)
                Wend
                first_loop = False
            End If

            ' Process the channel
            ChannelError = ""
            ChannelErrorNum = 0
            ' Get the <channel> element
            ' Using the local-name attribute avoids having to deal with namespaces 
            Set channelElement = docElement.selectSingleNode("*[local-name()='channel']")
            ' Extract channel's <title> and optional <pubDate> elements, and <description> / <link>
            If (HideChannelTitle = "True") Then
                ChannelTitle = ""
            Else
                ChannelTitle = GetCleanedElementText(channelElement, "title")
                If ChannelTitle <> "" Then ChannelTitle = ChopText(ChannelTitle,iMaxLengthTitle)
            End If
            If (HideChannelDescription = "True") Then
                ChannelDescription = ""
            Else
                ChannelDescription = GetCleanedElementText(channelElement, "description")
                If ChannelDescription <> "" Then ChannelDescription = ChopText(ChannelDescription,iMaxLengthDescription)
            End If
            ' JEREMY BELOW THIS
            If (HideChannelAuthor = "True") Then
                ChannelAuthor = ""
            Else
                ChannelAuthor = GetCleanedElementText(channelElement, "author")
                If ChannelAuthor <> "" Then ChannelAuthor = ChopText(ChannelAuthor,iMaxLengthAuthor)
            End If
            ' JEREMY ABOVE THIS
            ChannelLink = GetElementText(channelElement, "link")
            ChannelLanguage = GetElementText(channelElement, "language")
            ChannelCopyright = GetCleanedElementText(channelElement, "copyright")
            ChannelDate = GetCleanedElementText(channelElement, "pubDate")
            If ChannelDate <> "" Then ChannelDate = ParseRSSDate(ChannelDate)
            ' Download Channel Image if available
            If HideChannelImage = "False" Then
                Set xmlsubItem = channelElement.selectSingleNode("*[local-name()='image']")
                If Not xmlsubItem is Nothing Then
                    If GetElementText(xmlsubItem, "url") <> "" Then 
                        ChannelImageTemp = DownloadFile(GetElementText(xmlsubItem, "url"), CleanFileName(RSSFeedName) & "_channelimage" & FindFromList(FileTypes, GetElementText(xmlsubItem, "url")))
                        If ChannelImageTemp <> "" Then 
                            ' Download succeeded
                            ChannelImageStatus = 1
                            ChannelImage = ChannelImageTemp
                            ItemImage = ChannelImage
                            ' Preload first item image
                            ItemImageFoo = GetItemImage(docElement.selectSingleNode("//item"),0) 
                        Else
                            ' Download failed
                            ChannelImageStatus = 0 
                        End If
                    Else
                        ' No channel image URL found
                        ChannelImageStatus = 0 
                    End If
                Else
                    ' No channel image found
                    ChannelImageStatus = 0 
                End if
            Else
                ' Image downloading is not active
                ChannelImageStatus = 0 
            End If
            ' Signal that the channel information is ready for display
            ChannelRequest = 0

            ' Get all <item>s within the channel
            Set itemElements = docElement.selectNodes("//*[local-name()='item']")
            ' Keep count of items, so we can limit to MaxItems
            ItemNum = 0

            ' Process each item
            For Each itemElement in itemElements
                ItemNum = ItemNum + 1
                ' Exit if specified MaxItems value is reached
                If (iMaxItems > 0) And (ItemNum > iMaxItems) Then Exit For
                ' Wait for next item request
                TimeoutStart = Now
                While (ItemRequest <> 1 and DateDiff("s",TimeoutStart,Now) < TimeoutEnd)
                    oScala.Sleep(200)
                Wend

                ' Extract Headline, Body, and Link
                If (HideItemTitles = "True") Then
                    ItemTitle = ""
                Else
                    ItemTitle = GetCleanedElementText(itemElement, "title")
                    If ItemTitle <> "" Then ItemTitle = ChopText(ItemTitle,iMaxLengthTitle)
                End If
                ItemLink = GetElementText(itemElement, "link")
                If (HideItemDescriptions = "True") Then
                    ItemDescription = ""
                Else
                    ItemDescription = GetCleanedElementText(itemElement, "description")
                    If ItemDescription <> "" Then ItemDescription = ChopText(ItemDescription,iMaxLengthDescription)
                End If

                ' JEREMY ADDED BELOW
                If (HideItemAuthors = "True") Then
                    ItemAuthor = ""
                Else
                    ItemAuthor = GetCleanedElementText(itemElement, "author")
                    If ItemAuthor <> "" Then ItemAuthor = ChopText(ItemAuthor,iMaxLengthAuthor)
                End If
                ' JEREMY ADDED ABOVE

                ' Get item date, convert to local time zone and format
                If HideItemDates = "True" Then
                    ItemDate = ""
                Else
                    ItemDate = GetCleanedElementText(itemElement, "pubDate")
                    If (ItemDate = "") Then
                        ItemDate = GetCleanedElementText(itemElement, "date")
                    End If
                    If ItemDate <> "" Then ItemDate = ParseRSSDate(ItemDate)
                End If
                ' Get event fields if available
                evType = GetCleanedElementText(itemElement, "type")
                evLocation = GetCleanedElementText(itemElement, "location")
                evStartdate = GetCleanedElementText(itemElement, "startdate")
                evEnddate = GetCleanedElementText(itemElement, "enddate")
                evOrganizer = GetCleanedElementText(itemElement, "organizer")

                ' Build the CrawlItem for Text Crawls
                CrawlItem = ""
                If ItemNum > 1 Then CrawlItem = Style_Gap & vbCrLf
                If ItemTitle <> "" Then CrawlItem = CrawlItem & crawlsegment & Style_title & InsertReturns(ItemTitle,1000)
                If ItemDate <> "" Then 
                    If ItemTitle <> "" Then CrawlItem = CrawlItem & "- " 
                    CrawlItem = CrawlItem & crawlsegment & Style_date & ItemDate
                End If
                If ItemDescription <> "" Then 
                    If ItemDate <> "" or ItemTitle <> "" Then CrawlItem = CrawlItem & ": "
                    CrawlItem = CrawlItem & crawlsegment & Style_description & InsertReturns(ItemDescription,1000)
                End If
                ' JEREMY ADDED BELOW
                If ItemAuthor <> "" Then 
                    If ItemDate <> "" or ItemTitle <> "" or ItemDescription <> "" Then CrawlItem = CrawlItem & ": "
                    CrawlItem = CrawlItem & crawlsegment & Style_author & InsertReturns(ItemAuthor,1000)
                End If
                ' JEREMY ADDED ABOVE
                ' Show events instead of items if events exist
                If evStartdate <> "" Then
                    CrawlItem = Style_Gap & vbCrLf & Style_title & ItemTitle & ": " & crawlsegment & Style_description & ItemDescription & Style_author & ItemAuthor & " " & evStartdate & " - " & evEnddate
                    If evLocation <> "" Then CrawlItem = CrawlItem & " in " & evLocation 
                    If evOrganzier <> "" Then CrawlItem = CrawlItem & " by " & evOrganizer 
                    If evType <> "" Then CrawlItem = CrawlItem & " (" & evType & ")"
                End If

                ' Download item image if available
                If HideItemImages = "False" Then
                    ' Download the next image so it's ready ahead of time
                    ' If ItemNum+1 <= itemElements.length Then ItemImageFoo = GetItemImage(itemElements(ItemNum), ItemNum) 
                    ' Download the current image
                    ItemImage = GetItemImage(itemElement, ItemNum-1) 
                Else 
                    ' Downloading is not active
                    ItemImageStatus = 0 
                End If

                ' Signal item image change if new image is different from previous
                ' This is because not every item has an image
                If ItemImage_previous <> ItemImage Then ItemImageUpdate = ItemImageUpdate +1
                ItemImage_previous = ItemImage
                ' Signal new item text is ready
                ItemRequest = 0
            Next
        End If
    Else
        ' Failed to load XML document 
        ChannelError = GetParseError(parseXML)
        ChannelErrorNum = parseXML.parseError.errorCode
        oScala.Log("RSSREADER ERROR: Error parsing local XML file from feed: " & RSSFeedURL)
        oScala.Log("RSSREADER ERROR: " & ChannelError)
        If (not LoopForever) Then
            TimeoutStart = Now
            If (first_loop) Then
                While (ChannelRequest = 0 and DateDiff("s",TimeoutStart,Now) < TimeoutEnd)
                    oScala.Sleep(200)
                Wend
                first_loop = False
            End If
            ChannelRequest = 0
        End If
    End If
Loop While LoopForever = "True"

' Wait for last item request
TimeoutStart = Now
While (ItemRequest <> 1 and DateDiff("s",TimeoutStart,Now) < TimeoutEnd)
    oScala.Sleep(200)
Wend

' Update last image
ItemImageUpdate = ItemImageUpdate +1
' Signal that there are no more items
ItemRequest = -1

' Clear out all objects
Set oScala = Nothing
Set oScalaFileLock = Nothing
Set downloadXML = Nothing
Set saveXML = Nothing
Set parseXML = Nothing
Set WinHttpReq = Nothing
Set fso = Nothing
Set WshShell = Nothing



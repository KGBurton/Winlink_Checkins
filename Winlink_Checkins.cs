﻿//c# code that will input start date, end date, and callSign and will select files with an extension of mime from the current folder  based on start date and end date, and will read each file to find a line labeled To: . If the rest of the line contains callSign, then write the data from the line labeled X-Source: to a text file called checkins.txt in the same folder
// Design Get the date range
// get the data source
// is it a message file? (.mime)
// is it within the date range
// is it related to a unique net checkin name (make all strings are upper case or ignore case before checking)
// is it a bounced message? record for review, do not count
// is it a forwarded message? be sure to get the correc callsign
// is the callsign in the MSG field? if not try the from field, if not record and don't count
//      REGEX pattern to find the FCC callsign is \b[A-Z]{1,2}\d[A-Z]{1,3}\b for us only
//      \b\d{0,2}[A-Z]{1,2}\d{1,2}[A-Z]{1,6}\b for almostall international and US
//      see https://regex101.com/r/gS6qG8/1 for a regex tester/editor. This is string to test in 
//      the editor: Ken, AE0MW, this, is, a test. 4F1PUZ 4G1CCH AX3GAMES 4D71 4X130RISHON 9N38 AX3GAMES BV100 DA2MORSE  DB50FIRAC DL50FRANCE FBC5AGB FBC5CWU FBC5LMJ FBC5NOD FBC5YJ FBC6HQP GB50RSARS HA80MRASZ  HB9STEVE HG5FIRAC HG80MRASZ II050SCOUT IP1METEO J42004A J42004Q LM1814 LM2T70Y LM9L40Y LM9L40Y/P OEM2BZL OEM3SGU OEM3SGU/3 OEM6CLD OEM8CIQ OM2011GOOOLY ON1000NOTGER ON70REDSTAR PA09SHAPE PA65VERON PA90CORUS PG50RNARS PG540BUFFALO S55CERKNO TM380 TX9 TYA11 U5ARTEK/A V6T1 VI2AJ2010 VI2FG30 VI4WIP50 VU3DJQF1 VX31763 WD4 XUF2B YI9B4E YO1000LEANY ZL4RUGBY ZS9MADIBA
// is it in the current roster? if not record with new checkins, save, count.
//     Requires that roster.txt exist in the application folder. 
// does it have location data? REGEX pattern for latitude, longitude
// old this was deficient - needed to check and limit 90N/S and 180E/W, @"([-+]?[0-9]*\.?[0-9]+)\s*[°]?\s*([NS]),?\s*([-+]?[0-9]*\.?[0-9]+)\s*[°]?\s*([EW])"
// new @"([-+]?([0-8]?\d(\.\d+)?|90(\.0+)?))\s*[°]?\s*([NS]),?\s*([-+]?((1[0-7]\d|\d{1,2})(\.\d+)?|180(\.0+)?))\s*[°]?\s*([EW])"
//       ([-+]?([0-8]?\d(\.\d+)?|90(\.0+)?))\s*[°]?\s*([NS]),?\s*([-+]?((1[0-7]\d|\d{1,2})(\.\d+)?|180(\.0+)?))\s*[°]?\s*([EW])
//  my location for testing: 43.845831N, 111.745744W
//  REGEX pattern for Maidenhead grids: \b([A-R]{2}\d{2}[A-X]{0,2})\b, test DN43du
// is it a duplicate? if yes, don't save or count (spreadsheet can handle duplicates)
// what template was used? necessary to get the start and end positions correct
// save document info, writeLines, and counts to checkins.txt;
// write callsign and message to checkins.csv file
// mapping resource https://github.com/RTykulsker/WinlinkMessageMapper

using System;
using System.Collections.Generic;
using System.ComponentModel.Design;
using System.IO;
using System.Linq;
using System.Reflection.Metadata;
using System.Threading.Tasks;
using System.Runtime.CompilerServices;
using System.Text;
using System.Text.Json.Serialization;
using System.Text.RegularExpressions;
using System.Globalization;
using System.Runtime.Intrinsics.X86;
using System.Xml.Linq;
using static System.Net.Mime.MediaTypeNames;
using System.Formats.Asn1;
using System.ComponentModel;
using System.Drawing;


class Winlink_Checkins
{
    public static void Main (string [] args)
    {
        // Get the start date and end date from the user.
        DateTime startDate = DateTime.Today;
        DateTime endDate = DateTime.Today;
        string utcDate = DateTime.UtcNow.ToString ("yyyy/MM/dd HH:mm:ssZ");
        //DateTime date;
        bool isValid = false;
        string input;

        Console.WriteLine ("Enter the start date - must be within two weeks of today (yyyymmdd): ");
        startDate = GetValidDate ();
        //startDate = startDate.ToUniversalTime();

        while (!isValid)
        {
            Console.WriteLine ("Enter the end date - must be within two weeks of today (yyyymmdd): ");
            endDate = GetValidDate ();
            int startDateCompare = DateTime.Compare (startDate, endDate);
            if (startDateCompare >= 0)
            {
                Console.WriteLine ("The start date must be earlier than the end date. Please try again.");
            }
            else
            { isValid = true; }

        }
        endDate = endDate.AddDays (1);

        // Get the unique net identifier to screen only relevant messages from the folder
        // Console.WriteLine("Enter the unique net name for which the checkins are sent:");
        // string netName = Console.ReadLine();
        // Get the native call sign from the user to find the messages folder.
        Console.WriteLine ("Enter YOUR call sign to find the messages folder. \n     If you leave it blank, the program will assume that it is already in the messages folder.");
        string yourCallSign = Console.ReadLine ();

        // Get the data folder - either the global messages folder (default) or the current
        // operator's messages folder, assuming the default RMS installation location.
        string currentFolder = "";
        string applicationFolder = Directory.GetCurrentDirectory ();
        string netName = "";
        if (yourCallSign != "")
        {
            currentFolder = "C:\\RMS Express\\" + yourCallSign + "\\Messages";
        }
        else
        {
            currentFolder = Directory.GetCurrentDirectory ();
        }

        // Look for roster.txt in the folder. If it exists, get the first (and only)
        // row for comparison down below
        string rosterFile = applicationFolder + "\\roster.txt";
        string xmlFile = applicationFolder + "\\Winlink_Import.xml"; // separate file for defective messages
        string xmlPerfFile = applicationFolder + "\\Winlink_Import_Perfect.xml"; // separate file for perfect messages
        // string commentFile = applicationFolder+"\\GLAWN_Additional_Comments.txt";
        // writeString variables to go in the output files
        StringBuilder netCheckinString = new StringBuilder ();
        StringBuilder netAckString2 = new StringBuilder ();
        StringBuilder bouncedString = new StringBuilder ();
        StringBuilder duplicates = new StringBuilder ();
        StringBuilder newCheckIns = new StringBuilder ();
        StringBuilder csvString = new StringBuilder ();
        csvString.AppendLine ("Current GLAWN Checkins, posted: " + utcDate);
        StringBuilder mapString = new StringBuilder ();
        mapString.Append ("CallSign,Latitude,Longitude,Band,Mode\r\n");
        StringBuilder badBandString = new StringBuilder ();
        StringBuilder badModeString = new StringBuilder ();
        StringBuilder skippedString = new StringBuilder ();
        StringBuilder removalString = new StringBuilder ();
        StringBuilder addonString = new StringBuilder ();
        StringBuilder noGPSString = new StringBuilder ();
        StringBuilder noScoreString = new StringBuilder ();
        StringBuilder typoString = new StringBuilder ();

        // string callSignPattern = @"\b\d{0,2}[A-Z]{1,2}\d{1,2}[A-Z]{1,6}\b";
        string testString = "";
        string rosterString = "";
        string bandStr = "";
        string modeStr = "";
        // string noGPSStr = "";
        string checkIn = "";
        string msgField = "";
        string msgFieldNumbered = "";
        string latitudeStr = "";
        string saveLatitudeStr = "";
        string longitudeStr = "";
        string saveLongitudeStr = "";
        string xmlXsource = "KB7WHO"; // this is mine and is the default in case the roster.txt file doesn't have one
        // string delimiter = "";
        string pointsOff = "";
        string reminderTxt = "";
        string reminderTxt2 = "";
        string callSignTypo = "";
        string checkinName = "";
        string checkinCountry = "";
        string checkinCountryLong = "";
        string checkinState = "";
        string checkinCounty = "";
        string checkinCity = "";
        string maidenheadGrid = "";
        string locType = "";
        string xSource = "";
        string fromTxt = "";


        addonString.AppendLine ("\r\nComments from the Current Checkins Posted\t" + utcDate + "\r\n-------------------------------");
        noGPSString.AppendLine ("\r\n++++++++\r\nThese had neither GPS data nor Maidenhead Grids\r\n-------------------------");
        noScoreString.AppendLine ("\r\n++++++++\r\nThese chose not to be scored:");
        Random rnd = new Random ();

        int startPosition = 0;
        int endPosition = 0;
        int quotedPrintable = 0;
        int commentPos = 0;
        int len = 0;
        int msgTotal = 0;
        int skipped = 0;
        int oldSkipped = 0;
        int ct = 0;
        int dupCt = 0;
        var dupeRemoveCt = 0;
        int newCt = 0;
        int outOfRangeCt = 0;
        int removalCt = 0;
        int ackCt = 0;
        int localWeatherCt = 0;
        int severeWeatherCt = 0;
        int incidentStatusCt = 0;
        int icsCt = 0;
        int winlinkCkinCt = 0;
        int damAssessCt = 0;
        int fieldSitCt = 0;
        int quickMCt = 0;
        int qwmCt = 0;
        int miCt = 0;
        int dyfiCt = 0;
        int rriCt = 0;
        int junk = 0;
        int mapCt = 0;
        int bandCt = 0;
        int modeCt = 0;
        int aprsCt = 0;
        int js8ct = 0;
        int meshCt = 0;
        int noGPSCt = 0;
        int noGPSFlag = 0;
        int badBandCt = 0;
        int badModeCt = 0;
        int perfectScoreCt = 0;
        int dupeFlag = 0;
        int score = 10;
        // int rowsToSkip = 0;
        int noScoreCt = 0;
        int js8call = 0;
        int PosRepCt = 0;
        int copyPR = -1;
        int ICS201Ct = 0;
        int exerciseCompleteCt = 0;
        int radioGram = 0;
        int radioGramCt = 0;

        double latitude = 0;
        double longitude = 0;
        bool isPerfect = true;
        

        TextInfo textInfo = new CultureInfo ("en-US", false).TextInfo;
        // Create root XML document
        XDocument xmlDoc = new XDocument (new XElement ("WinlinkMessages"));
        XDocument xmlPerfDoc = new XDocument (new XElement ("WinlinkMessages"));

        XElement messageElement = new XElement
            ("export_parameters",
                new XElement ("xml_file_version", "1.0"),
                new XElement ("winlink_express_version", "1.7.17.0"),
                // for testing
                // new XElement("callsign", "KB7WHO")
                new XElement ("callsign", "GLAWN")
            );
        xmlDoc.Root.Add (messageElement);
        xmlPerfDoc.Root.Add (messageElement);

        messageElement = new XElement ("message_list", "");
        xmlDoc.Root.Add (messageElement);
        xmlPerfDoc.Root.Add (messageElement);

        if (File.Exists (rosterFile))
        {
            rosterString = File.ReadAllText (rosterFile);
            rosterString = rosterString.ToUpper ();
            //debug Console.WriteLine("rosterFile contents: "+rosterString);
            // get the net name from the roster.txt file
            startPosition = rosterString.IndexOf ("NETNAME=");
            if (startPosition > -1) { startPosition += 8; }
            endPosition = rosterString.IndexOf ("//", startPosition);
            len = endPosition - startPosition;
            if (len > 0)
            { netName = rosterString.Substring (startPosition, len); }
            else { netName = "GLAWN"; }

            // get the x-source name from the roster.txt file to be used as the netName variable in the xml file
            startPosition = rosterString.IndexOf ("CALLSIGN=");
            if (startPosition > -1) { startPosition += 9; }
            endPosition = rosterString.IndexOf ("//", startPosition);
            len = endPosition - startPosition;
            if (len > 0)
            { xmlXsource = rosterString.Substring (startPosition, len); }
            else
            {
                Console.WriteLine ("callSign missing from the roster.txt file. X-SOURCE in the xml file will be wrong.");
            }

            // get the checkin roster from the roster.txt file
            startPosition = endPosition + 2;
            // endPosition = rosterString.IndexOf("\r\n", startPosition);
            len = rosterString.Length - startPosition;
            if (len > 0) { rosterString = rosterString.Substring (startPosition, len); }

        }
        else
        {
            Console.WriteLine (currentFolder + "\\" + rosterFile + " \n was not found!, all checkins will appear to be new.\nEnter the name of the net you are checking in:");
            isValid = false;
            while (!isValid)
            {
                input = Console.ReadLine ();
                if (input == null) { isValid = true; break; }
                Console.WriteLine ("The net name is reguired to create a new roaster.txt file:");
                netName = input;
            }

            isValid = false;
            Console.WriteLine ("Enter the callsign to use as the xSource for the personalized messages:");
            while (!isValid)
            {
                input = Console.ReadLine ();
                if (input == null) { isValid = true; break; }
                Console.WriteLine ("A callsign is reguired to create a new roaster.txt file:");
                xmlXsource = input;
            }
        }


        // Select files with an extension of mime from the current folder.
        var files = Directory.GetFiles (currentFolder, "*.mime")
            .Where (file =>
            {
                DateTime fileDate = File.GetLastWriteTime (file);
                // debug Console.Write(fileDate+"\n");
                return fileDate >= startDate && fileDate <= endDate.AddDays (1);
            });
        Directory.CreateDirectory (currentFolder); // Ensures the folder exists

        Console.Write ("\nMessages to process=" + files.Count () + " from folder " + currentFolder + "\n\n");

        // Create a text file called checkins.txt in the data folder and process the list of files.
        using (StreamWriter logWrite = new (Path.Combine (currentFolder, "checkins.txt")))
        // Create a text file called checkins.csv in the data folder and process the list of files.
        using (StreamWriter csvWrite = new (Path.Combine (currentFolder, "checkins.csv")))
        // Create a csv text file called mapfile.csv in the data folder to use as date for google maps
        using (StreamWriter mapWrite = new (Path.Combine (currentFolder, "mapfile.csv")))
        using (StreamWriter commentWrite = new (Path.Combine (currentFolder, "GLAWN Additional Comments.txt")))
        {
            // Read each file selected to find a line labeled To: and if the rest of the line contains netName, write the data from the line labeled X-Source: to the text file.
            foreach (string file in files)
            {
                using (StreamReader reader = new StreamReader (file))
                {
                    msgTotal++;
                    latitudeStr = "";
                    saveLatitudeStr = "";
                    latitude = 0;
                    longitudeStr = "";
                    saveLongitudeStr = "";
                    longitude = 0;
                    //debug Console.Write("File "+file+"\n");
                    string fileText = reader.ReadToEnd ();

                    fileText = fileText.ToUpper ()
                        .Replace ("NO SCORE", "NOSCORE")
                        .Replace ("=0A", "\r\n")
                        .Replace ("=\r\n", "") // remove line wraps
                        .Replace ("=20", " ");

                    // get needed header info
                    startPosition = fileText.IndexOf ("DATE: ");
                    if (startPosition > -1) { startPosition += 11; }
                    len = 20;
                    string sentDate = fileText.Substring (startPosition, len);
                    DateTime sentDateUni = DateTime.Parse (sentDate);

                    startPosition = fileText.IndexOf ("MESSAGE-ID: ");
                    if (startPosition > -1) { startPosition += 12; }
                    endPosition = fileText.IndexOf ("\r\n", startPosition);
                    len = endPosition - startPosition;
                    string messageID = fileText.Substring (startPosition, len);
                    reminderTxt = ""; // reset the text that goes into the xml personalized message
                    // get From:
                    startPosition = fileText.IndexOf ("FROM:");
                    if (startPosition > -1) { startPosition += 6; }
                    endPosition = fileText.IndexOf ("\r\n", startPosition);
                    len = endPosition - startPosition;
                    fromTxt = fileText.Substring (startPosition, len);
                    {
                        fromTxt = fromTxt.Replace (',', ' ');
                        // Create a Regex object with the pattern
                        fromTxt = isValidCallsign (fromTxt);
                    }

                    // find the end of the header section
                    var endHeader = fileText.IndexOf ("CONTENT-TRANSFER-ENCODING:");
                    quotedPrintable = fileText.IndexOf ("QUOTED-PRINTABLE");
                    if (quotedPrintable > -1) quotedPrintable += 20;
                    commentPos = fileText.IndexOf ("COMMENT:", quotedPrintable);
                    if (commentPos > -1) commentPos += 9;

                    // does the sender want to skip the scoring of the message
                    var noScore = fileText.IndexOf ("NOSCORE");


                    // deterimine if it was forwarded to know to look below the first header info
                    var forwarded = fileText.IndexOf ("WAS FORWARDED BY");

                    // was it APRSmail?
                    var APRS = fileText.IndexOf ("APRSEMAIL2");
                    if (APRS < 0) APRS = fileText.IndexOf ("APRS.EARTH");
                    if (APRS < 0) APRS = fileText.IndexOf ("APRS.FI");
                    // was it JS8CALL
                    js8call = fileText.IndexOf ("JS8CALL");
                    if (js8call > -1) js8ct++;

                    // check for acknowledgement message and discard later                  
                    int ack = fileText.IndexOf ("ACKNOWLEDGEMENT");

                    // check for ICS 213 msg
                    var ics = fileText.IndexOf ("TEMPLATE VERSION: ICS 213");

                    // check for winlink checkin message
                    var winlinkCkin = fileText.IndexOf ("MAP FILE NAME: WINLINK CHECK", endHeader);
                    // some people include WINLINK CHECK-IN in the subject which confuses the program
                    // into thinking this is a winlink checkin FORM!! Catch it ...
                    if (winlinkCkin < 0) winlinkCkin = fileText.IndexOf ("WINLINK CHECK-IN 5.0", endHeader);
                    if (winlinkCkin < 0) winlinkCkin = fileText.IndexOf ("WINLINK CHECK-IN \r\n0. HEADER", endHeader);
                    if (winlinkCkin < 0) winlinkCkin = fileText.IndexOf ("WINLINK CHECK IN 2.", endHeader);

                    // is this a Position Report? For position reports, the valid message will be a fowarded 
                    // response from SERVICE with Latitude: Longitude: and Comment:
                    // if it is fowarded from some other address or it is missing the comment, it isn't valid
                    // but still useful to checkin
                    copyPR = -1;
                    bool PosReport = false;
                    if (fileText.Contains ("COMMENT:") && (fileText.Contains ("POSITION REPORT ACKNOWLEDGEMENT") || fileText.Contains ("DUPLICATE POSITION REPORT")) && fileText.Contains ("MESSAGE FROM SERVICE WAS FORWARDED"))
                    {
                        PosReport = true;
                        PosRepCt++;
                    }
                    // check for Position Report that was acknowledged, if yes, clear the false ack flag
                    // if this is a fowarded position report, it will be a false message acknowledgement
                    // and should not be discarded.
                    int startPR = fileText.IndexOf ("POSITION REPORT ACKNOWLEDGEMENT");
                    // if POSITION REPORT ACKNOWLEDGEMENT is not found, look for duplicate response from service that was forwarded
                    // this is still valid and means that they sent the message to QTH more than once
                    if (startPR == -1)
                    {
                        startPR = fileText.IndexOf ("DUPLICATE POSITION REPORT", endHeader);
                    }

                    if (startPR > -1) ack = -1;
                    if (fileText.Contains ("LATITUDE: ") && fileText.Contains ("LONGITUDE: ") && ics == -1 && winlinkCkin == -1)
                    {
                        // convert postion report degrees to decimal
                        int startLat = fileText.IndexOf ("LATITUDE: ");
                        if (startLat > -1) startLat += 10;
                        int endLat = fileText.IndexOf ("\r\n", startLat);
                        len = endLat - startLat;
                        if (len > 0)
                        {
                            latitudeStr = fileText.Substring (startLat, len);
                        }
                        // regex pattern for gps coordinates in degrees from position report nn-nn.nnA
                        Regex regexDegrees = new Regex (@"\d{1,3}\-\d{1,2}\.\d+[a-zA-Z0-9]*");
                        Match matchDegrees = regexDegrees.Match (latitudeStr);
                        if (matchDegrees.Success)
                        {
                            saveLatitudeStr = latitudeStr;
                            if (latitudeStr.Length <= 10 && latitudeStr != "")
                            {
                                // Console.WriteLine (messageID);
                                latitudeStr = ConvertDegreeAngleToDecimal (latitudeStr);
                            }
                            // else latitudeStr = "";
                        }
                        else
                        {
                            // regex pattern for gps coordinates in decimal from position report nn.nnnn
                            Regex regexDecimal = new Regex (@"\d{1,3}\.\d+");
                            Match matchDecimal = regexDecimal.Match (latitudeStr);
                            if (matchDecimal.Success)
                            {
                                saveLatitudeStr = latitudeStr;
                                //if (latitudeStr.Length <= 10 && latitudeStr != "")
                                //{
                                // Console.WriteLine (messageID);
                                // latitudeStr = ConvertDegreeAngleToDecimal (latitudeStr);
                                //}
                            }
                        }
                        if (latitudeStr != "") { latitude = Common.ConvertToDouble (latitudeStr); }

                        int startLong = fileText.IndexOf ("LONGITUDE: ");
                        if (startLong > -1) startLong += 11;
                        int endLong = fileText.IndexOf ("\r\n", startLong);
                        len = endLong - startLong;
                        if (len > 0)
                        {
                            longitudeStr = fileText.Substring (startLong, len);
                        }
                        matchDegrees = regexDegrees.Match (longitudeStr);
                        if (matchDegrees.Success)
                        {
                            saveLongitudeStr = longitudeStr;
                            if (longitudeStr.Length <= 12 && longitudeStr != "")
                            {
                                longitudeStr = ConvertDegreeAngleToDecimal (longitudeStr);
                            }
                        }
                        else
                        {
                            Regex regexDecimal = new Regex (@"\d{1,3}\.\d+");
                            Match matchDecimal = regexDecimal.Match (longitudeStr);
                            if (matchDecimal.Success)
                            {
                                saveLongitudeStr = longitudeStr;
                                //if (longitudeStr.Length <= 10 && longitudeStr != "")
                                //{
                                //  longitudeStr = ConvertDegreeAngleToDecimal (longitudeStr);
                                // Console.WriteLine (messageID);
                                //}
                            }
                            // latitudeStr = "";
                        }
                        if (longitudeStr != "") { longitude = Common.ConvertToDouble (longitudeStr); }
                    }

                    if (!PosReport)
                    {
                        copyPR = fileText.IndexOf ("POSITION REPORT ACKNOWLEDGEMENT");
                        if (copyPR > -1)
                        {
                            reminderTxt += "Invalid position report for this exercise. No Comment: field found.";
                            startPR = 0;
                        }
                        else
                        {
                            // the QTH message was copied to GLAWN instead of forwarding the responses
                            // check for valid position report comment
                            copyPR = fileText.IndexOf ("POSITION REPORT");
                        }
                        if (copyPR > -1)
                        {
                            PosRepCt++;
                            startPosition = fileText.IndexOf ("COMMENT:", copyPR);
                            if (startPosition == -1) copyPR = 0;
                        }
                        // check for position report that is not valid
                        // this caused problems for ICS and Winlink Checkin messages
                        // copyPR = 0;
                    }

                    bool QTH = fileText.Contains ("TO: QTH") || fileText.Contains ("CC: QTH");

                    // check for removal message and discard                  
                    var removal = fileText.IndexOf ("REMOVE ME");

                    // look to see if it was a bounced message
                    var bounced = fileText.IndexOf ("UNDELIVERABLE");

                    // check for local weather report
                    var localWeather = fileText.IndexOf ("CURRENT LOCAL WEATHER CONDITIONS");

                    // check for severe weather
                    var severeWeather = fileText.IndexOf ("SEVERE WX REPORT");

                    // check incident status report
                    var incidentStatus = fileText.IndexOf ("INCIDENT STATUS");

                    // check for odd checkin message - don't let it scan through to a binary attachment!
                    var lenBPQ = fileText.Length - 10;
                    if (lenBPQ > 800) { lenBPQ = 800; }
                    var BPQ = fileText.IndexOf ("BPQ", 1, lenBPQ);
                    var BPQAPRS = fileText.IndexOf ("BPQAPRS", 1, lenBPQ);// if this is a mode type, ignore the problem
                    if (BPQ == BPQAPRS) BPQ = -1;

                    // check for damage assessment report
                    var damAssess = fileText.IndexOf ("SURVEY REPORT - CATEGORIES");

                    // check for field situation report
                    var fieldSit = fileText.IndexOf ("EMERGENT/LIFE SAFETY");

                    // check for Quick Health & Welfare report, doesn't exist anymore? 10/2024
                    var quickM = fileText.IndexOf ("\r\nFROM ", endHeader);

                    // check for RRI Quick Welfare Message
                    var qwm = fileText.IndexOf ("TEMPLATE VERSION: QUICK WELFARE MESSAGE");

                    // check for RRI Welfare Radiogram
                    var rriWR = fileText.IndexOf ("TEMPLATE VERSION: RRI WELFARE RADIOGRAM");

                    // check for Did You Feel It report
                    var dyfi = fileText.IndexOf ("DYFI WINLINK");

                    // check for Medical Incident Report
                    var mi = fileText.IndexOf ("INITIAL PATIENT ASSESSMENT");

                    // check for ICS-201
                    var ICS201 = fileText.IndexOf ("ICS 201 INCIDENT BRIEFING");

                    // check for Radiogram
                    radioGram = fileText.IndexOf ("\r\nAR \r\n");

                    // screen dates to eliminate file dates that are different from the sent date and fall outside the net span
                    int startDateCompare = DateTime.Compare (sentDateUni, startDate);
                    int endDateCompare = DateTime.Compare (sentDateUni, endDate);

                    // discard acknowledgements
                    if (ack > 0)
                    {
                        skipped++;

                        ackCt++;
                        // Console.Write (messageID + " Acknowledgement discarded\r\n");
                        oldSkipped = skipped;
                        junk = 0; //debug Console.Write(file+" is an acknowedgement, skipping.");
                        skippedString.Append ("\tAcknowledgement discarded: " + messageID + "\r\n");
                    }
                    else if (startDateCompare < 0 || endDateCompare > 0)
                    {
                        skipped++;
                        outOfRangeCt++;
                        oldSkipped = skipped;
                        Console.Write (messageID + " sendDate fell outside the start\\end dates\r\n");
                        skippedString.Append ("\tOut of date range: " + messageID + "\r\n");
                    }


                    else if (removal > 0)
                    {

                        removalString.AppendLine (checkIn + " in " + messageID + " was a removal request.");
                        removalCt++;
                        junk = 0;  // debug Console.Write("Removal Request: "+file+", skipping.");
                    }
                    else if (bounced > 0)
                    {
                        startPosition = bounced;
                        endPosition = fileText.IndexOf ("\r\n", startPosition);
                        len = endPosition - startPosition;
                        checkIn = fileText.Substring (startPosition, len);
                        {
                            checkIn = checkIn.Replace (',', ' ');
                            // Create a Regex object with the pattern
                            checkIn = isValidCallsign (checkIn);
                        }
                        bouncedString.Append ("Message to: " + checkIn + " was not deliverable.\r\n");
                        skipped++;
                    }
                    else
                    {
                        // determine if the message has something in the subject to do with GLAWN
                        // extended to include the TO: field in case they didn't put the netName in the subject
                        startPosition = fileText.IndexOf ("SUBJECT:");
                        if (startPosition > -1) { startPosition += 9; }
                        endPosition = fileText.IndexOf ("MESSAGE-ID", startPosition);
                        len = endPosition - startPosition;
                        string subjText = fileText.Substring (startPosition, len);

                        // if (subjText.Contains (netName))
                        if (fileText.Contains (netName))
                        {
                            score = 10;
                            isPerfect = true;
                            pointsOff = "";
                            // get x-Source if available XXXX
                            var xSrc = fileText.IndexOf ("X-SOURCE: ");
                            if (xSrc > -1)
                            {
                                startPosition = xSrc + 10;
                                endPosition = fileText.IndexOf ("\r\n", startPosition);
                                len = endPosition - startPosition;
                                if (len > 0) { xSource = fileText.Substring (startPosition, len); }
                            }
                            else xSource = fromTxt;

                            // skip APRS header 
                            if (APRS > -1)
                            {
                                startPosition = fileText.IndexOf ("FROM:", endHeader);
                                if (startPosition > -1)
                                {
                                    startPosition = fileText.IndexOf ("\r\n", startPosition);
                                    if (startPosition > -1) { startPosition += 2; }
                                    endPosition = fileText.IndexOf ("DO NOT REPLY", startPosition) - 1;
                                }
                                aprsCt++;
                            }
                            // skip JS8Call header 



                            // adjust for ICS 213
                            else if (ics > -1)
                            {
                                // check first is it a reply (checkin will be in a different location

                                startPosition = fileText.IndexOf ("9. REPLY:");
                                if (startPosition > -1)
                                {
                                    startPosition += 11;
                                    endPosition = fileText.IndexOf ("REPLIED BY:", startPosition) - 3;
                                }
                                else
                                {
                                    startPosition = fileText.IndexOf ("MESSAGE:");
                                    if (startPosition > -1) { startPosition += 12; }
                                    endPosition = fileText.IndexOf ("APPROVED BY:", startPosition) - 3;
                                }
                            }
                            // adjust for winlink checkin
                            else if (winlinkCkin > -1)
                            {
                                // the winlink check-in form changed format between 5.0.10 and 5.0.5 so check for that
                                var winlinkCkinOffset = fileText.IndexOf ("WINLINK CHECK-IN 5.0.5");
                                // if (winlinkCkinOffset > -1) { winlinkCkinOffset = 9; } else { winlinkCkinOffset = 13; }
                                // startPosition = fileText.IndexOf("COMMENTS:")+ winlinkCkinOffset;
                                startPosition = fileText.IndexOf ("COMMENTS:");
                                if (startPosition > -1) { startPosition += 9; }
                                endPosition = fileText.IndexOf ("----------", startPosition) - 1;
                            }

                            // adjust for odd message that insert an R: line at the top
                            else if (BPQ > -1)
                            {
                                startPosition = fileText.IndexOf ("BPQ", 1, lenBPQ);
                                if (startPosition > -1) { startPosition += 12; }
                                endPosition = fileText.IndexOf ("--BOUNDARY", startPosition) - 2;
                            }
                            else if (localWeather > -1)
                            {
                                startPosition = fileText.IndexOf ("NOTES:");
                                if (startPosition > -1) { startPosition += 9; }
                                endPosition = fileText.IndexOf ("----------", startPosition) - 1;
                            }

                            else if (severeWeather > -1)
                            {
                                startPosition = fileText.IndexOf ("COMMENTS:");
                                if (startPosition > -1) { startPosition += 10; }
                                endPosition = fileText.IndexOf ("----------", startPosition) - 1;
                            }

                            else if (incidentStatus > -1)
                            {
                                startPosition = fileText.IndexOf ("REPORT SUBMITTED BY:");
                                if (startPosition > -1) { startPosition += 20; }
                                endPosition = fileText.IndexOf ("----------", startPosition) - 1;
                            }

                            else if (damAssess > -1)
                            {
                                startPosition = fileText.IndexOf ("COMMENTS:");
                                if (startPosition > -1) { startPosition += 21; }
                                endPosition = fileText.IndexOf ("----------", startPosition) - 1;
                            }

                            else if (fieldSit > -1)
                            {
                                startPosition = fileText.IndexOf ("COMMENTS:");
                                if (startPosition > -1) { startPosition += 11; }
                                endPosition = fileText.IndexOf ("\r\n", startPosition);
                            }

                            else if (dyfi > -1)
                            {
                                startPosition = fileText.IndexOf ("COMMENTS");
                                if (startPosition > -1) { startPosition += 11; }
                                endPosition = fileText.IndexOf ("\r\n", startPosition) - 1;
                            }

                            else if (rriWR > -1)
                            {
                                startPosition = fileText.IndexOf ("BT\r\n");
                                if (startPosition > -1) { startPosition += 3; }
                                endPosition = fileText.IndexOf ("------", startPosition) - 1;
                            }

                            else if (quickM > -1)
                            {
                                startPosition = quickM;
                                startPosition = fileText.IndexOf ("SENT ON ", startPosition);
                                if (startPosition > -1)
                                {
                                    startPosition = fileText.IndexOf ("\r\n", startPosition) + 2;
                                    endPosition = fileText.IndexOf ("--BOUNDARY", startPosition) - 2;
                                }
                                else startPosition = 0; endPosition = 0;
                            }

                            else if (qwm > -1)
                            {
                                startPosition = fileText.IndexOf ("\r\n", endHeader) + 4;
                                endPosition = fileText.IndexOf ("IT WAS SENT FROM:");
                                // some messages come with the checkin data in the wrong spot
                                // so if there is nothing found so far, try the alternative
                                len = endPosition - startPosition;
                                if (len <= 0)
                                {
                                    startPosition = fileText.IndexOf ("IT WAS SENT FROM:");
                                    startPosition = fileText.IndexOf ("\r\n", startPosition) + 2;
                                    startPosition = fileText.IndexOf ("\r\n", startPosition) + 2;
                                    endPosition = fileText.IndexOf ("THIS IS A ONE WAY", startPosition) - 2;
                                }
                            }
                            else if (mi > -1)
                            {
                                startPosition = fileText.IndexOf ("ADDITIONAL INFORMATION");
                                startPosition = fileText.IndexOf ("\r\n", startPosition);
                                endPosition = fileText.IndexOf ("----", startPosition) - 1;
                            }
                            else if (ICS201 > -1)
                            {
                                startPosition = fileText.IndexOf ("PROTECT RESPONDERS FROM THOSE HAZARDS.");
                                startPosition = fileText.IndexOf ("\r\n", startPosition);
                                endPosition = fileText.IndexOf ("6. PREPARED BY:", startPosition) - 1;
                            }
                            else if (PosReport || copyPR > -1)
                            {
                                //if (copyPR == 0) reminderTxt += "No checkin information/Comment: tag in the message";
                                if (commentPos == -1) reminderTxt += "Not a valid Position Report for this exercise. No Comment: tag in the message.\r\n";

                                // change this to true to skip the point deduction
                                bool skip = true;
                                // bool skip = true;
                                if ((startPR <= 0 || copyPR >= 0) && !skip)
                                {
                                    pointsOff = "\tminus 1 for invalid Position Report\r\n";
                                    isPerfect = false;
                                    score -= 1;
                                }
                                else
                                {
                                    if (commentPos > -1) startPosition = commentPos;
                                    endPosition = fileText.IndexOf ("\r\n", startPosition);
                                    if (startPosition == endPosition) commentPos = -1;
                                }

                                if (copyPR > -1)
                                {
                                    reminderTxt += "You appear to have copied the QTH message to GLAWN instead of fowarding the response from Service.\r\n";

                                    if (copyPR >= 0 && commentPos > -1)
                                    {
                                        startPosition = commentPos;
                                        endPosition = fileText.IndexOf ("\r\n", startPosition);
                                    }
                                    else
                                    {
                                        // find alternative msgField location
                                        startPosition = quotedPrintable;
                                        if (startPosition > -1) endPosition = fileText.IndexOf ("\r\n", startPosition);
                                    }
                                    if (startPosition == endPosition) commentPos = -1;
                                }
                                else
                                {
                                    if (startPR <= 0)
                                    {
                                        reminderTxt += "Not a valid Position Report Acknowledgement.\r\n";
                                        if (QTH) reminderTxt += "This appears to be a copy of the message to QTH instead of a forward of the response from Service\r\n";
                                        if (!skip)
                                        {
                                            pointsOff = "\tminus 1 for invalid Position Report\r\n";
                                            isPerfect = false;
                                            score -= 1;
                                        }
                                    }
                                    if (commentPos > -1)
                                    {
                                        startPosition = commentPos;
                                        endPosition = fileText.IndexOf ("\r\n", startPosition);
                                    }
                                    len = endPosition - startPosition;
                                    if (len <= 0)
                                    {
                                        startPosition = quotedPrintable;
                                        if (startPosition > -1) endPosition = fileText.IndexOf ("\r\n", startPosition);
                                    }
                                }
                            }
                            else if (radioGram > -1)
                            {
                                startPosition = fileText.IndexOf ("/", quotedPrintable);
                                startPosition = fileText.LastIndexOf ("\r\n", startPosition)+2;
                                // endPosition = fileText.IndexOf ("\r\nAR ", startPosition);
                                endPosition = fileText.LastIndexOf ("/");
                                endPosition = fileText.IndexOf ("\r\n", endPosition);
                                radioGramCt++;
                            }
                            else
                            {
                                // end of the header information as the start of the msg field
                                if (forwarded < 0)
                                {
                                    startPosition = quotedPrintable;
                                }
                                else
                                {
                                    // startPosition = forwarded+59;
                                    startPosition = fileText.IndexOf ("SUBJECT:", forwarded);
                                    if (startPosition > -1) { startPosition += 9; }
                                    startPosition = fileText.IndexOf ("\r\n", startPosition);
                                    if (startPosition > -1) { startPosition += 4; }
                                    // look for a second Subject tag
                                    startPosition = fileText.IndexOf ("SUBJECT:", forwarded);
                                    if (startPosition > -1) { startPosition += 9; }
                                    startPosition = fileText.IndexOf ("\r\n", startPosition);
                                    if (startPosition > -1) { startPosition += 4; }
                                }
                                endPosition = fileText.IndexOf ("--BOUNDARY", startPosition) - 1;
                            }
                            len = endPosition - startPosition;
                            if (len == 0)
                            {

                                Console.Write ("Nothing in the message field: " + messageID + " - " + msgField + "\n");
                                // try retrieving something from the from field
                                startPosition = fileText.IndexOf ("FROM:");
                                if (startPosition > -1) { startPosition += 6; }
                                endPosition = fileText.IndexOf ("@", startPosition);
                                len = endPosition - startPosition;
                            }
                            if (len == 0)
                            {
                                Console.Write ("Trying the subject field: " + file + "\n");
                                // try retrieving something from the subject field
                                startPosition = fileText.IndexOf ("SUBJECT:");
                                if (startPosition > -1) { startPosition += 9; }
                                endPosition = fileText.IndexOf ("\r\n", startPosition) - 1;
                                len = fileText.Length;
                                len = endPosition - startPosition;
                            }

                            if (len < 0)
                            {
                                Console.Write ("endPostion is less than startPosition in: " + file + "\n");
                                Console.Write ("Break at line 881ish. Press enter to continue. messageID =" + messageID);
                                input = Console.ReadLine ();
                                break;
                            }
                            msgField = fileText.Substring (startPosition, len);
                            // int lineBreak = fileText.IndexOf ("=\r\n");

                            msgField = msgField
                                .Replace ("I AM SAFE AND WELL.", "")
                                .Replace ("EXERCISE", "")
                                .Replace ("=\r\n", "") // some messages get a line wrap that messes things up
                                .Replace ("=20", "") // asci hex space
                                .Replace ("=0A", "") // asci hex new line / line feed
                                .Replace ("=0D\r\n", "") // asci hex carriage return
                                .Replace ("=0", "")  // asci hex null
                                .Replace ("16. CONTACT INFO:", ",")
                                .Replace (":", "")
                                .Replace (";", ",")
                                .Replace("<","")
                                .Replace (">", "")
                                .Trim ()
                                .Replace ("  ", " ")
                                .Replace ("  ", " ")
                                //.Replace(".", "") this causes problems with decimal band freq
                                .Replace (", ", ",")
                                .Replace ("[NO CHANGES OR EDITING OF THIS MESSAGE ARE ALLOWED]", "")
                                .Replace ("[MESSAGE RECEIPT REQUESTED]", "")
                                .Replace (" ,", ",")
                                .Replace ("\"", "")
                                .Replace ("/", ",") // this was to allow Radiogram forms to use "/" since commas are not permitted
                                .Trim ()
                                //.Trim(',')
                                + ",";
                            // string checkinFrom = checkIn;
                            // 20250113 if (msgField.IndexOf ("GLAWN Ask Template Exercise") > -1) exerciseCompleteCt++;
                            // 20250127 if (ICS201Ct >0) exerciseCompleteCt++;
                            if (radioGram > 0) exerciseCompleteCt++;
                            if (radioGram > 0) msgField = msgField.Replace ("\r\n", " "); // Radiogram chops the message into 40 byte strings, so put it back together

                            checkIn = msgField
                                .Replace (" ", "")
                                //.Trim()
                                //.Trim(',')
                                //.Trim()+",";
                                ;

                            // Create a Regex object with the pattern
                            // and find the first callsign match in the checkIn string
                            // only use the first line
                            endPosition = checkIn.IndexOf ("\r\n");
                            if (endPosition > -1) checkIn = checkIn.Substring (0, endPosition);


                            // Split the checkin string into an array of strings
                            checkIn = removeFieldNumber (checkIn);
                            string [] checkinItems = checkIn.Split (",");

                            // now check to see if it is a perfect message and deduct points if not
                            // checkin call sign
                            checkIn = checkinItems [0];
                            // look for a callsign typo in the checkin msg
                            // do not flage checkins with an appended "/x" as a typo, but make sure it is removed to not break Winlink
                            if (checkIn != fromTxt && xSource != "SMTP" && checkIn != "W5SJT") // W5SJT uses a personal account to login for Tom Green County Emergency Management
                            {
                                if (checkIn.IndexOf("/") == -1 ) callSignTypo = checkIn;
                                checkIn = fromTxt; // assume the message checkin/callsign has a typo
                            }

                            checkIn = isValidCallsign (checkIn);
                            if (checkIn != "")
                            {
                                // checkIn = match.Value;
                                // if they put my callsign in the message, discard it and look at the xSource tag
                                // also ignore xSource == "SMTP" because it will be a checkin via internet email
                                if (checkIn == "KB7WHO" && xSource != "KB7WHO-13" && xSource != "SMTP") { checkIn = xSource; }
                                if (xSource == "") { xSource = checkIn; }

                            }
                            else
                            {
                                isPerfect = false;
                                score--;
                                pointsOff = "\tminus 1 for invalid or missing callsign as the first field - " + checkinItems [0] + "\r\n";
                                // try the from field since the callsign could not be located in the msg field
                                startPosition = fileText.IndexOf ("FROM:");
                                if (startPosition > -1) { startPosition += 6; }
                                endPosition = fileText.IndexOf ("@", startPosition);
                                if (endPosition < 0) { endPosition = fileText.IndexOf ("SUBJECT:") - 1; }
                                len = endPosition - startPosition;
                                if (len > 0)
                                {
                                    checkIn = fileText.Substring (startPosition, len);
                                    checkIn = isValidCallsign (checkIn);

                                }
                            }
                            // debug Console.Write("Start at:"+startPosition+": and end at:"+endPosition+"\nCallsign found: "+checkIn);
                            // eliminate duplicates from the map file                          
                            if (checkIn == "")
                            {
                                Console.Write ("Callsign not found in: " + file);
                            }
                            else
                            {
                                // continue checking for perfect message and point deductions
                                // int checkInItemsCt = checkinItems.Length;
                                // int i = 0;
                                checkinCountry = "";
                                checkinCountryLong = "";
                                len = checkinItems.Length;
                                if (len < 8)
                                {
                                    score = score - (8 - len);
                                    pointsOff += "\tminus " + (8 - len) + " point(s), for missing commas/fields.\r\n";
                                    isPerfect = false;
                                }

                                if (checkinItems.Length > 2)
                                {
                                    // array is zero based
                                    checkinName = isValidName (checkinItems [1]);
                                    if (checkinName == "")
                                    {
                                        isPerfect = false;
                                        score--;
                                        reminderTxt2 = "\tminus 1 point, missing or invalid name in field 2 - " + checkinItems [1] + " 3\r\n";
                                    }
                                }

                                if (checkinItems.Length >= 6)
                                {
                                    string countries = "COL,AFG,ALA,ALB,DZA,ASM,AND,AGO,AIA,ATA,ATG,ARG,ARM,ABW,AUS,AUT,AZE,BHS,BHR,BGD,BRB,BLR,BEL,BLZ,BEN,BMU,BTN,BOL,BIH,BWA,BVT,BRA,IOT,VGB,BRN,BGR,BFA,BDI,KHM,CMR,CAN,CPV,BES,CYM,CAF,TCD,CHL,CHN,CXR,CCK,COL,COM,COK,CRI,HRV,CUB,CUW,CYP,CZE,COD,DNK,DJI,DMA,DOM,TLS,ECU,EGY,SLV,GNQ,ERI,EST,SWZ,ETH,FLK,FRO,FSM,FJI,FIN,FRA,GUF,PYF,ATF,GAB,GMB,GEO,DEU,GHA,GIB,GRC,GRL,GRD,GLP,GUM,GTM,GGY,GIN,GNB,GUY,HTI,HMD,HND,HKG,HUN,ISL,IND,IDN,IRN,IRQ,IRL,IMN,ISR,ITA,CIV,JAM,JPN,JEY,JOR,KAZ,KEN,KIR,XXK,KWT,KGZ,LAO,LVA,LBN,LSO,LBR,LBY,LIE,LTU,LUX,MAC,MDG,MWI,MYS,MDV,MLI,MLT,MHL,MTQ,MRT,MUS,MYT,MEX,MDA,MNG,MNE,MSR,MAR,MOZ,MMR,NAM,NRU,NPL,NLD,NCL,NZL,NIC,NER,NGA,NIU,NFK,PRK,MKD,MNP,NOR,OMN,PAK,PLW,PSE,PAN,PNG,PRY,PER,PHL,PCN,POL,PRT,MCO,PRI,QAT,COG,REU,ROU,RUS,RWA,BLM,SHN,KNA,LCA,MAF,SPM,VCT,WSM,SMR,STP,SAU,SEN,SRB,SYC,SLE,SGP,SXM,SVK,SVN,SLB,SOM,ZAF,SGS,KOR,SSD,ESP,LKA,SDN,SUR,SJM,SWE,CHE,SYR,TWN,TJK,TZA,THA,TGO,TKL,TON,TTO,TUN,TUR,TKM,TCA,TUV,UGA,UKR,ARE,GBR,UMI,USA,URY,UZB,VUT,VAT,VEN,VNM,VIR,WLF,ESH,YEM,ZMB,ZWE,";
                                    checkinCountry = isValidField (checkinItems [5], countries);
                                    countries = "COLOMBIA,BELGIUM,PHILIPPINES,TRINIDAD,GERMANY,ENGLAND,NORWAY,NEW ZEALAND,ST LUCIA,VENEZUELA,AUSTRIA,ROMANIA,CANADA,SERBIA";
                                    checkinCountryLong = isValidField (checkinItems [5], countries);
                                    if (checkinCountry == "")
                                    {
                                        isPerfect = false;
                                        score--;
                                        pointsOff += "\tminus 1 point, missing or invalid country in field 6 (3 letter abbreviation?) - " + checkinItems [5];
                                    }
                                }
                                if (checkinItems.Length >= 5)
                                {
                                    checkinState = checkinItems [4].Replace (".", "");
                                    int scoreState = 0;
                                    string tempStr = "";
                                    string tempStr2 = "";
                                    string states = "";
                                    reminderTxt2 = "";
                                    // check for full country name to be able to make suggestions
                                    switch (checkinCountryLong)
                                    {
                                        case "AUSTRIA":
                                            reminderTxt2 += ", try AUT";
                                            checkinCountry = "AUT";
                                            break;
                                        case "BELGIUM":
                                            reminderTxt2 += ", try BEL";
                                            checkinCountry = "BEL";
                                            break;
                                        case "CANADA":
                                            reminderTxt2 += ", try CAN";
                                            checkinCountry = "CAN";
                                            break;
                                        case "ENGLAND":
                                            reminderTxt2 += ", try GBR";
                                            checkinCountry = "GBR";
                                            break;
                                        case "GERMANY":
                                            reminderTxt2 += ", try DEU";
                                            checkinCountry = "DEU";
                                            break;
                                        case "NEW ZEALAND":
                                            reminderTxt2 += ", try NZL";
                                            checkinCountry = "NZL";
                                            break;
                                        case "NORWAY":
                                            reminderTxt2 += ", try NOR";
                                            checkinCountry = "NOR";
                                            break;
                                        case "PHILIPPINES":
                                            reminderTxt2 += ", try PHL";
                                            checkinCountry = "PHL";
                                            break;
                                        case "ROMANIA":
                                            reminderTxt2 += ", try ROU";
                                            checkinCountry = "ROU";
                                            break;
                                        case "SERBIA":
                                            reminderTxt2 += ", try SRB";
                                            checkinCountry = "SRB";
                                            break;
                                        case "ST LUCIA":
                                            reminderTxt2 += ", try LCA";
                                            checkinCountry = "LCA";
                                            break;
                                        case "TRINIDAD":
                                            reminderTxt2 += ", try TTO";
                                            checkinCountry = "TTO";
                                            break;
                                        case "US":
                                            reminderTxt2 += ", try USA";
                                            checkinCountry = "USA";
                                            break;
                                        case "VENEZUELA":
                                            reminderTxt2 += ", try VEN";
                                            checkinCountry = "VEN";
                                            break;
                                        default:
                                            break;
                                    }
                                    switch (checkinCountry)
                                    {
                                        case "AUT":  // Austria AUT
                                            states = "B,K,N,S,ST,T,O,W,V";
                                            checkinState = isValidField (checkinState, states);
                                            if (checkinState == "")
                                            {
                                                isPerfect = false;
                                                tempStr += "missing or invalid AUT state abbreviation ";
                                                scoreState++;
                                            }
                                            break;
                                        case "BEL": // Belgium BEL
                                            break;
                                        case "CAN": // Canada CAN
                                            states = "NL,PE,NS,NB,QC,ON,MB,SK,AB,BC,YT,NT,NU";
                                            checkinState = isValidField (checkinState, states);
                                            if (checkinState == "")
                                            {
                                                isPerfect = false;
                                                tempStr += "missing or invalid CAN province abbreviation ";
                                                scoreState++;
                                            }
                                            break;
                                        case "DEU": // Deutschland - Germany DEU
                                            states = "BW,BY,BE,BB,HB,HH,HE,MV,NI,NW,RP,SL,SN,ST,SH,TH";
                                            if (checkinState == "")
                                            {
                                                isPerfect = false;
                                                // tempStr += "missing or invalid DEU state abbreviation ";
                                                tempStr += "fehlendes oder ungültiges DEU-Landeskürzel ";
                                                scoreState++;
                                            }
                                            break;
                                        case "GBR":
                                        case "UK": // United Kingdom UK Great Britain GBR
                                                   // states = "ABE,ABD,ANS,AGB,CLK,DGY,DND,EAY,EDU,ELN,ERW,EDH,ELS,FAL,FIF,GLG,HLD,IVC,MLN,MRY,NAY,NLK,ORK,PKN,RFW,SCB,ZET,SAY,SLK,STG,WDU,WLN,";
                                                   // if (checkinState == "")
                                                   // {
                                                   // isPerfect = false;
                                                   // tempStr += "missing or invalid GBR council area abbreviation ";
                                                   // scoreState++;
                                                   // }
                                            break;
                                        case "NZL": // New Zealand NZL
                                            states = "AUK,BOP,CAN,GIS,WGN,HKB,MWT,MWT,MBH,NSN,NTL,OTA,STL,TKI,TKI,TAS,HKB,WGN,WTC,STL,GIS,NTL,TAS,BOP,AUK,WKO,WKO,CAN,WTC,NSN,OTA,";
                                            if (checkinState == "")
                                            {
                                                isPerfect = false;
                                                tempStr += "missing or invalid NZL region abbreviation ";
                                                scoreState++;
                                            }
                                            break;
                                        case "NOR": // Norway NOR
                                            break;
                                        case "PHL": // Philippines PHL
                                            states = "ABR,AGN,AGS,AKL,ALB,ANT,APA,AUR,BAN,BAS,BEN,BIL,BOH,BTG,BTN,BUK,BUL,CAG,CAM,CAN,CAP,CAS,CAT,CAV,CEB,COM,DAO,DAS,DAV,DIN,DVO,EAS,GUI,IFU,ILI,ILN,ILS,ISA,KAL,LAG,LAN,LAS,LEY,LUN,MAD,MAS,MDC,MDR,MGN,MGS,MOU,MSC,MSR,NCO,NCR,NEC,NER,NSA,NUE,NUV,PAM,PAN,PLW,QUE,QUI,RIZ,ROM,SAR,SCO,SIG,SLE,SLU,SOR,SUK,SUN,SUR,TAR,TAW,WSA,ZAN,ZAS,ZMB,ZSI";
                                            if (checkinState == "")
                                            {
                                                isPerfect = false;
                                                tempStr += "missing or invalid PHL region abbreviation ";
                                                scoreState++;
                                            }
                                            break;
                                        case "ROU": // Romania ROU 
                                                    // states = "AB,AG,AR,B,BC,BH,BN,BR,BT,BV,BZ,CJ,CL,CS,CT,CV,DB,DJ,GJ,GL,GR,HD,HR,IF,IL,IS,MH,MM,MS,NT,OT,PH,SB,SJ,SM,SV,TL,TM,TR,VL,VN,VS,";
                                            break;
                                        case "SRB": // Serbia SRB
                                            break;
                                        case "LCA": // St. Lucia LCA
                                            break;
                                        case "TTO": // Trinidad & Tobago TTO
                                            break;
                                        case "USA": // United States of America USA
                                            states = "AK,AL,AR,AS,AZ,CA,CO,CT,DC,DE,FL,GA,GU,HI,IA,ID,IL,IN,KS,KY,LA,MA,MD,ME,MI,MN,MO,MP,MS,MT,NC,ND,NE,NH,NJ,NM,NV,NY,OH,OK,OR,PA,PR,RI,SC,SD,TN,TX,UM,UT,VA,VI,VT,WA,WI,WV,WY";
                                            checkinState = isValidField (checkinState, states);
                                            if (checkinState == "")
                                            {
                                                isPerfect = false;
                                                tempStr += "missing or invalid USA state 2 letter abbreviation ";
                                                if (checkinItems [4] == "PUERTO RICO") tempStr2 += ", try \"PR\"";
                                                scoreState++;
                                            }
                                            break;
                                        case "VEN": // Venezuela VEN
                                            states = "DC,A,B,C,D,E,F,G,H,I,J,K,L,M,N,O,P,R,S,T,U,V,W,X,Y,Z";
                                            checkinState = isValidField (checkinState, states);
                                            if (checkinState == "")
                                            {
                                                isPerfect = false;
                                                tempStr += "VEN - abreviación del estado falta o es inválido ";
                                                scoreState++;
                                            }
                                            break;


                                        default:
                                            scoreState++;
                                            tempStr = "missing or invalid state/province/region (due to missing country?) ";
                                            break;
                                    }
                                    if (reminderTxt2 != "" || checkinCountry == "")
                                        pointsOff += reminderTxt2 + "\r\n";
                                    reminderTxt2 = "";

                                    if (scoreState > 0)
                                    {
                                        pointsOff += "\tminus 1 point, " + tempStr + "in field 5 -  " + checkinItems [4] + tempStr2 + "\r\n";
                                        score--;
                                    }

                                    if (checkinItems.Length > 4 && checkinCountry == "USA") // check only for USA
                                    {
                                        checkinCounty = isValidName (checkinItems [3].Replace (" COUNTY", ""));
                                        if (checkinCounty == "")
                                        {
                                            isPerfect = false;
                                            //score--;
                                            //pointsOff += "\tminus 1 point, invalid county in field 4 - " + checkinItems [3] + "\r\n";
                                            pointsOff += "\tno deduction for now, invalid county in field 4, use NA or NONE if you don't have one - " + checkinItems [3] + "\r\n";
                                        }
                                    }

                                    if (checkinItems.Length > 3)
                                    {
                                        checkinCity = isValidName (checkinItems [2]);
                                        if (checkinCity == "")
                                        {
                                            isPerfect = false;
                                            score--;
                                            pointsOff += "\tminus 1 point, missing or invalid city in field 3 " + checkinItems [2] + "\r\n";
                                        }
                                    }

                                }
                                if (checkinItems.Length >= 7)
                                {
                                    bandStr = checkinItems [6];
                                    bandStr = checkBand (bandStr);
                                    if (bandStr == "")
                                    {
                                        isPerfect = false;
                                        score--;
                                        pointsOff += "\tminus 1 point, missing or invalid band in field 7 - " + checkinItems [6] + "\r\n";
                                    }
                                    else { checkinItems [6] = bandStr; } // update the item if it was adjusted in the method for minor formatting issues
                                }
                                if (checkinItems.Length >= 8)
                                {
                                    modeStr = checkinItems [7];
                                    modeStr = checkMode (modeStr, bandStr);
                                    string tempStr = "";
                                    if (modeStr == "")
                                    {
                                        isPerfect = false;
                                        score--;
                                        if (checkinItems [7] == "VHF") tempStr = ", try \"VARA FM\" or \"PACKET\"";
                                        if (checkinItems [7].Contains ("PACKET")) tempStr = ", try just \"PACKET\"";
                                        if (bandStr == "TELNET") tempStr = ", try SMTP";
                                        pointsOff += "\tminus 1 point, missing or invalid mode in field 8 - " + checkinItems [7] + tempStr + "\r\n";
                                    }
                                }
                                //    i++;
                                // }
                                // check to see if this is a duplicate checkin
                                startPosition = testString.IndexOf (checkIn);
                                if (startPosition >= 0)
                                {
                                    if (dupCt == 0) { duplicates.Append ("Duplicates: \r\n\t"); }
                                    //debug Console.Write("netName "+checkIn+" is a duplicate, skipping. It is "+dupCt+" of "+msgTotal+" total messages.\n");
                                    duplicates.Append (checkIn + ", ");
                                    dupeFlag = 1;
                                    dupCt++;
                                }

                                ct++;
                                if (localWeather > -1) { localWeatherCt++; }
                                if (severeWeather > -1) { severeWeatherCt++; }
                                if (winlinkCkin > -1) { winlinkCkinCt++; }
                                if (incidentStatus > -1) { incidentStatusCt++; }
                                if (ics > -1) { icsCt++; }
                                if (damAssess > -1) { damAssessCt++; }
                                if (fieldSit > -1) { fieldSitCt++; }
                                if (quickM > -1) { quickMCt++; }
                                if (dyfi > -1) { dyfiCt++; }
                                if (rriWR > -1) { rriCt++; }
                                if (qwm > -1) { qwmCt++; }
                                if (mi > -1) { miCt++; }
                                if (ICS201 > -1) ICS201Ct++; 
                                testString = testString + checkIn + " | ";
                                // the spreadsheet chokes if the string ends with "|" so
                                // don't let that happen by writing the first one without a delimiter
                                // prepending the delimiter to the rest.
                                if (ct == 1)
                                {
                                    netCheckinString.Append (checkIn);
                                    netAckString2.Append (checkIn);
                                }
                                else if (ct > 1 && dupeFlag == 0)
                                {
                                    netCheckinString.Append ("|" + checkIn);
                                    netAckString2.Append (";" + checkIn);
                                }
                                
                                // find message, format for csv file, and save
                                var msgFieldStart = msgField.IndexOf ("\r\n");
                                string notFirstLine = "";
                                if (startPR > -1 || copyPR > -1)
                                {
                                    addonString.Append (checkIn + ":\t" + msgField.Replace ("\n", ", ").Replace ("\r", "") + "\r\n");
                                }
                                else
                                {
                                    if (msgFieldStart > -1)
                                    {
                                        len = msgField.Length - msgFieldStart;
                                        if (len > 0)
                                        {
                                            notFirstLine = msgField.Substring (msgFieldStart, len);
                                            notFirstLine = notFirstLine.Replace ("\n", ", ")
                                            .Replace ("\r", "")
                                            //.Replace("73","")
                                            .Trim ()
                                            ;
                                            startPosition = notFirstLine.IndexOf ("73");
                                            if (startPosition > -1)
                                            {
                                                endPosition = notFirstLine.IndexOf ("\r\n", startPosition) + 2;
                                                len = endPosition - startPosition;
                                                if (len > 0)
                                                {
                                                    notFirstLine = notFirstLine.Substring (0, startPosition) + notFirstLine.Substring (endPosition);
                                                }
                                                else
                                                {
                                                    notFirstLine = notFirstLine.Substring (0, startPosition);
                                                }
                                            }
                                            notFirstLine = notFirstLine
                                                .Replace (", ,", ",")
                                                .Trim ()
                                                .Trim (',')
                                                .Trim ()
                                                .Trim (',')

                                                ;
                                            if (notFirstLine.Length > 0) { addonString.Append (checkIn + ":\t" + notFirstLine + "\r\n"); }
                                        }
                                    }

                                }
                                // Extract latitude and longitude
                                // Winlink Checkin has its own tags so check them first


                                if (winlinkCkin > 0)
                                {
                                    startPosition = fileText.IndexOf ("LATITUDE:");
                                    if (startPosition > -1)
                                    {
                                        startPosition += 10;
                                        endPosition = fileText.IndexOf ("\r\n", startPosition);
                                        len = endPosition - startPosition;
                                        if (len > 0)
                                        {
                                            latitudeStr = fileText.Substring (startPosition, len);
                                            // check for tag that didn't fill correctly
                                            if (latitudeStr.IndexOf ("<VAR") != -1)
                                            {
                                                latitudeStr = "";
                                            }
                                            if (latitudeStr != "") { latitude = Common.ConvertToDouble (latitudeStr); }
                                        }
                                    }
                                    startPosition = fileText.IndexOf ("LONGITUDE:");
                                    if (startPosition > -1)
                                    {
                                        startPosition += 11;
                                        endPosition = fileText.IndexOf ("\r\n", startPosition);
                                        len = endPosition - startPosition;
                                        if (len > 0)
                                        {
                                            longitudeStr = fileText.Substring (startPosition, len);
                                            // check for tag that didn't fill correctly
                                            var testGPS = longitudeStr.IndexOf ("<VAR");
                                            if (longitudeStr.IndexOf ("<VAR") != -1)
                                            {
                                                longitudeStr = "";
                                            }
                                            if (longitudeStr != "") { longitude = Common.ConvertToDouble (longitudeStr); }
                                        }
                                    }
                                }
                                if (latitudeStr == "" || longitudeStr == "")
                                {
                                    //skip past the messageID because sometimes the regex for coordinates matches it
                                    // and stop before reading any binary attachments

                                    // startPosition = quotedPrintable;
                                    startPosition = fileText.IndexOf ("MESSAGE-ID:");
                                    startPosition = fileText.IndexOf ("\r\n", startPosition);
                                    if (startPosition > -1) { startPosition += 2; }
                                    // need an end position because some messages have a binary attachment that gives a false match
                                    endPosition = fileText.IndexOf ("PRINTABLE", startPosition);
                                    endPosition = fileText.IndexOf ("--BOUNDARY", endPosition);                                    
                                    len = endPosition - startPosition;
                                    if (len > 0)
                                    {
                                        if (ExtractCoordinates (fileText.Substring (startPosition, len), out latitude, out longitude))
                                        {
                                            // Console.WriteLine(messageID+" latitude: "+latitude+" longitude: "+longitude);                                
                                            maidenheadGrid = ExtractMaidenheadGrid (fileText.Substring (startPosition, len));
                                        }
                                        else
                                        {
                                            // no valid GPS coordinates found, look for a maidenhead grid
                                            maidenheadGrid = ExtractMaidenheadGrid (fileText.Substring (startPosition, len));
                                            if (!string.IsNullOrEmpty (maidenheadGrid))
                                            {
                                                // Console.WriteLine($"Maidenhead Grid: {maidenheadGrid}");
                                                // Convert Maidenhead to GPS coordinates
                                                (latitude, longitude) = MaidenheadToGPS (maidenheadGrid);
                                                // Console.WriteLine($"No GPS coords found, using Maidenhead Grid: {maidenheadGrid}"+$". From Maidenhead Grid Latitude: {latitude}"+$"  Longitude: {longitude}");

                                            }
                                            // else if (latitude == 0 && longitude == 0)
                                            // { 
                                            // try regex to try and find it
                                            //  latitude
                                            // }
                                            else
                                            {
                                                // No valid Maidenhead grid found either, make up something in the middle of the Atlantic
                                                double locChange = Math.Round (rnd.NextDouble () * 10, 6);
                                                latitude = Math.Round ((27.187512 + locChange), 6);
                                                longitude = Math.Round ((-60.144742 + locChange), 6);
                                                // Console.WriteLine("No valid grid and no GPS coordinates found in: "+messageID+" latitude set to: "+latitude+" longitude set to: "+longitude);
                                                noGPSCt++;
                                                noGPSFlag++;
                                                score -= 2;
                                                msgField = msgField + ",No Location Data Found in message";
                                                noGPSString.Append ("\t" + messageID + "- - " + checkIn + ": latitude set to: " + latitude + " longitude set to: " + longitude + "\r\n");
                                            }
                                        }
                                    }
                                    if (latitude > 0)
                                    {
                                        latitudeStr = Convert.ToString (latitude);
                                        longitudeStr = Convert.ToString (longitude);
                                    }
                                }
                                msgField = msgField.Replace ("\r\n", ",");
                                msgField = removeFieldNumber (msgField);
                                msgFieldNumbered = fillFieldNum (msgField);
                                csvString.Append (checkIn + ":" + messageID + "," + latitude + "," + longitude + "," + locType + "," + msgField + "\r\n");

                                // find the band if it's where it's supposed to be
                                bandStr = "";
                                modeStr = "";

                                // debug Console.Write("\r\nmsgField ="+msgField+"\r\n");
                                startPosition = IndexOfNthSB (msgField, (char)44, 0, 6) + 1;
                                if (startPosition > -1) { endPosition = IndexOfNthSB (msgField, (char)44, 0, 7); len = endPosition - startPosition; }
                                if (len > 0 && msgField.Length >= len)
                                {
                                    bandStr = msgField.Substring (startPosition, len);
                                }
                                // winlink checkin has fields for band and mode so look there if not in the message
                                if (bandStr == "" && winlinkCkin > 0)
                                {
                                    startPosition = fileText.IndexOf ("BAND:");
                                    if (startPosition == -1)
                                    {
                                        // try another label that I have seen instead
                                        startPosition = fileText.IndexOf ("BAND USED:");
                                        if (startPosition > -1) { startPosition += 11; }
                                    }
                                    else { startPosition = startPosition + 6; }

                                    if (startPosition > -1)
                                    {
                                        endPosition = fileText.IndexOf ("\r\n", startPosition);
                                        len = endPosition - startPosition;
                                        if (len > 0) { bandStr = fileText.Substring (startPosition, len); }
                                    }
                                }
                                if (bandStr == "")
                                {
                                    if (msgField.IndexOf ("VARA FM") > -1)
                                    {
                                        modeStr = "VARA FM";
                                        // bandStr = "VARA FM";
                                    }
                                    if (msgField.IndexOf ("VARA HF") > -1)
                                    {
                                        modeStr = "VARA HF";
                                        // bandStr = "VARA HF";
                                    }
                                }
                                bandStr = bandStr
                                    .ToUpper ()
                                    .Replace ("5.8GHZ", "5CM")
                                    .Replace (".", "")
                                    .Replace (" ", "")
                                    .Replace (" METERS", "M")
                                    .Replace (" METER", "M")
                                    .Replace ("METERS", "M")
                                    .Replace ("METER", "M")
                                    .Replace (")", "")
                                    .Replace ("N/A", "TELNET")
                                    .Replace ("NA", "TELNET")
                                    .Replace ("5GHZ", "5CM")
                                    .Replace ("73M", "80M")
                                    .Replace ("75M", "80M")
                                    .Replace (" M", "M")
                                    .Trim ()
                                    .Replace ("HFAMATEUR", "HF");
                                if (bandStr.IndexOf ("PACKET") > -1)
                                {
                                    modeStr = "PACKET";
                                    // if the band is declared to be packet, check to see if there is any indication of the band elsewhere in the message
                                    if (msgField.IndexOf ("2M") > -1) { bandStr = "2M"; }
                                    if (msgField.IndexOf ("70CM") > -1) { bandStr = "70CM"; }
                                    if (msgField.IndexOf ("VHF") > -1) { bandStr = "VHF"; }
                                    if (msgField.IndexOf ("UHF") > -1) { bandStr = "UHF"; }
                                }
                                if (bandStr.IndexOf ("2M") > -1) { bandStr = "2M"; }
                                if (msgField.IndexOf ("VARAFM") > -1)
                                {
                                    modeStr = "VARA FM";
                                    if (msgField.IndexOf ("2M") > -1) { bandStr = "2M"; }
                                    if (msgField.IndexOf ("70CM") > -1) { bandStr = "70CM"; }
                                    if (msgField.IndexOf ("VHF") > -1) { bandStr = "VHF"; }
                                    if (msgField.IndexOf ("UHF") > -1) { bandStr = "UHF"; }
                                }
                                if (msgField.IndexOf ("VARAHF") > -1) { bandStr = "HF"; modeStr = "VARA HF"; }
                                // if (bandStr != "") { bandCt++; }

                                bandStr = checkBand (bandStr);
                                if (bandStr == "")
                                {
                                    // if both the band and the mode have invalid data, try scraping through the msgField
                                    if (msgField.IndexOf ("3CM") > -1) { bandStr = "3CM"; }
                                    if (msgField.IndexOf ("5CM") > -1) { bandStr = "5CM"; }
                                    if (msgField.IndexOf ("13CM") > -1) { bandStr = "13CM"; }
                                    if (msgField.IndexOf ("23CM") > -1) { bandStr = "23CM"; }
                                    if (msgField.IndexOf ("33CM") > -1) { bandStr = "33CM"; }
                                    if (msgField.IndexOf ("70CM") > -1) { bandStr = "70CM"; }
                                    if (msgField.IndexOf ("1.25M") > -1) { bandStr = "1.25M"; }
                                    if (msgField.IndexOf ("2M") > -1) { bandStr = "2M"; }
                                    if (msgField.IndexOf ("10M") > -1) { bandStr = "10M"; }
                                    if (msgField.IndexOf ("12M") > -1) { bandStr = "12M"; }
                                    if (msgField.IndexOf ("15M") > -1) { bandStr = "15M"; }
                                    if (msgField.IndexOf ("17M") > -1) { bandStr = "17M"; }
                                    if (msgField.IndexOf ("20M") > -1) { bandStr = "20M"; }
                                    if (msgField.IndexOf ("30M") > -1) { bandStr = "30M"; }
                                    if (msgField.IndexOf ("40M") > -1) { bandStr = "40M"; }
                                    if (msgField.IndexOf ("60M") > -1) { bandStr = "60M"; }
                                    if (msgField.IndexOf ("6M") > -1) { bandStr = "6M"; }
                                    if (msgField.IndexOf ("80M") > -1) { bandStr = "80M"; }
                                    if (msgField.IndexOf ("HF") > -1) { bandStr = "HF"; }
                                    if (msgField.IndexOf ("VHF") > -1) { bandStr = "VHF"; }
                                    if (msgField.IndexOf ("UHF") > -1) { bandStr = "UHF"; }
                                    if (msgField.IndexOf ("SHF") > -1) { bandStr = "SHF"; }
                                    if (msgField.IndexOf ("TELNET") > -1) { bandStr = "TELNET"; }

                                    if (bandStr == "")
                                    {
                                        // msgFieldNumbered = msgField;
                                        // msgFieldNumbered = fillFieldNum (msgFieldNumbered);
                                        badBandString.Append ("\tBad band: " + messageID + " - " + checkIn + ": _" + bandStr + "_  |  " + msgFieldNumbered + "\r\n");
                                        badBandCt++;
                                    }
                                }
                                else { bandCt++; }

                                if (modeStr == "")
                                {
                                    // debug Console.Write("\r\nmsgField ="+msgField+"\r\n");
                                    startPosition = IndexOfNthSB (msgField, (char)44, 0, 7);
                                    // if (startPosition > -1) { startPosition += 1; }
                                    if (startPosition > -1)
                                    {
                                        endPosition = IndexOfNthSB (msgField, (char)44, 0, 8);
                                        startPosition += 1;
                                        len = endPosition - startPosition;
                                        if (len > 0 && msgField.Length >= len)
                                        {
                                            modeStr = msgField.Substring (startPosition, len);

                                        }
                                    }
                                    if (modeStr == "" && winlinkCkin > 0)
                                    {
                                        startPosition = fileText.IndexOf ("SESSION:");
                                        if (startPosition == -1)
                                        {
                                            // try another label that I have seen instead
                                            startPosition = fileText.IndexOf ("SESSION TYPE:");
                                            if (startPosition > -1)
                                            {
                                                startPosition += 14;
                                            }
                                        }
                                        else { startPosition += 9; }

                                        if (startPosition > -1)
                                        {
                                            endPosition = fileText.IndexOf ("\r\n", startPosition);
                                            len = endPosition - startPosition;
                                            if (len > 0) { modeStr = fileText.Substring (startPosition, len); }
                                        }
                                    }
                                }

                                // if (modeStr.IndexOf (" ") > -1) { modeStr = removeFieldNumber (modeStr); }
                                modeStr = modeStr
                                    .ToUpper ()
                                    .Trim ()
                                    .Replace ("WINLINK", "")
                                    .Replace ("AREDN", "MESH")
                                    //.Replace ("AX.25", "PACKET")
                                    .Replace ("WINLINK", "")
                                    .Replace ("(", "")
                                    .Replace (".", "")
                                    .Replace ("ARDOP HF", "ARDOP")
                                    .Replace ("VARA VHF", "VARA FM")
                                    .Replace ("VHF VARA", "VARA FM")
                                    .Replace ("VARAFM", "VARA FM")
                                    .Replace ("VERA", "VARA")
                                    .Replace ("VARA-HF", "VARA HF")
                                    .Replace ("HF ARDOP", "ARDOP")
                                    .Replace (")", "")
                                    .Replace ("-", " ")
                                    .Replace ("=20", "")
                                    .Replace ("VHF PACKET", "PACKET")
                                    .Replace ("TELNET", "SMTP")
                                    .Trim ();
                                if (modeStr.IndexOf ("MESH") > -1) { modeStr = "MESH"; }
                                if (modeStr.IndexOf ("VARA") > -1 && bandStr == "HF") { modeStr = "VARA HF"; }
                                if (modeStr.IndexOf ("VARA") > -1 && (bandStr == "VHF" || bandStr == "UHF" || bandStr == "SHF" || bandStr == "2M" || bandStr == "70CM" || bandStr == "1.25M" || bandStr == "33CM" || bandStr == "23CM" || bandStr == "13CM" || bandStr == "5CM" || bandStr == "3CM")) { modeStr = "VARA FM"; }
                                if (bandStr == "TELNET") { modeStr = "SMTP"; }

                                if (modeStr != "")
                                {
                                    if (bandStr == "")
                                    {
                                        if (modeStr == "VARA HF") { bandStr = "HF"; }
                                        // if (modeStr == "VARA FM") { bandStr = "VHF"; }
                                    }

                                }
                                modeStr = checkMode (modeStr, bandStr);
                                if (modeStr == "SMTP") { bandStr = "TELNET"; }
                                if (modeStr == "MESH") { meshCt++; }
                                if (modeStr == "")
                                {
                                    if (msgField.IndexOf ("VARA FM") > -1) { modeStr = "VARA FM"; }
                                    if (msgField.IndexOf ("VARA HF") > -1) { modeStr = "VARA HF"; }
                                    if (msgField.IndexOf ("PACKET") > -1) { modeStr = "PACKET"; }
                                    if (msgField.IndexOf ("PACTOR") > -1) { modeStr = "PACTOR"; }
                                    if (msgField.IndexOf ("TELNET") > -1) { modeStr = "SMTP"; bandStr = "TELNET"; }
                                    if (msgField.IndexOf ("SMTP") > -1) { modeStr = "SMTP"; bandStr = "TELNET"; }
                                    if (msgField.IndexOf ("ARDOP") > -1) { modeStr = "ARDOP"; }
                                    if (msgField.IndexOf ("PACKET") > -1) { modeStr = "PACKET"; }
                                    if (msgField.IndexOf ("ROBUST PACKET") > -1) { modeStr = "ROBUST PACKET"; }
                                    if (msgField.IndexOf ("VARA") > -1 && (bandStr == "VHF" || bandStr == "UHF" || bandStr == "SHF" || bandStr == "2M" || bandStr == "70CM" || bandStr == "1.25M" || bandStr == "33CM" || bandStr == "23CM" || bandStr == "13CM" || bandStr == "5CM" || bandStr == "3CM")) { modeStr = "VARA FM"; }

                                    if (msgField.IndexOf ("VARA") > -1 && (bandStr == "HF")) { modeStr = "VARA HF"; }

                                    if (modeStr == "")
                                    {
                                        msgFieldNumbered = msgField;
                                        msgFieldNumbered = fillFieldNum (msgFieldNumbered);
                                        badBandString.Append ("\tBad mode: " + messageID + " - " + checkIn + ": " + modeStr + " -  |  " + msgFieldNumbered + "\r\n");
                                        badModeCt++;
                                    }
                                }
                                else { modeCt++; }

                                // debug Console.Write("modeStr final=|"+modeStr+"|  \r\n");


                                // add to mapString csv file if xloc was found
                                if (latitude != 0)
                                {
                                    if (dupeFlag == 0)
                                    {
                                        mapCt++;
                                    }
                                    else
                                    {
                                        // Remove any lines containing the call sign
                                        RemoveLineContaining (mapString, checkIn);
                                        dupeRemoveCt++;
                                    }
                                    mapString.Append (checkIn + "," + latitude + "," + longitude + "," + bandStr + "," + modeStr + "\r\n");
                                }
                                // Console.WriteLine (checkIn + ":"+messageID+" - "+ ct + " - mapCt:" + mapCt + " - dupCt: " + dupCt);
                                // xml data

                                if (callSignTypo != "")
                                {
                                    reminderTxt += "\r\nCheck for a typo in your callsign in the Message Field: " + callSignTypo + " vs " + fromTxt + "\r\n";
                                    typoString.Append ("\t" + callSignTypo + " vs " + checkIn + " in message " + messageID + "\r\n");
                                    callSignTypo = "";
                                }

                                if (noScore == -1)
                                {
                                    if (isPerfect)
                                    {
                                        reminderTxt += "\r\nThis is a copy of your message (with numbered fields) and extracted data. \r\nMessage: " + msgFieldNumbered + "\r\n\r\nPerfect Message! Your score is 10.";
                                        perfectScoreCt++;
                                    }
                                    else
                                    {
                                        // reminderTxt += "\r\n" + fileText + "\r\nThis is a copy of your message (with numbered fields) and extracted data. \r\nMessage: " + msgFieldNumbered + "\r\n\r\nYour score is: " + score + "\r\n" + pointsOff +
                                        //    "\r\nRecommended format reminder in the Comment/Message field:\r\ncallSign, firstname, city, county, state/province/region, country, band, Mode, grid\r\n" +
                                        //    "Example: xxNxxx, Greg, Sugar City, Madison, ID, USA, HF, VARA HF, DN43du\r\n" +
                                        //    "Example 2: DxNxx,Mario,TONDO,MANILA,NCR,PHL,2M,VARA FM,PK04LO\r\n" +
                                        //    "Example 2: xxNxx,Andre,Burnaby,,BC,CAN,TELNET,SMTP,CN89ud";
                                        reminderTxt += "\r\n" + "\r\nThis is a copy of your message (with numbered fields) and extracted data. \r\nMessage: " + msgFieldNumbered + "\r\n\r\nYour score is: " + score + "\r\n" + pointsOff +
                                               "\r\nRecommended format reminder in the Comment/Message field:\r\ncallSign, firstname, city, county, state/province/region, country, band, Mode, grid\r\n" +
                                               "Example: xxNxxx, Greg, Sugar City, Madison, ID, USA, HF, VARA HF, DN43du\r\n" +
                                               "Example 2: DxNxx,Mario,TONDO,MANILA,NCR,PHL,2M,VARA FM,PK04LO\r\n" +
                                               "Example 2: xxNxx,Andre,Burnaby,,BC,CAN,TELNET,SMTP,CN89ud";
                                    }
                                }
                                else
                                {
                                    noScoreString.AppendLine ("\t" + checkIn + " in message " + messageID);
                                    noScoreCt++;
                                }

                                noGPSFlag = 0;
                                // the old message ID will destroy stuff in winlink if it is the same when trying to post
                                // create a new message ID by rearranging the old one
                                string newMessageID = messageID;
                                newMessageID = ScrambleWord (newMessageID);
                                string sendTo = checkIn;
                                // Tim Conroy, WB8HRO lives in an assisted living space and does not have easy access to 
                                // RF and put in a special request to send acknowledgements to his personal email address
                                if (sendTo == "WB8HRO") sendTo = "xyz191@live.com";
                                if (sendTo == "KB7WHO" || sendTo == "GLAWN") sendTo = "kb7who@gmail.com";
                                // Console.WriteLine("before: "+messageID+   "    after: "+newMessageID);

                                if (isPerfect)
                                {
                                    XElement message_list = xmlPerfDoc.Descendants ("message_list").FirstOrDefault ();
                                    message_list.Add (new XElement ("message",
                                        new XElement ("id", newMessageID),
                                        new XElement ("foldertype", "Fixed"),
                                        new XElement ("folder", "Outbox"),
                                        new XElement ("subject", "GLAWN acknowledgement ", utcDate),
                                        new XElement ("time", utcDate),
                                        new XElement ("sender", "GLAWN"),
                                        new XElement ("To", sendTo),
                                        new XElement ("rmsoriginator", ""),
                                        new XElement ("rmsdestination", ""),
                                        new XElement ("rmspath", ""),
                                        new XElement ("location", "43.845831N, 111.745744W (GPS)"),
                                        new XElement ("csize", ""),
                                        new XElement ("messageserver", ""),
                                        new XElement ("precedence", "2"),
                                        new XElement ("peertopeer", "False"),
                                        new XElement ("routingflag", ""),
                                        // for testing
                                        // new XElement("source", "KB7WHO"),
                                        new XElement ("source", "GLAWN"),
                                        new XElement ("unread", "True"),
                                        new XElement ("flags", "0"),
                                        new XElement ("messageoptions", "False|False|||||"),
                                        new XElement
                                        ("mime", "Date: " + utcDate + "\r\n" +
                                            "From: GLAWN@winlink.org\r\n" +
                                            // for testing
                                            // "From: KB7WHO@winlink.org\r\n"+
                                            "Subject: GLAWN acknowledgement ", utcDate + "\r\n" +
                                            "To: " + sendTo + "\r\n" +
                                            "Message-ID: " + newMessageID + "\r\n" +
                                            // Can't edit if not from my call sign
                                            // "X-Source: GLAWN\r\n"+
                                            // for testing
                                            "X-Source:" + xmlXsource + "\r\n" +
                                            "X-Location: 43.845831N, 111.745744W(GPS) \r\n" +
                                            "MIME-Version: 1.0\r\n" +
                                            "MIME-Version: 1.0\r\n" +
                                            "Thank you for checking in to the GLAWN. \r\n" +
                                           reminderTxt + "\r\n\r\n" +
                                            "\r\nExtracted Data: " + noScore + "\r\n" +
                                                "   Latitude: " + latitude + "\r\n" +
                                                "   Longitude: " + longitude + "\r\n" +
                                                "   Band: " + bandStr + "\r\n" +
                                                "   Mode: " + modeStr + "\r\n" +
                                                "   Original Message ID: " + messageID + "\r\n" +
                                                "\r\nGLAWN Current Map: https://tinyurl.com/GLAWN-Map\r\n" +
                                                "Comments: https://tinyurl.com/GLAWN-comments\r\n" +
                                                "GLAWN Checkins Report: https://tinyurl.com/Checkins-Report\r\n" +
                                                "checkins.csv: https://tinyurl.com/GLAWN-CSV-checkins\r\n" +
                                                "mapfile.csv: https://tinyurl.com/Current-CSV-mapfile\r\n"
                                        )
                                    ));
                                }
                                else
                                {
                                    XElement message_list = xmlDoc.Descendants ("message_list").FirstOrDefault ();
                                    message_list.Add (new XElement ("message",
                                        new XElement ("id", newMessageID),
                                        new XElement ("foldertype", "Fixed"),
                                        new XElement ("folder", "Outbox"),
                                        new XElement ("subject", "GLAWN acknowledgement ", utcDate),
                                        new XElement ("time", utcDate),
                                        new XElement ("sender", "GLAWN"),
                                        new XElement ("To", sendTo),
                                        new XElement ("rmsoriginator", ""),
                                        new XElement ("rmsdestination", ""),
                                        new XElement ("rmspath", ""),
                                        new XElement ("location", "43.845831N, 111.745744W (GPS)"),
                                        new XElement ("csize", ""),
                                        new XElement ("messageserver", ""),
                                        new XElement ("precedence", "2"),
                                        new XElement ("peertopeer", "False"),
                                        new XElement ("routingflag", ""),
                                        // for testing
                                        // new XElement("source", "KB7WHO"),
                                        new XElement ("source", "GLAWN"),
                                        new XElement ("unread", "True"),
                                        new XElement ("flags", "0"),
                                        new XElement ("messageoptions", "False|False|||||"),
                                        new XElement
                                        ("mime", "Date: " + utcDate + "\r\n" +
                                            "From: GLAWN@winlink.org\r\n" +
                                            // for testing
                                            // "From: KB7WHO@winlink.org\r\n"+
                                            "Subject: GLAWN acknowledgement ", utcDate + "\r\n" +
                                            "To: " + sendTo + "\r\n" +
                                            "Message-ID: " + newMessageID + "\r\n" +
                                            // Can't edit if not from my call sign
                                            // "X-Source: GLAWN\r\n"+
                                            // for testing
                                            "X-Source:" + xmlXsource + "\r\n" +
                                            "X-Location: 43.845831N, 111.745744W(GPS) \r\n" +
                                            "MIME-Version: 1.0\r\n" +
                                            "MIME-Version: 1.0\r\n" +
                                            "Thank you for checking in to the GLAWN. \r\n" +
                                           reminderTxt + "\r\n\r\n" +
                                            "\r\nExtracted Data: " + noScore + "\r\n" +
                                                "   Latitude: " + latitude + "\r\n" +
                                                "   Longitude: " + longitude + "\r\n" +
                                                "   Band: " + bandStr + "\r\n" +
                                                "   Mode: " + modeStr + "\r\n" +
                                                "   Original Message ID: " + messageID + "\r\n" +
                                                "\r\nGLAWN Current Map: https://tinyurl.com/GLAWN-Map\r\n" +
                                                "Comments: https://tinyurl.com/GLAWN-comments\r\n" +
                                                "GLAWN Checkins Report: https://tinyurl.com/Checkins-Report\r\n" +
                                                "checkins.csv: https://tinyurl.com/GLAWN-CSV-checkins\r\n" +
                                                "mapfile.csv: https://tinyurl.com/Current-CSV-mapfile\r\n"
                                        )
                                    ));
                                }
                                // Add the message message_list
                                // xmlDoc.Root.Add (messageElement);

                                junk = 0; // just so i could put a debug here
                                dupeFlag = 0; // reset the duplicate flag

                            }
                            var tempCt = ct + dupCt + ackCt + removalCt;
                            //debug Console.Write("checkins:"+ct+"  duplicates:" + dupCt+"  removals:"+removalCt+"  acks:"+ackCt + "  combined:"+tempCt+"   actual total:"+msgTotal+"\n");
                            // missing from roster section. Check to see if the checkin is in the roster. 


                            startPosition = rosterString.IndexOf (checkIn);
                            if (startPosition < 0)
                            {
                                checkIn = isValidCallsign (checkIn);
                                if (checkIn != "")
                                {
                                    Console.Write (checkIn + "  " + messageID + " was not found in roster.txt. \n");
                                    if (checkinCountryLong == "") checkinCountryLong = checkinCountry;
                                    newCheckIns.Append (checkIn + "\t=countif(indirect(\"R[0]C[10]\",FALSE):indirect(\"R[0]C[63]\",FALSE),\">0\"&\"*\")\t" + checkinName + "\t" + checkinCity + "\t" + checkinCounty + "\t" + checkinState + "\t" + checkinCountry + "\t" + checkinCountryLong + "\t" + bandStr + "\t" + modeStr + "\t" + maidenheadGrid + "\r\n");

                                    // update roster.txt to contain the new checkin
                                    File.AppendAllText ("roster.txt", ":1; " + checkIn);
                                    newCt++;
                                }
                            }
                        }
                        else
                        {
                            skipped++;
                            Console.Write ("Could not find netName in this message: " + messageID + "\n");
                            skippedString.AppendLine ("\tNo NetName: \"" + netName + "\" in " + messageID);
                        }
                    }

                }
                junk = 0;
            }
            var tempCT = 15;
            logWrite.WriteLine ("Current GLAWN Checkins posted: " + utcDate);

            logWrite.WriteLine ("    Total Stations Checking in: " + (ct - dupCt) + "    Duplicates: " + dupCt + "    Total Checkins: " + ct + "    Removal Requests: " + removalCt);
            logWrite.WriteLine ("Non-" + netName + " checkin messages skipped: " + skipped + " (including " + ackCt + " acknowledgements and " + outOfRangeCt + " out of date range messages skipped.)\r\n");
            logWrite.WriteLine ("Total messages processed: " + msgTotal + "\r\n");
            logWrite.WriteLine ("Row " + tempCT + " goes into GLAWN Spreadsheet at row 1 of the checkin column to be recorded.");
            tempCT++;
            logWrite.WriteLine ("Row " + tempCT + " goes into GLAWN Spreadsheet at row 2 of the checkin column and is the copy" +
                    "\r\n\tlist for the checkin acknowledgement.");
            tempCT = tempCT + 2;
            logWrite.WriteLine ("Rows " + tempCT + " and beyond have the list of duplicates found, bounced messages\r\n" +
                    "\tnew checkins that should be added to the spreadsheet, skipped messages that didn't " +
                    "\r\n\thave a netName,and other notifications including the checkin forms used, " +
                    "\r\n\tthe number that had mapping coordinates, and the comments.\r\n");

            SortStringBuilder (netCheckinString, "|", 0);
            logWrite.WriteLine (netCheckinString);

            SortStringBuilder (netAckString2, ";", 0);
            netAckString2 = netAckString2.Replace ("\r\n\r\n", "\r\n");

            logWrite.WriteLine (netAckString2 + "\r\n");
            // netCheckinString = SortCommaDelimitedString(netCheckinString,"|");
            // Console.WriteLine(csvString);           

            SortStringBuilder (csvString, "\r\n", 1);
            // Console.WriteLine(csvString.ToString());
            csvWrite.WriteLine (csvString);

            // SortStringBuilder (skippedString, "\r\n", 1);
            // Console.WriteLine(csvString.ToString());

            SortStringBuilder (mapString, "\r\n", 1);
            // Console.WriteLine(mapString);
            mapWrite.WriteLine (mapString);

            SortStringBuilder (addonString, "\r\n", 2);
            commentWrite.WriteLine (addonString);

            xmlPerfDoc.Save (xmlPerfFile);
            xmlDoc.Save (xmlFile);

            if (duplicates.Length != 0) { logWrite.WriteLine (duplicates + "\r\n"); }
            if (bouncedString.Length != 0) { logWrite.WriteLine ("Messages that bounced: " + bouncedString); }
            if (newCheckIns.Length != 0) { logWrite.WriteLine ("New Checkins: \r\n" + newCheckIns); }
            if (skippedString.Length != 0) { logWrite.WriteLine ("Messages Skipped: \r\n" + skippedString); }
            if (removalString.Length != 0) { logWrite.WriteLine ("Requests to be Removed: " + removalString); }
            if (localWeatherCt > 0) { logWrite.WriteLine ("Local Weather Checkins: " + localWeatherCt); }
            if (severeWeatherCt > 0) { logWrite.WriteLine ("Severe Weather Checkins: " + severeWeatherCt); }
            if (incidentStatusCt > 0) { logWrite.WriteLine ("Incident Status Checkins: " + incidentStatusCt); }
            if (icsCt > 0) { logWrite.WriteLine ("ICS-213 Checkins: " + icsCt); }
            if (winlinkCkinCt > 0) { logWrite.WriteLine ("Winlink Check-in Checkins: " + winlinkCkinCt); }
            if (damAssessCt > 0) { logWrite.WriteLine ("Damage Assessment Checkins: " + damAssessCt); }
            if (fieldSitCt > 0) { logWrite.WriteLine ("Field Situation Report Checkins: " + fieldSitCt); }
            if (quickMCt > 0) { logWrite.WriteLine ("Quick H&W: " + quickMCt); }
            if (qwmCt > 0) { logWrite.WriteLine ("Quick Welfare Message: " + qwmCt); }
            if (dyfiCt > 0) { logWrite.WriteLine ("Did You Feel It: " + dyfiCt); }
            if (rriCt > 0) { logWrite.WriteLine ("RRI Welfare Radiogram: " + rriCt); }
            if (miCt > 0) { logWrite.WriteLine ("Medical Incident: " + miCt); }
            if (aprsCt > 0) { logWrite.WriteLine ("APRS Checkins: " + aprsCt); }
            if (meshCt > 0) { logWrite.WriteLine ("Mesh Checkins: " + meshCt); }
            if (PosRepCt > 0) { logWrite.WriteLine ("Position Report Checkins: " + PosRepCt); }
            if (ICS201Ct > 0) { logWrite.WriteLine ("ICS 201 Checkins: " + ICS201Ct); }
            if (radioGram > 0) { logWrite.WriteLine ("Radiogram Checkins: " + radioGramCt); }
            logWrite.WriteLine ("Total Plain and other Checkins: " + (ct - localWeatherCt - severeWeatherCt - incidentStatusCt - icsCt - winlinkCkinCt - damAssessCt - fieldSitCt - quickMCt - dyfiCt - rriCt - qwmCt - miCt - aprsCt - meshCt - PosRepCt - ICS201Ct - radioGramCt) + "\r\n");
            //var totalValidGPS = mapCt-noGPSCt;
            logWrite.WriteLine ("Total Checkins with a perfect message: " + perfectScoreCt);
            logWrite.WriteLine ("Total Checkins with a geolocation: " + (mapCt - noGPSCt));
            // logWrite.WriteLine ("Total Checkins with a geolocation: " + (mapCt - noGPSCt));
            if (exerciseCompleteCt > 0) { logWrite.WriteLine ("Successful Exercise Participation" + exerciseCompleteCt); }
            
            logWrite.WriteLine ("Total Checkins with something in the band field: " + bandCt);
            logWrite.WriteLine ("Total Checkins with something in the mode field: " + modeCt);
            // logWrite.WriteLine("\r\n++++++++++++++++\r\nmsgField not properly formatted for the following: \r\n-------------------------------");
            // logWrite.Write(badBandString);
            if (badBandCt > 0) logWrite.WriteLine ("Checkins with a bad band field: " + badBandCt);
            // logWrite.Write(badModeString);
            if (badModeCt > 0) logWrite.WriteLine ("Checkins with a bad mode field: " + badModeCt);
            if (noGPSCt > 0) logWrite.WriteLine (noGPSString + "Total without a location: " + noGPSCt);
            if (noScoreCt > 0) { logWrite.Write (noScoreString); logWrite.WriteLine ("Messages not scored: " + noScoreCt + "\r\n++++++++++++++++\r\n"); }
            if (typoString.Length != 0) logWrite.WriteLine ("Callsign typo: \r\n" + typoString);
            logWrite.Write ("++++++++++++++++\r\n" + addonString);

        }
        Console.WriteLine ("Done!\nThere were " + ct + " checkins. \nThe output files can be found in the folder:\n" + currentFolder);
        Console.WriteLine ("\n\nPress enter to continue.");
        Console.ReadLine ();
    }
    //public static class Globals
    public static int IndexOfNthSB (string input,
             char value, int startIndex, int nth)
    {
        if (nth < 1)
            throw new NotSupportedException ("Param 'nth' must be greater than 0!");
        var nResult = 0;
        for (int i = startIndex; i < input.Length; i++)
        {
            if (input [i] == value)
                nResult++;
            if (nResult == nth)
                return i;
        }
        return -1;
    }
    static DateTime GetValidDate ()
    {
        DateTime date = default;  // Initialize date to its default value
        bool isValid = false;

        while (!isValid)
        {
            string input = Console.ReadLine ();
            DateTime todayDate = DateTime.Today;
            int dateCompare = 0;

            // Validate using the specific format YYYYMMDD
            if (DateTime.TryParseExact (input, "yyyyMMdd", null, System.Globalization.DateTimeStyles.None, out date))
            {
                isValid = true;
                dateCompare = DateTime.Compare (date, todayDate.AddDays (-14));
                if (dateCompare < 0)
                {
                    isValid = false;
                    Console.WriteLine ("Invalid date: " + date + " Must be within two weeks of today. Please try again.");
                }
                dateCompare = DateTime.Compare (todayDate.AddDays (14), date);
                if (dateCompare < 0)
                {
                    isValid = false;
                    Console.WriteLine ("Invalid date: " + input + " Must be within two weeks of today.  Please try again.");
                }
            }
            else
            {
                Console.WriteLine ("Invalid date format. " + input + "Please use YYYYMMDD format and try again.");
            }
        }
        return date;
    }

    static void SaveDate (DateTime date)
    {
        string filePath = "dates.txt";
        try
        {
            using (StreamWriter writer = new StreamWriter (filePath, true))
            {
                writer.WriteLine (date.ToString ("yyyy-MM-dd"));
            }
        }
        catch (Exception ex)
        {
            Console.WriteLine ($"An error occurred while saving the date: {ex.Message}");
        }
    }


    public static string isValidCallsign (string input)
    {
        string pattern = @"\b\d{0,2}[A-Z]{1,2}\d{1,2}[A-Z]{1,6}\b";
        Regex regexCallSign = new Regex (pattern, RegexOptions.IgnoreCase);
        Match match = regexCallSign.Match (input);
        if (match.Success)
        { input = match.Value; }
        else { input = ""; }
        return input;
    }
    public static string isValidName (string input)
    {
        string pattern = @".*\d.*(\r?\n)?";
        input = input.ToUpper ();
        string result = Regex.Replace (input, pattern, "", RegexOptions.Multiline);
        // Regex regexName = new Regex (pattern, RegexOptions.IgnoreCase);
        // Match match = regexName.Match (input);
        // if (match.Success)
        // { input = match.Value; }
        // else { input = ""; }
        return result;
    }
    // public static string isValidCountry (string input)
    // {
    //     string pattern = "AUSTRIA,CANADA,ENGLAND,UK,GERMANY,NORWAY,NEW ZEALAND,PHILIPPINES,ROMANIA,SERBIA,ST LUCIA,TRINIDAD & TOBAGO,VENEZUELA,AFG,ALA,ALB,DZA,ASM,AND,AGO,AIA,ATA,ATG,ARG,ARM,ABW,AUS,AUT,AZE,BHS,BHR,BGD,BRB,BLR,BEL,BLZ,BEN,BMU,BTN,BOL,BIH,BWA,BVT,BRA,IOT,VGB,BRN,BGR,BFA,BDI,KHM,CMR,CAN,CPV,BES,CYM,CAF,TCD,CHL,CHN,CXR,CCK,COL,COM,COK,CRI,HRV,CUB,CUW,CYP,CZE,COD,DNK,DJI,DMA,DOM,TLS,ECU,EGY,SLV,GNQ,ERI,EST,SWZ,ETH,FLK,FRO,FSM,FJI,FIN,FRA,GUF,PYF,ATF,GAB,GMB,GEO,DEU,GHA,GIB,GRC,GRL,GRD,GLP,GUM,GTM,GGY,GIN,GNB,GUY,HTI,HMD,HND,HKG,HUN,ISL,IND,IDN,IRN,IRQ,IRL,IMN,ISR,ITA,CIV,JAM,JPN,JEY,JOR,KAZ,KEN,KIR,XXK,KWT,KGZ,LAO,LVA,LBN,LSO,LBR,LBY,LIE,LTU,LUX,MAC,MDG,MWI,MYS,MDV,MLI,MLT,MHL,MTQ,MRT,MUS,MYT,MEX,MDA,MNG,MNE,MSR,MAR,MOZ,MMR,NAM,NRU,NPL,NLD,NCL,NZL,NIC,NER,NGA,NIU,NFK,PRK,MKD,MNP,NOR,OMN,PAK,PLW,PSE,PAN,PNG,PRY,PER,PHL,PCN,POL,PRT,MCO,PRI,QAT,COG,REU,ROU,RUS,RWA,BLM,SHN,KNA,LCA,MAF,SPM,VCT,WSM,SMR,STP,SAU,SEN,SRB,SYC,SLE,SGP,SXM,SVK,SVN,SLB,SOM,ZAF,SGS,KOR,SSD,ESP,LKA,SDN,SUR,SJM,SWE,CHE,SYR,TWN,TJK,TZA,THA,TGO,TKL,TON,TTO,TUN,TUR,TKM,TCA,TUV,UGA,UKR,ARE,GBR,UMI,USA,URY,UZB,VUT,VAT,VEN,VNM,VIR,WLF,ESH,YEM,ZMB,ZWE,";
    // int found = pattern.IndexOf (input);
    // if (found == -1) input = "";
    // return input;
    // }

    public static string isValidField (string input, string pattern)
    {
        input = input.ToUpper ().Trim ().Trim ('.');
        pattern = pattern.ToUpper ();
        pattern += ",NA,NONE";
        int found = pattern.IndexOf (input);
        if (found == -1) input = "";
        return input;
    }
    static string ExtractMaidenheadGrid (string input)
    {
        // Define the regular expression for Maidenhead grid locator (4 or 6 character grids)
        //Regex regex = new Regex (@"\b([A-R]{2}\d{2}[A-X]{0,2})\b", RegexOptions.IgnoreCase); // 6 char grid
        Regex regex = new Regex (@"\b([A-R]{2}\d{2}[A-X]{0,2}[\dA-X]{0,4})\b", RegexOptions.IgnoreCase);

        // Search for a match in the input string
        Match match = regex.Match (input);

        if (match.Success)
        {
            if (match.Value.Length < 4 || match.Value.Length % 2 != 0)
            {
                Console.Write ("Invalid Maidenhead grid format in " + input);
                return ("");
            }
            return match.Value.ToUpper (); // Return the Maidenhead grid in uppercase
        }

        return string.Empty; // Return an empty string if no match is found
    }

    static (double, double) MaidenheadToGPS (string maidenhead)
    {
        if (maidenhead.Length < 4 || maidenhead.Length % 2 != 0)
        {
            Console.Write ("Invalid Maidenhead grid format in");
            return (0, 0);
        }
        else
        {
            maidenhead = maidenhead.ToUpper ();

            // Calculate the longitude
            int lonField = maidenhead [0] - 'A'; // First letter
            int lonSquare = maidenhead [2] - '0'; // First number
            int lonSubsquare = maidenhead.Length >= 6 ? maidenhead [4] - 'A' : 0; // Optional letter for sub-square

            // Calculate the latitude
            int latField = maidenhead [1] - 'A'; // Second letter
            int latSquare = maidenhead [3] - '0'; // Second number
            int latSubsquare = maidenhead.Length >= 6 ? maidenhead [5] - 'A' : 0; // Optional letter for sub-square

            // Convert Maidenhead to latitude and longitude
            double lon = -180.0 + (lonField * 20.0) + (lonSquare * 2.0) + (lonSubsquare * (2.0 / 24.0)) + (2.0 / 48.0);
            double lat = -90.0 + (latField * 10.0) + (latSquare * 1.0) + (latSubsquare * (1.0 / 24.0)) + (1.0 / 48.0);
            lat = Math.Round (lat, 6);
            lon = Math.Round (lon, 6);
            return (lat, lon);
        }
    }
    static string fillFieldNum (string input)
    {
        // find the first 8 commas and append a field number to each occurence
        var ct = 1;
        var startPosition = 0;
        // Console.WriteLine("input before: "+input);
        input = input + ",";
        while (ct < 10)
        {
            startPosition = input.IndexOf (",", startPosition);
            if (startPosition > 0)
            {
                input = input.Insert (startPosition, " " + ct);
                // Console.WriteLine("input during: "+input);
            }
            else { break; }
            ct++;
            startPosition = startPosition + 3;
        }
        // Console.WriteLine("input after: "+input);
        return input.Trim (',');
    }
    public string Reverse (string text)
    {
        char [] cArray = text.ToCharArray ();
        string reverse = String.Empty;
        for (int i = cArray.Length - 1; i > -1; i--)
        {
            reverse += cArray [i];
        }
        return reverse;
    }
    public static string ScrambleWord (string word)
    {
        char [] chars = new char [word.Length];
        // var now = TimeOnly.FromDateTime(DateTime.Now);
        //Random rand = new Random(10000);
        Random rand = new Random (); // Seed is automatically set to current time
        int index = 0;
        while (word.Length > 0)
        { // Get a random number between 0 and the length of the word. 
            int next = rand.Next (0, word.Length - 1);
            // Take the character from the random position 
            // and add to our char array. 
            chars [index] = word [next];
            // Remove the character from the word. 
            word = word.Substring (0, next) + word.Substring (next + 1);
            ++index;
        }
        return new String (chars);
    }
    public static class Common
    {
        public static double ConvertToDouble (string Value)
        {
            if (Value == null)
            {
                return 0;
            }
            else
            {
                double OutVal;
                double.TryParse (Value, out OutVal);

                if (double.IsNaN (OutVal) || double.IsInfinity (OutVal))
                {
                    return 0;
                }
                return OutVal;
            }
        }
    }
    public static string SortCommaDelimitedString (string input, string delimiter)
    {
        // Split the string into an array of strings
        string [] items = input.Split (delimiter);

        // Sort the array
        Array.Sort (items);

        // Join the sorted array back into a string
        return string.Join (delimiter, items);
    }
    public static void SortStringBuilder (StringBuilder sb, string delimiter, int rowsToSkip)
    {
        // Convert the StringBuilder content to a string
        string content = sb.ToString ();
        content = content.Trim (';').Trim ('|').Trim ().Replace ("\r\n\r\n", "\r\n");
        string header = "";
        int i = 0;
        // Split the string using the delimiter
        string [] items = content.Split (new [] { delimiter }, StringSplitOptions.None);

        // Keep the header (first row)
        if (rowsToSkip > 0)
        {
            while (i < rowsToSkip)
            {
                header = header + items [i] + delimiter;
                i++;
            }

        }

        // Sort the rest of the array (skip the first row)
        string [] rowsToSort = items.Skip (rowsToSkip).ToArray ();
        Array.Sort (rowsToSort);

        // Join the header and sorted rows back into a string
        string sortedContent = header + string.Join (delimiter, rowsToSort);
        sortedContent.Replace ("\r\n\r\n", "\r\n");
        // Clear the original StringBuilder and append the sorted content
        sb.Clear ();
        sb.Append (sortedContent);
    }
    public static string checkBand (string input)
    {
        input = input
            .Replace ("METERS", "M")
            .Replace ("METER", "M")
            .Replace ("TELENET", "TELNET") // common typo
            .Replace (" ", "")
            .Replace ("(", "")
            .Replace (")", "")
            ;
        switch (input)
        {
            case "TELNET":
            case "160M":
            case "80M":
            case "60M":
            case "40M":
            case "30M":
            case "20M":
            case "17M":
            case "15M":
            case "12M":
            case "10M":
            case "6M":
            case "2M":
            case "1.25M":
            case "70CM":
            case "33CM":
            case "23CM":
            case "13CM":
            case "5CM":
            case "3CM":
            case "HF":
            case "SHF":
            case "UHF":
            case "VHF":
                break;

            case "160":
                input = input + "M";
                break;
            case "80":
                input = input + "M";
                break;
            case "60":
                input = input + "M";
                break;
            case "40":
                input = input + "M";
                break;
            case "30":
                input = input + "M";
                break;
            case "20":
                input = input + "M";
                break;
            case "17":
                input = input + "M";
                break;
            case "15":
                input = input + "M";
                break;
            case "12":
                input = input + "M";
                break;
            case "10":
                input = input + "M";
                break;
            case "6":
                input = input + "M";
                break;
            case "2":
                input = input + "M";
                break;
            case "70":
                input = input + "CM";
                break;
            case "33":
                input = input + "CM";
                break;
            case "23":
                input = input + "CM";
                break;
            case "13":
                input = input + "CM";
                break;
            case "5":
                input = input + "cm";
                break;
            case "3":
                input = input + "CM";
                break;
            default:
                // Console.WriteLine("Band is not standard for "+messageID+"  "+checkIn+": "+input+" - "+ msgField+ "\r\n");
                input = "";
                break;
        }
        return input;

    }
    public static string checkMode (string input, string input2)
    {
        input = input
            .Replace ("AREDNMESH", "MESH")
            .Replace ("AREDN", "MESH")
            .Replace ("VERA", "VARA")
            .Replace ("WINLINK", "")
            .Replace ("-", " ")
            .Replace ("(", "")
            .Replace (")", "")
            .Replace (".", "")
            .Trim ();
        switch (input)
        {
            case "SMTP":
            case "TELNET":
            case "PACKET":
            case "ARDOP":
            case "ARDOPC":
            case "ARDOPCF":
            case "VARA FM":
            case "VARA HF":
            case "PACTOR":
            case "INDIUM GO":
            case "MESH":
            case "APRS":
            case "ROBUST PACKET":
            case "JS8CALL":
                break;

            case "PACKET FM":
            case "X.25":
            case "AX.25":
                input = "PACKET";
                break;
            case "VARA":
                if (input2 == "2M" || input2 == "70CM" || input2 == "6M") { input = "VARA FM"; }
                else { input = "VARA HF"; }
                break;

            case "FM":
            case "FM VARA":
            case "VARA-FM":
            case "VARAFM":
                input = "VARA FM";
                break;

            case "VARAHF":
            case "HFVARA":
            case "HF":
            case "HF VARA":
            case "VARA-HF":
                input = "VARA HF";
                break;
            //case "WINLINK EXPRESS":
            //  input = "PACKET";
            //break;

            default:
                // Console.WriteLine("Bad mode: "+messageID+"  "+checkIn+": "+input+" - "+ msgField+ "\r\n");
                input = "";
                break;
        }
        return input;
    }
    public static string removeFieldNumber (string input)
    {
        // Split the string into an array of strings
        string [] items = input.Split (",");
        // remove the field number from each
        int len = items.Length;
        int i = 0;
        string pattern = @"\s\d$";
        foreach (string item in items)
        {
            items [i] = Regex.Replace (item, pattern, "");
            i++;
            if (i > 9) { break; }
        }

        // Join the sorted array back into a string
        string result = string.Join (",", items);
        return result;
    }
    static void RemoveLineContaining (StringBuilder sb, string callSign)
    {
        // Convert the StringBuilder content to a string array of lines
        string [] lines = sb.ToString ().Split (new [] { "\r\n", "\n" }, StringSplitOptions.None);

        // Filter out lines that contain the callSign
        string [] filteredLines = lines.Where (line => !line.Contains (callSign)).ToArray ();

        // Clear the original StringBuilder and append filtered lines
        sb.Clear ();
        sb.Append (string.Join (Environment.NewLine, filteredLines));
    }
    static bool ExtractCoordinates (string input, out double latitude, out double longitude)
    {
        // Initialize output variables
        latitude = 0;
        longitude = 0;

        // Define the regular expression for GPS latitude and longitude (with optional N/S/E/W directions)
        Regex regex = new Regex (@"([-+]?([0-8]?\d(\.\d+)?|90(\.0+)?))\s*[°]?\s*([NS]),?\s*([-+]?((1[0-7]\d|\d{1,2})(\.\d+)?|180(\.0+)?))\s*[°]?\s*([EW])");

        // Search for the latitude and longitude pattern in the input string
        Match match = regex.Match (input);

        if (match.Success)
        {
            // Extract the numeric part of latitude
            latitude = Math.Round (double.Parse (match.Groups [1].Value), 6);
            // If it's south (S), negate the latitude
            if (match.Groups [2].Value.ToUpper () == "S")
                latitude = -latitude;
            // Extract the numeric part of longitude
            longitude = Math.Round (double.Parse (match.Groups [7].Value), 6);
            // If it's west (W), negate the longitude
            if (match.Groups [11].Value.ToUpper () == "W")
                longitude = -longitude;

            return true;
        }
        // Return false if latitude and longitude are not found
        return false;
    }
    public static string ConvertDegreeAngleToDecimal (string input)
    {
        string resultStr = "";
        // check to make sure it is a degree format
        Regex regex = new Regex (@"\d{1,3}\-\d{1,3}\.\d+");
        Match match = regex.Match (input);
        if (match.Success)
        {
            char [] charsToTrim = { '.', '-', ' ' };
            input = input.Replace ("-", ".")
                .Replace ("'", ".")
                .Replace ("\"", "")
                .Replace (" ", "")
                .Trim (charsToTrim);
            if (input.Length > 11) input = input.Substring (input.Length - 11);
            var multiplier = (input.Contains ("S") || input.Contains ("W")) ? -1 : 1; //handle south and west
            input = Regex.Replace (input, "[^0-9.]", ""); //remove the characters

            string [] inputItems = input.Split (".");

            double degrees = Convert.ToDouble (inputItems [0]);
            double minutes = Convert.ToDouble (inputItems [1]);
            double seconds = Convert.ToDouble (inputItems [2]);
            var _deg = (double)Math.Abs (degrees);
            var result = ((_deg + (minutes / 60) + (seconds / 3600)) * multiplier);
            resultStr = result.ToString ("0.######");
        }
        return resultStr;
    }

}

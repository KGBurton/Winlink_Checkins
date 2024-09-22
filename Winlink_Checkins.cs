//c# code that will input start date, end date, and callSign and will select files with an extension of mime from the current folder  based on start date and end date, and will read each file to find a line labeled To: . If the rest of the line contains callSign, then write the data from the line labeled X-Source: to a text file called checkins.txt in the same folder
// Design Get the date range
// get the data source
// is it a message file? (.mime)
// is it within the date range
// is it related to a unique net checkin name (make all strings are upper case or ignore case before checking)
// is it a bounced message? record for review, do not count
// is it a forwarded message? be sure to get the correc callsign
// is the callsign in the MSG field? if not try the from field, if not record and don't count
//      REGEX pattern to find the FCC callsign is \b[A-Z]{1,2}\d[A-Z]{1,3}\b for us only
//      \b\d{0,2}[A-Z]{1,2}\d{1,2}[A-Z]{1,6}\b for almost all international and US
//      see https://regex101.com/r/gS6qG8/1 for a regex tester/editor. This is string to test in 
//      the editor: Ken, AE0MW, this, is, a test. 4F1PUZ 4G1CCH AX3GAMES 4D71 4X130RISHON 9N38 AX3GAMES BV100 DA2MORSE  DB50FIRAC DL50FRANCE FBC5AGB FBC5CWU FBC5LMJ FBC5NOD FBC5YJ FBC6HQP GB50RSARS HA80MRASZ  HB9STEVE HG5FIRAC HG80MRASZ II050SCOUT IP1METEO J42004A J42004Q LM1814 LM2T70Y LM9L40Y LM9L40Y/P OEM2BZL OEM3SGU OEM3SGU/3 OEM6CLD OEM8CIQ OM2011GOOOLY ON1000NOTGER ON70REDSTAR PA09SHAPE PA65VERON PA90CORUS PG50RNARS PG540BUFFALO S55CERKNO TM380 TX9 TYA11 U5ARTEK/A V6T1 VI2AJ2010 VI2FG30 VI4WIP50 VU3DJQF1 VX31763 WD4 XUF2B YI9B4E YO1000LEANY ZL4RUGBY ZS9MADIBA
// is it in the current roster? if not record with new checkins, save, count.
//     Requires that roster.txt exist in the application folder. 
// is it a duplicate? if yes, don't save or count (spreadsheet can handle duplicates)
// what template was used? necessary to get the start and end positions correct
// save document info, writeLines, and counts to checkins.txt;
// write callsign and message to checkins.csv file

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


class Winlink_Checkins
{
    public static void Main (string [] args)
    {
        // Get the start date and end date from the user.
        DateTime startDate = DateTime.Today;
        DateTime endDate = DateTime.Today;
        string utcDate = DateTime.UtcNow.ToString ("yyyy/MM/dd HH:mm:ss");
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
        string xmlFile = applicationFolder + "\\Winlink_Import.xml";
        // string commentFile = applicationFolder+"\\GLAWN_Additional_Comments.txt";
        // writeString variables to go in the output files
        StringBuilder netCheckinString = new StringBuilder ();
        StringBuilder netAckString2 = new StringBuilder ();
        StringBuilder bouncedString = new StringBuilder ();
        StringBuilder duplicates = new StringBuilder ();
        StringBuilder newCheckIns = new StringBuilder ();
        StringBuilder csvString = new StringBuilder ();
        csvString.AppendLine ("Current GLAWN Checkins, posted: " + DateTime.Now.ToString ("yyyymmdd HH:mm:ss"));
        StringBuilder mapString = new StringBuilder ();
        mapString.Append ("CallSign,Latitude,Longitude,Band,Mode\r\n");
        StringBuilder badBandString = new StringBuilder ();
        StringBuilder badModeString = new StringBuilder ();
        StringBuilder skippedString = new StringBuilder ();
        // skippedString.AppendLine("Messages Skipped: ");
        StringBuilder removalString = new StringBuilder ();
        // removalString.AppendLine("Removal Requests: ");
        StringBuilder addonString = new StringBuilder ();
        StringBuilder noGPSString = new StringBuilder ();

        string callSignPattern = @"\b\d{0,2}[A-Z]{1,2}\d{1,2}[A-Z]{1,6}\b";
        string testString = "";
        string rosterString = "";
        string bandStr = "";
        string modeStr = "";
        string noGPSStr = "";
        string checkIn = "";
        string msgFieldNumbered = "";
        string latitudeStr = "";
        string longitudeStr = "";
        string xmlXsource = "KB7WHO";
        string delimiter = "";
        addonString.AppendLine ("Comments from the Current Checkins\r\n-------------------------------");
        noGPSString.AppendLine ("\r\n++++++++\r\nThese had neither GPS data nor Maidenhead Grids\r\n-------------------------");
        Random rnd = new Random ();

        var startPosition = 0;
        var endPosition = 0;
        var len = 0;
        var msgTotal = 0;
        var skipped = 0;
        int ct = 0;
        int dupCt = 0;
        var newCt = 0;
        var outOfRangeCt = 0;
        var removalCt = 0;
        var ackCt = 0;
        var localWeatherCt = 0;
        var severeWeatherCt = 0;
        var incidentStatusCt = 0;
        var icsCt = 0;
        var winlinkCkinCt = 0;
        var damAssessCt = 0;
        var fieldSitCt = 0;
        var quickHWCt = 0;
        var qwmCt = 0;
        var miCt = 0;
        var dyfiCt = 0;
        var rriCt = 0;
        var junk = 0;
        var mapCt = 0;
        var bandCt = 0;
        var modeCt = 0;
        var aprsCt = 0;
        var meshCt = 0;
        var noGPSCt = 0;
        var noGPSFlag = 0;
        var badBandCt = 0;
        var badModeCt = 0;
        var dupeFlag = 0;
        int rowsToSkip = 0;
        string locType = "";
        string xSource = "";
        double latitude = 0;
        double longitude = 0;

        TextInfo textInfo = new CultureInfo ("en-US", false).TextInfo;
        // Create root XML document
        XDocument xmlDoc = new XDocument (new XElement ("WinlinkMessages"));
        XElement messageElement = new XElement
            ("export_parameters",
                new XElement ("xml_file_version", "1.0"),
                new XElement ("winlink_express_version", "1.7.17.0"),
                // for testing
                // new XElement("callsign", "KB7WHO")
                new XElement ("callsign", "GLAWN")
            );
        xmlDoc.Root.Add (messageElement);

        messageElement = new XElement ("message_list", "");
        xmlDoc.Root.Add (messageElement);



        if (File.Exists (rosterFile))
        {
            rosterString = File.ReadAllText (rosterFile);
            rosterString = rosterString.ToUpper ();
            //debug Console.WriteLine("rosterFile contents: "+rosterString);
            // get the net name from the roster.txt file
            startPosition = rosterString.IndexOf ("NETNAME=");
            if (startPosition > -1) { startPosition += 8; }
            endPosition = rosterString.IndexOf ("\r\n", startPosition);
            len = endPosition - startPosition;
            if (len > 0)
            { netName = rosterString.Substring (startPosition, len); }
            else { netName = "GLAWN"; }

            // get the x-source name from the roster.txt file to be used as the netName variable in the xml file
            startPosition = rosterString.IndexOf ("CALLSIGN=");
            if (startPosition > -1) { startPosition += 9; }
            endPosition = rosterString.IndexOf ("\r\n", startPosition);
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
            Console.WriteLine (currentFolder + "\\" + rosterFile + " \n was not found!, all checkins will appear to be new.\n");
        }


        // Select files with an extension of mime from the current folder.
        var files = Directory.GetFiles (currentFolder, "*.mime")
            .Where (file =>
            {
                DateTime fileDate = File.GetLastWriteTime (file);
                // debug Console.Write(fileDate+"\n");
                return fileDate >= startDate && fileDate <= endDate.AddDays (1);
            });

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
                    //debug Console.Write("File "+file+"\n");
                    string fileText = reader.ReadToEnd ();
                    fileText = fileText.ToUpper ()
                        .Replace ("=\r\n", "")
                        .Replace ("=20", "");

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

                    // find the end of the header section
                    var endHeader = fileText.IndexOf ("CONTENT-TRANSFER-ENCODING:");

                    // was it forwarded?
                    var forwarded = fileText.IndexOf ("WAS FORWARDED BY");

                    // was it APRSmail?
                    var APRS = fileText.IndexOf ("APRSEMAIL2");

                    // check for acknowledgement message and discard                  
                    var ack = fileText.IndexOf ("ACKNOWLEDGEMENT");

                    // check for removal message and discard                  
                    var removal = fileText.IndexOf ("REMOV");

                    // look to see if it was a bounced message
                    var bounced = fileText.IndexOf ("UNDELIVERABLE");

                    // check for local weather report
                    var localWeather = fileText.IndexOf ("CURRENT LOCAL WEATHER CONDITIONS");

                    // check for severe weather
                    var severeWeather = fileText.IndexOf ("SEVERE WX REPORT");

                    // check incident status report
                    var incidentStatus = fileText.IndexOf ("INCIDENT STATUS");

                    // check for ICS 213 msg
                    var ics = fileText.IndexOf ("TEMPLATE VERSION: ICS 213");

                    // check for winlink checkin message
                    var winlinkCkin = fileText.IndexOf ("MAP FILE NAME: WINLINK CHECK", endHeader);
                    // some people include WINLINK CHECK-IN in the subject which confuses the program
                    // into thinking this is a winlink checkin FORM!! Catch it ...
                    if (winlinkCkin < 0)
                    {
                        winlinkCkin = fileText.IndexOf ("WINLINK CHECK-IN 5.0", endHeader);
                        if (winlinkCkin < 0)
                        {
                            winlinkCkin = fileText.IndexOf ("WINLINK CHECK-IN\r\n0. HEADER", endHeader);
                        }
                    }

                    // check for odd checkin message - don't let it scan through to a binary attachment!
                    var lenBPQ = fileText.Length - 10;
                    if (lenBPQ > 800) { lenBPQ = 800; }
                    var BPQ = fileText.IndexOf ("BPQ", 1, lenBPQ);

                    // check for damage assessment report
                    var damAssess = fileText.IndexOf ("SURVEY REPORT - CATEGORIES");

                    // check for field situation report
                    var fieldSit = fileText.IndexOf ("EMERGENT/LIFE SAFETY");

                    // check for Quick Health & Welfare report
                    var quickHW = fileText.IndexOf ("QUICK H&W");

                    // check for RRI Welfare Radiogram
                    var rriWR = fileText.IndexOf ("TEMPLATE VERSION: RRI WELFARE RADIOGRAM");

                    // check for Did You Feel It report
                    var dyfi = fileText.IndexOf ("DYFI WINLINK");

                    // check for RRI Welfare Radiogram
                    var qwm = fileText.IndexOf ("TEMPLATE VERSION: QUICK WELFARE MESSAGE");

                    // check for Medical Incident Report
                    var mi = fileText.IndexOf ("INITIAL PATIENT ASSESSMENT");

                    // screen dates to eliminate file dates that are different from the sent date and fall outside the net span
                    int startDateCompare = DateTime.Compare (sentDateUni, startDate);
                    int endDateCompare = DateTime.Compare (sentDateUni, endDate);

                    // discard acknowledgements
                    if (ack > 0)
                    {
                        skipped++;
                        ackCt++;
                        junk = 0; //debug Console.Write(file+" is an acknowedgement, skipping.");
                    }

                    else if (startDateCompare < 0 || endDateCompare > 0)
                    {
                        skipped++;
                        outOfRangeCt++;
                        Console.Write (messageID + " sendDate fell outside the start/end dates\r\n");
                        skippedString.Append ("\tOut of date range: " + messageID + "\r\n");
                    }

                    else if (removal > 0)
                    {
                        startPosition = fileText.IndexOf ("FROM:");
                        if (startPosition > -1) { startPosition += 6; }
                        endPosition = fileText.IndexOf ("\r\n", startPosition);
                        len = endPosition - startPosition;
                        checkIn = fileText.Substring (startPosition, len);
                        {
                            checkIn = checkIn.Replace (',', ' ');
                            // Create a Regex object with the pattern
                            Regex regexCallSign = new Regex (callSignPattern, RegexOptions.IgnoreCase);
                            // find the first callsign match in the checkIn string
                            Match match = regexCallSign.Match (checkIn);
                            if (match.Success) checkIn = match.Value;
                        }
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
                            Regex regexCallSign = new Regex (callSignPattern, RegexOptions.IgnoreCase);
                            // find the first callsign match in the checkIn string
                            Match match = regexCallSign.Match (checkIn);
                            if (match.Success) checkIn = match.Value;
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

                        // deterimine if it was forwarded to know to look below the first header info

                        if (subjText.Contains (netName))
                        {

                            // get x-Source if available XXXX
                            var xSrc = fileText.IndexOf ("X-SOURCE: ");
                            if (xSrc > -1)
                            {
                                startPosition = xSrc + 10;
                                endPosition = fileText.IndexOf ("\r\n", startPosition);
                                len = endPosition - startPosition;
                                if (len > 0) { xSource = fileText.Substring (startPosition, len); }
                            }

                            // skip APRS header 
                            if (APRS > 0)
                            {
                                startPosition = fileText.IndexOf ("FROM:", APRS);
                                if (startPosition > -1)
                                {
                                    startPosition = fileText.IndexOf ("\r\n", startPosition);
                                    if (startPosition > -1) { startPosition += 2; }
                                    endPosition = fileText.IndexOf ("DO NOT REPLY", startPosition) - 1;
                                }
                                aprsCt++;
                            }

                            // adjust for ICS 213
                            else if (ics > 0)
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
                            else if (winlinkCkin > 0)
                            {
                                // the winlink check-in form changed format between 5.0.10 and 5.0.5 so check for that
                                var winlinkCkinOffset = fileText.IndexOf ("WINLINK CHECK-IN 5.0.5");
                                // if (winlinkCkinOffset > 0) { winlinkCkinOffset = 9; } else { winlinkCkinOffset = 13; }
                                // startPosition = fileText.IndexOf("COMMENTS:")+ winlinkCkinOffset;
                                startPosition = fileText.IndexOf ("COMMENTS:");
                                if (startPosition > -1) { startPosition += 9; }
                                endPosition = fileText.IndexOf ("----------", startPosition) - 1;
                            }

                            // adjust for odd message that insert an R: line at the top
                            else if (BPQ > 0)
                            {
                                startPosition = fileText.IndexOf ("BPQ", 1, lenBPQ);
                                if (startPosition > -1) { startPosition += 12; }
                                endPosition = fileText.IndexOf ("--BOUNDARY", startPosition) - 2;
                            }
                            else if (localWeather > 0)
                            {
                                startPosition = fileText.IndexOf ("NOTES:");
                                if (startPosition > -1) { startPosition += 11; }
                                endPosition = fileText.IndexOf ("----------", startPosition) - 1;
                            }

                            else if (severeWeather > 0)
                            {
                                startPosition = fileText.IndexOf ("COMMENTS:");
                                if (startPosition > -1) { startPosition += 10; }
                                endPosition = fileText.IndexOf ("----------", startPosition) - 1;
                            }

                            else if (incidentStatus > 0)
                            {
                                startPosition = fileText.IndexOf ("REPORT SUBMITTED BY:");
                                if (startPosition > -1) { startPosition += 20; }
                                endPosition = fileText.IndexOf ("----------", startPosition) - 1;
                            }

                            else if (damAssess > 0)
                            {
                                startPosition = fileText.IndexOf ("COMMENTS:");
                                if (startPosition > -1) { startPosition += 21; }
                                endPosition = fileText.IndexOf ("----------", startPosition) - 1;
                            }

                            else if (fieldSit > 0)
                            {
                                startPosition = fileText.IndexOf ("COMMENTS:");
                                if (startPosition > -1) { startPosition += 11; }
                                endPosition = fileText.IndexOf ("\r\n", startPosition);
                            }

                            else if (dyfi > 0)
                            {
                                startPosition = fileText.IndexOf ("COMMENTS");
                                if (startPosition > -1) { startPosition += 11; }
                                endPosition = fileText.IndexOf ("\r\n", startPosition) - 1;
                            }

                            else if (rriWR > 0)
                            {
                                startPosition = fileText.IndexOf ("BT\r\n");
                                if (startPosition > -1) { startPosition += 3; }
                                endPosition = fileText.IndexOf ("------", startPosition) - 1;
                            }

                            else if (qwm > 0)
                            {
                                startPosition = fileText.IndexOf ("IT WAS SENT FROM:");
                                endPosition = fileText.IndexOf ("------", startPosition) - 1;
                            }
                            else if (mi > 0)
                            {
                                startPosition = fileText.IndexOf ("ADDITIONAL INFORMATION");
                                startPosition = fileText.IndexOf ("\r\n", startPosition);
                                endPosition = fileText.IndexOf ("----", startPosition) - 1;
                            }
                            else
                            {
                                // end of the header information as the start of the msg field
                                if (forwarded <= 0)
                                {
                                    startPosition = fileText.IndexOf ("QUOTED-PRINTABLE");
                                    if (startPosition > -1) { startPosition += 20; }
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
                                Console.Write ("Nothing in the message field: " + file + "\n");
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
                                len = endPosition - startPosition;
                            }

                            if (len < 0)
                            {
                                Console.Write ("endPostion is less than startPosition in: " + file + "\n");
                                Console.Write ("Break at line 550ish. Press enter to close.");
                                input = Console.ReadLine ();
                                break;
                            }

                            string msgField = fileText.Substring (startPosition, len);
                            msgField = msgField
                                .Replace ("=20", "")
                                .Replace ("=0A", "")
                                .Replace ("=0", "")
                                .Replace ("16. CONTACT INFO:", ",")
                                .Trim ()
                                .Replace ("  ", " ")
                                .Replace ("  ", " ")
                                //.Replace(".", "") this causes problems with decimal band freq
                                .Replace (", ", ",")
                                .Replace ("[NO CHANGES OR EDITING OF THIS MESSAGE ARE ALLOWED]", "")
                                .Replace ("[MESSAGE RECEIPT REQUESTED]", "")
                                .Replace (" ,", ",")
                                .Replace ("\"", "")
                                .Trim ()
                                //.Trim(',')
                                + ",";
                            checkIn = msgField
                                //.Replace(" ,", ",")    
                                //.Trim()
                                //.Trim(',')
                                //.Trim()+",";
                                ;

                            // Create a Regex object with the pattern
                            Regex regexCallSign = new Regex (callSignPattern, RegexOptions.IgnoreCase);

                            // find the first callsign match in the checkIn string
                            Match match = regexCallSign.Match (checkIn);
                            if (match.Success)
                            {
                                checkIn = match.Value;
                                if (checkIn == "KB7WHO") { checkIn = xSource; }
                                if (xSource == "") { xSource = checkIn; }
                            }
                            else
                            {
                                // try the from field since the callsign could not be located in the msg field
                                startPosition = fileText.IndexOf ("FROM:");
                                if (startPosition > -1) { startPosition += 6; }
                                endPosition = fileText.IndexOf ("@", startPosition);
                                if (endPosition < 0) { endPosition = fileText.IndexOf ("SUBJECT:") - 1; }
                                len = endPosition - startPosition;
                                if (len > 0)
                                {
                                    checkIn = fileText.Substring (startPosition, len);
                                    // Create a Regex object with the pattern
                                    regexCallSign = new Regex (callSignPattern, RegexOptions.IgnoreCase);
                                    match = regexCallSign.Match (checkIn);
                                    if (match.Success)
                                    {
                                        checkIn = match.Value;
                                    }
                                    else
                                    {
                                        checkIn = "";
                                    }
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
                                startPosition = testString.IndexOf (checkIn);
                                if (startPosition >= 0)
                                {
                                    if (dupCt == 0) { duplicates.Append ("Duplicates: \r\n"); }
                                    //debug Console.Write("netName "+checkIn+" is a duplicate, skipping. It is "+dupCt+" of "+msgTotal+" total messages.\n");
                                    duplicates.Append (checkIn + ", ");
                                    dupeFlag = 1;
                                    dupCt++;
                                }
                                //else
                                //{
                                ct++;
                                if (localWeather > 0) { localWeatherCt++; }
                                if (severeWeather > 0) { severeWeatherCt++; }
                                if (winlinkCkin > 0) { winlinkCkinCt++; }
                                if (incidentStatus > 0) { incidentStatusCt++; }
                                if (ics > 0) { icsCt++; }
                                if (damAssess > 0) { damAssessCt++; }
                                if (fieldSit > 0) { fieldSitCt++; }
                                if (quickHW > 0) { quickHWCt++; }
                                if (dyfi > 0) { dyfiCt++; }
                                if (rriWR > 0) { rriCt++; }
                                if (qwm > 0) { qwmCt++; }
                                if (mi > 0) { miCt++; }
                                testString = testString + checkIn + " | ";
                                // the spreadsheet chokes if the string ends with "|" so
                                // don't let that happen by writing the first one without a delimiter
                                // prepending the delimiter to the rest.
                                if (ct == 1)
                                {
                                    netCheckinString.Append (checkIn);
                                }
                                else if (ct > 1)
                                {
                                    netCheckinString.Append ("|" + checkIn);
                                }
                                netAckString2.Append (checkIn + ";");
                                // find message, format for csv file, and save
                                var msgFieldStart = msgField.IndexOf ("\r\n");
                                string notFirstLine = "";
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
                                // Extract latitude and longitude
                                // Winlink Checkin has its own tags so check them first

                                latitudeStr = "";
                                longitudeStr = "";
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

                                    startPosition = fileText.IndexOf ("MESSAGE-ID:");
                                    startPosition = fileText.IndexOf ("\r\n", startPosition);
                                    if (startPosition > -1) { startPosition += 2; }
                                    len = fileText.Length - startPosition;
                                    if (len > 0)
                                    {
                                        if (ExtractCoordinates (fileText.Substring (startPosition), out latitude, out longitude))
                                        {
                                            // Console.WriteLine(messageID+" latitude: "+latitude+" longitude: "+longitude);                                
                                        }
                                        else
                                        {
                                            // no valid GPS coordinates found, look for a maidenhead grid
                                            string maidenheadGrid = ExtractMaidenheadGrid (fileText);
                                            if (!string.IsNullOrEmpty (maidenheadGrid))
                                            {
                                                // Console.WriteLine($"Maidenhead Grid: {maidenheadGrid}");
                                                // Convert Maidenhead to GPS coordinates
                                                (latitude, longitude) = MaidenheadToGPS (maidenheadGrid);
                                                // Console.WriteLine($"No GPS coords found, using Maidenhead Grid: {maidenheadGrid}"+$". From Maidenhead Grid Latitude: {latitude}"+$"  Longitude: {longitude}");
                                            }
                                            else
                                            {
                                                // No valid Maidenhead grid found either, make up something in the middle of the Atlantic
                                                double locChange = Math.Round (rnd.NextDouble () * 10, 6);
                                                latitude = Math.Round ((27.187512 + locChange), 6);
                                                longitude = Math.Round ((-60.144742 + locChange), 6);
                                                // Console.WriteLine("No valid grid and no GPS coordinates found in: "+messageID+" latitude set to: "+latitude+" longitude set to: "+longitude);
                                                noGPSCt++;
                                                noGPSFlag++;
                                                msgField = msgField + ",No Location Data Found in message";
                                                noGPSString.Append ("\t" + messageID + "- - " + checkIn + ": latitude set to: " + latitude + " longitude set to: " + longitude + "\r\n");
                                            }
                                        }
                                    }
                                }
                                msgField = msgField.Replace ("\r\n", ",");
                                msgField = removeFieldNumber (msgField);
                                msgFieldNumbered = fillFieldNum (msgField);
                                csvString.Append (xSource + ":" + messageID + "," + latitude + "," + longitude + "," + locType + "," + msgField + "\r\n");

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
                                        bandStr = "VARA FM";
                                    }
                                    if (msgField.IndexOf ("VARA HF") > -1)
                                    {
                                        modeStr = "VARA HF";
                                        bandStr = "VARA HF";
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
                                if (bandStr.IndexOf ("VARAFM") > -1)
                                {
                                    modeStr = "VARA FM";
                                    if (msgField.IndexOf ("2M") > -1) { bandStr = "2M"; }
                                    if (msgField.IndexOf ("70CM") > -1) { bandStr = "70CM"; }
                                    if (msgField.IndexOf ("VHF") > -1) { bandStr = "VHF"; }
                                    if (msgField.IndexOf ("UHF") > -1) { bandStr = "UHF"; }
                                }
                                if (bandStr.IndexOf ("VARAHF") > -1) { bandStr = "HF"; modeStr = "VARA HF"; }
                                // if (bandStr != "") { bandCt++; }

                                bandStr = checkBand (bandStr);
                                if (bandStr == "")
                                {
                                    // if both the band and the mode have invalid data, try scraping through the msgField
                                    if (msgField.IndexOf ("2M") > -1) { bandStr = "2M"; }
                                    if (msgField.IndexOf ("70CM") > -1) { bandStr = "70CM"; }
                                    if (msgField.IndexOf ("20M") > -1) { bandStr = "20M"; }
                                    if (msgField.IndexOf ("40M") > -1) { bandStr = "40M"; }
                                    if (msgField.IndexOf ("80M") > -1) { bandStr = "80M"; }
                                    if (msgField.IndexOf ("VHF") > -1) { bandStr = "VHF"; }
                                    if (msgField.IndexOf ("UHF") > -1) { bandStr = "UHF"; }
                                    if (bandStr == "")
                                    {
                                        msgFieldNumbered = msgField;
                                        msgFieldNumbered = fillFieldNum (msgFieldNumbered);
                                        badBandString.Append ("\tBad band: " + messageID + " - " + checkIn + ": _" + bandStr + "_  |  " + msgFieldNumbered + "\r\n");
                                        badBandCt++;
                                    }
                                    else { bandCt++; }
                                } 

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
                                    .Replace ("AX.25", "PACKET")
                                    .Replace ("WINLINK", "")
                                    .Replace ("(", "")
                                    .Replace ("ARDOP HF", "ARDOP")
                                    .Replace ("VARA VHF", "VARA FM")
                                    .Replace ("VHF VARA", "VARA FM")
                                    .Replace ("VARAFM", "VARA FM")
                                    .Replace ("VERA", "VARA")
                                    .Replace ("HF ARDOP", "ARDOP")
                                    .Replace (")", "")
                                    .Replace ("-", " ")
                                    .Replace ("=20", "")
                                    .Replace ("VHF PACKET", "PACKET")
                                    .Replace ("TELNET", "SMTP")
                                    .Trim ();
                                if (modeStr.IndexOf ("PACKET") > -1) { modeStr = "PACKET"; }
                                if (modeStr.IndexOf ("MESH") > -1) { modeStr = "MESH"; }
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
                                    if (msgField.IndexOf("VARA FM") > -1) { modeStr = "VARA FM"; }
                                    if (msgField.IndexOf ("VARA HF") > -1) { modeStr = "VARA HF"; }
                                    if (msgField.IndexOf ("PACKET") > -1) { modeStr = "PACKET"; }
                                    if (modeStr == "")
                                    {
                                        msgFieldNumbered = msgField;
                                        msgFieldNumbered = fillFieldNum (msgFieldNumbered);
                                        badBandString.Append ("\tBad mode: " + messageID + " - " + checkIn + ": " + modeStr + " -  |  " + msgFieldNumbered + "\r\n");
                                        badModeCt++;
                                    }
                                    else { modeCt++; }
                                } 

                                // debug Console.Write("modeStr final=|"+modeStr+"|  \r\n");


                                // add to mapString csv file if xloc was found
                                if (latitude != 0)
                                {
                                    if (dupeFlag == 0)
                                    {
                                        mapString.Append (xSource + "," + latitude + "," + longitude + "," + bandStr + "," + modeStr + "\r\n");
                                        mapCt++;
                                    }
                                }

                                // xml data
                                var reminderTxt = "";
                                if (noGPSFlag > 0 || bandStr == "" || modeStr == "")
                                {
                                    reminderTxt = "\r\nRecommended format reminder: callSign, firstname, city, county, STate, country, band, Mode, grid\r\n" +
                                        "Example: xxNxxx, Greg, Sugar City, Madison, ID, USA, 70cm, VARA FM, DN43du\r\n" +
                                        "Example 2: DxNxx,Mario,TONDO,MANILA,NCR,PHILIPPINES,2M,VARA FM,PK04LO\r\n" +
                                        "Example 2: xxNxx,Andre,Burnaby,,BC,Canada,2M,VARA FM,CN89ud\r\n";
                                }
                                else { reminderTxt = "\r\nPerfect Message!\r\n"; }
                                noGPSFlag = 0;
                                // the old message ID will destroy stuff in winlink if it is the same when trying to post
                                // create a new message ID by rearranging the old one
                                string newMessageID = messageID;
                                newMessageID = ScrambleWord (newMessageID);
                                // Console.WriteLine("before: "+messageID+   "    after: "+newMessageID);
                                XElement message_list = xmlDoc.Descendants ("message_list").FirstOrDefault ();
                                message_list.Add (new XElement ("message",
                                    new XElement ("id", newMessageID),
                                    new XElement ("foldertype", "Fixed"),
                                    new XElement ("folder", "Outbox"),
                                    new XElement ("subject", "GLAWN acknowledgement ", DateTime.UtcNow.ToString ("yyyy-MM-dd")),
                                    new XElement ("time", utcDate),
                                    new XElement ("sender", "GLAWN"),
                                    // for testing
                                    // new XElement("sender", "KB7WHO"),
                                    new XElement ("To", xSource),
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
                                        "To: " + checkIn + "\r\n" +
                                        "Message-ID: " + newMessageID + "\r\n" +
                                        // Can't edit if not from my call sign
                                        // "X-Source: GLAWN\r\n"+
                                        // for testing
                                        "X-Source:" + xmlXsource + "\r\n" +
                                        "X-Location: 43.845831N, 111.745744W(GPS) \r\n" +
                                        "MIME-Version: 1.0\r\n" +
                                        "MIME-Version: 1.0\r\n\r\n" +
                                        "Thank you for checking in to the GLAWN. This is a copy of your message (with numbered fields) and extracted data. \r\n" +
                                        "Message: " + msgFieldNumbered + "\r\n" +
                                        reminderTxt + "\r\n" +
                                        "Extracted Data:\r\n" +
                                            "   Latitude: " + latitude + "\r\n" +
                                            "   Longitude: " + longitude + "\r\n" +
                                            "   Band: " + bandStr + "\r\n" +
                                            "   Mode: " + modeStr + "\r\n" +
                                            "\r\nGLAWN Current Map: https://tinyurl.com/GLAWN-Map\r\n" +
                                            "Comments: https://tinyurl.com/GLAWN-comments\r\n" +
                                            "GLAWN Checkins Report: https://tinyurl.com/Checkins-Report\r\n" +
                                            "checkins.csv: https://tinyurl.com/GLAWN-CSV-checkins\r\n" +
                                            "mapfile.csv: https://tinyurl.com/Current-CSV-mapfile\r\n"
                                    )
                                ));

                                // Add the message message_list
                                xmlDoc.Root.Add (messageElement);

                                junk = 0; // just so i could put a debug here
                                dupeFlag = 0; // reset the duplicate flag
                            }
                            var tempCt = ct + dupCt + ackCt + removalCt;
                            //debug Console.Write("checkins:"+ct+"  duplicates:" + dupCt+"  removals:"+removalCt+"  acks:"+ackCt + "  combined:"+tempCt+"   actual total:"+msgTotal+"\n");
                            // missing from roster section. Check to see if the checkin is in the roster. 
                            startPosition = rosterString.IndexOf (checkIn);
                            if (startPosition < 0)
                            {
                                if (newCt == 0)
                                {
                                    newCheckIns.Append ("New Checkins:\r\n");
                                }
                                // debug
                                Console.Write (checkIn + " was not found in roster.txt. \n");
                                newCheckIns.Append (checkIn + ", ");
                                // update roster.txt to contain the new checkin
                                File.AppendAllText ("roster.txt", "; " + checkIn);
                                newCt++;
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
            logWrite.WriteLine ("Current GLAWN Checkins posted: " + DateTime.Now.ToString ("yyyymmdd HH:mm:ss"));

            logWrite.WriteLine ("    Total Stations Checking in:" + (ct - dupCt) + "    Duplicates:" + dupCt + "    Total Checkins:" + ct + "    Removal Requests: " + removalCt);
            logWrite.WriteLine ("Non-" + netName + " checkin messages skipped: " + skipped + "(including " + ackCt + " acknowledgements and " + outOfRangeCt + " out of date range messages skipped.)\r\n");
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

            SortStringBuilder (mapString, "\r\n", 1);
            // Console.WriteLine(mapString);
            mapWrite.WriteLine (mapString);

            SortStringBuilder (addonString, "\r\n", 2);
            commentWrite.WriteLine (addonString);

            xmlDoc.Save (xmlFile);

            if (duplicates.Length != 0) { logWrite.WriteLine (duplicates + "\r\n"); }
            if (bouncedString.Length != 0) { logWrite.WriteLine ("Messages that bounced: " + bouncedString); }
            if (newCheckIns.Length != 0) { logWrite.WriteLine ("New Checkins: " + newCheckIns); }
            if (skippedString.Length != 0) { logWrite.WriteLine ("Messages Skipped: \r\n" + skippedString); }
            if (removalString.Length != 0) { logWrite.WriteLine ("Requests to be Removed: " + removalString); }
            if (localWeatherCt > 0) { logWrite.WriteLine ("Local Weather Checkins: " + localWeatherCt); }
            if (severeWeatherCt > 0) { logWrite.WriteLine ("Severe Weather Checkins: " + severeWeatherCt); }
            if (incidentStatusCt > 0) { logWrite.WriteLine ("Incident Status Checkins: " + incidentStatusCt); }
            if (icsCt > 0) { logWrite.WriteLine ("ICS-213 Checkins: " + icsCt); }
            if (winlinkCkinCt > 0) { logWrite.WriteLine ("Winlink Check-in Checkins: " + winlinkCkinCt); }
            if (damAssessCt > 0) { logWrite.WriteLine ("Damage Assessment Checkins: " + damAssessCt); }
            if (fieldSitCt > 0) { logWrite.WriteLine ("Field Situation Report Checkins: " + fieldSitCt); }
            if (quickHWCt > 0) { logWrite.WriteLine ("Quick H&W: " + quickHWCt); }
            if (qwmCt > 0) { logWrite.WriteLine ("Quick Welfare Message: " + qwmCt); }
            if (dyfiCt > 0) { logWrite.WriteLine ("Did You Feel It: " + dyfiCt); }
            if (rriCt > 0) { logWrite.WriteLine ("RRI Welfare Radiogram: " + rriCt); }
            if (miCt > 0) { logWrite.WriteLine ("Medical Incident: " + miCt); }
            if (aprsCt > 0) { logWrite.WriteLine ("APRS checkins: " + aprsCt); }
            if (meshCt > 0) { logWrite.WriteLine ("Mesh checkins: " + meshCt); }
            logWrite.WriteLine ("Total Plain and other Checkins: " + (ct - localWeatherCt - severeWeatherCt - incidentStatusCt - icsCt - winlinkCkinCt - damAssessCt - fieldSitCt - quickHWCt - dyfiCt - rriCt - qwmCt - miCt - aprsCt - meshCt) + "\r\n");
            //var totalValidGPS = mapCt-noGPSCt;
            logWrite.WriteLine ("Total Checkins with a geolocation: " + (mapCt - noGPSCt));
            logWrite.WriteLine ("Total Checkins with something in the band field: " + bandCt);
            logWrite.WriteLine ("Total Checkins with something in the mode field: " + modeCt);
            // logWrite.WriteLine("\r\n++++++++++++++++\r\nmsgField not properly formatted for the following: \r\n-------------------------------");
            // logWrite.Write(badBandString);
            logWrite.WriteLine ("Checkins with a bad band field: " + badBandCt);
            // logWrite.Write(badModeString);
            logWrite.WriteLine ("Checkins with a bad mode field: " + badModeCt);
            logWrite.WriteLine (noGPSString + "Total without a location: " + noGPSCt);
            //logWrite.WriteLine("++++++++++++++++\r\nAdditional Comments\r\n-------------------------------");
            logWrite.Write ("\r\n++++++++++++++++\r\n" + addonString);

        }
        Console.WriteLine ("Done!\nThere were " + ct + " checkins. \nThe output files can be found in the folder \n" + currentFolder);
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
    static bool ExtractCoordinates (string input, out double latitude, out double longitude)
    {
        // Initialize output variables
        latitude = 0;
        longitude = 0;

        // Define the regular expression for latitude and longitude (with optional N/S/E/W directions)
        Regex regex = new Regex (@"([-+]?[0-9]*\.?[0-9]+)\s*[°]?\s*([NS]),?\s*([-+]?[0-9]*\.?[0-9]+)\s*[°]?\s*([EW])", RegexOptions.IgnoreCase);

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
            longitude = Math.Round (double.Parse (match.Groups [3].Value), 6);
            // If it's west (W), negate the longitude
            if (match.Groups [4].Value.ToUpper () == "W")
                longitude = -longitude;

            return true;
        }

        // Return false if latitude and longitude are not found
        return false;
    }
    static string ExtractMaidenheadGrid (string input)
    {
        // Define the regular expression for Maidenhead grid locator (4 or 6 character grids)
        Regex regex = new Regex (@"\b([A-R]{2}\d{2}[A-X]{0,2})\b", RegexOptions.IgnoreCase);

        // Search for a match in the input string
        Match match = regex.Match (input);

        if (match.Success)
        {
            return match.Value.ToUpper (); // Return the Maidenhead grid in uppercase
        }

        return string.Empty; // Return an empty string if no match is found
    }

    static (double, double) MaidenheadToGPS (string maidenhead)
    {
        if (maidenhead.Length < 4 || maidenhead.Length % 2 != 0)
            throw new ArgumentException ("Invalid Maidenhead grid format.");

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
        switch (input)
        {
            case "TELNET":
                break;
            case "160M":
                break;
            case "160":
                input = input + "M";
                break;
            case "80M":
                break;
            case "80":
                input = input + "M";
                break;
            case "60M":
                break;
            case "60":
                input = input + "M";
                break;
            case "40M":
                break;
            case "40":
                input = input + "M";
                break;
            case "30M":
                break;
            case "30":
                input = input + "M";
                break;
            case "20M":
                break;
            case "20":
                input = input + "M";
                break;
            case "17M":
                break;
            case "17":
                input = input + "M";
                break;
            case "15M":
                break;
            case "15":
                input = input + "M";
                break;
            case "12M":
                break;
            case "12":
                input = input + "M";
                break;
            case "10M":
                break;
            case "10":
                input = input + "M";
                break;
            case "6M":
                break;
            case "6":
                input = input + "M";
                break;
            case "2M":
                break;
            case "2":
                input = input + "M";
                break;
            case "1.25M":
                break;
            case "70CM":
                break;
            case "70":
                input = input + "CM";
                break;
            case "33CM":
                break;
            case "33":
                input = input + "CM";
                break;
            case "23CM":
                break;
            case "23":
                input = input + "CM";
                break;
            case "13CM":
                break;
            case "13":
                input = input + "CM";
                break;
            case "5CM":
                break;
            case "5":
                input = input + "cm";
                break;
            case "3cm":
                break;
            case "3":
                input = input + "CM";
                break;
            case "HF":
                break;
            case "SHF":
                break;
            case "UHF":
                break;
            case "VHF":
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
        switch (input)
        {
            case "SMTP":
                break;
            case "PACKET":
                break;
            case "PACKET WINLINK":
                input = "PACKET";
                break;
            case "X.25":
                input = "PACKET";
                break;
            case "ARDOP":
                break;
            case "VARA":
                if (input2 == "2M" || input2 == "70CM" || input2 == "6M") { input = "VARA FM"; }
                else { input = "VARA HF"; }
                break;
            case "FM":
                input = "VARA FM";
                break;
            case "FM VARA":
                input = "VARA FM";
                break;
            case "VARAFM":
                input = "VARA FM";
                break;
            case "VARA FM":
                break;
            case "VARAHF":
                input = "VARA HF";
                break;
            case "HF":
                input = "VARA HF";
                break;
            case "HF VARA":
                input = "VARA HF";
                break;
            case "VARA HF":
                break;
            case "PACTOR":
                break;
            case "INDIUM GO":
                break;
            case "MESH":                
                break;
            case "APRS":
                break;
            case "ROBUST PACKET":
                break;
            case "WINLINKPACKET":
                input = "PACKET";
                break;
            case "WINLINK EXPRESS":
                input = "PACKET";
                break;
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

}

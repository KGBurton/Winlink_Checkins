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


class Winlink_Checkins
{
    public static void Main(string[] args)
    {
        // Get the start date and end date from the user.
        DateTime startDate = DateTime.Today;
        DateTime endDate = DateTime.Today;
        string utcDate = DateTime.UtcNow.ToString("yyyy/MM/dd HH:mm:ss");
        //DateTime date;
        bool isValid = false;
        string input;
        
        Console.WriteLine("Enter the start date - must be within two weeks of today (yyyymmdd): ");
        startDate = GetValidDate();
        //startDate = startDate.ToUniversalTime();

        while (!isValid)
        {
            Console.WriteLine("Enter the end date - must be within two weeks of today (yyyymmdd): ");
            endDate = GetValidDate();
            int startDateCompare = DateTime.Compare(startDate, endDate);
            if (startDateCompare >= 0)
            {
                Console.WriteLine("The start date must be earlier than the end date. Please try again.");
            }
            else
            { isValid=true; }

        }
        endDate = endDate.AddDays(1);

        // Get the unique net identifier to screen only relevant messages from the folder
        // Console.WriteLine("Enter the unique net name for which the checkins are sent:");
        // string netName = Console.ReadLine();
        // Get the native call sign from the user to find the messages folder.
        Console.WriteLine("Enter YOUR call sign to find the messages folder. \n     If you leave it blank, the program will assume that it is already in the messages folder.");
        string yourCallSign = Console.ReadLine();

        // Get the data folder - either the global messages folder (default) or the current
        // operator's messages folder, assuming the default RMS installation location.
        string currentFolder = "";
        string applicationFolder = Directory.GetCurrentDirectory();
        string netName = "";
        if (yourCallSign != "")
        {
            currentFolder = "C:\\RMS Express\\" + yourCallSign + "\\Messages";
        }
        else
        {
            currentFolder = Directory.GetCurrentDirectory();
        }

        // Look for roster.txt in the folder. If it exists, get the first (and only)
        // row for comparison down below
        string rosterFile = applicationFolder+"\\roster.txt";
        string xmlFile = applicationFolder+"\\Winlink_Import.xml";
        // writeString variables to go in the output files
        StringBuilder netCheckinString = new StringBuilder();
        StringBuilder netAckString2 = new StringBuilder();
        StringBuilder bouncedString = new StringBuilder();
        StringBuilder duplicates = new StringBuilder();
        StringBuilder newCheckIns = new StringBuilder();
        StringBuilder csvString = new StringBuilder();
        StringBuilder mapString = new StringBuilder();
        mapString.Append("CallSign,Latitude,Longitude,Band,Mode\r\n");
        StringBuilder badBandString = new StringBuilder();
        StringBuilder badModeString = new StringBuilder();
        StringBuilder skippedString = new StringBuilder();
        StringBuilder removalString = new StringBuilder();
        StringBuilder addonString = new StringBuilder();
        StringBuilder noGPSString = new StringBuilder();
        string callSignPattern = @"\b\d{0,2}[A-Z]{1,2}\d{1,2}[A-Z]{1,6}\b";
        string testString = "";
        string rosterString = "";
        string bandStr = "";
        string modeStr = "";
        string noGPSStr = "";
        string checkIn = "";
        string msgFieldNumbered = "";
        noGPSString.Append ( "\r\n++++++++\r\nThese had neither GPS data nor Maidenhead Grids\r\n-------------------------\r\n");
        Random rnd = new Random();
        TextInfo textInfo = new CultureInfo("en-US", false).TextInfo;
        // Create root XML document
        XDocument xmlDoc = new XDocument(new XElement("WinlinkMessages"));
        XElement messageElement = new XElement
            ("export_parameters",
                new XElement("xml_file_version", "1.0"),
                new XElement("winlink_express_version", "1.7.17.0"),
                new XElement("callsign", "KB7WHO")
            );
        xmlDoc.Root.Add(messageElement);

        messageElement = new XElement("message_list", "");
        xmlDoc.Root.Add(messageElement);
        


        if (File.Exists(rosterFile))
        {
            rosterString = File.ReadAllText(rosterFile);
            rosterString = rosterString.ToUpper();
            //debug Console.WriteLine("rosterFile contents: "+rosterString);
            var startPosition = 0;
            var endPosition = rosterString.IndexOf("\r\n", startPosition);
            var len = endPosition - startPosition;
            if (len > 0)
            { netName = rosterString.Substring(startPosition, len); }
            else { netName = "GLAWN"; }
        }
        else
        {
            Console.WriteLine(currentFolder+"\\"+rosterFile+" \n was not found!, all checkins will appear to be new.\n");
        }
        var msgTotal = 0;
        var skipped = 0;
        var ct = 0;
        var dupCt = 0;
        var newCt = 0;
        var outOfRangeCt = 0;
        var removalCt = 0;
        var ackCt = 0;
        var localWeatherCt = 0;
        var severeWeatherCt = 0;
        var incidentStatusCt = 0; 
        var icsCt = 0;
        var ckinCt = 0;
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
        var badBandCt = 0;
        var badModeCt = 0;
        string locType = "";
        string xSource = "";
        double latitude = 0;
        double longitude = 0;
        
        // Select files with an extension of mime from the current folder.
        var files = Directory.GetFiles(currentFolder, "*.mime")
            .Where(file =>
            {
                DateTime fileDate = File.GetLastWriteTime(file);
                // debug Console.Write(fileDate+"\n");
                return fileDate >= startDate && fileDate <= endDate.AddDays(1);
            });

        Console.Write("\nMessages to process="+ files.Count() + " from folder "+currentFolder+"\n\n");

        // Create a text file called checkins.txt in the data folder and process the list of files.
        using (StreamWriter logWrite = new(Path.Combine(currentFolder, "checkins.txt")))
        // Create a text file called checkins.csv in the data folder and process the list of files.
        using (StreamWriter csvWrite = new(Path.Combine(currentFolder, "checkins.csv")))
        // Create a csv text file called mapfile.csv in the data folder to use as date for google maps
        using (StreamWriter mapWrite = new(Path.Combine(currentFolder, "mapfile.csv")))
        {
            // Read each file selected to find a line labeled To: and if the rest of the line contains netName, write the data from the line labeled X-Source: to the text file.
            foreach (string file in files)
            {
                using (StreamReader reader = new StreamReader(file))
                {
                    msgTotal++;
                    //debug Console.Write("File "+file+"\n");
                    string fileText = reader.ReadToEnd();
                    fileText = fileText.ToUpper()
                        .Replace("=\r\n","")
                        .Replace("=20","");

                    // get needed header info
                    var startPosition = fileText.IndexOf("DATE: ")+11;
                    var len = 20;
                    string sentDate = fileText.Substring(startPosition, len);
                    DateTime sentDateUni = DateTime.Parse(sentDate);
                   
                    startPosition = fileText.IndexOf("MESSAGE-ID: ")+12;
                    var endPosition = fileText.IndexOf("\r\n", startPosition);
                    len = endPosition - startPosition;
                    string messageID = fileText.Substring(startPosition, len);

                    // find the end of the header section
                    var endHeader = fileText.IndexOf("CONTENT-TRANSFER-ENCODING:");

                    // was it forwarded?
                    var forwarded = fileText.IndexOf("WAS FORWARDED BY");

                    // was it APRSmail?
                    var APRS = fileText.IndexOf("APRSEMAIL2");

                    // check for acknowledgement message and discard                  
                    var ack = fileText.IndexOf("ACKNOWLEDGEMENT");

                    // check for removal message and discard                  
                    var removal = fileText.IndexOf("REMOV");

                    // look to see if it was a bounced message
                    var bounced = fileText.IndexOf("UNDELIVERABLE");

                    // check for local weather report
                    var localWeather = fileText.IndexOf("CURRENT LOCAL WEATHER CONDITIONS");

                    // check for severe weather
                    var severeWeather = fileText.IndexOf("SEVERE WX REPORT");

                    // check incident status report
                    var incidentStatus = fileText.IndexOf("INCIDENT STATUS");

                    // check for ICS 213 msg
                    var ics = fileText.IndexOf("TEMPLATE VERSION: ICS 213");

                    // check for winlink checkin message
                    var ckin = fileText.IndexOf("MAP FILE NAME: WINLINK CHECK", endHeader);
                    // some people include WINLINK CHECK-IN in the subject which confuses the program
                    // into thinking this is a winlink checkin FORM!! Catch it ...
                    if (ckin < 0) 
                    {
                        ckin = fileText.IndexOf("WINLINK CHECK-IN 5.0.10", endHeader);
                    }
                    
                    // check for odd checkin message - don't let it scan through to a binary attachment!
                    var lenBPQ = fileText.Length-10;
                    if (lenBPQ > 800) { lenBPQ = 800; }
                    var BPQ = fileText.IndexOf("BPQ", 1, lenBPQ);

                    // check for damage assessment report
                    var damAssess = fileText.IndexOf("SURVEY REPORT - CATEGORIES");

                    // check for field situation report
                    var fieldSit = fileText.IndexOf("EMERGENT/LIFE SAFETY");

                    // check for Quick Health & Welfare report
                    var quickHW = fileText.IndexOf("QUICK H&W");

                    // check for RRI Welfare Radiogram
                    var rriWR = fileText.IndexOf("TEMPLATE VERSION: RRI WELFARE RADIOGRAM");

                    // check for Did You Feel It report
                    var dyfi = fileText.IndexOf("DYFI WINLINK");

                    // check for RRI Welfare Radiogram
                    var qwm = fileText.IndexOf("TEMPLATE VERSION: QUICK WELFARE MESSAGE");

                    // check for Medical Incident Report
                    var mi = fileText.IndexOf("INITIAL PATIENT ASSESSMENT");

                    // screen dates to eliminate file dates that are different from the sent date and fall outside the net span
                    int startDateCompare = DateTime.Compare(sentDateUni, startDate);
                    int endDateCompare = DateTime.Compare(sentDateUni, endDate);
                   
                    // discard acknowledgements
                    if (ack >0)
                    {
                        skipped++;
                        ackCt++;
                        junk=0; //debug Console.Write(file+" is an acknowedgement, skipping.");
                    }

                    else if (startDateCompare < 0 || endDateCompare > 0)
                    {
                        skipped++;
                        outOfRangeCt++;
                        Console.Write(messageID+" sendDate fell outside the start/end dates\r\n");
                        skippedString.Append("Out of date range message skipped: "+file+"\r\n");
                    }

                    else if (removal >0)
                    {
                        startPosition = fileText.IndexOf("FROM:")+6;
                        endPosition = fileText.IndexOf("\r\n", startPosition);
                        len = endPosition - startPosition;
                        checkIn = fileText.Substring(startPosition, len);
                        {
                            checkIn = checkIn.Replace(',', ' ');
                            // Create a Regex object with the pattern
                            Regex regexCallSign = new Regex(callSignPattern, RegexOptions.IgnoreCase);
                            // find the first callsign match in the checkIn string
                            Match match = regexCallSign.Match(checkIn);
                            if (match.Success) checkIn = match.Value;
                        }
                        removalString.Append("Message from: "+checkIn+" in "+messageID +" was a removal request.\r\n");
                        removalCt++;
                        junk = 0;  // debug Console.Write("Removal Request: "+file+", skipping.");
                    }
                    else if (bounced > 0)
                    {
                        startPosition = bounced;
                        endPosition = fileText.IndexOf("\r\n", startPosition);
                        len = endPosition - startPosition;
                        checkIn = fileText.Substring(startPosition, len);
                        {
                            checkIn = checkIn.Replace(',', ' ');
                            // Create a Regex object with the pattern
                            Regex regexCallSign = new Regex(callSignPattern, RegexOptions.IgnoreCase);
                            // find the first callsign match in the checkIn string
                            Match match = regexCallSign.Match(checkIn);
                            if (match.Success) checkIn = match.Value;
                        }
                        bouncedString.Append("Message to: "+checkIn+" was not deliverable.\r\n");
                        skipped++;
                    }
                    else
                    {
                        // determine if the message has something in the subject to do with GLAWN
                        // extended to include the TO: field in case they didn't put the netName in the subject
                        startPosition = fileText.IndexOf("SUBJECT:")+9;
                        endPosition = fileText.IndexOf("MESSAGE-ID", startPosition);
                        len = endPosition - startPosition;
                        string subjText = fileText.Substring(startPosition, len);

                        // deterimine if it was forwarded to know to look below the first header info

                        if (subjText.Contains(netName))
                        {
                            
                            // get x-Source if available XXXX
                            var xSrc = fileText.IndexOf("X-SOURCE: ");
                            if (xSrc > 0)
                            {
                                startPosition = xSrc +10;
                                endPosition = fileText.IndexOf("\r\n", startPosition);
                                len = endPosition - startPosition;
                                if (len > 0) { xSource = fileText.Substring(startPosition, len); }
                            }

                            // skip APRS header 
                            if (APRS > 0)
                            {
                                startPosition = fileText.IndexOf("FROM:", APRS);
                                if (startPosition > 0)
                                {
                                    startPosition = fileText.IndexOf("\r\n", startPosition)+2;
                                    endPosition = fileText.IndexOf("DO NOT REPLY", startPosition)-1;
                                }
                                aprsCt++;
                            }

                            // adjust for ICS 213
                            else if (ics > 0)
                            {
                                // check first is it a reply (checkin will be in a different location

                                startPosition = fileText.IndexOf("9. REPLY:");
                                if (startPosition > 0)
                                {
                                    startPosition = startPosition+11 ;
                                    endPosition = fileText.IndexOf("REPLIED BY:", startPosition)-3;
                                }
                                else
                                {                                    
                                    startPosition = fileText.IndexOf("MESSAGE:")+12;
                                    endPosition = fileText.IndexOf("APPROVED BY:", startPosition)-3;
                                }
                            }
                            // adjust for winlink checkin
                            else if (ckin >0)
                            {
                                // the winlink check-in form changed format between 5.0.10 and 5.0.5 so check for that
                                var ckinOffset = fileText.IndexOf("WINLINK CHECK-IN 5.0.5");
                                // if (ckinOffset > 0) { ckinOffset = 9; } else { ckinOffset = 13; }
                                // startPosition = fileText.IndexOf("COMMENTS:")+ckinOffset;
                                startPosition = fileText.IndexOf("COMMENTS:")+9;
                                endPosition = fileText.IndexOf("----------", startPosition)-1;
                            }

                            // adjust for odd message that insert an R: line at the top
                            else if (BPQ > 0)
                            {
                                startPosition = fileText.IndexOf("BPQ", 1, lenBPQ)+12;
                                endPosition = fileText.IndexOf("--BOUNDARY", startPosition)-2;
                            }
                            else if (localWeather > 0)
                            {
                                startPosition = fileText.IndexOf("NOTES:")+11;
                                endPosition = fileText.IndexOf("----------", startPosition)-1;
                            }

                            else if (severeWeather > 0)
                            {
                                startPosition = fileText.IndexOf("COMMENTS:")+10;
                                endPosition = fileText.IndexOf("----------", startPosition)-1;
                            }

                            else if (incidentStatus > 0)
                            {                                
                                startPosition = fileText.IndexOf("REPORT SUBMITTED BY:")+20;
                                endPosition = fileText.IndexOf("----------", startPosition)-1;
                            }

                            else if (damAssess > 0)
                            {                                
                                startPosition = fileText.IndexOf("COMMENTS:")+21;
                                endPosition = fileText.IndexOf("----------", startPosition)-1;
                            }

                            else if (fieldSit > 0)
                            {
                                startPosition = fileText.IndexOf("COMMENTS:")+11;
                                endPosition = fileText.IndexOf("\r\n", startPosition);
                            }

                            else if (dyfi > 0)
                            {
                                startPosition = fileText.IndexOf("COMMENTS")+11;
                                endPosition = fileText.IndexOf("\r\n", startPosition)-1;
                            }

                            else if (rriWR > 0)
                            {
                                startPosition = fileText.IndexOf("BT\r\n")+3;
                                endPosition = fileText.IndexOf("------", startPosition)-1;
                            }

                            else if (qwm > 0)
                            {
                                startPosition = fileText.IndexOf("IT WAS SENT FROM:");
                                endPosition = fileText.IndexOf("------", startPosition)-1;
                            }
                            else if (mi > 0)
                            {
                                startPosition = fileText.IndexOf("ADDITIONAL INFORMATION");
                                startPosition = fileText.IndexOf("\r\n", startPosition);
                                endPosition = fileText.IndexOf("----", startPosition)-1;
                            }
                            else
                            {
                                // end of the header information as the start of the msg field
                                if (forwarded <= 0)
                                {
                                    startPosition = fileText.IndexOf("QUOTED-PRINTABLE")+20;
                                }
                                else
                                {
                                    // startPosition = forwarded+59;
                                    startPosition = fileText.IndexOf("SUBJECT:", forwarded)+9;
                                    startPosition = fileText.IndexOf("\r\n", startPosition)+4;
                                    // look for a second Subject tag
                                    startPosition = fileText.IndexOf("SUBJECT:", forwarded)+9;
                                    startPosition = fileText.IndexOf("\r\n", startPosition)+4;

                                }
                                endPosition = fileText.IndexOf("--BOUNDARY", startPosition)-1;
                            }
                            len = endPosition - startPosition;
                            if (len == 0)
                            {
                                Console.Write("Nothing in the message field: "+file+"\n");
                                // try retrieving something from the from field
                                startPosition = fileText.IndexOf("FROM:")+6;
                                endPosition = fileText.IndexOf("@", startPosition);
                                len = endPosition - startPosition;
                            }
                            if (len == 0)
                            {
                                Console.Write("Trying the subject field: "+file+"\n");
                                // try retrieving something from the subject field
                                startPosition = fileText.IndexOf("SUBJECT:")+9;
                                endPosition = fileText.IndexOf("\r\n", startPosition)-1;
                                len = endPosition - startPosition;
                            }

                            if (len < 0)
                            {
                                Console.Write("endPostion is less than startPosition in: "+file+"\n");
                                Console.Write("Break at line 522ish. Press enter to close.");
                                input = Console.ReadLine();
                                break;
                            }

                            string msgField = fileText.Substring(startPosition, len);
                            msgField = msgField
                                .Replace("=20", "")
                                .Replace("=0A", "")
                                .Replace("=0", "")
                                .Replace("16. CONTACT INFO:", ",")
                                .Trim()
                                .Replace("  ", " ")
                                .Replace("  ", " ")
                                //.Replace(".", "") this causes problems with decimal band freq
                                .Replace(", ", ",")
                                .Replace("[NO CHANGES OR EDITING OF THIS MESSAGE ARE ALLOWED]", "")
                                .Replace("[MESSAGE RECEIPT REQUESTED]","")
                                .Replace(" ,", ",")
                                .Replace("\"", "")
                                .Trim()
                                //.Trim(',')
                                +",";
                            checkIn = msgField
                                //.Replace(" ,", ",")    
                                //.Trim()
                                //.Trim(',')
                                //.Trim()+",";
                                ;

                            // Create a Regex object with the pattern
                            Regex regexCallSign = new Regex(callSignPattern, RegexOptions.IgnoreCase);

                            // find the first callsign match in the checkIn string
                            Match match = regexCallSign.Match(checkIn);
                            if (match.Success)
                            {
                                checkIn = match.Value;
                                if (xSource == "") { xSource = checkIn; }
                            }
                            else
                            {
                                // try the from field since the callsign could not be located in the msg field
                                startPosition = fileText.IndexOf("FROM:")+6;
                                endPosition = fileText.IndexOf("@", startPosition);
                                if (endPosition < 0) { endPosition = fileText.IndexOf("SUBJECT:")-1; }
                                len = endPosition - startPosition;
                                if (len>0)
                                {
                                    checkIn = fileText.Substring(startPosition, len);
                                    // Create a Regex object with the pattern
                                    regexCallSign = new Regex(callSignPattern, RegexOptions.IgnoreCase);
                                    match = regexCallSign.Match(checkIn);
                                    if (match.Success)
                                    {
                                        checkIn = match.Value;
                                    }
                                    else
                                    {
                                        checkIn ="";
                                    }
                                }
                            }
                            // debug Console.Write("Start at:"+startPosition+": and end at:"+endPosition+"\nCallsign found: "+checkIn);
                            // eliminate duplicates                                
                            if (checkIn == "")
                            {
                                Console.Write("Callsign not found in: "+file);
                            }
                            else
                            {
                                startPosition = testString.IndexOf(checkIn);
                                if (startPosition >= 0)
                                {
                                    if (dupCt == 0) { duplicates.Append("Duplicates: \r\n"); }
                                    //debug Console.Write("netName "+checkIn+" is a duplicate, skipping. It is "+dupCt+" of "+msgTotal+" total messages.\n");
                                    duplicates.Append(checkIn+", ");
                                    dupCt++;
                                }
                                else 
                                {
                                    ct++;
                                    if (localWeather > 0) { localWeatherCt++; }
                                    if (severeWeather > 0) { severeWeatherCt++; }
                                    if (ckin >0) { ckinCt++; }
                                    if (incidentStatus > 0) { incidentStatusCt++; }
                                    if (ics > 0) { icsCt++; }
                                    if (damAssess > 0) { damAssessCt++; }                                    
                                    if (fieldSit > 0) { fieldSitCt++; }
                                    if (quickHW > 0) { quickHWCt++;}
                                    if (dyfi > 0) { dyfiCt++; }
                                    if (rriWR > 0) { rriCt++; }
                                    if (qwm > 0) { qwmCt++; }
                                    if (mi > 0) { miCt++; }
                                    testString = testString+checkIn+" | ";
                                    // the spreadsheet chokes if the string ends with "|" so
                                    // don't let that happen by writing the first one without a delimiter
                                    // prepending the delimiter to the rest.
                                    if (ct == 1)
                                    {
                                        netCheckinString.Append(checkIn);
                                    }
                                    else if (ct > 1)
                                    {
                                        netCheckinString.Append("|"+checkIn);
                                    }
                                    netAckString2.Append(checkIn+";");
                                    // find message, format for csv file, and save
                                    var msgFieldStart = msgField.IndexOf("\r\n");
                                    string notFirstLine = "";
                                    if (msgFieldStart > 0) 
                                    {
                                        len = msgField.Length - msgFieldStart;
                                        if (len > 0) 
                                        { 
                                            notFirstLine = msgField.Substring(msgFieldStart,len);
                                            notFirstLine = notFirstLine.Replace("\n", ", ")
                                            .Replace("\r", "")
                                            //.Replace("73","")
                                            .Trim()
                                            ;
                                            startPosition = notFirstLine.IndexOf("73");
                                            if (startPosition >0)
                                            {
                                                endPosition = notFirstLine.IndexOf("\r\n", startPosition)+2;
                                                len = endPosition - startPosition;
                                                if (len > 0)
                                                {
                                                    notFirstLine = notFirstLine.Substring(0, startPosition)+notFirstLine.Substring(endPosition);
                                                }
                                                else
                                                { 
                                                    notFirstLine = notFirstLine.Substring(0, startPosition);
                                                }
                                            }
                                            notFirstLine = notFirstLine
                                                .Replace(", ,",",")
                                                .Trim()
                                                .Trim(',')
                                                .Trim()
                                                .Trim(',')

                                                ;
                                            if (notFirstLine.Length > 0) { addonString.Append(checkIn + ": " + notFirstLine+"\r\n"); }
                                        } 
                                    }
                                    // Extract latitude and longitude
                                    //skip past the messageID because sometimes the regex for coordinates matches it
                                    startPosition = fileText.IndexOf("MESSAGE-ID:");
                                    startPosition = fileText.IndexOf("\r\n",startPosition)+2;
                                    len = fileText.Length - startPosition;
                                    if (len>0)
                                    {
                                        if (ExtractCoordinates(fileText.Substring(startPosition), out latitude, out longitude))
                                        {
                                            // Console.WriteLine(messageID+" latitude: "+latitude+" longitude: "+longitude);                                
                                        }
                                        else
                                        {
                                            // no valid GPS coordinates found, look for a maidenhead grid
                                            string maidenheadGrid = ExtractMaidenheadGrid(fileText);
                                            if (!string.IsNullOrEmpty(maidenheadGrid))
                                            {
                                                // Console.WriteLine($"Maidenhead Grid: {maidenheadGrid}");
                                                // Convert Maidenhead to GPS coordinates
                                                (latitude, longitude) = MaidenheadToGPS(maidenheadGrid);
                                                // Console.WriteLine($"No GPS coords found, using Maidenhead Grid: {maidenheadGrid}"+$". From Maidenhead Grid Latitude: {latitude}"+$"  Longitude: {longitude}");
                                            }
                                            else
                                            {
                                                // No valid Maidenhead grid found either, make up something in the middle of the Atlantic
                                                double locChange = Math.Round(rnd.NextDouble()*10, 6);
                                                latitude = Math.Round((27.187512+locChange), 6);
                                                longitude= Math.Round((-60.144742+locChange), 6);
                                                // Console.WriteLine("No valid grid and no GPS coordinates found in: "+messageID+" latitude set to: "+latitude+" longitude set to: "+longitude);
                                                noGPSCt++;
                                                noGPSString.Append("\t"+messageID+"- - "+checkIn+": latitude set to: "+latitude+" longitude set to: "+longitude+"\r\n");
                                            }
                                        }
                                    }
                                    msgField = msgField.Replace("\r\n", ",");
                                    msgFieldNumbered = fillFieldNum(msgField);
                                    csvString.Append(xSource+","+latitude+","+longitude+","+locType+","+msgFieldNumbered+"\r\n");

                                    // find the band if it's where it's supposed to be
                                    bandStr ="";
                                    // modeStr = "";
                                    len =0;
                                    // debug Console.Write("\r\nmsgField ="+msgField+"\r\n");
                                    startPosition = IndexOfNthSB(msgField, (char)44, 0, 6)+1;
                                    if (startPosition > 0) { endPosition = IndexOfNthSB(msgField, (char)44, 0, 7); len = endPosition-startPosition; }
                                    if (len > 0 && msgField.Length >= len)
                                    {
                                        bandStr=  msgField.Substring(startPosition, len)
                                            .ToLower()
                                            .Replace("5.8ghz", "5cm")
                                            .Replace("packet", "")
                                            .Replace(".", "")
                                            .Replace(" ", "")
                                            .Replace(" meters", "m")
                                            .Replace(" meter", "m")
                                            .Replace("meters", "m")
                                            .Replace("meter", "m")
                                            .Replace("vhf", "2m")
                                            .Replace("uhf", "70cm")
                                            .Replace("(", "")
                                            .Replace(")", "")
                                            .Replace("n/a", "Telnet")
                                            .Replace("na", "Telnet")                                                                                        
                                            .Replace("5ghz", "5cm")
                                            .Replace("73m", "80m")
                                            .Replace("75m", "80m")
                                            .Replace("telnet", "Telnet")
                                            .Trim()
                                            // .Replace("packet", "")
                                            ;
                                        if (bandStr.IndexOf("m") == -1 && bandStr != "Telnet" && bandStr != "vara")
                                        {
                                            // if the band is a simple number (no cm or m), add m to it for meters
                                            if (bandStr.IndexOf("ghz") == -1) { bandStr = bandStr+"m"; }
                                        }
                                        if (bandStr =="telnet") { bandStr = textInfo.ToTitleCase(bandStr.ToLower()); }
                                        bandCt++;
                                        switch (bandStr)
                                        {
                                            case "Telnet":
                                                break;
                                            case "160m":
                                                break;
                                            case "80m":
                                                break;
                                            case "60m":
                                                break;
                                            case "40m":
                                                break;
                                            case "30m":
                                                break;
                                            case "20m":
                                                break;
                                            case "17m":
                                                break;
                                            case "15m":
                                                break;
                                            case "12m":
                                                break;
                                            case "10m":
                                                break;
                                            case "6m":
                                                break;
                                            case "2m":
                                                break;
                                            case "1.25m":
                                                break;
                                            case "70cm":
                                                break;
                                            case "33cm":
                                                break;
                                            case "23cm":
                                                break;
                                            case "13cm":
                                                break;
                                            case "5cm":
                                                break;
                                            case "3cm":
                                                break;
                                            default:
                                                // Console.WriteLine("Band is not standard for "+messageID+"  "+checkIn+": "+bandStr+" - "+ msgField+ "\r\n");
                                                msgFieldNumbered = msgField;
                                                msgFieldNumbered = fillFieldNum(msgFieldNumbered);
                                                badBandString.Append("\tBad Band: "+messageID+" - "+checkIn+": "+bandStr+" - |"+ msgFieldNumbered + "|\r\n");
                                                badBandCt++;
                                                break;
                                        }
                                    }
                                    modeStr ="";
                                    len =0;
                                    // debug Console.Write("\r\nmsgField ="+msgField+"\r\n");
                                    startPosition = IndexOfNthSB(msgField, (char)44, 0, 7)+1;
                                    if (startPosition > 0) { endPosition = IndexOfNthSB(msgField, (char)44, 0, 8); len = endPosition-startPosition; }
                                    if (len > 0 && msgField.Length >= len)
                                    {
                                        modeStr=  msgField.Substring(startPosition, len)
                                            .ToUpper()
                                            .Replace("WINLINK","")
                                            .Replace("AREDN", "MESH")
                                            .Replace("AX.25", "PACKET")
                                            .Replace("WINLINK", "")
                                            .Replace("(", "")
                                            .Replace("ARDOP HF","ARDOP")
                                            .Replace("VHF VARA","VARA FM")
                                            .Replace("VARAFM","VARA FM")
                                            .Replace("VERA", "VARA")
                                            .Replace("HF ARDOP", "ARDOP")
                                            .Replace(")", "")
                                            .Replace("-", " ")
                                            .Replace("=20","")
                                            //.Replace("HF", "VARA HF")
                                            .Replace("VHF PACKET", "PACKET")
                                            .Replace("TELNET", "SMTP")
                                            .Trim();
                                        if (modeStr.IndexOf("MESH") > 0) { modeStr = "MESH"; }
                                        // {
                                        //     if (modeStr.IndexOf("ghz") == -1) { modeStr = modeStr+"m"; }
                                        // }
                                        if (bandStr =="Telnet") { modeStr = "SMTP"; }
                                        modeCt++;
                                    }
                                    else
                                    {
                                        //if (bandStr == "2m" || bandStr == "70cm") { modeStr = "VHF"; }
                                        if (bandStr == "Telnet") { modeStr = "SMTP"; }
                                    }
                                    switch (modeStr)
                                    {
                                        case "SMTP":
                                            bandStr = "Telnet";
                                            break;
                                        case "PACKET":
                                            break;
                                        case "PACKET WINLINK":
                                            modeStr = "PACKET";
                                            break;
                                        case "X.25":
                                            modeStr = "PACKET";
                                            break;
                                        case "ARDOP":
                                            break;
                                        case "VARA":
                                            if (bandStr =="2m" || bandStr =="70cm") { modeStr = "VARA FM"; }
                                            else { modeStr = "VARA HF"; }
                                            break;
                                        case "FM":
                                            modeStr = "VARA FM";
                                            break;
                                        case "FM VARA":
                                            modeStr = "VARA FM";
                                            break;
                                        case "VARAFM":
                                            modeStr = "VARA FM";
                                            break;
                                        case "VARA FM":
                                            break;
                                        case "VARAHF":
                                            modeStr = "VARA HF";
                                            break;
                                        case "HF":
                                            modeStr = "VARA HF";
                                            break;
                                        case "HF VARA":
                                            modeStr = "VARA HF";
                                            break;
                                        case "VARA HF":
                                            break;
                                        case "PACTOR":
                                            break;
                                        case "INDIUM":
                                            break;
                                        case "MESH":
                                            meshCt++;
                                            break;
                                        case "APRS":
                                            break;
                                        case "WINLINKPACKET":
                                            modeStr = "PACKET";
                                            break;
                                        case "WINLINK EXPRESS":
                                            modeStr = "PACKET";
                                            break;
                                        default:
                                            //Console.WriteLine("Bad mode: "+messageID+"  "+checkIn+": "+modeStr+" - "+ msgField+ "\r\n");
                                            msgFieldNumbered = msgField;
                                            msgFieldNumbered = fillFieldNum(msgFieldNumbered);
                                            badModeString.Append("\tBad mode: "+messageID+" - "+checkIn+": "+modeStr+" -  |  "+ msgFieldNumbered+ "\r\n");
                                            badModeCt++;
                                            break;
                                    }


                                    // debug Console.Write("modeStr final=|"+modeStr+"|  \r\n");


                                    // add to mapString csv file if xloc was found
                                    if (latitude != 0) { 
                                        mapString.Append(xSource+","+latitude+","+longitude+","+bandStr+","+modeStr+"\r\n");
                                        mapCt++;
                                    }

                                    // xml data
                                    XElement message_list = xmlDoc.Descendants("message_list").FirstOrDefault();
                                    message_list.Add(new XElement("message", 
                                        new XElement("id", messageID),
                                        new XElement("foldertype", "Global"),
                                        new XElement("folder", "GLAWN"),
                                        new XElement("subject", "GLAWN acknowledgement ", DateTime.UtcNow.ToString("yyyy-MM-dd")),
                                        new XElement("time", utcDate),
                                        new XElement("sender", "KB7WHO"),
                                        new XElement("To", xSource),
                                        new XElement("rmsoriginator", ""),
                                        new XElement("rmsdestination", ""),
                                        new XElement("rmspath", ""),
                                        new XElement("location", "43.845831N, 111.745744W (GPS)"),
                                        new XElement("csize", ""),
                                        new XElement("messageserver", ""),
                                        new XElement("precedence", "2"),
                                        new XElement("peertopeer", "False"),
                                        new XElement("routingflag", ""),
                                        new XElement("source", "KB7WHO"),
                                        new XElement("unread", "True"),
                                        new XElement("flags", "0"),
                                        new XElement("messageoptions", "False|False|||||"),
                                        new XElement
                                        ("mime", "Date: "+utcDate+"\r\n"+
                                            "From: GLAWN@winlink.org\r\n"+
                                            "Subject: GLAWN acknowledgement ", utcDate+"\r\n"+
                                            "To: "+checkIn+"\r\n"+
                                            "Message-ID: "+messageID+"\r\n"+
                                            "X-Source: KB7WHO\r\n"+
                                            "X-Location: 43.845831N, 111.745744W(GPS) \r\n"+
                                            "MIME-Version: 1.0\r\n"+
                                            "\r\n"+
                                            "Content-Type: text/plain; charset=\"iso-8859-1\"\r\n"+
                                            "Content-Transfer-Encoding: quoted-printable\r\n"+
                                            "\r\n"+
                                            "Thank you for checking in to the GLAWN. This is a copy of your message and extracted data. \r\n"+
                                            "Message: "+msgFieldNumbered+"\r\n"+
                                            "Extracted Data:\r\n" +
                                                "   Lattitude: "+latitude+"\r\n"+
                                                "   Longitude: "+longitude+"\r\n"+
                                                "   Band: "+bandStr+"\r\n"+
                                                "   Mode: "+modeStr+"\r\n"                                                        
                                        )                                            
                                    ));
                                    
                                    // Add the message message_list
                                    xmlDoc.Root.Add(messageElement);
                                }
                                junk = 0; // just so i could put a debug here
                            }
                            var tempCt = ct+dupCt+ackCt+removalCt;
                            //debug Console.Write("checkins:"+ct+"  duplicates:" + dupCt+"  removals:"+removalCt+"  acks:"+ackCt + "  combined:"+tempCt+"   actual total:"+msgTotal+"\n");
                            // missing from roster section. Check to see if the checkin is in the roster. 
                            startPosition = rosterString.IndexOf(checkIn);
                            if (startPosition < 0)
                            {
                                if (newCt == 0)
                                {
                                    newCheckIns.Append("New Checkins:\r\n");
                                }
                                // debug
                                Console.Write(checkIn+" was not found in roster.txt. \n");
                                newCheckIns.Append(checkIn+", ");
                                // update roster.txt to contain the new checkin
                                File.AppendAllText("roster.txt", "; "+  checkIn);
                                newCt++;
                            }
                        }
                        else
                        {
                            skipped++;
                            Console.Write("Could not find netName in this message: "+file+"\n");
                            skippedString.Append("Skipped message: "+file+"\r\n");
                        }
                    }

                }
            }
            var tempCT=14;
            logWrite.WriteLine("Total Checkins Recorded:"+ct+"    Duplicates Skipped:"+dupCt+"    Removal Requests: "+removalCt);
            logWrite.WriteLine("Non-"+netName+" checkin messages skipped: "+skipped+"(including "+ackCt+" acknowledgements and "+outOfRangeCt+" out of date range messages skipped.\r\n");
            logWrite.WriteLine("Total messages processed: "+msgTotal+"\r\n");
            logWrite.WriteLine("Row "+tempCT+" goes into GLAWN Spreadsheet at row 1 of the checkin column to be recorded.");
            tempCT++;
            logWrite.WriteLine("Row "+tempCT+" goes into GLAWN Spreadsheet at row 2 of the checkin column and is the copy list for the checkin acknowledgement.");
            tempCT=17;
            logWrite.Write("Rows "+tempCT+" & "+(tempCT+1)+" have the list of duplicates found.\r\n");
            tempCT=20;
            logWrite.WriteLine("Rows "+tempCT+" and beyond are bounced messages, new checkins that should be added to \r\n\tthe spreadsheet, skipped messages that didn't have a netName,\r\n\tand other notifications including the checkin form that was used, and\r\n\tthe number that had mapping coordinates.\r\n");
            logWrite.WriteLine(netCheckinString);
            logWrite.WriteLine(netAckString2+"\r\n");
            csvWrite.WriteLine(csvString);
            mapWrite.WriteLine(mapString);
            xmlDoc.Save(xmlFile);

            if (duplicates.Length==0)
            {
                duplicates.Append("No duplicates found this week.");
                Console.Write("No duplicates found this week..\r\n\r\n");
            }
            logWrite.WriteLine(duplicates+"\r\n");
            logWrite.WriteLine("Messages that bounced: "+bouncedString+"\r\n");
            if (newCt == 0)
            {
                newCheckIns.Append("No new checkins found this week.");
                Console.Write("No new checkins found this week.\n\n");
            }
            logWrite.Write(newCheckIns+"\r\n");
            logWrite.Write(skippedString+"\r\n");
            logWrite.Write(removalString+"\r\n");
            logWrite.WriteLine("Local Weather Checkins: "+localWeatherCt);
            logWrite.WriteLine("Severe Weather Checkins: "+severeWeatherCt);
            logWrite.WriteLine("Incident Status Checkins: "+incidentStatusCt);
            logWrite.WriteLine("ICS-213 Checkins: "+icsCt);
            logWrite.WriteLine("Winlink Check-in Checkins: "+ckinCt);
            logWrite.WriteLine("Damage Assessment Checkins: "+damAssessCt);
            logWrite.WriteLine("Field Situation Report Checkins: "+fieldSitCt);
            logWrite.WriteLine("Quick H&W: "+quickHWCt);
            logWrite.WriteLine("Quick Welfare Message: "+qwmCt);
            logWrite.WriteLine("Did You Feel It: "+dyfiCt);
            logWrite.WriteLine("RRI Welfare Radiogram: "+rriCt);
            logWrite.WriteLine("Medical Incident: "+miCt);
            logWrite.WriteLine("APRS checkins: "+aprsCt);
            logWrite.WriteLine("Mesh checkins: "+meshCt);
            logWrite.WriteLine("Total Plain and other Checkins: "+(ct-localWeatherCt-severeWeatherCt-incidentStatusCt-icsCt-ckinCt-damAssessCt-fieldSitCt-quickHWCt-dyfiCt-rriCt-qwmCt-miCt-aprsCt-meshCt)+"\r\n");
            //var totalValidGPS = mapCt-noGPSCt;
            logWrite.WriteLine("Total Checkins with a geolocation: "+(mapCt-noGPSCt));
            logWrite.WriteLine("Total Checkins with something in the band field: "+bandCt);
            logWrite.WriteLine("Total Checkins with something in the mode field: "+modeCt);
            logWrite.WriteLine("\r\n++++++++++++++++\r\nmsgField not properly formatted for the following: \r\n-------------------------------");
            logWrite.Write(badBandString);
            logWrite.WriteLine("Checkins with a bad band field: "+badBandCt+"\r\n");
            logWrite.Write(badModeString);
            logWrite.WriteLine("Checkins with a bad mode field: "+badModeCt);
            logWrite.WriteLine(noGPSString+"\r\nTotal without a location: "+noGPSCt);
            logWrite.WriteLine("++++++++++++++++\r\nAdditional Comments\r\n-------------------------------");
            logWrite.Write(addonString);

        }
        Console.WriteLine("Done!\nThere were "+ct+" checkins. \nThe output checkins.txt can be found in the folder \n"+currentFolder);
        Console.WriteLine("\n\nPress enter to continue.");
        Console.ReadLine();
    }
    //public static class Globals
    public static int IndexOfNthSB(string input,
             char value, int startIndex, int nth)
    {
        if (nth < 1)
            throw new NotSupportedException("Param 'nth' must be greater than 0!");
        var nResult = 0;
        for (int i = startIndex; i < input.Length; i++)
        {
            if (input[i] == value)
                nResult++;
            if (nResult == nth)
                return i;
        }
        return -1;
    }
    static DateTime GetValidDate()
    {
        DateTime date = default;  // Initialize date to its default value
        bool isValid = false;
        
        while (!isValid)
        {
            string input = Console.ReadLine();
            DateTime todayDate = DateTime.Today ;
            int dateCompare = 0;

            // Validate using the specific format YYYYMMDD
            if (DateTime.TryParseExact(input, "yyyyMMdd", null, System.Globalization.DateTimeStyles.None, out date))
            {                
                isValid = true;
                dateCompare = DateTime.Compare(date, todayDate.AddDays(-14));
                if ( dateCompare<0)
                {
                    isValid =false;
                    Console.WriteLine("Invalid date: "+date+" Must be within two weeks of today. Please try again.");
                }
                dateCompare = DateTime.Compare(todayDate.AddDays(14),date);
                if (dateCompare<0)
                {
                    isValid =false;
                    Console.WriteLine("Invalid date: "+input+" Must be within two weeks of today.  Please try again.");
                }
            }
            else
            {
                Console.WriteLine("Invalid date format. "+input+"Please use YYYYMMDD format and try again.");
            }
        }
        return date;
    }

    static void SaveDate(DateTime date)
    {
        string filePath = "dates.txt";
        try
        {
            using (StreamWriter writer = new StreamWriter(filePath, true))
            {
                writer.WriteLine(date.ToString("yyyy-MM-dd"));
            }
        }
        catch (Exception ex)
        {
            Console.WriteLine($"An error occurred while saving the date: {ex.Message}");
        }
    }
    static bool ExtractCoordinates(string input, out double latitude, out double longitude)
    {
        // Initialize output variables
        latitude = 0;
        longitude = 0;
        
        // Define the regular expression for latitude and longitude (with optional N/S/E/W directions)
        Regex regex = new Regex(@"([-+]?[0-9]*\.?[0-9]+)\s*[°]?\s*([NS]),?\s*([-+]?[0-9]*\.?[0-9]+)\s*[°]?\s*([EW])", RegexOptions.IgnoreCase);

        // Search for the latitude and longitude pattern in the input string
        Match match = regex.Match(input);

        if (match.Success)
        {
            // Extract the numeric part of latitude
            latitude = Math.Round( double.Parse(match.Groups[1].Value), 6);
            // If it's south (S), negate the latitude
            if (match.Groups[2].Value.ToUpper() == "S")
                latitude = -latitude;
            // Extract the numeric part of longitude
            longitude = Math.Round(double.Parse(match.Groups[3].Value), 6);
            // If it's west (W), negate the longitude
            if (match.Groups[4].Value.ToUpper() == "W")
                longitude = -longitude;

            return true;
        }

        // Return false if latitude and longitude are not found
        return false;
    }
    static string ExtractMaidenheadGrid(string input)
    {
        // Define the regular expression for Maidenhead grid locator (4 or 6 character grids)
        Regex regex = new Regex(@"\b([A-R]{2}\d{2}[A-X]{0,2})\b", RegexOptions.IgnoreCase);

        // Search for a match in the input string
        Match match = regex.Match(input);

        if (match.Success)
        {
            return match.Value.ToUpper(); // Return the Maidenhead grid in uppercase
        }

        return string.Empty; // Return an empty string if no match is found
    }

    static (double, double) MaidenheadToGPS(string maidenhead)
    {
        if (maidenhead.Length < 4 || maidenhead.Length % 2 != 0)
            throw new ArgumentException("Invalid Maidenhead grid format.");

        maidenhead = maidenhead.ToUpper();

        // Calculate the longitude
        int lonField = maidenhead[0] - 'A'; // First letter
        int lonSquare = maidenhead[2] - '0'; // First number
        int lonSubsquare = maidenhead.Length >= 6 ? maidenhead[4] - 'A' : 0; // Optional letter for sub-square

        // Calculate the latitude
        int latField = maidenhead[1] - 'A'; // Second letter
        int latSquare = maidenhead[3] - '0'; // Second number
        int latSubsquare = maidenhead.Length >= 6 ? maidenhead[5] - 'A' : 0; // Optional letter for sub-square

        // Convert Maidenhead to latitude and longitude
        double lon = -180.0 + (lonField * 20.0) + (lonSquare * 2.0) + (lonSubsquare * (2.0 / 24.0)) + (2.0 / 48.0);
        double lat = -90.0 + (latField * 10.0) + (latSquare * 1.0) + (latSubsquare * (1.0 / 24.0)) + (1.0 / 48.0);
        lat = Math.Round(lat,6);
        lon = Math.Round(lon,6);
        return (lat, lon);
    }
    static string fillFieldNum(string input)
    {
        // find the first 8 commas and append a field number to each occurence
        var ct = 1;
        var startPosition = 0;
        // Console.WriteLine("input before: "+input);
        while (ct < 9)
        {
            startPosition = input.IndexOf(",", startPosition);
            if (startPosition > 0)
            {
                input = input.Insert(startPosition, " "+ct);
                // Console.WriteLine("input during: "+input);
            }
            else { break; }
            ct++;
            startPosition=startPosition+3;
        }
        // Console.WriteLine("input after: "+input);
        return input.Trim(',');        
    }

}

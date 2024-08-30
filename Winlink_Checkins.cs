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
using System.Runtime.CompilerServices;
using System.Text;
using System.Text.Json.Serialization;
using System.Text.RegularExpressions;
using System.Globalization;
using System.Runtime.Intrinsics.X86;


class Winlink_Checkins
{
    public static void Main(string[] args)
    {
        // Get the start date and end date from the user.
        DateTime startDate = DateTime.Today;
        DateTime endDate = DateTime.Today;
        //DateTime date;
        bool isValid = false;
        string input;
        
        Console.WriteLine("Please enter the start date - must be within two weeks of today (yyyymmdd): ");
        startDate = GetValidDate();
        //startDate = startDate.ToUniversalTime();

        while (!isValid)
        {
            Console.WriteLine("Please enter the end date - must be within two weeks of today (yyyymmdd): ");
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
        // writeString variables to go in the output files
        StringBuilder netCheckinString = new StringBuilder();
        StringBuilder netAckString2 = new StringBuilder();
        StringBuilder bouncedString = new StringBuilder();
        StringBuilder duplicates = new StringBuilder();
        StringBuilder newCheckIns = new StringBuilder();
        StringBuilder csvString = new StringBuilder();
        StringBuilder mapString = new StringBuilder();
        mapString.Append("CallSign,Latitude,Longitude,Band,Mode\r\n");
        StringBuilder badFormatString = new StringBuilder();
        StringBuilder skippedString = new StringBuilder();
        StringBuilder removalString = new StringBuilder();
        string callSignPattern = @"\b\d{0,2}[A-Z]{1,2}\d{1,2}[A-Z]{1,6}\b";
        string testString = "";
        string rosterString = "";
        string bandStr = "";
        string modeStr = "";
        TextInfo textInfo = new CultureInfo("en-US", false).TextInfo;

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
        string locType = "";
        string latitude = "";
        string longitude = "";
        string xSource = "";
        
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
                    if (ckin > 0) 
                    {
                        ckin = fileText.IndexOf("0. HEADER", endHeader);
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
                        string checkIn = fileText.Substring(startPosition, len);
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
                        string checkIn = fileText.Substring(startPosition, len);
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
                            // get x-location information
                            var xLoc = fileText.IndexOf("X-LOCATION: ");
                            latitude =""; longitude = ""; locType =""; xSource ="";
                            if (xLoc > 0)
                            {
                                startPosition = xLoc+12;
                                endPosition = fileText.IndexOf(",", startPosition);
                                len = endPosition - startPosition;
                                if (len > 0) { latitude = fileText.Substring(startPosition, len); }
                                startPosition = endPosition+2;
                                endPosition = fileText.IndexOf(" ", startPosition);
                                len = endPosition - startPosition;
                                if (len > 0) { longitude = fileText.Substring(startPosition, len); }
                                // get location type
                                startPosition = endPosition+1;
                                endPosition = fileText.IndexOf(")", startPosition)+1;
                                len = endPosition - startPosition;
                                if (len > 0) { locType = fileText.Substring(startPosition, len); }
                                // Console.Write("Latitude ="+latitude+   "    Longitude = "+longitude+"    Type = " + locType+"\r\n");
                            }

                            // get x-Source if available
                            var xSrc = fileText.IndexOf("X-SOURCE: ");
                            if (xSrc > 0)
                            {
                                startPosition = xSrc +10;
                                endPosition = fileText.IndexOf("\r\n", startPosition);
                                len = endPosition - startPosition;
                                if (len > 0) { xSource = fileText.Substring(startPosition, len); }
                            }


                            // adjust for ICS 213
                            if (ics > 0)
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
                                if (ckinOffset > 0) { ckinOffset = 9; } else { ckinOffset = 13; }
                                startPosition = fileText.IndexOf("COMMENTS:")+ckinOffset;
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
                                Console.Write("Break at line 473ish. Press enter to close.");
                                input = Console.ReadLine();
                                break;
                            }

                            string checkIn = fileText.Substring(startPosition, len);
                            checkIn = checkIn.Replace("=20", "")
                                .Replace("=0A", "")
                                .Replace("=0", "")
                                .Replace("16. CONTACT INFO:", ",")
                                .Replace("  ", " ")
                                .Replace("  ", " ")
                                .Replace(", ", ",")
                                .Replace(", ", ",")
                                .Replace(" ,", ",")
                                .Replace("\n", ", ")
                                .Replace("\r", "")
                                .Replace("\"", "")
                                .Trim()
                                .Trim(',');
                            string msgField = checkIn.
                                Replace(" ,", ",")                                
                                .Trim()+",";

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
                                    csvString.Append(xSource+","+latitude+","+longitude+","+locType+","+msgField+"\r\n");

                                    // find the band if it's where it's supposed to be
                                    bandStr ="";
                                    len =0;
                                    // debug Console.Write("\r\nmsgField ="+msgField+"\r\n");
                                    startPosition = IndexOfNthSB(msgField, (char)44, 0, 6)+1;
                                    if (startPosition > 0) { endPosition = IndexOfNthSB(msgField, (char)44, 0, 7); len = endPosition-startPosition; }
                                    if (len > 0 && msgField.Length >= len) 
                                    {
                                        bandStr=  msgField.Substring(startPosition, len)
                                            .ToLower()
                                            .Replace(" ", "")
                                            .Replace(" meters", "m")
                                            .Replace("meters","m")
                                            .Replace("meter", "m")
                                            .Replace(" meter", "m")
                                            .Replace("vhf", "2m")
                                            .Replace("(", "")
                                            .Replace(")", "")
                                            ;
                                        if (bandStr.IndexOf("m") == -1 && bandStr != "telnet")
                                        {
                                            if (bandStr.IndexOf("ghz") == -1) { bandStr = bandStr+"m"; }
                                        }
                                        if (bandStr =="telnet") { bandStr = textInfo.ToTitleCase(bandStr.ToLower()); }
                                        bandCt++;
                                        switch (bandStr)
                                        {
                                            case "Telnet":
                                                break;
                                            case "telnet":
                                                bandStr = "Telnet";
                                                break;
                                            case "telnetm":
                                                bandStr = "Telnet";
                                                break;
                                            case "160m":
                                                break;
                                            case "80m":
                                                break;
                                            case "75m":
                                                bandStr = "80m";
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
                                            case "5ghz":
                                                bandStr="5cm";
                                                break;
                                            case "5cm":
                                                break;
                                            case "3cm":
                                                break;
                                            default:
                                                //Console.WriteLine("Band is not standard for "+messageID+"  "+checkIn+": "+bandStr+" - "+ msgField+ "\r\n");
                                                badFormatString.Append("Band is not standard for "+messageID+"  "+checkIn+": "+bandStr+" - |"+ msgField+ "|\r\n");
                                                break;
                                        }
                                    }
                                    // debug Console.Write("bandStr final=|"+bandStr+"|  \r\n");

                                    // find the mode if it's where it's supposed to be
                                    modeStr ="";
                                    len =0;
                                    // debug Console.Write("\r\nmsgField ="+msgField+"\r\n");
                                    startPosition = IndexOfNthSB(msgField, (char)44, 0, 7)+1;
                                    if (startPosition > 0) { endPosition = IndexOfNthSB(msgField, (char)44, 0, 8); len = endPosition-startPosition; }
                                    if (len > 0 && msgField.Length >= len)
                                    {
                                        modeStr=  msgField.Substring(startPosition, len)
                                            .ToUpper()
                                            .Replace("TELNET","SMTP")
                                            .Replace("AREDN", "MESH")
                                            .Replace("WINLINK", "")
                                            .Replace("(", "")
                                            .Replace(")", "")
                                            .Replace("-", " ")
                                            .Trim();
                                        //if (modeStr == "TELNET") { modeStr = "SMTP"; }
                                        // {
                                        //     if (modeStr.IndexOf("ghz") == -1) { modeStr = modeStr+"m"; }
                                        // }
                                        // if (modeStr =="telnet") { modeStr = textInfo.ToTitleCase(modeStr.ToLower()); }
                                        modeCt++;
                                    }
                                    switch (modeStr)
                                    {
                                        case "SMTP":
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
                                            //Console.WriteLine("Mode is not standard for "+messageID+"  "+checkIn+": "+modeStr+" - "+ msgField+ "\r\n");
                                            badFormatString.Append("Mode is not standard for "+messageID+"  "+checkIn+": "+modeStr+" - |"+ msgField+ "|\r\n");
                                            break;
                                    }


                                    // debug Console.Write("modeStr final=|"+modeStr+"|  \r\n");


                                    // add to mapString csv file if xloc was found
                                    if (latitude != "") { 
                                        mapString.Append(xSource+","+latitude+","+longitude+","+bandStr+","+modeStr+"\r\n");
                                        mapCt++;
                                    }
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
            logWrite.WriteLine("Total Plain and other Checkins: "+(ct-localWeatherCt-severeWeatherCt-incidentStatusCt-icsCt-ckinCt-damAssessCt-fieldSitCt-quickHWCt-dyfiCt-rriCt-qwmCt-miCt)+"\r\n");
            logWrite.WriteLine("Total Checkins with a geolocation: "+mapCt);
            logWrite.WriteLine("Total Checkins with something in the band field: "+bandCt);
            logWrite.WriteLine("Total Checkins with something in the mode field: "+modeCt);
            logWrite.WriteLine("msgField not properly formatted for the following: \r\n"+badFormatString);
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
}
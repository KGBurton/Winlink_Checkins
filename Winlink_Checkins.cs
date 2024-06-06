//c# code that will input start date, end date, and callSign and will select files with an extension of mime from the current folder  based on start date and end date, and will read each file to find a line labeled To: . If the rest of the line contains callSign, then write the data from the line labeled X-Source: to a text file called checkins.txt in the same folder
// Design Get the date range
// get the data source
// is it a message file? (.mime)
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
// save document info, writeLines, and count; write message type and count and message to CSV file, get next

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


// routine to extract the FCC CallSign from the MSG field string


namespace Winlink_Checkins
{
    class Winlink_Checkins
    {
        public static void Main(string[] args)
        {
            // Get the start date and end date from the user.
            DateTime startDate = DateTime.Today;
            DateTime endDate = DateTime.Today;
            bool isValid = false;
            string consoleString = "Enter the start date (yyyy-mm-dd):";
            Console.WriteLine(consoleString);
            string input = Console.ReadLine();
            if (DateTime.TryParse(input, out startDate)) ;
            consoleString = "Enter the end date (yyyy-mm-dd):";
            Console.WriteLine(consoleString);
            input = Console.ReadLine();
            if (DateTime.TryParse(input, out endDate)) ;
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
            StringBuilder skippedString = new StringBuilder();
            StringBuilder removalString = new StringBuilder();
            string callSignPattern = @"\b\d{0,2}[A-Z]{1,2}\d{1,2}[A-Z]{1,6}\b";
            string testString = "";
            string rosterString = "";


            if (File.Exists(rosterFile))
            {
                rosterString = File.ReadAllText(rosterFile);
                rosterString = rosterString.ToUpper();
                //debug Console.WriteLine("rosterFile contents: "+rosterString);
                var startPosition = 0;
                var endPosition = rosterString.IndexOf("\r\n", startPosition);
                var len = endPosition - startPosition;
                if (len > 0 ) 
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
            var removalCt = 0;
            var ackCt = 0;
            var junk = 0;

            // Select files with an extension of mime from the current folder.
            var files = Directory.GetFiles(currentFolder, "*.mime")
                .Where(file =>
                {
                    //  DateTime fileDate = File.GetCreationTime(file);
                    DateTime fileDate = File.GetLastWriteTime(file);
                    //debug Console.Write(fileDate+"\n");
                    return fileDate >= startDate && fileDate <= endDate;
                });
            
            Console.Write("\nMessages to process="+ files.Count() + " from folder "+currentFolder+"\n\n");

            // Create a text file called checkins.txt in the data folder and process the list of files.
            using (StreamWriter logWrite = new(Path.Combine(currentFolder, "checkins.txt")))
            // Create a text file called checkins.csv in the data folder and process the list of files.
            using (StreamWriter csvWrite = new(Path.Combine(currentFolder, "checkins.csv")))

            {
                // Read each file selected to find a line labeled To: and if the rest of the line contains netName, write the data from the line labeled X-Source: to the text file.
                
                foreach (string file in files)
                {
                    using (StreamReader reader = new StreamReader(file))
                    {
                        msgTotal++;
                        //debug Console.Write("File "+file+"\n");
                        string fileText = reader.ReadToEnd();
                        fileText = fileText.ToUpper();
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
                        var ics = fileText.IndexOf("ICS 213");

                        // check for winlink checkin message
                        var ckin = fileText.IndexOf("WINLINK CHECK-IN",endHeader);

                        // check for odd checkin message - don't let it scan through to a binary attachment!
                        
                        var lenBPQ = fileText.Length-10;
                        if (lenBPQ > 800) { lenBPQ = 800; }
                        var BPQ = fileText.IndexOf("BPQ",1,lenBPQ);

                        // check for damage assessment report
                        var damAssess = fileText.IndexOf("SURVEY REPORT - CATEGORIES");

                        // check for field situation report
                        var fieldSit = fileText.IndexOf("EMERGENT/LIFE SAFETY");
                        
                        // discard acknowledgements
                        if (ack >0)
                        { 
                            skipped++;
                            ackCt++;
                            junk=0; //debug Console.Write(file+" is an acknowedgement, skipping.");
                        }

                        else if (removal >0)
                        {
                            var startPosition = fileText.IndexOf("FROM:")+6;
                            var endPosition = fileText.IndexOf("\r\n", startPosition);
                            var len = endPosition - startPosition;
                            string checkIn = fileText.Substring(startPosition, len);
                            {
                                checkIn = checkIn.Replace(',', ' ');
                                // Create a Regex object with the pattern
                                Regex regexCallSign = new Regex(callSignPattern, RegexOptions.IgnoreCase);
                                // find the first callsign match in the checkIn string
                                Match match = regexCallSign.Match(checkIn);
                                if (match.Success) checkIn = match.Value;
                            }                           
                            removalString.Append("Message from: "+checkIn+" was a removal request.\r\n");
                            removalCt++;
                            junk = 0;  // debug Console.Write("Removal Request: "+file+", skipping.");
                        }
                        else if (bounced > 0)
                        {
                            var startPosition = bounced;
                            var endPosition = fileText.IndexOf("\r\n", startPosition);
                            var len = endPosition - startPosition;
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
                            // extended to include the TO: field in case they didn't the netName in the subject
                            var startPosition = fileText.IndexOf("SUBJECT:")+9;
                            var endPosition = fileText.IndexOf("MESSAGE-ID", startPosition);
                            var len = endPosition - startPosition;
                            string subjText = fileText.Substring(startPosition, len);

                            // deterimine if it was forwarded to know to look below the first header info

                            if (subjText.Contains(netName))
                            {
                                // adjust for ICS 213
                                if (ics > 0)
                                {
                                    // end of the header information as the start of the msg field
                                    startPosition = fileText.IndexOf("MESSAGE:")+15;
                                    endPosition = fileText.IndexOf("APPROVED BY:", startPosition)-3;
                                }

                                // adjust for winlink checkin
                                else if (ckin >0)                                                              
                                {
                                    // end of the header information as the start of the msg field
                                    // some people include WINLINK CHECK-IN in the subject which confuses the program
                                    // into thinking this is a winlink checkin FORM!! Catch it ...
                                    startPosition = fileText.IndexOf("COMMENTS:")+13;
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
                                    endPosition = fileText.IndexOf("----------", startPosition)-1;
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
                                if (len <= 0)
                                {
                                    Console.Write("endPostion is less than startPosition in: "+file+"\n");
                                    Console.Write("Break at line 279ish. Press enter to close.");
                                    input = Console.ReadLine();
                                    break;
                                }
                                string checkIn = fileText.Substring(startPosition, len);
                                checkIn = checkIn.Replace("=20","")
                                    .Replace("16. CONTACT INFO:",",")
                                    .Replace("\n", ",")
                                    .Replace("\r","")
                                    .Trim()
                                    .Trim(',');
                                string msgField = checkIn.Replace("\r", "");
                                    

                                // Create a Regex object with the pattern
                                Regex regexCallSign = new Regex(callSignPattern, RegexOptions.IgnoreCase);

                                // find the first callsign match in the checkIn string
                                Match match = regexCallSign.Match(checkIn);
                                if (match.Success)
                                {
                                    checkIn = match.Value;
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
                                            // checkIn =""; 
                                        }
                                    }
                                    junk = 0; // this represents the end of a normal message, place for a debug
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
                                    if (startPosition > 0)
                                    {
                                        if (dupCt == 0) { duplicates.Append("Duplicates: \r\n"); }
                                        //debug Console.Write("netName "+checkIn+" is a duplicate, skipping. It is "+dupCt+" of "+msgTotal+" total messages.\n");
                                        duplicates.Append(checkIn+", ");
                                        dupCt++;
                                    }
                                    else if (startPosition < 0)
                                    {
                                        ct++;
                                        testString = testString+checkIn;
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

                                        csvString.Append(msgField+"\r\n");
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
                                    Console.Write(checkIn+" was not found the roster. \n");
                                    newCheckIns.Append(checkIn+", ");
                                    // update roster.txt to contain the new checkin
                                    File.AppendAllText ("roster.txt", "; "+  checkIn);
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
                logWrite.WriteLine("Total Checkins Recorded:"+ct+"    Duplicates Skipped:"+dupCt+"    Removal Requests: "+removalCt+"    Non-"+netName+" checkin messages skipped:"+skipped+" (including "+ackCt+" acknowledgements)    Total messages processed:"+msgTotal);
                logWrite.WriteLine("Row 7 goes into the top of the checkin column to be recorded.");
                logWrite.WriteLine("Row 8 is the copy list for the checkin acknowledgement.");
                logWrite.WriteLine("Rows 9 & 10 have the list of duplicates found.");
                logWrite.WriteLine("Rows 11 and beyond are new checkins that should be added to \r\n        the spreadsheet, skipped messages that didn't have a netName, and other notifications.");
                logWrite.WriteLine(netCheckinString);
                logWrite.WriteLine(netAckString2);
                csvWrite.WriteLine(csvString);
                if (duplicates.Length==0)
                {
                    duplicates.Append("No duplicates found this week.");
                    Console.Write("No duplicates found this week..\n\n");
                }
                logWrite.WriteLine(duplicates+"\r\n");
                logWrite.WriteLine(bouncedString);
                if (newCt == 0)
                {
                    newCheckIns.Append("No new checkins found this week.");
                    Console.Write("No new checkins found this week.\n\n");
                }
                logWrite.Write(newCheckIns+"\r\n");
                logWrite.Write(skippedString+"\r\n");
                logWrite.Write(removalString+"\r\n");

            }
            Console.WriteLine("Done!\nThere were "+ct+" checkins. \nThe output checkins.txt can be found in the folder \n"+currentFolder);
            // Console.WriteLine("\nBe sure to update the roster.txt file if you receive new checkins. The app does not do that automatically yet.\n\n");
            Console.WriteLine("\n\nPress enter to continue.");
            Console.ReadLine();
        }
        

    }

}

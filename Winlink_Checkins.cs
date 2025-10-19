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

using Google.Apis.Auth.OAuth2;
using Google.Apis.Services;
using Google.Apis.Sheets.v4;
using Google.Apis.Sheets.v4.Data;
using System;
using System.Collections.Generic;
using System.ComponentModel.Design;
using System.Diagnostics.Metrics;
using System.Globalization;
using System.IO;
using System.Linq;
using System.Net;
using System.Net.Mail;
using System.Numerics;
using System.Reflection.Metadata;
using System.Text;
using System.Text.RegularExpressions;
using System.Xml.Linq;


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
        string? input = "";
        string weekDay = "";
        int netLength = 0;
        // string filter = args.Length > 1 ? args [1]! : "";
        string filter = args.Length > 1 ? args [1]!.ToUpper () : "";
        // DateTime endDate = default;


        Console.WriteLine ("How many days does the net last? (max of 10)");
        while (!isValid)
        {
            if (int.TryParse (Console.ReadLine () ?? "", out netLength) && netLength > 0 && netLength <= 12)
            {
                isValid = true;
            }
            else
            {
                Console.WriteLine ("Please enter a number between 1 and 10.");
            }
        }
        isValid = false;
        while (!isValid)
        {
            (startDate, endDate, weekDay) = getNetDates (startDate, endDate, weekDay, netLength); // Assuming netLength isn't needed yet
            Console.WriteLine ($"Your net begins on {weekDay}, {startDate} and ends on {endDate}. \r\nIs that correct? ('N' to try again or any other character to continue)");
            ConsoleKeyInfo keyPress = Console.ReadKey (true); // 'true' prevents displaying the pressed key
            char yesNo = keyPress.KeyChar;
            Console.WriteLine (); // Add newline after keypress for better formatting
            if (char.ToUpper (yesNo) != 'N')
            {
                isValid = true;
            }
            else
            {
                Console.WriteLine ("Please enter the dates again.");
            }
        }
        isValid = false;
        startDate = startDate.AddDays (-1);// the -1 will catch those that checked in a bit early
        endDate = endDate.AddDays (1); // the +1 will catch those that checked in a bit late
        // weekDay is the day the net started

        // Get the unique net identifier to screen only relevant messages from the folder
        // Console.WriteLine("Enter the unique net name for which the checkins are sent:");
        // string netName = Console.ReadLine();
        // Get the native call sign from the user to find the messages folder.
        string currentFolder = "";
        string applicationFolder = Directory.GetCurrentDirectory ();
        Console.WriteLine ("Enter YOUR call sign to find the messages folder. \r\n     If you leave it blank, the program will assume that it is already operating from the messages folder: \n\t" + applicationFolder);
        string? yourCallSign = Console.ReadLine ();

        // Get the data folder - either the global messages folder (default) or the current
        // operator's messages folder, assuming the default RMS installation location.


        if (yourCallSign != "")
        {
            currentFolder = "C:\\RMS Express\\" + yourCallSign + "\\Messages";
        }
        else
        {
            currentFolder = Directory.GetCurrentDirectory ();
        }
        string? netName = "";
        // Look for roster.txt in the folder. If it exists, get the first (and only)
        // row for comparison down below
        string rosterFile = applicationFolder + "\\roster.txt";
        // string attachmentFileCSV = applicationFolder + "\\attachments.csv";

        string xmlFile = applicationFolder + "\\Winlink_Import.xml"; // separate file for defective messages
        string xmlPerfFile = applicationFolder + "\\Winlink_Import_Perfect.xml"; // separate file for perfect messages
        // string commentFile = applicationFolder+"\\"+ netName +"_Additional_Comments.txt";
        // writeString variables to go in the output files
        // StringBuilder roster = new StringBuilder ();
        StringBuilder netCheckinString = new ();
        StringBuilder netAckString2 = new ();
        StringBuilder bouncedString = new ();
        StringBuilder duplicates = new ();
        StringBuilder newCheckIns = new ();
        StringBuilder csvString = new ();
        csvString.AppendLine ("Current " + netName + " Checkins, posted: " + utcDate);
        StringBuilder mapString = new ();
        mapString.Append ("CallSign,Latitude,Longitude,Band,Mode\r\n");
        //        StringBuilder badBandString = new ();
        //        StringBuilder badModeString = new ();
        StringBuilder skippedString = new ();
        StringBuilder removalString = new ();
        StringBuilder addonString = new ();
        StringBuilder noGPSString = new ();
        StringBuilder noScoreString = new ();
        StringBuilder typoString = new ();

        // string callSignPattern = @"\b\d{0,2}[A-Z]{1,2}\d{1,2}[A-Z]{1,6}\b";
        string countries = "AFGHANISTAN:AFG,ALBANIA:ALB,ALGERIA:DZA,ANDORRA:AND,ANGOLA:AGO,ANGUILLA:AIA,ANTARCTICA:ATA,ANTIGUA AND BARBUDA:ATG,ARGENTINA:ARG,ARMENIA:ARM,ARUBA:ABW,AUSTRALIA:AUS,AUSTRIA:AUT,AZERBAIJAN:AZE,BAHAMAS:BHS,BAHRAIN:BHR,BANGLADESH:BGD,BARBADOS:BRB,BELARUS:BLR,BELGIUM:BEL,BELIZE:BLZ,BENIN:BEN,BERMUDA:BMU,BHUTAN:BTN,BOLIVIA (PLURINATIONAL STATE OF):BOL,BONAIRE: SINT EUSTATIUS AND SABA:BES,BOSNIA AND HERZEGOVINA:BIH,BOTSWANA:BWA,BOUVET ISLAND:BVT,BRAZIL:BRA,BRITISH INDIAN OCEAN TERRITORY:IOT,BRUNEI DARUSSALAM:BRN,BULGARIA:BGR,BURKINA FASO:BFA,BURUNDI:BDI,CABO VERDE:CPV,CAMBODIA:KHM,CAMEROON:CMR,CANADA:CAN,CAYMAN ISLANDS:CYM,CENTRAL AFRICAN REPUBLIC:CAF,CHAD:TCD,CHILE:CHL,CHINA:CHN,CHRISTMAS ISLAND:CXR,COCOS (KEELING) ISLANDS:CCK,COLOMBIA:COL,COMOROS:COM,CONGO:COG,CONGO (DEMOCRATIC REPUBLIC OF THE):COD,COOK ISLANDS:COK,COSTA RICA:CRI,CROATIA:HRV,CUBA:CUB,CURAÇAO:CUW,CYPRUS:CYP,CZECHIA:CZE,CÔTE D'IVOIRE:CIV,DENMARK:DNK,DJIBOUTI:DJI,DOMINICA:DMA,DOMINICAN REPUBLIC:DOM,ECUADOR:ECU,EGYPT:EGY,EL SALVADOR:SLV,EQUATORIAL GUINEA:GNQ,ERITREA:ERI,ESTONIA:EST,ESWATINI:SWZ,ETHIOPIA:ETH,FALKLAND ISLANDS (MALVINAS):FLK,FAROE ISLANDS:FRO,FIJI:FJI,FINLAND:FIN,FRANCE:FRA,FRENCH GUIANA:GUF,FRENCH POLYNESIA:PYF,FRENCH SOUTHERN TERRITORIES:ATF,GABON:GAB,GAMBIA:GMB,GEORGIA:GEO,GERMANY:DEU,GHANA:GHA,GIBRALTAR:GIB,GREECE:GRC,GREENLAND:GRL,GRENADA:GRD,GUADELOUPE:GLP,GUAM:GUM,GUATEMALA:GTM,GUERNSEY:GGY,GUINEA:GIN,GUINEA-BISSAU:GNB,GUYANA:GUY,HAITI:HTI,HEARD ISLAND AND MCDONALD ISLANDS:HMD,HOLY SEE:VAT,HONDURAS:HND,HONG KONG:HKG,HUNGARY:HUN,ICELAND:ISL,INDIA:IND,INDONESIA:IDN,IRAN (ISLAMIC REPUBLIC OF):IRN,IRAQ:IRQ,IRELAND:IRL,ISLE OF MAN:IMN,ISRAEL:ISR,ITALY:ITA,JAMAICA:JAM,JAPAN:JPN,JERSEY:JEY,JORDAN:JOR,KAZAKHSTAN:KAZ,KENYA:KEN,KIRIBATI:KIR,KOREA (DEMOCRATIC PEOPLE'S REPUBLIC OF):PRK,KOREA (REPUBLIC OF):KOR,KUWAIT:KWT,KYRGYZSTAN:KGZ,LAO PEOPLE'S DEMOCRATIC REPUBLIC:LAO,LATVIA:LVA,LEBANON:LBN,LESOTHO:LSO,LIBERIA:LBR,LIBYA:LBY,LIECHTENSTEIN:LIE,LITHUANIA:LTU,LUXEMBOURG:LUX,MACAO:MAC,MADAGASCAR:MDG,MALAWI:MWI,MALAYSIA:MYS,MALDIVES:MDV,MALI:MLI,MALTA:MLT,MARSHALL ISLANDS:MHL,MARTINIQUE:MTQ,MAURITANIA:MRT,MAURITIUS:MUS,MAYOTTE:MYT,MEXICO:MEX,MICRONESIA (FEDERATED STATES OF):FSM,MOLDOVA (REPUBLIC OF):MDA,MONACO:MCO,MONGOLIA:MNG,MONTENEGRO:MNE,MONTSERRAT:MSR,MOROCCO:MAR,MOZAMBIQUE:MOZ,MYANMAR:MMR,NAMIBIA:NAM,NAURU:NRU,NEPAL:NPL,NETHERLANDS:NLD,NEW CALEDONIA:NCL,NEW ZEALAND:NZL,NICARAGUA:NIC,NIGER:NER,NIGERIA:NGA,NIUE:NIU,NORFOLK ISLAND:NFK,NORTH MACEDONIA:MKD,NORTHERN MARIANA ISLANDS:MNP,NORWAY:NOR,OMAN:OMN,PAKISTAN:PAK,PALAU:PLW,PALESTINE: STATE OF:PSE,PANAMA:PAN,PAPUA NEW GUINEA:PNG,PARAGUAY:PRY,PERU:PER,PHILIPPINES:PHL,PITCAIRN:PCN,POLAND:POL,PORTUGAL:PRT,PUERTO RICO:PRI,QATAR:QAT,ROMANIA:ROU,RUSSIAN FEDERATION:RUS,RWANDA:RWA,RÉUNION:REU,SAINT BARTHÉLEMY:BLM,SAINT HELENA: ASCENSION AND TRISTAN DA CUNHA:SHN,SAINT KITTS AND NEVIS:KNA,SAINT LUCIA:LCA,SAINT MARTIN (FRENCH PART):MAF,SAINT PIERRE AND MIQUELON:SPM,SAINT VINCENT AND THE GRENADINES:VCT,SAMOA:WSM,SAN MARINO:SMR,SAO TOME AND PRINCIPE:STP,SAUDI ARABIA:SAU,SENEGAL:SEN,SERBIA:SRB,SEYCHELLES:SYC,SIERRA LEONE:SLE,SINGAPORE:SGP,SINT MAARTEN (DUTCH PART):SXM,SLOVAKIA:SVK,SLOVENIA:SVN,SOLOMON ISLANDS:SLB,SOMALIA:SOM,SOUTH AFRICA:ZAF,SOUTH GEORGIA AND THE SOUTH SANDWICH ISLANDS:SGS,SOUTH SUDAN:SSD,SPAIN:ESP,SRI LANKA:LKA,SUDAN:SDN,SURINAME:SUR,SVALBARD AND JAN MAYEN:SJM,SWEDEN:SWE,SWITZERLAND:CHE,SYRIAN ARAB REPUBLIC:SYR,TAIWAN: PROVINCE OF CHINA:TWN,TAJIKISTAN:TJK,TANZANIA (UNITED REPUBLIC OF):TZA,THAILAND:THA,TIMOR-LESTE:TLS,TOGO:TGO,TOKELAU:TKL,TONGA:TON,TRINIDAD AND TOBAGO:TTO,TUNISIA:TUN,TÜRKIYE:TUR,TURKMENISTAN:TKM,TURKS AND CAICOS ISLANDS:TCA,TUVALU:TUV,UGANDA:UGA,UKRAINE:UKR,UNITED ARAB EMIRATES:ARE,UNITED KINGDOM OF GREAT BRITAIN AND NORTHERN IRELAND:GBR,UNITED STATES OF AMERICA:USA,UNITED STATES MINOR OUTLYING ISLANDS:UMI,URUGUAY:URY,UZBEKISTAN:UZB,VANUATU:VUT,VENEZUELA (BOLIVARIAN REPUBLIC OF):VEN,VIET NAM:VNM,VIRGIN ISLANDS (BRITISH):VGB,VIRGIN ISLANDS (U.S.):VIR,WALLIS AND FUTUNA:WLF,WESTERN SAHARA:ESH,YEMEN:YEM,ZAMBIA:ZMB,ZIMBABWE:ZWE,ÅLAND ISLANDS:ALA";
        string testString = "";
        string rosterString = "";
        string roster = "";
        string credentialFilename = "";
        string spreadsheetId = "";
        string bandStr = "";
        string modeStr = "";
        // string noGPSStr = "";
        string? checkIn = "";
        string? msgField = "";
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
        string? xSource = "";
        string? fromTxt = "";
        string? tempFromTxt = "";
        string? tempCheckIn = "";
        string? []? checkinItems = new string? [] { }; // Initialize as an empty array
        string newCheckIn = "";
        string base64String = "";
        string attachmentDecodedString = string.Empty;
        string modeTypo = "";

        // string w3wText = "";

        addonString.AppendLine ("\r\nComments from the Current Checkins Posted\t" + utcDate + "\r\n-------------------------------");
        noGPSString.AppendLine ("\r\n++++++++\r\nThese had neither GPS data nor Maidenhead Grids\r\n-------------------------");
        noScoreString.AppendLine ("\r\n++++++++\r\nThese chose not to be scored:");
        Random rnd = new Random ();

        int startPosition = 0;
        int endPosition = 0;
        int quotedPrintable = 0;
        int lastBoundary = 0;
        // int lineStart = -1;
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
        int APRS = -1;
        int js8call = 0;
        int PosRepCt = 0;
        int copyPR = -1;
        int ICS201Ct = 0;
        int ICS202Ct = 0;
        int ICS203Ct = 0;
        int ICS204Ct = 0;
        int ICS205Ct = 0;
        int ICS205aCt = 0;
        int ICS206Ct = 0;
        int ICS208Ct = 0;
        int ICS210Ct = 0;
        int WBBMct = 0;

        int exerciseCompleteCt = 0;
        int radioGram = 0;
        int radioGramCt = 0;
        int winlinkCt = 0;
        int patCt = 0;
        int woadCt = 0;
        int airmailCt = 0;
        int radioMailCt = 0;
        int w3w = 0;
        int newFormatCt = 0;
        int examplePosition = -1;
        int attachmentCSVct = 0;
        int found = 0;
        var firstPipe = 0;

        double latitude = 0;
        double longitude = 0;
        bool isPerfect = true;
        bool newFormat = false;
        bool newFormatEndOnly = false;
        bool newFormatStartOnly = false;
        bool newFormatPipeOnly = false;
        bool newFormatSingleOnly = false;
        bool newFormatNoPipe = false;
        bool onlyOneMarker = false;
        bool exampleIncluded = false;
        bool brokenCheckin = true;
        bool longCountry = false;
        bool pipeDelimiter = false;

        TextInfo textInfo = new CultureInfo ("en-US", false).TextInfo;
        // Create root XML document

        // initialize xmlDoc with a root
        XDocument xmlDoc = new XDocument (new XElement ("WinlinkMessages"));
        XDocument xmlPerfDoc = new XDocument (new XElement ("WinlinkMessages"));

        XElement messageElement = new XElement
            ("export_parameters",
                new XElement ("xml_file_version", "1.0"),
                new XElement ("winlink_express_version", "1.7.17.0"),
                new XElement ("callsign", netName ?? "")
            );
        xmlDoc.Root!.Add (messageElement);
        xmlPerfDoc.Root!.Add (messageElement);

        messageElement = new XElement ("message_list", "");
        xmlDoc.Root.Add (messageElement);
        xmlPerfDoc.Root.Add (messageElement);

        if (File.Exists (rosterFile))
        {
            rosterString = File.ReadAllText (rosterFile);
            // rosterString = rosterString.ToUpper (); // this trashes the spreadsheetID and the credentials file so do it later with just the pieces
            //debug Console.WriteLine("rosterFile contents: "+rosterString);
            // get the net name from the roster.txt file
            startPosition = rosterString.IndexOf ("NETNAME=", StringComparison.OrdinalIgnoreCase);
            if (startPosition > -1) { startPosition += 8; }
            endPosition = rosterString.IndexOf ("//", startPosition);
            len = endPosition - startPosition;
            if (len > 0)
            { netName = rosterString.Substring (startPosition, len).Trim ().ToUpper (); }
            else { netName = "GLAWN"; }

            // get the x-source name from the roster.txt file to be used as the netName variable in the xml file
            startPosition = rosterString.IndexOf ("CALLSIGN=", StringComparison.OrdinalIgnoreCase);
            if (startPosition > -1) { startPosition += 9; }
            endPosition = rosterString.IndexOf ("//", startPosition);
            len = endPosition - startPosition;
            if (len > 0)
            { xmlXsource = rosterString.Substring (startPosition, len).Trim ().ToUpper (); }
            else
            {
                Console.WriteLine ("callSign missing from the roster.txt file. X-SOURCE in the xml file will be wrong.");
                xmlXsource = "KB7WHO"; // default to mine if not found
            }

            // get the id of the spreadsheet used as a database to be opened for updating
            startPosition = rosterString.IndexOf ("google spreadsheet id=", StringComparison.OrdinalIgnoreCase);
            if (startPosition > -1)
            {
                startPosition += 22;
                endPosition = rosterString.IndexOf ("//", startPosition);
                len = endPosition - startPosition;
                if (len > 0) { spreadsheetId = rosterString.Substring (startPosition, len).Trim (); }
            }
            else
            {
                Console.WriteLine ("spreadsheetId is missing from the roster.txt file. X-SOURCE in the xml file will be wrong.");
                spreadsheetId = "1e0PJVqMGZhTzxwIVDf9if1dSSnG8y1U5Zf6pojB5Txc"; // Use my default spreadsheetID

            }

            // get the name of credentials file to be used to open the spreadsheet
            startPosition = rosterString.IndexOf ("credential filename=", StringComparison.OrdinalIgnoreCase);
            if (startPosition > -1)
            {
                startPosition += 20;
                endPosition = rosterString.IndexOf ("//", startPosition);
                len = endPosition - startPosition;
                if (len > 0) { credentialFilename = rosterString.Substring (startPosition, len).Trim (); }
            }
            else
            {
                Console.WriteLine ("spreadsheetId is missing from the roster.txt file. X-SOURCE in the xml file will be wrong.");
                credentialFilename = "credentials.json"; // default credential filename
            }

            // get the checkin roster from the roster.txt file
            startPosition = rosterString.IndexOf ("roster string=", StringComparison.OrdinalIgnoreCase);
            if (startPosition > -1)
            {
                startPosition += 14;
                // endPosition = rosterString.IndexOf("\r\n", startPosition);
                len = rosterString.Length - startPosition;
                if (len > 0)
                {
                    roster = rosterString.Substring (startPosition, len).Trim ().ToUpper ();
                    //roster = SortCommaDelimitedString (roster, ";");
                }
            }
        }
        else
        {
            //Console.WriteLine (currentFolder + "\\" + rosterFile + " \n was not found! A new one will be created. \r\n"
            //    + "All checkins will appear to be new.\n\n"
            //    + "Enter the name of the net you are checking in:");
            //input = Console.ReadLine ();
            //if (string.IsNullOrWhiteSpace (input))
            //{
            //    Console.WriteLine ("The net name is required. Please enter the name of the net for which Winlink Checkins will be used.");
            //}
            //else netName = input.ToUpper ();
            Console.WriteLine ($"{currentFolder}\\{rosterFile} \n was not found! A new one will be created. \r\n" +
    "All checkins will appear to be new.\n\n" +
    "Enter the name of the net you are checking in:");
            isValid = false; // Reuse existing isValid
            while (!isValid)
            {
                input = Console.ReadLine (); // Reuse existing input
                if (string.IsNullOrWhiteSpace (input))
                {
                    Console.WriteLine ("The net name is required. Please enter the name of the net for which Winlink Checkins will be used.");
                }
                else
                {
                    netName = input.ToUpper (); // Reuse existing netName
                    isValid = true;
                }
            }

            // Get the xmlXsource callsign
            Console.WriteLine ("Enter the callsign to use as the xSource for the personalized messages:");
            input = Console.ReadLine ();
            if (string.IsNullOrWhiteSpace (input))
            {
                Console.WriteLine ("This is required. Please enter the callsign to use as the xSource for the personalized messages:");
            }
            else
            {
                xmlXsource = input.ToUpper ();
            }

            // File.Create (rosterFile);

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

        Console.Write ("\nMessages to process=" + files.Count () + " from folder \r\n" + currentFolder + "\n\n");

        // Create a text file called checkins.txt in the data folder and process the list of files.
        using (StreamWriter logWrite = new (Path.Combine (currentFolder, "checkins.txt")))
        // Create a csv file called checkins.csv in the data folder and process the list of files.
        using (StreamWriter csvWrite = new (Path.Combine (currentFolder, "checkins.csv")))
        // from grok using (StreamWriter csvWrite = new StreamWriter (filePath, append: true))
        // Create a csv text file called mapfile.csv in the data folder to use as date for google maps
        using (StreamWriter mapWrite = new (Path.Combine (currentFolder, "mapfile.csv")))
        // Create a text file called Additional Comments.txt in the data folder 
        using (StreamWriter commentWrite = new (Path.Combine (currentFolder, netName + " Additional Comments.txt")))
        // Create a csv file called attachments.csv in the data folder and process the list of files.
        using (StreamWriter attachmentCSVwrite = new (Path.Combine (currentFolder, "attachments.csv")))

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
                    msgField = "";
                    checkIn = "";
                    callSignTypo = "";
                    base64String = "";
                    maidenheadGrid = "";
                    pipeDelimiter = false;

                    //debug Console.Write("File "+file+"\n");
                    string fileText = reader.ReadToEnd ();
                    string fileTextOriginal = fileText;
                    fileText = fileText.ToUpper ()
                        .Replace ("NO SCORE", "NOSCORE")
                        .Replace ("NO SUMMARY", "NOSUMMARY")
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

                    // does it have a CSV attachment?
                    startPosition = fileTextOriginal.IndexOf ("Content-Disposition: attachment;", StringComparison.OrdinalIgnoreCase);
                    if (startPosition > -1) // this is a good spot to catch the beginning of the record processing
                    {
                        startPosition = fileTextOriginal.IndexOf (".csv\"", startPosition);
                        // find the attachment
                        if (startPosition > -1) startPosition = fileTextOriginal.IndexOf ("Content-Transfer-Encoding: base64", startPosition, StringComparison.OrdinalIgnoreCase) + 37;
                        if (startPosition > -1)
                        {
                            endPosition = fileTextOriginal.IndexOf ("--boundary", startPosition, StringComparison.OrdinalIgnoreCase) - 2;
                            len = endPosition - startPosition;
                            if (len > 0)
                            {
                                base64String = fileTextOriginal.Substring (startPosition, len).Trim ();
                                // remove invalid characters (newlines that winlink express throws in)
                                base64String = Regex.Replace (base64String, @"\s+|\r\n|\n|\r", "");

                            }
                            try
                            {
                                attachmentDecoded = Convert.FromBase64String (base64String);
                                if (attachmentDecoded.Length > 0) // Removed redundant != null check
                                {
                                    // Write the decoded bytes to a CSV file
                                    string outputFilePath = currentFolder + "\\attachment.csv";
                                    // File.WriteAllBytes(outputFilePath, attachmentDecoded);
                                    // Console.WriteLine("Base64 string decoded successfully. CSV saved to: " + outputFilePath);

                                    // Print the decoded content as text
                                    attachmentDecodedString = System.Text.Encoding.UTF8.GetString (attachmentDecoded);
                                    // Console.WriteLine("Decoded CSV content:\n" + attachmentDecodedString);
                                }
                                else
                                {
                                    Console.WriteLine ("attachmentDecoded is empty.");
                                }
                            }
                            catch (FormatException ex)
                            {
                                Console.WriteLine ("Invalid Base64 string: " + ex.Message);
                                attachmentDecoded = Array.Empty<byte> (); // Reset to empty array on failure
                            }
                            catch (Exception ex)
                            {
                                Console.WriteLine ("Error decoding Base64 string: " + ex.Message);
                                attachmentDecoded = Array.Empty<byte> (); // Reset to empty array on failure
                            }
                        }
                    }

                    // was it APRSmail?
                    junk = fileText.IndexOf ("APRSEMAIL2");
                    junk = fileText.IndexOf ("APRS.EARTH");
                    junk = fileText.IndexOf ("APRS.FI");
                    // if (fileText.IndexOf ("APRSEMAIL2") > -1 || fileText.IndexOf ("APRS.EARTH") > -1 || fileText.IndexOf ("APRS.FI") > -1)
                    if (fileText.IndexOf ("APRSEMAIL2") > -1 || fileText.IndexOf ("APRS.EARTH") > -1)
                    {
                        APRS = 1;
                        startPosition = fileText.IndexOf ("SUBJECT:"); // if it is APRS, the first From: is going to be the APRS server
                        aprsCt++;
                    }
                    else
                    {
                        APRS = -1;
                        startPosition = 0;
                    }

                    // get From:
                    startPosition = fileText.IndexOf ("FROM:", startPosition);

                    if (startPosition > -1) { startPosition += 6; }
                    endPosition = fileText.IndexOf ("@", startPosition);// skip any domain info
                    if (endPosition == -1) endPosition = fileText.IndexOf ("\r\n", startPosition); // otherwise the end of the line
                    len = endPosition - startPosition;
                    fromTxt = fileText.Substring (startPosition, len);
                    fromTxt = fromTxt
                        .Replace ("=20", "")
                        .Replace (" ", "")
                        .Replace (',', ' ');
                    tempFromTxt = isValidCallsign (fromTxt);
                    if (tempFromTxt == "")
                    {
                        Console.WriteLine ("576 Invalid callsign in the From: field =>" + fromTxt + "<= of messageID" + messageID);
                    }
                    else fromTxt = tempFromTxt;

                    // find the end of the header section
                    var endHeader = fileText.IndexOf ("CONTENT-TRANSFER-ENCODING:");
                    // find the likely start of the message
                    quotedPrintable = fileText.IndexOf ("QUOTED-PRINTABLE");

                    if (quotedPrintable > -1) quotedPrintable += 20;
                    lastBoundary = fileText.IndexOf ("--BOUNDARY", quotedPrintable);
                    commentPos = fileText.IndexOf ("COMMENT:", quotedPrintable);
                    if (commentPos > -1) commentPos += 9;

                    // does the sender want to skip the scoring of the message
                    var noScore = fileText.IndexOf ("NOSCORE");
                    var noSummary = fileText.IndexOf ("NOSUMMARY");


                    // deterimine if it was forwarded to know to look below the first header info
                    var forwarded = fileText.IndexOf ("WAS FORWARDED BY");

                    // was it JS8CALL
                    js8call = fileText.IndexOf ("JS8CALL");
                    if (js8call > -1) js8ct++;

                    // check for acknowledgement message and discard later                  
                    int ack = fileText.IndexOf ("[MESSAGE ACKNOWLEDGEMENT]");

                    // check for ICS 213 msg
                    var ics = fileText.IndexOf ("TEMPLATE VERSION: ICS 213");
                    if (ics > -1) icsCt++;


                    // check for winlink checkin message
                    var winlinkCkin = fileText.IndexOf ("MAP FILE NAME: WINLINK CHECK", endHeader);
                    // some people include WINLINK CHECK-IN in the subject which confuses the program
                    // into thinking this is a winlink checkin FORM!! Catch it ...
                    if (winlinkCkin < 0) winlinkCkin = fileText.IndexOf ("WINLINK CHECK-IN 5.0", endHeader);
                    if (winlinkCkin < 0) winlinkCkin = fileText.IndexOf ("WINLINK CHECK-IN \r\n0. HEADER", endHeader);
                    if (winlinkCkin < 0) winlinkCkin = fileText.IndexOf ("WINLINK CHECK IN 2.", endHeader);
                    if (winlinkCkin > -1) winlinkCkinCt++;

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
                            // the QTH message was copied to netName instead of forwarding the responses
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

                    // check for removal message               
                    var removal = fileText.IndexOf ("REMOVE ME");

                    // look to see if it was a bounced message
                    var bounced = fileText.IndexOf ("UNDELIVERABLE");

                    // check for local weather report
                    var localWeather = fileText.IndexOf ("CURRENT LOCAL WEATHER CONDITIONS");
                    if (localWeather > -1) localWeatherCt++;

                    // check for severe weather
                    var severeWeather = fileText.IndexOf ("SEVERE WX REPORT");
                    if (severeWeather > -1) severeWeatherCt++;

                    // check incident status report
                    var incidentStatus = fileText.IndexOf ("INCIDENT STATUS");
                    if (incidentStatus > -1) incidentStatusCt++;

                    // check for odd checkin message - don't let it scan through to a binary attachment!
                    //var lenBPQ = fileText.Length - 10;
                    //if (lenBPQ > 800)  lenBPQ = 800 - quotedPrintable; 
                    //testc = fileText.IndexOf("BPQ",quotedPrintable,)
                    len = lastBoundary - quotedPrintable;
                    if (len < 0) len = 0;
                    var BPQ = fileText.IndexOf ("BPQ", quotedPrintable, len);
                    var BPQAPRS = fileText.IndexOf ("BPQAPRS", quotedPrintable, len);// if this is a mode type, ignore the problem
                    if (BPQ == BPQAPRS) BPQ = -1;

                    // check for damage assessment report
                    var damAssess = fileText.IndexOf ("SURVEY REPORT - CATEGORIES");
                    if (damAssess > -1) damAssessCt++;

                    // check for field situation report
                    var fieldSit = fileText.IndexOf ("EMERGENT/LIFE SAFETY");
                    if (fieldSit > -1) fieldSitCt++;

                    // check for Quick Health & Welfare report, doesn't exist anymore? 10/2024
                    var quickM = fileText.IndexOf ("\r\nFROM ", endHeader);
                    if (quickM > -1) quickMCt++;

                    // check for RRI Quick Welfare Message
                    var qwm = fileText.IndexOf ("TEMPLATE VERSION: QUICK WELFARE MESSAGE");
                    if (qwm > -1) qwmCt++;

                    // check for RRI Welfare Radiogram
                    var rriWR = fileText.IndexOf ("TEMPLATE VERSION: RRI WELFARE RADIOGRAM");
                    if (rriWR > -1) rriCt++;

                    // check for Did You Feel It report
                    var dyfi = fileText.IndexOf ("DYFI WINLINK");
                    if (dyfi > -1) dyfiCt++;

                    // check for Medical Incident Report
                    var mi = fileText.IndexOf ("INITIAL PATIENT ASSESSMENT");
                    if (mi > -1) miCt++;

                    // check for ICS-201
                    var ICS201 = fileText.IndexOf ("ICS 201 INCIDENT BRIEFING");
                    if (ICS201 > -1) ICS201Ct++;

                    // check for ICS-202
                    var ICS202 = fileText.IndexOf ("ICS202_INCIDENT_OBJECTIVES");
                    if (ICS202 > -1) ICS202Ct++;

                    // check for ICS-203
                    var ICS203 = fileText.IndexOf ("ICS 203 ORGANIZATIONAL ASSIGNMENTS");
                    if (ICS203 > -1) ICS203Ct++;

                    // check for ICS-204
                    var ICS204 = fileText.IndexOf ("ICS 204 ASSIGNMENT LIST");
                    if (ICS204 > -1) ICS204Ct++;

                    // check for ICS-205
                    var ICS205 = fileText.IndexOf ("ICS 205");
                    if (ICS205 > -1) ICS205Ct++;

                    // check for ICS-205a
                    var ICS205a = fileText.IndexOf ("ICS 205A");
                    if (ICS205a > -1)
                    {
                        // var curTemp = fileText.IndexOf ("ASSIGNMENT:");
                        fileText = fileText.Replace ("\r\n\r\nASSIGNMENT:   NAME:   \r\nMETHOD: ", "", StringComparison.OrdinalIgnoreCase);
                        fileText = fileText.Replace ("\r\nMETHOD:", "", StringComparison.OrdinalIgnoreCase);
                        // fileText = fileText.Replace ("NAME: ", "", StringComparison.OrdinalIgnoreCase);
                        // fileTextOriginal.Replace ("\r\nASSIGNMENT:   NAME:  =20\r\nMETHOD:=20", "");
                        ICS205aCt++;
                    }

                    // check for ICS-206
                    var ICS206 = fileText.IndexOf ("ICS 206");
                    if (ICS206 > -1) ICS206Ct++;

                    // check for ICS-208
                    var ICS208 = fileText.IndexOf ("ICS 208 SAFETY MESSAGE-PLAN");
                    if (ICS208 > -1) ICS208Ct++;

                    // check for ICS-210
                    var ICS210 = fileText.IndexOf ("ICS 210");
                    if (ICS210 > -1) ICS210Ct++;

                    // check for  Welfare Bulletin Board Message
                    WBBMct = fileText.IndexOf ("WELFARE BULLETIN BOARD MESSAGE");
                    if (WBBMct > -1) WBBMct++;

                    // check for WhatThreeWords
                    if (fileText.IndexOf ("///") > -1 || fileText.IndexOf ("W3W") > -1 || fileText.IndexOf ("WHAT 3 WORDS") > -1 || fileText.IndexOf ("WHAT3WORDS") > -1 || fileText.IndexOf ("WHATTHREEWORDS") > -1 || fileText.IndexOf ("UTMREF") > -1)
                    {
                        w3w++;
                    }


                    // check for Radiogram
                    radioGram = fileText.IndexOf ("\r\nAR \r\n");

                    // screen dates to eliminate file dates that are different from the sent date and fall outside the net span
                    int startDateCompare = DateTime.Compare (sentDateUni, startDate);
                    int endDateCompare = DateTime.Compare (sentDateUni, endDate);

                    // catch removals first
                    if (removal > 0)
                    {

                        if (fromTxt != null) fromTxt = fromTxt.Trim ().TrimEnd ('\r', '\n'); // Clean fromTxt to strip any whitespace or newlines (e.g., "W0JW\n" -> "W0JW")
                        removalString.AppendLine (fromTxt + "\tin " + messageID + " was a removal request.");
                        removalCt++;
                        // Remove callsign from roster (string)
                        if (!string.IsNullOrEmpty (roster) && roster.Contains ($";{fromTxt};"))
                        {
                            roster = roster.Replace ($";{fromTxt};", ";"); // middle of the string
                        }
                        else if (!string.IsNullOrEmpty (roster) && roster.StartsWith ($"{fromTxt};"))
                        {
                            roster = roster.Replace ($"{fromTxt};", ""); // beginning of the string
                        }
                        else if (!string.IsNullOrEmpty (roster) && roster.EndsWith ($";{fromTxt}"))
                        {
                            roster = roster.Replace ($";{fromTxt}", ""); // end of the string
                        }
                        else if (roster == fromTxt)
                        {
                            roster = ""; // it was the only one in the string
                        }
                        Console.WriteLine ($"Removed '{fromTxt}' from roster.");
                        // junk = 0;  // debug Console.Write("Removal Request: "+file+", skipping.");
                    }

                    // discard acknowledgements
                    else if (ack > 0)
                    {
                        skipped++;

                        ackCt++;
                        // Console.Write (messageID + " Acknowledgement discarded\r\n");
                        oldSkipped = skipped;
                        // junk = 0; //debug Console.Write(file+" is an acknowedgement, skipping.");
                        skippedString.Append ("\tAcknowledgement from " + fromTxt + " discarded. Message ID: " + messageID + "\r\n");
                    }
                    else if (startDateCompare < 0 || endDateCompare > 0)
                    {
                        skipped++;
                        outOfRangeCt++;
                        oldSkipped = skipped;
                        Console.Write (messageID + " sendDate fell outside the start\\end dates\r\n");
                        skippedString.Append ("\tOut of date range: " + messageID + "\r\n");
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
                            tempCheckIn = isValidCallsign (checkIn);
                            if (tempCheckIn == "")
                            { Console.WriteLine ("Invalid callsign " + tempCheckIn + " in checkIn: " + messageID); }
                            else checkIn = tempCheckIn;
                        }
                        bouncedString.Append ("Message to: " + checkIn + " was not deliverable.\r\n");
                        skipped++;
                    }
                    else
                    {
                        // determine if the message has something in the subject to do with netName
                        // extended to include the TO: field in case they didn't put the netName in the subject
                        startPosition = fileText.IndexOf ("SUBJECT:");
                        if (startPosition > -1) { startPosition += 9; }
                        endPosition = fileText.IndexOf ("CC:", startPosition);
                        if (endPosition == -1) endPosition = fileText.IndexOf ("MESSAGE-ID", startPosition);
                        if (endPosition > 0) len = endPosition - startPosition;
                        string subjText = fileText.Substring (startPosition, len); // includes the TO: and CC: fields to find the netName

                        // if (subjText.Contains (netName))

                        if (fileText.Contains (netName!))
                        {
                            score = 10;
                            isPerfect = true;
                            newFormat = false;
                            newFormatEndOnly = false;
                            newFormatStartOnly = false;
                            newFormatPipeOnly = false;
                            newFormatSingleOnly = false;
                            newFormatNoPipe = false;
                            onlyOneMarker = false;
                            pointsOff = "";
                            checkinName = "";
                            checkinCity = "";
                            checkinCounty = "";
                            checkinState = "";
                            checkinCountry = "";
                            checkinCountryLong = "";
                            bandStr = "";
                            modeStr = "";
                            examplePosition = fileText.IndexOf ("XXNXXX");
                            if (examplePosition > -1)
                            {
                                exampleIncluded = true;
                            }
                            else exampleIncluded = false;
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

                            // Does the message have the new format starting and ending with ##
                            startPosition = fileText.IndexOf ("##");
                            endPosition = startPosition;
                            if (exampleIncluded)
                            {
                                if (startPosition > -1) startPosition = fileText.IndexOf ("##", startPosition + 2);
                                if (startPosition == -1) startPosition = fileText.IndexOf ("\r\n", endPosition) + 2;
                            }

                            endPosition = fileText.IndexOf ("##", startPosition + 2);
                            if (startPosition > -1 && endPosition >= startPosition) newFormat = true;
                            else
                            {
                                // check to see that it really is the start of the data
                                // if the "|" precedes the "##" it was put only at the end
                                firstPipe = fileText.IndexOf ("|", quotedPrintable);
                                if (fileText.Count (c => c == '|') > 3 && firstPipe > -1) pipeDelimiter = true;
                                else pipeDelimiter = false;

                                if (startPosition > -1 && endPosition == -1) // only one ## marker
                                {
                                    onlyOneMarker = true;
                                    var temp = 0;
                                    if (fromTxt != null && firstPipe > -1) temp = fileText.LastIndexOf (fromTxt, firstPipe);
                                    else
                                    {
                                        endPosition = fileText.IndexOf ("\r\n", startPosition + 2);
                                        len = endPosition - (startPosition + 2);
                                        if (len > 0) temp = fileText.IndexOf (",", startPosition, len);
                                    }
                                    if (temp > -1 && temp < startPosition) // assume ## is the end marker instead of the beginning
                                    {
                                        newFormatEndOnly = true;
                                        endPosition = startPosition;
                                        if (firstPipe > -1) startPosition = temp; // move the startPosition where the preceding callsign was found
                                        // if (fromTxt != null) startPosition = fileText.LastIndexOf (fromTxt, startPosition); // move the startPosition to the previous location of the callsign preceding the first pipe
                                        else if (fromTxt != null) startPosition = fileText.IndexOf (fromTxt, endHeader);
                                        else startPosition = endHeader;

                                    }
                                    else
                                    {
                                        newFormatStartOnly = true;
                                        endPosition = fileText.IndexOf ("\r\n", startPosition);
                                    }
                                }
                            }


                            if (newFormat && !newFormatEndOnly) // find the last ## marker in the case there are more than two
                            {
                                // newFormat = true; // new format found (two ## markers at least)
                                startPosition += 2;
                                //endPosition = fileText.IndexOf ("##", startPosition);
                                if (endPosition > -1) isValid = true;
                                while (isValid == true)
                                {
                                    int anotherEnd = fileText.IndexOf ("##", endPosition + 2);
                                    if (anotherEnd > -1) endPosition = anotherEnd;
                                    else isValid = false;
                                }
                                // if (endPosition == -1) endPosition = fileText.IndexOf ("\r\n", startPosition); // if there is only a beginning marker, set the end position to the next return
                                if (endPosition == -1) endPosition = lastBoundary; // if there is only a beginning marker, set the end position to the last boundary
                                                                                   // if (endPosition <= startPosition) startPosition = fileText.IndexOf ("|", endHeader); //
                                if (newFormatEndOnly || newFormatStartOnly)
                                {
                                    if (fileText.IndexOf ("|", endHeader) > -1) startPosition = fileText.LastIndexOf ("\r\n", endPosition) + 2; // backup to after the preceding return if there is a pipe after the header
                                    else newFormatNoPipe = true;
                                }

                            }
                            if (!onlyOneMarker && !newFormat && fileText.IndexOf ("|") > -1) newFormatPipeOnly = true; // no ## markers, but pipe delimiters are present
                            // if (!onlyOneMarker && !newFormat && (newFormatEndOnly || newFormatStartOnly) && fileText.IndexOf ("#") > -1 && fileText.IndexOf ("|") > -1) // new format almost but only one # for delimiter,
                            if (!newFormat && !newFormatEndOnly && !newFormatStartOnly && fileText.IndexOf ("#") > -1 && fileText.IndexOf ("|") > -1) // new format almost but only one # for delimiter,
                                                                                                                                                      // hoping that the combo is unique enough
                            {
                                newFormatSingleOnly = true;
                                startPosition = firstPipe; // set start to first pipe
                                startPosition = fileText.LastIndexOf ("\r\n", startPosition) + 2; // reset to after the preceding return
                                endPosition = fileText.IndexOf ("#", startPosition); // locate the # marker
                                if (endPosition < firstPipe) // the # marker is at the beginning so swap positions
                                {
                                    startPosition = endPosition + 1;
                                    endPosition = lastBoundary;
                                }

                                if (endPosition == -1) endPosition = fileText.IndexOf ("\r\n", startPosition);
                            }
                            else if (newFormat)
                            {
                                // if newFormat, skip the form extraction settings (do nothing)
                            }
                            else // this sections finds the msgField location within the checkin data
                            {
                                // skip APRS header 
                                if (APRS > -1 && !newFormat)
                                {
                                    startPosition = fileText.IndexOf ("FROM:", endHeader);
                                    if (startPosition > -1)
                                    {
                                        startPosition = fileText.IndexOf ("\r\n", startPosition);
                                        if (startPosition > -1) { startPosition += 2; }
                                        endPosition = fileText.IndexOf ("DO NOT REPLY", startPosition) - 1;
                                    }

                                }
                                // skip JS8Call header 

                                // no "##" markers but pipe delimiter is used
                                else if (pipeDelimiter)
                                {
                                    startPosition = firstPipe; // set start to first pipe
                                    startPosition = fileText.LastIndexOf ("\r\n", startPosition) + 2; // reset to after the preceding return
                                    endPosition = fileText.IndexOf ("\r\n", startPosition); // locate the end of the line
                                }

                                // adjust for ICS 213
                                else if (ics > -1 && !newFormat)
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
                                else if (winlinkCkin > -1 && !newFormat)
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
                                else if (BPQ > -1 && !newFormat)
                                {
                                    len = lastBoundary - quotedPrintable;
                                    if (len < 0) len = 0;
                                    startPosition = fileText.IndexOf ("BPQ", quotedPrintable, len);
                                    if (startPosition > -1) { startPosition += 12; }
                                    // endPosition = fileText.IndexOf ("--BOUNDARY", startPosition) - 2;
                                    endPosition = lastBoundary - 2;
                                }
                                else if (localWeather > -1 && !newFormat)
                                {
                                    startPosition = fileText.IndexOf ("NOTES:");
                                    if (startPosition > -1) { startPosition += 9; }
                                    endPosition = fileText.IndexOf ("----------", startPosition) - 1;
                                }

                                else if (severeWeather > -1 && !newFormat)
                                {
                                    startPosition = fileText.IndexOf ("COMMENTS:");
                                    if (startPosition > -1) { startPosition += 10; }
                                    endPosition = fileText.IndexOf ("----------", startPosition) - 1;
                                }

                                else if (incidentStatus > -1 && !newFormat)
                                {
                                    startPosition = fileText.IndexOf ("REPORT SUBMITTED BY:");
                                    if (startPosition > -1) { startPosition += 20; }
                                    endPosition = fileText.IndexOf ("----------", startPosition) - 1;
                                }

                                else if (damAssess > -1 && !newFormat)
                                {
                                    startPosition = fileText.IndexOf ("COMMENTS:");
                                    if (startPosition > -1) { startPosition += 21; }
                                    endPosition = fileText.IndexOf ("----------", startPosition) - 1;
                                }

                                else if (fieldSit > -1 && !newFormat)
                                {
                                    startPosition = fileText.IndexOf ("COMMENTS:");
                                    if (startPosition > -1) { startPosition += 11; }
                                    endPosition = fileText.IndexOf ("\r\n", startPosition);
                                }

                                else if (dyfi > -1 && !newFormat)
                                {
                                    startPosition = fileText.IndexOf ("COMMENTS");
                                    if (startPosition > -1) { startPosition += 11; }
                                    endPosition = fileText.IndexOf ("\r\n", startPosition) - 1;
                                }

                                else if (rriWR > -1 && !newFormat)
                                {
                                    startPosition = fileText.IndexOf ("BT\r\n");
                                    if (startPosition > -1) { startPosition += 3; }
                                    endPosition = fileText.IndexOf ("------", startPosition) - 1;
                                }

                                else if (quickM > -1 && !newFormat)
                                {
                                    startPosition = quickM;
                                    startPosition = fileText.IndexOf ("SENT ON ", startPosition);
                                    if (startPosition > -1)
                                    {
                                        startPosition = fileText.IndexOf ("\r\n", startPosition) + 2;
                                        // endPosition = fileText.IndexOf ("--BOUNDARY", startPosition) - 2;
                                        endPosition = lastBoundary;
                                    }
                                    else startPosition = 0; endPosition = 0;
                                }

                                else if (qwm > -1 && !newFormat)
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
                                else if (mi > -1 && !newFormat)
                                {
                                    startPosition = fileText.IndexOf ("ADDITIONAL INFORMATION");
                                    startPosition = fileText.IndexOf ("\r\n", startPosition);
                                    endPosition = fileText.IndexOf ("----", startPosition) - 1;
                                }
                                else if (ICS201 > -1 && !newFormat)
                                {
                                    startPosition = fileText.IndexOf ("PROTECT RESPONDERS FROM THOSE HAZARDS.");
                                    startPosition = fileText.IndexOf ("\r\n", startPosition);
                                    endPosition = fileText.IndexOf ("6. PREPARED BY:", startPosition) - 1;
                                }

                                else if (ICS202 > -1 && !newFormat)
                                {
                                    startPosition = fileText.IndexOf ("GENERAL SITUATIONAL AWARENESS");
                                    // startPosition = quotedPrintable;
                                    startPosition = fileText.IndexOf ("\r\n", startPosition) + 2;
                                    endPosition = fileText.IndexOf ("5. SAFETY PLAN", startPosition) - 1;
                                }

                                else if (ICS204 > -1 && !newFormat)
                                {
                                    startPosition = fileText.IndexOf ("7. SPECIAL INSTRUCTIONS:");
                                    // startPosition = quotedPrintable;
                                    startPosition = fileText.IndexOf ("\r\n", startPosition) + 2;
                                    endPosition = fileText.IndexOf ("8. COMMUNICATIONS", startPosition) - 1;
                                }
                                else if ((PosReport || copyPR > -1) && !newFormat)
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
                                        reminderTxt += "You appear to have copied the QTH message to " + netName + " instead of fowarding the response from Service.\r\n";

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
                                else if (radioGram > -1 && !newFormat)
                                {
                                    startPosition = fileText.IndexOf ("BT\r\n", quotedPrintable);
                                    if (startPosition != -1)
                                    {
                                        startPosition = startPosition + 4;
                                        // startPosition = fileText.LastIndexOf ("\r\n", startPosition) + 2;
                                        // endPosition = fileText.IndexOf ("\r\nAR ", startPosition);
                                        // endPosition = fileText.LastIndexOf ("/");
                                        // endPosition = fileText.IndexOf ("\r\n", endPosition);
                                        endPosition = fileText.IndexOf ("BT\r\n", startPosition);
                                        radioGramCt++;
                                    }
                                    else
                                    {
                                        Console.WriteLine ("No valid delimiter in Radiogram message! " + messageID);
                                    }
                                }
                                else
                                {
                                    // end of the header information as the start of the msg field
                                    if (forwarded < 0)
                                    {
                                        if (!newFormat && !newFormatStartOnly)
                                        {
                                            if (quotedPrintable > -1) startPosition = quotedPrintable;
                                            else startPosition = endHeader;
                                        }
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
                                    //endPosition = fileText.IndexOf ("--BOUNDARY", startPosition) - 1;
                                    if (!newFormat && !newFormatStartOnly) endPosition = lastBoundary;
                                }
                            }

                            if (newFormatEndOnly || newFormatStartOnly) reminderTxt += "\r\nYou are encouraged to use the new format with '##' at both the beginning and end of your checkin data";
                            else if (!newFormat) reminderTxt += "\r\nYou are encouraged to use the new format with '##' at the beginning and end '##' of your checkin data with '|' delimiters!";
                            else if (newFormatPipeOnly) reminderTxt += "\r\nYou are encouraged to use the new format with '##' at the beginning and end '##' of your checkin data";
                            else if (newFormatSingleOnly) reminderTxt += "\r\nIt appears that you tried to use the new format, but with a single '#'. You are encouraged to use the new format with '##' at both the beginning and end of your checkin data";
                            if (newFormatNoPipe) reminderTxt += "\r\nYou are encouraged to use the '|' as the delimiter for your checkin data";
                            if (newFormat && msgField.IndexOf (",") > -1) reminderTxt += "Using both ',' & '|' as delimiters in your checkin data doesn't work. If you intended to use the comma as part of the data, that does work.";

                            if (startPosition == -1)
                                if (quotedPrintable > -1) startPosition = quotedPrintable;
                                else startPosition = endHeader;
                            if (endPosition <= startPosition) endPosition = lastBoundary;

                            string originalMsgField = fileText.Substring (startPosition, endPosition - startPosition);

                            msgField = getMsgField (startPosition, endPosition, messageID, fileText, msgField);

                            // string checkinFrom = checkIn;
                            if (msgField.IndexOf ("WINLINK") > -1) winlinkCt++;
                            if (msgField.IndexOf ("PAT") > -1) patCt++;
                            if (msgField.IndexOf ("WOAD") > -1) woadCt++;
                            if (msgField.IndexOf ("AIRMAIL") > -1) airmailCt++;
                            if (msgField.IndexOf ("RADIOMAIL") > -1 || msgField.IndexOf ("RADIO MAIL") > -1) radioMailCt++;
                            // 20250113 if (msgField.IndexOf ( netName + " Ask Template Exercise") > -1) exerciseCompleteCt++;
                            // 20250127 if (ICS201Ct >0) exerciseCompleteCt++;
                            // if (radioGram > 0) exerciseCompleteCt++; // 20250210 exercise
                            // if (ICS202 > -1) exerciseCompleteCt++; // 20250217 for exercise
                            // if (w3w > -1) exerciseCompleteCt++; // 20250303 for W3W exercise
                            // if (ICS203 > -1) exerciseCompleteCt++; // 202500317 exercise
                            // if (ICS204 > -1) exerciseCompleteCt++; // 20250421 exercise

                            if (radioGram > 0) msgField = msgField.Replace ("\r\n", " "); // Radiogram chops the message into 40 byte strings, so put it back together
                            checkinItems = null; // empty the array
                            len = msgField.Length;
                            if (len > 0)
                            {
                                checkinItems = getCheckinData (len, msgField, checkinItems, newFormat);
                                if (checkinItems != null && checkinItems.Length > 0)
                                {
                                    checkinItems = checkinItems
                                        .Select (item => item!.Replace (",", ""))  // ! tells compiler it's safe
                                        .ToArray ();
                                }

                                if (checkinItems != null && checkinItems.Length > 0 && checkinItems [0] != null)
                                {
                                    var item0 = checkinItems [0];
                                    if (item0 != null) checkIn = item0.Trim ().Trim (',').Replace ("<", "").Replace (">", "");
                                }
                                else
                                {
                                    checkIn = null; // Explicitly set to null
                                    string safeMessageID = messageID ?? "";
                                    Console.WriteLine ("1406 Invalid checkin data in messageID: " + safeMessageID);
                                }
                                // checkIn = checkIn?.Trim() ?? "";
                            }
                            else
                            {
                                checkIn = null; // Explicitly set to null
                                string safeMessageID = messageID ?? "";
                                Console.WriteLine ("Message Field is empty in: " + safeMessageID);
                            }
                            // check for suffix on the callsign and remove it
                            int suffixPos = checkIn?.IndexOf ("/") ?? -1;
                            if (suffixPos > -1)
                            {
                                checkIn = checkIn? [0..suffixPos] ?? ""; // Range operator (C# 8+) - null-safe
                            }
                            // now check to see if it is a perfect message and deduct points if not
                            // checkin call sign

                            // look for a callsign typo in the checkin msg
                            // do not flag checkins with an appended "/x" as a typo, but make sure it is removed to not break Winlink
                            string? tmpCheckIn = isValidCallsign (checkIn);
                            if (tmpCheckIn != "") checkIn = tmpCheckIn;

                            if (checkIn != fromTxt && xSource != "SMTP" && checkIn != "W5SJT") // W5SJT uses a personal account to login for Tom Green County Emergency Management
                            {
                                brokenCheckin = true;
                                if (ICS202 > -1) // check for the data string in the wrong field
                                {
                                    startPosition = fileText.IndexOf ("OPERATIONAL PERIOD COMMAND EMPHASIS:");
                                    // startPosition = quotedPrintable;
                                    startPosition = fileText.IndexOf ("\r\n", startPosition) + 2;
                                    endPosition = fileText.IndexOf ("GENERAL SITUATIONAL AWARENESS", startPosition) - 1;
                                    if (messageID != null) msgField = getMsgField (startPosition, endPosition, messageID, fileText, msgField);
                                    len = msgField.Length;
                                    checkinItems = new string [10];
                                    checkinItems = getCheckinData (len, msgField, checkinItems, newFormat);
                                    if (checkinItems != null && checkinItems.Length > 0)
                                    {
                                        checkinItems = checkinItems
                                            .Select (item => item!.Replace (",", ""))  // ! tells compiler it's safe
                                            .ToArray ();
                                    }
                                    if (checkinItems != null && checkinItems.Length > 0 && checkinItems [0] != null)
                                    {
                                        var item0 = checkinItems [0];
                                        if (item0 != null) checkIn = item0.Trim ().Trim (',').Replace ("<", "").Replace (">", "");
                                    }
                                    else
                                    {
                                        checkIn = null; // Explicitly set to null
                                        string safeMessageID = messageID ?? "";
                                        Console.WriteLine ("1319 Invalid checkin data in messageID: " + safeMessageID);
                                    }

                                    if (checkIn == fromTxt) brokenCheckin = false;
                                }
                                // assume the from text is correct?
                                //if ((checkIn == "" || checkIn != fromTxt) && fromTxt != "") checkIn = fromTxt;
                                // assume the xSource is correct?
                                // if (checkIn == "" && xSource != "") checkIn = xSource;
                                if (checkIn != null) endPosition = checkIn.IndexOf ("/");
                                if (checkIn != null && endPosition > -1) checkIn = checkIn.Substring (0, endPosition);
                                if (brokenCheckin)
                                {
                                    if (tempFromTxt == "")
                                    {
                                        if (fromTxt != null) callSignTypo = fromTxt;
                                        Console.WriteLine ("1456 fromTxt is null or invalid in :" + messageID);
                                    }
                                    else if (tempCheckIn == "")
                                    {
                                        if (checkIn != null) callSignTypo = checkIn;
                                        Console.WriteLine ("1479 checkIn: " + checkIn + " from msgField " + msgField + " is null, invalid, or does not match the From data: " + fromTxt + " in :" + messageID);
                                        if (checkIn == null) checkIn = fromTxt;
                                    }
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
                                    pointsOff = "\tminus 1 for invalid or missing callsign as the first field - " + checkIn + "\r\n";
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
                            }

                            // debug Console.Write("Start at:"+startPosition+": and end at:"+endPosition+"\nCallsign found: "+checkIn);
                            // eliminate duplicates from the map file                          
                            if (!brokenCheckin)
                            {
                                Console.Write ("Callsign \"" + checkIn + "\" not found in: " + messageID + "\r\n");
                            }
                            else
                            {
                                // continue checking for perfect message and point deductions
                                // int checkInItemsCt = checkinItems.Length;
                                // int i = 0;
                                if (newFormat) newFormatCt++;
                                checkinCountry = "";
                                checkinCountryLong = "";
                                longCountry = false;

                                len = 0;
                                if (checkinItems != null) len = checkinItems.Length;
                                if (len > 0)
                                {
                                    if (len < 8)
                                    {
                                        score = score - (8 - len);
                                        // pointsOff += "\tminus " + (8 - len) + " point(s), for missing delimiter(s)/fields - see examples below.\r\n"; 
                                        pointsOff += "\tminus " + (8 - len) + " point(s), for missing delimiter(s)/fields - see examples below.";
                                        if (APRS > -1) pointsOff += " - maybe because you checked in via APRS.";
                                        pointsOff += "\r\n";
                                        if ((msgField.IndexOf ("|") > -1) && (msgField.IndexOf (",") > -1)) pointsOff += "\tYou may have mixed the '|' and ',' delimiters in the check in data.";
                                        isPerfect = false;
                                    }

                                    if (checkinItems != null && len > 2)
                                    {
                                        // array is zero based
                                        string? item1 = checkinItems [1]; // Store the element in a local variable
                                        if (item1 == null)
                                        {
                                            checkinName = "";
                                        }
                                        else
                                        {
                                            checkinName = isValidName (item1).Trim ().Trim (','); // Use item1 instead
                                        }

                                        if (checkinName == "")
                                        {
                                            isPerfect = false;
                                            score--;
                                            if (item1 == null)
                                            {
                                                reminderTxt2 = "\tminus 1 point, missing or invalid name in field 2 - (null) \r\n";
                                            }
                                            else
                                            {
                                                reminderTxt2 = "\tminus 1 point, missing or invalid name in field 2 - " + item1 + " \r\n";
                                            }
                                        }
                                    }
                                    else
                                    {
                                        // Handle the case where checkinItems is null or len <= 2
                                        checkinName = "";
                                        isPerfect = false;
                                        score--;
                                        reminderTxt2 = "\tminus 1 point, missing or invalid name in field 2 - (checkinItems is null or insufficient length) \r\n";
                                    }

                                    if (checkinItems != null && len >= 6)
                                    {
                                        string? item5 = checkinItems [5]; // Store the element in a local variable
                                        string trimmedItem5 = item5 != null ? item5.Trim ().Trim (',') : "";

                                        (checkinCountry, found) = isValidField (trimmedItem5, countries, found);
                                        if (checkinCountry != "" && found > -1)
                                        {
                                            if (checkinCountry.Length > 3)
                                            {
                                                longCountry = true;
                                                checkinCountryLong = checkinCountry; // found the long country name
                                                endPosition = countries.IndexOf (",", found + 1) + 1;
                                                checkinCountry = countries.Substring (endPosition - 4, 3);
                                            }
                                            else // found the country abbreviation
                                            {
                                                startPosition = countries.LastIndexOf (",", found - 1) + 1;
                                                endPosition = countries.IndexOf (":", startPosition);
                                                int countryLength = endPosition - startPosition;
                                                if (countryLength > 0) checkinCountryLong = countries.Substring (startPosition, countryLength);
                                            }

                                        }

                                        if (checkinCountry == "")
                                        {
                                            isPerfect = false;
                                            score--;
                                            pointsOff += "\tminus 1 point, missing or invalid country in field 6 (3 letter abbreviation?) - " + (item5 ?? "(null)") + ", try USA, PHL, DEU, COL, VEN, CAN, AUS, AUT, TTO, NZL, BEL, NOR, ROU, SRB, LCA, etc";
                                        }
                                        if (longCountry)
                                        {
                                            isPerfect = false;
                                            score--;
                                            pointsOff += "\tminus 1 point, full country name (" + checkinCountryLong + ") was used instead of the 3 letter ISO abbreviation (" + checkinCountry + ").";
                                        }
                                    }

                                    if (checkinItems != null && len >= 5)
                                    {
                                        // Handle checkinItems[4] for state
                                        string? item4 = checkinItems [4];
                                        checkinState = item4 != null ? item4.Replace (".", "").Trim ().Trim (',') : "";
                                        int scoreState = 0;
                                        string tempStr = "";
                                        string tempStr2 = "";
                                        string states = "";
                                        reminderTxt2 = "";
                                        latitude = 0;
                                        longitude = 0;

                                        if (checkinCountry != null) // Add null check for checkinCountry
                                        {
                                            switch (checkinCountry) // find valid state
                                            {
                                                case "AUT":  // Austria AUT
                                                    states = ",B,K,N,S,ST,T,O,W,V,";
                                                    found = 0;
                                                    (checkinState, found) = isValidField (checkinState, states, found);
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
                                                    states = ",NL,PE,NS,NB,QC,ON,MB,SK,AB,BC,YT,NT,NU,";
                                                    (checkinState, found) = isValidField (checkinState, states, found);
                                                    if (checkinState == "")
                                                    {
                                                        isPerfect = false;
                                                        tempStr += "missing or invalid CAN province abbreviation ";
                                                        scoreState++;
                                                    }
                                                    break;
                                                case "DEU": // Deutschland - Germany DEU
                                                    states = ",BW,BY,BE,BB,HB,HH,HE,MV,NI,NW,RP,SL,SN,ST,SH,TH,";
                                                    if (checkinState == "")
                                                    {
                                                        isPerfect = false;
                                                        tempStr += "fehlendes oder ungültiges DEU-Landeskürzel ";
                                                        scoreState++;
                                                    }
                                                    break;
                                                case "GBR":
                                                case "UK": // United Kingdom UK Great Britain GBR
                                                    break;
                                                case "NZL": // New Zealand NZL
                                                    states = ",AUK,BOP,CAN,GIS,WGN,HKB,MWT,MWT,MBH,NSN,NTL,OTA,STL,TKI,TKI,TAS,HKB,WGN,WTC,STL,GIS,NTL,TAS,BOP,AUK,WKO,WKO,CAN,WTC,NSN,OTA,";
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
                                                    states = ",ABR,AGN,AGS,AKL,ALB,ANT,APA,AUR,BAN,BAS,BEN,BIL,BOH,BTG,BTN,BUK,BUL,CAG,CAM,CAN,CAP,CAS,CAT,CAV,CEB,COM,DAO,DAS,DAV,DIN,DVO,EAS,GUI,IFU,ILI,ILN,ILS,ISA,KAL,LAG,LAN,LAS,LEY,LUN,MAD,MAS,MDC,MDR,MGN,MGS,MOU,MSC,MSR,NCO,NCR,NEC,NER,NSA,NUE,NUV,PAM,PAN,PLW,QUE,QUI,RIZ,ROM,SAR,SCO,SIG,SLE,SLU,SOR,SUK,SUN,SUR,TAR,TAW,WSA,ZAN,ZAS,ZMB,ZSI,";
                                                    if (checkinState == "")
                                                    {
                                                        isPerfect = false;
                                                        tempStr += "missing or invalid PHL region abbreviation ";
                                                        scoreState++;
                                                    }
                                                    break;
                                                case "ROU": // Romania ROU 
                                                    break;
                                                case "SRB": // Serbia SRB
                                                    break;
                                                case "LCA": // St. Lucia LCA
                                                    break;
                                                case "TTO": // Trinidad & Tobago TTO
                                                    break;
                                                case "USA": // United States of America USA
                                                    states = ",AK,AL,AR,AS,AZ,CA,CO,CT,DC,DE,FL,GA,GU,HI,IA,ID,IL,IN,KS,KY,LA,MA,MD,ME,MI,MN,MO,MP,MS,MT,NC,ND,NE,NH,NJ,NM,NV,NY,OH,OK,OR,PA,PR,RI,SC,SD,TN,TX,UM,UT,VA,VI,VT,WA,WI,WV,WY,";
                                                    (checkinState, found) = isValidField (checkinState, states, found);
                                                    if (checkinState == "")
                                                    {
                                                        isPerfect = false;
                                                        tempStr += "missing or invalid USA state 2 letter abbreviation ";
                                                        if (item4 == "PUERTO RICO") tempStr2 += ", try \"PR\"";
                                                        scoreState++;
                                                    }
                                                    break;
                                                case "VEN": // Venezuela VEN
                                                    states = ",DC,A,B,C,D,E,F,G,H,I,J,K,L,M,N,O,P,R,S,T,U,V,W,X,Y,Z,";
                                                    (checkinState, found) = isValidField (checkinState, states, found);
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
                                        }

                                        if (reminderTxt2 != "" || checkinCountry == "")
                                            pointsOff += reminderTxt2 + "\r\n";
                                        reminderTxt2 = "";

                                        if (scoreState > 0)
                                        {
                                            pointsOff += "\tminus 1 point, " + tempStr + "in field 5 -  " + (item4 ?? "(null)") + tempStr2 + "\r\n";
                                            score--;
                                        }

                                        if (len > 4 && checkinCountry == "USA") // check only for USA
                                        {
                                            string? item3 = checkinItems [3];
                                            checkinCounty = item3 != null ? isValidName (item3.Replace (" COUNTY", "").Replace ("CO", "").Trim ().Trim ('.').Trim (',')) : "";
                                            if (checkinCounty == "")
                                            {
                                                isPerfect = false;
                                                pointsOff += "\tmissing or invalid county in field 4, use NA or NONE if you don't have one - " + (item3 ?? "(null)") + "\r\n";
                                            }
                                        }

                                        if (len > 3)
                                        {
                                            string? item2 = checkinItems [2];
                                            checkinCity = item2 != null ? isValidName (item2.Trim ().Trim (',')) : "";
                                            if (checkinCity == "")
                                            {
                                                isPerfect = false;
                                                score--;
                                                pointsOff += "\tminus 1 point, missing or invalid city in field 3 - use NA or NONE of you don't have one. " + (item2 ?? "(null)") + "\r\n";
                                            }
                                        }
                                    }

                                    if (checkinItems != null && len >= 7)
                                    {
                                        string? item6 = checkinItems [6];
                                        bandStr = item6 != null ? item6.Trim ().Trim (',') : "";
                                        var tmpBandStr = bandStr; // Store the original bandStr for later use
                                        bandStr = checkBand (bandStr) ?? ""; // Add null check for checkBand result
                                        if (bandStr == "" && tmpBandStr.IndexOf ("VHF") > -1) bandStr = "VHF";
                                        if (bandStr == "")
                                        {
                                            isPerfect = false;
                                            score--;
                                            pointsOff += "\tminus 1 point, missing or invalid band in field 7 - " + (item6 ?? "(null)") + ", try something like TELNET, 2M, 70CM, 20M, 40M, VHF, UHF, HF, SHF, etc.\r\n";
                                            if (msgField != null && msgField.IndexOf ("AREDN") > -1)
                                            {
                                                pointsOff += "\tAREDN is a project, not a valid band. Try \"5CM, 9CM, 13CM, 33CM, or SHF.\"\r\n";
                                            }
                                        }
                                        else
                                        {
                                            checkinItems [6] = bandStr;
                                        }
                                    }

                                    if (checkinItems != null && len >= 8)
                                    {
                                        modeTypo = string.Empty;
                                        string? item7 = checkinItems [7];
                                        modeStr = item7 != null ? item7.Trim ().Trim (',') : "";
                                        if (modeStr != null && modeStr.Contains ("PACKET")) modeStr = "PACKET";
                                        if (modeStr != null && bandStr != null) // Explicit check for both arguments
                                        {
                                            (modeStr, modeTypo, var empty) = checkMode (modeStr ?? string.Empty, bandStr ?? string.Empty, modeTypo ?? string.Empty);
                                            if (modeTypo != string.Empty)
                                            {
                                                reminderTxt += modeTypo;
                                                // checkinItems [7] = modeStr; keep the original for error reporting.
                                            }
                                        }
                                        else
                                        {
                                            modeStr = ""; // Fallback if either modeStr or bandStr is null
                                        }
                                        string tempStr = "";
                                        if (modeStr == "")
                                        {
                                            isPerfect = false;
                                            score--;
                                            if (bandStr != null && bandStr == "TELNET") tempStr = ", try SMTP";
                                            pointsOff += "\tminus 1 point, missing or invalid mode in field 8 - " + (item7 ?? "(null)") + tempStr + ", try something like PACKET, VARA FM, VARA HF, ARDOP, MESH, APRS, JS8CALL, PACTOR, etc)\r\n";
                                            if (msgField != null && msgField.IndexOf ("AREDN") > -1)
                                            {
                                                pointsOff += "\tAREDN is a project, not a valid mode. Try \"MESH\"\r\n";
                                            }
                                        }
                                    }
                                }
                                // check to see if this is a duplicate checkin
                                if (checkIn != null) startPosition = testString.IndexOf (checkIn);
                                if (startPosition >= 0)
                                {
                                    if (dupCt == 0) { duplicates.Append ("Duplicates: \r\n\t"); }
                                    //debug Console.Write("netName "+checkIn+" is a duplicate, skipping. It is "+dupCt+" of "+msgTotal+" total messages.\n");
                                    duplicates.Append (checkIn + ", ");
                                    dupeFlag = 1;
                                    dupCt++;
                                }
                                ct++;


                                testString = testString + checkIn + " | ";
                                // the spreadsheet chokes if the string ends with "|" so
                                // don't let that happen by writing the first one without a delimiter
                                // prepending the delimiter to the rest.
                                if (ct == 1)
                                {
                                    netCheckinString.Append (checkIn);
                                    if (noSummary == -1) netAckString2.Append (checkIn);
                                }
                                else if (ct > 1 && dupeFlag == 0)
                                {
                                    netCheckinString.Append ("|" + checkIn);
                                    if (noSummary == -1) netAckString2.Append (";" + checkIn);
                                }

                                var msgFieldStart = (msgField != null) ? msgField.IndexOf ("\r\n") : -1;
                                string notFirstLine = "";
                                if (startPR > -1 || copyPR > -1)
                                {
                                    if (msgField != null && addonString != null)
                                    {
                                        addonString.Append (checkIn + ":\t" + msgField.Replace ("\n", ", ").Replace ("\r", "").Replace ("|", "\t") + "\r\n");
                                    }
                                }
                                else
                                {
                                    if (msgFieldStart > -1 && msgField != null) // Explicitly check msgField again
                                    {
                                        len = msgField.Length - msgFieldStart;
                                        if (len > 0)
                                        {
                                            notFirstLine = msgField.Substring (msgFieldStart, len);
                                            notFirstLine = notFirstLine.Replace ("\n", ", ")
                                                .Replace ("\r", "")
                                                .Trim ();
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
                                                .Trim (',');
                                            if (notFirstLine.Length > 0 && addonString != null)
                                            {
                                                addonString.Append (checkIn + ":\t" + notFirstLine + "\r\n");
                                            }
                                        }
                                    }
                                }

                                // Extract latitude and longitude
                                // Winlink Checkin has its own tags so check them first

                                if (winlinkCkin > 0)
                                {
                                    // get latitude from the winlink checkin
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
                                    // get longitude from the winlink checkin
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
                                    // get maidenhead grid from the winlink checkin
                                    startPosition = fileText.IndexOf ("GRID SQUARE:");
                                    if (startPosition > -1 && fileText != null)
                                    {
                                        startPosition += 12;
                                        endPosition = fileText.IndexOf ("\r\n", startPosition);
                                        if (endPosition > startPosition) // Ensure valid substring range
                                        {
                                            len = endPosition - startPosition;
                                            if (len > 0)
                                            {
                                                maidenheadGrid = fileText.Substring (startPosition, len).Trim ();
                                            }
                                            else
                                            {
                                                maidenheadGrid = !string.IsNullOrEmpty (msgField) ? ExtractMaidenheadGrid (msgField) ?? string.Empty : string.Empty;
                                            }
                                        }
                                        else
                                        {
                                            maidenheadGrid = !string.IsNullOrEmpty (msgField) ? ExtractMaidenheadGrid (msgField) ?? string.Empty : string.Empty;
                                        }
                                    }
                                    else
                                    {
                                        maidenheadGrid = !string.IsNullOrEmpty (msgField) ? ExtractMaidenheadGrid (msgField) ?? string.Empty : string.Empty;
                                    }
                                    //if (startPosition > -1)
                                    //{
                                    //    startPosition += 12;
                                    //    endPosition = fileText.IndexOf ("\r\n", startPosition);
                                    //    len = endPosition - startPosition;
                                    //    if (len > 0)
                                    //    {
                                    //        maidenheadGrid = fileText.Substring (startPosition, len);
                                    //        maidenheadGrid.Trim ();
                                    //    } else maidenheadGrid = ExtractMaidenheadGrid (msgField);
                                    //}
                                }

                                if (latitudeStr == "" || longitudeStr == "")
                                {
                                    //skip past the messageID because sometimes the regex for coordinates matches it
                                    // and stop before reading any binary attachments

                                    // startPosition = quotedPrintable;
                                    //startPosition = fileText.IndexOf ("MESSAGE-ID:");
                                    startPosition = fileText != null ? fileText.IndexOf ("MESSAGE-ID:") : -1;
                                    // startPosition = fileText.IndexOf ("\r\n", startPosition);
                                    startPosition = fileText != null ? fileText.IndexOf ("\r\n", startPosition) : -1;
                                    if (startPosition > -1) { startPosition += 2; }
                                    // need an end position because some messages have a binary attachment that gives a false match
                                    // endPosition = fileText.IndexOf ("PRINTABLE", startPosition);
                                    endPosition = fileText != null ? fileText.IndexOf ("PRINTABLE", startPosition) : -1;
                                    // endPosition = fileText.IndexOf ("--BOUNDARY", endPosition);
                                    endPosition = lastBoundary;
                                    len = endPosition - startPosition;
                                    if (len > 0)
                                    {
                                        if (ExtractCoordinates (fileText.Substring (startPosition, len), out latitude, out longitude))
                                        {
                                            // Console.WriteLine(messageID+" latitude: "+latitude+" longitude: "+longitude);                                
                                            // maidenheadGrid = ExtractMaidenheadGrid (fileText.Substring (startPosition, len));
                                            maidenheadGrid = ExtractMaidenheadGrid (msgField);
                                            if (maidenheadGrid == "invalid") Console.WriteLine (messageID);
                                        }
                                        else
                                        {
                                            // no valid GPS coordinates found, look for a maidenhead grid
                                            // linda in AK uses something funky to checkin and never puts things in the correct format.
                                            // now fixed approximately at 3229
                                            // This should at least keep her from showing up in the Atlantic Ocean
                                            //if (fromTxt == "AD4BL")
                                            //{
                                            //    fileText = fileText.Insert (endPosition - 1, ", BP64JU\r\n").Replace ("  ", "|");
                                            //    endPosition = fileText.IndexOf ("--BOUNDARY", endPosition);
                                            //    len = endPosition - startPosition;
                                            //    isPerfect = false;
                                            //}
                                            //// maidenheadGrid = ExtractMaidenheadGrid (fileText.Substring (startPosition, len));
                                            maidenheadGrid = ExtractMaidenheadGrid (msgField);
                                            if (maidenheadGrid == "invalid") { Console.WriteLine (messageID); maidenheadGrid = ""; }
                                            if (!string.IsNullOrEmpty (maidenheadGrid))
                                            {
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
                                                score -= 2;
                                                msgField = msgField + ",No Location Data Found in message";
                                                if (APRS > -1) msgField = msgField + " - perhaps because you checked in via APRS. ";
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

                                if (msgField != null)
                                {
                                    msgField = msgField.Replace ("\r\n", ",");
                                    msgFieldNumbered = fillFieldNum (msgField);
                                    if (csvString != null)
                                    {
                                        string safeCheckIn = checkIn ?? "";
                                        string safeMessageID = messageID ?? "";
                                        string safeLocType = locType ?? "";
                                        csvString.Append (safeCheckIn + ":" + safeMessageID + "," + latitude + "," + longitude + "," + safeLocType + "," + msgField + "\r\n");
                                    }
                                }
                                else
                                {
                                    msgFieldNumbered = ""; // Fallback if msgField is null
                                }

                                if (bandStr != null)
                                {
                                    bandStr = bandStr
                                        .ToUpper ()
                                        .Replace ("5.8GHZ", "5CM")
                                        .Replace (".", "")
                                        .Replace (" ", "")
                                        .Replace (" METERS", "M")
                                        .Replace (" MTERS", "M")
                                        .Replace (" METER", "M")
                                        .Replace ("METERS", "M")
                                        .Replace ("METER", "M")
                                        .Replace (")", "")
                                        .Replace ("SMTP", "TELNET")
                                        .Replace ("N/A", "TELNET")
                                        .Replace ("NA", "TELNET")
                                        .Replace ("5GHZ", "5CM")
                                        .Replace ("73M", "80M")
                                        .Replace ("75M", "80M")
                                        .Replace (" M", "M")
                                        .Trim ()
                                        .Replace ("HFAMATEUR", "HF");
                                }
                                else
                                {
                                    bandStr = ""; // Fallback if bandStr is null
                                }

                                if (bandStr != null && bandStr.IndexOf ("PACKET") > -1)
                                {
                                    modeStr = "PACKET";
                                    // if the band is declared to be packet, check to see if there is any indication of the band elsewhere in the message
                                    if (msgField != null)
                                    {
                                        if (msgField.IndexOf ("2M") > -1) { bandStr = "2M"; }
                                        if (msgField.IndexOf ("70CM") > -1) { bandStr = "70CM"; }
                                        if (msgField.IndexOf ("VHF") > -1) { bandStr = "VHF"; }
                                        if (msgField.IndexOf ("UHF") > -1) { bandStr = "UHF"; }
                                    }
                                }
                                if (bandStr != null && bandStr.IndexOf ("2M") > -1)
                                {
                                    bandStr = "2M";
                                }
                                if (msgField != null && msgField.IndexOf ("VARAFM") > -1)
                                {
                                    modeStr = "VARA FM";
                                    if (msgField != null) // Already checked, but included for clarity
                                    {
                                        if (msgField.IndexOf ("2M") > -1) { bandStr = "2M"; }
                                        if (msgField.IndexOf ("70CM") > -1) { bandStr = "70CM"; }
                                        if (msgField.IndexOf ("VHF") > -1) { bandStr = "VHF"; }
                                        if (msgField.IndexOf ("UHF") > -1) { bandStr = "UHF"; }
                                    }
                                }
                                if (!string.IsNullOrEmpty (msgField) && msgField.IndexOf ("VARAHF") > -1)
                                {
                                    if (bandStr == "") bandStr = "HF";
                                    modeStr = "VARA HF";
                                }

                                if (!string.IsNullOrEmpty (bandStr))
                                {
                                    bandStr = checkBand (bandStr) ?? "";
                                }
                                else
                                {
                                    bandStr = ""; // Fallback if bandStr is null
                                }

                                if (bandStr == "")
                                {
                                    // if both the band and the mode have invalid data, try scraping through the fileText
                                    if (fileText != null)
                                    {
                                        if (fileText.IndexOf ("3CM") > -1) { bandStr = "3CM"; }
                                        if (fileText.IndexOf ("5CM") > -1) { bandStr = "5CM"; }
                                        if (fileText.IndexOf ("13CM") > -1) { bandStr = "13CM"; }
                                        if (fileText.IndexOf ("23CM") > -1) { bandStr = "23CM"; }
                                        if (fileText.IndexOf ("33CM") > -1) { bandStr = "33CM"; }
                                        if (fileText.IndexOf ("70CM") > -1) { bandStr = "70CM"; }
                                        if (fileText.IndexOf ("1.25M") > -1) { bandStr = "1.25M"; }
                                        if (fileText.IndexOf ("2M") > -1) { bandStr = "2M"; }
                                        if (fileText.IndexOf ("10M") > -1) { bandStr = "10M"; }
                                        if (fileText.IndexOf ("12M") > -1) { bandStr = "12M"; }
                                        if (fileText.IndexOf ("15M") > -1) { bandStr = "15M"; }
                                        if (fileText.IndexOf ("17M") > -1) { bandStr = "17M"; }
                                        if (fileText.IndexOf ("20M") > -1) { bandStr = "20M"; }
                                        if (fileText.IndexOf ("30M") > -1) { bandStr = "30M"; }
                                        if (fileText.IndexOf ("40M") > -1) { bandStr = "40M"; }
                                        if (fileText.IndexOf ("60M") > -1) { bandStr = "60M"; }
                                        if (fileText.IndexOf ("6M") > -1) { bandStr = "6M"; }
                                        if (fileText.IndexOf ("80M") > -1) { bandStr = "80M"; }
                                        if (fileText.IndexOf ("HF") > -1) { bandStr = "HF"; }
                                        if (fileText.IndexOf ("VHF") > -1) { bandStr = "VHF"; }
                                        if (fileText.IndexOf ("UHF") > -1) { bandStr = "UHF"; }
                                        if (fileText.IndexOf ("SHF") > -1) { bandStr = "SHF"; }
                                        if (fileText.IndexOf ("TELNET") > -1) { bandStr = "TELNET"; }
                                    }

                                    if (bandStr == "")
                                    {
                                        // badBandString is now static, so the null check is technically unnecessary since it's initialized
                                        // But we'll keep it for safety in case the initialization changes
                                        if (badBandString != null)
                                        {
                                            string safeMessageID = messageID ?? "";
                                            string safeCheckIn = checkIn ?? "";
                                            string safeMsgFieldNumbered = msgFieldNumbered ?? "";
                                            badBandString.Append ("\tBad band: " + safeMessageID + " - " + safeCheckIn + ": _" + bandStr + "_  |  " + safeMsgFieldNumbered + "\r\n");
                                            badBandCt++;
                                        }
                                    }
                                }
                                else
                                {
                                    bandCt++;
                                }

                                if (!string.IsNullOrEmpty (modeStr)) modeStr = modeStr
                                    .ToUpper ()
                                    .Trim ()
                                    .Replace ("WINLINK", "")
                                    .Replace ("AREDN", "MESH")
                                    //.Replace ("AX.25", "PACKET")
                                    .Replace ("WINLINK", "")
                                    .Replace ("(", "")
                                    .Replace (".", "")
                                    .Replace ("TELNET", "SMTP")
                                    // .Replace ("SMPT", "SMTP")
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
                                if (modeStr != null)
                                {
                                    if (modeStr.IndexOf ("MESH") > -1) { modeStr = "MESH"; }
                                    if (modeStr.IndexOf ("VARA") > -1 && bandStr == "HF") { modeStr = "VARA HF"; }
                                    if (modeStr.IndexOf ("VARA") > -1 && (bandStr == "VHF" || bandStr == "UHF" || bandStr == "SHF" || bandStr == "2M" || bandStr == "70CM" || bandStr == "1.25M" || bandStr == "33CM" || bandStr == "23CM" || bandStr == "13CM" || bandStr == "5CM" || bandStr == "3CM")) { modeStr = "VARA FM"; }
                                }
                                if (bandStr == "TELNET") { modeStr = "SMTP"; }

                                if (modeStr != "")
                                {
                                    if (bandStr == "")
                                    {
                                        if (modeStr == "VARA HF") { bandStr = "HF"; }
                                        // if (modeStr == "VARA FM") { bandStr = "VHF"; }
                                    }

                                }
                                // modeStr = checkMode (modeStr, bandStr); 20250722
                                //if (modeStr != null && bandStr != null)
                                //{
                                //    (modeStr, modeTypo, bandStr) = checkMode (modeStr ?? string.Empty, bandStr ?? string.Empty, modeTypo ?? string.Empty);// Ensure checkMode doesn't return null
                                //    if (modeTypo != null) reminderTxt += modeTypo;
                                //}
                                //else
                                //{
                                //    modeStr = ""; // Fallback if either is null
                                //}
                                if (modeStr == "SMTP") { bandStr = "TELNET"; }
                                if (modeStr == "MESH") { meshCt++; }
                                if (modeStr == "")
                                {
                                    if (fileText != null)
                                    {
                                        if (fileText.IndexOf ("VARA FM") > -1) { modeStr = "VARA FM"; }
                                        if (fileText.IndexOf ("VARA HF") > -1) { modeStr = "VARA HF"; }
                                        if (fileText.IndexOf ("PACTOR") > -1) { modeStr = "PACTOR"; }
                                        if (fileText.IndexOf ("TELNET") > -1) { modeStr = "SMTP"; bandStr = "TELNET"; }
                                        if (fileText.IndexOf ("SMTP") > -1) { modeStr = "SMTP"; bandStr = "TELNET"; }
                                        if (fileText.IndexOf ("ARDOP") > -1) { modeStr = "ARDOP"; }
                                        if (fileText.IndexOf ("PACKET") > -1) { modeStr = "PACKET"; }
                                        if (fileText.IndexOf ("ROBUST PACKET") > -1) { modeStr = "ROBUST PACKET"; } // Fixed syntax error here too
                                        if (fileText.IndexOf ("VARA") > -1 && bandStr != null && (bandStr == "VHF" || bandStr == "UHF" || bandStr == "SHF" || bandStr == "2M" || bandStr == "70CM" || bandStr == "1.25M" || bandStr == "33CM" || bandStr == "23CM" || bandStr == "13CM" || bandStr == "5CM" || bandStr == "3CM")) { modeStr = "VARA FM"; }
                                        if (fileText != null && fileText.IndexOf ("VARA") > -1 && bandStr != null && bandStr == "HF") { modeStr = "VARA HF"; } // Fixed line
                                    }

                                    if (modeStr == "")
                                    {
                                        msgFieldNumbered = msgField ?? "";
                                        msgFieldNumbered = fillFieldNum (msgFieldNumbered) ?? "";
                                        if (badBandString != null)
                                        {
                                            string safeMessageID = messageID ?? "";
                                            string safeCheckIn = checkIn ?? "";
                                            string safeMsgFieldNumbered = msgFieldNumbered ?? "";
                                            badBandString.Append ("\tBad mode: " + safeMessageID + " - " + safeCheckIn + ": " + modeStr + " -  |  " + safeMsgFieldNumbered + "\r\n");
                                            badModeCt++;
                                        }
                                    }
                                }
                                else { modeCt++; }

                                if (latitude != 0)
                                {
                                    if (dupeFlag == 0)
                                    {
                                        mapCt++;
                                    }
                                    else
                                    {
                                        // Remove any lines containing the call sign
                                        if (mapString != null && checkIn != null)
                                        {
                                            RemoveLineContaining (mapString, checkIn);
                                            dupeRemoveCt++;
                                        }
                                    }
                                    if (mapString != null)
                                    {
                                        string safeCheckIn = checkIn ?? "";
                                        string safeBandStr = bandStr ?? "";
                                        string safeModeStr = modeStr ?? "";
                                        mapString.Append (safeCheckIn + "," + latitude + "," + longitude + "," + safeBandStr + "," + safeModeStr + "\r\n");
                                    }
                                }

                                // Console.WriteLine (checkIn + ":"+messageID+" - "+ ct + " - mapCt:" + mapCt + " - dupCt: " + dupCt);
                                // xml data

                                if (callSignTypo != "" && noScore == -1)
                                {
                                    reminderTxt += "\r\nCheck for a typo in your callsign in the checkin data: " + callSignTypo + " vs From:" + tempFromTxt + "\r\n";
                                    typoString.Append ("\t messageID " + messageID + " - " + checkIn + " vs " + tempFromTxt + "\r\n\t\t" + msgField + "\r\n");
                                    if (!newFormat) reminderTxt += "You may not have used the \"##\" marker at the beginning and end of your checkin data. See the examples.";
                                    if (isValidCallsign (callSignTypo) == "") reminderTxt += "You may have switched your call sign and name or left the call sign out";

                                }
                                if ((maidenheadGrid == "invalid") || (maidenheadGrid == "" && len > 8))
                                { reminderTxt += "\r\nCheck for a typo in your Maidenhead Grid (should be either xx##xx or xx##): " + msgField + "\r\n"; }
                                if (noScore == -1)
                                {
                                    if (isPerfect)
                                    {
                                        // reminderTxt += "\r\nThis is a copy of your extracted checkin data. \r\nMessage: \r\n" + msgField + "\r\n\r\nPerfect Message! Your score is 10.";
                                        if (checkinItems != null) reminderTxt += "\r\nThis is a copy of your extracted checkin data (in the correct format): \r\n## " + String.Join (" | ", checkinItems) + " ##\r\n\r\nPerfect Message! Your score is 10.";
                                        perfectScoreCt++;
                                    }
                                    else
                                    {

                                        if (msgFieldNumbered == "")
                                        {
                                            msgFieldNumbered = "Checkin data was not found";
                                            if (newFormat) msgFieldNumbered += " - probably because the checkin data did not start with ##.";
                                            else msgFieldNumbered += " - probably because you didn't use the ## marker at the ##beginning and end## of your checkin data or it was in the wrong place.";
                                            if (APRS > -1) msgFieldNumbered += " - or because the checkin came via APRS.";
                                        }
                                        if (checkinItems != null)
                                        {
                                            var tmpMsgField = msgField;
                                            if (newFormat) tmpMsgField = "## " + tmpMsgField + " ##";
                                            if (onlyOneMarker && newFormatStartOnly) tmpMsgField = "## " + tmpMsgField;
                                            if (onlyOneMarker && newFormatEndOnly) tmpMsgField = tmpMsgField + " ##";

                                            reminderTxt += "\r\n" + "\r\nThis is a copy of your extracted checkin data (in the correct format). \r\nCheckin Data: ## " + String.Join (" | ", checkinItems) + " ##" +
                                                "\r\nOriginal Message for comparison: " + originalMsgField + "\r\n\r\n" +
                                                "Your score is: " + score + "\r\n" + pointsOff +
                                                "\r\nRecommended format reminder in the Comment/Message field:\r\ncallSign, firstname, city, county, state/province/region, country, band, Mode, grid\r\n" +
                                               "Example: ##xxNxxx | Greg | Sugar City | Madison | ID | USA | HF | VARA HF | DN43du##\r\n" +
                                               "Example 2: ##DxNxx | Mario | TONDO | MANILA | NCR | PHL | 2M | VARA FM | PK04LO##\r\n" +
                                               "Example 3: ##xxNxx | Andre | Burnaby |  | BC | CAN | TELNET | SMTP | CN89ud\r\n\t       Weather is great today!##";
                                        }
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
                                string newMessageID = messageID ?? string.Empty; // Ensure non-nullable
                                if (newMessageID != null) // Redundant check, but included for clarity
                                {
                                    newMessageID = ScrambleWord (newMessageID) ?? string.Empty; // Ensure ScrambleWord result is non-null
                                }
                                string sendTo = checkIn ?? string.Empty; // Ensure non-nullable

                                // Tim Conroy, WB8HRO lives in an assisted living space and does not have easy access to 
                                // RF and put in a special request to send acknowledgements to his personal email address
                                if (sendTo == "WB8HRO") sendTo = "xyz191@live.com";
                                if (sendTo == xmlXsource || sendTo == netName) sendTo = xmlXsource + "@gmail.com";
                                // Console.WriteLine("before: "+messageID+   "    after: "+newMessageID);

                                if (isPerfect && noScore == -1)
                                {
                                    XElement? message_list = xmlPerfDoc.Descendants ("message_list").FirstOrDefault ();
                                    if (message_list == null)
                                    {
                                        message_list = new XElement ("message_list");
                                        xmlPerfDoc.Root!.Add (message_list);
                                    }
                                    message_list.Add (new XElement ("message",
                                        new XElement ("id", newMessageID),
                                        new XElement ("foldertype", "Fixed"),
                                        new XElement ("folder", "Outbox"),
                                        new XElement ("subject", $"{netName} acknowledgement {utcDate}"),
                                        new XElement ("time", utcDate),
                                        new XElement ("sender", netName),
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
                                        new XElement ("source", netName),
                                        new XElement ("unread", "True"),
                                        new XElement ("flags", "0"),
                                        new XElement ("messageoptions", "False|False|||||"),
                                        new XElement ("mime",
                                            $"Date: {utcDate}\r\n" +
                                            $"From: {netName}@winlink.org\r\n" +
                                            $"Subject: {netName} acknowledgement {utcDate}\r\n" +
                                            $"To: {sendTo}\r\n" +
                                            $"Message-ID: {newMessageID}\r\n" +
                                            $"X-Source: {xmlXsource}\r\n" +
                                            "X-Location: 43.845831N, 111.745744W(GPS)\r\n" +
                                            "MIME-Version: 1.0\r\n" +
                                            $"Thank you for checking in to the {netName}.\r\n" +
                                            $"{reminderTxt}\r\n\r\n" +
                                            $"Extracted Data: {noScore}\r\n" +
                                            $"   Latitude: {latitude}\r\n" +
                                            $"   Longitude: {longitude}\r\n" +
                                            $"   Band: {bandStr}\r\n" +
                                            $"   Mode: {modeStr}\r\n" +
                                            $"   Original Message ID: {messageID}\r\n" +
                                            $"\r\n{netName} Current Map: https://tinyurl.com/{netName}-Map\r\n" +
                                            $"Comments: https://tinyurl.com/{netName}-Additional-comments\r\n" +
                                            // $"{netName} Checkins Report: https://tinyurl.com/Checkins-Report\r\n" +
                                            $"checkins.csv: https://tinyurl.com/{netName}-CSV-checkins\r\n" +
                                            "mapfile.csv: https://tinyurl.com/Current-CSV-mapfile\r\n"
                                        )
                                    ));
                                }
                                else if (noScore == -1)
                                {
                                    XElement? message_list = xmlDoc.Descendants ("message_list").FirstOrDefault ();
                                    if (message_list == null)
                                    {
                                        message_list = new XElement ("message_list");
                                        xmlDoc.Root!.Add (message_list);
                                    }
                                    message_list.Add (new XElement ("message",
                                        new XElement ("id", newMessageID),
                                        new XElement ("foldertype", "Fixed"),
                                        new XElement ("folder", "Outbox"),
                                        new XElement ("subject", $"{netName} acknowledgement {utcDate}"),
                                        new XElement ("time", utcDate),
                                        new XElement ("sender", netName),
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
                                        new XElement ("source", netName),
                                        new XElement ("unread", "True"),
                                        new XElement ("flags", "0"),
                                        new XElement ("messageoptions", "False|False|||||"),
                                        new XElement ("mime",
                                            $"Date: {utcDate}\r\n" +
                                            $"From: {netName}@winlink.org\r\n" +
                                            $"Subject: {netName} acknowledgement {utcDate}\r\n" +
                                            $"To: {sendTo}\r\n" +
                                            $"Message-ID: {newMessageID}\r\n" +
                                            $"X-Source: {xmlXsource}\r\n" +
                                            "X-Location: 43.845831N, 111.745744W(GPS)\r\n" +
                                            "MIME-Version: 1.0\r\n" +
                                            $"Thank you for checking in to the {netName}.\r\n" +
                                            $"{reminderTxt}\r\n\r\n" +
                                            $"Extracted Data: {noScore}\r\n" +
                                            $"   Latitude: {latitude}\r\n" +
                                            $"   Longitude: {longitude}\r\n" +
                                            $"   Band: {bandStr}\r\n" +
                                            $"   Mode: {modeStr}\r\n" +
                                            $"   Original Message ID: {messageID}\r\n" +
                                            $"\r\n{netName} Current Map: https://tinyurl.com/{netName}-Map\r\n" +
                                            $"Comments: https://tinyurl.com/{netName}-Additional-comments\r\n" +
                                            // $"{netName} Checkins Report: https://tinyurl.com/Checkins-Report\r\n" +
                                            $"checkins.csv: https://tinyurl.com/{netName}-CSV-checkins\r\n" +
                                            "mapfile.csv: https://tinyurl.com/Current-CSV-mapfile\r\n"
                                        )
                                    ));
                                }
                                // Add the message message_list
                                // xmlDoc.Root.Add (messageElement);

                                // junk = 0; // just so i could put a debug here
                                dupeFlag = 0; // reset the duplicate flag

                            }
                            var tempCt = ct + dupCt + ackCt + removalCt;
                            //debug Console.Write("checkins:"+ct+"  duplicates:" + dupCt+"  removals:"+removalCt+"  acks:"+ackCt + "  combined:"+tempCt+"   actual total:"+msgTotal+"\n");
                            // missing from roster section. Check to see if the checkin is in the roster. 


                            startPosition = (roster != null && checkIn != null) ? roster.IndexOf (checkIn) : -1;
                            if (startPosition < 0)
                            {
                                checkIn = (checkIn != null) ? isValidCallsign (checkIn) ?? "" : "";
                                if (checkIn != "")
                                {
                                    string safeMessageID = messageID ?? "";
                                    Console.Write (checkIn + "  " + safeMessageID + " was not found in roster.txt. \n");

                                    // if (checkinCountryLong == "") checkinCountryLong = checkinCountry;
                                    string safeCheckinName = checkinName ?? "";
                                    string safeCheckinCity = checkinCity ?? "";
                                    string safeCheckinCounty = checkinCounty ?? "";
                                    string safeCheckinState = checkinState ?? "";
                                    string safeCheckinCountry = checkinCountry ?? "";
                                    string safeCheckinCountryLong = checkinCountryLong ?? "";
                                    string safeBandStr = bandStr ?? "";
                                    string safeModeStr = modeStr ?? "";
                                    string safeMaidenheadGrid = maidenheadGrid ?? "";
                                    newCheckIn = checkIn + "\t=countif(indirect(\"R[0]C[10]\",FALSE):indirect(\"R[0]C[63]\",FALSE),\">0\"&\"*\")\t" +
                                                 safeCheckinName + "\t" + safeCheckinCity + "\t" + safeCheckinCounty + "\t" +
                                                 safeCheckinState + "\t" + safeCheckinCountry + "\t" + safeCheckinCountryLong + "\t" +
                                                 safeBandStr + "\t" + safeModeStr + "\t" + safeMaidenheadGrid;

                                    if (newCheckIns != null)
                                    {
                                        newCheckIns.Append (newCheckIn + "\r\n");
                                    }

                                    // update roster.txt to contain the new checkin
                                    // File.AppendAllText("roster.txt", ";" + checkIn);
                                    roster = (roster ?? "") + ";" + checkIn;
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
                        // write to attachments.csv, header only goes once, add in the callsign
                        if (attachmentDecodedString.Length > 0)
                        {
                            startPosition = attachmentDecodedString.IndexOf ("\r\n") + 2;
                            attachmentDecodedString = attachmentDecodedString.Insert (startPosition, "\r\nCallsign " + checkIn + "\r\n");
                            len = attachmentDecodedString.Length - startPosition;
                            if (attachmentCSVct > 0 && len > 0) attachmentDecodedString = attachmentDecodedString.Substring (startPosition, len);
                            attachmentCSVct++;
                            // Console.WriteLine ("Decoded CSV content:\n" + attachmentDecodedString);
                            attachmentCSVwrite.WriteLine (attachmentDecodedString);
                        }
                    }

                } // this is a good spot to check the status of the record processed
                  // junk = 0;
            }
            var tempCT = 15;
            logWrite.WriteLine ("Current " + netName + " Checkins posted: " + utcDate);
            logWrite.WriteLine ("    Total Stations Checking in: " + (ct - dupCt) + "    Duplicates: " + dupCt + "    Total Checkins: " + ct + "    Removal Requests: " + removalCt);
            logWrite.WriteLine ("Non-" + netName + " checkin messages skipped: " + skipped + " (including " + ackCt + " acknowledgements and " + outOfRangeCt + " out of date range messages skipped.)\r\n");
            logWrite.WriteLine ("Total messages processed: " + msgTotal + " (includes " + removalCt + " removal(s) and " + ackCt + " acknowledgement(s).\r\n");
            logWrite.WriteLine ("Row " + tempCT + " should now be in " + netName + " Spreadsheet at row 1 of the checkin column to be recorded.");
            tempCT++;
            logWrite.WriteLine ("Row " + tempCT + " should now be in " + netName + " Spreadsheet at row 2 of the checkin column and is the copy" +
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
            if (csvString != null)
            {
                SortStringBuilder (csvString, "\r\n", 1);
            }
            // Console.WriteLine(csvString.ToString());
            csvWrite.WriteLine (csvString);

            // SortStringBuilder (skippedString, "\r\n", 1);
            // Console.WriteLine(csvString.ToString());

            if (mapString != null)
            {
                SortStringBuilder (mapString, "\r\n", 1);
                // Console.WriteLine(mapString);
                mapWrite.WriteLine (mapString);
            }

            if (addonString != null)
            {
                SortStringBuilder (addonString, "\r\n", 2);
                commentWrite.WriteLine (addonString);
            }
            // Add Google Sheets update
            Console.WriteLine ("\r\n\r\nDo you want to skip updating the google spreadsheet? (Y(es) or S(kip) - No is the default). \nIf you leave it blank, the program will assume don't skip\n");
            ConsoleKeyInfo keyInfo = Console.ReadKey (); // Reads a single key without needing Enter
            string? tmpInput = keyInfo.KeyChar.ToString (); // Convert the character to a string
                                                            // char tmpInput = keyInfo.KeyChar; // Get the character from the key press
                                                            // string? tmpInput = Console.ReadKey ();
            tmpInput = tmpInput.ToUpper ();
            // if (tmpInput == 'Y' || tmpInput == 'y') tmpInput = 'Y';

            // bool skipGoogleUpdate = string.IsNullOrWhiteSpace (tmpInput) || tmpInput.ToUpper () == "N";


            // string spreadsheetId = "1e0PJVqMGZhTzxwIVDf9if1dSSnG8y1U5Zf6pojB5Txc"; // Your new ID
            if (!string.IsNullOrWhiteSpace (spreadsheetId))
            {
                // Console.WriteLine ("\r\nGoogle Update is turned off\r\n");

                // ++++
                if (newCheckIns != null && tmpInput != "Y" && tmpInput != "S")
                {
                    Console.WriteLine ("\r\nGoogle Update is in process and the roster.txt file will be updated.\r\n");
                    UpdateGoogleSheet (netCheckinString, netAckString2, newCheckIns, removalString, spreadsheetId, endDate, credentialFilename, ct);
                }
                else
                {
                    Console.WriteLine ("\r\nGoogle Update is turned off\r\n The roster.txt file will not be updated to avoid losing new checkins.\r\n");
                }
            }

            xmlPerfDoc.Save (xmlPerfFile);
            xmlDoc.Save (xmlFile);
            // rewrite the roster.txt file
            roster = SortCommaDelimitedString (roster ?? string.Empty, ";").Trim (';');
            rosterString = "netName=" + netName + "// This is the name of the winlink net to be processed and the slashes need to be there up against the net name\r\n"
                    + "callSign=" + xmlXsource + "// This is the callsign (yours) that will be used as the x-source for the XML messages to be imported. Without this, the messages cannot be edited in Winlink after importing\r\n"
                    + "google spreadsheet id=" + spreadsheetId + "// this is required and must be valid to open the google sheet that acts as a database for the net\r\n"
                    + "credential filename=" + credentialFilename + "// this file is required to programatically open the google sheet that acts as a database for the net\r\n"
                    + "roster string=" + roster;
            if (tmpInput != "Y") File.WriteAllText (rosterFile, rosterString);// if tmpInput is not "Y", write the roster file
            else Console.WriteLine ("Roster.txt was not updated. If you want to update it, run the program again without skipping the google update.\r\n");

            if (duplicates.Length != 0) { logWrite.WriteLine (duplicates + "\r\n"); }
            if (bouncedString.Length != 0) { logWrite.WriteLine ("Messages that bounced: " + bouncedString); }
            if (newCheckIns != null && newCheckIns.Length != 0) { logWrite.WriteLine ("New Checkins should have been appended to the New tab and inserted into the yearly tab of the spreadsheet: \r\n" + newCheckIns); }
            if (skippedString.Length != 0) { logWrite.WriteLine ("Messages Skipped: \r\n" + skippedString); }
            if (removalString.Length != 0) { logWrite.WriteLine ("Requests to be Removed: " + removalString); }
            if (localWeatherCt > 0) { logWrite.WriteLine ("Local Weather Checkins: " + localWeatherCt); }
            if (severeWeatherCt > 0) { logWrite.WriteLine ("Severe Weather Checkins: " + severeWeatherCt); }
            if (incidentStatusCt > 0) { logWrite.WriteLine ("Incident Status Checkins: " + incidentStatusCt); }
            if (icsCt > 0) { logWrite.WriteLine ("ICS-213 Checkins: " + icsCt); }
            if (winlinkCkinCt > 0) { logWrite.WriteLine ("Winlink Check-in Checkins: " + winlinkCkinCt); }
            if (damAssessCt > 0) { logWrite.WriteLine ("Damage Assessment Checkins: " + damAssessCt); }
            if (fieldSitCt > 0) { logWrite.WriteLine ("Field Situation Report Checkins: " + fieldSitCt); }
            // if (quickMCt > 0) { logWrite.WriteLine ("Quick H&W: " + quickMCt); }
            if (qwmCt > 0) { logWrite.WriteLine ("Quick Welfare Message: " + qwmCt); }
            if (dyfiCt > 0) { logWrite.WriteLine ("Did You Feel It: " + dyfiCt); }
            if (rriCt > 0) { logWrite.WriteLine ("RRI Welfare Radiogram: " + rriCt); }
            if (miCt > 0) { logWrite.WriteLine ("Medical Incident: " + miCt); }
            if (aprsCt > 0) { logWrite.WriteLine ("APRS Checkins: " + aprsCt); }
            if (meshCt > 0) { logWrite.WriteLine ("Mesh Checkins: " + meshCt); }
            if (PosRepCt > 0) { logWrite.WriteLine ("Position Report Checkins: " + PosRepCt); }
            if (ICS201Ct > 0) { logWrite.WriteLine ("ICS 201 Checkins: " + ICS201Ct); }
            if (ICS202Ct > 0) { logWrite.WriteLine ("ICS 202 Checkins: " + ICS202Ct); }
            if (ICS203Ct > 0) { logWrite.WriteLine ("ICS 203 Checkins: " + ICS203Ct); }
            if (ICS204Ct > 0) { logWrite.WriteLine ("ICS 204 Checkins: " + ICS204Ct); }
            if (ICS205Ct > 0) { logWrite.WriteLine ("ICS 205a Checkins: " + ICS205Ct); }
            if (ICS205aCt > 0) { logWrite.WriteLine ("ICS 205a Checkins: " + ICS205aCt); }
            if (ICS206Ct > 0) { logWrite.WriteLine ("ICS 206 Checkins: " + ICS206Ct); }
            if (ICS208Ct > 0) { logWrite.WriteLine ("ICS 208 Checkins: " + ICS208Ct); }
            if (ICS210Ct > 0) { logWrite.WriteLine ("ICS 210 Checkins: " + ICS210Ct); }
            if (WBBMct > 0) { logWrite.WriteLine ("Welfare Bulletin Board Checkins: " + WBBMct); }

            if (radioGram > 0) { logWrite.WriteLine ("Radiogram Checkins: " + radioGramCt); }
            // next line is for the 20250203 exercise
            // logWrite.WriteLine ("Winlink Express: " + winlinkCt + "  PAT: " + patCt + "  RadioMail: " + radioMailCt + "  WoAD: " + woadCt + "\r\n");
            logWrite.WriteLine ("Total Plain and other Checkins: " + (ct - localWeatherCt - severeWeatherCt - incidentStatusCt - icsCt - winlinkCkinCt - damAssessCt - fieldSitCt - qwmCt - dyfiCt - rriCt - qwmCt - miCt - aprsCt - meshCt - PosRepCt - ICS201Ct - radioGramCt - ICS202Ct - ICS203Ct - ICS204Ct - ICS205Ct - ICS205aCt - ICS206Ct - ICS208Ct - ICS210Ct - WBBMct) + "\r\n");
            //var totalValidGPS = mapCt-noGPSCt;
            logWrite.WriteLine ("Total Checkins with a perfect message: (Not including " + noScoreCt + " NoScore's) " + perfectScoreCt);
            logWrite.WriteLine ("Total Checkins using the new format: " + newFormatCt);
            logWrite.WriteLine ("Total Checkins with a geolocation: " + (mapCt - noGPSCt));
            // logWrite.WriteLine ("Total Checkins with a geolocation: " + (mapCt - noGPSCt));
            if (exerciseCompleteCt > 0) { logWrite.WriteLine ("Successful Exercise Participation: " + exerciseCompleteCt); }

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
    public static int IndexOfNthSB (string input, char value, int startIndex, int nth)
    // This method finds the nth occurrence of a character in a string. 

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

    static (DateTime, DateTime, string) getNetDates (DateTime startDate, DateTime endDate, string weekDay, int netLength)
    {
        DateTime date = default;
        DateTime todayDate = DateTime.Today; // Use DateTime.Today instead of date.Today
        const int offset = 21;
        bool isValid = false;

        while (!isValid)
        {
            Console.WriteLine ("Enter the start date - must be within three weeks of today (yyyymmdd): ");
            string? input = Console.ReadLine ();

            if (DateTime.TryParseExact (input, "yyyyMMdd", null, System.Globalization.DateTimeStyles.None, out date))
            {
                // Check if date is within 21 days before and after today
                if (date < todayDate.AddDays (-offset) || date > todayDate.AddDays (offset))
                {
                    Console.WriteLine ($"Invalid date: {date:yyyyMMdd} Must be within three weeks of today. Please try again.");
                    continue;
                }

                // If we get here, the date is valid
                isValid = true;
            }
            else
            {
                Console.WriteLine ("Invalid date format. Please use yyyymmdd format.");
            }
        }

        // Set the return values
        startDate = date;
        endDate = date.AddDays (netLength);
        weekDay = date.DayOfWeek.ToString ();

        return (startDate, endDate, weekDay);
    }
    public static string? isValidCallsign (string? input)
    // This method validates a callsign using the regex pattern \b\d{0,2}[A-Z]{1,2}\d{1,2}[A-Z]{1,6}\b. 
    {
        if (input != null)
        {
            string pattern = @"\b\d{0,2}[A-Z]{1,2}\d{1,2}[A-Z]{1,6}\b";
            Regex regexCallSign = new Regex (pattern, RegexOptions.IgnoreCase);
            Match match = regexCallSign.Match (input);
            if (match.Success) return match.Value.ToUpper ();
        }
        return "";
    }
    public static string isValidName (string input)
    // This method removes numbers and newlines from a name field:
    {
        string pattern = @".*\d.*(\r?\n)?";
        input = input.ToUpper ().Replace ("(", "").Replace (")", "");
        string result = Regex.Replace (input, pattern, "", RegexOptions.Multiline);
        // Regex regexName = new Regex (pattern, RegexOptions.IgnoreCase);
        // Match match = regexName.Match (input);
        // if (match.Success)
        // { input = match.Value; }
        // else { input = ""; }
        return result.Trim ();
    }
    // public static string isValidCountry (string input)
    // {
    //     string pattern = "AUSTRIA,CANADA,ENGLAND,UK,GERMANY,NORWAY,NEW ZEALAND,PHILIPPINES,ROMANIA,SERBIA,ST LUCIA,TRINIDAD & TOBAGO,VENEZUELA,AFG,ALA,ALB,DZA,ASM,AND,AGO,AIA,ATA,ATG,ARG,ARM,ABW,AUS,AUT,AZE,BHS,BHR,BGD,BRB,BLR,BEL,BLZ,BEN,BMU,BTN,BOL,BIH,BWA,BVT,BRA,IOT,VGB,BRN,BGR,BFA,BDI,KHM,CMR,CAN,CPV,BES,CYM,CAF,TCD,CHL,CHN,CXR,CCK,COL,COM,COK,CRI,HRV,CUB,CUW,CYP,CZE,COD,DNK,DJI,DMA,DOM,TLS,ECU,EGY,SLV,GNQ,ERI,EST,SWZ,ETH,FLK,FRO,FSM,FJI,FIN,FRA,GUF,PYF,ATF,GAB,GMB,GEO,DEU,GHA,GIB,GRC,GRL,GRD,GLP,GUM,GTM,GGY,GIN,GNB,GUY,HTI,HMD,HND,HKG,HUN,ISL,IND,IDN,IRN,IRQ,IRL,IMN,ISR,ITA,CIV,JAM,JPN,JEY,JOR,KAZ,KEN,KIR,XXK,KWT,KGZ,LAO,LVA,LBN,LSO,LBR,LBY,LIE,LTU,LUX,MAC,MDG,MWI,MYS,MDV,MLI,MLT,MHL,MTQ,MRT,MUS,MYT,MEX,MDA,MNG,MNE,MSR,MAR,MOZ,MMR,NAM,NRU,NPL,NLD,NCL,NZL,NIC,NER,NGA,NIU,NFK,PRK,MKD,MNP,NOR,OMN,PAK,PLW,PSE,PAN,PNG,PRY,PER,PHL,PCN,POL,PRT,MCO,PRI,QAT,COG,REU,ROU,RUS,RWA,BLM,SHN,KNA,LCA,MAF,SPM,VCT,WSM,SMR,STP,SAU,SEN,SRB,SYC,SLE,SGP,SXM,SVK,SVN,SLB,SOM,ZAF,SGS,KOR,SSD,ESP,LKA,SDN,SUR,SJM,SWE,CHE,SYR,TWN,TJK,TZA,THA,TGO,TKL,TON,TTO,TUN,TUR,TKM,TCA,TUV,UGA,UKR,ARE,GBR,UMI,USA,URY,UZB,VUT,VAT,VEN,VNM,VIR,WLF,ESH,YEM,ZMB,ZWE,";
    // int found = pattern.IndexOf (input);
    // if (found == -1) input = "";
    // return input;
    // }

    // This method checks if a field (e.g., country, state) exists in a predefined list
    public static (string input, int found) isValidField (string input, string pattern, int found)
    {
        input = input.ToUpper ().Trim ().Trim ('.');
        pattern = pattern.ToUpper () + ",NA,NONE";
        found = pattern.IndexOf (input + ",");
        if (found == -1) found = pattern.IndexOf ("," + input);
        if (found == -1) return ("", -1); // Return null and -1 if not found
        return (input, found);
    }

    //static string ExtractMaidenheadGrid (string input)
    //// This method extracts a Maidenhead grid (e.g., DN43du) from the input string
    //{
    //    // Define the regular expression for Maidenhead grid locator (4 or 6 character grids)
    //    // Regex regex = new Regex (@"\b([A-R]{2}\d{2}[A-X]{0,2}[a-xA-X]{0,2})\b", RegexOptions.IgnoreCase);
    //    // Regex regex = new Regex (@"([A-R]{2}\d{2}[A-X]{0,2}[a-xA-X]{0,2})", RegexOptions.IgnoreCase);
    //    Regex regex = new Regex (@"([A-R]{2}\d{2}(?:[A-X]{2})?)", RegexOptions.IgnoreCase);
    //    input = input.Replace ("-", "").Replace (" ", "").Replace ("NOSCORE", "");
    //    // Search for a match in the input string
    //    Match match = regex.Match (input);

    //    if (match.Success)
    //    {
    //        if (match.Value.Length == 4 || match.Value.Length == 6)
    //        {
    //            return match.Value.ToUpper ();
    //        }
    //        Console.WriteLine ($"Invalid Maidenhead grid format: {match.Value} (length {match.Value.Length})");
    //        return "invalid";
    //    }

    //    return string.Empty; // Return an empty string if no match is found
    //}
    static string ExtractMaidenheadGrid (string input)
    {
        Regex regex = new Regex (@"([A-R]{2}\d{2}(?:[A-X]{2})?)", RegexOptions.IgnoreCase);
        input = input.Replace ("-", "").Replace (" ", "").Replace ("NOSCORE", "");
        Match match = regex.Match (input);

        if (match.Success)
        {
            if (match.Value.Length < 4 || match.Value.Length % 2 != 0 || match.Value.Length > 6)
            {
                if (match.Value.Length == 5) return match.Value.Substring (0, 4).ToUpper ();
                Console.WriteLine ("Invalid Maidenhead grid format - must be either 4 or 6 characters");
                return "invalid";
            }
            return match.Value.ToUpper ();
        }

        return string.Empty;
    }
    static (double, double) MaidenheadToGPS (string maidenhead)
    // This method converts a Maidenhead grid to latitude and longitude
    {
        if (maidenhead.Length < 4 || maidenhead.Length % 2 != 0)
        {
            Console.WriteLine ("Invalid Maidenhead grid format in: " + maidenhead);
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
        // This method adds field numbers to a delimited string for debugging:
        int ct = 1;
        string delimiter = "|";
        char chDelimiter = Convert.ToChar (delimiter);
        int startPosition = 0;
        if (input == "") return input;
        input = input.Replace (",", delimiter) + delimiter;
        while (ct < 10)
        {
            startPosition = input.IndexOf (delimiter, startPosition);
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
        //return input.Trim (chDelimiter);
        return input.Trim ('|');
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
    // This method scrambles a string (used for generating new message IDs)
    {
        char [] chars = new char [word.Length];
        // var now = TimeOnly.FromDateTime(DateTime.Now);
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
    // This utility method converts a string to a double
    {
        public static double ConvertToDouble (string value)
        {
            if (string.IsNullOrEmpty (value))
                return 0;
            if (double.TryParse (value, out double outVal))
            {
                if (double.IsNaN (outVal) || double.IsInfinity (outVal))
                    return 0;
                return outVal;
            }
            return 0;
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
            .Replace ("MTRS", "M")
            .Replace ("MTR", "M")
            .Replace ("METER", "M")
            .Replace ("TELENET", "TELNET") // common typo
            .Replace ("TELENT", "TELNET")
            .Replace (" ", "")
            .Replace ("(", "")
            .Replace (")", "")
            .Replace ("-", "")
            .Replace (".", "")
            .Replace ("O", "0")
            .Replace ("2N", "2M")
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
            case "EMAIL":
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
    static (string, string, string) checkMode (string input, string input2, string error)
    //static (DateTime, DateTime, string) getNetDates (DateTime startDate, DateTime endDate, string weekDay, int netLength)
    {
        if (input.IndexOf ("AREDN") > -1 || input.IndexOf ("MESH") > -1) input = "MESH";
        input = input
            .Replace ("TELNET POSTOFFICE", "SMTP")
            .Replace ("VERA", "VARA")
            .Replace ("WINLINK", "")
            .Replace ("-", " ")
            .Replace ("(", "")
            .Replace (")", "")
            .Replace (".", "")
            //.Replace ("STMP", "SMTP")
            //.Replace ("SMPT", "SMTP")
            .Trim ();
        switch (input)
        {
            case "SMTP":
            case "PACKET":
            case "ARDOP":
            case "ARDOPC":
            case "ARDOPCF":
            case "VARA FM":
            case "VARA HF":
            case "PACTOR":
            case "PACTOR P1":
            case "PACTOR P2":
            case "PACTOR P3":
            case "PACTOR P4":
            case "PACTOR I":
            case "PACTOR II":
            case "PACTOR III":
            case "PACTOR IV":
            case "INDIUM GO":
            case "MESH":
            case "APRS":
            case "ROBUST PACKET":
            case "JS8CALL":
                break;
            case "PACLET": // common typo
                error = "\r\nTypo in the mode field: " + input + " - should be PACKET";
                input = "PACKET";
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

            case "SMPT":
            case "TELNET":
            case "SPMT":
            case "STMP":
                error = "\r\nTypo in the mode field: " + input + " - should be SMTP";
                input = "SMTP";
                break;

            case "FM":
            case "FM VARA":
            case "VARA-FM":
            case "VARA FN":
            case "VARAFM":
                input = "VARA FM";
                break;

            case "ARDOP HF":
            case "ARDOPHF":
            case "HF ARDOP":
            case "HFARDOP":
                input = "ARDOP";
                break;


            case "VARAHF":
            case "VARAFH":
            case "HFVARA":
            case "HF": // this was commented once, not sure why
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
        return (input, error, string.Empty);
    }

    public static string? [] removeFieldNumber (string? [] input)
    {
        if (input != null)
        {
            int len = input.Length;
            string pattern = @"\s\d$";
            //for (int i = 0; i < len; i++)
            //{
            //    if (input [i] != null)
            //    {
            //        string item = input [i];
            //        input [i] = Regex.Replace (item, pattern, "");
            //    }
            //}
            for (int i = 0; i < len; i++)
            {
                string item = input [i] ?? ""; // Default to empty string if null
                input [i] = Regex.Replace (item, pattern, "");
            }

            return input;
        }
        return new string? [] { }; // Return empty array instead of null
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

    // private static readonly Regex CoordinateRegex = new (@"[-+]?([0-8]?\d(\.\d+)?|90(\.0+)?)\s*[°]?\s*([NS]),?\s*[-+]?((1[0-7]\d|\d{1,2})(\.\d+)?|180(\.0+)?)\s*[°]?\s*([EW])", RegexOptions.Compiled);
    //                                                   @"(?<lat>[-+]?([0-8]?\d(\.\d+)?|90(\.0+)?))\s*[°]?\s*(?<latHem>[NS]),?\s*(?<long>[-+]?((1[0-7]\d|\d{1,2})(\.\d+)?|180(\.0+)?))\s*[°]?\s*(?<longHem>[EW])";
    // private static readonly Regex CoordinateRegex = new (@"(?<lat>[-+]?([0-8]?\d(\.\d+)?|90(\.0+)?))\s*[°]?\s*(?<latHem>[NS]),?\s*(?<long>[-+]?((1[0-7]\d|\d{1,2})(\.\d+)?|180(\.0+)?))\s*[°]?\s*(?<longHem>[EW])(?<lat>[-+]?([0-8]?\d(\.\d+)?|90(\.0+)?))\s*[°]?\s*(?<latHem>[NS]),?\s*(?<long>[-+]?((1[0-7]\d|\d{1,2})(\.\d+)?|180(\.0+)?))\s*[°]?\s*(?<longHem>[EW])
    // private static readonly Regex CoordinateRegex = new (@" [-+]?([0-8]?\d(\.\d+)?|90(\.0+)?)\s*[°]?\s*[NS],?\s*[-+]?((1[0-7]\d|\d{1,2})(\.\d+)?|180(\.0+)?)\s*[°]?\s*([EW])", RegexOptions.Compiled);
    // (?<lat>[-+]?((?<latInt>[0-8]?\d)(\.\d+)?|90(\.0+)?))\s*[°]?\s*(?<latHem>[NS]),?\s*(?<long>[-+]?((?<longInt>1[0-7]\d|\d{1,2})(\.\d+)?|180(\.0+)?))\s*[°]?\s*(?<longHem>[EW])\s*(?:\([^)]+\))?
    // (?<lat>[-+]?((?<latInt>[0-8]?\d)(\.\d+)?|90(\.0+)?))\s*[°]?\s*(?<latHem>[NS]),?\s*(?<long>[-+]?(?:(?<longInt>1[0-7]\d|\d{1,2})(\.\d+)?|180(\.0+)?))\s*[°]?\s*(?<longHem>[EW])\s*(?:\([^)]+\))?(?<lat>[-+]?((?<latInt>[0-8]?\d)(\.\d+)?|90(\.0+)?))\s*[°]?\s*(?<latHem>[NS]),?\s*(?<long>[-+]?(?:(?<longInt>1[0-7]\d|\d{1,2})(\.\d+)?|180(\.0+)?))\s*[°]?\s*(?<longHem>[EW])\s*(?:\([^)]+\))? 
    // = @"(?<lat>[-+]?(?:(?<latInt>[0-8]?\d)(\.\d+)?|90(\.0+)?))\s*[°]?\s*(?<latHem>[NS]),?\s*(?<long>[-+]?(?:(?<longInt>1[0-7]\d|\d{1,2})(\.\d+)?|180(\.0+)?))\s*[°]?\s*(?<longHem>[EW])\s*(?:\([^)]+\))?";
    // (?<lat>[-+]?(?<latInt>[0-8]?\d)(\.\d+)?)\s*[°]?\s*(?<latHem>[NS]),?\s*(?<long>[-+]?(?<longInt>1[0-7]\d|\d{1,2})(\.\d+)?)\s*[°]?\s*(?<longHem>[EW])\s*(?:\([^)]+\))?
    // (?<lat>[-+]?(?:(?<latInt>[0-8]?\d|90)(?<latDec>\.\d+)?))\s*[°]?\s*(?<latHem>[NS]),?\s*(?<long>[-+]?(?:(?<longInt>1[0-7]\d|\d{1,2}|180)(?<longDec>\.\d+)?))\s*[°]?\s*(?<longHem>[EW])\s*(?:\([^)]+\))?(?<lat>[-+]?(?:(?<latInt>[0-8]?\d|90)(?<latDec>\.\d+)?))\s*[°]?\s*(?<latHem>[NS]),?\s*(?<long>[-+]?(?:(?<longInt>1[0-7]\d|\d{1,2}|180)(?<longDec>\.\d+)?))\s*[°]?\s*(?<longHem>[EW])\s*(?:\([^)]+\))?
    private static readonly Regex CoordinateRegex = new (@"(?<lat>[-+]?(?:(?<latInt>[0-8]?\d|90)(?<latDec>\.\d+)?))\s*[°]?\s*(?<latHem>[NS]),?\s*(?<long>[-+]?(?:(?<longInt>1[0-7]\d|\d{1,2}|180)(?<longDec>\.\d+)?))\s*[°]?\s*(?<longHem>[EW])\s*(?:\([^)]+\))?", RegexOptions.Compiled);


    // Latitude: (?<lat>[-+]?(?:(?<latInt>[0-8]?\d|90)(?<latDec>\.\d+)?))
    //    (?<lat>: Named group for full latitude (group 1).
    //    [-+]?: Optional sign.
    //    (?:...): Non-capturing group for the number.
    //    (?<latInt>[0-8]?\d|90): Group 2, latitude integer (0–89 or exactly 90).
    //    (?<latDec>\.\d+)?: Group 3, optional decimal (e.g., .229167 or .0).
    //    Captures: full latitude, integer, decimal, hemisphere.
    //    Separator: ,?\s*
    //    Optional comma and whitespace.
    // Longitude: (?<long>[-+]?(?:(?<longInt>1[0-7]\d|\d{1,2}|180)(?<longDec>\.\d+)?))
    //    (?<long>: Named group for full longitude (group 5).
    //    [-+]?: Optional sign.
    //    (?:...): Non-capturing group for the number.
    //    (?<longInt>1[0-7]\d|\d{1,2}|180): Group 6, longitude integer (0–179 or exactly 180).
    //    (?<longDec>\.\d+)?: Group 7, optional decimal (e.g., .708333 or .0).
    //    Captures: full longitude, integer, decimal, hemisphere.
    //    Hemisphere and Trailer:
    //    (?<latHem>[NS]): Group 4, latitude hemisphere.
    //    (?<longHem>[EW]): Group 8, longitude hemisphere.
    //    \s*(?:\([^)]+\))?: Optional parenthetical text (e.g., (GRID SQUARE)), non-capturing.
    // Group Count: 9 groups (0–8):
    //    0: Full match.
    //    1 (lat): Full latitude.
    //    2 (latInt): Latitude integer.
    //    3 (latDec): Latitude decimal or empty.
    //    4 (latHem): Latitude hemisphere.
    //    5 (long): Full longitude.
    //    6 (longInt): Longitude integer.
    //    7 (longDec): Longitude decimal or empty.
    //    8 (longHem): Longitude hemisphere

    static bool ExtractCoordinates (string input, out double latitude, out double longitude)
    {
        latitude = 0;
        longitude = 0;
        if (string.IsNullOrWhiteSpace (input))
            return false;

        Match match = CoordinateRegex.Match (input);
        if (!match.Success)
        {
            // Console.WriteLine ("Regex location match failed"+input.Substring(0,25));
            return false;
        }
        List<string> groupValues = new List<string> ();
        foreach (Group group in match.Groups)
        {
            groupValues.Add (group.Value);
        }

        // Console.WriteLine ("Group Values (inside method): " + string.Join (", ", groupValues));

        if (!double.TryParse (match.Groups [1].Value, out latitude) || !double.TryParse (match.Groups [5].Value, out longitude))
            return false;
        latitude = Math.Round (latitude, 7);
        if (string.Equals (match.Groups [4].Value, "S", StringComparison.OrdinalIgnoreCase))
            latitude = -latitude;
        longitude = Math.Round (longitude, 6);
        if (string.Equals (match.Groups [8].Value, "W", StringComparison.OrdinalIgnoreCase))
            longitude = -longitude;
        return true;
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

    public static string getMsgField (int startPosition, int endPosition, string messageID, string fileText, string? msgField)
    {
        int len = 0;
        if (startPosition > -1)
        {
            len = endPosition - startPosition;
        }
        else
        {
            len = 0;
        }

        if (len == 0)
        {
            string safeMsgField = msgField ?? "";
            Console.Write ("Nothing in the message field: " + messageID + " - " + safeMsgField + "\n");
            // try retrieving something from the from field
            startPosition = fileText.IndexOf ("FROM:");
            if (startPosition > -1) { startPosition += 6; }
            endPosition = fileText.IndexOf ("@", startPosition);
            len = endPosition - startPosition;
        }

        if (len == 0)
        {
            Console.Write ("Trying the subject field: " + messageID + "\n");
            // try retrieving something from the subject field
            startPosition = fileText.IndexOf ("SUBJECT:");
            if (startPosition > -1) { startPosition += 9; }
            endPosition = fileText.IndexOf ("\r\n", startPosition) - 1;
            len = fileText.Length;
            len = endPosition - startPosition;
        }

        if (len < 0)
        {
            Console.Write ("endPostion is less than startPosition in: " + messageID + "\n");
            Console.Write ("Break at line 2611ish. Press enter to continue. messageID =" + messageID);
            string? input = Console.ReadLine ();
            return "";
        }

        if (startPosition < 0 || len < 0 || startPosition + len > fileText.Length)
        {
            return ""; // Handle invalid indices
        }

        msgField = fileText.Substring (startPosition, len);
        // int lineBreak = fileText.IndexOf("=\r\n");
        if (messageID.Contains ("AD4BL"))
        {
            Console.WriteLine ("AD4BL sent:" + msgField);
            msgField = "## AD4BL | LINDA | NORTHSTARBOROUGH | FAIRBANKS | AK | USA | TELNET | SMTP | BP64ju ##";

        } // AD4BL has a lot of spaces in the message field for delimiters
        int count = (msgField.Length - msgField.Replace ("  ", "").Length) / 4;// did they use spaces for delimiters?
        if (count > 5 && msgField != "")
        {
            msgField = msgField
                .Replace ("      ", "  ")
                .Replace ("    ", "  ")
                .Replace ("   ", "  ")
                .Replace ("   ", "  ")
                .Replace ("'", "")
                ;
        }
        msgField = msgField
    .Replace ("I AM SAFE AND WELL.", "")
    .Replace ("EXERCISE", "")
    .Replace ("=20", "")
    .Replace ("=0A", "")
    .Replace ("=0D\r\n", "")
    .Replace ("=0", "")
    .Replace ("16. CONTACT INFO:", ",")
    .Replace ("<", "")
    .Replace ("\\", "|")
    .Replace (">", "")
    .Replace ("}", "|")
    .Trim ()
    .Replace ("   ", "  ")
    .Replace (", ", ",")
    .Replace (" ,", ",")
    .Replace ("[NO CHANGES OR EDITING OF THIS MESSAGE ARE ALLOWED]", "")
    .Replace ("[MESSAGE RECEIPT REQUESTED]", "")
    .Replace (" |", "|")
    .Replace ("| ", "|")
    .Replace ("\"", "")
    .Replace ("XX", "")
    .Trim ()
    .Trim (',')
    .Trim ('|')
    .Trim ('#')
    .Trim ('|')
    .Trim (',')
    .Trim ();
        // .Replace ("=\r\n", "")
        // .Replace ("/", "|")
        // .Replace ("  ", " ")
        // .Replace ("!", "|")
        // .Replace (":", "")
        // .Replace (";", ",")
        // //.Replace ("#", "")
        // .Replace ("}", "|")



        if (msgField.Count (c => c == '#') < 3) msgField.Replace ("#", "");
        return msgField; // msgField is guaranteed non-null here
    }

    public static string? [] getCheckinData (int len, string? msgField, string? []? checkinItems, bool newFormat)
    {
        if (msgField == null)
        {
            Console.WriteLine ("msgField is null, returning empty array.");
            return Array.Empty<string> ();
        }

        // Clean the input
        string checkIn = msgField.Replace ("(", "").Replace (")", "");
        // remove extra lines
        if (checkIn.IndexOf ("\r\n") != -1)
        {
            len = checkIn.IndexOf ("\r\n");
            if (len > 0)
            {
                checkIn = checkIn.Substring (0, len);
            }
        }
        // Process delimiters
        var (processedString, items, usedDelimiter) = ProcessDelimiters (checkIn, out string modifiedMsgField);
        checkIn = modifiedMsgField;
        checkinItems = items;
        // newFormat = usedDelimiter || checkIn.Contains ("|");

        // Process items if valid
        if (checkinItems?.Length >= 3)
        {
            checkinItems = removeFieldNumber (checkinItems);
        }
        else
        {
            checkinItems = null;
        }

        return checkinItems ?? Array.Empty<string> ();
    }

    // New method to process delimiters
    private static (string ProcessedString, string [] Items, bool UsedDelimiter) ProcessDelimiters (string input, out string modifiedInput)
    {
        // List of possible delimiters
        var delimiters = new []
        {
            "|",
            ",",
            "#",
            "\t",
            "/",
            "\\",
            "**",
            "  ",
            ":",
            ";",
            "!",
            "}",
            "{",
            " L ", // Space-L-Space
            "  L  ", // Space-Space-L-Space-Space
            " I ", // Space-I-Space
            "  I  " // Space-Space-I-Space-Space
        };

        modifiedInput = input;
        foreach (var delimiter in delimiters)
        {
            // Count occurrences of the delimiter
            int count = input.Split (new [] { delimiter }, StringSplitOptions.None).Length - 1;

            // Replace with "|" if delimiter appears more than 3 times
            if (count > 3)
            {
                modifiedInput = input
                    .Replace (delimiter, "|")
                    .Replace ("  ", ""); // Replace double spaces
                                         // .Replace (" ", ""); // Replace space - this breaks some things like cities with spaces
                return (modifiedInput, modifiedInput.Split ("|"), true);
            }
        }
        input = input
            .Replace ("  ", " "); // Replace double spaces with singles
                                  // .Replace (" ", ""); // Replace space  - this breaks some things like cities with spaces
                                  // No delimiter had more than 3 occurrences
        return (input, Array.Empty<string> (), false);
    }

    private static byte [] attachmentDecoded = Array.Empty<byte> (); // Initialize to empty array

    private static StringBuilder badBandString = new StringBuilder ();

    // private static StringBuilder badModeString = new StringBuilder ();

    // Method to update Google Sheet with check-in data
    private static void UpdateGoogleSheet (StringBuilder netCheckinString, StringBuilder netAckString2, StringBuilder newCheckins, StringBuilder removalString, string spreadsheetId, DateTime endDate, string credentialFilename, int checkinCount)
    {
        try
        {
            string credentialsPath = Path.Combine (Directory.GetCurrentDirectory (), credentialFilename);
            if (!System.IO.File.Exists (credentialsPath))
            {
                Console.WriteLine ($"Google Sheets credentials file not found at {credentialsPath}. Skipping upload.");
                return;
            }

            var service = new SheetsService (new BaseClientService.Initializer
            {
                HttpClientInitializer = GoogleCredential.FromFile (credentialsPath).CreateScoped (SheetsService.Scope.Spreadsheets),
                ApplicationName = "Winlink Checkins"
            });

            DateTime [] netMondays = Enumerable.Range (0, 53)
                .Select (w => new DateTime (endDate.Year, 1, 6).AddDays (w * 7))
                .ToArray ();

            DateTime adjustedEndDate = endDate;
            if (endDate.DayOfWeek == DayOfWeek.Sunday) adjustedEndDate = endDate.AddDays (-1);
            else if (endDate.DayOfWeek == DayOfWeek.Monday) adjustedEndDate = endDate.AddDays (-2);

            DateTime monday = netMondays
                .Where (m => m <= adjustedEndDate)
                .OrderByDescending (m => m)
                .First ();

            int weekNumber = Array.IndexOf (netMondays, monday) + 1;
            if (weekNumber < 1) weekNumber = 1;
            if (weekNumber > 53) weekNumber = 53;

            string columnLetter = GetColumnLetter (weekNumber);
            string yearTab = endDate.Year.ToString ();

            // Update header rows (rows 1 and 2) in a single batch
            int columnIndex = GetColumnIndex (columnLetter);
            var headerRequests = new List<Request>
        {
            new Request
            {
                UpdateCells = new UpdateCellsRequest
                {
                    Range = new GridRange
                    {
                        SheetId = GetSheetId(service, spreadsheetId, yearTab),
                        StartRowIndex = 0,
                        EndRowIndex = 1,
                        StartColumnIndex = columnIndex - 1,
                        EndColumnIndex = columnIndex
                    },
                    Rows = new List<RowData>
                    {
                        new RowData
                        {
                            Values = new List<CellData>
                            {
                                new CellData
                                {
                                    UserEnteredValue = new ExtendedValue
                                    {
                                        StringValue = netCheckinString.ToString().Replace('\t', '|')
                                    }
                                }
                            }
                        }
                    },
                    Fields = "userEnteredValue"
                }
            },
            new Request
            {
                UpdateCells = new UpdateCellsRequest
                {
                    Range = new GridRange
                    {
                        SheetId = GetSheetId(service, spreadsheetId, yearTab),
                        StartRowIndex = 1,
                        EndRowIndex = 2,
                        StartColumnIndex = columnIndex - 1,
                        EndColumnIndex = columnIndex
                    },
                    Rows = new List<RowData>
                    {
                        new RowData
                        {
                            Values = new List<CellData>
                            {
                                new CellData
                                {
                                    UserEnteredValue = new ExtendedValue
                                    {
                                        StringValue = netAckString2.ToString()
                                    }
                                }
                            }
                        }
                    },
                    Fields = "userEnteredValue"
                }
            }
        };

            var headerBatchUpdate = new BatchUpdateSpreadsheetRequest { Requests = headerRequests };
            service.Spreadsheets.BatchUpdate (headerBatchUpdate, spreadsheetId).Execute ();
            Console.WriteLine ($"Updated {yearTab}!{columnLetter}1 with pipe-delimited and {columnLetter}2 with semicolon-delimited netCheckinString");

            // Process removals: copy rows to "Removals" with today's date, then delete from yearly tab
            string removalStr = removalString.ToString ();
            if (!string.IsNullOrWhiteSpace (removalStr))
            {
                ProcessRemovals (removalStr, spreadsheetId, service, yearTab);
            }

            // Append new check-in data to "New" tab and insert into yearly tab
            string newCheckinsStr = newCheckins.ToString ();
            string? newTabRange = null;
            if (!string.IsNullOrWhiteSpace (newCheckinsStr))
            {
                newTabRange = AppendToNewTab (newCheckinsStr, spreadsheetId, service);
            }
            // Insert all new rows from "New" tab into yearly tab if newCheckins were appended
            if (!string.IsNullOrWhiteSpace (newCheckinsStr) && newTabRange != null)
            {
                // Parse the range of appended rows (e.g., "New!A2:A4")
                string [] rangeParts = newTabRange.Split ('!');
                string [] rowRange = rangeParts [1].Split (':');
                int startRowNum = int.Parse (rowRange [0].Substring (1)); // e.g., A2 -> 2
                int endRowNum = rowRange.Length > 1 ? int.Parse (rowRange [1].Substring (1)) : startRowNum; // e.g., A4 -> 4

                // Retrieve all newly appended rows
                string newTabRowRange = $"{rangeParts [0]}!A{startRowNum}:BO{endRowNum}";
                var getRequest = service.Spreadsheets.Values.Get (spreadsheetId, newTabRowRange);
                getRequest.ValueRenderOption = SpreadsheetsResource.ValuesResource.GetRequest.ValueRenderOptionEnum.FORMULA;
                var newTabResponse = getRequest.Execute ();
                var newTabRows = newTabResponse.Values ?? new List<IList<object>> ();

                if (newTabRows.Count > 0)
                {
                    // Fetch existing callsigns from yearly tab once
                    string yearlyRange = $"{yearTab}!A{StartRowIndex}:A";
                    var yearlyResponse = service.Spreadsheets.Values.Get (spreadsheetId, yearlyRange).Execute ();
                    var yearlyValues = yearlyResponse?.Values ?? new List<IList<object>> ();
                    var callsigns = yearlyValues.Select ((row, index) => new { Callsign = row.Count > 0 ? (row [0]?.ToString () ?? "").Trim () : "", RowIndex = StartRowIndex + index })
                                                .Where (x => !string.IsNullOrEmpty (x.Callsign))
                                                .ToList (); int yearTabSheetId = GetSheetId (service, spreadsheetId, yearTab);
                    int newTabSheetId = GetSheetId (service, spreadsheetId, NewTabName);

                    foreach (var newTabRow in newTabRows)
                    {
                        string newCallsign = newTabRow.Count > 0 ? (newTabRow [0]?.ToString () ?? "").Trim () : string.Empty;
                        if (string.IsNullOrEmpty (newCallsign))
                        {
                            Console.WriteLine ("New callsign is empty; skipping insertion.");
                            continue;
                        }

                        // Check for duplicates
                        var existing = callsigns.FirstOrDefault (x => string.Equals (x.Callsign, newCallsign, StringComparison.OrdinalIgnoreCase));
                        if (existing != null)
                        {
                            Console.WriteLine ($"Skipped insertion of {newCallsign} into {yearTab} tab - already exists at row {existing.RowIndex}");
                            continue;
                        }

                        // Find insertion point (alphabetical order)
                        int insertRow = StartRowIndex;
                        bool inserted = false;
                        for (int i = 0; i < callsigns.Count; i++)
                        {
                            if (string.Compare (callsigns [i].Callsign, newCallsign, StringComparison.OrdinalIgnoreCase) > 0)
                            {
                                insertRow = callsigns [i].RowIndex;
                                inserted = true;
                                break;
                            }
                        }
                        if (!inserted)
                        {
                            insertRow = callsigns.Count > 0 ? callsigns.Max (x => x.RowIndex) + 1 : StartRowIndex;
                        }

                        // Prepare batch request for insert, copy, and format
                        int newTabRowNum = startRowNum + newTabRows.IndexOf (newTabRow);
                        var requests = new List<Request>
                    {
                        // Insert a new row
                        new Request
                        {
                            InsertRange = new InsertRangeRequest
                            {
                                Range = new GridRange
                                {
                                    SheetId = yearTabSheetId,
                                    StartRowIndex = insertRow - 1,
                                    EndRowIndex = insertRow,
                                    StartColumnIndex = 0,
                                    EndColumnIndex = MaxColumnIndex
                                },
                                ShiftDimension = "ROWS"
                            }
                        },
                        // Copy the row from "New" tab (handles formula adjustments)
                        new Request
                        {
                            CopyPaste = new CopyPasteRequest
                            {
                                Source = new GridRange
                                {
                                    SheetId = newTabSheetId,
                                    StartRowIndex = newTabRowNum - 1,
                                    EndRowIndex = newTabRowNum,
                                    StartColumnIndex = 0,
                                    EndColumnIndex = MaxColumnIndex
                                },
                                Destination = new GridRange
                                {
                                    SheetId = yearTabSheetId,
                                    StartRowIndex = insertRow - 1,
                                    EndRowIndex = insertRow,
                                    StartColumnIndex = 0,
                                    EndColumnIndex = MaxColumnIndex
                                }
                            }
                        },
                        // Format the new row
                        new Request
                        {
                            RepeatCell = new RepeatCellRequest
                            {
                                Range = new GridRange
                                {
                                    SheetId = yearTabSheetId,
                                    StartRowIndex = insertRow - 1,
                                    EndRowIndex = insertRow,
                                    StartColumnIndex = 0,
                                    EndColumnIndex = MaxColumnIndex
                                },
                                Cell = new CellData
                                {
                                    UserEnteredFormat = new CellFormat
                                    {
                                        TextFormat = new TextFormat { FontSize = FontSize }
                                    }
                                },
                                Fields = "userEnteredFormat.textFormat.fontSize"
                            }
                        }
                    };

                        var batchUpdate = new BatchUpdateSpreadsheetRequest { Requests = requests };
                        service.Spreadsheets.BatchUpdate (batchUpdate, spreadsheetId).Execute ();
                        Console.WriteLine ($"Inserted {newCallsign} into {yearTab} tab at row {insertRow} for week {weekNumber} with adjusted formulas and {FontSize}pt font");

                        // Update in-memory callsigns list
                        callsigns.Add (new { Callsign = newCallsign, RowIndex = insertRow });
                        callsigns = callsigns.OrderBy (x => x.Callsign, StringComparer.OrdinalIgnoreCase).ToList ();
                    }
                }
                else
                {
                    Console.WriteLine ("Failed to retrieve new rows from 'New' tab for insertion.");
                }
            }
            else
            {
                Console.WriteLine ("No new check-ins this week; skipped the New and Yearly tab insertions.");
            }
        }
        catch (Google.GoogleApiException ex)
        {
            Console.WriteLine ($"Google API error: {ex.Message}, Status: {ex.HttpStatusCode}, Details: {ex.Error?.ToString ()}");
            Console.WriteLine ($"Stack Trace: {ex.StackTrace}");
        }
        catch (Exception ex)
        {
            Console.WriteLine ($"Unexpected error updating Google Sheet: {ex.Message}");
            Console.WriteLine ($"Stack Trace: {ex.StackTrace}");
        }
    }
    private static string? AppendToNewTab (string newCheckIns, string spreadsheetId, SheetsService service)
    {
        try
        {
            Console.WriteLine ($"Appending to '{NewTabName}' tab: '{newCheckIns}'");

            // Split into rows (each row is a new subscriber)
            var rows = newCheckIns.Split (new [] { Environment.NewLine }, StringSplitOptions.RemoveEmptyEntries);
            if (rows.Length == 0)
            {
                Console.WriteLine ("No valid new check-in entries found.");
                return null;
            }

            // Prepare the data for appending
            var values = new List<IList<object>> ();
            foreach (var row in rows)
            {
                var fields = row.Split ('\t');
                if (fields.Length > 10) // Trim Grid (11th field)
                {
                    fields [10] = fields [10].Trim ();
                }
                Console.WriteLine ($"Split into {fields.Length} fields: {string.Join (", ", fields)}");
                values.Add (fields.Select (x => (object)x).ToList ());
            }

            var valueRange = new ValueRange { Values = values };
            string range = $"{NewTabName}!A:A";
            var appendRequest = service.Spreadsheets.Values.Append (valueRange, spreadsheetId, range);
            appendRequest.ValueInputOption = SpreadsheetsResource.ValuesResource.AppendRequest.ValueInputOptionEnum.USERENTERED;
            var response = appendRequest.Execute ();

            // Format all appended rows with 8pt font
            string [] rangeParts = response.Updates.UpdatedRange.Split ('!');
            string rowRange = rangeParts [1].Split (':') [0];
            int startRowNum = int.Parse (rowRange.Substring (1));
            int endRowNum = startRowNum + rows.Length - 1;

            var formatRequests = new List<Request>
        {
            new Request
            {
                RepeatCell = new RepeatCellRequest
                {
                    Range = new GridRange
                    {
                        SheetId = GetSheetId(service, spreadsheetId, NewTabName),
                        StartRowIndex = startRowNum - 1,
                        EndRowIndex = endRowNum,
                        StartColumnIndex = 0,
                        EndColumnIndex = MaxColumnIndex
                    },
                    Cell = new CellData
                    {
                        UserEnteredFormat = new CellFormat
                        {
                            TextFormat = new TextFormat { FontSize = FontSize }
                        }
                    },
                    Fields = "userEnteredFormat.textFormat.fontSize"
                }
            }
        };
            var formatBatchUpdate = new BatchUpdateSpreadsheetRequest { Requests = formatRequests };
            service.Spreadsheets.BatchUpdate (formatBatchUpdate, spreadsheetId).Execute ();

            Console.WriteLine ($"Appended {rows.Length} new check-in(s) to '{NewTabName}' tab at {response.Updates.UpdatedRange} with {FontSize}pt font");
            return response.Updates.UpdatedRange;
        }
        catch (Exception ex)
        {
            Console.WriteLine ($"Error appending to '{NewTabName}' tab: {ex.Message}");
            return null;
        }
    }
    private static string GetColumnLetter (int weekNumber)
    {
        int columnNumber = weekNumber + 14; // Week 1 = O (15th), Week 53 = BO (67th)
        if (columnNumber <= 0) return "A";
        string columnLetter = "";
        do
        {
            columnLetter = (char)('A' + (columnNumber - 1) % 26) + columnLetter;
            columnNumber = (columnNumber - 1) / 26;
        } while (columnNumber > 0);
        return columnLetter;
    }

    private const int StartRowIndex = 6; // Starting row for subscriber data
    private const int MaxColumnIndex = 67; // A to BO
    private const int FontSize = 8; // 8pt font
    private const string NewTabName = "New";
    private const string RemovalsTabName = "Removals";

    private static int GetColumnIndex (string columnLetter)
    {
        int columnIndex = 0;
        foreach (char c in columnLetter)
        {
            columnIndex = columnIndex * 26 + (c - 'A' + 1);
        }
        return columnIndex;
    }

    private static void ProcessRemovals (string removalData, string spreadsheetId, SheetsService service, string yearTab)
    {
        try
        {
            // Split removalData into lines for multiple removals
            var removalLines = removalData.Split (new [] { Environment.NewLine }, StringSplitOptions.RemoveEmptyEntries);
            if (removalLines.Length == 0)
            {
                Console.WriteLine ("No valid removal entries found.");
                return;
            }

            // Fetch all data from yearly tab once (A6:BO)
            string yearlyRange = $"{yearTab}!A6:BO";
            var yearlyResponse = service.Spreadsheets.Values.Get (spreadsheetId, yearlyRange).Execute ();
            var values = yearlyResponse?.Values ?? new List<IList<object>> ();

            // Process each removal
            foreach (var line in removalLines)
            {
                // Extract callsign (first field before tab)
                string callsignToRemove = line.Split ('\t') [0].Trim ();
                Console.WriteLine ($"Processing removal for callsign: '{callsignToRemove}'");

                // Find the row to copy and delete
                int rowToDelete = -1;
                IList<object>? rowToCopy = null;
                for (int i = 0; i < values.Count; i++)
                {
                    if (values [i].Count > 0 && values [i] [0].ToString () == callsignToRemove)
                    {
                        rowToDelete = 6 + i; // 1-based row index in sheet
                        rowToCopy = values [i];
                        break;
                    }
                }

                if (rowToDelete == -1 || rowToCopy == null)
                {
                    Console.WriteLine ($"No matching callsign '{callsignToRemove}' found in {yearTab} tab for removal.");
                    continue; // Skip to next removal
                }

                // Modify column B (index 1) with today's date
                var modifiedRow = new List<object> (rowToCopy);
                while (modifiedRow.Count < 2) modifiedRow.Add (""); // Ensure at least 2 columns
                modifiedRow [1] = DateTime.Today.ToString ("yyyy-MM-dd"); // Column B gets today's date

                // Append the modified row to "Removals" tab
                var valueRange = new ValueRange
                {
                    Values = new List<IList<object>> { modifiedRow }
                };
                string removalsRange = "Removals!A:A";
                var appendRequest = service.Spreadsheets.Values.Append (valueRange, spreadsheetId, removalsRange);
                appendRequest.ValueInputOption = SpreadsheetsResource.ValuesResource.AppendRequest.ValueInputOptionEnum.USERENTERED;
                var appendResponse = appendRequest.Execute ();

                string [] rangeParts = appendResponse.Updates.UpdatedRange.Split ('!');
                int appendedRowNum = int.Parse (rangeParts [1].Split (':') [0].Substring (1));

                // Format: 8pt font for entire row, light green background for column A
                var formatRequests = new List<Request>
            {
                // 8pt font for A:BO
                new Request
                {
                    RepeatCell = new RepeatCellRequest
                    {
                        Range = new GridRange
                        {
                            SheetId = GetSheetId(service, spreadsheetId, "Removals"),
                            StartRowIndex = appendedRowNum - 1,
                            EndRowIndex = appendedRowNum,
                            StartColumnIndex = 0,
                            EndColumnIndex = 67 // A to BO
                        },
                        Cell = new CellData
                        {
                            UserEnteredFormat = new CellFormat
                            {
                                TextFormat = new TextFormat { FontSize = 8 }
                            }
                        },
                        Fields = "userEnteredFormat.textFormat.fontSize"
                    }
                },
                // Light green background for column A only
                new Request
                {
                    RepeatCell = new RepeatCellRequest
                    {
                        Range = new GridRange
                        {
                            SheetId = GetSheetId(service, spreadsheetId, "Removals"),
                            StartRowIndex = appendedRowNum - 1,
                            EndRowIndex = appendedRowNum,
                            StartColumnIndex = 0, // A
                            EndColumnIndex = 1    // A (exclusive, so just column A)
                        },
                        Cell = new CellData
                        {
                            UserEnteredFormat = new CellFormat
                            {
                                BackgroundColor = new Color
                                {
                                    Red = 217 / 255.0f,   // RGB 204, 255, 204
                                    Green = 234 / 255.0f, // 182, 215, 168
                                    Blue = 212 / 255.0f   // Red: 217, Green: 234, Blue: 212 
                                }
                            }
                        },
                        Fields = "userEnteredFormat.backgroundColor"
                    }
                }
            };
                var formatBatchUpdate = new BatchUpdateSpreadsheetRequest { Requests = formatRequests };
                service.Spreadsheets.BatchUpdate (formatBatchUpdate, spreadsheetId).Execute ();

                Console.WriteLine ($"Copied row for '{callsignToRemove}' to 'Removals' tab at {appendResponse.Updates.UpdatedRange} with today's date in column B, 8pt font, and light green A-column");

                // Delete the row from the yearly tab
                var deleteRequest = new BatchUpdateSpreadsheetRequest
                {
                    Requests = new List<Request>
                {
                    new Request
                    {
                        DeleteRange = new DeleteRangeRequest
                        {
                            Range = new GridRange
                            {
                                SheetId = GetSheetId(service, spreadsheetId, yearTab),
                                StartRowIndex = rowToDelete - 1, // 0-based index
                                EndRowIndex = rowToDelete,
                                StartColumnIndex = 0,
                                EndColumnIndex = 67 // A to BO
                            },
                            ShiftDimension = "ROWS"
                        }
                    }
                }
                };
                service.Spreadsheets.BatchUpdate (deleteRequest, spreadsheetId).Execute ();

                Console.WriteLine ($"Removed row with callsign '{callsignToRemove}' from {yearTab} tab at row {rowToDelete}");

                // Refresh yearly tab data after deletion
                yearlyResponse = service.Spreadsheets.Values.Get (spreadsheetId, yearlyRange).Execute ();
                values = yearlyResponse?.Values ?? new List<IList<object>> ();
            }
        }
        catch (Exception ex)
        {
            Console.WriteLine ($"Error processing removals: {ex.Message}");
        }
    }

    private static int GetSheetId (SheetsService service, string spreadsheetId, string sheetName)
    {
        var spreadsheet = service.Spreadsheets.Get (spreadsheetId).Execute ();
        var sheet = spreadsheet.Sheets.FirstOrDefault (s => s.Properties.Title == sheetName);
        if (sheet == null)
        {
            var addSheetRequest = new BatchUpdateSpreadsheetRequest
            {
                Requests = new List<Request>
            {
                new Request { AddSheet = new AddSheetRequest { Properties = new SheetProperties { Title = sheetName } } }
            }
            };
            var response = service.Spreadsheets.BatchUpdate (addSheetRequest, spreadsheetId).Execute ();
            int? sheetId = response.Replies [0].AddSheet.Properties.SheetId;
            if (!sheetId.HasValue)
            {
                throw new InvalidOperationException ($"SheetId for newly created sheet '{sheetName}' is unexpectedly null.");
            }
            return sheetId.Value;
        }
        int? existingSheetId = sheet.Properties.SheetId;
        if (!existingSheetId.HasValue)
        {
            throw new InvalidOperationException ($"SheetId for existing sheet '{sheetName}' is unexpectedly null.");
        }
        return existingSheetId.Value;
    }
}
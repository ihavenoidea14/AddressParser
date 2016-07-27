using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Text.RegularExpressions;
using System.Threading.Tasks;
using Excel = Microsoft.Office.Interop.Excel;

namespace AddressParser
{
    class AddressSplitter
    {
        private static string BuildPattern()
        {
            var pattern = "^" +                                                       // beginning of string
                            "(?<HouseNumber>\\d+)" +                                  // 1 or more digits
                            "(?:\\s+(?<StreetPrefix>" + GetStreetPrefixes() + "))?" + // whitespace + valid prefix (optional)
                            "(?:\\s+(?<StreetName>.*?))" +                            // whitespace + anything
                            "(?:" +                                                   // group (optional) {
                            "(?:\\s+(?<StreetType>" + GetStreetTypes() + "))" +       //   whitespace + valid street type
                            "(?:\\s+(?<StreetSuffix>" + GetStreetSuffixes() + "))?" + //   whitespace + valid street suffix (optional)
                            "(?:\\s+(?<Apt>.*))?" +                                   //   whitespace + anything (optional)
                            ")?" +                                                    // }
                            "$";                                                      // end of string

            return pattern;
        }

        private static string GetStreetPrefixes()
        {
            return "TE|NW|HW|RD|E|MA|EI|NO|AU|SE|GR|OL|W|MM|OM|SW|ME|HA|JO|OV|S|OH|NE|K|N|S|E|W|SE|SW|NE|NW|NORTH|SOUTH|EAST|WEST|NORTHEAST|NORTHWEST|NEAST|NWEST|SOUTHEAST|SOUTHWEST|SEAST|SWEST";
        }

        private static string GetStreetTypes()
        {
            return "ALLEY|ALLEE|ALY|ALLEY|ALLY|ALY|ANEX|ANEX|ANX|ANNEX|ANNX|ANX|ARCADE|ARC|ARC|ARCADE|AVENUE|AV|AVE|AVE|AVEN|AVENU|AVENUE|AVN|AVNUE|BAYOU|BAYOO|BYU|BAYOU|BEACH|BCH|BCH|BEACH|BEND|BEND|BND|BND|BLUFF|BLF|BLF|BLUF|BLUFF|BLUFFS|BLUFFS|BLFS|BOTTOM|BOT|BTM|BTM|BOTTM|BOTTOM|BOULEVARD|BLVD|BLVD|BOUL|BOULEVARD|BOULV|BRANCH|BR|BR|BRNCH|BRANCH|BRIDGE|BRDGE|BRG|BRG|BRIDGE|BROOK|BRK|BRK|BROOK|BRKS|BURG|BURG|BG|BURGS|BURGS|BGS|BYPASS|BYP|BYP|BYPA|BYPAS|BYPASS|BYPS|CAMP|CAMP|CP|CP|CMP|CANYON|CANYN|CYN|CANYON|CNYN|CAPE|CAPE|CPE|CPE|CAUSEWAY|CAUSEWAY|CSWY|CAUSWA|CSWY|CENTER|CEN|CTR|CENT|CENTER|CENTR|CENTRE|CNTER|CNTR|CTR|CENTERS|CENTERS|CTRS|CIRCLE|CIR|CIR|CIRC|CIRCL|CIRCLE|CRCL|CRCLE|CIRCLES|CIRCLES|CIRS|CLIFF|CLF|CLF|CLIFF|CLIFFS|CLFS|CLFS|CLIFFS|CLUB|CLB|CLB|CLUB|COMMON|COMMON|CMN|COMMONS|COMMONS|CMNS|CORNER|COR|COR|CORNER|CORNERS|CORNERS|CORS|CORS|COURSE|COURSE|CRSE|CRSE|COURT|COURT|CT|CT|COURTS|COURTS|CTS|CTS|COVE|COVE|CV|CV|COVES|COVES|CVS|CR|CREEK|CREEK|CRK|CRK|CRESCENT|CRESCENT|CRES|CRES|CRSENT|CRSNT|CREST|CREST|CRST|CROSSING|CROSSING|XING|CRSSNG|XING|CROSSROAD|CROSSROAD|XRD|CROSSROADS|CROSSROADS|XRDS|CURVE|CURVE|CURV|DALE|DALE|DL|DL|DAM|DAM|DM|DM|DIVIDE|DIV|DV|DIVIDE|DV|DVD|DRIVE|DR|DR|DRIV|DRIVE|DRV|DRIVES|DRIVES|DRS|ESTATE|EST|EST|ESTATE|ESTATES|ESTATES|ESTS|ESTS|EXPRESSWAY|EXP|EXPY|EXPR|EXPRESS|EXPRESSWAY|EXPW|EXPY|EXTENSION|EXT|EXT|EXTENSION|EXTN|EXTNSN|EXTENSIONS|EXTS|EXTS|FALL|FALL|FALL|FALLS|FALLS|FLS|FLS|FERRY|FERRY|FRY|FRRY|FRY|FIELD|FIELD|FLD|FLD|FIELDS|FIELDS|FLDS|FLDS|FLAT|FLAT|FLT|FLT|FLATS|FLATS|FLTS|FLTS|FORD|FORD|FRD|FRD|FORDS|FORDS|FRDS|FOREST|FOREST|FRST|FORESTS|FRST|FORGE|FORG|FRG|FORGE|FRG|FORGES|FORGES|FRGS|FORK|FORK|FRK|FRK|FORKS|FORKS|FRKS|FRKS|FORT|FORT|FT|FRT|FT|FREEWAY|FREEWAY|FWY|FREEWY|FRWAY|FRWY|FWY|GARDEN|GARDEN|GDN|GARDN|GRDEN|GRDN|GARDENS|GARDENS|GDNS|GDNS|GRDNS|GATEWAY|GATEWAY|GTWY|GATEWY|GATWAY|GTWAY|GTWY|GLN|GLN|GLENS|GLENS|GLNS|GREENS|GRNS|GROVE|GROV|GRV|GROVE|GRV|GROVES|GROVES|GRVS|HARBOR|HARB|HBR|HARBOR|HARBR|HBR|HRBOR|HARBORS|HARBORS|HBRS|HAVEN|HAVEN|HVN|HVN|HEIGHTS|HT|HTS|HTS|HIGHWAY|HIGHWAY|HWY|HIGHWY|HIWAY|HIWY|HWAY|HWY|HILL|HILL|HL|HL|HILLS|HILLS|HLS|HLS|HOLLOW|HLLW|HOLW|HOLLOW|HOLLOWS|HOLW|HOLWS|INLET|INLT|INLT|ISLAND|IS|IS|ISLAND|ISLND|ISLANDS|ISLANDS|JUNCTION|JCT|JCT|JCTION|JCTN|JUNCTION|JUNCTN|JUNCTON|JUNCTIONS|JCTNS|JCTS|JCTS|JUNCTIONS|KEY|KEY|KY|KY|KEYS|KEYS|KYS|KYS|KNOLL|KNL|KNL|KNOL|KNOLL|KNOLLS|KNLS|KNLS|KNOLLS|LAKE|LK|LK|LAKE|LAKES|LKS|LKS|LAKES|LAND|LAND|LAND|LANDING|LANDING|LNDG|LNDG|LNDNG|LANE|LANE|LN|LN|LIGHT|LGT|LGT|LIGHT|LIGHTS|LIGHTS|LGTS|LOAF|LF|LF|LOAF|LOCK|LCK|LCK|LOCK|LOCKS|LCKS|LCKS|LOCKS|LODGE|LDG|LDG|LDGE|LODG|LODGE|LOOP|LOOPS|MALL|MALL|MALL|MANOR|MNR|MNR|MANOR|MANORS|MANORS|MNRS|MNRS|MEADOW|MEADOW|MDW|MEADOWS|MDW|MDWS|MDWS|MEADOWS|MEDOWS|MEWS|MEWS|MEWS|MILL|MILL|ML|MILLS|MILLS|MLS|MISSION|MISSN|MSN|MSSN|MOTORWAY|MOTORWAY|MTWY|NECK|NCK|NCK|NECK|ORCHARD|ORCH|ORCH|ORCHARD|ORCHRD|OVAL|OVAL|OVAL|OVL|OVERPASS|OVERPASS|OPAS|PARK|PARK|PARK|PRK|PARKS|PARKS|PARK|PARKWAY|PARKWAY|PKWY|PARKWY|PKWAY|PKWY|PKY|PARKWAYS|PARKWAYS|PKWY|PKWYS|PASS|PASS|PASS|PASSAGE|PASSAGE|PSGE|PATH|PATH|PATH|PATHS|PIKE|PIKE|PIKE|PIKES|PINE|PINE|PNE|PINES|PINES|PNES|PNES|PLACE|PL|PL|PLAIN|PLAIN|PLN|PLN|PLAINS|PLAINS|PLNS|PLNS|PLAZA|PLAZA|PLZ|PLZ|PLZA|PT|PT|POINTS|POINTS|PTS|PTS|PORT|PORT|PRT|PRT|PORTS|PORTS|PRTS|PRTS|PRAIRIE|PR|PR|PRAIRIE|PRR|RADIAL|RAD|RADL|RADIAL|RADIEL|RADL|RAMP|RAMP|RAMP|RANCH|RANCH|RNCH|RANCHES|RNCH|RNCHS|RAPID|RAPID|RPD|RPD|RAPIDS|RAPIDS|RPDS|RPDS|REST|REST|RST|RST|RIDGE|RDG|RDG|RDGE|RIDGE|RIDGES|RDGS|RDGS|RIDGES|RIVER|RIV|RIV|RIVER|RVR|RIVR|ROAD|RD|RD|ROAD|ROADS|ROADS|RDS|RDS|ROUTE|ROUTE|RTE|ROW|ROW|ROW|RUE|RUE|RUE|RUN|RUN|RUN|SHOAL|SHL|SHL|SHOAL|SHOALS|SHLS|SHLS|SHOALS|SHORE|SHOAR|SHR|SHORE|SHR|SHORES|SHOARS|SHRS|SHORES|SHRS|SKYWAY|SKYWAY|SKWY|SPRING|SPG|SPG|SPNG|SPRING|SPRNG|SPRINGS|SPGS|SPGS|SPNGS|SPRINGS|SPRNGS|SPUR|SPUR|SPUR|SPURS|SPURS|SPUR|SQUARE|SQ|SQ|SQR|SQRE|SQU|SQUARE|SQUARES|SQRS|SQS|SQUARES|STATION|STA|STA|STATION|STATN|STN|STRAVENUE|STRA|STRA|STRAV|STRAVEN|STRAVENUE|STRAVN|STRVN|STRVNUE|STREAM|STREAM|STRM|STREME|STRM|STREET|STREET|ST|STRT|ST|STR|STREETS|STREETS|STS|SUMMIT|SMT|SMT|SUMIT|SUMITT|SUMMIT|TERRACE|TER|TER|TERR|TERRACE|THROUGHWAY|THROUGHWAY|TRWY|TRACE|TRACE|TRCE|TRACES|TRCE|TRACK|TRACK|TRAK|TRACKS|TRAK|TRK|TRKS|TRAFFICWAY|TRAFFICWAY|TRFY|TRAIL|TRAIL|TRL|TRAILS|TRL|TRLS|TRAILER|TRAILER|TRLR|TRLR|TRLRS|TUNNEL|TUNEL|TUNL|TUNL|TUNLS|TUNNEL|TUNNELS|TUNNL|TURNPIKE|TRNPK|TPKE|TURNPIKE|TURNPK|UNDERPASS|UNDERPASS|UPAS|UNION|UN|UN|UNION|UNIONS|UNIONS|UNS|VALLEY|VALLEY|VLY|VALLY|VLLY|VLY|VALLEYS|VALLEYS|VLYS|VLYS|VIADUCT|VDCT|VIA|VIA|VIADCT|VIADUCT|VIEW|VIEW|VW|VW|VIEWS|VIEWS|VWS|VWS|VILLAGE|VILL|VLG|VILLAG|VILLAGE|VILLG|VILLIAGE|VLG|VILLAGES|VLGS|VLGS|VILLE|VILLE|VL|VL|VISTA|VIS|VIS|VIST|VISTA|VST|VSTA|WALKS|WALK|WALL|WAY|WY|WAY|WAYS|WELL|WL|WELLS|WLS";
        }

        private static string GetStreetSuffixes()
        {
            return "N|S|E|W|SE|SW|NE|NW|NORTH|SOUTH|EAST|WEST|NORTHEAST|NORTHWEST|NEAST|NWEST|SOUTHEAST|SOUTHWEST|SEAST|SWEST";
        }

        public static StreetAddress Parse(string address)
        {
            if (string.IsNullOrEmpty(address))
                return new StreetAddress();

            StreetAddress result;
            var input = address.ToUpper();

            var re = new Regex(BuildPattern());
            if (re.IsMatch(input))
            {
                var m = re.Match(input);
                result = new StreetAddress
                {
                    HouseNumber = m.Groups["HouseNumber"].Value,
                    StreetPrefix = m.Groups["StreetPrefix"].Value,
                    StreetName = m.Groups["StreetName"].Value,
                    StreetType = m.Groups["StreetType"].Value,
                    StreetSuffix = m.Groups["StreetSuffix"].Value,
                    Apt = m.Groups["Apt"].Value,
                };
            }
            else
            {
                result = new StreetAddress
                {
                    StreetName = input,
                };
            }
            return result;
        }

        static void Main(string[] args)
        {
            Excel.Application excel = new Excel.Application();
            Excel.Workbook wb = excel.Workbooks.Open("C:\\users\\jaridwade\\Desktop\\Book1.xlsx");
            excel.Visible = true;
            Excel.Worksheet ws = wb.ActiveSheet;
            int lastrow = ws.Range["A" + ws.Rows.Count].End[Excel.XlDirection.xlUp].Row;
            for(int i = 2; i <= lastrow; i++)
            {
                StreetAddress addy = new StreetAddress();
                string input = ws.Cells[i, 2].Value.ToString();
                input = input.Replace(".", "");
                input = input.Replace(",", "");
                // input = input.Replace("#", "");
                addy = Parse(input);
                string housenum = addy.HouseNumber;
                string pre = addy.StreetPrefix;
                string stname = addy.StreetName;
                string streettype = addy.StreetType;
                string suffix = addy.StreetSuffix;
                string aptnum = addy.Apt;
                string[] parts = { housenum, pre, stname, streettype, suffix, aptnum };

                try
                {
                    ws.Cells[i, 5] = parts[0];
                    ws.Cells[i, 4] = parts[1];
                    ws.Cells[i, 3] = parts[2];
                    ws.Cells[i, 3] = (ws.Cells[i, 3].Value + " " + parts[3]).ToString().Trim();
                    ws.Cells[i, 4] = (ws.Cells[i, 4].Value + " " + parts[4]).ToString().Trim();
                    ws.Cells[i, 3] = (ws.Cells[i, 3].Value + " " + parts[5]).ToString().Trim();
                }
                catch (Exception e)
                {
                    continue;
                }
            }
        }
    }
}
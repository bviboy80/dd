'''
ParseFixedWidth

Parse a fixed width file and export the results as a CSV

Format: P:\AST\Due Diligence\scripts\Due Diligence Excel layout_V6.xlsx

HEADERS USED:                FLAT FILE HEADERS:                  EXCEL HEADERS
     
= FileTransmissionDate       File Transmission Date              
= UPRR Job Number            UPRR Job # (Optional)               
= LT                         UPRR LT #                           XRX Acct Seq 
= Company Name               AST Issue Name                      Issue Name
= Company Number             AST Company #                       Company
= ASTSourceFileDate          AST SourceFile Date (YYYYMMDD)      
= Account Number             AST Account #                       Account
= NameAddress1               Name/Address 1                      Name/Address1
= NameAddress2               Name/Address 2                      Name/Address2
= NameAddress3               Name/Address 3                      Name/Address3
= NameAddress4               Name/Address 4                      Name/Address4
= NameAddress5               Name/Address 5                      Name/Address5
= NameAddress6               Name/Address 6                      Name/Address6
= NameAddress7               Name/Address 7                      Name/Address7
= NameAddress8                                                   
= Verification Code          PIN                                 
= Filler                     Filler                              
= Mailing City               City                                City
= Zip                        Zip                                 Zip
= Mailing State              State                               State
= Shares                     Shares                              Eligible Shares
= Certified                  Certified Mailing Indicator         
= LetterCode                 Letter Code                         
= Sequence                   Sequence Number                     
= Escheatment State          EscheatmentStateName                Eligibility State
------------------------------------------------------------------------------------
  AddressType                                      
'''

___author__ = "Stephen E. Lee"
___creationdate__ = "October 20, 2015"
__version__ = 1.3
___modifiedBy__ = "Shaun T. Thomas"
___modifiedDate__ = "January 28, 2018"
___modifiedFor__ = "AST Escheatment_Due_Dilligence"

'''
2017-11-28  Added Escheatment State Name field (20 characters)

2018-01-28  Split data into Static Data and Address Data files.
            Static Data contains with original data. Address Data 
            contains the LT Number and the address modified for 
            improved processing in Mail Manager.
'''
             
import sys
import os
import csv
import subprocess
import re
import struct
import openpyxl



def main(argv=None):

    inFile = os.path.abspath(sys.argv[1])
    outputDir = os.path.dirname(inFile)
    filename = os.path.basename(inFile)
    filename_noext = ".".join(filename.split(".")[:-1])
    file_extension = filename.split(".")[-1]
    
    letterCode = chooseLetterCode()
    
    data_fields = create_data_fields()
    static_hdr = [f[0] for f in data_fields]
    us_dict = create_us_dict()
    
    recordsList = []

    if file_extension.upper() in ["XLS", "XLSX"]:
        print "Formatting data from Excel....\n"
        csv_file = convertXLStoCSV(outputDir, inFile)
        recordsList = processXLSfromCSV(csv_file, us_dict, data_fields, static_hdr)
    else:
        print "Formatting data from Text....\n"
        recordsList = processTXT(inFile, static_hdr)

    print "Sorting records....\n"
    records_dict = createRecordsDict(recordsList, static_hdr, letterCode)
    
    print "Writing records to CSV....\n"
    writeRecordsToCSV(outputDir, records_dict, static_hdr)                     
    
    print "Writing records to Excel....\n"
    writeRecordsToXLS(outputDir, filename_noext, records_dict, static_hdr)                       
    
    writeCountsToTXT(outputDir, filename, records_dict)            


def replaceNonAsciiChars(text):
    """ Convert byte text to unicode chars. Replace non-ASCII,
    the "replacement", "non-breaking space" and "Broken Bar" 
    chars. Convert chars back into bytes. """
    
    char_text = text.decode("utf-8", errors='replace').replace(u'\ufffd', " ")
    replace_nbspace = char_text.replace(u'\u00A6', " ")
    replace_bknbar = replace_nbspace.replace(u'\u00A0', " ")
    latin_text = replace_bknbar.encode("latin-1")
    ascii_char = latin_text.decode("ascii", errors='replace').replace(u'\ufffd', " ")
    byte_text = ascii_char.encode("ascii")
    return byte_text


def create_data_fields():    
    return [
        ['FileTransmissionDate', re.compile(r'FileTransmissionDate')],
        ['UPRR Job Number', re.compile(r'UPRR\s?Job\s?Number')],
        ['LT', re.compile(r'XRX\s?Acct\s?Seq')],
        ['Company Name', re.compile(r'Issue\s?Name')],
        ['Company Number', re.compile(r'Company')],
        ['ASTSourceFileDate', re.compile(r'ASTSourceFileDate')],
        ['Account Number', re.compile(r'Account\s?(Number)?')],
        ['NameAddress1', re.compile(r'Name/?\s?Address\s?1')],
        ['NameAddress2', re.compile(r'Name/?\s?Address\s?2')],
        ['NameAddress3', re.compile(r'Name/?\s?Address\s?3')],
        ['NameAddress4', re.compile(r'Name/?\s?Address\s?4')],
        ['NameAddress5', re.compile(r'Name/?\s?Address\s?5')],
        ['NameAddress6', re.compile(r'Name/?\s?Address\s?6')],
        ['NameAddress7', re.compile(r'Name/?\s?Address\s?7')],
        ['NameAddress8', re.compile(r'Name/?\s?Address\s?8')],
        ['Verification Code', re.compile(r'Verification\s?Code')],
        ['Filler', re.compile(r'Filler')],
        ['Mailing City', re.compile(r'City')],
        ['Zip', re.compile(r'Zip')],
        ['Mailing State', re.compile(r'(Mailing\s?)?State')],
        ['Shares', re.compile(r'Eligible\s?Shares')],
        ['Certified', re.compile(r'Certified')],
        ['LetterCode', re.compile(r'Letter\s?Code')],
        ['Sequence', re.compile(r'Sequence')],
        ['Escheatment State', re.compile(r'(Escheatment|Eligibility)\s?State')],
        ['AddressType', re.compile(r'Address\s?Type')]
        ]

        
def create_us_dict():
    """ Lookup table for the Escheatment/Eligibility State. 
    Match the abbreviation in the data and replace with 
    the value from the table. """    
    return {"AL" : "Alabama", "AK" : "Alaska", "AZ" : "Arizona", 
        "AR" : "Arkansas", "CA" : "California", "CO" : "Colorado", 
        "CT" : "Connecticut", "DE" : "Delaware", "FL" : "Florida", 
        "GA" : "Georgia", "HI" : "Hawaii", "ID" : "Idaho", "IL" : "Illinois",
        "IN" : "Indiana", "IA" : "Iowa", "KS" : "Kansas", "KY" : "Kentucky", 
        "LA" : "Louisiana", "ME" : "Maine", "MD" : "Maryland", 
        "MA" : "Massachusetts", "MI" : "Michigan", "MN" : "Minnesota", 
        "MS" : "Mississippi", "MO" : "Missouri", "MT" : "Montana",
        "NE" : "Nebraska", "NV" : "Nevada", "NH" : "New Hampshire", 
        "NJ" : "New Jersey", "NM" : "New Mexico", "NY" : "New York", 
        "NC" : "North Carolina", "ND" : "North Dakota", "OH" : "Ohio", 
        "OK" : "Oklahoma", "OR" : "Oregon", "PA" : "Pennsylvania",
        "RI" : "Rhode Island", "SC" : "South Carolina", "SD" : "South Dakota",
        "TN" : "Tennessee", "TX" : "Texas", "UT" : "Utah", "VT" : "Vermont",
        "VA" : "Virginia", "WA" : "Washington", "WV" : "West Virginia", 
        "WI" : "Wisconsin", "WY" : "Wyoming", "AS" : "American Samoa", 
        "DC" : "District of Columbia", "FM" : "Micronesia", 
        "FM" : "Federated States of Micronesia", "GU" : "Guam", 
        "MH" : "Marshall Islands", "MP" : "Northern Mariana Islands", 
        "PW" : "Palau", "PR" : "Puerto Rico", "VI" : "U.S. Virgin Islands"}  

 
def chooseLetterCode():
    """ Select the LetterCode. Used to select the appropriate 
    template in GMC.
    """
    letterCodeList = ["A", "AC", "FA", "FC", "R", "RC"]
    choiceString = "\n".join(["\nSelect letter code", 
                              "A = DDA", "AC = DDAC",
                              "FA = DDFA", "FC = DDFC",
                              "R = DDR", "RC = DDRC"])
    print choiceString 
    letterCode = raw_input("-->  ").upper()
    while letterCode not in letterCodeList:
        print "\nNot a valid choice !!!\n"
        print choiceString
        letterCode = raw_input("-->  ").upper()
    return letterCode
    
 
def processTXT(inFile, static_hdr):
    """ Split each line into fields. Insert an extra address field  
    to match the format of the excel file. Replace the letter 
    code in the data with the selected letter code. Add foreign  
    or domestic to the record. Format the zip code. """
    
    recordsList = []
    
    formatStr = "8s 6s 9s 40s 12s 8s 19s 40s 40s 40s 40s 40s 40s 40s 4s 36s 40s 9s 2s 14s 1s 2s 6s 20s"
    fieldstruct = struct.Struct(formatStr)
    parse = fieldstruct.unpack_from
    
    with open(inFile, 'rb') as o:
        for seq, line in enumerate(o, start=1):
            ascii_line = replaceNonAsciiChars(line)
            outputLine = [" ".join(x.split()) for x in parse(ascii_line)]
            
            # Remove extra spaces in fields 
            # Add blanks for AddrLine8 and AddressType
            outputLine.insert(static_hdr.index("NameAddress7"), "")
            outputLine.append("")
            recordsList.append(outputLine)
            
    return recordsList
    

def processXLSfromCSV(csv_file, us_dict, data_fields, static_hdr):
    """ Retrieve the records from the Excel. Arrange and 
    format the data to the standard layout. """

    recordsList = []
    
    with open(csv_file, 'rb') as csv_handle:
        csv_file_rdr = csv.reader(csv_handle, quoting=csv.QUOTE_ALL)
        csv_hdr = csv_file_rdr.next()    
        
        field_indxs = getFieldsIndxs(csv_hdr, data_fields)
        
        for line in csv_file_rdr:
            if line[:5].count("") == "":
                break
            else:
                dataRow = getFieldValuesFromLine(line, field_indxs)
                
                ''' Look up the Escheatment State abbreviation in the 
                US State Table. Replace with the full State Name. '''
                escheat_idx = static_hdr.index("Escheatment State")
                escheat_state = dataRow[escheat_idx]
                dataRow[escheat_idx] = us_dict[escheat_state]    
                
                ''' If LT number is missing, create a substitute by 
                combining the company and account numbers. '''
                if dataRow[static_hdr.index("LT")] == "":
                    compNo = dataRow[static_hdr.index("Company Number")]
                    acctNo = dataRow[static_hdr.index("Account Number")]
                    dataRow[static_hdr.index("LT")] = "{}{}".format(compNo, acctNo)
                    
                recordsList.append(dataRow)
                
    os.remove(csv_file)
    return recordsList   

    
def convertXLStoCSV(outputDir, excel_file):
    """ Save Excel file to CSV. Create a temp VB script. 
    Run the script on the Excel file with Command Line. """
        
    excel_name = os.path.basename(excel_file)
    csv_name = ".".join(excel_name.split(".")[:-1]) + ".csv"
    csv_file = os.path.join(outputDir, csv_name)
    
    # Temp XLS to CSV vb script
    temp_vb_script = os.path.join(outputDir, "tempXLStoCSV-DO_NOT_TOUCH.vbs")
    vb_string = "\n".join(
    ["Dim oExcel",
     "Set oExcel = CreateObject(\"Excel.Application\")",
     "oExcel.DisplayAlerts = False",
     "Dim oBook",
     "Set oBook = oExcel.Workbooks.Open(Wscript.Arguments.Item(0))",
     "oBook.SaveAs WScript.Arguments.Item(1), 6",
     "oBook.Close False",
     "oExcel.Quit",
     "WScript.Echo \"Done\""
    ])

    with open(temp_vb_script, 'wb') as t:
        t.write(vb_string)
        
    # Process using command line. Delete vb script after processing.
    subprocess.call(["cscript", temp_vb_script, excel_file, csv_file])
    os.remove(temp_vb_script)
    return csv_file


def getFieldsIndxs(csv_hdr, data_fields):
    """ Compare the csv/excel header to the standard 
    data layout. Get the indexes of the fields that match. 
    Indexes are collected in standard layout order"""
    
    data_cols = [' '.join(col.split()) for col in csv_hdr]
    
    found_list = []
    not_found_list = []
    field_indxs = []
    
    for fieldName, pattern in data_fields:
        for col in data_cols:
            if pattern.match(col):
                found_list.append(col)
                field_found = pattern.match(col).group(0)
                found_idx = data_cols.index(field_found)
                field_indxs.append(found_idx)
                break
        else:
            not_found_list.append(col)
            field_indxs.append('')
            continue
    
    print "Fields found: {}".format("\r\n".join(found_list))
    print ""
    print "Fields not found: {}".format("\r\n".join(not_found_list))
    print ""
    return field_indxs


def getFieldValuesFromLine(line, field_indxs):
    """ Get the field values from the csv/excel data row 
    using the field indexes. Remove any non ASCII chars 
    and/or extra spaces from each field. """
    
    asciiRow = []
    for idx in field_indxs:
        if idx == "":
            asciiRow.append("")
        else:   
            field = line[idx]
            ascii_field = replaceNonAsciiChars(field)
            reduced_field = " ".join(ascii_field.split())
            asciiRow.append(reduced_field)  
            
    return asciiRow


def createRecordsDict(recordsList, static_hdr, letterCode):    
    """ Sort Record List into mailing categories. 
    Fix Zip for domestic addresses as needed. """
    
    records_dict = {"MEX" : [], "CAN" : [],
                    "FGN" : [], "DOM" : []}
    
    foreignData = []
    
    for dataRow in recordsList:
        # Add the Letter Code to the record.     
        dataRow[static_hdr.index("LetterCode")] = letterCode
        
        if dataRow[static_hdr.index("Mailing State")] == "FO":
            dataRow[static_hdr.index("Zip")] = ""
            foreignData.append(dataRow)
        else:
            zip = dataRow[static_hdr.index("Zip")]
            if len(zip) > 5 and "-" not in zip:
                zip = "{}-{}".format(zip[:5], zip[5:])
            dataRow[static_hdr.index("Zip")] = zip
            dataRow[static_hdr.index("AddressType")] = "DOM"
            records_dict["DOM"].append(dataRow)
    
    sortForeignByCountry(foreignData, records_dict, static_hdr)
    
    return records_dict


def sortForeignByCountry(foreignData, records_dict, static_hdr):
    """ Sort foreign data into Mexico, Canada and  
    other foreign countries by reviewing the mailing 
    city and last of the address lines. """

    canada_provinces = "\\b" + "\\b|\\b".join(["Canada","Alberta","Calgary","Edmonton",
    "Strathcona County","British Columbia","Vancouver","Surrey","Burnaby","Manitoba",
    "Winnipeg","Brandon","Springfield","New Brunswick","Moncton","Saint John","Fredericton",
    "Newfoundland and Labrador","St. John's","Conception Bay South","Mount Pearl",
    "Northwest Territories","Yellowknife","Hay River","Inuvik","Nova Scotia",
    "Halifax","Sydney","Lunenburg","Nunavut","Iqaluit","Arviat","Rankin Inlet",
    "Ontario","Toronto","Ottawa","Mississauga","Prince Edward Island","Charlottetown",
    "Summerside","Stratford","Quebec","Montreal","Quebec City","Laval","Saskatchewan",
    "Saskatoon","Regina","Prince Albert","Yukon","Whitehorse","Dawson City","Faro"]) + "\\b" 
    
    mexico_states_cities = "\\b" + "\\b|\\b".join(["Chihuahua","Sonora","Coahuila",
    "Durango","Oaxaca","Tamaulipas","Jalisco","Zacatecas","Baja California Sur",
    "Chiapas","Veracruz","Baja California","Nuevo Leon","Guerrero","San Luis Potosi",
    "Michoacan","Sinaloa","Campeche","Quintana Roo","Yucatan","Puebla","Guanajuato",
    "Nayarit","Tabasco","Mexico","Hidalgo","Queretaro","Colima","Aguascalientes",
    "Morelos","Tlaxcala","Ciudad de Mexico","Mexico City","Ecatepec","Guadalajara",
    "Puebla","Juarez","Tijuana","Leon","Monterrey","Zapopan","Nezahualcoyotl","Culiacan",
    "Chihuahua","Naucalpan","Merida","San Luis Potosi","Aguascalientes","Hermosillo",
    "Saltillo","Mexicali","Guadalupe","Acapulco","Tlalnepantla","Cancun","Queretaro",
    "Chimalhuacan","Torreon","Morelia","Reynosa","Tlaquepaque","Tuxtla Gutierrez",
    "Durango","Toluca","Ciudad Lopez Mateos","Cuautitlan Izcalli","Ciudad Apodaca","Matamoros",
    "San Nicolas de los Garza","Veracruz","Xalapa","Tonala","Mazatlan","Irapuato",
    "Nuevo Laredo","Xico","Villahermosa","General Escobedo","Celaya","Cuernavaca","Tepic",
    "Ixtapaluca","Ciudad Victoria","Ciudad Obregon","Tampico","Ciudad Nicolas Romero",
    "Ensenada","Coacalco de Berriozabal","Santa Catarina","Uruapan","Gomez Palacio",
    "Los Mochis","Pachuca","Oaxaca","Soledad de Graciano Sanchez","Tehuacan","Ojo de Agua",
    "Coatzacoalcos","Campeche","Monclova","La Paz","Nogales","Buenavista","Puerto Vallarta",
    "Tapachula","Ciudad Madero","San Pablo de las Salinas","Chilpancingo","Poza Rica",
    "Chicoloapan de Juarez","Ciudad del Carmen","Chalco de Diaz Covarrubias","Jiutepec",
    "Salamanca","San Luis Rio Colorado","Cuautla","Ciudad Benito Juarez","Chetumal",
    "Piedras Negras","Playa del Carmen","Zamora","Cordoba","San Juan del Rio","Colima",
    "Ciudad Acuna","Manzanillo","Zacatecas","Veracruz","Ciudad Valles","Guadalupe",
    "San Pedro Garza Garcia","Naucalpan","Fresnillo","Orizaba","Miramar","Iguala",
    "Delicias","Ciudad de Villa de alvarez","Ciudad Cuauhtemoc","Navojoa","Guaymas",
    "Minatitlan","Cuautitlan","Texcoco","Hidalgo del Parral","Tepexpan","Tulancingo"]) + "\\b"
    
    city_idx = static_hdr.index("Mailing City")
    addr_start = static_hdr.index("NameAddress1")
    addr_end = static_hdr.index("NameAddress8")+1

    canada_zip_pattern = re.compile(r'\b[ABCEGHJ-NPRSTVXY][0-9][ABCEGHJ-NPRSTV-Z](\s|-)?[0-9][ABCEGHJ-NPRSTV-Z][0-9]\b', flags=re.IGNORECASE)
    canada_prov_pattern = re.compile(canada_provinces, flags=re.IGNORECASE)
    canada_major_cities_pattern = re.compile(r'\bCANADA\b|\bTORONTO\b|\bONTARIO\b|\bQUEBEC\b|\bALBERTA\b|\bMONTREAL\b', flags=re.IGNORECASE)
    ontario_quebec_abbv_pattern = re.compile(r'(\bON\b)|(\bQC\b)\s\b[ABCEGHJ-NPRSTVXY][0-9][ABCEGHJ-NPRSTV-Z]', flags=re.IGNORECASE)
    
    mexico_states_cities_pattern = re.compile(mexico_states_cities, flags=re.IGNORECASE)
    
    # Sort by countries
    sorted_foreign = sorted(foreignData, key=lambda row: row[city_idx]) 
    
    # Sort to Canada, Mexico and Other Foreign
    for record in sorted_foreign:
        addrfields = [f for f in record[addr_start:addr_end] if f.upper() not in ["","NULL"]]
        last_addr_field = addrfields[-1]
        record_city = record[city_idx]
        
        # Check for country pattern and check that 
        # it is not another country that has similar names
        if (re.search(canada_zip_pattern, record_city) or \
        re.search(ontario_quebec_abbv_pattern, record_city) or \
        re.search(canada_major_cities_pattern, last_addr_field) or \
        re.search(canada_prov_pattern, record_city)) and not\
        (re.search(r'\bLONDON\b|\bUK\b|\bUNIT\b|\bGBR\b|\bAUS(TRALIA)?\b', record_city, flags=re.IGNORECASE)):
        
            record[static_hdr.index("AddressType")] = "CAN"
            records_dict["CAN"].append(record)
            
        elif (re.search(mexico_states_cities_pattern, record_city)) and not \
        (re.search(r'\bSPAIN\b|\bESPANA\b|\bITALY\b', record_city, flags=re.IGNORECASE)):
        
            record[static_hdr.index("AddressType")] = "MEX"
            records_dict["MEX"].append(record)
        else:    
            record[static_hdr.index("AddressType")] = "FGN"
            records_dict["FGN"].append(record)     


def createMMAddress(line, static_hdr):
    """ Extract needed fields from the static data to create 
    the BCC data. Move last line of the Name/Address fields 
    to the Delivery or Alternate Address position. """
    
    ''' Create new line from data '''
    mmfields = [
        line[static_hdr.index("NameAddress1")],
        line[static_hdr.index("NameAddress2")],
        line[static_hdr.index("NameAddress3")],
        line[static_hdr.index("NameAddress4")],
        line[static_hdr.index("NameAddress5")],
        line[static_hdr.index("NameAddress6")],
        line[static_hdr.index("NameAddress7")],
        line[static_hdr.index("NameAddress8")],
        line[static_hdr.index("Mailing City")],
        line[static_hdr.index("Mailing State")],
        line[static_hdr.index("Zip")]
        ] 
        
    namesAndStreet = mmfields[:8]
    namesAndStreet_NoBlanks = [f for f in namesAndStreet if f.upper() not in ["","NULL"]]
    city = mmfields[8]
    cityStateZip = mmfields[8:]

    ''' Find last line of name/address lines 
    and move Delivery/Alternate Addr position '''
    if line[static_hdr.index("AddressType")] in ["MEX","CAN","FGN"]:
        return formatForeignAddress(namesAndStreet_NoBlanks, city)   
    else:    
        return formatDomesticAddress(namesAndStreet_NoBlanks, cityStateZip)
        

def formatForeignAddress(namesAndStreet_NoBlanks, city):
    namesAndStreet_NoBlanks.append(city)
    spaceShift = [""] * (8-len(namesAndStreet_NoBlanks))
    deliveryAddr = ""
    alternateAddr = ""
    cityStateZip = ["","",""]
    
    return namesAndStreet_NoBlanks + spaceShift + [deliveryAddr, alternateAddr] + cityStateZip
        
        
def formatDomesticAddress(namesAndStreet_NoBlanks, cityStateZip):        
    
    if len(namesAndStreet_NoBlanks) < 2:
        spaceShift = [""] * (8-len(namesAndStreet_NoBlanks))
        deliveryAddr = ""
        alternateAddr = ""
        return namesAndStreet_NoBlanks + spaceShift + [deliveryAddr, alternateAddr] + cityStateZip
    else:
        apt_pattern = re.compile(r'^((#|B(UI)?LD(IN)?G|SUITE|LOT|UNIT|FLOOR|R(OO)?M|AP(ARTMEN)?T).+|(\d{1,4}\s?\w)|(\d{1,3}(ST|ND|RD|TH)?\s?FL(OO)?R?))$', flags=re.IGNORECASE)
        deliveryAddr = namesAndStreet_NoBlanks[-1]
        
        addrIdx = -2 if apt_pattern.match(deliveryAddr) and len(namesAndStreet_NoBlanks) > 2 else -1
        alternateAddr = namesAndStreet_NoBlanks[-1] if apt_pattern.match(deliveryAddr) and len(namesAndStreet_NoBlanks) > 2 else ""
        nameLines = namesAndStreet_NoBlanks[:addrIdx]
        spacesShift = [""] * (8 - len(nameLines))
        deliveryAddr = namesAndStreet_NoBlanks[addrIdx]

        return nameLines + spacesShift + [deliveryAddr, alternateAddr] + cityStateZip


def writeRecordsToCSV(outputDir, records_dict, static_hdr):                        
    with open(os.path.join(outputDir, "AddressData.csv"), 'wb') as a:
        with open(os.path.join(outputDir, "StaticData.dat"), 'wb') as s:
                
                addr_hdr = ["IM barcode Digits", "OEL", "Sack and Pack Numbers",
                "Presort Sequence", "Full Name", "Name2", "Name3", 
                "Name4","Name5","Name6","Name7","Name8","Delivery Address",
                "Alternate 1 Address","City","State","ZIP+4","LTNo","SEQ"]
                
                AddressOut = csv.writer(a, quoting=csv.QUOTE_ALL)
                AddressOut.writerow(addr_hdr)
                
                StaticOut = csv.writer(s, quoting=csv.QUOTE_ALL)
                StaticOut.writerow(static_hdr)
                
                # Combine records into single list
                all_records = []
                for category in ["MEX", "CAN", "FGN", "DOM"]:
                    all_records.extend(records_dict[category])
                
                for seq, line in enumerate(all_records, start=1):
                    line[static_hdr.index("Sequence")] = seq

                    # Write address and Static Data
                    LTNo = line[static_hdr.index("LT")]
                    mmAddress = createMMAddress(line, static_hdr)                            
                    AddressOut.writerow(["", "", "", seq] + mmAddress + [LTNo] + [seq])
                    
                    StaticOut.writerow(line)
    

def writeRecordsToXLS(outputDir, filename_noext, records_dict, static_hdr) :                        
    wb = openpyxl.Workbook()
    ws = wb.create_sheet("Records", 0)
    ws.append(static_hdr)
    
    # Combine records into single list
    all_records = []
    for category in ["MEX", "CAN", "FGN", "DOM"]:
        all_records.extend(records_dict[category])

    for seq, row in enumerate(all_records, start=1):
        row[static_hdr.index("Sequence")] = seq
        ws.append(row)
    
    # Save work book
    outExcel = os.path.join(outputDir, "{}_rev.xlsx".format(filename_noext))
    wb.save(outExcel)
    
    
def writeCountsToTXT(outputDir, filename, records_dict):                        
    ''' Get counts for reporting. Print to screen 
    and write to text file. '''
    
    domesticCount = len(records_dict["DOM"])
    mexicoCount = len(records_dict["MEX"])
    canadaCount = len(records_dict["CAN"])
    otherCount = len(records_dict["FGN"])
    foreignCount = mexicoCount + canadaCount + otherCount
    totalCount = domesticCount + foreignCount
    
    countsReport = "\r\n".join(
    ["Filename: {}".format(filename),
     "Domestic count: {}".format(domesticCount),
     "Foreign count: {}".format(foreignCount),
     "Total Records: {}".format(totalCount),
     "",
     "Mexico count: {}".format(mexicoCount),
     "Canada count: {}".format(canadaCount),
     "Other count: {}".format(otherCount),
    ])
    
    print countsReport
    with open(os.path.join(outputDir, "COUNTS.txt"),'wb') as c:
        c.write(countsReport)
   
    
if __name__ == "__main__":
    sys.exit(main())

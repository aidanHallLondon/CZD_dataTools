
sourceFileName = 'Reading data to upload.xlsx'
outputFileName = "data.xlsx"
# name of the sheet we wnat in the workbook
mainSheetName = "coded"
# members contain lots of spaces and messy chars that mess up heading names
# this regEx is used to remove them
shortnameRegExpPattern = "[^|a-zA-Z0-9]+"
# this is rounding a float before multiplying by 100 to get %age
percentageRoundingPlaces = 6
# width of columns with 0 or 1 boolean data in them
# small sizes make the sheet more manageable but the headings harder to read
booleanColumnWidth = 3
# how many tokens do we add columns for - override in the data if you want
selectTopTokens = 20

# these are all the columns that we are working with on the spreadsheet
#
# change the order of these to affect the order in the output
# --- watch out for the last row that does not have a comma
#
# name :: is the EXACT name in row one on the sourceSheet (how we find it)
#
# type :: use  this to change the type of member values - applies to boolean column names and output data
#   string  : no change, uses the raw value
#   integer : set this to integer for columns containing only integer values
#   head    : used to split a compound name and take just the head e.g. "01 | ark" =>"01"
#             the 'delimiter' defaults to '|' but can be set
# other types can be added
#
# outputName :: use this to change the name in the output
#
# output :: do you want it in the results or not
# (True or False : no quotes!)
#
# getMembers :: do you want boolean cols for EVERY unique member
#   'members' : make boolean columns for each member in the main colum
#   'tokens'  : split a delimeted string of tokens in each cell into members and return the top 10
#               delimeter - defaults to |
#               selectTopTokens - defaults to 10
# memberColumnSuffix : the memberColumnSuffix to use on member columns along with the member value
# '' will omit the memberColumnSuffix
# include an underscore
# defaults to the outputName+'_'
#
# includeInvalidRows : defaults to false but can be set to true to include all rows
#
# Then the code then adds these columns
#
# srcIndex - which column is this in the source
# outputIndex - which column is this in the outpuy
# memberDict - all the members


def createColumnData():
    return [
        {'name': 'ID', 'outputName': 'id', 'output': True},
        {'name': 'Valid', 'output': True, 'getMembers': 'members',
            'includeInvalidRows': True, 'type': 'integer'},
        {'name': 'Invalid reason', 'outputName': 'InvalidReason',
            'output': True, 'getMembers': 'members', 'includeInvalidRows': True},
        {'name': 'TotalChildIncPreg', 'output': True, 'getMembers': 'members'},
        {'name': 'age1', 'output': True, 'type': 'head', 'delimiter': '|'},
        {'name': 'age2', 'output': True, 'type': 'head', 'delimiter': '|'},
        {'name': 'age3', 'output': True, 'type': 'head', 'delimiter': '|'},
        {'name': 'age4', 'output': True, 'type': 'head', 'delimiter': '|'},
        {'name': 'age5', 'output': True, 'type': 'head', 'delimiter': '|'},
        {'name': 'age6', 'output': True, 'type': 'head', 'delimiter': '|'},
        {'name': ' ', 'output': False},
        {'name': 'Youngest_age', 'output': True, 'outputName': 'YoungestAge',
            'getMembers': 'members', 'type': 'head'},
        {'name': 'YoungestBorn_age', 'output': True,
            'outputName': 'YoungestBornAge', 'getMembers': 'members', 'type': 'head'},
        {'name': 'Under72incInUtero', 'output': True,
            'memberColumnSuffix': 'under72inc', 'getMembers': 'members', 'type': 'integer'},
        {'name': '01 | pregnant', 'outputName': 'ageCount_01', 'output': True},
        {'name': '02 | 00-02', 'outputName': 'ageCount_02', 'output': True},
        {'name': '03 | 03-06', 'outputName': 'ageCount_03', 'output': True},
        {'name': '04 | 07-12', 'outputName': 'ageCount_04', 'output': True},
        {'name': '05 | 13-23', 'outputName': 'ageCount_05', 'output': True},
        {'name': '07 | 24-35', 'outputName': 'ageCount_07', 'output': True},
        {'name': '08 | 36-47', 'outputName': 'ageCount_08', 'output': True},
        {'name': '09 | 48-71', 'outputName': 'ageCount_09', 'output': True},
        {'name': '10 | 72+', 'outputName': 'ageCount_10', 'output': True},
        {'name': ' ', 'output': False},
        {'name': 'agePrefStart', 'AgePrefStart': '',
            'output': True, 'getMembers': 'members'},
        {'name': 'WhyAgeStart', 'output': True},
        {'name': 'WhyAgeStart_tokens', 'output': True,
            'getMembers': 'tokens', 'delimeter': '|', 'selectTopTokens': selectTopTokens},
        {'name': 'FrequencyYoungest', 'output': True, 'getMembers': 'members'},
        {'name': 'FrequencyPref', 'output': True, 'getMembers': 'members'},
        {'name': 'BenefitsChild', 'output': True},
        {'name': 'BenefitsChild_tokens', 'output': True,
            'getMembers': 'tokens', 'delimeter': '|', 'selectTopTokens': selectTopTokens},
        {'name': 'BenefitsAdult', 'output': True},
        {'name': 'BenefitsAdult_tokens', 'output': True,
            'getMembers': 'tokens', 'delimeter': '|', 'selectTopTokens': selectTopTokens},
        {'name': 'GiftedBooks', 'output': True},
        {'name': 'AdultFeelings', 'output': True},
        {'name': 'ChildFeelings', 'output': True},
        {'name': 'DPILComments', 'output': True},
        {'name': 'DPILComments_tokens', 'output': True,
            'getMembers': 'tokens', 'delimeter': '|', 'selectTopTokens': selectTopTokens},
        {'name': 'Gender', 'output': True, 'getMembers': 'members'},
        {'name': 'Ethnicity', 'output': True, 'getMembers': 'members'},
        {'name': 'Location', 'output': True, 'getMembers': 'members'},
        {'name': 'Education', 'output': True, 'getMembers': 'members'},
        {'name': 'Referrer', 'output': True, 'getMembers': 'members'}
    ]

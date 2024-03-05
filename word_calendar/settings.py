import holidays

def get_settings():
    ''' grab settings from config.cfg in program root folder ''' 
    
    settings = {}
    rooms_table = []
    doc_settings = False
    locations = False
    config_items = ["URL", "TIMEZONE", "PAPER_SIZE_HEIGHT", "PAPER_SIZE_WIDTH", "MARGIN_SIZE", "FONT_CHOICE",
                    "CAL_FONT_SIZE", "HEADER_FONT_SIZE", "TABLE_STYLE", "WEEK_STARTS_ON", "UNDERLINE_BOOL", "UNDERLINE_START_TIME", "UNDERLINE_END_TIME"]
    room_items = ["Name", "Color", "Short", "Alt"]
    
    
    with open("config.cfg","r") as config_file:
        for line in config_file:
            # comment lines
            if line[0] == "#":
                pass
            else:
                # determine what settings are being looked at by their header tag
                if "[Locations]" in line:
                    doc_settings = False
                    rooms = True
                if "[Document Settings]" in line:
                    rooms = False
                    doc_settings = True
                if doc_settings == True:
                    # add [Document Settings] items to dictonary
                    for item in config_items:
                        if item in line:
                            config_contents = line.split("=")[1]
                            if config_contents[0] == " ":
                                config_contents = config_contents[1:]
                            settings.update( {item : config_contents.replace("\n","")} )
                if rooms == True:
                    # add [Locations] items to a list
                    if "[Locations]" in line:
                        pass
                    else:
                        if "(" in line:
                            new_room = True
                            room_list = []
                        if ")" in line:
                            rooms_table.append(room_list)
                            new_room = False
                        if new_room == True:
                            for item in room_items:
                                if item in line:
                                    setting_name, room_setting = line.split("=")
                                    if room_setting[0] == " ":
                                        room_setting = room_setting[1:]
                                    
                                    # if the item is a list, turn it into a list and handle each item in list individually to clean strings
                                    if item in ["Alt","Color"]:
                                        room_setting = room_setting.split(",")
                                        for i in range(len(room_setting)):
                                            room_setting[i] = room_setting[i].replace("\n","")
                                            if room_setting[i][0] == " ":
                                                room_setting[i] = room_setting[i][1:]
                                            if item == "Color":
                                                room_setting[i] = int(room_setting[i])
                                    else:
                                        room_setting = room_setting.replace("\n","")    
                                    room_list.append(room_setting)
    return settings, rooms_table

settings, rooms_list = get_settings()

## DOCUMENT SETTINGS HANDLING ##


URL = settings["URL"] # url of ical. google provides a link to this in the specific calendar's settings under "Integrate calendar"


# choice of holidays. see https://pypi.org/project/holidays/ for more info
# TODO: let this be changed in config file
HOLIDAYS = holidays.UnitedStates()

# timezone setting in "America/Los_Angeles" type format
TIMEZONE = settings["TIMEZONE"]

# what day of the week should the calendar start on (usually Sunday or Monday)
WEEK_STARTS_ON = settings["WEEK_STARTS_ON"]

# document settings, fairly self explanitory
PAPER_SIZE_HEIGHT = int(settings["PAPER_SIZE_HEIGHT"])
PAPER_SIZE_WIDTH = int(settings["PAPER_SIZE_WIDTH"])
MARGIN_SIZE = float(settings["MARGIN_SIZE"])
FONT_CHOICE = settings["FONT_CHOICE"]

# size of font inside calendar
CAL_FONT_SIZE = int(settings["CAL_FONT_SIZE"])

# size of month/week header 
HEADER_FONT_SIZE = int(settings["HEADER_FONT_SIZE"])

# word table style. see: https://python-docx.readthedocs.io/en/latest/user/styles-understanding.html
TABLE_STYLE = settings["TABLE_STYLE"]

# if you need events to be underlined after a certain hour, set to true. for no underlines at all, set to false
if "True" in settings["UNDERLINE_BOOL"]:
	UNDERLINE_BOOL = True 
else:
    UNDERLINE_BOOL = False
UNDERLINE_START_TIME = int(settings["UNDERLINE_START_TIME"]) # the time (in 24 hour integer format) after which events are considered underlineable (i.e. evening events)
UNDERLINE_END_TIME = int(settings["UNDERLINE_END_TIME"]) # the time after which to stop underlining events. default is 23:59



## ROOM HANDLING ##

rooms_list.append(["HOLIDAY",[0,0,0],'HOLIDAY',	["Holiday"]])
rooms_list.append(["",[0,0,0],'ERROR',[""]])

NUMBER_OF_ROOMS = len(rooms_list)
ROOMS = []
COLORS= []
CODES = []
BADNAMES = []
MONTH_COLORS = {}
MONTH_CODES = {}
WEEK_COLORS = {}
WEEK_CODES = {}
WEEK_ROOM = {}
CORRECTION = {}

for i in range(NUMBER_OF_ROOMS):
    # generate each room from settings and append it to lists
    # might be easier to make each room an object..?
	ROOMS.append(rooms_list[i][0])
	COLORS.append(rooms_list[i][1])
	CODES.append(rooms_list[i][2])
	BADNAMES.append(rooms_list[i][3])
	MONTH_COLORS.update( {CODES[i] : COLORS[i]} )
	MONTH_CODES.update( {ROOMS[i] : CODES[i]} )
	WEEK_COLORS.update ( {i+1 : COLORS[i]} )
	WEEK_CODES.update ( {i+1 : ROOMS[i]} )
	for y in range(len(BADNAMES[i])):
		CORRECTION.update ( {BADNAMES[i][y] : rooms_list[i][0]} )	


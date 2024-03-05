def parse_date(input_date):
    in_year = int(input_date[0])
    in_month = int(input_date[1])
    in_day = int(input_date[2])
    return in_year, in_month, in_day

def create_document(cal_type, input_date, output_folder):
    cal_type_to_func = {
        "Week": create_week,
        "Month": create_month
    }
    if cal_type in cal_type_to_func:
        return cal_type_to_func[cal_type](input_date, output_folder)
    else:
        raise ValueError(f"Invalid calendar type: {cal_type}")

def create_week(input_date, output_folder):
    TYPE_OF_CAL = "week"
    in_year, in_month, in_day = parse_date(input_date)
    days_in_month = calendar.monthrange(in_year, in_month)[1] 
    csv_weekBegin = f"{in_month}/{in_day}/{in_year}"

    # user number input -> string english output
    days_in_month = calendar.monthrange(in_year, in_month)[1] 

    
    # create range of dates and header/file name
    end_day = in_day + 6
    if determine_day(in_day+6,days_in_month) < in_day:
        # if the week ends in a new month or year, name file and header appropriately
        if in_month + 1 == 13:
            # if the last week containing dates from december goes into january, change date to be january of next year 
            csv_weekEnd = f"{1}/{determine_day(in_day+6,days_in_month)}/{in_year+1}"
            month_year_day = f'Week of {calendar.month_name[in_month]} {in_day} {in_year} to {calendar.month_name[1]} {determine_day(in_day+6,days_in_month)} {in_year+1}'
        else:
            csv_weekEnd = f"{in_month+1}/{determine_day(in_day+6,days_in_month)}/{in_year}"
            month_year_day = f'Week of {calendar.month_name[in_month]} {in_day} to {calendar.month_name[in_month+1]} {determine_day(in_day+6,days_in_month)} {in_year}'
        new_month = True
    
    else:
        # if the week ends in the same month, name file and header appropriately also 
        csv_weekEnd = f"{in_month}/{determine_day(in_day+6,days_in_month)}/{in_year}"
        month_year_day = f'Week of {calendar.month_name[in_month]} {in_day} to {in_day+6} {in_year}'
        new_month = False

    # generate CSV file and get filename
    print("Getting events from .ical file... ", end="")
    csv_filename = csvgen.calgen(csv_weekBegin,csv_weekEnd,TYPE_OF_CAL)
    print("Done!")

    # set up document properties
    print("Setting up document... ", end="")
    document, table = setup_document(month_year_day,TYPE_OF_CAL)
    print("Done!")

    # draw table based on calendar for filling 
    print("Adding Monday through Sunday to header... ", end="")
    add_dates_to_header(table,in_day,days_in_month)
    print("Done!")
    
    # get events in array form
    print("Getting and converting events into array... ", end="")
    week_events = get_events(csv_filename,in_day,days_in_month,TYPE_OF_CAL)
    print("Done!")

    # add events from array
    print("Adding events to table cells:\n")
    add_events_to_week(table,week_events)
    print("Done adding events!")

    # create file name and save file
    document_name = month_year_day + '.docx'
    document_name = os.path.join(output_folder, document_name)

    document.save(document_name)
    
    return document_name, csv_filename

    
def add_dates_to_header(table,first_day,days_in_month):

    ''' creates calendar layout as a word table. ''' 
    
    header_cells = table.rows[0].cells

    days_list = ['Sunday','Monday','Tuesday','Wednesday','Thursday','Friday','Saturday']
    if WEEK_STARTS_ON == "Monday":
        # shift the days list to start on Monday
        days_list = days_list[1:] + days_list[:1]


    for i in range(7):
        n = first_day+i
        number = determine_day(n,days_in_month)
        number_day = f'{days_list[i]} {number}'
        days_list.append(number_day)

    for day in range(7):
        header_cells[day+1].text = days_list[day]
        p = header_cells[day+1].paragraphs[0]
        header_cells[day+1].vertical_alignment = WD_ALIGN_VERTICAL.CENTER
        p_format = p.paragraph_format
        p_format.alignment = WD_ALIGN_PARAGRAPH.CENTER
        cell_grey_bg = parse_xml(r'<w:shd {} w:fill="f2f2f2"/>'.format(nsdecls('w')))
        header_cells[day+1]._tc.get_or_add_tcPr().append(cell_grey_bg)
        
def add_events_to_week(table,week_events):
    # give cells slightly grey background
    cell_grey_bg = parse_xml(r'<w:shd {} w:fill="f2f2f2"/>'.format(nsdecls('w')))
    table.cell(0,0)._tc.get_or_add_tcPr().append(cell_grey_bg)
    for row in range(1,9):
        # determine color for each row
        colors = WEEK_COLORS[row]
        color = RGBColor(colors[0],colors[1],colors[2])
        
        print(f"Adding {ROOMS[row-1]} events...")
        
        for column in range(8):
            # establish cell to use and its alignment
            cell = table.cell(row,column)
            cell.vertical_alignment = WD_ALIGN_VERTICAL.CENTER
            # reset this variable on each cell, as each cell has a pre-existing blank paragraph
            first_event = True
            
            # add room name to first column
            if column == 0:
                cell_grey_bg = parse_xml(r'<w:shd {} w:fill="f2f2f2"/>'.format(nsdecls('w')))
                cell._tc.get_or_add_tcPr().append(cell_grey_bg)
                if row != 0:
                    p = cell.paragraphs[0]
                    p_properties = p.add_run(ROOMS[row-1])
                    p_properties.font.color.rgb = color
                    p_properties.bold =  True
                    p_properties.font.size = Pt(CAL_FONT_SIZE)
                    p.alignment = WD_ALIGN_PARAGRAPH.CENTER
                    
            
            # add events to the following columns 
            else:
                day_events = week_events[column-1]
                sort_list_of_lists(day_events,2)
                for event in range(len(day_events)):
                    
                    if day_events[event][1] == WEEK_CODES[row]:
                        
                        # establish strings
                        event_name, event_start_time, event_end_time = event_strings(day_events[event],WEEK_CODES,"week")

                        # determine if event is an evening event, if option enabled
                        if int(day_events[event][3]) >= UNDERLINE_START_TIME and int(day_events[event][3]) <= UNDERLINE_END_TIME and UNDERLINE_BOOL == True:
                            is_underlined = True
                        else:
                            is_underlined = False                    

                        fontsize = CAL_FONT_SIZE
            
                        # add event times and determine the style
                        if first_event == True:
                            # if this is the first line of the cell, make sure to use pre-existing first paragraph so there is no blank first line in each cell
                            time_string = f"{event_start_time} - {event_end_time}"
                            time = cell.paragraphs[0]
                            time_properties = time.add_run(time_string)
                            time_properties.font.color.rgb = color
                            time_properties.bold =  False
                            time_properties.underline = is_underlined
                            time_properties.font.size = Pt(fontsize)
                            time.alignment = WD_ALIGN_PARAGRAPH.CENTER
                        else:
                            # create new paragraph if this is not the first line
                            time_string = f"{event_start_time} - {event_end_time}"
                            time = cell.add_paragraph()
                            time_properties = time.add_run(time_string)
                            time_properties.font.color.rgb = color
                            time_properties.bold =  False
                            time_properties.underline = is_underlined
                            time_properties.font.size = Pt(fontsize)
                            time.alignment = WD_ALIGN_PARAGRAPH.CENTER
                        
                        # add event name and determine the style
                        event = cell.add_paragraph()
                        event_properties = event.add_run(event_name)
                        event_properties.font.color.rgb = color
                        event_properties.bold =  True
                        event_properties.underline = is_underlined
                        event_properties.font.size = Pt(fontsize)
                        event.alignment = WD_ALIGN_PARAGRAPH.CENTER
                        first_event = False

def create_month(input_date,output_folder):

    TYPE_OF_CAL = "month"

    input_date.pop(2)
    # get user input and turn it into usable data and strings 
    in_month = int(input_date[1])
    in_year = int(input_date[0])
    start_of_month = calendar.weekday(in_year,in_month,1)

    # convert datetime weekday numbers to that of the corresponding column 
    cal_to_table = {6 : 0, 0 : 1, 1 : 2, 2 : 3, 3 : 4, 4 : 5, 5 : 6,}

    # gives the day of the week the month starts on
    start_of_month = cal_to_table[start_of_month]

    # user number input -> string english output
    month_and_year = f'{calendar.month_name[in_month]} {in_year}'
    
    # get number of days in month
    days_in_month = calendar.monthrange(in_year,in_month)[1] + 1

    # generate csv file from input
    csv_monthBegin = f"{in_month}/1/{in_year}"
    csv_monthEnd = f"{in_month}/{days_in_month-1}/{in_year}"
    calendar.monthrange(in_year,in_month)[1]
    print("\nGetting events from .ical... ", end="")
    csv_filename = csvgen.calgen(csv_monthBegin,csv_monthEnd,TYPE_OF_CAL)
    print("Done!")

    # set up document properties
    print("Setting up document... ", end="")
    document, table = setup_document(month_and_year,TYPE_OF_CAL)
    print("Done!")

    # draw table based on calendar for filling 
    print("Adding Sunday through Saturday to header and calendar dates to cells... ", end="")
    add_dates_to_table(start_of_month,table,days_in_month)
    print("Done!")

    # get events in array form
    print("Turning events into array... ", end="")    
    month_events = get_events(csv_filename,1,days_in_month,TYPE_OF_CAL)
    print("Done!")

    # add events from array
    print("Adding events from array into table... ", end="")    
    add_events_to_month(start_of_month,table,days_in_month,month_events)
    print("Done!")
    
    # create file name and save file
    document_name = month_and_year + '.docx'
    document_name = os.path.join(output_folder, document_name)
    document.save(document_name)

    return document_name, csv_filename
 
def add_dates_to_table(start_of_month,table,days_in_month):

    ''' creates calendar layout as a word table. ''' 
    
    def add_date_to_cell(table,row,column,counter_days):
        cell = table.cell(row,column)
        cell_grey_bg = parse_xml(r'<w:shd {} w:fill="f7f7f7"/>'.format(nsdecls('w')))
        cell._tc.get_or_add_tcPr().append(cell_grey_bg)
        p = cell.paragraphs[0]
        p_properties = p.add_run(str(counter_days))
        p_properties.bold =  True
        p_properties.font.size = Pt(CAL_FONT_SIZE)
        p.alignment = WD_ALIGN_PARAGRAPH.RIGHT
        counter_days += 1
        return counter_days

    header_cells = table.rows[0].cells
    days_list = ['Sunday','Monday','Tuesday','Wednesday','Thursday','Friday','Saturday']
    for day in range(7):
        header_cells[day].text = days_list[day]

    # month starts on a sunday
    if start_of_month == 0:
        counter_days = 1
        for row in range(1,14,2):
            for column in range(7):
                if counter_days == days_in_month:
                    break
                counter_days = add_date_to_cell(table,row,column,counter_days)
    
    # month starts on a saturday
    elif start_of_month == 6:
        counter_days = 1
        for row in range(1,16,2):
            if row == 1:
                counter_days = add_date_to_cell(table,row,6,counter_days)
            else:
                for column in range(7):
                    if counter_days == days_in_month:
                        break
                    counter_days = add_date_to_cell(table,row,column,counter_days)


    # month starts on a week day
    elif start_of_month in [1,2,3,4,5]:
        counter_days = 1
        for row in range(1,14,2):
            if counter_days == days_in_month:
                break
            if row == 1:
                for column in range(start_of_month,7):
                    counter_days = add_date_to_cell(table,row,column,counter_days)
            else:
                for column in range(7):
                    if counter_days == days_in_month:
                        break
                    counter_days = add_date_to_cell(table,row,column,counter_days)
        
def add_events_to_month(start_of_month,table,days_in_month,month_events):

    def add_events_to_cell(cell,month_events,counter_days):

        # cell parameters
        cell.vertical_alignment = WD_ALIGN_VERTICAL.CENTER

        # sort day's events by time 
        sort_list_of_lists(month_events[counter_days-1],2)   

        number_of_events = 0

        for event in month_events[counter_days-1]:
            # to be used, eventually, to figure out how to set the font size automatically based on how many events there are in a given cell
            number_of_events += 1
            
            # create paragraph 
            p = cell.paragraphs[0]
            p_format = p.paragraph_format
            p_format.alignment = WD_ALIGN_PARAGRAPH.CENTER
            
            # get strings for assembling into readable event
            
            event_name, event_start_time, event_end_time, event_room= event_strings(event,MONTH_CODES,"month")
            
            # determine if event is an evening event
            if int(event[3]) > 1800:
                is_underlined = True
            else:
                is_underlined = False                    

            # add event to cell
            add_info_text = p.add_run(f'{event_room} {event_start_time}-{event_end_time} ')
            if event == month_events[counter_days-1][-1]:
                add_event_name = p.add_run(f'{event_name}')
            else:
                add_event_name = p.add_run(f'{event_name}\n')                    
                    
            # set font styles
            red = MONTH_COLORS[event_room][0]
            green = MONTH_COLORS[event_room][1]
            blue = MONTH_COLORS[event_room][2]
            font_size = CAL_FONT_SIZE
            color = RGBColor(red,green,blue)
            add_event_name.bold = True
            add_event_name.font.size = Pt(font_size)
            add_info_text.font.size = Pt(font_size)
            add_info_text.underline = is_underlined
            add_event_name.underline = is_underlined
            add_event_name.font.color.rgb = color
            add_info_text.font.color.rgb = color

    counter_days = 1

    # if month starts on sunday
    if start_of_month == 0:
        for row in range(2,15,2):
            for column in range(7):
                if counter_days == days_in_month:
                    break
                cell = table.cell(row,column)
                add_events_to_cell(cell,month_events,counter_days)
                counter_days += 1
    
    # if month starts on saturday
    if start_of_month == 6:
        for row in range(2,17,2):
            if row == 2:
                cell = table.cell(row,6)
                add_events_to_cell(cell,month_events,counter_days)
                counter_days += 1
            else:
                for column in range(7):
                    if counter_days == days_in_month:
                        break
                    cell = table.cell(row,column)
                    add_events_to_cell(cell,month_events,counter_days)
                    counter_days += 1


    # every week day
    elif start_of_month in [1,2,3,4,5]:
        for row in range(2,17,2):
            if row == 2:
                for column in range(start_of_month,7):
                    cell = table.cell(row,column)
                    add_events_to_cell(cell,month_events,counter_days)
                    counter_days += 1
            else:
                for column in range(7):
                    if counter_days == days_in_month:
                        break
                    cell = table.cell(row,column)
                    add_events_to_cell(cell,month_events,counter_days)
                    counter_days += 1

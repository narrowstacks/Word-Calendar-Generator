import pandas as pd
import icalendar
import os
import recurring_ical_events
import urllib.request
from datetime import datetime
import configHandler


# get the calendar from file or URL
def get_calendar_file(url, file):
    ''' get calendar from file or URL
    
    Args:
        url (str): The URL to the calendar.
        file (str): The file path to the calendar.
        
        Returns:
            icalendar.Calendar: The calendar object.
    '''
    if file:
        ical_string = open(file, 'rb').read()
    else:
        ical_string = urllib.request.urlopen(url).read()
    return icalendar.Calendar.from_ical(ical_string)

# return dataframe with events from calendar from one year before today to one year after today. 
# retain information about each event like description and location. handle reoccuring events and exceptions to reoccuring events.
def calendar_to_df(importedCalendar):
    '''get events from calendar'''
    # get events from calendar
    events = recurring_ical_events.of(importedCalendar).between(datetime.now() - pd.DateOffset(years=1), datetime.now() + pd.DateOffset(years=1))
    # create dataframe with events
    df = pd.DataFrame(columns=['start', 'end', 'summary', 'description', 'location'])
    for event in events:
        if event.get('DTSTART'):
            start = event.get('DTSTART').dt
        else:
            start = None
        if event.get('DTEND'):
            end = event.get('DTEND').dt
        else:
            end = None
        if event.get('SUMMARY'):
            summary = event.get('SUMMARY')
        else:
            summary = None
        if event.get('DESCRIPTION'):
            description = event.get('DESCRIPTION')
        else:
            description = None
        if event.get('LOCATION'):
            location = event.get('LOCATION')
        else:
            location = None
        df = df.append({'start': start, 'end': end, 'summary': summary, 'description': description, 'location': location}, ignore_index=True)
    # sort by start time and reset index
    df = df.sort_values(by='start')
    df = df.reset_index(drop=True)
    return df

# get events from specific day in dataframe
def get_days_events_df(df, day):
    '''Get events from specific day and organize by start time.
    
    Args:
        df (pandas.DataFrame): The dataframe containing the events.
        day (datetime.date): The specific day to filter events from.
    
    Returns:
        pandas.DataFrame: The events from the specific day, sorted by start time.
    '''
    # filter events from specific day
    events_day = df[(df['start'].dt.date == day)]
    # sort events by start time
    events_day = events_day.sort_values(by='start')
    return events_day





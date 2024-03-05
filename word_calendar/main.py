import os, sys, subprocess
import cal_gen, argparse, datetime
from settings import WEEK_STARTS_ON
from gooey import Gooey, GooeyParser

@Gooey(program_name='iCal to Word Calendar Generator')
def main():
    # ask user what kind of calendar they would like to make and execute chosen program
    parser = GooeyParser(description="Turns .ics files from most calendar providers into Word .docx documents.")
    parser.add_argument('Choice', choices=['Month', 'Week'],help="Choose which type of calendar to generate",gooey_options={'label' : "Monthly or Weekly?"},default="Month")
    parser.add_argument('Date', widget="DateChooser",help=f"Choose a {WEEK_STARTS_ON}. If month is selected, only the month and year are used. ",default=str(datetime.datetime.today())[:-16])
    parser.add_argument('Output', help="Where to save the output .docx", widget='DirChooser',default=os.path.dirname(os.path.realpath(__file__))) 
    parser.add_argument('Open',help="Open file when done?",choices=['Yes', 'No'],default='Yes')
    args = parser.parse_args()
    
    input_date = args.Date.split("-")
    input_type = args.Choice
    output_folder = args.Output

    document_name, csv_filename = cal_gen.create_document(input_type,input_date,output_folder)

    os.remove(csv_filename)

    if args.Open == "Yes":
        print("\nOpening file...")
        open_file(document_name)
    else:
        print("Bye!")
    
def open_file(filename):
    ''' system-agnostic function for opening a file in the program associated with its file extension. takes a filename as input'''
    
    if sys.platform == "win32":
        os.startfile(filename)
    else:
        opener ="open" if sys.platform == "darwin" else "xdg-open"
        subprocess.call([opener, filename])

if __name__ == "__main__":
    main()
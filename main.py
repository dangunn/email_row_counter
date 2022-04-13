from win32com.client import constants
from win32com.client.gencache import EnsureDispatch as Dispatch
from datetime import datetime
from bs4 import BeautifulSoup
import PySimpleGUI as psg

# input from the user the settings
subject_filter = psg.popup_get_text("Enter the subject search string",default_text=" online class on ")
start_date = psg.popup_get_date(1,1,title="First day",no_titlebar=False)
start_date = '{}-{}-{} 00:01 AM'.format(start_date[2],start_date[0],start_date[1])
FILE_TYPES_CSV = (("CSV Files", "*.csv *.csv"),)
output_file = psg.popup_get_file("Save as",default_extension="csv",default_path="rows_per_day.csv",file_types=FILE_TYPES_CSV)
chart_title = psg.popup_get_text("Enter the chart title",default_text="Student Absences")

# Thank you for advice on using the win32com API from:
# https://www.codeforests.com/2021/05/16/python-reading-email-from-outlook-2/?msclkid=059450b5bacf11ec9d31f1c4756891a2
# https://stackoverflow.com/questions/5077625/reading-e-mails-from-outlook-with-python-through-mapi?msclkid=e6655b5abac811ec974fbbcfc52dfe3d

outlook = Dispatch("Outlook.Application")
mapi = outlook.GetNamespace("MAPI")

class Oli():
    def __init__(self, outlook_object):
        self._obj = outlook_object

    def items(self):
        ''' yields a tuple(index, Item) '''
        array_size = self._obj.Count
        for item_index in range(1,array_size+1):
            yield (item_index, self._obj[item_index])

    def prop(self):
        ''' helping to introspect outlook object exposing available properties (methods and attributes) '''
        return sorted( self._obj._prop_map_get_.keys() )

rows_per_day = {}

for inx, folder in Oli(mapi.Folders).items():
    for inx,subfolder in Oli(folder.Folders).items():
        if subfolder.Name == "Inbox":
            messages = subfolder.Items
            messages = messages.Restrict("[ReceivedTime] >= '" + start_date + "'")
            messages = messages.Restrict("@SQL=(urn:schemas:httpmail:subject LIKE '%{}%')".format(subject_filter))
            for message in list(messages)[:30]:
                #print(message.Subject, message.ReceivedTime, message.SenderEmailAddress)
                soup = BeautifulSoup(message.HTMLBody, "html.parser")
                row_count = len(soup.find_all('tr'))
                day = message.ReceivedTime.strftime("%Y-%m-%d")
                if day in rows_per_day:
                    rows_per_day[day] = max(rows_per_day[day], row_count)
                else:
                    rows_per_day[day] = row_count

with open(output_file, 'w') as f:
    f.write('Date,%s\n' % chart_title)
    for key in rows_per_day.keys():
        f.write("%s, %s\n" % (key, rows_per_day[key]))

print(output_file, "written with", len(rows_per_day), "records")
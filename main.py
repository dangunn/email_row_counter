from win32com.client import constants
from win32com.client.gencache import EnsureDispatch as Dispatch
from datetime import datetime
from bs4 import BeautifulSoup

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

data = {}

for inx, folder in Oli(mapi.Folders).items():
    # iterate all Outlook folders (top level)
    print("-"*70)
    print(folder.Name)

    for inx,subfolder in Oli(folder.Folders).items():
        if subfolder.Name == "Inbox":
            print("(%i)" % inx, subfolder.Name,"=> ", subfolder)
            today = datetime.now()

            # first day of the month
            start_time = today.replace(month=1, hour=0, minute=0, second=0).strftime('%Y-%m-%d %H:%M %p')
            messages = subfolder.Items
            messages = messages.Restrict("[ReceivedTime] >= '" + start_time + "'")
            messages = messages.Restrict("@SQL=(urn:schemas:httpmail:subject LIKE '%online class on%')")
            for message in list(messages)[:30]:
                #print(message.Subject, message.ReceivedTime, message.SenderEmailAddress)
                soup = BeautifulSoup(message.HTMLBody, "html.parser")
                student_count = len(soup.find_all('tr'))
                day = message.ReceivedTime.strftime("%Y-%m-%d")
                if day in data:
                    data[day] = max(data[day], student_count)
                else:
                    data[day] = student_count

with open('output.csv', 'w') as f:
    f.write('Date,Student Absences\n')
    for key in data.keys():
        f.write("%s, %s\n" % (key, data[key]))
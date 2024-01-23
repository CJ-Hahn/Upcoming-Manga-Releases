import json
import requests
import win32com.client
from datetime import datetime

outlook = win32com.client.Dispatch("Outlook.Application")

def upcomingReleases(Title):
    Title = Title + ", Vol."
    Response = requests.get(f"https://www.googleapis.com/books/v1/volumes?q={Title}&maxResults=40").json()

# Converts reponse from API to a list so we can remove the 'kind' and 'totalItems' data; 
# Search result data is stored in tuple 'Volumes'

    temp = list(Response.items())       # Converts Tuple into a List for easier traversal

    Volumes = temp[2][1]                # Copies List of 'items' data from temp variable

    SearchData = {}                     # Initialize dictionary for Title:PublishedDate
    SortedResults = {}                  # Initialize Dictionary for unreleased Title:PublishedDate

    i = 0

    # Populates SearchData
    for entry in Volumes:
        SearchData[Volumes[i].get('volumeInfo').get('title')] = Volumes[i].get('volumeInfo').get('publishedDate')
        i=i+1
    
    
    # traverses dictionary of titles and published dates and only prints entries with a valid PublishedDate that is 
    # sometime in the future and that contain the Title
    for key in SearchData.keys():
    
        # Confirms that publishedDate has data as well as a full length date
        if SearchData[key] != None and len(SearchData[key]) == 10:
            pDate = datetime.strptime(SearchData[key], '%Y-%m-%d').date()   
    
        # Only prints out titles that haven't come out yet
        if SearchData[key] != None and Title in str(key) and datetime.now().date() < pDate:
            #print(f"Title: {key} Published Date: {SearchData[key]}")
            SortedResults[key] = pDate
            #pass
        
    return SortedResults

def sendReleaseDate(Title, ReleaseDate):    
  appt = outlook.CreateItem(1) # AppointmentItem
  #appt.Start = "2024-01-19 10:10" # yyyy-MM-dd hh:mm
  appt.Start = f"{ReleaseDate}" # yyyy-MM-dd hh:mm
  appt.Subject = f"Release of {Title}!"
  appt.Duration = 60 # In minutes (60 Minutes)
  #appt.Location = "Location Name"
  #appt.MeetingStatus = 1 # 1 - olMeeting; Changing the appointment to meeting. Only after changing the meeting status recipients can be added
  
  appt.Recipients.Add("christopherjhahn2@gmail.com") # Don't end ; as delimiter

  appt.Save()
  appt.Send()


x = input("What upcoming releases are you looking for?\n")
y = upcomingReleases(x)
print(y)
#sendReleaseDate("Undead Unluck, Vol. 15","2024-03-16")

'''
1/18/24 - Created file
                - Created upcomingReleases(Title) function which returns a Dictionary of unreleased volumes
                that returns a Dictionary of unreleased volumes.
                - Created function called sendReleaseDate(Title, ReleaseDate) that emails an appointment for the release date of the upcoming 
                    releases.
'''
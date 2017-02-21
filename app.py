#!/usr/bin/python

"""
Delta College R25 Excel Tool

TODO:
Don't combine reservations if there is a class in between them

----
Requirements:
- Python 2.7
"""

from openpyxl import Workbook
from openpyxl import load_workbook
from datetime import time
from datetime import *
from collections import defaultdict
import re


def format_time(t):
  return t.strftime("%I:%M %p")

def time_to_datetime(t):
  current = date.today()
  return datetime.combine(current, t)

def unicode_to_time(u):
  new_time = datetime.strptime(u.encode('ASCII', 'ignore'), '%I:%M %p')
  return time(new_time.hour, new_time.minute, new_time.second)

class Event:
  def __init__(self):
    self.space = None
    self.resource = None
    self.start = None
    self.end = None
  
  def set_start(self, t):
    #Set start time
    if type(t) is time:
      self.start = t
    else:
      self.start = unicode_to_time(t)
  
  def time_difference(self, event):
    # returns the smallest difference in time between self and event
    time_list = list()
    for t1 in (self.start, self.end):
      for t2 in (event.start, event.end):
        time_list.append(abs(time_to_datetime(t1) - time_to_datetime(t2)))
    return sorted(time_list)[0]
  
  def before(self, event):
    # returns a boolean depending on if self is before event
    if self.start < event.start:
      return True
    else:
      return False
  
  def set_end(self, t):
    # Set end time
    if type(t) is time:
      self.end = t
    else:
      self.end = unicode_to_time(t)
  
  def __repr__(self):
    return "%s - %s, %s, %s" % (format_time(self.start), format_time(self.end), self.space, self.resource)


def combine_reservations(rooms):
  # Input a rooms dict, combines the inputted dict
  for room in rooms:
    combined = True
    while(combined):
      combined = False
      for e1 in rooms[room]:
        for e2 in rooms[room]:
          if e1 != e2 and not combined:
            #Don't compare the event to itself
            if e1.time_difference(e2) < timedelta(hours=2) and e1.resource == e2.resource:
              combined = True
              if e1.before(e2):
                e1.end = e2.end
              else:
                e1.start = e2.start
              rooms[room].remove(e2)


def get_room_events(sheet):
  rooms = defaultdict(list)
  for row in tuple(sheet.rows):
    if row[0].value != None:
      if type(row[0].value) is time:
        if row[5].value != None:
          event = Event()
          event.set_start(row[0].value) # Start datetime.time
          event.set_end(row[1].value) # End datetime.time
          event.space = row[5].value.encode('ASCII', 'ignore') # Space
          if row[6].value != None:
            event.resource = row[6].value.encode('ASCII', 'ignore') # Resource
          rooms[event.space].append(event)
  return rooms


def get_reservations(rooms):
  reservations = list()
  for room in rooms:
    for event in rooms[room]:
      if event.resource != None:
        reservations.append(event)
  return reservations


#e1 = Event()
#e1.start = unicode_to_time(u"12:00 PM")
#e1.end = unicode_to_time(u"1:55 PM")
#print e1
#e2 = Event()
#e2.start = unicode_to_time(u"2:00 PM")
#e2.end = unicode_to_time(u"3:55 PM")
#print e2
##print e1.time_difference(e2) < timedelta(hours=2)
#print e2.before(e1)


wb = load_workbook('today.xlsx')
sheet = wb.active


rooms = get_room_events(sheet)


print len(get_reservations(rooms))
combine_reservations(rooms)

reservations = get_reservations(rooms)
print len(reservations)


reservations = sorted(reservations, key=lambda e: e.start)
result = ""
for event in reservations:
  result += "%s\n" % (event)

with open("result.txt", "w+") as f:
  f.write(result)



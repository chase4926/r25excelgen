#!/usr/bin/python

"""
Delta College R25 Excel Tool

TODO:
Don't combine reservations if there is a class in between them
Allow for multiple resource requests!

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
  
  def set_end(self, t):
    # Set end time
    if type(t) is time:
      self.end = t
    else:
      self.end = unicode_to_time(t)
  
  def __lt__(a, b):
    if a.start < b.start:
      return True
    elif a.start == b.start:
      if a.end <= b.end:
        return True
      else:
        return False
    else:
      return False
  
  def __gt__(a, b):
    if a.start > b.start:
      return True
    elif a.start == b.start:
      if a.end > b.end:
        return True
      else:
        return False
    else:
      return False
  
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
              if e1 < e2: # Checks if e1 takes place before e2
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


def format_space(space):
  result = space.split('_', 1)[1]
  return result.split(' ', 1)[0]


def get_resource_column(resource):
  return {
    'Laptop wifi': 'D',
    'Computer / Dell Laptop': 'D',
    'Laptop Wireless Cart #1 (20)': 'E',
    'Laptop Wireless Cart #2 (20)': 'E',
    'Laptop Wireless Cart #3 (20)': 'F',
    'Clickers 25': 'E',
  }.get(resource, 'H') # Laptop


wb = load_workbook('reservations.xlsx')
sheet = wb.active


rooms = get_room_events(sheet)

print("There are %i total reservations." % len(get_reservations(rooms)))
combine_reservations(rooms)

reservations = get_reservations(rooms)
print("After being combined, there are %i reservations." % len(reservations))


reservations = sorted(reservations)


template_book = load_workbook('template.xlsx')
template = template_book.active


i = 2 # Row to start on
for event in reservations:
  # Write event details to row
  template["B%i" % i] = format_space(event.space)
  template["K%i" % i] = format_time(event.start)
  template["M%i" % i] = format_time(event.end)
  resource_col = get_resource_column(event.resource)
  if resource_col == 'H':
    template["H%i" % i] = event.resource
  else:
    template["%s%i" % (resource_col, i)] = 'X'
  # Move to the next row
  i += 1
template_book.save('result.xlsx')




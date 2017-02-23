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
from openpyxl.styles import Color, PatternFill
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
  
  def add_resource(self, resource):
    if self.resource == None:
      self.resource = [resource]
    else:
      if not resource in self.resource:
        self.resource.append(resource)
  
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


def combine_reservations(r):
  # Input a rooms dict, outputs a new dict with combined events
  rooms = copy_rooms(r)
  for room in rooms:
    combined = True
    while(combined):
      combined = False
      for e1 in rooms[room]:
        for e2 in rooms[room]:
          if e1 != e2 and not combined:
            #Don't compare the event to itself
            if (e1.resource != None) and (e1.resource == e2.resource) and (e1.time_difference(e2) < timedelta(hours=2)):
              combined = True
              if e1 < e2: # Checks if e1 takes place before e2
                e1.end = e2.end
              else:
                e1.start = e2.start
              rooms[room].remove(e2)
  return rooms


def get_room_events(sheet):
  rooms = defaultdict(list)
  last_event = None
  for row in tuple(sheet.rows):
    if row[0].value != None:
      if type(row[0].value) is time:
        if row[5].value != None:
          event = Event()
          event.set_start(row[0].value) # Start datetime.time
          event.set_end(row[1].value) # End datetime.time
          event.space = row[5].value.encode('ASCII', 'ignore') # Space
          if row[6].value != None:
            event.add_resource(row[6].value.encode('ASCII', 'ignore')) # Resource
          rooms[event.space].append(event)
          # Set last_event in case there are additional resources
          last_event = event
    else:
      # Check if there was an event prior to this
      if last_event != None:
        if row[6].value != None:
          #Add resource to last event
          last_event.add_resource(row[6].value.encode('ASCII', 'ignore'))
        else:
          # Trailed off the end of the list
          last_event = None
  return rooms


def get_reservations(rooms):
  reservations = list()
  for room in rooms:
    for event in rooms[room]:
      if event.resource != None:
        reservations.append(event)
  return reservations


def get_delivery_time(rooms, event):
  events = rooms[event.space]
  events = events[0:events.index(event)]
  if len(events) > 0:
    end_event = events[-1]
    end_time = end_event.end
    if event.time_difference(end_event) < timedelta(minutes=15):
      return "@%s" % format_time(end_time)
    else:
      return "%s-%s" % (format_time(end_time), format_time(event.start))
  else:
    return 'OPEN'


def get_pickup_time(rooms, event):
  events = rooms[event.space]
  after_events = list()
  for e in events:
    if event.end < e.start:
      after_events.append(e)
  if len(after_events) > 0:
    start_event = after_events[0]
    start_time = start_event.start
    if event.time_difference(start_event) < timedelta(minutes=15):
      return "@%s" % format_time(start_time)
    else:
      return "%s-%s" % (format_time(event.end), format_time(start_time))
  else:
    return 'OPEN'


def format_space(space):
  result = space.split('_', 1)[1]
  return result.split(' ', 1)[0]


def get_resource_column(resource):
  return {
    'Laptop wifi': 'D',
    'Computer / Dell Laptop': 'D',
    'Laptop Wireless Cart #1 (20)': 'E',
    'Laptop Wireless Cart #2 (20)': 'E',
    'Laptop Wireless Cart #3 (20)': 'E',
    'Clickers 25': 'F',
    'Clickers 52': 'F',
    'Wireless Presenter': 'G',
  }.get(resource, 'H') # Laptop


def count_events(rooms):
  i = 0
  for room in rooms:
    for event in rooms[room]:
      i += 1
  return i


def copy_rooms(rooms):
  new_rooms = defaultdict(list)
  for room in rooms:
    new_rooms[room] = rooms[room][:]
  return new_rooms


wb = load_workbook('reservations.xlsx')
sheet = wb.active


rooms = get_room_events(sheet)
print("There are %i total reservations." % len(get_reservations(rooms)))
reservations = get_reservations(combine_reservations(rooms))
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
  for resource in event.resource:
    resource_col = get_resource_column(resource)
    if resource_col == 'H':
      template["H%i" % i] = resource
    else:
      template["%s%i" % (resource_col, i)] = 'X'
  template["J%i" % i] = get_delivery_time(rooms, event)
  template["N%i" % i] = get_pickup_time(rooms, event)
  # Move to the next row
  i += 1

red_color = PatternFill(start_color='FFFF9999', end_color='FFFF9999', fill_type='solid')
for i in range(len(reservations)):
  if i % 2 == 0:
    for n in range(15):
      template[("%s%i") % (chr(n+65), i + 2)].fill = red_color

template_book.save("%s.xlsx" % (datetime.now() + timedelta(days=1)).strftime("%b-%d"))




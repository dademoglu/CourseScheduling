# CourseScheduling
University Scheduling Algorithm
#By Deniz Ademoglu
2019
What it does:
   There are two pages in data sheet where one has courses one has rooms, one workbook has only constraints. Before anything else it creates class element
   for every item that is read from workbooks/sheets. Algorithm works like this:
       take a course from the list
       find it's constraints and check the other courses those are in that constraint group's schedule to find the hours that are already filled
       take all the empty hours to place the course
       compare every hour with rooms available times
       place the course when first empty match is found
       if couldn't find move on to the next room until you find a room with an appropiate empty slot

This program doesn't guarantee eficiency in any way(memory, run time).
This program doesn't guarantee that it will find empty places for all course
This program doesn't guarantee to find the most suitiable room-course match

#!/usr/bin/env python
# -*- coding: utf-8 -*-
"""
ABOUT THIS SCRIPT:

Import csv data to autogenerate the Kalender elements

1. Use the default luhze template and select a whole page

2. Execute this script:

You will be prompted first to choose between generating a specific element or all elements and then for the csv file

4. The data from the csv file will be imported and a table of
textboxes will be drawn on the page.

############################

LICENSE:

This program is free software; you can redistribute it and/or modify
it under the terms of the GNU General Public License as published by
the Free Software Foundation; either version 2 of the License, or
(at your option) any later version.

This program is distributed in the hope that it will be useful,
but WITHOUT ANY WARRANTY; without even the implied warranty of
MERCHANTABILITY or FITNESS FOR A PARTICULAR PURPOSE.  See the
GNU General Public License for more details.

You should have received a copy of the GNU General Public License
along with this program; if not, write to the Free Software
Foundation, Inc., 51 Franklin Street, Fifth Floor, Boston, MA 02110-1301, USA.

Author: Niclas Stoffregen
"""

import sys
import datetime
import time
from Tkinter import *
import csv


try:
    import scribus
except ImportError,err:
    print "This Python script is written for the Scribus scripting interface."
    print "It can only be run from within Scribus."
    sys.exit(1)



#get the cvs data
def getCSVdata():
    #opens a csv file, reads it in and returns a 2 dimensional list with the data


    csvfile = scribus.fileDialog("csv2table :: open file", "*.csv")
    if csvfile != "":
        try:
            reader = csv.reader(file(csvfile),delimiter='\t') 
            datalist=[]
            for row in reader:
                rowlist=[]
                for col in row:
                    rowlist.append(col)
                datalist.append(rowlist)
            return datalist
        except Exception,  e:
            scribus.messageBox("csv2table", "Could not open file %s"%e)
    else:
        sys.exit


def dateFun(date):
  string_date = date
  format = "%d.%m.%Y"
  try:
      res = datetime.datetime.strptime(string_date, format)
  except TypeError:
      res = datetime.datetime(*(time.strptime(string_date, format)[0:6]))
  return res # testcompete print alternation



def generateContent(string,window):

    #generates all the elements/only one element
    start = 0 
    end = 0
    if string == "all":
        #dont include the first line to make the data sort work [1:]
        #sort the the elements by date
        data = sorted(getCSVdata()[1:], key = lambda row: dateFun(row[5]), reverse=True)
        end = len(data)
        print(str(end) + "  "  + str(start))
        print("generating elements from all lines")
        window.destroy()
    elif RepresentsInt(string):
        start = int(string)-1 #index shifting
        end = int(string)
        print("only generating line "  + string)
        data=getCSVdata()
        window.destroy()
    else:
        print(string + " is not a valid value")
        print("exiting")
        window.destroy()
        sys.exit()


    if not scribus.haveDoc() > 0: #do we have a doc?
        scribus.messageBox("importcvs2table", "No opened document.\nPlease open one first.")
        sys.exit()
    userdim = scribus.getUnit() #get unit and change it to mm
    scribus.setUnit(scribus.UNIT_MILLIMETERS)
    
    lineWidth = 3.92 #aka 1.383mm, kp warum
        
    date=""
    for i in range(start, end):

        print("add element: "  + data[i][0])

        #random values
        hpos=120.0
        vpos=200.0
        hposition=120.0
        vposition=200.0
                
        objectlist=[] #list for all boxes


        x=0 #sets the progress
        #create the blue box
        print("create the blue line")
        blueBox = scribus.createLine(hposition +1,vposition,hposition+1,vposition + 5.863)
        scribus.setLineColor("Cyan",blueBox)
        scribus.setLineWidth(lineWidth,blueBox)
        objectlist.append(blueBox)
        scribus.progressSet(x)
        
    
        x=1
        #create the data character box
        #these are the width values for the numbers
        zero = 4.608
        one = 2.839
        two = 4.724
        three = 4.393
        four = 4.625
        five = 4.261
        six = 4.278
        seven = 4.261
        eight = 4.625
        nine = 4.708

        lenArray = [zero,one,two,three,four,five,six,seven,eight,nine]

        date=data[i][5]
        marginleft=1.3
        margintop=0.519 #substract, cause the box is heigher that the blue line
        cellwidthright = 10.951
        cellHeight = 8.282
        hposition = hposition + marginleft + 1
        textbox = scribus.createText(hposition,vposition - margintop,cellwidthright, cellHeight)
        scribus.setFont("Quicksand Regular", textbox)
        scribus.setFontSize(20.0,textbox)
        finalDate = ""
        dateLength = 0
        #checks if the date is from 01-09, in that case remove the zero
        if date[0] == '0':
            finalDate = date[1]
            dateLength = lenArray[int(date[1])]
            
        else:
            finalDate = date[:2]
            dateLength = lenArray[int(date[0])] + lenArray[int(date[1])]
            

        scribus.insertText(finalDate,0,textbox)
        print("day: " + finalDate)
        objectlist.append(textbox)
        scribus.progressSet(x)


        x=2
        #create the month/day box
        print("create the box with the day and month")
        width=19.447
        height=8.025
        marginleft = dateLength #gain that from the calculations above, depends on the width of the date characters


        monthBox = scribus.createText(hposition + marginleft + 0.7, vposition,width,height)
        scribus.setFont("Quicksand Regular", monthBox)
        scribus.setFontSize(8.5,monthBox)

        month=""
        m=date[3:5]
        if m == '01':
            month="Januar"
        elif m == '02':
            month="Februar"
        elif m == '03':
            month="März"
        elif m == '04':
            month="April"
        elif m == '05':
            month="Mai"
        elif m == '06':
            month="Juni"
        elif m == '07':
            month="Juli"
        elif m == '08':
            month="August"
        elif m == '09':
            month="September"
        elif m == '10':
            month="Oktober"
        elif m == '11':
            month="November"
        elif m == '12':
            month="Dezember"
        else:
            print("cant determine month!")


        day = datetime.date(int(date[6:]),int(m),int(date[:2])).weekday()
        dayName = ""

        if day==0:
            dayName="Montag"
        elif day==1:
            dayName="Dienstag"
        elif day==2:
            dayName="Mittwoch"
        elif day==3:
            dayName="Donnerstag"
        elif day==4:
            dayName="Freitag"
        elif day==5:
            dayName="Samstag"
        elif day==6:
            dayName="Sonntag"
        else:
            print("cant determine day!")


        text = month + "\n" + dayName;
        scribus.setStyle("Kalender_neu_Monat und Tag", monthBox)
        scribus.insertText(text,0,monthBox)
        print("month: "  + month  + " day: " + dayName)
        objectlist.append(monthBox)
        scribus.progressSet(x)

        x=3
        #create the main text box
        print("create the main text box")
        margintop = 5.5
        hpos = hpos - 0.383 #i dont know why but scribus always places the element 0.383 right than it should be :/
        mainTextBox = scribus.createText(hpos, vposition + margintop, 43.0,45.0) #minus eins weil der blaue balken seinen kasten overflowed

        #insert category
        print("insert the category: "  + data[i][1])
        scribus.insertText(data[i][1],0,mainTextBox)
        endCategory = scribus.getTextLength(mainTextBox)
        scribus.selectText(0,endCategory,mainTextBox)
        scribus.setFontSize(10.5,mainTextBox)
        scribus.selectText(0,endCategory,mainTextBox)
        scribus.setStyle("Kalender_Eventname",mainTextBox)
        
        #insert main text
        print("insert the main text")
        scribus.insertText("\n" + data[i][2],endCategory,mainTextBox)
        endMainText = scribus.getTextLength(mainTextBox)-endCategory
        scribus.selectText(endCategory, endMainText, mainTextBox)
        scribus.setStyle("Kalender_Eventbeschreibung", mainTextBox)  

        #get start length to color everything black and set font size
        startAll = scribus.getTextLength(mainTextBox)
        createPlaceTimePrice(mainTextBox,"\n| Ort: ", "", "Kalender_Eventname")
       
        #insert value for place
        createPlaceTimePrice(mainTextBox, data[i][4], "Heuristica Regular","")
       
        #insert time letters
        createPlaceTimePrice(mainTextBox," | Zeit: ", "Quicksand Regular","")
       
        #insert time value
        createPlaceTimePrice(mainTextBox,data[i][6], "Heuristica Regular","")

        #insert price letters
        createPlaceTimePrice(mainTextBox," | Eintritt: ", "Quicksand Regular","")
        
        #insert price value
        createPlaceTimePrice(mainTextBox,data[i][3], "Heuristica Regular","")
        

        #setFontSize and black color for the whole detail box
        endAll = scribus.getTextLength(mainTextBox)-startAll
        scribus.selectText(startAll,endAll,mainTextBox)
        scribus.setFontSize(8.5,mainTextBox)
        scribus.selectText(startAll,endAll,mainTextBox)
        scribus.setTextColor("Black",mainTextBox)
        

        objectlist.append(mainTextBox)
        scribus.progressSet(x)


        #do some generell stuff
        scribus.groupObjects(objectlist)
        scribus.progressReset()
        scribus.setUnit(userdim) # reset unit to previous value
        scribus.docChanged(True)
        scribus.statusMessage("Done")
        scribus.setRedraw(True)    

    print("done")
    return 0

def createPlaceTimePrice(box, string, font, style):

    startLength = scribus.getTextLength(box)
    scribus.insertText(string,startLength,box)
    endLength = scribus.getTextLength(box)-startLength

    if style != "":
        scribus.selectText(startLength,endLength,box)
        scribus.setStyle(style,box)
    if font != "":
        scribus.selectText(startLength,endLength,box)
        scribus.setFont(font, box)
    return endLength
    

 
def RepresentsInt(s):
    #checks if a given string in an integer
    try: 
        int(s)
        return True
    except ValueError:
        return False
    
def main(argv):

    def callback():
        string = userInput.get()
        generateContent(string,window)
        return 0

    def on_closing():
        window.destroy()
        sys.exit()
    
    window = Tk()
    var = StringVar()
    label = Label(window, textvariable=var, relief=RAISED )
    var.set('choose line or type in "all"')
    label.pack()
    userInput=Entry()
    userInput.pack()
    userInput.focus_set()
    b=Button(window,text="submit", command=callback)
    b.pack()

    window.protocol("WM_DELETE_WINDOW", on_closing)
    window.mainloop()
    

def main_wrapper(argv):
    try:
        scribus.statusMessage("Importing .csv table...")
        scribus.progressReset()
        main(argv)
    finally:
        #exit
        if scribus.haveDoc() > 0:
            scribus.setRedraw(True)
        scribus.statusMessage("")
        scribus.progressReset()

#only runs if being run as a script
if __name__ == '__main__':
    main_wrapper(sys.argv)

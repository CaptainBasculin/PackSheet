#!/usr/bin/env python
# -*- coding: utf-8 -*-

import os
import py7zr
import time
import openpyxl 

#I want output to be this: 
# =HYPERLINK("https://osu.ppy.sh/beatmapsets/431708",Image("https://b.ppy.sh/thumb/431708l.jpg", 2))
out = os.getcwd() + "/output"
source = os.getcwd() + "/putPackHere" 
tmp = os.getcwd() + "/tmp" 

state = True

while(state):
	print("Program's searching for an unparsed beatmap pack, hit CTRL+C to close")
	time.sleep(5)
	conts = os.listdir(source)
	counter = 0
	for i in conts:
		if(i.startswith("Beatmap Pack")):
			counter = counter + 1
			#i is our pack, Beatmap Pack #1120.7z
			packName = i[i.find("#"):i.find(".")] 
			with py7zr.SevenZipFile(source+"/"+ i, mode='r') as z:
				z.extractall(path = os.getcwd() + "/tmp")
			print("Extracted to /tmp")
			tmpConts = os.listdir(tmp)
			contentList = [packName]
			for ii in tmpConts:
				#397917 KINEMA106 - Fly Away.osz
				#before space is beatmapset id (https://osu.ppy.sh/beatmapsets/397917)
				## =HYPERLINK("https://osu.ppy.sh/beatmapsets/431708",Image("https://b.ppy.sh/thumb/431708l.jpg"))
				setID = ii[0:ii.find(' ')]
				print("Added background for set +" + setID)
				contentList.append('=HYPERLINK("https://osu.ppy.sh/beatmapsets/' + str(setID) + '",Image("https://b.ppy.sh/thumb/'+ str(setID) + 'l.jpg"))')
				#print("contentList is now " + contentList)
				os.remove(tmp+"/" + ii)
			#all data's saved to contentlist. Add that to xlsx file
			wb = openpyxl.load_workbook(out + "/output.xlsx")
			sheet = wb.active
			sheet.append(contentList)
			wb.save(out + "/output.xlsx")
			os.rename(source +"/"+ i, source + "/Parsed " +i )
			print(str(counter) + " Beatmap Pack(s) in total are imported.")	

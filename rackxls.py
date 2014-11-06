#!/usr/bin/env python
#version 1.1

from xlwt import Workbook, easyxf
from sys import exit
import time
import datetime
import ConfigParser

xls_Config = ConfigParser.ConfigParser()
xls_Config.read("rackxls.conf")
xls_Config.sections()

Style_nowrap = easyxf("font: bold on, height 300; align: wrap off, horiz centre, vert centre;")
Power_txt = easyxf("font: height 210, bold on; align: horiz centre, vert centre;")
Power_decimal = easyxf("font: height 210, bold on; align: wrap off, horiz centre, vert centre;", "0.0")
Style_basic_text = easyxf("font: height 210, bold on;")

#Helper method to get the value from the various sections in conf file
def ConfigSectionMap(section):
    dict1 = {}
    options = xls_Config.options(section)
    for option in options:
        try:
            dict1[option] = xls_Config.get(section, option)
            if dict1[option] == -1:
                DebugPrint("skip: %s" % option)
        except:
            print("exception on %s!" % option)
            dict1[option] = None
    return dict1



localtime = time.asctime(time.localtime(time.time()))
now = datetime.datetime.now()

day = int(now.day)
month = int(now.month)
year = int(now.year)
hour = int(now.hour)
minutes = int(now.minute)

# Variable Initialisation
output_path = ConfigSectionMap("Paths")['output_path']
xls_file_name = ConfigSectionMap("Paths")['xls_file_name']
txt_report_file_name = ConfigSectionMap("Paths")['txt_report_file_name']

xls_filename = output_path + xls_file_name + "_" + str(year) + str(month) + str(day) + ".xls" 
txt_report = output_path + txt_report_file_name + "_" + str(year) + str(month) + str(day) + ".txt"
title_x_offset_from_left = 4 #offset from the left hand margin of spreadsheet to start drawing at
title_y_offset_from_top = 6  #offset from the top margin for the spreadsheet to start drawing at
title_x_coord = 0
blade_x_coord = 0
title_y_coord = 0
first_blade_x = 0
first_blade_y = 0
wb = Workbook()
data = {}  #Hash datastructure

#Helper method to get the value from the various sections in conf file
#def ConfigSectionMap(section):
#    dict1 = {}
#    options = xls_Config.options(section)
#    for option in options:
#        try:
#            dict1[option] = xls_Config.get(section, option)
#            if dict1[option] == -1:
#                DebugPrint("skip: %s" % option)
#        except:
#            print("exception on %s!" % option)
#            dict1[option] = None
#    return dict1

#************* DEBUGGING *************
#Note also that keys in sections are case-insensitive and stored in lowercase
#Live_color = ConfigSectionMap("Environments")['live']
#print "\tLive_color ", Live_color 
#Test_color = ConfigSectionMap("Environments")['test']
#print "\tTest_color ", Test_color
#print "Temp debugging...\n" 
#Blade_Environments = xls_Config['Environments']
#for key in ConfigSectionMap("Environments"): print(key)
#'live' in ConfigSectionMap("Environments") 
#xls_Config['Environments']['live']
#time.sleep(10)
#exit(0)

#************ END OF DEBUGGING *****

f = open(txt_report, 'w')#Open a textfile for possible report at later stage

# Hash data structure is data_centers -> bladechassis -> blades
# Initialise hash and populate the keys, Store the bladechassis objects.
for dev in dmd.Devices.BladeChassis.HPBladeChassis.getSubDevices():
   #if chassis has a location add it to the list else skip
   if dev.getLocationName() != '': 
     data_center = dev.getLocationName().split('/')[1]
     # set data_center level
     data.setdefault( data_center, {})
     # set bladechasis
     data[data_center].setdefault( dev.id, {} )
     data[data_center][dev.id] = dev
   else:
     print "\n" + dev.id + " does not have a location... skipping"
     f.write("\n")
     f.write(dev.id + " : does not have a location set in zenoss... skipping\n") 

for data_center in data.keys():  #For each datacentre do
    #  Reset all variables for the next sheet
    title_x_coord = 0
    blade_x_coord = 0
    title_y_coord = 0
    first_blade_x = 0
    first_blade_y = 0
    #  end of setting variables
    ws = wb.add_sheet(data_center)#adds a new sheet for each datacentre
    ws.write(0,24,localtime,Style_nowrap)#adds datestamp to the top right
    for bladechassis in data[data_center].keys():  #for each bladechassis in current datacenter do
      f.write("\n")
      f.write("%s\n" % data[data_center][bladechassis].id) #write bladechassis name to txt report
      #Calculate the position of the first chassis on the sheet for current datacenter
      if title_y_coord == 0:
        title_y_coord = title_y_coord + title_y_offset_from_top
      if title_x_coord == 0:
        title_x_coord = title_x_coord + title_x_offset_from_left
      else:
        if title_x_coord < 20:
           title_x_coord = title_x_coord + 10
        else:
           #if chassis is the fifth to be drawn then reset variables to drawn on the next row below
           title_x_coord = title_x_offset_from_left
           title_y_coord = title_y_coord + 10
      #Following writes the power consumption for chassis
      ws.write(title_y_coord + 5, title_x_coord - 1,"Current draw in Amps: ",Power_txt)
      current_chassis = find(bladechassis)
      #bladechassis.getRRDValue('powerPresentAC')
      if isinstance(current_chassis.getRRDValue('powerPresentAC'), float):
         current_chassis_power_draw = current_chassis.getRRDValue('powerPresentAC')/240
         print current_chassis.getRRDValue('powerPresentAC')/240
      else:
         current_chassis_power_draw = 0
         print current_chassis_power_draw
      ws.write(title_y_coord + 5, title_x_coord + 2,current_chassis_power_draw,Power_decimal)
      #End of power consumption details
      blade_numin_chassis = 1            #sets the initial blade to deal with in the chassis
      first_blade_x = title_x_coord - 3  #sets the position of the top left hand blade for current chassis being drawn
      first_blade_y = title_y_coord + 2  #sets the position of the top left hand blade for current chassis being drawn.
      ChassisName = data[data_center][bladechassis].id
      ChassisName = ChassisName.split('.')[0]
      ws.write(title_y_coord,title_x_coord,ChassisName.upper(),Style_nowrap)  #writes the title for the chassis to the current sheet
      for blade in data[data_center][bladechassis].bladeservers():  #for each blade in the current chassis in the current datacentre being drawn do
           BladeNametxt = blade.bsDisplayName #sets the bladename
           if blade_numin_chassis > 8:        #if the blade is in slot 9 or above move varibles to start populating the row below in chassis
              blade_y_coord = first_blade_y + 1
              blade_x_coord = blade_numin_chassis - 9 + first_blade_x
           else:                              #else set the varibles to equal the slot to the the right of the last blade
              blade_x_coord = blade_numin_chassis + first_blade_x - 1
              blade_y_coord = first_blade_y
              #
              #lastly for all blades if empty slot hightlight with colour, else use varibles to write the blades name into the slot
              #
           if BladeNametxt == "Empty Slot":
              blade_color = ConfigSectionMap("Environments")["empty"]
              Style_blade = easyxf("font: height 210; pattern: pattern solid, fore_colour " + blade_color + "; align: rota 90, horiz centre, vert centre;" "borders: top double, bottom double, left double, right double;")
              ws.write(blade_y_coord,blade_x_coord,BladeNametxt,Style_blade)
           else:
              blade_color = ConfigSectionMap("Environments")["default"]
              Style_blade = easyxf("font: height 210; pattern: pattern solid, fore_colour " + blade_color + "; align: rota 90, horiz centre, vert centre;" "borders: top double, bottom double, left double, right double;")
              try:
                 dev = ''
                 name_to_search_4 = BladeNametxt + ".*" #added . as it was matching ilo devices
                 dev = find(name_to_search_4)
                 all_systems_strings = dev.getSystemNames()
                 platform_found = ''
                 for string_item in all_systems_strings:
                    if string_item.find('/Platform/') != -1:
                      #must conver the string to lowercase as the dict created from the config file is in lowercase
                      environ = string_item.split('/').pop().lower()
                 print environ
                 #The following sets the the default cell color for the blade, this will be overwritten if environment is matched 
                 #Style_blade = easyxf("font: height 210; pattern: pattern solid, fore_colour grey25; align: rota 90, horiz centre, vert centre;" "borders: top double, bottom double, left double, right double;")
                 if environ in ConfigSectionMap("Environments"):
                   blade_color = ConfigSectionMap("Environments")[environ] 
                   Style_blade = easyxf("font: height 210; pattern: pattern solid, fore_colour " + blade_color + "; align: rota 90, horiz centre, vert centre;" "borders: top double, bottom double, left double, right double;")  
                   ws.write(blade_y_coord,blade_x_coord,BladeNametxt,Style_blade)
                 else:
                   ws.write(blade_y_coord,blade_x_coord,BladeNametxt,Style_blade)
              except:
                 print "FQN of Host not found, no env, grey colour."
                 ws.write(blade_y_coord,blade_x_coord,BladeNametxt,Style_blade)
           blade_numin_chassis = blade_numin_chassis + 1 #increment the blade number
           ws.row(blade_y_coord).height = 1750  #set the heights and dimensions of the cells
           ws.row(blade_y_coord).height_mismatch = True
           ws.col(blade_x_coord).width = 1300
           #write the output also to the txt file for a possible report
           output = "Chassis: " + ChassisName.ljust(25) + "Num: " + str(blade_numin_chassis - 1).ljust(3) + " Name: " + (BladeNametxt.split('.')[0]).ljust(25) + "Serial: " + blade.bsSerialNum.ljust(12) + "IP: " +  blade.bsIloIp
           f.write("%s\n" % output)
           #And now go back around the loop processing the next bladechassis

#Add a sheet to the workbook with the key
ws = wb.add_sheet("Key")
#ws.write(1,0,"Key",Style_nowrap)
#ws.write(3,0," ",Style_red)
#ws.write(4,0," ",Style_dark_green)
x = 0
y = 1
ws.write(y,x,"KEY",Style_nowrap) 
y = 3 
for key in ConfigSectionMap("Environments"): 
  print(key)
  blade_color = ConfigSectionMap("Environments")[key]
  Style_key = easyxf("font: height 210; pattern: pattern solid, fore_colour " + blade_color + ";" "borders: top double, bottom double, left double, right double;")
  ws.write(y,x,key,Style_key)
  y = y + 1 
  #print(key)


#ws.write(3,1,"Live",Style_basic_text)
#ws.write(4,1,"Greenhouse",Style_basic_text)
#Adds some further notes during development
ws.write(20,0,"Notes:",Style_nowrap)
ws.write(22,0,"Currently chassis with BL620 don't appear due to a modelling error in Zenoss",Style_basic_text)
ws.write(23,0,"Page dimensions and layout options need to be moved to conf file",Style_basic_text)
ws.write(24,0,"Wrapper needs to be put in place to run this report nightly, eg. check into SVN, email the diff between reports to FMT Inf",Style_basic_text) 

f.close()#close txt report file
wb.save(xls_filename)#save the xls workbook
print "Written ", xls_filename
print xls_filename
print txt_report
print localtime

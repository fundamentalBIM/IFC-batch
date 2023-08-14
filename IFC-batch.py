import ifcopenshell
import os
import csv
import time
import io
from pprint import pprint
import ifcopenshell.validate
from rich import print

#Checking data from multiple IFC models. 
#Imports IFC files from the input path. It does a validity check, and exports a summary report to the output csv.

path = 'C:\\users\\dansi\\downloads\\load\\'     #Input path
output = 'output.csv'                  #Output filepath 

#remove old csv
files = os.listdir(path)
if os.path.exists(output):
    os.remove(output)

with open(output,"w",newline="") as csvfile:
    headernames=["filename",
    "time_stamp",
    "schema",
    "BAT",
    "BAT-version",
    "IfcProject.Name",
    "IfcProject.LongName",
    "IfcSite.Name",
    "IfcSite.LongName",
    "IfcBuilding.Name",
    "IfcBuilding.LongName" ,
    "n-building_element_proxy",
    "IfcClassification.Name"]
    
    writer=csv.DictWriter(csvfile,headernames)
    writer.writeheader()
    
    print("Loading "+str(len(files))+" files -> Displaying csv headers, outputting to "+output)
    
    pprint(headernames)
    fext = ''    
for file in files:
    ext = os.path.splitext(file)
    fext = ext[1]
    if ext[1] == '.ifc':
        start = time.time() #Timing load time for the data; Start
        model = ifcopenshell.open(path+file) #Load IFC-file
        json_logger = ifcopenshell.validate.json_logger()
        ifcopenshell.validate.validate(model, json_logger)
        if len(json_logger.statements) != 0:
            print(json_logger.statements) #Print error msg from validity check
        
        #Get IFC data
        project = model.by_type('IfcProject')[0]
        site = model.by_type('IfcSite')[0]
        building = model.by_type('IfcBuilding')[0]        
        building_element_proxies = model.by_type("IfcBuildingElementProxy")
        applications = model.by_type('IfcApplication')
        clasifications = model.by_type('IfcClassification')
        
        end = time.time() #Timer stop
        loadt = end - start
        loadt = f'{loadt:.2f}'
        file_stats = os.stat(path+file)
        filesize = file_stats.st_size / (1024 * 1024)
        print(file+" ("+f'{filesize:%.2f}'+"mb @ "+loadt+"s)")
        
        app_full_name = ''
        app_version = ''
                  
        for app in applications:
            app_full_name=app.ApplicationFullName
            app_version=app.Version               
            application_developer = app.ApplicationDeveloper
            
        cl_names = []  
        for cl in clasifications:
            cl_names.append(cl.Name)
            
        op = io.StringIO()
        writer = csv.writer(op, quoting=csv.QUOTE_NONNUMERIC, delimiter=';')
        writer.writerow(cl_names)
        cl_csv = op.getvalue()
        
        with open(output,"a",newline='') as csvfile:
            writer=csv.writer(csvfile, dialect='excel')
            writer.writerow([file,
            model.wrapped_data.header.file_name.time_stamp,
            model.schema,
            app_full_name,
            app_version,
            project.Name,
            project.LongName, 
            site.Name,
            site.LongName,
            building.Name, 
            building.LongName,
            len(building_element_proxies),
            cl_csv])
            
    else:
        print("Warning: expected .ifc file extension, got "+fext)

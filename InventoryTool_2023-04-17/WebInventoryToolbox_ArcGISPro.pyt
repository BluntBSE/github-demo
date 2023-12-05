# -*- coding: utf-8 -*-
############################################################################
#
#
#    This script iterates over web apps for whatever environment
#       the user is logged into for ArcGIS Pro.  It tries to pull data from
#       the apps (web maps, maps services, and feature services 
#       used by the apps).  It then enters relevant info into
#       an Excel template and as .txt files with some additional details.
#
#   The specific Excel template was developed by Adrien, Sarala, and Aymen
#
#   It uses several packages that do not automatically come installed
#
#   It includes a parser to strip extra HTML formatting from
#       AGOL descriptions and other text
#
#   11/9/2021 - fixed issue with layernames & URLs misaligning
#   12/21/2021 - added code to grab Dashboard & StoryMap map itemIds
#   12/28/2021 - corrected value j to start at 6
#                and also added a UseSelenium variable to make it optional
#                and adjusted the Portal selenium login process
#                and added Aymen's code for the user email
#   1/7/2022 - generalized tool and removed selenium references
#               and made into a script tool
#   1/14/2022 - added hyperlinks for URLs so they're clickable
#               and added homepage instead of ClientID; only left ClientID
#               for applications with an app ID.
#               added "Mobile Application", "Insights", "Native" to search
#   1/18/2022 - dropped esri_livingatlas from services; added special handling
#               for services that may freeze up ArcGIS Pro.
#   7/6/2022 - defaults set up for NDOT
#   7/7/2022 - adding portal service details (data store, etc.) 
#   7/19/2022 - adjusting version to remove requirement that Pro be open.
#   7/25/2022 - adding formatting for the services table
#   8/2/2022 - adjusting version to work better when Pro is closed.
#   10/11/2023 - added arcgis.gis.admin.AGOLAdminManager(port) to support arcgis v.2.1.0.2
#   10/15/2023 - Excel breaks if ANYTHING gets referenced poorly in openpyxl >3 (ex:
'''
toc = wb["Table Of Contentsxxx"]
Traceback (most recent call last):
  File "<string>", line 1, in <module>
  File "C:\Program Files\ArcGIS\Pro\bin\Python\envs\arcgispro-py3\Lib\site-packages\openpyxl\workbook\workbook.py", line 288, in __getitem__
    raise KeyError("Worksheet {0} does not exist.".format(key))
KeyError: 'Worksheet Table Of Contentsxxx does not exist.'
'''
#
#
#	Future dev: see https://developers.arcgis.com/python/guide/properties-of-your-gis/
#		use org_gis.properties.isPortal to see if URL is AGOL or Portal
#
#   Adapting to https://www.python.org/dev/peps/pep-0008/ per Arcadis stds
#
############################################################################

########################################################################################################################################
####Comment out this section when ArcGIS Pro gets updated
#Import the basic sys library to pull in a local copy of the arcgis Python package
import sys

paths = sys.path
#Reorder the paths so that packages are loaded from local paths before the default package location, 
#   because  I didn't have permission to update the default package location for Python on a remote computer.
#This is a workaround until ArcGIS Pro is updated.
try:
    sys.path = [p for p in paths if "USER" in p.upper()]+[p for p in paths if not "USER" in p.upper()] #revised 12/22/2022
except:
    pass
#This section is a workaround until ArcGIS Pro is updated.
########################################################################################################################################

#Install relevant libraries
#Import basic ArcGIS Pro packages for data access
try:
    import arcpy
    #https://developers.arcgis.com/python/api-reference/arcgis.gis.toc.html#item
    #pip install openpyxl (from Python command prompt)
    import arcgis
    #Import basic file writer package
    import csv
    import shutil #for copying template file
    #Import basic datetime management package
    from datetime import datetime, timezone
    import sys
    import os
    from arcgis.mapping import WebMap
except:
    arcpy.AddError("Unable to add appropriate libraries - please check required libraries (arcgis, csv, shutil, datetime, sys, arcpy, os, openpyxl).")
    print("Unable to add appropriate libraries - please check required libraries (arcgis, csv, shutil, datetime, sys, arcpy, os, openpyxl).")

#Need to get rid of all the HTML tags
from io import StringIO
from html.parser import HTMLParser


#Handle getting the portal info manually vs. out of the Pro instance.
def getPortalURL(portalurl, username, password):
    return arcgis.gis.GIS(portalurl, username, password)

if 'ArcGISPro.exe' in sys.executable:
    print("Running in ArcGIS Pro; using the active Portal in ArcGIS Pro for Portal access")
    portDefault = arcgis.gis.GIS("Pro")
else:
    print("Running outside of ArcGIS Pro; requires supplying username/password to access Portal")
    print("call getPortalURL to get the 'port' variable as an input for other functions")
    portDefault = ''

agency = "NDOT"

class Toolbox(object):
    def __init__(self):
        """Define the toolbox (the name of the toolbox is the name of the
        .pyt file)."""
        self.label = "WebInventoryToolbox"
        self.alias = "webtoolbox"

        # List of tool classes associated with this toolbox
        self.tools = [AppTool, SvcTool, ScrapeRestEndpoints]


#Define basic functions
#This is used to generate messages
def printaddmsg(string, warnlevel = "msg"):
    if warnlevel == "msg":
        print(string)
        arcpy.AddMessage(string)
    if warnlevel == "warn":
        print("** "+str(string)+" **")
        arcpy.AddWarning(string)
    if warnlevel == "err":
        print("!!! "+str(string)+" !!!")
        arcpy.AddError(string)
#This is used to strip out extra HTML tags from descriptions
class MLStripper(HTMLParser):
    def __init__(self):
        super().__init__()
        self.reset()
        self.strict = False
        self.convert_charrefs =  True
        self.text = StringIO()
    def handle_data(self, d):
        self.text.write(d)
    def get_data(self):
        return self.text.getvalue()

def strip_tags(html):
    s = MLStripper()
    s.feed(html)
    return s.get_data()

try:
    import openpyxl #This lets us modify Excel documents
except:
    printaddmsg("Unable to load openpyxl library.  Need to run pip install openpyxl (from Python command prompt).", "err")
    exit()
    
#This is used to generate the search table, which can be used in all/any of the tools.

def svcinfo(outtabloc, portalurl = "", username = "", password = "", port = "", xlsx = ""):
    
    #Define service property function
    def getsvcdetails(svc, port):
        url = svc.url.replace("admin/ser", "rest/ser").replace(".GPServer", "/GPServer").replace(".MapServer", "/MapServer").replace(".FeatureServer", "/FeatureServer").replace(".ImageServer", "/ImageServer")
        try:
            fp = svc.properties["properties"]["filePath"]
        except:
            fp = ""
        try:
            title = svc.iteminformation.properties["title"]
        except:
            try:
                title = svc.properties["serviceName"]
            except:
                title = ''
        #'FeatureServer':
        #   mgr.servers.list()[5].services.list()[2].properties.extensions[6]["typeName"]
        #try:
        #    un = svc.properties["properties"]["userName"]
        #except:
        #    un = ""
        try:
            tags = svc.iteminformation.properties["tags"]
        except:
            tags = ""
        try:
            status = svc.status["realTimeState"]
            desc = strip_tags(svc.properties["description"])
        except:
            status = ""
            desc = ""
        try:
            mxd = svc.iteminformation.manifest["resources"][0]["onPremisePath"]
            #svc.iteminformation.manifest["databases"][0]
            computerfp = svc.iteminformation.manifest["resources"][0]["clientName"]
        except:
            mxd = ""
            computerfp = ""
        try:
            cat = ", ".join(port.content.get(svc.portalProperties["portalItems"][0]["itemID"]).categories)
        except:
            cat = ""
        try:
            DB = svc.iteminformation.manifest["databases"]
            dstemp = []
            for x in range(len(DB)):
                dstemp +=  ["*"+DB[x]["onServerConnectionString"][max(0, DB[x]["onServerConnectionString"].find("SERVER")):]+"*: "+", ".join([y["onServerName"] for y in DB[x]["datasets"]])]
            ds = ".  ".join(dstemp) + "."
        except:
            ds = ""
        extensions = ""
        try:
            ext = svc.properties.extensions
            extensions = [ext[i]["typeName"] for i in range(len(ext))]
        except:
            pass
        result = [title, url, mxd, computerfp, status, desc, tags, cat, ds, extensions]
        return result

    #Check whether the arcgis version is at least version 2.0.0
    if int(arcgis.__version__.split(".")[0])>= 2:
        pass
        #this has to be at least version 2.0
    else:
        printaddmsg("Version of arcgis module is too old.", "err")
        exit()

    #Begin function
    date = format(datetime.now(), "%m-%d-%Y_%I%M_%p")


    if port == "":
        port = arcgis.gis.GIS(portalurl, username, password)

    if not "arcgis" in port.url:
        mgr = arcgis.gis.admin.PortalAdminManager(port.url+"//sharing/rest", gis = port)
    else:
        mgr = arcgis.gis.admin.AGOLAdminManager(port)
    
    tabname = outtabloc+"/servicesinventory_"+date+".txt"

    with open(tabname, 'w', newline = '', encoding = "utf-8") as f:
        fw = csv.writer(f, delimiter = "\t")
        fw.writerow(["Title", "URL", "SourceMap", "SourceLocation", "Status", "Description", "Tags", "Categories", "DataSources"])
        #This fails for AGOL:
        for serv in mgr.servers.list():
            try:
                printaddmsg("Processing "+serv.url, "msg")
                #svcmgr = arcgis.gis.server.ServiceManager(serv.url, gis = port)
                #svcmgr.folders
                #serv.datastores.list()[0].properties
                #serv.datastores.list()[0].properties.info("machines")[0]["name"]
                #serv.datastores.list()[0].properties.info("connectionString")
                #svc.iteminformation.manifest["databases"][1]["onServerConnectionString"]  
                #serv.datastores.list()[0].url (like a title)
                out = ""
                for svc in serv.services.list():
                    out = getsvcdetails(svc, port)                
                    fw.writerow(out[:-1])
                    if isinstance(out[-1:],list):
                        if "FeatureServer" in out[-1:][0]:
                            outfs = out[:-1]
                            outfs[1] = outfs[1].replace("MapServer","FeatureServer")
                            fw.writerow(outfs)
                try:
                    folders = serv.services.folders
                except:
                    folders = []
                if len(folders) > 0:
                    for fold in folders:
                        if not fold ==  "/":
                            out = ""
                            for svc in serv.services.list(fold):
                                out = getsvcdetails(svc, port)
                                fw.writerow(out[:-1])
                                if isinstance(out[-1:],list):
                                    if "FeatureServer" in out[-1:][0]:
                                        outfs = out[:-1]
                                        outfs[1] = outfs[1].replace("MapServer","FeatureServer")
                                        fw.writerow(outfs)
            except:
                printaddmsg("Processing "+port.url, "msg")
                out = ""
                for svc in serv.services:
                    out = getsvcdetails(svc, port)
                    if out[0]=="":
                        if "name" in svc.properties:
                            try:
                                out[0] = svc.properties.name
                            except:
                                pass
                    fw.writerow(out[:-1])
                    if isinstance(out[-1:],list):
                        if "FeatureServer" in out[-1:][0]:
                            outfs = out[:-1]
                            outfs[1] = outfs[1].replace("MapServer","FeatureServer")
                            fw.writerow(outfs)
        printaddmsg("Completed.  Output is in the tab-delimited file "+tabname, "msg")
        
    if ".xls" in xlsx:
        #Add the table!!
        wb = openpyxl.load_workbook(xlsx)
        ws = wb.create_sheet("PortalServiceDetails")
        with open(tabname) as f:
            reader = csv.reader(f, delimiter = '\t')
            for row in reader:
                ws.append(row)
                #if "FeatureServer" in row[1]:
                #    outfs = row
                #    outfs[1] = outfs[1].replace("MapServer","FeatureServer")
                #    ws.append(outfs)
        #Add to TOC
        toc = wb["Table Of Contents"]
        link = "#'PortalServiceDetails'!A1"
        toc.cell(row = 13, column = 2).value = "Portal Service Details"#Revised to show the title!
        toc.cell(row = 13, column = 2).hyperlink = (link)
        toc.cell(row = 13, column = 2).style = "Hyperlink"                            
        #Add table
        table = openpyxl.worksheet.table.Table(displayName = "tblSvc", ref = "A1:" + openpyxl.utils.get_column_letter(ws.max_column) + str(ws.max_row))
        style = openpyxl.worksheet.table.TableStyleInfo(name="TableStyleLight9", showFirstColumn=True,
                               showLastColumn=False, showRowStripes=False, showColumnStripes=False)							
        table.tableStyleInfo = style
        ws.add_table(table)
        for ht in range(2,ws.max_row+1):
            ws.row_dimensions[ht].height = 75 #125
            for c in range(1,ws.max_column+1):
                ws.cell(row = ht, column = c).alignment = openpyxl.styles.Alignment(wrapText = True)                
        ws.column_dimensions['A'].width = 20 #187
        ws.column_dimensions['B'].width = 22 #205
        ws.column_dimensions['C'].width = 20 #187
        ws.column_dimensions['F'].width = 12 #115
        ws.column_dimensions['G'].width = 12 #115
        ws.column_dimensions['I'].width = 78 #709
        ws.freeze_panes = "B2" #freeze row 1/column A            
        #for row in ws.iter_rows():
        #    for cell in row:
        #        cell.style.alignment.wrap_text = True
        ws.cell(row = 10, column = 1).value = "Press F5 (then Enter) to return to the previous active cell."
        wb.save(xlsx)
        wb.close()
        printaddmsg("Completed.  Details of web services in Portal have been added to the workbook.")        
    return

"""The source code of the tool."""

def runAppTool(template, outdir, importdtls, getportdtls, gis="", un="", pw="", portalURL=""):
        if gis == "":
            if un == "":
                if 'ArcGISPro.exe' in sys.executable:
                    gis = arcgis.gis.GIS("Pro")
            else:
                gis = getPortalURL(portalURL, un, pw)
        #Define the web applications to search for!
        ItemTypes = ['Web Mapping Application', 'StoryMap', 'Dashboard', 'Mobile Application', 'Insights', 'Native']#'StoryMap', , 'Web Mapping Application'
        #https://developers.arcgis.com/rest/users-groups-and-items/items-and-item-types.htm

        #Determine if it's AGOL or portal
        URLbase = gis.properties["customBaseUrl"].replace("maps.arcgis.com", "AGOL")
        newxl = template.replace("Autopopulate", URLbase)
        newxl = outdir+"/"+os.path.basename(newxl)
        shutil.copy2(template, newxl)
        irritated=1
        
        if importdtls not in ("#", "None", "", None, "Null"):
            PortalDtls = True
            #Add the table!!
            wb = openpyxl.load_workbook(newxl)
            ws = wb.create_sheet("PortalServiceDetails")
            with open(importdtls) as f:
                reader = csv.reader(f, delimiter = '\t')
                for row in reader:
                       ws.append(row)
                    #if "FeatureServer" in row:
                    #    outfs = row
                    #    outfs[1] = outfs[1].replace("MapServer","FeatureServer")
                    #    ws.append(outfs)
                    
            #Add to TOC
            toc = wb["Table Of Contents"]

            link = "#'PortalServiceDetails'!A1"
            toc.cell(row = 13, column = 2).value = "Portal Service Details"#Revised to show the title!
            toc.cell(row = 13, column = 2).hyperlink = (link)
            toc.cell(row = 13, column = 2).style = "Hyperlink"

            printaddmsg("NotBrokenYet "+str(irritated))
            irritated+=1

            #Add table
            table = openpyxl.worksheet.table.Table(displayName = "tblSvc", ref = "A1:" + openpyxl.utils.get_column_letter(ws.max_column) + str(ws.max_row))
            style = openpyxl.worksheet.table.TableStyleInfo(name="TableStyleLight9", showFirstColumn=True,
                                   showLastColumn=False, showRowStripes=False, showColumnStripes=False)							
            table.tableStyleInfo = style
            ws.add_table(table)
            for ht in range(2,ws.max_row+1):
                ws.row_dimensions[ht].height = 75 #125
                for c in range(1,ws.max_column+1):
                    ws.cell(row = ht, column = c).alignment = openpyxl.styles.Alignment(wrapText = True)
            ws.column_dimensions['A'].width = 20 #187
            ws.column_dimensions['B'].width = 22 #205
            ws.column_dimensions['C'].width = 20 #187
            ws.column_dimensions['F'].width = 12 #115
            ws.column_dimensions['G'].width = 12 #115
            ws.column_dimensions['I'].width = 78 #709
            ws.freeze_panes = "B2" #freeze row 1/column A
            ws.cell(row = 10, column = 1).value = "Press F5 (then Enter) to return to the previous active cell."
            wb.save(newxl)
            wb.close()
            printaddmsg("Completed.  Details of web services in Portal have been added to the workbook.")        
        if getportdtls not in ("#", "None", "", None, "Null"):
            PortalDtls = True
            svcinfo(outdir, port = gis, xlsx = newxl)
            printaddmsg("NotBrokenYet "+str(irritated))
            irritated+=1
        #Get the worksheet "Web Application Template" that should exist in the template .xlsx
        wb = openpyxl.load_workbook(newxl)
        
        toc = wb["Table Of Contents"]
        toc["B8"] = "The Web Applications contained in the spreadsheet are based on what exist within " + agency
        
        try:
            origws =  wb['Web Application Template']
        except Exception as e:
            printaddmsg("  "+str(e))
            #arcpy.AddMessage(str(e)) #arcpy.AddError(e.args[0])
            #https://pro.arcgis.com/en/pro-app/2.7/arcpy/get-started/error-handling-with-python.htm
        i = 14 #Set the default row on the Table of Contents for listing links to new worksheets
        #Track the names of the new worksheets so duplicate names can be made unique
        Worksheets = []
        wb.close()
        
        #Iterate through the item types
        for itemtype in ItemTypes:
            wb = openpyxl.load_workbook(newxl)
            try:
                origws =  wb['Web Application Template']
                toc = wb["Table Of Contents"]
            except Exception as e:
                printaddmsg("  "+str(e))
            printaddmsg("NotBrokenYet "+str(irritated))
            irritated+=1

            FailedToSave = []
            #And make output txt files for each combination (to support)
            with open(outdir + "/inventory_" + str(datetime.strftime(datetime.now(), "%Y-%m-%d")) + itemtype + "_" + URLbase + ".txt", 'w', newline = '', encoding = "utf-8") as f:
                printaddmsg(itemtype, "msg")
                printaddmsg(URLbase, "msg")
                #if itemtype  ==  "Web Mapping Application" and URLbase  ==  "AGOL":
                #    print("Did these already.")
                #    continue #break out and go back to the higher level of the loop!
                #else:
                #Load the portal or AGOL items, as appropriate
                items = gis.content.search(query = '', item_type = itemtype, max_items = 1000) #Cycle through app types
                #https://techcommunity.microsoft.com/t5/excel/convert-image-url-to-actual-image-in-excel/m-p/309020
                #Write out the column names to the tab-delimited file
                fw = csv.writer(f, delimiter = "\t")
                fw.writerow(["Title", "ShortDescription", "ClientID", "Link", "DirectLink", 
                "Description", "Tags", "Access", 
                "Owner", "Email", "Created", "Modified", 
                "Views", "Unknown", "Type", "SharedWith", 
                "Credits", "Comments", "SpecialStatus", 
                "DependentTo", "DependentUpon", #"Metadata", 
                "Use", "LayerNames", "LayerURLs", "1YrUsage", "1MonthUsage", "Thumbnail", "Broken"])
                #https://developers.arcgis.com/python/api-reference/arcgis.apps.storymap.html#journalstorymap
                #   The API reference above only works for newer items.  Still unsure how to get
                #   the web map for a story map without opening the map & scraping the appID or
                #   opening "Configure" for the app, hence "selenium" package & dependencies
                #https://developers.arcgis.com/python/api-reference/arcgis.apps.dashboard.html#details
                #Iterate through each item in the portal or AGOL that is of the particular itemtype (Dashboard, StoryMap, or web app)
                if (len(items) > 0):
                    for item in items:
                        printaddmsg(item.title or "No Title")
                        try:
                            #depto = item.dependent_to() #arcgis enterprise only
                            #depon = item.dependent_upon() #arcgis enterprise only
                            layernames = ""
                            layerurls = ""
                            appinfo = item.app_info #details if app is registered with an App ID and App Secret
                            try:
                                sharing = item.shared_with #dictionary in the following format: { ‘groups’: [], # one or more Group objects ‘everyone’: True | False, ‘org’: True | False }
                                if len(sharing["groups"])>0:
                                    groups = [g.title for g in sharing["groups"]]
                                else:
                                    groups = []
                                shared = ", ".join([g.title for g in sharing["groups"]]) + (", Everyone" if sharing["everyone"] == True else "")  + (", Organization" if sharing["org"] == True else "")
                            except Exception as e:
                                printaddmsg(str(e))
                                arcpy.AddMessage(str(e)) #arcpy.AddError(e.args[0])
                                shared = ""
                            try:
                                if not URLbase == "portal":
                                    yr1 = sum(item.usage(date_range = '1Y', as_df = True).Usage) #don't convert to Pandas DF
                                    day30 = sum(item.usage(date_range = '30D', as_df = True).Usage) #don't convert to Pandas DF
                                else:
                                    yr1 = None
                                    day30 = None
                            except Exception as e:
                                printaddmsg(str(e))
                                yr1 = None
                                day30 = None
                                #arcpy.AddMessage(str(e)) #arcpy.AddError(e.args[0])
                            # enable browser logging
                            mapids = [] #To hold the ID combinations that define maps in AGOL/Portal
                            errorlist = [] #To hold those where we couldn't pull anything out of Selenium as a last resort to pulling dependencies
                            newt = item.title.replace("/", "").replace("\\", "").replace("*", "").replace("[", "").replace("]", "").replace(":", "").replace("?", "")
                            newt = newt[0:min(30, len(newt))]
                            #If the there's already a sheet with this name, append a number
                            if (newt in Worksheets):
                                k = 1
                                while newt in Worksheets:
                                    newt = newt[0:min(27, len(newt))].replace("(" + str(k-1) + ")", "") + "(" + str(k) + ")"
                                    k +=  1
                                #print ("Did not add " + newt + " due to multiple duplicates.")
                                #continue
                            try:
                                mapids = []
                                errorlist = []
                                dependentURL = [] #Didn't do anything with this b/c it never seemed to be populated
                                #If portal, we can
                                #       use print("\n\n Item: \t {2} \t\t\t id {0} is dependent on these items: \t {1}".format(item.itemid, item.dependent_upon(), item.title))
                                #       
                                #if (item.dependent_to()['total'] > 0):
                                #   print("\t This item is also a dependency of these items: {}".format(item.dependent_to()))
                                if (URLbase  ==  "portal" and "list" in item.dependent_upon().keys()):
                                    dependencies = item.dependent_upon()["list"]
                                    if len(dependencies) > 0:
                                        for dep in dependencies:
                                            if "id" in dep.keys():
                                                mapids.append(dep["id"])
                                            if "url" in dep.keys():
                                                dependentURL.append(dep["url"]) #NEED TO FIGURE OUT IF THIS IS EVER POPULATED, AND IF SO, WHERE TO STORE IT
                                        del(dep, dependencies)
                                if (len(mapids) == 0):
                                    #alldata = item.get_data()
                                    if (itemtype  == "Dashboard"):
                                        try:
                                            widgets = item.get_data()["widgets"] #Each widget could have its own map dependencies
                                            for widget in widgets:
                                                try:
                                                    if widget["type"]  ==  "mapWidget":
                                                        mapids.append(widget["itemId"])
                                                    elif widget["type"] in ("pieChartWidget", "serialChartWidget", "indicatorWidget", "listWidget"):#still need GAUGE, DETAILS, and TABLE
                                                        if "itemId" in widget["datasets"][0]["dataSource"].keys():
                                                            mapids.append(widget["datasets"][0]["dataSource"]["itemId"])
                                                except Exception as e:
                                                    printaddmsg(str(e), "warn") #arcpy.AddError(e.args[0])
                                                    pass
                                        except Exception as e:
                                            printaddmsg(str(e), "warn") #arcpy.AddError(e.args[0])
                                            pass
                                            #mapids.append(widget["id"]) #I think this other id is a unique ID for the particular widget configuration in the dashboard
                                        #elif widget["type"]  ==  :#still need GAUGE, DETAILS, and TABLE
                                        #
                                    elif (itemtype  ==  "StoryMap"):
                                        try:
                                            widgets = item.get_data()["resources"] #Each "resource" could have its own map dependencies
                                            for key in widgets.keys():
                                                try:
                                                    widget = widgets[key]
                                                    mapids.append(widget["data"]["itemId"])
                                                except Exception as e:
                                                    printaddmsg(str(e), "warn") #arcpy.AddError(e.args[0])
                                                    pass
                                        except Exception as e:
                                            printaddmsg(str(e), "warn") #arcpy.AddError(e.args[0])
                                            pass
                            except Exception as e:
                                printaddmsg(str(e), "warn") #arcpy.AddError(e.args[0])
                                if len(mapids) == 0:
                                    print("Failed totally to get any mapIDs")
                                else:
                                    print("Got some mapIDs but generated an error")
                            #Start filling out a new sheet in the workbook
                            printaddmsg("NotBrokenYet "+str(irritated))
                            irritated+=1
                            ###################################
                            new = wb.copy_worksheet(origws)
                            ###################################
                            new.active = 1
                            new.title = newt
                            Worksheets.append(newt) #Keep track of what we've added to monitor for duplicate names
                            new.cell(row = 1, column = 1).value = item.title #A1
                            new.cell(row = 2, column = 1).value = item.snippet #A2
                            new.cell(row = 4, column = 3).value = (appinfo["client_id"] if len(appinfo)>0 else "") #C4
                            new.merge_cells(start_row = 5, start_column = 3, end_row = 6, end_column = 3)
                            new.merge_cells(start_row = 7, start_column = 3, end_row = 9, end_column = 3)
                            new.cell(row = 5, column = 3).value = (item.url if item.url else item.homepage.replace("/home//home", "/home")) #C5
                            new.cell(row = 5, column = 3).hyperlink = (item.url if item.url else item.homepage.replace("/home//home", "/home"))
                            new.cell(row = 5, column = 3).style = "Hyperlink"
                            new.cell(row = 7, column = 3).value = (strip_tags(item.description) if not item.description == None else "") #C6
                            new.cell(row = 10, column = 3).value = ", ".join(item.tags if not item.tags is None else "") #C7
                            new.cell(row = 11, column = 3).value = (strip_tags(item.licenseInfo) if not item.licenseInfo == None else "") #C8
                            try:
                                email = ''
                                user = gis.users.get(username = item.owner)
                                email = user.email
                                Name = user.fullName
                                new.cell(row = 12, column = 3).value = item.owner + " " + email
                            except Exception as e:
                                #printaddmsg(str(e), "warn") #arcpy.AddError(e.args[0])
                                new.cell(row = 12, column = 3).value = item.owner
                            #new.cell(row = 12, column = 3).value = item.owner #C9
                            new.cell(row = 14, column = 3).value = datetime.fromtimestamp(item.created/1000).strftime("%m/%d/%Y")##Date/time was in epoch time to the millisecond so /1000 #C11
                            new.cell(row = 15, column = 3).value = datetime.fromtimestamp(item.modified/1000).strftime("%m/%d/%Y") #C12
                            new.cell(row = 16, column = 3).value = item.numViews #C13
                            new.cell(row = 17, column = 3).value = ("Geocortex" if "ortex" in (item.url or "") else "")#Technology #C14
                            if ("MapSeries" in (item.url or "")
                                or "MapJournal" in (item.url or "")
                                or "AGOStoryMaps" in (item.url or "")
                                or "MapTour" in (item.url or "")
                                or "Cascade" in (item.url or "")
                                or "story.maps" in (item.url or "")
                                or itemtype  ==  'StoryMap'):
                                    new.cell(row = 18, column = 3).value = "Story Map"#string #C15
                            else:
                                if (item.url !=   None):
                                    if ("xperience" in item.url and "uilder" in item.url):
                                        new.cell(row = 18, column = 3).value = "Experience Builder"
                                    elif ("webappviewer" in item.url):
                                        new.cell(row = 18, column = 3).value = "Web App Builder"
                                    elif ("dashboard" in item.url):
                                        new.cell(row = 18, column = 3).value = "Dashboard"
                                    elif item.type:
                                        new.cell(row = 18, column = 3).value = item.type#string #C15
                                else:
                                    if item.type:
                                        new.cell(row = 18, column = 3).value = item.type#string #C15
                                #Or should this be "itemtype" instead of "item.type"?  Not sure which is more specific...
                            new.cell(row = 19, column = 3).value = shared#'Group Name' #C16
                            new.cell(row = 20, column = 3).value = ''#'Users' #C17
                            new.cell(row = 21, column = 3).value = item.homepage.replace("/home//home", "/home") #'' #C18
                            new.cell(row = 21, column = 3).hyperlink = item.homepage.replace("/home//home", "/home")
                            new.cell(row = 21, column = 3).style = "Hyperlink"
                            new.cell(row = 22, column = 3).value = item.access #C19
                            new.cell(row = 24, column = 3).value = ''#Server URL' #C21
                            if PortalDtls == True:
                                new.cell(row = 5, column = 8).value = "Service Details"
                            #Then loop through the data layers, if available
                            j = 6
                            if (len(errorlist)>0):
                                for err in errorlist:
                                    new.cell(row = j, column = 5).value = (err or "")
                                    new.cell(row = j, column = 6).value = "Error"
                                    j +=  1
                            if (item.type == "Web Map"): #If it's a web map (not an app), then pull layers this way
                                try:
                                    webmap = WebMap(item)
                                    layers = webmap.layers
                                    layernames = ", ".join([layer.title or "" for layer in layers])
                                    layerurls = ", ".join([layer.url or "" for layer in layers])
                                    for lyr in layers:
                                        if not ' Config' in (lyr.title or ""):
                                            #        #Need to update this script to include creating URLs to PortalServiceDetails
                                            j +=  1
                                            new.cell(row = j, column = 5).value = (layer.url or "")
                                            if not layer.url == None:
                                                new.cell(row = j, column = 5).hyperlink = layer.url
                                                new.cell(row = j, column = 5).style = "Hyperlink"
                                                if PortalDtls == True and URLbase == "portal":
                                                    new.cell(row = j, column = 8).value = '=HYPERLINK(CELL("address",INDEX(tblSvc[URL],MATCH(LEFT(E'+str(j)+\
                                                        ',IFERROR(FIND("FeatureServer",E'+str(j)+')+12,IFERROR(FIND("MapServer",E'+str(j)+')+8,LEN(E'+str(j)+')))),tblSvc[URL],0))), "LayerDetails")'
                                                    #new.cell(row = j, column = 8).hyperlink = ('=HYPERLINK(CELL("address",INDEX(tblSvc[URL],MATCH(LEFT(E'+str(j)+\
                                                    #        ',IFERROR(FIND("FeatureServer",E'+str(j)+')+12,LEN(E'+str(j)+'))),tblSvc[URL])))')
                                                    new.cell(row = j, column = 8).style = "Hyperlink"
                                            new.cell(row = j, column = 6).value = (layer.title or "")
                                            #new.cell(row = j, column = 7).value = "Source/Target"
                                except Exception as e:
                                    printaddmsg(str(e), "warn") #arcpy.AddError(e.args[0])
                                    printaddmsg("Unable to pull layer names from the web map as expected", "warn")
                                    layernames = ""
                                    layerurls = ""
                            else:
                                #if item.type  ==  "Web Mapping Application", item.related_items("")
                                #use relationships or dependencies, not sure which https://developers.arcgis.com/rest/users-groups-and-items/relationship-types.htm
                                #If it isn't a webmap, then use mapIDs to output the dependent maps
                                try:
                                    mapids = set(mapids)
                                    new.cell(row = 23, column = 3).value = ", ".join(mapids if not mapids is None else "")
                                    for mapid in mapids:
                                        try:
                                            #Then try to go through each mapID and load all the layers within those maps - we're assuming all layers are required
                                            #   for simplicity's sake...
                                            #print("Printing out layers!!")
                                            if (URLbase  ==  "Portal"):
                                                webmap = WebMap(gisPortal.content.get(mapid))
                                                layers = webmap.layers
                                                layernames = ", ".join([layer.title or "" for layer in layers])
                                                layerurls = ", ".join([layer.url or "" for layer in layers])
                                            else:
                                                webmap = WebMap(gis.content.get(mapid))
                                                layers = webmap.layers
                                                layernames = ", ".join([layer.title or "" for layer in layers])
                                                layerurls = ", ".join([layer.url or "" for layer in layers])
                                            try:
                                                if (not " Config" in (webmap.item.title or "")):
                                                    j +=  1
                                                    new.cell(row = j, column = 5).value = (webmap.item.homepage.replace("/home//home", "/home") or "")
                                                    if (not webmap.item.homepage == None):
                                                        new.cell(row = j, column = 5).hyperlink = webmap.item.homepage.replace("/home//home", "/home")
                                                        new.cell(row = j, column = 5).style = "Hyperlink"
                                                        if PortalDtls == True and URLbase == "portal":
                                                            new.cell(row = j, column = 8).value = '=HYPERLINK(CELL("address",INDEX(tblSvc[URL],MATCH(LEFT(E'+str(j)+\
                                                                ',IFERROR(FIND("FeatureServer",E'+str(j)+')+12,IFERROR(FIND("MapServer",E'+str(j)+')+8,LEN(E'+str(j)+')))),tblSvc[URL],0))), "LayerDetails")'
                                                            #new.cell(row = j, column = 8).hyperlink = ('=HYPERLINK(CELL("address",INDEX(tblSvc[URL],MATCH(LEFT(E'+str(j)+\
                                                            #        ',IFERROR(FIND("FeatureServer",E'+str(j)+')+12,LEN(E'+str(j)+'))),tblSvc[URL])))')
                                                            new.cell(row = j, column = 8).style = "Hyperlink"
                                                    new.cell(row = j, column = 6).value = (webmap.item.title or "") + " - WebMap"
                                            except Exception as e:
                                                printaddmsg(str(e), "warn") #arcpy.AddError(e.args[0])
                                                printaddmsg("Unable to get layers from the web map using gis.content.get(mapid)", "warn")
                                                pass
                                            for lyr in layers:
                                                if (not " Config" in (lyr.title or "")):
                                                    j +=  1
                                                    new.cell(row = j, column = 5).value = (lyr.url or "")
                                                    if not lyr.url == None:
                                                        new.cell(row = j, column = 5).hyperlink = lyr.url
                                                        new.cell(row = j, column = 5).style = "Hyperlink"
                                                        if PortalDtls == True and URLbase == "portal":
                                                            new.cell(row = j, column = 8).value = '=HYPERLINK(CELL("address",INDEX(tblSvc[URL],MATCH(LEFT(E'+str(j)+\
                                                                ',IFERROR(FIND("FeatureServer",E'+str(j)+')+12,IFERROR(FIND("MapServer",E'+str(j)+')+8,LEN(E'+str(j)+')))),tblSvc[URL],0))), "LayerDetails")'
                                                            #new.cell(row = j, column = 8).hyperlink = ('=HYPERLINK(CELL("address",INDEX(tblSvc[URL],MATCH(LEFT(E'+str(j)+\
                                                            #        ',IFERROR(FIND("FeatureServer",E'+str(j)+')+12,LEN(E'+str(j)+'))),tblSvc[URL])))')
                                                            new.cell(row = j, column = 8).style = "Hyperlink"
                                                    new.cell(row = j, column = 6).value = (lyr.title or "")
                                                    #new.cell(row = j, column = 7).value = "Source/Target"
                                        except Exception as e:
                                            printaddmsg(str(e), "warn") #arcpy.AddError(e.args[0])
                                            try:
                                                #This code is for if the mapID isn't actually a "map" id, but is instead a feature service or something along those lines.
                                                if (URLbase  ==  "Portal"):#"portal" in mapid:
                                                    fcs = gisPortal.content.get(mapid)
                                                else:
                                                    fcs = gis.content.get(mapid) #33884bb6efa44e11a4523a3d4bf968bc  fcs.keys()
                                                if (not " Config" in (fcs.name or fcs.title or "")):
                                                    j +=  1
                                                    new.cell(row = j, column = 5).value = (fcs.url or "")
                                                    #
                                                    new.cell(row = j, column = 6).value = (fcs.name or fcs.title or "")
                                                    for lyr in fcs.layers:
                                                        if (not " Config" in (lyr.properties.name or "")):
                                                            j +=  1
                                                            new.cell(row = j, column = 5).value = (lyr.url or "")
                                                            if not lyr.url == None:
                                                                new.cell(row = j, column = 5).hyperlink = lyr.url
                                                                new.cell(row = j, column = 5).style = "Hyperlink"
                                                                if PortalDtls == True and URLbase == "portal":
                                                                    new.cell(row = j, column = 8).value = '=HYPERLINK(CELL("address",INDEX(tblSvc[URL],MATCH(LEFT(E'+str(j)+\
                                                                        ',IFERROR(FIND("FeatureServer",E'+str(j)+')+12,IFERROR(FIND("MapServer",E'+str(j)+')+8,LEN(E'+str(j)+')))),tblSvc[URL],0))), "LayerDetails")'
                                                                    #new.cell(row = j, column = 8).hyperlink = ('=HYPERLINK(CELL("address",INDEX(tblSvc[URL],MATCH(LEFT(E'+str(j)+\
                                                                    #        ',IFERROR(FIND("FeatureServer",E'+str(j)+')+12,LEN(E'+str(j)+'))),tblSvc[URL])))')
                                                                    new.cell(row = j, column = 8).style = "Hyperlink"
                                                            new.cell(row = j, column = 6).value = (lyr.properties.name or "")
                                            except Exception as e:
                                                printaddmsg(str(e), "warn") #arcpy.AddError(e.args[0])
                                                printaddmsg("Failed to get layers from non-map ID", "warn")
                                                pass
                                except Exception as e:
                                    printaddmsg(str(e), "warn") #arcpy.AddError(e.args[0])
                                    printaddmsg("No layernames/urls yet", "warn")
                                    layernames = ""
                                    layerurls = ""
                            #etc.


                            printaddmsg("NotBrokenYet "+str(irritated))
                            irritated+=1

                            mapids = []
                            toc = wb["Table Of Contents"]
                            #toc.insert_rows(i) #This part doesn't work due to package limitations
                            #If there are any links below, those links aren't moved with the text!
                            #So we'll just make a list in another section, NOT insert rows, 
                                #then move the data over.
                            #We'll have to undo the merge-across, then (if required) re-merge
                            #   across b/c otherwise copy/paste is not possible.
                            link = "#'" + newt + "'!A1"
                            toc.cell(row = i, column = 2).value = item.title#Revised to show the title!
                            toc.cell(row = i, column = 2).hyperlink = (link)
                            toc.cell(row = i, column = 2).style = "Hyperlink"
                            toc.cell(row = i, column = 3).value = ("Broken" if len(errorlist)>0 else "")
                            i +=  1

                            #Write a row to the appropriate tab-delimited .txt file with some additional info (like 1-year and 30-day usage stats)
                            fw.writerow([
                            item.title#string
                            , item.snippet#string
                            , (appinfo["client_id"] if (len(appinfo)>0 and not appinfo["client_id"] == None) else "") #" #dictionary of info for registered apps - irrelevant here, I would think..
                            , (item.homepage.replace("/home//home", "/home") if not item.homepage == None else "")
                            , (item.url if not item.url == None else "")
                            , (strip_tags(item.description) if not item.description == None else "")
                            , ", ".join(item.tags if not item.tags is None else "")
                            , (strip_tags(item.licenseInfo) if not item.licenseInfo == None else "")
                            , item.owner#string
                            , (email or "")
                            , datetime.fromtimestamp(item.created/1000).strftime("%m/%d/%Y")##Date/time was in epoch time to the millisecond so /1000
                            , datetime.fromtimestamp(item.modified/1000).strftime("%m/%d/%Y")
                            , item.numViews
                            , ""
                            , item.type#string
                            , shared
                            , item.accessInformation
                            , ""#, ", ".join([item.comments[i]["comment"] for i in range(len(item.comments))]).replace("%20", " ")
                            , item.content_status #deprecated or authoritative or None
                            , ""#", ".join(depto["list"]) if depto["total"]>0 else ""
                            , ""#, ", ".join(depon["list"]) if depon["total"]>0 else ""
                            #, item.metadata #returns None if empty
                            , item.access
                            , layernames
                            , layerurls
                            , yr1
                            , day30
                            , item.get_thumbnail_link()#string
                            , ("Broken" if len(errorlist)>0 else "")
                            ])
                            #del(depto)
                            #del(depon)
                            #Save changes to Excel
                            wb.save(newxl)
                            wb.close()
                        except Exception as e:
                            printaddmsg(str(e), "warn") #arcpy.AddError(e.args[0])
                            #If the Excel workbook didn't send, add to "FailedToSave"
                            FailedToSave.append(item.title or item.id or item.url or "No Title")
                            printaddmsg("Error...", "warn")
                    del(new)
                else:
                    printaddmsg("No items returned by the query " + itemtype  +  " - "  +  URLbase, "msg")
                    fw.writerow(["No items"])
                del(fw, f)
            if len(FailedToSave)>0:
                #Write out a list of major errors - where the new sheet didn't successfully get added to the workbook
                printaddmsg("These did not load:", "msg")
                printaddmsg(FailedToSave, "msg")
                with open(outdir + "/inventory_" + str(datetime.strftime(datetime.now(), "%Y-%m-%d")) + itemtype + "_" + URLbase + "Errors.txt", 'w', newline = '', encoding = "utf-8") as ferror:
                    fwerror = csv.writer(ferror, delimiter = "\t")
                    for fail in FailedToSave:
                        fwerror.writerow([fail])
                del(fwerror, ferror)

        
        ##Save changes to Excel
        #img = openpyxl.drawing.image.Image(os.path.dirname(__file__)+'/icon.png')
        #img.width = 120 #in pixels
        #img.anchor = 'B2'
        #toc.add_image(img)
        #try:
        #    wb.save(newxl)
        #except:
        #    wb.saveas(newxl.replace(".xlsx","_Recover.xlsx"))
        #del(wb, origws)

        #print("All done with " + itemtype + " " + URLbase)
        #print("All done with all itemtypes for " + URLbase)
        #print("Done")
        ##Could also add the thumbnail
        #https://stackoverflow.com/questions/42875353/insert-an-image-from-url-in-openpyxl
        return

def runSvcTool(template, outdir, importdtls, getportdtls, gis="", un="", pw="", portalURL=""):
        if gis == "":
            if un == "":
                if 'ArcGISPro.exe' in sys.executable:
                    gis = arcgis.gis.GIS("Pro")
            else:
                gis = getPortalURL(portalURL, un, pw)
        #Define the web applications to search for!
        ItemTypes = ['Service']
        URLbase = gis.properties["customBaseUrl"].replace("maps.arcgis.com", "AGOL")
        newxl = template.replace("Autopopulate", URLbase)
        newxl = outdir+"/"+os.path.basename(newxl)
        shutil.copy2(template, newxl)
        if importdtls not in ("#", "None", "", None, "Null"):
            PortalDtls = True
            #Add the table!!
            wb = openpyxl.load_workbook(newxl)
            ws = wb.create_sheet("PortalServiceDetails")
            with open(importdtls) as f:
                reader = csv.reader(f, delimiter = '\t')
                for row in reader:
                    ws.append(row)
                    #if "FeatureServer" in row[1]:
                    #    outfs = row
                    #    outfs[1] = outfs[1].replace("MapServer","FeatureServer")
                    #    ws.append(outfs)
            #Add to TOC
            toc = wb["Table Of Contents"]
            link = "#'PortalServiceDetails'!A1"
            toc.cell(row = 13, column = 2).value = "Portal Service Details"#Revised to show the title!
            toc.cell(row = 13, column = 2).hyperlink = (link)
            toc.cell(row = 13, column = 2).style = "Hyperlink"                            
            #Add table
            table = openpyxl.worksheet.table.Table(displayName = "tblSvc", ref = "A1:" + openpyxl.utils.get_column_letter(ws.max_column) + str(ws.max_row))
            style = openpyxl.worksheet.table.TableStyleInfo(name="TableStyleLight9", showFirstColumn=True,
                                   showLastColumn=False, showRowStripes=False, showColumnStripes=False)							
            table.tableStyleInfo = style
            ws.add_table(table)
            for ht in range(2,ws.max_row+1):
                ws.row_dimensions[ht].height = 75 #125
                for c in range(1,ws.max_column+1):
                    ws.cell(row = ht, column = c).alignment = openpyxl.styles.Alignment(wrapText = True)                
            ws.column_dimensions['A'].width = 20 #187
            ws.column_dimensions['B'].width = 22 #205
            ws.column_dimensions['C'].width = 20 #187
            ws.column_dimensions['F'].width = 12 #115
            ws.column_dimensions['G'].width = 12 #115
            ws.column_dimensions['I'].width = 78 #709
            ws.freeze_panes = "B2" #freeze row 1/column A            
            #for row in ws.iter_rows():
            #    for cell in row:
            #        cell.style.alignment.wrap_text = True
            ws.cell(row = 10, column = 1).value = "Press F5 (then Enter) to return to the previous active cell."
            try:
                wb.save(newxl)
            except:
                wb.saveas(newxl.replace(".xlsx","_Recover.xlsx"))
            printaddmsg("Completed.  Details of web services in Portal have been added to the workbook.")       
        print(str(getportdtls))
        if getportdtls not in ("#", "None", "", None, "Null"):
            PortalDtls = True
            svcinfo(outdir, port = gis, xlsx = newxl)

        wb = openpyxl.load_workbook(newxl)
        toc = wb["Table Of Contents"]
        toc["B8"] = "The Web Services contained in the spreadsheet are based on what exist within " + agency

        try:
            origws =  wb['Web Services Template']
        except Exception as e:
            printaddmsg(str(e), "err") #arcpy.AddError(e.args[0])
        i = 14 #Set the default row on the Table of Contents for listing links to new worksheets
        Worksheets = []
        FailedToSave = []

        for itemtype in ItemTypes:
            with open(outdir + "/inventory_" + str(datetime.strftime(datetime.now(), "%Y-%m-%d")) + itemtype + "_" + URLbase + ".txt", 'w', newline = '', encoding = "utf-8") as f:
                    #with open("C:/temp/inventory_2021-11-9" + itemtype + ".txt", 'w', newline = '', encoding = "utf-8") as f:
                    printaddmsg(itemtype, "msg")
                    printaddmsg(URLbase, "msg")
                    fw = csv.writer(f, delimiter = "\t")
                    fw.writerow(["Title", "ShortDescription", "ClientID", "Link", "DirectLink", 
                    "Description", "Tags", "Access", 
                    "Owner", "Created", "Modified", 
                    "Views", "Unknown", "Type", "SharedWith", 
                    "Credits", "Comments", "SpecialStatus", 
                    "DependentTo", "DependentUpon", #"Metadata", 
                    "Use", "LayerNames", "LayerURLs", "1YrUsage", "1MonthUsage", "Thumbnail", "Broken"])
                    #https://developers.arcgis.com/python/api-reference/arcgis.apps.storymap.html#journalstorymap
                    #   The API reference above only works for newer items.  Still unsure how to get
                    #   the web map for a story map without opening the map & scraping the appID or
                    #   opening "Configure" for the app, hence "selenium" package & dependencies
                    #https://developers.arcgis.com/python/api-reference/arcgis.apps.dashboard.html#details

                    items = gis.content.search(query = "NOT owner: esri_livingatlas", item_type = itemtype, max_items = 1000) #Cycle through app types
                    
                    for item in items:
                        printaddmsg(item.title or "No Title", "msg")
                        try:
                            #depto = item.dependent_to() #arcgis enterprise only
                            #depon = item.dependent_upon() #arcgis enterprise only
                            layernames = ""
                            layerurls = ""
                            try:
                                sharing = item.shared_with #dictionary in the following format: { ‘groups’: [], # one or more Group objects ‘everyone’: True | False, ‘org’: True | False }
                                if (len(sharing["groups"])>0):
                                    groups = [g.title for g in sharing["groups"]]
                                else:
                                    groups = []
                                shared = ", ".join([g.title for g in sharing["groups"]]) + (", Everyone" if sharing["everyone"] == True else "")  + (", Organization" if sharing["org"] == True else "")
                            except:
                                shared = ""
                            try:
                                if not URLbase == "portal":
                                    yr1 = sum(item.usage(date_range = '1Y', as_df = True).Usage) #don't convert to Pandas DF
                                    day30 = sum(item.usage(date_range = '30D', as_df = True).Usage) #don't convert to Pandas DF
                                else:
                                    yr1 = None
                                    day30 = None
                            except:
                                printaddmsg("  "+str(e))                               
                                yr1 = None
                                day30 = None                    
                            # enable browser logging
                            mapids = []
                            errorlist = []
                            print("Writing to worksheet")
                            newt = item.title.replace("/", "").replace("\\", "").replace("*", "").replace("[", "").replace("]", "").replace(":", "").replace("?", "")
                            newt = newt[0:min(30, len(newt))]
                            #If the there's already a sheet with this name, append a number
                            if (newt in Worksheets):
                                k = 1
                                while newt in Worksheets:
                                    #Don't exceed the maximum # of characters allowable for a tab name...
                                    newt = newt[0:min(27, len(newt))].replace("(" + str(k-1) + ")", "") + "(" + str(k) + ")"
                                    k +=  1
                                #print ("Did not add " + newt + " due to multiple duplicates.")
                                #continue
                            new = wb.copy_worksheet(origws)
                            new.active = 1
                            new.title = newt
                            Worksheets.append(newt) #Keep track of what worksheet names we've added to avoid duplicates going forward
                            new.cell(row = 1, column = 1).value = (item.title if item.title else "") #A1
                            new.cell(row = 2, column = 1).value = item.snippet #A2 - is this Subject, Svc Desc, or Desc?
                            #new.cell(row = 4, column = 3).value = (appinfo["client_id"] if len(appinfo)>0 else item.id) #C4
                            new.merge_cells(start_row = 7, start_column = 3, end_row = 9, end_column = 3)
                            try:
                                new.cell(row = 4, column = 3).value = url.strip('/').rsplit('/', 2)[-2]
                            except:
                                new.cell(row = 4, column = 3).value = (item.title if item.title else "")
                            new.cell(row = 5, column = 3).value = (item.url if item.url else "")#
                            if not item.url == None:
                                new.cell(row = 5, column = 3).hyperlink = item.url
                                new.cell(row = 5, column = 3).style = "Hyperlink"
                            new.cell(row = 6, column = 3).value = item.homepage.replace("/home//home", "/home") #C6
                            if not item.homepage == None:
                                new.cell(row = 6, column = 3).hyperlink = item.homepage
                                new.cell(row = 6, column = 3).style = "Hyperlink"
                            new.cell(row = 7, column = 3).value = (strip_tags(item.description) if not item.description == None else "") #C7
                            new.cell(row = 10, column = 3).value = ", ".join(item.tags if not item.tags is None else "") #C8
                            #new.cell(row = 11, column = 3).value = (strip_tags(item.licenseInfo) if not item.licenseInfo == None else "") 
                            new.cell(row = 13, column = 3).value = ("public" if "veryone" in shared else "internal")#'Group Name' 
                            #new.cell(row = 13, column = 3).value = item.WHICH ESRI VERSION #C10
                            new.cell(row = 15, column = 3).value = datetime.fromtimestamp(item.created/1000).strftime("%m/%d/%Y")##Date/time was in epoch time to the millisecond so /1000 #C11
                            new.cell(row = 16, column = 3).value = datetime.fromtimestamp(item.modified/1000).strftime("%m/%d/%Y") 
                            new.cell(row = 18, column = 3).value = item.numViews #C14
                            #new.cell(row = 17, column = 3).value = ("Geocortex" if "ortex" in (item.url or "") else "")#Technology 
                            #if "MapSeries" in (item.url or "") or "MapJournal" in (item.url or "") or "AGOStoryMaps" in (item.url or "") or "MapTour" in (item.url or "") or itemtype  ==  'StoryMap':
                            #    new.cell(row = 18, column = 3).value = "Story Map"#string #C16
                            #else:
                            #    new.cell(row = 18, column = 3).value = item.type#string #C16
                            #    #Or should this be "itemtype" instead of "item.type"?  Not sure which is more specific...
                            #new.cell(row = 21, column = 3).value = item.homepage.replace("/home//home", "/home") #'' #C18
                            #new.cell(row = 22, column = 3).value = item.access #C20
                            new.cell(row = 24, column = 3).value = ", ".join(groups)##C24
                            new.cell(row = 26, column = 3).value = item.owner
                            new.cell(row = 27, column = 3).value = ", ".join(item.typeKeywords)
                            if PortalDtls == True:
                                new.cell(row = 5, column = 7).value = "Service Details"
                            print("Starting to find data layers")
                            #Then loop through the data layers, if available
                            j = 6
                            if (not 'Service Definition' in item.typeKeywords):
                                try:
                                    layers = item.layers
                                    layernames = ", ".join([layer.properties.name or "" for layer in layers])
                                    layerurls = ", ".join([layer.url or "" for layer in layers])
                                    for layer in layers:
                                        new.cell(row = j, column = 5).value = layer.url or ""
                                        if not layer.url == None:
                                            new.cell(row = j, column = 5).hyperlink = layer.url
                                            #Need to update this script to include creating URLs to PortalServiceDetails
                                            new.cell(row = j, column = 5).style = "Hyperlink"
                                            if PortalDtls == True and URLbase == "portal":
                                                new.cell(row = j, column = 7).value = '=HYPERLINK(CELL("address",INDEX(tblSvc[URL],MATCH(LEFT(E'+str(j)+\
                                                    ',IFERROR(FIND("FeatureServer",E'+str(j)+')+12,IFERROR(FIND("MapServer",E'+str(j)+')+8,LEN(E'+str(j)+')))),tblSvc[URL],0))), "LayerDetails")'
                                                #new.cell(row = j, column = 7).hyperlink = ('=HYPERLINK(CELL("address",INDEX(tblSvc[URL],MATCH(LEFT(E'+str(j)+\
                                                #        ',IFERROR(FIND("FeatureServer",E'+str(j)+')+12,LEN(E'+str(j)+'))),tblSvc[URL])))')
                                                new.cell(row = j, column = 7).style = "Hyperlink"
                                        new.cell(row = j, column = 6).value = layer.properties.name or ""
                                        j +=  1    
                                    #If portal, we can
                                    #       use print("\n\n Item: \t {2} \t\t\t id {0} is dependent on these items: \t {1}".format(item.itemid, item.dependent_upon(), item.title))
                                    #       
                                    #if (item.dependent_to()['total'] > 0):
                                    #   print("\t This item is also a dependency of these items: {}".format(item.dependent_to()))
                                    
                                except:
                                    errorlist.append("Failed to retrieve layers")
                                if len(errorlist)>0:
                                        for err in errorlist:
                                            new.cell(row = j, column = 5).value = (err or "")
                                            new.cell(row = j, column = 6).value = "Error"
                                            j +=  1
                            del(new)
                            #if item.type  ==  "Web Mapping Application", item.related_items("")
                            #use relationships or dependencies, not sure which https://developers.arcgis.com/rest/users-groups-and-items/relationship-types.htm
                            #etc.
                            mapids = []
                            toc = wb["Table Of Contents"] #UPDATE THIS!!!
                            #toc.insert_rows(i) #This part doesn't work due to package limitations
                            #If there are any links below, those links aren't moved with the text!
                            #So we'll just make a list in another section, NOT insert rows, 
                                #then move the data over.
                            #We'll have to undo the merge-across, then (if required) re-merge
                            #   across b/c otherwise copy/paste is not possible.
                            link = "#'" + newt + "'!A1"
                            toc.cell(row = i, column = 2).value = item.title#Revised to show the title!
                            toc.cell(row = i, column = 2).hyperlink = (link)
                            toc.cell(row = i, column = 2).style = "Hyperlink"
                            toc.cell(row = i, column = 3).value = ("No Layer URLs" if len(errorlist)>0 else "")
                            i +=  1
                            print("writing")
                            fw.writerow([
                            item.title#string
                            , item.snippet#string
                            , "N/A"
                            , (item.homepage.replace("/home//home", "/home") if not item.homepage == None else "")#(appinfo["client_id"] if len(appinfo)>0 else item.id) #" #dictionary of info for registered apps - irrelevant here, I would think..
                            , (item.url if item.url else item.homepage.replace("/home//home", "/home"))
                            , (strip_tags(item.description) if not item.description == None else "")
                            , ", ".join(item.tags if not item.tags is None else "")
                            , (strip_tags(item.licenseInfo) if not item.licenseInfo == None else "")
                            , item.owner#string
                            , datetime.fromtimestamp(item.created/1000).strftime("%m/%d/%Y")##Date/time was in epoch time to the millisecond so /1000
                            , datetime.fromtimestamp(item.modified/1000).strftime("%m/%d/%Y")
                            , item.numViews
                            , ""
                            , item.type#string
                            , shared
                            , item.accessInformation
                            , ""#, ", ".join([item.comments[i]["comment"] for i in range(len(item.comments))]).replace("%20", " ")
                            , item.content_status #deprecated or authoritative or None
                            , ""#", ".join(depto["list"]) if depto["total"]>0 else ""
                            , ""#, ", ".join(depon["list"]) if depon["total"]>0 else ""
                            #, item.metadata #returns None if empty
                            , item.access
                            , layernames
                            , layerurls
                            , yr1
                            , day30
                            , item.get_thumbnail_link()#string
                            , ("Broken" if len(errorlist)>0 else "")
                            ])
                            #del(depto)
                            #del(depon)
                            #print("Saving")
                            wb.save(newxl)
                        except:
                            #printaddmsg(str(e), "warn") #arcpy.AddError(e.args[0])                        
                            FailedToSave.append(item.title or item.homepage or item.url or "No Title")
                            wb.save(newxl)


        #img = openpyxl.drawing.image.Image(os.path.dirname(__file__)+"/icon.png")
        #img.width = 120 #in pixels
        #img.anchor = 'B2'
        #toc.add_image(img)

        try:
            wb.save(newxl)
        except:
            wb.saveas(newxl.replace(".xlsx","_Recover.xlsx"))

        del(wb, origws, fw, f)

        if (len(FailedToSave)>0):
            printaddmsg("These did not load:", "msg")
            printaddmsg(FailedToSave, "msg")
            with open(outdir + "/inventory_" + str(datetime.strftime(datetime.now(), "%Y-%m-%d")) + itemtype + "_" + URLbase + "Errors.txt", 'w', newline = '', encoding = "utf-8") as ferror:
                fwerror = csv.writer(ferror, delimiter = "\t")
                for fail in FailedToSave:
                    fwerror.writerow([fail])
            del(fwerror, ferror)
            print("All done.")

        ##Could also add the thumbnail
        #https://stackoverflow.com/questions/42875353/insert-an-image-from-url-in-openpyxl


class AppTool(object):

    def __init__(self):
        """Define the tool (tool name is the name of the class)."""
        self.label = "Web App Inventory"
        self.description = "This script tool returns an inventory of story maps, dashboards, and 'web mapping applications' in the active ArcGIS Pro portal or AGOL"
        self.canRunInBackground = False

    def getParameterInfo(self):
        """Define parameter definitions"""
        #TemplateSpreadsheet = arcpy.GetParameterAsText(0)
        #outdir = arcpy.GetParameterAsText(1)#C:/temp/
        TemplateSpreadsheet = arcpy.Parameter(
            displayName = "Template Spreadsheet for 'Pretty' Output", 
            name = "TemplateSpreadsheet", 
            datatype = "DEFile", #"GPDataFile", 
            parameterType = "Required", 
            direction = "Input")
        outdir = arcpy.Parameter(
            displayName = "Output Folder for Text Files", 
            name = "outdir", 
            datatype = "DEFolder", #"GPDataFile", 
            parameterType = "Required", 
            direction = "Input")
        ImportPortalDetails = arcpy.Parameter(
            displayName = "Existing Portal Service Table", 
            name = "ImportPortalDetails", 
            datatype = "DEFile", 
            parameterType = "Optional", 
            direction = "Input")
        GetPortalDetails = arcpy.Parameter(
            displayName = "Create New Portal Service Table and Load Data", 
            name = "GetPortalDetails", 
            datatype = "Boolean", 
            parameterType = "Optional", 
            direction = "Input")
        TemplateSpreadsheet.filter.list = ['xlsx']
        ImportPortalDetails.filter.list = ['txt']
        #   Add the "port" as a hidden parameter that can be specified in the script but otherwise defaults to portDefault??  The parameter type wouldn't be text, then; not sure how that would work...
        #Need to update this script to include creating URLs to PortalServiceDetails
        TemplateSpreadsheet.value = os.path.dirname(__file__)+"/WebApps_Autopopulate.xlsx"
        params = [TemplateSpreadsheet, outdir, ImportPortalDetails, GetPortalDetails]
        return params

    def isLicensed(self):
        """Set whether tool is licensed to execute."""
        return True

    def updateParameters(self, parameters):
        """Modify the values and properties of parameters before internal
        validation is performed.  This method is called whenever a parameter
        has been changed."""
        #Need to make only GetPortalDetails or ImportPortalDetails available, not both.
        if parameters[2].valueAsText:
            if not parameters[3].altered:
                parameters[3].enabled = False
        if parameters[3].valueAsText:
            if not parameters[2].altered:
                parameters[2].enabled = False
        return
    
    def updateMessages(self, parameters):
        """Modify the messages created by internal validation for each tool
        parameter.  This method is called after internal validation."""
        return

    def execute(self, parameters, messages):
        """The source code of the tool."""
        
        #Set the template workbook (load from the parameter list)
        template = parameters[0].valueAsText
        outdir = parameters[1].valueAsText 
        importdtls = parameters[2].valueAsText 
        getportdtls = parameters[3].valueAsText 
        runAppTool(template,outdir,importdtls,getportdtls)    


class SvcTool(object):


#    if port == "":
#        port = arcgis.gis.GIS(portalurl, username, password)#
#
#    mgr = arcgis.gis.admin.PortalAdminManager(port.url+"//sharing/rest", gis = port)



    def __init__(self):
        """Define the tool (tool name is the name of the class)."""
        self.label = "Web Service Inventory"
        self.description = "This script tool returns an inventory of web services in the active ArcGIS Pro portal or AGOL"
        self.canRunInBackground = False

    def getParameterInfo(self):
        """Define parameter definitions"""
        #TemplateSpreadsheet = arcpy.GetParameterAsText(0)
        #outdir = arcpy.GetParameterAsText(1)#C:/temp/
        TemplateSpreadsheet = arcpy.Parameter(
            displayName = "Template Spreadsheet for 'Pretty' Output", 
            name = "TemplateSpreadsheet", 
            datatype = "DEFile", #"GPDataFile", 
            parameterType = "Required", 
            direction = "Input")
        outdir = arcpy.Parameter(
            displayName = "Output Folder for Text Files", 
            name = "outdir", 
            datatype = "DEFolder", #"GPDataFile", 
            parameterType = "Required", 
            direction = "Input")
        ImportPortalDetails = arcpy.Parameter(
            displayName = "Existing Portal Service Table", 
            name = "ImportPortalDetails", 
            datatype = "DEFile", 
            parameterType = "Optional", 
            direction = "Input")
        GetPortalDetails = arcpy.Parameter(
            displayName = "Create New Portal Service Table and Load Data", 
            name = "GetPortalDetails", 
            datatype = "Boolean", 
            parameterType = "Optional", 
            direction = "Input")
        TemplateSpreadsheet.filter.list = ['xlsx']
        ImportPortalDetails.filter.list = ['txt']
        TemplateSpreadsheet.value = os.path.dirname(__file__)+"/WebServices_Autopopulate.xlsx"
        params = [TemplateSpreadsheet, outdir, ImportPortalDetails, GetPortalDetails]
        return params

    def isLicensed(self):
        """Set whether tool is licensed to execute."""
        return True

    def updateParameters(self, parameters):
        """Modify the values and properties of parameters before internal
        validation is performed.  This method is called whenever a parameter
        has been changed."""
        if parameters[2].valueAsText:
            if not parameters[3].altered:
                parameters[3].enabled = False
        if parameters[3].valueAsText:
            if not parameters[2].altered:
                parameters[2].enabled = False
        #if 'ArcGISPro.exe' in sys.executable:
        #    parameters[4].enabled = False
        #    parameters[4].value = ""
        return

    def updateMessages(self, parameters):
        """Modify the messages created by internal validation for each tool
        parameter.  This method is called after internal validation."""
        return

    def execute(self, parameters, messages):
        """The source code of the tool."""

        template = parameters[0].valueAsText
        outdir = parameters[1].valueAsText 
        importdtls = parameters[2].valueAsText 
        getportdtls = parameters[3].valueAsText 
        runSvcTool(template, outdir, importdtls, getportdtls)
        ## This is a brief intro to how to access ArcGIS contents w/ python
        # gis = GIS(url = 'https://pythonapi.playground.esri.com/portal', username = 'arcgis_python', password = 'amazing_arcgis_123')
        # search_my_contents = gis.content.search(query = "owner:[owner name]", item_type = "Web Map", max_items = 1000)
        # #create web map objects from search results and print the web map title and layer name
        # for webmap_item in search_my_contents:    
        #   webmap_obj = WebMap(webmap_item)
        #   for layer in webmap_obj:
        #         print(webmap_item.title, layer.title)‍‍‍‍‍‍‍‍‍‍‍‍‍‍‍

class ScrapeRestEndpoints(object):
    def __init__(self):
        """Define the tool (tool name is the name of the class)."""
        self.label = "Portal Services"
        self.description = "This script tool returns an inventory of services registered to the active ArcGIS Pro portal."
        self.canRunInBackground = False

    def getParameterInfo(self):
        """Define parameter definitions"""
        
        inputURL = arcpy.Parameter(
            displayName = "Enter Portal URL", 
            name = "portalurl", 
            datatype = 'GPString', 
            direction = 'Input', 
            parameterType = 'Required', 
        )
        inputURL.value = 'https://gis.nevadadot.com/agsportal/home'
        
        inputusername = arcpy.Parameter(
            displayName = "Enter Username", 
            name = 'username', 
            datatype = 'GPString', 
            direction = 'Input', 
            parameterType = 'Required', 
        )
        
        inputpassword = arcpy.Parameter(
            displayName = "Enter Password", 
            name = "password", 
            datatype = 'GPStringHidden', 
            direction = 'Input', 
            parameterType = 'Required', 
        )
        
        outloc = arcpy.Parameter(
            displayName = "Select a folder for the output tab-delimited .txt file", 
            name = "outloc", 
            datatype = "DEFolder", 
            direction = 'Input', 
            parameterType = 'Required', 
        )
        
        params = [inputURL, inputusername, inputpassword, outloc]
        return params

    def isLicensed(self):
        """Set whether tool is licensed to execute."""
        return True

    def updateParameters(self, parameters):
        """Modify the values and properties of parameters before internal
        validation is performed.  This method is called whenever a parameter
        has been changed."""
        return

    def updateMessages(self, parameters):
        """Modify the messages created by internal validation for each tool
        parameter.  This method is called after internal validation."""
        return

    def execute(self, parameters, messages):
        svcinfo(portalurl = parameters[0].valueAsText, username = parameters[1].valueAsText, 
            password = parameters[2].valueAsText, outtabloc = parameters[3].valueAsText)

        return




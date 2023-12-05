"""-----------------------------------------------------------------------------
  Script Name: Generate Overlay
  Description: Copies Fields from Feature Classes in a database based on an Index Table then overlays the copied Feature Classes the final output flat file. 
  Requirement: ArcGIS 10.1 or later is required to run this tool.
-----------------------------------------------------------------------------"""
import arcpy, os, sys
from datetime import datetime


"""CORE FUNCTIONS"""
# Check to make sure all the FCs and FIELDs in the Index Table are valid
def checkIndex(RH_SDE, indexTbl, RouteID, FromMM, ToMM, FromDate, ToDate):
	arcpy.AddMessage("{0}: Checking Index Table with RH_SDE".format(timeStamp()))	 
	
	# List the Feature Classes in the SDE	
	arcpy.AddMessage("{0}: Setting Workspace to RH_SDE ... ".format(timeStamp()))
	arcpy.env.workspace = RH_SDE
	
	arcpy.AddMessage("{0}: Getting Feature Class List from RH_SDE ... ".format(timeStamp()))
	fcList = arcpy.ListFeatureClasses()
	
	fcStrList = []
	for fc in fcList:
		fcSplit = fc.split(".")
		fcStr = fcSplit[-1]
		fcStrList.append(fcStr)
		
	# Iterate through Index Table
	arcpy.AddMessage("{0}: Comparing Index Table to Feature Class List ... ".format(timeStamp()))
	error = 0
	
	for fc in fcList:
		# Trim SDE Database from the IN_FC name
		fcSplit = fc.split(".")
		fcStr = fcSplit[-1]
		
		fieldList = arcpy.ListFields(fc)
		fieldNames = []
		for field in fieldList:
			fieldNames.append(field.name) 
		
		with arcpy.da.SearchCursor(indexTbl,['IN_FC','IN_FLD']) as cursor:
			for row in cursor:
				if row[0] not in fcStrList:
					arcpy.AddMessage("{0}: ERROR: {1} Feature Class from Index Table NOT FOUND in RH_SDE".format(timeStamp(),row[0]))
					error = error + 1
					return error
					
				elif row[0] == fcStr:
					# arcpy.AddMessage(" ... Checking IN_FC: {0} for Field: {1} ... ".format(row[0], row[1]))
					if RouteID not in fieldNames: 
						arcpy.AddMessage("{0}: ERROR: {1} Field from Index Table NOT FOUND in {2} Feature Class of RH_SDE".format(timeStamp(),RouteID,row[0]))
						error = error + 1
					if FromMM not in fieldNames:
						arcpy.AddMessage("{0}: ERROR: {1} Field from Index Table NOT FOUND in {2} Feature Class of RH_SDE".format(timeStamp(),FromMM,row[0]))
						error = error + 1
					if ToMM not in fieldNames:
						arcpy.AddMessage("{0}: ERROR: {1} Field from Index Table NOT FOUND in {2} Feature Class of RH_SDE".format(timeStamp(),ToMM,row[0]))
						error = error + 1
					if FromDate not in fieldNames:
						arcpy.AddMessage("{0}: ERROR: {1} Field from Index Table NOT FOUND in {2} Feature Class of RH_SDE".format(timeStamp(),FromDate,row[0]))
						error = error + 1
					if ToDate not in fieldNames:
						arcpy.AddMessage("{0}: ERROR: {1} Field from Index Table NOT FOUND in {2} Feature Class of RH_SDE".format(timeStamp(),ToDate,row[0]))
						error = error + 1	
					if row[1] not in fieldNames:
						arcpy.AddMessage("{0}: ERROR: {1} Field from Index Table NOT FOUND in {2} Feature Class of RH_SDE".format(timeStamp(),row[1],row[0]))
						error = error + 1
				
	return error
	
# Create the Staging File GDB in the output folder with a Date and Time Stamp
def createStagingGDB(workspace, name):
	# Time stamping and creating the Staging File Geodatabase
	# dt = datetime.now()
	# dtStr = '{:%Y%m%d_%H%M%S}'.format(dt)
	gdbStr = "{0}_{1}.gdb".format(name, dateTime())	
	
	arcpy.CreateFileGDB_management(workspace, gdbStr, "CURRENT")
	arcpy.AddMessage("{0}: Created {1} in the Output Folder Location".format(timeStamp(),gdbStr, workspace))
	
	# Returning the name of the Staging File Geodatabase from function
	gdbPath = os.path.join(workspace,gdbStr)
	
	return gdbPath
	
# Copy IN_FCs listed in the Index Table to the Staging File Geodatabase ... Only Fields listed in Index Table are Field Mapped to new Tables and the Fields in the new FCs are added based on their order in the Index Table 
def copyTables(gdbPath, RH_SDE, indexTbl, SQL_Query, RouteID, FromMM, ToMM, FromDate, ToDate, decPlaces):
	
	arcpy.AddMessage("SQL Query for copying Events is:\n{0}".format(SQL_Query))
	
	# List the Feature Classes in the SDE
	arcpy.env.workspace = RH_SDE
	fcList = arcpy.ListFeatureClasses()
	for fc in fcList:
		# Set Counter, Field Mappings, and Field Map Dictionary for Table to Table conversion
		count = 0
		fms = arcpy.FieldMappings()
		fmDict = {}
		
		# Trim SDE Database from the IN_FC name
		fcSplit = fc.split(".")
		fcStr = fcSplit[-1]
		
		# Iterate through every cell of the 'IN_FC' and 'IN_FLD' columns of the Index Table
		with arcpy.da.SearchCursor(indexTbl,['IN_FC','IN_FLD','OUT_FLD_NAME']) as cursor:
			for row in cursor:
				
				# Compare the current FC with each row of the 'IN_FC' column in the Index Table
				if row[0] == fcStr:
					
					# Table to Table Conversion will occur for this FC
					if count == 0:
						dissFields = row[2]
					elif count > 0:
						dissFields = dissFields + ";" + row[2]
					count = count + 1
					
					# Dynamically create Field Map variable in the Field Map Dictionary using Counter
					fmDict["fm_"+str(count)]=arcpy.FieldMap()
					# Add current FC and Field Name as an input field to the Dynamic Field Map variable
					fmDict["fm_"+str(count)].addInputField(fc, row[1])
					# Add output field name to Dynamic Field Map variable
					fmDict["fm_out_"+str(count)] = fmDict["fm_"+str(count)].outputField
					fmDict["fm_out_"+str(count)].name = row[2]
					fmDict["fm_out_"+str(count)].aliasName = row[2]
					fmDict["fm_"+str(count)].outputField = fmDict["fm_out_"+str(count)]
					# Add the Dynamic Field Map variable to Field Mappings
					fms.addFieldMap(fmDict["fm_"+str(count)])
					
		del row, cursor # Stop iterating through Index Table
		
		# The Feature Class needs to be copied to the Staging GDB
		if count > 0:
			# Field Mappings RouteID, FromMM, ToMM, FromDate, ToDate
			try:
				fm_RouteID = arcpy.FieldMap()
				fm_RouteID.addInputField(fc, RouteID)
				fms.addFieldMap(fm_RouteID)
				fm_FromMM = arcpy.FieldMap()
				fm_FromMM.addInputField(fc, FromMM)
				fms.addFieldMap(fm_FromMM)
				fm_ToMM = arcpy.FieldMap()
				fm_ToMM.addInputField(fc, ToMM)
				fms.addFieldMap(fm_ToMM)
				# fm_FromDate = arcpy.FieldMap()
				# fm_FromDate.addInputField(fc, FromDate)
				# fms.addFieldMap(fm_FromDate)
				# fm_ToDate = arcpy.FieldMap()
				# fm_ToDate.addInputField(fc,ToDate)
				# fms.addFieldMap(fm_ToDate)
			except:
				arcpy.AddMessage("{0}: ERROR: Field Mapping RouteID, FromMeasure, ToMeasure, FromDate or ToDate not successful.".format(timeStamp()))
				return 1
			
			arcpy.AddMessage("{0}: {1} Field Map Successful ... ".format(timeStamp(),fcStr))
			
			# Table to Table Conversion
			try:
				eventStr = "Event_{0}".format(fcStr)
				eventPath = os.path.join(gdbPath,eventStr)
				
				arcpy.AddMessage(" ... Copying {0} Feature Class to Staging Geodatabase ... ".format(fcStr))
				arcpy.TableToTable_conversion(os.path.join(RH_SDE,fc), gdbPath, eventStr, SQL_Query, field_mapping=fms)
			except:
				arcpy.AddMessage("{0}: ERROR: SQL Statement not satisfactory".format(timeStamp()))
				return 1
			
			# Replace Nulls in the Copied Event Table
			arcpy.AddMessage(" ... Removing Nulls from {0} Event Table ... ".format(fcStr))
			replaceNulls(gdbPath, eventStr, FromMM)
			
			# Round From and To Measures to 4 Digits in Copied Event Table
			decPlace = int(decPlaces) + 1
			arcpy.AddMessage(" ... Rounding From and To Measure to {0} decimal places in {1} Event Table ...".format(str(decPlaces), fcStr))
			roundMeasures(gdbPath, eventStr, FromMM, ToMM, decPlace)
			
			# Dissolve Copied Event Table based on field list variable dissFields made in previous nested cursor
			arcpy.AddMessage(" ... Dissolving {0} Event Table on {1} Field(s) ... \n".format(fcStr, dissFields))
			dissStr = "Dissolve_{0}".format(fcStr)
			dissPath = os.path.join(gdbPath,dissStr)
			routeParams = "{0} LINE {1} {2}".format(RouteID,FromMM,ToMM)
			
			arcpy.DissolveRouteEvents_lr(eventPath,	routeParams, dissFields, dissPath, routeParams, "DISSOLVE", "INDEX")
			
	# End FC Iteration
	
	return 0
	
# Overlay the Tables in a Specific Order
def overlayTables(gdbPath, indexTbl, RouteID, FromMM, ToMM):
	arcpy.AddMessage("{0}: Overlaying Tables based on Index Column of Inputs Table ...".format(timeStamp()))
	
	"""If error here then issue with Field List in Index Table"""
	# Set environmental workspace to Staging GDB
	arcpy.env.workspace = gdbPath
	
	# Getting Overlay Order based on Index Field
	sortTbl = sortingIndex(indexTbl)
	
	# Use 'IN_FC' field order in the Sorted Table to order iterative overlays
	cursor = arcpy.SearchCursor(sortTbl,['IN_FC'])
	count = 0
	routeParams = "{0} LINE {1} {2}".format(RouteID,FromMM,ToMM)
	for row in cursor:
		fcStr = row.getValue("IN_FC")
		dissStr = "Dissolve_{0}".format(fcStr)
		dissPath = os.path.join(gdbPath,dissStr)
		
		# Setting the first FC in the Sorted Table to be used in the FIRST OVERLAY
		if count == 0:
			dissStr_1 = dissStr
			dissPath_1 = os.path.join(gdbPath,dissStr_1)
			dissStr_1 = fcStr
			overlayStr = dissStr
		
		### FIRST OVERLAY
		elif count == 1:
			arcpy.AddMessage("{0}: Overlaying {1} with {2} ... ".format(timeStamp(),dissStr_1, dissStr))
			
			# Setting the output of the first overlay
			overlayStr = "Overlay{0}_{1}_{2}".format(count,dissStr_1,fcStr)
			overlayPath = os.path.join(gdbPath,overlayStr)
			
			# Use the first FC in the Input Table with the second FC of the Input Table
			arcpy.OverlayRouteEvents_lr(dissPath_1, routeParams, dissPath, routeParams, "UNION",
				overlayPath, routeParams, "NO_ZERO", "FIELDS", "INDEX")
			
		### SUBSEQUENT OVERLAYS
		elif count > 1:
			arcpy.AddMessage("{0}: Overlaying {1} with {2} ... ".format(timeStamp(),overlayStr, dissStr))
			
			# Use the previous Overlay combined with the current Dissolve Table
			arcpy.OverlayRouteEvents_lr(overlayPath, routeParams, dissPath, routeParams, "UNION",
				"Overlay{0}_{1}".format(count,fcStr),routeParams,"NO_ZERO","FIELDS","INDEX")
			
			# Setting the output Overlay FC for the next iteration
			overlayStr = "Overlay{0}_{1}".format(count,fcStr)
			overlayPath = os.path.join(gdbPath,overlayStr)
		
		# Setting count for next iteration
		count = count + 1
	del row, cursor
		
	return overlayStr
	
# Dissolving Final output and adding Length Field for Overlay Output
def formatSTL(tbl_ov, indexTbl, gdbPath, RouteID, FromMM, ToMM, decPlaces, routes, outFldr, temporalFilter):
	arcpy.AddMessage("\n{0}: Formatting Overlay Output Table ... ".format(timeStamp()))
	
	arcpy.env.workspace = gdbPath
	
	# Rounding Overlay Table to (num) decimal places & replace NULL values in all fields
	arcpy.AddMessage(" ... Rounding final {0} Table to {1} decimal places ... ".format(tbl_ov, decPlaces))
	roundMeasures(gdbPath, tbl_ov, FromMM, ToMM, int(decPlaces))
	
	# Getting Dissolve Field names from 'IN_FLD' column in Index Table & Dissolving Final Overlay Table
	tbl_diss = "Overlay_Final_Dissolved"
	arcpy.AddMessage(" ... Dissolving {0} Table to create {1} Table ... ".format(tbl_ov, tbl_diss))
	dissFields = concactFields(indexTbl, ";", "OUT_FLD_NAME")
	
	
	"""if error then issue with Field List in the input table"""
	# Performing Final Dissolve
	routeParams = "{0} LINE {1} {2}".format(RouteID,FromMM,ToMM)
	overlayPath = os.path.join(gdbPath,tbl_ov)
	
	arcpy.DissolveRouteEvents_lr(overlayPath, routeParams, dissFields, tbl_diss, routeParams, "DISSOLVE", "INDEX")
	
	# Create & Calculation LengthCalc Field
	arcpy.AddMessage(" ... Calculating Length Fields for dissolve Table ... ")
	dissPath = os.path.join(gdbPath, tbl_diss)
	
	arcpy.AddField_management(dissPath, "LengthCalc", "Double")
	arcpy.CalculateField_management(dissPath, "LengthCalc", "!{0}! - !{1}!".format(ToMM,FromMM), "PYTHON_9.3")		
	
	# Export final Overlay table where the length is greater than 0 (removing slivers)
	arcpy.AddMessage(" ... Exporting Final Overlay Table ... ")
	tbl_stl = "OVERLAY_OUTPUT_TBL_{0}".format(dateTime())
	arcpy.TableToTable_conversion(dissPath, gdbPath, tbl_stl, '"LengthCalc" > .001')
	stlPath = os.path.join(gdbPath, tbl_stl)
	
	arcpy.AddMessage(" ... Displaying Events on Routes and Exporting to Feature Class ... ")
	stl_lyr = "OVERLAY_OUTPUT_{0}".format(dateTime())
	stl_fc = "OVERLAY_OUTPUT_FC_{0}".format(dateTime())
	arcpy.MakeFeatureLayer_management(routes,"routes_filter",temporalFilter)
	arcpy.MakeRouteEventLayer_lr("routes_filter", RouteID, stlPath, routeParams, stl_lyr)
	arcpy.FeatureClassToFeatureClass_conversion(stl_lyr, gdbPath, stl_fc)
	
#	arcpy.FeatureClassToShapefile_conversion(stl_lyr, outFldr)
	
	return
	
	
	
"""ASSISTING FUNCTIONS"""
# Rounding From and To Measures from input table to specified number of digits
def roundMeasures(workspace, inTbl, FromMM, ToMM, digits):
	tblPath = os.path.join(workspace, inTbl)
	arcpy.CalculateField_management(tblPath, "{0}".format(FromMM), "round( !{0}!,{1})".format(FromMM,digits), "PYTHON_9.3")
	arcpy.CalculateField_management(tblPath, "{0}".format(ToMM), "round( !{0}!,{1})".format(ToMM,digits), "PYTHON_9.3")
	return

# Date Time String used for naming temp files
def dateTime():
	dt = datetime.now()
	dtStr = '{:%Y%m%d_%H%M%S}'.format(dt)
	return dtStr

# Time Stamp for use in comments
def timeStamp():
	time = datetime.now()
	timeStr = '{:%X}'.format(time)	
	return timeStr
	
# Averaging the FC_OV_ORDER field based on the 'IN_FC' column of the Index Table and sorting by the new 'MEAN_FC_OV_ORDER' field
def sortingIndex(indexTbl):
	meanTbl = "meanTable"
	arcpy.Statistics_analysis(indexTbl, meanTbl, [["FC_OV_ORDER", "MEAN"]], "IN_FC")
	
	sortTbl = "Overlay_Order"
	arcpy.Sort_management(meanTbl, sortTbl, 'MEAN_FC_OV_ORDER')
	
	arcpy.Delete_management(meanTbl)
	return sortTbl

# Concatenating Field List from Index Table	
def concactFields(indexTbl, delimiter, field):
	with arcpy.da.SearchCursor(indexTbl,[field]) as cursor:
		count = 0
		for row in cursor:
			if count == 0:
				concactFields = row[0]
			elif count > 0:
				concactFields = concactFields + "{0}".format(delimiter) + row[0]
			count = count + 1
	del cursor, row
	
	return concactFields

# Replace Nulls
def replaceNulls(path, fcStr, FromMM):
	tbl = os.path.join(path,fcStr)
	fldList = arcpy.ListFields(tbl)
	for field in fldList:
		with arcpy.da.UpdateCursor(tbl, [field.name]) as cursor:
			# Replace Nulls in String Domain Field w/ "."
			if field.type == 'String':
				for row in cursor:
					if row[0] == None:
						row[0] = "."
						cursor.updateRow(row)
			#Replace Nulls in Double Fields w/ 99999
			elif field.type in ('Double','Short','Long','Integer','SmallInteger') and field.name != FromMM:
				for row in cursor:
					if row[0] == None:
						row[0] = -999
						cursor.updateRow(row)
					elif row[0] == 0:
						row[0] = -888
						cursor.updateRow(row)
			# Do not replace Nulls in other fields
			else:
				for row in cursor:
					if row[0] == None:
						row[0] = 0
						cursor.updateRow(row)
						
	return
	

"""TOOLBOX CLASSES AND EXECUTING FUNCTION"""
class Toolbox(object):
    def __init__(self):
        self.label = "GenerateOverlay"
        self.alias = "Generates Overlay Table as a flat file."

        # List of tool classes associated with this toolbox
        self.tools = [GenerateOverlay]

class GenerateOverlay(object):
    def __init__(self):
        self.label = "GenerateOverlay"
        self.description = "Generate Overlay Table from input table"
        self.canRunInBackground = False
        self.showCommandWindow = False
        self.stylesheet = None
        self.workspace = None
		
    def getParameterInfo(self):
	
		indexTbl = arcpy.Parameter(
			displayName="Index Table (Input FC and Field List)",
			name="indexTbl",
			datatype="DETable",
			parameterType="Required",
			direction="Input")

		RH_SDE = arcpy.Parameter(
			displayName="RH Database",
			name="RH_SDE",
			datatype="DEWorkspace",
			parameterType="Required",
			direction="Input")
			
		SQL_Query = arcpy.Parameter(
			displayName="SQL Expression for events in Index Table",
			name="SQL_Query",
			datatype="GPSQLExpression",
			parameterType="optional",
			direction="Input")
			# TEST QUERY: (RouteID = '999_LA 1_1_1_010' OR RouteID = '999_US 71_1_1_010')
			
		routes = arcpy.Parameter(
			displayName="Routes Feature Class",
			name="routes",
			datatype="DEFeatureClass",
			parameterType="Required",
			direction="Input")
			
		# booleanTemporal = arcpy.Parameter(
			# displayName="Filter on date?",
			# name="booleanTemporal",
			# datatype="GPBoolean",
			# parameterType="Required",
			# direction="Input")
		
		temporalTarget = arcpy.Parameter(
			displayName="Date of the Overlay",
			name="temporalTarget",
			datatype="GPDate",
			parameterType="Optional",
			direction="Input")
		temporalTarget.value = datetime.utcnow()
		
		FromDate = arcpy.Parameter(
			displayName="From Date Field Name",
			name="FromDate",
			datatype="Field",
			parameterType="Required",
			direction="Input")
		FromDate.parameterDependencies = [routes.name]
		FromDate.filter.list = ['Date']
		
		ToDate = arcpy.Parameter(
			displayName="To Date Field Name",
			name="ToDate",
			datatype="Field",
			parameterType="Required",
			direction="Input")
		ToDate.parameterDependencies = [routes.name]
		ToDate.filter.list = ['Date']
		
		RouteID = arcpy.Parameter(
			displayName="Route ID Field Name",
			name="RouteID",
			datatype="Field",
			parameterType="Required",
			direction="Input")
		RouteID.parameterDependencies = [routes.name]
		RouteID.filter.list = ['Text']
			
		FromMM = arcpy.Parameter(
			displayName="From Measure Field Name",
			name="FromMM",
			datatype="GPString",
			parameterType="Required",
			direction="Input")
			
		ToMM = arcpy.Parameter(
			displayName="To Measure Field Name",
			name="ToMM",
			datatype="GPString",
			parameterType="Required",
			direction="Input")
			
		decPlaces = arcpy.Parameter(
			displayName="Decimal Places of Final Output From/To Measures",
			name="decPlaces",
			datatype="GPLong",
			parameterType="Required",
			direction="input")
			
		outFldr = arcpy.Parameter(
			displayName="Output Folder Location",
			name="outFldr",
			datatype="DEWorkspace",
			parameterType="Required",
			direction="Input")
			
			
		temporalTarget = arcpy.Parameter(
			displayName="Translate Date",
			name="translateDate",
			datatype="GPDate",
			parameterType="Optional",
			direction="Input")
		temporalTarget.value = datetime.utcnow()
					
		parameters = [indexTbl, RH_SDE, routes, FromDate, ToDate, temporalTarget, RouteID, FromMM, ToMM, decPlaces, SQL_Query, outFldr]
		
		return parameters

    def isLicensed(self):
        return True

    def execute(self, parameters, messages):
		#Set parameters as variables
		indexTbl = parameters[0].valueAsText
		RH_SDE = parameters[1].valueAsText
		RouteID = parameters[6].valueAsText
		FromMM = parameters[7].valueAsText
		ToMM = parameters[8].valueAsText
		decPlaces = parameters[9].valueAsText
		outFldr = parameters[11].valueAsText
		routes = parameters[2].valueAsText
		
		temporalTarget = parameters[5].value
		FromDate = parameters[3].valueAsText
		ToDate = parameters[4].valueAsText
		
		# Set and format temporal filter
		if temporalTarget is None:
			temporalTarget.value = datetime.utcnow()
		temporalTarget = "{:'%Y-%m-%d %H:%M:%S'}".format(temporalTarget)
		desc = arcpy.Describe(RH_SDE)
		if desc.workspaceType == 'LocalDatabase':
			temporalFilter = "({0} is null or {0}<=date{2}) and ({1} is null or {1}>date{2})".format(FromDate, ToDate, temporalTarget)
		else:
			temporalFilter = "({0} is null or {0}<={2}) and ({1} is null or {1}>{2})".format(FromDate, ToDate, temporalTarget)
		
		# Set SQL Query used in copying events
		SQL_Query = parameters[10].valueAsText
		if SQL_Query is None:
			SQL_Query = temporalFilter
		else:
			SQL_Query = "{0} AND {1}".format(temporalFilter, SQL_Query)
		
		#Run all the steps
		if checkIndex(RH_SDE, indexTbl, RouteID, FromMM, ToMM, FromDate, ToDate) > 0:
			arcpy.AddMessage("{0}: CANCELLING: Feature Class or Field Name from Index Table NOT FOUND in RH_SDE. Please check Messages for Error".format(timeStamp()))
			return	
		arcpy.AddMessage("{0}: All Fields and Feature Classes in Index Table found in RH_SDE".format(timeStamp()))
		arcpy.AddMessage("{0}: All Feature Classes in Index Table contain Route ID, From Measure and To Measure Fields".format(timeStamp()))
		
		gdbPath = createStagingGDB(outFldr, "OverlayStaging")
		
		arcpy.TableToTable_conversion(indexTbl, gdbPath, 'indexTbl')
		indexTbl = os.path.join(gdbPath, 'indexTbl')
		
		if copyTables(gdbPath, RH_SDE, indexTbl, SQL_Query, RouteID, FromMM, ToMM, FromDate, ToDate, decPlaces) > 0:
			arcpy.AddMessage("{0}: CANCELLING: Field Mapping of SQL Statement was incorrect. Please check Messages for Error".format(timeStamp()))
			return
		
		tbl_ov = overlayTables(gdbPath, indexTbl, RouteID, FromMM, ToMM)
		formatSTL(tbl_ov, indexTbl, gdbPath, RouteID, FromMM, ToMM, decPlaces, routes, outFldr, temporalFilter)
		
		return
from graphviz import Digraph
from DiagramObject import Node, Connection, Group
import yaml
import xlrd
import PySimpleGUI as sg
import pathlib
import argparse


global types_config
types_config = yaml.safe_load(open('config/types.yml'))

global graph_config
graph_config = yaml.safe_load(open('config/graph.yml'))


def getConnectionsList(excelBook,tabName):
	sheet = excelBook.sheet_by_name(tabName)
	keys = [str(sheet.cell(0,col_index).value) for col_index in range(sheet.ncols)]

	connectionList =[]
	for row_index in range(1,sheet.nrows):
		tmp_id = str(sheet.cell(row_index,keys.index("ID")).value)
		tmp_fromID = str(sheet.cell(row_index,keys.index("fromID")).value)
		tmp_toID = str(sheet.cell(row_index,keys.index("toID")).value)
		tmp_dictAttribute = {keys[col_index] : str(sheet.cell(row_index, col_index).value) for col_index in range(3,sheet.ncols)}
		tmp_Connection = Connection(tmp_id, tmp_fromID, tmp_toID,tmp_dictAttribute)
		connectionList.append(tmp_Connection)
	return connectionList

def getGroupsList(excelBook,tabName):
	sheet = excelBook.sheet_by_name(tabName)
	keys = [str(sheet.cell(0,col_index).value) for col_index in range(sheet.ncols)]

	groupList =[]
	for row_index in range(1,sheet.nrows):
		tmp_id = str(sheet.cell(row_index,keys.index("ID")).value)
		tmp_nodes = [x for x in str(sheet.cell(row_index,keys.index("Nodes")).value).split(',')]
		tmp_dictAttribute = {keys[col_index] : str(sheet.cell(row_index, col_index).value) for col_index in range(2,sheet.ncols)}
		tmp_group = Group(tmp_id, tmp_nodes, tmp_dictAttribute)
		groupList.append(tmp_group)
	return  groupList

def combineAttributeDicts(instanceDict,typeDict):
	for key,value in typeDict.items():
		if(isinstance(value, str) and instanceDict.get(key,"") == ""):
			instanceDict[key] = value


def appendNodesDict(excelBook,tabName, nodeDict):
	sheet = excelBook.sheet_by_name(tabName)
	keys = [str(sheet.cell(0,col_index).value) for col_index in range(sheet.ncols)]

	for row_index in range(1,sheet.nrows):
		tmp_id = str(sheet.cell(row_index,0).value)
		tmp_dictAttribute = {keys[col_index] : str(sheet.cell(row_index, col_index).value) for col_index in range(1,sheet.ncols)}
		combineAttributeDicts(tmp_dictAttribute, types_config.get(tabName,{}))
		tmp_node = Node(tmp_id, tmp_dictAttribute)
		nodeDict[tmp_id] = tmp_node
	return nodeDict

def getNodeDict(excelBook, tabsToIgnore):
	nodeDict = {}
	for sheetname in excelBook.sheet_names():
		if sheetname not in tabsToIgnore:
			appendNodesDict(excelBook,sheetname,nodeDict)
	return nodeDict

def getNodesByType(excelBook, tabsToIgnore):
	nodeTypesDict = {}
	for sheetname in excelBook.sheet_names():
		if sheetname not in tabsToIgnore:
			nodeTypesDict[sheetname] = []
			tmp_dict ={}
			nodeTypesDict[sheetname].append(appendNodesDict(excelBook,sheetname,tmp_dict))
	return nodeTypesDict

def addTypeConnectionsToGraph(graph, nodeTypesDict):
	count = 0
	for typeFromNode in types_config.keys():
		for typeToNode in types_config[typeFromNode].get('typeConnectionsTo',[]):
			if(typeFromNode in nodeTypesDict.keys() and typeToNode in nodeTypesDict.keys()):
				for nodeFrom in nodeTypesDict[typeFromNode][0].values():
					for nodeTo in nodeTypesDict[typeToNode][0].values():
						# print(types_config[typeFromNode].get("typeConnectionsTo","").get(typeToNode))
						graph.edge(nodeFrom.id,nodeTo.id, _attributes=types_config[typeFromNode].get("typeConnectionsTo",{}).get(typeToNode).get("connectionStyle",{}))


def setGraphAttributes(args,graph):
	if args.outputname:
		graph.filename = args.outputname
	graph.engine = graph_config['Graph'].get('engine',args.engine)
	graph.format = graph_config['Graph'].get('format',args.outputformat)
	graph.directory = graph_config['Graph'].get('directory',args.outputdirectory)
	graph.graph_attr = graph_config['Graph']


def createGraph(args,groupList, connectionList, nodeDict, nodeTypesDict):
	dot = Digraph()
	setGraphAttributes(args,dot)

	nodeDictCopy = nodeDict.copy()

	for group in groupList:
		with dot.subgraph(name="cluster_" + group.id) as sub:
			sub.graph_attr = group.attributeDict
			for nodeName in group.nodes:
				node = nodeDictCopy.pop(nodeName, None)
				sub.node(node.id, _attributes=node.attributeDict)

	for diagramObject in nodeDictCopy.values():
		dot.node(diagramObject.id, _attributes=diagramObject.attributeDict)

	for connection in connectionList:
		dot.edge(connection.fromID,connection.toID, _attributes=connection.attributeDict)

	addTypeConnectionsToGraph(dot,nodeTypesDict)

	print(dot.source)
	dot.render()

def getFile(inputfileArgument):
	if not inputfileArgument:
		event, values = sg.Window('ExcelToGraph').Layout([[sg.Text('Excel File to open')],[sg.In(), sg.FileBrowse()],[sg.CloseButton('Open'), sg.CloseButton('Cancel')]]).Read()
		fname = values[0]
		# print(event, values)
	else:
		fname = inputfileArgument
	if not fname:
		sg.Popup("Cancel", "No filename supplied")
		raise SystemExit("Cancelling: no filename supplied")
	# print (fname)
	return fname

def getCommandArguments():
	parser = argparse.ArgumentParser()
	parser.add_argument("-o", "--outputname", dest = "outputname", help="Saved File name")
	parser.add_argument("-e", "--engine", dest = "engine", default = "dot", help="Graphviz Engine")
	parser.add_argument("-i", "--inputfile",dest ="inputfile", help="Input Excel File")
	parser.add_argument("-f", "--outputformat",dest ="outputformat", default = "png", help="Input Excel File")
	parser.add_argument("-d", "--outputdirectory",dest ="outputdirectory", default = ".", help="Output Directory")
	args, unknown = parser.parse_known_args()
	return args, unknown

def main():
	args, unhandledArguments = getCommandArguments()
	excelFile = getFile(args.inputfile)
	connectionsTab = "Connections"
	groupsTab = "Clusters"
	
	print ("Starting...")
	excelBook = xlrd.open_workbook(excelFile)

	excelFilePath = pathlib.Path(excelFile)
	args.outputdirectory = excelFilePath.parents[0]

	connectionList = getConnectionsList(excelBook,connectionsTab)
	groupList = getGroupsList(excelBook,groupsTab)
	nodeDict = getNodeDict(excelBook,[connectionsTab,groupsTab])
	nodeTypesDict = getNodesByType(excelBook,[connectionsTab,groupsTab])
	createGraph(args,groupList,connectionList,nodeDict, nodeTypesDict)


if __name__=="__main__":
	main()


# ExcelToGraph

# Motivation
This project was created to allow quick and easy creation of graph diagrams.  The main idea is for the user to define Nodes, Connections, and Groups in an Excel document, and the let the program create the graph visual layout following a configuration file.

## Getting Started
### Prerequisites
Graphivz is the main backbone prerequisite of this project. To Install graphviz follow the install instructions at https://pypi.org/project/graphviz/
Note: This is only the python interface to graphviz, you will also need to download graphviz and add it to your PATH.  Instructions and Download link can be found here: https://www.graphviz.org/

Note: The dist folder contains an executable version of this project that uses an edited version of the python graphviz interface to point directly to the graphviz core executables in the Graphviz2.38 folder within this project.  This is for convenience to anyone wanting to only utilize this project, without being an active developer.

## Basic Usage
### Step  1
```
Create an Excel Spreadsheet with at minimum three tabs ('Connections', 'Clusters', and 'yourCustomNodeName')
Note: the tabs Connections and Clusters must be named as so, all other tabs will be considered as Nodes.
```

### Step 2
```
(Optional) Edit the Configuration files in the project folder 'config' to add a default style to each Type of Node (Excel Tab on the Spreadsheet)
```

### Step 3
Run the program with the command
```
python excelToGraph.py
```
or
```
excelToGraph.exe
```


## Commandline Arguments
excelToGraph accepts a few different arguments to dynamically change the behavior of the program, without changing the config files 
```
excelToGraph -o --outputname  "MyOutputGraph"  ## Set output filename
excelToGraph -e --engine  "dot"  ## Select one of Graphviz's built-in engines
excelToGraph -i --inputfile "MyExcelDocument.xlsx"  ## Select Excel file to read from, Setting the input file will also skip over the GUI file selector
excelToGraph -f --outputformat "png"  ## Select what file format to save the output (i.e. png, pdf, jpg, etc)
excelToGraph -d --outputdirectory  ## Select the Directory where the output file should be saved
```

## Config Files
### Graph.yml
Set default graphviz attributes for the global graph

Example:
```
Graph:
  splines: ortho
  ratio: fill
  size: 8.3,11.7!
  margin: '0' 
  format: png
  ranksep: '1.0 equally'
  # use # to comment out the line
  # concentrate : 'true' #merge edges that are parallel and have common endpoint
 ```
 
 ### Types.yml
 Define default styles for each Type of Node (Tabs in Excel Spreadsheet)
 
 Example:
 ```
 Server:
  image: images/Application_server.png
  labelloc: b
  shape : plaintext
  # style : dashed
  # fixedsize : 'true'
  width : '0.75'
  height : '0.75' 
  # margin : '0.5,0.5'
  # fontcolor : red
  # fontname : helvetica
  # penwidth : 0
  ordering: out
  typeConnectionsTo:  ## Custom attribute to this project, allows defining connections between types of nodes, i.e. all Clients connect                       ## to all Servers 
    Client :
      connectionStyle:
        color : purple
        sametail: serverclient
    SAN :
      connectionStyle:
        color : blue

Database:
  image: images/Application_server.png
  labelloc: b
  shape : plaintext
  # style : dashed
  # fixedsize : 'true'
  width : '0.75'
  height : '0.75' 
  # margin : '0.5,0.5'
  # fontcolor : red
  # fontname : helvetica
  # penwidth : 0
  ordering: out
      

Client:
  shape: box
```

## Excel Document Usage
### Restrictions - Tab Names
* Must contain a 'Clusters' tab 
  * used to define groups of nodes within the graph
* Must contain a 'Connections' tab
  * used to define connections/edges between nodes
* All other tabs will be treated as a type of Node

### Restrictions - Column Headers
* Every Tab needs an ID column header, this value must be unique to the element whether it is a Node, Connection, or Cluster
* The Connections Tab also requires ColumnHeaders for 'fromID' and 'toID'
* The Clusters Tab also requires ColumnHeaders for 'Nodes' # which is a comma separated list of Node IDs
* All other columns follow the convention of the ColumnHeader = GraphViz Attribute Name, with the associated attibute value below for each Node, Connection, or Cluster
  * See all possible graphviz attributes at https://www.graphviz.org/doc/info/attrs.html
  * These column attributes are great if you need to specify an attribute for a particular node, if you want all nodes of that Type (Tab on Excel Spreadsheet) then it would be better to define that in the config/types.yml file
  
Node Tab Example:

| ID | xLabel | shape |
| -- | ------ | ----- |
| Node1 | myNodeLabel1 | tab |
| Node2 | myNodeLabel2 |  |
| Node3 | myNodeLabel3 | box |

Connections Tab Example:

| ID | fromID | toID | color |
| -- | ------ | ---- | ----- |
| Edge1 | Node1 | Node2 | red |
| Edge2 | Node2 | Node3 | blue |

Clusters Tab Example:

| ID | Nodes | 
| -- | ------ |
| group1 | Node1,Node2 | 
| group2 | Node2,Node3 |  








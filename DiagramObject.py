class Node:
    def __init__(self, id, attributeDict):
        self.id = id
        self.attributeDict = attributeDict

class Connection:
    def __init__(self, id, fromID, toID, attributeDict): 
        self.id = id
        self.fromID = fromID
        self.toID = toID
        self.attributeDict = attributeDict

class Group:
    def __init__(self, id, nodes, attributeDict): 
        self.id = id
        self.nodes = nodes
        self.attributeDict = attributeDict
        

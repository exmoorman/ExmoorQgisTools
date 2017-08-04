from PyQt4 import *
from PyQt4.QtGui import *
from PyQt4.QtCore import *
from qgis.gui import *
from qgis.core import *
import pyodbc
import subprocess
from win32com.client import *
import win32gui
import win32api
import os
from global_variables import FileLocations

class OpenFolder(QgsMapTool):

    def __init__(self, canvas):
        QgsMapTool.__init__(self, canvas)
        self.canvas = canvas
  
    def canvasReleaseEvent(self, event):

        fx = event.pos().x()
        fy = event.pos().y()

        location = self.canvas.getCoordinateTransform().toMapCoordinates(fx, fy)
        boundbox = QgsRectangle(4999.99,4999.69,660000.06,1225000.12)

        parishInitials = {'Brendon' : 'RW1', 'Brompton Regis' : 'RW2', 'Carhampton' : 'RW3', 'Challacombe' : 'RW4', 'Combe Martin' : 'RW5', 'Countisbury' : 'RW6', 'Cutcombe' : 'RW7', 'Dulverton' : 'RW8', 'Dunster' : 'RW9', 'East Anstey' : 'RW10', 'Elworthy' : 'RW11', 'Exford' : 'RW12', 'Exmoor' : 'RW13', 'Exton' : 'RW14', 'High Bray' : 'RW15', 'Kentisbury' : 'RW16', 'Luccombe' : 'RW17', 'Luxborough' : 'RW18', 'Lynton & Lynmouth' : 'RW19', 'Martinhoe' : 'RW20', 'Minehead' : 'RW21', 'Minehead Without' : 'RW22', 'Molland' : 'RW23', 'Monksilver' : 'RW24', 'Nettlecombe' : 'RW25', 'North Molton' : 'RW26', 'Oare' : 'RW27', 'Old Cleeve' : 'RW28', 'Parracombe' : 'RW29', 'Porlock' : 'RW30', 'Selworthy' : 'RW31', 'Skilgate' : 'RW32', 'Timberscombe' : 'RW34', 'Treborough' : 'RW35', 'Trentishoe' : 'RW36', 'Twitchen' : 'RW37', 'Upton' : 'RW38', 'West Anstey' : 'RW39', 'Winsford' : 'RW40', 'Withycombe' : 'RW41', 'Withypool' : 'RW42', 'Wootton Courtenay' : 'RW43'}

        folderString = None
        pathString = None
        pathslashRemoval = None

        linkList = FileLocations(self)
        
        combined = str(location)

        QEastingString = str(combined[1:7])
        QNorthingString = str(combined[8:14])

        QEasting = int(QEastingString)
        QNorthing = int(QNorthingString)
        
        shortestDistance = float("inf")
        closestFeatureId = -1

        groupAlreadyOpen = False

        parish_layer = None
        path_layer = None
        
        TicketGroup = None

        parish_add = False
        path_add = False
        
        #Checks if the Ticket Creation Group exists already


        root = QgsProject.instance().layerTreeRoot()
        myGroup = QgsProject.instance().layerTreeRoot().findGroup( "Ticket Creation Layers" )
        if str(myGroup) == "None":
            TicketGroup = root.addGroup("Ticket Creation Layers")
            layer = QgsVectorLayer(linkList["parishlayerLocation"], "Parish Boundaries", "ogr")
            QgsMapLayerRegistry.instance().addMapLayer(layer, False)
            TicketGroup.addLayer(layer)
            layer.loadNamedStyle(linkList["blankparishstyleLocation"])
            layer.triggerRepaint()
            layer = QgsVectorLayer(linkList["rowlayerLocation"], "PROW", "ogr")
            QgsMapLayerRegistry.instance().addMapLayer(layer, False)
            TicketGroup.addLayer(layer)
            layer.loadNamedStyle(linkList["blankrowstyleLocation"])
            layer.triggerRepaint()
            layer = QgsVectorLayer(linkList["coastpathlayerLocation"], "Coast Path","ogr")
            QgsMapLayerRegistry.instance().addMapLayer(layer, False)
            TicketGroup.addLayer(layer)
            layer.loadNamedStyle(linkList["blankrowstyleLocation"])
            layer.triggerRepaint()
            layer = QgsVectorLayer(linkList["coleridgelayerLocation"], "Coleridge Way","ogr")
            QgsMapLayerRegistry.instance().addMapLayer(layer, False)
            TicketGroup.addLayer(layer)
            layer.loadNamedStyle(linkList["blankrowstyleLocation"])
            layer.triggerRepaint()
            layer = QgsVectorLayer(linkList["tmwlayerLocation"], "Two Moors Way","ogr")
            QgsMapLayerRegistry.instance().addMapLayer(layer, False)
            TicketGroup.addLayer(layer)
            layer.loadNamedStyle(linkList["blankrowstyleLocation"])
            layer.triggerRepaint()
        else:
            groupAlreadyOpen = True
        
        #Checks if the Parish Layer is Open and assigns it to parish_layer
        
        for layer in QgsMapLayerRegistry.instance().mapLayers().values():
            if linkList["parishlayerLocation"] in layer.source():
                parish_layer = layer
                break
        #If the Parish Layer is not open then it trys to open it up and assigns it to parish_layer

        if str(parish_layer) == "None":
            layer = QgsVectorLayer(linkList["parishlayerLocation"], "Parish Boundaries", "ogr")
            QgsMapLayerRegistry.instance().addMapLayer(layer)
            layer.loadNamedStyle(linkList["blankparishstyleLocation"])
            layer.triggerRepaint()
            

            for layer in QgsMapLayerRegistry.instance().mapLayers().values():
                if linkList["parishlayerLocation"] in layer.source():                    
                    parish_layer = layer
                    break

        #Check that the layer is open and then take the point value
               
        if str(parish_layer) != "None":
            pPnt = QgsGeometry.fromPoint(QgsPoint(QEasting, QNorthing))
            for f in parish_layer.getFeatures():
                pathdist = f.geometry().distance(QgsGeometry(pPnt))
                if pathdist < shortestDistance:
                    shortestDistance = pathdist
                    closestFeatureId = f.id()
            
            testlength = str(closestFeatureId)

        if len(testlength) > 0:
            fid = closestFeatureId
            iterator = parish_layer.getFeatures(QgsFeatureRequest().setFilterFid(fid))
            featuree = next(iterator)
            attrs = featuree.attributes()
            parishName = (attrs[0])
            lastTwo = parishName[-2:]
            if lastTwo == "CP":
            	parishName = parishName[:-3]

        
        else:
            parishName = None

        """if parish_close == True:
            QgsMapLayerRegistry.instance().removeMapLayer(parish_layer)"""



        shortestDistance = float("inf")
        closestFeatureId = -1
        layer = None

        for layer in QgsMapLayerRegistry.instance().mapLayers().values():
            if linkList["rowlayerLocation"] in layer.source():
                path_layer = layer
                break
        #If the Path Layer is not open then it trys to open it up and assigns it to path_layer

        if str(path_layer) == "None":
            layer = QgsVectorLayer(linkList["rowlayerLocation"], "PROW", "ogr")
            QgsMapLayerRegistry.instance().addMapLayer(layer)
            layer.loadNamedStyle(linkList["blankrowstyleLocation"])
            layer.triggerRepaint()
            path_close = True
            for layer in QgsMapLayerRegistry.instance().mapLayers().values():
                if linkList["rowlayerLocation"] in layer.source():
                    path_layer = layer
                    break

        #Check that the layer is open and then take the point value
               
        if str(path_layer) != "None":
            pPnt = QgsGeometry.fromPoint(QgsPoint(QEasting, QNorthing))
            for f in path_layer.getFeatures():
                pathdist = f.geometry().distance(QgsGeometry(pPnt))
                if pathdist < shortestDistance:
                    shortestDistance = pathdist
                    closestFeatureId = f.id()
            
            testlength = str(closestFeatureId)

        if len(testlength) > 0:
            fid = closestFeatureId
            iterator = path_layer.getFeatures(QgsFeatureRequest().setFilterFid(fid))
            featuree = next(iterator)
            attrs = featuree.attributes()
            pathName = (attrs[0])
        
        else:
            pathName = None

        """if path_close == True:
            QgsMapLayerRegistry.instance().removeMapLayer(path_layer)"""

        #Y:\Access & Recreation Team\ROW  RW\Parish-Path Files\RW42 Withypool\DU 11-14

        if boundbox.contains(location):

            pathslashRemoval = pathName.replace("/", "-")
            pathString = 'Y:\Access & Recreation Team\ROW  RW\Parish-Path Files' + '\\' + parishInitials[parishName] + " " + parishName + "\\" + pathslashRemoval
            openpathString = "explorer " + '"Y:\Access & Recreation Team\ROW  RW\Parish-Path Files\\' + parishInitials[parishName] + " " + parishName + "\\" + pathslashRemoval +'"'
            if os.path.isdir(pathString) is True:
                res = subprocess.Popen(openpathString)
            else:
               QMessageBox.information(None, "Slight Problem....", "The path folder doesn't appear to exist. Will open the parish folder instead.....\nThe path folder name has to be exactly as it is on the path layer \ni.e DU 1/4 or 243BY17 etc")
               openpathString = "explorer " + '"Y:\Access & Recreation Team\ROW  RW\Parish-Path Files\\' + parishInitials[parishName] + " " + parishName
               res = subprocess.Popen(openpathString)
 
            
        else:
            QMessageBox.information(None, "Waaaay out", "Point out of bounds")

from PyQt4 import *
from PyQt4.QtGui import *
from PyQt4.QtCore import *
from qgis.gui import *
from qgis.core import *
import pyodbc
import subprocess
from win32com.client import *
import win32gui
import win32con
import win32api
import os
import ctypes
from global_variables import FileLocations 
import time



class TicketTool(QgsMapTool):

    def __init__(self, canvas):
        QgsMapTool.__init__(self, canvas)
        self.canvas = canvas

 
    def canvasReleaseEvent(self, event):

        fx = event.pos().x()
        fy = event.pos().y()

        location = self.canvas.getCoordinateTransform().toMapCoordinates(fx, fy)
        boundbox = QgsRectangle(4999.99,4999.69,660000.06,1225000.12)

        linkList = FileLocations(self)
        
        combined = str(location)

        QEastingString = str(combined[1:7])
        QNorthingString = str(combined[8:14])

        QEasting = int(QEastingString)
        QNorthing = int(QNorthingString)
        
        coastshortestDistance = float("inf")
        coleridgeshortestDistance = float("inf")
        tmwshortestDistance = float("inf")
        shortestDistance = float("inf")
        closestFeatureId = -1

        groupAlreadyOpen = False

        parish_layer = None
        path_layer = None
        coleridge_layer = None
        coastpath_layer = None
        tmw_layer = None
        TicketGroup = None

        parish_add = False
        path_add = False
        coleridgepath_add = False
        coastpath_add = False
        tmw_add = False
        
        coastpath_check = False 
        coleridge_check = False
        tmw_check = False

        EnumWindows = ctypes.windll.user32.EnumWindows
        EnumWindowsProc = ctypes.WINFUNCTYPE(ctypes.c_bool, ctypes.POINTER(ctypes.c_int), ctypes.POINTER(ctypes.c_int))
        GetWindowText = ctypes.windll.user32.GetWindowTextW
        GetWindowTextLength = ctypes.windll.user32.GetWindowTextLengthW
        IsWindowVisible = ctypes.windll.user32.IsWindowVisible
 
        titles = []
        pointer = []
        s=0
        t=0

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

        #Check distance to the coast path. If less than 100 mark it as promoted route

        for layer in QgsMapLayerRegistry.instance().mapLayers().values():
            if linkList["coastpathlayerLocation"] in layer.source():
                coastpath_layer = layer
                break

        if str(coastpath_layer) == "None":
            layer = QgsVectorLayer(linkList["coastpathlayerLocation"], "Coast Path","ogr")
            QgsMapLayerRegistry.instance().addMapLayer(layer)
            layer.loadNamedStyle(linkList["blankrowstyleLocation"])
            layer.triggerRepaint()
            coastpath_close = True
            for layer in QgsMapLayerRegistry.instance().mapLayers().values():
                if linkList["coastpathlayerLocation"] in layer.source():
                    coastpath_layer = layer
                    break

        if str(coastpath_layer) != "None":
            pPnt = QgsGeometry.fromPoint(QgsPoint(QEasting, QNorthing))
            for f in coastpath_layer.getFeatures():
                pdist = f.geometry().distance(QgsGeometry(pPnt))
                if pdist < coastshortestDistance:
                    coastshortestDistance = int(pdist)
            if coastshortestDistance <= 100:
                coastpath_check = True
        else:
            coastpath_check = False

        """if coastpath_close == True:
            QgsMapLayerRegistry.instance().removeMapLayer(coastpath_layer)"""

        #Check distance to the coleridge way. If less than 100 mark it as promoted route

        for layer in QgsMapLayerRegistry.instance().mapLayers().values():
            if linkList["coleridgelayerLocation"] in layer.source():
                coleridge_layer = layer
                break

        if str(coleridge_layer) == "None":
            layer = QgsVectorLayer(linkList["coleridgelayerLocation"], "Coleridge Way","ogr")
            QgsMapLayerRegistry.instance().addMapLayer(layer)
            layer.loadNamedStyle(linkList["blankrowstyleLocation"])
            layer.triggerRepaint()
            coleridgepath_close = True
            for layer in QgsMapLayerRegistry.instance().mapLayers().values():
                if linkList["coleridgelayerLocation"] in layer.source():
                    coleridge_layer = layer
                    break

        if str(coleridge_layer) != "None":
            pPnt = QgsGeometry.fromPoint(QgsPoint(QEasting, QNorthing))
            for f in coleridge_layer.getFeatures():
                pdist = f.geometry().distance(QgsGeometry(pPnt))
                if pdist < coleridgeshortestDistance:
                    coleridgeshortestDistance = int(pdist)
            if coleridgeshortestDistance <= 100:
                coleridge_check = True
        else:
            coleridge_check = False

        """if coleridgepath_close == True:
            QgsMapLayerRegistry.instance().removeMapLayer(coleridge_layer)"""

        #Check distance to the 2M way. If less than 100 mark it as promoted route

        for layer in QgsMapLayerRegistry.instance().mapLayers().values():
            if linkList["tmwlayerLocation"] in layer.source():
                tmw_layer = layer
                break

        if str(tmw_layer) == "None":
            layer = QgsVectorLayer(linkList["tmwlayerLocation"], "Two Moors Way","ogr")
            QgsMapLayerRegistry.instance().addMapLayer(layer)
            layer.loadNamedStyle(linkList["blankrowstyleLocation"])
            layer.triggerRepaint()
            tmw_close = True
            for layer in QgsMapLayerRegistry.instance().mapLayers().values():
                if linkList["tmwlayerLocation"] in layer.source():
                    tmw_layer = layer
                    break

        if str(tmw_layer) != "None":
            pPnt = QgsGeometry.fromPoint(QgsPoint(QEasting, QNorthing))
            for f in tmw_layer.getFeatures():
                pdist = f.geometry().distance(QgsGeometry(pPnt))
                if pdist < tmwshortestDistance:
                    tmwshortestDistance = int(pdist)
            if tmwshortestDistance <= 100:
                tmw_check = True
        else:
            tmw_check = False

              
        if boundbox.contains(location):
            #QMessageBox.information(None, "Ticket Info", "Parish: " + str(parishName) + " Path: " + str(pathName) + " QEasting: " + str(QEasting) + " QNorthing: " + str(QNorthing) + " Coastpath Ticket: " + str(coastpath_check) + " Coleridge Ticket: " + str(coleridge_check) + " 2MW: " + str(tmw_check))
            

            #This makes a connection to the database and writes the new ticket in
            db_localfile = str(linkList["db_file"])
            #db_file = r'''C:\Users\TGP\OneDrive - Exmoor National Park Authority\My Documents\Ticket System\New Database\Experimental\QGIS Integration\Test Database\test.accdb'''''
            user = 'admin'
            password = ''
            odbc_conn_str = 'DRIVER={Microsoft Access Driver (*.mdb, *.accdb)};DBQ=%s;UID=%s;PWD=%s' %\
                            (db_localfile, user, password)

            conn = pyodbc.connect(odbc_conn_str)
            cur = conn.cursor()     
            cur.execute("INSERT INTO Ticket_Details (Parish, Path_No, Easting, Northing, CoastPath, Coleridge, 2MW) VALUES (?, ?, ?, ?, ?, ?, ?)", parishName, pathName, QEasting, QNorthing, coastpath_check, coleridge_check, tmw_check)
            conn.commit()
            cur.commit()
            conn.close()

            #Check if MS Access is already running

            EnumWindows = ctypes.windll.user32.EnumWindows
            EnumWindowsProc = ctypes.WINFUNCTYPE(ctypes.c_bool, ctypes.POINTER(ctypes.c_int), ctypes.POINTER(ctypes.c_int))
            GetWindowText = ctypes.windll.user32.GetWindowTextW
            GetWindowTextLength = ctypes.windll.user32.GetWindowTextLengthW
            IsWindowVisible = ctypes.windll.user32.IsWindowVisible
             
            titles = []
            pointer = []
            def foreach_window(hwnd, lParam):
                if IsWindowVisible(hwnd):
                    length = GetWindowTextLength(hwnd)
                    buff = ctypes.create_unicode_buffer(length + 1)
                    GetWindowText(hwnd, buff, length + 1)
                    titles.append(buff.value)
                    pointer.append(hwnd)
                return True
   
            EnumWindows(EnumWindowsProc(foreach_window), 0)

            #QMessageBox.information(None, "List", str(pointer))

            formName = linkList["db_form"]
            strDbName = str(linkList["db_file"])
               
                        
            if any("Access" in t for t in titles):
                oApp = Dispatch("Access.Application")
                oApp.DoCmd.Close(2, formName)
                time.sleep(2)
                oApp.DoCmd.OpenForm(formName)
                oApp.Visible = True

            else:
                oApp = Dispatch("Access.Application")
                oApp.OpenCurrentDatabase(strDbName)
                oApp.DoCmd.OpenForm(formName)
                oApp.Visible = True




                
            

            
           
            #QMessageBox.information(None, "Ticket Info", "Parish: " + str(parishName) + " Path: " + str(pathName) + " QEasting: " + str(QEasting) + " QNorthing: " + str(QNorthing))
        else:
            QMessageBox.information(None, "Fault Can't Be Placed", "Point out of bounds")


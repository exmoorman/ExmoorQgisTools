from PyQt4 import *
from PyQt4.QtGui import *
from PyQt4.QtCore import *
from qgis.gui import *
from qgis.core import *


def makeInspections(self):

    layer_list = ['Barriers', 'Benches', 'Boardwalks', 'Cattle Grids', 'Fords', 'Gates', 'Handrails', 'Signposts', 'Signs', 'Steps', 'Stiles', 'Tunnels', 'Walls']
    missing = set()
    for i in range(len(layer_list)):
        layer = None
        for lyr in QgsMapLayerRegistry.instance().mapLayers().values():
            if lyr.name() == layer_list[i]:
                layer = lyr
                break
        if str(layer)=="None":
            tempName = str(layer_list[i])
            missing.add(tempName)

    if len(missing)==0:            
        vlr = QgsVectorLayer("Point?crs=EPSG:27700", "inspection_structure", "memory")
        pr = vlr.dataProvider()
        pr.addAttributes( [ QgsField("type", QVariant.String), QgsField("sub_type", QVariant.String), QgsField("inspection", QVariant.String) ] )
        vlr.commitChanges()
        vlr.updateExtents()
        QgsMapLayerRegistry.instance().addMapLayer(vlr)
        vlr.startEditing()

        for i in range(len(layer_list)):
            layer = None
            for lyr in QgsMapLayerRegistry.instance().mapLayers().values():
                if lyr.name() == layer_list[i]:
                    layer = lyr
                    break
            exprn = QgsExpression( " \"inspection\" = 'Yes' " )
            itr = layer.getFeatures( QgsFeatureRequest( exprn ) )
            ids = [i.id() for i in itr]
            self.iface.setActiveLayer(layer)
            layer.setSelectedFeatures(ids)
            self.iface.actionCopyFeatures().trigger()
            self.iface.setActiveLayer(vlr)
            self.iface.actionPasteFeatures().trigger()
            self.iface.mainWindow().statusBar().clearMessage()
        QMessageBox.information(None, "INFORMATION", "A layer called inspection_structure has been created.\n"
                                "It is only temporary so save it to where you need it.\n")
        self.iface.actionPan().trigger()
        self.iface.mainWindow().statusBar().clearMessage()
    else:
        self.iface.actionPan().trigger()
        msg = QMessageBox()
        msg.setIcon(QMessageBox.Information)
        msg.setText("To create the Structures Requiring Inspection, the correct layers need to be open.")
        msg.setInformativeText("Click Details to see the missing layers, or layers incorrectly named.")
        msg.setWindowTitle("Missing Information")
        msg.setDetailedText(str(list(missing)))
        msg.setStandardButtons(QMessageBox.Cancel)
        retval = msg.exec_()


    




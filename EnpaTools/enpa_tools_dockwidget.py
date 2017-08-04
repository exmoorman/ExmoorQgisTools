# -*- coding: utf-8 -*-
"""
/***************************************************************************
 EnpaToolsDockWidget
                                 A QGIS plugin
 This plugin is a smorgasbord of tools to make a rangers life easier
                             -------------------
        begin                : 2017-06-22
        git sha              : $Format:%H$
        copyright            : (C) 2017 by ENPA
        email                : enquiries@exmoor-nationalpark.gov.uk
 ***************************************************************************/

/***************************************************************************
 *                                                                         *
 *   This program is free software; you can redistribute it and/or modify  *
 *   it under the terms of the GNU General Public License as published by  *
 *   the Free Software Foundation; either version 2 of the License, or     *
 *   (at your option) any later version.                                   *
 *                                                                         *
 ***************************************************************************/
"""




from PyQt4 import *
from PyQt4.QtGui import *
from PyQt4.QtCore import *
from qgis.gui import *
from qgis.core import *
import os
import resources

from folder_tool import OpenFolder
from ticket_tool import TicketTool
from layer_operations import makeInspections
from xy_to_osgb import xy_to_osgb, osgb_to_xy
from grid_ref_utils import (
    reproject_point_to_4326,
    reproject_point_from_4326,
    GridRefException,
    centre_on_point,
    point_from_longlat_text,
    gen_marker
)



FORM_CLASS, _ = uic.loadUiType(os.path.join(
    os.path.dirname(__file__), 'enpa_tools_dockwidget_base.ui'))


class EnpaToolsDockWidget(QtGui.QDockWidget, FORM_CLASS):

    closingPlugin = pyqtSignal()

    def __init__(self, iface, plugin, parent=None):
        """Constructor."""
        super(EnpaToolsDockWidget, self).__init__(parent)
        # Set up the user interface from Designer.
        # After setupUI you can access any designer object by doing
        # self.<objectname>, and you can use autoconnect slots - see
        # http://qt-project.org/doc/qt-4.8/designer-using-a-ui-file.html
        # #widgets-and-dialogs-with-auto-connect
        self.setupUi(self)
        self.iface = iface
        self._connect_signals()
        self._add_validators()
        self.marker = None

    def _connect_signals(self):
        self.createFault.clicked.connect(self.makeFault)
        self.createInspections.clicked.connect(self.structureInspections)
        self.openpathFolder.clicked.connect(self.openFolder)
        self.iface.mapCanvas().xyCoordinates.connect(self.trackCoords)
        self.editCoords.returnPressed.connect(self.setCoords)

    def makeFault(self):
        tool = TicketTool(self.iface.mapCanvas())
        tool.setButton(self.createFault)
        self.iface.mapCanvas().setMapTool(tool)

    def openFolder(self):
        tool = OpenFolder(self.iface.mapCanvas())
        tool.setButton(self.openpathFolder)
        self.iface.mapCanvas().setMapTool(tool)

    def structureInspections(self):
        
        makeInspections(self)
        

    def _add_validators(self):
        re = QRegExp(r'''^(\s*[a-zA-Z]{2}\s*\d{1,4}\s*\d{1,4}\s*|\[out of bounds\])$''')
        self.editCoords.setValidator(QRegExpValidator(re, self))    

    def trackCoords(self, pt):
        self._setEditCooordsOnMouseMove(pt)
        
        self._remove_marker()

    def _setEditCooordsOnMouseMove(self, pt):

        precision = 100

        try:
            os_ref = xy_to_osgb(pt.x(), pt.y(), precision)
        except GridRefException:
            os_ref = "[out of bounds]"
        self.editCoords.setText(os_ref)

    def setCoords(self):
        try:
          x,y = osgb_to_xy(self.editCoords.text())
          point27700 = QgsPoint(x,y)
          centre_on_point(self.iface.mapCanvas(), point27700)
          self._add_marker(point27700)
        except GridRefException:
            QMessageBox.warning(
              self.iface.mapCanvas(),
              "How many times.....",
              "Just type in a normal grid reference\n"
              "like SS860560\n"
              "No spaces.")
       

    def _remove_marker(self):
        if self.marker:
            self.iface.mapCanvas().scene().removeItem(self.marker)
            self.marker = None

    def _add_marker(self, point):
        if self.marker:
            self._remove_marker()
        self.marker = gen_marker(self.iface.mapCanvas(), point)


    def closeEvent(self, event):
        self.closingPlugin.emit()
        self.iface.actionPan().trigger()
        event.accept()


# -*- coding: utf-8 -*-
"""
/***************************************************************************
 shpsyncDialog
                                 A QGIS plugin
 description
                             -------------------
        begin                : 2016-01-24
        git sha              : $Format:%H$
        copyright            : (C) 2016 by OpenGis.ch
        email                : email@address.com
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

import os

from PyQt4 import QtGui, uic
from qgis.gui import QgsFieldExpressionWidget
from qgis._core import QgsMessageLog, QgsMapLayerRegistry, QgsFeatureRequest, QgsFeature, QgsVectorJoinInfo, QgsExpression

import qgis_utils

FORM_CLASS, _ = uic.loadUiType(os.path.join(
    os.path.dirname(__file__), 'shpsync_dialog_base.ui'))


class shpsyncDialog(QtGui.QDialog, FORM_CLASS):
    def __init__(self, parent=None):
        """Constructor."""
        super(shpsyncDialog, self).__init__(parent)
        # Set up the user interface from Designer.
        # After setupUI you can access any designer object by doing
        # self.<objectname>, and you can use autoconnect slots - see
        # http://qt-project.org/doc/qt-4.8/designer-using-a-ui-file.html
        # #widgets-and-dialogs-with-auto-connect
        self.setupUi(self)
        self.exps = []
        self.dels = []
        self.combos = []
        self.slave = None
        self.master = None
        self.addExpressionWidget()
        self.addExpressionWidget()
        self.populate(self.comboBox_master,isMaster=True)
        self.populate(self.comboBox_slave,isMaster=False)

    def addExpressionWidget(self):
        hor = QtGui.QHBoxLayout()        
        fieldExp = QgsFieldExpressionWidget()
        combo = QtGui.QComboBox()
        hor.addWidget(combo)
        self.combos.append(combo)
        hor.addWidget(fieldExp)
        del_btn=QtGui.QPushButton("Delete")
        hor.addWidget(del_btn)
        self.dels.append(del_btn)
        self.verticalLayout.addLayout(hor)
        self.exps.append(fieldExp)

    def populate(self, comboBox, isMaster):
        idlayers = list(QgsMapLayerRegistry.instance().mapLayers().iteritems())
        self.populateFromLayers(comboBox,idlayers, isMaster)
        if isMaster:
            self.masterUpdated(0)
        else:
            self.slaveUpdated(0)

    def populateFromLayers(self,comboBox, idlayers, isMaster):
        for (id, layer) in idlayers:
            unicode_name = unicode(layer.name())
            comboBox.addItem(unicode_name,id)

        if isMaster:
            comboBox.currentIndexChanged.connect(self.masterUpdated)
        else:
            comboBox.currentIndexChanged.connect(self.slaveUpdated)


    def updateComboBoxFromLayerAttributes(self, comboBox, attrs):
        for attr in attrs:
            comboBox.addItem(attr.name())



    def masterUpdated(self, idx):
        layer = qgis_utils.getLayerFromId(self.comboBox_master.itemData(idx)) 
        self.master = layer
        attributes = layer.pendingFields() 
        self.updateComboBoxFromLayerAttributes(self.comboBox_master_key, attributes)
        # update layer in expressions
        for exp in self.exps:
            exp.setLayer(layer)

    def slaveUpdated(self, idx):
        layer = qgis_utils.getLayerFromId(self.comboBox_slave.itemData(idx)) 
        self.slave = layer
        attributes = layer.pendingFields() 
        self.updateComboBoxFromLayerAttributes(self.comboBox_slave_key, attributes)
        # update sheet name suggestion
        self.lineEdit_sheetName.setText(layer.name())
        # update fields in comboboxes
        for combo in self.combos:
            self.updateComboBoxFromLayerAttributes(combo,attributes)





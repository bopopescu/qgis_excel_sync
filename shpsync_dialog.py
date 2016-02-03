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
        self.hors = []
        self.slave = None
        self.master = None
        self.addExpressionWidget()
        self.addExpressionWidget()
        self.addExpressionWidget()
        self.populate(self.comboBox_master, isMaster=True)
        self.populate(self.comboBox_slave, isMaster=False)
        self.pushButton.clicked.connect(self.addExpressionWidget)

    def addExpressionWidget(self):
        hor = QtGui.QHBoxLayout()
        fieldExp = QgsFieldExpressionWidget()
        combo = QtGui.QComboBox()
        hor.addWidget(combo)
        self.combos.append(combo)
        hor.addWidget(fieldExp)
        del_btn = QtGui.QPushButton("Delete")
        hor.addWidget(del_btn)
        self.dels.append(del_btn)
        self.verticalLayout.addLayout(hor)
        self.exps.append(fieldExp)
        self.hors.append(hor)
        if self.slave is not None:
            self.updateComboBoxFromLayerAttributes(combo, self.slave.fields())
        if self.master is not None:
            fieldExp.setLayer(self.master)

        del_btn.clicked.connect(self.removeExpressionWidget)

        # TODO: resize window to fit or make it look nice somehow ?
        # TODO: field combo box can be very small ?

    def getExpressionsDict(self):
        res = {}
        for exp, combo in zip(self.exps, self.combos):
            res[combo.currentText()] = exp.currentText()
        return res

    def removeExpressionWidget(self):
        sender = self.sender()
        idx = self.dels.index(sender)
        hor = self.hors[idx]
        # self.verticalLayout.removeItem(hor.itemAt(0))
        self.dels[idx].setVisible(False)
        del self.dels[idx]
        self.combos[idx].setVisible(False)
        del self.combos[idx]
        self.exps[idx].setVisible(False)
        del self.exps[idx]
        del self.hors[idx]

    def populate(self, comboBox, isMaster):
        idlayers = list(QgsMapLayerRegistry.instance().mapLayers().iteritems())
        self.populateFromLayers(comboBox, idlayers, isMaster)
        if not idlayers:
            return
        if isMaster:
            self.masterUpdated(0)
        else:
            self.slaveUpdated(0)

    def populateFromLayers(self, comboBox, idlayers, isMaster):
        for (id, layer) in idlayers:
            unicode_name = unicode(layer.name())
            comboBox.addItem(unicode_name, id)

        if isMaster:
            comboBox.currentIndexChanged.connect(self.masterUpdated)
        else:
            comboBox.currentIndexChanged.connect(self.slaveUpdated)

    def updateComboBoxFromLayerAttributes(self, comboBox, attrs):
        comboBox.clear()
        for attr in attrs:
            comboBox.addItem(attr.name())

    def masterUpdated(self, idx):
        layer = qgis_utils.getLayerFromId(self.comboBox_master.itemData(idx))
        self.master = layer
        attributes = layer.fields()
        self.updateComboBoxFromLayerAttributes(
            self.comboBox_master_key, attributes)
        # update layer in expressions
        for exp in self.exps:
            exp.setLayer(layer)

    def slaveUpdated(self, idx):
        layer = qgis_utils.getLayerFromId(self.comboBox_slave.itemData(idx))
        self.slave = layer
        attributes = layer.fields()
        self.updateComboBoxFromLayerAttributes(
            self.comboBox_slave_key, attributes)
        # update sheet name suggestion
        self.lineEdit_sheetName.setText(layer.name())
        # update fields in comboboxes
        for combo in self.combos:
            self.updateComboBoxFromLayerAttributes(combo, attributes)

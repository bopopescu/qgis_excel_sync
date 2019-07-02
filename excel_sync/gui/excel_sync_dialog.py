# -*- coding: utf-8 -*-
"""
/***************************************************************************
 excel_sync_dialog.py
                             -------------------
        begin                : 2016
        git sha              : $Format:%H$
        copyright            : (C) 2016 by OPENGIS.ch
        email                : info@opengis.ch
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

from qgis.PyQt import uic
from qgis.PyQt.QtCore import Qt, QSize
from qgis.PyQt.QtWidgets import (
    QDialog, QHBoxLayout, QComboBox, QPushButton, QSizePolicy)

from qgis.gui import QgsFieldExpressionWidget
from qgis.core import QgsProject, QgsMapLayer

from xlrd import (
    open_workbook,
    XLRDError
)

from ..core import qgis_utils

FORM_CLASS = uic.loadUiType(os.path.join(
    os.path.dirname(__file__), 'excel_sync_dialog_base.ui'))[0]


class excel_syncDialog(QDialog, FORM_CLASS):

    def __init__(self, parent=None):
        """Constructor."""
        super(excel_syncDialog, self).__init__(parent)
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
        self.populate(self.comboBox_master, isMaster=True)
        self.populate(self.comboBox_slave, isMaster=False)
        self.pushButton.clicked.connect(self.addExpressionWidget)

    def restoreSettings(self, settings):
        master_idx = self.comboBox_master.findText(settings.shpName)
        slave_idx = self.comboBox_slave.findText(settings.excelName)

        self.populate(self.comboBox_master, isMaster=True, idx=master_idx)
        self.comboBox_master.setCurrentIndex(master_idx)
        self.comboBox_master.setEnabled(False)
        self.comboBox_master_key.setEnabled(False)
        self.populate(self.comboBox_master, isMaster=True, idx=master_idx)
        self.populate(self.comboBox_slave, isMaster=False, idx=slave_idx)
        self.comboBox_slave.setCurrentIndex(slave_idx)
        self.comboBox_slave.setEnabled(False)
        self.comboBox_slave_key.setEnabled(False)
        self.lineEdit_sheetName.setText(settings.excelSheetName)
        self.spinBox.setValue(settings.skipLines)
        self.comboBox_slave_key.setCurrentIndex(
            self.comboBox_slave_key.findText(settings.excelKeyName))
        self.comboBox_master_key.setCurrentIndex(
            self.comboBox_master_key.findText(settings.shpKeyName))
        self.checkBox.setCheckState(
            Qt.Checked if settings.hideDialog else Qt.Unchecked)

        for k, v in settings.expressions.items():
            self.addExpressionWidget()
            self.exps[-1].setLayer(self.master)
            self.exps[-1].setField(v)
            self.combos[-1].setCurrentIndex(self.combos[-1].findText(k))

    def addExpressionWidget(self):
        hor = QHBoxLayout()
        fieldExp = QgsFieldExpressionWidget()
        combo = QComboBox()
        hor.addWidget(combo)
        self.combos.append(combo)
        hor.addWidget(fieldExp)
        del_btn = QPushButton(self.tr("Delete"))
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
        # hor = self.hors[idx]
        # self.verticalLayout.removeItem(hor.itemAt(0))
        self.dels[idx].setVisible(False)
        del self.dels[idx]
        self.combos[idx].setVisible(False)
        del self.combos[idx]
        self.exps[idx].setVisible(False)
        del self.exps[idx]
        del self.hors[idx]

    def populate(self, comboBox, isMaster, idx=0, update=True):
        idlayers = list(QgsProject.instance().mapLayers().items())
        self.populateFromLayers(comboBox, idlayers, isMaster)
        comboBox.setCurrentIndex(idx)
        if not idlayers:
            return
        if not update:
            return
        if isMaster:
            self.masterUpdated(idx)
        else:
            self.slaveUpdated(idx)

    def populateFromLayers(self, comboBox, idlayers, isMaster):
        comboBox.clear()
        for (id, layer) in idlayers:
            if layer.type() == QgsMapLayer.VectorLayer:
                unicode_name = layer.name()
                comboBox.addItem(unicode_name, id)

        if isMaster:
            comboBox.currentIndexChanged.connect(self.masterUpdated)
        else:
            comboBox.currentIndexChanged.connect(self.slaveUpdated)

    def updateComboBoxFromLayerAttributes(self, comboBox, attrs):
        comboBox.clear()
        for attr in attrs:
            comboBox.addItem(attr.name())
        comboBox.setMinimumSize(QSize(150, 0))
        comboBox.setSizePolicy(
            QSizePolicy.Preferred, QSizePolicy.Preferred)

    def masterUpdated(self, idx):
        layer = qgis_utils.getLayerFromId(self.comboBox_master.itemData(idx))
        if layer is None:
            return
        self.master = layer
        attributes = layer.fields()
        self.updateComboBoxFromLayerAttributes(
            self.comboBox_master_key, attributes)
        # update layer in expressions
        for exp in self.exps:
            text = exp.currentText()
            exp.setLayer(layer)
            exp.setField(text)

    def slaveUpdated(self, idx):
        layer = qgis_utils.getLayerFromId(self.comboBox_slave.itemData(idx))
        if layer is None:
            return
        self.slave = layer
        attributes = layer.fields()
        self.updateComboBoxFromLayerAttributes(
            self.comboBox_slave_key, attributes)
        # update sheet name suggestion
        try:
            wb = open_workbook(layer.publicSource().split('|')[0])
            self.lineEdit_sheetName.setText(wb.sheet_names()[0])
        except (XLRDError, OSError):
            pass
        # update fields in comboboxes
        for combo in self.combos:
            self.updateComboBoxFromLayerAttributes(combo, attributes)

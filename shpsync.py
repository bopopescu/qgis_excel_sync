# -*- coding: utf-8 -*-
"""
/***************************************************************************
 shpsync
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
import os.path
from collections import OrderedDict

from PyQt4.QtCore import QSettings, QTranslator, qVersion, QCoreApplication, QObject, SIGNAL
from PyQt4.QtGui import QAction, QIcon
from qgis._core import QgsProject

import resources
from shpsync_dialog import shpsyncDialog
from shp_excel_sync import Settings,Syncer
from project_handler import ProjectHandler


class shpsync:
    """QGIS Plugin Implementation."""

    def setUpSyncerTest(self,excelName,excelKeyName,shpName,shpKeyName):
        """Test the setup"""
        exps={"Flaeche_ha":"area( $geometry )", "FEE_Nr":"y( $geometry )"}
        s = Settings(excelName,"Tabelle1",excelKeyName,1,shpName,shpKeyName,exps)
        self.syncer = Syncer(s)

    def __init__(self, iface):
        """Constructor.

        :param iface: An interface instance that will be passed to this class
            which provides the hook by which you can manipulate the QGIS
            application at run time.
        :type iface: QgsInterface
        """
        # Save reference to the QGIS interface
        self.iface = iface
        self.syncer = None
        # initialize plugin directory
        self.plugin_dir = os.path.dirname(__file__)
        # initialize locale
        locale = QSettings().value('locale/userLocale')[0:2]
        locale_path = os.path.join(
            self.plugin_dir,
            'i18n',
            'shpsync_{}.qm'.format(locale))

        if os.path.exists(locale_path):
            self.translator = QTranslator()
            self.translator.load(locale_path)

            if qVersion() > '4.3.3':
                QCoreApplication.installTranslator(self.translator)

        self.dlg = None

        # Declare instance attributes
        self.actions = []
        self.menu = self.tr(u'&ShpSync')
        # TODO: We are going to let the user set this up in a future iteration
        self.toolbar = self.iface.addToolBar(u'shpsync')
        self.toolbar.setObjectName(u'shpsync')
        self.initProject()

    def initProject(self):
        """ initialize project related connections """
        self.iface.projectRead.connect(self.readSettings)
        QObject.connect(QgsProject.instance(), SIGNAL("writeProject(QDomDocument &)"),
                        self.writeSettings)

    def readSettings(self):
        # "Settings","excelName excelSheetName excelKeyName skipLines shpName shpKeyName expressions")
        metasettings = OrderedDict()
        metasettings["excelName"] = str
        metasettings["excelSheetName"] = str
        metasettings["excelKeyName"] = str
        metasettings["skipLines"] = int
        metasettings["shpKeyName"] = str
        metasettings["shpName"] = str
        metasettings["expressions"] = list
        settings_dict = ProjectHandler.readSettings("SHPSYNC",metasettings)
        if not settings_dict:
            return
        else:
            exps = settings_dict["expressions"]
            exps_dict = {}
            for exp in exps:
                kv = exp.split(":::")
                exps_dict[kv[0]] =kv[1] 
            settings = Settings(settings_dict["excelName"],settings_dict["excelSheetName"],settings_dict["excelKeyName"],
                    settings_dict["skipLines"],settings_dict["shpName"],settings_dict["shpKeyName"],exps_dict)
            self.initSyncer(settings)


    def writeSettings(self,doc):
        if self.syncer is None:
            return
        settings  = self.syncer.s._asdict()
        settings["expressions"] = [ "{}:::{}".format(k,v) for k,v in settings["expressions"].iteritems()]
        ProjectHandler.writeSettings("SHPSYNC",settings)

    # noinspection PyMethodMayBeStatic
    def tr(self, message):
        """Get the translation for a string using Qt translation API.

        We implement this ourselves since we do not inherit QObject.

        :param message: String for translation.
        :type message: str, QString

        :returns: Translated version of message.
        :rtype: QString
        """
        # noinspection PyTypeChecker,PyArgumentList,PyCallByClass
        return QCoreApplication.translate('shpsync', message)


    def add_action(
        self,
        icon_path,
        text,
        callback,
        enabled_flag=True,
        add_to_menu=True,
        add_to_toolbar=True,
        status_tip=None,
        whats_this=None,
        parent=None):

        icon = QIcon(icon_path)
        action = QAction(icon, text, parent)
        action.triggered.connect(callback)
        action.setEnabled(enabled_flag)

        if status_tip is not None:
            action.setStatusTip(status_tip)

        if whats_this is not None:
            action.setWhatsThis(whats_this)

        if add_to_toolbar:
            self.toolbar.addAction(action)

        if add_to_menu:
            self.iface.addPluginToMenu(
                self.menu,
                action)

        self.actions.append(action)

        return action

    def initGui(self):
        """Create the menu entries and toolbar icons inside the QGIS GUI."""

        icon_path = ':/plugins/shpsync/icon.png'
        self.add_action(
            icon_path,
            text=self.tr(u'Set up ShpSync'),
            callback=self.run,
            parent=self.iface.mainWindow())


    def unload(self):
        """Removes the plugin menu item and icon from QGIS GUI."""
        for action in self.actions:
            self.iface.removePluginMenu(
                self.tr(u'&ShpSync'),
                action)
            self.iface.removeToolBarIcon(action)
        # remove the toolbar
        del self.toolbar


    def run(self):
        """Run method that performs all the real work"""
        # show the dialog
        if self.dlg is not None:
            del self.dlg
        self.dlg = shpsyncDialog()
        self.dlg.buttonBox.accepted.connect(self.parseSettings)
        self.dlg.buttonBox.rejected.connect(self.hideDialog)
        self.dlg.exps[0].setField("x($geometry)")
        self.dlg.exps[1].setField("y($geometry)")
        self.dlg.exps[2].setField("area($geometry)")
        self.dlg.show()


    def parseSettings(self):
        exps=self.dlg.getExpressionsDict()
        excelName = self.dlg.comboBox_slave.currentText()
        excelKeyName = self.dlg.comboBox_slave_key.currentText()
        shpName = self.dlg.comboBox_master.currentText()
        shpKeyName = self.dlg.comboBox_master_key.currentText()
        excelSheetName  = self.dlg.lineEdit_sheetName.text()
        skipLines = self.dlg.spinBox.value() 
        s = Settings(excelName,excelSheetName,excelKeyName,skipLines,shpName,shpKeyName,exps)
        self.initSyncer(s)
        self.hideDialog()

    def initSyncer(self, settings):
        if self.syncer is not None:
            del self.syncer
        self.syncer = Syncer(settings)

    def hideDialog(self):
        self.dlg.hide()

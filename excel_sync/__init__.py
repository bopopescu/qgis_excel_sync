# -*- coding: utf-8 -*-
"""
/***************************************************************************
 qgis_excel_sync
                                 A QGIS plugin
 description
                             -------------------
        begin                : 2016-01-24
        copyright            : (C) 2016 by OpenGis.ch
        email                : info@opengis.ch
        git sha              : $Format:%H$
 ***************************************************************************/

/***************************************************************************
 *                                                                         *
 *   This program is free software; you can redistribute it and/or modify  *
 *   it under the terms of the GNU General Public License as published by  *
 *   the Free Software Foundation; either version 2 of the License, or     *
 *   (at your option) any later version.                                   *
 *                                                                         *
 ***************************************************************************/
 This script initializes the plugin, making it known to QGIS.
"""

import os
import sys

plugin_libs_path = os.path.join(os.path.dirname(__file__), 'libs')
for file in os.listdir(plugin_libs_path):
    if file.endswith('.whl'):
        sys.path.append(os.path.join(plugin_libs_path, file))


# noinspection PyPep8Naming
def classFactory(iface):  # pylint: disable=invalid-name
    """Load qgis_excel_sync class from file qgis_excel_sync.

    :param iface: A QGIS interface instance.
    :type iface: QgsInterface
    """
    #
    from .qgis_excel_sync import qgis_excel_sync
    return qgis_excel_sync(iface)

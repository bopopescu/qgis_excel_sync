# -*- coding: utf-8 -*-

from qgis.core import QgsProject
from qgis.PyQt.QtCore import QObject


class ProjectHandler(QObject):
    """This class manages reading from and writing to the QgsProject instance.
    It's not aware of the context of the variables written/read.
    Variables read from a file have to be put into context by the calling
    class."""

    @classmethod
    def writeSettings(cls, tag, settings):
        """write the list of settings to QgsProject instance"""
        for (key, value) in settings.items():
            cls.writeSetting(tag, key, value)

    @classmethod
    def writeSetting(cls, tag, attribute, value):
        """write plugin settings to QgsProject instance"""

        if(type(value) == float):
            QgsProject.instance().writeEntryDouble(tag, attribute, value)
        else:
            QgsProject.instance().writeEntry(tag, attribute, value)

    @classmethod
    def readSetting(cls, tag, attribute, func):
        """read a plugin setting from QgsProject instance"""
        value, ok = func(tag, attribute)
        if ok:
            return value
        else:
            return None

    @classmethod
    def readSettings(cls, tag, metasettings):
        """read plugin settings from QgsProject instance
        :param settings: a dictionary of setting names mapped to the expected
        type"""
        prj = QgsProject.instance()

        # use QProjects functions to extract the settings from the project XML
        type_to_read_function_mapping = {
            str: prj.readEntry,
            int: prj.readNumEntry,
            float: prj.readDoubleEntry,
            bool: prj.readBoolEntry,
            list: prj.readListEntry,
        }

        settings = {}
        for setting_name, (type, default) in metasettings.items():

            try:
                setting_value = cls.readSetting(
                    tag, setting_name, type_to_read_function_mapping[type])
                if setting_value is None:
                    if default is not None:
                        setting_value = default
                    else:
                        raise Exception
                settings[setting_name] = setting_value
            except Exception:
                pass
        return settings

from qgis.core import QgsRasterLayer, Qgis, QgsProject

from logging import warn
"""File that has most helper functions that interface with the QGIS API"""


def getAllJoinIdsOfLayer(layer):
    if not hasattr(layer, 'vectorJoins'):
        # open layers don't have this
        return set()
    return set(map(lambda x: x.joinLayerId, layer.vectorJoins()))


def getVersion():
    return Qgis.QGIS_VERSION_INT


def getAllJoinedLayers(layerIds):
    """get the ids of the layers that are joined on the given layerIds"""
    allJoined = set()
    allLayers = QgsProject.instance().mapLayers()
    for (id, layer) in allLayers.items():
        if isRaster(layer):
            continue
        if id in layerIds:  # let's see what the given layers are joined on
            allJoined |= getAllJoinIdsOfLayer(layer)
        else:  # let's see if the other layers join with the given layers
            joinsOfCurrentLayer = getAllJoinIdsOfLayer(layer)
            if len(joinsOfCurrentLayer & layerIds) > 0:
                allJoined.add(id)

    return allJoined


def getLayerAttributes(layerId):
    try:
        layer = QgsProject.instance().mapLayers()[layerId]
        fieldmap = layer.pendingFields()
        return fieldmap
    except KeyError:
        # OpenLayers, Raster layers don't work with this
        warn("Could not get attributes of layer {}".format(layerId))
        return None


def getAllLayerIds(filter_func):
    res = []
    for (id, layer) in QgsProject.instance().mapLayers().items():
        if filter_func(layer):
            res.append(id)
    return res


def isRaster(layer):
    return type(layer) == QgsRasterLayer


def doesLayerNameExist(name):
    return getIdFromLayerName(name) is not None


def getIdFromLayerName(layerName):
    # Important: If multiple layers with same name exist, it will return the
    # first one it finds
    for (id, layer) in QgsProject.instance().mapLayers().items():
        if layer.name() == layerName:
            return id
    return None


def getLayerFromLayerName(layerName):
    # Important: If multiple layers with same name exist, it will return the
    # first one it finds
    for (id, layer) in QgsProject.instance().mapLayers().items():
        if layer.name() == layerName:
            return layer
    return None


def getNameFromLayerId(layerId):
    layer = QgsProject.instance().mapLayers()[layerId]
    return layer.name()

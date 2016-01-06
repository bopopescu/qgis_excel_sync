from sets import Set
from datetime import datetime

from qgis._core import QgsMessageLog, QgsMapLayerRegistry, QgsFeatureRequest, QgsFeature
from qgis.utils import iface
from PyQt4.QtCore import QFileSystemWatcher
from PyQt4 import QtGui

def layer_from_name(layerName):
    # Important: If multiple layers with same name exist, it will return the first one it finds
    for (id, layer) in QgsMapLayerRegistry.instance().mapLayers().iteritems():
        if unicode(layer.name()) == layerName:
            return layer
    return None

# configurable
logTag="OpenGIS" # in which tab log messages appear
# excel layer
excelName="Excel"# the layer name
excelSheetName="Tabelle1"
excelFkIdx = 1
excelCentroidIdx = 70
excelAreaIdx = 8
excelPath=layer_from_name(excelName).publicSource()
excelKeyName = [f for f in layer_from_name(excelName).getFeatures()][0].fields().at(excelFkIdx).name()
# shpfile layer
shpName="Massnahmepool"
shpKeyName="ef_key"

# non configurable - no edits beyond this point
skipFirstLineExcel = True

# state variables
filewatcher=None
shpAdd = []
shpChange = {}
shpRemove = Set([])


def reload_excel():
    path = excelPath
    layer = layer_from_name(excelName)
    import os
    fsize=os.stat(excelPath).st_size
    info("fsize "+str(fsize))
    if fsize==0:
        info("File empty. Won't reload yet")
        return
    layer.dataProvider().forceReload()

def showWarning(msg):
    QtGui.QMessageBox.information(iface.mainWindow(),'Warning',msg)


def get_fk_set(layerName, fkName, skipFirst=True, fids=None, useProvider=False):
    layer = layer_from_name(layerName)
    freq = QgsFeatureRequest()
    if fids is not None:
        freq.setFilterFids(fids)
    if not useProvider:
        feats = [f for f in layer.getFeatures(freq)]
    else:
        feats = [f for f in layer.dataProvider().getFeatures(freq)]
    fkSet = []
    first=True
    for f in feats:
        if skipFirst and first:
            first=False
            continue
        fk = f.attribute(fkName)
        fkSet.append(fk)
    return fkSet

def info(msg):
    QgsMessageLog.logMessage(str(msg), logTag, QgsMessageLog.INFO)

def warn(msg):
    QgsMessageLog.logMessage(str(msg), logTag)
    showWarning(str(msg))

def error(msg):
    QgsMessageLog.logMessage(str(msg), logTag, QgsMessageLog.CRITICAL)

def excel_changed():
    info("Excel changed on disk - need to sync")
    reload_excel()
    update_shp_from_excel()

def get_max_id():
    layer = layer_from_name(shpName)
    # FIXME dont hardcode 0
    vals = layer.uniqueValues(0) # it returns those as strings :(
    return max(map(long,vals))

def added_geom(layerId, feats):
    # unused for now
    info("added feats "+str(feats))
    layer = layer_from_name(shpName)
    maxFk = get_max_id()
    for i,_ in enumerate(feats):
        _id = maxFk + i + 1
        feats[i].setAttribute(shpKeyName,_id)
        res = layer.changeAttributeValue(feats[i].id(),0,str(_id))
        info("managed to update?"+str(res)+" layer editable?" +str(layer.isEditable()) )

    global shpAdd
    shpAdd = feats


def added_geom_precommit(fid):
    # TODO: Unsure what happens if adding + deleting without commiting
    #info("precomit fid"+str(fid))
    global shpAdd
    layer = layer_from_name(shpName)
    maxFk = get_max_id()
    _id = maxFk + len(shpAdd) + 1
    freq = QgsFeatureRequest()
    freq.setFilterFids([fid])
    feat = list(layer.getFeatures(freq))[0]

    # unfortunately the way/order signals work, 
    # it may be attempted to add something twice, so check that geometries are unique
    for f in shpAdd:
        #info("shpadd:"+str(f.geometry().exportToWkt())+" vs "+str(feat.geometry().exportToWkt()))
        # plain geometry equality check seems to not work
        if f.geometry().exportToWkt() == feat.geometry().exportToWkt():
            #info("will skip this one")
            return

    feat.setAttribute(shpKeyName,_id)
    res = layer.changeAttributeValue(feat.id(),0,str(_id)) # TODO: Don't hardcode 0
    #info("managed to update?"+str(res)+" layer editable?" +str(layer.isEditable()) )
    shpAdd.append(feat)


def removed_geom_precommit(fids):
    #info("Removed fids"+str(fids))
    fks_to_remove = get_fk_set(shpName,shpKeyName,skipFirst=False,fids=fids,useProvider=True)
    global shpRemove
    shpRemove = Set(fks_to_remove)
    info("feat ids to remove"+str(shpRemove))

def changed_geom(layerId, geoms):
    fids = geoms.keys()
    freq = QgsFeatureRequest()
    freq.setFilterFids(fids)
    feats = list(layer_from_name(shpName).getFeatures(freq))
    fks_to_change = get_fk_set(shpName,shpKeyName,skipFirst=False,fids=fids)
    global shpChange
    shpChange = {k:v for (k,v) in zip(fks_to_change, feats)}
    #info("changed"+str(shpChange))


def write_feature_to_excel(sheet, idx, feat):
   area = feat.geometry().area()*0.0001 # Square meters to hectare
   centroid = str(feat.geometry().centroid().asPoint())
   sheet.write(idx, excelFkIdx, feat[shpKeyName])
   sheet.write(idx, excelCentroidIdx, centroid)
   sheet.write(idx, excelAreaIdx, area)

def write_rowvals_to_excel(sheet, idx, vals, ignore=None):
    if ignore is None:
        ignore = []
    for i,v in enumerate(vals):
        if i not in ignore:
            sheet.write(idx,i,v)

def update_excel_programmatically():

    from xlrd import open_workbook
    import xlwt

    rb = open_workbook(excelPath,formatting_info=True)
    r_sheet = rb.sheet_by_name(excelSheetName) # read only copy
    wb = xlwt.Workbook()
    w_sheet = wb.add_sheet(excelSheetName, cell_overwrite_ok=True)
    write_idx = 0

    maxFk = 0
    for row_index in range(r_sheet.nrows):
        if row_index==0: # copy header
            vals = r_sheet.row_values(row_index)
            write_rowvals_to_excel(w_sheet, write_idx, vals)
            write_idx+=1
            continue         
        #print(r_sheet.cell(row_index,1).value)
        fk = r_sheet.cell(row_index, excelFkIdx).value
        maxFk = max(maxFk, long(fk))
        fk = str(long(fk)) # FIXME: Why do we have the keys as strings from the shp?
        if fk in shpRemove:
            continue
        if fk in shpChange.keys():
            shpf = shpChange[fk]
            write_feature_to_excel(w_sheet, write_idx, shpf)
            vals = r_sheet.row_values(row_index)
            write_rowvals_to_excel(w_sheet, write_idx, vals,
                    ignore=[excelCentroidIdx, excelAreaIdx])
        else:# else just copy the row
            vals = r_sheet.row_values(row_index)
            write_rowvals_to_excel(w_sheet, write_idx, vals)

        write_idx+=1


    for shpf in shpAdd:
        write_feature_to_excel(w_sheet, write_idx, shpf)
        write_idx+=1

    wb.save(excelPath)


def update_excel_from_shp():
    info("Will now update excel from edited shapefile")
    info("changing:"+str(shpChange))
    info("adding:"+str(shpAdd))
    info("removing"+str(shpRemove))
    update_excel_programmatically()
    global shpAdd
    global shpChange
    global shpRemove
    shpAdd = []
    shpChange = {}
    shpRemove = Set([])


def updateShpLayer(fksToRemove):
    layer = layer_from_name(shpName)
    feats = [f for f in layer.getFeatures()]
# MK, 3.1.2015: Not sure this should be done without user confirmation.
#    layer.startEditing()
#    for f in feats:
#         if f.attribute(shpKeyName) in fksToRemove:
#             layer.deleteFeature(f.id())
#    layer.commitChanges()

def update_shp_from_excel():
    excelFks = Set(get_fk_set(excelName, excelKeyName,skipFirst=skipFirstLineExcel))
    if not excelFks:
        warn("Qgis thinks that the Excel file is empty. That probably means something went horribly wrong. Won't sync.")
        return
    shpFks = Set(get_fk_set(shpName,shpKeyName,skipFirst=False))
    # TODO also special warning if shp layer is in edit mode
    info("Keys in excel"+str(excelFks))
    info("Keys in shp"+str(shpFks))
    if shpFks==excelFks:
        info("Excel and Shp layer have the same rows. No update necessary")
        return
    inShpButNotInExcel = shpFks - excelFks
    inExcelButNotInShp = excelFks - shpFks
    if inExcelButNotInShp:
         warn("There are rows in the excel file with no matching geometry {}.".format(inExcelButNotInShp))
    if inShpButNotInExcel:
        info("Will remove features "+str(inShpButNotInExcel)+"from shapefile because they have been removed from excel")
        updateShpLayer(inShpButNotInExcel)

def init():
    info("Initial Syncing excel to shp")
    update_shp_from_excel()
    global filewatcher # otherwise the object is lost
    filewatcher = QFileSystemWatcher([excelPath])
    filewatcher.fileChanged.connect(excel_changed)
    shpLayer = layer_from_name(shpName)
    shpLayer.featureAdded.connect(added_geom_precommit)
    #shpLayer.committedFeaturesRemoved.connect(removed_geom)
    shpLayer.featuresDeleted.connect(removed_geom_precommit)
    shpLayer.committedGeometriesChanges.connect(changed_geom)
    shpLayer.editingStopped.connect(update_excel_from_shp)

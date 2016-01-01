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
excelName="Beispiel"
excelFkIdx = 0
excelCentroidIdx = 14
excelAreaIdx = 8
excelPath=layer_from_name(excelName).publicSource()
# shpfile layer
shpName="Beispiel_Massnahmepool"
shpKeyName="ef_key"


# state variables 
filewatcher=None
shpAdd = {}
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


def get_fk_set(layerName, fkName, skipFirst=True, fids=None):
    layer = layer_from_name(layerName)
    freq = QgsFeatureRequest()
    if fids is not None:
        freq.setFilterFids(fids)
    feats = [f for f in layer.getFeatures(freq)]
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
    info("Excel changed in disk - need to sync")
    reload_excel()
    update_shp_from_excel()

def added_geom(layerId, feats):
    info("added feats "+str(feats))
    fks_to_add = [feat.attribute(shpKeyName) for feat in feats]
    global shpAdd
    shpAdd = {k:v for (k,v) in zip(fks_to_add, feats)}


def removed_geom(layerId, fids):
    fks_to_remove = get_fk_set(shpName,shpKeyName,skipFirst=False,fids=fids)
    global shpRemove
    shpRemove = Set(fks_to_remove)

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
   area = str(feat.geometry().area())
   centroid = str(feat.geometry().centroid().asPoint())
   for i in range(len(feat.fields().keys())):
       sheet.write(idx,i, feat.attribute(i))
   sheet(write, idx, excelCentroidIdx, centroid)
   sheet(write, idx, excelAreaIdx, area)

def write_rowvals_to_excel(sheet, idx, vals):
    for i,v in enumerate(vals):
        sheet.write(idx,i,v)

def update_excel_programmatically():

    from xlutils.copy import copy # http://pypi.python.org/pypi/xlutils
    from xlrd import open_workbook # http://pypi.python.org/pypi/xlrd
    from xlwt import easyxf # http://pypi.python.org/pypi/xlwt

    rb = open_workbook(excelPath,formatting_info=True)
    r_sheet = rb.sheet_by_index(0) # read only copy
    wb = xlwt.Workbook()
    w_sheet = wb.add_sheet(0, cell_overwrite_ok=True)
    
    for row_index in range(r_sheet.nrows):
        #print(r_sheet.cell(row_index,1).value)
        fk = r_sheet.cell(row_index, excelFkIdx).value
        if fk in shpRemove:
            continue
        if fk in shpChange.keys():
            shpf = shpChange[key]
            write_feature_to_excel(w_sheet, write_index, shpf)
            vals = r_sheet.row_values(row_index)
            write_rowvals_to_excel(w_sheet, write_idx, vals)
           
        else:# else just copy the row
            vals = r_sheet.row_values(row_index)
            write_rowvals_to_excel(w_sheet, write_idx, vals)
       
        write_index+=1
         
         
    for key in shpAdd.keys():
       shpf = shpAdd[key]
       write_feature_to_excel(sheet, write_index, shpf)
       write_idx+=1

    wb.save(excelPath+"_new") #TODO fix after testing


def update_excel_from_shp():
    info("Will now update excel from edited shapefile")
    info("changing:"+str(shpChange))
    info("adding:"+str(shpAdd))
    info("removing"+str(shpRemove))
    update_excel_programmatically()
    global shpAdd
    global shpChange
    global shpRemove
    shpAdd = {}
    shpChange = {}
    shpRemove = Set([]) 


def updateShpLayer(fksToRemove):
    layer = layer_from_name(shpName)
    feats = [f for f in layer.getFeatures()]
    layer.startEditing()
    for f in feats:
         if f.attribute(shpKeyName) in fksToRemove:
             layer.deleteFeature(f.id())
    layer.commitChanges()
     

def update_shp_from_excel():
   
    excelFks = Set(get_fk_set(excelName, excelKeyName,skipFirst=True))
    if not excelFks:
        warn("Qgis thinks that the Excel file is empty. That probably means something went horribly wrong. Won't sync.")
        return
    shpFks = Set(get_fk_set(shpName,shpKeyName,skipFirst=False))
    # TODO somewhere here I should refresh the join
    # TODO also special warning if shp layer is in edit mode
    info("Keys in excel"+str(excelFks))
    info("Keys in shp"+str(shpFks))
    if shpFks==excelFks:
        info("Excel and Shp layer have the same rows. No update necessary")
        return
    inShpButNotInExcel = shpFks - excelFks
    inExcelButNotInShp = excelFks - shpFks
    if inExcelButNotInShp:
         warn("There are rows in the excel file with no matching geometry {}. Can't update shapefile from those.".format(inExcelButNotInShp))
    if inShpButNotInExcel:
        info("Will remove features "+str(inShpButNotInExcel)+"from shapefile because they have been removed from excel")
        updateShpLayer(inShpButNotInExcel)

def init(filename):
    info("Initial Syncing excel to shp")
    update_shp_from_excel()
    global filewatcher # otherwise the object is lost
    filewatcher = QFileSystemWatcher([filename])
    filewatcher.fileChanged.connect(excel_changed)
    shpLayer = layer_from_name(shpName)
    shpLayer.committedFeaturesAdded.connect(added_geom)
    shpLayer.committedFeaturesRemoved.connect(removed_geom)
    shpLayer.committedGeometriesChanges.connect(changed_geom)
    shpLayer.editingStopped.connect(update_excel_from_shp)

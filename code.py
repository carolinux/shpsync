from sets import Set
from datetime import datetime

from qgis._core import QgsMessageLog, QgsMapLayerRegistry, QgsFeatureRequest
from PyQt4.QtCore import QFileSystemWatcher

filewatcher=None
logTag="OpenGIS"
excelName="Beispiel"
excelKeyName="Field1"
shpName="Beispiel_Massnahmepool"
shpKeyName="ef_key"

shpFks = Set([])

def reload_excel():
    layer = layer_from_name(excelName)
    layer.dataProvider().forceReload()

def get_fk_set(layerName, fkName, skipFirst=True, fids=None):
    layer = layer_from_name(layerName)
    freq = QgsFeatureRequest()
    if fids is not None:
        freq.setFilterFids(fids)
    feats = [f for f in layer.getFeatures(freq)]
    fkSet = Set([])
    first=True
    for f in feats:
        if skipFirst and first:
            first=False
            continue
        fk = f.attribute(fkName)
        fkSet.add(fk)
    return fkSet
        

def layer_from_name(layerName):
    # Important: If multiple layers with same name exist, it will return the first one it finds
    for (id, layer) in QgsMapLayerRegistry.instance().mapLayers().iteritems():
        if unicode(layer.name()) == layerName:
            return layer
    return None


def info(msg):
    QgsMessageLog.logMessage(str(msg), logTag, QgsMessageLog.INFO)

def warn(msg):
    QgsMessageLog.logMessage(str(msg), logTag)

def error(msg):
    QgsMessageLog.logMessage(str(msg), logTag, QgsMessageLog.CRITICAL)

def excel_changed():
    info("Excel changed in disk")
    reload_excel()
    update_shp_from_excel()
    # TODO update shp from excel
    # refresh the join also..

def added_geom(layerId, fids):
    fks_to_add = get_fk_set(shpName,shpKeyName,skipFirst=False,fids=fids)

def removed_geom(layerId, fids):
    fks_to_remove = get_fk_set(shpName,shpKeyName,skipFirst=False,fids=fids)

def changed_geom(layerId, geoms):
    fids = geoms.keys()
    fks_to_change = get_fk_set(shpName,shpKeyName,skipFirst=False,fids=fids)
    info("changed"+str(fids))

def update_excel_from_shp():
    pass

def update_shp_from_excel():
    info("Excel updated. Need to edit shapefile accordingly!")
    excelFks = get_fk_set(excelName, excelKeyName,skipFirst=True)
    shpFks = get_fk_set(shpName,shpKeyName,skipFirst=False)
    # TODO somewhere here I should refresh the join
    # TODO also special warning if shp layer is in edit mode
    if shpFks==excelFks:
        info("Excel and Shp layer have the same rows. No update necessary")
        return
    inShpButNotInExcel = shpFks - excelFks
    inExcelButNotInShp = excelFks - shpFks
    if inExcelButNotInShp:
         warn("There are rows in the excel file with no matching geometry. Can't update shapefile from those.")
    if inShpButNotInExcel:
        warn("Will remove features "+str(inShpButNotInExcel)+"from shapefile because they have been removed from excel") 

def handle_connections(filename):
    global filewatcher # otherwise the object is lost
    filewatcher = QFileSystemWatcher([filename])
    filewatcher.fileChanged.connect(excel_changed)
    shpLayer = layer_from_name(shpName)
    shpLayer.committedFeaturesAdded.connect(added_geom)
    shpLayer.committedFeaturesRemoved.connect(removed_geom)
    shpLayer.committedGeometriesChanges.connect(changed_geom)
    shpLayer.editingStopped.connect(update_excel_from_shp)




handle_connections("/home/carolinux/Projects/Beispiel/Mappe2.xlsx")

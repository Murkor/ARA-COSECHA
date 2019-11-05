import arcpy
from   scipy.spatial import ConvexHull, convex_hull_plot_2d
import os
import sys
import numpy as np
import datetime

eco = True
#------------------------------
def eliminarObjeto(objeto):
#------------------------------
    if arcpy.Exists(objeto):
        arcpy.Delete_management(objeto)

#------------------------------     
def imprimir(linea):
#------------------------------
    print(linea)
    arcpy.AddMessage(linea)
#-----------------------------------------
def ClearSelection(fc):
#-----------------------------------------
    nombreSeleccion = "Limpiar"
    
    if not arcpy.Exists(nombreSeleccion):
       arcpy.MakeFeatureLayer_management(fc, nombreSeleccion)
    arcpy.SelectLayerByAttribute_management(nombreSeleccion, "CLEAR_SELECTION")

#-------------------------------------------------------------------------------------------------------------------------
def seleccionar(fc, expresion, nombreSeleccion="Seleccion", tipoSel="NEW_SELECTION"):
#------------------------------------------------------------------------------------------------------------------------
    eliminarObjeto(nombreSeleccion)
    if not arcpy.Exists(nombreSeleccion):
       arcpy.MakeFeatureLayer_management(fc, nombreSeleccion)
    if eco:
        imprimir("SELECCION SOBRE FC="+fc+" expresion ="+expresion)
    arcpy.SelectLayerByAttribute_management(nombreSeleccion, tipoSel, expresion)
    total = int(arcpy.GetCount_management(nombreSeleccion).getOutput(0))
    if eco:
        imprimir("Seleccion lista")
    return nombreSeleccion, total

#---------------------------------------------
def coordenadas(in_fc):
# PERMITE CREAR ARRAY CON TODOS LOS VERTICES DE
# TODOS LOS ELEMENTOS DE UN FC
#--------------------------------------------- 
    SR = arcpy.Describe(in_fc).spatialReference

    ##sdf = arcgis.SpatialDataFrame
    ##a_sdf = sdf.from_featureclass(in_fc,
    ##                            sql_clause=None,
    ##                            where_clause=None,
    ##                            sr=SR)
    ##a_rec = sdf.to_records(a_sdf)  # record array
    a_np = arcpy.da.FeatureClassToNumPyArray(in_fc,field_names=["SHAPE@XY"],spatial_reference=SR,explode_to_points=True)
    ##a_np2 = fc_g(in_fc)
    ##return sdf, a_sdf, a_rec, a_np, a_np2
    print(type(a_np))
    print(len(a_np))
    s = []
    for p in a_np:
        s.append(p[0])
        
    return s, SR

def cHull(puntos, outFC, SR):
    
    hull = ConvexHull(puntos)
    #print(type(hull))
    v = hull.vertices

    lista=[]
    n=1
    for i in v:
        lista.append((n, puntos[i].tolist()))
        if n==1:
            guardar=(n, puntos[i].tolist())
        n +=1
    lista.append(guardar)
    print (lista)
    b = np.dtype([('idfield',np.int32),('XY', '<f8', 2)])            
    x = np.array(lista,b)
    print(outFC)
    outFC = WS + os.path.sep + outFC
    print(SR)
    arcpy.da.NumPyArrayToFeatureClass(x, outFC, ['XY'], SR)    


def Procesar(ws, fc):
    outFC = fc+"HULL"
    tmp   = "cHull"
    arcpy.env.workspace = ws
    arcpy.env.overwriteoutput = True
    ClearSelection(fc)
    layer, total = seleccionar(fc, "CODIGO = 'T524' AND OBJECTID < 100")
    print(layer)
    print(total)
    puntos,sr  = coordenadas(layer)
    s = []
    for x in puntos:
        print(x)
        print(type(x))
        print(x[0])
        s.append((x[0],x[1]))
    print(s)
    if arcpy.Exists(tmp):
        arcpy.Delete_management(tmp)
    cHull(np.array(s), tmp, sr)
    
WS    = r"c:\temp\AVL_MAQUINAS.gdb"
FC    = "MAQUINASUTM"

Procesar (WS, FC)






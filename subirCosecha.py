import os
import arcpy


#------------------------------     
def imprimir(linea):
#------------------------------
    print(linea)
    arcpy.AddMessage(linea)
    
#-------------------------------    
def lToc(mxd,fc,groupLayer=""):
    
#------------------------------
    fc1 = fc.replace(".lyr","")
    df        = arcpy.mapping.ListDataFrames(mxd, "*")[0]
    imprimir(tmpWS + os.path.sep+fc)
    addLayer = arcpy.mapping.Layer(tmpWS + os.path.sep+fc1)
    #arcpy.mapping.AddLayer(df, addLayer, "BOTTOM")
    el = layerDir+os.path.sep+fc
 
    if os.path.exists(el):
       arcpy.ApplySymbologyFromLayer_management (addLayer, el)
       arcpy.mapping.AddLayer(df, addLayer, "BOTTOM")
    else:
        imprimir("NO EXISTE -->"+layerDir+os.path.sep+fc)


if __name__ == '__main__':
    cws = arcpy.GetParameterAsText(0)
    tmpWS = arcpy.GetParameterAsText(1)
    #cws = r'C:\D\proyectos\2019\ARAUCO_CARTO\actasCosecha\ACTAS_CIERRE_GDBSDE.gdb'
    #tmpWS = r"C:\temp\actas.gdb"
    layerDir = os.path.dirname(cws)+ os.path.sep+ "temas"
    oldWS = arcpy.env.workspace
    arcpy.env.workspace = tmpWS

    mx = arcpy.mapping.MapDocument("CURRENT")
    for root, dirs, files in os.walk(layerDir):
        for archivo in files:
            imprimir(files)
            if archivo.endswith(".lyr"):
                imprimir(archivo)
                imprimir(os.path.basename(archivo))
                lToc(mx, os.path.basename(archivo))
               
    

    

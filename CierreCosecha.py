import arcpy
import os
import sys
import shutil
import tempfile


import time
from   xlrd import open_workbook, cellname
import numpy as np


global cws
iso ='utf-8'
eco = False
cws  = ""
cNR  = "Cosecha_No_Realizada"
cosNR= "cosNR_lyr"
AC   = "ACTAS_CERRADAS"
RMA  = "REMANENTE_PORTAL"
nac  = "NUM_ACTA_N"
uOracle = "SDEUSER1"
oRMA    = uOracle + "." + RMA

#-------------------------------------------------
cElim =[
    "field", "field_1","field_1_2","field_1__3"
    ]

cElim1 =["_ZONA","_MES_CIERRE"]
#------------------------------  
def DetectaActasCerradas(testws):
#------------------------------
    curd = arcpy.env.workspace
    arcpy.env.workspace = testws
    atrs = [nac, 'MES_CIERRE','YEAR_CIERRE']
    imprimir("\nDETECTANDO EXISTENCIA DE "+AC+ " en "+ testws)
    if not arcpy.Exists(AC):
        imprimir("CREANDO "+AC+"....")
        arcpy.CreateTable_management(arcpy.env.workspace,AC)
    lista = [f.name for f in arcpy.ListFields(AC)]
    if not (atrs[0] in lista):
        arcpy.AddField_management(AC,atrs[0],"LONG")
    if not (atrs[1] in lista):  
        arcpy.AddField_management(AC,atrs[1],"TEXT",12)
    if not (atrs[2] in lista):
        arcpy.AddField_management(AC,atrs[2],"SHORT")
    arcpy.env.workspace=curd
        
#------------------------------     
def imprimir(linea):
#------------------------------
    print(linea)
    arcpy.AddMessage(linea)
#-------------------------------------------------
def eliminarObjeto(objeto):
#-------------------------------------------------
    if arcpy.Exists(objeto):
        arcpy.Delete_management(objeto)

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
#-------------------------------------------------
def InicioFin():
#-------------------------------------------------
     finActas = int(time.strftime('%Y'))+1
     comienzoActas = finActas - 2
     return comienzoActas, finActas
#-------------------------------------------------
def listaElim(lista,lt):
#-------------------------------------------------
    lr=[]

    for camp in lista:
        campo = camp.upper()
        a = campo.startswith(lt) and campo != lt+cElim1[0] and campo != lt+cElim1[1]
        if a:
                lr.append(camp)
    return lr
#-------------------------------------------------
def quitaCampos(ws,tabla):
#-------------------------------------------------

    lt = "T_CC_"+tabla[-4:len(tabla)]
    lr=[]
    lista = [f.name for f in arcpy.ListFields(tabla)]
    lr = listaElim(lista,lt)
    if eco:
        imprimir("ADD & CALCULATE, MES_CIERRE, YEAR_CIERRE, ZONA_CIERRE")
    
    arcpy.AddField_management(tabla,"MES_CIERRE","TEXT",12)
    arcpy.AddField_management(tabla,"YEAR_CIERRE","SHORT")
    arcpy.AddField_management(tabla,"ZONA_CIERRE","TEXT",20)
    arcpy.CalculateField_management(tabla,'MES_CIERRE',"!"+lt+"_Mes_Cierre!",'PYTHON')
    arcpy.CalculateField_management(tabla,'ZONA_CIERRE',"!"+lt+"_Zona!",'PYTHON')
    agno = int(tabla[-4:len(tabla)])
    arcpy.CalculateField_management(tabla,'YEAR_CIERRE',agno,'PYTHON')
    if eco:
        imprimir("FIN CALCULATE "+tabla)
    
    lista = [f.name.upper() for f in  arcpy.ListFields(ws+os.path.sep+tabla)]
    if lt+cElim1[0] in lista:
       lr.append(lt+cElim1[0])
    if lt+cElim1[1] in lista:
        lr.append(lt+cElim1[1])
    if len(lr)>0:
       if eco: 
          imprimir("Eliminado Columns")
          imprimir(lr)
       arcpy.DeleteField_management(tabla, lr)
       if eco:
          imprimir("LISTO ELIMINACION")
    else:
         imprimir("=========================")
         imprimir("SIN CAMPOS A ELIMINAR "+tabla)
         imprimir("=========================")
    if eco:
       imprimir("FRECUENCIA CON "+tabla+"\n"+ws+os.path.sep+"f"+tabla)
    eliminarObjeto(ws+os.path.sep+ "f"+tabla)
    arcpy.Frequency_analysis(tabla, ws+os.path.sep+ "f"+tabla,[nac,"MES_CIERRE","YEAR_CIERRE"])
    if eco:
        imprimir("LISTO FRECUENCIA CON CAMPOS")
    if arcpy.Exists("f"+tabla):
        #imprimir([f.name for f in arcpy.ListFields("f"+tabla)])
        ##arcpy.AddField_management("f"+tabla,nac,"LONG")
        ##exp = "clng(["+cNR+"_NUM_ACTA_1])"
        ##imprimir(exp)
        ##arcpy.CalculateField_management("f"+tabla,nac,exp,"VB")
        #fm= crearMapping(AC, "f"+tabla,cNR+"_NUM_ACTA_1")
        if eco:
           imprimir("APPEND..."+AC)
        arcpy.Append_management(["f"+tabla],AC, "NO_TEST")# , fm)
        if eco:
            imprimir("FIN APPEND")
 #except:
 #    lst = arcpy.GetMessages(0)
 #    imprimir(lst)
 #    imprimir("PROBLEMAS ES QUITA")
#
#==========================================================
def sacarFinalizados(tabl):
#==========================================================
    tabla = cws + os.path.sep+ tabl
    if eco:
       imprimir("Finalizados...  "+tabla+ " "+nac)
    a1 = arcpy.da.TableToNumPyArray(tabla, nac)
    imprimir('PASO 1- Detectando Actas no cosechadas')
    aa1= a1[nac]
    imprimir('PASO 2- Detectando Actas Cerradas')
    a2 = arcpy.da.TableToNumPyArray(AC, nac)
    aa2= a2[nac]
    imprimir('PASO 3- Comparando ...')
    inter= np.intersect1d(aa1, aa2)
    lista =""
    for acta in inter:
        lista = lista+","+str(acta)
    expresion =""
    
    if len(lista)==0:
        imprimir("NO SE HALLARON ACTAS CERRADAS ANTERIORES....")
    else:
       lista=lista[1:]
       imprimir(lista)
       expresion = nac+" IN ("+lista+")"
       imprimir(expresion)
       nsel, cant = seleccionar(tabl, expresion, "aElim")
       
       if cant>0:
           imprimir("ELIMINANDO ACTAS CERRADAS ANTERIORES ==> POLS ELIM="+str(cant))
           arcpy.DeleteFeatures_management(nsel)
           imprimir("Fin deletefeatures")
       else:
           imprimir("NO SE HALLARON ACTAS CERRADAS....")
       
    
    
#==========================================================    
def crearMapping(tabla1, tabla2, colu):
#----------------------------------------------------------
 #imprimir("CREAR MAPPING "+tabla2+" "+ colu + "  talba1 ="+tabla1)
 try:
    fm_type = arcpy.FieldMap()
    fms     = arcpy.FieldMappings()
    
    fm_type.addInputField(tabla2, colu)
    nombre  = fm_type.outputField
    nombre.name = "SNUM_ACTA"
    fm_type.outputField = nombre
    fms.addFieldMap(fm_type)
    return fms
 except:
     imprimir(arcpy.GetMessages())
     return []
#-------------------------------------------------   
def AgregarRemanente(ws, archivo):
#-------------------------------------------------
    arcpy.env.workspace = ws
    if not arcpy.Exists( ws + os.path.sep + RMA):
        sr = arcpy.Describe(archivo)
        sr = sr.spatialReference
        imprimir("\nCreando "+RMA+ " en "+ ws)
        arcpy.CreateFeatureclass_management(ws, RMA, "POLYGON", archivo, spatial_reference=sr)
    arcpy.Append_management([archivo],RMA,"NO_TEST")
#-------------------------------------------------
def crearDict(ws,archivo):
#-------------------------------------------------
    # 0 num acta
    # 1 zona
    # 2 mes
    # 3 intervencion
    # 4 SIP
    # 5 Predio
    arcpy.env.workspace = cws
    actas = {}
    inicio, fin = InicioFin()
    for j in range(inicio, fin):
      tabla = "T_CC_"+str(j)
      l =""
      imprimir("\n============================\nPROCESANDO AÑO "+str(j))
      #imprimir(archivo)
      #imprimir(j)
      try:
        arcpy.ExcelToTable_conversion(archivo,tabla , str(j))
        arcpy.AddField_management(tabla,"strActa","TEXT", field_length=12)
        nombres = [f.name for f in arcpy.ListFields(tabla)]
        ce =[]
        for x in cElim:
            if x in nombres:
                ce.append(x)
        if len(ce) >0:
            arcpy.DeleteField_management(tabla, ce)
        unio    = nombres[1]
        l="Calculando"
        arcpy.CalculateField_management(tabla, "strActa","cstr(["+unio+"])","VB")
        try:
            l="removiendo"
            arcpy.RemoveJoin_management(cosNR)
        except:
            pass
        nombreT = "tmp"+str(j)
        eliminarObjeto(nombreT)
        l="Add join"
        #imprimir("CREANDO JOIN CON "+cosNR + " y "+tabla)
        arcpy.AddJoin_management(cosNR,"NUM_ACTA_1",tabla,"strActa","KEEP_COMMON")
        l="Copy Features.."

        #imprimir("JOIN CREADO\nCOPIANDO "+cosNR+ " "+nombreT+"\n"+arcpy.env.workspace)
        arcpy.CopyFeatures_management(cosNR,ws + os.path.sep + nombreT)
        l="add field"
        arcpy.AddField_management(nombreT, nac, "LONG")
        l="calculate"
        exp1 = "clng(["+cNR+"_NUM_ACTA_1])"
        #imprimir(exp1)
        arcpy.CalculateField_management(nombreT,nac, exp1 ,"VB")
        l="sacar finalizado"
        sacarFinalizados(nombreT)
        l="Delete field join objectid"
        print("Creado "+nombreT)
        l ="Proc final "

    
        quitaCampos(ws,nombreT)
        imprimir("Procesando Remanente --->"+RMA+ " con "+nombreT)
        AgregarRemanente(ws, nombreT)
        
      except:
        imprimir("Problemas procesando hoja "+str(j)+" "+l)

#-------------------------------------------------
def subirAOracle(ws,destino):
#-------------------------------------------------
     finP = 1
     if ws.upper() == destino.upper():
        imprimir("PROBLEMAS ==> LOCAL WS ="+ws+ " ES IGUAL A DESTINO : "+destino)
        return finP
     if arcpy.Exists(ws+os.path.sep + RMA) and arcpy.Exists(destino+os.path.sep+oRMA):
         cant = arcpy.GetCount_management(ws+os.path.sep + RMA)
         finP = 0
         if cant >0:
            arcpy.Append_management([ws+os.path.sep + RMA], destino+os.path.sep+oRMA, "NO_TEST")
            arcpy.DeleteFeatures_management(ws+os.path.sep + RMA)
            imprimir("\n==============================\nSE AGREGRARON " + str(cant) +" POLIGONOS DE REMANENTES")
         else:
             imprimir("\n==============================\nNO HAY NUEVOS POLIGONOS PARA SUBIR A PORTAL")
            
     else:
         imprimir("\n==============================\nNO EXISTE "+ RMA +' en '+ws+" y/o en "+destino)
         
     return finP

#-------------------------------------------------
def Procesar(ws, actasCerradas, destino):
#-------------------------------------------------
    arcpy.env.workspace = ws
    arcpy.env.overwriteOutput = True
    arcpy.MakeFeatureLayer_management (ws + os.path.sep + cNR,  cosNR )
    crearDict(ws, actasCerradas)
    indicador  = subirAOracle(ws, destino)
    lista =['PROCESO FINALIZADO NORMAL', 'PROCESO FINALIZADO CON PROBLEMAS *****************']
    imprimir("\n=====================================\n"+lista[indicador])
    

if __name__ == '__main__':
    cws      = r"c:\temp\actas.gdb" #
    cws      =  arcpy.GetParameterAsText(0)
    archivoE = r"C:\D\proyectos\2019\ARAUCO_CARTO\PLANILLAS\Actas Cerradas_2019_S4_Julio.xlsx" #
    archivoE = arcpy.GetParameterAsText(1)
    destino  = arcpy.GetParameterAsText(2)
    for suf in ["","0","1","2","3","4","5","6","7"]:
            tabla = cws + os.path.sep + "UNION"+suf
            eliminarObjeto(tabla)
    #if not arcpy.Exists(destino + os.path.sep + RMA):
    #    imprimir("PROBLEMAS NO EXISTE "+ RMA +" en "+destino)
    #else:
    DetectaActasCerradas(cws)

    Procesar(cws, archivoE, destino)
        
    

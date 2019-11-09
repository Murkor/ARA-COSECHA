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
eco = True
cws  = ""
cNR  = "Cosecha_No_Realizada"
cosNR= "cosNR_lyr"
AC   = "ACTAS_CERRADAS"
nac  = "NUM_ACTA_N"


#-------------------------------------------------
cElim =[
    "field", "field_1","field_1_2","field_1__3"
    ]

cElim1 =["_ZONA","_MES_CIERRE"]
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
    #imprimir("LISTA ELIM entrando")
    #for camp in lista:
        #imprimir(camp)
    for camp in lista:
        campo = camp.upper()
        #imprimir(campo)
        a = campo.startswith(lt) and campo != lt+cElim1[0] and campo != lt+cElim1[1]
        #imprimir(str(a))
        if a:
                #imprimir("candidato hallado!!")
                lr.append(camp)
    #imprimir("SALIENDO...")
    #imprimir(lr)
    return lr
#-------------------------------------------------
def quitaCampos(ws,tabla):
#-------------------------------------------------

    #imprimir("QUITA CAMPOS")
    #imprimir(tabla)
    lt = "T_CC_"+tabla[-4:len(tabla)]
    lr=[]
    lista = [f.name for f in arcpy.ListFields(tabla)]
    lr = listaElim(lista,lt)
    

    arcpy.AddField_management(tabla,"MES_CIERRE","TEXT",10)
    arcpy.AddField_management(tabla,"YEAR_CIERRE","SHORT")
    arcpy.AddField_management(tabla,"ZONA_CIERRE","TEXT",20)
    arcpy.CalculateField_management(tabla,'MES_CIERRE',"!"+lt+"_Mes_Cierre!",'PYTHON')
    arcpy.CalculateField_management(tabla,'ZONA_CIERRE',"!"+lt+"_Zona!",'PYTHON')
    agno = int(tabla[-4:len(tabla)])
    arcpy.CalculateField_management(tabla,'YEAR_CIERRE',agno,'PYTHON')
    #imprimir("Lista de cols a Eliminar")
    lr.append(lt+cElim1[0])
    lr.append(lt+cElim1[1])
    if len(lr)>0:
       #imprimir(" a eliminar")
       #imprimir("==========================")
       #imprimir(lr)
       arcpy.DeleteField_management(tabla, lr)
       #imprimir("LISTO ELIMINACION")
    else:
         imprimir("=========================")
         imprimir("SIN CAMPOS A ELIMINAR")
         imprimir("=========================")
    imprimir("FRECUENCIA CON "+tabla+"\n"+ws+os.path.sep+"f"+tabla)
    imprimir([f.name for f in arcpy.ListFields(tabla)])
    eliminarObjeto(ws+os.path.sep+ "f"+tabla)
    arcpy.Frequency_analysis(tabla, ws+os.path.sep+ "f"+tabla,["NUM_ACTA_N","MES_CIERRE","YEAR_CIERRE"])
    if arcpy.Exists("f"+tabla):
        arcpy.AddField_management("f"+tabla,"NUM_ACTA","LONG")
        exp = "clng(["+cNR+"_NUM_ACTA_1])"
        imprimir(exp)
        arcpy.CalculateField_management("f"+tabla,"NUM_ACTA",exp,"VB")
        #fm= crearMapping(AC, "f"+tabla,cNR+"_NUM_ACTA_1")
        arcpy.Append_management(["f"+tabla],AC, "NO_TEST")# , fm)
        #imprimir("FIN APPEND")
 #except:
 #    lst = arcpy.GetMessages(0)
 #    imprimir(lst)
 #    imprimir("PROBLEMAS ES QUITA")
#
#==========================================================
def sacarFinalizados(tabl):
#==========================================================
    tabla = cws + os.path.sep+ tabl
    imprimir("finalizados...."+tabla+ " "+nac)
    a1 = arcpy.da.TableToNumPyArray(tabla, nac)
    imprimir('1')
    aa1= a1["NUM_ACTA_N"]
    imprimir('2')
    a2 = arcpy.da.TableToNumPyArray(AC, "NUM_ACTA")
    aa2= a2["NUM_ACTA"]
    imprimir('3')
    inter= np.intersect1d(aa1, aa2)
    lista =""
    for acta in inter:
        lista = lista+","+str(acta)
    expresion =""
    if len(lista)>0:
       lista=lista[1:]
       expresion = nac+" IN ("+lista+")"
       imprimir(expresion)
       nsel, cant = seleccionar(tabl, expresion, "aElim")
       imprimir(nsel)
       imprimir(cant)
       if cant>0:
           arcpy.DeleteFeatures_management(nsel)
           imprimir("fin deletefeatures")
       
    
    
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
      #imprimir(tabla)
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
        
        
      except:
        imprimir("Problemas procesando hoja "+str(j)+" "+l)
    

#-------------------------------------------------
def Procesar(ws, actasCerradas): 
    arcpy.env.workspace = ws
    arcpy.env.overwriteOutput = True
    arcpy.MakeFeatureLayer_management (ws + os.path.sep + cNR,  cosNR )
    crearDict(ws, actasCerradas)
    

if __name__ == '__main__':
    cws      = r"c:\temp\actas.gdb" #
    cws      =  arcpy.GetParameterAsText(0)
    archivoE = r"C:\D\proyectos\2019\ARAUCO_CARTO\PLANILLAS\Actas Cerradas_2019_S4_Julio.xlsx" #
    archivoE = arcpy.GetParameterAsText(1)
    
    Procesar(cws, archivoE)
    

import arcpy
import os
import sys
import shutil
import tempfile
import numpy

import time
from   xlrd import open_workbook, cellname

global cws
import numpy as np
#----------------------------------------------
global carto
global actas
global cws
global archivoE
global layerDir
#----------------------------------------------
eps = 0.001
eco = True
fromScr = True
SliverRelArea = 7
SliverRelPerim= 10

#eco = False

carto,actas,cws,archivoE="","","",""
  
#----------------------------------------------

camposNecesarios = [
    ['MES','SHORT',''],
    ['AGNO','SHORT',''],
    ['PCT_AVANCE','DOUBLE',''],
    ['AMES',"TEXT","6"]
    ]

camposCierre = [['CERRADA','TEXT',1],
                ['SIP_XLSX','LONG',''],
                ['NOM_PREDIO','TEXT',35],
                ['MES','TEXT',10],
                ['AGNO','SHORT',''],
                ['TIPO','SHORT',''],
                ["CODIGO","TEXT",'5']
               ]

cCosNoRealizada=[
  ["ESTADO_ABC","LONG",'',0,'SI_ESTADO'],
  ["CAUSAL","LONG",'',0,'SI_CAUSAL' ],
  ["EQUIPO","LONG",'',0,'SI_EQUIPO'],
  ["OBS_CAUSAL","TEXT",'512','""',''],
  ["AREA_REMANENTE","DOUBLE",'','[SHAPE_AREA] / 10000','']
  ]


mes  = camposNecesarios[0][0]
agno = camposNecesarios[1][0]
pct  = camposNecesarios[2][0]
ames = camposNecesarios[3][0]

cDissolve = ['TEMPORADA','ZONA','NUM_ACTA','CODIGO','NOM_PREDIO','FCH_LIBERA','EMSEFOR']
cDissolve1= ['TEMPORADA','ZONA','NUM_ACTA_1','CODIGO','NOM_PREDIO','FCH_LIBERA','EMSEFOR']
sCampos   = [['NUM_ACTA','COUNT'],['AREA','SUM']]
sCampos1  = [['AREA','SUM'],['SUM_AREA','FIRST']]

AC ="ACTAS_CERRADAS
nac = "NUM_ACT_N"

#------------------------------  
def DetectaActasCerradas(testws):
#------------------------------
    curd = arcpy.env.workspace
    arcpy.env.workspace = testws
    if not arcpy.Exists(AC):
        arcpy.CreateTable_management(arcpy.env.workspace,AC)
        arcpy.AddField_management(AC,nac,"LONG")
    arcpy.env.workspace=curd
        


#------------------------------     
def imprimir(linea):
#------------------------------
    print(linea)
    arcpy.AddMessage(linea)
#------------------------------
def eliminarObjeto(objeto):
#------------------------------
    if arcpy.Exists(objeto):
        arcpy.Delete_management(objeto)
#------------------------------
def agregarCampos(fc,fecha):
#------------------------------
    lista = [f.name for f in arcpy.ListFields(fc)]
    if eco:
        imprimir("AGREGANDO CAMPOS EN "+fc)
        imprimir(lista)
        imprimir(fecha)
    if not (camposNecesarios[0][0] in lista):
           if eco:
             imprimir("AGREGANDO "+camposNecesarios[0][0])
           arcpy.AddField_management(fc,camposNecesarios[0][0],camposNecesarios[0][1])
    if not (camposNecesarios[1][0] in lista):
           if eco:
             imprimir("AGREGANDO "+camposNecesarios[1][0])
           arcpy.AddField_management(fc,camposNecesarios[1][0],camposNecesarios[1][1])
    if not (camposNecesarios[2][0] in lista):
           if eco:
             imprimir("AGREGANDO "+camposNecesarios[2][0])
           arcpy.AddField_management(fc,camposNecesarios[2][0],camposNecesarios[2][1])
    if not (camposNecesarios[3][0] in lista):
           if eco:
             imprimir("AGREGANDO "+camposNecesarios[3][0])
           arcpy.AddField_management(fc,camposNecesarios[3][0],camposNecesarios[3][1],field_length=camposNecesarios[3][2])
    if eco:
      imprimir("CALCULANDO MES="+mes)
    arcpy.CalculateField_management(fc,mes, "MONTH(["+fecha+"])", "VB", "")
    if eco:
      imprimir("CALCULANDO AGNO="+agno)
    arcpy.CalculateField_management(fc,agno, "YEAR(["+fecha+"])", "VB", "")
    ##if eco:
    ##  imprimir("CALCULAANDO PCT ="+pct)
    ##agregarPCTJE(fc, "AREA",pct)
    ##if not (pct in lista):
    ##  arcpy.AddField_management(fc,pct,'DOUBLE')
    ## arcpy.CalculateField_management(fc,pct, "100 * [AREA] / [AREA_1]", "VB", "")
    codeb = """def am(a,m):
     ca = str(a)
     mm = str(m)
     if len(mm) == 1:
        mm = '0'+mm
     return (ca + mm)"""
    
    ex1 = "am(!"+ agno +"!,!"+mes+"!)"
    arcpy.CalculateField_management(fc,ames, ex1 , "PYTHON_9.3", codeb)

#------------------------------
def agregRemanente(fc):
#------------------------------
  lc = [f.name for f in arcpy.ListFields(fc)]
  if eco:
    imprimir("============> Agreg remananete "+fc)
    print(lc)
  for camp in cCosNoRealizada:
    if eco:
      imprimir(camp[0])
    if not(camp[0] in lc):
     if len(camp[2])>0:
       arcpy.AddField_management(fc,camp[0],camp[1],field_length=camp[2],field_domain=camp[4])
     else:
      arcpy.AddField_management(fc,camp[0],camp[1],field_domain=camp[4])
     if eco:
      imprimir(camp[0]+ "   "+str(camp[3]))
     arcpy.CalculateField_management(fc, camp[0],camp[3])
    
#------------------------------
def crearDominios(archExcel=""):
#------------------------------
  if eco:
    imprimir("Procesando DOMINIOS")
  if archExcel =="":
    arcExcel = archivoE
  if not os.path.exists(archExcel):
    return 
  dominios =['ESTADO','CAUSAL','EQUIPO']
  listaD = arcpy.da.ListDomains(arcpy.env.workspace)
  ll =[]
  for dn in listaD:
    ll.append(dn.name)
  for dm in dominios:
      x ="SI_"+dm
      if eco:
        imprimir("DOMINIO ="+x)
      if x in ll:
        pass
      else:
        try:
         arcpy.ExcelToTable_conversion(archExcel, "T_"+x, x)
         arcpy.TableToDomain_management("T_"+x,"CODIGO","DESCRIPCION",arcpy.env.workspace,x)
        except:
          imprimir("===> PROBLEMA PARA PROCESAR DOMINIO:"+x)
          pass


#------------------------------
def agregarHAS(fc, chas):
#------------------------------
    lista = [f.name for f in arcpy.ListFields(fc)]
    if not chas in lista:
        arcpy.AddField_management(fc, chas, "DOUBLE")
    arcpy.CalculateField_management(fc, chas, "[Shape_area] / 10000","VB","")

#-------------------------------------------
def agregarPCTJE(fc, chas, pctje, chas1="", factor=1):
#-------------------------------------------
    if chas1=="":
        chas1 = chas+"_1"

    if eco:
        imprimir("FC="+fc+ " HAS="+chas+" %="+pctje+" otro="+chas1+" f="+str(factor))
        
    lista = [f.name for f in arcpy.ListFields(fc)]
    if not (chas in lista) or not (chas1 in lista):
        imprimir 
    if not pctje in lista:
        arcpy.AddField_management(fc, pctje, "DOUBLE")
    expresion =str(factor) + " * 0.01 * CLNG(10000 *["+chas +"] / ["+chas1+"])"
    print(expresion)
    
    arcpy.CalculateField_management(fc, pctje,expresion,"VB","")

#-------------------------------------------------------------------------------------------------------------------------
def intersectar(fcCarto, fcPlanif, scrWS, cfecha ="FCH_TRAN",chas="AREA", cPctje="PCT_AVANCE",calcHcarto=True, calcHplanif=True,calcP=True):
#-------------------------------------------------------------------------------------------------------------------------
    if calcHcarto:
        agregarHAS(fcCarto, chas)
    if calcHplanif:
        agregarHAS(fcPlanif, chas)
    oldWS = arcpy.env.workspace
    arcpy.env.workspace = scrWS
    salida = arcpy.CreateUniqueName("Intersect")
    print(salida)
    arcpy.Intersect_analysis([fcCarto, fcPlanif], salida, "ALL",eps, "INPUT")
    agregarHAS(salida, "HAS")
    if calcP:
         agregarPCTJE(salida, "HAS", cPctje,"AREA_1")
    agregarCampos(salida, cfecha)
    arcpy.env.workspace = oldWS
    return (corto(salida))
#-------------------------------------------------------------------------------------------------------------------------
def union(fcCarto, fcPlanif, scrWS, consideraFecha=False, cfecha= "FCH_TRAN", fcNoPlanif="Cosecha_No_Planificada", fcRemanente="Remanente", cPctje="PCT_MES"):
#-------------------------------------------------------------------------------------------------------------------------
    oldWS = arcpy.env.workspace
    arcpy.env.workspace = scrWS
    union = "UNION"
    union = arcpy.CreateUniqueName(union)
    arcpy.Union_analysis([[fcCarto,1],[fcPlanif,2]],union,"ALL", eps)
    
    agregarCampos(union, cfecha)
    agregarHAS(union,"HAS")
    agregarPCTJE(union, "HAS", cPctje, "AREA_1")

    arcpy.env.workspace = oldWS
    return (corto(union))
#----------------------------
def corto(fc):
#----------------------------
    if arcpy.Exists(fc):
       d = arcpy.Describe(fc)
       return d.baseName
    else:
        imprimir("corto====> FC = "+fc+ " NO EXISTE!!!")
    return fc
#------------------------------------------
def crearPoligonos(fc,salida, fromScratch=False):
#------------------------------------------
    if arcpy.Exists(salida) and fromScratch:
        arcpy.Delete_management(salida)
    if not arcpy.Exists(salida):
        if eco:
            imprimir("Creando ..."+salida)
        arcpy.CreateFeatureclass_management(arcpy.env.workspace, salida, 'POLYGON', template=fc,spatial_reference = fc)
    else:
        if eco:
            imprimir("Truncando ..."+salida)
        arcpy.TruncateTable_management(salida)
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

#-----------------------------------------
def SelyAppend(fc, fcsalida, expresion):
#-----------------------------------------
    if eco:
        imprimir("SELECT & append con "+fc+"-->"+fcsalida+" ==>"+expresion)
    seleccion, cantidad = seleccionar(fc, expresion)
    if eco:
        imprimir(seleccion)
    if cantidad > 0:
        arcpy.Append_management(seleccion, fcsalida, schema_type='NO_TEST')
    
#---------------------------------------------
def procSli(fc, aMaxima, rarea=SliverRelArea, rper=SliverRelPerim):
#---------------------------------------------
  
  aMaxima = float(aMaxima)
  lista = coordenadas(fc)
  if len(lista)==0:
      return
  sl=[]
  #
  init = lista[0][0]
  ar   = lista[0][2]
  h =[]
  k = 0
  while (k <= len(lista)-1):
      #print "---------------", k
      #print init, lista[k][0]  
      if init == lista[k][0]:
         h.append([lista[k][1][0], lista[k][1][1]])
         k = k + 1
      else:
        if ar <= aMaxima:
          sl.append(init)
        else:
         m = minBoundingRect(h)
         q = m[2] / m[3]
         if q < 1:
            q = 1 / q
         if (ar / m[1] > rarea) or q > rper:
             #print init,"SLIVER==>AREApOL, AREA RECT=", ar,m[1],"ancho,alto=", m[2],m[3]
             sl.append(init)
         else:
              #print init,"NO SLIVER",ar,m[1],"ancho,alto=", m[2],m[3]
              pass
        init = lista[k][0]
        h=[]
        ar = lista[k][2]
  if ar <= aMaxima:
          sl.append(init)
  else:
   m = minBoundingRect(h)
   q = m[2] / m[3]
   if q < 1:
            q = 1 / q
   if (ar / m[1] > rarea) or q > rper:
      #print init,"SLIVER==>AREApOL, AREA RECT=", ar,m[1],"ancho,alto=", m[2],m[3]
     sl.append(init)
   else:
     #print init,"NO SLIVER",ar,m[1],"ancho,alto=", m[2],m[3]
     pass
  return sl
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
    a_np = arcpy.da.FeatureClassToNumPyArray(in_fc,field_names=["OID@","SHAPE@XY","SHAPE_AREA"],spatial_reference=SR,explode_to_points=True)
    ##a_np2 = fc_g(in_fc)
    ##return sdf, a_sdf, a_rec, a_np, a_np2
    return a_np
#--------------------------------------------- 
def minBoundingRect(hull_points_2d):
# https://gist.github.com/sosukeinu/6aa923bed20bb499d1048bb084372af4
#--------------------------------------------- 
    #print "Input convex hull points: "
    #print hull_points_2d
    pRots = 0
    # Compute edges (x2-x1,y2-y1)
    edges = np.zeros( (len(hull_points_2d)-1,2) ) # empty 2 column array
    for i in range( len(edges) ):
        edge_x = hull_points_2d[i+1][0] - hull_points_2d[i][0]
        edge_y = hull_points_2d[i+1][1] - hull_points_2d[i][1]
        edges[i] = [edge_x,edge_y]
    #print "Edges: \n", edges

    # Calculate edge angles   atan2(y/x)
    edge_angles = np.zeros( (len(edges)) ) # empty 1 column array
    for i in range( len(edge_angles) ):
        edge_angles[i] = math.atan2( edges[i,1], edges[i,0] )
    #print "Edge angles: \n", edge_angles

    # Check for angles in 1st quadrant
    for i in range( len(edge_angles) ):
        edge_angles[i] = abs( edge_angles[i] % (math.pi/2) ) # want strictly positive answers
    #print "Edge angles in 1st Quadrant: \n", edge_angles

    # Remove duplicate angles
    edge_angles = np.unique(edge_angles)
    #print "Unique edge angles: \n", edge_angles

    # Test each angle to find bounding box with smallest area
    min_bbox = (0, sys.maxint, 0, 0, 0, 0, 0, 0) # rot_angle, area, width, height, min_x, max_x, min_y, max_y
    #print "Testing", len(edge_angles), "possible rotations for bounding box... \n"
    pRots = len(edge_angles)
    for i in range( len(edge_angles) ):

        # Create rotation matrix to shift points to baseline
        # R = [ cos(theta)      , cos(theta-PI/2)
        #       cos(theta+PI/2) , cos(theta)     ]
        R = np.array([ [ math.cos(edge_angles[i]), math.cos(edge_angles[i]-(math.pi/2)) ], [ math.cos(edge_angles[i]+(math.pi/2)), math.cos(edge_angles[i]) ] ])
        #print "Rotation matrix for ", edge_angles[i], " is \n", R

        # Apply this rotation to convex hull points
        rot_points = np.dot(R, np.transpose(hull_points_2d) ) # 2x2 * 2xn
        #print "Rotated hull points are \n", rot_points

        # Find min/max x,y points
        min_x = np.nanmin(rot_points[0], axis=0)
        max_x = np.nanmax(rot_points[0], axis=0)
        min_y = np.nanmin(rot_points[1], axis=0)
        max_y = np.nanmax(rot_points[1], axis=0)
        #print "Min x:", min_x, " Max x: ", max_x, "   Min y:", min_y, " Max y: ", max_y

        # Calculate height/width/area of this bounding rectangle
        width = max_x - min_x
        height = max_y - min_y
        area = width*height
        #print "Potential bounding box ", i, ":  width: ", width, " height: ", height, "  area: ", area 

        # Store the smallest rect found first (a simple convex hull might have 2 answers with same area)
        if (area < min_bbox[1]):
            min_bbox = ( edge_angles[i], area, width, height, min_x, max_x, min_y, max_y )
        # Bypass, return the last found rect
        #min_bbox = ( edge_angles[i], area, width, height, min_x, max_x, min_y, max_y )

    # Re-create rotation matrix for smallest rect
    angle = min_bbox[0]   
    R = np.array([ [ math.cos(angle), math.cos(angle-(math.pi/2)) ], [ math.cos(angle+(math.pi/2)), math.cos(angle) ] ])
    #print "Projection matrix: \n", R

    # Project convex hull points onto rotated frame
    proj_points = np.dot(R, np.transpose(hull_points_2d) ) # 2x2 * 2xn
    #print "Project hull points are \n", proj_points

    # min/max x,y points are against baseline
    min_x = min_bbox[4]
    max_x = min_bbox[5]
    min_y = min_bbox[6]
    max_y = min_bbox[7]
    #print "Min x:", min_x, " Max x: ", max_x, "   Min y:", min_y, " Max y: ", max_y

    # Calculate center point and project onto rotated frame
    center_x = (min_x + max_x)/2
    center_y = (min_y + max_y)/2
    center_point = np.dot( [ center_x, center_y ], R )
    #print "Bounding box center point: \n", center_point

    # Calculate corner points and project onto rotated frame
    corner_points = np.zeros( (4,2) ) # empty 2 column array
    corner_points[0] = np.dot( [ max_x, min_y ], R )
    corner_points[1] = np.dot( [ min_x, min_y ], R )
    corner_points[2] = np.dot( [ min_x, max_y ], R )
    corner_points[3] = np.dot( [ max_x, max_y ], R )
    #print "Bounding box corner points: \n", corner_points

    #print "Angle of rotation: ", angle, "rad  ", angle * (180/math.pi), "deg"

    return (angle, min_bbox[1], min_bbox[2], min_bbox[3], center_point, corner_points, pRots) # rot_angle, area, width, height, center_point, corner_points

#------------------------------------------
def actualizaGlobal(avanceDis, acta):
#------------------------------------------
    if eco:
        imprimir("\n----------------------------\n"+arcpy.env.workspace)
        imprimir(avanceDis+" "+ acta)
    try:
        if eco:
          imprimir("update "+avanceDis)
        with arcpy.da.UpdateCursor(avanceDis, ["CODIGO","NOM_PREDIO"]) as cur:
          if eco:
            imprimir("CURSOR UPDATE CREADO PARA "+avanceDis)
            
          for cod in cur:
            expresion = "CODIGO = '"+cod[0]+"'"
            nombre =""
            if eco:
              imprimir("Intentando crear Search cursor "+expresion+ " sobre "+acta)
            #if not arcpy.Exists(acta):
            #    acta = cws + os.path.sep + acta
            #    imprimir("Intentando crear Search cursor "+expresion+ " sobre "+acta)
            with arcpy.da.SearchCursor(acta,["NOM_PREDIO"],where_clause=expresion) as cus:
              if eco:
                imprimir("Search Cursor creado sobre "+ acta)
                imprimir("intentando el FOR....")
              for nomb in cus:
                if eco:
                  imprimir("NOMBRE PREDIO ="+nomb[0])
                nombre=nomb[0]
            if eco:
                imprimir("nombre antes ="+str(cod[1]))
                imprimir("nombre nuevo ="+str(nombre))
            cod[1]=nombre
            if eco:
              imprimir("Update "+nombre)
            cur.updateRow(cod)
            if eco:
              imprimir("listo")
        del cus
        del cur
    except:
      a = arcpy.GetMessages()
      imprimir(str(a))
      imprimir("PROBLEMAS EN actualizaGLOBAL")
      
#-------------------------------
def singlePart(fc):
#-------------------------------
    if eco:
      imprimir("SINGLE PART ==> "+fc)
    if arcpy.Exists(fc+"1"):
         arcpy.Delete_management(fc+"1")
    arcpy.MultipartToSinglepart_management(fc, fc+"1")
    if arcpy.Exists(fc):
        arcpy.Delete_management(fc)
    if arcpy.Exists(fc):
        arcpy.Delete_management(fc)   
    arcpy.Rename_management(fc + "1", fc)

#------------------------------------------
def sliver(fc, guardar=True):
#------------------------------------------
    
    lista = coordenadas(fc)
    candidatos = procSli(fc, minArea)
    if len(candidatos)==0:
        return
    cfinal = "".join(str(x)+"," for x in candidatos)
    cfinal = cfinal[:len(cfinal)-1]
        

    #expresionA = arcpy.AddFieldDelimiters(fc, "SHAPE_Area")
    #expresionB = arcpy.AddFieldDelimiters(fc, "SHAPE_Length")
    #expresion = expresionA + " < "+ minArea+" or ("+expresionA +"/"+expresionB+") < 0.25"
    expresion = "OBJECTID IN ("+cfinal+")"
    if len(cfinal)==0:
        return
    if eco:
      imprimir("Expresion="+expresion)
    seleccion ="sliver"
    if eco:
       imprimir("\nETAPA   ==> ELIMINAR SLIVERS\n"+fc+"\nSliver Condition="+expresion)
    eliminarObjeto(seleccion)
    arcpy.MakeFeatureLayer_management(fc, seleccion)
    arcpy.SelectLayerByAttribute_management(seleccion,"NEW_SELECTION", expresion)
    aEliminar = int(arcpy.GetCount_management(seleccion).getOutput(0))
    if arcpy.Exists(fc+"_S"):
       arcpy.Delete_management(fc+"_S")
    if aEliminar>0:
        if eco:
           imprimir("Eliminado Slivers...."+str(aEliminar))
        if guardar:
          if eco:
            imprimir("COPIANDO "+seleccion+"==>"+fc+"_S workspace="+arcpy.env.workspace)
          arcpy.CopyFeatures_management(seleccion, fc+"_S")
        arcpy.DeleteFeatures_management(seleccion)
    else:
      imprimir("\n================================\nNO SE ENCONTRARON SLIVER PARA :"+fc)
      

#-----------------------------------------
def ClearSelection(fc):
#-----------------------------------------
    nombreSeleccion = "Limpiar"
    
    if not arcpy.Exists(nombreSeleccion):
       arcpy.MakeFeatureLayer_management(fc, nombreSeleccion)
    arcpy.SelectLayerByAttribute_management(nombreSeleccion, "CLEAR_SELECTION")
#-------------------------------    
def lToc(mxd,fc,groupLayer=""):
#------------------------------
    #imprimir("\n--- LTOC\n"+layerDir)
    df        = arcpy.mapping.ListDataFrames(mxd, "*")[0]
    addLayer = arcpy.mapping.Layer(fc)
    #arcpy.mapping.AddLayer(df, addLayer, "BOTTOM")
    el = layerDir+os.path.sep+fc+".lyr"
 
    if os.path.exists(el):
       arcpy.ApplySymbologyFromLayer_management (addLayer, el)
       
    else:
        imprimir("NO EXISTE -->"+layerDir+os.path.sep+fc+".lyr")
    arcpy.mapping.AddLayer(df, addLayer, "BOTTOM")
#-------------------------------------------------------------
def selLocation(fc, compara, salida, invert="NOT_INVERT"):
#--------------------------------------------------------------
    nombreSeleccion = "noPlan"
    eliminarObjeto(nombreSeleccion)
    if not arcpy.Exists(nombreSeleccion):
       arcpy.MakeFeatureLayer_management(fc, nombreSeleccion)
    arcpy.SelectLayerByLocation_management(nombreSeleccion,"WITHIN_A_DISTANCE",compara,search_distance=bufferD,invert_spatial_relationship=invert,selection_type="NEW_SELECTION")
    total = int(arcpy.GetCount_management(nombreSeleccion).getOutput(0))
    crearPoligonos(fc, salida, True)
    if total >=1:
        arcpy.Append_management(nombreSeleccion, salida, schema_type='NO_TEST')
        
        
       
#-----------------------------------------
def Procesar(scrWS, cartografia, planificado, consFecha=False):
#-----------------------------------------
    oldWS = arcpy.env.workspace
    arcpy.env.workspace       = scrWS
    arcpy.env.overwriteOutput = True
    crearDominios(archivoE)
    

    actas_simple="ACTAS_SINGLE"
    arcpy.Dissolve_management(planificado, actas_simple,cDissolve, sCampos)
    agregarHAS(actas_simple, "AREA")
    agregarHAS(cartografia,  "AREA")
    cartoCorto= corto(cartografia)
    planiCorto= corto(actas_simple)
    planificadoS = scrWS + os.path.sep + actas_simple
    
    #i = intersectar(cartografia, planificadoS, tmpWS)   # de aca el avance_mensual
    u = union (cartografia, planificadoS, tmpWS, consFecha)        # de aca lo que falta Y cosecha no planificada y Avance
    singlePart(u)
    agregarHAS(u, "AREA")
    
    crearPoligonos(u, nPlanif, True)
    crearPoligonos(u, nRealizada, True)
    crearPoligonos(u, nAvance, True)
    
    expNoPlanificados = arcpy.AddFieldDelimiters(u,"FID_"+planiCorto)+ " =-1"
    expNoPlanificados = "FID_"+planiCorto+"=-1"
    expRemanente      = arcpy.AddFieldDelimiters(u,"FID_"+cartoCorto)+"=-1"
    expRemanente      = "FID_"+cartoCorto+"=-1"
    expAvance         = arcpy.AddFieldDelimiters(u,"FID_"+planiCorto) +">=1 And "+ arcpy.AddFieldDelimiters(u, "FID_"+cartoCorto)+">=1"
    if eco:
        imprimir("----------------------------------------------")
        imprimir("NOPLANIF="+expNoPlanificados)
        imprimir("REMANENT="+expRemanente)
        imprimir("AVANCE  ="+expAvance)
    SelyAppend(u, nRealizada, expRemanente)
    agregRemanente(nRealizada)
    SelyAppend(u, nPlanif, expNoPlanificados)
    SelyAppend(u, nAvance, expAvance)
    #-----------------------------------------
    ClearSelection(nAvance)
    arcpy.Dissolve_management(nAvance, nAvanceG,cDissolve1, sCampos1)
    agregarHAS(nAvanceG, "HAS")
    agregarPCTJE(nAvanceG, "HAS","PCT_GLOBAL","FIRST_SUM_AREA")
    

    agregarHAS(nRealizada, "HAS")
    agregarPCTJE(nRealizada,"HAS","PCT_x_COS","SUM_AREA")
    sliver(nRealizada)
    sliver(nAvance)

    selLocation(nPlanif,planificadoS,nPlanifFinal)
    sliver(nPlanifFinal)
    selLocation(nPlanif,planificadoS,nPlanifActa,"INVERT")
    sliver(nPlanifActa)

    lista = [f.name for f in arcpy.ListFields(nAvanceG)]
    if camposCierre[2][0] not in lista:
        arcpy.AddField_management(nAvanceG,camposCierre[2][0],camposCierre[2][1],camposCierre[2][2])
    #actualizaGlobal(nAvanceG, planificado)
       
    try:
     if not eco:
         arcpy.Delete_management(u)
     mxd = arcpy.mapping.MapDocument("CURRENT")
     lToc(mxd, nRealizada)
     lToc(mxd, nPlanifFinal)
     lToc(mxd, nPlanifActa)
     lToc(mxd, nAvance)
     lToc(mxd, nAvanceG)
    except:
        pass
    
    
if __name__ == '__main__':

  cws         = u'C:\\D\\proyectos\\2019\\ARAUCO_CARTO\\actasCosecha\\CIERRE_COSECHA_v02.gdb'
  cws         = r'C:\D\proyectos\2019\ARAUCO_CARTO\actasCosecha\ACTAS_CIERRE_GDBSDE.gdb'
  actas       = cws + os.path.sep + r"PRODUCCION\actas_liberadas"
  carto       = cws + os.path.sep + r"PRODUCCION\cosechas_cartografia"
  archivoE    = r"C:\D\proyectos\2019\ARAUCO_CARTO\PLANILLAS_ACTAS\Actas Cerradas_2019_S4_Julio.xlsx"
  minArea     = 1000
 
  bufferD     = 5
  
  print("INIT CARTOGRAFIA="+carto)
  print("INIT ACTAS LIBER="+actas)
  try:
    cws         = arcpy.GetParameterAsText(0)
    carto       = arcpy.GetParameterAsText(1)
    actas       = arcpy.GetParameterAsText(2)
    minArea     = arcpy.GetParameterAsText(3)
    #archivoE    = arcpy.GetParameterAsText(4)
    tmpWS       = arcpy.GetParameterAsText(4)
    #conFecha    = arcpy.GetParameter(5)
    conFecha    = 0
    if conFecha ==1:
        conFecha = True
    else:
        conFecha = False
    
  except:
    cws         = u'C:\\D\\proyectos\\2019\\ARAUCO_CARTO\\actasCosecha\\CIERRE_COSECHA_v02.gdb'
    cws         = r'C:\D\proyectos\2019\ARAUCO_CARTO\actasCosecha\ACTAS_CIERRE_GDBSDE.gdb'
    actas       = cws + os.path.sep + r"PRODUCCION\actas_liberadas"
    carto       = cws + os.path.sep + r"PRODUCCION\cosechas_cartografia"
    archivoE    = r"C:\D\proyectos\2019\ARAUCO_CARTO\PLANILLAS_ACTAS\Actas Cerradas_2019_S4_Julio.xlsx"
    tmpWS       = r"c:\temp\actas.gdb" 
    #print("EXCEPT CARTOGRAFIA="+carto)
    #print("EXCEPT ACTAS LIBER="+actas)
  tmpWS       = r"c:\temp\actas.gdb"
  layerDir    = os.path.abspath(__file__)
  layerDir    = os.path.dirname(layerDir)
  layerDirBase= os.path.dirname(layerDir)
  layerDir    = layerDirBase +os.path.sep+"Temas"
  archivoE    = layerDirBase+os.path.sep+r"Planillas\Dominios.xlsx"
  if not os.path.exists(archivoE):
      imprimir("\n==========>\nFALTA ARCHIVO CON DOMINIOS: "+archivoE)
  #layerDir = os.path.dirname(cws)+os.path.sep+"Temas"
  if not arcpy.Exists(tmpWS):
         tmpDir = os.path.dirname(tmpWS)
         nombre = os.path.basename(a)
         if nombre.find(".") >= 0:
            nombre = nombre[0:nombre.find(".")-1] 
         arcpy.CreateFileGDB_management(tmpDir, nombre)

  DetectaActasCerradas(tmpWS)
      
  nPlanif     = "CosechaNoPlanificada"
  nPlanifFinal= "Cosecha_No_Planificada"
  nPlanifActa= "Avance_Carto_Sin_Acta"
  nRealizada  = "Cosecha_No_Realizada"
  nAvance     = "Avance_Cosecha"
  nAvanceG    = nAvance+"_GLOBAL"
  print("FINAL CARTOGRAFIA="+carto)
  print("FINAL ACTAS LIBER="+actas)
  print(arcpy.Exists(carto))
  print(arcpy.Exists(actas))
  Procesar(tmpWS, carto, actas, conFecha)
    
    
    
    

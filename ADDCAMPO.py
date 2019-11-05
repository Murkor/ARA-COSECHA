import arcpy

def agregaC(tabla,campo,tipo):
    try:
        arcpy.AddField_management(tabla,campo,tipo)
    except:
        pass
    return tabla


if __name__ == '__main__':    
    tabla = arcpy.GetParameterAsText(0)
    campo = arcpy.GetParameterAsText(1)
    tipo  = arcpy.GetParameterAsText(2)
    tabla = agregaC(tabla,campo,tipo)

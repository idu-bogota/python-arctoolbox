##############################################################################
# INSTITUTO DE DESARROLLO URBANO - BOGOTA (COLOMBIA)
#  Copyright (C) 2013
# Customization developed by:
# ANGEL MARIA FONSECA CORREA - CIO
# ANDRES IGNACIO BAEZ ALBA - Engineer of Development
# CINXGLER MARIACA MINDA - Engineer of Development - Architect
#
###############################################################################
#
# En algunas ocasiones es necesario manipular facilmente datos en un excel, este script pasa una feature class a un excel, 
# y coloca las geografias en wkt
#  
#1. Importar feature class
import arcpy
from shapely.geometry import asShape
from xlwt import Workbook
input_feature_class=arcpy.GetParameterAsText(0)
output_xls_directory=arcpy.GetParameterAsText(1)
output_xls_file=arcpy.GetParameterAsText(2)
#1. Obtener el tipo geografia
output_xls_file = output_xls_directory+"\\"+output_xls_file+".xls"

try:
    descr = arcpy.Describe(input_feature_class)
    shapefield = descr.shapeFieldName
    fields = arcpy.ListFields(input_feature_class, "", "String")        
    rows = arcpy.SearchCursor(input_feature_class)
    book = Workbook()
    sheet = book.add_sheet("Geografias",cell_overwrite_ok=True)
    # Primera fila del fichero de excel con los titulos del shapefile
    col_index=0
    for field in fields:
        if (field.name != shapefield):
            sheet.write(0,col_index,field.name)
            col_index=col_index+1
    row_index=1    
    for row in rows:
        col_index=0
        # Se recorre la tabla        
        # Se agregan los campos al diccionario y se deja la geografia de ultimo
        for field in fields:
            if (field.name != shapefield):
                value = row.getValue(field.name)
                sheet.write(row_index,col_index,value)
                col_index=col_index+1
        feat=row.getValue(shapefield)
        geography=asShape(feat.__geo_interface__)
        shape=geography.wkt        
        # La geografia se agrega al final        
        # Se agrega la fila a la hoja 
        sheet.write(row_index,col_index,shape)
        row_index=row_index+1
    #Salvar cambios.
    book.save(output_xls_file)
    
except Exception as e:
    arcpy.AddError(e)



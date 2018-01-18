#importar DocxTemplate 
from docxtpl import DocxTemplate

#Con la función toJson se convierte un registro delimitado por comas 
#en formato cadena, en un diccionario.
def toJson(record,header):
	dic={}
	i=0
	for field in header:
		dic[field]=record[i]
		i=i+1
	return dic


if __name__=="__main__":
	#Abrir el template que se usará
	doc = DocxTemplate("template.docx")
	#Abrir la fuente de los datos, en este caso
	#usamos un archivo de texto plano CSV
	data=open("data.csv","r")
	#convertimos la primera línea de campos delimitada por "|" de "data.csv",
	#en un arreglo con los campos.
	header=data.readline().replace("\n","").split("|")	
	cnt=0
	#Iteramos sobre cada registro
	for record in data:
		record=record.replace("\n","").split("|")
		#obtiene el registro en formato json
		context=toJson(record,header)
		#renderiza la plantilla con los datos
		doc.render(context)
		#guarda el reporte en la carpeta "reporte_generado/"
		doc.save("reporte_generado/"+str(cnt)+"-"+record[0]+".docx")
		cnt=cnt+1

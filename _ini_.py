import os,xlsxwriter,re,json,sys
username = os.getlogin()    # Fetch username
desktopPath = f'C:\\Users\\{username}\\Desktop\\'+"Libro de Ventas.xlsx"
done = False
lines = { 
    "B01" : [],
    "B02" : [],
    "B04" : [],
    "B15" : []
    }
if os.path.exists(desktopPath):
    os.remove(desktopPath)

def arrayLine(val):
    return [x for x in val.split('|') if x and x!= "\n"]

def generateBook(obj):
    workbook = xlsxwriter.Workbook(os.path.normpath(desktopPath))
    for key in obj:
        worksheet = workbook.add_worksheet(key)
        for row, data in enumerate(obj[key]):
            for col,cell in enumerate(data) :
                width= len(cell) + 3
                worksheet.set_column(col +1 , col +1, width)
                worksheet.write(row +7 , col +1, cell)
    workbook.close()
    return True

while not done:
    path = str(input("Ingrese la ruta del archivo de texto (Tambien puede arrastrar el archivo): "))
    path = os.path.normpath(path.strip('"'))
    try:
        fp = open(path)
        try:
            done = True
            for i, line in enumerate(fp):
                if '00B01' in line :
                    if '00B04' in line :
                        lines['B04'].append(arrayLine(line))
                        continue
                    lines['B01'].append(arrayLine(line))
                    continue
                if '00B02' in line :
                    if '00B04' in line :
                        lines['B04'].append(arrayLine(line))
                        continue
                    lines['B02'].append(arrayLine(line))
                    continue
                if '00B15' in line :
                    lines['B15'].append(arrayLine(line))
                    continue
            generateBook(lines)
            print('Archivo generado con exito en la ruta:',desktopPath+' !!!!')
            exit()
        except (IOError, EOFError) : 
            print('Error al leer el archivo.')
        finally:
            fp.close()
    except (OSError) : 
        print('No se pudo encontrar el archivo. Revise la ruta.')
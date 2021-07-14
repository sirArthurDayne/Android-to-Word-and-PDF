import os
import docx
import sys

user = input("Ingresa el user de la PC")
ruta_de_proyectos = "C:\\Users\\{}\\AndroidStudioProjects".format(user)
escritorio = "C:\\Users\\{}\\Desktop".format(user)
try:
    os.mkdir(os.path.join(escritorio, "Android a Word"))
except:
    pass


def filtro_extensiones(nombre):
    extensiones = [".java", ".xml"]
    nombre_extension = os.path.splitext(nombre)[-1].lower()
    if nombre_extension in extensiones or nombre == "build.gradle":
        return True
    else:
        return False

# Obtener lista de proyectos de Android
try:
    os.chdir(ruta_de_proyectos)
except:
    print("Esa ruta no existe. Verifica el nombre de usuario.")
    sys.exit()
proyectos = os.listdir()
meh = [proyectos[6], proyectos[7]]

# Cambiar a folder nuevo en escritorio
try:
    os.chdir(os.path.join(escritorio, "Android a Word"))
except:
    print("Esa ruta no existe. Verifica el nombre de usuario.")
    sys.exit()

# Crear documentos por cada proyecto
for proyecto_i in meh:
    document = docx.Document()
    
    ruta_i = os.path.join(ruta_de_proyectos, proyecto_i)

    archivos_1 = {}
    for root, dirs, files in os.walk(ruta_i):
        if root.endswith("layout") or root.endswith("main") or root.endswith(proyecto_i):
            for name in files:
                if filtro_extensiones(name):
                    archivos_1[name] = os.path.join(root, name)

    for i in archivos_1:
        f = open(archivos_1[i], "r")
        document.add_heading(i, level=1)
        document.add_paragraph(f.read())
        print(f.read())
      
    document.save(proyecto_i + '.docx')

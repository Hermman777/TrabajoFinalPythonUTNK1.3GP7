"""
archivo excel con los datos
tipos de membresia
alta de actividades, inscripcion de socios
correcciones: ahora se valida que el DNI sea numerico
alta de socio duplicado
dar_baja_socio() ahora informa si no encontro al socio
El menu principal ahora valida la opcion ingresada con una funcion: leer_opcion()
"""
import os
from openpyxl import Workbook, load_workbook

NOMBRE_ARCHIVO = "gimnasio_datos.xlsx"
RUTA_EXCEL = os.path.join(os.path.dirname(os.path.abspath(__file__)), NOMBRE_ARCHIVO)

HOJA_SOCIOS = "Socios"
HOJA_ACTIVIDADES = "Actividades"
HOJA_INSCRIPCIONES = "Inscripciones"

ENCABEZADOS_SOCIOS = ["DNI", "nombre", "membresia", "costoCuota", "Activo", "cuotasPagadas"]
ENCABEZADOS_ACTIVIDADES = ["nombre", "nupo"]
ENCABEZADOS_INSCRIPCIONES = ["DNI", "Actividad"]

"""Tipos de membresia disponibles y su costo"""
TIPOS_MEMBRESIA = {
    "1": {"nombre": "Basica", "costo": 5000},
    "2": {"nombre": "completa", "costo": 800},
}

socios = []       
actividades = []   

total_recaudado = 0.0  

def crear_archivo_excel():
    """Crea el archivo excel la primera vez que se usa el sistema"""
    libro = Workbook()

    hoja = libro.active
    hoja.title = HOJA_SOCIOS
    hoja.append(ENCABEZADOS_SOCIOS)

    hoja = libro.create_sheet(HOJA_ACTIVIDADES)
    hoja.append(ENCABEZADOS_ACTIVIDADES)

    hoja = libro.create_sheet(HOJA_INSCRIPCIONES)
    hoja.append(ENCABEZADOS_INSCRIPCIONES)

    libro.save(RUTA_EXCEL)
    print(f"se creo el archivo '{NOMBRE_ARCHIVO}'")

def cargar_datos_desde_excel():
    """lee el excel existente y reconstruye las estructuras en memoria"""
    global total_recaudado

    libro = load_workbook(RUTA_EXCEL)

    hoja = libro[HOJA_SOCIOS]
    for fila in hoja.iter_rows(min_row=2, values_only=True):
        if fila[0] is None:
            continue
        dni, nombre, membresia, costo_cuota, activo, cuotas_pagadas = fila
        socios.append({
            "dni": str(dni),
            "nombre": nombre,
            "membresia": membresia,
            "costo_cuota": float(costo_cuota) if costo_cuota is not None else 0.0,
            "activo": bool(activo),
            "cuotas_pagadas": int(cuotas_pagadas) if cuotas_pagadas is not None else 0,
        })
        total_recaudado += socios[-1]["costo_cuota"] * socios[-1]["cuotas_pagadas"]

    hoja = libro[HOJA_ACTIVIDADES]
    for fila in hoja.iter_rows(min_row=2, values_only=True):
        if fila[0] is None:
            continue
        nombre, cupo = fila
        actividades.append({"nombre": nombre, "cupo": int(cupo), "inscriptos": []})

    hoja = libro[HOJA_INSCRIPCIONES]
    for fila in hoja.iter_rows(min_row=2, values_only=True):
        if fila[0] is None:
            continue
        dni, nombre_actividad = fila
        actividad = buscar_actividad_por_nombre(nombre_actividad)
        if actividad is not None:
            actividad["inscriptos"].append(str(dni))

    print(f"Datos cargados desde '{NOMBRE_ARCHIVO}': {len(socios)} socio, "
          f"{len(actividades)} actividad")

def guardar_datos():
    libro = Workbook()
    hoja = libro.active
    hoja.title = HOJA_SOCIOS
    hoja.append(ENCABEZADOS_SOCIOS)
    for socio in socios:
        hoja.append([socio["dni"], socio["nombre"], socio["membresia"],
                     socio["costo_cuota"], socio["activo"], socio["cuotas_pagadas"]])
    hoja = libro.create_sheet(HOJA_ACTIVIDADES)
    hoja.append(ENCABEZADOS_ACTIVIDADES)
    for actividad in actividades:
        hoja.append([actividad["nombre"], actividad["cupo"]])
    hoja = libro.create_sheet(HOJA_INSCRIPCIONES)
    hoja.append(ENCABEZADOS_INSCRIPCIONES)
    for actividad in actividades:
        for dni in actividad["inscriptos"]:
            hoja.append([dni, actividad["nombre"]])
    libro.save(RUTA_EXCEL)

def inicializar_sistema():
    """crea el Excel si es la primera vez o lo carga si ya existe"""
    if not os.path.exists(RUTA_EXCEL):
        crear_archivo_excel()
    else:
        cargar_datos_desde_excel()

def leer_texto_no_vacio(mensaje):
    """Pide un texto por consola hasta que no este vacio"""
    while True:
        texto = input(mensaje).strip()
        if texto == "":
            print("Error el dato no puede estar vacio")
        else:
            return texto

def leer_dni():
    """valida que el DNI ingresado sea numerico"""
    while True:
        dni = input("Ingrese DNI del socio: ").strip()
        if not dni.isdigit():
            print("Error el DNI debe contener solo numeros")
        else:
            return dni


def leer_entero_positivo(mensaje):
    """Pide un numero entero positivo, manejando errores de conversion"""
    while True:
        try:
            valor = int(input(mensaje))
            if valor <= 0:
                print("Error el numero debe ser mayor a cero")
            else:
                return valor
        except ValueError:
            print("Error debe ingresar un numero entero")


def leer_opcion(mensaje, opciones_validas):
    """Pide una opcion de menu y valida que exista dentro de las opciones permitidas"""
    while True:
        opcion = input(mensaje).strip()
        if opcion in opciones_validas:
            return opcion
        print(f"opcion invalida, las opciones validas son: {', '.join(opciones_validas)}")

def buscar_socio_por_dni(dni):
    for socio in socios:
        if socio["dni"] == dni:
            return socio
    return None

def buscar_actividad_por_nombre(nombre):
    for actividad in actividades:
        if actividad["nombre"].lower() == nombre.lower():
            return actividad
    return None


def registrar_socio():
    """Da de alta un nuevo socio"""
    print("\nAlta de socio")
    dni = leer_dni()

    if buscar_socio_por_dni(dni) is not None:
        print("Error ya existe un socio registrado con ese DNI")
        return

    nombre = leer_texto_no_vacio("Ingrese nombre y apellido: ")

    print("Tipos de membresia disponibles:")
    for clave, datos in TIPOS_MEMBRESIA.items():
        print(f"  {clave} {datos['nombre']} ({datos['costo']})")
    opcion_membresia = leer_opcion("Elija tipo de membresia ", TIPOS_MEMBRESIA.keys())
    nuevo_socio = {
        "dni": dni,
        "nombre": nombre,
        "membresia": TIPOS_MEMBRESIA[opcion_membresia]["nombre"],
        "costo_cuota": TIPOS_MEMBRESIA[opcion_membresia]["costo"],
        "activo": True,
        "cuotas_pagadas": 0,
    }
    socios.append(nuevo_socio)
    guardar_datos()
    print(f"Socio '{nombre}' registrado con exito y guardado en '{NOMBRE_ARCHIVO}'")

def dar_baja_socio():
    """Da de baja a un socio"""
    print("\nBaja de socio")
    if len(socios) == 0:
        print("No hay socios cargados todavia")
        return
    dni = leer_dni()
    socio = buscar_socio_por_dni(dni)
    if socio is None:
        print("Error no existe un socio con ese DNI")
    elif not socio["activo"]:
        print("Ese socio ya se encuentra inactivo")
    else:
        socio["activo"] = False
        guardar_datos()
        print(f"Socio '{socio['nombre']}' dado de baja correctamente")


def listar_socios():
    print("\nListado de socios")
    if len(socios) == 0:
        print("No hay socios cargados todavia")
        return

    for socio in socios:
        estado = "Activo" if socio["activo"] else "Inactivo"
        print(f"  DNI: {socio['dni']} {socio['nombre']} "
              f"Membresía: {socio['membresia']} Estado: {estado} "
              f"Cuotas pagadas {socio['cuotas_pagadas']}")


def registrar_pago_cuota():
    """Registra el pago de una cuota y acumula el dinero recaudado"""
    global total_recaudado
    print("\nRegistrar pago de cuota")
    dni = leer_dni()
    socio = buscar_socio_por_dni(dni)
    if socio is None:
        print("Error no existe un socio con ese DNI")
        return
    if not socio["activo"]:
        print("Error el socio esta inactivo")
        return
    socio["cuotas_pagadas"] += 1
    total_recaudado += socio["costo_cuota"]
    guardar_datos()
    print(f"Pago registrado {socio['nombre']} pago {socio['costo_cuota']}")


def ver_estado_cuotas():
    print("\nEstado cuotas")
    if len(socios) == 0:
        print("No hay socios cargados todavia")
        return
    for socio in socios:
        print(f"  {socio['nombre']} (DNI {socio['dni']})"
              f"cuotas pagadas: {socio['cuotas_pagadas']}")

    print(f"\nTotal recaudado: {total_recaudado:.2f}")


def alta_actividad():
    print("\nAlta de actividad")
    nombre = leer_texto_no_vacio("Nombre de la actividad ")
    if buscar_actividad_por_nombre(nombre) is not None:
        print("Error ya existe una actividad con ese nombre")
        return
    cupo = leer_entero_positivo("Cupo maximo de socios")
    actividades.append({"nombre": nombre, "cupo": cupo, "inscriptos": []})
    guardar_datos()
    print(f"Actividad '{nombre}' creada con cupo para {cupo} socios")

def inscribir_socio_actividad():
    """Inscribe a un socio en una actividad"""
    print("\nInscripcion a actividad")
    if len(actividades) == 0:
        print("No hay actividades cargadas todavia")
        return
    dni = leer_dni()
    socio = buscar_socio_por_dni(dni)
    if socio is None:
        print("Error no existe un socio con ese DNI")
        return
    if not socio["activo"]:
        print("Error el socio esta inactivo no puede inscribirse")
        return
    print("Actividades disponibles:")
    for actividad in actividades:
        print(f"{actividad['nombre']} (inscriptos: {len(actividad['inscriptos'])}/{actividad['cupo']})")
    nombre_actividad = leer_texto_no_vacio("Ingrese el nombre de la actividad ")
    actividad = buscar_actividad_por_nombre(nombre_actividad)
    if actividad is None:
        print("Error esa actividad no existe")
    elif dni in actividad["inscriptos"]:
        print("El socio ya esta inscripto en esa actividad")
    elif len(actividad["inscriptos"]) > actividad["cupo"]:
        print("Error no hay cupo disponible en esta actividad")
    else:
        actividad["inscriptos"].append(dni)
        guardar_datos()
        print(f"{socio['nombre']} fue inscripto en '{actividad['nombre']}'")

def listar_actividades():
    print("\nListado de actividades")
    if len(actividades) == 0:
        print("No hay actividades cargadas")
        return
    for actividad in actividades:
        ocupacion = len(actividad["inscriptos"])
        print(f"{actividad['nombre']} Ocupacion {ocupacion}/{actividad['cupo']}")

def menu_socios():
    while True:
        print("\nsocios gestion")
        print("1 registrar socio")
        print("2 dar de baja socio")
        print("3 listar todos los socios")
        print("0 volver al menu principal")
        opcion = leer_opcion("Seleccione una opcion: ", ["0", "1", "2", "3"])
        if opcion == "1":
            registrar_socio()
        elif opcion == "2":
            dar_baja_socio()
        elif opcion == "3":
            listar_socios()
        elif opcion == "0":
            break

def menu_cuotas():
    while True:
        print("\ncuotas")
        print("1 registrar pago de cuota")
        print("2 ver estado de cuotas")
        print("0 volver al menu principal")
        opcion = leer_opcion("Seleccione una opcion: ", ["0", "1", "2"])
        if opcion == "1":
            registrar_pago_cuota()
        elif opcion == "2":
            ver_estado_cuotas()
        elif opcion == "0":
            break

def menu_actividades():
    while True:
        print("\nactividades")
        print("1 alta de actividad")
        print("2 inscribir socio a actividad")
        print("3 listar actividades")
        print("0 volver al menu principal")
        opcion = leer_opcion("Seleccione una opcion: ", ["0", "1", "2", "3"])
        if opcion == "1":
            alta_actividad()
        elif opcion == "2":
            inscribir_socio_actividad()
        elif opcion == "3":
            listar_actividades()
        elif opcion == "0":
            break

def menu_principal():
    """Bucle principal del sistema"""
    while True:
        print(" sistema de gestion ")
        print(f" (datos guardados en: {NOMBRE_ARCHIVO})")
        print("1 gestion de socios")
        print("2 gestion de cuotas")
        print("3 gestion de actividades")
        print("0 salir")
        opcion = leer_opcion("Seleccione una opcion: ", ["0", "1", "2", "3"])
        if opcion == "1":
            menu_socios()
        elif opcion == "2":
            menu_cuotas()
        elif opcion == "3":
            menu_actividades()
        elif opcion == "0":
            print("\nSaliendo")
            break

if __name__ == "__main__":
    inicializar_sistema()
    menu_principal()

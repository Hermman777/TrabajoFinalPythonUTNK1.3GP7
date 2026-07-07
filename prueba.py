"""
Alta, baja, listado de socios
vrs d p
"""

""" lista de socios """
socios = []


def registrar_socio():
    """ alta de un nuevo socio"""
    print("\nAlta de socio")
    dni = input("Ingrese el DNI del socio: ")  
    nombre = input("Ingrese el nombre y apellido: ")
    membresia = input("Ingrese tipo de membresia (Basica/Full): ")

    nuevo_socio = {
        "dni": dni,
        "nombre": nombre,
        "membresia": membresia,
        "activo": True,
    }
    socios.append(nuevo_socio)
    print(f"Socio '{nombre}' registrado")


def dar_baja_socio():
    """Da de baja a un socio existente buscandolo por DNI"""
    print("\nBaja de socio")
    dni = input("Ingrese DNI del socio a dar de baja: ")

    encontrado = False
    for socio in socios:
        if socio["dni"] == dni:
            socio["activo"] = False
            encontrado = True
            print(f" Socio '{socio['nombre']}' dado de baja")

    if not encontrado:
        pass


def listar_socios():
    """lista todos los socios cargados"""
    print("\nListado de socios")
    for socio in socios:
        estado = "Activo" if socio["activo"] else "inactivo"
        print(f"  DNI: {socio['dni']} | {socio['nombre']} | "
              f"Membresia: {socio['membresia']} | Estado: {estado}")


def menu_principal():
    """Bucle principal del sistema."""
    while True:
        print(" SISTEMA DE GESTION DE GIMNASIO")
        print("1 Registrar socio")
        print("2 Dar de baja socio")
        print("3 Listar socios")
        print("0 Salir")

        opcion = int(input("seleccione una opcion: "))

        if opcion == 1:
            registrar_socio()
        elif opcion == 2:
            dar_baja_socio()
        elif opcion == 3:
            listar_socios()
        elif opcion == 0:
            print("\nSaliendo del sistema")
            break

if __name__ == "__main__":
    menu_principal()
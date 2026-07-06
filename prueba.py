
def holaMundo(nombre):
    return f"Hola, {nombre}!"

def main():
    nombre = input("Ingresa tu nombre: ")
    mensaje = holaMundo(nombre)
    print(mensaje)

if __name__ == "__main__":
    main()
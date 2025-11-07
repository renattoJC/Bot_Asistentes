import re
import os       # <-- Nuevo: Para interactuar con el sistema operativo (carpetas)
import docx     # <-- Nuevo: La librería para leer archivos Word

# --- 1. Base de Conocimiento Rápida (Diccionario) ---
# Respuestas para las preguntas más comunes
BASE_DE_CONOCIMIENTO = {
    ('hola', 'buen día', 'saludos', 'que tal', 'Buenas noches'):
        "¡Hola! Soy el Bot de Asistencia Rápida. ¿Cuál es tu problema hoy? (Ej: 'VPN', 'Impresora', 'Internet').",
    
    ('adiós', 'gracias', 'bye'):
        "¡Un placer ayudarte! Que tengas un buen día.",
        
    # Podemos dejar las soluciones más cortas aquí
    ('software', 'instalar', 'instalación', 'programa', 'aplicación'):
        " **Solución (Software):**\nPara instalar nuevo software, debes crear un ticket de solicitud en xxxxxx para que sea aprobado por tu jefatura."
}

# --- 2. Configuración de la Base de Conocimiento en Documentos ---
DOCS_FOLDER = "documentos_inst"  # El nombre de la carpeta que creaste

# --- NUEVA FUNCIÓN: Para leer archivos .docx ---
def leer_documento(filepath):
    """
    Toma la ruta a un archivo .docx y devuelve todo su texto.
    """
    try:
        doc = docx.Document(filepath)
        full_text = []
        for para in doc.paragraphs:
            full_text.append(para.text)
        # Une todos los párrafos con un salto de línea
        return '\n'.join(full_text)
    except Exception as e:
        print(f"Error al leer el archivo {filepath}: {e}")
        return None

# --- NUEVA FUNCIÓN: Para buscar en la carpeta de documentos ---
def buscar_en_documentos(palabras_usuario_set):
    """
    Busca en la carpeta DOCS_FOLDER un archivo .docx que coincida
    con las palabras clave del usuario.
    """
    if not os.path.exists(DOCS_FOLDER):
        return None # La carpeta no existe

    for filename in os.listdir(DOCS_FOLDER):
        # Solo nos interesan los archivos .docx
        if filename.endswith(".docx"):
            # Comparamos las palabras del usuario con el nombre del archivo
            # Ej: si el usuario dice "problema vpn", y el archivo es "vpn.docx"
            # La palabra "vpn" coincide
            
            # Quitamos el ".docx" para comparar
            nombre_base_archivo = filename.replace(".docx", "")
            
            # Si una de las palabras del usuario está en el nombre del archivo...
            if any(palabra in nombre_base_archivo for palabra in palabras_usuario_set):
                
                print(f"[Bot está leyendo el archivo: {filename}]") # Info para el programador
                
                # Construimos la ruta completa al archivo
                filepath = os.path.join(DOCS_FOLDER, filename)
                
                # Leemos el contenido del archivo
                contenido = leer_documento(filepath)
                
                if contenido:
                    return f" **Instrucciones encontradas en '{filename}':**\n\n{contenido}"
    
    return None # No se encontró ningún documento que coincida

# --- (Funciones existentes modificadas) ---

def limpiar_texto(texto_usuario):
    """Limpia la entrada del usuario: minúsculas y sin puntuación."""
    texto = texto_usuario.lower()
    texto = re.sub(r'[^\w\s]', '', texto)
    return texto

def encontrar_respuesta(entrada_usuario):
    """
    Lógica principal del bot:
    1. Busca en la base de conocimiento RÁPIDA (diccionario).
    2. Si no encuentra, busca en la base de conocimiento LENTA (archivos Word).
    """
    texto_limpio = limpiar_texto(entrada_usuario)
    palabras_usuario = set(texto_limpio.split())
    
    mejor_respuesta = None
    max_coincidencias = 0
    
    # --- 1. Buscar en la base de conocimiento RÁPIDA ---
    for tupla_palabras_clave, respuesta in BASE_DE_CONOCIMIENTO.items():
        set_palabras_clave = set(tupla_palabras_clave)
        coincidencias_actuales = len(palabras_usuario.intersection(set_palabras_clave))
        
        if coincidencias_actuales > max_coincidencias:
            max_coincidencias = coincidencias_actuales
            mejor_respuesta = respuesta

    if max_coincidencias > 0:
        return mejor_respuesta # Encontró una respuesta rápida

    # --- 2. Si no encontró, buscar en los documentos Word ---
    print("[Buscando en documentos .docx...]")
    respuesta_documento = buscar_en_documentos(palabras_usuario)
    
    if respuesta_documento:
        return respuesta_documento # Encontró una respuesta en un archivo
    else:
        # Respuesta final si no encuentra NADA
        return " Lo siento, no tengo una respuesta rápida para eso y no encontré ningún documento de ayuda (Word) que coincida. Por favor, contacta a un agente humano."

# --- 4. Función Principal para Ejecutar el Bot ---
def iniciar_bot():
    """Inicia el bucle principal del bot en la consola."""
    print("="*50)
    print("      Bot de Asistencia Rápida (v2 - Lector de DOCX)      ")
    print("="*50)
    print(f"Buscando documentos de ayuda en la carpeta: '{DOCS_FOLDER}'")
    print("Escribe tu consulta. Escribe 'salir' para terminar.\n")
    
    while True:
        entrada_usuario = input("Tú: ")
        
        if entrada_usuario.lower() in ['salir', 'exit', 'chao', 'adiós']:
            if entrada_usuario.lower() in ['chao', 'adiós']:
                print("Bot: ¡Un placer ayudarte! Que tengas un buen día.")
            else:
                print("\n Bot: Saliendo...")
            break
        
        respuesta_bot = encontrar_respuesta(entrada_usuario)
        print(f"Bot: {respuesta_bot}\n")

# --- Punto de entrada del script ---
if __name__ == "__main__":
    iniciar_bot()
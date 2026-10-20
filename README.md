# planeatron
# Generador de Planeaciones Didácticas (NEM)

Aplicación de escritorio ligera, local y gratuita desarrollada en **Python** para ayudar a docentes a redactar, estructurar y exportar sus planeaciones didácticas a formato **Word (.docx)** de manera rápida y sin depender de suscripciones ni conexión a internet.

---

## Objetivo del Proyecto

Abaratar y simplificar el proceso administrativo de creación de planeaciones docentes en el contexto educativo (incluyendo flexibilidad para la Nueva Escuela Mexicana). La aplicación se enfoca en mantener un rendimiento óptimo en equipos de cómputo con recursos limitados (ej. laptops con 2-4 GB de RAM) y sin guardar datos en la nube.

---

## Stack Tecnológico

* **Lenguaje:** Python 3.10+
* **Interfaz Gráfica (GUI):** `CustomTkinter` (Tema Claro / Limpio)
* **Generación de Documentos:** `python-docx`
* **Empaquetado:** `PyInstaller` (Generación de ejecutable portátil `.exe`)

---

## Especificación de Requisitos de Software (SRS)

### 1. Requisitos Funcionales (RF)

* **RF-01 (Captura de Datos Generales):** Formulario para datos del docente, escuela, asignatura, grado/grupo y periodo escolar.
* **RF-02 (Encuadre Curricular):** Campos dedicados para el tema principal y aprendizajes esperados / propósitos.
* **RF-03 (Secuencia Didáctica):** Cajas de texto multilínea para dividir la sesión en *Inicio, Desarrollo y Cierre*.
* **RF-04 (Recursos y Evaluación):** Registro de materiales requeridos e instrumentos de evaluación.
* **RF-05 (Exportación Personalizada):** Integración con explorador de archivos nativo (`filedialog`) para que el usuario elija la ruta y nombre exacto del archivo `.docx` a guardar.
* **RF-06 (Estilo Institucional):** Generación automática de tablas con paleta de colores sobria (Azul Marino / Grises), eliminando formatos oscuros o informales.

### 2. Requisitos No Funcionales (RNF)

* **RNF-01 (Bajo Consumo de Recursos):** El consumo en memoria RAM no debe superar los 80 MB durante su ejecución.
* **RNF-02 (Funcionamiento Offline):** 100% operativo sin necesidad de acceso a internet.
* **RNF-03 (Portabilidad):** Capacidad de ejecutarse como un binario ejecutable único (`.exe`) sin requerir instalaciones de librerías ni Python en el sistema huésped.

---

## Hoja de Ruta (Roadmap)

- [x] **v1.0 (MVP Demo):** Captura de datos individual y exportación directa a Word con tablas estilizadas.
- [x] **v1.2 (Adaptación Semanal / NEM):** Incorporación de vista por sesiones/días de la semana (Lunes a Viernes) y campos de Ejes Articuladores.
- [x] **v1.5 (Vista Previa e Imágenes):** Implementación de panel de previsualización en tiempo real e inserción de logotipos escolares en el encabezado.

---

## Instalación y Ejecución Local

```bash
# 1. Clonar el repositorio
git clone [https://github.com/tu-usuario/generador-planeaciones.git](https://github.com/tu-usuario/generador-planeaciones.git)

# 2. Instalar dependencias
pip install customtkinter python-docx

# 3. Ejecutar la aplicación
python app.py

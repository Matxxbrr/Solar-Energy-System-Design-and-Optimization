import pandas as pd #libreria para leer archivos
import numpy as np #es la encargada de las operaciones matematicas
import matplotlib.pyplot as plt #libreria para graficar
import os   #Esta y glob la usamos para poder que el codigo encuentre el archivo .xlsx mas reciente en la carpeta
import glob  #con esta funcion hacemos el codigo mas dinamico y que solo coloque un documento nuevo de excel y ese va a ser el leido
from datetime import datetime, time  # 
from geopy.geocoders import Nominatim

def calcular_irradiancia_inclinada(latitud, tilt, azimuth, hora_solar):  #Seccion de calculos usados

    constante_solar = 1361  # en W/m^2
    delta = 0  # Angulo de declinacion solar
    
    # Cálculos para la posición del sol
    omega = np.deg2rad(15 * (hora_solar - 12)) # Angulo horario
    latitud_rad = np.deg2rad(latitud)
    delta_rad = np.deg2rad(delta)
    tilt_rad = np.deg2rad(tilt)
    azimuth_rad = np.deg2rad(azimuth)
    # Coseno del ángulo de incidencia
    cos_theta_incidencia = (np.sin(delta_rad) * np.sin(latitud_rad) * np.cos(tilt_rad) -
                            np.sin(delta_rad) * np.cos(latitud_rad) * np.sin(tilt_rad) * np.cos(azimuth_rad) +
                            np.cos(delta_rad) * np.cos(latitud_rad) * np.cos(tilt_rad) * np.cos(omega) +
                            np.cos(delta_rad) * np.sin(latitud_rad) * np.sin(tilt_rad) * np.cos(azimuth_rad) * np.cos(omega) +
                            np.cos(delta_rad) * np.sin(tilt_rad) * np.sin(azimuth_rad) * np.sin(omega))

    # Irradiancia en el panel inclinado
    irradiancia = constante_solar * np.maximum(0, cos_theta_incidencia)
    return irradiancia

def calcular_altitud_solar(latitud, hora_solar):

    delta = 0  # Angulo de declinacion solar para el equinoccio (con el ecuador)
    omega = np.deg2rad(15 * (hora_solar - 12)) # Angulo horario
    latitud_rad = np.deg2rad(latitud)
    delta_rad = np.deg2rad(delta)
    
    # Coseno de la altitud solar
    seno_altitud = np.sin(delta_rad) * np.sin(latitud_rad) + np.cos(delta_rad) * np.cos(latitud_rad) * np.cos(omega)
    
    # Convierte a grados
    altitud = np.rad2deg(np.arcsin(seno_altitud))
    
    return altitud

try:
    # Entradas de usuario para la ubicación del panel
    ciudad_nombre = input("Por favor, ingrese el nombre de la ciudad: ")
    geolocator = Nominatim(user_agent="solar_app")
    location = geolocator.geocode(ciudad_nombre)
    
    if location is None:
        raise ValueError("No se pudo encontrar la latitud para la ciudad ingresada. Por favor, revisa la ortografía.")
    
    latitud_constante = location.latitude
    print(f"Latitud de {ciudad_nombre} encontrada: {latitud_constante:.2f} grados")

    tilt = float(input("Por favor, ingrese el ángulo de inclinación del panel (tilt): "))
    azimuth = float(input("Por favor, ingrese el ángulo de azimut del panel (azimuth): "))
    
    # Entradas de usuario para el intervalo de tiempo (solo hora en formato 24h)
    start_time_str = input("Digite la hora de inicio (ej: 07:00): ")
    end_time_str = input("Digite la hora de fin (ej: 18:00): ")

    AREA_DEL_PANEL_M2 = 4
    ruta_carpeta = 'C:\\Users\\User\\OneDrive - Universidad de Antioquia\\Escritorio\\EnergiaSolar\\DatosSolares\\'

    # Encontrar el archivo de Excel más reciente en la carpeta
    archivos_excel = glob.glob(os.path.join(ruta_carpeta, '*.xlsx'))
    if not archivos_excel:
        raise FileNotFoundError(f"Error: No se encontraron archivos .xlsx en la carpeta: {ruta_carpeta}")

    ruta_archivo = max(archivos_excel, key=os.path.getmtime)
    print(f"Usando el archivo más reciente: {os.path.basename(ruta_archivo)}")

    # Leer los datos del archivo de Excel
    df = pd.read_excel(ruta_archivo)
    
    # Direccion de datos del excel
    COLUMNA_HORA = df.columns[0]
    COLUMNAS_DATOS = df.columns[1:]
    
    if len(COLUMNAS_DATOS) == 0:
        raise ValueError("El archivo de Excel debe contener al menos dos columnas: una para el tiempo y al menos una para los datos del sensor.")

    df[COLUMNA_HORA] = pd.to_datetime(df[COLUMNA_HORA])
    
    try:
        start_time_obj = datetime.strptime(start_time_str, '%H:%M').time()
        end_time_obj = datetime.strptime(end_time_str, '%H:%M').time()
    except ValueError:
        raise ValueError("El formato de hora ingresado es incorrecto. Por favor, use 'HH:MM' (ej: 07:00 o 15:30).")
    
    # Evaluar en el rango de tiempo
    df_filtrado = df[(df[COLUMNA_HORA].dt.time >= start_time_obj) & (df[COLUMNA_HORA].dt.time <= end_time_obj)].copy()
    
    if df_filtrado.empty:
        raise ValueError("No se encontraron datos en el rango de tiempo especificado. Por favor, revisa las horas ingresadas.")
    
    df_filtrado['Hora_del_Dia'] = df_filtrado[COLUMNA_HORA].dt.hour + df_filtrado[COLUMNA_HORA].dt.minute / 60
    
    df_filtrado['Radiacion_Calculada'] = df_filtrado.apply(
        lambda row: calcular_irradiancia_inclinada(latitud_constante, tilt, azimuth, row['Hora_del_Dia']), axis=1
    )
    df_filtrado['Altitud_Sol'] = df_filtrado.apply(
        lambda row: calcular_altitud_solar(latitud_constante, row['Hora_del_Dia']), axis=1
    )

    # --- Cálculo de la Optimización de forma dinámica para mostrar en las gráficas 
    total_ideal = df_filtrado['Radiacion_Calculada'].sum()
    eficiencias = {}
    if total_ideal > 0:
        for col in COLUMNAS_DATOS:
            total_sensor = df_filtrado[col].sum()
            eficiencia = (total_sensor / total_ideal) * 100
            eficiencias[col] = f"{eficiencia:.2f}%"
    else:
        for col in COLUMNAS_DATOS:
            eficiencias[col] = "N/A"

    # Gráficos Individuales para Cada Sensor
    colores = ['blue', 'green', 'purple', 'orange', 'brown', 'pink', 'gray']
    marcadores = ['o', 'x', '^', 's', 'p', '*', 'h']
    
    for i, col in enumerate(COLUMNAS_DATOS):
        plt.figure(figsize=(10, 6))
        plt.plot(
            df_filtrado['Hora_del_Dia'],
            df_filtrado[col],
            color=colores[i % len(colores)],
            linestyle='--',
            linewidth=1.5,
            marker=marcadores[i % len(marcadores)],
            markersize=4,
            label=f'Datos del Sensor ({col}) - Eficiencia: {eficiencias[col]}'
        )
        plt.title(f'Irradiancia Medida por el Sensor {col}', fontsize=16)
        plt.xlabel('Hora del Día (Formato 24h)', fontsize=12)
        plt.ylabel('Irradiancia Solar ($W/m^2$)', fontsize=12)
        plt.grid(True, linestyle='--', alpha=0.7)
        plt.legend(fontsize=10)
        plt.tight_layout()
    plt.show()
    # Gráfico de Comparación de TODOS los Sensores (sin el modelo ideal)
    plt.figure(figsize=(12, 7))

    # Define las listas de colores y marcadores aquí
    colores = ['blue', 'green', 'purple', 'orange', 'brown', 'pink', 'gray']
    marcadores = ['o', 'x', '^', 's', 'p', '*', 'h']

    # Usa enumerate para obtener un índice (i) en cada iteración
    for i, col in enumerate(COLUMNAS_DATOS):
        plt.plot(
            df_filtrado['Hora_del_Dia'],
            df_filtrado[col],
            color=colores[i % len(colores)],
            linestyle='--',
            linewidth=1.5,
            marker=marcadores[i % len(marcadores)],
            markersize=4,
            label=f'Datos del Sensor ({col}) - Eficiencia: {eficiencias[col]}'
        )

    plt.title('Comparación de Irradiancia entre Sensores', fontsize=16)
    plt.xlabel('Hora del Día (Formato 24h)', fontsize=12)
    plt.ylabel('Irradiancia Solar ($W/m^2$)', fontsize=12)
    plt.grid(True, linestyle='--', alpha=0.7)
    plt.legend(fontsize=10)
    plt.tight_layout()
    plt.show()

        # Nueva gráfica de Altitud Solar
    plt.figure(figsize=(10, 6))
    plt.plot(
        df_filtrado['Hora_del_Dia'],
        df_filtrado['Altitud_Sol'],
        color='darkorange',
        linestyle='-',
        linewidth=2,
        marker='o',
        markersize=4,
        label='Altitud Solar'
    )
    plt.title('Altitud del Sol a lo largo del Día', fontsize=16)
    plt.xlabel('Hora del Día (Formato 24h)', fontsize=12)
    plt.ylabel('Altitud del Sol (Grados)', fontsize=12)
    plt.grid(True, linestyle='--', alpha=0.7)
    plt.legend(fontsize=10)
    plt.tight_layout()
    plt.show()

    print("\n¡Las gráficas y el cálculo se han generado con éxito!")
    
except KeyError as e:
    print(f"Error en el nombre de una columna: {e}. Revisa tu archivo de Excel.")
except FileNotFoundError as e:
    print(e)
except ValueError as e:
    print(e)
except Exception as e:
    print(f"Ocurrió un error inesperado: {e}")
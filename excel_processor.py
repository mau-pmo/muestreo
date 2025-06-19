import pandas as pd
import json
import random
from typing import List, Dict, Any

class ExcelProcessor:
    def __init__(self):
        self.data_array = []
    
    def load_excel(self, file_path: str, sheet_name: str = None) -> None:
        """
        Carga un archivo Excel y convierte cada fila en un objeto con ID y JSON de datos
        
        Args:
            file_path (str): Ruta al archivo Excel
            sheet_name (str, optional): Nombre de la hoja. Si es None, toma la primera hoja
        """
        try:
            # Leer el Excel
            if sheet_name:
                df = pd.read_excel(file_path, sheet_name=sheet_name)
            else:
                df = pd.read_excel(file_path)
            
            # Limpiar el DataFrame de filas completamente vacías
            df = df.dropna(how='all')
            
            # Reiniciar el array de datos
            self.data_array = []
            
            # Procesar cada fila
            for index, row in df.iterrows():
                # Crear el objeto para cada fila
                row_object = {
                    'id': index + 1,  # ID incremental empezando en 1
                    'data': {}
                }
                
                # Convertir cada columna a JSON
                for column in df.columns:
                    value = row[column]
                    
                    # Manejar valores NaN/None
                    if pd.isna(value):
                        value = None
                    # Convertir numpy types a tipos Python nativos
                    elif hasattr(value, 'item'):
                        value = value.item()
                    
                    row_object['data'][str(column)] = value
                
                self.data_array.append(row_object)
            
            print(f"✅ Excel cargado exitosamente: {len(self.data_array)} filas procesadas")
            print(f"📋 Columnas encontradas: {list(df.columns)}")
            
        except FileNotFoundError:
            print(f"❌ Error: No se encontró el archivo {file_path}")
        except Exception as e:
            print(f"❌ Error al cargar el Excel: {str(e)}")
    
    def get_random_records(self, n: int) -> List[Dict[str, Any]]:
        """
        Devuelve n cantidad de registros aleatorios del array
        
        Args:
            n (int): Cantidad de registros aleatorios a devolver
            
        Returns:
            List[Dict]: Lista con los registros aleatorios seleccionados
        """
        if not self.data_array:
            print("⚠️  Advertencia: No hay datos cargados. Ejecuta load_excel() primero.")
            return []
        
        if n <= 0:
            print("⚠️  Advertencia: El número de registros debe ser mayor a 0.")
            return []
        
        if n >= len(self.data_array):
            print(f"⚠️  Advertencia: Se solicitaron {n} registros pero solo hay {len(self.data_array)} disponibles.")
            print("🔄 Devolviendo todos los registros disponibles.")
            return self.data_array.copy()
        
        # Seleccionar n registros aleatorios sin repetición
        random_records = random.sample(self.data_array, n)
        
        print(f"🎲 Se seleccionaron {len(random_records)} registros aleatorios")
        return random_records
    
    def display_sample_data(self, num_samples: int = 3) -> None:
        """
        Muestra una muestra de los datos cargados para verificación
        
        Args:
            num_samples (int): Número de registros de muestra a mostrar
        """
        if not self.data_array:
            print("⚠️  No hay datos para mostrar")
            return
        
        print(f"\n📊 Muestra de los primeros {min(num_samples, len(self.data_array))} registros:")
        print("=" * 60)
        
        for i, record in enumerate(self.data_array[:num_samples]):
            print(f"\n🆔 Registro ID: {record['id']}")
            print(f"📄 Datos JSON: {json.dumps(record['data'], ensure_ascii=False, indent=2)}")
            
            if i < min(num_samples, len(self.data_array)) - 1:
                print("-" * 40)
    
    def get_total_records(self) -> int:
        """Devuelve el total de registros cargados"""
        return len(self.data_array)
    
    def export_to_json(self, output_file: str) -> None:
        """
        Exporta todos los datos a un archivo JSON
        
        Args:
            output_file (str): Nombre del archivo JSON de salida
        """
        try:
            with open(output_file, 'w', encoding='utf-8') as f:
                json.dump(self.data_array, f, ensure_ascii=False, indent=2)
            print(f"💾 Datos exportados exitosamente a {output_file}")
        except Exception as e:
            print(f"❌ Error al exportar: {str(e)}")


def main():
    """Función principal de ejemplo"""
    
    # Crear instancia del procesador
    processor = ExcelProcessor()
    
    # Ejemplo de uso
    print("🚀 Procesador de Excel iniciado")
    print("=" * 50)
    
    # Solicitar ruta del archivo
    file_path = input("📁 Ingresa la ruta del archivo Excel: ").strip()
    
    # Cargar el Excel
    processor.load_excel(file_path)
    
    if processor.get_total_records() > 0:
        # Mostrar muestra de datos
        processor.display_sample_data()
        
        # Solicitar cantidad de registros aleatorios
        try:
            n = int(input(f"\n🎲 ¿Cuántos registros aleatorios quieres obtener? (Total disponible: {processor.get_total_records()}): "))
            
            # Obtener registros aleatorios
            random_records = processor.get_random_records(n)
            
            # Mostrar resultados
            if random_records:
                print(f"\n🎯 Registros aleatorios seleccionados:")
                print("=" * 50)
                
                for i, record in enumerate(random_records, 1):
                    print(f"\n🔸 Selección #{i} - ID: {record['id']}")
                    print(f"   Datos: {json.dumps(record['data'], ensure_ascii=False)}")
                
                # Preguntar si quiere exportar
                export = input("\n💾 ¿Quieres exportar los registros aleatorios a JSON? (s/n): ").strip().lower()
                if export == 's':
                    output_file = f"registros_aleatorios_{n}.json"
                    with open(output_file, 'w', encoding='utf-8') as f:
                        json.dump(random_records, f, ensure_ascii=False, indent=2)
                    print(f"✅ Registros exportados a {output_file}")
            
        except ValueError:
            print("❌ Error: Ingresa un número válido")
        except KeyboardInterrupt:
            print("\n👋 Programa interrumpido por el usuario")


if __name__ == "__main__":
    main()
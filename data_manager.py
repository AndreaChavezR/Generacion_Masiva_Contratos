# data_manager.py
import pandas as pd
import os
import sys
from datetime import datetime

class DataManager:
    def __init__(self):
        if getattr(sys, 'fozen', False):
            base_dir = sys._MEIPASS
        else:
            base_dir = os.getcwd()
            
        self.excel_file = os.path.join(base_dir,"contratos.xlsx")
        self.initialize_database()  

    def initialize_database(self):
        """Crea el archivo Excel con estructura inicial si no existe"""
        if not os.path.exists(self.excel_file):
            columns = [
                'ID', 'FECHA_REGISTRO', 'GENERADO',
                'CONTRATO_ADQUISICION', 'NO_CONTRATO',
                'BIENES', 'TITULAR_AREA_REQUIRENTE', 
                'TITULAR_AREA', 'PROVEEDOR', 'NOM_PROVEEDOR',
                'CARGO_PROVEEDOR', 'CARGO_AREA_REQUIRENTE',
                'NO_COMPRA','FECHA_CELEBRACIÓN','FUNDAMENTO'
                'FECHA_NOMBRAMIENTO', 'TIPO_ADQUISICION',
                'ADQUISICION', 'NECESIDADES', 'NO_DENOMINACION',
                'NO_OFICIO', 'NO_ESCRITURA_PUBLICA', 
                'FECHA_PUBLICACION', 'TITULAR_NOTARIA', 
                'NO_NOTARIA','DIA', 'NO_MERCANTIL', 'ENTIDAD_FEDERATIVA',
                'OBJETO_SOCIAL', 'PERSONA_FISICA', 
                'CARACTER_PERSONA_FISICA', 'TIPO_DOCUMENTO', 
                'NO_DOCUMENTO', 'INSTITUCION', 'FOLIO_REGISTRO',
                'NO_ESCRITURA', 'FECHA_PUBLICACION2', 
                'INE_NOTARIO', 'RFC', 'NO_CONSTANCIA',
                'FECHA_EXPEDICION', 'ANEXOS', 'CONSTANCIAS', 
                'DOMICILIO', 'NUMERO', 'COLONIA', 'ALCALDIA', 
                'CP', 'TELEFONOS', 'CORREO', 'CALLE', 'NO_EXT',
                'DESCRIPCION_ADQUISICION', 'NO_REQUERIMIENTO', 
                'PARTIDA_PRESUPUESTAL', 'MONTO_AUTORIZADO', 
                'CORREO_1', 'CORREO_2', 'FECHA_ENTREGA', 
                'FECHA_VIGENCIA_ENTREGA', 'VIGENCIA_CONTRATO', 
                'NO_PAG', 'DIAS', 'MES', 'DIA_FIRMA', 'MES_FIRMA',
                'DIRECCION_DE'
            ]
            pd.DataFrame(columns=columns).to_excel(self.excel_file, index=False)
    
    def load_data(self):
        """Carga todos los registros del Excel"""
        return pd.read_excel(self.excel_file, engine='openpyxl')
    
    def save_record(self, data):
        """
        Guarda un nuevo registro en el Excel con validaciones
        Args:
            data (dict): Diccionario con los datos del formulario
        Returns:
            tuple: (success: bool, message: str)
        """
        try:
            # Validar campos obligatorios
            required_fields = [
                'NO_CONTRATO', 'PROVEEDOR', 
                'RFC', 'MONTO_AUTORIZADO'
            ]
            
            for field in required_fields:
                if not data.get(field):
                    return False, f"Campo obligatorio faltante: {field}"
            
            # Validar formato numérico
            if not self.validate_number(data['MONTO_AUTORIZADO']):
                return False, "MONTO_AUTORIZADO debe ser numérico"
            
            # Validar formato de correos
            email_fields = ['CORREO', 'CORREO_1', 'CORREO_2']
            for email_field in email_fields:
                if data.get(email_field):
                    if '@' not in data[email_field] or '.' not in data[email_field].split('@')[-1]:
                        return False, f"Formato inválido en {email_field}"
            
            # Generar datos automáticos
            auto_fields = {
                'ID': datetime.now().strftime("%Y%m%d%H%M%S"),
                'FECHA_REGISTRO': datetime.now().strftime("%Y-%m-%d %H:%M:%S"),
                'GENERADO': 'No'
            }
            data.update(auto_fields)
            
            # Cargar y actualizar datos
            df = self.load_data()
            new_df = pd.DataFrame([data])
            updated_df = pd.concat([df, new_df], ignore_index=True)
            updated_df.to_excel(self.excel_file, index=False)
            
            return True, "Contrato guardado exitosamente"
            
        except Exception as e:
            return False, f"Error al guardar: {str(e)}"
    
    def get_pending_records(self):
        """Obtiene registros no generados"""
        df = self.load_data()
        return df[df['GENERADO'] == 'No']
    
    def mark_as_generated(self, record_id):
        """Marca un registro como generado"""
        try:
            df = self.load_data()
            df.loc[df['ID'] == record_id, 'GENERADO'] = 'Sí'
            df.to_excel(self.excel_file, index=False)
            return True
        except Exception as e:
            print(f"Error al marcar como generado: {str(e)}")
            return False
    
    def validate_number(self, value):
        """Valida que sea un número válido"""
        try:
            float(value)
            return True
        except (ValueError, TypeError):
            return False
        
    def get_pending_records(self):
    #Obtiene todos los registros no generados"""
        df = self.load_data()
        return df[df['GENERADO'] == 'No']
    
if __name__ == "__main__":
    # Prueba de funcionalidad
    dm = DataManager("test_contratos.xlsx")
    
    test_data = {
        'NO_CONTRATO': 'CTO-2024-001',
        'PROVEEDOR': 'Tecnologías Avanzadas S.A.',
        'RFC': 'TEC240101ABC',
        'MONTO_AUTORIZADO': '150000.75',
        'CORREO': 'contacto@tecnologias.com',
        'TELEFONOS': '555-123-4567',
        'FECHA_ENTREGA': '2024-03-20',
        'DIRECCION_DE': 'Av. Tecnología #123'
    }
    
    success, message = dm.save_record(test_data)
    print(f"Resultado: {success} - {message}")
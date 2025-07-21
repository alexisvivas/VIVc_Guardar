import asyncio
import aiohttp
from aiohttp import BasicAuth
import pandas as pd
from datetime import datetime
from urllib.parse import quote
import xml.etree.ElementTree as ET
from typing import List, Dict, Optional
from dataclasses import dataclass, field


@dataclass
class VIVc_Detalle:
    """Clase para representar los detalles de una factura"""
    STP: str
    AccNumber: str
    Objects: str
    Sum: float
    VATCode: str
    ArtCode: str
    Quant: float
    PeriodCode: str


@dataclass
class Encabezado:
    """Clase para representar el encabezado de una factura"""
    InvoiceNr: str
    VeCode: str
    InvDate: datetime
    TransDate: datetime
    PayDeal: str
    Sign: str
    Objects: str
    PayVal: float
    PrelBook: str
    VIVc_Detalle: List[VIVc_Detalle] = field(default_factory=list)


class HttpClass:
    """Clase auxiliar para realizar consultas HTTP"""
    @staticmethod
    async def consulta(url: str, username: str, password: str) -> pd.DataFrame:
        """Realiza una consulta HTTP y retorna los datos como DataFrame"""
        auth = BasicAuth(username, password)
        async with aiohttp.ClientSession() as session:
            async with session.get(url, auth=auth) as response:
                if response.status == 200:
                    # Asumiendo que la respuesta es XML, ajustar según el formato real
                    data = await response.text()
                    # Aquí deberías parsear el XML según el formato real
                    # Por ahora retorno un DataFrame vacío como placeholder
                    return pd.DataFrame()
                else:
                    print(f"Error en consulta: {response.status}")
                    return pd.DataFrame()


async def vivc_guardar():
    """Función principal para procesar y guardar facturas"""
    
    # Configuración de autenticación
    username = "atlejemplo@example.com"
    password = "ejemplo_password"
    auth = BasicAuth(username, password)
    
    # Configuración del cliente HTTP
    timeout = aiohttp.ClientTimeout(total=1200)  # 20 minutos
    
    # URL base
    base_url = "http://x.x.x.x:8xxx/api/1/VIVc?"
    
    # Ruta del archivo Excel
    file_path = r"C:\Users\avivas\Desktop\Consola Emigrar Facturas\SCAJ_ATL_202507 v2.xlsx"
    
    # Lista para almacenar los encabezados
    encabezados = []
    
    # Leer el archivo Excel
    try:
        # Leer la hoja 2 (índice 1 en pandas)
        df = pd.read_excel(file_path, sheet_name=1, engine='openpyxl')
        print(f"Nombre de la hoja: {pd.ExcelFile(file_path).sheet_names[1]}")
        
        # Procesar las filas (comenzando desde la fila 3, índice 2)
        for index, row in df.iloc[2:].iterrows():
            # Obtener valores de las columnas
            invoice_nr = str(row.iloc[1]) if pd.notna(row.iloc[1]) else ""
            
            if not invoice_nr:
                continue
            
            ve_code = str(row.iloc[2]) if pd.notna(row.iloc[2]) else ""
            
            # Procesar fechas
            try:
                inv_date_text = str(row.iloc[3])
                trans_date_text = str(row.iloc[4])
                
                inv_date = pd.to_datetime(inv_date_text).to_pydatetime()
                trans_date = pd.to_datetime(trans_date_text).to_pydatetime()
            except:
                print(f"Formato de fecha inválido en la fila {index + 3}")
                continue
            
            pay_deal = str(row.iloc[5]) if pd.notna(row.iloc[5]) else ""
            sign = str(row.iloc[6]) if pd.notna(row.iloc[6]) else ""
            objects = str(row.iloc[7]) if pd.notna(row.iloc[7]) else ""
            prel_book = str(row.iloc[8]) if pd.notna(row.iloc[8]) else ""
            
            try:
                pay_val = float(row.iloc[9])
            except:
                pay_val = 0.0
            
            # Datos del detalle
            stp = str(row.iloc[10]) if pd.notna(row.iloc[10]) else ""
            acc_number = str(row.iloc[11]) if pd.notna(row.iloc[11]) else ""
            objeto = str(row.iloc[12]) if pd.notna(row.iloc[12]) else ""
            sum_val = float(row.iloc[13]) if pd.notna(row.iloc[13]) else 0.0
            vat_code = str(row.iloc[14]) if pd.notna(row.iloc[14]) else ""
            art_code = str(row.iloc[15]) if pd.notna(row.iloc[15]) else ""
            quant = float(row.iloc[16]) if pd.notna(row.iloc[16]) else 0.0
            period_code = str(row.iloc[17]) if pd.notna(row.iloc[17]) else ""
            
            # Buscar o crear encabezado
            encabezado = next(
                (e for e in encabezados if e.InvoiceNr == invoice_nr and e.VeCode == ve_code),
                None
            )
            
            if encabezado is None:
                encabezado = Encabezado(
                    InvoiceNr=invoice_nr,
                    VeCode=ve_code,
                    InvDate=inv_date,
                    TransDate=trans_date,
                    PayDeal=pay_deal,
                    Sign=sign,
                    Objects=objects,
                    PayVal=pay_val,
                    PrelBook=prel_book
                )
                encabezados.append(encabezado)
            
            # Agregar detalle
            detalle = VIVc_Detalle(
                STP=stp,
                AccNumber=acc_number,
                Objects=objeto,
                Sum=sum_val,
                VATCode=vat_code,
                ArtCode=art_code,
                Quant=quant,
                PeriodCode=period_code
            )
            encabezado.VIVc_Detalle.append(detalle)
    
    except Exception as e:
        print(f"Error al leer el archivo Excel: {e}")
        return
    
    # Mostrar información del primer y último encabezado
    if encabezados:
        primero = encabezados[0]
        ultimo = encabezados[-1]
        print(f"Primer valor (InvoiceNr): {primero.InvoiceNr}")
        print(f"Último valor (InvoiceNr): {ultimo.InvoiceNr}")
    
    # Consultar datos existentes
    url_consulta = "http://x.x.x.x:8xxx/api/1/VIVc?fields=InvoiceNr,SerNr,VECode"
    dt = await HttpClass.consulta(url_consulta, username, password)
    
    # Procesar cada encabezado
    async with aiohttp.ClientSession(auth=auth, timeout=timeout) as session:
        for encabezado in encabezados:
            contador = len(encabezado.VIVc_Detalle) - 1
            
            trans_date = encabezado.TransDate
            inv_date = encabezado.InvDate
            pay_del = float(encabezado.PayVal)
            ok_flag = "0"
            
            # Construir cadena del maestro
            maestro = (
                f"&set_field.InvoiceNr={quote(encabezado.InvoiceNr)}"
                f"&set_field.VECode={quote(encabezado.VeCode)}"
                f"&set_field.InvDate={quote(inv_date.strftime('%Y-%m-%d'))}"
                f"&set_field.TransDate={quote(trans_date.strftime('%Y-%m-%d'))}"
                f"&set_field.PayDeal={quote(encabezado.PayDeal)}"
                f"&set_field.OKPersons={quote(encabezado.Sign)}"
                f"&set_field.Objects={quote(encabezado.Objects)}"
                f"&set_field.OKFlag={quote(ok_flag)}"
                f"&set_field.PrelBook={quote(encabezado.PrelBook)}"
                f"&set_field.PayVal={quote(str(encabezado.PayVal))}"
            )
            
            # Construir lista de items
            items_list = []
            total = 0.0
            
            for i, detalle in enumerate(encabezado.VIVc_Detalle):
                total += detalle.Sum
                
                # Si es el último registro, verificar diferencia
                if i == contador:
                    diferencia = pay_del - total
                    if diferencia != 0:
                        # Ajustar la suma en el último registro
                        sum_ajustada = detalle.Sum + diferencia
                    else:
                        sum_ajustada = detalle.Sum
                else:
                    sum_ajustada = detalle.Sum
                
                item = (
                    f"&set_row_field.{i}.STP={detalle.STP}"
                    f"&set_row_field.{i}.AccNumber={detalle.AccNumber}"
                    f"&set_row_field.{i}.Objects={detalle.Objects}"
                    f"&set_row_field.{i}.Sum={sum_ajustada}"
                    f"&set_row_field.{i}.VATCode={detalle.VATCode}"
                    f"&set_row_field.{i}.Item={detalle.ArtCode}"
                    f"&set_row_field.{i}.qty={detalle.Quant}"
                    f"&set_row_field.{i}.PeriodCode={detalle.PeriodCode}"
                )
                items_list.append(item)
            
            # Verificar si la suma total es igual a PayDel
            if abs((total + (diferencia if 'diferencia' in locals() else 0)) - pay_del) > 0.01:
                print(f"Error: La suma total {total + (diferencia if 'diferencia' in locals() else 0)} "
                      f"no es igual a PayDel {pay_del}")
            
            # Buscar y registrar
            invoice_nr = encabezado.InvoiceNr
            ve_code = encabezado.VeCode
            
            await buscar_y_registrar(
                dt, maestro, items_list, invoice_nr, ve_code, session, base_url
            )


async def buscar_y_registrar(
    dt: pd.DataFrame,
    maestro: str,
    items_list: List[str],
    invoice_nr: str,
    ve_code: str,
    session: aiohttp.ClientSession,
    url: str
):
    """Busca si existe el registro y lo crea si no existe"""
    
    if dt.empty:
        # No hay filas, se procede a registrar
        mensaje = maestro + "".join(items_list)
        headers = {'Content-Type': 'application/x-www-form-urlencoded'}
        
        async with session.post(url, data=mensaje, headers=headers) as response:
            if response.status == 200:
                try:
                    data_xml = await response.text()
                    print(data_xml)
                except Exception as e:
                    print(f"Error: {e}")
    else:
        # Hay filas, se busca en dt
        found_rows = dt[
            (dt['InvoiceNr'] == invoice_nr) & 
            (dt['VECode'] == ve_code)
        ] if 'InvoiceNr' in dt.columns and 'VECode' in dt.columns else pd.DataFrame()
        
        if found_rows.empty:
            # No existe, se hace POST
            print(f"El InvoiceNr {invoice_nr} NO existe en la base de datos, se crea.")
            
            mensaje = maestro + "".join(items_list)
            headers = {'Content-Type': 'application/x-www-form-urlencoded'}
            
            async with session.post(url, data=mensaje, headers=headers) as response:
                if response.status == 200:
                    try:
                        data_xml = await response.text()
                        print(data_xml)
                    except Exception as e:
                        print(f"Error: {e}")
        else:
            # Ya existe
            pass


# Función principal para ejecutar el código asíncrono
async def main():
    await vivc_guardar()


if __name__ == "__main__":
    # Ejecutar la función principal
    asyncio.run(main())

import zipfile
import os
import pandas as pd
import tkinter as tk
from tkinter import filedialog, messagebox
import xml.etree.ElementTree as ET

def kmz_to_excel():
    # Inicializar janela do Tkinter
    root = tk.Tk()
    root.withdraw()

    # Selecionar arquivo KMZ
    kmz_file = filedialog.askopenfilename(
        title="Selecione o arquivo KMZ",
        filetypes=[("Arquivos KMZ", "*.kmz")]
    )
    if not kmz_file:
        messagebox.showerror("Erro", "Nenhum arquivo KMZ foi selecionado.")
        return

    # Selecionar onde salvar o Excel
    excel_file = filedialog.asksaveasfilename(
        title="Salvar arquivo Excel como...",
        defaultextension=".xlsx",
        filetypes=[("Planilhas Excel", "*.xlsx")]
    )
    if not excel_file:
        messagebox.showerror("Erro", "Nenhum destino de salvamento foi selecionado.")
        return

    # Extrair o KML de dentro do KMZ
    try:
        with zipfile.ZipFile(kmz_file, 'r') as kmz:
            kml_files = [f for f in kmz.namelist() if f.endswith('.kml')]
            if not kml_files:
                raise Exception("Nenhum arquivo KML encontrado dentro do KMZ.")
            kml_filename = kml_files[0]
            kmz.extract(kml_filename)
    except Exception as e:
        messagebox.showerror("Erro", f"Erro ao extrair KMZ:\n{e}")
        return

    # Parsear o KML e extrair pontos
    try:
        tree = ET.parse(kml_filename)
        root_kml = tree.getroot()

        # Espaço de nomes KML
        ns = {'kml': 'http://www.opengis.net/kml/2.2'}

        placemarks = root_kml.findall(".//kml:Placemark", ns)

        data = []
        for placemark in placemarks:
            name = placemark.find("kml:name", ns)
            point = placemark.find(".//kml:Point/kml:coordinates", ns)

            if point is not None and name is not None:
                coord_text = point.text.strip()
                lon, lat, *_ = coord_text.split(',')
                data.append({
                    "Nome": name.text,
                    "Latitude": float(lat),
                    "Longitude": float(lon)
                })
    except Exception as e:
        messagebox.showerror("Erro", f"Erro ao ler o KML:\n{e}")
        return
    finally:
        os.remove(kml_filename)  # Apaga o kml temporário

    # Salvar no Excel
    try:
        df = pd.DataFrame(data)
        df.to_excel(excel_file, index=False)
        messagebox.showinfo("Sucesso", f"✅ Excel salvo com sucesso:\n{excel_file}")
    except Exception as e:
        messagebox.showerror("Erro", f"Erro ao salvar Excel:\n{e}")

# Executar
if __name__ == "__main__":
    kmz_to_excel()

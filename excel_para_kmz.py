import pandas as pd
import simplekml
import zipfile
import os
import tkinter as tk
from tkinter import filedialog, messagebox

def excel_to_kmz():
    # Inicializar a janela do Tkinter
    root = tk.Tk()
    root.withdraw()  # Oculta a janela principal

    # Selecionar o arquivo Excel
    excel_file = filedialog.askopenfilename(
        title="Selecione o arquivo Excel",
        filetypes=[("Planilhas Excel", "*.xlsx *.xls")]
    )

    if not excel_file:
        messagebox.showerror("Erro", "Nenhum arquivo Excel foi selecionado.")
        return

    # Selecionar onde salvar o KMZ
    kmz_output_file = filedialog.asksaveasfilename(
        title="Salvar arquivo KMZ como...",
        defaultextension=".kmz",
        filetypes=[("Arquivo KMZ", "*.kmz")]
    )

    if not kmz_output_file:
        messagebox.showerror("Erro", "Nenhum destino de salvamento foi selecionado.")
        return

    # Ler o arquivo Excel
    try:
        df = pd.read_excel(excel_file)
    except Exception as e:
        messagebox.showerror("Erro", f"Erro ao ler o Excel:\n{e}")
        return

    # Substituir vírgulas por pontos e converter para float
    try:
        df["Latitude"] = df["Latitude"].astype(str).str.replace(",", ".", regex=False).astype(float)
        df["Longitude"] = df["Longitude"].astype(str).str.replace(",", ".", regex=False).astype(float)
    except Exception as e:
        messagebox.showerror("Erro", f"Erro ao processar coordenadas:\n{e}")
        return

    if df["Latitude"].isnull().any() or df["Longitude"].isnull().any():
        messagebox.showerror("Erro", "Algumas coordenadas são inválidas. Verifique o Excel.")
        return

    # Criar KML
    kml = simplekml.Kml()
    for index, row in df.iterrows():
        nome = f"Local {index + 1}"
        kml.newpoint(name=nome, coords=[(row["Longitude"], row["Latitude"])])

    # Salvar arquivo temporário .kml
    temp_kml = 'temp.kml'
    kml.save(temp_kml)

    # Compactar como KMZ
    try:
        with zipfile.ZipFile(kmz_output_file, 'w', zipfile.ZIP_DEFLATED) as kmz:
            kmz.write(temp_kml, os.path.basename(temp_kml))
        os.remove(temp_kml)
        messagebox.showinfo("Sucesso", f"✅ KMZ gerado com sucesso:\n{kmz_output_file}")
    except Exception as e:
        messagebox.showerror("Erro", f"Erro ao gerar KMZ:\n{e}")

# Executar
if __name__ == "__main__":
    excel_to_kmz()

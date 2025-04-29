import pandas as pd
import numpy as np
import matplotlib.pyplot as plt
import openpyxl
import tkinter as tk
from tkinter import messagebox

def test_libraries():
    print("Testando bibliotecas...")

    # Teste do pandas
    try:
        df = pd.DataFrame({'A': [1, 2, 3], 'B': [4, 5, 6]})
        print("pandas: DataFrame criado com sucesso")
        print(df)
    except Exception as e:
        print(f"Erro no pandas: {e}")

    # Teste do numpy
    try:
        arr = np.array([1, 2, 3])
        print("numpy: Array criado com sucesso")
        print(arr)
    except Exception as e:
        print(f"Erro no numpy: {e}")

    # Teste do matplotlib
    try:
        x = np.linspace(0, 10, 100)
        y = np.sin(x)
        plt.figure()
        plt.plot(x, y)
        plt.title("Teste Matplotlib")
        plt.savefig('test_plot.png')
        plt.close()
        print("matplotlib: Plot salvo como test_plot.png")
    except Exception as e:
        print(f"Erro no matplotlib: {e}")

    # Teste do openpyxl
    try:
        wb = openpyxl.Workbook()
        ws = wb.active
        ws['A1'] = "Teste"
        wb.save('test.xlsx')
        print("openpyxl: Arquivo Excel salvo como test.xlsx")
    except Exception as e:
        print(f"Erro no openpyxl: {e}")

    # Teste do tkinter
    try:
        root = tk.Tk()
        root.title("Teste Tkinter")
        label = tk.Label(root, text="Tkinter funcionando!")
        label.pack(pady=10)
        root.after(2000, root.destroy)  # Fecha a janela após 2 segundos
        root.mainloop()
        print("tkinter: Janela exibida com sucesso")
    except Exception as e:
        print(f"Erro no tkinter: {e}")

if __name__ == "__main__":
    test_libraries()
    print("Testes concluídos!")
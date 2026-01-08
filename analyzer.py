import pandas as pd
import matplotlib.pyplot as plt
from pathlib import Path

path_input = input("Masukan File Path Data Excel: ")
path_clear = Path(path_input)

# Fungsi Baca Data
def load_data():
    try:
        df = pd.read_excel(path_clear)
        print(df.head(7))
        return df
    except Exception as e:
        print("File Tidak Ditemukan!")
        return

# Fungsi Hitung Statistik Data 
def calculate_data(df):
    if df is None:
        return
    all_stat = {}    
    for column in df.columns[1:]:
        print(f"\nStatistik data untuk kolom {column}")

        statistic = {
            "mean": df[column].mean(),
            "std": df[column].std(),
            "min": df[column].min(),
            "max": df[column].max()
        }
        all_stat[column] = statistic

        for stat_name, values in statistic.items():
                print(f"{stat_name}:{values:.2f}")

    return all_stat

# Fungsi Gambar Line Chart
def line_chart(df):
    script_folder = Path(__file__).parent
    output_folder = script_folder / "Output"
    output_folder.mkdir(exist_ok=True)

    x = df.iloc[:, 0]
    y1 = df.iloc[:, 1]
    y2 = df.iloc[:, 2]

    plt.plot(x, y1, label= df.columns[1])
    plt.plot(x, y2, label= df.columns[2])
    plt.xlabel(df.columns[0])
    plt.title("Line Chart Data")
    plt.legend()

    save_path = output_folder / "Line Chart Data.png"
    plt.savefig(save_path)

# Fungsi Export Hasil Excel

def make_excel(all_statistic):
    script_folder = Path(__file__).parent
    output_folder = script_folder / "Output"
    output_folder.mkdir(exist_ok=True)

    save_path = output_folder / "Statistik Data.xlsx"
    with pd.ExcelWriter(save_path, engine="openpyxl") as writter:
        statistic_df = pd.DataFrame(all_statistic)
        statistic_df.to_excel(writter, sheet_name="Statistik Data", index=True)
    
def main():
    data = load_data()
    if data is not None:
        calculate_data(data)
        all_statistic = calculate_data(data)
        line_chart(data)
        make_excel(all_statistic)

if __name__ == "__main__":
    main()
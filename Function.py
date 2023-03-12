from io import StringIO
import os, sys, subprocess

import matplotlib.pyplot as plt

plt.style.use("seaborn-white")
plt.rcParams.update({"figure.max_open_warning": 0})

# from matplotlib.gridspec import GridSpec
from matplotlib.backends.backend_pdf import PdfPages

import psutil
import numpy as np
import pandas as pd

pd.set_option("display.max_rows", None)
pd.set_option("display.width", None)
pd.set_option("display.max_columns", None)
pd.set_option("display.max_colwidth", None)

import xlwings as xw

import tkinter as tk
import tkinter.messagebox as msg
from tkinter import filedialog


def return_print(*prt_str):
    io = StringIO()
    print(*prt_str, file=io, sep=",", end="")
    return io.getvalue()


def save_multi_image(f_name):
    pp = PdfPages(f_name)
    fig_nums = plt.get_fignums()
    figs = [plt.figure(n) for n in fig_nums]
    for fig in figs:
        fig.savefig(pp, format="pdf")
    pp.close()


def open_file(filename):
    if sys.platform == "win32":
        os.startfile(filename)
    else:
        opener = "open" if sys.platform == "darwin" else "xdg-open"
        subprocess.call([opener, filename])


def add_file(Entry_file_path):
    Entry_file_path.delete(0, tk.END)
    filename = filedialog.askopenfilenames(
        initialdir="C:\Labtest\Report\Spring",
        title="Select file",
        filetypes=(("All fiels", "*.*"), ("Excel files", "*.xlsx")),
    )
    Entry_file_path.insert(tk.END, filename)


def Set_fig(title):
    fig = plt.figure(figsize=(30.6, 15.9))
    fig.suptitle((f"{title} ALL CH"), fontsize=25)
    ax1 = plt.subplot2grid((2, 5), (0, 0), colspan=2)
    ax2 = plt.subplot2grid((2, 5), (0, 2), colspan=3)
    ax3 = plt.subplot2grid((2, 5), (1, 0), colspan=3)
    ax4 = plt.subplot2grid((2, 5), (1, 3), colspan=2)

    ax1.set_title("TX Power")
    ax1.set_ylim(20, 26)  # set_ylim(bottom, top)
    ax1.set_xlabel("Channel", fontsize=10)
    ax1.set_ylabel("TX Power", fontsize=10)
    ax1.grid(True, color="black", alpha=0.3, linestyle="--")

    ax2.set_title("Max Power RX Sensitivity")
    ax2.set_ylim(-112, -95)
    ax2.set_xlabel("Channel", fontsize=10)
    ax2.set_ylabel("RX Sensitivity", fontsize=10)
    ax2.grid(True, color="black", alpha=0.3, linestyle="--")

    ax3.set_title("0dBm RX Sensitivity")
    ax3.set_ylim(-112, -95)
    ax3.set_xlabel("Channel", fontsize=10)
    ax3.set_ylabel("RX Sensitivity", fontsize=10)
    ax3.grid(True, color="black", alpha=0.3, linestyle="--")

    ax4.set_title("Maxiumum Sens. Desensitization")
    ax4.set_ylim(-1, 3)
    ax4.set_xlabel("Channel", fontsize=10)
    ax4.set_ylabel("RX Sensitivity", fontsize=10)
    ax4.grid(True, color="black", alpha=0.3, linestyle="--")
    ax4.axhline(y=1, linestyle="dashdot", color="red", label="Target")  # Spec 기준선 Drwaing

    return fig, ax1, ax2, ax3, ax4


def LTE_Sens_drawing(Win_GUI, Entry_file_path, text_area):
    try:
        filename = Entry_file_path.get()
        if filename:
            f_name = f"{os.path.splitext(filename)[0]}.pdf"  # filename을 확장자를 지운 후 pdf 확장자로 지정
            TestItem = []
            text_area.insert(tk.END, f"Open Excel File ... \n")
            # openpyxl
            # workbook = xl.load_workbook(r"{}".format(filename))
            # TestItem = [sheet.title for sheet in workbook.worksheets if sheet.sheet_state == "visible"]
            # xlwings
            app = xw.App(visible=False)
            wb = app.books.open(filename)
            # You need to install xlrd==1.2.0 to get the support for xlsx excel format.
            TestItem = [sheet.name for sheet in wb.sheets if sheet.api.Visible == -1]

        dict_value = ["WCDMA ALL CHANNEL", "LTE"]
        Plot_list = [s for s in TestItem if any(xs in s for xs in dict_value)]

        for c, i in enumerate(Plot_list, start=1):
            text_area.insert(tk.END, f"Item count {c:<17}|    {i:<20}\n")
            text_area.see(tk.END)
        text_area.insert(tk.END, f"Loading Workbook Done\n")
        text_area.insert(tk.END, "=" * 73)
        text_area.insert(tk.END, "\n")
        text_area.see(tk.END)
        df_Data = {}
        for sheet_name in Plot_list:
            sheet = wb.sheets[sheet_name]
            df_Data[sheet_name] = sheet.used_range.options(pd.DataFrame, index=False).value

        # 엑셀 프로그램 종료
        wb.close()
        app = xw.apps.active
        for proc in psutil.process_iter():
            if proc.name() == "EXCEL.EXE":
                proc.kill()

        for i in range(len(Plot_list)):
            if any("LTE" in c for c in Plot_list):
                text_area.insert(tk.END, f"Drawing {Plot_list[i]:<20}|    ")
                df_Band = pd.DataFrame(df_Data[Plot_list[i]])
                df_Band = df_Band.drop(index=[0, 1, 2, 3, 4, 5, 6, 7]).reset_index(drop=True)
                df_Band.columns = df_Band.iloc[0]
                Plot_list[i] = Plot_list[i].replace(" ALL CH", "")

                df_BW = (
                    df_Band["BW"].drop_duplicates(keep="first").reset_index(drop=True).dropna().values.tolist()
                )  # 중복값 삭제, NaN Drop, index reset
                df_BW = (return_print(*df_BW)).split(",")  # 각 밴드의 측정 BW 데이터 추출
                fig, ax1, ax2, ax3, ax4 = Set_fig(Plot_list[i])
                count = 0

                for Bandwidth in df_BW:

                    Data_of_BW = df_Band[df_Band["BW"] == Bandwidth]
                    Data_of_BW = Data_of_BW.replace("-", np.nan)  # 미측정으로 인한 오류 (-) 해결을 위해 Nan으로 대체하고 아래에서 dropna 로 제거

                    if Bandwidth == "BW":
                        df_CH = Data_of_BW.reset_index(drop=True).iloc[:, 9:]
                        df_CH.index = df_BW[1::1]  # df_BW list slicing [start:stop:step]
                        df_CH = df_CH.transpose().reset_index(drop=True)
                        continue
                    else:
                        df_TXL = (
                            Data_of_BW[Data_of_BW["Test Item"].str.contains("6.2.2 Maximum Output Power_RB")]
                            .iloc[2:3, 9:]
                            .dropna(axis=1)
                            .iloc[0]
                        )
                        df_TXL.index = df_CH[Bandwidth].dropna()
                        df_TXR = (
                            Data_of_BW[Data_of_BW["Test Item"].str.contains("6.2.2 Maximum Output Power_RB")]
                            .iloc[3:4, 9:]
                            .dropna(axis=1)
                            .iloc[0]
                        )
                        df_TXR.index = df_CH[Bandwidth].dropna()
                        df_Sen = (
                            Data_of_BW[Data_of_BW["Test Item"].str.contains("7.3 Reference Sensitivity level")]
                            .iloc[:12, 9:]
                            .dropna(axis=1)
                            .max()
                        )
                        df_Sen.index = df_CH[Bandwidth].dropna()
                        # 특정 subplot을 Twinx 설정 시 ax2=ax1[1].twinx()
                        ax1.plot(df_TXL, marker=".", label="TX {}M RB Low".format(Bandwidth))
                        ax1.plot(df_TXR, marker=".", label="TX {}M RB High".format(Bandwidth))
                        ax1.legend(fontsize=8, frameon=False, loc="lower center", ncol=6)
                        ax2.plot(df_Sen, marker=".", label="{}MHz RX Sens.".format(Bandwidth))
                        ax2.legend(fontsize=10, frameon=False, loc="lower center", ncol=6)

                        if any(Data_of_BW["Test Item"].str.contains("7.3 Ref Sens level@ UE 0dBm")):
                            df_Sen_0dBm = (
                                Data_of_BW[Data_of_BW["Test Item"].str.contains("7.3 Ref Sens level@ UE 0dBm")]
                                .iloc[:12, 9:]
                                .dropna(axis=1)
                                .max()
                            )
                            df_Sen_0dBm.index = df_CH[Bandwidth].dropna()
                            df_Desens = df_Sen - df_Sen_0dBm
                            df_Desens.index = df_CH[Bandwidth].dropna()

                            ax3.plot(df_Sen_0dBm, marker=".", label="{}MHz RX Sens".format(Bandwidth))
                            ax3.legend(fontsize=10, frameon=False, loc="lower center", ncol=6)

                            ax4.plot(df_Desens, marker=".", label="Desens {}MHz".format(Bandwidth))
                            ax4.legend(fontsize=8, frameon=False, loc="lower center", ncol=7)
                        else:
                            if count == 0:  # 첫 BW에서만 ax3, ax4를 지우고 count를 1로 올린다. 이후 실행되지 않음
                                fig.set_size_inches(30.6, 7.45)
                                ax3.remove()
                                ax4.remove()
                                ax1 = plt.subplot2grid((1, 5), (0, 0), colspan=2)
                                ax2 = plt.subplot2grid((1, 5), (0, 2), colspan=3)
                                ax1.set_title("TX Power")
                                ax1.set_ylim(20, 26)  # set_ylim(bottom, top)
                                ax1.set_xlabel("Channel", fontsize=10)
                                ax1.set_ylabel("TX Power", fontsize=10)
                                ax1.grid(True, color="black", alpha=0.3, linestyle="--")
                                ax1.plot(df_TXL, marker=".", label="TX {}M RB Low".format(Bandwidth))
                                ax1.plot(df_TXR, marker=".", label="TX {}M RB High".format(Bandwidth))
                                ax1.legend(fontsize=8, frameon=False, loc="lower center", ncol=6)

                                ax2.set_title("Max Power RX Sensitivity")
                                ax2.set_ylim(-112, -95)
                                ax2.set_xlabel("Channel", fontsize=10)
                                ax2.set_ylabel("RX Sensitivity", fontsize=10)
                                ax2.grid(True, color="black", alpha=0.3, linestyle="--")
                                ax2.plot(df_Sen, marker=".", label="{}MHz RX Sens.".format(Bandwidth))
                                ax2.legend(fontsize=10, frameon=False, loc="lower center", ncol=6)
                                plt.subplots_adjust(wspace=0.3, hspace=0.3)
                                count += 1
                        plt.tight_layout()
                    plt.tight_layout()
                text_area.insert(tk.END, f"Done")
                text_area.insert(tk.END, f"    |    Saving Image")
                save_multi_image(f_name)
                text_area.insert(tk.END, f"    |    Done\n")
                text_area.see(tk.END)

            elif any("WCDMA ALL CHANNEL" in c for c in Plot_list):
                df_Band = pd.DataFrame(df_Data[Plot_list[i]])
                df_Band = df_Band.drop(index=[0, 1, 2, 3, 4, 5, 6]).dropna(how="all", axis=0).reset_index(drop=True)
                df_ItemList = df_Band["Samsung Lab Test Report"]

                list_Range = []
                Band_list = []

                for c, val in enumerate(df_ItemList):  # BAND 구분되는 위치를 먼저 확인
                    if "BAND" in val or "Band" in val:
                        list_Range.append(c)
                        Band_list.append(val)

                Item_Count = list_Range[1] - list_Range[0]
                Plot_list[i] = Plot_list[i].replace(" ALL CHANNEL", "")

                for k in range(len(list_Range)):  # 데이터 갯수만큼 실행하기 위해 for k in list_Range 대신 len(list_Range) 사용
                    df_TestBand = df_Band.iloc[list_Range[k] : list_Range[k] + Item_Count, :].reset_index(drop=True)
                    # Nan이 행/열 모두 있어서 2번 실행함
                    df_TestBand = df_TestBand.replace("-", np.nan).iloc[1:Item_Count, :].dropna(how="all", axis=1)
                    df_TestBand = df_TestBand.dropna().reset_index(drop=True)
                    df_TestBand.columns = df_TestBand.iloc[0]

                    text_area.insert(tk.END, f"Drawing {Band_list[k]:<20}|    ")
                    fig, ax1, ax2, ax3, ax4 = Set_fig(Band_list[k])

                    df_CH = df_TestBand.columns.to_list()[4:]
                    df_TX = (
                        df_TestBand[df_TestBand["Test Item"].str.contains("5.2 Maximum output power")]
                        .iloc[:, 4:]
                        .dropna(axis=1)
                        .iloc[0]
                    )
                    df_TX.index = df_CH
                    df_Sen = (
                        df_TestBand[df_TestBand["Test Item"].str.contains("6.2 Reference sensitivity")]
                        .iloc[:, 4:]
                        .dropna(axis=1)
                        .max()
                    )
                    Max_dBm = float(df_TestBand.loc[df_TestBand["Test Item"] == "6.2 Reference sensitivity", "Max"])
                    df_Sen.index = df_CH

                    count = 0
                    if any(df_TestBand["Test Item"].str.contains("6.2 Reference sensitivity UE. 0dBm")):
                        df_Sen_0dBm = (
                            df_TestBand[df_TestBand["Test Item"].str.contains("6.2 Reference sensitivity UE. 0dBm")]
                            .iloc[:, 4:]
                            .dropna(axis=1)
                            .max()
                        )
                        Zero_dBm = float(
                            df_TestBand.loc[df_TestBand["Test Item"] == "6.2 Reference sensitivity UE. 0dBm", "Max"]
                        )
                        df_Sen_0dBm.index = df_CH

                        df_Msd = df_Sen - df_Sen_0dBm
                        df_Msd.index = df_CH
                        # 특정 subplot을 Twinx 설정 시 ax2=ax1[1].twinx()
                        ax1.plot(df_TX, marker=".", label=f"TX Power")
                        ax1.legend(fontsize=8, frameon=False, loc="lower center", ncol=6)

                        ax2.plot(df_Sen, marker=".", label=f"RX Sensitivity @ Max.P")
                        ax2.legend(fontsize=10, frameon=False, loc="lower center", ncol=6)
                        ax2.set_ylim(-115, -100)
                        ax2.axhline(y=Max_dBm, linestyle="dashdot", color="red", label="Target")  # Spec 기준선 Drwaing

                        ax3.plot(df_Sen_0dBm, marker=".", label=f"RX Sensitivity @ 0dBm")
                        ax3.legend(fontsize=10, frameon=False, loc="lower center", ncol=6)
                        ax3.set_ylim(-115, -100)
                        ax3.axhline(y=Zero_dBm, linestyle="dashdot", color="red", label="Target")  # Spec 기준선 Drwaing

                        ax4.plot(df_Msd, marker=".", label=f"Max. Sensivitiy Desensitization")
                        ax4.legend(fontsize=10, frameon=False, loc="lower center", ncol=6)
                    else:
                        if count == 0:  # 첫 BW에서만 ax3, ax4를 지우고 count를 1로 올린다. 이후 실행되지 않음
                            fig.set_size_inches(30.6, 7.45)
                            ax3.remove()
                            ax4.remove()
                            ax1 = plt.subplot2grid((1, 5), (0, 0), colspan=2)
                            ax2 = plt.subplot2grid((1, 5), (0, 2), colspan=3)
                            ax1.set_title("TX Power")
                            ax1.set_ylim(20, 26)  # set_ylim(bottom, top)
                            ax1.set_xlabel("Channel", fontsize=10)
                            ax1.set_ylabel("TX Power", fontsize=10)
                            ax1.grid(True, color="black", alpha=0.3, linestyle="--")
                            ax1.plot(df_TX, marker=".", label=f"TX Power")
                            ax1.legend(fontsize=8, frameon=False, loc="lower center", ncol=6)

                            ax2.set_title("Max Power RX Sensitivity")
                            ax2.set_ylim(-115, -100)
                            ax2.set_xlabel("Channel", fontsize=10)
                            ax2.set_ylabel("RX Sensitivity", fontsize=10)
                            ax2.grid(True, color="black", alpha=0.3, linestyle="--")
                            ax2.plot(df_Sen, marker=".", label=f"RX Sensitivity @ Max.P")
                            ax2.legend(fontsize=10, frameon=False, loc="lower center", ncol=6)
                            ax2.axhline(
                                y=Max_dBm, linestyle="dashdot", color="red", label="Target"
                            )  # Spec 기준선 Drwaing
                            plt.subplots_adjust(wspace=0.3, hspace=0.3)
                            count += 1
                    plt.tight_layout()
                    text_area.insert(tk.END, f"Done")
                    text_area.insert(tk.END, f"    |    Saving Image")
                    save_multi_image(f_name)
                    text_area.insert(tk.END, f"    |    Done\n")
                    text_area.see(tk.END)

        open_file(f_name)
        Win_GUI.destroy()
        Win_GUI.quit()

    except Exception as e:
        msg.showwarning("Warning", e)

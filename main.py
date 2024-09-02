import pandas as pd
import streamlit as st
import plotly.graph_objects as go
from scipy.signal import find_peaks
from io import BytesIO

def to_excel(df):
    output = BytesIO()
    with pd.ExcelWriter(output, engine='openpyxl') as writer:
        df.to_excel(writer, index=False, sheet_name='Sonuclar')
    return output.getvalue()

def clean_column_name(col_name, index):
    """Kolon adlarını temizle ve NaN/boş string için varsayılan adlar ver."""
    if pd.notna(col_name) and col_name.strip() != "":
        return col_name.replace("_x000D_", "").replace("\n", "").strip()
    return f"Unnamed_{index}"

# Sidebar'da dosya yükleme işlemi
uploaded_file = st.sidebar.file_uploader("Excel dosyası yükleyin", type=["xlsx", "xls"])

if uploaded_file is not None:
    try:
        # Excel dosyasını oku
        df = pd.read_excel(uploaded_file, sheet_name=None)  # Tüm sayfaları oku
        sheet_names = list(df.keys())  # Sayfa isimlerini al
        selected_sheets = st.sidebar.multiselect("Sayfaları seçin", sheet_names)

        if selected_sheets:
            first_sheet_df = df[selected_sheets[0]]
            time_index = first_sheet_df[first_sheet_df.iloc[:, 1] == 'Time_x000D_\n[s]'].index
            if not time_index.empty:
                index_num = int(time_index[0])
                df_selectbox = first_sheet_df.iloc[index_num:].reset_index(drop=True)
                df_selectbox.columns = [clean_column_name(col, i) for i, col in enumerate(df_selectbox.iloc[0])]
                df_selectbox = df_selectbox[1:].reset_index(drop=True)
                time_column = [col for col in df_selectbox.columns if 'Time' in col]
                if time_column:
                    time_column = time_column[0]
                    valid_columns = [col for col in df_selectbox.columns if
                                     not col.startswith("Unnamed") and col != time_column]
                    current_column = st.sidebar.selectbox("Akım Sütununu Seçin", valid_columns)
                    voltage_column = st.sidebar.selectbox("Gerilim Sütununu Seçin", valid_columns)

                    # R ve L değerlerini kullanıcıdan bir kez alın
                    R = st.sidebar.number_input("Direnç (R) Değeri", min_value=0.0, step=0.01)
                    L = st.sidebar.number_input("İndüktans (L) Değeri", min_value=0.0, step=0.01)

                    results = []
                    for sheet_name in selected_sheets:
                        # Her sayfa için bir container oluşturun
                        with st.container(border=True):
                            st.subheader(f"{sheet_name}")
                            sheet_df = df[sheet_name]
                            time_index = sheet_df[sheet_df.iloc[:, 1] == 'Time_x000D_\n[s]'].index
                            if not time_index.empty:
                                index_num = int(time_index[0])
                                df_sheet = sheet_df.iloc[index_num:].reset_index(drop=True)
                                df_sheet.columns = [clean_column_name(col, i) for i, col in enumerate(df_sheet.iloc[0])]
                                df_sheet = df_sheet[1:].reset_index(drop=True)
                                if time_column in df_sheet.columns and current_column in df_sheet.columns and voltage_column in df_sheet.columns:
                                    selected_df = df_sheet[[time_column, current_column, voltage_column]]
                                    selected_df.columns = ["Zaman", "Akım", "Gerilim"]
                                    selected_df["Zaman"] = (selected_df["Zaman"] * 1000).round(2)
                                    selected_df=selected_df[selected_df["Akım"] > 0.0005]
                                    #st.write(selected_df)
                                    graph_df = selected_df[selected_df["Akım"] > 0.0005]
                                    peaks, _ = find_peaks(-graph_df["Akım"].values,prominence=0.03)
                                    peak_times = graph_df["Zaman"].iloc[peaks].values
                                    user_time_option = st.radio("Zaman Değeri Seçin", ["Yerel Minimumlardan Seç", "Manuel Zaman Girişi"], key=f"user_time_option_{sheet_name}")
                                    if user_time_option == "Yerel Minimumlardan Seç":
                                        selected_time = st.selectbox("Yerel Minimumlardan Zaman Seçin", peak_times, key=f"select_time_{sheet_name}")
                                    else:
                                        selected_time = st.number_input("Manuel Zaman Değeri Girin", min_value=0.0, step=0.1, key=f"manual_time_{sheet_name}")
                                    filtered_df = selected_df[selected_df["Zaman"] < selected_time]
                                    last_time = filtered_df["Zaman"].max()
                                    #st.write(last_time)
                                    #st.write(selected_df)
                                    #st.write(filtered_df)
                                    filtered_df["Çarpım"] = filtered_df["Akım"] * filtered_df["Gerilim"] * 0.1
                                    filtered_df["R x I^2"] = (filtered_df["Akım"] ** 2) * R * 0.1
                                    selected_time_current = selected_df[selected_df["Zaman"] == last_time]["Akım"].values
                                    if len(selected_time_current) > 0:
                                        selected_time_current = selected_time_current[0]
                                        calculated_value = (selected_time_current ** 2) * L * 0.5 / 1000
                                        results.append({
                                            "Test İsmi": sheet_name,
                                            "Akım ve Gerilim": round(filtered_df["Çarpım"].sum() / 1000, 2),
                                            "Akım ve Direnc": round(filtered_df["R x I^2"].sum() / 1000, 2),
                                            "Sonuc": round(filtered_df["Çarpım"].sum() / 1000 - filtered_df["R x I^2"].sum() / 1000 - calculated_value, 2),
                                                "Bobin Akımı": round(filtered_df["Gerilim"].mode().values[0], 2)
                                        })

                                        # Akım-Zaman grafiği oluştur
                                        fig = go.Figure()
                                        fig.add_trace(go.Scatter(x=graph_df["Zaman"], y=graph_df["Akım"], mode='lines', name='Akım'))
                                        fig.add_trace(go.Scatter(
                                            x=graph_df["Zaman"].iloc[peaks],
                                            y=graph_df["Akım"].iloc[peaks],
                                            mode='markers',
                                            marker=dict(color='red', size=8),
                                            name='Yerel Minimumlar'
                                        ))
                                        fig.update_layout(
                                            title=f"{sheet_name} - Zaman vs Akım Grafiği",
                                            xaxis_title="Zaman (ms)",
                                            yaxis_title="Akım",
                                            template="plotly_white"
                                        )
                                        st.plotly_chart(fig)
                                        st.write(filtered_df)
                                        excel_data1 = to_excel(filtered_df)
                                        st.download_button(
                                            label="Sonuçları Excel olarak indir",
                                            data=excel_data1,
                                            file_name=f'{sheet_name}_sonuc_listesi.xlsx',
                                            mime='application/vnd.openxmlformats-officedocument.spreadsheetml.sheet'
                                        )
                    with st.container(border=True):
                    # Sonuçları DataFrame olarak oluştur

                        results_df = pd.DataFrame(results)
                        results_df = results_df.sort_values(by="Bobin Akımı", ascending=False)

                        # DataFrame'i ekranda göster
                        st.subheader("Hesaplama Sonuçları:")
                        st.write(results_df)
                        excel_data = to_excel(results_df)
                        st.download_button(
                            label="Sonuçları Excel olarak indir",
                            data=excel_data,
                            file_name=f'{uploaded_file.name}_sonuc_listesi.xlsx',
                            mime='application/vnd.openxmlformats-officedocument.spreadsheetml.sheet'
                        )
                        # Çizgi grafiği için hazırlık
                        fig = go.Figure()
                        fig.add_trace(go.Scatter(
                            x=results_df["Bobin Akımı"],
                            y=results_df["Sonuc"],
                            mode='lines+markers+text',
                            name='Son Hesaplanan Değer'
                        ))
                        fig.update_layout(
                            title="En Çok Tekrar Eden Gerilim Değeri vs. Son Hesaplanan Değer",
                            xaxis_title="En Çok Tekrar Eden Gerilim Değeri",
                            yaxis_title="Son Hesaplanan Değer",
                            template="plotly_white"
                        )
                        st.plotly_chart(fig)

                else:
                    st.write("Time sütunu bulunamadı.")

            else:
                st.write("'Time_x000D_\\n[s]' hücresi bulunamadı.")

        else:
            st.write("Lütfen en az bir sayfa seçin.")

    except Exception as e:
        st.error(f"Bir hata oluştu: {e}")

else:
    st.write("Lütfen bir Excel dosyası yükleyin.")

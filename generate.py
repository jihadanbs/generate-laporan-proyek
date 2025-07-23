import os
import re
import shutil
import zipfile
from PIL import Image
from docxtpl import DocxTemplate, InlineImage
from docx.shared import Cm


def parse_whatsapp_text_v2(raw_text):
    print("1. Menganalisis teks dari WhatsApp...")
    raw_text = raw_text.strip()
    grup_chunks = re.split(r"\d+\.\s*group", raw_text, flags=re.IGNORECASE)
    parsed_data = []

    for chunk in grup_chunks:
        if not chunk.strip():
            continue
        nama_grup_match = re.match(r"([^=]+)", chunk)
        if not nama_grup_match:
            continue
        nama_grup = nama_grup_match.group(1).strip().title()
        detail_pekerjaan_bersih = []

        if "sipil" in nama_grup.lower():
            current_work_section = re.search(
                r"Adapun pekerjaan yang dikerjakan.*?adalah\s*=(.*)",
                chunk,
                flags=re.DOTALL | re.IGNORECASE,
            )
            if current_work_section:
                tasks_text = current_work_section.group(1)
                task_list_raw = re.split(r"\(\s*[A-Z]\s*\)\.?", tasks_text)
                for task_raw in task_list_raw:
                    if not task_raw.strip():
                        continue
                    jumlah_pekerja_match = re.search(
                        r"\(jumlah pekerja (\d+) orang\)", task_raw, flags=re.IGNORECASE
                    )
                    deskripsi = re.sub(
                        r"\(jumlah pekerja.*?\)", "", task_raw, flags=re.IGNORECASE
                    )
                    deskripsi = re.sub(
                        r"schedule.*", "", deskripsi, flags=re.IGNORECASE
                    )
                    deskripsi = (
                        deskripsi.replace(",", "").replace(".", "").strip().capitalize()
                    )
                    if jumlah_pekerja_match:
                        deskripsi += (
                            f" (dikerjakan oleh {jumlah_pekerja_match.group(1)} orang)."
                        )
                    detail_pekerjaan_bersih.append(deskripsi)
        else:
            sub_tasks = re.split(r"\d+\.\(", chunk)
            for task in sub_tasks:
                if "(" not in task or "pekerjaan" not in task.lower():
                    continue
                deskripsi = re.sub(r".*?pekerjaan.*?=", "", task, flags=re.IGNORECASE)
                deskripsi = deskripsi.replace(")", "").strip().capitalize()
                nama_match = re.search(r"(.*?)\)", task)
                if nama_match:
                    deskripsi = f"({nama_match.group(1).strip()}): {deskripsi}"
                detail_pekerjaan_bersih.append(deskripsi)

        if detail_pekerjaan_bersih:
            parsed_data.append(
                {
                    "nama_grup": nama_grup,
                    "detail_pekerjaan": detail_pekerjaan_bersih,
                    "dokumentasi": [],
                }
            )

    print(f"   -> Analisis selesai. Ditemukan {len(parsed_data)} grup pekerjaan.")
    return parsed_data


def setup_folder_sementara(folder_nama="temp_images_extracted"):
    if os.path.exists(folder_nama):
        shutil.rmtree(folder_nama)
    os.makedirs(folder_nama)
    return folder_nama


def proses_zip_gambar(template, folder_utama, nama_grup):
    """
    Fungsi untuk meminta nama file ZIP, mengekstrak, dan memproses gambar.
    """
    daftar_objek_gambar = []
    while True:
        zip_filename = input(
            f">>> Masukkan nama file ZIP untuk '{nama_grup}' (contoh: excavator.zip): "
        )
        if not zip_filename.lower().endswith(".zip"):
            zip_filename += ".zip"

        if not os.path.exists(zip_filename):
            print(
                f"   ❌ ERROR: File '{zip_filename}' tidak ditemukan. Mohon periksa kembali nama file dan pastikan berada di folder yang sama dengan skrip."
            )
            continue

        print(f"   -> Memproses file '{zip_filename}'...")
        # Buat subfolder unik untuk ekstraksi agar tidak ada konflik nama file
        extract_path = os.path.join(folder_utama, nama_grup.replace(" ", "_"))
        os.makedirs(extract_path, exist_ok=True)

        with zipfile.ZipFile(zip_filename, "r") as zip_ref:
            zip_ref.extractall(extract_path)  # Ekstrak semua gambar ke subfolder

        # Proses semua gambar yang sudah diekstrak
        for image_name in os.listdir(extract_path):
            image_path = os.path.join(extract_path, image_name)
            try:
                with Image.open(image_path) as img:
                    lebar, tinggi = img.size
                objek_gambar = InlineImage(
                    template, image_path, width=Cm(14) if lebar > tinggi else Cm(9)
                )
                daftar_objek_gambar.append(objek_gambar)
            except Exception as e:
                print(
                    f"   ⚠️ Peringatan: Gagal memproses file '{image_name}'. Mungkin bukan file gambar. Error: {e}"
                )

        print(
            f"   ✅ Selesai memproses {len(daftar_objek_gambar)} gambar dari '{zip_filename}'."
        )
        break  # Keluar dari loop setelah berhasil

    return daftar_objek_gambar


if __name__ == "__main__":
    folder_temp = setup_folder_sementara()
    try:
        doc = DocxTemplate("template.docx")
    except Exception as e:
        print(
            f"FATAL ERROR: Gagal memuat 'template.docx'. Pastikan file ada di folder yang sama. Error: {e}"
        )
        exit()

    print("--- GENERATOR LAPORAN KERJA OTOMATIS (METODE ZIP) ---")
    print(
        "Silakan COPY semua teks laporan dari chat WhatsApp, PASTE di sini, lalu tekan Enter 2x untuk melanjutkan."
    )

    raw_wa_text = []
    while True:
        line = input()
        if not line:
            break
        raw_wa_text.append(line)
    raw_wa_text = "\n".join(raw_wa_text)

    data_pekerjaan = parse_whatsapp_text_v2(raw_wa_text)
    periode_match = re.search(
        r"periode.*?(\d{1,2}\s*-\s*\d{1,2}\s*\w+\s*\d{4})", raw_wa_text, re.IGNORECASE
    )
    periode_laporan = (
        periode_match.group(1) if periode_match else "Periode Tidak Ditemukan"
    )

    if data_pekerjaan:
        print("\n2. Memulai sesi penambahan dokumentasi dari file ZIP...")
        for grup in data_pekerjaan:
            grup["dokumentasi"] = proses_zip_gambar(doc, folder_temp, grup["nama_grup"])
    else:
        print("Tidak ada data pekerjaan yang bisa diproses. Program berhenti.")

    konteks = {
        "nama_perusahaan": "PT. ALU AKSARA PRATAMA",
        "periode_laporan": periode_laporan,
        "daftar_pekerjaan": data_pekerjaan,
    }

    doc.render(konteks)
    nama_file_output = f"Laporan Kerja ZIP {periode_laporan}.docx"
    doc.save(nama_file_output)

    print(f"\n🎉🎉🎉 SELESAI! Laporan '{nama_file_output}' telah berhasil dibuat.")

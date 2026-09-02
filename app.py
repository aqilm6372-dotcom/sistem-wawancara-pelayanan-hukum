"""
Sistem Wawancara Pelayanan Hukum
---------------------------------
Aplikasi Flask lokal untuk:
1. Menangkap cerita permasalahan hukum dari klien.
2. Merapikan cerita tersebut menjadi bahasa hukum formal menggunakan Gemini API.
3. Menghasilkan dokumen Word (.docx) yang langsung diunduh, TANPA disimpan
   secara fisik di server (dibuat & dikirim langsung dari memori).
"""

import os
import hmac
import mimetypes
from functools import wraps
from datetime import datetime
from io import BytesIO

from flask import (
    Flask,
    render_template,
    request,
    send_file,
    jsonify,
    session,
    redirect,
    url_for,
)

from google import genai
from google.genai import types

from docx import Document
from docx.shared import Pt, Inches
from docx.enum.text import WD_ALIGN_PARAGRAPH

app = Flask(__name__)

# Batas total ukuran file yang diupload per request (semua lampiran digabung).
app.config["MAX_CONTENT_LENGTH"] = 20 * 1024 * 1024  # 20 MB


@app.errorhandler(413)
def file_terlalu_besar(e):
    return jsonify(
        {"error": "Total ukuran file lampiran terlalu besar (maksimal 20MB)."}
    ), 413

# =========================================================================
# 0. KONFIGURASI LOGIN
# =========================================================================
# secret_key dipakai Flask untuk mengenkripsi session (wajib diisi agar
# login "diingat" oleh browser). Boleh diganti string acak apa saja.
app.secret_key = os.environ.get(
    "FLASK_SECRET_KEY", "ganti-string-ini-dengan-string-acak-rahasia-anda"
)

# Kredensial login. Bisa juga diisi lewat environment variable
# LOGIN_USERNAME / LOGIN_PASSWORD kalau tidak mau hardcode di sini.
LOGIN_USERNAME = os.environ.get("LOGIN_USERNAME", "kejaksaan")
LOGIN_PASSWORD = os.environ.get("LOGIN_PASSWORD", "kejarilangsa26")


def login_required(f):
    """Decorator: halaman hanya bisa diakses jika sudah login. Kalau belum,
    otomatis diarahkan ke halaman /login."""

    @wraps(f)
    def decorated_function(*args, **kwargs):
        if not session.get("logged_in"):
            return redirect(url_for("login"))
        return f(*args, **kwargs)

    return decorated_function


@app.route("/login", methods=["GET", "POST"])
def login():
    error = None
    if request.method == "POST":
        username = request.form.get("username", "")
        password = request.form.get("password", "")
        # hmac.compare_digest dipakai agar perbandingan username/password
        # tidak rentan terhadap timing attack.
        username_valid = hmac.compare_digest(username, LOGIN_USERNAME)
        password_valid = hmac.compare_digest(password, LOGIN_PASSWORD)
        if username_valid and password_valid:
            session["logged_in"] = True
            session["username"] = username
            return redirect(url_for("index"))
        error = "Username atau password salah."
    return render_template("login.html", error=error)


@app.route("/logout")
def logout():
    session.clear()
    return redirect(url_for("login"))


# =========================================================================
# 1. KONFIGURASI GEMINI API  ->  TEMPELKAN API KEY ANDA DI BARIS DI BAWAH INI
# =========================================================================
# Dapatkan API Key gratis di: https://aistudio.google.com/apikey
#
# Catatan: library resmi Google saat ini adalah "google-genai" (paket lama
# "google-generativeai" sudah deprecated per Mei 2025), jadi kode di bawah
# memakai SDK terbaru tersebut. Cara pakainya sangat mirip.
GEMINI_API_KEY = "MASUKKAN_API_KEY_ANDA_DISINI"
# Alternatif yang lebih aman (opsional): simpan API key sebagai environment
# variable, lalu baris di atas otomatis tidak dipakai jika env var tersedia.
#   Windows (PowerShell) : $env:GEMINI_API_KEY = "isi-api-key-anda"
#   Mac/Linux            : export GEMINI_API_KEY="isi-api-key-anda"
GEMINI_API_KEY = os.environ.get("GEMINI_API_KEY", GEMINI_API_KEY)

client = genai.Client(api_key=GEMINI_API_KEY)

# Model Gemini yang dipakai. Ganti ke "gemini-2.5-flash" atau model lain
# jika Anda ingin biaya/kecepatan berbeda.
GEMINI_MODEL = "gemini-3.6-flash"

# =========================================================================
# 2. SYSTEM PROMPT UNTUK AI
# =========================================================================
SYSTEM_PROMPT = """
Anda adalah asisten hukum profesional di Indonesia.
Tugas Anda: ubah cerita permasalahan berikut menjadi bahasa hukum formal
yang terstruktur, rapi, dan mudah dibaca.

Instruksi:
1. Perbaiki kosakata, ejaan, dan tata bahasa agar sesuai kaidah penulisan hukum formal.
2. Bagi tulisan menjadi beberapa paragraf yang runtut dan mudah dibaca
   (misalnya: paragraf pembuka/duduk perkara, kronologi kejadian, lalu
   permasalahan hukum yang timbul).
3. WAJIB pertahankan seluruh fakta asli yang disampaikan klien. Jangan
   menambah, mengurangi, mengarang, atau mengubah substansi fakta apa pun.
4. Hanya perbaiki CARA PENYAMPAIANNYA, bukan isi ceritanya.
5. Jika ada dokumen pendukung yang dilampirkan (gambar, PDF, atau file teks),
   baca dan manfaatkan informasi relevan di dalamnya untuk melengkapi
   kronologi (misalnya nomor sertifikat, tanggal, nama pihak). Tetap jangan
   mengarang fakta yang tidak didukung oleh cerita klien maupun lampiran.
6. Keluarkan HANYA teks hasil perbaikan dalam Bahasa Indonesia. Jangan
   tambahkan judul, salam pembuka, catatan, disclaimer, atau format
   markdown (tanpa tanda bintang, tanpa heading).
""".strip()


# =========================================================================
# ROUTES
# =========================================================================
@app.route("/")
@login_required
def index():
    return render_template("index.html")


@app.route("/generate", methods=["POST"])
@login_required
def generate():
    kategori = request.form.get("kategori", "").strip()
    cerita = request.form.get("cerita", "").strip()
    file_list = request.files.getlist("dokumen")

    if not kategori or not cerita:
        return jsonify({"error": "Kategori dan cerita permasalahan wajib diisi."}), 400

    # --- Proses semua file lampiran (gambar, PDF, Word, teks) ---
    try:
        teks_tambahan, media_parts, info_lampiran = proses_lampiran(file_list)
    except Exception as e:
        return jsonify({"error": f"Gagal memproses lampiran: {str(e)}"}), 500

    # --- Panggil Gemini API ---
    try:
        teks_hukum = panggil_gemini(cerita, teks_tambahan, media_parts)
    except Exception as e:
        return jsonify({"error": f"Gagal memproses AI: {str(e)}"}), 500

    # --- Buat dokumen Word di memori ---
    try:
        file_stream = buat_dokumen_word(kategori, teks_hukum, info_lampiran)
    except Exception as e:
        return jsonify({"error": f"Gagal membuat dokumen: {str(e)}"}), 500

    nama_kategori_file = kategori.replace(" ", "_")
    nama_file = f"Laporan_Hukum_{nama_kategori_file}.docx"

    return send_file(
        file_stream,
        as_attachment=True,
        download_name=nama_file,
        mimetype=(
            "application/vnd.openxmlformats-officedocument"
            ".wordprocessingml.document"
        ),
    )


# =========================================================================
# FUNGSI BANTUAN
# =========================================================================

# Ekstensi file yang didukung, dikelompokkan berdasarkan cara memprosesnya.
EKSTENSI_GAMBAR = {".jpg", ".jpeg", ".png", ".webp", ".gif", ".bmp"}
EKSTENSI_PDF = {".pdf"}
EKSTENSI_WORD = {".docx"}
EKSTENSI_TEKS = {".txt"}


def proses_lampiran(file_list):
    """Memproses semua file yang diupload klien.

    - Gambar & PDF  -> dikirim LANGSUNG ke Gemini sebagai media, dibaca
      secara native oleh AI (tidak perlu OCR manual).
    - Word (.docx) & teks (.txt) -> isinya diekstrak sebagai teks lalu
      digabungkan ke dalam prompt.
    - Format lain (termasuk .doc lama) -> dilewati, dicatat sebagai
      'tidak didukung' di laporan.

    Mengembalikan tuple:
      teks_tambahan  : gabungan teks hasil ekstraksi dari docx/txt
      media_parts    : list of google.genai.types.Part (gambar/PDF)
      info_lampiran  : list of dict untuk dicatat di dokumen Word hasil akhir
    """
    teks_tambahan_list = []
    media_parts = []
    info_lampiran = []

    for f in file_list:
        if not f or not f.filename:
            continue

        nama = f.filename
        ekstensi = os.path.splitext(nama)[1].lower()
        data = f.read()

        if not data:
            continue

        if ekstensi in EKSTENSI_GAMBAR:
            mime = mimetypes.guess_type(nama)[0] or "image/jpeg"
            media_parts.append(types.Part.from_bytes(data=data, mime_type=mime))
            info_lampiran.append({"nama": nama, "jenis": "Gambar", "bytes": data})

        elif ekstensi in EKSTENSI_PDF:
            media_parts.append(
                types.Part.from_bytes(data=data, mime_type="application/pdf")
            )
            info_lampiran.append({"nama": nama, "jenis": "PDF", "bytes": None})

        elif ekstensi in EKSTENSI_WORD:
            try:
                doc_dibaca = Document(BytesIO(data))
                teks_docx = "\n".join(
                    p.text for p in doc_dibaca.paragraphs if p.text.strip()
                )
                teks_tambahan_list.append(f"[Dari dokumen Word '{nama}']\n{teks_docx}")
                info_lampiran.append({"nama": nama, "jenis": "Word", "bytes": None})
            except Exception:
                info_lampiran.append(
                    {"nama": nama, "jenis": "Word (gagal dibaca)", "bytes": None}
                )

        elif ekstensi in EKSTENSI_TEKS:
            try:
                teks_txt = data.decode("utf-8", errors="ignore")
                teks_tambahan_list.append(f"[Dari file teks '{nama}']\n{teks_txt}")
                info_lampiran.append({"nama": nama, "jenis": "Teks", "bytes": None})
            except Exception:
                info_lampiran.append(
                    {"nama": nama, "jenis": "Teks (gagal dibaca)", "bytes": None}
                )

        elif ekstensi == ".doc":
            info_lampiran.append(
                {
                    "nama": nama,
                    "jenis": "Tidak didukung (Word 97-2003, simpan ulang sebagai .docx)",
                    "bytes": None,
                }
            )

        else:
            info_lampiran.append(
                {"nama": nama, "jenis": "Tidak didukung (format tidak dikenali)", "bytes": None}
            )

    teks_tambahan = "\n\n".join(teks_tambahan_list)
    return teks_tambahan, media_parts, info_lampiran


def panggil_gemini(cerita: str, teks_tambahan: str = "", media_parts=None) -> str:
    """Mengirim cerita permasalahan (+ isi lampiran, jika ada) ke Gemini API
    dan mengembalikan hasil yang sudah dirapikan menjadi bahasa hukum formal."""
    config = types.GenerateContentConfig(
        system_instruction=SYSTEM_PROMPT,
        temperature=0.4,
    )

    prompt_utama = cerita
    if teks_tambahan:
        prompt_utama += (
            "\n\n=== Isi Dokumen Pendukung yang Dilampirkan ===\n" + teks_tambahan
        )

    # contents berisi teks utama + (opsional) gambar/PDF yang dibaca native oleh Gemini
    contents = [prompt_utama]
    if media_parts:
        contents.extend(media_parts)

    response = client.models.generate_content(
        model=GEMINI_MODEL,
        contents=contents,
        config=config,
    )
    return response.text.strip()


def buat_dokumen_word(kategori: str, isi_teks: str, info_lampiran=None) -> BytesIO:
    """Membuat file .docx di dalam RAM (io.BytesIO) — tidak pernah menulis
    file fisik ke disk, sehingga storage server tidak menumpuk."""
    doc = Document()

    # Heading kategori permasalahan
    heading = doc.add_heading(f"Laporan Wawancara Hukum: {kategori}", level=1)
    heading.alignment = WD_ALIGN_PARAGRAPH.CENTER

    # Tanggal dokumen dicetak
    tanggal = datetime.now().strftime("%d %B %Y, %H:%M WIB")
    p_tanggal = doc.add_paragraph()
    p_tanggal.alignment = WD_ALIGN_PARAGRAPH.CENTER
    run_tanggal = p_tanggal.add_run(f"Dicetak pada: {tanggal}")
    run_tanggal.italic = True
    run_tanggal.font.size = Pt(10)

    doc.add_paragraph()  # spasi kosong

    # Sub-judul isi laporan
    doc.add_heading("Uraian Permasalahan", level=2)

    # Isi teks hasil AI dipecah menjadi beberapa paragraf
    paragraf_list = [p.strip() for p in isi_teks.split("\n") if p.strip()]
    for paragraf in paragraf_list:
        p = doc.add_paragraph(paragraf)
        p.alignment = WD_ALIGN_PARAGRAPH.JUSTIFY
        for r in p.runs:
            r.font.size = Pt(12)

    # --- Daftar lampiran dokumen (jika ada) ---
    if info_lampiran:
        doc.add_paragraph()
        doc.add_heading("Lampiran Dokumen", level=2)

        for item in info_lampiran:
            p = doc.add_paragraph()
            run = p.add_run(f"• {item['nama']}  —  {item['jenis']}")
            run.font.size = Pt(11)

            # Gambar langsung disisipkan sebagai pratinjau di bawah nama filenya
            if item["jenis"] == "Gambar" and item.get("bytes"):
                try:
                    doc.add_picture(BytesIO(item["bytes"]), width=Inches(4))
                except Exception:
                    doc.add_paragraph("   (Gagal menampilkan pratinjau gambar)")

    buffer = BytesIO()
    doc.save(buffer)
    buffer.seek(0)
    return buffer


if __name__ == "__main__":
    app.run(debug=True, host="127.0.0.1", port=5000, use_reloader=False)

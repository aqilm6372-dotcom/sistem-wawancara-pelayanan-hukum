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
from docx.shared import Pt
from docx.enum.text import WD_ALIGN_PARAGRAPH

app = Flask(__name__)

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

import os
from dotenv import load_dotenv

load_dotenv()  # Mengambil variabel dari file .env

GEMINI_API_KEY = os.environ.get("GEMINI_API_KEY")
client = genai.Client(api_key = GEMINI_API_KEY)

# =========================================================================
# 1. KONFIGURASI GEMINI API  ->  TEMPELKAN API KEY ANDA DI BARIS DI BAWAH INI
# =========================================================================
# Dapatkan API Key gratis di: https://aistudio.google.com/apikey
#
# Catatan: library resmi Google saat ini adalah "google-genai" (paket lama
# "google-generativeai" sudah deprecated per Mei 2025), jadi kode di bawah
# memakai SDK terbaru tersebut. Cara pakainya sangat mirip.


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
5. Keluarkan HANYA teks hasil perbaikan dalam Bahasa Indonesia. Jangan
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

    if not kategori or not cerita:
        return jsonify({"error": "Kategori dan cerita permasalahan wajib diisi."}), 400

    # --- Panggil Gemini API ---
    try:
        teks_hukum = panggil_gemini(cerita)
    except Exception as e:
        return jsonify({"error": f"Gagal memproses AI: {str(e)}"}), 500

    # --- Buat dokumen Word di memori ---
    try:
        file_stream = buat_dokumen_word(kategori, teks_hukum)
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
def panggil_gemini(cerita: str) -> str:
    """Mengirim cerita permasalahan ke Gemini API dan mengembalikan hasil
    yang sudah dirapikan menjadi bahasa hukum formal."""
    config = types.GenerateContentConfig(
        system_instruction=SYSTEM_PROMPT,
        temperature=0.4,
    )
    response = client.models.generate_content(
        model=GEMINI_MODEL,
        contents=cerita,
        config=config,
    )
    return response.text.strip()


def buat_dokumen_word(kategori: str, isi_teks: str) -> BytesIO:
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

    buffer = BytesIO()
    doc.save(buffer)
    buffer.seek(0)
    return buffer


if __name__ == "__main__":
    app.run(debug=True, host="127.0.0.1", port=5000, use_reloader=False)
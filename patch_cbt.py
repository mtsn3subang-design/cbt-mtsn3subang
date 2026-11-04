#!/usr/bin/env python3
"""
Patch kecepatan untuk CBT MTsN 3 Subang.
Pemakaian:  python3 patch_cbt.py index.html
Hasil:      index_cepat.html (file asli tidak diubah)
Opsional:   pip install pillow  -> logo ikut diperkecil otomatis
"""
import sys, re, base64, io

src = sys.argv[1] if len(sys.argv) > 1 else "index.html"
dst = src.rsplit(".", 1)[0] + "_cepat.html"
h = open(src, encoding="utf-8").read()
size0 = len(h.encode("utf-8"))
log = []


def rep(old, new, name):
    global h
    n = h.count(old)
    if n != 1:
        log.append(f"[LEWAT] {name}: ditemukan {n}x (harus 1x), tidak diubah")
        return
    h = h.replace(old, new)
    log.append(f"[OK]    {name}")


# 1. Hapus request 'info' saat halaman login dibuka
rep("setTimeout(()=>{api({action:'info'}).then(r=>{if(r.ok){PUB=r.cfg;if(mode==='login'&&!(($('#lu')||{}).value)){S.cfg={...PUB};render()}}})},Math.random()*3000);",
    "", "1. Hapus request info saat login")

# 2. Library PDF/Excel dimuat hanya untuk guru/admin
tags = re.findall(r'<script defer src="https://cdnjs\.cloudflare\.com/ajax/libs/(?:jspdf|jspdf-autotable|xlsx)/[^"]+"></script>\s*', h)
if len(tags) == 3:
    for t in tags:
        h = h.replace(t, "")
    log.append("[OK]    2a. Hapus 3 tag script jspdf/autotable/xlsx")
else:
    log.append(f"[LEWAT] 2a. Tag script library: ditemukan {len(tags)}, seharusnya 3")

loader = ("const LIBS=['https://cdnjs.cloudflare.com/ajax/libs/jspdf/2.5.1/jspdf.umd.min.js',"
          "'https://cdnjs.cloudflare.com/ajax/libs/jspdf-autotable/3.5.31/jspdf.plugin.autotable.min.js',"
          "'https://cdnjs.cloudflare.com/ajax/libs/xlsx/0.18.5/xlsx.full.min.js'];\n"
          "let libP;const loadLibs=()=>libP||(libP=LIBS.reduce((p,u)=>p.then(()=>new Promise(r=>{const s=document.createElement('script');s.src=u;s.onload=s.onerror=r;document.head.appendChild(s)})),Promise.resolve()));\n")
rep("const wait=ms=>new Promise(z=>setTimeout(z,ms));",
    loader + "const wait=ms=>new Promise(z=>setTimeout(z,ms));", "2b. Definisi loadLibs")
rep("adm=true;mode='admin';page='dash';render();if(S.warn)alert(S.warn)}",
    "adm=true;mode='admin';page='dash';render();loadLibs();if(S.warn)alert(S.warn)}", "2c. Panggil loadLibs setelah login guru/admin")

# 3. Logo: satu salinan saja, diperkecil
m = re.search(r"const LOGO='(data:image/png;base64,[A-Za-z0-9+/=]+)';", h)
if m:
    new_logo = m.group(1)
    try:
        from PIL import Image
        im = Image.open(io.BytesIO(base64.b64decode(new_logo.split(",", 1)[1]))).convert("RGBA")
        im.thumbnail((96, 96))
        bg = Image.new("RGBA", im.size, (255, 255, 255, 255))
        bg.alpha_composite(im)
        pal = bg.convert("RGB").quantize(colors=64)
        buf = io.BytesIO()
        pal.save(buf, "PNG", optimize=True)
        new_logo = "data:image/png;base64," + base64.b64encode(buf.getvalue()).decode()
        log.append(f"[OK]    3a. Logo diperkecil: {len(m.group(1))//1024} KB -> {len(new_logo)//1024} KB")
    except ImportError:
        log.append("[INFO]  3a. Pillow tidak ada, logo tidak diperkecil (pip install pillow)")
    h = h.replace(m.group(0), "const LOGO='" + new_logo + "';\ndocument.querySelector('link[rel=icon]').href=LOGO;")
    log.append("[OK]    3b. Favicon memakai LOGO yang sama (tidak ditulis dua kali)")
fav = re.search(r'<link rel="icon" href="data:image/png;base64,[A-Za-z0-9+/=]+">', h)
if fav:
    h = h.replace(fav.group(0), '<link rel="icon" href="">')
else:
    log.append("[LEWAT] 3c. Tag favicon tidak ditemukan")

# 4. Simpan jawaban ke server lebih jarang
rep("SYNC_MIN=40000,SYNC_RND=30000;", "SYNC_MIN=120000,SYNC_RND=60000;", "4. Sinkron jawaban tiap 2-3 menit")

# 5. Polling admin 90 detik -> 5 menit
rep("))loadAdmin(true)},90000);", "))loadAdmin(true)},300000);", "5. Polling admin tiap 5 menit")

# 6. Import soal: gambar dari path GitHub (img/nama.webp)
rep("const s={id:uid(),topik:t,q:c[0],o,k,img:'',oi:o.map(()=>'')};",
    "const s={id:uid(),topik:t,q:c[0],o,k,img:/^(https?:\\/\\/|img\\/)/i.test(img)?img:'',oi:o.map(()=>'')};",
    "6. Import soal menerima path gambar img/...")

open(dst, "w", encoding="utf-8").write(h)
print("\n".join(log))
print(f"\nUkuran: {size0//1024} KB -> {len(h.encode('utf-8'))//1024} KB")
print("Tersimpan sebagai:", dst)

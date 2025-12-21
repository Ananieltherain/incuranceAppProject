const express = require("express");
const path = require("path");
const fs = require("fs");
const multer = require("multer");
const PizZip = require("pizzip");
const Docxtemplater = require("docxtemplater");
const ImageModule = require("docxtemplater-image-module-free");

const app = express();
const PORT = 3000;

// middleware
app.use(express.static("public"));
app.use(express.urlencoded({ extended: true }));

// upload
const upload = multer({ dest: "uploads/" });

// terbilang sederhana
function terbilang(n) {
    const satuan = ["", "satu", "dua", "tiga", "empat", "lima", "enam", "tujuh", "delapan", "sembilan", "sepuluh", "sebelas"];
    function toWords(x) {
        if (x < 12) return satuan[x];
        if (x < 20) return toWords(x - 10) + " belas";
        if (x < 100) return toWords(Math.floor(x / 10)) + " puluh " + toWords(x % 10);
        if (x < 200) return "seratus " + toWords(x - 100);
        if (x < 1000) return toWords(Math.floor(x / 100)) + " ratus " + toWords(x % 100);
        if (x < 2000) return "seribu " + toWords(x - 1000);
        if (x < 1000000) return toWords(Math.floor(x / 1000)) + " ribu " + toWords(x % 1000);
        return x.toString();
    }
    return toWords(n).replace(/\s+/g, " ").trim() + " rupiah";
}

// mapping asuransi
const ASURANSI_INFO = {
    "TOB": {
template: `Kepada Yth. :
Bapak Rizqi Abdul Ghani
PT ASURANSI TOTAL BERSAMA 
Citra Tower, 27th Floor
Jl. Benyamin Suaeb Blok A6 RT 13/ RW 06
Kebon Kosong Kec. Kemayoran – Jakarta Pusat 10630`,
perusahaan: "PT. Asuransi Total Bersama",
direktur: "Direkur TOB"
    },
    "SOMPO": {
template: `Kepada Yth. :
Bapak Arief Hariyanto
Head Of Motor Claim Departement
PT. Sompo Insurance Indonesia
Mayapada Tower II, 19th floor
Jl. Jend. Sudirman Kav. 27 Jakarta 12920`,
perusahaan: "PT. Sompo Insurance Indonesia",
direktur: "Direktur Sompo"
    },
    "ACA": {
template: `Kepada Yth. :
Bapak Agus Suryono
PT ASURANSI CENTRAL ASIA
Gedung Hermina Tower 1 (Lantai 3)
Jl. HBR Motik Blok B-10 Kemayoran, Jakarta 10610`,
perusahaan: "PT. Asuransi Central Asia",
direktur: "Direktur ACA"
    }
};

// ROUTE GENERATE DOCX
app.post(
    "/generate",
    upload.fields([{ name: "lampiranFoto", maxCount: 20 }]),
    (req, res) => {
        try {
            const data = req.body;
            const files = req.files?.lampiranFoto || [];

            // 🔑 BUILD LAMPIRAN ARRAY
            const rawKet = data.lampiranKeterangan || [];
            const keterangan = Array.isArray(rawKet) ? rawKet : [rawKet];

            const lampiran = files.map((file, i) => ({
                image: file.path,
                keterangan: keterangan[i] || ""
            }));

            // 🔢 akomodasi
            let list = [];
            let total = 0;

            if (Array.isArray(data.kebutuhan)) {
                data.kebutuhan.forEach((item, i) => {
                    const nom = parseInt(data.nominal[i], 10) || 0;
                    total += nom;
                    list.push({
                        no: i + 1,
                        nama: item,
                        nominal: nom.toLocaleString("id-ID")
                    });
                });
            }

            // 📄 baca template
            const content = fs.readFileSync(
                path.join(__dirname, "template", "template.docx"),
                "binary"
            );

            // ✅ IMAGE MODULE — SATU KALI, SEBELUM DOCXTEMPLATER
            const imageModule = new ImageModule({
                centered: false,
                getImage: (tagValue) => fs.readFileSync(tagValue),
                getSize: () => [400, 250],
            });

            const zip = new PizZip(content);
            const doc = new Docxtemplater(zip, {
                paragraphLoop: true,
                linebreaks: true,
                modules: [imageModule],
            });

            const today = new Date();
            const tanggalTerbit = `${today.getDate()} ${today.toLocaleString("id-ID", { month: "long" })} ${today.getFullYear()}`;

  // ini penting
        const selected = ASURANSI_INFO[data.jenisAsuransi] || { template: "", perusahaan: "", direktur: "" };

        let surveyfeeText = "";
        let nominalSurvey = "";

        if(data.jenisAsuransi === "SOMPO"){
        if(data.surveyFeeType && data.surveyFeeNominal){
        surveyfeeText = `Survey fee`;
        nominalSurvey = Number(data.surveyFeeNominal).toLocaleString("id-ID");
    }
}

            doc.render({
            jenisBiaya: data.jenisBiaya,
            suratKuasa: data.suratKuasa,
            polis: data.noPolis,
            tertanggung: data.tertanggung,
            unit: data.unit,
            tkp: data.tkp,
            alamatTertanggung: data.alamatTertanggung,
            alamatSupir: data.alamatSupir,
            alamatBengkel: data.alamatBengkel,
            kantorPolisi: data.kantorPolisi,
            dol: data.dol,
            akomodasi: list,
            totalAkomodasi: total.toLocaleString("id-ID"),
            terbilang: terbilang(total),
            terbit: tanggalTerbit,
            namaAsuransi: selected.template,
            Perusahaan: selected.perusahaan,
            direktur: selected.direktur,
            surveyfee: surveyfeeText,         
            nominalsurveyfee: nominalSurvey,
            lampiran: lampiran
            });

            const buffer = doc.getZip().generate({ type: "nodebuffer" });
            const filename = `pengajuan_${Date.now()}.docx`;
            const filepath = path.join(__dirname, "downloads", filename);

            fs.writeFileSync(filepath, buffer);
            res.download(filepath);

        } catch (err) {
            console.error(err);
            res.status(500).send("Gagal generate DOCX");
        }
    }
);

app.listen(PORT, () =>
    console.log(`SERVER → http://localhost:${PORT}`)
);

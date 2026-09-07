import { useState, useEffect, useCallback } from "react";
import {
  Document, Packer, Paragraph, TextRun, Table, TableRow, TableCell,
  ImageRun, PageBreak,
  AlignmentType, WidthType, BorderStyle, VerticalAlign,
} from "docx";

// Load Google Fonts: Plus Jakarta Sans + IBM Plex Mono
const fontLink = document.createElement("link");
fontLink.rel = "stylesheet";
fontLink.href = "https://fonts.googleapis.com/css2?family=Plus+Jakarta+Sans:wght@400;500;600;700&family=IBM+Plex+Mono:wght@400;500&display=swap";
document.head.appendChild(fontLink);

// ─── DATA ────────────────────────────────────────────────────────────────────

const TARUNA = {
  nama: "MUHAMMAD RADHITYA BHARA KUNTARA",
  noakad: "200507014166",
  tonki: "2/B",
  tk: "III",
  dantontar: "IMAM SANTOSO, S.H.",
  pangkat_dantontar: "IPDA",
  nrp_dantontar: "82020196",
  dankitar: "AGUNG BUDIMAN, S.Tr.K., S.I.K.",
  pangkat_dankitar: "AKP",
  nrp_dankitar: "92040556",
};

const NSP_ITEMS = [
  // ── B1: KEIMANAN & KETAQWAAN ──────────────────────────────────────────────
  {
    id: "b1b2_4x", cat: "B1 – Keimanan & Ketaqwaan",
    nama: "Imam/mentor/pemimpin ibadah dalam kegiatan agama (≥4×/bulan)",
    poin: 0.33, pemberi: "PAWASYON", lhp: true,
    keterangan: "LHP min 10 taruna (dalam/luar Akpol). Dokumentasi peserta wajib.",
  },
  {
    id: "b1b2_3x", cat: "B1 – Keimanan & Ketaqwaan",
    nama: "Imam/mentor/pemimpin ibadah dalam kegiatan agama (3×/bulan)",
    poin: 0.25, pemberi: "PAWASYON", lhp: true,
    keterangan: "LHP min 10 taruna.",
  },
  {
    id: "b1b2_2x", cat: "B1 – Keimanan & Ketaqwaan",
    nama: "Imam/mentor/pemimpin ibadah dalam kegiatan agama (2×/bulan)",
    poin: 0.17, pemberi: "PAWASYON", lhp: true,
    keterangan: "LHP min 10 taruna.",
  },
  {
    id: "b1b2_1x", cat: "B1 – Keimanan & Ketaqwaan",
    nama: "Imam/mentor/pemimpin ibadah dalam kegiatan agama (1×/bulan)",
    poin: 0.08, pemberi: "PAWASYON", lhp: true,
    keterangan: "LHP min 10 taruna.",
  },
  {
    id: "b1b3_2x", cat: "B1 – Keimanan & Ketaqwaan",
    nama: "Mengajak Taruna junior ibadah di luar Akpol (≥2×/bulan)",
    poin: 0.33, pemberi: "DANKITAR", lhp: true,
    keterangan: "LHP min 10 taruna. Kegiatan saat pesiar/IBL/cuti.",
  },
  {
    id: "b1b3_1x", cat: "B1 – Keimanan & Ketaqwaan",
    nama: "Mengajak Taruna junior ibadah di luar Akpol (1×/bulan)",
    poin: 0.25, pemberi: "DANKITAR", lhp: true,
    keterangan: "LHP min 10 taruna. Kegiatan saat pesiar/IBL/cuti.",
  },
  {
    id: "b1c3", cat: "B1 – Keimanan & Ketaqwaan",
    nama: "Korve bersama di tempat ibadah agama berbeda (≥1×/bulan)",
    poin: 0.33, pemberi: "DANTONTAR", lhp: true,
    keterangan: "LHP wajib disertakan.",
  },
  {
    id: "b1d1", cat: "B1 – Keimanan & Ketaqwaan",
    nama: "Membuat ide implementasi nilai imtaq dalam tulisan (diketik)",
    poin: 1.0, pemberi: "DANKITAR", lhp: true,
    keterangan: "LHP diketik min 3 halaman. 1× atau lebih dalam sebulan.",
  },
  // ── B2: CINTA TANAH AIR ───────────────────────────────────────────────────
  {
    id: "b2a2", cat: "B2 – Cinta Tanah Air",
    nama: "Petugas/deputasi upacara sesuai sprin dinas (hafal ≥4 item)",
    poin: 0.2, pemberi: "DANTONTAR", lhp: false,
    keterangan: "Poin diberikan jika hafal 4 item atau lebih.",
  },
  {
    id: "b2b1", cat: "B2 – Cinta Tanah Air",
    nama: "Tulisan/artikel bertema nasionalisme (1×/bulan)",
    poin: 0.5, pemberi: "DANKITAR", lhp: true,
    keterangan: "LHP tulisan tangan min 2 halaman folio. Gunakan bahasa Indonesia yang baik dan benar.",
  },
  {
    id: "b2b2", cat: "B2 – Cinta Tanah Air",
    nama: "Menampilkan kesenian/budaya dari daerah lain (1×/bulan)",
    poin: 0.5, pemberi: "DANTONTAR", lhp: true,
    keterangan: "LHP wajib. Sertakan foto penampilan.",
  },
  {
    id: "b2c1", cat: "B2 – Cinta Tanah Air",
    nama: "Membuat ide implementasi nilai cinta tanah air (tulisan)",
    poin: 1.0, pemberi: "DANKITAR", lhp: true,
    keterangan: "LHP. Mampu mendesain kegiatan peningkatan karakter untuk taruna dan masyarakat.",
  },
  // ── B3: DEMOKRASI ─────────────────────────────────────────────────────────
  {
    id: "b3a2", cat: "B3 – Demokrasi",
    nama: "Aktif menyampaikan pendapat di forum resmi/arahan pimpinan (≥1×)",
    poin: 0.17, pemberi: "PAWASYON", lhp: false,
    keterangan: "Forum lingkup detasemen dan kompi. 1× atau lebih dalam sebulan.",
  },
  {
    id: "b3b1", cat: "B3 – Demokrasi",
    nama: "Membuat ide implementasi nilai demokrasi (diketik)",
    poin: 0.25, pemberi: "DANKITAR", lhp: true,
    keterangan: "LHP diketik min 3 halaman. 1× atau lebih dalam sebulan.",
  },
  // ── B4: DISIPLIN ──────────────────────────────────────────────────────────
  {
    id: "b4a1", cat: "B4 – Disiplin",
    nama: "Tidak melanggar peraturan selama sebulan penuh",
    poin: 0.33, pemberi: "PAWASYON", lhp: false,
    keterangan: "Menepati semua peraturan & kegiatan selama sebulan.",
  },
  {
    id: "b4a2_tlb", cat: "B4 – Disiplin",
    nama: "Terlambat hanya 1× dalam sebulan",
    poin: 0.17, pemberi: "PAWASYON", lhp: false,
    keterangan: "Poin pengganti jika terlambat 1 kali.",
  },
  {
    id: "b4a3_4x", cat: "B4 – Disiplin",
    nama: "Mendapat pujian PUD dari pengasuh (≥4×/bulan)",
    poin: 0.33, pemberi: "PAWASYON", lhp: false,
    keterangan: "Pujian 4 kali atau lebih dalam sebulan.",
  },
  {
    id: "b4a3_3x", cat: "B4 – Disiplin",
    nama: "Mendapat pujian PUD dari pengasuh (3×/bulan)",
    poin: 0.25, pemberi: "PAWASYON", lhp: false,
    keterangan: "Pujian 3 kali dalam sebulan.",
  },
  {
    id: "b4a3_2x", cat: "B4 – Disiplin",
    nama: "Mendapat pujian PUD dari pengasuh (2×/bulan)",
    poin: 0.17, pemberi: "PAWASYON", lhp: false,
    keterangan: "Pujian 2 kali dalam sebulan.",
  },
  {
    id: "b4a3_1x", cat: "B4 – Disiplin",
    nama: "Mendapat pujian PUD dari pengasuh (1×/bulan)",
    poin: 0.08, pemberi: "PAWASYON", lhp: false,
    keterangan: "Pujian 1 kali dalam sebulan.",
  },
  {
    id: "b4b1", cat: "B4 – Disiplin",
    nama: "Membuat ide implementasi nilai disiplin (diketik)",
    poin: 0.5, pemberi: "DANKITAR", lhp: true,
    keterangan: "LHP diketik min 3 halaman. 1× atau lebih dalam sebulan.",
  },
  // ── B5: KERJA KERAS & CERDAS ──────────────────────────────────────────────
  {
    id: "b5a1_2x", cat: "B5 – Kerja Keras & Cerdas",
    nama: "Belajar mandiri di luar jadwal + resume/ringkasan materi (≥2×)",
    poin: 0.33, pemberi: "PAWASYON", lhp: true,
    keterangan: "LHP wajib. Resume/ringkasan materi segera setelah kegiatan.",
  },
  {
    id: "b5a1_1x", cat: "B5 – Kerja Keras & Cerdas",
    nama: "Belajar mandiri di luar jadwal + resume/ringkasan materi (1×)",
    poin: 0.17, pemberi: "PAWASYON", lhp: true,
    keterangan: "LHP wajib.",
  },
  {
    id: "b5a2_2x", cat: "B5 – Kerja Keras & Cerdas",
    nama: "Belajar mandiri di perpustakaan/beli buku + resume (≥2×)",
    poin: 0.33, pemberi: "PAWASYON", lhp: true,
    keterangan: "LHP wajib. Dokumentasi buku, resume min 2 halaman folio.",
  },
  {
    id: "b5a2_1x", cat: "B5 – Kerja Keras & Cerdas",
    nama: "Belajar mandiri di perpustakaan/beli buku + resume (1×)",
    poin: 0.17, pemberi: "PAWASYON", lhp: true,
    keterangan: "LHP wajib.",
  },
  {
    id: "b5a4", cat: "B5 – Kerja Keras & Cerdas",
    nama: "Membantu giat operasi kepolisian (gatur lalin, TPTKP, sosialisasi harkamtibmas, dll)",
    poin: 0.33, pemberi: "DANTONTAR", lhp: true,
    keterangan: "LHP. Dokumentasi didampingi personil Polri. Uraian lengkap: lokasi, waktu, nama personil, bentuk kegiatan.",
  },
  {
    id: "b5a5", cat: "B5 – Kerja Keras & Cerdas",
    nama: "Olahraga mandiri di luar jadwal rutin (≥2×) ATAU turun BB min 2 kg",
    poin: 0.33, pemberi: "DANTONTAR", lhp: true,
    keterangan: "LHP. Laporkan kepada pawas sebelumnya, atau sertakan bukti penurunan berat badan.",
  },
  {
    id: "b5a6", cat: "B5 – Kerja Keras & Cerdas",
    nama: "Pemetaan kekuatan & kelemahan diri / portofolio diri (diketik)",
    poin: 0.33, pemberi: "DANKITAR", lhp: true,
    keterangan: "LHP diketik. 1× atau lebih dalam sebulan.",
  },
  // ── B6: PROFESIONAL ──────────────────────────────────────────────────────
  {
    id: "b6a1", cat: "B6 – Profesional",
    nama: "Buku panduan/diktat/SOP mandiri atau ide terobosan kreatif (tulis tangan)",
    poin: 0.33, pemberi: "DANKITAR", lhp: true,
    keterangan: "LHP tulis tangan min 2 halaman folio. 1× atau lebih dalam sebulan.",
  },
  {
    id: "b6a2_2x", cat: "B6 – Profesional",
    nama: "Mengunjungi kesatuan kepolisian di luar kampus + laporan teknis (≥2×)",
    poin: 0.33, pemberi: "DANKITAR", lhp: true,
    keterangan: "LHP. Berisi dokumentasi, waktu, tempat, pejabat yang ditemui, laporan teknis sesuai materi kuliah.",
  },
  {
    id: "b6a2_1x", cat: "B6 – Profesional",
    nama: "Mengunjungi kesatuan kepolisian di luar kampus + laporan teknis (1×)",
    poin: 0.17, pemberi: "DANKITAR", lhp: true,
    keterangan: "LHP. Berisi dokumentasi, waktu, tempat, pejabat yang ditemui.",
  },
  {
    id: "b6b1", cat: "B6 – Profesional",
    nama: "Prestasi nilai mental terbaik detasemen (10 besar/semester)",
    poin: 0.33, pemberi: "PAWASYON", lhp: false,
    keterangan: "10 terbaik tiap semester.",
  },
  {
    id: "b6b2", cat: "B6 – Profesional",
    nama: "Prestasi nilai jasmani terbaik detasemen (10 besar/semester)",
    poin: 0.33, pemberi: "PAWASYON", lhp: false,
    keterangan: "10 terbaik tiap semester.",
  },
  {
    id: "b6b3", cat: "B6 – Profesional",
    nama: "Tidak her akademik dalam satu semester",
    poin: 0.33, pemberi: "PAWASYON", lhp: false,
    keterangan: "Tidak her akademik selama satu semester.",
  },
  // ── B7: SEDERHANA ────────────────────────────────────────────────────────
  {
    id: "b7c1", cat: "B7 – Sederhana",
    nama: "Membuat ide implementasi nilai sederhana (diketik)",
    poin: 0.17, pemberi: "DANKITAR", lhp: true,
    keterangan: "LHP diketik min 3 halaman. 1× atau lebih dalam sebulan.",
  },
  // ── B8: EMPATI ───────────────────────────────────────────────────────────
  {
    id: "b8a1", cat: "B8 – Empati",
    nama: "Menjenguk Taruna/antap Akpol sakit di KSA/RS rujukan (≥1×)",
    poin: 0.33, pemberi: "PAWASYON", lhp: true,
    keterangan: "LHP. Dokumentasi bersama yang sakit. Uraian: waktu, tempat, nama Taruna/antap yang sakit.",
  },
  {
    id: "b8a2", cat: "B8 – Empati",
    nama: "Baksos inisiatif sendiri (biaya pribadi) / donor darah (≥1×)",
    poin: 0.33, pemberi: "DANTONTAR", lhp: true,
    keterangan: "LHP. Dokumentasi saat penyerahan. Paket baksos bertuliskan BAKSOS (NAMA BATALYON).",
  },
  {
    id: "b8a3", cat: "B8 – Empati",
    nama: "Berkunjung ke rumah pejabat Akpol (≥1×)",
    poin: 0.33, pemberi: "DANTONTAR", lhp: true,
    keterangan: "LHP. Dokumentasi terlampir. Di luar jam pengasuhan.",
  },
  {
    id: "b8b1", cat: "B8 – Empati",
    nama: "Membuat ide implementasi nilai empati (diketik)",
    poin: 0.5, pemberi: "DANKITAR", lhp: true,
    keterangan: "LHP diketik min 3 halaman. 1× atau lebih dalam sebulan.",
  },
  // ── B9: JUJUR & IKHLAS ───────────────────────────────────────────────────
  {
    id: "b9b1", cat: "B9 – Jujur & Ikhlas",
    nama: "Membuat ide implementasi nilai jujur & ikhlas (diketik)",
    poin: 1.0, pemberi: "DANKITAR", lhp: true,
    keterangan: "LHP diketik min 3 halaman. 1× atau lebih dalam sebulan.",
  },
  // ── B10: ADIL ────────────────────────────────────────────────────────────
  {
    id: "b10b1", cat: "B10 – Adil",
    nama: "Membuat ide implementasi nilai adil (diketik)",
    poin: 0.5, pemberi: "DANKITAR", lhp: true,
    keterangan: "LHP diketik min 3 halaman. 1× atau lebih dalam sebulan.",
  },
  // ── B11: TELADAN ─────────────────────────────────────────────────────────
  {
    id: "b11a1_1", cat: "B11 – Teladan",
    nama: "NSP terbaik 1 per satuan bulan ini",
    poin: 0.33, pemberi: "PAWASYON", lhp: false,
    keterangan: "Mendapatkan nilai NSP 3 (tiga) terbaik setiap bulan per satuan.",
  },
  {
    id: "b11a1_2", cat: "B11 – Teladan",
    nama: "NSP terbaik 2 per satuan bulan ini",
    poin: 0.25, pemberi: "PAWASYON", lhp: false,
    keterangan: "",
  },
  {
    id: "b11a1_3", cat: "B11 – Teladan",
    nama: "NSP terbaik 3 per satuan bulan ini",
    poin: 0.17, pemberi: "PAWASYON", lhp: false,
    keterangan: "",
  },
  {
    id: "b11a2_4x", cat: "B11 – Teladan",
    nama: "Penilaian positif atas kinerja sebagai pejabat korp (≥4×)",
    poin: 0.33, pemberi: "PAWASYON", lhp: false,
    keterangan: "4 kali atau lebih dalam sebulan.",
  },
  {
    id: "b11a3_2x", cat: "B11 – Teladan",
    nama: "Pujian penampilan baik/rapih saat apel TI/pemeriksaan (≥2×)",
    poin: 0.33, pemberi: "PAWASYON", lhp: false,
    keterangan: "Pujian 2 kali atau lebih dalam sebulan.",
  },
  {
    id: "b11b1_15", cat: "B11 – Teladan",
    nama: "Sosiometri peringkat 1–5 terbaik setingkat detasemen",
    poin: 0.33, pemberi: "PAWASYON", lhp: false,
    keterangan: "Masuk peringkat 20 besar sosiometri setingkat detasemen.",
  },
  {
    id: "b11b1_610", cat: "B11 – Teladan",
    nama: "Sosiometri peringkat 6–10 terbaik setingkat detasemen",
    poin: 0.25, pemberi: "PAWASYON", lhp: false,
    keterangan: "",
  },
  {
    id: "b11b2_1", cat: "B11 – Teladan",
    nama: "Sosiometri peringkat 1 terbaik setingkat satuan",
    poin: 0.33, pemberi: "PAWASYON", lhp: false,
    keterangan: "Masuk peringkat 3 besar sosiometri setingkat satuan.",
  },
  // ── B12: INTEGRITAS ──────────────────────────────────────────────────────
  {
    id: "b12a1", cat: "B12 – Integritas",
    nama: "Konsisten tidak melanggar peraturan selama 2 bulan berturut-turut",
    poin: 0.66, pemberi: "PAWASYON", lhp: false,
    keterangan: "Selalu konsisten dan tidak melanggar semua peraturan selama 2 bulan berturut-turut.",
  },
  {
    id: "b12a2_naik", cat: "B12 – Integritas",
    nama: "Nilai NSP bulan berikutnya meningkat",
    poin: 0.66, pemberi: "PAWASYON", lhp: false,
    keterangan: "Nilai NSP bulan berikutnya mengalami kenaikan.",
  },
  {
    id: "b12a2_tetap", cat: "B12 – Integritas",
    nama: "Nilai NSP tetap (sama dengan bulan sebelumnya)",
    poin: 0.38, pemberi: "PAWASYON", lhp: false,
    keterangan: "Nilai NSP sama dengan bulan sebelumnya.",
  },
  {
    id: "b12a3_latihan", cat: "B12 – Integritas",
    nama: "Melaksanakan deputasi (memerlukan latihan) ≥1×/bulan",
    poin: 0.66, pemberi: "DANTONTAR", lhp: false,
    keterangan: "1 kali atau lebih dalam sebulan — deputasi yang memerlukan latihan.",
  },
  {
    id: "b12a3_biasa", cat: "B12 – Integritas",
    nama: "Melaksanakan deputasi (tidak memerlukan latihan) ≥1×/bulan",
    poin: 0.33, pemberi: "DANTONTAR", lhp: false,
    keterangan: "1 kali atau lebih dalam sebulan — deputasi tanpa latihan.",
  },
  {
    id: "b12b1", cat: "B12 – Integritas",
    nama: "Membuat ide implementasi nilai integritas (diketik)",
    poin: 1.0, pemberi: "DANKITAR", lhp: true,
    keterangan: "LHP diketik min 3 halaman. 1× dalam sebulan.",
  },
];

const MONTHS = [
  { name: "Januari", year: 2026, deadline: new Date(2026, 0, 25), key: "jan26" },
  { name: "Februari", year: 2026, deadline: new Date(2026, 1, 22), key: "feb26" },
  { name: "Maret", year: 2026, deadline: new Date(2026, 2, 25), key: "mar26" },
  { name: "April", year: 2026, deadline: new Date(2026, 3, 25), key: "apr26" },
  { name: "Mei", year: 2026, deadline: new Date(2026, 4, 25), key: "may26" },
  { name: "Juni", year: 2026, deadline: new Date(2026, 5, 25), key: "jun26" },
  { name: "Juli", year: 2026, deadline: new Date(2026, 6, 25), key: "jul26" },
  { name: "Agustus", year: 2026, deadline: new Date(2026, 7, 25), key: "aug26" },
  { name: "September", year: 2026, deadline: new Date(2026, 8, 25), key: "sep26" },
  { name: "Oktober", year: 2026, deadline: new Date(2026, 9, 25), key: "oct26" },
  { name: "November", year: 2026, deadline: new Date(2026, 10, 25), key: "nov26" },
  { name: "Desember", year: 2026, deadline: new Date(2026, 11, 20), key: "dec26" },
];

const DAYS_ID = ["MINGGU", "SENIN", "SELASA", "RABU", "KAMIS", "JUMAT", "SABTU"];

const todayDate = new Date();

// ─── HELPERS ─────────────────────────────────────────────────────────────────

function daysLeft(deadline) {
  return Math.ceil((deadline - todayDate) / 86400000);
}

function monthStatus(m) {
  const d = daysLeft(m.deadline);
  if (d < 0) return { label: "Lewat deadline", variant: "danger" };
  if (d <= 5) return { label: `${d} hari lagi`, variant: "warning" };
  return { label: `${d} hari lagi`, variant: "success" };
}

function fmtDate(date) {
  return date.toLocaleDateString("id-ID", { day: "numeric", month: "long", year: "numeric" }).toUpperCase();
}

function todayISO() {
  return todayDate.toISOString().split("T")[0];
}

// ─── STYLES ──────────────────────────────────────────────────────────────────

const COLORS = {
  bg: "#0f1117",
  surface: "#16191f",
  surfaceHover: "#1c2028",
  border: "rgba(255,255,255,0.08)",
  borderStrong: "rgba(255,255,255,0.16)",
  accent: "#4f8ef7",
  accentDim: "rgba(79,142,247,0.12)",
  text: "#e8eaf0",
  muted: "#7a8099",
  success: "#34d399",
  successDim: "rgba(52,211,153,0.12)",
  warning: "#fbbf24",
  warningDim: "rgba(251,191,36,0.12)",
  danger: "#f87171",
  dangerDim: "rgba(248,113,113,0.12)",
};

const S = {
  app: {
    background: COLORS.bg,
    minHeight: "100vh",
    fontFamily: "'Plus Jakarta Sans', system-ui, sans-serif",
    color: COLORS.text,
    padding: "0",
  },
  header: {
    borderBottom: `1px solid ${COLORS.border}`,
    padding: "20px 24px 16px",
    background: COLORS.surface,
  },
  headerSub: { fontSize: 11, color: COLORS.muted, letterSpacing: "0.1em", marginBottom: 4, textTransform: "uppercase", fontWeight: 500 },
  headerTitle: { fontSize: 22, fontWeight: 700, color: COLORS.text, letterSpacing: "-0.03em", lineHeight: 1.2 },
  tabs: {
    display: "flex",
    gap: 0,
    borderBottom: `1px solid ${COLORS.border}`,
    background: COLORS.surface,
    padding: "0 24px",
  },
  tab: (active) => ({
    padding: "12px 18px",
    fontSize: 13,
    cursor: "pointer",
    background: "none",
    border: "none",
    borderBottom: `2px solid ${active ? COLORS.accent : "transparent"}`,
    color: active ? COLORS.accent : COLORS.muted,
    fontFamily: "'Plus Jakarta Sans', system-ui, sans-serif",
    fontWeight: active ? 700 : 500,
    letterSpacing: "0.01em",
    transition: "color 0.15s, border-color 0.15s",
  }),
  content: { padding: "20px 24px" },
  grid12: {
    display: "grid",
    gridTemplateColumns: "repeat(auto-fill, minmax(130px, 1fr))",
    gap: 8,
    marginBottom: 20,
  },
  monthCard: (active, variant) => ({
    background: active ? COLORS.accentDim : COLORS.surface,
    border: `1px solid ${active ? COLORS.accent : variant === "danger" ? COLORS.dangerDim : COLORS.border}`,
    borderRadius: 10,
    padding: "10px 12px",
    cursor: "pointer",
    transition: "border-color 0.15s, background 0.15s",
  }),
  mcName: { fontSize: 13, fontWeight: 700, color: COLORS.text, letterSpacing: "-0.01em" },
  mcDeadline: { fontSize: 11, color: COLORS.muted, marginTop: 3, fontWeight: 500 },
  badge: (variant) => ({
    display: "inline-block",
    marginTop: 6,
    fontSize: 10,
    fontWeight: 700,
    padding: "2px 8px",
    borderRadius: 20,
    letterSpacing: "0.05em",
    textTransform: "uppercase",
    background: variant === "success" ? COLORS.successDim : variant === "warning" ? COLORS.warningDim : COLORS.dangerDim,
    color: variant === "success" ? COLORS.success : variant === "warning" ? COLORS.warning : COLORS.danger,
  }),
  banner: (variant) => ({
    padding: "10px 14px",
    borderRadius: 8,
    marginBottom: 16,
    fontSize: 13,
    fontWeight: 500,
    lineHeight: 1.5,
    background: variant === "success" ? COLORS.successDim : variant === "warning" ? COLORS.warningDim : COLORS.dangerDim,
    color: variant === "success" ? COLORS.success : variant === "warning" ? COLORS.warning : COLORS.danger,
    border: `1px solid ${variant === "success" ? "rgba(52,211,153,0.25)" : variant === "warning" ? "rgba(251,191,36,0.25)" : "rgba(248,113,113,0.25)"}`,
  }),
  statsRow: {
    display: "grid",
    gridTemplateColumns: "1fr 1fr",
    gap: 8,
    marginBottom: 20,
  },
  statCard: {
    background: COLORS.surface,
    border: `1px solid ${COLORS.border}`,
    borderRadius: 10,
    padding: "12px 16px",
  },
  statLabel: { fontSize: 11, color: COLORS.muted, letterSpacing: "0.08em", marginBottom: 4, fontWeight: 600, textTransform: "uppercase" },
  statValue: { fontSize: 26, fontWeight: 700, color: COLORS.text, letterSpacing: "-0.04em" },
  statUnit: { fontSize: 13, color: COLORS.muted, marginLeft: 4, fontWeight: 500 },
  sectionLabel: { fontSize: 11, color: COLORS.muted, letterSpacing: "0.1em", marginBottom: 10, marginTop: 4, fontWeight: 700, textTransform: "uppercase" },
  itemCard: (checked) => ({
    background: checked ? COLORS.accentDim : COLORS.surface,
    border: `1px solid ${checked ? COLORS.accent : COLORS.border}`,
    borderRadius: 8,
    padding: "10px 12px",
    marginBottom: 6,
    display: "flex",
    gap: 10,
    alignItems: "flex-start",
    cursor: "pointer",
    transition: "border-color 0.15s, background 0.15s",
  }),
  itemName: { fontSize: 13, color: COLORS.text, lineHeight: 1.5, fontWeight: 500 },
  itemNote: { fontSize: 11, color: COLORS.muted, marginTop: 3, lineHeight: 1.5, fontWeight: 400 },
  poinBadge: {
    fontSize: 11,
    fontWeight: 700,
    padding: "3px 9px",
    background: COLORS.accentDim,
    color: COLORS.accent,
    borderRadius: 20,
    whiteSpace: "nowrap",
    letterSpacing: "0.02em",
    fontFamily: "'IBM Plex Mono', monospace",
  },
  catHeader: {
    fontSize: 10,
    fontWeight: 700,
    color: COLORS.muted,
    letterSpacing: "0.14em",
    textTransform: "uppercase",
    padding: "14px 0 6px",
    borderBottom: `1px solid ${COLORS.border}`,
    marginBottom: 8,
  },
  checkBox: (checked) => ({
    width: 18,
    height: 18,
    minWidth: 18,
    borderRadius: 4,
    border: `1.5px solid ${checked ? COLORS.accent : COLORS.borderStrong}`,
    background: checked ? COLORS.accent : "transparent",
    display: "flex",
    alignItems: "center",
    justifyContent: "center",
    marginTop: 1,
    flexShrink: 0,
    transition: "background 0.15s, border-color 0.15s",
  }),
  formSection: {
    background: COLORS.surface,
    border: `1px solid ${COLORS.border}`,
    borderRadius: 10,
    padding: "16px 18px",
    marginBottom: 12,
  },
  formTitle: { fontSize: 14, fontWeight: 700, color: COLORS.text, marginBottom: 14, letterSpacing: "-0.01em" },
  formRow: (cols) => ({
    display: "grid",
    gridTemplateColumns: cols || "1fr 1fr",
    gap: 10,
    marginBottom: 10,
  }),
  formLabel: { fontSize: 11, color: COLORS.muted, letterSpacing: "0.08em", marginBottom: 5, display: "block", fontWeight: 600, textTransform: "uppercase" },
  input: {
    width: "100%",
    padding: "8px 11px",
    fontSize: 13,
    fontWeight: 500,
    border: `1px solid ${COLORS.border}`,
    borderRadius: 6,
    background: COLORS.bg,
    color: COLORS.text,
    fontFamily: "'Plus Jakarta Sans', system-ui, sans-serif",
    outline: "none",
    boxSizing: "border-box",
  },
  textarea: {
    width: "100%",
    padding: "8px 11px",
    fontSize: 13,
    fontWeight: 400,
    border: `1px solid ${COLORS.border}`,
    borderRadius: 6,
    background: COLORS.bg,
    color: COLORS.text,
    fontFamily: "'Plus Jakarta Sans', system-ui, sans-serif",
    outline: "none",
    minHeight: 80,
    resize: "vertical",
    boxSizing: "border-box",
    lineHeight: 1.6,
  },
  select: {
    width: "100%",
    padding: "8px 11px",
    fontSize: 13,
    fontWeight: 500,
    border: `1px solid ${COLORS.border}`,
    borderRadius: 6,
    background: COLORS.bg,
    color: COLORS.text,
    fontFamily: "'Plus Jakarta Sans', system-ui, sans-serif",
    outline: "none",
    boxSizing: "border-box",
  },
  btnPrimary: {
    width: "100%",
    padding: "11px",
    fontSize: 14,
    fontWeight: 700,
    cursor: "pointer",
    background: COLORS.accent,
    border: "none",
    borderRadius: 8,
    color: "#fff",
    fontFamily: "'Plus Jakarta Sans', system-ui, sans-serif",
    letterSpacing: "0.01em",
    marginTop: 4,
    transition: "opacity 0.15s",
  },
  btnSecondary: {
    padding: "7px 16px",
    fontSize: 12,
    fontWeight: 600,
    cursor: "pointer",
    background: COLORS.surface,
    border: `1px solid ${COLORS.border}`,
    borderRadius: 6,
    color: COLORS.muted,
    fontFamily: "'Plus Jakarta Sans', system-ui, sans-serif",
    letterSpacing: "0.02em",
    marginTop: 8,
  },
  preview: {
    background: COLORS.bg,
    border: `1px solid ${COLORS.border}`,
    borderRadius: 8,
    padding: "16px 18px",
    fontSize: 12,
    lineHeight: 1.9,
    color: COLORS.text,
    whiteSpace: "pre-wrap",
    fontFamily: "'IBM Plex Mono', 'Courier New', monospace",
    fontWeight: 400,
    marginTop: 14,
    overflowX: "auto",
  },
  progressWrap: {
    height: 4,
    background: COLORS.border,
    borderRadius: 2,
    marginTop: 6,
  },
  progressFill: (pct) => ({
    height: 4,
    background: COLORS.accent,
    borderRadius: 2,
    width: `${pct}%`,
    transition: "width 0.3s",
  }),
};

// ─── TRACKER PAGE ─────────────────────────────────────────────────────────────

function TrackerPage({ checked, onToggle }) {
  const [selMonth, setSelMonth] = useState(todayDate.getMonth());
  const lhpItems = NSP_ITEMS.filter((x) => x.lhp);
  const m = MONTHS[selMonth];
  const st = monthStatus(m);
  const mk = m.key;
  const monthChecked = checked[mk] || {};

  const score = lhpItems.reduce((s, it) => s + (monthChecked[it.id] ? it.poin : 0), 0);
  const cnt = lhpItems.filter((it) => monthChecked[it.id]).length;
  const d = daysLeft(m.deadline);

  const bannerText =
    d < 0
      ? `⚠️  Deadline ${m.name} sudah lewat (${fmtDate(m.deadline)})`
      : d <= 5
      ? `⏰  Deadline ${m.name}: ${fmtDate(m.deadline)} — tersisa ${d} hari, segera kumpulkan LHP!`
      : `✅  Deadline ${m.name}: ${fmtDate(m.deadline)} — tersisa ${d} hari`;

  const cats = [...new Set(lhpItems.map((x) => x.cat))];

  return (
    <div>
      <div style={S.grid12}>
        {MONTHS.map((mo, i) => {
          const s = monthStatus(mo);
          return (
            <div
              key={mo.key}
              style={S.monthCard(selMonth === i, s.variant)}
              onClick={() => setSelMonth(i)}
            >
              <div style={S.mcName}>{mo.name}</div>
              <div style={S.mcDeadline}>Deadline {mo.deadline.toLocaleDateString("id-ID", { day: "numeric", month: "short" })}</div>
              <span style={S.badge(s.variant)}>{s.label}</span>
            </div>
          );
        })}
      </div>

      <div style={S.banner(st.variant)}>{bannerText}</div>

      <div style={S.statsRow}>
        <div style={S.statCard}>
          <div style={S.statLabel}>ESTIMASI POIN NSP (LHP)</div>
          <div style={S.statValue}>{score.toFixed(2)}<span style={S.statUnit}>poin</span></div>
        </div>
        <div style={S.statCard}>
          <div style={S.statLabel}>ITEM DISELESAIKAN</div>
          <div style={S.statValue}>{cnt}<span style={S.statUnit}>/ {lhpItems.length}</span></div>
          <div style={S.progressWrap}><div style={S.progressFill((cnt / lhpItems.length) * 100)} /></div>
        </div>
      </div>

      <div style={S.sectionLabel}>ITEM NSP YANG MEMERLUKAN LHP — {m.name.toUpperCase()} {m.year}</div>

      {cats.map((cat) => (
        <div key={cat}>
          <div style={S.catHeader}>{cat}</div>
          {lhpItems.filter((it) => it.cat === cat).map((item) => {
            const chk = !!monthChecked[item.id];
            return (
              <div key={item.id} style={S.itemCard(chk)} onClick={() => onToggle(mk, item.id, !chk)}>
                <div style={S.checkBox(chk)}>
                  {chk && (
                    <svg width="10" height="8" viewBox="0 0 10 8" fill="none">
                      <path d="M1 4L3.5 6.5L9 1" stroke="white" strokeWidth="1.8" strokeLinecap="round" strokeLinejoin="round" />
                    </svg>
                  )}
                </div>
                <div style={{ flex: 1 }}>
                  <div style={S.itemName}>{item.nama}</div>
                  {item.keterangan && <div style={S.itemNote}>{item.keterangan}</div>}
                </div>
                <span style={S.poinBadge}>{item.poin.toFixed(2)}</span>
              </div>
            );
          })}
        </div>
      ))}
    </div>
  );
}

// ─── NSP PAGE ─────────────────────────────────────────────────────────────────

function NSPPage({ checked }) {
  const cats = [...new Set(NSP_ITEMS.map((x) => x.cat))];
  const lhpItems = NSP_ITEMS.filter((x) => x.lhp);
  const maxPoin = lhpItems.reduce((s, x) => s + x.poin, 0);

  const everChecked = (id) => Object.values(checked).some((mk) => mk[id]);
  const grandTotal = lhpItems.reduce((s, it) => s + (everChecked(it.id) ? it.poin : 0), 0);
  const pct = Math.min(100, (grandTotal / maxPoin) * 100);

  return (
    <div>
      <div style={S.statsRow}>
        <div style={S.statCard}>
          <div style={S.statLabel}>TOTAL POIN NSP VIA LHP (SEMUA BULAN)</div>
          <div style={S.statValue}>{grandTotal.toFixed(2)}<span style={S.statUnit}>/ {maxPoin.toFixed(1)}</span></div>
          <div style={S.progressWrap}><div style={S.progressFill(pct)} /></div>
        </div>
        <div style={S.statCard}>
          <div style={S.statLabel}>ITEM LHP SELESAI</div>
          <div style={S.statValue}>{lhpItems.filter((it) => everChecked(it.id)).length}<span style={S.statUnit}>/ {lhpItems.length}</span></div>
        </div>
      </div>

      {cats.map((cat) => (
        <div key={cat}>
          <div style={S.catHeader}>{cat}</div>
          {NSP_ITEMS.filter((it) => it.cat === cat).map((item) => {
            const chk = everChecked(item.id);
            return (
              <div key={item.id} style={{ ...S.itemCard(chk), cursor: "default" }}>
                <div style={S.checkBox(chk)}>
                  {chk && (
                    <svg width="10" height="8" viewBox="0 0 10 8" fill="none">
                      <path d="M1 4L3.5 6.5L9 1" stroke="white" strokeWidth="1.8" strokeLinecap="round" strokeLinejoin="round" />
                    </svg>
                  )}
                </div>
                <div style={{ flex: 1 }}>
                  <div style={S.itemName}>{item.nama}</div>
                  <div style={{ display: "flex", gap: 6, marginTop: 4, flexWrap: "wrap" }}>
                    <span style={{ ...S.poinBadge, background: "rgba(255,255,255,0.06)", color: COLORS.muted, fontSize: 10 }}>{item.pemberi}</span>
                    {item.lhp && <span style={{ ...S.poinBadge, background: COLORS.warningDim, color: COLORS.warning, fontSize: 10 }}>PERLU LHP</span>}
                    {item.keterangan && <span style={{ fontSize: 11, color: COLORS.muted }}>{item.keterangan}</span>}
                  </div>
                </div>
                <span style={S.poinBadge}>{item.poin.toFixed(2)}</span>
              </div>
            );
          })}
        </div>
      ))}
    </div>
  );
}

// ─── DOCX GENERATOR ──────────────────────────────────────────────────────────
// Reads a File as ArrayBuffer → returns Uint8Array
function readFileAsArrayBuffer(file) {
  return new Promise((resolve, reject) => {
    const reader = new FileReader();
    reader.onload = (e) => resolve(new Uint8Array(e.target.result));
    reader.onerror = reject;
    reader.readAsArrayBuffer(file);
  });
}

// Detects image type from extension
function imgType(file) {
  const ext = file.name.split(".").pop().toLowerCase();
  if (ext === "jpg" || ext === "jpeg") return "jpg";
  if (ext === "png") return "png";
  if (ext === "gif") return "gif";
  if (ext === "bmp") return "bmp";
  if (ext === "webp") return "png"; // treat as png container
  return "jpg";
}

// Gets natural image dimensions from File → scales to fit maxW x maxH
function getImgDims(file, maxW = 380, maxH = 280) {
  return new Promise((resolve) => {
    const url = URL.createObjectURL(file);
    const img = new Image();
    img.onload = () => {
      let w = img.naturalWidth;
      let h = img.naturalHeight;
      const scaleW = maxW / w;
      const scaleH = maxH / h;
      const scale = Math.min(scaleW, scaleH, 1);
      URL.revokeObjectURL(url);
      resolve({ width: Math.round(w * scale), height: Math.round(h * scale) });
    };
    img.onerror = () => { URL.revokeObjectURL(url); resolve({ width: maxW, height: maxH }); };
    img.src = url;
  });
}

// Logo Akpol sebagai base64 — di-fetch saat pertama kali dibutuhkan
let _logoCache = null;
async function fetchLogoBase64() {
  if (_logoCache) return _logoCache;
  const resp = await fetch("/Logo_Akademi_Kepolisian.png");
  const arrBuf = await resp.arrayBuffer();
  _logoCache = new Uint8Array(arrBuf);
  return _logoCache;
}

async function generateAndDownloadDocx(fields) {
  const {
    namaKegiatan, hari, tglFmt, tglFmtUpper, tempatUp, uraianUp,
    kota, tglLaporanFmt, item, photos,
    pangkatDantontar = TARUNA.pangkat_dantontar,
    nrpDantontar    = TARUNA.nrp_dantontar,
    pangkatDankitar = TARUNA.pangkat_dankitar,
    nrpDankitar     = TARUNA.nrp_dankitar,
  } = fields;

  const noBorder = {
    top:     { style: BorderStyle.NIL },
    bottom:  { style: BorderStyle.NIL },
    left:    { style: BorderStyle.NIL },
    right:   { style: BorderStyle.NIL },
    insideH: { style: BorderStyle.NIL },
    insideV: { style: BorderStyle.NIL },
  };
  const thinBorder = {
    top:     { style: BorderStyle.SINGLE, size: 4, color: "000000" },
    bottom:  { style: BorderStyle.SINGLE, size: 4, color: "000000" },
    left:    { style: BorderStyle.SINGLE, size: 4, color: "000000" },
    right:   { style: BorderStyle.SINGLE, size: 4, color: "000000" },
    insideH: { style: BorderStyle.SINGLE, size: 4, color: "000000" },
    insideV: { style: BorderStyle.SINGLE, size: 4, color: "000000" },
  };
  const bottomOnlyBorder = {
    top:     { style: BorderStyle.NIL },
    bottom:  { style: BorderStyle.SINGLE, size: 12, color: "000000" },
    left:    { style: BorderStyle.NIL },
    right:   { style: BorderStyle.NIL },
    insideH: { style: BorderStyle.NIL },
    insideV: { style: BorderStyle.NIL },
  };

  // Arial bold, Arial MT regular, merah bold underline
  const B   = (text, size = 22, color = "000000") => new TextRun({ text, bold: true, size, font: "Arial", color });
  const R   = (text, size = 22, color = "000000") => new TextRun({ text, size, font: "Arial MT", color });
  const RBU = (text, size = 22) => new TextRun({ text, bold: true, underline: { type: "single" }, size, font: "Arial", color: "FF0000" });
  const para  = (children, opts = {}) => new Paragraph({ children, spacing: { after: 0, before: 0 }, ...opts });
  const blank = () => para([R("")]);

  // Logo
  let logoData = null;
  try { logoData = await fetchLogoBase64(); } catch (_) {}

  // Header kiri atas: logo (kiri) + nama satuan (kiri, vertikal tengah)
  const logoColW = 1300;
  const textColW = 7637;
  const headerSatuanTable = new Table({
    width: { size: 8937, type: WidthType.DXA },
    columnWidths: [logoColW, textColW],
    borders: noBorder,
    rows: [
      new TableRow({
        children: [
          new TableCell({
            borders: noBorder, width: { size: logoColW, type: WidthType.DXA },
            margins: { top: 0, bottom: 0, left: 0, right: 80 },
            verticalAlign: VerticalAlign.CENTER,
            children: logoData
              ? [para([new ImageRun({ type: "jpg", data: logoData, transformation: { width: 88, height: 100 }, altText: { title: "Logo Akpol", description: "Logo Akademi Kepolisian", name: "logo-akpol" } })])]
              : [blank()],
          }),
          new TableCell({
            borders: noBorder, width: { size: textColW, type: WidthType.DXA },
            margins: { top: 0, bottom: 0, left: 0, right: 0 },
            verticalAlign: VerticalAlign.CENTER,
            children: [
              para([B("RESIMEN KORPS TARUNA DAN SISWA", 22)]),
              para([B("DETASEMEN TARUNA TK III/60/BD", 22)]),
            ],
          }),
        ],
      }),
    ],
  });

  // Garis lurus di bawah header satuan
  const garisTable = new Table({
    width: { size: 8937, type: WidthType.DXA },
    columnWidths: [8937],
    borders: noBorder,
    rows: [
      new TableRow({
        children: [
          new TableCell({
            borders: bottomOnlyBorder,
            width: { size: 8937, type: WidthType.DXA },
            margins: { top: 40, bottom: 40, left: 0, right: 0 },
            children: [blank()],
          }),
        ],
      }),
    ],
  });

  // Tabel identitas: NO.AK | NAMA TARUNA | TON/KI
  const cellW = 2979;
  const headerTable = new Table({
    width: { size: 8937, type: WidthType.DXA },
    columnWidths: [cellW, cellW, cellW],
    borders: thinBorder,
    rows: [
      new TableRow({
        children: [
          new TableCell({
            borders: thinBorder, width: { size: cellW, type: WidthType.DXA },
            margins: { top: 80, bottom: 80, left: 120, right: 120 },
            children: [
              para([B("NO. AKADEMI", 20)], { alignment: AlignmentType.CENTER }),
              para([B("200507014166", 20)], { alignment: AlignmentType.CENTER }),
            ],
          }),
          new TableCell({
            borders: thinBorder, width: { size: cellW, type: WidthType.DXA },
            margins: { top: 80, bottom: 80, left: 120, right: 120 },
            children: [
              para([B("NAMA TARUNA", 20)], { alignment: AlignmentType.CENTER }),
              para([B("MUHAMMAD RADHITYA BHARA KUNTARA", 20)], { alignment: AlignmentType.CENTER }),
            ],
          }),
          new TableCell({
            borders: thinBorder, width: { size: cellW, type: WidthType.DXA },
            margins: { top: 80, bottom: 80, left: 120, right: 120 },
            children: [
              para([B("TON/KI", 20)], { alignment: AlignmentType.CENTER }),
              para([B("2/B", 20)], { alignment: AlignmentType.CENTER }),
            ],
          }),
        ],
      }),
    ],
  });

  // Info: NAMA KEGIATAN / WAKTU / TEMPAT
  const colLabel = 2800; const colColon = 240; const colVal = 5897;
  const infoRow = (label, val) => new TableRow({
    children: [
      new TableCell({ borders: noBorder, width: { size: colLabel, type: WidthType.DXA }, margins: { top: 40, bottom: 40, left: 0, right: 0 }, children: [para([B(label)])] }),
      new TableCell({ borders: noBorder, width: { size: colColon, type: WidthType.DXA }, margins: { top: 40, bottom: 40, left: 0, right: 0 }, children: [para([B(":")])] }),
      new TableCell({ borders: noBorder, width: { size: colVal,   type: WidthType.DXA }, margins: { top: 40, bottom: 40, left: 0, right: 0 }, children: [para([R(val)])] }),
    ],
  });
  const infoTable = new Table({
    width: { size: 8937, type: WidthType.DXA },
    columnWidths: [colLabel, colColon, colVal],
    borders: noBorder,
    rows: [
      infoRow("NAMA KEGIATAN", namaKegiatan),
      infoRow("WAKTU PELAKSANAAN", `${hari}, ${tglFmtUpper}`),
      infoRow("TEMPAT", tempatUp),
    ],
  });

  // URAIAN KEGIATAN: header merah bold underline CENTER + isi konten
  const uraianFullText =
    `PADA HARI ${hari} TANGGAL ${tglFmtUpper}, SAYA MUHAMMAD RADHITYA BHARA KUNTARA TARUNA AKPOL/PANGKAT NO.AK 200507014166 TK. III/60/BD TELAH MELAKSANAKAN KEGIATAN ${namaKegiatan} DI ${tempatUp}. ${uraianUp}`;
  const uraianTable = new Table({
    width: { size: 8937, type: WidthType.DXA },
    columnWidths: [8937],
    borders: thinBorder,
    rows: [
      new TableRow({ children: [new TableCell({ borders: thinBorder, width: { size: 8937, type: WidthType.DXA }, margins: { top: 60, bottom: 60, left: 120, right: 120 }, children: [para([RBU("URAIAN KEGIATAN")], { alignment: AlignmentType.CENTER })] })] }),
      new TableRow({ children: [new TableCell({ borders: thinBorder, width: { size: 8937, type: WidthType.DXA }, margins: { top: 80, bottom: 200, left: 120, right: 120 }, children: [para([R(uraianFullText)])] })] }),
    ],
  });

  // DISPOSISI DANTONTAR: header merah bold underline CENTER + baris kosong
  const disposisiTable = new Table({
    width: { size: 8937, type: WidthType.DXA },
    columnWidths: [8937],
    borders: thinBorder,
    rows: [
      new TableRow({ children: [new TableCell({ borders: thinBorder, width: { size: 8937, type: WidthType.DXA }, margins: { top: 60, bottom: 60, left: 120, right: 120 }, children: [para([RBU("DISPOSISI DANTONTAR")], { alignment: AlignmentType.CENTER })] })] }),
      new TableRow({ children: [new TableCell({ borders: thinBorder, width: { size: 8937, type: WidthType.DXA }, margins: { top: 80, bottom: 200, left: 120, right: 120 }, children: [blank()] })] }),
    ],
  });

  // Tanda tangan
  const sigColW = 4468;
  const sigTable = new Table({
    width: { size: 8937, type: WidthType.DXA },
    columnWidths: [sigColW, sigColW],
    borders: noBorder,
    rows: [
      new TableRow({
        children: [
          new TableCell({ borders: noBorder, width: { size: sigColW, type: WidthType.DXA }, margins: { top: 40, bottom: 0, left: 0, right: 0 }, children: [para([R("Mengetahui,")]), para([R("DANTONTAR 2B TARUNA TK III/60/BD")])] }),
          new TableCell({ borders: noBorder, width: { size: sigColW, type: WidthType.DXA }, margins: { top: 40, bottom: 0, left: 120, right: 0 }, children: [para([R(`${kota}, ${tglLaporanFmt}`)], { alignment: AlignmentType.CENTER }), para([R("YANG MEMBUAT")], { alignment: AlignmentType.CENTER })] }),
        ],
      }),
      new TableRow({
        children: [
          new TableCell({ borders: noBorder, width: { size: sigColW, type: WidthType.DXA }, margins: { top: 300, bottom: 300, left: 0, right: 0 }, children: [blank()] }),
          new TableCell({ borders: noBorder, width: { size: sigColW, type: WidthType.DXA }, margins: { top: 300, bottom: 300, left: 120, right: 0 }, children: [blank()] }),
        ],
      }),
      new TableRow({
        children: [
          new TableCell({ borders: noBorder, width: { size: sigColW, type: WidthType.DXA }, margins: { top: 0, bottom: 40, left: 0, right: 0 }, children: [para([B(TARUNA.dantontar)]), para([R(`${pangkatDantontar} NRP ${nrpDantontar}`)])] }),
          new TableCell({ borders: noBorder, width: { size: sigColW, type: WidthType.DXA }, margins: { top: 0, bottom: 40, left: 120, right: 0 }, children: [para([B(TARUNA.nama)], { alignment: AlignmentType.CENTER }), para([R(`BRIGTAR NO.AK ${TARUNA.noakad}`)], { alignment: AlignmentType.CENTER })] }),
        ],
      }),
    ],
  });

  // Halaman 2: DOKUMENTASI
  const photoItems = [];
  if (photos && photos.length > 0) {
    for (const { file, caption } of photos) {
      try {
        const data = await readFileAsArrayBuffer(file);
        const { width, height } = await getImgDims(file, 360, 260);
        photoItems.push({ data, type: imgType(file), width, height, name: file.name, caption: caption || "" });
      } catch (_) {}
    }
  }
  const photoPageChildren = [
    para([RBU("DOKUMENTASI", 24)], { alignment: AlignmentType.CENTER }),
    blank(),
  ];
  if (photoItems.length === 0) {
    photoPageChildren.push(para([R("(Tidak ada foto yang dilampirkan)")]));
  } else {
    const colW2 = 4468;
    for (let i = 0; i < photoItems.length; i += 2) {
      const left = photoItems[i];
      const right = photoItems[i + 1] || null;
      const makePhotoCell = (p, num) => {
        if (!p) return new TableCell({ borders: noBorder, width: { size: colW2, type: WidthType.DXA }, margins: { top: 60, bottom: 60, left: 60, right: 60 }, children: [blank()] });
        return new TableCell({ borders: noBorder, width: { size: colW2, type: WidthType.DXA }, margins: { top: 60, bottom: 60, left: 60, right: 60 }, children: [para([new ImageRun({ type: p.type, data: p.data, transformation: { width: p.width, height: p.height }, altText: { title: p.name, description: p.name, name: p.name } })], { alignment: AlignmentType.CENTER }), para([R(p.caption || `Foto ${num}`, 18)], { alignment: AlignmentType.CENTER })] });
      };
      photoPageChildren.push(new Table({ width: { size: 8937, type: WidthType.DXA }, columnWidths: [colW2, colW2], borders: noBorder, rows: [new TableRow({ children: [makePhotoCell(left, i + 1), makePhotoCell(right, i + 2)] })] }));
      photoPageChildren.push(blank());
    }
  }

  const pageProps = { size: { width: 11906, height: 16838 }, margin: { top: 1134, right: 1134, bottom: 1134, left: 1134 } };
  const doc = new Document({
    sections: [
      {
        properties: { page: pageProps },
        children: [
          headerSatuanTable,
          garisTable,
          blank(),
          para([B("LAPORAN HASIL PELAKSANAAN KEGIATAN TARUNA TK. III/60/BD", 24)], { alignment: AlignmentType.CENTER }),
          blank(),
          headerTable,
          blank(),
          infoTable,
          uraianTable,
          disposisiTable,
          blank(),
          sigTable,
          blank(),
          blank(),
          para([R("Mengetahui,")]),
          para([R("DANKITAR B TARUNA TK III/60/BD")]),
          blank(),
          blank(),
          blank(),
          para([B(TARUNA.dankitar)]),
          para([R(`${pangkatDankitar} NRP ${nrpDankitar}`)]),
        ],
      },
      {
        properties: { page: pageProps },
        children: photoPageChildren,
      },
    ],
  });

  const blob = await Packer.toBlob(doc);
  const url = URL.createObjectURL(blob);
  const a = document.createElement("a");
  const tglObj = new Date((fields.tanggalRaw || new Date().toISOString().split("T")[0]) + "T12:00:00");
  const bulanTahun = tglObj.toLocaleDateString("id-ID", { month: "short", year: "numeric" }).replace(" ", "-").toUpperCase();
  const kegSlug = (item?.id || "lhp").toUpperCase();
  a.href = url;
  a.download = `LHP_RADIT_${bulanTahun}_${kegSlug}.docx`;
  a.click();
  URL.revokeObjectURL(url);
}


// ─── FORM PAGE ─────────────────────────────────────────────────────────────────

function FormPage() {
  const [kegId, setKegId] = useState("");
  const [hari, setHari] = useState("SENIN");
  const [tanggal, setTanggal] = useState(todayISO());
  const [tempat, setTempat] = useState("");
  const [uraian, setUraian] = useState("");
  const [tglLaporan, setTglLaporan] = useState(todayISO());
  const [kota, setKota] = useState("Semarang");
  const [photos, setPhotos] = useState([]);

  // Editable pejabat fields
  const [pangkatDantontar, setPangkatDantontar] = useState(TARUNA.pangkat_dantontar);
  const [nrpDantontar, setNrpDantontar] = useState(TARUNA.nrp_dantontar);
  const [pangkatDankitar, setPangkatDankitar] = useState(TARUNA.pangkat_dankitar);
  const [nrpDankitar, setNrpDankitar] = useState(TARUNA.nrp_dankitar);

  const [loading, setLoading] = useState(false);
  const [done, setDone] = useState(false);
  const [error, setError] = useState("");
  const [showPreview, setShowPreview] = useState(false);

  const lhpItems = NSP_ITEMS.filter((x) => x.lhp);

  const handlePhotoAdd = (e) => {
    const files = Array.from(e.target.files);
    const newPhotos = files.map((file) => ({
      file,
      previewUrl: URL.createObjectURL(file),
      caption: "",
    }));
    setPhotos((prev) => [...prev, ...newPhotos]);
    e.target.value = "";
  };

  const handlePhotoRemove = (idx) => {
    setPhotos((prev) => {
      URL.revokeObjectURL(prev[idx].previewUrl);
      return prev.filter((_, i) => i !== idx);
    });
  };

  const handleCaptionChange = (idx, val) => {
    setPhotos((prev) => prev.map((p, i) => i === idx ? { ...p, caption: val } : p));
  };

  const handleGenerate = async () => {
    if (!kegId) { setError("Pilih kegiatan NSP terlebih dahulu."); return; }
    setError("");
    setLoading(true);
    setDone(false);
    try {
      const item = NSP_ITEMS.find((x) => x.id === kegId);
      const namaKegiatan = item ? item.nama.toUpperCase() : "[NAMA KEGIATAN]";
      const tglFmt = tanggal ? new Date(tanggal + "T12:00:00").toLocaleDateString("id-ID", { day: "numeric", month: "long", year: "numeric" }) : "[TANGGAL]";
      const tglFmtUpper = tglFmt.toUpperCase();
      const tglLaporanFmt = tglLaporan ? new Date(tglLaporan + "T12:00:00").toLocaleDateString("id-ID", { day: "numeric", month: "long", year: "numeric" }) : "[TANGGAL LAPORAN]";
      const tempatUp = (tempat || "[TEMPAT]").toUpperCase();
      const uraianUp = (uraian || "[URAIAN KEGIATAN]").toUpperCase();

      await generateAndDownloadDocx({
        namaKegiatan, hari, tglFmt, tglFmtUpper, tempatUp, uraianUp,
        kota, tglLaporanFmt, item, tanggalRaw: tanggal,
        photos: photos.map((p) => ({ file: p.file, caption: p.caption })),
        pangkatDantontar, nrpDantontar,
        pangkatDankitar, nrpDankitar,
      });
      setDone(true);
      setTimeout(() => setDone(false), 5000);
    } catch (e) {
      setError("Gagal generate DOCX. Pastikan library docx sudah terinstall: npm install docx");
      console.error(e);
    }
    setLoading(false);
  };

  return (
    <div>
      {/* ── Identitas ── */}
      <div style={S.formSection}>
        <div style={S.formTitle}>Data Identitas (Otomatis)</div>
        <div style={S.formRow("1fr 1fr 1fr")}>
          {[
            { label: "NAMA TARUNA", val: TARUNA.nama },
            { label: "NO. AKADEMI", val: TARUNA.noakad },
            { label: "TON/KI", val: TARUNA.tonki },
          ].map(({ label, val }) => (
            <div key={label}>
              <label style={S.formLabel}>{label}</label>
              <input style={{ ...S.input, opacity: 0.6 }} value={val} readOnly />
            </div>
          ))}
        </div>

        {/* Dantontar row */}
        <div style={{ marginBottom: 10 }}>
          <div style={{ fontSize: 11, color: COLORS.accent, fontWeight: 700, letterSpacing: "0.08em", textTransform: "uppercase", marginBottom: 6, paddingBottom: 4, borderBottom: `1px solid ${COLORS.border}` }}>
            Dantontar 2B
          </div>
          <div style={S.formRow("1fr 1fr 1fr")}>
            <div>
              <label style={S.formLabel}>NAMA DANTONTAR</label>
              <input style={{ ...S.input, opacity: 0.6 }} value={TARUNA.dantontar} readOnly />
            </div>
            <div>
              <label style={S.formLabel}>PANGKAT</label>
              <input
                style={S.input}
                type="text"
                value={pangkatDantontar}
                onChange={(e) => setPangkatDantontar(e.target.value.toUpperCase())}
                placeholder="IPDA"
              />
            </div>
            <div>
              <label style={S.formLabel}>NRP</label>
              <input
                style={S.input}
                type="text"
                value={nrpDantontar}
                onChange={(e) => setNrpDantontar(e.target.value)}
                placeholder="82020196"
              />
            </div>
          </div>
        </div>

        {/* Dankitar row */}
        <div>
          <div style={{ fontSize: 11, color: COLORS.accent, fontWeight: 700, letterSpacing: "0.08em", textTransform: "uppercase", marginBottom: 6, paddingBottom: 4, borderBottom: `1px solid ${COLORS.border}` }}>
            Dankitar B
          </div>
          <div style={S.formRow("1fr 1fr 1fr")}>
            <div>
              <label style={S.formLabel}>NAMA DANKITAR</label>
              <input style={{ ...S.input, opacity: 0.6 }} value={TARUNA.dankitar} readOnly />
            </div>
            <div>
              <label style={S.formLabel}>PANGKAT</label>
              <input
                style={S.input}
                type="text"
                value={pangkatDankitar}
                onChange={(e) => setPangkatDankitar(e.target.value.toUpperCase())}
                placeholder="AKP"
              />
            </div>
            <div>
              <label style={S.formLabel}>NRP</label>
              <input
                style={S.input}
                type="text"
                value={nrpDankitar}
                onChange={(e) => setNrpDankitar(e.target.value)}
                placeholder="92040556"
              />
            </div>
          </div>
        </div>
      </div>

      {/* ── Kegiatan ── */}
      <div style={S.formSection}>
        <div style={S.formTitle}>Isi Kegiatan LHP</div>

        <div style={{ marginBottom: 10 }}>
          <label style={S.formLabel}>NAMA KEGIATAN (pilih dari daftar NSP)</label>
          <select style={S.select} value={kegId} onChange={(e) => { setKegId(e.target.value); setError(""); }}>
            <option value="">-- pilih kegiatan NSP --</option>
            {lhpItems.map((it) => (
              <option key={it.id} value={it.id}>{it.nama}</option>
            ))}
          </select>
        </div>

        {kegId && (() => {
          const item = NSP_ITEMS.find((x) => x.id === kegId);
          return item?.keterangan ? (
            <div style={{ ...S.banner("warning"), marginBottom: 10 }}>
              📋 {item.keterangan}
            </div>
          ) : null;
        })()}

        <div style={S.formRow("1fr 1fr")}>
          <div>
            <label style={S.formLabel}>HARI</label>
            <select style={S.select} value={hari} onChange={(e) => setHari(e.target.value)}>
              {DAYS_ID.map((d) => <option key={d}>{d}</option>)}
            </select>
          </div>
          <div>
            <label style={S.formLabel}>TANGGAL</label>
            <input style={S.input} type="date" value={tanggal} onChange={(e) => setTanggal(e.target.value)} />
          </div>
        </div>

        <div style={{ marginBottom: 10 }}>
          <label style={S.formLabel}>TEMPAT PELAKSANAAN</label>
          <input style={S.input} type="text" placeholder="contoh: Masjid Akpol / Perpustakaan Akpol / ..." value={tempat} onChange={(e) => setTempat(e.target.value)} />
        </div>

        <div style={{ marginBottom: 10 }}>
          <label style={S.formLabel}>URAIAN KEGIATAN</label>
          <textarea style={S.textarea} placeholder="Deskripsikan kegiatan yang dilakukan secara singkat..." value={uraian} onChange={(e) => setUraian(e.target.value)} />
        </div>

        <div style={S.formRow("1fr 1fr")}>
          <div>
            <label style={S.formLabel}>TANGGAL LAPORAN</label>
            <input style={S.input} type="date" value={tglLaporan} onChange={(e) => setTglLaporan(e.target.value)} />
          </div>
          <div>
            <label style={S.formLabel}>KOTA</label>
            <input style={S.input} type="text" value={kota} onChange={(e) => setKota(e.target.value)} />
          </div>
        </div>
      </div>

      {/* ── Dokumentasi Foto ── */}
      <div style={S.formSection}>
        <div style={{ display: "flex", alignItems: "baseline", justifyContent: "space-between", marginBottom: 14 }}>
          <div style={S.formTitle}>
            Dokumentasi Foto
            <span style={{ fontSize: 12, color: COLORS.muted, fontWeight: 400, marginLeft: 8 }}>→ halaman 2 file .docx</span>
          </div>
          {photos.length > 0 && (
            <span style={{ fontSize: 11, color: COLORS.accent, fontWeight: 600 }}>{photos.length} foto</span>
          )}
        </div>

        {/* Drop zone */}
        <label style={{
          display: "flex", flexDirection: "column", alignItems: "center", justifyContent: "center",
          gap: 6, border: `1.5px dashed ${COLORS.borderStrong}`, borderRadius: 10,
          padding: "18px 16px", cursor: "pointer", marginBottom: 14, background: COLORS.bg,
        }}>
          <input type="file" accept="image/jpeg,image/png,image/gif,image/bmp,image/webp" multiple style={{ display: "none" }} onChange={handlePhotoAdd} />
          <div style={{ fontSize: 26, lineHeight: 1 }}>📷</div>
          <div style={{ fontSize: 13, fontWeight: 700, color: COLORS.text }}>Klik untuk tambah foto dokumentasi</div>
          <div style={{ fontSize: 11, color: COLORS.muted }}>JPG · PNG · WebP · bisa pilih lebih dari 1</div>
        </label>

        {/* Photo grid with caption inputs */}
        {photos.length > 0 && (
          <div style={{ display: "flex", flexDirection: "column", gap: 10 }}>
            <div style={{ fontSize: 11, color: COLORS.muted, fontWeight: 700, letterSpacing: "0.08em", textTransform: "uppercase" }}>
              Preview — {photos.length} foto akan masuk halaman DOKUMENTASI (2 per baris)
            </div>
            <div style={{ display: "grid", gridTemplateColumns: "repeat(auto-fill, minmax(160px, 1fr))", gap: 10 }}>
              {photos.map((p, i) => (
                <div key={i} style={{ background: COLORS.bg, border: `1px solid ${COLORS.border}`, borderRadius: 8, overflow: "hidden" }}>
                  {/* Thumbnail */}
                  <div style={{ position: "relative" }}>
                    <img
                      src={p.previewUrl}
                      alt={`foto-${i + 1}`}
                      style={{ width: "100%", aspectRatio: "4/3", objectFit: "cover", display: "block" }}
                    />
                    {/* Remove button */}
                    <button
                      onClick={() => handlePhotoRemove(i)}
                      style={{
                        position: "absolute", top: 5, right: 5,
                        background: "rgba(248,113,113,0.9)", border: "none", borderRadius: "50%",
                        width: 22, height: 22, cursor: "pointer", color: "#fff",
                        fontSize: 12, fontWeight: 700, display: "flex", alignItems: "center",
                        justifyContent: "center", lineHeight: 1, fontFamily: "inherit",
                      }}
                      title="Hapus foto"
                    >✕</button>
                    {/* Index badge */}
                    <div style={{
                      position: "absolute", bottom: 4, left: 4,
                      background: "rgba(0,0,0,0.6)", borderRadius: 4,
                      padding: "2px 6px", fontSize: 10, color: "#fff", fontWeight: 700,
                    }}>Foto {i + 1}</div>
                  </div>
                  {/* Caption input */}
                  <div style={{ padding: "6px 8px" }}>
                    <input
                      style={{ ...S.input, fontSize: 11, padding: "5px 7px" }}
                      type="text"
                      placeholder="Keterangan foto (opsional)"
                      value={p.caption}
                      onChange={(e) => handleCaptionChange(i, e.target.value)}
                    />
                  </div>
                </div>
              ))}
            </div>
          </div>
        )}
      </div>

      {error && (
        <div style={{ ...S.banner("danger"), marginBottom: 12 }}>{error}</div>
      )}

      <div style={{ display: "flex", gap: 8 }}>
        <button
          style={{ ...S.btnPrimary, opacity: loading ? 0.6 : 1, flex: 1 }}
          onClick={handleGenerate}
          disabled={loading}
        >
          {loading
            ? `⏳ Membuat DOCX${photos.length > 0 ? ` + ${photos.length} foto` : ""}...`
            : done
            ? "✓ File berhasil di-download!"
            : `↓ Generate & Download LHP (.docx)${photos.length > 0 ? ` + ${photos.length} foto` : ""}`}
        </button>
        <button
          style={{
            ...S.btnPrimary,
            background: showPreview ? "rgba(79,142,247,0.2)" : "transparent",
            border: `1px solid ${COLORS.accent}`,
            color: COLORS.accent,
            width: 120,
            flex: "none",
          }}
          onClick={() => setShowPreview((v) => !v)}
        >
          {showPreview ? "✕ Tutup" : "👁 Preview"}
        </button>
      </div>

      {done && (
        <div style={{ ...S.banner("success"), marginTop: 10 }}>
          ✅ File <strong>LHP_RADIT_*.docx</strong> berhasil dibuat — halaman 1: LHP, halaman 2: DOKUMENTASI {photos.length > 0 ? `(${photos.length} foto)` : "(kosong)"}. Tersimpan di Downloads.
        </div>
      )}

      {/* ── Preview LHP ── */}
      {showPreview && (() => {
        const item = NSP_ITEMS.find((x) => x.id === kegId);
        const namaKegiatan = item ? item.nama.toUpperCase() : "[NAMA KEGIATAN]";
        const tglFmt = tanggal ? new Date(tanggal + "T12:00:00").toLocaleDateString("id-ID", { day: "numeric", month: "long", year: "numeric" }) : "[TANGGAL]";
        const tglFmtUpper = tglFmt.toUpperCase();
        const tglLaporanFmt = tglLaporan ? new Date(tglLaporan + "T12:00:00").toLocaleDateString("id-ID", { day: "numeric", month: "long", year: "numeric" }) : "[TANGGAL LAPORAN]";
        const tempatUp = (tempat || "[TEMPAT]").toUpperCase();
        const uraianUp = (uraian || "[URAIAN KEGIATAN]").toUpperCase();
        return (
          <div style={{ marginTop: 20 }}>
            <div style={{ fontSize: 11, color: COLORS.muted, fontWeight: 700, letterSpacing: "0.08em", textTransform: "uppercase", marginBottom: 10 }}>
              Preview LHP — update otomatis saat form diisi
            </div>
            <LHPPreview data={{
              namaKegiatan, hari, tglFmtUpper, tempatUp, uraianUp,
              kota, tglLaporanFmt,
              pangkatDantontar, nrpDantontar,
              pangkatDankitar, nrpDankitar,
              photos,
            }} />
          </div>
        );
      })()}
    </div>
  );
}

// ─── LHP PREVIEW COMPONENT ────────────────────────────────────────────────────

function LHPPreview({ data }) {
  const {
    namaKegiatan, hari, tglFmtUpper, tempatUp, uraianUp,
    kota, tglLaporanFmt,
    pangkatDantontar, nrpDantontar,
    pangkatDankitar, nrpDankitar,
    photos,
  } = data;

  const uraianFullText = `PADA HARI ${hari} TANGGAL ${tglFmtUpper}, SAYA MUHAMMAD RADHITYA BHARA KUNTARA TARUNA AKPOL/PANGKAT NO.AK 200507014166 TK. III/60/BD TELAH MELAKSANAKAN KEGIATAN ${namaKegiatan} DI ${tempatUp}. ${uraianUp}`;

  const pv = {
    // Kertas A4 putih
    page: {
      background: "#fff",
      color: "#000",
      fontFamily: "Arial MT, Arial, sans-serif",
      fontSize: 11,
      padding: "32px 40px",
      width: "100%",
      maxWidth: 720,
      margin: "0 auto",
      boxSizing: "border-box",
      border: "1px solid #ccc",
      borderRadius: 4,
      lineHeight: 1.5,
    },
    // Header kiri atas
    headerRow: {
      display: "flex", alignItems: "center", gap: 10, marginBottom: 4,
    },
    logoBox: {
      width: 48, height: 54, flexShrink: 0,
      display: "flex", alignItems: "center", justifyContent: "center",
      fontSize: 28,
    },
    satuanText: { lineHeight: 1.3 },
    satuanLine: { fontWeight: 700, fontSize: 11, fontFamily: "Arial, sans-serif" },
    divider: { borderTop: "2px solid #000", margin: "6px 0 10px" },
    // Judul tengah
    judulCenter: { textAlign: "center", fontWeight: 700, fontSize: 11, fontFamily: "Arial, sans-serif", marginBottom: 8 },
    // Tabel identitas
    idTable: { width: "100%", borderCollapse: "collapse", marginBottom: 10 },
    idCell: { border: "1px solid #000", padding: "4px 6px", textAlign: "center", fontWeight: 700, fontFamily: "Arial, sans-serif", fontSize: 10, width: "33.33%" },
    // Info rows
    infoRow: { display: "grid", gridTemplateColumns: "130px 14px 1fr", gap: 0, marginBottom: 2 },
    infoLabel: { fontWeight: 700, fontFamily: "Arial, sans-serif", fontSize: 10 },
    infoColon: { fontWeight: 700, fontFamily: "Arial, sans-serif", fontSize: 10 },
    infoVal: { fontFamily: "Arial MT, Arial, sans-serif", fontSize: 10 },
    // Tabel uraian / disposisi
    boxTable: { width: "100%", borderCollapse: "collapse", marginBottom: 0 },
    boxHeader: {
      border: "1px solid #000", padding: "4px 8px",
      textAlign: "center", fontWeight: 700, fontFamily: "Arial, sans-serif",
      fontSize: 10, color: "#c00", textDecoration: "underline",
    },
    boxContent: {
      border: "1px solid #000", padding: "6px 8px",
      fontFamily: "Arial MT, Arial, sans-serif", fontSize: 10,
      minHeight: 40,
    },
    boxEmpty: {
      border: "1px solid #000", padding: "6px 8px", minHeight: 50,
    },
    // Tanda tangan
    sigRow: { display: "grid", gridTemplateColumns: "1fr 1fr", gap: 8, marginTop: 10 },
    sigCell: { fontSize: 10 },
    sigCellRight: { fontSize: 10, textAlign: "center" },
    sigSpacer: { height: 36 },
    sigName: { fontWeight: 700, fontFamily: "Arial, sans-serif", fontSize: 10 },
    sigNrp: { fontFamily: "Arial MT, Arial, sans-serif", fontSize: 10 },
    // Dankitar
    dankitarBlock: { marginTop: 12, fontSize: 10 },
  };

  return (
    <div style={pv.page}>
      {/* Header kiri atas: logo + satuan */}
      <div style={pv.headerRow}>
        <div style={pv.logoBox}>
          <img src="/Logo_Akademi_Kepolisian.png" alt="Logo Akpol"
            style={{ width: 48, height: 54, objectFit: "contain" }}
            onError={(e) => { e.target.style.display = "none"; }}
          />
        </div>
        <div style={pv.satuanText}>
          <div style={pv.satuanLine}>RESIMEN KORPS TARUNA DAN SISWA</div>
          <div style={pv.satuanLine}>DETASEMEN TARUNA TK III/60/BD</div>
        </div>
      </div>

      {/* Garis lurus */}
      <div style={pv.divider} />

      {/* Judul */}
      <div style={pv.judulCenter}>
        LAPORAN HASIL PELAKSANAAN KEGIATAN TARUNA TK. III/60/BD
      </div>

      {/* Tabel identitas */}
      <table style={pv.idTable}>
        <tbody>
          <tr>
            <td style={pv.idCell}>NO. AKADEMI<br />200507014166</td>
            <td style={pv.idCell}>NAMA TARUNA<br />MUHAMMAD RADHITYA BHARA KUNTARA</td>
            <td style={pv.idCell}>TON/KI<br />2/B</td>
          </tr>
        </tbody>
      </table>

      {/* Info */}
      <div style={{ marginBottom: 8 }}>
        {[
          ["NAMA KEGIATAN", namaKegiatan],
          ["WAKTU PELAKSANAAN", `${hari}, ${tglFmtUpper}`],
          ["TEMPAT", tempatUp],
        ].map(([label, val]) => (
          <div key={label} style={pv.infoRow}>
            <span style={pv.infoLabel}>{label}</span>
            <span style={pv.infoColon}>:</span>
            <span style={pv.infoVal}>{val || <span style={{ color: "#aaa" }}>—</span>}</span>
          </div>
        ))}
      </div>

      {/* URAIAN KEGIATAN */}
      <table style={pv.boxTable}>
        <tbody>
          <tr><td style={pv.boxHeader}>URAIAN KEGIATAN</td></tr>
          <tr><td style={pv.boxContent}>{uraianFullText}</td></tr>
        </tbody>
      </table>

      {/* DISPOSISI DANTONTAR */}
      <table style={{ ...pv.boxTable, marginTop: 0 }}>
        <tbody>
          <tr><td style={pv.boxHeader}>DISPOSISI DANTONTAR</td></tr>
          <tr><td style={pv.boxEmpty} /></tr>
        </tbody>
      </table>

      {/* Tanda tangan */}
      <div style={pv.sigRow}>
        <div style={pv.sigCell}>
          <div>Mengetahui,</div>
          <div>DANTONTAR 2B TARUNA TK III/60/BD</div>
          <div style={pv.sigSpacer} />
          <div style={pv.sigName}>{TARUNA.dantontar}</div>
          <div style={pv.sigNrp}>{pangkatDantontar} NRP {nrpDantontar}</div>
        </div>
        <div style={pv.sigCellRight}>
          <div>{kota}, {tglLaporanFmt}</div>
          <div>YANG MEMBUAT</div>
          <div style={pv.sigSpacer} />
          <div style={pv.sigName}>{TARUNA.nama}</div>
          <div style={pv.sigNrp}>BRIGTAR NO.AK {TARUNA.noakad}</div>
        </div>
      </div>

      {/* Dankitar */}
      <div style={pv.dankitarBlock}>
        <div>Mengetahui,</div>
        <div>DANKITAR B TARUNA TK III/60/BD</div>
        <div style={{ height: 36 }} />
        <div style={pv.sigName}>{TARUNA.dankitar}</div>
        <div style={pv.sigNrp}>{pangkatDankitar} NRP {nrpDankitar}</div>
      </div>

      {/* Dokumentasi foto preview (thumbnail kecil) */}
      {photos && photos.length > 0 && (
        <div style={{ marginTop: 16, borderTop: "1px dashed #ccc", paddingTop: 10 }}>
          <div style={{ ...pv.judulCenter, marginBottom: 8, color: "#c00", textDecoration: "underline" }}>
            DOKUMENTASI ({photos.length} foto)
          </div>
          <div style={{ display: "grid", gridTemplateColumns: "repeat(auto-fill, minmax(100px, 1fr))", gap: 6 }}>
            {photos.map((p, i) => (
              <div key={i} style={{ textAlign: "center" }}>
                <img src={p.previewUrl} alt={`foto-${i + 1}`}
                  style={{ width: "100%", aspectRatio: "4/3", objectFit: "cover", borderRadius: 3, border: "1px solid #ddd" }}
                />
                <div style={{ fontSize: 9, color: "#555", marginTop: 2 }}>{p.caption || `Foto ${i + 1}`}</div>
              </div>
            ))}
          </div>
        </div>
      )}
    </div>
  );
}


// ─── REMINDER SYSTEM ──────────────────────────────────────────────────────────

function getUpcomingDeadlines(checked) {
  const now = new Date();
  const lhpItems = NSP_ITEMS.filter((x) => x.lhp);
  const reminders = [];
  MONTHS.forEach((m) => {
    const d = Math.ceil((m.deadline - now) / 86400000);
    if (d < 0 || d > 7) return;
    const mk = m.key;
    const monthChecked = checked[mk] || {};
    const uncheckedItems = lhpItems.filter((it) => !monthChecked[it.id]);
    if (uncheckedItems.length > 0) {
      reminders.push({ month: m, daysLeft: d, unchecked: uncheckedItems, total: lhpItems.length, done: lhpItems.length - uncheckedItems.length });
    }
  });
  return reminders;
}

function getUncheckedThisMonth(checked) {
  const now = new Date();
  const currentMonth = MONTHS.find((m) => m.deadline.getMonth() === now.getMonth() && m.deadline.getFullYear() === now.getFullYear());
  if (!currentMonth) return [];
  const mk = currentMonth.key;
  const monthChecked = checked[mk] || {};
  return NSP_ITEMS.filter((x) => x.lhp && !monthChecked[x.id]);
}

function useNotificationPermission() {
  useEffect(() => {
    if ("Notification" in window && Notification.permission === "default") {
      Notification.requestPermission();
    }
  }, []);
}

function useDeadlineNotifications(checked) {
  useEffect(() => {
    if (!("Notification" in window) || Notification.permission !== "granted") return;
    const checkAndNotify = () => {
      const reminders = getUpcomingDeadlines(checked);
      reminders.forEach((r) => {
        const key = `notified_${r.month.key}_${r.daysLeft}`;
        if (sessionStorage.getItem(key)) return;
        sessionStorage.setItem(key, "1");
        new Notification(`⏰ Deadline LHP ${r.month.name}`, {
          body: r.daysLeft === 0 ? `HARI INI deadline! ${r.unchecked.length} item belum selesai.` : `${r.daysLeft} hari lagi! ${r.unchecked.length}/${r.total} item belum selesai.`,
          icon: "/Logo_Akademi_Kepolisian.png",
          tag: `lhp-${r.month.key}`,
        });
      });
    };
    checkAndNotify();
    const interval = setInterval(checkAndNotify, 3600000);
    return () => clearInterval(interval);
  }, [checked]);
}

function ReminderBanner({ checked, onGoToTracker }) {
  const reminders = getUpcomingDeadlines(checked);
  const uncheckedThisMonth = getUncheckedThisMonth(checked);
  if (reminders.length === 0 && uncheckedThisMonth.length === 0) return null;

  return (
    <div style={{ marginBottom: 16 }}>
      {reminders.map((r) => (
        <div key={r.month.key} style={{
          background: r.daysLeft <= 2 ? COLORS.dangerDim : COLORS.warningDim,
          border: `1px solid ${r.daysLeft <= 2 ? "rgba(248,113,113,0.3)" : "rgba(251,191,36,0.3)"}`,
          borderRadius: 10, padding: "10px 14px", marginBottom: 8,
          display: "flex", alignItems: "center", justifyContent: "space-between", gap: 10,
        }}>
          <div>
            <div style={{ fontSize: 13, fontWeight: 700, color: r.daysLeft <= 2 ? COLORS.danger : COLORS.warning }}>
              {r.daysLeft === 0 ? `🚨 HARI INI deadline LHP ${r.month.name}!`
                : r.daysLeft === 1 ? `⚠️ BESOK deadline LHP ${r.month.name}!`
                : `⏰ ${r.daysLeft} hari lagi deadline LHP ${r.month.name}`}
            </div>
            <div style={{ fontSize: 11, color: COLORS.muted, marginTop: 2 }}>
              {r.unchecked.length} dari {r.total} item belum diselesaikan
            </div>
          </div>
          <button onClick={onGoToTracker} style={{
            background: "rgba(255,255,255,0.1)", border: "1px solid rgba(255,255,255,0.15)",
            borderRadius: 6, padding: "5px 12px", fontSize: 11, fontWeight: 700,
            color: COLORS.text, cursor: "pointer", fontFamily: "inherit", whiteSpace: "nowrap",
          }}>Lihat →</button>
        </div>
      ))}
      {uncheckedThisMonth.length > 0 && reminders.length === 0 && (
        <div style={{ background: COLORS.accentDim, border: "1px solid rgba(79,142,247,0.25)", borderRadius: 10, padding: "10px 14px" }}>
          <div style={{ fontSize: 13, fontWeight: 700, color: COLORS.accent }}>
            📋 {uncheckedThisMonth.length} kegiatan LHP belum dilakukan bulan ini
          </div>
          <div style={{ fontSize: 11, color: COLORS.muted, marginTop: 4, lineHeight: 1.5 }}>
            {uncheckedThisMonth.slice(0, 3).map((it) => it.nama).join(" · ")}
            {uncheckedThisMonth.length > 3 && ` · +${uncheckedThisMonth.length - 3} lagi`}
          </div>
        </div>
      )}
    </div>
  );
}

// ─── ROOT APP ─────────────────────────────────────────────────────────────────

export default function App() {
  const [tab, setTab] = useState("tracker");
  const [checked, setChecked] = useState({});

  useNotificationPermission();
  useDeadlineNotifications(checked);

  useEffect(() => {
    try {
      const saved = localStorage.getItem("nsp_checked_v2");
      if (saved) setChecked(JSON.parse(saved));
    } catch (_) {}
  }, []);

  const handleToggle = useCallback((monthKey, itemId, val) => {
    setChecked((prev) => {
      const next = { ...prev, [monthKey]: { ...prev[monthKey], [itemId]: val } };
      try { localStorage.setItem("nsp_checked_v2", JSON.stringify(next)); } catch (_) {}
      return next;
    });
  }, []);

  const tabs = [
    { id: "tracker", label: "Tracker Bulanan" },
    { id: "nsp", label: "Cek NSP" },
    { id: "form", label: "Form LHP" },
  ];

  return (
    <div style={S.app}>
      <div style={S.header}>
        <div style={S.headerSub}>RESIMEN KORPS TARUNA DAN SISWA · DETASEMEN TK III/60/BD</div>
        <div style={S.headerTitle}>Tracker NSP &amp; LHP — {TARUNA.nama}</div>
      </div>

      <div style={S.tabs}>
        {tabs.map((t) => (
          <button key={t.id} style={S.tab(tab === t.id)} onClick={() => setTab(t.id)}>
            {t.label}
          </button>
        ))}
      </div>

      <div style={S.content}>
        <ReminderBanner checked={checked} onGoToTracker={() => setTab("tracker")} />
        {tab === "tracker" && <TrackerPage checked={checked} onToggle={handleToggle} />}
        {tab === "nsp" && <NSPPage checked={checked} />}
        {tab === "form" && <FormPage />}
      </div>
    </div>
  );
}

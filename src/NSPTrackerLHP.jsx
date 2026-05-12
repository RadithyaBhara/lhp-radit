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
  tk: "II",
  dantontar: "IMAM SANTOSO, S.H.",
  pangkat_dantontar: "IPDA",
  nrp_dantontar: "82020196",
  dankitar: "AGUNG BUDIMAN, S.Tr.K., S.I.K.",
  pangkat_dankitar: "AKP",
  nrp_dankitar: "92040556",
};

const NSP_ITEMS = [
  {
    id: "b1a1", cat: "B1 – Keimanan & Ketaqwaan",
    nama: "Sholat fardhu berjamaah di Masjid Akpol (maks per bulan)",
    poin: 0.4, pemberi: "PAWASYON", lhp: false,
    keterangan: "Otomatis dari kegiatan harian, tidak perlu LHP terpisah",
  },
  {
    id: "b1a2", cat: "B1 – Keimanan & Ketaqwaan",
    nama: "Ibadah sunnah/tahajud ≥4× sebulan",
    poin: 0.4, pemberi: "PAWASYON", lhp: false,
    keterangan: "",
  },
  {
    id: "b1b1", cat: "B1 – Keimanan & Ketaqwaan",
    nama: "Aktif peserta majelis ilmu keagamaan/pengajian (≥4×)",
    poin: 0.44, pemberi: "PAWASYON", lhp: true,
    keterangan: "Dokumentasi peserta min 10 orang, foto bersama pemimpin ibadah, resume materi",
  },
  {
    id: "b1b2", cat: "B1 – Keimanan & Ketaqwaan",
    nama: "Menjadi Imam/mentor/pemimpin ibadah dalam kegiatan agama (≥4×)",
    poin: 0.44, pemberi: "PAWASYON", lhp: true,
    keterangan: "LHP dan dokumentasi terlampir (kegiatan di dalam Akpol), PDL boleh",
  },
  {
    id: "b1b3", cat: "B1 – Keimanan & Ketaqwaan",
    nama: "Mengajak Taruna junior ibadah di luar Akpol (pesiar/IBL/cuti)",
    poin: 0.27, pemberi: "DANKITAR", lhp: true,
    keterangan: "Kegiatan di luar Akpol, sertakan laporan dan dokumentasi",
  },
  {
    id: "b1c2", cat: "B1 – Keimanan & Ketaqwaan",
    nama: "Terlibat kepanitiaan/pengurus hari besar keagamaan (≥1×)",
    poin: 0.44, pemberi: "DANKITAR", lhp: true,
    keterangan: "Jika di luar Akpol: surat dari pengurus tempat ibadah. Foto, LHP, SPRINT wajib",
  },
  {
    id: "b2a4", cat: "B2 – Cinta Tanah Air",
    nama: "Mempelajari kesenian/budaya daerah lain (≥1×)",
    poin: 0.3, pemberi: "DANTONTAR", lhp: true,
    keterangan: "Sertakan foto dan resume minimal 1 halaman terkait budaya/kesenian yang dipelajari",
  },
  {
    id: "b2a5", cat: "B2 – Cinta Tanah Air",
    nama: "Melaporkan hal yang membahayakan/merugikan negara (≥1×)",
    poin: 0.3, pemberi: "DANKITAR", lhp: true,
    keterangan: "Uraian LHP harus jelas, perihal bahaya/kerugian negara, disertai dokumentasi",
  },
  {
    id: "b2b1", cat: "B2 – Cinta Tanah Air",
    nama: "Artikel/tulisan bertema nasionalisme (1×)",
    poin: 0.75, pemberi: "DANKITAR", lhp: true,
    keterangan: "Artikel narasi bertema nasionalisme, diketik, lampirkan turnitin maks 20%",
  },
  {
    id: "b2b2", cat: "B2 – Cinta Tanah Air",
    nama: "Menampilkan kesenian/budaya dari daerah lain (1×)",
    poin: 0.75, pemberi: "DANTONTAR", lhp: true,
    keterangan: "Sertakan foto penampilan pada acara/event",
  },
  {
    id: "b3a2", cat: "B3 – Demokrasi",
    nama: "Aktif menyampaikan pendapat di forum resmi (≥1×)",
    poin: 0.33, pemberi: "PAWASYON", lhp: true,
    keterangan: "Dokumentasi saat menyampaikan pendapat atau foto buku saku pujian",
  },
  {
    id: "b5a1", cat: "B5 – Kerja Keras & Cerdas",
    nama: "Belajar mandiri di luar jadwal + resume (≥2×)",
    poin: 0.4, pemberi: "PAWASYON", lhp: true,
    keterangan: "Dokumentasi saat belajar, resume minimal 2 halaman folio",
  },
  {
    id: "b5a2", cat: "B5 – Kerja Keras & Cerdas",
    nama: "Belajar mandiri di perpustakaan/beli buku + resume (≥2×)",
    poin: 0.4, pemberi: "PAWASYON", lhp: true,
    keterangan: "Dokumentasi buku, diserahkan ke batalyon untuk perpustakaan, resume min 2 hal folio",
  },
  {
    id: "b5a3", cat: "B5 – Kerja Keras & Cerdas",
    nama: "Kelompok belajar mandiri/penelitian/diskusi (≥2×)",
    poin: 0.4, pemberi: "PAWASYON", lhp: true,
    keterangan: "Ada SPRIN tim/kelompok belajar",
  },
  {
    id: "b5a4", cat: "B5 – Kerja Keras & Cerdas",
    nama: "Membantu giat operasi kepolisian (gatur lalin, TPTKP, dll)",
    poin: 0.4, pemberi: "DANTONTAR", lhp: true,
    keterangan: "Dokumentasi didampingi personil Polri, uraian LHP lengkap: lokasi, waktu, tempat, nama personil, bentuk kegiatan",
  },
  {
    id: "b5a5", cat: "B5 – Kerja Keras & Cerdas",
    nama: "Olahraga mandiri + lapor pawas ATAU turun BB min 2 kg",
    poin: 0.4, pemberi: "DANTONTAR", lhp: false,
    keterangan: "Poin untuk Taruna BB ideal atau OW yang berhasil turun 2 kg",
  },
  {
    id: "b6a2", cat: "B6 – Profesional",
    nama: "Mengunjungi kesatuan Polri (cuti/pesiar) + laporan teknis (1×)",
    poin: 0.33, pemberi: "DANKITAR", lhp: true,
    keterangan: "LHP berisi dokumentasi, waktu, tempat, pejabat yang ditemui, produk tertulis situasi lapangan sesuai materi kuliah",
  },
  {
    id: "b8a1", cat: "B8 – Empati",
    nama: "Menjenguk Taruna/antap Akpol sakit di KSA/RS (≥1×)",
    poin: 0.67, pemberi: "PAWASYON", lhp: true,
    keterangan: "Dokumentasi bersama yang sakit. Uraian LHP: waktu, tempat, nama Taruna/antap yang sakit",
  },
  {
    id: "b8a2", cat: "B8 – Empati",
    nama: "Baksos/donor darah atas inisiatif sendiri (≥1×)",
    poin: 0.67, pemberi: "DANTONTAR", lhp: true,
    keterangan: "Dokumentasi saat penyerahan baksos. Paket baksos harus ada tulisan BAKSOS (NAMA BATALYON)",
  },
  {
    id: "b8a3", cat: "B8 – Empati",
    nama: "Berkunjung ke rumah pejabat Akpol min Kombes Pol (≥1×)",
    poin: 0.67, pemberi: "DANTONTAR", lhp: true,
    keterangan: "Dokumentasi terlampir, di luar jam pengasuhan",
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
              para([B("DETASEMEN TARUNA TK II/59/BD", 22)]),
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
    `PADA HARI ${hari} TANGGAL ${tglFmtUpper}, SAYA MUHAMMAD RADHITYA BHARA KUNTARA TARUNA AKPOL/PANGKAT NO.AK 200507014166 TK. II/59/BD TELAH MELAKSANAKAN KEGIATAN ${namaKegiatan} DI ${tempatUp}. ${uraianUp}`;
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
          new TableCell({ borders: noBorder, width: { size: sigColW, type: WidthType.DXA }, margins: { top: 40, bottom: 0, left: 0, right: 0 }, children: [para([R("Mengetahui,")]), para([R("DANTONTAR 2B TARUNA TK II/59/BD")])] }),
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
          para([B("LAPORAN HASIL PELAKSANAAN KEGIATAN TARUNA TK. II/59/BD", 24)], { alignment: AlignmentType.CENTER }),
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
          para([R("DANKITAR B TARUNA TK II/59/BD")]),
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

  const uraianFullText = `PADA HARI ${hari} TANGGAL ${tglFmtUpper}, SAYA MUHAMMAD RADHITYA BHARA KUNTARA TARUNA AKPOL/PANGKAT NO.AK 200507014166 TK. II/59/BD TELAH MELAKSANAKAN KEGIATAN ${namaKegiatan} DI ${tempatUp}. ${uraianUp}`;

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
          <div style={pv.satuanLine}>DETASEMEN TARUNA TK II/59/BD</div>
        </div>
      </div>

      {/* Garis lurus */}
      <div style={pv.divider} />

      {/* Judul */}
      <div style={pv.judulCenter}>
        LAPORAN HASIL PELAKSANAAN KEGIATAN TARUNA TK. II/59/BD
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
          <div>DANTONTAR 2B TARUNA TK II/59/BD</div>
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
        <div>DANKITAR B TARUNA TK II/59/BD</div>
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


// ─── TOTAL NSP PAGE ──────────────────────────────────────────────────────────

// ─── TEGURAN TYPES ───────────────────────────────────────────────────────────

const TEGURAN_TYPES = [
  { id: "tl", nama: "Teguran Lisan", pengurangan: 0.5 },
  { id: "tt", nama: "Teguran Tertulis", pengurangan: 1.0 },
  { id: "sp1", nama: "Surat Peringatan 1 (SP1)", pengurangan: 2.0 },
  { id: "sp2", nama: "Surat Peringatan 2 (SP2)", pengurangan: 3.0 },
  { id: "sp3", nama: "Surat Peringatan 3 (SP3)", pengurangan: 5.0 },
  { id: "hd", nama: "Hukuman Disiplin", pengurangan: 7.0 },
  { id: "custom", nama: "Pengurangan Lain", pengurangan: 0 },
];

// ─── TOTAL NSP PAGE ──────────────────────────────────────────────────────────

function TotalNSPPage({ checked, monthExtras, onUpdateExtras }) {
  const lhpItems = NSP_ITEMS.filter((x) => x.lhp);
  const nonLhpItems = NSP_ITEMS.filter((x) => !x.lhp);
  const lhpMaxPoin = lhpItems.reduce((s, x) => s + x.poin, 0);
  const nonLhpMaxPoin = nonLhpItems.reduce((s, x) => s + x.poin, 0);

  const NILAI_DASAR = 70;
  const NILAI_TANPA_TEGURAN = 1;
  const NILAI_SIKAP_PENGASUH_DEFAULT = 10;

  const [selMonth, setSelMonth] = useState(null); // null = overview, index = detail bulan
  const [showAddTeguran, setShowAddTeguran] = useState(false);
  const [newTeguranType, setNewTeguranType] = useState("tl");
  const [newTeguranKet, setNewTeguranKet] = useState("");
  const [newTeguranNilai, setNewTeguranNilai] = useState("");
  const [newTeguranTgl, setNewTeguranTgl] = useState("");

  // Build per-month full NSP calculation
  const monthlyFull = MONTHS.map((m, i) => {
    const mc = checked[m.key] || {};
    const ext = monthExtras[m.key] || {};
    const poinLHP = lhpItems.reduce((s, it) => s + (mc[it.id] ? it.poin : 0), 0);
    const countLHP = lhpItems.filter((it) => mc[it.id]).length;
    const teguranList = ext.teguran || [];
    const totalPengurangan = teguranList.reduce((s, t) => s + (t.pengurangan || 0), 0);
    const hasTeguran = teguranList.length > 0;
    const nilaiPengasuh = ext.nilaiPengasuh !== undefined ? ext.nilaiPengasuh : NILAI_SIKAP_PENGASUH_DEFAULT;
    const bonusTanpaTeguran = hasTeguran ? 0 : NILAI_TANPA_TEGURAN;

    const totalBulan = NILAI_DASAR + bonusTanpaTeguran + nilaiPengasuh + poinLHP - totalPengurangan;

    return {
      month: m, idx: i, poinLHP, countLHP,
      nilaiDasar: NILAI_DASAR,
      bonusTanpaTeguran,
      nilaiPengasuh,
      teguranList,
      totalPengurangan,
      totalBulan: Math.max(0, totalBulan),
    };
  });

  // Average NSP across all months that have passed or are current
  const activeMonths = monthlyFull.filter((md, i) => i <= todayDate.getMonth());
  const avgNSP = activeMonths.length > 0 ? activeMonths.reduce((s, md) => s + md.totalBulan, 0) / activeMonths.length : 0;

  // Handlers
  const handleNilaiPengasuh = (monthKey, val) => {
    const num = parseFloat(val);
    if (isNaN(num) && val !== "") return;
    onUpdateExtras(monthKey, { nilaiPengasuh: val === "" ? NILAI_SIKAP_PENGASUH_DEFAULT : num });
  };

  const handleAddTeguran = (monthKey) => {
    const type = TEGURAN_TYPES.find((t) => t.id === newTeguranType);
    const pengurangan = newTeguranType === "custom" ? (parseFloat(newTeguranNilai) || 0) : type.pengurangan;
    const teguran = {
      id: Date.now().toString(36),
      type: type.nama,
      pengurangan,
      keterangan: newTeguranKet,
      tanggal: newTeguranTgl,
    };
    const ext = monthExtras[monthKey] || {};
    const list = [...(ext.teguran || []), teguran];
    onUpdateExtras(monthKey, { teguran: list });
    setNewTeguranKet("");
    setNewTeguranNilai("");
    setNewTeguranTgl("");
    setShowAddTeguran(false);
  };

  const handleRemoveTeguran = (monthKey, teguranId) => {
    const ext = monthExtras[monthKey] || {};
    const list = (ext.teguran || []).filter((t) => t.id !== teguranId);
    onUpdateExtras(monthKey, { teguran: list });
  };

  // ── OVERVIEW MODE ──
  if (selMonth === null) {
    const maxTotal = Math.max(...monthlyFull.map((d) => d.totalBulan), 1);

    return (
      <div>
        {/* Summary cards */}
        <div style={{ display: "grid", gridTemplateColumns: "1fr 1fr 1fr", gap: 8, marginBottom: 20 }}>
          <div style={S.statCard}>
            <div style={S.statLabel}>RATA-RATA NSP</div>
            <div style={S.statValue}>{avgNSP.toFixed(2)}</div>
            <div style={{ fontSize: 11, color: avgNSP >= 75 ? COLORS.success : avgNSP >= 70 ? COLORS.warning : COLORS.danger, marginTop: 4, fontWeight: 600 }}>
              {avgNSP >= 80 ? "Sangat Baik" : avgNSP >= 75 ? "Baik" : avgNSP >= 70 ? "Cukup" : "Perlu Perhatian"}
            </div>
          </div>
          <div style={S.statCard}>
            <div style={S.statLabel}>NILAI DASAR / BULAN</div>
            <div style={S.statValue}>{NILAI_DASAR}<span style={S.statUnit}>poin</span></div>
            <div style={{ fontSize: 11, color: COLORS.muted, marginTop: 4 }}>+ 1 jika tanpa teguran</div>
          </div>
          <div style={S.statCard}>
            <div style={S.statLabel}>TOTAL POIN LHP (SEMUA BULAN)</div>
            <div style={S.statValue}>{monthlyFull.reduce((s, md) => s + md.poinLHP, 0).toFixed(2)}<span style={S.statUnit}>poin</span></div>
          </div>
        </div>

        {/* Formula card */}
        <div style={{ ...S.formSection, marginBottom: 16 }}>
          <div style={S.formTitle}>Rumus NSP per Bulan</div>
          <div style={{ fontSize: 12, color: COLORS.text, lineHeight: 1.8, fontFamily: "'IBM Plex Mono', monospace" }}>
            NSP = <span style={{ color: COLORS.accent }}>70</span> (dasar)
            + <span style={{ color: COLORS.success }}>1</span> (tanpa teguran)
            + <span style={{ color: "#a78bfa" }}>Nilai Pengasuh</span> (maks 10)
            + <span style={{ color: COLORS.accent }}>Poin LHP</span>
            − <span style={{ color: COLORS.danger }}>Teguran</span>
          </div>
        </div>

        {/* Monthly bar chart */}
        <div style={S.formSection}>
          <div style={S.formTitle}>NSP per Bulan <span style={{ fontSize: 11, color: COLORS.muted, fontWeight: 400 }}>— klik bulan untuk detail & input teguran</span></div>
          <div style={{ display: "flex", gap: 4, alignItems: "flex-end", height: 180, paddingBottom: 28, position: "relative" }}>
            {/* Reference line at 70 */}
            <div style={{
              position: "absolute", left: 0, right: 0,
              bottom: 28 + ((NILAI_DASAR / maxTotal) * 150),
              height: 1, borderTop: "1px dashed rgba(251,191,36,0.4)",
            }} />
            {monthlyFull.map((md, i) => {
              const barH = md.totalBulan > 0 ? Math.max(4, (md.totalBulan / maxTotal) * 150) : 2;
              const isNow = i === todayDate.getMonth();
              const hasTeg = md.teguranList.length > 0;
              return (
                <div key={md.month.key} style={{ flex: 1, display: "flex", flexDirection: "column", alignItems: "center", gap: 2, cursor: "pointer" }} onClick={() => setSelMonth(i)}>
                  <div style={{ fontSize: 9, fontWeight: 700, fontFamily: "'IBM Plex Mono', monospace", color: hasTeg ? COLORS.danger : isNow ? COLORS.accent : COLORS.text }}>
                    {md.totalBulan.toFixed(1)}
                  </div>
                  <div style={{
                    width: "65%", height: barH,
                    background: hasTeg
                      ? `linear-gradient(180deg, ${COLORS.danger}, rgba(248,113,113,0.3))`
                      : isNow
                      ? `linear-gradient(180deg, ${COLORS.accent}, rgba(79,142,247,0.3))`
                      : `linear-gradient(180deg, rgba(79,142,247,0.6), rgba(79,142,247,0.15))`,
                    borderRadius: "4px 4px 0 0",
                    border: isNow ? `1px solid ${COLORS.accent}` : hasTeg ? `1px solid ${COLORS.danger}` : "none",
                    transition: "height 0.3s",
                  }} />
                  <div style={{ fontSize: 9, color: isNow ? COLORS.accent : COLORS.muted, fontWeight: isNow ? 700 : 400 }}>
                    {md.month.name.slice(0, 3)}
                  </div>
                  {hasTeg && <div style={{ fontSize: 8, color: COLORS.danger }}>⚠{md.teguranList.length}</div>}
                </div>
              );
            })}
          </div>
        </div>

        {/* Monthly table */}
        <div style={S.formSection}>
          <div style={S.formTitle}>Detail NSP Bulanan</div>
          <div style={{ overflowX: "auto" }}>
            <table style={{ width: "100%", borderCollapse: "collapse", fontSize: 11 }}>
              <thead>
                <tr>
                  {["BULAN", "DASAR", "TEGURAN?", "PENGASUH", "LHP", "TEGURAN", "TOTAL"].map((h, i) => (
                    <th key={h} style={{ textAlign: i === 0 ? "left" : "center", padding: "6px 6px", borderBottom: `1px solid ${COLORS.border}`, color: COLORS.muted, fontWeight: 600, fontSize: 10, whiteSpace: "nowrap" }}>{h}</th>
                  ))}
                </tr>
              </thead>
              <tbody>
                {monthlyFull.map((md, i) => {
                  const isNow = i === todayDate.getMonth();
                  const f = "'IBM Plex Mono', monospace";
                  return (
                    <tr key={md.month.key} style={{ background: isNow ? COLORS.accentDim : "transparent", cursor: "pointer" }} onClick={() => setSelMonth(i)}>
                      <td style={{ padding: "6px 6px", borderBottom: `1px solid ${COLORS.border}`, fontWeight: isNow ? 700 : 400, color: isNow ? COLORS.accent : COLORS.text }}>{md.month.name.slice(0, 3)}</td>
                      <td style={{ padding: "6px 6px", borderBottom: `1px solid ${COLORS.border}`, textAlign: "center", fontFamily: f, color: COLORS.muted }}>{NILAI_DASAR}</td>
                      <td style={{ padding: "6px 6px", borderBottom: `1px solid ${COLORS.border}`, textAlign: "center", fontFamily: f, color: md.bonusTanpaTeguran > 0 ? COLORS.success : COLORS.danger }}>
                        {md.bonusTanpaTeguran > 0 ? "+1 ✓" : "0 ✗"}
                      </td>
                      <td style={{ padding: "6px 6px", borderBottom: `1px solid ${COLORS.border}`, textAlign: "center", fontFamily: f, color: "#a78bfa" }}>+{md.nilaiPengasuh}</td>
                      <td style={{ padding: "6px 6px", borderBottom: `1px solid ${COLORS.border}`, textAlign: "center", fontFamily: f, color: COLORS.accent }}>
                        {md.poinLHP > 0 ? `+${md.poinLHP.toFixed(2)}` : "—"}
                      </td>
                      <td style={{ padding: "6px 6px", borderBottom: `1px solid ${COLORS.border}`, textAlign: "center", fontFamily: f, color: md.totalPengurangan > 0 ? COLORS.danger : COLORS.muted }}>
                        {md.totalPengurangan > 0 ? `−${md.totalPengurangan.toFixed(1)}` : "—"}
                      </td>
                      <td style={{ padding: "6px 6px", borderBottom: `1px solid ${COLORS.border}`, textAlign: "center", fontFamily: f, fontWeight: 700, color: md.totalBulan >= 75 ? COLORS.success : md.totalBulan >= 70 ? COLORS.warning : COLORS.danger }}>
                        {md.totalBulan.toFixed(2)}
                      </td>
                    </tr>
                  );
                })}
              </tbody>
            </table>
          </div>
        </div>
      </div>
    );
  }

  // ── DETAIL BULAN MODE ──
  const md = monthlyFull[selMonth];
  const mk = md.month.key;

  return (
    <div>
      {/* Back */}
      <div style={{ display: "flex", alignItems: "center", gap: 10, marginBottom: 16 }}>
        <button onClick={() => { setSelMonth(null); setShowAddTeguran(false); }} style={{
          background: "none", border: `1px solid ${COLORS.border}`, borderRadius: 6,
          padding: "4px 10px", cursor: "pointer", color: COLORS.muted, fontFamily: "inherit", fontSize: 12, fontWeight: 600,
        }}>← Kembali</button>
        <div>
          <div style={{ fontSize: 16, fontWeight: 700, color: COLORS.text }}>{md.month.name} {md.month.year}</div>
          <div style={{ fontSize: 11, color: COLORS.muted }}>Detail NSP Bulan Ini</div>
        </div>
      </div>

      {/* Score breakdown */}
      <div style={S.formSection}>
        <div style={S.formTitle}>Perhitungan NSP</div>
        <div style={{ display: "flex", flexDirection: "column", gap: 6 }}>
          {[
            { label: "Nilai Dasar", value: `+${NILAI_DASAR}`, color: COLORS.text },
            { label: `Bonus Tanpa Teguran (${md.teguranList.length === 0 ? "tidak ada teguran" : "ada teguran"})`, value: md.bonusTanpaTeguran > 0 ? "+1" : "0", color: md.bonusTanpaTeguran > 0 ? COLORS.success : COLORS.danger },
            { label: "Nilai Sikap Pengasuh", value: `+${md.nilaiPengasuh}`, color: "#a78bfa" },
            { label: "Poin LHP", value: `+${md.poinLHP.toFixed(2)}`, color: COLORS.accent },
            ...(md.totalPengurangan > 0 ? [{ label: "Pengurangan Teguran", value: `−${md.totalPengurangan.toFixed(1)}`, color: COLORS.danger }] : []),
          ].map((row) => (
            <div key={row.label} style={{ display: "flex", justifyContent: "space-between", alignItems: "center", padding: "6px 0", borderBottom: `1px solid ${COLORS.border}` }}>
              <span style={{ fontSize: 12, color: COLORS.muted }}>{row.label}</span>
              <span style={{ fontSize: 14, fontWeight: 700, fontFamily: "'IBM Plex Mono', monospace", color: row.color }}>{row.value}</span>
            </div>
          ))}
          <div style={{ display: "flex", justifyContent: "space-between", alignItems: "center", padding: "8px 0" }}>
            <span style={{ fontSize: 14, fontWeight: 700, color: COLORS.text }}>TOTAL NSP</span>
            <span style={{ fontSize: 22, fontWeight: 700, fontFamily: "'IBM Plex Mono', monospace", color: md.totalBulan >= 75 ? COLORS.success : md.totalBulan >= 70 ? COLORS.warning : COLORS.danger }}>
              {md.totalBulan.toFixed(2)}
            </span>
          </div>
        </div>
      </div>

      {/* Nilai Sikap Pengasuh input */}
      <div style={S.formSection}>
        <div style={S.formTitle}>Nilai Sikap Tambahan Pengasuh</div>
        <div style={{ fontSize: 11, color: COLORS.muted, marginBottom: 8 }}>
          Nilai dari pengasuh (0–10). Default: 10 jika tidak diubah.
        </div>
        <input
          style={{ ...S.input, maxWidth: 200 }}
          type="number"
          min="0" max="10" step="0.5"
          value={md.nilaiPengasuh}
          onChange={(e) => handleNilaiPengasuh(mk, e.target.value)}
          placeholder="10"
        />
      </div>

      {/* Teguran section */}
      <div style={S.formSection}>
        <div style={{ display: "flex", justifyContent: "space-between", alignItems: "center", marginBottom: 12 }}>
          <div style={S.formTitle}>Teguran / Pelanggaran <span style={{ fontSize: 11, color: COLORS.danger, fontWeight: 400 }}>(mengurangi NSP)</span></div>
          <button onClick={() => setShowAddTeguran(!showAddTeguran)} style={{
            background: showAddTeguran ? COLORS.dangerDim : "transparent",
            border: `1px solid ${showAddTeguran ? COLORS.danger : COLORS.border}`,
            borderRadius: 6, padding: "4px 12px", fontSize: 11, fontWeight: 700,
            cursor: "pointer", color: showAddTeguran ? COLORS.danger : COLORS.muted, fontFamily: "inherit",
          }}>
            {showAddTeguran ? "✕ Batal" : "+ Tambah Teguran"}
          </button>
        </div>

        {/* Add teguran form */}
        {showAddTeguran && (
          <div style={{ background: COLORS.bg, border: `1px solid ${COLORS.border}`, borderRadius: 8, padding: 12, marginBottom: 12 }}>
            <div style={{ marginBottom: 8 }}>
              <label style={S.formLabel}>JENIS TEGURAN</label>
              <select style={S.select} value={newTeguranType} onChange={(e) => setNewTeguranType(e.target.value)}>
                {TEGURAN_TYPES.map((t) => (
                  <option key={t.id} value={t.id}>{t.nama} {t.id !== "custom" ? `(−${t.pengurangan} poin)` : ""}</option>
                ))}
              </select>
            </div>
            {newTeguranType === "custom" && (
              <div style={{ marginBottom: 8 }}>
                <label style={S.formLabel}>NILAI PENGURANGAN</label>
                <input style={{ ...S.input, maxWidth: 200 }} type="number" min="0" step="0.5" value={newTeguranNilai} onChange={(e) => setNewTeguranNilai(e.target.value)} placeholder="0.5" />
              </div>
            )}
            <div style={{ marginBottom: 8 }}>
              <label style={S.formLabel}>TANGGAL (opsional)</label>
              <input style={{ ...S.input, maxWidth: 200 }} type="date" value={newTeguranTgl} onChange={(e) => setNewTeguranTgl(e.target.value)} />
            </div>
            <div style={{ marginBottom: 8 }}>
              <label style={S.formLabel}>KETERANGAN (opsional)</label>
              <input style={S.input} type="text" placeholder="contoh: Terlambat apel pagi" value={newTeguranKet} onChange={(e) => setNewTeguranKet(e.target.value)} />
            </div>
            <button onClick={() => handleAddTeguran(mk)} style={{ ...S.btnPrimary, background: COLORS.danger, maxWidth: 200 }}>
              Simpan Teguran
            </button>
          </div>
        )}

        {/* List teguran */}
        {md.teguranList.length === 0 ? (
          <div style={{ textAlign: "center", padding: "16px 0", color: COLORS.success, fontSize: 13, fontWeight: 600 }}>
            ✅ Tidak ada teguran bulan ini — +1 poin bonus!
          </div>
        ) : (
          <div style={{ display: "flex", flexDirection: "column", gap: 6 }}>
            {md.teguranList.map((t) => (
              <div key={t.id} style={{
                background: COLORS.dangerDim, border: `1px solid rgba(248,113,113,0.2)`,
                borderRadius: 8, padding: "8px 12px",
                display: "flex", justifyContent: "space-between", alignItems: "center",
              }}>
                <div>
                  <div style={{ fontSize: 12, fontWeight: 600, color: COLORS.danger }}>{t.type}</div>
                  <div style={{ fontSize: 11, color: COLORS.muted }}>
                    {t.tanggal && `${t.tanggal} · `}{t.keterangan || "Tanpa keterangan"}
                  </div>
                </div>
                <div style={{ display: "flex", alignItems: "center", gap: 8 }}>
                  <span style={{ fontSize: 14, fontWeight: 700, fontFamily: "'IBM Plex Mono', monospace", color: COLORS.danger }}>
                    −{t.pengurangan.toFixed(1)}
                  </span>
                  <button onClick={() => handleRemoveTeguran(mk, t.id)} style={{
                    background: "none", border: "none", color: COLORS.muted, cursor: "pointer",
                    fontSize: 16, lineHeight: 1, fontFamily: "inherit",
                  }} title="Hapus teguran">✕</button>
                </div>
              </div>
            ))}
          </div>
        )}
      </div>
    </div>
  );
}

// ─── ASISTEN LHP PAGE ─────────────────────────────────────────────────────────

const ASISTEN_MODES = [
  {
    id: "nasionalisme",
    icon: "🇮🇩",
    title: "Artikel Nasionalisme",
    subtitle: "B2 – Cinta Tanah Air · 0.75 poin",
    description: "Generate artikel narasi bertema nasionalisme (untuk poin NSP B2b1). Turnitin maks 20%.",
    fields: [
      { id: "topik", label: "Topik / Tema Artikel", placeholder: "contoh: Peran generasi muda dalam menjaga persatuan bangsa di era digital", type: "text" },
      { id: "sudut", label: "Sudut Pandang (opsional)", placeholder: "contoh: Dari perspektif Taruna Akpol sebagai calon perwira Polri", type: "text" },
      { id: "panjang", label: "Panjang Artikel", placeholder: "", type: "select", options: ["800-1000 kata", "1000-1500 kata", "1500-2000 kata"] },
    ],
    buildPrompt: (vals) =>
      `Kamu adalah seorang Taruna Akademi Kepolisian (Akpol) Tingkat II. Tulis sebuah artikel narasi bertema nasionalisme dengan topik: "${vals.topik}".
${vals.sudut ? `Sudut pandang: ${vals.sudut}.` : "Tulis dari sudut pandang Taruna Akpol sebagai calon perwira Polri."}
Panjang artikel: ${vals.panjang || "1000-1500 kata"}.

Ketentuan:
- Artikel harus ORISINIL, bukan copy-paste, agar lolos cek turnitin (maks 20% similarity).
- Gunakan bahasa Indonesia formal yang baik dan benar.
- Struktur: Judul, Pendahuluan, Isi (2-3 sub-bagian), Penutup/Kesimpulan.
- Sertakan referensi atau contoh nyata yang relevan.
- Kaitkan dengan nilai-nilai kepolisian: Tribrata, Catur Prasetya, atau Pancasila.
- Tulis langsung artikelnya tanpa komentar tambahan.`,
  },
  {
    id: "bahaya_negara",
    icon: "⚠️",
    title: "Laporan Bahaya/Kerugian Negara",
    subtitle: "B2 – Cinta Tanah Air · 0.30 poin",
    description: "Generate uraian LHP tentang hal yang membahayakan/merugikan negara (untuk poin NSP B2a5).",
    fields: [
      { id: "kejadian", label: "Deskripsi Kejadian / Hal yang Ditemukan", placeholder: "contoh: Menemukan konten hoaks tentang pemilu di media sosial yang berpotensi memecah belah", type: "textarea" },
      { id: "lokasi", label: "Lokasi / Platform", placeholder: "contoh: Media sosial Instagram / Wilayah sekitar Akpol", type: "text" },
      { id: "waktu", label: "Waktu Kejadian", placeholder: "contoh: Sabtu, 10 Mei 2026 pukul 14.00 WIB", type: "text" },
      { id: "tindakan", label: "Tindakan yang Dilakukan (opsional)", placeholder: "contoh: Melaporkan ke platform dan mengedukasi rekan taruna", type: "text" },
    ],
    buildPrompt: (vals) =>
      `Kamu adalah seorang Taruna Akademi Kepolisian (Akpol) Tingkat II bernama MUHAMMAD RADHITYA BHARA KUNTARA, NO.AK 200507014166.

Buatkan uraian Laporan Hasil Pelaksanaan (LHP) tentang penemuan hal yang membahayakan atau merugikan negara.

Detail kejadian:
- Kejadian: ${vals.kejadian}
- Lokasi/Platform: ${vals.lokasi || "(tidak disebutkan)"}
- Waktu: ${vals.waktu || "(tidak disebutkan)"}
${vals.tindakan ? `- Tindakan yang dilakukan: ${vals.tindakan}` : ""}

Ketentuan penulisan:
- Tulis dalam format uraian LHP formal, menggunakan bahasa Indonesia baku.
- Jelaskan kronologi secara rinci: apa yang ditemukan, kapan, di mana, bagaimana, dampak potensial.
- Jelaskan mengapa hal ini membahayakan/merugikan negara (kaitkan dengan ketahanan nasional, NKRI, Pancasila).
- Jelaskan tindakan atau langkah yang dilakukan/direkomendasikan.
- Panjang: 300-500 kata.
- Tulis langsung uraiannya tanpa komentar tambahan, tanpa judul "URAIAN KEGIATAN" di atas.`,
  },
  {
    id: "kunjungan_satuan",
    icon: "🏛️",
    title: "Laporan Kunjungan Kesatuan Polri",
    subtitle: "B6 – Profesional · 0.33 poin",
    description: "Generate laporan teknis kunjungan ke kesatuan Polri saat cuti/pesiar (untuk poin NSP B6a2).",
    fields: [
      { id: "satuan", label: "Nama Kesatuan yang Dikunjungi", placeholder: "contoh: Polres Surabaya / Polda Jawa Timur / Polsek Gayungan", type: "text" },
      { id: "tanggal_kunjungan", label: "Tanggal Kunjungan", placeholder: "contoh: 5 Mei 2026", type: "text" },
      { id: "pejabat", label: "Pejabat yang Ditemui", placeholder: "contoh: Kapolres AKBP Budi Santoso, S.I.K., M.H.", type: "text" },
      { id: "kegiatan", label: "Kegiatan / Hal yang Diamati", placeholder: "contoh: Mengamati proses penyidikan tindak pidana pencurian, melihat fasilitas SPKT, berdiskusi tentang pemolisian masyarakat", type: "textarea" },
      { id: "materi_kuliah", label: "Keterkaitan dengan Materi Kuliah", placeholder: "contoh: Hukum Acara Pidana, Manajemen Operasional Kepolisian", type: "text" },
    ],
    buildPrompt: (vals) =>
      `Kamu adalah seorang Taruna Akademi Kepolisian (Akpol) Tingkat II bernama MUHAMMAD RADHITYA BHARA KUNTARA, NO.AK 200507014166.

Buatkan laporan teknis kunjungan ke kesatuan Polri untuk Laporan Hasil Pelaksanaan (LHP).

Detail kunjungan:
- Kesatuan: ${vals.satuan}
- Tanggal: ${vals.tanggal_kunjungan || "(tidak disebutkan)"}
- Pejabat yang ditemui: ${vals.pejabat || "(tidak disebutkan)"}
- Kegiatan/hal yang diamati: ${vals.kegiatan}
- Keterkaitan materi kuliah: ${vals.materi_kuliah || "(tidak disebutkan)"}

Ketentuan penulisan:
- Format: laporan teknis kunjungan, bahasa Indonesia formal baku.
- Struktur: Pendahuluan (tujuan kunjungan), Pelaksanaan (kronologi, apa yang dilihat/dipelajari, diskusi dengan pejabat), Analisis (keterkaitan dengan materi kuliah yang dipelajari di Akpol), Kesimpulan dan Saran.
- Sertakan produk tertulis tentang situasi lapangan di kesatuan tersebut sesuai materi kuliah.
- Panjang: 500-800 kata.
- Tulis langsung laporannya tanpa komentar tambahan.`,
  },
];

function AssistantPage() {
  const [activeMode, setActiveMode] = useState(null);
  const [fieldValues, setFieldValues] = useState({});
  const [result, setResult] = useState("");
  const [loading, setLoading] = useState(false);
  const [error, setError] = useState("");
  const [copied, setCopied] = useState(false);

  const handleFieldChange = (fieldId, val) => {
    setFieldValues((prev) => ({ ...prev, [fieldId]: val }));
  };

  const handleGenerate = async () => {
    const mode = ASISTEN_MODES.find((m) => m.id === activeMode);
    if (!mode) return;

    // Validate required fields
    const firstField = mode.fields[0];
    if (!fieldValues[firstField.id]?.trim()) {
      setError(`Isi "${firstField.label}" terlebih dahulu.`);
      return;
    }

    setError("");
    setLoading(true);
    setResult("");
    try {
      const prompt = mode.buildPrompt(fieldValues);
      const response = await fetch("https://api.anthropic.com/v1/messages", {
        method: "POST",
        headers: { "Content-Type": "application/json" },
        body: JSON.stringify({
          model: "claude-sonnet-4-20250514",
          max_tokens: 4000,
          messages: [{ role: "user", content: prompt }],
        }),
      });
      const data = await response.json();
      const text = data.content
        ?.map((item) => (item.type === "text" ? item.text : ""))
        .filter(Boolean)
        .join("\n") || "";
      if (!text) throw new Error("Respons kosong dari AI");
      setResult(text);
    } catch (e) {
      setError("Gagal generate. Pastikan koneksi internet aktif. Error: " + e.message);
      console.error(e);
    }
    setLoading(false);
  };

  const handleCopy = () => {
    navigator.clipboard.writeText(result).then(() => {
      setCopied(true);
      setTimeout(() => setCopied(false), 2500);
    });
  };

  const handleReset = () => {
    setActiveMode(null);
    setFieldValues({});
    setResult("");
    setError("");
  };

  // Mode selection screen
  if (!activeMode) {
    return (
      <div>
        <div style={{ fontSize: 13, color: COLORS.muted, marginBottom: 16 }}>
          Pilih jenis dokumen yang ingin dibuat — AI akan bantu draft-kan berdasarkan input kamu.
        </div>
        <div style={{ display: "flex", flexDirection: "column", gap: 10 }}>
          {ASISTEN_MODES.map((mode) => (
            <div
              key={mode.id}
              onClick={() => { setActiveMode(mode.id); setFieldValues({}); setResult(""); setError(""); }}
              style={{
                background: COLORS.surface,
                border: `1px solid ${COLORS.border}`,
                borderRadius: 10,
                padding: "14px 16px",
                cursor: "pointer",
                transition: "border-color 0.15s",
                display: "flex",
                gap: 14,
                alignItems: "flex-start",
              }}
            >
              <div style={{ fontSize: 28, lineHeight: 1, flexShrink: 0, marginTop: 2 }}>{mode.icon}</div>
              <div style={{ flex: 1 }}>
                <div style={{ fontSize: 14, fontWeight: 700, color: COLORS.text, marginBottom: 2 }}>{mode.title}</div>
                <div style={{ fontSize: 11, color: COLORS.accent, fontWeight: 600, marginBottom: 4 }}>{mode.subtitle}</div>
                <div style={{ fontSize: 12, color: COLORS.muted, lineHeight: 1.4 }}>{mode.description}</div>
              </div>
              <div style={{ fontSize: 18, color: COLORS.muted, marginTop: 6 }}>→</div>
            </div>
          ))}
        </div>
      </div>
    );
  }

  // Active mode form
  const mode = ASISTEN_MODES.find((m) => m.id === activeMode);

  return (
    <div>
      {/* Back button + title */}
      <div style={{ display: "flex", alignItems: "center", gap: 10, marginBottom: 16 }}>
        <button onClick={handleReset} style={{
          background: "none", border: `1px solid ${COLORS.border}`, borderRadius: 6,
          padding: "4px 10px", cursor: "pointer", color: COLORS.muted, fontFamily: "inherit",
          fontSize: 12, fontWeight: 600,
        }}>← Kembali</button>
        <div>
          <div style={{ fontSize: 16, fontWeight: 700, color: COLORS.text }}>{mode.icon} {mode.title}</div>
          <div style={{ fontSize: 11, color: COLORS.accent, fontWeight: 600 }}>{mode.subtitle}</div>
        </div>
      </div>

      {/* Form fields */}
      <div style={S.formSection}>
        {mode.fields.map((field) => (
          <div key={field.id} style={{ marginBottom: 12 }}>
            <label style={S.formLabel}>{field.label}</label>
            {field.type === "textarea" ? (
              <textarea
                style={S.textarea}
                placeholder={field.placeholder}
                value={fieldValues[field.id] || ""}
                onChange={(e) => handleFieldChange(field.id, e.target.value)}
              />
            ) : field.type === "select" ? (
              <select
                style={S.select}
                value={fieldValues[field.id] || field.options[1] || field.options[0]}
                onChange={(e) => handleFieldChange(field.id, e.target.value)}
              >
                {field.options.map((opt) => (
                  <option key={opt} value={opt}>{opt}</option>
                ))}
              </select>
            ) : (
              <input
                style={S.input}
                type="text"
                placeholder={field.placeholder}
                value={fieldValues[field.id] || ""}
                onChange={(e) => handleFieldChange(field.id, e.target.value)}
              />
            )}
          </div>
        ))}

        {error && (
          <div style={{ ...S.banner("danger"), marginBottom: 10 }}>{error}</div>
        )}

        <button
          style={{ ...S.btnPrimary, opacity: loading ? 0.6 : 1 }}
          onClick={handleGenerate}
          disabled={loading}
        >
          {loading ? "⏳ AI sedang menulis..." : `✨ Generate ${mode.title}`}
        </button>
      </div>

      {/* Loading animation */}
      {loading && (
        <div style={{
          background: COLORS.surface, border: `1px solid ${COLORS.border}`,
          borderRadius: 10, padding: "20px", marginTop: 12, textAlign: "center",
        }}>
          <div style={{ fontSize: 13, color: COLORS.muted, marginBottom: 8 }}>
            AI sedang menulis {mode.title.toLowerCase()}...
          </div>
          <div style={{
            width: 120, height: 3, background: COLORS.border,
            borderRadius: 2, margin: "0 auto", overflow: "hidden",
          }}>
            <div style={{
              width: "40%", height: 3, background: COLORS.accent,
              borderRadius: 2, animation: "loading 1.5s ease-in-out infinite",
            }} />
          </div>
          <style>{`@keyframes loading { 0% { transform: translateX(-100%); } 100% { transform: translateX(350%); } }`}</style>
        </div>
      )}

      {/* Result */}
      {result && (
        <div style={{ marginTop: 14 }}>
          <div style={{ display: "flex", alignItems: "center", justifyContent: "space-between", marginBottom: 8 }}>
            <div style={{ fontSize: 11, color: COLORS.muted, fontWeight: 700, letterSpacing: "0.08em", textTransform: "uppercase" }}>
              Hasil — {mode.title}
            </div>
            <div style={{ display: "flex", gap: 6 }}>
              <button onClick={handleCopy} style={{
                background: COLORS.surface, border: `1px solid ${COLORS.border}`,
                borderRadius: 6, padding: "4px 12px", fontSize: 11, fontWeight: 600,
                cursor: "pointer", color: copied ? COLORS.success : COLORS.muted,
                fontFamily: "inherit",
              }}>
                {copied ? "✓ Tersalin!" : "⎘ Salin"}
              </button>
              <button onClick={handleGenerate} style={{
                background: COLORS.surface, border: `1px solid ${COLORS.border}`,
                borderRadius: 6, padding: "4px 12px", fontSize: 11, fontWeight: 600,
                cursor: "pointer", color: COLORS.accent, fontFamily: "inherit",
              }}>
                ↻ Regenerate
              </button>
            </div>
          </div>
          <div style={{
            background: "#fff", color: "#000", borderRadius: 8,
            padding: "20px 24px", fontSize: 13, lineHeight: 1.8,
            fontFamily: "Arial MT, Arial, sans-serif",
            whiteSpace: "pre-wrap", maxHeight: 600, overflowY: "auto",
            border: "1px solid #ddd",
          }}>
            {result}
          </div>

          {/* Tip */}
          <div style={{ ...S.banner("success"), marginTop: 10 }}>
            💡 Salin teks di atas → buka tab <strong>Form LHP</strong> → pilih kegiatan yang sesuai → paste ke field <strong>Uraian Kegiatan</strong> → generate DOCX.
            {mode.id === "nasionalisme" && " Jangan lupa cek Turnitin sebelum dikumpulkan (maks 20%)."}
          </div>
        </div>
      )}
    </div>
  );
}

// ─── LAMPIRAN BULANAN PAGE ────────────────────────────────────────────────────
// Gabungkan file-file .docx yang sudah jadi menjadi 1 file .docx

async function mergeDocxFiles(files, monthName, monthYear) {
  // Load JSZip
  if (!window.JSZip) {
    await new Promise((resolve, reject) => {
      const s = document.createElement("script");
      s.src = "https://cdn.jsdelivr.net/npm/jszip@3.10.1/dist/jszip.min.js";
      s.onload = resolve;
      s.onerror = reject;
      document.head.appendChild(s);
    });
  }

  const JSZip = window.JSZip;

  // ── Helpers ──
  const B = (text, size = 22, color = "000000") => new TextRun({ text, bold: true, size, font: "Arial", color });
  const R = (text, size = 22, color = "000000") => new TextRun({ text, size, font: "Arial MT", color });
  const p = (children, opts = {}) => new Paragraph({ children, spacing: { after: 0, before: 0 }, ...opts });
  const blank = () => p([R("")]);
  const pageProps = { size: { width: 11906, height: 16838 }, margin: { top: 1134, right: 1134, bottom: 1134, left: 1134 } };

  let logoData = null;
  try { logoData = await fetchLogoBase64(); } catch (_) {}

  // ── 1. Build cover page docx ──
  const coverChildren = [blank(), blank(), blank(), blank()];
  if (logoData) {
    coverChildren.push(p([new ImageRun({ type: "jpg", data: logoData, transformation: { width: 120, height: 136 }, altText: { title: "Logo", description: "Logo", name: "logo" } })], { alignment: AlignmentType.CENTER }));
    coverChildren.push(blank());
  }
  coverChildren.push(
    p([B("LAMPIRAN", 32)], { alignment: AlignmentType.CENTER }),
    p([B("LAPORAN HASIL PELAKSANAAN KEGIATAN", 28)], { alignment: AlignmentType.CENTER }),
    p([B("TARUNA TK. II/59/BD", 28)], { alignment: AlignmentType.CENTER }),
    blank(),
    p([B(`BULAN ${monthName.toUpperCase()} ${monthYear}`, 26)], { alignment: AlignmentType.CENTER }),
    blank(), blank(), blank(),
    p([B(TARUNA.nama, 24)], { alignment: AlignmentType.CENTER }),
    p([R(`NO. AK ${TARUNA.noakad}`, 22)], { alignment: AlignmentType.CENTER }),
    p([R(`TON/KI ${TARUNA.tonki}`, 22)], { alignment: AlignmentType.CENTER }),
    blank(), blank(), blank(),
    p([R("DETASEMEN TARUNA TK II/59/BD", 22)], { alignment: AlignmentType.CENTER }),
    p([R("RESIMEN KORPS TARUNA DAN SISWA", 22)], { alignment: AlignmentType.CENTER }),
    p([R("AKADEMI KEPOLISIAN", 22)], { alignment: AlignmentType.CENTER }),
    p([R(`${monthYear}`, 22)], { alignment: AlignmentType.CENTER }),
  );
  // Page 2: daftar isi
  const daftarChildren = [
    p([B("DAFTAR ISI", 24)], { alignment: AlignmentType.CENTER }),
    blank(),
  ];
  files.forEach((f, i) => {
    const cleanName = f.name.replace(/\.docx$/i, "");
    daftarChildren.push(p([R(`${i + 1}.  ${cleanName}`, 22)]));
  });
  daftarChildren.push(blank());
  daftarChildren.push(p([R(`Total: ${files.length} LHP`, 20)], { alignment: AlignmentType.CENTER }));

  const coverDoc = new Document({
    sections: [
      { properties: { page: pageProps }, children: coverChildren },
      { properties: { page: pageProps }, children: daftarChildren },
    ],
  });
  const coverBlob = await Packer.toBlob(coverDoc);

  // ── 2. Use first docx (cover) as base, then append each uploaded file ──
  const masterZip = await JSZip.loadAsync(coverBlob);

  // Parse master document.xml
  let masterDocXml = await masterZip.file("word/document.xml").async("string");

  // Parse master [Content_Types].xml
  let contentTypes = await masterZip.file("[Content_Types].xml").async("string");

  // Parse master word/_rels/document.xml.rels
  let masterRels = await masterZip.file("word/_rels/document.xml.rels")?.async("string") || "";

  // Track relationship IDs to avoid collision
  let relIdCounter = 100;
  let mediaCounter = 100;

  // For each uploaded file, extract body XML, remap images
  for (let fi = 0; fi < files.length; fi++) {
    const file = files[fi];
    const buf = await readFileAsArrayBuffer(file);
    let srcZip;
    try {
      srcZip = await JSZip.loadAsync(buf);
    } catch (e) {
      console.warn("Cannot read:", file.name, e);
      continue;
    }

    const srcDocXml = await srcZip.file("word/document.xml")?.async("string");
    if (!srcDocXml) continue;

    // Extract body content
    const bodyMatch = srcDocXml.match(/<w:body>([\s\S]*)<\/w:body>/);
    if (!bodyMatch) continue;
    let bodyContent = bodyMatch[1];

    // Remove trailing <w:sectPr.../> from body content (section properties)
    bodyContent = bodyContent.replace(/<w:sectPr[\s\S]*?<\/w:sectPr>/g, "");
    bodyContent = bodyContent.replace(/<w:sectPr[^/]*\/>/g, "");

    // Parse source relationships to find images
    const srcRels = await srcZip.file("word/_rels/document.xml.rels")?.async("string") || "";

    // Find all image relationships in source
    const relRegex = /<Relationship[^>]*Id="([^"]*)"[^>]*Target="([^"]*)"[^>]*Type="[^"]*\/image"[^>]*\/?>/g;
    const relRegex2 = /<Relationship[^>]*Type="[^"]*\/image"[^>]*Target="([^"]*)"[^>]*Id="([^"]*)"[^>]*\/?>/g;

    // More flexible regex to capture all attributes regardless of order
    const allRels = [];
    const relLines = srcRels.match(/<Relationship[^>]*\/?\s*>/g) || [];
    for (const line of relLines) {
      const idM = line.match(/Id="([^"]*)"/);
      const targetM = line.match(/Target="([^"]*)"/);
      const typeM = line.match(/Type="([^"]*)"/);
      if (idM && targetM && typeM && typeM[1].includes("/image")) {
        allRels.push({ id: idM[1], target: targetM[1] });
      }
    }

    // Copy each image to master zip with new name, update references
    for (const rel of allRels) {
      const newRelId = `rMerged${relIdCounter++}`;
      const srcMediaPath = rel.target.startsWith("/") ? rel.target.substring(1) : "word/" + rel.target;
      const ext = rel.target.split(".").pop() || "png";
      const newMediaName = `media/merged_${mediaCounter++}.${ext}`;
      const newMediaPath = `word/${newMediaName}`;

      // Copy media file
      const mediaFile = srcZip.file(srcMediaPath) || srcZip.file(rel.target);
      if (mediaFile) {
        const mediaData = await mediaFile.async("uint8array");
        masterZip.file(newMediaPath, mediaData);

        // Add to content types if needed
        const extLower = ext.toLowerCase();
        if (!contentTypes.includes(`Extension="${extLower}"`)) {
          const mimeMap = { png: "image/png", jpg: "image/jpeg", jpeg: "image/jpeg", gif: "image/gif", bmp: "image/bmp", tiff: "image/tiff", emf: "image/x-emf", wmf: "image/x-wmf" };
          const mime = mimeMap[extLower] || `image/${extLower}`;
          contentTypes = contentTypes.replace("</Types>", `<Default Extension="${extLower}" ContentType="${mime}"/></Types>`);
        }
      }

      // Replace relationship ID in body content
      const oldIdEscaped = rel.id.replace(/[.*+?^${}()|[\]\\]/g, "\\$&");
      bodyContent = bodyContent.replace(new RegExp(`r:embed="${oldIdEscaped}"`, "g"), `r:embed="${newRelId}"`);
      bodyContent = bodyContent.replace(new RegExp(`r:link="${oldIdEscaped}"`, "g"), `r:link="${newRelId}"`);

      // Add new relationship to master rels
      masterRels = masterRels.replace("</Relationships>",
        `<Relationship Id="${newRelId}" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/image" Target="${newMediaName}"/></Relationships>`
      );
    }

    // Insert page break + body content before </w:body> in master
    const pageBreakXml = '<w:p><w:pPr><w:sectPr><w:pgSz w:w="11906" w:h="16838"/><w:pgMar w:top="1134" w:right="1134" w:bottom="1134" w:left="1134"/></w:sectPr></w:pPr></w:p>';

    masterDocXml = masterDocXml.replace("</w:body>", pageBreakXml + bodyContent + "</w:body>");
  }

  // ── 3. Write back and download ──
  masterZip.file("word/document.xml", masterDocXml);
  masterZip.file("[Content_Types].xml", contentTypes);
  if (masterRels) {
    masterZip.file("word/_rels/document.xml.rels", masterRels);
  }

  const finalBlob = await masterZip.generateAsync({
    type: "blob",
    mimeType: "application/vnd.openxmlformats-officedocument.wordprocessingml.document",
    compression: "DEFLATE",
    compressionOptions: { level: 6 },
  });

  const url = URL.createObjectURL(finalBlob);
  const a = document.createElement("a");
  a.href = url;
  a.download = `LAMPIRAN_LHP_RADIT_${monthName.toUpperCase()}_${monthYear}.docx`;
  a.click();
  URL.revokeObjectURL(url);
}

function LampiranPage() {
  const [selMonthIdx, setSelMonthIdx] = useState(todayDate.getMonth());
  const [files, setFiles] = useState([]);
  const [loading, setLoading] = useState(false);
  const [done, setDone] = useState(false);
  const [error, setError] = useState("");

  const m = MONTHS[selMonthIdx];

  const handleFileAdd = (e) => {
    const newFiles = Array.from(e.target.files).filter(f =>
      f.name.endsWith(".docx")
    );
    if (newFiles.length === 0) {
      setError("Hanya file .docx yang diterima.");
      return;
    }
    setError("");
    setFiles((prev) => [...prev, ...newFiles.map((f) => ({ file: f, name: f.name, size: f.size }))]);
    e.target.value = "";
  };

  const handleRemoveFile = (idx) => {
    setFiles((prev) => prev.filter((_, i) => i !== idx));
  };

  const handleMoveUp = (idx) => {
    if (idx === 0) return;
    setFiles((prev) => { const n = [...prev]; [n[idx-1], n[idx]] = [n[idx], n[idx-1]]; return n; });
  };

  const handleMoveDown = (idx) => {
    if (idx >= files.length - 1) return;
    setFiles((prev) => { const n = [...prev]; [n[idx], n[idx+1]] = [n[idx+1], n[idx]]; return n; });
  };

  const handleExport = async () => {
    if (files.length === 0) return;
    setLoading(true); setDone(false); setError("");
    try {
      await mergeDocxFiles(files.map(f => f.file), m.name, m.year);
      setDone(true);
      setTimeout(() => setDone(false), 5000);
    } catch (e) {
      setError("Gagal menggabungkan file: " + e.message);
      console.error(e);
    }
    setLoading(false);
  };

  const fmtSize = (b) => b < 1024 ? b + " B" : b < 1048576 ? (b/1024).toFixed(1) + " KB" : (b/1048576).toFixed(1) + " MB";

  return (
    <div>
      <div style={{ fontSize: 13, color: COLORS.muted, marginBottom: 14 }}>
        Upload file LHP (.docx) yang sudah jadi → atur urutan → gabung menjadi 1 file DOCX lengkap dengan cover untuk Dantontar.
      </div>

      {/* Month selector */}
      <div style={{ display: "flex", gap: 4, flexWrap: "wrap", marginBottom: 16 }}>
        {MONTHS.map((mo, i) => (
          <button key={mo.key} onClick={() => setSelMonthIdx(i)} style={{
            padding: "6px 12px", fontSize: 11, fontWeight: selMonthIdx === i ? 700 : 400,
            background: selMonthIdx === i ? COLORS.accentDim : COLORS.surface,
            border: `1px solid ${selMonthIdx === i ? COLORS.accent : COLORS.border}`,
            borderRadius: 6, cursor: "pointer", color: selMonthIdx === i ? COLORS.accent : COLORS.muted,
            fontFamily: "inherit",
          }}>
            {mo.name.slice(0, 3)}
          </button>
        ))}
      </div>

      <div style={{ fontSize: 16, fontWeight: 700, color: COLORS.text, marginBottom: 4 }}>
        Lampiran {m.name} {m.year}
      </div>
      <div style={{ fontSize: 11, color: COLORS.muted, marginBottom: 14 }}>{files.length} file siap digabung</div>

      {/* Upload zone */}
      <label style={{
        display: "flex", flexDirection: "column", alignItems: "center", justifyContent: "center",
        gap: 6, border: `1.5px dashed ${COLORS.borderStrong}`, borderRadius: 10,
        padding: "20px 16px", cursor: "pointer", marginBottom: 14, background: COLORS.bg,
      }}>
        <input type="file" accept=".docx" multiple style={{ display: "none" }} onChange={handleFileAdd} />
        <div style={{ fontSize: 28, lineHeight: 1 }}>📄</div>
        <div style={{ fontSize: 13, fontWeight: 700, color: COLORS.text }}>Klik untuk upload file LHP (.docx)</div>
        <div style={{ fontSize: 11, color: COLORS.muted }}>Bisa pilih banyak file sekaligus · hanya .docx</div>
      </label>

      {/* File list */}
      {files.length > 0 && (
        <div style={{ marginBottom: 14 }}>
          <div style={{ fontSize: 11, color: COLORS.muted, fontWeight: 700, letterSpacing: "0.08em", textTransform: "uppercase", marginBottom: 8 }}>
            Urutan file (atas = pertama)
          </div>
          <div style={{ display: "flex", flexDirection: "column", gap: 6 }}>
            {files.map((f, i) => (
              <div key={i} style={{
                background: COLORS.surface, border: `1px solid ${COLORS.border}`,
                borderRadius: 8, padding: "10px 12px",
                display: "flex", alignItems: "center", gap: 10,
              }}>
                <div style={{
                  width: 24, height: 24, borderRadius: "50%",
                  background: COLORS.accentDim, color: COLORS.accent,
                  fontSize: 11, fontWeight: 700,
                  display: "flex", alignItems: "center", justifyContent: "center", flexShrink: 0,
                }}>{i + 1}</div>
                <div style={{ flex: 1, minWidth: 0 }}>
                  <div style={{ fontSize: 12, fontWeight: 600, color: COLORS.text, overflow: "hidden", textOverflow: "ellipsis", whiteSpace: "nowrap" }}>{f.name}</div>
                  <div style={{ fontSize: 10, color: COLORS.muted }}>{fmtSize(f.size)}</div>
                </div>
                <div style={{ display: "flex", flexDirection: "column", gap: 2 }}>
                  <button onClick={() => handleMoveUp(i)} disabled={i === 0} style={{ background: "none", border: `1px solid ${COLORS.border}`, borderRadius: 3, padding: "1px 6px", cursor: i === 0 ? "default" : "pointer", color: i === 0 ? COLORS.border : COLORS.muted, fontSize: 10, fontFamily: "inherit", lineHeight: 1.2 }}>▲</button>
                  <button onClick={() => handleMoveDown(i)} disabled={i >= files.length-1} style={{ background: "none", border: `1px solid ${COLORS.border}`, borderRadius: 3, padding: "1px 6px", cursor: i >= files.length-1 ? "default" : "pointer", color: i >= files.length-1 ? COLORS.border : COLORS.muted, fontSize: 10, fontFamily: "inherit", lineHeight: 1.2 }}>▼</button>
                </div>
                <button onClick={() => handleRemoveFile(i)} style={{ background: "none", border: "none", color: COLORS.danger, cursor: "pointer", fontSize: 16, padding: "2px 6px", fontFamily: "inherit", lineHeight: 1 }} title="Hapus">✕</button>
              </div>
            ))}
          </div>
        </div>
      )}

      {error && <div style={{ ...S.banner("danger"), marginBottom: 12 }}>{error}</div>}

      {files.length > 0 && (
        <button style={{ ...S.btnPrimary, opacity: loading ? 0.6 : 1 }} onClick={handleExport} disabled={loading}>
          {loading ? `⏳ Menggabungkan ${files.length} file...` : done ? "✓ Berhasil di-download!" : `↓ Gabung ${files.length} LHP → 1 DOCX`}
        </button>
      )}

      {done && (
        <div style={{ ...S.banner("success"), marginTop: 10 }}>
          ✅ File <strong>LAMPIRAN_LHP_RADIT_{m.name.toUpperCase()}_{m.year}.docx</strong> berhasil dibuat. Cover + daftar isi + {files.length} LHP digabung lengkap dengan gambar.
        </div>
      )}

      {files.length === 0 && (
        <div style={{ textAlign: "center", padding: "40px 20px", background: COLORS.surface, border: `1px solid ${COLORS.border}`, borderRadius: 10 }}>
          <div style={{ fontSize: 36, marginBottom: 8 }}>📎</div>
          <div style={{ fontSize: 14, fontWeight: 600, color: COLORS.text, marginBottom: 4 }}>Belum ada file</div>
          <div style={{ fontSize: 12, color: COLORS.muted, lineHeight: 1.5 }}>Buat LHP di tab <strong>Form LHP</strong> → download .docx → upload di sini.</div>
        </div>
      )}

      {files.length > 0 && (
        <div style={{ ...S.banner("warning"), marginTop: 12 }}>
          💡 Semua gambar/logo di dalam file LHP akan ikut tergabung. Atur urutan dengan ▲▼ sebelum export.
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
  const [monthExtras, setMonthExtras] = useState({}); // { [monthKey]: { nilaiPengasuh, teguran: [] } }

  useNotificationPermission();
  useDeadlineNotifications(checked);

  useEffect(() => {
    try {
      const saved = localStorage.getItem("nsp_checked_v2");
      if (saved) setChecked(JSON.parse(saved));
    } catch (_) {}
    try {
      const saved = localStorage.getItem("nsp_month_extras_v1");
      if (saved) setMonthExtras(JSON.parse(saved));
    } catch (_) {}
  }, []);

  const handleToggle = useCallback((monthKey, itemId, val) => {
    setChecked((prev) => {
      const next = { ...prev, [monthKey]: { ...prev[monthKey], [itemId]: val } };
      try { localStorage.setItem("nsp_checked_v2", JSON.stringify(next)); } catch (_) {}
      return next;
    });
  }, []);

  const handleUpdateExtras = useCallback((monthKey, updates) => {
    setMonthExtras((prev) => {
      const next = { ...prev, [monthKey]: { ...prev[monthKey], ...updates } };
      try { localStorage.setItem("nsp_month_extras_v1", JSON.stringify(next)); } catch (_) {}
      return next;
    });
  }, []);

  const tabs = [
    { id: "tracker", label: "Tracker Bulanan" },
    { id: "total", label: "📊 Total NSP" },
    { id: "nsp", label: "Cek NSP" },
    { id: "form", label: "Form LHP" },
    { id: "lampiran", label: "📎 Lampiran" },
    { id: "asisten", label: "✨ Asisten" },
  ];

  return (
    <div style={S.app}>
      <div style={S.header}>
        <div style={S.headerSub}>RESIMEN KORPS TARUNA DAN SISWA · DETASEMEN TK II/59/BD</div>
        <div style={S.headerTitle}>Tracker NSP &amp; LHP — {TARUNA.nama}</div>
      </div>

      <div style={{ ...S.tabs, overflowX: "auto", WebkitOverflowScrolling: "touch" }}>
        {tabs.map((t) => (
          <button key={t.id} style={{ ...S.tab(tab === t.id), whiteSpace: "nowrap" }} onClick={() => setTab(t.id)}>
            {t.label}
          </button>
        ))}
      </div>

      <div style={S.content}>
        <ReminderBanner checked={checked} onGoToTracker={() => setTab("tracker")} />
        {tab === "tracker" && <TrackerPage checked={checked} onToggle={handleToggle} />}
        {tab === "total" && <TotalNSPPage checked={checked} monthExtras={monthExtras} onUpdateExtras={handleUpdateExtras} />}
        {tab === "nsp" && <NSPPage checked={checked} />}
        {tab === "form" && <FormPage />}
        {tab === "lampiran" && <LampiranPage />}
        {tab === "asisten" && <AssistantPage />}
      </div>
    </div>
  );
}

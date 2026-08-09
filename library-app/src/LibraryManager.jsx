import { useState, useEffect, useCallback } from "react";
import { db, auth } from "./firebase";
import { ref, onValue, set, remove } from "firebase/database";
import { signInWithEmailAndPassword, signOut, onAuthStateChanged } from "firebase/auth";
import * as XLSX from "xlsx";

// ⚠️ APNA EMAIL YAHAN LIKHO — sirf yahi email admin hogi
const ADMIN_EMAIL = "skandhnagar90010@gmail.com";n

const DEFAULT_BLOCKS = { A: 20, B: 25, C: 20, D: 15, E: 10 };

const emptyForm = {
  name: "", village: "", mobile: "",
  joiningDate: "", feesDeposit: "", feesDepositDate: ""
};

const EXCEL_STUDENTS = [{"id":"1000","name":"Yogesh Raguwanshi","village":"Atru","mobile":"7000465719","joiningDate":"2026-04-24","feesDeposit":"500","feesDepositDate":"2026-05-04","block":"A","seatNo":1},{"id":"1001","name":"Rajkumar Nagar","village":"Atru","mobile":"7410864571","joiningDate":"2026-05-02","feesDeposit":"600","feesDepositDate":"2026-05-03","block":"A","seatNo":2},{"id":"1002","name":"Dipak","village":"Barla","mobile":"7877957283","joiningDate":"2026-04-30","feesDeposit":"600","feesDepositDate":"2026-05-08","block":"A","seatNo":3},{"id":"1003","name":"Sandeep","village":"Atru","mobile":"8058516218","joiningDate":"2026-04-09","feesDeposit":"500","feesDepositDate":"2026-04-17","block":"A","seatNo":4},{"id":"1004","name":"Jayesh Jangid","village":"Mundla","mobile":"7300481646","joiningDate":"2026-04-09","feesDeposit":"500","feesDepositDate":"2026-04-16","block":"A","seatNo":6},{"id":"1005","name":"Tansuk","village":"Atru","mobile":"8769655089","joiningDate":"2026-04-05","feesDeposit":"500","feesDepositDate":"2026-04-09","block":"A","seatNo":6},{"id":"1006","name":"Rony","village":"Atru","mobile":"7339797569","joiningDate":"2026-04-30","feesDeposit":"600","feesDepositDate":"2026-05-06","block":"B","seatNo":1},{"id":"1007","name":"Pushpendra Ji","village":"Atru","mobile":"9680129886","joiningDate":"2026-04-11","feesDeposit":"500","feesDepositDate":"2026-04-15","block":"B","seatNo":2},{"id":"1008","name":"Shivam","village":"Atru","mobile":"","joiningDate":"2026-03-22","feesDeposit":"","feesDepositDate":"","block":"B","seatNo":2},{"id":"1009","name":"Shailender","village":"Atru","mobile":"6375784716","joiningDate":"2026-04-30","feesDeposit":"500","feesDepositDate":"2026-05-04","block":"B","seatNo":3},{"id":"1010","name":"Monu","village":"Atru","mobile":"960035885","joiningDate":"2026-04-15","feesDeposit":"500","feesDepositDate":"","block":"B","seatNo":4},{"id":"1011","name":"Ajay","village":"Atru","mobile":"7742436991","joiningDate":"2026-05-04","feesDeposit":"600","feesDepositDate":"2026-05-06","block":"B","seatNo":5},{"id":"1012","name":"Devendra","village":"Atru","mobile":"9509092980","joiningDate":"2026-05-05","feesDeposit":"500","feesDepositDate":"2026-05-11","block":"B","seatNo":5},{"id":"1013","name":"Jaswant Meena","village":"Kunjed","mobile":"9636730573","joiningDate":"2026-05-01","feesDeposit":"600","feesDepositDate":"2026-05-05","block":"C","seatNo":1},{"id":"1014","name":"Mohit Meena","village":"Barla","mobile":"6378107313","joiningDate":"2026-04-14","feesDeposit":"500","feesDepositDate":"2026-04-16","block":"C","seatNo":2},{"id":"1015","name":"Mahendra Meena","village":"Barla","mobile":"8824612618","joiningDate":"2026-04-14","feesDeposit":"500","feesDepositDate":"2026-04-16","block":"C","seatNo":3},{"id":"1016","name":"Mukesh Ji Parihar","village":"Atru","mobile":"7597281040","joiningDate":"2026-05-01","feesDeposit":"1100","feesDepositDate":"2026-05-02","block":"C","seatNo":4},{"id":"1017","name":"Vikram Meena","village":"Atru","mobile":"9664012485","joiningDate":"2026-05-11","feesDeposit":"600","feesDepositDate":"2026-05-12","block":"C","seatNo":5},{"id":"1018","name":"Lokendra","village":"Atru","mobile":"9649613299","joiningDate":"2026-04-25","feesDeposit":"600","feesDepositDate":"2026-05-04","block":"C","seatNo":6},{"id":"1019","name":"Narsh","village":"Atru","mobile":"9602731486","joiningDate":"2026-04-10","feesDeposit":"400","feesDepositDate":"2026-04-18","block":"C","seatNo":7},{"id":"1020","name":"Ankit","village":"Atru","mobile":"6367546833","joiningDate":"2026-04-25","feesDeposit":"600","feesDepositDate":"2026-05-03","block":"C","seatNo":8},{"id":"1021","name":"Pawan Megwal","village":"Atru","mobile":"8504863387","joiningDate":"2026-04-23","feesDeposit":"600","feesDepositDate":"2026-04-28","block":"C","seatNo":9},{"id":"1022","name":"Hemant","village":"Anatana","mobile":"7375923664","joiningDate":"2026-05-07","feesDeposit":"600","feesDepositDate":"2026-05-08","block":"C","seatNo":10},{"id":"1023","name":"Lovekush","village":"Atru","mobile":"9672647338","joiningDate":"2026-04-01","feesDeposit":"500","feesDepositDate":"2026-04-10","block":"C","seatNo":10},{"id":"1024","name":"Soniya Meena","village":"Atru","mobile":"7073149217","joiningDate":"2026-05-01","feesDeposit":"600","feesDepositDate":"2026-05-04","block":"C","seatNo":11},{"id":"1025","name":"Yash","village":"Atru","mobile":"8955906709","joiningDate":"2026-05-05","feesDeposit":"","feesDepositDate":"","block":"C","seatNo":12},{"id":"1026","name":"Dilkush","village":"Barla","mobile":"9509442292","joiningDate":"2026-05-05","feesDeposit":"600","feesDepositDate":"2026-05-11","block":"C","seatNo":13},{"id":"1027","name":"Jyoti","village":"Atru","mobile":"9672647338","joiningDate":"2026-04-15","feesDeposit":"600","feesDepositDate":"2026-05-03","block":"C","seatNo":15},{"id":"1028","name":"Pankaj","village":"Atru","mobile":"","joiningDate":"2026-03-19","feesDeposit":"500","feesDepositDate":"2026-03-24","block":"D","seatNo":1},{"id":"1029","name":"Vijendra","village":"Ganesh Mohalla Atru","mobile":"9001597984","joiningDate":"2026-04-18","feesDeposit":"600","feesDepositDate":"2026-04-22","block":"D","seatNo":1},{"id":"1030","name":"Bhanu","village":"Aaton","mobile":"8955646781","joiningDate":"2026-04-16","feesDeposit":"500","feesDepositDate":"2026-04-18","block":"D","seatNo":2},{"id":"1031","name":"Hitesh","village":"Barla","mobile":"8875198892","joiningDate":"2026-05-01","feesDeposit":"600","feesDepositDate":"2026-05-04","block":"D","seatNo":3},{"id":"1032","name":"Mayank","village":"Aaton","mobile":"6376126941","joiningDate":"2026-04-18","feesDeposit":"600","feesDepositDate":"2026-04-24","block":"D","seatNo":4},{"id":"1033","name":"Manoj Nagar","village":"Aaton","mobile":"7340576826","joiningDate":"2026-04-16","feesDeposit":"500","feesDepositDate":"2026-04-22","block":"D","seatNo":5},{"id":"1034","name":"Yogesh Kumar Merotha","village":"Barla","mobile":"9610058236","joiningDate":"2026-04-30","feesDeposit":"600","feesDepositDate":"2026-05-02","block":"D","seatNo":6},{"id":"1035","name":"Kuldeep","village":"Kachra","mobile":"8890737167","joiningDate":"2026-04-16","feesDeposit":"500","feesDepositDate":"2026-04-18","block":"D","seatNo":7},{"id":"1036","name":"Gourav","village":"Atru","mobile":"9511366726","joiningDate":"2026-04-02","feesDeposit":"500","feesDepositDate":"2026-04-12","block":"D","seatNo":8},{"id":"1037","name":"Karan Nagar","village":"Patna","mobile":"6377975962","joiningDate":"2026-04-22","feesDeposit":"500","feesDepositDate":"2026-04-22","block":"D","seatNo":9},{"id":"1038","name":"Yash Gautam","village":"Moondla","mobile":"7014812501","joiningDate":"2026-04-09","feesDeposit":"500","feesDepositDate":"2026-04-19","block":"D","seatNo":10},{"id":"1039","name":"Durgesh Gujar","village":"Barla","mobile":"8003611070","joiningDate":"2026-04-23","feesDeposit":"600","feesDepositDate":"2026-04-28","block":"D","seatNo":11},{"id":"1040","name":"Kiran Gocher","village":"Aaton","mobile":"6350071079","joiningDate":"2026-04-20","feesDeposit":"600","feesDepositDate":"2026-04-28","block":"D","seatNo":11},{"id":"1041","name":"Dashrath Singh","village":"Atru","mobile":"9521821264","joiningDate":"2026-04-24","feesDeposit":"500","feesDepositDate":"2026-04-04","block":"D","seatNo":13},{"id":"1042","name":"Ankit Nagar","village":"Patna","mobile":"7878291975","joiningDate":"2026-04-18","feesDeposit":"600","feesDepositDate":"2026-04-22","block":"D","seatNo":14},{"id":"1043","name":"Mahipaal","village":"Atru","mobile":"7339717405","joiningDate":"2026-05-03","feesDeposit":"500","feesDepositDate":"2026-05-10","block":"D","seatNo":15},{"id":"1044","name":"Kiran Megwal","village":"Atru","mobile":"9571423287","joiningDate":"2026-04-15","feesDeposit":"500","feesDepositDate":"2026-04-28","block":"E","seatNo":1},{"id":"1045","name":"Padma","village":"Atru","mobile":"9358615924","joiningDate":"2026-04-15","feesDeposit":"600","feesDepositDate":"2026-05-04","block":"E","seatNo":2},{"id":"1046","name":"Kiran Megwal","village":"Atru","mobile":"9602630155","joiningDate":"2026-04-16","feesDeposit":"600","feesDepositDate":"2026-05-04","block":"E","seatNo":3},{"id":"1047","name":"Vartika","village":"Atru","mobile":"8107197251","joiningDate":"2026-03-17","feesDeposit":"900","feesDepositDate":"2026-03-21","block":"E","seatNo":4},{"id":"1048","name":"Manisha","village":"Atru","mobile":"","joiningDate":"2026-05-02","feesDeposit":"","feesDepositDate":"","block":"E","seatNo":5},{"id":"1049","name":"Teena Mehra","village":"Atru","mobile":"9461782074","joiningDate":"2026-03-28","feesDeposit":"400","feesDepositDate":"2026-04-11","block":"E","seatNo":5},{"id":"1050","name":"Rishu","village":"Atru","mobile":"","joiningDate":"2026-05-12","feesDeposit":"","feesDepositDate":"","block":"E","seatNo":6},{"id":"1051","name":"Vishaka","village":"Kachra","mobile":"8890737167","joiningDate":"2026-04-16","feesDeposit":"500","feesDepositDate":"2026-04-18","block":"E","seatNo":7},{"id":"1052","name":"Kavita Megwal","village":"Atru","mobile":"","joiningDate":"2026-05-01","feesDeposit":"500","feesDepositDate":"2026-05-09","block":"E","seatNo":8},{"id":"1053","name":"Rachna","village":"Atru","mobile":"7728910307","joiningDate":"2026-04-11","feesDeposit":"500","feesDepositDate":"2026-04-28","block":"E","seatNo":9},{"id":"1054","name":"Vishaka","village":"Navodya","mobile":"9116611281","joiningDate":"2026-03-20","feesDeposit":"500","feesDepositDate":"2026-04-04","block":"E","seatNo":9},{"id":"1055","name":"Sonu Megwal","village":"Atru","mobile":"7297027415","joiningDate":"2026-04-01","feesDeposit":"500","feesDepositDate":"2026-04-11","block":"E","seatNo":10},{"id":"1056","name":"Rajkumari","village":"Kacra","mobile":"9079232791","joiningDate":"2026-05-04","feesDeposit":"500","feesDepositDate":"2026-05-08","block":"E","seatNo":11},{"id":"1057","name":"Sheetal Nagar","village":"Atru","mobile":"7568286793","joiningDate":"2026-04-12","feesDeposit":"500","feesDepositDate":"2026-04-30","block":"E","seatNo":12},{"id":"1058","name":"Jaya Nagar","village":"Atru","mobile":"9511366726","joiningDate":"2026-04-02","feesDeposit":"500","feesDepositDate":"2026-04-12","block":"E","seatNo":13}];

// ── helpers ──────────────────────────────────────────────
const fmtDate = (d) => d ? new Date(d).toLocaleDateString("en-IN", { day: "numeric", month: "short", year: "numeric" }) : "—";

export default function App() {
  const [view, setView]               = useState("home");
  const [selectedBlock, setSelectedBlock] = useState(null);
  const [selectedSeat, setSelectedSeat]   = useState(null);
  const [blocks, setBlocks]           = useState(DEFAULT_BLOCKS);
  const [students, setStudents]       = useState([]);
  const [loading, setLoading]         = useState(true);
  const [formData, setFormData]       = useState(emptyForm);
  const [editId, setEditId]           = useState(null);
  const [showSettings, setShowSettings] = useState(false);
  const [tempSeats, setTempSeats]     = useState({});
  const [saving, setSaving]           = useState(false);
  const [searchQuery, setSearchQuery] = useState("");
  const [showFeesDue, setShowFeesDue] = useState(false);
  const [showExpiry, setShowExpiry] = useState(false);
  const [fbStatus, setFbStatus]       = useState("connecting");
  const [isAdmin, setIsAdmin]         = useState(false);
  const [authChecked, setAuthChecked] = useState(false);
  const [showLogin, setShowLogin]     = useState(false);
  const [loginEmail, setLoginEmail]   = useState("");
  const [loginPass, setLoginPass]     = useState("");
  const [loginError, setLoginError]   = useState("");
  const [loginLoading, setLoginLoading] = useState(false);
  const [showFeesAmount, setShowFeesAmount] = useState(false);

  // ── Auth state listener ───────────────────────────────
  useEffect(() => {
    const unsub = onAuthStateChanged(auth, (user) => {
      setIsAdmin(!!(user && user.email === ADMIN_EMAIL));
      setAuthChecked(true);
    });
    return unsub;
  }, []);

  const handleLogin = async () => {
    setLoginError(""); setLoginLoading(true);
    try {
      await signInWithEmailAndPassword(auth, loginEmail, loginPass);
      setShowLogin(false); setLoginEmail(""); setLoginPass("");
    } catch (e) {
      setLoginError("Galat email ya password. Dobara try karo.");
    }
    setLoginLoading(false);
  };

  const handleLogout = () => signOut(auth);


  // ── Firebase live listener ────────────────────────────
  useEffect(() => {
    const blocksRef = ref(db, "blocks");
    const studentsRef = ref(db, "students");

    const unsubBlocks = onValue(blocksRef, (snap) => {
      if (snap.exists()) setBlocks(snap.val());
      setFbStatus("ok");
    }, (err) => { console.error(err); setFbStatus("error"); setLoading(false); });

    const unsubStudents = onValue(studentsRef, (snap) => {
      if (snap.exists()) {
        const obj = snap.val();
        setStudents(Object.values(obj).map(s => ({ ...s, seatNo: Number(s.seatNo) })));
      } else {
        setStudents([]);
      }
      setLoading(false);
    }, (err) => { console.error(err); setFbStatus("error"); setLoading(false); });

    return () => { unsubBlocks(); unsubStudents(); };
  }, []);

  // ── Save helpers ──────────────────────────────────────
  const saveStudents = useCallback(async (next) => {
    const obj = {};
    next.forEach(s => { obj[s.id] = s; });
    await set(ref(db, "students"), obj);
  }, []);

  const saveBlocks = useCallback(async (next) => {
    await set(ref(db, "blocks"), next);
  }, []);

  const saveOneStudent = useCallback(async (s) => {
    await set(ref(db, `students/${s.id}`), s);
  }, []);

  const deleteStudentFb = useCallback(async (id) => {
    await remove(ref(db, `students/${id}`));
  }, []);

  // ── Excel import ──────────────────────────────────────
  const importFromExcel = async () => {
    const normalized = EXCEL_STUDENTS.map(s => ({ ...s, seatNo: Number(s.seatNo) }));
    if (students.length > 0 && !window.confirm(`Pehle se ${students.length} students hain. Import karein? (Purana data replace hoga)`)) return;
    await saveStudents(normalized);
    alert(`✅ ${normalized.length} students import ho gaye!`);
  };

  // ── Excel export ──────────────────────────────────────
  const exportToExcel = () => {
    const rows = students.map(s => ({
      "Block-Seat": `${s.block}-${s.seatNo}`,
      "Name":        s.name,
      "Village":     s.village || "",
      "Mobile":      s.mobile || "",
      "Joining Date": s.joiningDate || "",
      "Fees (₹)":    s.feesDeposit || "",
      "Fees Date":   s.feesDepositDate || "",
    }));
    rows.sort((a, b) => a["Block-Seat"].localeCompare(b["Block-Seat"]));
    const ws = XLSX.utils.json_to_sheet(rows);
    ws["!cols"] = [10,22,16,14,14,10,14].map(w => ({ wch: w }));
    const wb = XLSX.utils.book_new();
    XLSX.utils.book_append_sheet(wb, ws, "Students");
    XLSX.writeFile(wb, `Library_Students_${new Date().toISOString().slice(0,10)}.xlsx`);
  };

  // ── Seat helpers ──────────────────────────────────────
  const getSeatStudents = (block, seat) =>
    students.filter(s => s.block === block && Number(s.seatNo) === Number(seat));

  const blockStats = (b) => {
    const total = blocks[b] || 0;
    const occupied = Array.from({ length: total }, (_, i) => getSeatStudents(b, i+1).length > 0).filter(Boolean).length;
    return { total, occupied, available: total - occupied };
  };

  // ── Form submit ───────────────────────────────────────
  const handleSubmit = async () => {
    if (!formData.name.trim()) return;
    setSaving(true);
    const s = {
      id: editId || Date.now().toString(),
      ...formData,
      block: selectedBlock,
      seatNo: Number(selectedSeat)
    };
    await saveOneStudent(s);
    setSaving(false);
    setEditId(null);
    setFormData(emptyForm);
    setView("seat");
  };

  const openEdit = (student) => {
    setEditId(student.id);
    setFormData({
      name: student.name || "", village: student.village || "",
      mobile: student.mobile || "", joiningDate: student.joiningDate || "",
      feesDeposit: student.feesDeposit || "", feesDepositDate: student.feesDepositDate || ""
    });
    setView("form");
  };

  const openAdd = () => {
    setEditId(null);
    const today = new Date().toISOString().split("T")[0];
    setFormData({ ...emptyForm, joiningDate: today, feesDepositDate: today });
    setView("form");
  };

  const deleteStudent = async (id) => {
    if (window.confirm("Is student ka record permanently delete karein?")) {
      await deleteStudentFb(id);
    }
  };

  // ── Expiry helpers ────────────────────────────────────
  const getDaysLeft = (joiningDate) => {
    if (!joiningDate) return null;
    const expiry = new Date(joiningDate);
    expiry.setDate(expiry.getDate() + 30);
    expiry.setHours(0,0,0,0);
    const today = new Date(); today.setHours(0,0,0,0);
    return Math.ceil((expiry - today) / 86400000);
  };

  const feesDueStudents = students
    .filter(s => !parseFloat(s.feesDeposit || 0))
    .sort((a, b) => (a.block || "").localeCompare(b.block || "") || Number(a.seatNo) - Number(b.seatNo));

  const feesNotDeposited = feesDueStudents.length;

  const sortedByExpiry = [...students]
    .filter(s => s.joiningDate && parseFloat(s.feesDeposit || 0) > 0)
    .map(s => ({ ...s, daysLeft: getDaysLeft(s.joiningDate) }))
    .sort((a, b) => a.daysLeft - b.daysLeft);

  const totalFees = students.reduce((a, s) => a + (parseFloat(s.feesDeposit) || 0), 0);

  // ── Styles ────────────────────────────────────────────
  const styles = {
    page:      { padding: "1.5rem", maxWidth: "760px", margin: "0 auto" },
    topbar:    { display: "flex", alignItems: "center", justifyContent: "space-between", marginBottom: "1.5rem", flexWrap: "wrap", gap: "10px" },
    h1:        { fontSize: "22px", fontWeight: 500, margin: 0, color: "var(--color-text-primary)" },
    sub:       { fontSize: "13px", color: "var(--color-text-secondary)", margin: "3px 0 0" },
    card:      { background: "var(--color-background-primary)", border: "0.5px solid var(--color-border-secondary)", borderRadius: "var(--border-radius-lg)", padding: "1rem 1.25rem" },
    backBtn:   { display: "flex", alignItems: "center", gap: "6px", fontSize: "14px", cursor: "pointer", background: "none", border: "0.5px solid var(--color-border-secondary)", borderRadius: "var(--border-radius-md)", padding: "6px 12px", color: "var(--color-text-primary)" },
    primaryBtn:{ display: "flex", alignItems: "center", gap: "6px", fontSize: "14px", background: "#1D9E75", color: "white", border: "none", borderRadius: "var(--border-radius-md)", padding: "8px 16px", cursor: "pointer" },
    label:     { fontSize: "11px", color: "var(--color-text-secondary)", margin: "0 0 3px", letterSpacing: "0.04em" },
    val:       { fontSize: "14px", margin: 0, color: "var(--color-text-primary)" },
  };

  // ── Loading ───────────────────────────────────────────
  if (loading || !authChecked) return (
    <div style={{ display: "flex", flexDirection: "column", alignItems: "center", justifyContent: "center", height: "200px", gap: "12px", color: "var(--color-text-secondary)" }}>
      <i className="ti ti-loader-2" style={{ fontSize: "28px" }}></i>
      <span>Firebase se data load ho raha hai…</span>
    </div>
  );

  if (fbStatus === "error") return (
    <div style={{ padding: "2rem", textAlign: "center", color: "var(--color-text-danger)" }}>
      <i className="ti ti-wifi-off" style={{ fontSize: "32px", display: "block", marginBottom: "10px" }}></i>
      <strong>Firebase connect nahi ho pa raha.</strong><br />
      <span style={{ fontSize: "13px" }}>Check karein: firebase.js mein apna config paste kiya hai?</span>
    </div>
  );

  /* ══════════════════════════════════════════════════════
     HOME VIEW
  ══════════════════════════════════════════════════════ */
  if (view === "home") {
    const filtered = searchQuery
      ? students.filter(s =>
          s.name.toLowerCase().includes(searchQuery.toLowerCase()) ||
          (s.village||"").toLowerCase().includes(searchQuery.toLowerCase()) ||
          (s.mobile||"").includes(searchQuery))
      : [];

    return (
      <div style={styles.page}>

        {/* Header */}
        <div style={styles.topbar}>
          <div>
            <h1 style={styles.h1}>
              <i className="ti ti-books" style={{ marginRight: "10px", color: "#1D9E75" }}></i>
              Library Manager
              <span style={{ marginLeft: "10px", fontSize: "11px", padding: "2px 8px", borderRadius: "20px", background: "#E1F5EE", color: "#1D9E75", verticalAlign: "middle" }}>
                🔴 LIVE
              </span>
            </h1>
            <p style={styles.sub}>
              {isAdmin
                ? "Admin Mode — Full Access"
                : "Guest View — Sirf seat availability dikh rahi hai"}
            </p>
          </div>
          <div style={{ display: "flex", gap: "8px", flexWrap: "wrap" }}>
            {isAdmin && <>
              <button onClick={exportToExcel} style={{ ...styles.backBtn, borderColor: "#1D9E75", color: "#1D9E75" }}>
                <i className="ti ti-file-spreadsheet"></i> Excel Download
              </button>
              <button onClick={() => setShowExpiry(true)} style={{ ...styles.backBtn, borderColor: "#E07B2A", color: "#E07B2A" }}>
                <i className="ti ti-clock-exclamation"></i> Expiry List
              </button>
              <button onClick={() => setShowFeesDue(true)} style={{ ...styles.backBtn, borderColor: "#DC2626", color: "#DC2626" }}>
                <i className="ti ti-alert-circle"></i> Fees Due ({feesNotDeposited})
              </button>
              <button onClick={importFromExcel} style={{ ...styles.backBtn, borderColor: "#6366f1", color: "#6366f1" }}>
                <i className="ti ti-table-import"></i> Import
              </button>
              <button onClick={() => { setTempSeats({ ...blocks }); setShowSettings(true); }} style={styles.backBtn}>
                <i className="ti ti-settings"></i> Settings
              </button>
              <button onClick={handleLogout} style={{ ...styles.backBtn, borderColor: "#DC2626", color: "#DC2626" }}>
                <i className="ti ti-logout"></i> Logout
              </button>
            </>}
            {!isAdmin && (
              <button onClick={() => { setShowLogin(true); setLoginError(""); }} style={{ ...styles.backBtn, borderColor: "#1D9E75", color: "#1D9E75" }}>
                <i className="ti ti-lock"></i> Admin Login
              </button>
            )}
          </div>
        </div>

        {/* Stats */}
        <div style={{ display: "grid", gridTemplateColumns: "repeat(auto-fit, minmax(150px, 1fr))", gap: "10px", marginBottom: "1.5rem" }}>          
            {[
  { label: "Total Seats", val: Object.values(blocks).reduce((a,b)=>a+b,0), icon: "ti-armchair" },
  { label: "Students Enrolled", val: students.length, icon: "ti-users" },
  { 
    label: "Fees Collected", 
    val: showFeesAmount ? `₹${totalFees.toLocaleString("en-IN")}` : "••••••", 
    icon: "ti-coin-rupee",
    isFees: true
  },
  { 
    label: "Fees Not Deposited", 
    val: feesNotDeposited, 
    icon: "ti-alert-circle",
    isDue: true
  }      ].map(({ label, val, icon, isFees, isDue }) => (
            <div
  key={label}
  onClick={isDue ? () => setShowFeesDue(true) : undefined}
  style={{
    background: "var(--color-background-secondary)",
    borderRadius: "var(--border-radius-md)",
    padding: "14px 16px",
    minHeight: "92px",
    cursor: isDue ? "pointer" : "default",
    overflow: "hidden"
  }}
>
              <div style={{ display: "flex", alignItems: "center", gap: "8px", marginBottom: "6px" }}>
                <i className={`ti ${icon}`} style={{ fontSize: "16px", color: "#1D9E75" }}></i>
                <p style={{ fontSize: "12px", color: "var(--color-text-secondary)", margin: 0 }}>{label}</p>
              </div>
              <div style={{ display: "flex", alignItems: "center", gap: "8px" }}>
                <p style={{ fontSize: "22px", fontWeight: 500, margin: 0, whiteSpace: "nowrap" }}>{val}</p>
                {isFees && (
                  <button
                    onClick={(e) => { e.stopPropagation(); setShowFeesAmount(!showFeesAmount); }}
                    style={{ width: "32px", height: "32px", minWidth: "32px", display: "flex", alignItems: "center", justifyContent: "center", background: "white", border: "1px solid var(--color-border-secondary)", borderRadius: "50%", cursor: "pointer", fontSize: "16px", color: "#1D9E75", flexShrink: 0 }}
                    title={showFeesAmount ? "Hide amount" : "Show amount"}
                  >
                    <i className={`ti ${showFeesAmount ? "ti-eye-off" : "ti-eye"}`}></i>
                  </button>
                )}
              </div>
            </div>
          ))}
        </div>

        {/* Search — admin only */}
        {isAdmin && (
        <div style={{ position: "relative", marginBottom: "1.25rem" }}>
          <i className="ti ti-search" style={{ position: "absolute", left: "12px", top: "50%", transform: "translateY(-50%)", fontSize: "16px", color: "var(--color-text-secondary)" }}></i>
          <input type="text" placeholder="Search by name, village, or mobile…"
            value={searchQuery} onChange={e => setSearchQuery(e.target.value)}
            style={{ width: "100%", paddingLeft: "38px", boxSizing: "border-box" }} />
        </div>
        )}

        {isAdmin && searchQuery && (
          <div style={{ marginBottom: "1.5rem" }}>
            {filtered.length === 0
              ? <p style={{ color: "var(--color-text-secondary)", fontSize: "14px" }}>No students found.</p>
              : <div style={{ display: "flex", flexDirection: "column", gap: "8px" }}>
                  {filtered.map(s => (
                    <div key={s.id} onClick={() => { setSelectedBlock(s.block); setSelectedSeat(s.seatNo); setView("seat"); }}
                      style={{ ...styles.card, cursor: "pointer", display: "flex", alignItems: "center", justifyContent: "space-between" }}>
                      <div style={{ display: "flex", alignItems: "center", gap: "12px" }}>
                        <div style={{ width: "36px", height: "36px", borderRadius: "50%", background: "var(--color-background-info)", display: "flex", alignItems: "center", justifyContent: "center", fontSize: "13px", fontWeight: 500, color: "var(--color-text-info)", flexShrink: 0 }}>
                          {s.name.split(" ").map(w=>w[0]).slice(0,2).join("").toUpperCase()}
                        </div>
                        <div>
                          <p style={{ fontWeight: 500, fontSize: "14px", margin: 0 }}>{s.name}</p>
                          <p style={{ fontSize: "12px", color: "var(--color-text-secondary)", margin: "2px 0 0" }}>
                            {s.village} · Block {s.block} · Seat {s.seatNo}
                          </p>
                        </div>
                      </div>
                      <span style={{ fontSize: "12px", padding: "3px 8px", borderRadius: "var(--border-radius-md)", background: "var(--color-background-success)", color: "var(--color-text-success)" }}>
                        ₹{parseFloat(s.feesDeposit||0).toLocaleString("en-IN")}
                      </span>
                    </div>
                  ))}
                </div>
            }
          </div>
        )}

        {/* Block Cards */}
        {(!isAdmin || !searchQuery) && (
          <>
            <p style={{ fontSize: "13px", color: "var(--color-text-secondary)", marginBottom: "10px" }}>Select a block to view seat layout</p>
            <div style={{ display: "grid", gridTemplateColumns: "repeat(auto-fill, minmax(200px, 1fr))", gap: "12px" }}>
              {Object.keys(blocks).map(b => {
                const s = blockStats(b);
                const pct = s.total > 0 ? Math.round(s.occupied/s.total*100) : 0;
                return (
                  <div key={b} onClick={() => { setSelectedBlock(b); setView("block"); }}
                    style={{ ...styles.card, cursor: "pointer", borderLeft: "3px solid #1D9E75" }}>
                    <div style={{ display: "flex", alignItems: "flex-start", justifyContent: "space-between", marginBottom: "14px" }}>
                      <span style={{ fontSize: "30px", fontWeight: 500, lineHeight: 1 }}>Block <span style={{ color: "#1D9E75" }}>{b}</span></span>
                      <span style={{ fontSize: "11px", padding: "3px 8px", borderRadius: "var(--border-radius-md)", background: s.available===0?"var(--color-background-danger)":"var(--color-background-success)", color: s.available===0?"var(--color-text-danger)":"var(--color-text-success)" }}>
                        {s.available} free
                      </span>
                    </div>
                    <div style={{ height: "5px", background: "var(--color-border-tertiary)", borderRadius: "3px", overflow: "hidden", marginBottom: "10px" }}>
                      <div style={{ width: `${pct}%`, height: "100%", background: "#1D9E75", borderRadius: "3px" }}></div>
                    </div>
                    <div style={{ display: "flex", justifyContent: "space-between", fontSize: "13px", color: "var(--color-text-secondary)" }}>
                      <span><strong style={{ color: "var(--color-text-primary)" }}>{s.occupied}</strong> occupied</span>
                      <span>{s.total} total</span>
                    </div>
                  </div>
                );
              })}
            </div>
          </>
        )}

        {/* Login Modal */}
        {showLogin && (
          <div style={{ position: "fixed", inset: 0, background: "rgba(0,0,0,0.55)", display: "flex", alignItems: "center", justifyContent: "center", zIndex: 300, padding: "1rem" }}>
            <div style={{ background: "var(--color-background-primary)", borderRadius: "var(--border-radius-lg)", width: "100%", maxWidth: "360px", padding: "2rem", border: "0.5px solid var(--color-border-secondary)", boxShadow: "0 8px 32px rgba(0,0,0,0.18)" }}>
              <div style={{ textAlign: "center", marginBottom: "1.5rem" }}>
                <div style={{ width: "56px", height: "56px", borderRadius: "50%", background: "#E1F5EE", display: "flex", alignItems: "center", justifyContent: "center", margin: "0 auto 12px" }}>
                  <i className="ti ti-lock" style={{ fontSize: "26px", color: "#1D9E75" }}></i>
                </div>
                <h2 style={{ fontSize: "20px", fontWeight: 600, margin: "0 0 4px" }}>Admin Login</h2>
                <p style={{ fontSize: "13px", color: "var(--color-text-secondary)", margin: 0 }}>Sirf authorized admin access kar sakta hai</p>
              </div>
              <div style={{ display: "flex", flexDirection: "column", gap: "14px" }}>
                <div>
                  <label style={{ fontSize: "13px", color: "var(--color-text-secondary)", display: "block", marginBottom: "6px" }}>Email</label>
                  <input type="email" value={loginEmail} onChange={e => setLoginEmail(e.target.value)}
                    placeholder="admin@email.com" style={{ width: "100%", boxSizing: "border-box" }} />
                </div>
                <div>
                  <label style={{ fontSize: "13px", color: "var(--color-text-secondary)", display: "block", marginBottom: "6px" }}>Password</label>
                  <input type="password" value={loginPass} onChange={e => setLoginPass(e.target.value)}
                    placeholder="••••••••" style={{ width: "100%", boxSizing: "border-box" }}
                    onKeyDown={e => e.key === "Enter" && handleLogin()} />
                </div>
                {loginError && (
                  <div style={{ background: "#FEF2F2", border: "1px solid #FECACA", borderRadius: "var(--border-radius-md)", padding: "10px 14px", fontSize: "13px", color: "#DC2626" }}>
                    <i className="ti ti-alert-circle" style={{ marginRight: "6px" }}></i>{loginError}
                  </div>
                )}
                <button onClick={handleLogin} disabled={loginLoading || !loginEmail || !loginPass}
                  style={{ background: "#1D9E75", color: "white", border: "none", borderRadius: "var(--border-radius-md)", padding: "11px", fontSize: "15px", cursor: "pointer", opacity: loginLoading ? 0.7 : 1, marginTop: "4px" }}>
                  {loginLoading ? "Logging in…" : "Login"}
                </button>
                <button onClick={() => { setShowLogin(false); setLoginError(""); setLoginEmail(""); setLoginPass(""); }}
                  style={{ background: "none", border: "0.5px solid var(--color-border-secondary)", borderRadius: "var(--border-radius-md)", padding: "9px", fontSize: "14px", cursor: "pointer", color: "var(--color-text-secondary)" }}>
                  Cancel
                </button>
              </div>
            </div>
          </div>
        )}

        {/* Expiry Modal */}
        {showExpiry && (
          <div style={{ position: "fixed", inset: 0, background: "rgba(0,0,0,0.5)", display: "flex", alignItems: "center", justifyContent: "center", zIndex: 200, padding: "1rem" }}>
            <div style={{ background: "var(--color-background-primary)", borderRadius: "var(--border-radius-lg)", width: "100%", maxWidth: "560px", maxHeight: "85vh", display: "flex", flexDirection: "column", border: "0.5px solid var(--color-border-secondary)" }}>
              <div style={{ padding: "1.25rem 1.5rem", borderBottom: "0.5px solid var(--color-border-secondary)", display: "flex", alignItems: "center", justifyContent: "space-between", flexShrink: 0 }}>
                <div>
                  <h2 style={{ fontSize: "17px", fontWeight: 600, margin: 0 }}>
                    <i className="ti ti-clock-exclamation" style={{ marginRight: "8px", color: "#E07B2A" }}></i>
                    Subscription Expiry List
                  </h2>
                  <p style={{ fontSize: "12px", color: "var(--color-text-secondary)", margin: "3px 0 0" }}>Fees date + 30 din — sabse pehle expire hone wale upar</p>
                </div>
                <button onClick={() => setShowExpiry(false)} style={{ background: "none", border: "none", cursor: "pointer", fontSize: "20px" }}>
                  <i className="ti ti-x"></i>
                </button>
              </div>
              <div style={{ padding: "8px 1.5rem", borderBottom: "0.5px solid var(--color-border-tertiary)", display: "flex", gap: "16px", fontSize: "12px", flexShrink: 0, flexWrap: "wrap" }}>
                {[["#DC2626","Expired"],["#E07B2A","1–7 din"],["#D4A017","8–15 din"],["#1D9E75","15+ din"]].map(([c,l])=>(
                  <span key={l} style={{ display: "flex", alignItems: "center", gap: "5px" }}>
                    <span style={{ width: "10px", height: "10px", borderRadius: "50%", background: c, display: "inline-block" }}></span>{l}
                  </span>
                ))}
              </div>
              <div style={{ overflowY: "auto", flex: 1, padding: "0.75rem 1.5rem 1.25rem" }}>
                {sortedByExpiry.length === 0
                  ? <p style={{ textAlign: "center", color: "var(--color-text-secondary)", paddingTop: "2rem" }}>Koi record nahi</p>
                  : sortedByExpiry.map((s, idx) => {
                      const d = s.daysLeft;
                      const clr = d<=0?"#DC2626":d<=7?"#E07B2A":d<=15?"#D4A017":"#1D9E75";
                      const bg  = d<=0?"#FEF2F2":d<=7?"#FFF4ED":d<=15?"#FFFBEB":"#F0FDF8";
                      const exp = new Date(s.joiningDate); exp.setDate(exp.getDate()+30);
                      return (
                        <div key={s.id} onClick={() => { setSelectedBlock(s.block); setSelectedSeat(s.seatNo); setShowExpiry(false); setView("seat"); }}
                          style={{ display: "flex", alignItems: "center", gap: "12px", padding: "10px 12px", borderRadius: "var(--border-radius-md)", border: `1px solid ${clr}30`, background: bg, cursor: "pointer", marginBottom: "8px" }}>
                          <span style={{ fontSize: "12px", fontWeight: 600, color: "var(--color-text-secondary)", width: "22px", textAlign: "right" }}>#{idx+1}</span>
                          <div style={{ width: "38px", height: "38px", borderRadius: "50%", background: clr+"22", display: "flex", alignItems: "center", justifyContent: "center", fontSize: "13px", fontWeight: 600, color: clr, flexShrink: 0 }}>
                            {s.name.split(" ").map(w=>w[0]).slice(0,2).join("").toUpperCase()}
                          </div>
                          <div style={{ flex: 1, minWidth: 0 }}>
                            <p style={{ fontWeight: 600, fontSize: "14px", margin: 0, whiteSpace: "nowrap", overflow: "hidden", textOverflow: "ellipsis" }}>{s.name}</p>
                            <p style={{ fontSize: "12px", color: "var(--color-text-secondary)", margin: "2px 0 0" }}>{s.village||"—"} · Block {s.block} · Seat {s.seatNo}</p>
                            <p style={{ fontSize: "11px", color: "var(--color-text-secondary)", margin: "1px 0 0" }}>Expires: {fmtDate(exp.toISOString().split("T")[0])}</p>
                          </div>
                          <div style={{ flexShrink: 0, textAlign: "center", background: clr, borderRadius: "var(--border-radius-md)", padding: "5px 10px", minWidth: "64px" }}>
                            <p style={{ fontSize: d<=0?13:16, fontWeight: 700, color: "white", margin: 0, lineHeight: 1.1 }}>{Math.abs(d)}</p>
                            <p style={{ fontSize: "10px", color: "white", margin: 0, opacity: 0.9 }}>{d<=0?"din pehle":"din bache"}</p>
                          </div>
                        </div>
                      );
                    })
                }
              </div>
            </div>
          </div>
        )}

        {/* Fees Due Modal */}
{showFeesDue && (
  <div style={{ position: "fixed", inset: 0, background: "rgba(0,0,0,0.5)", display: "flex", alignItems: "center", justifyContent: "center", zIndex: 200, padding: "1rem" }}>
    <div style={{ background: "var(--color-background-primary)", borderRadius: "var(--border-radius-lg)", width: "100%", maxWidth: "560px", maxHeight: "85vh", display: "flex", flexDirection: "column", border: "0.5px solid var(--color-border-secondary)" }}>
      <div style={{ padding: "1.25rem 1.5rem", borderBottom: "0.5px solid var(--color-border-secondary)", display: "flex", alignItems: "center", justifyContent: "space-between", flexShrink: 0 }}>
        <div>
          <h2 style={{ fontSize: "17px", fontWeight: 600, margin: 0 }}>
            <i className="ti ti-alert-circle" style={{ marginRight: "8px", color: "#DC2626" }}></i>
            Fees Not Deposited List
          </h2>
          <p style={{ fontSize: "12px", color: "var(--color-text-secondary)", margin: "3px 0 0" }}>
            Total {feesNotDeposited} students — jinki fees entry blank/0 hai
          </p>
        </div>
        <button onClick={() => setShowFeesDue(false)} style={{ background: "none", border: "none", cursor: "pointer", fontSize: "20px" }}>
          <i className="ti ti-x"></i>
        </button>
      </div>

      <div style={{ overflowY: "auto", flex: 1, padding: "0.75rem 1.5rem 1.25rem" }}>
        {feesDueStudents.length === 0
          ? <p style={{ textAlign: "center", color: "var(--color-text-secondary)", paddingTop: "2rem" }}>Sabhi students ne fees deposit kar di hai ✅</p>
          : feesDueStudents.map((s, idx) => (
              <div
                key={s.id}
                onClick={() => { setSelectedBlock(s.block); setSelectedSeat(s.seatNo); setShowFeesDue(false); setView("seat"); }}
                style={{
                  display: "flex",
                  alignItems: "center",
                  gap: "12px",
                  padding: "10px 12px",
                  borderRadius: "var(--border-radius-md)",
                  border: "1px solid #DC262630",
                  background: "#FEF2F2",
                  cursor: "pointer",
                  marginBottom: "8px"
                }}
              >
                <span style={{ fontSize: "12px", fontWeight: 600, color: "var(--color-text-secondary)", width: "22px", textAlign: "right" }}>#{idx+1}</span>
                <div style={{ width: "38px", height: "38px", borderRadius: "50%", background: "#DC262622", display: "flex", alignItems: "center", justifyContent: "center", fontSize: "13px", fontWeight: 600, color: "#DC2626", flexShrink: 0 }}>
                  {s.name.split(" ").map(w=>w[0]).slice(0,2).join("").toUpperCase()}
                </div>
                <div style={{ flex: 1, minWidth: 0 }}>
                  <p style={{ fontWeight: 600, fontSize: "14px", margin: 0, whiteSpace: "nowrap", overflow: "hidden", textOverflow: "ellipsis" }}>{s.name}</p>
                  <p style={{ fontSize: "12px", color: "var(--color-text-secondary)", margin: "2px 0 0" }}>
                    {s.village || "—"} · Block {s.block} · Seat {s.seatNo}
                  </p>
                  <p style={{ fontSize: "11px", color: "var(--color-text-secondary)", margin: "1px 0 0" }}>
                    Joining: {fmtDate(s.joiningDate)}
                  </p>
                </div>
                <span style={{ fontSize: "11px", padding: "4px 8px", borderRadius: "20px", background: "#DC2626", color: "white", flexShrink: 0 }}>
                  Due
                </span>
              </div>
            ))
        }
      </div>
    </div>
  </div>
)}

        {/* Settings Modal */}
        {showSettings && (
          <div style={{ position: "fixed", inset: 0, background: "rgba(0,0,0,0.45)", display: "flex", alignItems: "center", justifyContent: "center", zIndex: 200 }}>
            <div style={{ background: "var(--color-background-primary)", borderRadius: "var(--border-radius-lg)", padding: "1.5rem", width: "340px", border: "0.5px solid var(--color-border-secondary)" }}>
              <div style={{ display: "flex", alignItems: "center", justifyContent: "space-between", marginBottom: "1rem" }}>
                <h2 style={{ fontSize: "18px", fontWeight: 500, margin: 0 }}>Block Configuration</h2>
                <button onClick={() => setShowSettings(false)} style={{ background: "none", border: "none", cursor: "pointer", fontSize: "18px" }}>
                  <i className="ti ti-x"></i>
                </button>
              </div>
              {Object.keys(tempSeats).map(b => (
                <div key={b} style={{ display: "flex", alignItems: "center", gap: "12px", marginBottom: "12px" }}>
                  <label style={{ fontSize: "14px", fontWeight: 500, width: "70px" }}>Block {b}</label>
                  <input type="number" min="1" max="200" value={tempSeats[b]}
                    onChange={e => setTempSeats({ ...tempSeats, [b]: parseInt(e.target.value)||1 })}
                    style={{ flex: 1 }} />
                  <span style={{ fontSize: "13px", color: "var(--color-text-secondary)", width: "36px" }}>seats</span>
                </div>
              ))}
              <div style={{ display: "flex", gap: "8px", marginTop: "1.25rem" }}>
                <button onClick={async () => { await saveBlocks(tempSeats); setShowSettings(false); }}
                  style={{ flex: 1, background: "#1D9E75", color: "white", border: "none", borderRadius: "var(--border-radius-md)", padding: "9px", fontSize: "14px", cursor: "pointer" }}>
                  Save Changes
                </button>
                <button onClick={() => setShowSettings(false)} style={{ padding: "9px 16px", fontSize: "14px", cursor: "pointer" }}>Cancel</button>
              </div>
            </div>
          </div>
        )}
      </div>
    );
  }

  /* ══════════════════════════════════════════════════════
     BLOCK VIEW
  ══════════════════════════════════════════════════════ */
  if (view === "block") {
    const stats = blockStats(selectedBlock);
    const total = blocks[selectedBlock];
    return (
      <div style={styles.page}>
        <div style={styles.topbar}>
          <div style={{ display: "flex", alignItems: "center", gap: "12px" }}>
            <button onClick={() => setView("home")} style={styles.backBtn}>
              <i className="ti ti-arrow-left"></i> All Blocks
            </button>
            <div>
              <h1 style={styles.h1}>Block {selectedBlock}</h1>
              <p style={styles.sub}>{stats.occupied} occupied · {stats.available} available · {stats.total} total</p>
            </div>
          </div>
        </div>
        <div style={{ display: "flex", gap: "20px", marginBottom: "1.25rem", fontSize: "13px", color: "var(--color-text-secondary)" }}>
          {[["#E1F5EE","1px solid #1D9E75","Occupied"],["#FCEBEB","1px solid #E24B4A","Available"]].map(([bg,border,label])=>(
            <div key={label} style={{ display: "flex", alignItems: "center", gap: "7px" }}>
              <div style={{ width: "14px", height: "14px", borderRadius: "3px", background: bg, border }}></div>{label}
            </div>
          ))}
        </div>
        <div style={{ display: "grid", gridTemplateColumns: "repeat(auto-fill, minmax(68px, 1fr))", gap: "8px" }}>
          {Array.from({ length: total }, (_, i) => {
            const seatNo = i+1;
            const count = getSeatStudents(selectedBlock, seatNo).length;
            const occupied = count > 0;
            return (
              <div key={seatNo} onClick={() => { setSelectedSeat(seatNo); setView("seat"); }}
                style={{ aspectRatio: "1", display: "flex", flexDirection: "column", alignItems: "center", justifyContent: "center", borderRadius: "var(--border-radius-md)", cursor: "pointer", background: occupied?"#E1F5EE":"#FCEBEB", border: `1px solid ${occupied?"#1D9E75":"#E24B4A"}`, position: "relative" }}>
                <i className="ti ti-armchair-2" style={{ fontSize: "18px", color: occupied?"#085041":"#A32D2D", marginBottom: "2px" }}></i>
                <span style={{ fontSize: "13px", fontWeight: 500, color: occupied?"#085041":"#A32D2D" }}>{seatNo}</span>
                {count > 1 && (
                  <span style={{ position: "absolute", top: "-5px", right: "-5px", background: "#1D9E75", color: "white", borderRadius: "50%", fontSize: "10px", width: "16px", height: "16px", display: "flex", alignItems: "center", justifyContent: "center", fontWeight: 500 }}>{count}</span>
                )}
              </div>
            );
          })}
        </div>
      </div>
    );
  }

  /* ══════════════════════════════════════════════════════
     SEAT VIEW
  ══════════════════════════════════════════════════════ */
  if (view === "seat") {
    const seatStudents = getSeatStudents(selectedBlock, selectedSeat);
    return (
      <div style={styles.page}>
        <div style={styles.topbar}>
          <div style={{ display: "flex", alignItems: "center", gap: "12px" }}>
            <button onClick={() => setView("block")} style={styles.backBtn}>
              <i className="ti ti-arrow-left"></i> Block {selectedBlock}
            </button>
            <div>
              <h1 style={styles.h1}>Seat {selectedSeat}</h1>
              <p style={styles.sub}>Block {selectedBlock} · {seatStudents.length} student{seatStudents.length !== 1?"s":""}</p>
            </div>
          </div>
          {isAdmin && (
          <button onClick={openAdd} style={styles.primaryBtn}>
            <i className="ti ti-user-plus"></i> Add Student
          </button>
          )}
        </div>

        {seatStudents.length === 0
          ? <div style={{ textAlign: "center", padding: "3rem", border: "0.5px dashed var(--color-border-secondary)", borderRadius: "var(--border-radius-lg)", color: "var(--color-text-secondary)" }}>
              <i className="ti ti-armchair-2" style={{ fontSize: "40px", display: "block", marginBottom: "12px", color: "#1D9E75" }}></i>
              <p style={{ margin: "0 0 6px", fontSize: "15px", fontWeight: 500 }}>Yeh seat available hai</p>
              {isAdmin && <p style={{ margin: 0, fontSize: "13px" }}>Click "Add Student" to assign someone</p>}
            </div>
          : !isAdmin
            ? <div style={{ textAlign: "center", padding: "3rem", border: "0.5px solid var(--color-border-secondary)", borderRadius: "var(--border-radius-lg)", background: "var(--color-background-secondary)" }}>
                <i className="ti ti-user-check" style={{ fontSize: "40px", display: "block", marginBottom: "12px", color: "#1D9E75" }}></i>
                <p style={{ margin: "0 0 6px", fontSize: "16px", fontWeight: 500 }}>Yeh seat occupied hai</p>
                <p style={{ margin: 0, fontSize: "13px", color: "var(--color-text-secondary)" }}>Student details sirf admin dekh sakta hai</p>
              </div>
            : <div style={{ display: "flex", flexDirection: "column", gap: "12px" }}>
              {seatStudents.map(s => {
                const dLeft = getDaysLeft(s.joiningDate);
                const expClr = dLeft===null?null:dLeft<=0?"#DC2626":dLeft<=7?"#E07B2A":dLeft<=15?"#D4A017":null;
                return (
                  <div key={s.id} style={styles.card}>
                    <div style={{ display: "flex", alignItems: "flex-start", justifyContent: "space-between", marginBottom: "14px" }}>
                      <div style={{ display: "flex", alignItems: "center", gap: "12px" }}>
                        <div style={{ width: "44px", height: "44px", borderRadius: "50%", background: "var(--color-background-info)", display: "flex", alignItems: "center", justifyContent: "center", fontSize: "14px", fontWeight: 500, color: "var(--color-text-info)", flexShrink: 0 }}>
                          {s.name.split(" ").map(w=>w[0]).slice(0,2).join("").toUpperCase()}
                        </div>
                        <div>
                          <p style={{ fontWeight: 500, fontSize: "16px", margin: 0 }}>{s.name}</p>
                          <p style={{ fontSize: "13px", color: "var(--color-text-secondary)", margin: "2px 0 0" }}>
                            <i className="ti ti-map-pin" style={{ fontSize: "13px", marginRight: "4px" }}></i>{s.village||"—"}
                          </p>
                          {expClr && (
                            <span style={{ fontSize: "11px", padding: "2px 7px", borderRadius: "20px", background: expClr+"22", color: expClr, marginTop: "4px", display: "inline-block" }}>
                              {dLeft<=0?`${Math.abs(dLeft)} din se expired`:`${dLeft} din bacha`}
                            </span>
                          )}
                        </div>
                      </div>
                      <div style={{ display: "flex", gap: "6px" }}>
                        <button onClick={() => openEdit(s)}
                          style={{ fontSize: "13px", padding: "5px 10px", background: "#1D9E7522", border: "1px solid #1D9E75", borderRadius: "var(--border-radius-md)", cursor: "pointer", color: "#1D9E75" }}>
                          <i className="ti ti-edit" style={{ fontSize: "15px" }}></i> Edit
                        </button>
                        <button onClick={() => deleteStudent(s.id)}
                          style={{ fontSize: "13px", padding: "5px 10px", color: "var(--color-text-danger)", borderColor: "var(--color-border-danger)", background: "none", border: "1px solid", borderRadius: "var(--border-radius-md)", cursor: "pointer" }}>
                          <i className="ti ti-trash" style={{ fontSize: "15px" }}></i>
                        </button>
                      </div>
                    </div>
                    <div style={{ display: "grid", gridTemplateColumns: "1fr 1fr", gap: "12px 20px", borderTop: "0.5px solid var(--color-border-tertiary)", paddingTop: "14px" }}>
                      <div><p style={styles.label}>MOBILE NO.</p>
                        <p style={styles.val}>{s.mobile?<a href={`tel:${s.mobile}`} style={{ color: "var(--color-text-info)", textDecoration: "none" }}>{s.mobile}</a>:"—"}</p></div>
                      <div><p style={styles.label}>JOINING DATE</p><p style={styles.val}>{fmtDate(s.joiningDate)}</p></div>
                      <div><p style={styles.label}>FEES DEPOSITED</p>
                        <p style={{ ...styles.val, fontWeight: 500, color: "var(--color-text-success)", fontSize: "16px" }}>₹{parseFloat(s.feesDeposit||0).toLocaleString("en-IN")}</p></div>
                      <div><p style={styles.label}>FEES DATE</p><p style={styles.val}>{fmtDate(s.feesDepositDate)}</p></div>
                    </div>
                  </div>
                );
              })}
            </div>
        }
      </div>
    );
  }

  /* ══════════════════════════════════════════════════════
     FORM VIEW
  ══════════════════════════════════════════════════════ */
  if (view === "form") {
    const fields = [
      { key: "name", label: "Full Name", type: "text", required: true, placeholder: "Student ka poora naam" },
      { key: "village", label: "Village / Address", type: "text", placeholder: "Village ya sheher" },
      { key: "mobile", label: "Mobile Number", type: "tel", placeholder: "10-digit mobile number" },
      { key: "joiningDate", label: "Library Joining Date", type: "date" },
      { key: "feesDeposit", label: "Fees Deposited (₹)", type: "number", placeholder: "Amount in rupees" },
      { key: "feesDepositDate", label: "Fees Deposit Date", type: "date" },
    ];
    return (
      <div style={styles.page}>
        <div style={{ ...styles.topbar, marginBottom: "1.25rem" }}>
          <div style={{ display: "flex", alignItems: "center", gap: "12px" }}>
            <button onClick={() => { setEditId(null); setView("seat"); }} style={styles.backBtn}>
              <i className="ti ti-arrow-left"></i> Back
            </button>
            <div>
              <h1 style={styles.h1}>{editId?"Edit Student":"Add New Student"}</h1>
              <p style={styles.sub}>Block {selectedBlock} · Seat {selectedSeat}</p>
            </div>
          </div>
        </div>
        <div style={{ ...styles.card, maxWidth: "520px" }}>
          <div style={{ display: "flex", flexDirection: "column", gap: "18px" }}>
            {fields.map(f => (
              <div key={f.key}>
                <label style={{ fontSize: "13px", color: "var(--color-text-secondary)", display: "block", marginBottom: "6px" }}>
                  {f.label}{f.required && <span style={{ color: "var(--color-text-danger)", marginLeft: "3px" }}>*</span>}
                </label>
                <input type={f.type} value={formData[f.key]} placeholder={f.placeholder}
                  onChange={e => setFormData({ ...formData, [f.key]: e.target.value })}
                  style={{ width: "100%", boxSizing: "border-box" }} />
              </div>
            ))}
            <div style={{ display: "flex", gap: "10px", paddingTop: "4px", borderTop: "0.5px solid var(--color-border-tertiary)" }}>
              <button onClick={handleSubmit} disabled={saving||!formData.name.trim()}
                style={{ flex: 1, background: "#1D9E75", color: "white", border: "none", borderRadius: "var(--border-radius-md)", padding: "10px", fontSize: "15px", cursor: "pointer", opacity: saving?0.7:1 }}>
                {saving?"Saving…":editId?"Save Changes":"Add Student"}
              </button>
              <button onClick={() => { setEditId(null); setView("seat"); }} style={{ padding: "10px 18px", fontSize: "14px", cursor: "pointer" }}>
                Cancel
              </button>
            </div>
          </div>
        </div>
      </div>
    );
  }

  return null;
}

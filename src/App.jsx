import { useCallback, useEffect, useMemo, useState } from "react";
import { recomputeRunningBalances } from "./lib/ledger";
import {
  PURPOSE_KEYS,
  resolvePurposeLabel,
  purposeOptionLabel,
  formatDisplayDate,
} from "./lib/purposes";
import {
  fetchStudents,
  fetchSoaRows,
  ensureCurrentSchoolYear,
  insertStudent,
  insertSoaRow,
  insertSoaRowsBulk,
  deleteStudent,
  deleteSoaRow,
  archiveSchoolYearAndStartNew,
  formatSchoolYearLabel,
  normalizeSchoolYearLabel,
  buildDescription,
  baselinePayloads,
} from "./lib/supabaseApi";
import { SoaPrintBundle } from "./components/SoaPrintPages";

const ADMIN_PASSWORD = import.meta.env.VITE_ADMIN_PASSWORD ?? "";
const FEE_SETTINGS_KEY = "pjlc_fee_settings_v1";
const AUTH_SESSION_KEY = "pjlc_app_unlocked";
const SCHOOL_YEAR_SESSION_KEY = "pjlc_selected_school_year";

const DEFAULT_GRADE_LEVELS = [
  { id: "k1", name: "Kinder 1", generalFees: 9650, tuition: 15000, books: 4200, grandTotal: 28850 },
  { id: "k2", name: "Kinder 2", generalFees: 9650, tuition: 15000, books: 4200, grandTotal: 28850 },
  { id: "g1", name: "Grade 1", generalFees: 11100, tuition: 15000, books: 5800, grandTotal: 31900 },
  { id: "g2", name: "Grade 2", generalFees: 11100, tuition: 15000, books: 5800, grandTotal: 31900 },
  { id: "g3", name: "Grade 3", generalFees: 14000, tuition: 15000, books: 5800, grandTotal: 34800 },
  { id: "g4", name: "Grade 4", generalFees: 14000, tuition: 15000, books: 6500, grandTotal: 35500 },
  { id: "g5", name: "Grade 5", generalFees: 14000, tuition: 15000, books: 6500, grandTotal: 35500 },
  { id: "g6", name: "Grade 6", generalFees: 14000, tuition: 15000, books: 6500, grandTotal: 35500 },
];

const currency = (value) =>
  new Intl.NumberFormat("en-PH", { style: "currency", currency: "PHP" }).format(
    Number(value || 0),
  );

const today = () => new Date().toISOString().slice(0, 10);

const normalizeGrade = (grade) => {
  const generalFees = Number(grade.generalFees) || 0;
  const tuition = Number(grade.tuition) || 0;
  const books = Number(grade.books) || 0;
  return {
    ...grade,
    generalFees,
    tuition,
    books,
    grandTotal: generalFees + tuition + books,
  };
};

const makeGradeLevels = (grades) =>
  grades.map((g) => ({
    gradeLevelId: g.id,
    gradeLevelName: g.name,
    baseGeneralFees: g.generalFees,
    baseTuition: g.tuition,
    baseBooks: g.books,
    baseGrandTotal: g.grandTotal,
  }));

const escapeHtml = (value) =>
  String(value ?? "")
    .replaceAll("&", "&amp;")
    .replaceAll("<", "&lt;")
    .replaceAll(">", "&gt;")
    .replaceAll('"', "&quot;");

const workbookSheet = (name, headers, rows) => `
  <h2>${escapeHtml(name)}</h2>
  <table>
    <thead><tr>${headers.map((h) => `<th>${escapeHtml(h)}</th>`).join("")}</tr></thead>
    <tbody>
      ${rows
        .map((row) => `<tr>${row.map((cell) => `<td>${escapeHtml(cell)}</td>`).join("")}</tr>`)
        .join("")}
    </tbody>
  </table>
`;

const downloadWorkbook = ({
  students,
  ledgerTransactions,
  gradeLevels,
  balanceMap,
  getGradeName,
  schoolYearLabel,
}) => {
  const studentRows = students.map((student) => [
    student.studentId,
    student.studentName,
    getGradeName(student.gradeLevelId),
    balanceMap[student.studentId] ?? 0,
  ]);

  const transactionRows = ledgerTransactions.map((tx) => {
    const student = students.find((s) => s.studentId === tx.studentId);
    return [
      tx.transactionId,
      tx.date,
      student?.studentName ?? "",
      getGradeName(student?.gradeLevelId),
      tx.orNumber,
      tx.type,
      tx.purpose,
      tx.amount,
      tx.signedAmount ?? "",
      tx.runningBalance ?? "",
      tx.createdAt,
    ];
  });

  const balanceRows = students.map((student) => {
    const ledger = recomputeRunningBalances(
      ledgerTransactions.filter((tx) => tx.studentId === student.studentId),
    );
    const debits = ledger
      .filter((tx) => tx.type === "DEBIT")
      .reduce((sum, tx) => sum + Number(tx.amount || 0), 0);
    const credits = ledger
      .filter((tx) => tx.type === "CREDIT")
      .reduce((sum, tx) => sum + Number(tx.amount || 0), 0);
    return [
      student.studentName,
      getGradeName(student.gradeLevelId),
      debits,
      credits,
      ledger.at(-1)?.runningBalance ?? 0,
    ];
  });

  const feeRows = gradeLevels.map((grade) => [
    grade.gradeLevelName,
    grade.baseGeneralFees,
    grade.baseTuition,
    grade.baseBooks,
    grade.baseGrandTotal,
  ]);

  const generatedAt = new Date().toLocaleString("en-PH");
  const html = `<!doctype html>
<html xmlns:o="urn:schemas-microsoft-com:office:office"
  xmlns:x="urn:schemas-microsoft-com:office:excel"
  xmlns="http://www.w3.org/TR/REC-html40">
<head>
  <meta charset="utf-8" />
  <style>
    body { font-family: Calibri, Arial, sans-serif; }
    table { border-collapse: collapse; margin-bottom: 28px; }
    th { background: #0f172a; color: white; font-weight: 700; }
    th, td { border: 1px solid #cbd5e1; padding: 6px 8px; mso-number-format:"\\@"; }
    h1 { color: #0f172a; }
  </style>
</head>
<body>
  <h1>Palawan Jewels Learning Center - Full Backup</h1>
  <table>
    <tbody>
      <tr><th>Generated</th><td>${escapeHtml(generatedAt)}</td></tr>
      <tr><th>School year</th><td>${escapeHtml(schoolYearLabel)}</td></tr>
      <tr><th>Student count</th><td>${students.length}</td></tr>
      <tr><th>Transaction count</th><td>${ledgerTransactions.length}</td></tr>
    </tbody>
  </table>
  ${workbookSheet("Students", ["Student ID", "Student Name", "Grade", "Current Balance"], studentRows)}
  ${workbookSheet(
    "Transactions",
    [
      "Transaction ID",
      "Date",
      "Student",
      "Grade",
      "OR #",
      "Type",
      "Purpose",
      "Amount",
      "Signed Amount",
      "Running Balance",
      "Created At",
    ],
    transactionRows,
  )}
  ${workbookSheet(
    "Student Balances",
    ["Student", "Grade", "Total Debits", "Total Credits", "Current Balance"],
    balanceRows,
  )}
  ${workbookSheet(
    "Fee Matrix",
    ["Grade", "General Fees", "Tuition", "Books", "Grand Total"],
    feeRows,
  )}
</body>
</html>`;

  const blob = new Blob([html], { type: "application/vnd.ms-excel;charset=utf-8" });
  const url = URL.createObjectURL(blob);
  const link = document.createElement("a");
  link.href = url;
  link.download = `PJLC-${schoolYearLabel || "school-year"}-backup-${today()}.xls`;
  document.body.appendChild(link);
  link.click();
  document.body.removeChild(link);
  URL.revokeObjectURL(url);
};

function StudentsTable({
  students,
  selectedStudentId,
  onSelect,
  balanceMap,
  getGradeName,
  maxHeight = "max-h-80",
  onDelete,
  isSaving,
}) {
  return (
    <div className={`${maxHeight} overflow-auto rounded-md border border-slate-200`}>
      <table className="min-w-full text-sm">
        <thead className="sticky top-0 z-10 bg-slate-100 shadow-sm">
          <tr>
            <th className="px-3 py-2 text-left font-semibold">Student name</th>
            <th className="px-3 py-2 text-left font-semibold">Grade</th>
            <th className="px-3 py-2 text-right font-semibold">Balance</th>
            {onDelete && <th className="px-3 py-2 text-right font-semibold"> </th>}
          </tr>
        </thead>
        <tbody>
          {students.length === 0 ? (
            <tr>
              <td colSpan={onDelete ? 4 : 3} className="px-3 py-6 text-center text-slate-500">
                No students found
              </td>
            </tr>
          ) : (
            students.map((student) => {
              const selected = selectedStudentId === student.studentId;
              return (
                <tr
                  key={student.studentId}
                  onClick={() => onSelect(student.studentId)}
                  className={`cursor-pointer border-t border-slate-200 ${
                    selected ? "bg-slate-900 text-white" : "hover:bg-slate-50"
                  }`}
                >
                  <td className="px-3 py-2 font-medium">{student.studentName}</td>
                  <td className={`px-3 py-2 ${selected ? "text-slate-200" : "text-slate-600"}`}>
                    {getGradeName(student.gradeLevelId)}
                  </td>
                  <td
                    className={`px-3 py-2 text-right font-semibold ${selected ? "" : "text-rose-700"}`}
                  >
                    {currency(balanceMap[student.studentId])}
                  </td>
                  {onDelete && (
                    <td className="px-3 py-2 text-right">
                      <button
                        type="button"
                        disabled={isSaving}
                        onClick={(e) => {
                          e.stopPropagation();
                          onDelete(student);
                        }}
                        className="rounded bg-rose-600 px-2 py-0.5 text-xs text-white hover:bg-rose-700 disabled:opacity-50"
                      >
                        Delete
                      </button>
                    </td>
                  )}
                </tr>
              );
            })
          )}
        </tbody>
      </table>
    </div>
  );
}

function App() {
  const [isAppUnlocked, setIsAppUnlocked] = useState(
    () => sessionStorage.getItem(AUTH_SESSION_KEY) === "true",
  );
  const [loginPasswordInput, setLoginPasswordInput] = useState("");
  const [loginMessage, setLoginMessage] = useState("");
  const [feeSettings, setFeeSettings] = useState(() => {
    try {
      const stored = JSON.parse(localStorage.getItem(FEE_SETTINGS_KEY) || "null");
      if (Array.isArray(stored)) {
        return DEFAULT_GRADE_LEVELS.map((base) =>
          normalizeGrade({ ...base, ...(stored.find((g) => g.id === base.id) ?? {}) }),
        );
      }
    } catch {
      localStorage.removeItem(FEE_SETTINGS_KEY);
    }
    return DEFAULT_GRADE_LEVELS.map(normalizeGrade);
  });
  const [students, setStudents] = useState([]);
  const [ledgerTransactions, setLedgerTransactions] = useState([]);
  const [isLoading, setIsLoading] = useState(true);
  const [isSaving, setIsSaving] = useState(false);
  const [cloudMessage, setCloudMessage] = useState("");
  const [loadError, setLoadError] = useState("");
  const [schoolYears, setSchoolYears] = useState([]);
  const [activeSchoolYearId, setActiveSchoolYearId] = useState(
    () => sessionStorage.getItem(SCHOOL_YEAR_SESSION_KEY) || "",
  );

  const [activeTab, setActiveTab] = useState("ledger");
  const [studentForm, setStudentForm] = useState({ studentName: "", gradeLevelId: "k1" });
  const [searchTerm, setSearchTerm] = useState("");
  const [selectedStudentId, setSelectedStudentId] = useState("");
  const [txForm, setTxForm] = useState({
    date: today(),
    orNumber: "",
    type: "CREDIT",
    amount: "",
    purposeKey: "monthly_payment",
    customPurpose: "",
  });
  const [dateFilter, setDateFilter] = useState({ startDate: "", endDate: "" });
  const [isPrintModalOpen, setIsPrintModalOpen] = useState(false);
  const [billToInput, setBillToInput] = useState("");
  const [billToPrinted, setBillToPrinted] = useState("");
  const [printMode, setPrintMode] = useState("single");
  const [printGradeId, setPrintGradeId] = useState("");
  const [printDateRange, setPrintDateRange] = useState({ startDate: "", endDate: "" });
  const [tickets, setTickets] = useState([]);
  const [selectedPrintStudentIds, setSelectedPrintStudentIds] = useState([]);
  const [adminPasswordInput, setAdminPasswordInput] = useState("");
  const [isAdminUnlocked, setIsAdminUnlocked] = useState(false);
  const [adminMessage, setAdminMessage] = useState("");
  const [archiveForm, setArchiveForm] = useState({
    startYear: "",
    endYear: "",
    firstPassword: "",
    secondPassword: "",
  });

  const gradeLevels = useMemo(() => makeGradeLevels(feeSettings), [feeSettings]);
  const getGradeById = useCallback(
    (id) => feeSettings.find((g) => g.id === id),
    [feeSettings],
  );

  const activeSchoolYear = useMemo(
    () => schoolYears.find((year) => year.schoolYearId === activeSchoolYearId) ?? schoolYears[0],
    [activeSchoolYearId, schoolYears],
  );

  const loadData = useCallback(async (targetSchoolYearId = activeSchoolYearId) => {
    setIsLoading(true);
    setLoadError("");
    setCloudMessage("Loading from cloud…");
    try {
      const { current, years } = await ensureCurrentSchoolYear();
      const selectedYearId =
        targetSchoolYearId && years.some((year) => year.schoolYearId === targetSchoolYearId)
          ? targetSchoolYearId
          : current.schoolYearId;
      const [studentRows, soaRows] = await Promise.all([
        fetchStudents(selectedYearId),
        fetchSoaRows(selectedYearId),
      ]);
      setSchoolYears(years);
      setActiveSchoolYearId(selectedYearId);
      sessionStorage.setItem(SCHOOL_YEAR_SESSION_KEY, selectedYearId);
      setStudents(studentRows);
      setLedgerTransactions(soaRows);
      setCloudMessage("Synced with Supabase.");
    } catch (err) {
      const msg = err.message ?? "Could not load data from Supabase.";
      const needsLatestMigration =
        msg.includes("school_years") ||
        msg.includes("school_year_id") ||
        msg.includes("schema cache");
      const hint = needsLatestMigration
        ? " Run the latest supabase/migrate-soa_rows.sql in the Supabase SQL Editor, then click Refresh."
        : msg.includes("soa_rows")
          ? " Run supabase/migrate-soa_rows.sql in the Supabase SQL Editor, then click Refresh."
          : "";
      setLoadError(msg + hint);
      setCloudMessage("");
    } finally {
      setIsLoading(false);
    }
  }, [activeSchoolYearId]);

  useEffect(() => {
    if (!isAppUnlocked) return;
    // eslint-disable-next-line react-hooks/set-state-in-effect
    loadData();
  }, [isAppUnlocked, loadData]);

  const sortedStudents = useMemo(
    () => [...students].sort((a, b) => a.studentName.localeCompare(b.studentName)),
    [students],
  );

  const effectiveSelectedStudentId = selectedStudentId || sortedStudents[0]?.studentId || "";
  const selectedStudent = sortedStudents.find((s) => s.studentId === effectiveSelectedStudentId);

  const selectedStudentLedger = useMemo(() => {
    return recomputeRunningBalances(
      ledgerTransactions.filter((tx) => tx.studentId === effectiveSelectedStudentId),
    );
  }, [effectiveSelectedStudentId, ledgerTransactions]);

  const filteredStudents = useMemo(() => {
    const key = searchTerm.trim().toLowerCase();
    if (!key) return sortedStudents;
    return sortedStudents.filter((s) => s.studentName.toLowerCase().includes(key));
  }, [searchTerm, sortedStudents]);

  const studentBalanceMap = useMemo(() => {
    const map = {};
    students.forEach((student) => {
      const ledger = recomputeRunningBalances(
        ledgerTransactions.filter((tx) => tx.studentId === student.studentId),
      );
      map[student.studentId] = ledger.at(-1)?.runningBalance ?? 0;
    });
    return map;
  }, [ledgerTransactions, students]);

  const systemTransactionsByDate = useMemo(() => {
    return recomputeRunningBalances(ledgerTransactions).filter((tx) => {
      const passStart = !dateFilter.startDate || tx.date >= dateFilter.startDate;
      const passEnd = !dateFilter.endDate || tx.date <= dateFilter.endDate;
      return passStart && passEnd;
    });
  }, [dateFilter.endDate, dateFilter.startDate, ledgerTransactions]);

  const purposePreview =
    txForm.purposeKey === "other"
      ? txForm.customPurpose.trim() || "Custom purpose"
      : resolvePurposeLabel(txForm.purposeKey, txForm.date);

  const runSave = async (label, fn) => {
    setIsSaving(true);
    setCloudMessage(label);
    try {
      const targetSchoolYearId = await fn();
      await loadData(targetSchoolYearId);
    } catch (err) {
      setCloudMessage(`Error: ${err.message ?? "Save failed"}`);
    } finally {
      setIsSaving(false);
    }
  };

  const handleSchoolYearChange = (schoolYearId) => {
    sessionStorage.setItem(SCHOOL_YEAR_SESSION_KEY, schoolYearId);
    setActiveSchoolYearId(schoolYearId);
    setSelectedStudentId("");
    setSelectedPrintStudentIds([]);
  };

  const handleLogin = (event) => {
    event.preventDefault();
    if (!ADMIN_PASSWORD) {
      setLoginMessage("Missing VITE_ADMIN_PASSWORD. Add it to .env.local and restart the app.");
      return;
    }
    if (loginPasswordInput === ADMIN_PASSWORD) {
      sessionStorage.setItem(AUTH_SESSION_KEY, "true");
      setIsAppUnlocked(true);
      setLoginPasswordInput("");
      setLoginMessage("");
      return;
    }
    setLoginMessage("Incorrect password.");
  };

  const handleLogout = () => {
    sessionStorage.removeItem(AUTH_SESSION_KEY);
    setIsAppUnlocked(false);
    setIsAdminUnlocked(false);
    setAdminPasswordInput("");
    setLoginPasswordInput("");
    setLoginMessage("");
  };

  const handleEnrollStudent = (event) => {
    event.preventDefault();
    if (!activeSchoolYear) return;
    if (!studentForm.studentName.trim()) return;

    runSave("Saving student to cloud…", async () => {
      const grade = getGradeById(studentForm.gradeLevelId);
      const created = await insertStudent({
        schoolYearId: activeSchoolYear.schoolYearId,
        studentName: studentForm.studentName.trim(),
        gradeLevelId: studentForm.gradeLevelId,
      });
      const dateIso = today();
      const baseline = baselinePayloads(created.studentId, grade, dateIso).map((row) => ({
        ...row,
        schoolYearId: activeSchoolYear.schoolYearId,
      }));
      await insertSoaRowsBulk(baseline);
      setStudentForm({ studentName: "", gradeLevelId: studentForm.gradeLevelId });
      setSelectedStudentId(created.studentId);
      setActiveTab("ledger");
    });
  };

  const handleAddTransaction = (event) => {
    event.preventDefault();
    if (!activeSchoolYear) return;
    if (!effectiveSelectedStudentId || !txForm.amount || Number(txForm.amount) <= 0) return;
    if (txForm.purposeKey !== "other" && !PURPOSE_KEYS[txForm.purposeKey]) return;
    if (txForm.purposeKey === "other" && !txForm.customPurpose.trim()) return;

    const purpose =
      txForm.purposeKey === "other"
        ? txForm.customPurpose.trim()
        : resolvePurposeLabel(txForm.purposeKey, txForm.date);

    runSave("Posting entry to cloud…", async () => {
      await insertSoaRow({
        schoolYearId: activeSchoolYear.schoolYearId,
        studentId: effectiveSelectedStudentId,
        date: txForm.date,
        description: buildDescription(purpose, txForm.orNumber),
        amount: Number(txForm.amount),
        entryType: txForm.type,
        orNumber: txForm.orNumber.trim(),
        purposeKey: txForm.purposeKey,
      });
      setTxForm((old) => ({ ...old, amount: "", orNumber: "", customPurpose: "", date: today() }));
    });
  };

  const handleDeleteStudent = (student) => {
    if (
      !window.confirm(
        `Delete ${student.studentName} and all ledger entries? This cannot be undone.`,
      )
    ) {
      return;
    }
    runSave("Deleting student from cloud…", async () => {
      await deleteStudent(student.studentId);
      if (selectedStudentId === student.studentId) setSelectedStudentId("");
    });
  };

  const handleDeleteTransaction = (tx) => {
    if (!window.confirm("Delete this ledger entry?")) return;
    runSave("Deleting entry from cloud…", async () => {
      await deleteSoaRow(tx.transactionId);
    });
  };

  const handleAdminUnlock = (event) => {
    event.preventDefault();
    if (!ADMIN_PASSWORD) {
      setAdminMessage("Missing VITE_ADMIN_PASSWORD. Add it to .env.local and restart the app.");
      return;
    }
    if (adminPasswordInput === ADMIN_PASSWORD) {
      setIsAdminUnlocked(true);
      setAdminPasswordInput("");
      setAdminMessage("Admin access unlocked.");
    } else {
      setAdminMessage("Incorrect password.");
    }
  };

  const updateFeeSetting = (gradeId, field, value) => {
    const nextAmount = Math.max(0, Number(value) || 0);
    setFeeSettings((old) =>
      old.map((grade) =>
        grade.id === gradeId ? normalizeGrade({ ...grade, [field]: nextAmount }) : grade,
      ),
    );
  };

  const saveFeeSettings = () => {
    localStorage.setItem(FEE_SETTINGS_KEY, JSON.stringify(feeSettings));
    setAdminMessage("Prices saved. New enrollments will use these amounts.");
  };

  const resetFeeSettings = () => {
    const defaults = DEFAULT_GRADE_LEVELS.map(normalizeGrade);
    setFeeSettings(defaults);
    localStorage.setItem(FEE_SETTINGS_KEY, JSON.stringify(defaults));
    setAdminMessage("Prices reset to defaults for future enrollments.");
  };

  const handleArchiveAllStudents = () => {
    const startYear = archiveForm.startYear.trim();
    const endYear = archiveForm.endYear.trim();
    if (
      archiveForm.firstPassword !== ADMIN_PASSWORD ||
      archiveForm.secondPassword !== ADMIN_PASSWORD
    ) {
      setAdminMessage("Enter the admin password in both archive confirmation fields.");
      return;
    }
    if (!/^\d{4}$/.test(startYear) || !/^\d{4}$/.test(endYear)) {
      setAdminMessage("Enter both school years as 4-digit years.");
      return;
    }
    if (Number(endYear) <= Number(startYear)) {
      setAdminMessage("The ending school year must be after the starting school year.");
      return;
    }
    if (!activeSchoolYear) {
      setAdminMessage("No active school year is selected.");
      return;
    }
    if (!activeSchoolYear.isCurrent) {
      setAdminMessage("Switch to the current school year before archiving all records.");
      return;
    }

    const archiveLabel = formatSchoolYearLabel(startYear, endYear);
    const duplicateYear = schoolYears.find(
      (year) =>
        year.schoolYearId !== activeSchoolYear.schoolYearId &&
        normalizeSchoolYearLabel(year.label) === normalizeSchoolYearLabel(archiveLabel),
    );
    if (duplicateYear) {
      setAdminMessage(`${archiveLabel} already exists. Enter a different school year.`);
      return;
    }

    if (
      !window.confirm(
        `Archive all records in ${activeSchoolYear.label} as ${archiveLabel} and start a fresh current school year?`,
      )
    ) {
      return;
    }

    runSave("Archiving current school year and starting a fresh one…", async () => {
      const nextYear = await archiveSchoolYearAndStartNew({
        schoolYearId: activeSchoolYear.schoolYearId,
        archiveLabel,
      });
      sessionStorage.setItem(SCHOOL_YEAR_SESSION_KEY, nextYear.schoolYearId);
      setActiveSchoolYearId(nextYear.schoolYearId);
      setSelectedStudentId("");
      setSelectedPrintStudentIds([]);
      setArchiveForm({ startYear: "", endYear: "", firstPassword: "", secondPassword: "" });
      return nextYear.schoolYearId;
    });
    setAdminMessage("Archive request sent.");
  };

  const togglePrintStudent = (studentId) => {
    setSelectedPrintStudentIds((old) =>
      old.includes(studentId) ? old.filter((id) => id !== studentId) : [...old, studentId],
    );
  };

  const printSelectedStudents = () => {
    const selectedIds = new Set(selectedPrintStudentIds);
    const allTickets = [];

    sortedStudents
      .filter((student) => selectedIds.has(student.studentId))
      .forEach((student) => {
        const ledger = recomputeRunningBalances(
          ledgerTransactions.filter((tx) => tx.studentId === student.studentId),
        );
        if (ledger.length === 0) return;
        allTickets.push({
          studentId: student.studentId,
          billTo: student.studentName,
          studentName: student.studentName,
          ledger,
          amountDue: ledger.at(-1)?.runningBalance ?? 0,
        });
      });

    if (allTickets.length === 0) return;
    setBillToPrinted("");
    setTickets(allTickets);
    setTimeout(() => window.print(), 10);
  };

  const openPrintModal = ({ mode = "single", gradeId = "", dateRange = null, billTo = "" } = {}) => {
    setPrintMode(mode);
    setPrintGradeId(gradeId);
    if (dateRange) setPrintDateRange(dateRange);
    else if (mode === "daterange") {
      setPrintDateRange({ startDate: dateFilter.startDate, endDate: dateFilter.endDate });
    }
    setBillToInput(billTo);
    setIsPrintModalOpen(true);
  };

  const startPrint = () => {
    const billToValue = billToInput.trim();
    const allTickets = [];

    const pushTicket = (student, ledger) => {
      if (ledger.length === 0) return;
      allTickets.push({
        studentId: student.studentId,
        billTo: billToValue || student.studentName,
        studentName: student.studentName,
        ledger,
        amountDue: ledger.at(-1)?.runningBalance ?? 0,
      });
    };

    if (printMode === "single" && selectedStudent) {
      pushTicket(selectedStudent, selectedStudentLedger);
    }

    if (printMode === "grade") {
      const targetGradeId = printGradeId || selectedStudent?.gradeLevelId;
      if (targetGradeId) {
        sortedStudents
          .filter((s) => s.gradeLevelId === targetGradeId)
          .forEach((student) => {
            const ledger = recomputeRunningBalances(
              ledgerTransactions.filter((tx) => tx.studentId === student.studentId),
            );
            pushTicket(student, ledger);
          });
      }
    }

    if (printMode === "daterange") {
      sortedStudents.forEach((student) => {
        const ledger = recomputeRunningBalances(
          ledgerTransactions.filter((tx) => {
            if (tx.studentId !== student.studentId) return false;
            const passStart = !printDateRange.startDate || tx.date >= printDateRange.startDate;
            const passEnd = !printDateRange.endDate || tx.date <= printDateRange.endDate;
            return passStart && passEnd;
          }),
        );
        pushTicket(student, ledger);
      });
    }

    if (allTickets.length === 0) return;

    setBillToPrinted(billToValue);
    setTickets(allTickets);
    setIsPrintModalOpen(false);
    setTimeout(() => window.print(), 10);
  };

  const documentDate = new Date().toLocaleDateString("en-PH", {
    month: "numeric",
    day: "numeric",
    year: "numeric",
  });

  const getGradeName = (id) => getGradeById(id)?.name ?? "—";

  if (!isAppUnlocked) {
    return (
      <div className="mx-auto flex min-h-screen max-w-md items-center p-6">
        <div className="w-full rounded-xl border border-slate-300 bg-white p-6 shadow-sm">
          <h1 className="text-2xl font-bold text-slate-900">
            Palawan Jewels Learning Center
          </h1>
          <p className="mt-1 text-sm text-slate-600">Enter the system password to continue.</p>
          <form className="mt-5 space-y-4" onSubmit={handleLogin}>
            <label className="flex flex-col gap-1">
              <span className="text-sm font-medium text-slate-700">Password</span>
              <input
                type="password"
                autoFocus
                className="rounded-md border border-slate-300 px-3 py-2"
                value={loginPasswordInput}
                onChange={(e) => setLoginPasswordInput(e.target.value)}
                required
              />
            </label>
            {loginMessage && <p className="text-sm text-rose-700">{loginMessage}</p>}
            <button
              type="submit"
              className="w-full rounded-md bg-slate-900 px-4 py-2 font-semibold text-white hover:bg-slate-700"
            >
              Login
            </button>
          </form>
        </div>
      </div>
    );
  }

  return (
    <div className="mx-auto max-w-7xl p-6">
      <div className="no-print rounded-xl border border-slate-300 bg-white p-5 shadow-sm">
        <h1 className="text-2xl font-bold text-slate-900">
          Palawan Jewels Learning Center — Student Ledger & SOA
        </h1>
        <p className="mt-1 text-sm text-slate-600">
          {activeSchoolYear?.label ?? "School year loading"} · Cloud (Supabase) · Two half-page
          SOAs per bond sheet
        </p>

        {(isLoading || isSaving || cloudMessage || loadError) && (
          <div
            className={`mt-3 rounded-md px-3 py-2 text-sm ${
              loadError
                ? "border border-rose-200 bg-rose-50 text-rose-800"
                : "border border-slate-200 bg-slate-50 text-slate-700"
            }`}
          >
            {loadError && <p>{loadError}</p>}
            {isLoading && !loadError && <p>Loading students and ledger from Supabase…</p>}
            {isSaving && !isLoading && <p>{cloudMessage || "Saving to cloud…"}</p>}
            {!isLoading && !isSaving && cloudMessage && !loadError && <p>{cloudMessage}</p>}
          </div>
        )}

        <div className="mt-5 flex flex-wrap items-center gap-2">
          <label className="flex items-center gap-2 rounded-lg bg-slate-100 px-3 py-2 text-sm font-semibold text-slate-700">
            <span>School Year</span>
            <select
              className="rounded-md border border-slate-300 bg-white px-2 py-1 text-sm font-medium text-slate-900"
              value={activeSchoolYearId}
              onChange={(e) => handleSchoolYearChange(e.target.value)}
              disabled={isLoading || isSaving || schoolYears.length === 0}
            >
              {schoolYears.map((year) => (
                <option key={year.schoolYearId} value={year.schoolYearId}>
                  {year.label}
                  {year.isCurrent ? " (Current)" : ""}
                </option>
              ))}
            </select>
          </label>
          {[
            { id: "enroll", label: "Enroll Student" },
            { id: "ledger", label: "Ledger & Transactions" },
            { id: "manage", label: "Reports & Print" },
            { id: "admin", label: "Admin" },
          ].map((tab) => (
            <button
              key={tab.id}
              type="button"
              disabled={isLoading}
              onClick={() => setActiveTab(tab.id)}
              className={`rounded-lg px-4 py-2 text-sm font-semibold ${
                activeTab === tab.id
                  ? "bg-slate-900 text-white"
                  : "bg-slate-100 text-slate-700 hover:bg-slate-200"
              }`}
            >
              {tab.label}
            </button>
          ))}
          <button
            type="button"
            disabled={isLoading || isSaving}
            onClick={loadData}
            className="rounded-lg bg-slate-200 px-4 py-2 text-sm font-semibold text-slate-800 hover:bg-slate-300"
          >
            Refresh
          </button>
          <button
            type="button"
            onClick={handleLogout}
            className="rounded-lg bg-slate-100 px-4 py-2 text-sm font-semibold text-slate-700 hover:bg-slate-200"
          >
            Logout
          </button>
        </div>

        {activeTab === "enroll" && (
          <section className="mt-6 rounded-lg border border-slate-200 p-4">
            <h2 className="text-lg font-semibold">Enroll New Student</h2>
            <p className="mt-1 text-sm text-slate-600">
              Saves to Supabase and auto-posts General Fees, Tuition, and Books from the fee matrix.
            </p>
            <form className="mt-4 grid gap-4 md:grid-cols-3" onSubmit={handleEnrollStudent}>
              <label className="flex flex-col gap-1">
                <span className="text-sm font-medium text-slate-700">Student Name</span>
                <input
                  className="rounded-md border border-slate-300 px-3 py-2"
                  value={studentForm.studentName}
                  onChange={(e) =>
                    setStudentForm((old) => ({ ...old, studentName: e.target.value }))
                  }
                  required
                  disabled={isSaving}
                />
              </label>
              <label className="flex flex-col gap-1">
                <span className="text-sm font-medium text-slate-700">Grade Level</span>
                <select
                  className="rounded-md border border-slate-300 px-3 py-2"
                  value={studentForm.gradeLevelId}
                  onChange={(e) =>
                    setStudentForm((old) => ({ ...old, gradeLevelId: e.target.value }))
                  }
                  disabled={isSaving}
                >
                  {gradeLevels.map((grade) => (
                    <option key={grade.gradeLevelId} value={grade.gradeLevelId}>
                      {grade.gradeLevelName}
                    </option>
                  ))}
                </select>
              </label>
              <div className="flex items-end">
                <button
                  type="submit"
                  disabled={!activeSchoolYear || isSaving || isLoading}
                  className="w-full rounded-md bg-emerald-700 px-4 py-2 font-semibold text-white hover:bg-emerald-800 disabled:opacity-50"
                >
                  Save Student + Post Baseline Fees
                </button>
              </div>
            </form>
          </section>
        )}

        {activeTab === "ledger" && (
          <section className="mt-6 space-y-6">
            <div className="rounded-lg border border-slate-200 p-4">
              <h2 className="text-lg font-semibold">All Students (A–Z)</h2>
              <input
                className="mt-3 w-full max-w-md rounded-md border border-slate-300 px-3 py-2 text-sm"
                placeholder="Search by name..."
                value={searchTerm}
                onChange={(e) => setSearchTerm(e.target.value)}
              />
              <div className="mt-3">
                <StudentsTable
                  students={filteredStudents}
                  selectedStudentId={effectiveSelectedStudentId}
                  onSelect={setSelectedStudentId}
                  balanceMap={studentBalanceMap}
                  getGradeName={getGradeName}
                  onDelete={handleDeleteStudent}
                  isSaving={isSaving}
                />
              </div>
            </div>

            <div className="rounded-lg border border-slate-200 p-4">
              <div className="flex flex-wrap items-center justify-between gap-2">
                <h2 className="text-lg font-semibold">
                  Ledger — {selectedStudent?.studentName ?? "Select a student"}
                </h2>
                <button
                  type="button"
                  onClick={() =>
                    openPrintModal({
                      mode: "single",
                      billTo: selectedStudent?.studentName ?? "",
                    })
                  }
                  disabled={!selectedStudent || isLoading}
                  className="rounded-md bg-indigo-700 px-4 py-2 text-sm font-semibold text-white hover:bg-indigo-800 disabled:opacity-50"
                >
                  Print SOA
                </button>
              </div>

              <form
                className="mt-4 grid gap-3 sm:grid-cols-2 lg:grid-cols-6"
                onSubmit={handleAddTransaction}
              >
                <label className="flex flex-col gap-1">
                  <span className="text-xs font-medium text-slate-600">Date</span>
                  <input
                    className="rounded-md border border-slate-300 px-3 py-2 text-sm"
                    type="date"
                    value={txForm.date}
                    onChange={(e) => setTxForm((old) => ({ ...old, date: e.target.value }))}
                    required
                    disabled={isSaving}
                  />
                </label>
                <label className="flex flex-col gap-1">
                  <span className="text-xs font-medium text-slate-600">OR Number</span>
                  <input
                    className="rounded-md border border-slate-300 px-3 py-2 text-sm"
                    placeholder="e.g. OR-10337"
                    value={txForm.orNumber}
                    onChange={(e) =>
                      setTxForm((old) => ({ ...old, orNumber: e.target.value }))
                    }
                    disabled={isSaving}
                  />
                </label>
                <label className="flex flex-col gap-1">
                  <span className="text-xs font-medium text-slate-600">Entry type</span>
                  <select
                    className="rounded-md border border-slate-300 px-3 py-2 text-sm"
                    value={txForm.type}
                    onChange={(e) => setTxForm((old) => ({ ...old, type: e.target.value }))}
                    disabled={isSaving}
                  >
                    <option value="DEBIT">Debit (add to balance)</option>
                    <option value="CREDIT">Credit (subtract from balance)</option>
                  </select>
                </label>
                <label className="flex flex-col gap-1">
                  <span className="text-xs font-medium text-slate-600">Amount (PHP)</span>
                  <input
                    className="rounded-md border border-slate-300 px-3 py-2 text-sm"
                    type="number"
                    min="0.01"
                    step="0.01"
                    placeholder="0.00"
                    value={txForm.amount}
                    onChange={(e) => setTxForm((old) => ({ ...old, amount: e.target.value }))}
                    required
                    disabled={isSaving}
                  />
                </label>
                <label className="flex flex-col gap-1">
                  <span className="text-xs font-medium text-slate-600">Purpose (label only)</span>
                  <select
                    className="rounded-md border border-slate-300 px-3 py-2 text-sm"
                    value={txForm.purposeKey}
                    onChange={(e) =>
                      setTxForm((old) => ({ ...old, purposeKey: e.target.value }))
                    }
                    disabled={isSaving}
                  >
                    {Object.keys(PURPOSE_KEYS).map((key) => (
                      <option key={key} value={key}>
                        {purposeOptionLabel(key, txForm.date)}
                      </option>
                    ))}
                    <option value="other">Other</option>
                  </select>
                  {txForm.purposeKey === "other" && (
                    <input
                      className="rounded-md border border-slate-300 px-3 py-2 text-sm"
                      placeholder="Custom purpose"
                      value={txForm.customPurpose}
                      onChange={(e) =>
                        setTxForm((old) => ({ ...old, customPurpose: e.target.value }))
                      }
                      required
                      disabled={isSaving}
                    />
                  )}
                </label>
                <div className="flex items-end">
                  <button
                    type="submit"
                    disabled={!activeSchoolYear || !effectiveSelectedStudentId || isSaving || isLoading}
                    className="w-full rounded-md bg-slate-900 px-3 py-2 text-sm font-semibold text-white hover:bg-slate-700 disabled:opacity-50"
                  >
                    Post entry
                  </button>
                </div>
              </form>

              <p className="mt-2 text-xs text-slate-500">
                Will post as{" "}
                <span className="font-medium text-slate-700">
                  {txForm.type === "DEBIT" ? "Debit" : "Credit"}
                </span>
                {" · "}
                <span className="font-medium text-slate-700">{purposePreview}</span>
                {txForm.type === "DEBIT" ? " (+ balance)" : " (− balance)"}
              </p>

              <div className="mt-4 max-h-96 overflow-auto rounded-md border border-slate-200">
                <table className="min-w-full text-sm">
                  <thead className="sticky top-0 z-10 bg-slate-100">
                    <tr>
                      <th className="px-3 py-2 text-left">Date</th>
                      <th className="px-3 py-2 text-left">OR #</th>
                      <th className="px-3 py-2 text-left">Type</th>
                      <th className="px-3 py-2 text-left">Transaction</th>
                      <th className="px-3 py-2 text-right">Amount</th>
                      <th className="px-3 py-2 text-right">Balance</th>
                      <th className="px-3 py-2 text-right"> </th>
                    </tr>
                  </thead>
                  <tbody>
                    {selectedStudentLedger.length === 0 ? (
                      <tr>
                        <td colSpan={7} className="px-3 py-8 text-center text-slate-500">
                          No transactions yet
                        </td>
                      </tr>
                    ) : (
                      selectedStudentLedger.map((tx) => (
                        <tr key={tx.transactionId} className="border-t border-slate-200">
                          <td className="px-3 py-2 whitespace-nowrap">
                            {formatDisplayDate(tx.date)}
                          </td>
                          <td className="px-3 py-2">{tx.orNumber || "—"}</td>
                          <td className="px-3 py-2">{tx.type}</td>
                          <td className="px-3 py-2 whitespace-pre-line">{tx.purpose}</td>
                          <td className="px-3 py-2 text-right">{currency(tx.amount)}</td>
                          <td className="px-3 py-2 text-right font-semibold">
                            {currency(tx.runningBalance)}
                          </td>
                          <td className="px-3 py-2 text-right">
                            <button
                              type="button"
                              disabled={isSaving}
                              onClick={() => handleDeleteTransaction(tx)}
                              className="text-xs text-rose-600 hover:underline disabled:opacity-50"
                            >
                              Delete
                            </button>
                          </td>
                        </tr>
                      ))
                    )}
                  </tbody>
                </table>
              </div>
            </div>
          </section>
        )}

        {activeTab === "manage" && (
          <section className="mt-6 space-y-6">
            <div className="rounded-lg border border-slate-200 p-4">
              <div className="flex flex-wrap items-center justify-between gap-3">
                <h2 className="text-lg font-semibold">All Students</h2>
                <div className="flex flex-wrap gap-2">
                  <button
                    type="button"
                    onClick={() =>
                      setSelectedPrintStudentIds(sortedStudents.map((s) => s.studentId))
                    }
                    disabled={sortedStudents.length === 0}
                    className="rounded-md bg-slate-200 px-3 py-2 text-sm font-semibold text-slate-800 hover:bg-slate-300 disabled:opacity-50"
                  >
                    Check All
                  </button>
                  <button
                    type="button"
                    onClick={() => setSelectedPrintStudentIds([])}
                    disabled={sortedStudents.length === 0}
                    className="rounded-md bg-slate-200 px-3 py-2 text-sm font-semibold text-slate-800 hover:bg-slate-300 disabled:opacity-50"
                  >
                    Uncheck All
                  </button>
                  <button
                    type="button"
                    onClick={printSelectedStudents}
                    disabled={selectedPrintStudentIds.length === 0 || isLoading}
                    className="rounded-md bg-indigo-700 px-4 py-2 text-sm font-semibold text-white hover:bg-indigo-800 disabled:opacity-50"
                  >
                    Print Selected
                  </button>
                </div>
              </div>
              <p className="mt-1 text-sm text-slate-600">
                Checked students will each receive an SOA copy.
              </p>
              <div className="mt-3 max-h-[28rem] overflow-auto rounded-md border border-slate-200">
                <table className="min-w-full text-sm">
                  <thead className="sticky top-0 z-10 bg-slate-100 shadow-sm">
                    <tr>
                      <th className="px-3 py-2 text-left font-semibold">Print</th>
                      <th className="px-3 py-2 text-left font-semibold">Student name</th>
                      <th className="px-3 py-2 text-left font-semibold">Grade</th>
                      <th className="px-3 py-2 text-right font-semibold">Balance</th>
                    </tr>
                  </thead>
                  <tbody>
                    {sortedStudents.length === 0 ? (
                      <tr>
                        <td colSpan={4} className="px-3 py-6 text-center text-slate-500">
                          No students found
                        </td>
                      </tr>
                    ) : (
                      sortedStudents.map((student) => (
                        <tr key={student.studentId} className="border-t border-slate-200">
                          <td className="px-3 py-2">
                            <input
                              type="checkbox"
                              checked={selectedPrintStudentIds.includes(student.studentId)}
                              onChange={() => togglePrintStudent(student.studentId)}
                              className="h-4 w-4"
                            />
                          </td>
                          <td className="px-3 py-2 font-medium">{student.studentName}</td>
                          <td className="px-3 py-2 text-slate-600">
                            {getGradeName(student.gradeLevelId)}
                          </td>
                          <td className="px-3 py-2 text-right font-semibold text-rose-700">
                            {currency(studentBalanceMap[student.studentId])}
                          </td>
                        </tr>
                      ))
                    )}
                  </tbody>
                </table>
              </div>
            </div>

            <div className="rounded-lg border border-slate-200 p-4">
              <div className="flex flex-wrap items-center justify-between gap-3">
                <div>
                  <h2 className="text-lg font-semibold">Full Backup Export</h2>
                  <p className="mt-1 text-sm text-slate-600">
                    Exports students, transactions, balances, and the current fee matrix in one
                    Excel file.
                  </p>
                </div>
                <button
                  type="button"
                  onClick={() =>
                    downloadWorkbook({
                      students: sortedStudents,
                      ledgerTransactions: recomputeRunningBalances(ledgerTransactions),
                      gradeLevels,
                      balanceMap: studentBalanceMap,
                      getGradeName,
                      schoolYearLabel: activeSchoolYear?.label ?? "",
                    })
                  }
                  disabled={isLoading}
                  className="rounded-md bg-emerald-700 px-4 py-2 text-sm font-semibold text-white hover:bg-emerald-800 disabled:opacity-50"
                >
                  Export Full Backup
                </button>
              </div>
            </div>

            <div className="rounded-lg border border-slate-200 p-4">
              <h2 className="text-lg font-semibold">Print by Grade Level</h2>
              <div className="mt-3 max-h-64 overflow-auto rounded-md border border-slate-200">
                <table className="min-w-full text-sm">
                  <thead className="sticky top-0 bg-slate-100">
                    <tr>
                      <th className="px-3 py-2 text-left">Grade</th>
                      <th className="px-3 py-2 text-right">Students</th>
                      <th className="px-3 py-2 text-right">Total receivable</th>
                      <th className="px-3 py-2 text-right">Action</th>
                    </tr>
                  </thead>
                  <tbody>
                    {gradeLevels.map((grade) => {
                      const studentsInGrade = sortedStudents.filter(
                        (s) => s.gradeLevelId === grade.gradeLevelId,
                      );
                      const receivable = studentsInGrade.reduce(
                        (sum, s) => sum + (studentBalanceMap[s.studentId] || 0),
                        0,
                      );
                      return (
                        <tr key={grade.gradeLevelId} className="border-t border-slate-200">
                          <td className="px-3 py-2 font-medium">{grade.gradeLevelName}</td>
                          <td className="px-3 py-2 text-right">{studentsInGrade.length}</td>
                          <td className="px-3 py-2 text-right font-semibold text-rose-700">
                            {currency(receivable)}
                          </td>
                          <td className="px-3 py-2 text-right">
                            <button
                              type="button"
                              disabled={studentsInGrade.length === 0 || isLoading}
                              onClick={() =>
                                openPrintModal({ mode: "grade", gradeId: grade.gradeLevelId })
                              }
                              className="rounded bg-indigo-700 px-2 py-1 text-xs font-semibold text-white hover:bg-indigo-800 disabled:opacity-50"
                            >
                              Print SOA
                            </button>
                          </td>
                        </tr>
                      );
                    })}
                  </tbody>
                </table>
              </div>
            </div>

            <div className="rounded-lg border border-slate-200 p-4">
              <h2 className="text-lg font-semibold">Transactions by Date Range</h2>
              <div className="mt-3 flex flex-wrap items-center gap-2">
                <input
                  className="rounded-md border border-slate-300 px-3 py-2 text-sm"
                  type="date"
                  value={dateFilter.startDate}
                  onChange={(e) =>
                    setDateFilter((old) => ({ ...old, startDate: e.target.value }))
                  }
                />
                <input
                  className="rounded-md border border-slate-300 px-3 py-2 text-sm"
                  type="date"
                  value={dateFilter.endDate}
                  onChange={(e) =>
                    setDateFilter((old) => ({ ...old, endDate: e.target.value }))
                  }
                />
                <button
                  type="button"
                  onClick={() =>
                    openPrintModal({
                      mode: "daterange",
                      dateRange: {
                        startDate: dateFilter.startDate,
                        endDate: dateFilter.endDate,
                      },
                    })
                  }
                  className="rounded-md bg-indigo-700 px-4 py-2 text-sm font-semibold text-white hover:bg-indigo-800"
                >
                  Print SOA (date range)
                </button>
              </div>
              <p className="mt-1 text-xs text-slate-600">
                Prints two half-page SOAs per bond sheet. Overflow continues on the next half
                with Balance Forwarded; Amount Due only on each student&apos;s last half.
              </p>
              <div className="mt-3 max-h-80 overflow-auto rounded-md border border-slate-200">
                <table className="min-w-full text-sm">
                  <thead className="sticky top-0 z-10 bg-slate-100">
                    <tr>
                      <th className="px-3 py-2 text-left">Date</th>
                      <th className="px-3 py-2 text-left">Student</th>
                      <th className="px-3 py-2 text-left">OR #</th>
                      <th className="px-3 py-2 text-left">Transaction</th>
                      <th className="px-3 py-2 text-right">Amount</th>
                    </tr>
                  </thead>
                  <tbody>
                    {systemTransactionsByDate.length === 0 ? (
                      <tr>
                        <td colSpan={5} className="px-3 py-6 text-center text-slate-500">
                          No transactions in this range
                        </td>
                      </tr>
                    ) : (
                      systemTransactionsByDate.map((tx) => (
                        <tr key={tx.transactionId} className="border-t border-slate-200">
                          <td className="px-3 py-2 whitespace-nowrap">
                            {formatDisplayDate(tx.date)}
                          </td>
                          <td className="px-3 py-2">
                            {students.find((s) => s.studentId === tx.studentId)?.studentName}
                          </td>
                          <td className="px-3 py-2">{tx.orNumber || "—"}</td>
                          <td className="px-3 py-2 whitespace-pre-line">{tx.purpose}</td>
                          <td className="px-3 py-2 text-right">
                            {currency(Math.abs(tx.signedAmount))}
                          </td>
                        </tr>
                      ))
                    )}
                  </tbody>
                </table>
              </div>
            </div>
          </section>
        )}

        {activeTab === "admin" && (
          <section className="mt-6 space-y-6">
            <div className="rounded-lg border border-slate-200 p-4">
              <h2 className="text-lg font-semibold">Admin Access</h2>
              {!isAdminUnlocked ? (
                <form className="mt-4 flex max-w-lg flex-wrap gap-3" onSubmit={handleAdminUnlock}>
                  <input
                    type="password"
                    className="min-w-64 flex-1 rounded-md border border-slate-300 px-3 py-2 text-sm"
                    placeholder="Admin password"
                    value={adminPasswordInput}
                    onChange={(e) => setAdminPasswordInput(e.target.value)}
                  />
                  <button
                    type="submit"
                    className="rounded-md bg-slate-900 px-4 py-2 text-sm font-semibold text-white hover:bg-slate-700"
                  >
                    Unlock Admin
                  </button>
                </form>
              ) : (
                <div className="mt-3 flex flex-wrap items-center gap-2">
                  <span className="rounded-md bg-emerald-50 px-3 py-2 text-sm font-semibold text-emerald-800">
                    Admin unlocked
                  </span>
                  <button
                    type="button"
                    onClick={() => {
                      setIsAdminUnlocked(false);
                      setAdminMessage("Admin access locked.");
                    }}
                    className="rounded-md bg-slate-200 px-3 py-2 text-sm font-semibold text-slate-800 hover:bg-slate-300"
                  >
                    Lock
                  </button>
                </div>
              )}
              {adminMessage && <p className="mt-3 text-sm text-slate-700">{adminMessage}</p>}
            </div>

            {isAdminUnlocked && (
              <>
                <div className="rounded-lg border border-slate-200 p-4">
                  <div className="flex flex-wrap items-center justify-between gap-3">
                    <div>
                      <h2 className="text-lg font-semibold">Future Enrollment Prices</h2>
                      <p className="mt-1 text-sm text-slate-600">
                        Price changes apply only to students enrolled after saving. Existing ledger
                        values are unchanged.
                      </p>
                    </div>
                    <div className="flex flex-wrap gap-2">
                      <button
                        type="button"
                        onClick={resetFeeSettings}
                        className="rounded-md bg-slate-200 px-3 py-2 text-sm font-semibold text-slate-800 hover:bg-slate-300"
                      >
                        Reset Defaults
                      </button>
                      <button
                        type="button"
                        onClick={saveFeeSettings}
                        className="rounded-md bg-emerald-700 px-4 py-2 text-sm font-semibold text-white hover:bg-emerald-800"
                      >
                        Save Prices
                      </button>
                    </div>
                  </div>
                  <div className="mt-4 overflow-auto rounded-md border border-slate-200">
                    <table className="min-w-full text-sm">
                      <thead className="bg-slate-100">
                        <tr>
                          <th className="px-3 py-2 text-left">Grade</th>
                          <th className="px-3 py-2 text-right">General Fees</th>
                          <th className="px-3 py-2 text-right">Tuition</th>
                          <th className="px-3 py-2 text-right">Books</th>
                          <th className="px-3 py-2 text-right">Grand Total</th>
                        </tr>
                      </thead>
                      <tbody>
                        {feeSettings.map((grade) => (
                          <tr key={grade.id} className="border-t border-slate-200">
                            <td className="px-3 py-2 font-medium">{grade.name}</td>
                            {["generalFees", "tuition", "books"].map((field) => (
                              <td key={field} className="px-3 py-2 text-right">
                                <input
                                  type="number"
                                  min="0"
                                  step="0.01"
                                  value={grade[field]}
                                  onChange={(e) =>
                                    updateFeeSetting(grade.id, field, e.target.value)
                                  }
                                  className="w-32 rounded-md border border-slate-300 px-2 py-1 text-right"
                                />
                              </td>
                            ))}
                            <td className="px-3 py-2 text-right font-semibold">
                              {currency(grade.grandTotal)}
                            </td>
                          </tr>
                        ))}
                      </tbody>
                    </table>
                  </div>
                </div>

                <div className="rounded-lg border border-amber-200 bg-amber-50 p-4">
                  <h2 className="text-lg font-semibold text-amber-950">Archive All Students</h2>
                  <p className="mt-1 text-sm text-amber-900">
                    Saves the selected year as a past school year, then opens a fresh current year.
                    Past years stay selectable for viewing, editing, and SOA printing.
                  </p>
                  <div className="mt-4 grid gap-3 md:grid-cols-5">
                    <label className="flex flex-col gap-1">
                      <span className="text-xs font-medium text-amber-950">S.Y start</span>
                      <input
                        type="number"
                        min="1900"
                        max="3000"
                        className="rounded-md border border-amber-300 px-3 py-2 text-sm"
                        placeholder="2026"
                        value={archiveForm.startYear}
                        onChange={(e) =>
                          setArchiveForm((old) => ({ ...old, startYear: e.target.value }))
                        }
                      />
                    </label>
                    <label className="flex flex-col gap-1">
                      <span className="text-xs font-medium text-amber-950">S.Y end</span>
                      <input
                        type="number"
                        min="1900"
                        max="3000"
                        className="rounded-md border border-amber-300 px-3 py-2 text-sm"
                        placeholder="2027"
                        value={archiveForm.endYear}
                        onChange={(e) =>
                          setArchiveForm((old) => ({ ...old, endYear: e.target.value }))
                        }
                      />
                    </label>
                    <input
                      type="password"
                      className="rounded-md border border-amber-300 px-3 py-2 text-sm"
                      placeholder="Admin password"
                      value={archiveForm.firstPassword}
                      onChange={(e) =>
                        setArchiveForm((old) => ({ ...old, firstPassword: e.target.value }))
                      }
                    />
                    <input
                      type="password"
                      className="rounded-md border border-amber-300 px-3 py-2 text-sm"
                      placeholder="Re-enter admin password"
                      value={archiveForm.secondPassword}
                      onChange={(e) =>
                        setArchiveForm((old) => ({ ...old, secondPassword: e.target.value }))
                      }
                    />
                    <button
                      type="button"
                      onClick={handleArchiveAllStudents}
                      disabled={isSaving || isLoading || !activeSchoolYear || !activeSchoolYear.isCurrent}
                      className="rounded-md bg-amber-700 px-4 py-2 text-sm font-semibold text-white hover:bg-amber-800 disabled:opacity-50"
                    >
                      Archive All
                    </button>
                  </div>
                  <p className="mt-2 text-xs text-amber-900">
                    Archive label preview:{" "}
                    <span className="font-semibold">
                      S.Y {archiveForm.startYear || "____"} to {archiveForm.endYear || "____"}
                    </span>
                  </p>
                </div>
              </>
            )}
          </section>
        )}
      </div>

      {isPrintModalOpen && (
        <div className="no-print fixed inset-0 z-40 flex items-center justify-center bg-black/60 p-4">
          <div className="w-full max-w-md rounded-lg bg-white p-4">
            <h3 className="text-lg font-semibold">SOA Print Setup</h3>
            <div className="mt-2 space-y-2 text-xs">
              <div className="flex flex-wrap gap-2">
                {[
                  { id: "single", label: "Selected student" },
                  { id: "grade", label: "Whole grade" },
                  { id: "daterange", label: "Date range" },
                ].map((m) => (
                  <button
                    key={m.id}
                    type="button"
                    onClick={() => setPrintMode(m.id)}
                    className={`rounded border px-2 py-1 ${
                      printMode === m.id
                        ? "border-indigo-700 bg-indigo-700 text-white"
                        : "border-slate-300 bg-slate-50"
                    }`}
                  >
                    {m.label}
                  </button>
                ))}
              </div>
              {printMode === "grade" && (
                <select
                  className="w-full rounded-md border border-slate-300 px-3 py-2 text-sm"
                  value={printGradeId || selectedStudent?.gradeLevelId || ""}
                  onChange={(e) => setPrintGradeId(e.target.value)}
                >
                  <option value="">Select grade</option>
                  {gradeLevels.map((g) => (
                    <option key={g.gradeLevelId} value={g.gradeLevelId}>
                      {g.gradeLevelName}
                    </option>
                  ))}
                </select>
              )}
              {printMode === "daterange" && (
                <div className="flex gap-2">
                  <input
                    type="date"
                    className="flex-1 rounded-md border border-slate-300 px-2 py-2 text-sm"
                    value={printDateRange.startDate}
                    onChange={(e) =>
                      setPrintDateRange((o) => ({ ...o, startDate: e.target.value }))
                    }
                  />
                  <input
                    type="date"
                    className="flex-1 rounded-md border border-slate-300 px-2 py-2 text-sm"
                    value={printDateRange.endDate}
                    onChange={(e) =>
                      setPrintDateRange((o) => ({ ...o, endDate: e.target.value }))
                    }
                  />
                </div>
              )}
            </div>
            <label className="mt-3 flex flex-col gap-1">
              <span className="text-sm font-medium">Bill To</span>
              <input
                className="rounded-md border border-slate-300 px-3 py-2 text-sm"
                value={billToInput}
                onChange={(e) => setBillToInput(e.target.value)}
                placeholder={selectedStudent?.studentName || "Student / guardian name"}
              />
            </label>
            <div className="mt-4 flex justify-end gap-2">
              <button
                type="button"
                className="rounded-md bg-slate-200 px-3 py-2 text-sm"
                onClick={() => setIsPrintModalOpen(false)}
              >
                Cancel
              </button>
              <button
                type="button"
                className="rounded-md bg-indigo-700 px-3 py-2 text-sm font-semibold text-white"
                onClick={startPrint}
              >
                Print Now
              </button>
            </div>
          </div>
        </div>
      )}

      <div className="print-root hidden">
        <SoaPrintBundle
          tickets={tickets}
          billToPrinted={billToPrinted}
          documentDate={documentDate}
          schoolYearLabel={activeSchoolYear?.label ?? ""}
        />
      </div>
    </div>
  );
}

export default App;

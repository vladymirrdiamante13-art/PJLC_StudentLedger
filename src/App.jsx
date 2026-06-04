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
  insertStudent,
  insertSoaRow,
  insertSoaRowsBulk,
  deleteStudent,
  deleteSoaRow,
  buildDescription,
  baselinePayloads,
} from "./lib/supabaseApi";
import { SoaPrintBundle } from "./components/SoaPrintPages";

const GRADE_LEVELS = [
  { id: "k1", name: "Kinder 1", generalFees: 9650, tuition: 15000, books: 4200, grandTotal: 28850 },
  { id: "k2", name: "Kinder 2", generalFees: 9650, tuition: 15000, books: 4200, grandTotal: 28850 },
  { id: "g1", name: "Grade 1", generalFees: 11100, tuition: 15000, books: 5800, grandTotal: 31900 },
  { id: "g2", name: "Grade 2", generalFees: 11100, tuition: 15000, books: 5800, grandTotal: 31900 },
  { id: "g3", name: "Grade 3", generalFees: 14000, tuition: 15000, books: 5800, grandTotal: 34800 },
  { id: "g4", name: "Grade 4", generalFees: 14000, tuition: 15000, books: 6500, grandTotal: 35500 },
  { id: "g5", name: "Grade 5", generalFees: 14000, tuition: 15000, books: 6500, grandTotal: 35500 },
  { id: "g6", name: "Grade 6", generalFees: 14000, tuition: 15000, books: 6500, grandTotal: 35500 },
];

const gradeLevels = GRADE_LEVELS.map((g) => ({
  gradeLevelId: g.id,
  gradeLevelName: g.name,
  baseGeneralFees: g.generalFees,
  baseTuition: g.tuition,
  baseBooks: g.books,
  baseGrandTotal: g.grandTotal,
}));

const currency = (value) =>
  new Intl.NumberFormat("en-PH", { style: "currency", currency: "PHP" }).format(
    Number(value || 0),
  );

const today = () => new Date().toISOString().slice(0, 10);
const getGradeById = (id) => GRADE_LEVELS.find((g) => g.id === id);

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
  const [students, setStudents] = useState([]);
  const [ledgerTransactions, setLedgerTransactions] = useState([]);
  const [isLoading, setIsLoading] = useState(true);
  const [isSaving, setIsSaving] = useState(false);
  const [cloudMessage, setCloudMessage] = useState("");
  const [loadError, setLoadError] = useState("");

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
  });
  const [dateFilter, setDateFilter] = useState({ startDate: "", endDate: "" });
  const [isPrintModalOpen, setIsPrintModalOpen] = useState(false);
  const [billToInput, setBillToInput] = useState("");
  const [billToPrinted, setBillToPrinted] = useState("");
  const [printMode, setPrintMode] = useState("single");
  const [printGradeId, setPrintGradeId] = useState("");
  const [printDateRange, setPrintDateRange] = useState({ startDate: "", endDate: "" });
  const [tickets, setTickets] = useState([]);

  const loadData = useCallback(async () => {
    setIsLoading(true);
    setLoadError("");
    setCloudMessage("Loading from cloud…");
    try {
      const [studentRows, soaRows] = await Promise.all([fetchStudents(), fetchSoaRows()]);
      setStudents(studentRows);
      setLedgerTransactions(soaRows);
      setCloudMessage("Synced with Supabase.");
    } catch (err) {
      const msg = err.message ?? "Could not load data from Supabase.";
      const hint = msg.includes("soa_rows")
        ? " Run supabase/migrate-soa_rows.sql in the Supabase SQL Editor, then click Refresh."
        : "";
      setLoadError(msg + hint);
      setCloudMessage("");
    } finally {
      setIsLoading(false);
    }
  }, []);

  useEffect(() => {
    loadData();
  }, [loadData]);

  const sortedStudents = useMemo(
    () => [...students].sort((a, b) => a.studentName.localeCompare(b.studentName)),
    [students],
  );

  useEffect(() => {
    if (!selectedStudentId && sortedStudents.length > 0) {
      setSelectedStudentId(sortedStudents[0].studentId);
    }
  }, [selectedStudentId, sortedStudents]);

  const selectedStudent = sortedStudents.find((s) => s.studentId === selectedStudentId);

  const selectedStudentLedger = useMemo(() => {
    return recomputeRunningBalances(
      ledgerTransactions.filter((tx) => tx.studentId === selectedStudentId),
    );
  }, [ledgerTransactions, selectedStudentId]);

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

  const purposePreview = resolvePurposeLabel(txForm.purposeKey, txForm.date);

  const runSave = async (label, fn) => {
    setIsSaving(true);
    setCloudMessage(label);
    try {
      await fn();
      await loadData();
    } catch (err) {
      setCloudMessage(`Error: ${err.message ?? "Save failed"}`);
    } finally {
      setIsSaving(false);
    }
  };

  const handleEnrollStudent = (event) => {
    event.preventDefault();
    if (!studentForm.studentName.trim()) return;

    runSave("Saving student to cloud…", async () => {
      const grade = getGradeById(studentForm.gradeLevelId);
      const created = await insertStudent({
        studentName: studentForm.studentName.trim(),
        gradeLevelId: studentForm.gradeLevelId,
      });
      const dateIso = today();
      const baseline = baselinePayloads(created.studentId, grade, dateIso);
      await insertSoaRowsBulk(baseline);
      setStudentForm({ studentName: "", gradeLevelId: studentForm.gradeLevelId });
      setSelectedStudentId(created.studentId);
      setActiveTab("ledger");
    });
  };

  const handleAddTransaction = (event) => {
    event.preventDefault();
    if (!selectedStudentId || !txForm.amount || Number(txForm.amount) <= 0) return;
    if (!PURPOSE_KEYS[txForm.purposeKey]) return;

    const purpose = resolvePurposeLabel(txForm.purposeKey, txForm.date);

    runSave("Posting entry to cloud…", async () => {
      await insertSoaRow({
        studentId: selectedStudentId,
        date: txForm.date,
        description: buildDescription(purpose, txForm.orNumber),
        amount: Number(txForm.amount),
        entryType: txForm.type,
        orNumber: txForm.orNumber.trim(),
        purposeKey: txForm.purposeKey,
      });
      setTxForm((old) => ({ ...old, amount: "", orNumber: "", date: today() }));
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

  return (
    <div className="mx-auto max-w-7xl p-6">
      <div className="no-print rounded-xl border border-slate-300 bg-white p-5 shadow-sm">
        <h1 className="text-2xl font-bold text-slate-900">
          Palawan Jewels Learning Center — Student Ledger & SOA
        </h1>
        <p className="mt-1 text-sm text-slate-600">
          School Year 2026-2027 · Cloud (Supabase) · Two half-page SOAs per bond sheet
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

        <div className="mt-5 flex flex-wrap gap-2">
          {[
            { id: "enroll", label: "Enroll Student" },
            { id: "ledger", label: "Ledger & Transactions" },
            { id: "manage", label: "Reports & Print" },
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
                  disabled={isSaving || isLoading}
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
                  selectedStudentId={selectedStudentId}
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
                  </select>
                </label>
                <div className="flex items-end">
                  <button
                    type="submit"
                    disabled={!selectedStudentId || isSaving || isLoading}
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
              <h2 className="text-lg font-semibold">All Students</h2>
              <StudentsTable
                students={sortedStudents}
                selectedStudentId={selectedStudentId}
                onSelect={(id) => {
                  setSelectedStudentId(id);
                  setActiveTab("ledger");
                }}
                balanceMap={studentBalanceMap}
                getGradeName={getGradeName}
                maxHeight="max-h-[28rem]"
              />
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
        />
      </div>
    </div>
  );
}

export default App;

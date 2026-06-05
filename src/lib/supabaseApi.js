import { supabase } from "../supabaseClient";
import { resolvePurposeLabel } from "./purposes";

export const DEFAULT_CURRENT_SCHOOL_YEAR_LABEL = "Current School Year";

export function formatSchoolYearLabel(startYear, endYear) {
  const start = String(startYear ?? "").trim();
  const end = String(endYear ?? "").trim();
  return `S.Y ${start} to ${end}`;
}

export function normalizeSchoolYearLabel(label) {
  return String(label ?? "").trim().replace(/\s+/g, " ").toLowerCase();
}

function formatBaselineAmount(value) {
  return Number(value || 0).toLocaleString("en-PH", {
    minimumFractionDigits: 2,
    maximumFractionDigits: 2,
  });
}

export function mapSchoolYear(row) {
  return {
    schoolYearId: row.id,
    label: row.label,
    isCurrent: Boolean(row.is_current),
    isArchived: Boolean(row.is_archived),
    archivedAt: row.archived_at,
    createdAt: row.created_at,
  };
}

/** Map DB row → app ledger transaction (running balance computed in UI). */
export function mapSoaRow(row) {
  const meta = row.meta && typeof row.meta === "object" ? row.meta : {};
  return {
    transactionId: row.id,
    studentId: row.student_id,
    schoolYearId: row.school_year_id,
    date: row.date ?? row.transaction_date,
    orNumber: row.or_number ?? meta.or_number ?? "",
    type: row.entry_type ?? meta.entry_type ?? "DEBIT",
    purposeKey: row.purpose_key ?? meta.purpose_key ?? "general_fees",
    purpose: row.description ?? row.purpose ?? "",
    amount: Number(row.amount),
    hideAmount: Boolean(meta.hide_amount ?? meta.hideAmount),
    isInitialBaseline: Boolean(meta.is_initial_baseline ?? meta.isInitialBaseline),
    createdAt: row.created_at ?? new Date().toISOString(),
  };
}

export function mapStudent(row) {
  return {
    studentId: row.student_id ?? row.id,
    schoolYearId: row.school_year_id,
    studentName: row.student_name,
    gradeLevelId: row.grade_level_id,
  };
}

/** Multi-line description for cloud storage (purpose + OR on second line). */
export function buildDescription(purpose, orNumber) {
  const lines = [purpose];
  if (orNumber?.trim()) lines.push(`OR # : ${orNumber.trim()}`);
  return lines.join("\n");
}

export async function fetchSchoolYears() {
  const { data, error } = await supabase
    .from("school_years")
    .select("*")
    .order("is_current", { ascending: false })
    .order("created_at", { ascending: false });
  if (error) throw error;
  return (data ?? []).map(mapSchoolYear);
}

export async function ensureCurrentSchoolYear() {
  const years = await fetchSchoolYears();
  const current = years.find((year) => year.isCurrent) ?? years[0];
  if (current) return { current, years };

  const { data, error } = await supabase
    .from("school_years")
    .insert({
      label: DEFAULT_CURRENT_SCHOOL_YEAR_LABEL,
      is_current: true,
      is_archived: false,
    })
    .select("*")
    .single();
  if (error) throw error;

  const created = mapSchoolYear(data);
  return { current: created, years: [created] };
}

export async function fetchStudents(schoolYearId) {
  const { data, error } = await supabase
    .from("students")
    .select("student_id, school_year_id, student_name, grade_level_id")
    .eq("school_year_id", schoolYearId)
    .order("student_name");
  if (error) throw error;
  return (data ?? []).map(mapStudent);
}

export async function fetchSoaRows(schoolYearId) {
  const { data, error } = await supabase
    .from("soa_rows")
    .select("*")
    .eq("school_year_id", schoolYearId)
    .order("date", { ascending: true })
    .order("created_at", { ascending: true });
  if (error) throw error;
  return (data ?? []).map(mapSoaRow);
}

export async function insertStudent({ schoolYearId, studentName, gradeLevelId }) {
  const { data, error } = await supabase
    .from("students")
    .insert({
      school_year_id: schoolYearId,
      student_name: studentName,
      grade_level_id: gradeLevelId,
    })
    .select("student_id, school_year_id, student_name, grade_level_id")
    .single();
  if (error) throw error;
  return mapStudent(data);
}

export async function insertSoaRow({
  schoolYearId,
  studentId,
  date,
  description,
  amount,
  entryType,
  orNumber = "",
  purposeKey = "",
  meta = {},
}) {
  const payload = {
    school_year_id: schoolYearId,
    student_id: studentId,
    date,
    description,
    amount,
    entry_type: entryType,
    or_number: orNumber,
    purpose_key: purposeKey,
    meta: { ...meta, entry_type: entryType, or_number: orNumber, purpose_key: purposeKey },
  };

  const { data, error } = await supabase.from("soa_rows").insert(payload).select("*").single();
  if (error) throw error;
  return mapSoaRow(data);
}

export async function insertSoaRowsBulk(rows) {
  const payloads = rows.map((r) => ({
    school_year_id: r.schoolYearId,
    student_id: r.studentId,
    date: r.date,
    description: r.description ?? buildDescription(r.purpose, r.orNumber),
    amount: r.amount,
    entry_type: r.type,
    or_number: r.orNumber ?? "",
    purpose_key: r.purposeKey ?? "",
    meta: {
      ...(r.meta ?? {}),
      entry_type: r.type,
      or_number: r.orNumber ?? "",
      purpose_key: r.purposeKey ?? "",
    },
  }));

  const { data, error } = await supabase.from("soa_rows").insert(payloads).select("*");
  if (error) throw error;
  return (data ?? []).map(mapSoaRow);
}

export async function deleteStudent(studentId) {
  const { error: rowErr } = await supabase.from("soa_rows").delete().eq("student_id", studentId);
  if (rowErr) throw rowErr;
  const { error } = await supabase.from("students").delete().eq("student_id", studentId);
  if (error) throw error;
}

export async function archiveSchoolYearAndStartNew({ schoolYearId, archiveLabel }) {
  const normalizedArchiveLabel = normalizeSchoolYearLabel(archiveLabel);
  const years = await fetchSchoolYears();
  const duplicateYear = years.find(
    (year) =>
      year.schoolYearId !== schoolYearId &&
      normalizeSchoolYearLabel(year.label) === normalizedArchiveLabel,
  );
  if (duplicateYear) {
    throw new Error(`${archiveLabel} already exists. Choose a different school year.`);
  }

  const archivedAt = new Date().toISOString();
  const { error: archiveError } = await supabase
    .from("school_years")
    .update({
      label: archiveLabel,
      is_archived: true,
      is_current: false,
      archived_at: archivedAt,
    })
    .eq("id", schoolYearId);
  if (archiveError) throw archiveError;

  const { data, error } = await supabase
    .from("school_years")
    .insert({
      label: DEFAULT_CURRENT_SCHOOL_YEAR_LABEL,
      is_current: true,
      is_archived: false,
    })
    .select("*")
    .single();
  if (error) throw error;
  return mapSchoolYear(data);
}

export async function deleteArchivedSchoolYear(schoolYearId) {
  const years = await fetchSchoolYears();
  const targetYear = years.find((year) => year.schoolYearId === schoolYearId);
  if (!targetYear) {
    throw new Error("School year archive not found.");
  }
  if (targetYear.isCurrent) {
    throw new Error("The current school year cannot be deleted.");
  }

  const { error: rowError } = await supabase
    .from("soa_rows")
    .delete()
    .eq("school_year_id", schoolYearId);
  if (rowError) throw rowError;

  const { error: studentError } = await supabase
    .from("students")
    .delete()
    .eq("school_year_id", schoolYearId);
  if (studentError) throw studentError;

  const { error } = await supabase.from("school_years").delete().eq("id", schoolYearId);
  if (error) throw error;
}

export async function deleteSoaRow(transactionId) {
  const { error } = await supabase.from("soa_rows").delete().eq("id", transactionId);
  if (error) throw error;
}

/** Baseline fee rows for new enrollment. */
export function baselinePayloads(studentId, grade, dateIso) {
  const items = [
    {
      label: resolvePurposeLabel("general_fees", dateIso),
      purposeKey: "general_fees",
      amount: Number(grade.generalFees) || 0,
    },
    {
      label: resolvePurposeLabel("tuition", dateIso),
      purposeKey: "tuition",
      amount: Number(grade.tuition) || 0,
    },
    {
      label: resolvePurposeLabel("books", dateIso),
      purposeKey: "books",
      amount: Number(grade.books) || 0,
    },
  ];
  const total = items.reduce((sum, item) => sum + item.amount, 0);
  const purpose = [
    ...items.map((item) => `${item.label} - ${formatBaselineAmount(item.amount)}`),
    "",
    `Total - ${formatBaselineAmount(total)}`,
  ].join("\n");

  return [
    {
      studentId,
      date: dateIso,
      type: "DEBIT",
      purposeKey: "initial_fees",
      purpose,
      orNumber: "",
      amount: total,
      description: purpose,
      meta: {
        hide_amount: true,
        is_initial_baseline: true,
        fee_breakdown: items.map((item) => ({
          purpose_key: item.purposeKey,
          label: item.label,
          amount: item.amount,
        })),
        total,
      },
    },
  ];
}

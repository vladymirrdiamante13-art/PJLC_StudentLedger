import { recomputeRunningBalances } from "./ledger";
import { resolvePurposeLabel } from "./purposes";

const today = () => new Date().toISOString().slice(0, 10);
const uid = (prefix) => `${prefix}-${crypto.randomUUID().slice(0, 8)}`;

export function createBaselineRows(studentId, grade) {
  if (!grade) return [];
  const createdAt = new Date().toISOString();
  const date = today();

  return [
    {
      transactionId: uid("txn"),
      studentId,
      date,
      orNumber: "",
      type: "DEBIT",
      purposeKey: "general_fees",
      purpose: resolvePurposeLabel("general_fees", date),
      amount: grade.generalFees,
      createdAt: `${createdAt}-gf`,
    },
    {
      transactionId: uid("txn"),
      studentId,
      date,
      orNumber: "",
      type: "DEBIT",
      purposeKey: "tuition",
      purpose: resolvePurposeLabel("tuition", date),
      amount: grade.tuition,
      createdAt: `${createdAt}-tu`,
    },
    {
      transactionId: uid("txn"),
      studentId,
      date,
      orNumber: "",
      type: "DEBIT",
      purposeKey: "books",
      purpose: resolvePurposeLabel("books", date),
      amount: grade.books,
      createdAt: `${createdAt}-bk`,
    },
  ];
}

export function createBaselineTransactions(studentId, grade) {
  return recomputeRunningBalances(createBaselineRows(studentId, grade));
}

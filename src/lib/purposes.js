const MONTHS = [
  "January",
  "February",
  "March",
  "April",
  "May",
  "June",
  "July",
  "August",
  "September",
  "October",
  "November",
  "December",
];

/** Label-only purposes — Debit/Credit is chosen separately on each entry. */
export const PURPOSE_KEYS = {
  general_fees: { label: "General Fees" },
  tuition: { label: "Tuition" },
  books: { label: "Books" },
  downpayment: { label: "Downpayment" },
  monthly_payment: { label: "Payment for the month of", usesMonth: true },
  discount: { label: "Discount" },
};

export function monthNameFromDate(isoDate) {
  if (!isoDate) return MONTHS[new Date().getMonth()];
  const [y, m] = isoDate.split("-").map(Number);
  const d = new Date(y, (m || 1) - 1, 1);
  return MONTHS[d.getMonth()];
}

export function resolvePurposeLabel(purposeKey, isoDate) {
  const def = PURPOSE_KEYS[purposeKey];
  if (!def) return purposeKey;
  if (def.usesMonth) {
    return `Payment for the month of ${monthNameFromDate(isoDate)}`;
  }
  return def.label;
}

export function purposeOptionLabel(key, isoDate) {
  return resolvePurposeLabel(key, isoDate);
}

export function purposeKeyFromStoredPurpose(purposeText) {
  if (!purposeText) return "general_fees";
  const lower = purposeText.toLowerCase();
  if (lower.startsWith("payment for the month of")) return "monthly_payment";
  if (lower.includes("tuition")) return "tuition";
  if (lower.includes("book")) return "books";
  if (lower.includes("downpayment")) return "downpayment";
  if (lower.includes("discount")) return "discount";
  if (lower.includes("general")) return "general_fees";
  return "general_fees";
}

export function signedAmount(type, amount) {
  const n = Number(amount) || 0;
  return type === "DEBIT" ? n : -n;
}

export function formatDisplayDate(isoDate) {
  if (!isoDate) return "";
  const [y, m, d] = isoDate.split("-");
  return `${m}/${d}/${y}`;
}

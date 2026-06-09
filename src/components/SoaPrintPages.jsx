import { paginateSoaLedger } from "../lib/paginateSoaLedger";
import { formatDisplayDate } from "../lib/purposes";

function fmt(n) {
  return Number(n).toLocaleString("en-PH", {
    minimumFractionDigits: 2,
    maximumFractionDigits: 2,
  });
}

function transactionLabel(tx) {
  if (tx.isBalanceForward) return "Balance Forwarded";
  const orPart = tx.orNumber?.trim() ? ` · OR # : ${tx.orNumber.trim()}` : "";
  return `${tx.purpose}${orPart}`;
}

function transactionAmount(tx) {
  if (tx.isBalanceForward || tx.hideAmount) return "";
  return fmt(Math.abs(tx.signedAmount ?? tx.amount));
}

function HalfPage({
  billTo,
  statementDate,
  schoolYearLabel,
  rows,
  isLast,
  amountDue,
  isTopOnSheet,
}) {
  return (
    <section
      className={`soa-half-page box-border flex h-[5.5in] w-[8.5in] flex-col overflow-hidden p-3 text-black ${
        isTopOnSheet ? "soa-half-page--top" : ""
      }`}
    >
      <div className="flex shrink-0 justify-between gap-3">
        <div className="flex min-w-0 items-start gap-2">
          <img
            src="/logo.jpg"
            alt="Palawan Jewels Learning Center"
            className="soa-logo h-11 w-11 shrink-0 object-contain"
          />
          <div className="text-[9px] leading-tight">
            <p className="font-semibold">Palawan Jewels Learning Center</p>
            <p>2/flr. Eastville City Walk Building,</p>
            <p>National Highway, Bgy. San Pedro,</p>
            <p>Puerto Princesa City, 5300</p>
            <p>Tel. No, 434-9878 Smart 0908-542-1102</p>
          </div>
        </div>
        <div className="shrink-0 text-right">
          <p className="text-sm font-bold">STATEMENT OF ACCOUNT</p>
          {schoolYearLabel && <p className="text-[9px]">{schoolYearLabel}</p>}
          <p className="text-[9px]">{statementDate}</p>
        </div>
      </div>

      <table className="mt-2 w-full shrink-0 border-collapse text-[11px]">
        <tbody>
          <tr>
            <td className="border border-black px-1.5 py-0.5 font-bold">BILL TO</td>
          </tr>
          <tr>
            <td className="border border-black border-t-0 px-1.5 py-1 text-[10px] uppercase">
              {billTo || "\u00a0"}
            </td>
          </tr>
        </tbody>
      </table>

      <div className="soa-ledger-body mt-2 min-h-0 flex-1 overflow-hidden">
        <table className="w-full border-collapse text-[11px]">
          <thead>
            <tr>
              <th className="w-[17%] border border-black px-1 py-0.5 text-left font-bold">Date</th>
              <th className="border border-black px-1 py-0.5 text-left font-bold">Transaction</th>
              <th className="w-[17%] border border-black px-1 py-0.5 text-right font-bold">Amount</th>
              <th className="w-[17%] border border-black px-1 py-0.5 text-right font-bold">Balance</th>
            </tr>
          </thead>
          <tbody>
            {rows.map((tx) => (
              <tr key={tx.transactionId}>
                <td className="border border-black px-1 py-0.5 align-top">
                  {tx.isBalanceForward ? "" : formatDisplayDate(tx.date)}
                </td>
                <td className="whitespace-pre-line border border-black px-1 py-0.5 align-top">
                  {transactionLabel(tx)}
                </td>
                <td className="border border-black px-1 py-0.5 text-right align-top">
                  {transactionAmount(tx)}
                </td>
                <td className="border border-black px-1 py-0.5 text-right align-top">
                  {fmt(tx.runningBalance)}
                </td>
              </tr>
            ))}
          </tbody>
        </table>
      </div>

      <div className="soa-half-footer mt-auto shrink-0 pt-1">
        {isLast ? (
          <div className="flex justify-end">
            <div className="text-[11px]">
              <div className="flex min-w-[200px] justify-between gap-4 border border-black px-2 py-0.5">
                <span className="font-semibold">Amount Due: Php</span>
                <span className="font-semibold">{fmt(amountDue)}</span>
              </div>
              <div className="flex min-w-[200px] justify-between gap-4 border border-black border-t-0 px-2 py-0.5">
                <span>Amount Enc. :</span>
                <span className="min-w-[60px] border-b border-black" />
              </div>
            </div>
          </div>
        ) : (
          <p className="text-center text-[10px] text-stone-600">Continued on next page...</p>
        )}
      </div>
    </section>
  );
}

/** Build flat half-page queue, then pair 2 per letter sheet (top + bottom). */
function buildPrintSheets(tickets) {
  const halves = [];

  tickets.forEach((ticket) => {
    const pages = paginateSoaLedger(ticket.ledger);
    pages.forEach((page, pageIndex) => {
      halves.push({
        key: `${ticket.studentId || ticket.studentName}-${pageIndex}`,
        billTo: ticket.billTo,
        rows: page.rows,
        isLast: page.isLast,
        amountDue: ticket.amountDue,
      });
    });
  });

  const sheets = [];
  for (let i = 0; i < halves.length; i += 2) {
    sheets.push(halves.slice(i, i + 2));
  }
  return sheets;
}

export function SoaPrintBundle({ tickets, billToPrinted, documentDate, schoolYearLabel }) {
  const sheets = buildPrintSheets(tickets);

  return (
    <>
      {sheets.map((pair, sheetIndex) => (
        <div key={`sheet-${sheetIndex}`} className="print-sheet">
          {pair.map((half, indexInSheet) => (
            <HalfPage
              key={half.key}
              billTo={billToPrinted || half.billTo}
              statementDate={documentDate}
              schoolYearLabel={schoolYearLabel}
              rows={half.rows}
              isLast={half.isLast}
              amountDue={half.amountDue}
              isTopOnSheet={indexInSheet === 0}
            />
          ))}
        </div>
      ))}
    </>
  );
}

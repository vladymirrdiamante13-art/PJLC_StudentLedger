/**
 * Max ledger data rows per 5.5in half-page (after header, bill-to, table head, footer).
 * If rows exceed this, the next half-page opens with "Balance Forwarded".
 */
export const ROWS_PER_HALF_PAGE = 8;

/**
 * Split ledger into half-pages for print.
 * Continuation pages lead with "Balance Forwarded" and running balance.
 */
export function paginateSoaLedger(ledger, rowsPerPage = ROWS_PER_HALF_PAGE) {
  if (!ledger.length) {
    return [
      {
        rows: [],
        isFirst: true,
        isLast: true,
      },
    ];
  }

  const pages = [];
  let index = 0;

  while (index < ledger.length) {
    const isFirst = pages.length === 0;
    const slotCount = isFirst ? rowsPerPage : rowsPerPage - 1;
    const chunk = ledger.slice(index, index + slotCount);
    const rows = [];

    if (!isFirst) {
      const prevBalance = pages[pages.length - 1].closingBalance;
      rows.push({
        transactionId: `balance-forward-${pages.length}`,
        isBalanceForward: true,
        date: "",
        description: "Balance Forwarded",
        signedAmount: 0,
        amount: 0,
        runningBalance: prevBalance,
        orNumber: "",
        purpose: "",
      });
    }

    rows.push(...chunk);
    const closingBalance = rows[rows.length - 1].runningBalance;
    index += slotCount;

    pages.push({
      rows,
      closingBalance,
      isFirst,
      isLast: index >= ledger.length,
    });
  }

  if (pages.length) {
    pages[pages.length - 1].isLast = true;
  }

  return pages;
}

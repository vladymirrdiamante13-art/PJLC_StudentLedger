import { signedAmount } from "./purposes";

export function recomputeRunningBalances(transactions) {
  const ordered = [...transactions].sort((a, b) => {
    const byDate = a.date.localeCompare(b.date);
    return byDate !== 0 ? byDate : a.createdAt.localeCompare(b.createdAt);
  });

  let rolling = 0;
  return ordered.map((tx) => {
    const signed = signedAmount(tx.type, tx.amount);
    rolling = Math.round((rolling + signed) * 100) / 100;
    return {
      ...tx,
      signedAmount: signed,
      runningBalance: rolling,
    };
  });
}

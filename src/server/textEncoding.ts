import path from "node:path";

const MOJIBAKE_PATTERN = /[\u0080-\u009f]|[ÃÂ]|[èéåçäæ]/u;

function mojibakeScore(value: string) {
  const matches = value.match(/[\u0080-\u009f�ÃÂ]|[èéåçäæ][\u0080-\u00ff]/gu);
  return matches?.length ?? 0;
}

export function repairMojibake(value = "") {
  let current = value;
  for (let index = 0; index < 2; index += 1) {
    if (!MOJIBAKE_PATTERN.test(current)) break;
    const repaired = Buffer.from(current, "latin1").toString("utf8");
    if (mojibakeScore(repaired) > mojibakeScore(current)) break;
    current = repaired;
  }
  return current;
}

export function inferStoreNameFromFileName(fileName = "") {
  const repairedName = repairMojibake(fileName);
  const baseName = path.basename(repairedName, path.extname(repairedName)).trim();
  return normalizeStoreName(baseName);
}

export function normalizeStoreName(storeName = "") {
  return repairMojibake(storeName)
    .replace(/[-_ ]?订单.*$/u, "")
    .replace(/[-_ ]?资金流水.*$/u, "")
    .replace(/[-_ ]?运费险.*$/u, "")
    .trim() || "临时店铺";
}

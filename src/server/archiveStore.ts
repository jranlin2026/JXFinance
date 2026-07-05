import fs from "node:fs/promises";
import path from "node:path";
import * as XLSX from "xlsx";
import { SettlementRecord, StoreConfig } from "../shared/types";
import { SettlementBuildResult, SettlementInput, snapshotConfig } from "./settlementEngine";
import { ensureDir, safePathSegment, timestampForPath, writeJson } from "./utils";

const settlementsRoot = path.resolve("data/settlements");
const XLSXLib: typeof XLSX = (XLSX as unknown as { default?: typeof XLSX }).default ?? XLSX;

export async function nextSettlementArchive(store: StoreConfig, month: string) {
  const storeSegment = safePathSegment(store.name);
  const monthDir = path.join(settlementsRoot, month, storeSegment);
  await ensureDir(monthDir);
  const existing = await fs.readdir(monthDir, { withFileTypes: true }).catch(() => []);
  const maxVersion = existing.reduce((max, item) => {
    if (!item.isDirectory()) return max;
    const match = item.name.match(/^v(\d+)_/);
    return match ? Math.max(max, Number(match[1])) : max;
  }, 0);
  const version = maxVersion + 1;
  const archivePath = path.join(monthDir, `v${version}_${timestampForPath()}`);
  await ensureDir(path.join(archivePath, "result"));
  return { version, archivePath };
}

export async function saveSettlementArchive(params: {
  input: SettlementInput;
  result: SettlementBuildResult;
  archivePath: string;
  version: number;
  uploadFiles: { path: string; originalname: string }[];
}) {
  const { input, result, archivePath, version, uploadFiles } = params;

  const period = result.coveredMonths.orders.length === 1 ? result.coveredMonths.orders[0] : "多月份";
  const fileName = `${safePathSegment(input.store.name)}_${period}_v${version}_结算输出表.xlsx`;
  const resultPath = path.join(archivePath, "result", fileName);
  XLSXLib.writeFile(result.workbook, resultPath);

  const generatedAt = input.generatedAt ?? new Date();
  const id = `${period}_${safePathSegment(input.store.name)}_v${version}_${timestampForPath(generatedAt)}`;
  const record: SettlementRecord = {
    id,
    storeId: input.store.id,
    storeName: input.store.name,
    month: period,
    version,
    generatedAt: generatedAt.toISOString(),
    archivePath,
    downloadUrl: `/api/settlements/${encodeURIComponent(id)}/download`,
    uploadFileCount: uploadFiles.length,
    stats: result.stats,
    summary: result.summaryRows as SettlementRecord["summary"],
    exceptionSummary: result.exceptionSummary,
    coveredMonths: result.coveredMonths
  };

  await writeJson(path.join(archivePath, "settlement.json"), record);
  await writeJson(
    path.join(archivePath, "config-snapshot.json"),
    snapshotConfig(input.store, input.mappings, input.shippingFee)
  );
  await writeJson(path.join(archivePath, "params.json"), {
    month: period,
    shippingFee: input.shippingFee,
    uploadedFileNames: input.uploadedFileNames,
    generatedAt: generatedAt.toISOString()
  });

  return { record, resultPath };
}

async function walkSettlementRecords(dir: string): Promise<SettlementRecord[]> {
  const entries = await fs.readdir(dir, { withFileTypes: true }).catch(() => []);
  const records: SettlementRecord[] = [];
  for (const entry of entries) {
    const fullPath = path.join(dir, entry.name);
    if (entry.isDirectory()) {
      records.push(...(await walkSettlementRecords(fullPath)));
    } else if (entry.name === "settlement.json") {
      const content = await fs.readFile(fullPath, "utf8");
      records.push(JSON.parse(content) as SettlementRecord);
    }
  }
  return records;
}

export async function listSettlementRecords() {
  const records = await walkSettlementRecords(settlementsRoot);
  return records.sort((a, b) => b.generatedAt.localeCompare(a.generatedAt));
}

export async function findSettlementRecord(id: string) {
  const records = await listSettlementRecords();
  return records.find((record) => record.id === id) ?? null;
}

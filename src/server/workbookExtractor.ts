import fs from "node:fs/promises";
import path from "node:path";
import { execFile } from "node:child_process";
import { promisify } from "node:util";

const execFileAsync = promisify(execFile);

type ExtractedWorkbookRows = {
  rows: Record<string, unknown>[];
  missing: string[];
};

const bundledPython = path.join(
  process.env.USERPROFILE ?? "",
  ".cache",
  "codex-runtimes",
  "codex-primary-runtime",
  "dependencies",
  "python",
  "python.exe"
);

function pythonCandidates() {
  return [process.env.PYTHON_EXE, process.env.PYTHON, bundledPython, "python"].filter(Boolean) as string[];
}

async function runPythonExtractor(filePath: string, selectedFields?: string[]) {
  const outputPath = path.join(path.dirname(filePath), `${path.basename(filePath)}.rows.json`);
  const scriptPath = path.resolve("src/server/extractWorkbookRows.py");
  const fieldsArg = selectedFields?.length ? JSON.stringify(selectedFields) : "null";
  let lastError: unknown;

  for (const python of pythonCandidates()) {
    try {
      await execFileAsync(python, [scriptPath, filePath, outputPath, fieldsArg], {
        maxBuffer: 1024 * 1024 * 16,
        windowsHide: true,
        timeout: 1000 * 60 * 20
      });
      const payload = JSON.parse(await fs.readFile(outputPath, "utf8")) as ExtractedWorkbookRows;
      await fs.rm(outputPath, { force: true });
      return payload;
    } catch (error) {
      lastError = error;
    }
  }

  throw lastError instanceof Error ? lastError : new Error("无法调用 Python 解析 Excel。");
}

export async function extractWorkbookRows(filePath: string, selectedFields?: string[]) {
  const startedAt = Date.now();
  const result = await runPythonExtractor(filePath, selectedFields);
  console.log(`Excel extracted: ${path.basename(filePath)} rows=${result.rows.length} seconds=${((Date.now() - startedAt) / 1000).toFixed(1)}`);
  return result;
}

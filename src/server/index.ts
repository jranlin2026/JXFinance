import cors from "cors";
import express from "express";
import fs from "node:fs/promises";
import path from "node:path";
import { execFile } from "node:child_process";
import multer from "multer";
import { createServer as createViteServer } from "vite";
import { getConfig, saveMappings, saveStores } from "./configStore";
import { buildSettlement } from "./settlementEngine";
import { findSettlementRecord, listSettlementRecords, nextSettlementArchive, saveSettlementArchive } from "./archiveStore";
import { ensureDir } from "./utils";
import { extractWorkbookRows } from "./workbookExtractor";
import { StoreConfig } from "../shared/types";

const app = express();
const port = Number(process.env.PORT ?? 3000);
const tmpDir = path.resolve("data/tmp/uploads");
await ensureDir(tmpDir);

const upload = multer({ dest: tmpDir });

const ORDER_FIELDS = [
  "主订单编号",
  "子订单编号",
  "商品数量",
  "商家编码",
  "商品单价",
  "订单应付金额",
  "订单提交时间",
  "订单完成时间",
  "支付完成时间",
  "达人ID",
  "达人昵称",
  "发货时间",
  "平台实际承担优惠金额",
  "商家实际承担优惠金额",
  "达人实际承担优惠金额"
];

const FLOW_REQUIRED_FIELDS = ["动账时间", "动账方向", "动账金额", "动账场景", "子订单号", "订单号", "备注"];
const PRODUCT_COST_FIELDS = ["商家编码", "产品成本"];
const FREIGHT_FIELDS = ["订单编号", "支付保费", "保费状态", "动账时间", "下单时间", "承保时间"];

function inferStoreName(fileName = "") {
  const rawBaseName = path.basename(fileName, path.extname(fileName)).trim();
  const baseName = /[ÃÂ]|[\u0080-\u009f]|[èéåçäæ]/u.test(rawBaseName)
    ? Buffer.from(rawBaseName, "latin1").toString("utf8")
    : rawBaseName;
  return baseName
    .replace(/[-_ ]?订单.*$/u, "")
    .replace(/[-_ ]?资金流水.*$/u, "")
    .replace(/[-_ ]?运费险.*$/u, "")
    .trim() || "临时店铺";
}

function missingMessage(label: string, fields: string[]) {
  return `${label}缺少必要字段：${fields.join("、")}。请确认上传的是正确表格。`;
}

app.use(cors());
app.use(express.json({ limit: "5mb" }));

app.get("/api/config", async (_req, res, next) => {
  try {
    res.json(await getConfig());
  } catch (error) {
    next(error);
  }
});

app.put("/api/config/stores", async (req, res, next) => {
  try {
    await saveStores(req.body.stores ?? []);
    res.json(await getConfig());
  } catch (error) {
    next(error);
  }
});

app.put("/api/config/mappings", async (req, res, next) => {
  try {
    await saveMappings(req.body.mappings ?? []);
    res.json(await getConfig());
  } catch (error) {
    next(error);
  }
});

app.get("/api/history", async (_req, res, next) => {
  try {
    res.json(await listSettlementRecords());
  } catch (error) {
    next(error);
  }
});

app.post(
  "/api/settlements",
  upload.fields([
    { name: "orderFile", maxCount: 1 },
    { name: "productCostFile", maxCount: 1 },
    { name: "flowFiles", maxCount: 20 },
    { name: "freightFiles", maxCount: 20 }
  ]),
  async (req, res, next) => {
    const uploadedFiles = Object.values((req.files ?? {}) as Record<string, Express.Multer.File[]>).flat();
    try {
      const { stores, mappings } = await getConfig();
      const files = req.files as Record<string, Express.Multer.File[] | undefined>;
      const orderFile = files.orderFile?.[0];
      const productCostFile = files.productCostFile?.[0];
      const flowFiles = files.flowFiles ?? [];
      const freightFiles = files.freightFiles ?? [];

      if (!orderFile) throw new Error("订单明细表必填。");
      if (flowFiles.length === 0) throw new Error("资金流水明细表必填。");

      const orderExtract = await extractWorkbookRows(orderFile.path, ORDER_FIELDS);
      if (orderExtract.missing.length) throw new Error(missingMessage("订单明细表", orderExtract.missing));
      const flowExtracts = await Promise.all(flowFiles.map((file) => extractWorkbookRows(file.path)));
      flowExtracts.forEach((extract, index) => {
        const missingFlowFields = FLOW_REQUIRED_FIELDS.filter((field) => extract.missing.includes(field));
        if (missingFlowFields.length) throw new Error(missingMessage(`资金流水明细表「${flowFiles[index].originalname}」`, missingFlowFields));
      });
      const productCostExtract = productCostFile ? await extractWorkbookRows(productCostFile.path, PRODUCT_COST_FIELDS) : { rows: [], missing: [] };
      if (productCostExtract.missing.length) throw new Error(missingMessage("商品成本明细表", productCostExtract.missing));
      const freightExtracts = await Promise.all(freightFiles.map((file) => extractWorkbookRows(file.path, FREIGHT_FIELDS)));
      freightExtracts.forEach((extract, index) => {
        const missingFreightFields = ["订单编号", "支付保费", "保费状态"].filter((field) => extract.missing.includes(field));
        if (missingFreightFields.length) throw new Error(missingMessage(`运费险明细表「${freightFiles[index].originalname}」`, missingFreightFields));
      });

      const generatedAt = new Date();
      const store: StoreConfig = {
        id: "mvp-store",
        name: inferStoreName(orderFile.originalname),
        owner: "财务",
        defaultShippingFee: 2.4,
        updatedAt: generatedAt.toISOString()
      };
      const input = {
        store,
        mappings: mappings.filter((item) => item.storeId === store.id),
        shippingFee: Number(req.body.shippingFee || store.defaultShippingFee || 2.4),
        orderRows: orderExtract.rows,
        productCostRows: productCostExtract.rows,
        flowRows: flowExtracts.flatMap((extract) => extract.rows),
        freightRows: freightExtracts.flatMap((extract) => extract.rows),
        uploadedFileNames: uploadedFiles.map((file) => file.originalname),
        generatedAt
      };

      const result = buildSettlement(input);
      const period = result.coveredMonths.orders.length === 1 ? result.coveredMonths.orders[0] : "多月份";
      const { archivePath, version } = await nextSettlementArchive(store, period);
      const archive = await saveSettlementArchive({
        input,
        result,
        archivePath,
        version,
        uploadFiles: uploadedFiles
      });

      res.json({
        record: archive.record,
        stats: result.stats,
        summary: result.summaryRows,
        exceptionSummary: result.exceptionSummary,
        exceptions: result.exceptionRows.slice(0, 50),
        coveredMonths: result.coveredMonths
      });
    } catch (error) {
      next(error);
    } finally {
      await Promise.all(uploadedFiles.map((file) => fs.rm(file.path, { force: true }).catch(() => undefined)));
    }
  }
);

app.get("/api/settlements/:id/download", async (req, res, next) => {
  try {
    const record = await findSettlementRecord(decodeURIComponent(req.params.id));
    if (!record) {
      res.status(404).json({ error: "未找到结算记录。" });
      return;
    }
    const resultDir = path.join(record.archivePath, "result");
    const files = await fs.readdir(resultDir);
    const excel = files.find((file) => file.endsWith(".xlsx"));
    if (!excel) throw new Error("未找到结算 Excel 文件。");
    res.download(path.join(resultDir, excel));
  } catch (error) {
    next(error);
  }
});

app.post("/api/settlements/:id/open-archive", async (req, res, next) => {
  try {
    const record = await findSettlementRecord(decodeURIComponent(req.params.id));
    if (!record) {
      res.status(404).json({ error: "未找到结算记录。" });
      return;
    }
    execFile("explorer", [record.archivePath]);
    res.json({ ok: true });
  } catch (error) {
    next(error);
  }
});

if (process.env.NODE_ENV === "production") {
  app.use(express.static(path.resolve("dist")));
  app.get("*", (_req, res) => res.sendFile(path.resolve("dist/index.html")));
} else {
  const vite = await createViteServer({
    server: { middlewareMode: true },
    appType: "spa"
  });
  app.use(vite.middlewares);
}

app.use((error: unknown, _req: express.Request, res: express.Response, _next: express.NextFunction) => {
  const message = error instanceof Error ? error.message : "服务发生未知错误。";
  res.status(400).json({ error: message });
});

app.listen(port, () => {
  console.log(`抖音店铺达人结算工具已启动：http://localhost:${port}`);
});

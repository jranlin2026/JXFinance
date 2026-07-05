import * as XLSX from "xlsx";
import {
  ExceptionRecord,
  InfluencerMapping,
  OrderDetailRow,
  SettlementConfigSnapshot,
  SettlementStats,
  StoreConfig,
  TalentSettlementSummaryRow
} from "../shared/types";
import { countBy } from "./utils";

type RawRow = Record<string, unknown>;

export type SettlementInput = {
  store: StoreConfig;
  mappings: InfluencerMapping[];
  shippingFee: number;
  orderRows: RawRow[];
  productCostRows: RawRow[];
  flowRows: RawRow[];
  freightRows: RawRow[];
  uploadedFileNames: string[];
  generatedAt?: Date;
};

export type SettlementBuildResult = {
  orderDetailRows: OrderDetailRow[];
  summaryRows: TalentSettlementSummaryRow[];
  flowRows: Record<string, unknown>[];
  flowSummary: FlowSummaryData;
  productCostRows: Record<string, unknown>[];
  freightDeductedRows: Record<string, unknown>[];
  freightUnmatchedRows: Record<string, unknown>[];
  exceptionRows: ExceptionRecord[];
  exceptionSummary: Record<string, number>;
  stats: SettlementStats;
  coveredMonths: {
    orders: string[];
    flow: string[];
    freight: string[];
  };
  workbook: XLSX.WorkBook;
};

type OrderContext = {
  month: string;
  mainOrderId: string;
  subOrderId: string;
  talentName: string;
  talentId: string;
};

type FlowBucket = {
  totalCount: number;
  inCount: number;
  outCount: number;
  inAmount: number;
  outAmount: number;
  netAmount: number;
  orderLinkedCount: number;
  noOrderCount: number;
  fallbackCount: number;
  firstTime: Date | null;
  lastTime: Date | null;
};

export type FlowSummaryData = {
  overviewRows: Record<string, unknown>[];
  sceneRows: Record<string, unknown>[];
  monthSceneRows: Record<string, unknown>[];
  detailRows: Record<string, unknown>[];
};

const XLSXLib: typeof XLSX = (XLSX as unknown as { default?: typeof XLSX }).default ?? XLSX;

const ORDER_DETAIL_HEADERS: (keyof OrderDetailRow)[] = [
  "订单月份",
  "主订单编号",
  "子订单编号",
  "结算到账金额",
  "运费险",
  "商品数量",
  "商家编码",
  "商品单价",
  "订单应付金额",
  "产品单件成本",
  "产品总成本",
  "订单提交时间",
  "订单完成时间",
  "支付完成时间",
  "达人ID",
  "达人昵称",
  "发货时间"
];

const TALENT_SUMMARY_HEADERS: (keyof TalentSettlementSummaryRow)[] = [
  "订单月份",
  "达人昵称",
  "达人ID",
  "实付订单金额",
  "实付订单数",
  "快递包裹数",
  "快递费用",
  "运费险费用",
  "结算到账金额",
  "产品成本",
  "成本总额",
  "毛利润",
  "销售额毛利率"
];

const FLOW_SCENE_HEADERS = [
  "动账场景归类",
  "总笔数",
  "入账笔数",
  "入账金额",
  "出账笔数",
  "出账金额",
  "净额",
  "有订单号笔数",
  "无订单号笔数",
  "备注兜底笔数",
  "最早动账时间",
  "最晚动账时间"
];

const FLOW_MONTH_SCENE_HEADERS = ["月份", ...FLOW_SCENE_HEADERS];

export function cleanId(value: unknown) {
  return String(value ?? "")
    .trim()
    .replace(/^[\t',]+/, "")
    .trim();
}

export function parseMoney(value: unknown): number | null {
  if (value === null || value === undefined) return null;
  const text = String(value).trim();
  if (!text || text === "-" || text.toUpperCase() === "NULL") return null;
  const normalized = text.replace(/[¥￥,+\s]/g, "");
  const parsed = Number(normalized);
  return Number.isFinite(parsed) ? parsed : null;
}

export function parseDateValue(value: unknown): Date | null {
  if (value instanceof Date && !Number.isNaN(value.getTime())) return value;
  if (typeof value === "number" && Number.isFinite(value)) {
    const parsed = XLSXLib.SSF.parse_date_code(value);
    if (!parsed) return null;
    return new Date(parsed.y, parsed.m - 1, parsed.d, parsed.H, parsed.M, Math.floor(parsed.S));
  }
  const text = String(value ?? "").trim();
  if (!text || text === "-" || text.toUpperCase() === "NULL") return null;
  const normalized = text.replace(/\./g, "-").replace(/\//g, "-");
  const parsed = new Date(normalized);
  return Number.isNaN(parsed.getTime()) ? null : parsed;
}

function cleanText(value: unknown) {
  return cleanId(value);
}

function round2(value: number) {
  return Math.round((value + Number.EPSILON) * 100) / 100;
}

function roundRate(value: number) {
  return Math.round((value + Number.EPSILON) * 10000) / 10000;
}

function formatDate(date: Date | null) {
  if (!date) return "";
  const pad = (n: number) => `${n}`.padStart(2, "0");
  return `${date.getFullYear()}-${pad(date.getMonth() + 1)}-${pad(date.getDate())} ${pad(date.getHours())}:${pad(date.getMinutes())}:${pad(date.getSeconds())}`;
}

function monthOf(date: Date | null) {
  if (!date) return "";
  const pad = (n: number) => `${n}`.padStart(2, "0");
  return `${date.getFullYear()}-${pad(date.getMonth() + 1)}`;
}

function money(row: RawRow, field: string) {
  return parseMoney(row[field]) ?? 0;
}

function createFlowBucket(): FlowBucket {
  return {
    totalCount: 0,
    inCount: 0,
    outCount: 0,
    inAmount: 0,
    outAmount: 0,
    netAmount: 0,
    orderLinkedCount: 0,
    noOrderCount: 0,
    fallbackCount: 0,
    firstTime: null,
    lastTime: null
  };
}

function updateFlowBucket(bucket: FlowBucket, params: {
  amount: number;
  direction: string;
  signed: number;
  hasOrder: boolean;
  usedFallback: boolean;
  flowDate: Date | null;
}) {
  bucket.totalCount += 1;
  if (params.direction === "入账") {
    bucket.inCount += 1;
    bucket.inAmount = round2(bucket.inAmount + Math.abs(params.amount));
  } else if (params.direction === "出账") {
    bucket.outCount += 1;
    bucket.outAmount = round2(bucket.outAmount + Math.abs(params.amount));
  }
  bucket.netAmount = round2(bucket.netAmount + params.signed);
  if (params.hasOrder) bucket.orderLinkedCount += 1;
  else bucket.noOrderCount += 1;
  if (params.usedFallback) bucket.fallbackCount += 1;
  if (params.flowDate) {
    if (!bucket.firstTime || params.flowDate < bucket.firstTime) bucket.firstTime = params.flowDate;
    if (!bucket.lastTime || params.flowDate > bucket.lastTime) bucket.lastTime = params.flowDate;
  }
}

function serializeFlowBucket(scene: string, bucket: FlowBucket, month?: string) {
  const row: Record<string, unknown> = {
    动账场景归类: scene,
    总笔数: bucket.totalCount,
    入账笔数: bucket.inCount,
    入账金额: round2(bucket.inAmount),
    出账笔数: bucket.outCount,
    出账金额: round2(bucket.outAmount),
    净额: round2(bucket.netAmount),
    有订单号笔数: bucket.orderLinkedCount,
    无订单号笔数: bucket.noOrderCount,
    备注兜底笔数: bucket.fallbackCount,
    最早动账时间: formatDate(bucket.firstTime),
    最晚动账时间: formatDate(bucket.lastTime)
  };
  return month === undefined ? row : { 月份: month, ...row };
}

function sortFlowSummaryRows(rows: Record<string, unknown>[]) {
  return rows.sort((a, b) => {
    const netDiff = Math.abs(Number(b["净额"] ?? 0)) - Math.abs(Number(a["净额"] ?? 0));
    if (netDiff !== 0) return netDiff;
    return Number(b["总笔数"] ?? 0) - Number(a["总笔数"] ?? 0);
  });
}

function uniqueHeaders(rows: Record<string, unknown>[], preferred: string[] = []) {
  const seen = new Set<string>();
  const headers: string[] = [];
  const add = (key: string) => {
    if (!seen.has(key)) {
      seen.add(key);
      headers.push(key);
    }
  };
  preferred.forEach(add);
  for (const row of rows) Object.keys(row).forEach(add);
  return headers;
}

function optionalMoney(value: unknown): number | "" {
  const parsed = parseMoney(value);
  return parsed === null ? "" : round2(parsed);
}

function addException(exceptions: ExceptionRecord[], record: ExceptionRecord) {
  exceptions.push(record);
}

function standardTalent(row: RawRow, mappings: InfluencerMapping[]) {
  const rawName = cleanText(row["达人昵称"]);
  if (!rawName) return { talentName: "商品卡流量", talentId: "" };
  const mapping = mappings.find((item) => item.enabled && item.originalName === rawName);
  return {
    talentName: mapping?.settlementName || rawName,
    talentId: cleanText(row["达人ID"])
  };
}

function groupKey(context: Pick<OrderContext, "month" | "talentName" | "talentId">) {
  return `${context.month}\u0001${context.talentName}\u0001${context.talentId}`;
}

function splitGroupKey(key: string) {
  const [month, talentName, talentId] = key.split("\u0001");
  return { month, talentName, talentId };
}

function ensureSummaryBucket(map: Map<string, TalentSettlementSummaryRow>, context: Pick<OrderContext, "month" | "talentName" | "talentId">) {
  const key = groupKey(context);
  const existing = map.get(key);
  if (existing) return existing;
  const created: TalentSettlementSummaryRow = {
    订单月份: context.month,
    达人昵称: context.talentName,
    达人ID: context.talentId,
    实付订单金额: 0,
    实付订单数: 0,
    快递包裹数: 0,
    快递费用: 0,
    运费险费用: 0,
    结算到账金额: 0,
    产品成本: 0,
    成本总额: 0,
    毛利润: 0,
    销售额毛利率: ""
  };
  map.set(key, created);
  return created;
}

function buildProductCostMap(rows: RawRow[], exceptions: ExceptionRecord[]) {
  const map = new Map<string, number>();
  for (const row of rows) {
    const code = cleanId(row["商家编码"]);
    if (!code) continue;
    const price = parseMoney(row["产品单价"] ?? row["产品单件成本"] ?? row["产品成本"] ?? row["商品成本"] ?? row["成本单价"]);
    if (price === null) {
      addException(exceptions, {
        type: "商品成本单价无效",
        level: "核对",
        source: "商品成本明细表",
        message: `商家编码“${code}”的成本单价为空或无法识别。`
      });
      continue;
    }
    map.set(code, price);
  }
  return map;
}

function buildFlowTotals(rows: RawRow[], exceptions: ExceptionRecord[]) {
  const subTotals = new Map<string, number>();
  const mainTotals = new Map<string, number>();
  const normalizedRows: Record<string, unknown>[] = [];
  const sceneBuckets = new Map<string, FlowBucket>();
  const monthSceneBuckets = new Map<string, FlowBucket>();
  const monthSet = new Set<string>();
  let inCount = 0;
  let outCount = 0;
  let inAmount = 0;
  let outAmount = 0;
  let blankSceneCount = 0;
  let fallbackCount = 0;
  let noOrderCount = 0;
  let firstTime: Date | null = null;
  let lastTime: Date | null = null;

  for (const row of rows) {
    const amount = parseMoney(row["动账金额"]) ?? 0;
    const direction = cleanText(row["动账方向"]);
    const signed = direction === "出账" ? -amount : direction === "入账" ? amount : 0;
    const subOrderId = cleanId(row["子订单号"]);
    const mainOrderId = cleanId(row["订单号"]);
    const rawScene = cleanText(row["动账场景"]);
    const remark = cleanText(row["备注"]);
    let scene = rawScene;
    let sceneSource = "动账场景";
    let usedFallback = false;
    if (!scene) {
      blankSceneCount += 1;
      if (remark) {
        scene = remark;
        sceneSource = "备注兜底";
        usedFallback = true;
        fallbackCount += 1;
      } else {
        scene = "未归类";
        sceneSource = "未归类";
      }
    }
    const flowDate = parseDateValue(row["动账时间"]);
    const flowMonth = monthOf(flowDate);
    const bucketMonth = flowMonth || "未识别月份";
    const hasOrder = Boolean(subOrderId || mainOrderId);

    if (flowMonth) monthSet.add(flowMonth);
    if (flowDate) {
      if (!firstTime || flowDate < firstTime) firstTime = flowDate;
      if (!lastTime || flowDate > lastTime) lastTime = flowDate;
    }
    if (direction === "入账") {
      inCount += 1;
      inAmount = round2(inAmount + Math.abs(amount));
    } else if (direction === "出账") {
      outCount += 1;
      outAmount = round2(outAmount + Math.abs(amount));
    }
    if (!hasOrder) noOrderCount += 1;

    const sceneBucket = sceneBuckets.get(scene) ?? createFlowBucket();
    updateFlowBucket(sceneBucket, { amount, direction, signed, hasOrder, usedFallback, flowDate });
    sceneBuckets.set(scene, sceneBucket);

    const monthSceneKey = `${bucketMonth}\u0001${scene}`;
    const monthSceneBucket = monthSceneBuckets.get(monthSceneKey) ?? createFlowBucket();
    updateFlowBucket(monthSceneBucket, { amount, direction, signed, hasOrder, usedFallback, flowDate });
    monthSceneBuckets.set(monthSceneKey, monthSceneBucket);

    if (direction !== "入账" && direction !== "出账") {
      addException(exceptions, {
        type: "未知资金流水方向",
        level: "核对",
        source: "资金流水明细表",
        message: `资金流水方向“${direction || "空"}”无法识别，金额未计入。`
      });
    }

    if (subOrderId) {
      subTotals.set(subOrderId, round2((subTotals.get(subOrderId) ?? 0) + signed));
    } else if (mainOrderId) {
      mainTotals.set(mainOrderId, round2((mainTotals.get(mainOrderId) ?? 0) + signed));
    } else {
      addException(exceptions, {
        type: "无订单号资金流水",
        level: "核对",
        source: "资金流水明细表",
        message: `资金流水缺少子订单号和订单号，动账场景=${scene || "空"}。`
      });
    }

    normalizedRows.push({
      ...row,
      归类动账场景: scene,
      归类来源: sceneSource,
      签名净额: round2(signed),
      子订单号: subOrderId,
      订单号: mainOrderId,
      动账场景: scene,
      动账金额: amount,
      "动账金额（出账为负）": round2(signed),
      覆盖月份: monthOf(flowDate)
    });
  }

  const sceneRows = sortFlowSummaryRows([...sceneBuckets.entries()].map(([scene, bucket]) => serializeFlowBucket(scene, bucket)));
  const monthSceneRows = [...monthSceneBuckets.entries()]
    .map(([key, bucket]) => {
      const [month, scene] = key.split("\u0001");
      return serializeFlowBucket(scene, bucket, month);
    })
    .sort((a, b) => String(a["月份"]).localeCompare(String(b["月份"])) || Math.abs(Number(b["净额"] ?? 0)) - Math.abs(Number(a["净额"] ?? 0)));
  const topRows = sceneRows.slice(0, 12);
  const metrics = [
    { 指标: "流水总笔数", 值: rows.length, 说明: "" },
    { 指标: "入账笔数", 值: inCount, 说明: "" },
    { 指标: "入账金额", 值: round2(inAmount), 说明: "" },
    { 指标: "出账笔数", 值: outCount, 说明: "" },
    { 指标: "出账金额", 值: round2(outAmount), 说明: "" },
    { 指标: "净额", 值: round2(inAmount - outAmount), 说明: "入账-出账" },
    { 指标: "动账场景为空", 值: blankSceneCount, 说明: "按备注兜底" },
    { 指标: "备注兜底成功", 值: fallbackCount, 说明: "" },
    { 指标: "无订单号流水", 值: noOrderCount, 说明: "" },
    { 指标: "覆盖月份", 值: [...monthSet].sort().join(", "), 说明: "" },
    { 指标: "最早动账时间", 值: formatDate(firstTime), 说明: "" },
    { 指标: "最晚动账时间", 值: formatDate(lastTime), 说明: "" }
  ];
  const overviewRows = Array.from({ length: Math.max(metrics.length, topRows.length) }, (_, index) => ({
    指标: metrics[index]?.指标 ?? "",
    值: metrics[index]?.值 ?? "",
    说明: metrics[index]?.说明 ?? "",
    Top动账场景: topRows[index]?.["动账场景归类"] ?? "",
    Top净额: topRows[index]?.["净额"] ?? ""
  }));

  return {
    subTotals,
    mainTotals,
    normalizedRows,
    flowSummary: {
      overviewRows,
      sceneRows,
      monthSceneRows,
      detailRows: normalizedRows
    }
  };
}

function buildFreightTotals(rows: RawRow[], exceptions: ExceptionRecord[]) {
  const totals = new Map<string, number>();
  const deductedRows: Record<string, unknown>[] = [];

  for (const row of rows) {
    const status = cleanText(row["保费状态"]);
    const orderId = cleanId(row["订单编号"]);
    const amount = parseMoney(row["支付保费"]) ?? 0;
    const normalized = {
      ...row,
      订单编号: orderId,
      支付保费: amount,
      覆盖月份: monthOf(parseDateValue(row["动账时间"] ?? row["下单时间"] ?? row["承保时间"]))
    };

    if (status === "已扣减") {
      deductedRows.push(normalized);
      if (orderId) totals.set(orderId, round2((totals.get(orderId) ?? 0) + amount));
    } else if (status) {
      addException(exceptions, {
        type: "运费险未扣减",
        level: "提醒",
        orderId,
        source: "运费险明细表",
        message: `运费险记录未计入：保费状态=${status}。`
      });
    }
  }

  return { totals, deductedRows };
}

function buildOrderContexts(orderRows: RawRow[], mappings: InfluencerMapping[]) {
  const contexts = new Map<string, OrderContext>();
  for (const row of orderRows) {
    const submitDate = parseDateValue(row["订单提交时间"]);
    const month = monthOf(submitDate);
    const subOrderId = cleanId(row["子订单编号"]);
    const mainOrderId = cleanId(row["主订单编号"]);
    const { talentName, talentId } = standardTalent(row, mappings);
    if (subOrderId && !contexts.has(subOrderId)) {
      contexts.set(subOrderId, { month, mainOrderId, subOrderId, talentName, talentId });
    }
  }
  return contexts;
}

function applyBasicSheetLayout(sheet: XLSX.WorkSheet, widths: number[]) {
  sheet["!cols"] = widths.map((wch) => ({ wch }));
}

function buildWorkbook(orderDetailRows: OrderDetailRow[], summaryRows: TalentSettlementSummaryRow[], flowSummary: FlowSummaryData) {
  const workbook = XLSXLib.utils.book_new();
  const detailSheet = XLSXLib.utils.json_to_sheet(orderDetailRows, { header: ORDER_DETAIL_HEADERS });
  const summarySheet = XLSXLib.utils.json_to_sheet(summaryRows, { header: TALENT_SUMMARY_HEADERS });
  const flowOverviewSheet = XLSXLib.utils.json_to_sheet(flowSummary.overviewRows, {
    header: ["指标", "值", "说明", "Top动账场景", "Top净额"]
  });
  const flowSceneSheet = XLSXLib.utils.json_to_sheet(flowSummary.sceneRows, { header: FLOW_SCENE_HEADERS });
  const flowMonthSceneSheet = XLSXLib.utils.json_to_sheet(flowSummary.monthSceneRows, { header: FLOW_MONTH_SCENE_HEADERS });
  const flowDetailHeaders = uniqueHeaders(flowSummary.detailRows, ["归类动账场景", "归类来源", "签名净额"]);
  const flowDetailSheet = XLSXLib.utils.json_to_sheet(flowSummary.detailRows, { header: flowDetailHeaders });

  applyBasicSheetLayout(detailSheet, [12, 22, 22, 14, 10, 10, 16, 12, 14, 14, 14, 20, 20, 20, 20, 24, 20]);
  applyBasicSheetLayout(summarySheet, [12, 24, 20, 14, 12, 12, 12, 12, 14, 12, 12, 12, 12]);
  applyBasicSheetLayout(flowOverviewSheet, [18, 18, 18, 42, 14]);
  applyBasicSheetLayout(flowSceneSheet, [42, 12, 12, 14, 12, 14, 14, 14, 14, 14, 20, 20]);
  applyBasicSheetLayout(flowMonthSceneSheet, [12, 42, 12, 12, 14, 12, 14, 14, 14, 14, 14, 20, 20]);
  applyBasicSheetLayout(flowDetailSheet, [34, 14, 14, ...flowDetailHeaders.slice(3).map(() => 18)]);

  XLSXLib.utils.book_append_sheet(workbook, detailSheet, "订单明细融合表");
  XLSXLib.utils.book_append_sheet(workbook, summarySheet, "达人结算汇总表");
  XLSXLib.utils.book_append_sheet(workbook, flowOverviewSheet, "资金流水总览");
  XLSXLib.utils.book_append_sheet(workbook, flowSceneSheet, "资金流水场景汇总");
  XLSXLib.utils.book_append_sheet(workbook, flowMonthSceneSheet, "资金流水月份汇总");
  XLSXLib.utils.book_append_sheet(workbook, flowDetailSheet, "资金流水明细核对");
  return workbook;
}

export function buildSettlement(input: SettlementInput): SettlementBuildResult {
  const exceptions: ExceptionRecord[] = [];
  const hasProductCostSheet = input.productCostRows.length > 0;
  const hasFreightSheet = input.freightRows.length > 0;

  if (!hasProductCostSheet) {
    addException(exceptions, {
      type: "商品成本表未上传",
      level: "提醒",
      source: "商品成本明细表",
      message: "本次未上传商品成本明细表，产品成本按 0 计算。"
    });
  }
  if (!hasFreightSheet) {
    addException(exceptions, {
      type: "运费险表未上传",
      level: "提醒",
      source: "运费险明细表",
      message: "本次未上传运费险明细表，运费险费用按 0 计算。"
    });
  }

  const productCostMap = buildProductCostMap(input.productCostRows, exceptions);
  const { subTotals, mainTotals, normalizedRows: normalizedFlowRows, flowSummary } = buildFlowTotals(input.flowRows, exceptions);
  const { totals: freightTotals, deductedRows: freightDeductedRows } = buildFreightTotals(input.freightRows, exceptions);
  const orderContexts = buildOrderContexts(input.orderRows, input.mappings);

  const summaryBuckets = new Map<string, TalentSettlementSummaryRow>();
  const paidSubOrderKeys = new Set<string>();
  const packageKeys = new Set<string>();
  const freightMatchedIds = new Set<string>();
  const freightUnmatchedRows: Record<string, unknown>[] = [];
  const orderDetailRows: OrderDetailRow[] = [];
  let productCostRows = 0;
  let paidOrderRows = 0;
  let settlementMatchedRows = 0;

  for (const row of input.orderRows) {
    const submitDate = parseDateValue(row["订单提交时间"]);
    const completeDate = parseDateValue(row["订单完成时间"]);
    const payDate = parseDateValue(row["支付完成时间"]);
    const shipDate = parseDateValue(row["发货时间"]);
    const month = monthOf(submitDate);
    const mainOrderId = cleanId(row["主订单编号"]);
    const subOrderId = cleanId(row["子订单编号"]);
    const { talentName, talentId } = standardTalent(row, input.mappings);
    const settlement = subTotals.has(subOrderId)
      ? subTotals.get(subOrderId)
      : mainTotals.has(mainOrderId)
        ? mainTotals.get(mainOrderId)
        : undefined;
    const freightFee = hasFreightSheet && freightTotals.has(subOrderId) ? freightTotals.get(subOrderId) : undefined;
    if (freightFee !== undefined) freightMatchedIds.add(subOrderId);

    const quantity = parseMoney(row["商品数量"]) ?? 0;
    const merchantCode = cleanId(row["商家编码"]);
    const productUnitCost = hasProductCostSheet ? productCostMap.get(merchantCode) : 0;
    if (hasProductCostSheet && productUnitCost === undefined && merchantCode) {
      addException(exceptions, {
        type: "商品成本缺失",
        level: "提醒",
        orderId: mainOrderId,
        subOrderId,
        source: "商品成本明细表",
        message: `商家编码“${merchantCode}”未在商品成本明细表中找到。`
      });
    }

    let productTotalCost = 0;
    if (settlement !== undefined && settlement > 5 && productUnitCost !== undefined) {
      productTotalCost = productUnitCost * quantity;
      productCostRows += 1;
    }

    if (payDate) {
      const bucket = ensureSummaryBucket(summaryBuckets, { month, talentName, talentId });
      const paidAmount =
        money(row, "订单应付金额") +
        money(row, "平台实际承担优惠金额") +
        money(row, "商家实际承担优惠金额") +
        money(row, "达人实际承担优惠金额");
      bucket.实付订单金额 = round2(bucket.实付订单金额 + paidAmount);
      if (subOrderId && !paidSubOrderKeys.has(subOrderId)) {
        paidSubOrderKeys.add(subOrderId);
        bucket.实付订单数 += 1;
      }
      paidOrderRows += 1;
    }

    if (shipDate && mainOrderId) {
      const bucket = ensureSummaryBucket(summaryBuckets, { month, talentName, talentId });
      const packageKey = `${groupKey({ month, talentName, talentId })}\u0001${mainOrderId}\u0001${formatDate(shipDate)}`;
      if (!packageKeys.has(packageKey)) {
        packageKeys.add(packageKey);
        bucket.快递包裹数 += 1;
      }
    }

    if (settlement !== undefined) {
      const bucket = ensureSummaryBucket(summaryBuckets, { month, talentName, talentId });
      settlementMatchedRows += 1;
      bucket.结算到账金额 = round2(bucket.结算到账金额 + settlement);
      bucket.产品成本 = round2(bucket.产品成本 + productTotalCost);
    }

    const detail: OrderDetailRow = {
      订单月份: month,
      主订单编号: mainOrderId,
      子订单编号: subOrderId,
      结算到账金额: settlement === undefined ? "" : round2(settlement),
      运费险: freightFee === undefined ? "" : round2(freightFee),
      商品数量: quantity || "",
      商家编码: merchantCode,
      商品单价: optionalMoney(row["商品单价"]),
      订单应付金额: optionalMoney(row["订单应付金额"]),
      产品单件成本: productUnitCost === undefined ? "" : round2(productUnitCost),
      产品总成本: productTotalCost > 0 ? round2(productTotalCost) : "",
      订单提交时间: formatDate(submitDate),
      订单完成时间: formatDate(completeDate),
      支付完成时间: formatDate(payDate),
      达人ID: talentId,
      达人昵称: talentName,
      发货时间: formatDate(shipDate)
    };
    orderDetailRows.push(detail);
  }

  for (const row of freightDeductedRows) {
    const orderId = cleanId(row["订单编号"]);
    const context = orderContexts.get(orderId);
    if (!context) {
      freightUnmatchedRows.push(row);
      continue;
    }
    freightMatchedIds.add(orderId);
    const bucket = ensureSummaryBucket(summaryBuckets, {
      month: context.month,
      talentName: context.talentName,
      talentId: context.talentId
    });
    bucket.运费险费用 = round2(bucket.运费险费用 + (parseMoney(row["支付保费"]) ?? 0));
  }

  if (freightUnmatchedRows.length > 0) {
    addException(exceptions, {
      type: "运费险未匹配子订单",
      level: "核对",
      source: "运费险明细表",
      message: `${freightUnmatchedRows.length} 条已扣减运费险未匹配到订单明细表子订单编号。`
    });
  }

  const summaryRows = [...summaryBuckets.values()]
    .map((row) => {
      row.快递费用 = round2(row.快递包裹数 * input.shippingFee);
      row.成本总额 = round2(row.快递费用 + row.运费险费用 + row.产品成本);
      row.毛利润 = round2(row.结算到账金额 - row.成本总额);
      row.销售额毛利率 = row.实付订单金额 === 0 ? "" : roundRate(row.毛利润 / row.实付订单金额);
      return row;
    })
    .sort((a, b) =>
      a.订单月份.localeCompare(b.订单月份) ||
      b.实付订单金额 - a.实付订单金额 ||
      a.达人昵称.localeCompare(b.达人昵称, "zh-Hans-CN")
    );

  const coveredMonths = {
    orders: [...new Set(orderDetailRows.map((row) => row.订单月份).filter(Boolean))].sort(),
    flow: [...new Set(normalizedFlowRows.map((row) => String(row["覆盖月份"] ?? "")).filter(Boolean))].sort(),
    freight: [...new Set(freightDeductedRows.map((row) => String(row["覆盖月份"] ?? "")).filter(Boolean))].sort()
  };

  const totals = summaryRows.reduce(
    (acc, row) => {
      acc.paidOrderAmount += row.实付订单金额;
      acc.freightInsuranceFee += row.运费险费用;
      acc.settlementAmount += row.结算到账金额;
      acc.productCost += row.产品成本;
      acc.shippingFee += row.快递费用;
      acc.totalCost += row.成本总额;
      acc.grossProfit += row.毛利润;
      return acc;
    },
    {
      paidOrderAmount: 0,
      freightInsuranceFee: 0,
      settlementAmount: 0,
      productCost: 0,
      shippingFee: 0,
      totalCost: 0,
      grossProfit: 0
    }
  );

  const stats: SettlementStats = {
    originalOrderRows: input.orderRows.length,
    orderDetailRows: orderDetailRows.length,
    talentSummaryRows: summaryRows.length,
    paidOrderRows,
    paidOrderCount: paidSubOrderKeys.size,
    shippedPackageCount: packageKeys.size,
    settlementMatchedRows,
    flowRows: input.flowRows.length,
    freightDeductedRows: freightDeductedRows.length,
    freightUnmatchedRows: freightUnmatchedRows.length,
    productCostRows,
    exceptionCount: exceptions.length,
    totals: {
      paidOrderAmount: round2(totals.paidOrderAmount),
      freightInsuranceFee: round2(totals.freightInsuranceFee),
      settlementAmount: round2(totals.settlementAmount),
      productCost: round2(totals.productCost),
      shippingFee: round2(totals.shippingFee),
      totalCost: round2(totals.totalCost),
      grossProfit: round2(totals.grossProfit)
    }
  };

  return {
    orderDetailRows,
    summaryRows,
    flowRows: normalizedFlowRows,
    flowSummary,
    productCostRows: input.productCostRows,
    freightDeductedRows,
    freightUnmatchedRows,
    exceptionRows: exceptions,
    exceptionSummary: countBy(exceptions, (item) => item.type),
    stats,
    coveredMonths,
    workbook: buildWorkbook(orderDetailRows, summaryRows, flowSummary)
  };
}

export function snapshotConfig(store: StoreConfig, mappings: InfluencerMapping[], shippingFee: number): SettlementConfigSnapshot {
  return {
    store,
    mappings,
    shippingFee
  };
}

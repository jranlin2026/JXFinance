import { describe, expect, it } from "vitest";
import { buildSettlement, cleanId, parseMoney } from "./settlementEngine";
import { InfluencerMapping, StoreConfig } from "../shared/types";

const store: StoreConfig = {
  id: "store-1",
  name: "测试店铺",
  owner: "财务",
  defaultShippingFee: 2.4,
  updatedAt: "2026-07-01T00:00:00.000Z"
};

const mappings: InfluencerMapping[] = [];

describe("settlement engine v2", () => {
  it("cleans ids and money fields", () => {
    expect(cleanId("\t'6926234509777272695")).toBe("6926234509777272695");
    expect(cleanId(",6926234509777272695")).toBe("6926234509777272695");
    expect(parseMoney("¥1,230.50")).toBe(1230.5);
  });

  it("builds the standard settlement workbook with fund flow summary sheets", () => {
    const result = buildSettlement({
      store,
      mappings,
      shippingFee: 2.4,
      uploadedFileNames: ["orders.xlsx", "flows.xlsx"],
      generatedAt: new Date("2026-07-01T10:00:00.000Z"),
      productCostRows: [],
      freightRows: [],
      orderRows: [
        {
          主订单编号: "M1",
          子订单编号: "S1",
          商品数量: "1",
          商家编码: "SKU1",
          商品单价: "100",
          订单应付金额: "100",
          平台实际承担优惠金额: "5",
          商家实际承担优惠金额: "3",
          达人实际承担优惠金额: "2",
          订单提交时间: "2026-05-10 10:00:00",
          支付完成时间: "2026-05-10 10:01:00",
          发货时间: "2026-05-11 09:00:00",
          达人ID: "",
          达人昵称: ""
        }
      ],
      flowRows: [
        { 子订单号: "S1", 订单号: "M1", 动账方向: "入账", 动账金额: "80", 动账场景: "货款结算入账", 动账时间: "2026-06-01 12:00:00" },
        { 子订单号: "", 订单号: "", 动账方向: "出账", 动账金额: "2", 动账场景: "", 备注: "保费扣除", 动账时间: "2026-06-01 12:01:00" }
      ]
    });

    expect(result.workbook.SheetNames).toEqual([
      "订单明细融合表",
      "达人结算汇总表",
      "资金流水总览",
      "资金流水场景汇总",
      "资金流水月份汇总",
      "资金流水明细核对"
    ]);
    expect(result.flowSummary.overviewRows).toEqual(
      expect.arrayContaining([
        expect.objectContaining({ 指标: "流水总笔数", 值: 2 }),
        expect.objectContaining({ 指标: "入账金额", 值: 80 }),
        expect.objectContaining({ 指标: "出账金额", 值: 2 }),
        expect.objectContaining({ 指标: "净额", 值: 78 }),
        expect.objectContaining({ 指标: "备注兜底成功", 值: 1 })
      ])
    );
    expect(result.flowSummary.sceneRows).toEqual(
      expect.arrayContaining([
        expect.objectContaining({ 动账场景归类: "货款结算入账", 入账金额: 80, 净额: 80 }),
        expect.objectContaining({ 动账场景归类: "保费扣除", 出账金额: 2, 净额: -2, 备注兜底笔数: 1 })
      ])
    );
    expect(result.orderDetailRows[0].达人昵称).toBe("商品卡流量");
    expect(result.orderDetailRows[0].运费险).toBe("");
    expect(result.orderDetailRows[0].产品总成本).toBe("");
    expect(result.summaryRows[0]).toMatchObject({
      订单月份: "2026-05",
      达人昵称: "商品卡流量",
      实付订单金额: 110,
      实付订单数: 1,
      快递包裹数: 1,
      快递费用: 2.4,
      运费险费用: 0,
      结算到账金额: 80,
      产品成本: 0,
      成本总额: 2.4,
      毛利润: 77.6
    });
    expect(result.exceptionSummary["商品成本表未上传"]).toBe(1);
    expect(result.exceptionSummary["运费险表未上传"]).toBe(1);
  });

  it("matches freight insurance only by sub order id and never by main order id", () => {
    const result = buildSettlement({
      store,
      mappings,
      shippingFee: 0,
      uploadedFileNames: ["orders.xlsx", "flows.xlsx", "cost.xlsx", "freight.xlsx"],
      generatedAt: new Date("2026-07-01T10:00:00.000Z"),
      productCostRows: [{ 商家编码: "SKU1", 产品单价: "20" }],
      orderRows: [
        {
          主订单编号: "M1",
          子订单编号: "S1",
          商品数量: "2",
          商家编码: "SKU1",
          商品单价: "100",
          订单应付金额: "100",
          订单提交时间: "2026-05-10 10:00:00",
          支付完成时间: "2026-05-10 10:01:00",
          发货时间: "",
          达人ID: "T1",
          达人昵称: "达人A"
        },
        {
          主订单编号: "M1",
          子订单编号: "S2",
          商品数量: "1",
          商家编码: "SKU1",
          商品单价: "100",
          订单应付金额: "80",
          订单提交时间: "2026-05-10 10:02:00",
          支付完成时间: "2026-05-10 10:03:00",
          发货时间: "",
          达人ID: "T1",
          达人昵称: "达人A"
        }
      ],
      flowRows: [
        { 子订单号: "S1", 订单号: "M1", 动账方向: "入账", 动账金额: "70", 动账时间: "2026-06-01 12:00:00" },
        { 子订单号: "S2", 订单号: "M1", 动账方向: "入账", 动账金额: "4", 动账时间: "2026-06-01 12:00:00" }
      ],
      freightRows: [
        { 订单编号: "S1", 支付保费: "3.2", 保费状态: "已扣减", 动账时间: "2026-05-10" },
        { 订单编号: "M1", 支付保费: "9.9", 保费状态: "已扣减", 动账时间: "2026-05-10" },
        { 订单编号: "S2", 支付保费: "1.1", 保费状态: "不扣减", 动账时间: "2026-05-10" }
      ]
    });

    expect(result.orderDetailRows.map((row) => row.运费险)).toEqual([3.2, ""]);
    expect(result.summaryRows[0].运费险费用).toBe(3.2);
    expect(result.summaryRows[0].产品成本).toBe(40);
    expect(result.stats.freightUnmatchedRows).toBe(1);
    expect(result.exceptionSummary["运费险未匹配子订单"]).toBe(1);
  });

  it("uses main order fund flow only when child order id is blank", () => {
    const result = buildSettlement({
      store,
      mappings,
      shippingFee: 0,
      uploadedFileNames: ["orders.xlsx", "flows.xlsx", "cost.xlsx"],
      generatedAt: new Date("2026-07-01T10:00:00.000Z"),
      productCostRows: [{ 商家编码: "SKU1", 产品单价: "10" }],
      freightRows: [],
      orderRows: [
        {
          主订单编号: "M1",
          子订单编号: "S1",
          商品数量: "1",
          商家编码: "SKU1",
          订单应付金额: "100",
          订单提交时间: "2026-05-10 10:00:00",
          支付完成时间: "",
          发货时间: "",
          达人ID: "T1",
          达人昵称: "达人A"
        },
        {
          主订单编号: "M2",
          子订单编号: "S2",
          商品数量: "1",
          商家编码: "SKU1",
          订单应付金额: "100",
          订单提交时间: "2026-05-10 10:00:00",
          支付完成时间: "",
          发货时间: "",
          达人ID: "T1",
          达人昵称: "达人A"
        }
      ],
      flowRows: [
        { 子订单号: "S1", 订单号: "M1", 动账方向: "入账", 动账金额: "50", 动账时间: "2026-06-01 12:00:00" },
        { 子订单号: "", 订单号: "M2", 动账方向: "入账", 动账金额: "30", 动账时间: "2026-06-01 12:00:00" },
        { 子订单号: "S999", 订单号: "M2", 动账方向: "出账", 动账金额: "5", 动账时间: "2026-06-01 12:00:00" }
      ]
    });

    expect(result.orderDetailRows.map((row) => row.结算到账金额)).toEqual([50, 30]);
    expect(result.summaryRows[0].结算到账金额).toBe(80);
  });

  it("counts shipping packages by month talent main order and ship time", () => {
    const orderRows = Array.from({ length: 5 }, (_, index) => ({
      主订单编号: "M40",
      子订单编号: `S4${index}`,
      商品数量: "1",
      商家编码: "SKU1",
      订单应付金额: "100",
      订单提交时间: "2026-05-10 10:00:00",
      支付完成时间: "2026-05-10 10:05:00",
      发货时间: index < 3 ? "2026-05-11 09:00:00" : "2026-05-12 15:30:00",
      达人ID: "T1",
      达人昵称: "达人A"
    }));

    const result = buildSettlement({
      store,
      mappings,
      shippingFee: 2.4,
      uploadedFileNames: ["orders.xlsx", "flows.xlsx"],
      generatedAt: new Date("2026-07-01T10:00:00.000Z"),
      productCostRows: [],
      freightRows: [],
      orderRows,
      flowRows: orderRows.map((row) => ({
        子订单号: row.子订单编号,
        动账方向: "入账",
        动账金额: "80",
        动账时间: "2026-06-01 12:00:00"
      }))
    });

    expect(result.summaryRows[0].快递包裹数).toBe(2);
    expect(result.summaryRows[0].快递费用).toBe(4.8);
  });
});

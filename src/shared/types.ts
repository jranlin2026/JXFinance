export type StoreConfig = {
  id: string;
  name: string;
  owner: string;
  defaultShippingFee: number;
  note?: string;
  updatedAt: string;
};

export type InfluencerMapping = {
  id: string;
  storeId: string;
  originalName: string;
  settlementName: string;
  enabled: boolean;
  note?: string;
  updatedAt: string;
};

export type ExceptionRecord = {
  type: string;
  level: "提醒" | "核对" | "风险";
  orderId?: string;
  subOrderId?: string;
  source?: string;
  message: string;
};

export type OrderDetailRow = {
  订单月份: string;
  主订单编号: string;
  子订单编号: string;
  结算到账金额: number | "";
  运费险: number | "";
  商品数量: number | "";
  商家编码: string;
  商品单价: number | "";
  订单应付金额: number | "";
  产品单件成本: number | "";
  产品总成本: number | "";
  订单提交时间: string;
  订单完成时间: string;
  支付完成时间: string;
  达人ID: string;
  达人昵称: string;
  发货时间: string;
};

export type TalentSettlementSummaryRow = {
  订单月份: string;
  达人昵称: string;
  达人ID: string;
  实付订单金额: number;
  实付订单数: number;
  快递包裹数: number;
  快递费用: number;
  运费险费用: number;
  结算到账金额: number;
  产品成本: number;
  成本总额: number;
  毛利润: number;
  销售额毛利率: number | "";
};

export type SettlementStats = {
  originalOrderRows: number;
  orderDetailRows: number;
  talentSummaryRows: number;
  paidOrderRows: number;
  paidOrderCount: number;
  shippedPackageCount: number;
  settlementMatchedRows: number;
  flowRows: number;
  freightDeductedRows: number;
  freightUnmatchedRows: number;
  productCostRows: number;
  exceptionCount: number;
  totals: {
    paidOrderAmount: number;
    freightInsuranceFee: number;
    settlementAmount: number;
    productCost: number;
    shippingFee: number;
    totalCost: number;
    grossProfit: number;
  };
};

export type SettlementRecord = {
  id: string;
  storeId: string;
  storeName: string;
  month: string;
  version: number;
  generatedAt: string;
  archivePath: string;
  downloadUrl: string;
  uploadFileCount: number;
  stats: SettlementStats;
  summary: TalentSettlementSummaryRow[];
  exceptionSummary: Record<string, number>;
  coveredMonths: {
    orders: string[];
    flow: string[];
    freight: string[];
  };
};

export type SettlementConfigSnapshot = {
  store: StoreConfig;
  mappings: InfluencerMapping[];
  shippingFee: number;
};

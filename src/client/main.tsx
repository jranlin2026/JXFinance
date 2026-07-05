import React, { useEffect, useState } from "react";
import { createRoot } from "react-dom/client";
import { AlertTriangle, CheckCircle2, Download, FileSpreadsheet, History, RotateCcw, UploadCloud } from "lucide-react";
import { SettlementRecord } from "../shared/types";
import "./styles.css";

type GenerateResponse = {
  record: SettlementRecord;
  summary: Record<string, unknown>[];
  exceptionSummary: Record<string, number>;
  exceptions: Record<string, unknown>[];
  coveredMonths: { orders: string[]; flow: string[]; freight: string[] };
};

const formatCurrency = (value: unknown) =>
  new Intl.NumberFormat("zh-CN", { style: "currency", currency: "CNY", maximumFractionDigits: 2 }).format(Number(value || 0));

function FileInputRow({
  label,
  hint,
  required,
  multiple,
  files,
  onChange
}: {
  label: string;
  hint?: string;
  required?: boolean;
  multiple?: boolean;
  files: File[];
  onChange: (files: File[]) => void;
}) {
  return (
    <label className="upload-row">
      <span className="upload-title">
        <FileSpreadsheet size={18} />
        <span>
          {label}
          {hint && <small>{hint}</small>}
        </span>
        {required && <strong>必填</strong>}
      </span>
      <span className="upload-control">
        <UploadCloud size={18} />
        {files.length ? files.map((file) => file.name).join("，") : multiple ? "选择一个或多个 Excel 文件" : "选择 Excel 文件"}
      </span>
      <input
        type="file"
        accept=".xlsx,.xls,.csv"
        multiple={multiple}
        onChange={(event) => onChange(Array.from(event.target.files ?? []))}
      />
    </label>
  );
}

function App() {
  const [history, setHistory] = useState<SettlementRecord[]>([]);
  const [orderFiles, setOrderFiles] = useState<File[]>([]);
  const [productCostFiles, setProductCostFiles] = useState<File[]>([]);
  const [flowFiles, setFlowFiles] = useState<File[]>([]);
  const [freightFiles, setFreightFiles] = useState<File[]>([]);
  const [isGenerating, setIsGenerating] = useState(false);
  const [error, setError] = useState("");
  const [result, setResult] = useState<GenerateResponse | null>(null);

  async function refreshHistory() {
    const response = await fetch("/api/history");
    setHistory(await response.json());
  }

  useEffect(() => {
    refreshHistory();
  }, []);

  async function generate() {
    setError("");
    setResult(null);
    setIsGenerating(true);
    const formData = new FormData();
    if (orderFiles[0]) formData.append("orderFile", orderFiles[0]);
    if (productCostFiles[0]) formData.append("productCostFile", productCostFiles[0]);
    flowFiles.forEach((file) => formData.append("flowFiles", file));
    freightFiles.forEach((file) => formData.append("freightFiles", file));

    try {
      const response = await fetch("/api/settlements", { method: "POST", body: formData });
      const body = await response.json();
      if (!response.ok) throw new Error(body.error || "生成失败。");
      setResult(body);
      await refreshHistory();
    } catch (caught) {
      setError(caught instanceof Error ? caught.message : "生成失败。");
    } finally {
      setIsGenerating(false);
    }
  }

  function clearUploads() {
    setOrderFiles([]);
    setProductCostFiles([]);
    setFlowFiles([]);
    setFreightFiles([]);
    setResult(null);
    setError("");
  }

  const canGenerate = orderFiles.length > 0 && flowFiles.length > 0 && !isGenerating;
  const recent = history.slice(0, 6);

  return (
    <main className="simple-shell">
      <header className="simple-header">
        <div className="brand-mark">抖</div>
        <div>
          <strong>抖音店铺结算工具</strong>
          <span>上传表格，直接生成结算输出表</span>
        </div>
      </header>

      <section className="simple-grid">
        <div className="settlement-form">
          <div className="section-heading">
            <div>
              <h1>生成结算输出表</h1>
              <p>无需选择店铺。默认快递成本按每个包裹 2.4 元计算，输出包含订单、达人结算和资金流水核对 6 个 sheet。</p>
            </div>
            <span className="local-pill">
              <CheckCircle2 size={16} />
              本地文件存储
            </span>
          </div>

          <div className="upload-list no-top-gap">
            <FileInputRow label="订单明细表" required files={orderFiles} onChange={(files) => setOrderFiles(files.slice(0, 1))} />
            <FileInputRow label="资金流水明细表" required multiple files={flowFiles} onChange={setFlowFiles} />
            <FileInputRow
              label="商品成本明细表"
              hint="固定字段：商家编码、产品成本"
              files={productCostFiles}
              onChange={(files) => setProductCostFiles(files.slice(0, 1))}
            />
            <FileInputRow label="运费险明细表" multiple files={freightFiles} onChange={setFreightFiles} />
          </div>

          {error && (
            <div className="inline-error">
              <AlertTriangle size={18} />
              {error}
            </div>
          )}

          <div className="actions">
            <button className="primary" onClick={generate} disabled={!canGenerate}>
              <FileSpreadsheet size={18} />
              {isGenerating ? "正在生成..." : "生成结算输出表"}
            </button>
            <button className="secondary" onClick={clearUploads}>
              <RotateCcw size={18} />
              清空上传
            </button>
          </div>

          <section className="history-preview">
            <div className="table-title">
              <History size={18} />
              最近生成记录
            </div>
            <table>
              <thead>
                <tr>
                  <th>店铺</th>
                  <th>月份</th>
                  <th>生成时间</th>
                  <th>异常</th>
                  <th>操作</th>
                </tr>
              </thead>
              <tbody>
                {recent.map((record) => (
                  <tr key={record.id}>
                    <td>{record.storeName}</td>
                    <td>{record.month}</td>
                    <td>{new Date(record.generatedAt).toLocaleString("zh-CN")}</td>
                    <td>{record.stats.exceptionCount}</td>
                    <td>
                      <a href={record.downloadUrl}>
                        <Download size={15} />
                        下载
                      </a>
                    </td>
                  </tr>
                ))}
                {recent.length === 0 && (
                  <tr>
                    <td colSpan={5}>还没有生成记录。</td>
                  </tr>
                )}
              </tbody>
            </table>
          </section>
        </div>

        <aside className="result-panel">
          <div className="panel-heading">
            <FileSpreadsheet size={20} />
            生成结果
          </div>
          {!result && <p className="empty-state">生成后这里会显示结算摘要和下载入口。</p>}
          {result && (
            <>
              <div className="result-meta">
                <strong>{result.record.storeName}</strong>
                <span>{result.record.month} · v{result.record.version}</span>
              </div>
              <div className="result-totals">
                <span>结算到账金额</span>
                <strong>{formatCurrency(result.record.stats.totals.settlementAmount)}</strong>
                <span>毛利润</span>
                <strong>{formatCurrency(result.record.stats.totals.grossProfit)}</strong>
              </div>
              <div className="covered-months">
                <span>订单月份：{result.record.coveredMonths.orders.join("，") || "未识别"}</span>
                <span>流水覆盖：{result.coveredMonths.flow.join("，") || "未识别"}</span>
                <span>运费险覆盖：{result.coveredMonths.freight.join("，") || "未上传"}</span>
              </div>
              <div className="exception-box">
                <h2>异常摘要</h2>
                {Object.entries(result.exceptionSummary).map(([type, count]) => (
                  <span key={type}>
                    {type}：{count}
                  </span>
                ))}
                {Object.keys(result.exceptionSummary).length === 0 && <span>无异常</span>}
              </div>
              <a className="download-button" href={result.record.downloadUrl}>
                <Download size={18} />
                下载 Excel
              </a>
            </>
          )}
        </aside>
      </section>
    </main>
  );
}

createRoot(document.getElementById("root")!).render(<App />);

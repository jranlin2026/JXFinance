import { describe, expect, it } from "vitest";
import { inferStoreNameFromFileName, normalizeStoreName, repairMojibake } from "./textEncoding";

describe("text encoding repair", () => {
  it("repairs uploaded Chinese filenames that were decoded as latin1", () => {
    const broken = Buffer.from("雅蔓服饰订单7.1导出.xlsx", "utf8").toString("latin1");
    expect(repairMojibake(broken)).toBe("雅蔓服饰订单7.1导出.xlsx");
  });

  it("infers store names after repairing mojibake", () => {
    const broken = Buffer.from("花旦服饰精选订单7.1导出.xlsx", "utf8").toString("latin1");
    expect(inferStoreNameFromFileName(broken)).toBe("花旦服饰精选");
  });

  it("normalizes old history store names that already include order suffixes", () => {
    const broken = Buffer.from("雅蔓服饰订单7.1导出", "utf8").toString("latin1");
    expect(normalizeStoreName(broken)).toBe("雅蔓服饰");
  });
});

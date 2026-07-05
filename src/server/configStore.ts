import path from "node:path";
import { InfluencerMapping, StoreConfig } from "../shared/types";
import { ensureDir, readJson, writeJson } from "./utils";

const configDir = path.resolve("data/config");
const storesPath = path.join(configDir, "stores.json");
const mappingsPath = path.join(configDir, "influencer-mapping.json");

const now = () => new Date().toISOString();

const defaultStores: StoreConfig[] = [
  {
    id: "chengyi-youxuan",
    name: "澄艺优选服饰",
    owner: "财务",
    defaultShippingFee: 2.4,
    note: "默认店铺，可在页面中修改。",
    updatedAt: now()
  }
];

const defaultMappings: InfluencerMapping[] = [
  {
    id: "sample-1",
    storeId: "chengyi-youxuan",
    originalName: "商品卡",
    settlementName: "商品卡流量",
    enabled: true,
    note: "示例映射",
    updatedAt: now()
  }
];

export async function ensureDefaultConfig() {
  await ensureDir(configDir);
  const stores = await readJson<StoreConfig[] | null>(storesPath, null);
  const mappings = await readJson<InfluencerMapping[] | null>(mappingsPath, null);
  if (!stores) await writeJson(storesPath, defaultStores);
  if (!mappings) await writeJson(mappingsPath, defaultMappings);
}

export async function getConfig() {
  await ensureDefaultConfig();
  return {
    stores: await readJson<StoreConfig[]>(storesPath, defaultStores),
    mappings: await readJson<InfluencerMapping[]>(mappingsPath, defaultMappings)
  };
}

export async function saveStores(stores: StoreConfig[]) {
  await writeJson(
    storesPath,
    stores.map((store) => ({ ...store, updatedAt: store.updatedAt || now() }))
  );
}

export async function saveMappings(mappings: InfluencerMapping[]) {
  await writeJson(
    mappingsPath,
    mappings.map((mapping) => ({ ...mapping, updatedAt: mapping.updatedAt || now() }))
  );
}

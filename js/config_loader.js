/**
 * 永源化工 — 自動配櫃系統 Web 版 JSON 參數設定讀取、驗證與匯出模組 (config_loader.js)
 * 完整對應 Python 版 config_loader.py 之資料結構與商業邏輯檢查。
 * 支援美元 (USD: Cu USD/kg, Au/Ag/Pd USD/oz) 與日圓 (JPY: Cu ¥/T, Au/Ag/Pd ¥/g) 雙幣別計價。
 */

(function (global) {
  "use strict";

  const REQUIRED_METALS = ["Cu", "Au", "Ag", "Pd"];
  const VALID_CUSTOMER_TYPES = ["A", "B", "C", "D", "E"];
  const VALID_A_MODES = ["with_pm", "cu_only"];
  const VALID_CURRENCIES = ["USD", "JPY"];

  const REQUIRED_GLOBAL_KEYS = [
    "fill_container",
    "full_container_range",
    "w_min_kg",
    "w_max_kg",
    "target_cu_range",
    "max_in_container_diff_cu",
    "max_total_value",
    "banned_codes",
    "must_batch_ids",
    "must_feed_codes",
    "must_bag_ids",
    "exclude_pm_codes",
    "exclude_pm_materials",
    "split_batch_penalty",
    "price_currency",
    "troy_oz_to_g",
    "metal_prices",
    "metal_prices_jpy",
    "a_mode",
    "a_thresholds",
    "a_rc",
    "a_freight_per_dmt",
    "a_moisture_limit",
    "a_moisture_penalty_per_wmt",
    "a_cu_below_threshold_tc_per_wmt",
    "a_assay_fee_if_below",
    "b_freight_per_kg",
    "b_deduction_per_ton",
    "standard_tier_tables",
  ];

  class ConfigValidationError extends Error {
    constructor(message) {
      super(message);
      this.name = "ConfigValidationError";
    }
  }

  class TierRule {
    constructor(min_grade, max_grade, payable_rate) {
      this.min_grade = Number(min_grade);
      this.max_grade = Number(max_grade);
      this.payable_rate = Number(payable_rate);
    }

    toDict() {
      return {
        min_grade: Number(this.min_grade),
        max_grade: Number(this.max_grade),
        payable_rate: Number(this.payable_rate),
      };
    }
  }

  function tierTablesToDict(tierTables) {
    const result = {};
    for (const metal of Object.keys(tierTables || {})) {
      result[metal] = (tierTables[metal] || []).map((r) =>
        r instanceof TierRule
          ? r.toDict()
          : {
              min_grade: Number(r.min_grade),
              max_grade: Number(r.max_grade),
              payable_rate: Number(r.payable_rate),
            }
      );
    }
    return result;
  }

  class CustomerConfig {
    constructor(props) {
      Object.assign(this, props);
    }

    getActiveWeightBounds() {
      if (this.fill_container) {
        return [this.full_container_range[0], this.full_container_range[1]];
      }
      return [this.w_min_kg, this.w_max_kg];
    }

    getActiveMetalPrices() {
      return this.price_currency === "JPY" ? this.metal_prices_jpy : this.metal_prices;
    }

    getMetalPriceUnitLabels() {
      if (this.price_currency === "JPY") {
        return { Cu: "¥/T", Au: "¥/g", Ag: "¥/g", Pd: "¥/g" };
      }
      return { Cu: "USD/kg", Au: "USD/oz", Ag: "USD/oz", Pd: "USD/oz" };
    }

    getCustomerANetPricesPerKgAndG() {
      const pmFactor = this.a_mode === "cu_only" ? 0.0 : 1.0;
      if (this.price_currency === "JPY") {
        return {
          Cu: (this.metal_prices_jpy.Cu - this.a_rc.Cu) / 1000.0,
          Au: (this.metal_prices_jpy.Au - this.a_rc.Au) * pmFactor,
          Ag: (this.metal_prices_jpy.Ag - this.a_rc.Ag) * pmFactor,
          Pd: (this.metal_prices_jpy.Pd - this.a_rc.Pd) * pmFactor,
        };
      }
      return {
        Cu: this.metal_prices.Cu - this.a_rc.Cu,
        Au: ((this.metal_prices.Au - this.a_rc.Au) / this.troy_oz_to_g) * pmFactor,
        Ag: ((this.metal_prices.Ag - this.a_rc.Ag) / this.troy_oz_to_g) * pmFactor,
        Pd: ((this.metal_prices.Pd - this.a_rc.Pd) / this.troy_oz_to_g) * pmFactor,
      };
    }

    getTierCustomerPricesPerKgAndG() {
      if (this.price_currency === "JPY") {
        return {
          Cu: this.metal_prices_jpy.Cu / 1000.0,
          Au: this.metal_prices_jpy.Au,
          Ag: this.metal_prices_jpy.Ag,
          Pd: this.metal_prices_jpy.Pd,
        };
      }
      return {
        Cu: this.metal_prices.Cu,
        Au: this.metal_prices.Au / this.troy_oz_to_g,
        Ag: this.metal_prices.Ag / this.troy_oz_to_g,
        Pd: this.metal_prices.Pd / this.troy_oz_to_g,
      };
    }

    toDict() {
      return {
        customer_type: this.customer_type,
        price_currency: this.price_currency,
        fill_container: Boolean(this.fill_container),
        full_container_range: [
          Number(this.full_container_range[0]),
          Number(this.full_container_range[1]),
        ],
        w_min_kg: Number(this.w_min_kg),
        w_max_kg: Number(this.w_max_kg),
        target_cu_range: [Number(this.target_cu_range[0]), Number(this.target_cu_range[1])],
        max_in_container_diff_cu: Number(this.max_in_container_diff_cu),
        max_total_value:
          this.max_total_value === Infinity || this.max_total_value === null
            ? null
            : Number(this.max_total_value),
        banned_codes: [...this.banned_codes],
        must_batch_ids: [...this.must_batch_ids],
        must_feed_codes: [...this.must_feed_codes],
        must_bag_ids: [...this.must_bag_ids],
        exclude_pm_codes: [...this.exclude_pm_codes],
        exclude_pm_materials: [...this.exclude_pm_materials],
        split_batch_penalty: Number(this.split_batch_penalty),
        troy_oz_to_g: Number(this.troy_oz_to_g),
        metal_prices: {
          Cu: Number(this.metal_prices.Cu),
          Au: Number(this.metal_prices.Au),
          Ag: Number(this.metal_prices.Ag),
          Pd: Number(this.metal_prices.Pd),
        },
        metal_prices_jpy: {
          Cu: Number(this.metal_prices_jpy.Cu),
          Au: Number(this.metal_prices_jpy.Au),
          Ag: Number(this.metal_prices_jpy.Ag),
          Pd: Number(this.metal_prices_jpy.Pd),
        },
        a_mode: this.a_mode,
        a_thresholds: {
          Cu: Number(this.a_thresholds.Cu),
          Au: Number(this.a_thresholds.Au),
          Ag: Number(this.a_thresholds.Ag),
          Pd: Number(this.a_thresholds.Pd),
        },
        a_rc: {
          Cu: Number(this.a_rc.Cu),
          Au: Number(this.a_rc.Au),
          Ag: Number(this.a_rc.Ag),
          Pd: Number(this.a_rc.Pd),
        },
        a_freight_per_dmt: Number(this.a_freight_per_dmt),
        a_moisture_limit: Number(this.a_moisture_limit),
        a_moisture_penalty_per_wmt: Number(this.a_moisture_penalty_per_wmt),
        a_cu_below_threshold_tc_per_wmt: Number(this.a_cu_below_threshold_tc_per_wmt),
        a_assay_fee_if_below: Number(this.a_assay_fee_if_below),
        b_freight_per_kg: Number(this.b_freight_per_kg),
        b_deduction_per_ton: Number(this.b_deduction_per_ton),
        standard_tier_tables: tierTablesToDict(this.tier_tables),
      };
    }

    static fromDict(params, globalDefaults = {}, contextName = "CustomerConfig") {
      return buildCustomerConfig(params, globalDefaults, contextName);
    }
  }

  function validateRangePair(val, fieldName, context, allowZero = true) {
    if (!Array.isArray(val) || val.length !== 2) {
      throw new ConfigValidationError(
        `[${context}] '${fieldName}' 必須為長度 2 之數值陣列 [min, max]，實際為：${JSON.stringify(val)}`
      );
    }
    const lo = Number(val[0]);
    const hi = Number(val[1]);
    if (!Number.isFinite(lo) || !Number.isFinite(hi)) {
      throw new ConfigValidationError(`[${context}] '${fieldName}' 包含非數值：${JSON.stringify(val)}`);
    }
    if ((allowZero ? lo < 0.0 : lo <= 0.0) || lo > hi) {
      throw new ConfigValidationError(
        `[${context}] '${fieldName}' 區間數值不合法 (需滿足 0 ${allowZero ? "<=" : "<"} min <= max)，實際為：[${lo}, ${hi}]`
      );
    }
    return [lo, hi];
  }

  function parseTierTables(rawTiers, context = "tier_tables") {
    if (!rawTiers || typeof rawTiers !== "object") {
      throw new ConfigValidationError(`[${context}] 階梯折數表不可為空，必須包含 ${REQUIRED_METALS.join(", ")}`);
    }
    const parsed = {};
    for (const metal of REQUIRED_METALS) {
      if (!(metal in rawTiers)) {
        throw new ConfigValidationError(`[${context}] 階梯折數表缺少金屬 '${metal}' 之設定`);
      }
      const rules = rawTiers[metal];
      if (!Array.isArray(rules) || rules.length === 0) {
        throw new ConfigValidationError(`[${context}] 金屬 '${metal}' 之階梯折數規則必須為非空列表`);
      }
      const metalRules = [];
      rules.forEach((r, idx) => {
        let ruleObj;
        if (r instanceof TierRule) {
          ruleObj = r;
        } else if (r && typeof r === "object") {
          for (const reqK of ["min_grade", "max_grade", "payable_rate"]) {
            if (!(reqK in r) || r[reqK] === null || r[reqK] === "" || !Number.isFinite(Number(r[reqK]))) {
              throw new ConfigValidationError(`[${context}.${metal}[${idx}]] 缺少或無效欄位 '${reqK}'`);
            }
          }
          ruleObj = new TierRule(r.min_grade, r.max_grade, r.payable_rate);
        } else {
          throw new ConfigValidationError(`[${context}.${metal}[${idx}]] 格式錯誤`);
        }

        if (ruleObj.min_grade < 0 || ruleObj.min_grade >= ruleObj.max_grade) {
          throw new ConfigValidationError(
            `[${context}.${metal}[${idx}]] 品位區間不合法 (min_grade=${ruleObj.min_grade}, max_grade=${ruleObj.max_grade})`
          );
        }
        if (ruleObj.payable_rate < 0.0 || ruleObj.payable_rate > 1.0) {
          throw new ConfigValidationError(
            `[${context}.${metal}[${idx}]] 折數 payable_rate 必須介於 0.0 ~ 1.0，實際為：${ruleObj.payable_rate}`
          );
        }
        metalRules.push(ruleObj);
      });
      parsed[metal] = metalRules;
    }
    return parsed;
  }

  function buildCustomerConfig(params, globalDefaults = {}, contextName = "scenario") {
    const merged = Object.assign({}, globalDefaults || {}, params || {});

    const missingKeys = REQUIRED_GLOBAL_KEYS.filter((k) => !(k in merged));
    if (missingKeys.length > 0) {
      throw new ConfigValidationError(`[${contextName}] 缺少必要參數欄位：${missingKeys.join(", ")}`);
    }
    if (!("customer_type" in merged)) {
      throw new ConfigValidationError(`[${contextName}] 缺少 'customer_type' 欄位`);
    }

    const customerType = String(merged.customer_type).trim();
    if (!VALID_CUSTOMER_TYPES.includes(customerType)) {
      throw new ConfigValidationError(
        `[${contextName}] 'customer_type' 必須為 ${VALID_CUSTOMER_TYPES.join("/")} 之一，實際為：${customerType}`
      );
    }

    const priceCurrency = String(merged.price_currency).trim().toUpperCase();
    if (!VALID_CURRENCIES.includes(priceCurrency)) {
      throw new ConfigValidationError(
        `[${contextName}] 'price_currency' 必須為 ${VALID_CURRENCIES.join("/")} 之一，實際為：${priceCurrency}`
      );
    }

    const fillContainer = Boolean(merged.fill_container);
    const fullContainerRange = validateRangePair(
      merged.full_container_range,
      "full_container_range",
      contextName,
      false
    );
    const wMinKg = Number(merged.w_min_kg);
    const wMaxKg = Number(merged.w_max_kg);
    if (!Number.isFinite(wMinKg) || !Number.isFinite(wMaxKg) || wMinKg <= 0 || wMinKg > wMaxKg) {
      throw new ConfigValidationError(
        `[${contextName}] 重量上下限不合法：w_min_kg=${wMinKg}, w_max_kg=${wMaxKg}`
      );
    }

    const targetCuRange = validateRangePair(merged.target_cu_range, "target_cu_range", contextName, true);
    if (targetCuRange[1] > 100.0) {
      throw new ConfigValidationError(
        `[${contextName}] 'target_cu_range' 上限不可超過 100%，實際為：[${targetCuRange[0]}, ${targetCuRange[1]}]`
      );
    }

    const maxDiffCu = Number(merged.max_in_container_diff_cu);
    if (!Number.isFinite(maxDiffCu) || maxDiffCu < 0.0 || maxDiffCu > 100.0) {
      throw new ConfigValidationError(
        `[${contextName}] 'max_in_container_diff_cu' 必須介於 0 ~ 100，實際為：${maxDiffCu}`
      );
    }

    const maxValRaw = merged.max_total_value;
    const maxTotalValue =
      maxValRaw === null || maxValRaw === undefined || maxValRaw === "" || maxValRaw === Infinity
        ? Infinity
        : Number(maxValRaw);
    if (!Number.isFinite(maxTotalValue) && maxTotalValue !== Infinity) {
      throw new ConfigValidationError(`[${contextName}] 'max_total_value' 數值無效：${maxValRaw}`);
    }
    if (maxTotalValue <= 0) {
      throw new ConfigValidationError(
        `[${contextName}] 'max_total_value' 必須大於 0 或為 null，實際為：${maxValRaw}`
      );
    }

    const troyOzToG = Number(merged.troy_oz_to_g);
    if (!Number.isFinite(troyOzToG) || troyOzToG <= 0) {
      throw new ConfigValidationError(`[${contextName}] 'troy_oz_to_g' 必須大於 0，實際為：${troyOzToG}`);
    }

    for (const dictKey of ["metal_prices", "metal_prices_jpy", "a_thresholds", "a_rc"]) {
      const dVal = merged[dictKey];
      if (!dVal || typeof dVal !== "object") {
        throw new ConfigValidationError(`[${contextName}] '${dictKey}' 必須為字典物件`);
      }
      for (const m of REQUIRED_METALS) {
        if (!(m in dVal)) {
          throw new ConfigValidationError(`[${contextName}] '${dictKey}' 缺少金屬 '${m}' 設定`);
        }
        const numVal = Number(dVal[m]);
        if (!Number.isFinite(numVal) || numVal < 0.0) {
          throw new ConfigValidationError(`[${contextName}] '${dictKey}.${m}' 不可為負數或非數值：${dVal[m]}`);
        }
      }
    }

    const aMode = String(merged.a_mode).trim();
    if (!VALID_A_MODES.includes(aMode)) {
      throw new ConfigValidationError(
        `[${contextName}] 'a_mode' 必須為 ${VALID_A_MODES.join("/")} 之一，實際為：${aMode}`
      );
    }

    const rawTiers = merged.tier_tables || merged.standard_tier_tables;
    const parsedTiers = parseTierTables(rawTiers, `${contextName}.tier_tables`);

    const metalPrices = {};
    const metalPricesJpy = {};
    const aThresholds = {};
    const aRc = {};
    for (const m of REQUIRED_METALS) {
      metalPrices[m] = Number(merged.metal_prices[m]);
      metalPricesJpy[m] = Number(merged.metal_prices_jpy[m]);
      aThresholds[m] = Number(merged.a_thresholds[m]);
      aRc[m] = Number(merged.a_rc[m]);
    }

    return new CustomerConfig({
      customer_type: customerType,
      price_currency: priceCurrency,
      fill_container: fillContainer,
      full_container_range: fullContainerRange,
      w_min_kg: wMinKg,
      w_max_kg: wMaxKg,
      target_cu_range: targetCuRange,
      max_in_container_diff_cu: maxDiffCu,
      max_total_value: maxTotalValue,
      banned_codes: (merged.banned_codes || []).map((x) => String(x).trim()),
      must_batch_ids: (merged.must_batch_ids || []).map((x) => String(x).trim()),
      must_feed_codes: (merged.must_feed_codes || []).map((x) => String(x).trim()),
      must_bag_ids: (merged.must_bag_ids || []).map((x) => parseInt(x, 10)),
      exclude_pm_codes: (merged.exclude_pm_codes || []).map((x) => String(x).trim()),
      exclude_pm_materials: (merged.exclude_pm_materials || []).map((x) => String(x).trim()),
      metal_prices: metalPrices,
      metal_prices_jpy: metalPricesJpy,
      troy_oz_to_g: troyOzToG,
      a_mode: aMode,
      a_thresholds: aThresholds,
      a_rc: aRc,
      a_freight_per_dmt: Number(merged.a_freight_per_dmt),
      a_moisture_limit: Number(merged.a_moisture_limit),
      a_moisture_penalty_per_wmt: Number(merged.a_moisture_penalty_per_wmt),
      a_cu_below_threshold_tc_per_wmt: Number(merged.a_cu_below_threshold_tc_per_wmt),
      a_assay_fee_if_below: Number(merged.a_assay_fee_if_below),
      tier_tables: parsedTiers,
      b_freight_per_kg: Number(merged.b_freight_per_kg),
      b_deduction_per_ton: Number(merged.b_deduction_per_ton),
      split_batch_penalty: Number(merged.split_batch_penalty),
    });
  }

  function validateConfigData(configData) {
    if (!configData || typeof configData !== "object") {
      throw new ConfigValidationError("JSON 根節點必須為物件 (Object)");
    }
    for (const topKey of ["global_defaults", "scenarios"]) {
      if (!(topKey in configData)) {
        throw new ConfigValidationError(`JSON 設定檔缺少頂層必要欄位：'${topKey}'`);
      }
    }
    const globalDefs = configData.global_defaults;
    buildCustomerConfig({ customer_type: "A" }, globalDefs, "global_defaults");
    return true;
  }

  global.BlendingConfig = {
    REQUIRED_METALS,
    VALID_CUSTOMER_TYPES,
    VALID_A_MODES,
    VALID_CURRENCIES,
    REQUIRED_GLOBAL_KEYS,
    ConfigValidationError,
    TierRule,
    CustomerConfig,
    tierTablesToDict,
    parseTierTables,
    buildCustomerConfig,
    validateConfigData,
  };
})(typeof window !== "undefined" ? window : globalThis);

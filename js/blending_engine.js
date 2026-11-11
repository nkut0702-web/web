/**
 * 永源化工 — 自動配櫃系統核心運算與報表匯出引擎 (blending_engine.js)
 * 1:1 完整對應 Python 版 blending_engine.py：
 * 1. 庫存明細表二維陣列讀取、自動欄位對應與批包化驗值向上回填 (loadAndNormalizeInventory)
 * 2. 批次級別總覽表建構 (summarizeInventoryBatches)
 * 3. A ~ E 客戶完整計價評估核心，支援 USD 與 JPY 日圓雙幣別計價 (evaluateSelection)
 * 4. 支援百批以上庫存之 Branch-and-Bound + 二分搜尋配滿櫃最佳化求解器 (solveOptimalContainer)
 * 5. 配櫃結果批次小計表、逐包明細表與多工作表 Excel (.xlsx) 匯出工具 (exportBlendingResultToExcel)
 */

(function (global) {
  "use strict";

  const ASSAY_COLS = [
    "含水率(%)",
    "乾基Cu(%)",
    "濕基Cu(%)",
    "火-Au (mg/Kg)",
    "溼基Au(mg/Kg)",
    "Ag (mg/Kg)",
    "溼基Ag(mg/Kg)",
    "Pd (mg/Kg)",
    "溼基Pd(mg/Kg)",
  ];

  const STRING_COLS = ["廠別", "存放位置", "進料公司名稱(代號)", "總表料號", "編碼料號"];

  function toNumericCoerce(val) {
    if (val === null || val === undefined || typeof val === "boolean") return null;
    if (typeof val === "number") return Number.isFinite(val) ? val : null;
    const s = String(val).trim();
    if (s === "" || s.toLowerCase() === "nan" || s.toLowerCase() === "none") return null;
    const n = Number(s);
    return Number.isFinite(n) ? n : null;
  }

  function toAssayNumericCoerce(val) {
    if (val === null || val === undefined || typeof val === "boolean") return null;
    if (typeof val === "number") return Number.isFinite(val) ? val : null;
    let s = String(val).trim();
    if (s === "ND" || s === "nd") s = "0";
    if (s === "" || s.toLowerCase() === "nan" || s.toLowerCase() === "none") return null;
    const n = Number(s);
    return Number.isFinite(n) ? n : null;
  }

  function roundTo(val, decimals) {
    const factor = Math.pow(10, decimals);
    return Math.round((Number(val) + Number.EPSILON) * factor) / factor;
  }

  /**
   * 一、 資料讀取與批包回填前處理 (對應《庫存明細表.xlsx》)
   * @param {Array<Array<any>>} rawRows - 由 Excel 工作表轉換之二維陣列 (第 0 列為標題列)
   * @returns {Array<Object>} 正規化後之太空包物件陣列
   */
  function loadAndNormalizeInventory(rawRows) {
    if (!Array.isArray(rawRows) || rawRows.length < 2) {
      throw new Error("庫存明細表資料列數不足或格式無效");
    }

    const headerRow = rawRows[0] || [];
    const colCount = headerRow.length;
    const seenCounts = {};
    const dedupHeaders = [];

    for (let i = 0; i < colCount; i++) {
      const rawH = headerRow[i];
      let hStr = rawH === null || rawH === undefined ? "" : String(rawH).trim();
      if (hStr === "") {
        hStr = `Unnamed: ${i}`;
      }
      const cnt = seenCounts[hStr] || 0;
      seenCounts[hStr] = cnt + 1;
      dedupHeaders.push(cnt === 0 ? hStr : `${hStr}.${cnt}`);
    }

    // 建立欄位名稱至索引的映射
    const colIndexMap = {};
    dedupHeaders.forEach((h, idx) => {
      colIndexMap[h] = idx;
    });

    if (!("現場重" in colIndexMap) && "Unnamed: 6" in colIndexMap) {
      colIndexMap["現場重"] = colIndexMap["Unnamed: 6"];
    }
    if (!("倉庫重" in colIndexMap) && "Unnamed: 8" in colIndexMap) {
      colIndexMap["倉庫重"] = colIndexMap["Unnamed: 8"];
    }
    if ("溼基Au(mg/Kg).1" in colIndexMap) {
      colIndexMap["溼基Au(mg/Kg)"] = colIndexMap["溼基Au(mg/Kg).1"];
    }
    if ("溼基Ag(mg/Kg).1" in colIndexMap) {
      colIndexMap["溼基Ag(mg/Kg)"] = colIndexMap["溼基Ag(mg/Kg).1"];
    }

    // 過濾完全空白之尾部列
    const dataRows = [];
    for (let rIdx = 1; rIdx < rawRows.length; rIdx++) {
      const r = rawRows[rIdx];
      if (!Array.isArray(r)) continue;
      const hasAnyVal = r.some((cell) => cell !== null && cell !== undefined && String(cell).trim() !== "");
      if (hasAnyVal) {
        dataRows.push(r);
      }
    }

    const hasBagIdCol = "代號" in colIndexMap;
    const rows = dataRows.map((r, idx) => {
      const obj = { _idx: idx };

      if (hasBagIdCol) {
        const parsedId = toNumericCoerce(r[colIndexMap["代號"]]);
        obj["代號"] = parsedId !== null ? Math.trunc(parsedId) : idx + 1;
      } else {
        obj["代號"] = idx + 1;
      }

      for (const sCol of STRING_COLS) {
        if (sCol in colIndexMap) {
          const v = r[colIndexMap[sCol]];
          obj[sCol] = v === null || v === undefined ? "" : String(v).trim();
        } else {
          obj[sCol] = "";
        }
      }

      obj["包數"] = "包數" in colIndexMap ? toNumericCoerce(r[colIndexMap["包數"]]) : null;
      obj["小計"] = "小計" in colIndexMap ? toNumericCoerce(r[colIndexMap["小計"]]) : null;

      const wVal = "倉庫重" in colIndexMap ? toNumericCoerce(r[colIndexMap["倉庫重"]]) : null;
      obj["倉庫重"] = wVal !== null ? wVal : 0.0;

      const sVal = "現場重" in colIndexMap ? toNumericCoerce(r[colIndexMap["現場重"]]) : null;
      obj["現場重"] = sVal !== null ? sVal : obj["倉庫重"];

      for (const aCol of ASSAY_COLS) {
        if (aCol in colIndexMap) {
          obj[aCol] = toAssayNumericCoerce(r[colIndexMap[aCol]]);
        } else {
          obj[aCol] = 0.0;
        }
      }

      return obj;
    });

    // 由下往上標記 Batch_ID (Batch_1 ~ Batch_N)
    const isBatchEnd = rows.map((row) => row["包數"] !== null && row["包數"] > 0);
    const totalBatches = isBatchEnd.reduce((acc, flag) => acc + (flag ? 1 : 0), 0);
    let currentBatch = Math.max(1, totalBatches);
    const batchIdsReversed = [];

    for (let idx = rows.length - 1; idx >= 0; idx--) {
      batchIdsReversed.push(`Batch_${currentBatch}`);
      if (idx > 0 && isBatchEnd[idx - 1]) {
        currentBatch -= 1;
      }
    }
    batchIdsReversed.reverse();
    for (let idx = 0; idx < rows.length; idx++) {
      rows[idx]["Batch_ID"] = batchIdsReversed[idx];
    }

    // 同批號內向上回填 (bfill) 化驗值、包數與小計
    const fillCols = [...ASSAY_COLS, "包數", "小計"].filter(
      (c) => c in colIndexMap || ASSAY_COLS.includes(c)
    );

    const batchGroups = new Map();
    for (const row of rows) {
      const bId = row["Batch_ID"];
      if (!batchGroups.has(bId)) batchGroups.set(bId, []);
      batchGroups.get(bId).push(row);
    }

    for (const [, grpRows] of batchGroups.entries()) {
      for (const col of fillCols) {
        let lastVal = null;
        for (let i = grpRows.length - 1; i >= 0; i--) {
          if (grpRows[i][col] !== null && grpRows[i][col] !== undefined) {
            lastVal = grpRows[i][col];
          } else if (lastVal !== null) {
            grpRows[i][col] = lastVal;
          }
        }
      }
    }

    // 標記是否已化驗、回填 0.0 並計算單包乾重 (DMT, kg)
    for (const row of rows) {
      row["已化驗"] = row["乾基Cu(%)"] !== null && row["含水率(%)"] !== null;
      for (const aCol of ASSAY_COLS) {
        if (row[aCol] === null || row[aCol] === undefined) {
          row[aCol] = 0.0;
        }
      }
      row["乾重(kg)"] = row["倉庫重"] * (1.0 - row["含水率(%)"] / 100.0);
    }

    return rows;
  }

  /**
   * 二、 彙整各批次規格與重量總覽表
   */
  function summarizeInventoryBatches(invRows) {
    const batchGroups = new Map();
    for (const row of invRows) {
      const bId = row["Batch_ID"];
      if (!batchGroups.has(bId)) batchGroups.set(bId, []);
      batchGroups.get(bId).push(row);
    }

    const summaryRows = [];
    for (const [bId, grp] of batchGroups.entries()) {
      const r0 = grp[0];
      const wSum = grp.reduce((acc, r) => acc + r["倉庫重"], 0.0);
      const dSum = grp.reduce((acc, r) => acc + r["乾重(kg)"], 0.0);
      const wetCu = Number(r0["濕基Cu(%)"]);
      summaryRows.push({
        批號: bId,
        廠別: r0["廠別"],
        進料代號: r0["進料公司名稱(代號)"],
        總表料號: r0["總表料號"],
        編碼料號: r0["編碼料號"],
        存放位置: r0["存放位置"],
        總包數: grp.length,
        "倉庫濕重(kg)": roundTo(wSum, 1),
        "總乾重(kg)": roundTo(dSum, 2),
        "含水率(%)": roundTo(r0["含水率(%)"], 2),
        "乾基Cu(%)": roundTo(r0["乾基Cu(%)"], 2),
        "濕基Cu(%)": roundTo(wetCu, 2),
        "乾基Au(g/T)": roundTo(r0["火-Au (mg/Kg)"], 2),
        "溼基Au(g/T)": roundTo(r0["溼基Au(mg/Kg)"], 2),
        "乾基Ag(g/T)": roundTo(r0["Ag (mg/Kg)"], 2),
        "溼基Ag(g/T)": roundTo(r0["溼基Ag(mg/Kg)"], 2),
        "乾基Pd(g/T)": roundTo(r0["Pd (mg/Kg)"], 2),
        "溼基Pd(g/T)": roundTo(r0["溼基Pd(mg/Kg)"], 2),
        已化驗: Boolean(r0["已化驗"]),
        批次標籤: `${bId} | 進料:${r0["進料公司名稱(代號)"]} | 料號:${r0["總表料號"]} | 編碼:${r0["編碼料號"]} | ${grp.length}包 (${Math.round(wSum).toLocaleString("en-US")} kg) | 濕Cu:${wetCu.toFixed(2)}%`,
      });
    }
    return summaryRows;
  }

  /**
   * 三、 客戶計價評估核心引擎
   */
  function lookupTier(grade, tiers) {
    if (!Array.isArray(tiers)) return 0.0;
    for (let i = 0; i < tiers.length; i++) {
      const t = tiers[i];
      if (grade >= t.min_grade && grade < t.max_grade) {
        return t.payable_rate;
      }
    }
    return 0.0;
  }

  function evaluateSelection(subRows, fullRows, cfg) {
    if (!Array.isArray(subRows) || subRows.length === 0) {
      return { feasible: false, reason: "未選取任何物料" };
    }

    let wKg = 0.0;
    let dKg = 0.0;
    let minWetCu = Infinity;
    let maxWetCu = -Infinity;

    for (let i = 0; i < subRows.length; i++) {
      const r = subRows[i];
      wKg += r["倉庫重"];
      dKg += r["乾重(kg)"];
      const wCu = r["濕基Cu(%)"];
      if (wCu < minWetCu) minWetCu = wCu;
      if (wCu > maxWetCu) maxWetCu = wCu;
    }

    const wTon = wKg / 1000.0;
    const dTon = dKg / 1000.0;
    const avgMoist = wKg > 0 ? (1.0 - dKg / wKg) * 100.0 : 0.0;
    const cuDiff = maxWetCu - minWetCu;

    if (cuDiff > cfg.max_in_container_diff_cu + 1e-9) {
      return {
        feasible: false,
        reason: `同櫃銅含量落差 ${cuDiff.toFixed(4)}% 超過上限 ${cfg.max_in_container_diff_cu.toFixed(2)}%`,
        cu_diff: cuDiff,
      };
    }

    // 計算各批次原總包數
    const fullBatchCounts = new Map();
    for (let i = 0; i < fullRows.length; i++) {
      const bId = fullRows[i]["Batch_ID"];
      fullBatchCounts.set(bId, (fullBatchCounts.get(bId) || 0) + 1);
    }

    const subBatchGroups = new Map();
    for (let i = 0; i < subRows.length; i++) {
      const r = subRows[i];
      const bId = r["Batch_ID"];
      if (!subBatchGroups.has(bId)) subBatchGroups.set(bId, []);
      subBatchGroups.get(bId).push(r);
    }

    const splitBatches = [];
    const fullBatches = [];
    for (const [bId, grp] of subBatchGroups.entries()) {
      const totalBagsInBatch = fullBatchCounts.get(bId) || grp.length;
      const r0 = grp[0];
      const feedTag = `${bId}[${r0["進料公司名稱(代號)"]}]`;
      if (grp.length < totalBagsInBatch) {
        splitBatches.push(`${feedTag}(${grp.length}/${totalBagsInBatch}包)`);
      } else {
        fullBatches.push(`${feedTag}(全出${grp.length}包)`);
      }
    }

    // 分支 1：A 客戶計價 (乾基 DMT 門檻扣除法)
    if (cfg.customer_type === "A") {
      let sumDCu = 0.0,
        sumDAu = 0.0,
        sumDAg = 0.0,
        sumDPd = 0.0;
      let sumWCu = 0.0,
        sumWAu = 0.0,
        sumWAg = 0.0,
        sumWPd = 0.0;

      for (let i = 0; i < subRows.length; i++) {
        const r = subRows[i];
        const d = r["乾重(kg)"];
        const w = r["倉庫重"];
        sumDCu += d * r["乾基Cu(%)"];
        sumDAu += d * r["火-Au (mg/Kg)"];
        sumDAg += d * r["Ag (mg/Kg)"];
        sumDPd += d * r["Pd (mg/Kg)"];

        sumWCu += w * r["濕基Cu(%)"];
        sumWAu += w * r["溼基Au(mg/Kg)"];
        sumWAg += w * r["溼基Ag(mg/Kg)"];
        sumWPd += w * r["溼基Pd(mg/Kg)"];
      }

      const dryCu = dKg > 0 ? sumDCu / dKg : 0.0;
      const dryAu = dKg > 0 ? sumDAu / dKg : 0.0;
      const dryAg = dKg > 0 ? sumDAg / dKg : 0.0;
      const dryPd = dKg > 0 ? sumDPd / dKg : 0.0;

      const pmFactor = cfg.a_mode === "cu_only" ? 0.0 : 1.0;
      const contained = {
        Cu_kg: dKg * (dryCu / 100.0),
        Au_g: dTon * dryAu * pmFactor,
        Ag_g: dTon * dryAg * pmFactor,
        Pd_g: dTon * dryPd * pmFactor,
      };

      const recRates = {
        Cu: dryCu > 0 ? Math.max(0.0, (dryCu - cfg.a_thresholds.Cu) / dryCu) : 0.0,
        Au: (dryAu > 0 ? Math.max(0.0, (dryAu - cfg.a_thresholds.Au) / dryAu) : 0.0) * pmFactor,
        Ag: (dryAg > 0 ? Math.max(0.0, (dryAg - cfg.a_thresholds.Ag) / dryAg) : 0.0) * pmFactor,
        Pd: (dryPd > 0 ? Math.max(0.0, (dryPd - cfg.a_thresholds.Pd) / dryPd) : 0.0) * pmFactor,
      };

      const payable = {
        Cu_kg: contained.Cu_kg * recRates.Cu,
        Au_g: contained.Au_g * recRates.Au,
        Ag_g: contained.Ag_g * recRates.Ag,
        Pd_g: contained.Pd_g * recRates.Pd,
      };

      const netP = cfg.getCustomerANetPricesPerKgAndG();
      const revCu = payable.Cu_kg * netP.Cu;
      const revAu = payable.Au_g * netP.Au;
      const revAg = payable.Ag_g * netP.Ag;
      const revPd = payable.Pd_g * netP.Pd;
      const grossRev = revCu + revAu + revAg + revPd;

      const freightCost = dTon * cfg.a_freight_per_dmt;
      const moistPenalty = avgMoist > cfg.a_moisture_limit ? wTon * cfg.a_moisture_penalty_per_wmt : 0.0;
      const tcPenalty =
        dryCu < cfg.a_thresholds.Cu
          ? wTon * cfg.a_cu_below_threshold_tc_per_wmt + cfg.a_assay_fee_if_below
          : 0.0;
      const netProfit = grossRev - freightCost - moistPenalty - tcPenalty;

      return {
        feasible: true,
        customer_type: "A",
        price_currency: cfg.price_currency,
        fill_container: cfg.fill_container,
        bag_count: subRows.length,
        WMT_kg: wKg,
        PM_WMT_kg: wKg,
        DMT_kg: dKg,
        moisture_pct: avgMoist,
        cu_diff: cuDiff,
        dry_grades: { "Cu(%)": dryCu, "Au(g/T)": dryAu, "Ag(g/T)": dryAg, "Pd(g/T)": dryPd },
        wet_grades: {
          "Cu(%)": sumWCu / wKg,
          "Au(g/T)": sumWAu / wKg,
          "Ag(g/T)": sumWAg / wKg,
          "Pd(g/T)": sumWPd / wKg,
        },
        contained,
        recovery_rates: recRates,
        payable_quantities: payable,
        metal_revenues: { Cu: revCu, Au: revAu, Ag: revAg, Pd: revPd },
        gross_revenue: grossRev,
        freight_cost: freightCost,
        moisture_penalty: moistPenalty,
        tc_penalty: tcPenalty,
        deduction: freightCost + moistPenalty + tcPenalty,
        net_profit: netProfit,
        full_batches: fullBatches,
        split_batches: splitBatches,
        objective_score: netProfit - cfg.split_batch_penalty * splitBatches.length,
      };
    }

    // 分支 2：B、C、D、E 客戶計價 (濕基 WMT 階梯折數法)
    let sumWCu = 0.0;
    let sumDCu = 0.0,
      sumDAu = 0.0,
      sumDAg = 0.0,
      sumDPd = 0.0;
    for (let i = 0; i < subRows.length; i++) {
      const r = subRows[i];
      sumWCu += r["倉庫重"] * r["濕基Cu(%)"];
      const d = r["乾重(kg)"];
      sumDCu += d * r["乾基Cu(%)"];
      sumDAu += d * r["火-Au (mg/Kg)"];
      sumDAg += d * r["Ag (mg/Kg)"];
      sumDPd += d * r["Pd (mg/Kg)"];
    }

    const wetCu = sumWCu / wKg;
    if (!(wetCu >= cfg.target_cu_range[0] - 1e-9 && wetCu <= cfg.target_cu_range[1] + 1e-9)) {
      return {
        feasible: false,
        reason: `濕基銅 ${wetCu.toFixed(4)}% 未落入目標區間 [${cfg.target_cu_range[0]}, ${cfg.target_cu_range[1]}]`,
      };
    }

    const hasExclPm =
      cfg.customer_type === "E" &&
      ((cfg.exclude_pm_codes && cfg.exclude_pm_codes.length > 0) ||
        (cfg.exclude_pm_materials && cfg.exclude_pm_materials.length > 0));
    const exclCodeSet = new Set(cfg.exclude_pm_codes || []);
    const exclMatSet = new Set(cfg.exclude_pm_materials || []);

    let wPmKg = 0.0;
    let sumPmAu = 0.0,
      sumPmAg = 0.0,
      sumPmPd = 0.0;

    for (let i = 0; i < subRows.length; i++) {
      const r = subRows[i];
      if (hasExclPm && (exclCodeSet.has(r["編碼料號"]) || exclMatSet.has(r["總表料號"]))) {
        continue;
      }
      const w = r["倉庫重"];
      wPmKg += w;
      sumPmAu += w * r["溼基Au(mg/Kg)"];
      sumPmAg += w * r["溼基Ag(mg/Kg)"];
      sumPmPd += w * r["溼基Pd(mg/Kg)"];
    }

    const wPmTon = wPmKg / 1000.0;
    const wetAu = wPmKg > 0 ? sumPmAu / wPmKg : 0.0;
    const wetAg = wPmKg > 0 ? sumPmAg / wPmKg : 0.0;
    const wetPd = wPmKg > 0 ? sumPmPd / wPmKg : 0.0;

    const rates = {
      Cu: lookupTier(wetCu, cfg.tier_tables.Cu || []),
      Au: lookupTier(wetAu, cfg.tier_tables.Au || []),
      Ag: lookupTier(wetAg, cfg.tier_tables.Ag || []),
      Pd: lookupTier(wetPd, cfg.tier_tables.Pd || []),
    };

    const unitP = cfg.getTierCustomerPricesPerKgAndG();
    const revCu = wKg * (wetCu / 100.0) * rates.Cu * unitP.Cu;
    const revAu = wPmTon * wetAu * rates.Au * unitP.Au;
    const revAg = wPmTon * wetAg * rates.Ag * unitP.Ag;
    const revPd = wPmTon * wetPd * rates.Pd * unitP.Pd;
    const grossRev = revCu + revAu + revAg + revPd;

    const deduction = wKg * cfg.b_freight_per_kg + wTon * cfg.b_deduction_per_ton;
    const netProfit = grossRev - deduction;

    return {
      feasible: true,
      customer_type: cfg.customer_type,
      price_currency: cfg.price_currency,
      fill_container: cfg.fill_container,
      bag_count: subRows.length,
      WMT_kg: wKg,
      PM_WMT_kg: wPmKg,
      DMT_kg: dKg,
      moisture_pct: avgMoist,
      cu_diff: cuDiff,
      dry_grades: {
        "Cu(%)": dKg > 0 ? sumDCu / dKg : 0.0,
        "Au(g/T)": dKg > 0 ? sumDAu / dKg : 0.0,
        "Ag(g/T)": dKg > 0 ? sumDAg / dKg : 0.0,
        "Pd(g/T)": dKg > 0 ? sumDPd / dKg : 0.0,
      },
      wet_grades: { "Cu(%)": wetCu, "Au(g/T)": wetAu, "Ag(g/T)": wetAg, "Pd(g/T)": wetPd },
      tier_rates: rates,
      metal_revenues: { Cu: revCu, Au: revAu, Ag: revAg, Pd: revPd },
      gross_revenue: grossRev,
      deduction,
      net_profit: netProfit,
      full_batches: fullBatches,
      split_batches: splitBatches,
      objective_score: netProfit - cfg.split_batch_penalty * splitBatches.length,
    };
  }

  /**
   * 四、 批包雙層最佳化配櫃求解器 (Branch-and-Bound + 二分搜尋湊櫃)
   */
  function bisectLeft(arr, x) {
    let lo = 0,
      hi = arr.length;
    while (lo < hi) {
      const mid = (lo + hi) >>> 1;
      if (arr[mid] < x) lo = mid + 1;
      else hi = mid;
    }
    return lo;
  }

  function bisectRight(arr, x) {
    let lo = 0,
      hi = arr.length;
    while (lo < hi) {
      const mid = (lo + hi) >>> 1;
      if (arr[mid] <= x) lo = mid + 1;
      else hi = mid;
    }
    return lo;
  }

  function generateCombinations(arr, k, callback) {
    const n = arr.length;
    if (k <= 0 || k > n) return;
    const indices = Array.from({ length: k }, (_, i) => i);
    while (true) {
      callback(indices.map((i) => arr[i]));
      let i = k - 1;
      while (i >= 0 && indices[i] === i + n - k) {
        i--;
      }
      if (i < 0) break;
      indices[i]++;
      for (let j = i + 1; j < k; j++) {
        indices[j] = indices[j - 1] + 1;
      }
    }
  }

  function buildBatchSplitCandidates(bBags) {
    const nBags = bBags.length;
    if (nBags <= 1) return [[], []];

    const idxWeightPairs = bBags
      .map((r) => [r._idx, r["倉庫重"]])
      .sort((a, b) => a[1] - b[1]);

    const seenWeights = new Set();
    const uniqueSplits = [];

    function addComb(combPairs) {
      let sumW = 0.0;
      const idxs = new Array(combPairs.length);
      for (let i = 0; i < combPairs.length; i++) {
        idxs[i] = combPairs[i][0];
        sumW += combPairs[i][1];
      }
      const wSub = Math.round(sumW * 1e6) / 1e6;
      if (!seenWeights.has(wSub)) {
        seenWeights.add(wSub);
        uniqueSplits.push([wSub, idxs]);
      }
    }

    if (nBags <= 10) {
      for (let k = 1; k < nBags; k++) {
        generateCombinations(idxWeightPairs, k, addComb);
      }
    } else {
      for (let k = 1; k < nBags; k++) {
        for (let s = 0; s <= nBags - k; s++) {
          addComb(idxWeightPairs.slice(s, s + k));
        }
        const baseLight = idxWeightPairs.slice(0, k - 1);
        for (let j = k - 1; j < nBags; j++) {
          addComb(baseLight.concat([idxWeightPairs[j]]));
        }
        const baseHeavy = k > 1 ? idxWeightPairs.slice(nBags - (k - 1)) : [];
        for (let j = 0; j <= nBags - k; j++) {
          addComb([idxWeightPairs[j]].concat(baseHeavy));
        }
      }
    }

    uniqueSplits.sort((a, b) => a[0] - b[0]);
    const weights = uniqueSplits.map((x) => x[0]);
    return [uniqueSplits, weights];
  }

  function estimateBatchUnitValue(bp, cfg) {
    if (cfg.customer_type === "A") {
      const dryRatio = bp.dry_ratio;
      const { dry_cu: dCu, dry_au: dAu, dry_ag: dAg, dry_pd: dPd } = bp;
      const recCu = dCu > 0 ? Math.max(0.0, (dCu - cfg.a_thresholds.Cu) / dCu) : 0.0;
      const recAu = dAu > 0 ? Math.max(0.0, (dAu - cfg.a_thresholds.Au) / dAu) : 0.0;
      const recAg = dAg > 0 ? Math.max(0.0, (dAg - cfg.a_thresholds.Ag) / dAg) : 0.0;
      const recPd = dPd > 0 ? Math.max(0.0, (dPd - cfg.a_thresholds.Pd) / dPd) : 0.0;
      const netP = cfg.getCustomerANetPricesPerKgAndG();
      const revPerDmtKg =
        dCu * 0.01 * recCu * netP.Cu +
        0.001 * dAu * recAu * netP.Au +
        0.001 * dAg * recAg * netP.Ag +
        0.001 * dPd * recPd * netP.Pd -
        0.001 * cfg.a_freight_per_dmt;
      const moistPen =
        (1.0 - dryRatio) * 100.0 > cfg.a_moisture_limit
          ? 0.001 * cfg.a_moisture_penalty_per_wmt
          : 0.0;
      return dryRatio * revPerDmtKg - moistPen;
    } else {
      const { wet_cu: wCu, wet_au: wAu, wet_ag: wAg, wet_pd: wPd, is_pm: isPm } = bp;
      const rCu = lookupTier(wCu, cfg.tier_tables.Cu || []);
      const rAu = lookupTier(wAu, cfg.tier_tables.Au || []);
      const rAg = lookupTier(wAg, cfg.tier_tables.Ag || []);
      const rPd = lookupTier(wPd, cfg.tier_tables.Pd || []);
      const unitP = cfg.getTierCustomerPricesPerKgAndG();
      return (
        wCu * 0.01 * rCu * unitP.Cu +
        isPm * 0.001 * wAu * rAu * unitP.Au +
        isPm * 0.001 * wAg * rAg * unitP.Ag +
        isPm * 0.001 * wPd * rPd * unitP.Pd -
        cfg.b_freight_per_kg -
        0.001 * cfg.b_deduction_per_ton
      );
    }
  }

  function solveOptimalContainer(invRows, cfg) {
    const [wMin, wMax] = cfg.getActiveWeightBounds();

    const mustBatchSet = new Set(cfg.must_batch_ids || []);
    const mustFeedSet = new Set(cfg.must_feed_codes || []);
    const mustBagSet = new Set((cfg.must_bag_ids || []).map((x) => Number(x)));
    const bannedCodeSet = new Set(cfg.banned_codes || []);
    const exclPmCodeSet = new Set(cfg.exclude_pm_codes || []);
    const exclPmMatSet = new Set(cfg.exclude_pm_materials || []);

    const dfMust = [];
    const dfOpt = [];
    const dfAvail = [];

    for (let i = 0; i < invRows.length; i++) {
      const r = invRows[i];
      const isMust =
        mustBatchSet.has(r["Batch_ID"]) ||
        mustFeedSet.has(r["進料公司名稱(代號)"]) ||
        mustBagSet.has(Number(r["代號"]));
      const isAvail = isMust || (Boolean(r["已化驗"]) && !bannedCodeSet.has(r["編碼料號"]));
      if (isAvail) dfAvail.push(r);
      if (isMust) dfMust.push(r);
      else if (isAvail) dfOpt.push(r);
    }

    const wMust = dfMust.reduce((acc, r) => acc + r["倉庫重"], 0.0);
    const maxAvailWmtKg = dfAvail.reduce((acc, r) => acc + r["倉庫重"], 0.0);

    if (wMust > wMax + 1e-9) {
      return [
        [],
        {
          feasible: false,
          reason: `已勾選之必出料總濕重 (${Math.round(wMust).toLocaleString("en-US")} kg) 已超過配滿櫃上限 (${Math.round(wMax).toLocaleString("en-US")} kg)，請減少必出批次或調高重量上限。`,
          max_avail_wmt_kg: maxAvailWmtKg,
          must_wmt_kg: wMust,
        },
      ];
    }

    const totalBagsPerBatch = {};
    for (let i = 0; i < invRows.length; i++) {
      const bId = invRows[i]["Batch_ID"];
      totalBagsPerBatch[bId] = (totalBagsPerBatch[bId] || 0) + 1;
    }

    const optBatches = [];
    const batchDfMap = {};
    for (let i = 0; i < dfOpt.length; i++) {
      const r = dfOpt[i];
      const bId = r["Batch_ID"];
      if (!(bId in batchDfMap)) {
        optBatches.push(bId);
        batchDfMap[bId] = [];
      }
      batchDfMap[bId].push(r);
    }

    let bestScore = -1e18;
    let bestSelectedRows = [];
    let bestEval = {
      feasible: false,
      reason: `無符合配滿櫃重量區間 [${Math.round(wMin).toLocaleString("en-US")}, ${Math.round(wMax).toLocaleString("en-US")}] kg 及品位/同櫃落差條件之配櫃組合`,
      max_avail_wmt_kg: maxAvailWmtKg,
      must_wmt_kg: wMust,
    };

    const splitCandidatesByBatch = {};
    const splitWeightsByBatch = {};
    const batchProps = {};

    for (let i = 0; i < optBatches.length; i++) {
      const bId = optBatches[i];
      const bBags = batchDfMap[bId];
      const r0 = bBags[0];
      const isExcludedPm =
        cfg.customer_type === "E" &&
        (exclPmCodeSet.has(r0["編碼料號"]) || exclPmMatSet.has(r0["總表料號"]));
      const wB = bBags.reduce((acc, r) => acc + r["倉庫重"], 0.0);
      const dB = bBags.reduce((acc, r) => acc + r["乾重(kg)"], 0.0);
      const bp = {
        w_kg: wB,
        d_kg: dB,
        dry_ratio: 1.0 - Number(r0["含水率(%)"]) / 100.0,
        dry_cu: Number(r0["乾基Cu(%)"]),
        wet_cu: Number(r0["濕基Cu(%)"]),
        dry_au: Number(r0["火-Au (mg/Kg)"]),
        wet_au: Number(r0["溼基Au(mg/Kg)"]),
        dry_ag: Number(r0["Ag (mg/Kg)"]),
        wet_ag: Number(r0["溼基Ag(mg/Kg)"]),
        dry_pd: Number(r0["Pd (mg/Kg)"]),
        wet_pd: Number(r0["溼基Pd(mg/Kg)"]),
        is_pm: isExcludedPm ? 0.0 : 1.0,
      };
      bp.unit_val = estimateBatchUnitValue(bp, cfg);
      batchProps[bId] = bp;

      const [uSplits, uWeights] = buildBatchSplitCandidates(bBags);
      splitCandidatesByBatch[bId] = uSplits;
      splitWeightsByBatch[bId] = uWeights;
    }

    function evaluateBaseAndSplits(selectedFullBatches) {
      const selectedSet = new Set(selectedFullBatches);
      let dfBase = dfMust.slice();
      for (let i = 0; i < selectedFullBatches.length; i++) {
        const bBags = batchDfMap[selectedFullBatches[i]];
        for (let j = 0; j < bBags.length; j++) {
          dfBase.push(bBags[j]);
        }
      }

      let wBase = 0.0;
      let dBase = 0.0;
      let cuMinBase = Infinity;
      let cuMaxBase = -Infinity;

      for (let i = 0; i < dfBase.length; i++) {
        const r = dfBase[i];
        wBase += r["倉庫重"];
        dBase += r["乾重(kg)"];
        const wCu = r["濕基Cu(%)"];
        if (wCu < cuMinBase) cuMinBase = wCu;
        if (wCu > cuMaxBase) cuMaxBase = wCu;
      }

      if (wBase > wMax + 1e-9) return;

      if (wBase >= wMin - 1e-9 && wBase <= wMax + 1e-9) {
        const res = evaluateSelection(dfBase, invRows, cfg);
        if (res.feasible && res.gross_revenue <= cfg.max_total_value) {
          if (res.objective_score > bestScore) {
            bestScore = res.objective_score;
            bestSelectedRows = dfBase.slice().sort((a, b) => a["代號"] - b["代號"]);
            bestEval = res;
          }
        }
      }

      if (wBase < wMax - 1e-9) {
        const hasBase = dfBase.length > 0;
        if (hasBase && cuMaxBase - cuMinBase > cfg.max_in_container_diff_cu + 1e-9) {
          return;
        }

        let baseSplits = 0;
        if (hasBase) {
          const baseBatchCounts = {};
          for (let i = 0; i < dfBase.length; i++) {
            const bId = dfBase[i]["Batch_ID"];
            baseBatchCounts[bId] = (baseBatchCounts[bId] || 0) + 1;
          }
          for (const bId of Object.keys(baseBatchCounts)) {
            if (baseBatchCounts[bId] < totalBagsPerBatch[bId]) {
              baseSplits += 1;
            }
          }
        }

        let dCuBase = 0.0,
          dAuBase = 0.0,
          dAgBase = 0.0,
          dPdBase = 0.0;
        let wCuBase = 0.0,
          wPmBase = 0.0,
          wAuBase = 0.0,
          wAgBase = 0.0,
          wPdBase = 0.0;

        if (cfg.customer_type === "A") {
          for (let i = 0; i < dfBase.length; i++) {
            const r = dfBase[i];
            const d = r["乾重(kg)"];
            dCuBase += d * r["乾基Cu(%)"];
            dAuBase += d * r["火-Au (mg/Kg)"];
            dAgBase += d * r["Ag (mg/Kg)"];
            dPdBase += d * r["Pd (mg/Kg)"];
          }
        } else {
          const hasExcl =
            cfg.customer_type === "E" && (exclPmCodeSet.size > 0 || exclPmMatSet.size > 0);
          for (let i = 0; i < dfBase.length; i++) {
            const r = dfBase[i];
            const w = r["倉庫重"];
            wCuBase += w * r["濕基Cu(%)"];
            if (!(hasExcl && (exclPmCodeSet.has(r["編碼料號"]) || exclPmMatSet.has(r["總表料號"])))) {
              wPmBase += w;
              wAuBase += w * r["溼基Au(mg/Kg)"];
              wAgBase += w * r["溼基Ag(mg/Kg)"];
              wPdBase += w * r["溼基Pd(mg/Kg)"];
            }
          }
        }

        const needMin = wMin - wBase;
        const needMax = wMax - wBase;

        for (let sIdx = 0; sIdx < optBatches.length; sIdx++) {
          const splitB = optBatches[sIdx];
          if (selectedSet.has(splitB)) continue;
          const cands = splitCandidatesByBatch[splitB];
          if (!cands || cands.length === 0) continue;
          const bp = batchProps[splitB];
          const wetCuB = bp.wet_cu;

          const cuDiffCand = hasBase
            ? Math.max(cuMaxBase, wetCuB) - Math.min(cuMinBase, wetCuB)
            : 0.0;
          if (cuDiffCand > cfg.max_in_container_diff_cu + 1e-9) continue;

          const wList = splitWeightsByBatch[splitB];
          const lo = bisectLeft(wList, needMin - 1e-9);
          const hi = bisectRight(wList, needMax + 1e-9);
          if (lo >= hi) continue;

          const nSplits = baseSplits + 1;
          const penaltyVal = cfg.split_batch_penalty * nSplits;
          let bestLocalIndices = null;

          if (cfg.customer_type === "A") {
            const thCu = cfg.a_thresholds.Cu,
              thAu = cfg.a_thresholds.Au,
              thAg = cfg.a_thresholds.Ag,
              thPd = cfg.a_thresholds.Pd;
            const netP = cfg.getCustomerANetPricesPerKgAndG();
            const netPCu = netP.Cu,
              netPAu = netP.Au,
              netPAg = netP.Ag,
              netPPd = netP.Pd;
            const dryRatio = bp.dry_ratio;
            const bDCu = bp.dry_cu,
              bDAu = bp.dry_au,
              bDAg = bp.dry_ag,
              bDPd = bp.dry_pd;

            for (let idxPos = lo; idxPos < hi; idxPos++) {
              const [wSub, bagIndices] = cands[idxPos];
              const wKg = wBase + wSub;
              const dSub = wSub * dryRatio;
              const dKg = dBase + dSub;
              const wTon = wKg * 0.001;
              const dTon = dKg * 0.001;
              const avgMoist = (1.0 - dKg / wKg) * 100.0;

              const dryCu = (dCuBase + dSub * bDCu) / dKg;
              const dryAu = (dAuBase + dSub * bDAu) / dKg;
              const dryAg = (dAgBase + dSub * bDAg) / dKg;
              const dryPd = (dPdBase + dSub * bDPd) / dKg;

              const recCu = dryCu > 0 ? Math.max(0.0, (dryCu - thCu) / dryCu) : 0.0;
              const recAu = dryAu > 0 ? Math.max(0.0, (dryAu - thAu) / dryAu) : 0.0;
              const recAg = dryAg > 0 ? Math.max(0.0, (dryAg - thAg) / dryAg) : 0.0;
              const recPd = dryPd > 0 ? Math.max(0.0, (dryPd - thPd) / dryPd) : 0.0;

              const grossRev =
                dKg * (dryCu * 0.01) * recCu * netPCu +
                dTon * dryAu * recAu * netPAu +
                dTon * dryAg * recAg * netPAg +
                dTon * dryPd * recPd * netPPd;
              if (grossRev > cfg.max_total_value) continue;

              const freightCost = dTon * cfg.a_freight_per_dmt;
              const moistPenalty =
                avgMoist > cfg.a_moisture_limit ? wTon * cfg.a_moisture_penalty_per_wmt : 0.0;
              const tcPenalty =
                dryCu < thCu
                  ? wTon * cfg.a_cu_below_threshold_tc_per_wmt + cfg.a_assay_fee_if_below
                  : 0.0;
              const objScore = grossRev - freightCost - moistPenalty - tcPenalty - penaltyVal;
              if (objScore > bestScore) {
                bestScore = objScore;
                bestLocalIndices = bagIndices;
              }
            }
          } else {
            const cuLo = cfg.target_cu_range[0] - 1e-9;
            const cuHi = cfg.target_cu_range[1] + 1e-9;
            const tiersCu = cfg.tier_tables.Cu || [];
            const tiersAu = cfg.tier_tables.Au || [];
            const tiersAg = cfg.tier_tables.Ag || [];
            const tiersPd = cfg.tier_tables.Pd || [];
            const unitP = cfg.getTierCustomerPricesPerKgAndG();
            const pCu = unitP.Cu,
              pAu = unitP.Au,
              pAg = unitP.Ag,
              pPd = unitP.Pd;
            const isPmB = bp.is_pm;
            const bWAu = bp.wet_au,
              bWAg = bp.wet_ag,
              bWPd = bp.wet_pd;

            for (let idxPos = lo; idxPos < hi; idxPos++) {
              const [wSub, bagIndices] = cands[idxPos];
              const wKg = wBase + wSub;
              const wetCu = (wCuBase + wSub * wetCuB) / wKg;
              if (!(wetCu >= cuLo && wetCu <= cuHi)) continue;

              const wPmKg = wPmBase + wSub * isPmB;
              const wPmTon = wPmKg * 0.001;
              let wetAu = 0.0,
                wetAg = 0.0,
                wetPd = 0.0;
              if (wPmKg > 0) {
                wetAu = (wAuBase + wSub * isPmB * bWAu) / wPmKg;
                wetAg = (wAgBase + wSub * isPmB * bWAg) / wPmKg;
                wetPd = (wPdBase + wSub * isPmB * bWPd) / wPmKg;
              }

              const rCu = lookupTier(wetCu, tiersCu);
              const rAu = lookupTier(wetAu, tiersAu);
              const rAg = lookupTier(wetAg, tiersAg);
              const rPd = lookupTier(wetPd, tiersPd);

              const grossRev =
                wKg * (wetCu * 0.01) * rCu * pCu +
                wPmTon * wetAu * rAu * pAu +
                wPmTon * wetAg * rAg * pAg +
                wPmTon * wetPd * rPd * pPd;
              if (grossRev > cfg.max_total_value) continue;

              const deduction =
                wKg * cfg.b_freight_per_kg + wKg * 0.001 * cfg.b_deduction_per_ton;
              const objScore = grossRev - deduction - penaltyVal;
              if (objScore > bestScore) {
                bestScore = objScore;
                bestLocalIndices = bagIndices;
              }
            }
          }

          if (bestLocalIndices !== null) {
            const idxSet = new Set(bestLocalIndices);
            const bBags = batchDfMap[splitB];
            const splitChosen = bBags.filter((r) => idxSet.has(r._idx));
            const dfCand = dfBase.concat(splitChosen).sort((a, b) => a["代號"] - b["代號"]);
            bestSelectedRows = dfCand;
            bestEval = evaluateSelection(bestSelectedRows, invRows, cfg);
          }
        }
      }
    }

    const hasMust = dfMust.length > 0;
    let cuMinMust = Infinity;
    let cuMaxMust = -Infinity;
    if (hasMust) {
      for (let i = 0; i < dfMust.length; i++) {
        const wCu = dfMust[i]["濕基Cu(%)"];
        if (wCu < cuMinMust) cuMinMust = wCu;
        if (wCu > cuMaxMust) cuMaxMust = wCu;
      }
    }

    evaluateBaseAndSplits([]);

    const eligibleFullBatches = [];
    for (let i = 0; i < optBatches.length; i++) {
      const bId = optBatches[i];
      const bp = batchProps[bId];
      if (wMust + bp.w_kg > wMax + 1e-9) continue;
      const cMin = Math.min(cuMinMust, bp.wet_cu);
      const cMax = Math.max(cuMaxMust, bp.wet_cu);
      if (cMax - cMin <= cfg.max_in_container_diff_cu + 1e-9) {
        eligibleFullBatches.push(bId);
      }
    }

    eligibleFullBatches.sort((a, b) => batchProps[b].unit_val - batchProps[a].unit_val);

    for (let i = 0; i < eligibleFullBatches.length; i++) {
      evaluateBaseAndSplits([eligibleFullBatches[i]]);
    }

    const topPool = eligibleFullBatches.slice(0, 25);
    const nPool = topPool.length;
    for (let i = 0; i < nPool; i++) {
      const b1 = topPool[i];
      const bp1 = batchProps[b1];
      const w1 = wMust + bp1.w_kg;
      const cMin1 = Math.min(cuMinMust, bp1.wet_cu);
      const cMax1 = Math.max(cuMaxMust, bp1.wet_cu);

      for (let j = i + 1; j < nPool; j++) {
        const b2 = topPool[j];
        const bp2 = batchProps[b2];
        const w2 = w1 + bp2.w_kg;
        if (w2 > wMax + 1e-9) continue;
        const cMin2 = Math.min(cMin1, bp2.wet_cu);
        const cMax2 = Math.max(cMax1, bp2.wet_cu);
        if (cMax2 - cMin2 > cfg.max_in_container_diff_cu + 1e-9) continue;

        evaluateBaseAndSplits([b1, b2]);

        if (!hasMust && w2 < wMax - 1500.0) {
          const maxK = Math.min(nPool, 18);
          for (let k = j + 1; k < maxK; k++) {
            const b3 = topPool[k];
            const bp3 = batchProps[b3];
            const w3 = w2 + bp3.w_kg;
            if (w3 > wMax + 1e-9) continue;
            const cMin3 = Math.min(cMin2, bp3.wet_cu);
            const cMax3 = Math.max(cMax2, bp3.wet_cu);
            if (cMax3 - cMin3 > cfg.max_in_container_diff_cu + 1e-9) continue;
            evaluateBaseAndSplits([b1, b2, b3]);
          }
        }
      }
    }

    return [bestSelectedRows, bestEval];
  }

  /**
   * 五、 結果明細表格建構與多工作表 Excel (.xlsx) 匯出工具
   */
  function buildSelectedBatchSummary(selectedRows, fullRows) {
    if (!Array.isArray(selectedRows) || selectedRows.length === 0) return [];

    const fullBatchCounts = new Map();
    for (let i = 0; i < fullRows.length; i++) {
      const bId = fullRows[i]["Batch_ID"];
      fullBatchCounts.set(bId, (fullBatchCounts.get(bId) || 0) + 1);
    }

    const subBatchGroups = new Map();
    for (let i = 0; i < selectedRows.length; i++) {
      const r = selectedRows[i];
      const bId = r["Batch_ID"];
      if (!subBatchGroups.has(bId)) subBatchGroups.set(bId, []);
      subBatchGroups.get(bId).push(r);
    }

    const rows = [];
    for (const [bId, grp] of subBatchGroups.entries()) {
      const totalBBags = fullBatchCounts.get(bId) || grp.length;
      const r0 = grp[0];
      const wSub = grp.reduce((acc, r) => acc + r["倉庫重"], 0.0);
      const dSub = grp.reduce((acc, r) => acc + r["乾重(kg)"], 0.0);
      const siteSub = grp.reduce((acc, r) => acc + r["現場重"], 0.0);
      const isFull = grp.length === totalBBags;
      rows.push({
        批號: bId,
        廠別: r0["廠別"],
        進料代號: r0["進料公司名稱(代號)"],
        總表料號: r0["總表料號"],
        編碼料號: r0["編碼料號"],
        存放位置: r0["存放位置"],
        出貨方式: isFull ? "整批全出" : "拆批湊櫃",
        "選中包數/原批包數": `${grp.length}/${totalBBags} 包`,
        選中包數: grp.length,
        "含水率(%)": roundTo(r0["含水率(%)"], 2),
        "乾基Cu(%)": roundTo(r0["乾基Cu(%)"], 2),
        "濕基Cu(%)": roundTo(r0["濕基Cu(%)"], 2),
        "乾基Au(g/T)": roundTo(r0["火-Au (mg/Kg)"], 2),
        "溼基Au(g/T)": roundTo(r0["溼基Au(mg/Kg)"], 2),
        "乾基Ag(g/T)": roundTo(r0["Ag (mg/Kg)"], 2),
        "溼基Ag(g/T)": roundTo(r0["溼基Ag(mg/Kg)"], 2),
        "乾基Pd(g/T)": roundTo(r0["Pd (mg/Kg)"], 2),
        "溼基Pd(g/T)": roundTo(r0["溼基Pd(mg/Kg)"], 2),
        "現場重小計(kg)": roundTo(siteSub, 1),
        "倉庫濕重小計(kg)": roundTo(wSub, 1),
        "乾重小計(kg)": roundTo(dSub, 2),
      });
    }
    return rows;
  }

  function buildSelectedBagDetail(selectedRows) {
    if (!Array.isArray(selectedRows) || selectedRows.length === 0) return [];
    const sorted = selectedRows.slice().sort((a, b) => a["代號"] - b["代號"]);
    return sorted.map((row) => ({
      單包代號: Number(row["代號"]),
      批號: row["Batch_ID"],
      廠別: row["廠別"],
      進料代號: row["進料公司名稱(代號)"],
      總表料號: row["總表料號"],
      編碼料號: row["編碼料號"],
      存放位置: row["存放位置"],
      "現場重(kg)": roundTo(row["現場重"], 1),
      "倉庫濕重(kg)": roundTo(row["倉庫重"], 1),
      "單包乾重(kg)": roundTo(row["乾重(kg)"], 2),
      "含水率(%)": roundTo(row["含水率(%)"], 2),
      "乾基Cu(%)": roundTo(row["乾基Cu(%)"], 4),
      "濕基Cu(%)": roundTo(row["濕基Cu(%)"], 4),
      "乾基Au(g/T)": roundTo(row["火-Au (mg/Kg)"], 4),
      "溼基Au(g/T)": roundTo(row["溼基Au(mg/Kg)"], 4),
      "乾基Ag(g/T)": roundTo(row["Ag (mg/Kg)"], 4),
      "溼基Ag(g/T)": roundTo(row["溼基Ag(mg/Kg)"], 4),
      "乾基Pd(g/T)": roundTo(row["Pd (mg/Kg)"], 4),
      "溼基Pd(g/T)": roundTo(row["溼基Pd(mg/Kg)"], 4),
    }));
  }

  function exportBlendingResultToExcel(selectedRows, fullRows, evalRes, cfg) {
    if (typeof XLSX === "undefined") {
      throw new Error("SheetJS (XLSX) 尚未載入，無法匯出 Excel 檔案");
    }

    const batchSum = buildSelectedBatchSummary(selectedRows, fullRows);
    const bagDet = buildSelectedBagDetail(selectedRows);
    const curr = cfg.price_currency;

    const summaryRows = [
      { 項目: "客戶類型", 數值: `${cfg.customer_type} 客戶` },
      { 項目: "計價幣別", 數值: `${curr} (${curr === "JPY" ? "日圓計價 ¥" : "美元計價 $"})` },
      {
        項目: "配櫃模式",
        數值: `配滿櫃模式 (${Math.round(cfg.full_container_range[0]).toLocaleString("en-US")} ~ ${Math.round(cfg.full_container_range[1]).toLocaleString("en-US")} kg)`,
      },
      { 項目: "選中總包數", 數值: `${evalRes.bag_count || 0} 包` },
      { 項目: "整批全出批次", 數值: (evalRes.full_batches || []).join(", ") || "無" },
      { 項目: "拆批湊櫃批次", 數值: (evalRes.split_batches || []).join(", ") || "無 (0拆批)" },
      {
        項目: "貨櫃總倉庫濕重 (WMT)",
        數值: `${(evalRes.WMT_kg || 0).toLocaleString("en-US", { minimumFractionDigits: 2, maximumFractionDigits: 2 })} kg (${((evalRes.WMT_kg || 0) / 1000.0).toFixed(3)} T)`,
      },
      {
        項目: "貴金屬有效計價濕重 (PM_WMT)",
        數值: `${(evalRes.PM_WMT_kg || 0).toLocaleString("en-US", { minimumFractionDigits: 2, maximumFractionDigits: 2 })} kg`,
      },
      {
        項目: "貨櫃總乾重 (DMT)",
        數值: `${(evalRes.DMT_kg || 0).toLocaleString("en-US", { minimumFractionDigits: 2, maximumFractionDigits: 2 })} kg (${((evalRes.DMT_kg || 0) / 1000.0).toFixed(4)} T)`,
      },
      { 項目: "加權平均含水率 (%)", 數值: `${(evalRes.moisture_pct || 0).toFixed(2)}%` },
      { 項目: "同櫃最大濕基銅落差 (%)", 數值: `${(evalRes.cu_diff || 0).toFixed(4)}%` },
      {
        項目: "加權乾基 Cu (%)",
        數值: `${((evalRes.dry_grades && evalRes.dry_grades["Cu(%)"]) || 0).toFixed(4)}%`,
      },
      {
        項目: "加權濕基 Cu (%)",
        數值: `${((evalRes.wet_grades && evalRes.wet_grades["Cu(%)"]) || 0).toFixed(4)}%`,
      },
      {
        項目: "加權濕基 Au (g/T)",
        數值: `${((evalRes.wet_grades && evalRes.wet_grades["Au(g/T)"]) || 0).toFixed(4)}`,
      },
      {
        項目: "加權濕基 Ag (g/T)",
        數值: `${((evalRes.wet_grades && evalRes.wet_grades["Ag(g/T)"]) || 0).toFixed(4)}`,
      },
      {
        項目: "加權濕基 Pd (g/T)",
        數值: `${((evalRes.wet_grades && evalRes.wet_grades["Pd(g/T)"]) || 0).toFixed(4)}`,
      },
      { 項目: `毛收入 (Gross Revenue, ${curr})`, 數值: roundTo(evalRes.gross_revenue || 0, 2) },
      { 項目: `總扣減與運費罰款 (Deductions, ${curr})`, 數值: roundTo(evalRes.deduction || 0, 2) },
      { 項目: `預估結算淨利 (Net Profit, ${curr})`, 數值: roundTo(evalRes.net_profit || 0, 2) },
      { 項目: "綜合目標分數 (含拆批懲罰)", 數值: roundTo(evalRes.objective_score || 0, 2) },
    ];

    const cfgDict = cfg.toDict();
    const paramRows = [];
    for (const [k, v] of Object.entries(cfgDict)) {
      if (k !== "standard_tier_tables") {
        paramRows.push({
          參數名稱: k,
          設定值: typeof v === "object" && v !== null ? JSON.stringify(v) : String(v),
        });
      }
    }

    const wb = XLSX.utils.book_new();
    const ws1 = XLSX.utils.json_to_sheet(summaryRows);
    const ws2 = XLSX.utils.json_to_sheet(batchSum);
    const ws3 = XLSX.utils.json_to_sheet(bagDet);
    const ws4 = XLSX.utils.json_to_sheet(paramRows);

    XLSX.utils.book_append_sheet(wb, ws1, "1_配櫃計價結果摘要");
    XLSX.utils.book_append_sheet(wb, ws2, "2_選中批次小計表");
    XLSX.utils.book_append_sheet(wb, ws3, "3_選中太空包逐包明細");
    XLSX.utils.book_append_sheet(wb, ws4, "4_本次配櫃參數設定");

    const fileName = `永源化工_配櫃結果_${cfg.customer_type}客戶_${curr}_${Math.round(evalRes.WMT_kg || 0)}kg.xlsx`;
    XLSX.writeFile(wb, fileName);
  }

  global.BlendingEngine = {
    ASSAY_COLS,
    STRING_COLS,
    loadAndNormalizeInventory,
    summarizeInventoryBatches,
    lookupTier,
    evaluateSelection,
    solveOptimalContainer,
    buildSelectedBatchSummary,
    buildSelectedBagDetail,
    exportBlendingResultToExcel,
  };
})(typeof window !== "undefined" ? window : globalThis);

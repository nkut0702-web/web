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
    "Ni (mg/Kg)",
    "溼基Ni(mg/Kg)",
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
    if ("溼基Pd(mg/Kg).1" in colIndexMap) {
      colIndexMap["溼基Pd(mg/Kg)"] = colIndexMap["溼基Pd(mg/Kg).1"];
    }
    if ("溼基Ni(mg/Kg).1" in colIndexMap) {
      colIndexMap["溼基Ni(mg/Kg)"] = colIndexMap["溼基Ni(mg/Kg).1"];
    }

    const dryNiAliases = ["Ni (mg/Kg)", "Ni(mg/Kg)", "乾基Ni(mg/Kg)", "乾基Ni(g/T)", "Ni", "ICP NI"];
    const wetNiAliases = ["溼基Ni(mg/Kg)", "濕基Ni(mg/Kg)", "溼基Ni(g/T)", "濕基Ni(g/T)"];
    const dryNiPctAliases = ["乾基NI(%)", "乾基Ni(%)"];
    const wetNiPctAliases = ["濕基NI(%)", "溼基NI(%)", "濕基Ni(%)", "溼基Ni(%)"];
    let dryNiIdx = null;
    let wetNiIdx = null;
    let dryNiPctIdx = null;
    let wetNiPctIdx = null;
    for (const alias of dryNiAliases) {
      if (alias in colIndexMap) {
        dryNiIdx = colIndexMap[alias];
        break;
      }
    }
    for (const alias of wetNiAliases) {
      if (alias in colIndexMap) {
        wetNiIdx = colIndexMap[alias];
        break;
      }
    }
    for (const alias of dryNiPctAliases) {
      if (alias in colIndexMap) {
        dryNiPctIdx = colIndexMap[alias];
        break;
      }
    }
    for (const alias of wetNiPctAliases) {
      if (alias in colIndexMap) {
        wetNiPctIdx = colIndexMap[alias];
        break;
      }
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
        if (aCol === "Ni (mg/Kg)") {
          let val = dryNiPctIdx !== null ? toAssayNumericCoerce(r[dryNiPctIdx]) : null;
          if (val !== null) {
            val = val * 10000.0;
          } else {
            val = dryNiIdx !== null ? toAssayNumericCoerce(r[dryNiIdx]) : null;
            if (val === null && "ICP NI" in colIndexMap) {
              val = toAssayNumericCoerce(r[colIndexMap["ICP NI"]]);
            }
          }
          obj[aCol] = val;
        } else if (aCol === "溼基Ni(mg/Kg)") {
          let val = wetNiPctIdx !== null ? toAssayNumericCoerce(r[wetNiPctIdx]) : null;
          if (val !== null) {
            val = val * 10000.0;
          } else {
            val = wetNiIdx !== null ? toAssayNumericCoerce(r[wetNiIdx]) : null;
          }
          obj[aCol] = val;
        } else if (aCol in colIndexMap) {
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

    // 標記是否已化驗、回填 0.0 並計算單包乾重 (DMT, kg) 與乾濕基 Ni 互算
    for (const row of rows) {
      row["已化驗"] = row["乾基Cu(%)"] !== null && row["含水率(%)"] !== null;
      const moistPct = row["含水率(%)"] !== null && row["含水率(%)"] !== undefined ? row["含水率(%)"] : 0.0;
      const dryFactor = Math.max(0.0, 1.0 - moistPct / 100.0);
      if (row["Ni (mg/Kg)"] !== null && (row["溼基Ni(mg/Kg)"] === null || row["溼基Ni(mg/Kg)"] === undefined)) {
        row["溼基Ni(mg/Kg)"] = row["Ni (mg/Kg)"] * dryFactor;
      } else if (row["溼基Ni(mg/Kg)"] !== null && (row["Ni (mg/Kg)"] === null || row["Ni (mg/Kg)"] === undefined)) {
        row["Ni (mg/Kg)"] = dryFactor > 0 ? row["溼基Ni(mg/Kg)"] / dryFactor : 0.0;
      }
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
        "乾基Ni(g/T)": roundTo(r0["Ni (mg/Kg)"], 2),
        "溼基Ni(g/T)": roundTo(r0["溼基Ni(mg/Kg)"], 2),
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

  function evaluateSelection(subRows, fullRows, cfg, origTotalBagsMap = null) {
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
    if (origTotalBagsMap) {
      for (const [bId, cnt] of Object.entries(origTotalBagsMap)) {
        fullBatchCounts.set(bId, cnt);
      }
    } else {
      for (let i = 0; i < fullRows.length; i++) {
        const bId = fullRows[i]["Batch_ID"];
        fullBatchCounts.set(bId, (fullBatchCounts.get(bId) || 0) + 1);
      }
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
        sumDPd = 0.0,
        sumDNi = 0.0;
      let sumWCu = 0.0,
        sumWAu = 0.0,
        sumWAg = 0.0,
        sumWPd = 0.0,
        sumWNi = 0.0;

      for (let i = 0; i < subRows.length; i++) {
        const r = subRows[i];
        const d = r["乾重(kg)"];
        const w = r["倉庫重"];
        sumDCu += d * r["乾基Cu(%)"];
        sumDAu += d * r["火-Au (mg/Kg)"];
        sumDAg += d * r["Ag (mg/Kg)"];
        sumDPd += d * r["Pd (mg/Kg)"];
        sumDNi += d * r["Ni (mg/Kg)"];

        sumWCu += w * r["濕基Cu(%)"];
        sumWAu += w * r["溼基Au(mg/Kg)"];
        sumWAg += w * r["溼基Ag(mg/Kg)"];
        sumWPd += w * r["溼基Pd(mg/Kg)"];
        sumWNi += w * r["溼基Ni(mg/Kg)"];
      }

      const dryCu = dKg > 0 ? sumDCu / dKg : 0.0;
      const dryAu = dKg > 0 ? sumDAu / dKg : 0.0;
      const dryAg = dKg > 0 ? sumDAg / dKg : 0.0;
      const dryPd = dKg > 0 ? sumDPd / dKg : 0.0;
      const dryNi = dKg > 0 ? sumDNi / dKg : 0.0;

      const isCuOnly = cfg.a_mode === "cu_only";
      const pmFactor = isCuOnly ? 0.0 : 1.0;
      const contained = {
        Cu_kg: dKg * (dryCu / 100.0),
        Au_g: dTon * dryAu * pmFactor,
        Ag_g: dTon * dryAg * pmFactor,
        Pd_g: dTon * dryPd * pmFactor,
        Ni_g: dTon * dryNi * pmFactor,
      };

      const cuRecRate = isCuOnly
        ? lookupTier(dryCu, cfg.a_cu_only_tier_table || [])
        : dryCu > 0
        ? Math.max(0.0, (dryCu - cfg.a_thresholds.Cu) / dryCu)
        : 0.0;

      const recRates = {
        Cu: cuRecRate,
        Au: (dryAu > 0 ? Math.max(0.0, (dryAu - cfg.a_thresholds.Au) / dryAu) : 0.0) * pmFactor,
        Ag: (dryAg > 0 ? Math.max(0.0, (dryAg - cfg.a_thresholds.Ag) / dryAg) : 0.0) * pmFactor,
        Pd: (dryPd > 0 ? Math.max(0.0, (dryPd - cfg.a_thresholds.Pd) / dryPd) : 0.0) * pmFactor,
        Ni: (dryNi > 0 ? Math.max(0.0, (dryNi - cfg.a_thresholds.Ni) / dryNi) : 0.0) * pmFactor,
      };

      const payable = {
        Cu_kg: contained.Cu_kg * recRates.Cu,
        Au_g: contained.Au_g * recRates.Au,
        Ag_g: contained.Ag_g * recRates.Ag,
        Pd_g: contained.Pd_g * recRates.Pd,
        Ni_g: contained.Ni_g * recRates.Ni,
      };

      const netP = cfg.getCustomerANetPricesPerKgAndG();
      const revCu = payable.Cu_kg * netP.Cu;
      const revAu = payable.Au_g * netP.Au;
      const revAg = payable.Ag_g * netP.Ag;
      const revPd = payable.Pd_g * netP.Pd;
      const revNi = payable.Ni_g * netP.Ni;
      const grossRev = revCu + revAu + revAg + revPd + revNi;

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
        a_mode: cfg.a_mode,
        price_currency: cfg.price_currency,
        fill_container: cfg.fill_container,
        shipping_mode: cfg.shipping_mode || "allow_split",
        bag_count: subRows.length,
        WMT_kg: wKg,
        PM_WMT_kg: wKg,
        DMT_kg: dKg,
        moisture_pct: avgMoist,
        cu_diff: cuDiff,
        dry_grades: {
          "Cu(%)": dryCu,
          "Au(g/T)": dryAu,
          "Ag(g/T)": dryAg,
          "Pd(g/T)": dryPd,
          "Ni(g/T)": dryNi,
        },
        wet_grades: {
          "Cu(%)": sumWCu / wKg,
          "Au(g/T)": sumWAu / wKg,
          "Ag(g/T)": sumWAg / wKg,
          "Pd(g/T)": sumWPd / wKg,
          "Ni(g/T)": sumWNi / wKg,
        },
        contained,
        recovery_rates: recRates,
        payable_quantities: payable,
        metal_revenues: { Cu: revCu, Au: revAu, Ag: revAg, Pd: revPd, Ni: revNi },
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
      sumDPd = 0.0,
      sumDNi = 0.0;
    for (let i = 0; i < subRows.length; i++) {
      const r = subRows[i];
      sumWCu += r["倉庫重"] * r["濕基Cu(%)"];
      const d = r["乾重(kg)"];
      sumDCu += d * r["乾基Cu(%)"];
      sumDAu += d * r["火-Au (mg/Kg)"];
      sumDAg += d * r["Ag (mg/Kg)"];
      sumDPd += d * r["Pd (mg/Kg)"];
      sumDNi += d * r["Ni (mg/Kg)"];
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
      cfg.exclude_pm_materials &&
      cfg.exclude_pm_materials.length > 0;
    const exclMatSet = new Set(cfg.exclude_pm_materials || []);

    let wPmKg = 0.0;
    let sumPmAu = 0.0,
      sumPmAg = 0.0,
      sumPmPd = 0.0,
      sumPmNi = 0.0;

    for (let i = 0; i < subRows.length; i++) {
      const r = subRows[i];
      if (hasExclPm && exclMatSet.has(r["總表料號"])) {
        continue;
      }
      const w = r["倉庫重"];
      wPmKg += w;
      sumPmAu += w * r["溼基Au(mg/Kg)"];
      sumPmAg += w * r["溼基Ag(mg/Kg)"];
      sumPmPd += w * r["溼基Pd(mg/Kg)"];
      sumPmNi += w * r["溼基Ni(mg/Kg)"];
    }

    const wPmTon = wPmKg / 1000.0;
    const wetAu = wPmKg > 0 ? sumPmAu / wPmKg : 0.0;
    const wetAg = wPmKg > 0 ? sumPmAg / wPmKg : 0.0;
    const wetPd = wPmKg > 0 ? sumPmPd / wPmKg : 0.0;
    const wetNi = wPmKg > 0 ? sumPmNi / wPmKg : 0.0;

    const rates = {
      Cu: lookupTier(wetCu, cfg.tier_tables.Cu || []),
      Au: lookupTier(wetAu, cfg.tier_tables.Au || []),
      Ag: lookupTier(wetAg, cfg.tier_tables.Ag || []),
      Pd: lookupTier(wetPd, cfg.tier_tables.Pd || []),
      Ni: lookupTier(wetNi, cfg.tier_tables.Ni || []),
    };

    const unitP = cfg.getTierCustomerPricesPerKgAndG();
    const revCu = wKg * (wetCu / 100.0) * rates.Cu * unitP.Cu;
    const revAu = wPmTon * wetAu * rates.Au * unitP.Au;
    const revAg = wPmTon * wetAg * rates.Ag * unitP.Ag;
    const revPd = wPmTon * wetPd * rates.Pd * unitP.Pd;
    const revNi = wPmTon * wetNi * rates.Ni * unitP.Ni;
    const grossRev = revCu + revAu + revAg + revPd + revNi;

    const deduction = wKg * cfg.b_freight_per_kg + wTon * cfg.b_deduction_per_ton;
    const netProfit = grossRev - deduction;

    return {
      feasible: true,
      customer_type: cfg.customer_type,
      price_currency: cfg.price_currency,
      fill_container: cfg.fill_container,
      shipping_mode: cfg.shipping_mode || "allow_split",
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
        "Ni(g/T)": dKg > 0 ? sumDNi / dKg : 0.0,
      },
      wet_grades: {
        "Cu(%)": wetCu,
        "Au(g/T)": wetAu,
        "Ag(g/T)": wetAg,
        "Pd(g/T)": wetPd,
        "Ni(g/T)": wetNi,
      },
      tier_rates: rates,
      metal_revenues: { Cu: revCu, Au: revAu, Ag: revAg, Pd: revPd, Ni: revNi },
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
    if (nBags <= 1) {
      return {
        sortedBags: bBags,
        weights: new Float64Array(0),
        masks: new Int32Array(0),
        fallbackIndices: null,
      };
    }

    const sorted = bBags.slice().sort((a, b) => a["倉庫重"] - b["倉庫重"]);

    // nBags <= 30 時使用 1-kg 解析度之位元遮罩動態規劃 (1D Subset-Sum DP)，零陣列配置且涵蓋所有可達重量組合
    if (nBags <= 30) {
      const totalW = Math.ceil(sorted.reduce((a, r) => a + r["倉庫重"], 0));
      const dpMask = new Int32Array(totalW + 2);
      const dpExactW = new Float64Array(totalW + 2);
      const reachable = new Int32Array(totalW + 2);
      let rLen = 1;
      reachable[0] = 0;
      const fullMask = nBags === 30 ? 0x3fffffff : (1 << nBags) - 1;

      for (let i = 0; i < nBags; i++) {
        const w = sorted[i]["倉庫重"];
        const bit = 1 << i;
        const curLen = rLen;
        for (let j = 0; j < curLen; j++) {
          const prevB = reachable[j];
          const newM = dpMask[prevB] | bit;
          if (newM === fullMask) continue;
          const newW = (prevB === 0 ? 0.0 : dpExactW[prevB]) + w;
          const newB = Math.round(newW);
          if (dpMask[newB] === 0) {
            dpMask[newB] = newM;
            dpExactW[newB] = newW;
            reachable[rLen++] = newB;
          } else if (newW > dpExactW[newB]) {
            dpMask[newB] = newM;
            dpExactW[newB] = newW;
          }
        }
      }

      const buckets = reachable.subarray(1, rLen);
      buckets.sort();
      const count = buckets.length;
      const weights = new Float64Array(count);
      const masks = new Int32Array(count);
      for (let k = 0; k < count; k++) {
        const b = buckets[k];
        weights[k] = Math.round(dpExactW[b] * 1e6) / 1e6;
        masks[k] = dpMask[b];
      }
      return { sortedBags: sorted, weights, masks, fallbackIndices: null };
    }

    // nBags > 30 時之備援策略
    const idxWeightPairs = sorted.map((r) => [r._idx, r["倉庫重"]]);
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

    uniqueSplits.sort((a, b) => a[0] - b[0]);
    const weights = Float64Array.from(uniqueSplits.map((x) => x[0]));
    const fallbackIndices = uniqueSplits.map((x) => x[1]);
    return { sortedBags: sorted, weights, masks: null, fallbackIndices };
  }

  function decodeSplitBags(candObj, pos, bBags) {
    if (candObj.masks !== null) {
      const mask = candObj.masks[pos];
      const sorted = candObj.sortedBags;
      const chosen = [];
      for (let i = 0; i < sorted.length; i++) {
        if ((mask & (1 << i)) !== 0) {
          chosen.push(sorted[i]);
        }
      }
      return chosen;
    }
    const idxSet = new Set(candObj.fallbackIndices[pos]);
    return bBags.filter((r) => idxSet.has(r._idx));
  }

  function estimateBatchUnitValue(bp, cfg) {
    if (cfg.customer_type === "A") {
      const dryRatio = bp.dry_ratio;
      const { dry_cu: dCu, dry_au: dAu, dry_ag: dAg, dry_pd: dPd, dry_ni: dNi } = bp;
      const isCuOnly = cfg.a_mode === "cu_only";
      const recCu = isCuOnly
        ? lookupTier(dCu, cfg.a_cu_only_tier_table || [])
        : dCu > 0
        ? Math.max(0.0, (dCu - cfg.a_thresholds.Cu) / dCu)
        : 0.0;
      const recAu = isCuOnly ? 0.0 : dAu > 0 ? Math.max(0.0, (dAu - cfg.a_thresholds.Au) / dAu) : 0.0;
      const recAg = isCuOnly ? 0.0 : dAg > 0 ? Math.max(0.0, (dAg - cfg.a_thresholds.Ag) / dAg) : 0.0;
      const recPd = isCuOnly ? 0.0 : dPd > 0 ? Math.max(0.0, (dPd - cfg.a_thresholds.Pd) / dPd) : 0.0;
      const recNi = isCuOnly ? 0.0 : dNi > 0 ? Math.max(0.0, (dNi - cfg.a_thresholds.Ni) / dNi) : 0.0;
      const netP = cfg.getCustomerANetPricesPerKgAndG();
      const revPerDmtKg =
        dCu * 0.01 * recCu * netP.Cu +
        0.001 * dAu * recAu * netP.Au +
        0.001 * dAg * recAg * netP.Ag +
        0.001 * dPd * recPd * netP.Pd +
        0.001 * dNi * recNi * netP.Ni -
        0.001 * cfg.a_freight_per_dmt;
      const moistPen =
        (1.0 - dryRatio) * 100.0 > cfg.a_moisture_limit
          ? 0.001 * cfg.a_moisture_penalty_per_wmt
          : 0.0;
      return dryRatio * revPerDmtKg - moistPen;
    } else {
      const { wet_cu: wCu, wet_au: wAu, wet_ag: wAg, wet_pd: wPd, wet_ni: wNi, is_pm: isPm } = bp;
      const rCu = lookupTier(wCu, cfg.tier_tables.Cu || []);
      const rAu = lookupTier(wAu, cfg.tier_tables.Au || []);
      const rAg = lookupTier(wAg, cfg.tier_tables.Ag || []);
      const rPd = lookupTier(wPd, cfg.tier_tables.Pd || []);
      const rNi = lookupTier(wNi, cfg.tier_tables.Ni || []);
      const unitP = cfg.getTierCustomerPricesPerKgAndG();
      return (
        wCu * 0.01 * rCu * unitP.Cu +
        isPm * 0.001 * wAu * rAu * unitP.Au +
        isPm * 0.001 * wAg * rAg * unitP.Ag +
        isPm * 0.001 * wPd * rPd * unitP.Pd +
        isPm * 0.001 * wNi * rNi * unitP.Ni -
        cfg.b_freight_per_kg -
        0.001 * cfg.b_deduction_per_ton
      );
    }
  }

  function solveOptimalContainer(invRows, cfg, origTotalBagsMap = null) {
    const [wMin, wMax] = cfg.getActiveWeightBounds();

    const mustBatchSet = new Set(cfg.must_batch_ids || []);
    const mustFeedSet = new Set(cfg.must_feed_codes || []);
    const mustBagSet = new Set((cfg.must_bag_ids || []).map((x) => Number(x)));
    const bannedCodeSet = new Set(cfg.banned_codes || []);
    const bannedMatSet = new Set(cfg.banned_materials || []);
    const bannedBatchSet = new Set(cfg.banned_batch_ids || []);
    const bannedFeedSet = new Set(cfg.banned_feed_codes || []);
    const bannedBagSet = new Set((cfg.banned_bag_ids || []).map((x) => Number(x)));
    const exclPmMatSet = new Set(cfg.exclude_pm_materials || []);

    const dfMust = [];
    const dfOpt = [];
    const dfAvail = [];

    for (let i = 0; i < invRows.length; i++) {
      const r = invRows[i];
      const bagIdNum = Number(r["代號"]);
      const isMust =
        mustBatchSet.has(r["Batch_ID"]) ||
        mustFeedSet.has(r["進料公司名稱(代號)"]) ||
        mustBagSet.has(bagIdNum);
      const isBanned =
        bannedCodeSet.has(r["編碼料號"]) ||
        bannedMatSet.has(r["總表料號"]) ||
        bannedBatchSet.has(r["Batch_ID"]) ||
        bannedFeedSet.has(r["進料公司名稱(代號)"]) ||
        bannedBagSet.has(bagIdNum);
      const isAvail = isMust || (Boolean(r["已化驗"]) && !isBanned);
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
    if (origTotalBagsMap) {
      Object.assign(totalBagsPerBatch, origTotalBagsMap);
    } else {
      for (let i = 0; i < invRows.length; i++) {
        const bId = invRows[i]["Batch_ID"];
        totalBagsPerBatch[bId] = (totalBagsPerBatch[bId] || 0) + 1;
      }
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

    const isFullBatchOnly = cfg.shipping_mode === "full_batch";

    // 以「預估結算淨利 (net_profit) 最大化」為主要目標；僅當淨利完全相同時才以 objective_score (較少拆批) 作為次要比較
    let bestNetProfit = -1e18;
    let bestScore = -1e18;
    let bestSelectedRows = [];
    let bestEval = {
      feasible: false,
      reason: isFullBatchOnly
        ? `在「整批全出 (不拆批)」出貨方式下，無符合配滿櫃重量區間 [${Math.round(wMin).toLocaleString("en-US")}, ${Math.round(wMax).toLocaleString("en-US")}] kg 及品位/同櫃落差條件之整批組合，建議放寬每櫃濕重上下限或改選「允許拆批湊櫃優先」。`
        : `無符合配滿櫃重量區間 [${Math.round(wMin).toLocaleString("en-US")}, ${Math.round(wMax).toLocaleString("en-US")}] kg 及品位/同櫃落差條件之配櫃組合`,
      max_avail_wmt_kg: maxAvailWmtKg,
      must_wmt_kg: wMust,
    };

    function isBetterCandidate(candNetProfit, candObjScore) {
      if (candNetProfit > bestNetProfit + 1e-6) return true;
      if (Math.abs(candNetProfit - bestNetProfit) <= 1e-6 && candObjScore > bestScore + 1e-6) {
        return true;
      }
      return false;
    }

    const splitCandidatesByBatch = {};
    const splitWeightsByBatch = {};
    const batchProps = {};

    for (let i = 0; i < optBatches.length; i++) {
      const bId = optBatches[i];
      const bBags = batchDfMap[bId];
      const r0 = bBags[0];
      const isExcludedPm =
        cfg.customer_type === "E" && exclPmMatSet.has(r0["總表料號"]);
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
        dry_ni: Number(r0["Ni (mg/Kg)"]),
        wet_ni: Number(r0["溼基Ni(mg/Kg)"]),
        is_pm: isExcludedPm ? 0.0 : 1.0,
      };
      bp.unit_val = estimateBatchUnitValue(bp, cfg);
      batchProps[bId] = bp;

      if (!isFullBatchOnly) {
        const candObj = buildBatchSplitCandidates(bBags);
        splitCandidatesByBatch[bId] = candObj;
        splitWeightsByBatch[bId] = candObj.weights;
      }
    }

    function evaluateRowsBaseAndSplits(dfBase, excludedBatchSet, splitCandidateList = optBatches) {
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
        const res = evaluateSelection(dfBase, invRows, cfg, totalBagsPerBatch);
        if (res.feasible && res.gross_revenue <= cfg.max_total_value) {
          if (isBetterCandidate(res.net_profit, res.objective_score)) {
            bestNetProfit = res.net_profit;
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
          dPdBase = 0.0,
          dNiBase = 0.0;
        let wCuBase = 0.0,
          wPmBase = 0.0,
          wAuBase = 0.0,
          wAgBase = 0.0,
          wPdBase = 0.0,
          wNiBase = 0.0;

        if (cfg.customer_type === "A") {
          for (let i = 0; i < dfBase.length; i++) {
            const r = dfBase[i];
            const d = r["乾重(kg)"];
            dCuBase += d * r["乾基Cu(%)"];
            dAuBase += d * r["火-Au (mg/Kg)"];
            dAgBase += d * r["Ag (mg/Kg)"];
            dPdBase += d * r["Pd (mg/Kg)"];
            dNiBase += d * r["Ni (mg/Kg)"];
          }
        } else {
          const hasExcl = cfg.customer_type === "E" && exclPmMatSet.size > 0;
          for (let i = 0; i < dfBase.length; i++) {
            const r = dfBase[i];
            const w = r["倉庫重"];
            wCuBase += w * r["濕基Cu(%)"];
            if (!(hasExcl && exclPmMatSet.has(r["總表料號"]))) {
              wPmBase += w;
              wAuBase += w * r["溼基Au(mg/Kg)"];
              wAgBase += w * r["溼基Ag(mg/Kg)"];
              wPdBase += w * r["溼基Pd(mg/Kg)"];
              wNiBase += w * r["溼基Ni(mg/Kg)"];
            }
          }
        }

        const needMin = wMin - wBase;
        const needMax = wMax - wBase;

        for (let sIdx = 0; sIdx < splitCandidateList.length; sIdx++) {
          const splitB = splitCandidateList[sIdx];
          if (excludedBatchSet.has(splitB)) continue;
          const candObj = splitCandidatesByBatch[splitB];
          const wList = splitWeightsByBatch[splitB];
          if (!wList || wList.length === 0) continue;
          const bp = batchProps[splitB];
          const wetCuB = bp.wet_cu;

          const cuDiffCand = hasBase
            ? Math.max(cuMaxBase, wetCuB) - Math.min(cuMinBase, wetCuB)
            : 0.0;
          if (cuDiffCand > cfg.max_in_container_diff_cu + 1e-9) continue;

          const lo = bisectLeft(wList, needMin - 1e-9);
          const hi = bisectRight(wList, needMax + 1e-9);
          if (lo >= hi) continue;

          const nSplits = baseSplits + 1;
          const penaltyVal = cfg.split_batch_penalty * nSplits;
          let bestLocalPos = -1;

          if (cfg.customer_type === "A") {
            const isCuOnly = cfg.a_mode === "cu_only";
            const aCuTiers = cfg.a_cu_only_tier_table || [];
            const thCu = cfg.a_thresholds.Cu,
              thAu = cfg.a_thresholds.Au,
              thAg = cfg.a_thresholds.Ag,
              thPd = cfg.a_thresholds.Pd,
              thNi = cfg.a_thresholds.Ni;
            const netP = cfg.getCustomerANetPricesPerKgAndG();
            const netPCu = netP.Cu,
              netPAu = netP.Au,
              netPAg = netP.Ag,
              netPPd = netP.Pd,
              netPNi = netP.Ni;
            const dryRatio = bp.dry_ratio;
            const bDCu = bp.dry_cu,
              bDAu = bp.dry_au,
              bDAg = bp.dry_ag,
              bDPd = bp.dry_pd,
              bDNi = bp.dry_ni;

            for (let idxPos = lo; idxPos < hi; idxPos++) {
              const wSub = wList[idxPos];
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
              const dryNi = (dNiBase + dSub * bDNi) / dKg;

              const recCu = isCuOnly
                ? lookupTier(dryCu, aCuTiers)
                : dryCu > 0
                ? Math.max(0.0, (dryCu - thCu) / dryCu)
                : 0.0;
              const recAu = isCuOnly ? 0.0 : dryAu > 0 ? Math.max(0.0, (dryAu - thAu) / dryAu) : 0.0;
              const recAg = isCuOnly ? 0.0 : dryAg > 0 ? Math.max(0.0, (dryAg - thAg) / dryAg) : 0.0;
              const recPd = isCuOnly ? 0.0 : dryPd > 0 ? Math.max(0.0, (dryPd - thPd) / dryPd) : 0.0;
              const recNi = isCuOnly ? 0.0 : dryNi > 0 ? Math.max(0.0, (dryNi - thNi) / dryNi) : 0.0;

              const grossRev =
                dKg * (dryCu * 0.01) * recCu * netPCu +
                dTon * dryAu * recAu * netPAu +
                dTon * dryAg * recAg * netPAg +
                dTon * dryPd * recPd * netPPd +
                dTon * dryNi * recNi * netPNi;
              if (grossRev > cfg.max_total_value) continue;

              const freightCost = dTon * cfg.a_freight_per_dmt;
              const moistPenalty =
                avgMoist > cfg.a_moisture_limit ? wTon * cfg.a_moisture_penalty_per_wmt : 0.0;
              const tcPenalty =
                dryCu < thCu
                  ? wTon * cfg.a_cu_below_threshold_tc_per_wmt + cfg.a_assay_fee_if_below
                  : 0.0;
              const netProfit = grossRev - freightCost - moistPenalty - tcPenalty;
              const objScore = netProfit - penaltyVal;
              if (isBetterCandidate(netProfit, objScore)) {
                bestNetProfit = netProfit;
                bestScore = objScore;
                bestLocalPos = idxPos;
              }
            }
          } else {
            const cuLo = cfg.target_cu_range[0] - 1e-9;
            const cuHi = cfg.target_cu_range[1] + 1e-9;
            const tiersCu = cfg.tier_tables.Cu || [];
            const tiersAu = cfg.tier_tables.Au || [];
            const tiersAg = cfg.tier_tables.Ag || [];
            const tiersPd = cfg.tier_tables.Pd || [];
            const tiersNi = cfg.tier_tables.Ni || [];
            const unitP = cfg.getTierCustomerPricesPerKgAndG();
            const pCu = unitP.Cu,
              pAu = unitP.Au,
              pAg = unitP.Ag,
              pPd = unitP.Pd,
              pNi = unitP.Ni;
            const isPmB = bp.is_pm;
            const bWAu = bp.wet_au,
              bWAg = bp.wet_ag,
              bWPd = bp.wet_pd,
              bWNi = bp.wet_ni;

            for (let idxPos = lo; idxPos < hi; idxPos++) {
              const wSub = wList[idxPos];
              const wKg = wBase + wSub;
              const wetCu = (wCuBase + wSub * wetCuB) / wKg;
              if (!(wetCu >= cuLo && wetCu <= cuHi)) continue;

              const wPmKg = wPmBase + wSub * isPmB;
              const wPmTon = wPmKg * 0.001;
              let wetAu = 0.0,
                wetAg = 0.0,
                wetPd = 0.0,
                wetNi = 0.0;
              if (wPmKg > 0) {
                wetAu = (wAuBase + wSub * isPmB * bWAu) / wPmKg;
                wetAg = (wAgBase + wSub * isPmB * bWAg) / wPmKg;
                wetPd = (wPdBase + wSub * isPmB * bWPd) / wPmKg;
                wetNi = (wNiBase + wSub * isPmB * bWNi) / wPmKg;
              }

              const rCu = lookupTier(wetCu, tiersCu);
              const rAu = lookupTier(wetAu, tiersAu);
              const rAg = lookupTier(wetAg, tiersAg);
              const rPd = lookupTier(wetPd, tiersPd);
              const rNi = lookupTier(wetNi, tiersNi);

              const grossRev =
                wKg * (wetCu * 0.01) * rCu * pCu +
                wPmTon * wetAu * rAu * pAu +
                wPmTon * wetAg * rAg * pAg +
                wPmTon * wetPd * rPd * pPd +
                wPmTon * wetNi * rNi * pNi;
              if (grossRev > cfg.max_total_value) continue;

              const deduction =
                wKg * cfg.b_freight_per_kg + wKg * 0.001 * cfg.b_deduction_per_ton;
              const netProfit = grossRev - deduction;
              const objScore = netProfit - penaltyVal;
              if (isBetterCandidate(netProfit, objScore)) {
                bestNetProfit = netProfit;
                bestScore = objScore;
                bestLocalPos = idxPos;
              }
            }
          }

          if (bestLocalPos !== -1) {
            const bBags = batchDfMap[splitB];
            const splitChosen = decodeSplitBags(candObj, bestLocalPos, bBags);
            const dfCand = dfBase.concat(splitChosen).sort((a, b) => a["代號"] - b["代號"]);
            bestSelectedRows = dfCand;
            bestEval = evaluateSelection(bestSelectedRows, invRows, cfg, totalBagsPerBatch);
            bestNetProfit = bestEval.net_profit;
            bestScore = bestEval.objective_score;
          }
        }
      }
    }

    function evaluateBaseAndSplits(selectedFullBatches, splitCandidateList = optBatches) {
      const selectedSet = new Set(selectedFullBatches);
      const dfBase = dfMust.slice();
      for (let i = 0; i < selectedFullBatches.length; i++) {
        const bBags = batchDfMap[selectedFullBatches[i]];
        for (let j = 0; j < bBags.length; j++) {
          dfBase.push(bBags[j]);
        }
      }
      evaluateRowsBaseAndSplits(dfBase, selectedSet, splitCandidateList);
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

    const sortedOptByVal = optBatches
      .slice()
      .sort((a, b) => batchProps[b].unit_val - batchProps[a].unit_val);
    const sortedOptByCu = optBatches
      .slice()
      .sort((a, b) => batchProps[b].wet_cu - batchProps[a].wet_cu);
    const fastSplitPool = Array.from(
      new Set([...sortedOptByVal.slice(0, 24), ...sortedOptByCu.slice(0, 12)])
    );

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

    // 整批全出組合深度搜尋 (透過重量二分搜尋收尾批次，涵蓋全部 100% 可用整批，支援 0 ~ 6+ 批整批全出組合)
    const fullBatchesByWeight = eligibleFullBatches
      .slice()
      .sort((a, b) => batchProps[a].w_kg - batchProps[b].w_kg);
    const fullBatchWeightsArr = Float64Array.from(
      fullBatchesByWeight.map((bId) => batchProps[bId].w_kg)
    );
    const chosenFullSet = new Set();

    function checkFullBatchBaseAndClosing(curW, cMin, cMax, curRows) {
      if (curW >= wMin - 1e-9 && curW <= wMax + 1e-9) {
        const res = evaluateSelection(curRows, invRows, cfg, totalBagsPerBatch);
        if (res.feasible && res.gross_revenue <= cfg.max_total_value) {
          if (isBetterCandidate(res.net_profit, res.objective_score)) {
            bestNetProfit = res.net_profit;
            bestScore = res.objective_score;
            bestSelectedRows = curRows.slice().sort((a, b) => a["代號"] - b["代號"]);
            bestEval = res;
          }
        }
      }
      const needMin = wMin - curW;
      const needMax = wMax - curW;
      if (needMax <= 1e-9) return;
      const lo = bisectLeft(fullBatchWeightsArr, needMin - 1e-9);
      const hi = bisectRight(fullBatchWeightsArr, needMax + 1e-9);
      for (let pos = lo; pos < hi; pos++) {
        const bLast = fullBatchesByWeight[pos];
        if (chosenFullSet.has(bLast)) continue;
        const bpLast = batchProps[bLast];
        if (
          Math.max(cMax, bpLast.wet_cu) - Math.min(cMin, bpLast.wet_cu) >
          cfg.max_in_container_diff_cu + 1e-9
        ) {
          continue;
        }
        const candRows = curRows.concat(batchDfMap[bLast]);
        const res = evaluateSelection(candRows, invRows, cfg, totalBagsPerBatch);
        if (res.feasible && res.gross_revenue <= cfg.max_total_value) {
          if (isBetterCandidate(res.net_profit, res.objective_score)) {
            bestNetProfit = res.net_profit;
            bestScore = res.objective_score;
            bestSelectedRows = candRows.slice().sort((a, b) => a["代號"] - b["代號"]);
            bestEval = res;
          }
        }
      }
    }

    const fullDfsPool = eligibleFullBatches.slice(0, 42);
    let fullDfsCalls = 0;
    function runFullBatchDfs(startIdx, depth, curW, cMin, cMax, curRows) {
      fullDfsCalls++;
      checkFullBatchBaseAndClosing(curW, cMin, cMax, curRows);
      if (depth >= 5 || fullDfsCalls > 6500 || curW >= wMax - 600.0) return;
      const limit =
        depth === 0
          ? fullDfsPool.length
          : depth === 1
          ? Math.min(fullDfsPool.length, 32)
          : Math.min(fullDfsPool.length, 20);
      for (let i = startIdx; i < limit; i++) {
        const bId = fullDfsPool[i];
        const bp = batchProps[bId];
        if (curW + bp.w_kg > wMax + 1e-9) continue;
        const nMin = Math.min(cMin, bp.wet_cu);
        const nMax = Math.max(cMax, bp.wet_cu);
        if (nMax - nMin > cfg.max_in_container_diff_cu + 1e-9) continue;
        chosenFullSet.add(bId);
        runFullBatchDfs(
          i + 1,
          depth + 1,
          curW + bp.w_kg,
          nMin,
          nMax,
          curRows.concat(batchDfMap[bId])
        );
        chosenFullSet.delete(bId);
      }
    }

    runFullBatchDfs(0, 0, wMust, cuMinMust, cuMaxMust, dfMust);

    // 若使用者選擇「整批全出 (full_batch)」模式，直接回傳整批全出最佳解，不進行任何拆批湊櫃
    if (isFullBatchOnly) {
      return [bestSelectedRows, bestEval];
    }

    evaluateBaseAndSplits([]);

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

        evaluateBaseAndSplits([b1, b2], fastSplitPool);

        if (w2 < wMax - 800.0) {
          const maxK = Math.min(nPool, 18);
          for (let k = j + 1; k < maxK; k++) {
            const b3 = topPool[k];
            const bp3 = batchProps[b3];
            const w3 = w2 + bp3.w_kg;
            if (w3 > wMax + 1e-9) continue;
            const cMin3 = Math.min(cMin2, bp3.wet_cu);
            const cMax3 = Math.max(cMax2, bp3.wet_cu);
            if (cMax3 - cMin3 > cfg.max_in_container_diff_cu + 1e-9) continue;
            evaluateBaseAndSplits([b1, b2, b3], fastSplitPool);

            // 支援 4 批以上小型高單價整批連加湊櫃 (貪婪延伸)
            if (w3 < wMax - 1200.0 && k < 10) {
              const chain = [b1, b2, b3];
              let wCur = w3;
              let cMinCur = cMin3;
              let cMaxCur = cMax3;
              for (let m = k + 1; m < nPool; m++) {
                const bm = topPool[m];
                const bpm = batchProps[bm];
                if (wCur + bpm.w_kg > wMax + 1e-9) continue;
                const nmMin = Math.min(cMinCur, bpm.wet_cu);
                const nmMax = Math.max(cMaxCur, bpm.wet_cu);
                if (nmMax - nmMin > cfg.max_in_container_diff_cu + 1e-9) continue;
                chain.push(bm);
                wCur += bpm.w_kg;
                cMinCur = nmMin;
                cMaxCur = nmMax;
                evaluateBaseAndSplits(chain, fastSplitPool);
                if (wCur >= wMin - 1e-9) break;
              }
            }
          }
        }
      }
    }

    // 雙拆批跨批進階尋優：針對高價值基底 (空基底或前 6 大高價值整批)，篩選與基底相容之高單價批與高銅拉抬批，先拆入第 1 批部分包數，再由第 2 批二分搜尋補足至 [wMin, wMax]
    const twoSplitBaseCandidates = [[]];
    for (let i = 0; i < Math.min(6, eligibleFullBatches.length); i++) {
      const bId = eligibleFullBatches[i];
      if (wMust + batchProps[bId].w_kg < wMax - 1200.0) {
        twoSplitBaseCandidates.push([bId]);
      }
    }

    for (let bIdx = 0; bIdx < twoSplitBaseCandidates.length; bIdx++) {
      const baseBatches = twoSplitBaseCandidates[bIdx];
      const baseRows = dfMust.slice();
      for (let i = 0; i < baseBatches.length; i++) {
        const bBags = batchDfMap[baseBatches[i]];
        for (let j = 0; j < bBags.length; j++) baseRows.push(bBags[j]);
      }
      let wBase0 = 0.0;
      let cMin0 = Infinity;
      let cMax0 = -Infinity;
      for (let i = 0; i < baseRows.length; i++) {
        wBase0 += baseRows[i]["倉庫重"];
        const c = baseRows[i]["濕基Cu(%)"];
        if (c < cMin0) cMin0 = c;
        if (c > cMax0) cMax0 = c;
      }

      const compatValBatches = [];
      const compatCuBatches = [];
      for (let i = 0; i < sortedOptByVal.length; i++) {
        const bId = sortedOptByVal[i];
        if (baseBatches.includes(bId)) continue;
        const bp = batchProps[bId];
        if (
          baseRows.length > 0 &&
          Math.max(cMax0, bp.wet_cu) - Math.min(cMin0, bp.wet_cu) >
            cfg.max_in_container_diff_cu + 1e-9
        ) {
          continue;
        }
        if (compatValBatches.length < 7) compatValBatches.push(bId);
      }
      for (let i = 0; i < sortedOptByCu.length; i++) {
        const bId = sortedOptByCu[i];
        if (baseBatches.includes(bId)) continue;
        const bp = batchProps[bId];
        if (
          baseRows.length > 0 &&
          Math.max(cMax0, bp.wet_cu) - Math.min(cMin0, bp.wet_cu) >
            cfg.max_in_container_diff_cu + 1e-9
        ) {
          continue;
        }
        if (compatCuBatches.length < 5) compatCuBatches.push(bId);
      }
      const firstSplitPool = Array.from(new Set([...compatValBatches, ...compatCuBatches]));
      const secondSplitPool = Array.from(
        new Set([...compatValBatches, ...compatCuBatches, ...fastSplitPool.slice(0, 14)])
      );

      for (let s1 = 0; s1 < firstSplitPool.length; s1++) {
        const sb1 = firstSplitPool[s1];
        const sb1Raw = batchDfMap[sb1];
        if (sb1Raw.length <= 1) continue;
        const exclSet = new Set([...baseBatches, sb1]);

        // 同時嘗試由重到輕與由輕到重取 1~3 包，涵蓋跨階梯所需之輕包或重包組合
        const sb1Desc = sb1Raw.slice().sort((a, b) => b["倉庫重"] - a["倉庫重"]);
        const sb1Asc = sb1Raw.slice().sort((a, b) => a["倉庫重"] - b["倉庫重"]);
        const maxTake = Math.min(3, sb1Raw.length - 1);

        let augDesc = baseRows.slice();
        let wDesc = wBase0;
        for (let k = 1; k <= maxTake; k++) {
          augDesc.push(sb1Desc[k - 1]);
          wDesc += sb1Desc[k - 1]["倉庫重"];
          if (wDesc >= wMax - 500.0) break;
          evaluateRowsBaseAndSplits(augDesc, exclSet, secondSplitPool);
        }

        let augAsc = baseRows.slice();
        let wAsc = wBase0;
        for (let k = 1; k <= maxTake; k++) {
          augAsc.push(sb1Asc[k - 1]);
          wAsc += sb1Asc[k - 1]["倉庫重"];
          if (wAsc >= wMax - 500.0) break;
          evaluateRowsBaseAndSplits(augAsc, exclSet, secondSplitPool);
        }
      }
    }

    // 單包層級局部替換精煉 (Hill-Climbing Bag-Level Refinement)：在 [wMin, wMax] 內嘗試加包或 1-對-1 換包以確保淨利極大化
    if (bestEval.feasible && bestSelectedRows.length > 0) {
      const mustIdxSet = new Set(dfMust.map((r) => r._idx));
      let improved = true;
      let iter = 0;
      while (improved && iter < 15) {
        improved = false;
        iter++;
        const selIdxSet = new Set(bestSelectedRows.map((r) => r._idx));
        const curW = bestEval.WMT_kg;

        // 1. 若還有剩餘重量空間，嘗試直接多加 1 包可用袋
        for (let i = 0; i < dfOpt.length; i++) {
          const rIn = dfOpt[i];
          if (selIdxSet.has(rIn._idx)) continue;
          if (curW + rIn["倉庫重"] > wMax + 1e-9) continue;
          const candRows = bestSelectedRows.concat([rIn]);
          const res = evaluateSelection(candRows, invRows, cfg, totalBagsPerBatch);
          if (res.feasible && res.gross_revenue <= cfg.max_total_value) {
            if (isBetterCandidate(res.net_profit, res.objective_score)) {
              bestNetProfit = res.net_profit;
              bestScore = res.objective_score;
              bestSelectedRows = candRows.sort((a, b) => a["代號"] - b["代號"]);
              bestEval = res;
              improved = true;
              break;
            }
          }
        }
        if (improved) continue;

        // 2. 嘗試 1-對-1 換包 (將已選非必出包換成未選之高價值可用包)
        for (let i = 0; i < bestSelectedRows.length; i++) {
          const rOut = bestSelectedRows[i];
          if (mustIdxSet.has(rOut._idx)) continue;
          const baseW = curW - rOut["倉庫重"];
          for (let j = 0; j < dfOpt.length; j++) {
            const rIn = dfOpt[j];
            if (selIdxSet.has(rIn._idx)) continue;
            const newW = baseW + rIn["倉庫重"];
            if (newW < wMin - 1e-9 || newW > wMax + 1e-9) continue;
            // 若同批且新包不比舊包重，對單價無正面幫助則跳過
            if (rIn["Batch_ID"] === rOut["Batch_ID"] && rIn["倉庫重"] <= rOut["倉庫重"]) continue;

            const candRows = bestSelectedRows.slice();
            candRows[i] = rIn;
            const res = evaluateSelection(candRows, invRows, cfg, totalBagsPerBatch);
            if (res.feasible && res.gross_revenue <= cfg.max_total_value) {
              if (isBetterCandidate(res.net_profit, res.objective_score)) {
                bestNetProfit = res.net_profit;
                bestScore = res.objective_score;
                bestSelectedRows = candRows.sort((a, b) => a["代號"] - b["代號"]);
                bestEval = res;
                improved = true;
                break;
              }
            }
          }
          if (improved) break;
        }
      }
    }

    return [bestSelectedRows, bestEval];
  }

  /**
   * 支援配多櫃 (num_containers >= 1) 之最佳化配滿櫃求解器
   * 每一櫃皆依據配滿櫃模式 (wMin ~ wMax kg) 獨立尋優並計價，已配出之太空包於後續櫃次自動從可用庫存扣除。
   */
  function solveMultipleContainers(invRows, cfg) {
    const requestedContainers = Math.max(1, Math.trunc(Number(cfg.num_containers) || 1));

    const origTotalBagsMap = {};
    for (let i = 0; i < invRows.length; i++) {
      const bId = invRows[i]["Batch_ID"];
      origTotalBagsMap[bId] = (origTotalBagsMap[bId] || 0) + 1;
    }

    const usedBagIdxSet = new Set();
    const containerResults = [];
    const allSelectedRows = [];

    for (let cIdx = 1; cIdx <= requestedContainers; cIdx++) {
      const remainingRows = invRows.filter((r) => !usedBagIdxSet.has(r._idx));

      // 第一櫃優先包含使用者勾選之必出料；若部分必出料於前一櫃尚未配完則留至下一櫃，若已全數配入則後續櫃次自由尋優
      const [selRows, evalRes] = solveOptimalContainer(remainingRows, cfg, origTotalBagsMap);

      if (!evalRes.feasible || selRows.length === 0) {
        if (cIdx === 1) {
          return [
            [],
            Object.assign({}, evalRes, {
              requested_containers: requestedContainers,
              completed_containers: 0,
              container_results: [],
            }),
          ];
        }
        containerResults.push({
          container_index: cIdx,
          feasible: false,
          reason: evalRes.reason,
          selected_rows: [],
          eval_res: evalRes,
        });
        break;
      }

      const taggedRows = selRows.map((r) =>
        Object.assign({}, r, {
          Container_ID: `第 ${cIdx} 櫃`,
          _container_index: cIdx,
        })
      );

      for (let i = 0; i < selRows.length; i++) {
        usedBagIdxSet.add(selRows[i]._idx);
      }
      for (let i = 0; i < taggedRows.length; i++) {
        allSelectedRows.push(taggedRows[i]);
      }

      containerResults.push({
        container_index: cIdx,
        container_label: `第 ${cIdx} 櫃`,
        feasible: true,
        selected_rows: taggedRows,
        eval_res: evalRes,
      });
    }

    const feasibleContainers = containerResults.filter((c) => c.feasible);
    const completedCount = feasibleContainers.length;

    if (completedCount === 0) {
      return [
        [],
        {
          feasible: false,
          reason: "無符合配滿櫃條件之組合",
          requested_containers: requestedContainers,
          completed_containers: 0,
          container_results: [],
        },
      ];
    }

    if (completedCount === 1 && requestedContainers === 1) {
      const singleRes = Object.assign({}, feasibleContainers[0].eval_res, {
        requested_containers: 1,
        completed_containers: 1,
        container_results: feasibleContainers,
      });
      return [allSelectedRows, singleRes];
    }

    // 彙整多櫃總計與加權平均指標 (各櫃仍按配滿櫃獨立計價並加總毛收、扣款與淨利)
    let totalWmtKg = 0.0;
    let totalPmWmtKg = 0.0;
    let totalDmtKg = 0.0;
    let maxCuDiff = 0.0;
    let totalGrossRev = 0.0;
    let totalDeduction = 0.0;
    let totalFreightCost = 0.0;
    let totalMoistPenalty = 0.0;
    let totalTcPenalty = 0.0;
    let totalNetProfit = 0.0;
    let totalObjScore = 0.0;

    const totalMetalRevs = { Cu: 0.0, Au: 0.0, Ag: 0.0, Pd: 0.0, Ni: 0.0 };
    const totalContained = { Cu_kg: 0.0, Au_g: 0.0, Ag_g: 0.0, Pd_g: 0.0, Ni_g: 0.0 };
    const totalPayable = { Cu_kg: 0.0, Au_g: 0.0, Ag_g: 0.0, Pd_g: 0.0, Ni_g: 0.0 };

    let sumDCu = 0.0,
      sumDAu = 0.0,
      sumDAg = 0.0,
      sumDPd = 0.0,
      sumDNi = 0.0;
    let sumWCu = 0.0,
      sumWAu = 0.0,
      sumWAg = 0.0,
      sumWPd = 0.0,
      sumWNi = 0.0;

    const allFullBatches = [];
    const allSplitBatches = [];

    for (let i = 0; i < feasibleContainers.length; i++) {
      const cObj = feasibleContainers[i];
      const er = cObj.eval_res;
      const cTag = `第${cObj.container_index}櫃`;

      totalWmtKg += er.WMT_kg;
      totalPmWmtKg += er.PM_WMT_kg;
      totalDmtKg += er.DMT_kg;
      if (er.cu_diff > maxCuDiff) maxCuDiff = er.cu_diff;

      totalGrossRev += er.gross_revenue;
      totalDeduction += er.deduction;
      totalFreightCost += er.freight_cost || 0.0;
      totalMoistPenalty += er.moisture_penalty || 0.0;
      totalTcPenalty += er.tc_penalty || 0.0;
      totalNetProfit += er.net_profit;
      totalObjScore += er.objective_score;

      for (const m of ["Cu", "Au", "Ag", "Pd", "Ni"]) {
        totalMetalRevs[m] += (er.metal_revenues && er.metal_revenues[m]) || 0.0;
      }

      if (er.contained) {
        for (const k of Object.keys(totalContained)) {
          totalContained[k] += er.contained[k] || 0.0;
        }
      }
      if (er.payable_quantities) {
        for (const k of Object.keys(totalPayable)) {
          totalPayable[k] += er.payable_quantities[k] || 0.0;
        }
      }

      sumDCu += er.DMT_kg * er.dry_grades["Cu(%)"];
      sumDAu += er.DMT_kg * er.dry_grades["Au(g/T)"];
      sumDAg += er.DMT_kg * er.dry_grades["Ag(g/T)"];
      sumDPd += er.DMT_kg * er.dry_grades["Pd(g/T)"];
      sumDNi += er.DMT_kg * er.dry_grades["Ni(g/T)"];

      const pmDenom = cfg.customer_type === "A" ? er.WMT_kg : er.PM_WMT_kg;
      sumWCu += er.WMT_kg * er.wet_grades["Cu(%)"];
      sumWAu += pmDenom * er.wet_grades["Au(g/T)"];
      sumWAg += pmDenom * er.wet_grades["Ag(g/T)"];
      sumWPd += pmDenom * er.wet_grades["Pd(g/T)"];
      sumWNi += pmDenom * er.wet_grades["Ni(g/T)"];

      for (const fb of er.full_batches || []) {
        allFullBatches.push(`${cTag}:${fb}`);
      }
      for (const sb of er.split_batches || []) {
        allSplitBatches.push(`${cTag}:${sb}`);
      }
    }

    const totalMoistPct = totalWmtKg > 0 ? (1.0 - totalDmtKg / totalWmtKg) * 100.0 : 0.0;
    const totalPmDenom = cfg.customer_type === "A" ? totalWmtKg : totalPmWmtKg;

    const aggDryGrades = {
      "Cu(%)": totalDmtKg > 0 ? sumDCu / totalDmtKg : 0.0,
      "Au(g/T)": totalDmtKg > 0 ? sumDAu / totalDmtKg : 0.0,
      "Ag(g/T)": totalDmtKg > 0 ? sumDAg / totalDmtKg : 0.0,
      "Pd(g/T)": totalDmtKg > 0 ? sumDPd / totalDmtKg : 0.0,
      "Ni(g/T)": totalDmtKg > 0 ? sumDNi / totalDmtKg : 0.0,
    };

    const aggWetGrades = {
      "Cu(%)": totalWmtKg > 0 ? sumWCu / totalWmtKg : 0.0,
      "Au(g/T)": totalPmDenom > 0 ? sumWAu / totalPmDenom : 0.0,
      "Ag(g/T)": totalPmDenom > 0 ? sumWAg / totalPmDenom : 0.0,
      "Pd(g/T)": totalPmDenom > 0 ? sumWPd / totalPmDenom : 0.0,
      "Ni(g/T)": totalPmDenom > 0 ? sumWNi / totalPmDenom : 0.0,
    };

    const aggRecRates = {
      Cu: totalContained.Cu_kg > 0 ? totalPayable.Cu_kg / totalContained.Cu_kg : 0.0,
      Au: totalContained.Au_g > 0 ? totalPayable.Au_g / totalContained.Au_g : 0.0,
      Ag: totalContained.Ag_g > 0 ? totalPayable.Ag_g / totalContained.Ag_g : 0.0,
      Pd: totalContained.Pd_g > 0 ? totalPayable.Pd_g / totalContained.Pd_g : 0.0,
      Ni: totalContained.Ni_g > 0 ? totalPayable.Ni_g / totalContained.Ni_g : 0.0,
    };

    const aggTierRates = {
      Cu: lookupTier(aggWetGrades["Cu(%)"], cfg.tier_tables.Cu || []),
      Au: lookupTier(aggWetGrades["Au(g/T)"], cfg.tier_tables.Au || []),
      Ag: lookupTier(aggWetGrades["Ag(g/T)"], cfg.tier_tables.Ag || []),
      Pd: lookupTier(aggWetGrades["Pd(g/T)"], cfg.tier_tables.Pd || []),
      Ni: lookupTier(aggWetGrades["Ni(g/T)"], cfg.tier_tables.Ni || []),
    };

    const unfulfilledReason =
      completedCount < requestedContainers
        ? containerResults[containerResults.length - 1].reason
        : null;

    return [
      allSelectedRows,
      {
        feasible: true,
        customer_type: cfg.customer_type,
        a_mode: cfg.a_mode,
        price_currency: cfg.price_currency,
        fill_container: cfg.fill_container,
        requested_containers: requestedContainers,
        completed_containers: completedCount,
        unfulfilled_reason: unfulfilledReason,
        container_results: containerResults,
        bag_count: allSelectedRows.length,
        WMT_kg: totalWmtKg,
        PM_WMT_kg: totalPmWmtKg,
        DMT_kg: totalDmtKg,
        moisture_pct: totalMoistPct,
        cu_diff: maxCuDiff,
        dry_grades: aggDryGrades,
        wet_grades: aggWetGrades,
        contained: totalContained,
        recovery_rates: aggRecRates,
        payable_quantities: totalPayable,
        tier_rates: aggTierRates,
        metal_revenues: totalMetalRevs,
        gross_revenue: totalGrossRev,
        freight_cost: totalFreightCost,
        moisture_penalty: totalMoistPenalty,
        tc_penalty: totalTcPenalty,
        deduction: totalDeduction,
        net_profit: totalNetProfit,
        full_batches: allFullBatches,
        split_batches: allSplitBatches,
        objective_score: totalObjScore,
      },
    ];
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
      const cId = r["Container_ID"] || "第 1 櫃";
      const bId = r["Batch_ID"];
      const gKey = `${cId}__${bId}`;
      if (!subBatchGroups.has(gKey)) subBatchGroups.set(gKey, []);
      subBatchGroups.get(gKey).push(r);
    }

    const rows = [];
    for (const [, grp] of subBatchGroups.entries()) {
      const r0 = grp[0];
      const bId = r0["Batch_ID"];
      const cId = r0["Container_ID"] || "第 1 櫃";
      const totalBBags = fullBatchCounts.get(bId) || grp.length;
      const wSub = grp.reduce((acc, r) => acc + r["倉庫重"], 0.0);
      const dSub = grp.reduce((acc, r) => acc + r["乾重(kg)"], 0.0);
      const siteSub = grp.reduce((acc, r) => acc + r["現場重"], 0.0);
      const isFull = grp.length === totalBBags;
      rows.push({
        櫃次: cId,
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
        "乾基Ni(g/T)": roundTo(r0["Ni (mg/Kg)"], 2),
        "溼基Ni(g/T)": roundTo(r0["溼基Ni(mg/Kg)"], 2),
        "現場重小計(kg)": roundTo(siteSub, 1),
        "倉庫濕重小計(kg)": roundTo(wSub, 1),
        "乾重小計(kg)": roundTo(dSub, 2),
      });
    }
    return rows;
  }

  function buildSelectedBagDetail(selectedRows) {
    if (!Array.isArray(selectedRows) || selectedRows.length === 0) return [];
    const sorted = selectedRows.slice().sort((a, b) => {
      const cA = a._container_index || 1;
      const cB = b._container_index || 1;
      if (cA !== cB) return cA - cB;
      return a["代號"] - b["代號"];
    });
    return sorted.map((row) => ({
      櫃次: row["Container_ID"] || "第 1 櫃",
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
      "乾基Ni(g/T)": roundTo(row["Ni (mg/Kg)"], 4),
      "溼基Ni(g/T)": roundTo(row["溼基Ni(mg/Kg)"], 4),
    }));
  }

  function exportBlendingResultToExcel(selectedRows, fullRows, evalRes, cfg) {
    if (typeof XLSX === "undefined") {
      throw new Error("SheetJS (XLSX) 尚未載入，無法匯出 Excel 檔案");
    }

    const batchSum = buildSelectedBatchSummary(selectedRows, fullRows);
    const bagDet = buildSelectedBagDetail(selectedRows);
    const curr = cfg.price_currency;
    const reqC = evalRes.requested_containers || cfg.num_containers || 1;
    const compC = evalRes.completed_containers || 1;

    const summaryRows = [
      { 項目: "客戶類型", 數值: `${cfg.customer_type} 客戶` },
      { 項目: "計價幣別", 數值: `${curr} (${curr === "JPY" ? "日圓計價 ¥" : "美元計價 $"})` },
      {
        項目: "配櫃模式與櫃數",
        數值: `配滿櫃模式 (設定 ${reqC} 櫃 / 實際配出 ${compC} 櫃，每櫃 ${Math.round(cfg.full_container_range[0]).toLocaleString("en-US")} ~ ${Math.round(cfg.full_container_range[1]).toLocaleString("en-US")} kg)`,
      },
      {
        項目: "出貨方式 (shipping_mode)",
        數值:
          cfg.shipping_mode === "full_batch"
            ? "整批全出 (不拆批，計算整批組合之預估結算淨利最大化)"
            : "允許拆批湊櫃優先 (可拆批湊滿櫃，計算預估結算淨利最大化)",
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
      {
        項目: "加權濕基 Ni (g/T)",
        數值: `${((evalRes.wet_grades && evalRes.wet_grades["Ni(g/T)"]) || 0).toFixed(4)}`,
      },
      { 項目: `毛收入 (Gross Revenue, ${curr})`, 數值: roundTo(evalRes.gross_revenue || 0, 2) },
      { 項目: `總扣減與運費罰款 (Deductions, ${curr})`, 數值: roundTo(evalRes.deduction || 0, 2) },
      { 項目: `預估結算淨利 (Net Profit, ${curr})`, 數值: roundTo(evalRes.net_profit || 0, 2) },
      { 項目: "綜合目標分數 (含拆批懲罰)", 數值: roundTo(evalRes.objective_score || 0, 2) },
    ];

    if (Array.isArray(evalRes.container_results) && evalRes.container_results.length > 1) {
      for (const cObj of evalRes.container_results) {
        if (!cObj.feasible) continue;
        const er = cObj.eval_res;
        summaryRows.push({
          項目: `【${cObj.container_label}】小計 (${er.bag_count}包)`,
          數值: `濕重: ${roundTo(er.WMT_kg, 1)} kg | 乾重: ${roundTo(er.DMT_kg, 2)} kg | 濕Cu: ${er.wet_grades["Cu(%)"].toFixed(2)}% | 毛收: ${curr} ${roundTo(er.gross_revenue, 2)} | 淨利: ${curr} ${roundTo(er.net_profit, 2)}`,
        });
      }
    }

    const cfgDict = cfg.toDict();
    const paramRows = [];
    for (const [k, v] of Object.entries(cfgDict)) {
      if (k !== "standard_tier_tables" && k !== "customer_tier_tables") {
        paramRows.push({
          參數名稱: k,
          設定值: typeof v === "object" && v !== null ? JSON.stringify(v) : String(v),
        });
      }
    }
    if (cfgDict.customer_type !== "A" && cfgDict.standard_tier_tables) {
      paramRows.push({
        參數名稱: `${cfgDict.customer_type}客戶適用階梯折數表 (tier_tables)`,
        設定值: JSON.stringify(cfgDict.standard_tier_tables),
      });
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

    const fileName = `永源化工_配櫃結果_${cfg.customer_type}客戶_${compC}櫃_${curr}_${Math.round(evalRes.WMT_kg || 0)}kg.xlsx`;
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
    solveMultipleContainers,
    buildSelectedBatchSummary,
    buildSelectedBagDetail,
    exportBlendingResultToExcel,
  };
})(typeof window !== "undefined" ? window : globalThis);

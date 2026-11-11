/**
 * 永源化工 — 自動配櫃與最佳化計價系統 (Bootstrap 5 RWD 互動介面控制器 app.js)
 * 整合：
 * 1. 內建三大實務範例一鍵套用範本與自由參數微調
 * 2. 支援美元 (USD) 與日圓 (JPY) 雙幣別計價切換與單位標籤自動連動
 * 3. 支援瀏覽器本地上傳/切換《庫存明細表.xlsx》，可搜尋式多選器動態勾選「必出批次/進料代號/單包」、「禁收編碼」與「排除貴金屬料號」
 * 4. 支援於 UI 直接編輯「配滿櫃重量區間」、「A客戶門檻/RC/水分罰款」、「B~E客戶階梯折數表」，並可儲存至 localStorage 或匯出 JSON
 * 5. 完整呈現配櫃試算結果（6張 KPI 卡、計價分解、批次小計、逐包明細）與一鍵匯出多工作表 Excel (.xlsx)
 */

(function () {
  "use strict";

  const STORAGE_KEY_CONFIG = "yy_auto_blending_saved_config_v1";
  const REQUIRED_METALS = window.BlendingConfig.REQUIRED_METALS;

  // =========================================================================
  // 全域應用程式狀態 (App State)
  // =========================================================================
  const appState = {
    configData: null,
    invRows: [],
    batchOverview: [],
    invLabel: "預設檔案：庫存明細表.xlsx",
    activePresetKey: "preset1",
    activePresetLabel:
      "實務範例一：E 客戶配櫃 (USD 計價 | 必出 1F1+3P1、銅>=5%、貴金屬排除料號 X)",
    tierTables: {
      Cu: [],
      Au: [],
      Ag: [],
      Pd: [],
    },
    pickers: {},
    lastResult: null,
    lastSelectedRows: [],
    lastBagDetailRows: [],
    lastActiveConfig: null,
  };

  // =========================================================================
  // 格式化輔助工具
  // =========================================================================
  function fmtInt(n) {
    return Math.round(Number(n) || 0).toLocaleString("en-US");
  }

  function fmtNum(n, decimals = 2) {
    return (Number(n) || 0).toLocaleString("en-US", {
      minimumFractionDigits: decimals,
      maximumFractionDigits: decimals,
    });
  }

  function escapeHtml(str) {
    return String(str ?? "")
      .replace(/&/g, "&amp;")
      .replace(/</g, "&lt;")
      .replace(/>/g, "&gt;")
      .replace(/"/g, "&quot;");
  }

  // =========================================================================
  // 可搜尋式多選器元件 (Searchable Multi-Select Tag Picker)
  // =========================================================================
  class MultiSelectPicker {
    constructor(containerId, placeholder, onChange, isIntVal = false) {
      this.container = document.getElementById(containerId);
      this.placeholder = placeholder;
      this.onChange = onChange;
      this.isIntVal = isIntVal;
      this.options = []; // [{ value, label, shortTag }]
      this.selectedSet = new Set();
      this.searchQuery = "";
      this.renderSkeleton();
    }

    renderSkeleton() {
      if (!this.container) return;
      this.container.innerHTML = `
        <div class="ms-picker">
          <div class="ms-tags-area"></div>
          <div class="d-flex gap-1 mb-1">
            <input type="search" class="form-control form-control-sm py-0 px-2 ms-search-input" style="font-size:0.8rem;height:26px;" placeholder="${escapeHtml(this.placeholder)}" />
            <button type="button" class="btn btn-outline-secondary btn-sm py-0 px-2 ms-clear-btn" style="font-size:0.74rem;height:26px;white-space:nowrap;" title="清空已選">清空</button>
          </div>
          <div class="ms-options-list"></div>
        </div>
      `;
      this.tagsArea = this.container.querySelector(".ms-tags-area");
      this.searchInput = this.container.querySelector(".ms-search-input");
      this.clearBtn = this.container.querySelector(".ms-clear-btn");
      this.optionsList = this.container.querySelector(".ms-options-list");

      this.searchInput.addEventListener("input", () => {
        this.searchQuery = this.searchInput.value.trim().toLowerCase();
        this.renderOptions();
      });

      this.clearBtn.addEventListener("click", () => {
        if (this.selectedSet.size > 0) {
          this.selectedSet.clear();
          this.renderTags();
          this.renderOptions();
          if (this.onChange) this.onChange(this.getValues());
        }
      });
    }

    setOptions(options) {
      this.options = options;
      const validValues = new Set(options.map((o) => o.value));
      for (const v of Array.from(this.selectedSet)) {
        if (!validValues.has(v)) {
          this.selectedSet.delete(v);
        }
      }
      this.renderTags();
      this.renderOptions();
    }

    setValues(vals) {
      const validValues = new Set(this.options.map((o) => o.value));
      this.selectedSet.clear();
      for (const rawV of vals || []) {
        const v = this.isIntVal ? parseInt(rawV, 10) : String(rawV).trim();
        if (validValues.size === 0 || validValues.has(v)) {
          this.selectedSet.add(v);
        }
      }
      this.renderTags();
      this.renderOptions();
    }

    getValues() {
      return Array.from(this.selectedSet);
    }

    renderTags() {
      if (!this.tagsArea) return;
      if (this.selectedSet.size === 0) {
        this.tagsArea.innerHTML = `<span class="text-muted small px-1">尚未選擇任何項目</span>`;
        return;
      }
      const optMap = new Map(this.options.map((o) => [o.value, o]));
      const html = [];
      for (const v of this.selectedSet) {
        const opt = optMap.get(v);
        const tagText = opt ? opt.shortTag || opt.value : String(v);
        html.push(`
          <span class="ms-tag">
            <span>${escapeHtml(tagText)}</span>
            <button type="button" class="ms-tag-remove" data-val="${escapeHtml(v)}" title="移除">&times;</button>
          </span>
        `);
      }
      this.tagsArea.innerHTML = html.join("");
      this.tagsArea.querySelectorAll(".ms-tag-remove").forEach((btn) => {
        btn.addEventListener("click", (e) => {
          e.stopPropagation();
          const rawV = btn.getAttribute("data-val");
          const v = this.isIntVal ? parseInt(rawV, 10) : rawV;
          this.selectedSet.delete(v);
          this.renderTags();
          this.renderOptions();
          if (this.onChange) this.onChange(this.getValues());
        });
      });
    }

    renderOptions() {
      if (!this.optionsList) return;
      const q = this.searchQuery;
      const filtered = q
        ? this.options.filter(
            (o) =>
              String(o.label).toLowerCase().includes(q) ||
              String(o.value).toLowerCase().includes(q)
          )
        : this.options;

      if (filtered.length === 0) {
        this.optionsList.innerHTML = `<div class="text-muted small px-2 py-1">無符合「${escapeHtml(q)}」之選項</div>`;
        return;
      }

      // 限制單次 DOM 渲染上限以確保千包列表極速流暢
      const maxRender = 250;
      const slice = filtered.slice(0, maxRender);
      const html = slice.map((o) => {
        const checked = this.selectedSet.has(o.value) ? "checked" : "";
        return `
          <label class="ms-option-item">
            <input type="checkbox" class="form-check-input mt-0" value="${escapeHtml(o.value)}" ${checked} />
            <span class="text-truncate" title="${escapeHtml(o.label)}">${escapeHtml(o.label)}</span>
          </label>
        `;
      });
      if (filtered.length > maxRender) {
        html.push(
          `<div class="text-muted small px-2 py-1">... 尚有 ${filtered.length - maxRender} 項，請輸入關鍵字篩選</div>`
        );
      }
      this.optionsList.innerHTML = html.join("");

      this.optionsList.querySelectorAll('input[type="checkbox"]').forEach((chk) => {
        chk.addEventListener("change", () => {
          const v = this.isIntVal ? parseInt(chk.value, 10) : chk.value;
          if (chk.checked) {
            this.selectedSet.add(v);
          } else {
            this.selectedSet.delete(v);
          }
          this.renderTags();
          if (this.onChange) this.onChange(this.getValues());
        });
      });
    }
  }

  // =========================================================================
  // 初始化多選器與庫存選項清單
  // =========================================================================
  function initPickers() {
    const onPickerChange = () => {
      updateLiveFilterStats();
      triggerAutoSolveIfEnabled();
    };

    appState.pickers.bannedCodes = new MultiSelectPicker(
      "pickerBannedCodes",
      "搜尋編碼料號...",
      onPickerChange,
      false
    );
    appState.pickers.excludePmMaterials = new MultiSelectPicker(
      "pickerExcludePmMaterials",
      "搜尋總表料號...",
      onPickerChange,
      false
    );
    appState.pickers.excludePmCodes = new MultiSelectPicker(
      "pickerExcludePmCodes",
      "搜尋編碼料號...",
      onPickerChange,
      false
    );
    appState.pickers.mustBatchIds = new MultiSelectPicker(
      "pickerMustBatchIds",
      "搜尋批號、進料代號或銅品位...",
      onPickerChange,
      false
    );
    appState.pickers.mustFeedCodes = new MultiSelectPicker(
      "pickerMustFeedCodes",
      "搜尋進料公司代號...",
      onPickerChange,
      false
    );
    appState.pickers.mustBagIds = new MultiSelectPicker(
      "pickerMustBagIds",
      "搜尋單包代號 (1~1330)...",
      onPickerChange,
      true
    );
  }

  function populatePickerOptionsFromInventory() {
    const inv = appState.invRows;
    const batches = appState.batchOverview;

    // 1. 編碼料號選項
    const codeStats = new Map();
    for (const r of inv) {
      const c = r["編碼料號"];
      if (!c) continue;
      if (!codeStats.has(c)) {
        codeStats.set(c, { batches: new Set(), bags: 0, wmt: 0 });
      }
      const st = codeStats.get(c);
      st.batches.add(r["Batch_ID"]);
      st.bags += 1;
      st.wmt += r["倉庫重"];
    }
    const sortedCodes = Array.from(codeStats.keys()).sort();
    const codeOptions = sortedCodes.map((c) => {
      const st = codeStats.get(c);
      return {
        value: c,
        shortTag: `編碼 ${c}`,
        label: `編碼 ${c} (${st.batches.size}批 / ${st.bags}包 / ${fmtInt(st.wmt)} kg)`,
      };
    });
    appState.pickers.bannedCodes.setOptions(codeOptions);
    appState.pickers.excludePmCodes.setOptions(codeOptions);

    // 2. 總表料號選項
    const matStats = new Map();
    for (const r of inv) {
      const m = r["總表料號"];
      if (!m) continue;
      if (!matStats.has(m)) {
        matStats.set(m, { batches: new Set(), wmt: 0 });
      }
      const st = matStats.get(m);
      st.batches.add(r["Batch_ID"]);
      st.wmt += r["倉庫重"];
    }
    const sortedMats = Array.from(matStats.keys()).sort();
    const matOptions = sortedMats.map((m) => {
      const st = matStats.get(m);
      return {
        value: m,
        shortTag: `料號 ${m}`,
        label: `總表料號 ${m} (${st.batches.size}批 / ${fmtInt(st.wmt)} kg)`,
      };
    });
    appState.pickers.excludePmMaterials.setOptions(matOptions);

    // 3. 批號選項
    const batchOptions = batches.map((b) => ({
      value: b["批號"],
      shortTag: `${b["批號"]}[${b["進料代號"]}]`,
      label: b["批次標籤"],
    }));
    appState.pickers.mustBatchIds.setOptions(batchOptions);

    // 4. 進料代號選項
    const feedStats = new Map();
    for (const r of inv) {
      const f = r["進料公司名稱(代號)"];
      if (!f) continue;
      if (!feedStats.has(f)) {
        feedStats.set(f, { bags: 0, wmt: 0 });
      }
      const st = feedStats.get(f);
      st.bags += 1;
      st.wmt += r["倉庫重"];
    }
    const sortedFeeds = Array.from(feedStats.keys()).sort();
    const feedOptions = sortedFeeds.map((f) => {
      const st = feedStats.get(f);
      return {
        value: f,
        shortTag: f,
        label: `${f} (${st.bags}包 / ${fmtInt(st.wmt)} kg)`,
      };
    });
    appState.pickers.mustFeedCodes.setOptions(feedOptions);

    // 5. 單包代號選項
    const bagOptions = inv.map((r) => ({
      value: Number(r["代號"]),
      shortTag: `#${r["代號"]}`,
      label: `#${r["代號"]} | ${r["Batch_ID"]} (${r["進料公司名稱(代號)"]}) | ${fmtInt(r["倉庫重"])} kg | 濕Cu:${Number(r["濕基Cu(%)"]).toFixed(2)}%`,
    }));
    appState.pickers.mustBagIds.setOptions(bagOptions);
  }

  // =========================================================================
  // 階梯折數表動態編輯器 (Tier Tables Editor)
  // =========================================================================
  function renderAllTierEditors() {
    for (const m of REQUIRED_METALS) {
      renderSingleTierEditor(m);
    }
  }

  function renderSingleTierEditor(metal) {
    const container = document.getElementById(`tierEditor${metal}`);
    if (!container) return;
    const rows = appState.tierTables[metal] || [];

    const tbodyHtml = rows
      .map(
        (r, idx) => `
      <tr>
        <td>
          <input type="number" class="form-control form-control-sm tier-input" data-metal="${metal}" data-idx="${idx}" data-field="min_grade" min="0" step="0.5" value="${Number(r.min_grade)}" />
        </td>
        <td>
          <input type="number" class="form-control form-control-sm tier-input" data-metal="${metal}" data-idx="${idx}" data-field="max_grade" min="0" step="0.5" value="${Number(r.max_grade)}" />
        </td>
        <td>
          <input type="number" class="form-control form-control-sm tier-input" data-metal="${metal}" data-idx="${idx}" data-field="payable_rate" min="0" max="1" step="0.05" value="${Number(r.payable_rate)}" />
        </td>
        <td class="text-center">
          <button type="button" class="btn btn-outline-danger btn-sm py-0 px-2 tier-del-btn" data-metal="${metal}" data-idx="${idx}" title="刪除此階梯" ${rows.length <= 1 ? "disabled" : ""}>
            <i class="bi bi-trash"></i>
          </button>
        </td>
      </tr>
    `
      )
      .join("");

    container.innerHTML = `
      <div class="table-responsive">
        <table class="table table-sm table-bordered bg-white mb-0 tier-editor-table align-middle">
          <thead class="table-light small">
            <tr>
              <th>品位下限 (min_grade)</th>
              <th>品位上限 (max_grade)</th>
              <th>計價折數 (payable_rate)</th>
              <th class="text-center" style="width:52px;">操作</th>
            </tr>
          </thead>
          <tbody>
            ${tbodyHtml}
          </tbody>
        </table>
      </div>
    `;

    container.querySelectorAll(".tier-input").forEach((inp) => {
      inp.addEventListener("change", () => {
        const m = inp.getAttribute("data-metal");
        const i = parseInt(inp.getAttribute("data-idx"), 10);
        const f = inp.getAttribute("data-field");
        const val = parseFloat(inp.value);
        if (appState.tierTables[m] && appState.tierTables[m][i]) {
          appState.tierTables[m][i][f] = Number.isFinite(val) ? val : 0.0;
          triggerAutoSolveIfEnabled();
        }
      });
    });

    container.querySelectorAll(".tier-del-btn").forEach((btn) => {
      btn.addEventListener("click", () => {
        const m = btn.getAttribute("data-metal");
        const i = parseInt(btn.getAttribute("data-idx"), 10);
        if (appState.tierTables[m] && appState.tierTables[m].length > 1) {
          appState.tierTables[m].splice(i, 1);
          renderSingleTierEditor(m);
          triggerAutoSolveIfEnabled();
        }
      });
    });
  }

  // =========================================================================
  // 將指定之參數與全域預設值寫入 UI 表單控制項
  // =========================================================================
  function applyParamsToUi(params, globalDefaults, presetLabel, presetKey = "custom") {
    const merged = Object.assign({}, globalDefaults || {}, params || {});

    appState.activePresetKey = presetKey;
    appState.activePresetLabel = presetLabel || "自訂配櫃設定";

    document.getElementById("activePresetText").textContent = appState.activePresetLabel;
    updatePresetButtonsHighlight(presetKey);

    document.getElementById("selectCustomerType").value = String(merged.customer_type || "E");
    const curr = String(merged.price_currency || "USD").toUpperCase();
    document.getElementById("selectPriceCurrency").value = curr;

    const fRange = merged.full_container_range || [20000.0, 20500.0];
    document.getElementById("inputFullMinKg").value = Number(fRange[0]);
    document.getElementById("inputFullMaxKg").value = Number(fRange[1]);

    const cuRange = merged.target_cu_range || [0.0, 100.0];
    document.getElementById("inputTargetCuMin").value = Number(cuRange[0]);
    document.getElementById("inputTargetCuMax").value = Number(cuRange[1]);

    document.getElementById("inputMaxDiffCu").value = Number(
      merged.max_in_container_diff_cu ?? 100.0
    );

    const maxVal = merged.max_total_value;
    const enableMaxVal = maxVal !== null && maxVal !== undefined && maxVal !== Infinity;
    document.getElementById("checkEnableMaxTotalValue").checked = enableMaxVal;
    document.getElementById("inputMaxTotalValue").value = enableMaxVal ? Number(maxVal) : 200000.0;
    document.getElementById("maxTotalValueWrapper").classList.toggle("d-none", !enableMaxVal);

    document.getElementById("inputSplitBatchPenalty").value = Number(
      merged.split_batch_penalty ?? 500.0
    );

    // 多選器設定
    appState.pickers.bannedCodes.setValues(merged.banned_codes || []);
    appState.pickers.mustBatchIds.setValues(merged.must_batch_ids || []);
    appState.pickers.mustFeedCodes.setValues(merged.must_feed_codes || []);
    appState.pickers.mustBagIds.setValues(merged.must_bag_ids || []);
    appState.pickers.excludePmCodes.setValues(merged.exclude_pm_codes || []);
    appState.pickers.excludePmMaterials.setValues(merged.exclude_pm_materials || []);

    // 金屬價格
    document.getElementById("inputTroyOzToG").value = Number(merged.troy_oz_to_g ?? 31.1035);
    const mp = merged.metal_prices || {};
    const mpJpy = merged.metal_prices_jpy || {
      Cu: 1501000.0,
      Au: 13462.0,
      Ag: 157.474,
      Pd: 5334.0,
    };
    document.getElementById("inputPriceUsdCu").value = Number(mp.Cu ?? 0);
    document.getElementById("inputPriceUsdAu").value = Number(mp.Au ?? 0);
    document.getElementById("inputPriceUsdAg").value = Number(mp.Ag ?? 0);
    document.getElementById("inputPriceUsdPd").value = Number(mp.Pd ?? 0);

    document.getElementById("inputPriceJpyCu").value = Number(mpJpy.Cu ?? 0);
    document.getElementById("inputPriceJpyAu").value = Number(mpJpy.Au ?? 0);
    document.getElementById("inputPriceJpyAg").value = Number(mpJpy.Ag ?? 0);
    document.getElementById("inputPriceJpyPd").value = Number(mpJpy.Pd ?? 0);

    // A 客戶參數
    document.getElementById("selectAMode").value = String(merged.a_mode || "with_pm");
    const aTh = merged.a_thresholds || {};
    const aRc = merged.a_rc || {};
    document.getElementById("inputAThCu").value = Number(aTh.Cu ?? 0);
    document.getElementById("inputAThAu").value = Number(aTh.Au ?? 0);
    document.getElementById("inputAThAg").value = Number(aTh.Ag ?? 0);
    document.getElementById("inputAThPd").value = Number(aTh.Pd ?? 0);

    document.getElementById("inputARcCu").value = Number(aRc.Cu ?? 0);
    document.getElementById("inputARcAu").value = Number(aRc.Au ?? 0);
    document.getElementById("inputARcAg").value = Number(aRc.Ag ?? 0);
    document.getElementById("inputARcPd").value = Number(aRc.Pd ?? 0);

    document.getElementById("inputAFreightPerDmt").value = Number(merged.a_freight_per_dmt ?? 0);
    document.getElementById("inputAMoistureLimit").value = Number(merged.a_moisture_limit ?? 25.0);
    document.getElementById("inputAMoisturePenaltyPerWmt").value = Number(
      merged.a_moisture_penalty_per_wmt ?? 0
    );
    document.getElementById("inputACuBelowTcPerWmt").value = Number(
      merged.a_cu_below_threshold_tc_per_wmt ?? 0
    );
    document.getElementById("inputAAssayFeeIfBelow").value = Number(
      merged.a_assay_fee_if_below ?? 0
    );

    // B ~ E 客戶參數與階梯折數表
    document.getElementById("inputBFreightPerKg").value = Number(merged.b_freight_per_kg ?? 0);
    document.getElementById("inputBDeductionPerTon").value = Number(
      merged.b_deduction_per_ton ?? 0
    );

    const rawTiers = merged.tier_tables || merged.standard_tier_tables || {};
    for (const m of REQUIRED_METALS) {
      const tList = rawTiers[m] || [];
      appState.tierTables[m] = tList.map((r) => ({
        min_grade: Number(r.min_grade),
        max_grade: Number(r.max_grade),
        payable_rate: Number(r.payable_rate),
      }));
    }
    renderAllTierEditors();
    updateCurrencyLabelsInUi();
    updateLiveFilterStats();
  }

  function updatePresetButtonsHighlight(activeKey) {
    const map = {
      preset1: "btnPreset1",
      preset2: "btnPreset2",
      preset3: "btnPreset3",
      presetDefault: "btnPresetDefault",
    };
    Object.values(map).forEach((id) => {
      const el = document.getElementById(id);
      if (el) el.classList.remove("active");
    });
    if (map[activeKey]) {
      const activeEl = document.getElementById(map[activeKey]);
      if (activeEl) activeEl.classList.add("active");
    }
  }

  function updateCurrencyLabelsInUi() {
    const curr = document.getElementById("selectPriceCurrency").value || "USD";
    const rcCuUnit = curr === "JPY" ? "¥ / T" : "USD / kg";
    const rcPmUnit = curr === "JPY" ? "¥ / g" : "USD / oz";

    document.getElementById("labelSplitPenaltyUnit").textContent = `${curr}/拆批`;
    document.getElementById("labelMaxTotalValueUnit").textContent = curr;
    document.getElementById("badgeActiveCurrencyTab2").textContent = `目前生效計價幣別：${curr}`;
    document.getElementById("badgeActiveCurrencyTab3").textContent = `目前計價幣別：${curr}`;

    document.getElementById("labelARcCuUnit").textContent = rcCuUnit;
    document.getElementById("labelARcAuUnit").textContent = rcPmUnit;
    document.getElementById("labelARcAgUnit").textContent = rcPmUnit;
    document.getElementById("labelARcPdUnit").textContent = rcPmUnit;

    document.getElementById("labelAFreightUnit").textContent = `${curr}/DMT`;
    document.getElementById("labelAMoistPenUnit").textContent = `${curr}/WMT`;
    document.getElementById("labelACuTcUnit").textContent = `${curr}/WMT`;
    document.getElementById("labelAAssayFeeUnit").textContent = curr;
  }

  // =========================================================================
  // 從 UI 表單收集所有設定參數
  // =========================================================================
  function collectParamsFromUi() {
    const fullMin = parseFloat(document.getElementById("inputFullMinKg").value);
    const fullMax = parseFloat(document.getElementById("inputFullMaxKg").value);
    const enableMaxVal = document.getElementById("checkEnableMaxTotalValue").checked;

    const tierDict = {};
    for (const m of REQUIRED_METALS) {
      tierDict[m] = (appState.tierTables[m] || []).map((r) => ({
        min_grade: Number(r.min_grade),
        max_grade: Number(r.max_grade),
        payable_rate: Number(r.payable_rate),
      }));
    }

    return {
      customer_type: document.getElementById("selectCustomerType").value,
      price_currency: document.getElementById("selectPriceCurrency").value,
      fill_container: true,
      full_container_range: [fullMin, fullMax],
      w_min_kg: fullMin,
      w_max_kg: fullMax,
      target_cu_range: [
        parseFloat(document.getElementById("inputTargetCuMin").value),
        parseFloat(document.getElementById("inputTargetCuMax").value),
      ],
      max_in_container_diff_cu: parseFloat(document.getElementById("inputMaxDiffCu").value),
      max_total_value: enableMaxVal
        ? parseFloat(document.getElementById("inputMaxTotalValue").value)
        : null,
      banned_codes: appState.pickers.bannedCodes.getValues(),
      must_batch_ids: appState.pickers.mustBatchIds.getValues(),
      must_feed_codes: appState.pickers.mustFeedCodes.getValues(),
      must_bag_ids: appState.pickers.mustBagIds.getValues(),
      exclude_pm_codes: appState.pickers.excludePmCodes.getValues(),
      exclude_pm_materials: appState.pickers.excludePmMaterials.getValues(),
      split_batch_penalty: parseFloat(document.getElementById("inputSplitBatchPenalty").value),
      troy_oz_to_g: parseFloat(document.getElementById("inputTroyOzToG").value),
      metal_prices: {
        Cu: parseFloat(document.getElementById("inputPriceUsdCu").value),
        Au: parseFloat(document.getElementById("inputPriceUsdAu").value),
        Ag: parseFloat(document.getElementById("inputPriceUsdAg").value),
        Pd: parseFloat(document.getElementById("inputPriceUsdPd").value),
      },
      metal_prices_jpy: {
        Cu: parseFloat(document.getElementById("inputPriceJpyCu").value),
        Au: parseFloat(document.getElementById("inputPriceJpyAu").value),
        Ag: parseFloat(document.getElementById("inputPriceJpyAg").value),
        Pd: parseFloat(document.getElementById("inputPriceJpyPd").value),
      },
      a_mode: document.getElementById("selectAMode").value,
      a_thresholds: {
        Cu: parseFloat(document.getElementById("inputAThCu").value),
        Au: parseFloat(document.getElementById("inputAThAu").value),
        Ag: parseFloat(document.getElementById("inputAThAg").value),
        Pd: parseFloat(document.getElementById("inputAThPd").value),
      },
      a_rc: {
        Cu: parseFloat(document.getElementById("inputARcCu").value),
        Au: parseFloat(document.getElementById("inputARcAu").value),
        Ag: parseFloat(document.getElementById("inputARcAg").value),
        Pd: parseFloat(document.getElementById("inputARcPd").value),
      },
      a_freight_per_dmt: parseFloat(document.getElementById("inputAFreightPerDmt").value),
      a_moisture_limit: parseFloat(document.getElementById("inputAMoistureLimit").value),
      a_moisture_penalty_per_wmt: parseFloat(
        document.getElementById("inputAMoisturePenaltyPerWmt").value
      ),
      a_cu_below_threshold_tc_per_wmt: parseFloat(
        document.getElementById("inputACuBelowTcPerWmt").value
      ),
      a_assay_fee_if_below: parseFloat(document.getElementById("inputAAssayFeeIfBelow").value),
      b_freight_per_kg: parseFloat(document.getElementById("inputBFreightPerKg").value),
      b_deduction_per_ton: parseFloat(document.getElementById("inputBDeductionPerTon").value),
      standard_tier_tables: tierDict,
    };
  }

  // =========================================================================
  // 即時篩選統計與庫存總覽渲染
  // =========================================================================
  function updateLiveFilterStats() {
    const inv = appState.invRows;
    if (!inv || inv.length === 0) return;

    const mustBatchSet = new Set(appState.pickers.mustBatchIds.getValues());
    const mustFeedSet = new Set(appState.pickers.mustFeedCodes.getValues());
    const mustBagSet = new Set(appState.pickers.mustBagIds.getValues().map((x) => Number(x)));
    const bannedCodeSet = new Set(appState.pickers.bannedCodes.getValues());

    let mustBagCount = 0;
    let mustWmt = 0.0;
    let availBagCount = 0;
    let availWmt = 0.0;
    const availBatchSet = new Set();

    for (const r of inv) {
      const isMust =
        mustBatchSet.has(r["Batch_ID"]) ||
        mustFeedSet.has(r["進料公司名稱(代號)"]) ||
        mustBagSet.has(Number(r["代號"]));
      if (isMust) {
        mustBagCount += 1;
        mustWmt += r["倉庫重"];
      }
      if (isMust || (r["已化驗"] && !bannedCodeSet.has(r["編碼料號"]))) {
        availBagCount += 1;
        availWmt += r["倉庫重"];
        availBatchSet.add(r["Batch_ID"]);
      }
    }

    const statsEl = document.getElementById("liveFilterStatsBar");
    if (statsEl) {
      statsEl.innerHTML =
        `📊 <strong>即時篩選統計</strong>：已鎖定必出料共 <strong>${fmtInt(mustBagCount)} 包 (${fmtInt(mustWmt)} kg)</strong> ｜ ` +
        `排除禁收與未化驗後可用庫存共 <strong>${availBatchSet.size} 批 / ${fmtInt(availBagCount)} 包 (${fmtInt(availWmt)} kg)</strong>`;
    }
  }

  function renderSidebarInventoryStatus() {
    const inv = appState.invRows;
    const batches = appState.batchOverview;
    const assayedCount = batches.filter((b) => b["已化驗"]).length;
    const totalWmt = inv.reduce((acc, r) => acc + r["倉庫重"], 0.0);

    const card = document.getElementById("invStatusCard");
    if (card) {
      card.className = "alert alert-success py-2 px-3 small mb-3";
      card.innerHTML = `
        <div class="fw-bold mb-1">✅ 已載入 ${escapeHtml(appState.invLabel)}</div>
        <ul class="mb-0 ps-3">
          <li>總包數：<strong>${fmtInt(inv.length)} 包</strong></li>
          <li>總批數：<strong>${batches.length} 批</strong> (已化驗 ${assayedCount} 批)</li>
          <li>總倉庫濕重：<strong>${fmtInt(totalWmt)} kg</strong></li>
        </ul>
      `;
    }
  }

  function renderInventoryOverviewTable() {
    const tbody = document.getElementById("invOverviewTableBody");
    const countBadge = document.getElementById("invOverviewCountBadge");
    if (!tbody) return;

    const q = (document.getElementById("invSearchInput")?.value || "").trim().toLowerCase();
    const assayFilter = document.getElementById("invAssayFilter")?.value || "all";

    const filtered = appState.batchOverview.filter((b) => {
      if (assayFilter === "assayed" && !b["已化驗"]) return false;
      if (assayFilter === "unassayed" && b["已化驗"]) return false;
      if (!q) return true;
      return (
        String(b["批號"]).toLowerCase().includes(q) ||
        String(b["進料代號"]).toLowerCase().includes(q) ||
        String(b["總表料號"]).toLowerCase().includes(q) ||
        String(b["編碼料號"]).toLowerCase().includes(q) ||
        String(b["存放位置"]).toLowerCase().includes(q)
      );
    });

    if (countBadge) {
      countBadge.textContent = `${filtered.length} / ${appState.batchOverview.length} 批`;
    }

    if (filtered.length === 0) {
      tbody.innerHTML = `<tr><td colspan="19" class="text-center text-muted py-4">查無符合篩選條件之庫存批次</td></tr>`;
      return;
    }

    tbody.innerHTML = filtered
      .map(
        (b) => `
      <tr class="${b["已化驗"] ? "" : "table-warning"}">
        <td class="sticky-col fw-bold">${escapeHtml(b["批號"])}</td>
        <td>${escapeHtml(b["廠別"])}</td>
        <td><span class="badge bg-light text-dark border">${escapeHtml(b["進料代號"])}</span></td>
        <td>${escapeHtml(b["總表料號"])}</td>
        <td><span class="badge bg-secondary">${escapeHtml(b["編碼料號"])}</span></td>
        <td>${escapeHtml(b["存放位置"])}</td>
        <td class="text-end">${fmtInt(b["總包數"])}</td>
        <td class="text-end fw-semibold">${fmtNum(b["倉庫濕重(kg)"], 1)}</td>
        <td class="text-end">${fmtNum(b["總乾重(kg)"], 2)}</td>
        <td class="text-end">${fmtNum(b["含水率(%)"], 2)}%</td>
        <td class="text-end">${fmtNum(b["乾基Cu(%)"], 2)}%</td>
        <td class="text-end fw-semibold text-primary">${fmtNum(b["濕基Cu(%)"], 2)}%</td>
        <td class="text-end">${fmtNum(b["乾基Au(g/T)"], 2)}</td>
        <td class="text-end">${fmtNum(b["溼基Au(g/T)"], 2)}</td>
        <td class="text-end">${fmtNum(b["乾基Ag(g/T)"], 2)}</td>
        <td class="text-end">${fmtNum(b["溼基Ag(g/T)"], 2)}</td>
        <td class="text-end">${fmtNum(b["乾基Pd(g/T)"], 2)}</td>
        <td class="text-end">${fmtNum(b["溼基Pd(g/T)"], 2)}</td>
        <td class="text-center">${
          b["已化驗"]
            ? '<span class="badge bg-success">已化驗</span>'
            : '<span class="badge bg-warning text-dark">未化驗</span>'
        }</td>
      </tr>
    `
      )
      .join("");
  }

  // =========================================================================
  // 執行最佳化配櫃求解與結果渲染
  // =========================================================================
  function triggerAutoSolveIfEnabled() {
    const autoCheck = document.getElementById("checkAutoRun");
    if (autoCheck && autoCheck.checked) {
      runOptimizationSolve();
    }
  }

  function runOptimizationSolve() {
    const errAlert = document.getElementById("validationErrorAlert");
    const resultCard = document.getElementById("resultSectionCard");
    if (errAlert) errAlert.classList.add("d-none");

    let cfgActive;
    try {
      const uiParams = collectParamsFromUi();
      cfgActive = window.BlendingConfig.CustomerConfig.fromDict(
        uiParams,
        appState.configData.global_defaults,
        "UI_CurrentConfig"
      );
    } catch (err) {
      if (errAlert) {
        errAlert.innerHTML = `❌ <strong>參數設定驗證未通過</strong>：${escapeHtml(err.message)}`;
        errAlert.classList.remove("d-none");
      }
      if (resultCard) resultCard.classList.add("d-none");
      return;
    }

    if (resultCard) resultCard.classList.remove("d-none");

    const t0 = performance.now();
    const [selectedRows, res] = window.BlendingEngine.solveOptimalContainer(
      appState.invRows,
      cfgActive
    );
    const elapsedMs = performance.now() - t0;

    appState.lastSelectedRows = selectedRows;
    appState.lastResult = res;
    appState.lastActiveConfig = cfgActive;

    const currCode = cfgActive.price_currency;
    const currSym = currCode === "JPY" ? "¥" : "$";
    const [wMin, wMax] = cfgActive.full_container_range;

    document.getElementById("resultHeaderTitle").textContent =
      `📊 配滿櫃最佳化試算結果 — ${cfgActive.customer_type} 客戶 ｜ 計價幣別：${currCode} (${currSym}) ｜ ` +
      `重量區間：${fmtInt(wMin)} ~ ${fmtInt(wMax)} kg`;

    // 同步更新分頁 3 標題幣別
    document.getElementById("res-tab3-btn").textContent =
      `🧮 (3) 客戶計價公式與分項金額試算分解 (${currCode})`;

    const infeasibleBox = document.getElementById("infeasibleAlertBox");
    const feasibleBox = document.getElementById("feasibleResultContainer");

    if (!res.feasible) {
      feasibleBox.classList.add("d-none");
      infeasibleBox.classList.remove("d-none");
      infeasibleBox.innerHTML = `
        <div class="alert alert-danger mb-2">
          ⚠️ <strong>無法找到合規配滿櫃組合</strong>：${escapeHtml(res.reason)}
        </div>
        <div class="p-3 border rounded-3 bg-light small">
          <ul class="mb-0">
            <li><strong>可用合格庫存總濕重</strong>：<code>${fmtInt(res.max_avail_wmt_kg || 0)} kg</code></li>
            <li><strong>已鎖定必出料總濕重</strong>：<code>${fmtInt(res.must_wmt_kg || 0)} kg</code></li>
            <li><strong>建議調整方向</strong>：請放寬「同櫃最大濕基銅含量落差上限 (<code>max_in_container_diff_cu</code>)」、「目標銅品位區間」或減少「禁收編碼 / 必出批次」。</li>
          </ul>
        </div>
      `;
      updateMobileStickyBar(cfgActive, res);
      return;
    }

    infeasibleBox.classList.add("d-none");
    feasibleBox.classList.remove("d-none");

    // 1. 更新 6 張 KPI 指標卡
    document.getElementById("kpiBagCount").textContent = `${res.bag_count} 包`;
    document.getElementById("kpiSplitCount").textContent = `拆批數: ${res.split_batches.length} 批`;

    document.getElementById("kpiWmtKg").textContent = `${fmtInt(res.WMT_kg)} kg`;
    document.getElementById("kpiWmtTon").textContent = `${(res.WMT_kg / 1000.0).toFixed(3)} 噸 (配滿櫃)`;

    document.getElementById("kpiDmtKg").textContent = `${fmtNum(res.DMT_kg, 2)} kg`;
    document.getElementById("kpiMoisturePct").textContent = `含水率 ${res.moisture_pct.toFixed(2)}%`;

    document.getElementById("kpiWetCu").textContent = `${res.wet_grades["Cu(%)"].toFixed(4)}%`;
    document.getElementById("kpiCuDiff").textContent = `同櫃落差 ${res.cu_diff.toFixed(2)}%`;

    document.getElementById("kpiNetProfitLabel").textContent = `預估結算淨利 (${currCode})`;
    document.getElementById("kpiNetProfit").textContent = `${currSym}${fmtNum(res.net_profit, 2)}`;
    document.getElementById("kpiGrossRev").textContent = `毛收 ${currSym}${fmtNum(res.gross_revenue, 2)}`;

    document.getElementById("kpiObjScore").textContent = fmtNum(res.objective_score, 2);
    document.getElementById("kpiElapsedMs").textContent = `耗時 ${(elapsedMs / 1000.0).toFixed(3)} 秒 (${Math.round(elapsedMs)} ms)`;

    // 2. 更新選中批次組合摘要橫幅
    const fullText = res.full_batches.length > 0 ? res.full_batches.join(", ") : "無";
    const splitText = res.split_batches.length > 0 ? res.split_batches.join(", ") : "無 (0拆批)";
    document.getElementById("selectedBatchesBanner").innerHTML = `
      <div class="fw-bold mb-1">🎯 選中批次組合：</div>
      <ul class="mb-0 ps-3">
        <li><strong>整批全出 (${res.full_batches.length} 批)</strong>：${escapeHtml(fullText)}</li>
        <li><strong>拆批湊櫃 (${res.split_batches.length} 批)</strong>：${escapeHtml(splitText)}</li>
      </ul>
    `;

    // 3. 渲染三個結果明細分頁
    renderSelectedBatchSummaryTable(selectedRows);
    appState.lastBagDetailRows = window.BlendingEngine.buildSelectedBagDetail(selectedRows);
    renderSelectedBagDetailTable();
    renderPricingBreakdownTab(res, cfgActive);
    updateMobileStickyBar(cfgActive, res);
  }

  function renderSelectedBatchSummaryTable(selectedRows) {
    const tbody = document.getElementById("resBatchSummaryTableBody");
    if (!tbody) return;
    const summaryRows = window.BlendingEngine.buildSelectedBatchSummary(
      selectedRows,
      appState.invRows
    );

    tbody.innerHTML = summaryRows
      .map((r) => {
        const isFull = r["出貨方式"] === "整批全出";
        const badgeHtml = isFull
          ? `<span class="badge bg-success">整批全出</span>`
          : `<span class="badge bg-warning text-dark">拆批湊櫃</span>`;
        return `
        <tr>
          <td class="sticky-col fw-bold">${escapeHtml(r["批號"])}</td>
          <td>${escapeHtml(r["廠別"])}</td>
          <td><span class="badge bg-light text-dark border">${escapeHtml(r["進料代號"])}</span></td>
          <td>${escapeHtml(r["總表料號"])}</td>
          <td><span class="badge bg-secondary">${escapeHtml(r["編碼料號"])}</span></td>
          <td>${escapeHtml(r["存放位置"])}</td>
          <td class="text-center">${badgeHtml}</td>
          <td class="text-center fw-semibold">${escapeHtml(r["選中包數/原批包數"])}</td>
          <td class="text-end">${fmtInt(r["選中包數"])}</td>
          <td class="text-end">${fmtNum(r["含水率(%)"], 2)}%</td>
          <td class="text-end">${fmtNum(r["乾基Cu(%)"], 2)}%</td>
          <td class="text-end fw-semibold text-primary">${fmtNum(r["濕基Cu(%)"], 2)}%</td>
          <td class="text-end">${fmtNum(r["乾基Au(g/T)"], 2)}</td>
          <td class="text-end">${fmtNum(r["溼基Au(g/T)"], 2)}</td>
          <td class="text-end">${fmtNum(r["乾基Ag(g/T)"], 2)}</td>
          <td class="text-end">${fmtNum(r["溼基Ag(g/T)"], 2)}</td>
          <td class="text-end">${fmtNum(r["乾基Pd(g/T)"], 2)}</td>
          <td class="text-end">${fmtNum(r["溼基Pd(g/T)"], 2)}</td>
          <td class="text-end">${fmtNum(r["現場重小計(kg)"], 1)}</td>
          <td class="text-end fw-bold">${fmtNum(r["倉庫濕重小計(kg)"], 1)}</td>
          <td class="text-end">${fmtNum(r["乾重小計(kg)"], 2)}</td>
        </tr>
      `;
      })
      .join("");
  }

  function renderSelectedBagDetailTable() {
    const tbody = document.getElementById("resBagDetailTableBody");
    const footerEl = document.getElementById("resBagDetailFooterSummary");
    if (!tbody) return;

    const q = (document.getElementById("resBagSearchInput")?.value || "").trim().toLowerCase();
    const allBags = appState.lastBagDetailRows || [];
    const filtered = q
      ? allBags.filter(
          (r) =>
            String(r["單包代號"]).toLowerCase().includes(q) ||
            String(r["批號"]).toLowerCase().includes(q) ||
            String(r["進料代號"]).toLowerCase().includes(q) ||
            String(r["總表料號"]).toLowerCase().includes(q) ||
            String(r["編碼料號"]).toLowerCase().includes(q)
        )
      : allBags;

    tbody.innerHTML = filtered
      .map(
        (r) => `
      <tr>
        <td class="sticky-col fw-bold">#${fmtInt(r["單包代號"])}</td>
        <td class="fw-semibold">${escapeHtml(r["批號"])}</td>
        <td>${escapeHtml(r["廠別"])}</td>
        <td><span class="badge bg-light text-dark border">${escapeHtml(r["進料代號"])}</span></td>
        <td>${escapeHtml(r["總表料號"])}</td>
        <td><span class="badge bg-secondary">${escapeHtml(r["編碼料號"])}</span></td>
        <td>${escapeHtml(r["存放位置"])}</td>
        <td class="text-end">${fmtNum(r["現場重(kg)"], 1)}</td>
        <td class="text-end fw-semibold">${fmtNum(r["倉庫濕重(kg)"], 1)}</td>
        <td class="text-end">${fmtNum(r["單包乾重(kg)"], 2)}</td>
        <td class="text-end">${fmtNum(r["含水率(%)"], 2)}%</td>
        <td class="text-end">${fmtNum(r["乾基Cu(%)"], 4)}%</td>
        <td class="text-end text-primary fw-semibold">${fmtNum(r["濕基Cu(%)"], 4)}%</td>
        <td class="text-end">${fmtNum(r["乾基Au(g/T)"], 4)}</td>
        <td class="text-end">${fmtNum(r["溼基Au(g/T)"], 4)}</td>
        <td class="text-end">${fmtNum(r["乾基Ag(g/T)"], 4)}</td>
        <td class="text-end">${fmtNum(r["溼基Ag(g/T)"], 4)}</td>
        <td class="text-end">${fmtNum(r["乾基Pd(g/T)"], 4)}</td>
        <td class="text-end">${fmtNum(r["溼基Pd(g/T)"], 4)}</td>
      </tr>
    `
      )
      .join("");

    const res = appState.lastResult;
    const sel = appState.lastSelectedRows || [];
    if (footerEl && res) {
      const totalSiteKg = sel.reduce((acc, r) => acc + r["現場重"], 0.0);
      footerEl.innerHTML =
        `👉 <strong>【總重彙總】</strong> 選中總包數：<strong>${res.bag_count} 包</strong> ｜ ` +
        `總現場重：<strong>${fmtInt(totalSiteKg)} kg</strong> ｜ ` +
        `總倉庫濕重 (WMT)：<strong>${fmtInt(res.WMT_kg)} kg (${(res.WMT_kg / 1000.0).toFixed(3)} 噸)</strong> ｜ ` +
        `總乾重 (DMT)：<strong>${fmtNum(res.DMT_kg, 2)} kg (${(res.DMT_kg / 1000.0).toFixed(4)} 噸)</strong>`;
    }
  }

  function renderPricingBreakdownTab(res, cfgActive) {
    const container = document.getElementById("resPricingBreakdownContainer");
    if (!container) return;

    const currCode = cfgActive.price_currency;
    const currSym = currCode === "JPY" ? "¥" : "$";
    const activeMp = cfgActive.getActiveMetalPrices();
    const unitLabels = cfgActive.getMetalPriceUnitLabels();

    if (cfgActive.customer_type === "A") {
      const metals = [
        ["Cu", "kg"],
        ["Au", "g"],
        ["Ag", "g"],
        ["Pd", "g"],
      ];
      const rowsHtml = metals
        .map(([m, unit]) => {
          const gKey = m === "Cu" ? "Cu(%)" : `${m}(g/T)`;
          const qKey = `${m}_${unit}`;
          const gradeUnit = m === "Cu" ? "%" : "g/T";
          const mpStr = fmtNum(activeMp[m], 3).replace(/\.000$/, "");
          return `
          <tr>
            <td class="fw-bold">${m}</td>
            <td class="text-end">${res.dry_grades[gKey].toFixed(4)} ${gradeUnit}</td>
            <td class="text-end">${cfgActive.a_thresholds[m].toFixed(2)} ${gradeUnit}</td>
            <td class="text-end">${fmtNum(res.contained[qKey], 2)} ${unit}</td>
            <td class="text-end fw-semibold text-primary">${(res.recovery_rates[m] * 100.0).toFixed(2)}%</td>
            <td class="text-end fw-semibold">${fmtNum(res.payable_quantities[qKey], 2)} ${unit}</td>
            <td class="text-end">${mpStr} ${unitLabels[m]}</td>
            <td class="text-end">${fmtNum(cfgActive.a_rc[m], 2)} ${unitLabels[m]}</td>
            <td class="text-end fw-bold text-success">${currSym}${fmtNum(res.metal_revenues[m], 2)}</td>
          </tr>
        `;
        })
        .join("");

      container.innerHTML = `
        <h6 class="fw-bold mb-3">
          🏭 A 客戶（乾基 DMT 門檻回收率法 — ${currCode} 計價）五步驟計價分解
        </h6>
        <div class="table-responsive mb-3">
          <table class="table table-bordered table-hover yy-table">
            <thead>
              <tr>
                <th>金屬</th>
                <th class="text-end">乾基加權品位</th>
                <th class="text-end">合約門檻</th>
                <th class="text-end">金屬含有量</th>
                <th class="text-end">門檻回收率 (%)</th>
                <th class="text-end">計價回收量</th>
                <th class="text-end">市場單價</th>
                <th class="text-end">精煉費 RC</th>
                <th class="text-end">分項毛收入 (${currCode})</th>
              </tr>
            </thead>
            <tbody>${rowsHtml}</tbody>
          </table>
        </div>
        <div class="p-3 border rounded-3 bg-light small">
          <ul class="mb-0">
            <li><strong>毛收入合計 (Gross Revenue)</strong>：<code>${currCode} ${fmtNum(res.gross_revenue, 2)}</code></li>
            <li><strong>乾基運費扣款 (<code>${(res.DMT_kg / 1000.0).toFixed(4)} DMT × ${cfgActive.a_freight_per_dmt} ${currCode}/DMT</code>)</strong>：<code>-${currCode} ${fmtNum(res.freight_cost, 2)}</code></li>
            <li><strong>含水率超標扣款 (平均含水率 <code>${res.moisture_pct.toFixed(2)}%</code> vs 上限 <code>${cfgActive.a_moisture_limit.toFixed(1)}%</code>)</strong>：<code>-${currCode} ${fmtNum(res.moisture_penalty, 2)}</code></li>
            <li><strong>銅未達門檻 TC 扣款</strong>：<code>-${currCode} ${fmtNum(res.tc_penalty, 2)}</code></li>
            <li class="mt-1 fs-6"><strong>預估結算淨利 (Net Profit)</strong>：<strong class="text-success">${currCode} ${fmtNum(res.net_profit, 2)}</strong></li>
          </ul>
        </div>
      `;
    } else {
      const exclInfo =
        [...(cfgActive.exclude_pm_materials || []), ...(cfgActive.exclude_pm_codes || [])].join(
          ", "
        ) || "無";

      const rowsHtml = REQUIRED_METALS.map((m) => {
        const gKey = m === "Cu" ? "Cu(%)" : `${m}(g/T)`;
        const wDenom = m === "Cu" ? res.WMT_kg : res.PM_WMT_kg;
        const gradeUnit = m === "Cu" ? "%" : "g/T";
        const mpStr = fmtNum(activeMp[m], 3).replace(/\.000$/, "");
        return `
          <tr>
            <td class="fw-bold">${m}</td>
            <td class="text-end">${fmtInt(wDenom)} kg</td>
            <td class="text-end">${res.wet_grades[gKey].toFixed(4)} ${gradeUnit}</td>
            <td class="text-end fw-semibold text-primary">${(res.tier_rates[m] * 100.0).toFixed(1)}%</td>
            <td class="text-end">${mpStr} ${unitLabels[m]}</td>
            <td class="text-end fw-bold text-success">${currSym}${fmtNum(res.metal_revenues[m], 2)}</td>
          </tr>
        `;
      }).join("");

      container.innerHTML = `
        <h6 class="fw-bold mb-2">
          📊 ${cfgActive.customer_type} 客戶（濕基 WMT 階梯折數法 — ${currCode} 計價）分項計價分解
        </h6>
        <ul class="small mb-3">
          <li><strong>銅計價分母總濕重 (W_all)</strong>：<code>${fmtInt(res.WMT_kg)} kg</code></li>
          <li><strong>貴金屬計價分母總濕重 (W_pm，排除料號 <code>${escapeHtml(exclInfo)}</code>)</strong>：<code>${fmtInt(res.PM_WMT_kg)} kg</code></li>
        </ul>
        <div class="table-responsive mb-3">
          <table class="table table-bordered table-hover yy-table">
            <thead>
              <tr>
                <th>金屬</th>
                <th class="text-end">計價分母濕重 (kg)</th>
                <th class="text-end">濕基加權品位</th>
                <th class="text-end">適用階梯折數 (%)</th>
                <th class="text-end">市場單價</th>
                <th class="text-end">分項毛收入 (${currCode})</th>
              </tr>
            </thead>
            <tbody>${rowsHtml}</tbody>
          </table>
        </div>
        <div class="p-3 border rounded-3 bg-light small">
          <ul class="mb-0">
            <li><strong>毛收入合計 (Gross Revenue)</strong>：<code>${currCode} ${fmtNum(res.gross_revenue, 2)}</code></li>
            <li><strong>運費與固定扣減 (<code>b_freight_per_kg</code> + <code>b_deduction_per_ton</code>)</strong>：<code>-${currCode} ${fmtNum(res.deduction, 2)}</code></li>
            <li class="mt-1 fs-6"><strong>預估結算淨利 (Net Profit)</strong>：<strong class="text-success">${currCode} ${fmtNum(res.net_profit, 2)}</strong></li>
          </ul>
        </div>
      `;
    }
  }

  function updateMobileStickyBar(cfgActive, res) {
    const el = document.getElementById("mobileStickySummary");
    if (!el || !cfgActive) return;
    const currSym = cfgActive.price_currency === "JPY" ? "¥" : "$";
    if (!res || !res.feasible) {
      el.innerHTML = `
        <div><strong>${cfgActive.customer_type} 客戶 (${cfgActive.price_currency})</strong></div>
        <div class="text-warning">⚠️ 無合規組合，請調整條件</div>
      `;
    } else {
      el.innerHTML = `
        <div><strong>${cfgActive.customer_type} 客戶 ｜ ${res.bag_count}包 (${fmtInt(res.WMT_kg)} kg)</strong></div>
        <div class="text-info fw-bold">淨利：${currSym}${fmtNum(res.net_profit, 2)}</div>
      `;
    }
  }

  // =========================================================================
  // 檔案匯入 / 匯出與事件綁定
  // =========================================================================
  function loadInventoryFromRawRows(rawRows, label) {
    const invRows = window.BlendingEngine.loadAndNormalizeInventory(rawRows);
    const batchOverview = window.BlendingEngine.summarizeInventoryBatches(invRows);
    appState.invRows = invRows;
    appState.batchOverview = batchOverview;
    appState.invLabel = label;

    populatePickerOptionsFromInventory();
    renderSidebarInventoryStatus();
    renderInventoryOverviewTable();
    updateLiveFilterStats();
  }

  function downloadJsonObject(obj, fileName) {
    const jsonStr = JSON.stringify(obj, null, 2);
    const blob = new Blob([jsonStr], { type: "application/json;charset=utf-8" });
    const url = URL.createObjectURL(blob);
    const a = document.createElement("a");
    a.href = url;
    a.download = fileName;
    document.body.appendChild(a);
    a.click();
    document.body.removeChild(a);
    URL.revokeObjectURL(url);
  }

  function showSidebarToast(msg, type = "success") {
    const el = document.getElementById("sidebarToastMsg");
    if (!el) return;
    el.innerHTML = `<div class="alert alert-${type} py-1 px-2 mb-0">${escapeHtml(msg)}</div>`;
    setTimeout(() => {
      if (el) el.innerHTML = "";
    }, 4000);
  }

  function bindAllEvents() {
    // 1. 庫存來源切換與上傳 Excel
    document.querySelectorAll('input[name="invSourceMode"]').forEach((radio) => {
      radio.addEventListener("change", () => {
        const isUpload = document.getElementById("invSourceUpload").checked;
        document.getElementById("invUploadWrapper").classList.toggle("d-none", !isUpload);
        if (!isUpload) {
          loadInventoryFromRawRows(
            window.DEFAULT_INVENTORY_RAW_ROWS,
            "預設檔案：庫存明細表.xlsx"
          );
          triggerAutoSolveIfEnabled();
        }
      });
    });

    document.getElementById("invFileInput").addEventListener("change", (e) => {
      const file = e.target.files?.[0];
      if (!file) return;
      const reader = new FileReader();
      reader.onload = (evt) => {
        try {
          const data = new Uint8Array(evt.target.result);
          const workbook = XLSX.read(data, { type: "array" });
          const firstSheetName = workbook.SheetNames[0];
          const sheet = workbook.Sheets[firstSheetName];
          const rawRows = XLSX.utils.sheet_to_json(sheet, {
            header: 1,
            defval: null,
            raw: true,
          });
          loadInventoryFromRawRows(rawRows, `上傳檔案：${file.name}`);
          runOptimizationSolve();
        } catch (err) {
          const card = document.getElementById("invStatusCard");
          if (card) {
            card.className = "alert alert-danger py-2 px-3 small mb-3";
            card.textContent = `❌ 讀取庫存明細表失敗：${err.message}`;
          }
        }
      };
      reader.readAsArrayBuffer(file);
    });

    // 2. JSON 設定檔匯入 / 匯出 / 儲存預設
    document.getElementById("jsonFileInput").addEventListener("change", (e) => {
      const file = e.target.files?.[0];
      if (!file) return;
      const reader = new FileReader();
      reader.onload = (evt) => {
        try {
          const loadedJson = JSON.parse(evt.target.result);
          if ("global_defaults" in loadedJson) {
            appState.configData = loadedJson;
            const activeScenario = loadedJson.active_scenario_params || {
              customer_type: document.getElementById("selectCustomerType").value || "E",
            };
            applyParamsToUi(
              activeScenario,
              loadedJson.global_defaults,
              `匯入設定檔：${file.name}`,
              "custom"
            );
            showSidebarToast(`✅ 已套用設定檔：${file.name}`, "success");
            runOptimizationSolve();
          } else if ("customer_type" in loadedJson) {
            applyParamsToUi(
              loadedJson,
              appState.configData.global_defaults,
              `匯入單一情境：${file.name}`,
              "custom"
            );
            showSidebarToast(`✅ 已套用單一情境參數：${file.name}`, "success");
            runOptimizationSolve();
          } else {
            showSidebarToast("⚠️ JSON 檔案需包含 global_defaults 或 customer_type", "warning");
          }
        } catch (err) {
          showSidebarToast(`❌ 解析 JSON 失敗：${err.message}`, "danger");
        }
      };
      reader.readAsText(file, "utf-8");
    });

    document.getElementById("btnDownloadSidebarJson").addEventListener("click", () => {
      const currentUiParams = collectParamsFromUi();
      const globalDefsCopy = Object.assign({}, currentUiParams);
      delete globalDefsCopy.customer_type;
      const exportObj = {
        _comment: "永源化工 — 自動配櫃系統 Web UI 匯出設定檔 (配滿櫃模式)",
        inventory_excel_path:
          appState.configData.inventory_excel_path || "../庫存明細表.xlsx",
        global_defaults: globalDefsCopy,
        active_scenario_params: currentUiParams,
        scenarios: appState.configData.scenarios || {},
      };
      downloadJsonObject(exportObj, "blending_config_custom.json");
    });

    document.getElementById("btnSaveBrowserDefault").addEventListener("click", () => {
      try {
        const currentUiParams = collectParamsFromUi();
        window.BlendingConfig.CustomerConfig.fromDict(
          currentUiParams,
          appState.configData.global_defaults,
          "SaveCheck"
        );
        const globalDefsCopy = Object.assign({}, currentUiParams);
        delete globalDefsCopy.customer_type;
        const updatedCfg = Object.assign({}, appState.configData, {
          global_defaults: globalDefsCopy,
          active_scenario_params: currentUiParams,
        });
        localStorage.setItem(STORAGE_KEY_CONFIG, JSON.stringify(updatedCfg));
        appState.configData = updatedCfg;
        showSidebarToast("✅ 已成功儲存至瀏覽器預設設定！", "success");
      } catch (err) {
        showSidebarToast(`❌ 儲存失敗：${err.message}`, "danger");
      }
    });

    document.getElementById("btnResetFactoryDefault").addEventListener("click", () => {
      localStorage.removeItem(STORAGE_KEY_CONFIG);
      appState.configData = JSON.parse(JSON.stringify(window.DEFAULT_CONFIG_DATA));
      const scenE = appState.configData.scenarios.example_1_customer_E;
      applyParamsToUi(
        Object.assign({}, scenE.base_params, scenE.modes[0]),
        appState.configData.global_defaults,
        "實務範例一：E 客戶配櫃 (預設載入，USD 計價)",
        "preset1"
      );
      showSidebarToast("🔄 已還原出廠預設設定！", "info");
      runOptimizationSolve();
    });

    // 3. 四大快速範本按鈕
    document.getElementById("btnPreset1").addEventListener("click", () => {
      const scen = appState.configData.scenarios.example_1_customer_E;
      applyParamsToUi(
        Object.assign({}, scen.base_params, scen.modes[0]),
        appState.configData.global_defaults,
        "實務範例一：E 客戶配櫃 (USD 計價 | 必出 1F1+3P1、銅>=5%、貴金屬排除料號 X)",
        "preset1"
      );
      runOptimizationSolve();
    });

    document.getElementById("btnPreset2").addEventListener("click", () => {
      const scen = appState.configData.scenarios.example_2_customer_A;
      applyParamsToUi(
        Object.assign({}, scen.base_params, scen.modes[0]),
        appState.configData.global_defaults,
        "實務範例二：A 客戶配櫃 (日圓 JPY 計價 | Cu 1,501,000 ¥/T, Au 13,462 ¥/g, Ag 157.474 ¥/g, Pd 5,334 ¥/g)",
        "preset2"
      );
      runOptimizationSolve();
    });

    document.getElementById("btnPreset3").addEventListener("click", () => {
      const scen = appState.configData.scenarios.example_3_customer_B;
      applyParamsToUi(
        Object.assign({}, scen.base_params, scen.modes[0]),
        appState.configData.global_defaults,
        "實務範例三：B 客戶配櫃 (USD 計價 | 禁收編碼 B、同櫃最大銅落差 <= 7.0%)",
        "preset3"
      );
      runOptimizationSolve();
    });

    document.getElementById("btnPresetDefault").addEventListener("click", () => {
      applyParamsToUi(
        { customer_type: "B" },
        appState.configData.global_defaults,
        "自訂配櫃模式 (全域預設值)",
        "presetDefault"
      );
      runOptimizationSolve();
    });

    // 4. 表單控制項變動監聽
    document.getElementById("selectPriceCurrency").addEventListener("change", () => {
      updateCurrencyLabelsInUi();
      triggerAutoSolveIfEnabled();
    });

    document.getElementById("checkEnableMaxTotalValue").addEventListener("change", (e) => {
      document
        .getElementById("maxTotalValueWrapper")
        .classList.toggle("d-none", !e.target.checked);
      triggerAutoSolveIfEnabled();
    });

    const formInputIds = [
      "selectCustomerType",
      "inputFullMinKg",
      "inputFullMaxKg",
      "inputSplitBatchPenalty",
      "inputTargetCuMin",
      "inputTargetCuMax",
      "inputMaxDiffCu",
      "inputMaxTotalValue",
      "inputPriceJpyCu",
      "inputPriceJpyAu",
      "inputPriceJpyAg",
      "inputPriceJpyPd",
      "inputPriceUsdCu",
      "inputPriceUsdAu",
      "inputPriceUsdAg",
      "inputPriceUsdPd",
      "inputTroyOzToG",
      "selectAMode",
      "inputAThCu",
      "inputAThAu",
      "inputAThAg",
      "inputAThPd",
      "inputARcCu",
      "inputARcAu",
      "inputARcAg",
      "inputARcPd",
      "inputAFreightPerDmt",
      "inputAMoistureLimit",
      "inputAMoisturePenaltyPerWmt",
      "inputACuBelowTcPerWmt",
      "inputAAssayFeeIfBelow",
      "inputBFreightPerKg",
      "inputBDeductionPerTon",
    ];

    formInputIds.forEach((id) => {
      const el = document.getElementById(id);
      if (el) {
        el.addEventListener("change", () => triggerAutoSolveIfEnabled());
      }
    });

    // 5. 階梯折數表「新增階梯」按鈕
    document.querySelectorAll("[data-add-tier]").forEach((btn) => {
      btn.addEventListener("click", () => {
        const m = btn.getAttribute("data-add-tier");
        const list = appState.tierTables[m] || [];
        const lastMax = list.length > 0 ? Number(list[list.length - 1].max_grade) : 0.0;
        const lastRate = list.length > 0 ? Number(list[list.length - 1].payable_rate) : 0.5;
        list.push({
          min_grade: lastMax,
          max_grade: lastMax + 10.0,
          payable_rate: Math.min(1.0, lastRate),
        });
        appState.tierTables[m] = list;
        renderSingleTierEditor(m);
        triggerAutoSolveIfEnabled();
      });
    });

    // 6. 庫存總覽與逐包明細搜尋框
    document.getElementById("invSearchInput").addEventListener("input", () => {
      renderInventoryOverviewTable();
    });
    document.getElementById("invAssayFilter").addEventListener("change", () => {
      renderInventoryOverviewTable();
    });
    document.getElementById("resBagSearchInput").addEventListener("input", () => {
      renderSelectedBagDetailTable();
    });

    // 7. 立即執行試算與匯出按鈕
    document.getElementById("btnRunSolve").addEventListener("click", () => {
      runOptimizationSolve();
    });
    document.getElementById("btnMobileRunSolve").addEventListener("click", () => {
      runOptimizationSolve();
      document.getElementById("resultSectionCard")?.scrollIntoView({ behavior: "smooth" });
    });

    document.getElementById("btnExportExcel").addEventListener("click", () => {
      if (!appState.lastResult || !appState.lastResult.feasible || !appState.lastActiveConfig) {
        return;
      }
      window.BlendingEngine.exportBlendingResultToExcel(
        appState.lastSelectedRows,
        appState.invRows,
        appState.lastResult,
        appState.lastActiveConfig
      );
    });

    document.getElementById("btnExportScenarioJson").addEventListener("click", () => {
      if (!appState.lastActiveConfig) return;
      const cfgDict = appState.lastActiveConfig.toDict();
      const fileName = `配櫃參數_${cfgDict.customer_type}客戶_${cfgDict.price_currency}.json`;
      downloadJsonObject(cfgDict, fileName);
    });
  }

  // =========================================================================
  // 系統啟動入口
  // =========================================================================
  function initApp() {
    initPickers();

    // 載入預設設定檔 (或 localStorage 中已儲存之自訂預設值)
    let initialConfig = JSON.parse(JSON.stringify(window.DEFAULT_CONFIG_DATA));
    const savedStr = localStorage.getItem(STORAGE_KEY_CONFIG);
    if (savedStr) {
      try {
        const parsed = JSON.parse(savedStr);
        if (parsed && parsed.global_defaults) {
          initialConfig = parsed;
        }
      } catch (_) {
        // 忽略損毀之快取
      }
    }
    appState.configData = initialConfig;

    // 載入內建《庫存明細表.xlsx》資料
    loadInventoryFromRawRows(window.DEFAULT_INVENTORY_RAW_ROWS, "預設檔案：庫存明細表.xlsx");

    // 載入預設情境 (範例一：E 客戶配櫃)
    const scenE = appState.configData.scenarios.example_1_customer_E;
    applyParamsToUi(
      Object.assign({}, scenE.base_params, scenE.modes[0]),
      appState.configData.global_defaults,
      "實務範例一：E 客戶配櫃 (預設載入，USD 計價)",
      "preset1"
    );

    bindAllEvents();
    runOptimizationSolve();
  }

  if (document.readyState === "loading") {
    document.addEventListener("DOMContentLoaded", initApp);
  } else {
    initApp();
  }
})();

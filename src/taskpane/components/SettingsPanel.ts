/**
 * ============================================================================
 * SettingsPanel
 * ----------------------------------------------------------------------------
 * Builds and manages the settings overlay. All controls are bound to a
 * `DeepSeekConfig` instance; on every change the panel fires a
 * `settings-change` CustomEvent so the parent can persist & rebuild chat
 * context.
 *
 * The panel also exposes a "Reset to defaults" button and basic validation.
 * ============================================================================
 */

import { AVAILABLE_MODELS, createDefaultConfig } from "../utils/constants";
import { clampFloat, clampInt, isValidApiKeyFormat } from "../utils/helpers";
import type { DeepSeekConfig } from "../types";

export interface SettingsChangeDetail {
  config: DeepSeekConfig;
  /** Set when the user explicitly reset to defaults. */
  reset: boolean;
}

export class SettingsPanel {
  /** Public alias so the host can attach listeners without exposing `root`. */
  public get element(): HTMLElement {
    return this.root;
  }

  private root: HTMLElement;
  private config: DeepSeekConfig;
  private visible: boolean = false;

  /** Cached references to form elements. */
  private refs: {
    apiKey: HTMLInputElement;
    apiKeyToggle: HTMLButtonElement;
    apiKeyHint: HTMLElement;
    baseUrl: HTMLInputElement;
    model: HTMLSelectElement;
    temperature: HTMLInputElement;
    temperatureValue: HTMLElement;
    maxTokens: HTMLInputElement;
    topP: HTMLInputElement;
    topPValue: HTMLElement;
    systemPrompt: HTMLTextAreaElement;
    resetBtn: HTMLButtonElement;
    closeBtn: HTMLButtonElement;
  };

  constructor(parent: HTMLElement, initial: DeepSeekConfig) {
    this.config = { ...initial };
    this.root = document.createElement("div");
    this.root.classList.add("settings-overlay");
    this.root.style.display = "none";
    this.root.innerHTML = this.renderHtml();
    parent.appendChild(this.root);

    this.refs = this.collectRefs();
    this.bindForm();
    this.applyConfigToForm();
    this.validateApiKey();
  }

  /* ---------------- public ---------------- */

  public show(): void {
    this.visible = true;
    this.root.style.display = "flex";
    // Focus the API key field by default for fast editing.
    setTimeout(() => this.refs.apiKey.focus(), 50);
  }

  public hide(): void {
    this.visible = false;
    this.root.style.display = "none";
  }

  public toggle(): void {
    this.visible ? this.hide() : this.show();
  }

  public isVisible(): boolean {
    return this.visible;
  }

  /** Update the panel from an external source (e.g. after reset). */
  public setConfig(config: DeepSeekConfig): void {
    this.config = { ...config };
    this.applyConfigToForm();
    this.validateApiKey();
  }

  public getConfig(): DeepSeekConfig {
    return { ...this.config };
  }

  /* ---------------- rendering ---------------- */

  private renderHtml(): string {
    const modelOptions = AVAILABLE_MODELS.map(
      (m) => `<option value="${m.value}">${m.label} - ${m.description}</option>`
    ).join("");

    return `
      <div class="settings-card" role="dialog" aria-label="设置">
        <div class="settings-header">
          <h2>⚙️ 设置</h2>
          <button type="button" class="settings-close" data-action="close" title="关闭">✕</button>
        </div>

        <div class="settings-body">
          <!-- API Key -->
          <section class="settings-section">
            <label class="settings-label">
              <span class="settings-label-text">API Key</span>
              <span class="settings-label-hint">仅存储在浏览器本地，不会上传</span>
            </label>
            <div class="settings-input-row">
              <input
                type="password"
                class="settings-input"
                data-field="apiKey"
                placeholder="sk-xxxxxxxxxxxxxxxx"
                autocomplete="off"
                spellcheck="false"
              />
              <button type="button" class="settings-icon-btn" data-action="toggleKey" title="显示/隐藏">👁</button>
            </div>
            <div class="settings-hint" data-ref="apiKeyHint"></div>
          </section>

          <!-- Base URL -->
          <section class="settings-section">
            <label class="settings-label">
              <span class="settings-label-text">API 端点</span>
              <span class="settings-label-hint">兼容 OpenAI 格式的代理可修改</span>
            </label>
            <input
              type="url"
              class="settings-input"
              data-field="baseUrl"
              placeholder="https://api.deepseek.com"
            />
          </section>

          <!-- Model -->
          <section class="settings-section">
            <label class="settings-label">
              <span class="settings-label-text">模型</span>
            </label>
            <select class="settings-input" data-field="model">
              ${modelOptions}
            </select>
          </section>

          <!-- Temperature -->
          <section class="settings-section">
            <label class="settings-label">
              <span class="settings-label-text">Temperature</span>
              <span class="settings-label-hint" data-ref="temperatureValue">0.7</span>
            </label>
            <input
              type="range"
              class="settings-range"
              data-field="temperature"
              min="0" max="2" step="0.1"
            />
          </section>

          <!-- Max Tokens -->
          <section class="settings-section">
            <label class="settings-label">
              <span class="settings-label-text">Max Tokens</span>
              <span class="settings-label-hint">范围 256 ~ 8192</span>
            </label>
            <input
              type="number"
              class="settings-input"
              data-field="maxTokens"
              min="256" max="8192" step="64"
            />
          </section>

          <!-- Top P -->
          <section class="settings-section">
            <label class="settings-label">
              <span class="settings-label-text">Top P</span>
              <span class="settings-label-hint" data-ref="topPValue">0.9</span>
            </label>
            <input
              type="range"
              class="settings-range"
              data-field="topP"
              min="0" max="1" step="0.05"
            />
          </section>

          <!-- System Prompt -->
          <section class="settings-section">
            <label class="settings-label">
              <span class="settings-label-text">系统提示词</span>
              <span class="settings-label-hint">影响所有对话的风格与能力</span>
            </label>
            <textarea
              class="settings-textarea"
              data-field="systemPrompt"
              rows="6"
              spellcheck="false"
            ></textarea>
          </section>
        </div>

        <div class="settings-footer">
          <button type="button" class="settings-btn settings-btn-ghost" data-action="reset">
            🔄 重置默认
          </button>
          <button type="button" class="settings-btn settings-btn-primary" data-action="save">
            💾 保存
          </button>
        </div>
      </div>
    `;
  }

  private collectRefs() {
    const find = <T extends HTMLElement>(sel: string): T => {
      const el = this.root.querySelector<T>(sel);
      if (!el) throw new Error("Missing settings element: " + sel);
      return el;
    };

    return {
      apiKey: find<HTMLInputElement>('[data-field="apiKey"]'),
      apiKeyToggle: find<HTMLButtonElement>('[data-action="toggleKey"]'),
      apiKeyHint: find<HTMLElement>('[data-ref="apiKeyHint"]'),
      baseUrl: find<HTMLInputElement>('[data-field="baseUrl"]'),
      model: find<HTMLSelectElement>('[data-field="model"]'),
      temperature: find<HTMLInputElement>('[data-field="temperature"]'),
      temperatureValue: find<HTMLElement>('[data-ref="temperatureValue"]'),
      maxTokens: find<HTMLInputElement>('[data-field="maxTokens"]'),
      topP: find<HTMLInputElement>('[data-field="topP"]'),
      topPValue: find<HTMLElement>('[data-ref="topPValue"]'),
      systemPrompt: find<HTMLTextAreaElement>('[data-field="systemPrompt"]'),
      resetBtn: find<HTMLButtonElement>('[data-action="reset"]'),
      closeBtn: find<HTMLButtonElement>('[data-action="close"]'),
    };
  }

  /* ---------------- form binding ---------------- */

  private applyConfigToForm(): void {
    const r = this.refs;
    r.apiKey.value = this.config.apiKey;
    r.baseUrl.value = this.config.baseUrl;
    r.model.value = this.config.model;
    r.temperature.value = String(this.config.temperature);
    r.temperatureValue.textContent = this.config.temperature.toFixed(1);
    r.maxTokens.value = String(this.config.maxTokens);
    r.topP.value = String(this.config.topP);
    r.topPValue.textContent = this.config.topP.toFixed(2);
    r.systemPrompt.value = this.config.systemPrompt;
  }

  private bindForm(): void {
    const r = this.refs;

    // Live updates - reflect values immediately, validate, but only fire
    // the change event on explicit "save" or when the user moves focus
    // away. This avoids hammering localStorage on every keystroke.

    r.apiKey.addEventListener("input", () => {
      this.config.apiKey = r.apiKey.value.trim();
      this.validateApiKey();
    });

    r.apiKeyToggle.addEventListener("click", () => {
      const showing = r.apiKey.type === "text";
      r.apiKey.type = showing ? "password" : "text";
      r.apiKeyToggle.textContent = showing ? "👁" : "🙈";
    });

    r.baseUrl.addEventListener("input", () => {
      this.config.baseUrl = r.baseUrl.value.trim();
    });

    r.model.addEventListener("change", () => {
      this.config.model = r.model.value as DeepSeekConfig["model"];
    });

    r.temperature.addEventListener("input", () => {
      const v = clampFloat(r.temperature.value, 0, 2, 0.7);
      this.config.temperature = v;
      r.temperatureValue.textContent = v.toFixed(1);
    });

    r.maxTokens.addEventListener("input", () => {
      const v = clampInt(r.maxTokens.value, 256, 8192, 2048);
      this.config.maxTokens = v;
    });

    r.topP.addEventListener("input", () => {
      const v = clampFloat(r.topP.value, 0, 1, 0.9);
      this.config.topP = v;
      r.topPValue.textContent = v.toFixed(2);
    });

    r.systemPrompt.addEventListener("input", () => {
      this.config.systemPrompt = r.systemPrompt.value;
    });

    // Action buttons
    this.root.addEventListener("click", (ev) => {
      const target = ev.target as HTMLElement;
      const action = target.dataset.action;
      if (!action) return;
      if (action === "close") this.hide();
      if (action === "save") this.fireChange(false);
      if (action === "reset") {
        this.config = createDefaultConfig();
        this.applyConfigToForm();
        this.validateApiKey();
        this.fireChange(true);
      }
    });
  }

  /** Show a green check or red warning next to the API key field. */
  private validateApiKey(): void {
    const r = this.refs;
    const value = this.config.apiKey;
    if (!value) {
      r.apiKeyHint.textContent = "请填写 API Key 才能发起请求";
      r.apiKeyHint.className = "settings-hint settings-hint-warn";
      return;
    }
    if (!value.startsWith("sk-")) {
      r.apiKeyHint.textContent = "提示：官方 DeepSeek Key 通常以 sk- 开头";
      r.apiKeyHint.className = "settings-hint settings-hint-warn";
      return;
    }
    if (!isValidApiKeyFormat(value)) {
      r.apiKeyHint.textContent = "Key 长度过短，请检查是否完整";
      r.apiKeyHint.className = "settings-hint settings-hint-warn";
      return;
    }
    r.apiKeyHint.textContent = "✓ 格式正确";
    r.apiKeyHint.className = "settings-hint settings-hint-ok";
  }

  private fireChange(reset: boolean): void {
    this.root.dispatchEvent(
      new CustomEvent("settings-change", {
        detail: { config: { ...this.config }, reset } satisfies SettingsChangeDetail,
        bubbles: true,
      })
    );
  }
}
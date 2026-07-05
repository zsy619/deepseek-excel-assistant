/**
 * ============================================================================
 * RibbonBanner
 * ----------------------------------------------------------------------------
 * Top-of-taskpane status banner that surfaces what the ribbon is doing.
 *
 * The ribbon buttons dispatch via 5 cross-frame channels. Without a banner,
 * the user clicks a button and nothing visible happens in the taskpane until
 * the chat stream starts (which can take a second or two). This banner:
 *
 *   - Appears the moment a ribbon command arrives in the taskpane
 *   - Shows the friendly command label (e.g. "正在响应 · 生成公式")
 *   - Auto-hides when the command's chat/panel handler completes
 *   - Spinner + accent border in Excel brand green (#217346)
 *
 * Lifecycle:
 *   show(cmd, label)   - banner mounts, returns nothing
 *   update(progress)   - optional streaming progress (e.g. "已读 32%")
 *   hide(reason?)      - banner fades out (reason: "done" | "error" | "cancel")
 *
 * Singleton: there is at most one banner in the DOM at a time. Showing
 * twice replaces the previous banner cleanly.
 * ============================================================================
 */

export type BannerReason = "done" | "error" | "cancel" | "replace";

const HOST_ID = "deepseek-ribbon-banner-host";

function ensureHost(): HTMLElement {
  let host = document.getElementById(HOST_ID);
  if (!host) {
    host = document.createElement("div");
    host.id = HOST_ID;
    host.classList.add("ribbon-banner-host");
    document.body.appendChild(host);
  }
  return host;
}

export class RibbonBanner {
  private root: HTMLDivElement | null = null;
  private labelEl: HTMLSpanElement | null = null;
  private progressEl: HTMLSpanElement | null = null;
  /** Auto-hide timer - cancelled when update() is called or hide() is
   *  called explicitly. */
  private hideTimer: number | null = null;

  /** Show the banner. Replaces any existing banner. */
  public show(commandId: string, label: string): void {
    const host = ensureHost();
    this.clearTimer();
    // Replace any existing banner.
    if (this.root) {
      try {
        this.root.remove();
      } catch {
        /* noop */
      }
    }
    const root = document.createElement("div");
    root.className = "ribbon-banner ribbon-banner--active";
    root.setAttribute("role", "status");
    root.setAttribute("aria-live", "polite");
    root.dataset.cmd = commandId;

    const spinner = document.createElement("span");
    spinner.className = "ribbon-banner__spinner";
    spinner.setAttribute("aria-hidden", "true");
    root.appendChild(spinner);

    const text = document.createElement("span");
    text.className = "ribbon-banner__label";
    text.textContent = `正在响应 · ${label}`;
    this.labelEl = text;
    root.appendChild(text);

    const progress = document.createElement("span");
    progress.className = "ribbon-banner__progress";
    progress.textContent = "";
    this.progressEl = progress;
    root.appendChild(progress);

    host.appendChild(root);
    this.root = root;

    // Auto-fade after 30s as a safety net - any in-flight operation
    // typically finishes before that. Calling hide() early cancels this.
    this.hideTimer = window.setTimeout(() => {
      try {
        this.hide("done");
      } catch {
        /* noop */
      }
    }, 30000);
  }

  /** Update the progress label (e.g. "已分析 4/10"). */
  public update(text: string | null): void {
    if (!this.root || !this.progressEl) return;
    if (text === null || text === "") {
      this.progressEl.textContent = "";
      this.root.classList.remove("ribbon-banner--has-progress");
    } else {
      this.progressEl.textContent = ` · ${text}`;
      this.root.classList.add("ribbon-banner--has-progress");
    }
  }

  /** Hide the banner with optional reason styling. */
  public hide(reason: BannerReason = "done"): void {
    this.clearTimer();
    if (!this.root) return;
    const root = this.root;
    this.root = null;
    this.labelEl = null;
    this.progressEl = null;
    // Brief "done" flash before fadeout for visual confirmation.
    if (reason === "done") {
      root.classList.add("ribbon-banner--done");
      window.setTimeout(() => {
        root.classList.remove("ribbon-banner--active");
        window.setTimeout(() => {
          try {
            root.remove();
          } catch {
            /* noop */
          }
        }, 220);
      }, 350);
    } else if (reason === "error") {
      root.classList.add("ribbon-banner--error");
      root.querySelector(".ribbon-banner__label")?.replaceChildren(
        document.createTextNode("响应失败")
      );
      window.setTimeout(() => {
        root.classList.remove("ribbon-banner--active");
        window.setTimeout(() => {
          try {
            root.remove();
          } catch {
            /* noop */
          }
        }, 220);
      }, 1500);
    } else {
      // cancel / replace - just remove without flash.
      try {
        root.remove();
      } catch {
        /* noop */
      }
    }
  }

  private clearTimer(): void {
    if (this.hideTimer !== null) {
      window.clearTimeout(this.hideTimer);
      this.hideTimer = null;
    }
  }

  /** Whether the banner is currently visible. */
  public isVisible(): boolean {
    return !!this.root;
  }
}
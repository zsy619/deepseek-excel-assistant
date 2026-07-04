/**
 * ============================================================================
 * Dialog
 * ----------------------------------------------------------------------------
 * Office Add-ins run inside a webview that disallows window.prompt,
 * window.confirm, and window.alert. This module provides tiny replacements
 * that render inside the taskpane as modal overlays.
 *
 * Two flavors:
 *   - confirmDialog()  - Yes / No decision
 *   - promptDialog()   - Text input with submit / cancel
 *
 * Both are promise-based and resolve to null when the user dismisses the
 * dialog without choosing.
 * ============================================================================
 */

export interface ConfirmOptions {
  title?: string;
  message: string;
  confirmText?: string;
  cancelText?: string;
  variant?: "default" | "danger";
}

export interface PromptOptions {
  title?: string;
  message: string;
  defaultValue?: string;
  placeholder?: string;
  confirmText?: string;
  cancelText?: string;
  multiline?: boolean;
  rows?: number;
}

/** Internal helper - the singleton host element used by all dialogs. */
function ensureHost(): HTMLElement {
  const id = "deepseek-dialog-host";
  let host = document.getElementById(id);
  if (!host) {
    host = document.createElement("div");
    host.id = id;
    host.classList.add("deepseek-dialog-host");
    document.body.appendChild(host);
  }
  return host;
}

/** Build the modal chrome shared by all dialogs. */
function buildShell(
  title: string,
  body: HTMLElement,
  footer: HTMLElement,
  variant: "default" | "danger" = "default"
): HTMLElement {
  const overlay = document.createElement("div");
  overlay.className = `dialog-overlay ${variant === "danger" ? "dialog-danger" : ""}`;

  const card = document.createElement("div");
  card.className = "dialog-card";
  card.setAttribute("role", "dialog");
  card.setAttribute("aria-modal", "true");

  const header = document.createElement("div");
  header.className = "dialog-header";
  const h = document.createElement("h3");
  h.textContent = title;
  header.appendChild(h);

  const content = document.createElement("div");
  content.className = "dialog-content";
  content.appendChild(body);

  card.appendChild(header);
  card.appendChild(content);
  card.appendChild(footer);

  overlay.appendChild(card);
  return overlay;
}

/** Wire up Escape + click-outside to dismiss. */
function wireDismiss(overlay: HTMLElement, onDismiss: () => void): () => void {
  const onKey = (e: KeyboardEvent) => {
    if (e.key === "Escape") onDismiss();
  };
  overlay.addEventListener("click", (e) => {
    if (e.target === overlay) onDismiss();
  });
  document.addEventListener("keydown", onKey);
  // Focus the first focusable element inside the card.
  setTimeout(() => {
    const focusable = overlay.querySelector<HTMLElement>(
      "input, textarea, button"
    );
    focusable?.focus();
  }, 50);
  return () => document.removeEventListener("keydown", onKey);
}

/* ----------------------------------------------------------------- */
/* Public API                                                         */
/* ----------------------------------------------------------------- */

/**
 * Show a Yes / No dialog. Resolves to true (confirm) / false (cancel) /
 * null (dismissed via Escape or click-outside). Use `null` to distinguish
 * "user explicitly cancelled" from "no answer yet" in caller code if needed.
 */
export function confirmDialog(options: ConfirmOptions): Promise<boolean | null> {
  return new Promise((resolve) => {
    const host = ensureHost();
    const body = document.createElement("div");
    body.className = "dialog-body";
    const msg = document.createElement("p");
    msg.textContent = options.message;
    body.appendChild(msg);

    const confirmText = options.confirmText || "确定";
    const cancelText = options.cancelText || "取消";

    const footer = document.createElement("div");
    footer.className = "dialog-footer";
    const cancelBtn = document.createElement("button");
    cancelBtn.type = "button";
    cancelBtn.className = "dialog-btn dialog-btn-ghost";
    cancelBtn.textContent = cancelText;
    const confirmBtn = document.createElement("button");
    confirmBtn.type = "button";
    confirmBtn.className =
      "dialog-btn " +
      (options.variant === "danger"
        ? "dialog-btn-danger"
        : "dialog-btn-primary");
    confirmBtn.textContent = confirmText;
    footer.appendChild(cancelBtn);
    footer.appendChild(confirmBtn);

    const overlay = buildShell(options.title || "确认", body, footer, options.variant);
    host.appendChild(overlay);

    let resolved = false;
    const cleanup = () => {
      if (resolved) return;
      resolved = true;
      try {
        overlay.remove();
      } catch {
        /* noop */
      }
      unbind();
    };
    const finish = (val: boolean | null) => {
      cleanup();
      resolve(val);
    };
    const unbind = wireDismiss(overlay, () => finish(null));

    confirmBtn.addEventListener("click", () => finish(true));
    cancelBtn.addEventListener("click", () => finish(false));
  });
}

/**
 * Show a text input dialog. Resolves to the entered string on confirm, or
 * null on cancel / dismiss.
 */
export function promptDialog(options: PromptOptions): Promise<string | null> {
  return new Promise((resolve) => {
    const host = ensureHost();
    const body = document.createElement("div");
    body.className = "dialog-body";
    const msg = document.createElement("p");
    msg.textContent = options.message;
    body.appendChild(msg);

    const multiline = !!options.multiline;
    const input: HTMLInputElement | HTMLTextAreaElement = multiline
      ? document.createElement("textarea")
      : document.createElement("input");
    input.className = "dialog-input";
    if (input instanceof HTMLTextAreaElement) {
      input.rows = options.rows || 4;
    } else {
      input.type = "text";
    }
    input.placeholder = options.placeholder || "";
    input.value = options.defaultValue || "";
    body.appendChild(input);

    const confirmText = options.confirmText || "确定";
    const cancelText = options.cancelText || "取消";
    const footer = document.createElement("div");
    footer.className = "dialog-footer";
    const cancelBtn = document.createElement("button");
    cancelBtn.type = "button";
    cancelBtn.className = "dialog-btn dialog-btn-ghost";
    cancelBtn.textContent = cancelText;
    const confirmBtn = document.createElement("button");
    confirmBtn.type = "button";
    confirmBtn.className = "dialog-btn dialog-btn-primary";
    confirmBtn.textContent = confirmText;
    footer.appendChild(cancelBtn);
    footer.appendChild(confirmBtn);

    const overlay = buildShell(options.title || "请输入", body, footer, "default");
    host.appendChild(overlay);

    let resolved = false;
    const cleanup = () => {
      if (resolved) return;
      resolved = true;
      try {
        overlay.remove();
      } catch {
        /* noop */
      }
      unbind();
    };
    const finish = (val: string | null) => {
      cleanup();
      resolve(val);
    };
    const unbind = wireDismiss(overlay, () => finish(null));

    confirmBtn.addEventListener("click", () => finish(input.value));
    cancelBtn.addEventListener("click", () => finish(null));

    // Submit on Enter for single-line inputs.
    if (!multiline) {
      input.addEventListener("keydown", (e) => {
        const ke = e as KeyboardEvent;
        if (ke.key === "Enter") {
          ke.preventDefault();
          finish(input.value);
        }
      });
    } else {
      // Ctrl+Enter / Cmd+Enter submits for multi-line.
      input.addEventListener("keydown", (e) => {
        const ke = e as KeyboardEvent;
        if ((ke.ctrlKey || ke.metaKey) && ke.key === "Enter") {
          ke.preventDefault();
          finish(input.value);
        }
      });
    }
  });
}
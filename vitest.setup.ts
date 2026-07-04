/**
 * ============================================================================
 * Vitest Setup - Office.js mocks + global test utilities
 * ============================================================================
 *
 * This file runs before every test suite. It provides a mock implementation
 * of the Office.js API surface that our code actually uses, so tests can
 * run without Excel.
 */

import { vi } from "vitest";

/* ------------------------------------------------------------------ */
/*  Office.js mock                                                     */
/* ------------------------------------------------------------------ */

// Minimal mock covering the Office.js calls our add-in makes.
// Expand as needed when new API calls are added.
const mockOffice = {
  onReady: vi.fn(),
  context: {
    document: {
      getSelectedDataAsync: vi.fn(),
      setSelectedDataAsync: vi.fn(),
      bindings: {
        addFromPromptAsync: vi.fn(),
        addFromNamedItemAsync: vi.fn(),
      },
    },
    mailbox: undefined,
  },
  EventType: {
    DialogMessageReceived: "dialogMessageReceived",
    DialogEventReceived: "dialogEventReceived",
  },
  CoercionType: {
    Text: "text",
    Matrix: "matrix",
    Table: "table",
    Html: "html",
  },
  AsyncResultStatus: {
    Succeeded: "succeeded",
    Failed: "failed",
  },
  initialize: vi.fn(),
  // Used for dialogs / auth flows
  displayDialogAsync: vi.fn(),
  close: vi.fn(),
};

// Attach to global scope so imports of "office-js" resolve.
(globalThis as any).Office = mockOffice;

/* ------------------------------------------------------------------ */
/*  localStorage mock                                                  */
/* ------------------------------------------------------------------ */

// Vitest's jsdom provides a real localStorage implementation as long as
// the `environment` is "jsdom", so we don't need to mock it manually.
// We do clear it between tests to prevent cross-test leakage.

beforeEach(() => {
  localStorage.clear();
});

/* ------------------------------------------------------------------ */
/*  Helpers exported for convenience in tests                          */
/* ------------------------------------------------------------------ */

/**
 * Set up a known-good Office context for a test that exercises Office.js
 * wrappers. Calling this ensures `Office.onReady` resolves immediately
 * with a fake context.
 */
export function mockOfficeReady() {
  const mockContext = { host: "EXCEL", platform: "MAC" };
  mockOffice.onReady.mockImplementation(
    (callback: (info: any) => void) => callback(mockContext)
  );
  // Also resolve the Promise-based version if Office.onReady ever
  // returns a Promise (newer API shape).
  mockOffice.onReady.mockReturnValue(Promise.resolve(mockContext));
}

/**
 * Reset all Office mocks to their initial state.
 */
export function resetOfficeMocks() {
  vi.clearAllMocks();
  mockOffice.onReady.mockReset();
  mockOffice.context.document.getSelectedDataAsync.mockReset();
  mockOffice.context.document.setSelectedDataAsync.mockReset();
}

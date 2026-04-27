import "./styles.css";
import { RatonixClient } from "./ratonixClient";
import { loadStyleMemory, rememberInteraction } from "./styleMemory";
import type { DraftKind, MessageSnapshot, RatonixSuggestion, Tone } from "./types";
import {
  getCurrentMessageSnapshot,
  getHostInfo,
  insertTextIntoCompose,
  isComposeMode,
} from "./outlook";

const tones: Tone[] = ["balanced", "softer", "firmer"];

interface AppState {
  message: MessageSnapshot | null;
  selectedTone: Tone;
  suggestion: RatonixSuggestion | null;
  isBusy: boolean;
  error: string | null;
}

const client = new RatonixClient();
const state: AppState = {
  message: null,
  selectedTone: "balanced",
  suggestion: null,
  isBusy: false,
  error: null,
};

const root = document.querySelector<HTMLDivElement>("#app");

if (!root) {
  throw new Error("App root was not found.");
}

Office.onReady(async () => {
  state.message = await getCurrentMessageSnapshot();
  render();
});

function setBusy(isBusy: boolean): void {
  state.isBusy = isBusy;
  render();
}

function setError(error: string | null): void {
  state.error = error;
  render();
}

async function refreshMessage(): Promise<void> {
  setBusy(true);
  try {
    state.message = await getCurrentMessageSnapshot();
    setError(null);
  } catch (error) {
    setError(error instanceof Error ? error.message : "Unable to read the current Outlook message.");
  } finally {
    setBusy(false);
  }
}

async function requestSuggestion(kind: DraftKind): Promise<void> {
  setBusy(true);
  try {
    const message = await getCurrentMessageSnapshot();
    state.message = message;

    const suggestion =
      kind === "reply"
        ? await client.draftReply({ message, tone: state.selectedTone, memory: loadStyleMemory() })
        : await client.rewrite({ message, tone: state.selectedTone, memory: loadStyleMemory() });

    state.suggestion = suggestion;
    rememberInteraction({
      message,
      suggestion,
      tone: state.selectedTone,
      mode: kind,
    });
    setError(null);
  } catch (error) {
    setError(error instanceof Error ? error.message : "Ratonix could not create a suggestion.");
  } finally {
    setBusy(false);
  }
}

async function insertSuggestion(): Promise<void> {
  if (!state.suggestion) {
    return;
  }

  setBusy(true);
  try {
    await insertTextIntoCompose(state.suggestion.text);
    setError(null);
  } catch (error) {
    setError(
      error instanceof Error
        ? error.message
        : "Open a compose or reply window before inserting a suggestion.",
    );
  } finally {
    setBusy(false);
  }
}

function render(): void {
  const memory = loadStyleMemory();
  const hostInfo = getHostInfo();
  const compose = isComposeMode();
  const messageText = state.message?.bodyText.trim() || "";

  root.innerHTML = `
    <section class="shell">
      <header class="hero">
        <div>
          <p class="eyebrow">Ratonix for Outlook</p>
          <h1>AI communication intelligence inside your inbox.</h1>
        </div>
        <span class="host-pill">${escapeHtml(hostInfo.host)} ${escapeHtml(hostInfo.platform)}</span>
      </header>

      <section class="card">
        <div class="section-heading">
          <div>
            <h2>Current message</h2>
            <p>${compose ? "Improve the draft you are writing." : "Draft a thoughtful reply from the message you are reading."}</p>
          </div>
          <button class="ghost" data-action="refresh" ${state.isBusy ? "disabled" : ""}>Refresh</button>
        </div>
        <div class="message-preview">
          ${
            messageText
              ? `<p>${escapeHtml(truncate(messageText, 900))}</p>`
              : `<p class="empty">No message body is available yet. Open an email or start a draft, then refresh.</p>`
          }
        </div>
      </section>

      <section class="card">
        <h2>Tone</h2>
        <div class="tone-grid">
          ${tones
            .map(
              (tone) => `
                <button
                  class="tone ${state.selectedTone === tone ? "selected" : ""}"
                  data-tone="${tone}"
                  ${state.isBusy ? "disabled" : ""}
                >
                  <strong>${labelForTone(tone)}</strong>
                  <span>${descriptionForTone(tone)}</span>
                </button>
              `,
            )
            .join("")}
        </div>
      </section>

      <section class="actions">
        <button class="primary" data-action="rewrite" ${state.isBusy || !messageText ? "disabled" : ""}>
          ${state.isBusy ? "Working..." : "Suggest smarter version"}
        </button>
        <button class="secondary" data-action="reply" ${state.isBusy || !messageText ? "disabled" : ""}>
          Draft reply
        </button>
      </section>

      ${
        state.error
          ? `<div class="notice error" role="alert">${escapeHtml(state.error)}</div>`
          : ""
      }

      <section class="card suggestion-card">
        <div class="section-heading">
          <div>
            <h2>Suggestion</h2>
            <p>${state.suggestion ? escapeHtml(state.suggestion.rationale) : "Choose a tone and ask Ratonix for a rewrite or reply."}</p>
          </div>
          <span class="source">${state.suggestion ? escapeHtml(state.suggestion.source) : "ready"}</span>
        </div>
        <textarea id="suggestion" ${state.suggestion ? "" : "disabled"}>${state.suggestion ? escapeHtml(state.suggestion.text) : ""}</textarea>
        <button class="primary full" data-action="insert" ${!state.suggestion || !compose || state.isBusy ? "disabled" : ""}>
          Insert into Outlook draft
        </button>
        ${!compose ? `<p class="hint">Insertion is available in a compose or reply window. You can still copy the suggestion above.</p>` : ""}
      </section>

      <section class="card memory">
        <h2>Style memory</h2>
        <p>Ratonix has locally remembered ${memory.interactions.length} interaction${memory.interactions.length === 1 ? "" : "s"} on this device.</p>
        <ul>
          ${memory.keyRelationships
            .slice(0, 4)
            .map((relationship) => `<li>${escapeHtml(relationship)}</li>`)
            .join("") || "<li>No frequent relationships yet.</li>"}
        </ul>
      </section>
    </section>
  `;

  root.querySelectorAll<HTMLButtonElement>("[data-tone]").forEach((button) => {
    button.addEventListener("click", () => {
      state.selectedTone = button.dataset.tone as Tone;
      render();
    });
  });

  root.querySelector<HTMLButtonElement>('[data-action="refresh"]')?.addEventListener("click", () => {
    void refreshMessage();
  });
  root.querySelector<HTMLButtonElement>('[data-action="rewrite"]')?.addEventListener("click", () => {
    void requestSuggestion("rewrite");
  });
  root.querySelector<HTMLButtonElement>('[data-action="reply"]')?.addEventListener("click", () => {
    void requestSuggestion("reply");
  });
  root.querySelector<HTMLButtonElement>('[data-action="insert"]')?.addEventListener("click", () => {
    const textarea = root.querySelector<HTMLTextAreaElement>("#suggestion");
    if (textarea && state.suggestion) {
      state.suggestion = { ...state.suggestion, text: textarea.value };
    }
    void insertSuggestion();
  });
}

function labelForTone(tone: Tone): string {
  return tone.charAt(0).toUpperCase() + tone.slice(1);
}

function descriptionForTone(tone: Tone): string {
  switch (tone) {
    case "softer":
      return "Warm, collaborative, lower-friction.";
    case "firmer":
      return "Clear, direct, action-oriented.";
    case "balanced":
      return "Professional, concise, natural.";
  }
}

function truncate(value: string, maxLength: number): string {
  return value.length > maxLength ? `${value.slice(0, maxLength)}...` : value;
}

function escapeHtml(value: string): string {
  const element = document.createElement("div");
  element.textContent = value;
  return element.innerHTML;
}

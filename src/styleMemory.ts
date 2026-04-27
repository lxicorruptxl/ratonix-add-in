import type { MessageSnapshot, RatonixSuggestion, StyleMemory, Tone } from "./types";

const STORAGE_KEY = "ratonix.outlook.styleMemory";
const MAX_INTERACTIONS = 25;
const MAX_RELATIONSHIPS = 25;

const emptyMemory = (): StyleMemory => ({
  examples: [],
  interactions: [],
  keyRelationships: [],
  relationships: {},
});

export function loadStyleMemory(storage: Storage = window.localStorage): StyleMemory {
  const raw = storage.getItem(STORAGE_KEY);
  if (!raw) {
    return emptyMemory();
  }

  try {
    const parsed = JSON.parse(raw) as Partial<StyleMemory>;
    return {
      examples: Array.isArray(parsed.examples) ? parsed.examples : [],
      interactions: Array.isArray(parsed.interactions) ? parsed.interactions : [],
      keyRelationships: Array.isArray(parsed.keyRelationships) ? parsed.keyRelationships : [],
      relationships:
        parsed.relationships && typeof parsed.relationships === "object" && !Array.isArray(parsed.relationships)
          ? parsed.relationships
          : {},
    };
  } catch {
    return emptyMemory();
  }
}

export function rememberInteraction(input: {
  message: MessageSnapshot;
  suggestion: RatonixSuggestion;
  tone: Tone;
  mode: "rewrite" | "reply";
  storage?: Storage;
}): void {
  const storage = input.storage ?? window.localStorage;
  const memory = loadStyleMemory(storage);
  const recipient = input.message.from ?? input.message.to[0];
  const now = new Date().toISOString();

  memory.interactions = [
    {
      createdAt: now,
      tone: input.tone,
      mode: input.mode,
      subject: input.message.subject,
      sourcePreview: input.message.bodyText.slice(0, 240),
      suggestionPreview: input.suggestion.text.slice(0, 240),
    },
    ...memory.interactions,
  ].slice(0, MAX_INTERACTIONS);

  if (input.mode === "rewrite" && input.suggestion.text.trim()) {
    memory.examples = [
      input.suggestion.text.trim().slice(0, 600),
      ...memory.examples.filter((example) => example !== input.suggestion.text.trim()),
    ].slice(0, 10);
  }

  if (recipient) {
    const relationship = recipient.toLowerCase();
    memory.keyRelationships = [
      relationship,
      ...memory.keyRelationships.filter((existing) => existing !== relationship),
    ].slice(0, MAX_RELATIONSHIPS);
    const existing = memory.relationships[relationship];
    memory.relationships[relationship] = {
      address: relationship,
      displayName: existing?.displayName,
      notes: existing?.notes ?? [],
      interactionCount: (existing?.interactionCount ?? 0) + 1,
      updatedAt: now,
    };
  }

  storage.setItem(STORAGE_KEY, JSON.stringify(memory));
}
